VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPIVAParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������Һ�������Ĳ�������"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11310
   Icon            =   "frmPIVAParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPRI 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   3120
      ScaleHeight     =   2055
      ScaleWidth      =   2535
      TabIndex        =   5
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdYes 
         Height          =   360
         Left            =   720
         Picture         =   "frmPIVAParaSet.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   810
      End
      Begin VB.CommandButton cmdNO 
         Height          =   360
         Left            =   1560
         Picture         =   "frmPIVAParaSet.frx":6DDC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   810
      End
      Begin MSComctlLib.ListView lvwPRI 
         Height          =   1305
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "˫���򰴻س���ȷ��"
         Top             =   120
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2302
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ColHdrIcons     =   "imgLvwSel"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   6135
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   9
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "��������(&0)"
      TabPicture(0)   =   "frmPIVAParaSet.frx":6F26
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra��Һ������"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkOpen"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra��ʾ����"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "��������(&1)"
      TabPicture(1)   =   "frmPIVAParaSet.frx":6F42
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "vsfBatch"
      Tab(1).Control(2)=   "cmdDel"
      Tab(1).Control(3)=   "cmdAdd"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "��ӡ����(&2)"
      TabPicture(2)   =   "frmPIVAParaSet.frx":6F5E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "lblNum"
      Tab(2).Control(2)=   "lblƱ��"
      Tab(2).Control(3)=   "cboNum"
      Tab(2).Control(4)=   "cmd��ӡ����"
      Tab(2).Control(5)=   "cboƱ������"
      Tab(2).Control(6)=   "fraƿǩ��ӡ��ʽ"
      Tab(2).Control(7)=   "fra���������ӡ��ʽ"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "���ȼ�����(&3)"
      TabPicture(3)   =   "frmPIVAParaSet.frx":6F7A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblpritip"
      Tab(3).Control(1)=   "vsfPri"
      Tab(3).Control(2)=   "vsfDept"
      Tab(3).Control(3)=   "chkAll"
      Tab(3).Control(4)=   "cmdIN"
      Tab(3).Control(5)=   "cmdDelPri"
      Tab(3).Control(6)=   "cmdAddPri"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "��������(&4)"
      TabPicture(4)   =   "frmPIVAParaSet.frx":6F96
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblvoltip"
      Tab(4).Control(1)=   "vsfVolume"
      Tab(4).Control(2)=   "cmdVolAdd"
      Tab(4).Control(3)=   "cmdVolDel"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "����ҩƷ����(&5)"
      TabPicture(5)   =   "frmPIVAParaSet.frx":6FB2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblMedi"
      Tab(5).Control(1)=   "vsfPrint"
      Tab(5).Control(2)=   "chkByMedi"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "������ҩƷ����(&6)"
      TabPicture(6)   =   "frmPIVAParaSet.frx":6FCE
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblNoMedi"
      Tab(6).Control(1)=   "vsfNoMedi"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "��ʾ��Դ����(&7)"
      TabPicture(7)   =   "frmPIVAParaSet.frx":6FEA
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Lvw��Դ����"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      Begin VB.Frame fra��ʾ���� 
         Caption         =   " ��ʾ���� "
         Height          =   855
         Left            =   120
         TabIndex        =   69
         Top             =   3720
         Width           =   8775
         Begin VB.ComboBox cboҩƷ������ʾ��ʽ 
            Height          =   300
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   300
            Width           =   2295
         End
         Begin VB.Label lalҩƷ������ʾ��ʽ 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ������ʾ��ʽ"
            Height          =   180
            Left            =   360
            TabIndex        =   70
            Top             =   360
            Width           =   1440
         End
      End
      Begin VB.CheckBox chkByMedi 
         Caption         =   "�Ƿ�������õĳ���ҩƷ����ҩƷ���˲���"
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   3855
      End
      Begin VB.Frame fra���������ӡ��ʽ 
         Caption         =   " ���������ӡ��ʽ"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   55
         Top             =   2280
         Width           =   6255
         Begin VB.ComboBox cbo��ҩ�� 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   660
            Width           =   2415
         End
         Begin VB.ComboBox cbo���͵� 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1035
            Width           =   2415
         End
         Begin VB.ComboBox cboSum 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "����ȷ�Ϻ�"
            Height          =   180
            Left            =   120
            TabIndex        =   64
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��ҩȷ�Ϻ�"
            Height          =   180
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ�����嵥"
            Height          =   180
            Left            =   3720
            TabIndex        =   62
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "���ܷ����嵥"
            Height          =   180
            Left            =   3720
            TabIndex        =   61
            Top             =   1095
            Width           =   1080
         End
         Begin VB.Label lblSum 
            AutoSize        =   -1  'True
            Caption         =   "���ܱ���"
            Height          =   180
            Left            =   3720
            TabIndex        =   60
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblSumPrint 
            AutoSize        =   -1  'True
            Caption         =   "��ӡ��ǩ��"
            Height          =   180
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame fraƿǩ��ӡ��ʽ 
         Caption         =   " ƿǩ��ӡ��ʽ"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   49
         Top             =   480
         Width           =   6255
         Begin VB.CheckBox chkManPrint 
            Caption         =   "�����ֹ����ƴ�ӡƿǩ���ɽ��в���"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1080
            Width           =   3375
         End
         Begin VB.ComboBox cbo��ǩ��ӡ 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkPrintLabelStep 
            Caption         =   "��ҩ��"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkPrintLabelStep 
            Caption         =   "��ҩ��"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   660
            Width           =   855
         End
         Begin VB.ComboBox cbo��ǩ��ӡ 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.ComboBox cboƱ������ 
         Height          =   300
         Left            =   -73500
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   4305
         Width           =   2415
      End
      Begin VB.CommandButton cmd��ӡ���� 
         Caption         =   "��ӡ����(&P)"
         Height          =   345
         Left            =   -70980
         TabIndex        =   44
         Top             =   4275
         Width           =   1155
      End
      Begin VB.ComboBox cboNum 
         Height          =   300
         Left            =   -73500
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Caption         =   " ��Ƭ���� "
         Height          =   615
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Width           =   8800
         Begin VB.ComboBox cbo���� 
            Height          =   300
            ItemData        =   "frmPIVAParaSet.frx":7006
            Left            =   1080
            List            =   "frmPIVAParaSet.frx":7013
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl����2 
            AutoSize        =   -1  'True
            Caption         =   "�ſ�Ƭ"
            Height          =   180
            Left            =   2040
            TabIndex        =   42
            Top             =   300
            Width           =   540
         End
         Begin VB.Label lbl����1 
            Caption         =   "������ʾ"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.CheckBox chkOpen 
         Caption         =   "���ý���ʱ��ο���"
         Height          =   180
         Left            =   360
         TabIndex        =   27
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Frame fra��Һ������ 
         Height          =   1215
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   8800
         Begin VB.CheckBox chk����ҽ�� 
            Caption         =   "���յ��ռ���ǰ��ҽ��"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txtDeff 
            Enabled         =   0   'False
            Height          =   270
            Left            =   3000
            TabIndex        =   30
            Text            =   "0"
            Top             =   795
            Width           =   375
         End
         Begin MSComCtl2.UpDown updDeff 
            Height          =   270
            Left            =   3480
            TabIndex        =   29
            Top             =   795
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   315
            Left            =   960
            TabIndex        =   32
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   84934658
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   315
            Left            =   3240
            TabIndex        =   33
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   84934658
            CurrentDate     =   36985
         End
         Begin VB.Label lblʱ����� 
            AutoSize        =   -1  'True
            Caption         =   "ҽ�����Ͳ��ڸ�ʱ�����Һҽ�������ٲ�����Һ����"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   4560
            TabIndex        =   38
            Top             =   360
            Width           =   4140
         End
         Begin VB.Label lbl����ҽ�� 
            AutoSize        =   -1  'True
            Caption         =   "��ѡʱ�������Ľ���������ʱ��������ĵ���ִ�е�ҽ����"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   3840
            TabIndex        =   37
            Top             =   840
            Width           =   4680
         End
         Begin VB.Label lblBegin 
            Caption         =   "��ʼʱ��"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblEnd 
            Caption         =   "����ʱ��"
            Height          =   255
            Left            =   2400
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblDeff 
            Caption         =   "Сʱ��"
            Height          =   255
            Left            =   2355
            TabIndex        =   34
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdVolDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   20
         Top             =   1560
         Width           =   1100
      End
      Begin VB.CommandButton cmdVolAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   19
         Top             =   1080
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddPri 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   15
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdDelPri 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   14
         Top             =   2400
         Width           =   1100
      End
      Begin VB.CommandButton cmdIN 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   -66600
         TabIndex        =   13
         Top             =   1800
         Width           =   1100
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Ӧ�������п��ҵ����ȼ�����"
         Height          =   250
         Left            =   -74880
         TabIndex        =   12
         Top             =   720
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Caption         =   " �������Ŀⷿѡ�� "
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   8800
         Begin VB.CheckBox chkCheck 
            Caption         =   "��˸�ҩ��������ҽ��"
            Height          =   255
            Left            =   4080
            TabIndex        =   67
            Top             =   240
            Width           =   3855
         End
         Begin VB.ComboBox CboStore 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   2280
         End
         Begin VB.Label lblStore 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Left            =   360
            TabIndex        =   11
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   -66480
         TabIndex        =   4
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   -66480
         TabIndex        =   3
         Top             =   1560
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDept 
         Height          =   4905
         Left            =   -74880
         TabIndex        =   16
         Top             =   1080
         Width           =   2400
         _cx             =   4233
         _cy             =   8652
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":7020
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPri 
         Height          =   4905
         Left            =   -72360
         TabIndex        =   17
         Top             =   1080
         Width           =   5505
         _cx             =   9710
         _cy             =   8652
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":70B6
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
      Begin VSFlex8Ctl.VSFlexGrid vsfVolume 
         Height          =   5145
         Left            =   -74880
         TabIndex        =   21
         Top             =   840
         Width           =   7995
         _cx             =   14102
         _cy             =   9075
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":716F
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   23
         Top             =   960
         Width           =   7995
         _cx             =   14102
         _cy             =   8864
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":7213
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
      Begin VSFlex8Ctl.VSFlexGrid vsfNoMedi 
         Height          =   5145
         Left            =   -74880
         TabIndex        =   25
         Top             =   840
         Width           =   8000
         _cx             =   14111
         _cy             =   9075
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":727C
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
      Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   66
         Top             =   960
         Width           =   8160
         _cx             =   14393
         _cy             =   8864
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":72E5
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
      Begin MSComctlLib.ListView Lvw��Դ���� 
         Height          =   5445
         Left            =   -74880
         TabIndex        =   72
         Top             =   480
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   9604
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�����������Ĺ�������(0������Ϊ�������δ��ڣ�������ҩʱ�䷶Χ����)"
         Height          =   180
         Left            =   -74760
         TabIndex        =   68
         Top             =   600
         Width           =   5850
      End
      Begin VB.Label lblƱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺͱ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74640
         TabIndex        =   48
         Top             =   4365
         Width           =   900
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "ƿǩ��ӡ����"
         Height          =   180
         Left            =   -74640
         TabIndex        =   47
         Top             =   4860
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�����ڰ�ҩ����ҩ���ӡ"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -70980
         TabIndex        =   46
         Top             =   4860
         Width           =   1980
      End
      Begin VB.Label lblNoMedi 
         AutoSize        =   -1  'True
         Caption         =   "�����������Ĳ��������õ�ҩƷ"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   2520
      End
      Begin VB.Label lblMedi 
         AutoSize        =   -1  'True
         Caption         =   "���ó���ҩƷ������Һ��������԰�ҩƷ���й��˺�����"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   4500
      End
      Begin VB.Label lblvoltip 
         AutoSize        =   -1  'True
         Caption         =   "����ĳ�����ҵ�������ĳ�����ο�����ҩ������"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   3780
      End
      Begin VB.Label lblpritip 
         AutoSize        =   -1  'True
         Caption         =   "��������ͬ��������ͬ��ҩƷ�����ȼ�"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   3060
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10170
      TabIndex        =   1
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8760
      TabIndex        =   0
      Top             =   6360
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgLvwSel 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmPIVAParaSet.frx":7479
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":7793
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":7AAD
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":7DFF
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   5880
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPIVAParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String                              'Ȩ�޴�
Public mlng�ⷿid As Long
Private mblnSetPara As Boolean
Private mRsDept As Recordset
Private mRsPC As Recordset
Private mRsType As Recordset
Private mintRow As Integer
Private mintCol As Integer
Private mblnPri As Boolean
Private mblnEdit As Boolean     '�Ƿ�༭���ȼ�
Private mstrSourceDep As String '��Դ����

Private Sub LoadStore()
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.id,B.���� From ��������˵�� A,���ű� B" & _
    " Where A.����ID=B.ID And A.��������='��������' And B.Id In (Select ����id From ������Ա Where ��Աid = [1])"
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������ĵĲ���", glngUserId)
    
    With Me.CboStore
        Do While Not rstemp.EOF
            .AddItem rstemp!����
            .ItemData(.NewIndex) = rstemp!Id
            rstemp.MoveNext
        Loop
        If rstemp.RecordCount > 0 Then .ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load��Դ����()
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select ���� || '-' || ���� ����, Id " & _
            " From ���ű� " & _
            " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And Id In (Select ����id From ��������˵�� Where �������� = '����' And ������� In (2,3)) And " & _
            " (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By ���� || '-' || ���� "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Load��Դ����")
    
    '��ʼ������
    With rsData
        If .EOF Then
            MsgBox "û�����ø��ಿ�ţ������Ź���", vbInformation, gstrSysName
            Exit Sub
        End If
        Lvw��Դ����.ListItems.Clear
        Do While Not .EOF
            Lvw��Դ����.ListItems.Add , "_" & !Id, !����, 1, 1
            If mstrSourceDep <> "" Then
                If InStr("," & mstrSourceDep & ",", "," & CStr(!Id) & ",") > 0 Then
                    Lvw��Դ����.ListItems("_" & !Id).Checked = True
                End If
            End If
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadParams()
    Dim int��ҩ�� As Integer
    Dim int��ҩ�� As Integer
    Dim int�������� As Integer
    Dim int�ϴ����� As Integer
    Dim int������� As Integer
    Dim intBarCode As Integer
    Dim strAutoPrint As String
    Dim intManPrint As Integer
    Dim str��ֹʱ�� As String
    Dim intҽ������ As Integer
    Dim str��Һ��ҩ;�� As String
    Dim str��Դ���� As String
    Dim rsData As ADODB.Recordset
    Dim str����ҽ�� As String
    Dim intCount As Integer
    Dim intOpen As Integer
    Dim lng����ID As Long
    Dim IntLocate As Integer
    Dim dateNow As Date
    Dim intNum As Integer
    Dim int��ҩ���� As Integer
    Dim i As Integer
    Dim int���� As Integer
    Dim intTPN As Integer
    Dim intSpecial As Integer
    
    On Error GoTo errHandle
    '����
    int��ҩ�� = Val(zlDatabase.GetPara("��ҩ���ӡ", glngSys, 1345, 0, Array(Label3, cbo��ҩ��, Label5), mblnSetPara))
    int��ҩ�� = Val(zlDatabase.GetPara("���ͺ��ӡ", glngSys, 1345, 0, Array(Label4, cbo���͵�, Label6), mblnSetPara))
    
    strAutoPrint = zlDatabase.GetPara("ƿǩ�Զ���ӡ", glngSys, 1345, "00|00", Array(chkPrintLabelStep(0), chkPrintLabelStep(1)), mblnSetPara)
    intManPrint = Val(zlDatabase.GetPara("ƿǩ�ֹ���ӡ", glngSys, 1345, "0", Array(chkManPrint), mblnSetPara))
    intCount = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\��Һ��Ƭ", "��Ƭ����", 3))
    
    int���� = Val(zlDatabase.GetPara("��ӡ��ǩ���Ƿ��ӡ���ܱ���", glngSys, 1345, 0, Array(lblSumPrint, cboSum, lblSum), mblnSetPara))
    
    Me.cboҩƷ������ʾ��ʽ.ListIndex = Val(zlDatabase.GetPara("ҩƷ������ʾ��ʽ", glngSys, 1345, 0, Array(lalҩƷ������ʾ��ʽ, cboҩƷ������ʾ��ʽ), mblnSetPara))
    
    '��������
    str��ֹʱ�� = zlDatabase.GetPara("������ֹʱ��", glngSys, 1345, "", Array(lblBegin, dtpBegin, lblEnd, dtpEnd), mblnSetPara)
    str����ҽ�� = zlDatabase.GetPara("�����յ��ռ���ǰҽ��", glngSys, 1345, 0, Array(chk����ҽ��, txtDeff, updDeff, lblDeff), mblnSetPara)
    
    intOpen = Val(zlDatabase.GetPara("���ý���ʱ�����", glngSys, 1345, 0, Array(chkOpen), mblnSetPara))
    lng����ID = Val(zlDatabase.GetPara("��������", glngSys, 1345, 0, Array(CboStore, lblStore), mblnSetPara))
    intNum = Val(zlDatabase.GetPara("ƿǩ��ӡ����", glngSys, 1345, 1, Array(lblNum, cboNum), mblnSetPara))
    Me.chkByMedi.Value = Val(zlDatabase.GetPara("�Ƿ����õĳ���ҩƷ����ҩƷ���˲���", glngSys, 1345, 0, Array(chkByMedi), mblnSetPara))
    Me.chkCheck.Value = Val(zlDatabase.GetPara("��˸�ҩ������������", glngSys, 1345, 0, Array(chkCheck), mblnSetPara))

    '��ʾ��Դ����
    mstrSourceDep = zlDatabase.GetPara("��ʾ��Դ����", glngSys, 1345, "")

    If lng����ID <> 0 Then                                  '��λҩ��
        '�����ڸ�ҩ������ʾ
        For IntLocate = 0 To Me.CboStore.ListCount - 1
            If Me.CboStore.ItemData(IntLocate) = lng����ID Then
                Me.CboStore.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (CboStore.ListCount - 1) Then
            MsgBox "�����������������ģ�ԭ�����õ�����������ʧЧ����", vbInformation, gstrSysName
            If CboStore.ListCount >= 1 Then CboStore.ListIndex = 0
        End If
    Else
        MsgBox "�������������ģ�", vbInformation, gstrSysName
    End If
    
    Me.chkOpen.Value = intOpen
    
    If InStr(1, str��ֹʱ��, "|") > 0 Then
        Me.dtpBegin.Value = Mid(str��ֹʱ��, 1, InStr(1, str��ֹʱ��, "|") - 1)
        Me.dtpEnd.Value = Mid(str��ֹʱ��, InStr(1, str��ֹʱ��, "|") + 1)
    End If
    
    Me.chk����ҽ��.Value = Mid(str����ҽ��, 1, 1)
    If InStr(1, str����ҽ��, "|") > 1 Then
        Me.txtDeff.Text = Mid(str����ҽ��, 3)
    Else
        Me.txtDeff.Text = 0
    End If
    
    ''��������
    If int��ҩ�� >= 0 And int��ҩ�� <= cbo��ҩ��.ListCount - 1 Then
        cbo��ҩ��.ListIndex = int��ҩ��
    End If
    
    If int���� >= 0 And int���� <= cboSum.ListCount - 1 Then
        cboSum.ListIndex = int����
    End If
    
    If int��ҩ�� >= 0 And int��ҩ�� <= cbo��ҩ��.ListCount - 1 Then
        cbo���͵�.ListIndex = int��ҩ��
    End If
    
    If InStr(1, strAutoPrint, "|") = 0 Or Len(strAutoPrint) <> 5 Then
        strAutoPrint = "00|00"
    End If
    
    If Mid(strAutoPrint, 1, 1) = 1 Then
        chkPrintLabelStep(0).Value = 1
        If Val(Mid(strAutoPrint, 2, 1)) = 1 Then
            cbo��ǩ��ӡ(0).ListIndex = 1
        Else
            cbo��ǩ��ӡ(0).ListIndex = 0
        End If
    End If
    
    If Mid(strAutoPrint, 4, 1) = 1 Then
        chkPrintLabelStep(1).Value = 1
        If Val(Mid(strAutoPrint, 5, 1)) = 1 Then
            cbo��ǩ��ӡ(1).ListIndex = 1
        Else
            cbo��ǩ��ӡ(1).ListIndex = 0
        End If
    End If
    
    cbo��ǩ��ӡ(0).Enabled = chkPrintLabelStep(0).Enabled And (chkPrintLabelStep(0).Value = 1)
    cbo��ǩ��ӡ(1).Enabled = chkPrintLabelStep(1).Enabled And (chkPrintLabelStep(1).Value = 1)
    
    vsfVolume.Enabled = mblnSetPara
    vsfPrint.Enabled = mblnSetPara
    vsfNoMedi.Enabled = mblnSetPara
    vsfPri.Enabled = mblnSetPara
    cmdAddPri.Enabled = mblnSetPara
    cmdIN.Enabled = mblnSetPara
    cmdDelPri.Enabled = mblnSetPara
    cmdVolAdd.Enabled = mblnSetPara
    cmdVolDel.Enabled = mblnSetPara
    
    
    If intManPrint < 0 Or intManPrint > 1 Then
        chkManPrint.Value = 0
    Else
        chkManPrint.Value = intManPrint
    End If
    
    If chkManPrint.Value = 1 Then
        cboSum.Enabled = True
    Else
        cboSum.Enabled = False
    End If
    
    '��Ƭ����
    Me.cbo����.Text = IIf(intCount = 0, 3, intCount)
    
    Me.cboNum.Text = IIf(intNum = 0, 3, intNum)

    '����ҩƷ��ӡ����
    gstrSQL = "select ҩƷid,���� from ��Һ���ȴ�ӡҩƷ"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "LoadҩƷ")
    
    Me.vsfPrint.rows = rsData.RecordCount + 2
    For i = 1 To rsData.RecordCount
        Me.vsfPrint.TextMatrix(i, vsfPrint.ColIndex("ҩƷid")) = rsData!ҩƷID
        Me.vsfPrint.TextMatrix(i, vsfPrint.ColIndex("ҩƷ���������")) = rsData!����
       
       rsData.MoveNext
    Next
    
    
    '��Һ������ҩƷ
    gstrSQL = "select ҩƷid,���� from ��Һ������ҩƷ"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "LoadҩƷ")
    
    Me.vsfNoMedi.rows = rsData.RecordCount + 2
    For i = 1 To rsData.RecordCount
        Me.vsfNoMedi.TextMatrix(i, vsfNoMedi.ColIndex("ҩƷid")) = rsData!ҩƷID
        Me.vsfNoMedi.TextMatrix(i, vsfNoMedi.ColIndex("ҩƷ���������")) = rsData!����
       
       rsData.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CboStore_Click()
    Call LoadBatchSet
    Call loadVolume
End Sub

Private Sub chkAll_Click()
    If mblnEdit Then
        If MsgBox("�뱣�����õ����ȼ����л����Һ����������ȼ����ý�ʧЧ���Ƿ��л���", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
           If Me.chkAll.Value = 0 Then
                Me.vsfPri.Left = Me.vsfDept.Width + Me.vsfDept.Left + 100
                Me.vsfPri.Width = Me.vsfPri.Width - Me.vsfDept.Width - 100
                Me.vsfDept.Visible = True
                Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))))
            Else
                Me.vsfPri.Width = Me.vsfPri.Width + Me.vsfDept.Width + 100
                Me.vsfPri.Left = Me.vsfDept.Left
                Me.vsfDept.Visible = False
                
                Call LoadVsfPRI(0)
            End If
            mblnEdit = False
            
        End If
    Else
        If Me.chkAll.Value = 0 Then
            Me.vsfPri.Left = Me.vsfDept.Width + Me.vsfDept.Left + 100
            Me.vsfPri.Width = Me.vsfPri.Width - Me.vsfDept.Width - 100
            Me.vsfDept.Visible = True
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))))
        Else
            Me.vsfPri.Width = Me.vsfPri.Width + Me.vsfDept.Width + 100
            Me.vsfPri.Left = Me.vsfDept.Left
            Me.vsfDept.Visible = False
            
            Call LoadVsfPRI(0)
        End If
    End If
End Sub

Private Sub chkManPrint_Click()
    If chkManPrint.Value = 1 Then
        cboSum.Enabled = True
    Else
        cboSum.Enabled = False
    End If
End Sub

Private Sub chkOpen_Click()
    Me.dtpBegin.Enabled = (Me.chkOpen.Value = 1)
    Me.dtpEnd.Enabled = (Me.chkOpen.Value = 1)
    Me.chk����ҽ��.Enabled = (Me.chkOpen.Value = 1)
    Me.updDeff.Enabled = (Me.chkOpen.Value = 1)
End Sub

Private Sub chkPrintLabelStep_Click(index As Integer)
    cbo��ǩ��ӡ(index).Enabled = (chkPrintLabelStep(index).Value = 1)
End Sub

Private Sub cmdAdd_Click()
    With vsfBatch
        If .rows > 2 Then
            If Trim(.TextMatrix(.rows - 1, .ColIndex("����ʱ�俪ʼ"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("����ʱ�����"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("��ҩʱ�俪ʼ"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("��ҩʱ�����"))) = "" Then
                Exit Sub
            End If
        End If
        
        .rows = .rows + 1
        
        If .rows > 3 Then
            .TextMatrix(.rows - 1, .ColIndex("����")) = Mid(.TextMatrix(.rows - 2, .ColIndex("����")), 1, Len(.TextMatrix(.rows - 2, .ColIndex("����"))) - 1) + 1 & "#"
        Else
            .TextMatrix(.rows - 1, .ColIndex("����")) = "0#"
        End If
        .TextMatrix(.rows - 1, .ColIndex("����")) = "��"
    End With
End Sub

Private Sub cmdAddPri_Click()
    If Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, Me.vsfPri.ColIndex("��ҩ����")) <> "" And Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, Me.vsfPri.ColIndex("Ƶ��")) <> "" Then
        Me.vsfPri.rows = Me.vsfPri.rows + 1
        Me.vsfPri.RowHeight(Me.vsfPri.rows - 1) = 250
        Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, vsfPri.ColIndex("���")) = Me.vsfPri.rows - 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim lngRow As Long
    Dim lngCur As Long
    
    With vsfBatch
        If .Row > 1 Then
            If MsgBox("�Ƿ�ɾ������(" & .TextMatrix(.Row, .ColIndex("����")) & ")��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            .Redraw = flexRDNone
            
            lngCur = .Row
            .RemoveItem .Row
            
            '�������κ�
            For lngRow = lngCur To .rows - 1
                .TextMatrix(lngRow, .ColIndex("����")) = Mid(.TextMatrix(lngRow, .ColIndex("����")), 1, Len(.TextMatrix(lngRow, .ColIndex("����"))) - 1) - 1 & "#"
            Next
            
            .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub cmdDelPri_Click()
    Dim i As Integer
    Dim intRow As Integer
    
    If Me.vsfPri.Row = 0 Then Exit Sub
    intRow = Me.vsfPri.Row
    Me.vsfPri.RemoveItem Me.vsfPri.Row
    
    '�������
    For i = intRow To Me.vsfPri.rows - 1
        Me.vsfPri.TextMatrix(i, Me.vsfPri.ColIndex("���")) = i
    Next
    
    mblnEdit = True
End Sub

Private Sub cmdIN_Click()
    Dim intCount As Integer
    Dim lngRow As Long
    
    If mblnSetPara Then
         '�������ȼ�����
        With vsfPri
            intCount = 1
            
            If .rows = 1 Then
                gstrSQL = "Zl_��ҺҩƷ���ȼ�_Save("
                '����id
                gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, 0, vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))) & "'"
                gstrSQL = gstrSQL & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "�������ȼ�")
            End If
            
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("��ҩ����")) <> "" And .TextMatrix(lngRow, .ColIndex("Ƶ��")) <> "" Then
                    
                    gstrSQL = "Zl_��ҺҩƷ���ȼ�_Save("
                    '����id
                    gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, 0, vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))) & "',"
                    '��������
                    gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, "���п���", vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("��������"))) & "',"
                    '��ҩ����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("��ҩ����")) & "',"
                    'Ƶ��
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("Ƶ��")) & "',"
                    '��Ч
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("��Ч"))) & ","
                    '���ȼ�
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("���")))
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������ȼ�")
                    intCount = intCount + 1
                End If
            Next
        End With
    End If
    
    mblnEdit = False
End Sub

Private Sub cmdNo_Click()
    picPRI.Visible = False
    CmdOK.Enabled = True
    CmdCancel.Enabled = True
End Sub

Private Sub cmdOk_Click()
    Dim strInput As String
    Dim lngRow As Long
    Dim strPrintLabel As String
    Dim intCount As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    
    'ƿǩ��ӡ��ʽ
    If chkPrintLabelStep(0).Value = 0 Then
        strPrintLabel = "00"
    Else
        strPrintLabel = "1" & cbo��ǩ��ӡ(0).ListIndex
    End If
    strPrintLabel = strPrintLabel & "|"
    If chkPrintLabelStep(1).Value = 0 Then
        strPrintLabel = strPrintLabel & "00"
    Else
        strPrintLabel = strPrintLabel & "1" & cbo��ǩ��ӡ(1).ListIndex
    End If

    '��ʾ��Դ����
    mstrSourceDep = ""
    With Me.Lvw��Դ����
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked Then
                If mstrSourceDep = "" Then
                    mstrSourceDep = Mid(.ListItems(i).Key, 2)
                Else
                    mstrSourceDep = mstrSourceDep & "," & Mid(.ListItems(i).Key, 2)
                End If
            End If
        Next
    End With

    '����˽�в���
    '��������
    zlDatabase.SetPara "��ҩ���ӡ", cbo��ҩ��.ListIndex, glngSys, 1345
    zlDatabase.SetPara "���ͺ��ӡ", cbo���͵�.ListIndex, glngSys, 1345
    zlDatabase.SetPara "ƿǩ�Զ���ӡ", strPrintLabel, glngSys, 1345
    zlDatabase.SetPara "ƿǩ�ֹ���ӡ", chkManPrint.Value, glngSys, 1345
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\��Һ��Ƭ", "��Ƭ����", Me.cbo����.Text
    zlDatabase.SetPara "��ӡ��ǩ���Ƿ��ӡ���ܱ���", cboSum.ListIndex, glngSys, 1345
    zlDatabase.SetPara "ҩƷ������ʾ��ʽ", cboҩƷ������ʾ��ʽ.ListIndex, glngSys, 1345
    
    '��������
    zlDatabase.SetPara "������ֹʱ��", Format(dtpBegin.Value, "hh:mm:ss") & "|" & Format(Me.dtpEnd.Value, "hh:mm:ss"), glngSys, 1345
    zlDatabase.SetPara "�����յ��ռ���ǰҽ��", chk����ҽ��.Value & "|" & Me.txtDeff.Text, glngSys, 1345
    zlDatabase.SetPara "���ý���ʱ�����", chkOpen.Value, glngSys, 1345
    zlDatabase.SetPara "��������", Me.CboStore.ItemData(Me.CboStore.ListIndex), glngSys, 1345
    zlDatabase.SetPara "ƿǩ��ӡ����", Me.cboNum.Text, glngSys, 1345
    zlDatabase.SetPara "�Ƿ����õĳ���ҩƷ����ҩƷ���˲���", chkByMedi.Value, glngSys, 1345
    zlDatabase.SetPara "��˸�ҩ������������", chkCheck.Value, glngSys, 1345
    
    If zlStr.IsHavePrivs(mstrPrivs, "���ù�������") Then
        With vsfBatch
            For lngRow = 2 To .rows - 1
                If IsDate(.TextMatrix(lngRow, .ColIndex("����ʱ�俪ʼ"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("����ʱ�����"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("��ҩʱ�俪ʼ"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("��ҩʱ�����"))) Then
                    
                    strInput = IIf(strInput = "", "", strInput & "|") & _
                        Mid(.TextMatrix(lngRow, .ColIndex("����")), 1, Len(.TextMatrix(lngRow, .ColIndex("����"))) - 1) & "," & _
                        .TextMatrix(lngRow, .ColIndex("����ʱ�俪ʼ")) & "-" & .TextMatrix(lngRow, .ColIndex("����ʱ�����")) & "," & _
                        .TextMatrix(lngRow, .ColIndex("��ҩʱ�俪ʼ")) & "-" & .TextMatrix(lngRow, .ColIndex("��ҩʱ�����")) & "," & _
                        IIf(.TextMatrix(lngRow, .ColIndex("���")) = "", 0, 1) & "," & _
                        IIf(.TextMatrix(lngRow, .ColIndex("����")) = "", 0, 1) & "," & _
                        .Cell(flexcpBackColor, lngRow, .ColIndex("��ɫ")) & "," & _
                        IIf(Trim(.TextMatrix(lngRow, .ColIndex("ҩƷ����"))) = "", Null, .TextMatrix(lngRow, .ColIndex("ҩƷ����")))
                End If
            Next
        End With
        
        '���strInputΪ�ձ�ʾɾ��������������
        gstrSQL = "Zl_��ҩ��������_Save("
        '������Ϣ
        gstrSQL = gstrSQL & "'" & strInput & "',"
        gstrSQL = gstrSQL & Me.CboStore.ItemData(Me.CboStore.ListIndex)
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������ҩ��������")
    End If
    
    '��ʾ��Դ����
    zlDatabase.SetPara "��ʾ��Դ����", mstrSourceDep, glngSys, 1345

    If mblnSetPara Then
        '������������
        With Me.vsfVolume
            For lngRow = 0 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("��������")) <> "" And .TextMatrix(lngRow, .ColIndex("����")) <> "" Then
                    
                    gstrSQL = "Zl_������������_Save("
                    '����id
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("����id")) & "',"
                    '��������
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("��������")) & "',"
                    '����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("��ҩ����")) & "',"
                    '����
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("����"))) & ","
                    '���ȼ�
                    gstrSQL = gstrSQL & lngRow & ","
                    '��������ID
                    gstrSQL = gstrSQL & Me.CboStore.ItemData(Me.CboStore.ListIndex)
                    gstrSQL = gstrSQL & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
                End If
            Next
        End With
        
        '���泣��ҩƷ
        With Me.vsfPrint
            For i = 1 To .rows - 1
                If (.TextMatrix(i, .ColIndex("ҩƷid")) <> "" And .TextMatrix(i, .ColIndex("ҩƷ���������")) <> "") Or i = 1 Then
                    gstrSQL = "Zl_��Һ���ȴ�ӡҩƷ_��ӡ����("
                    'ҩƷid
                    gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("ҩƷid"))) & ","
                    'ҩƷ����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("ҩƷ���������")) & "',"
                    gstrSQL = gstrSQL & i & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "���泣��ҩƷ")
                End If
            Next
        End With
        
        '���治����ҩƷ
        With Me.vsfNoMedi
            For i = 1 To .rows - 1
                If (.TextMatrix(i, .ColIndex("ҩƷid")) <> "" And .TextMatrix(i, .ColIndex("ҩƷ���������")) <> "") Or i = 1 Then
                    gstrSQL = "Zl_��Һ������ҩƷ_����("
                    'ҩƷid
                    gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("ҩƷid"))) & ","
                    'ҩƷ����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("ҩƷ���������")) & "',"
                    gstrSQL = gstrSQL & i & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "���治����ҩƷ")
                End If
            Next
        End With
    End If
    
    frmPIVAMain.mblnParamsRefresh = True
    
    Unload Me
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadVsfPRI(ByVal str����id As String)
    Dim rstemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select ����id,��������,��ҩ����,Ƶ��,��Ч,���ȼ� from ��ҺҩƷ���ȼ� where (����id=[1] or ����id='0') order by ���ȼ�"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ȼ�����", str����id)
    
    i = 1
    rstemp.Filter = "����id='" & str����id & "'"
    If rstemp.EOF Then rstemp.Filter = ""
    With Me.vsfPri
        .RowHeight(0) = 250
        
        If rstemp.RecordCount = 0 Then
            .rows = 1
            .rows = 2
            .TextMatrix(1, .ColIndex("���")) = 1
        Else
            .rows = rstemp.RecordCount + 1
        End If
       
        Do While Not rstemp.EOF
            .RowHeight(i) = 250
            .TextMatrix(i, .ColIndex("���")) = rstemp!���ȼ�
            .TextMatrix(i, .ColIndex("��ҩ����")) = rstemp!��ҩ����
            .TextMatrix(i, .ColIndex("Ƶ��")) = rstemp!Ƶ��
            .TextMatrix(i, .ColIndex("��Ч")) = rstemp!��Ч
            i = i + 1
            rstemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdVolAdd_Click()
    If Me.vsfVolume.TextMatrix(Me.vsfVolume.rows - 1, Me.vsfVolume.ColIndex("��������")) <> "" And Me.vsfVolume.TextMatrix(Me.vsfVolume.rows - 1, Me.vsfVolume.ColIndex("����")) <> "" Then
        Me.vsfVolume.rows = Me.vsfVolume.rows + 1
        Me.vsfVolume.RowHeight(Me.vsfVolume.rows - 1) = 250
    End If
End Sub

Private Sub cmdVolDel_Click()
    If Me.vsfVolume.Row = 0 Then Exit Sub
    Me.vsfVolume.RemoveItem Me.vsfVolume.Row
End Sub

Private Sub cmdYes_Click()
    Dim strIDS As String
    Dim strReturn As String
    
    strReturn = ReturnSelectedPri(1, strIDS)
    
    If mblnPri Then
        With Me.vsfPri
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("��������") Then
                .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
            End If
        End With
    Else
        With Me.vsfVolume
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("��������") Then
                .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
            End If
        End With
    End If
    
End Sub

Private Sub cmd��ӡ����_Click()
    Dim strBill As String
    Select Case cboƱ������.ListIndex
    Case 0
        '��Һƿ��ǩ
        strBill = "ZL1_BILL_1345_1"
    Case 1
        '��ҩҩƷ�����嵥
        strBill = "ZL1_INSIDE_1345_1"
    Case 2
        '����ҩƷ�����嵥
        strBill = "ZL1_INSIDE_1345_2"
    Case 3
        '��ҩ�����嵥
        strBill = "ZL1_BILL_1345_2"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Load()
    mblnSetPara = zlStr.IsHavePrivs(mstrPrivs, "��������")
    
    With cbo��ǩ��ӡ(0)
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .ListIndex = 0
    End With
    
    With cbo��ǩ��ӡ(1)
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .ListIndex = 0
    End With
    
    With cbo��ҩ��
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .AddItem "2-����ӡ"
    End With
    
    With cbo���͵�
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .AddItem "2-����ӡ"
    End With
    
    With cboSum
        .Clear
        .AddItem "0-��ʾ�Ƿ��ӡ"
        .AddItem "1-�Զ���ӡ"
        .AddItem "2-����ӡ"
    End With
    
    With cboƱ������
        .Clear
        .AddItem "1-��Һƿ��ǩ"
        .AddItem "2-��ҩҩƷ�����嵥"
        .AddItem "3-����ҩƷ�����嵥"
        .AddItem "4-��ҩ�����嵥"

        .ListIndex = 0
    End With
    
    With cboNum
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        
        .ListIndex = 0
    End With
        
    With vsfBatch
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(.ColIndex("����")) = True
        .MergeCol(.ColIndex("��ɫ")) = True
        .MergeCol(.ColIndex("����ʱ�俪ʼ")) = True
        .MergeCol(.ColIndex("����ʱ�����")) = True
        .MergeCol(.ColIndex("��ҩʱ�俪ʼ")) = True
        .MergeCol(.ColIndex("��ҩʱ�����")) = True
        .MergeCol(.ColIndex("���")) = True
        .MergeCol(.ColIndex("����")) = True
        .MergeCol(.ColIndex("ҩƷ����")) = True
        .MergeCells = flexMergeFixedOnly
    End With
    
    With cboҩƷ������ʾ��ʽ
        .Clear
        .AddItem "ҩƷ����+ҩƷ����", 0
        .AddItem "ҩƷ����", 1
        .AddItem "ҩƷ����", 2
    End With
    
    Call LoadStore
        
    '��ȡ����
    Call LoadBatchSet
    Call LoadParams
    Call LoadPRI
    
    Call loadVolume
    Call LoadDept
    Call Load��Դ����
    
    Call chkAll_Click
    
    Call chkOpen_Click
End Sub
Private Sub LoadBatchSet()
    '��ȡ��ҩ���Ĺ�������
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select ����,��ɫ, ��ҩʱ��, ��ҩʱ��, ���, ����,ҩƷ���� From ��ҩ�������� where ��������ID=[1] Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ���Ĺ�������", Me.CboStore.ItemData(Me.CboStore.ListIndex))
    
    With vsfBatch
        .rows = 2
        .ColComboList(.ColIndex("ҩƷ����")) = "   |����ҩ|Ӫ��ҩ|������"
        Do While Not rsTmp.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsTmp!���� & "#"
            .TextMatrix(.rows - 1, .ColIndex("����ʱ�俪ʼ")) = Mid(rsTmp!��ҩʱ��, 1, InStr(rsTmp!��ҩʱ��, "-") - 1)
            .TextMatrix(.rows - 1, .ColIndex("����ʱ�����")) = Mid(rsTmp!��ҩʱ��, InStr(rsTmp!��ҩʱ��, "-") + 1)
            .TextMatrix(.rows - 1, .ColIndex("��ҩʱ�俪ʼ")) = Mid(rsTmp!��ҩʱ��, 1, InStr(rsTmp!��ҩʱ��, "-") - 1)
            .TextMatrix(.rows - 1, .ColIndex("��ҩʱ�����")) = Mid(rsTmp!��ҩʱ��, InStr(rsTmp!��ҩʱ��, "-") + 1)
            .TextMatrix(.rows - 1, .ColIndex("���")) = IIf(rsTmp!��� = 0, "", "��")
            .TextMatrix(.rows - 1, .ColIndex("����")) = IIf(rsTmp!���� = 0, IIf(rsTmp!���� = 0, "��", ""), "��")
            .TextMatrix(.rows - 1, .ColIndex("ҩƷ����")) = NVL(rsTmp!ҩƷ����)
            
            If .TextMatrix(.rows - 1, .ColIndex("����")) = "" Then
                .Cell(flexcpBackColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &HE0E0E0
            Else
                .Cell(flexcpBackColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &H80000005
            End If
            
            .Cell(flexcpBackColor, .rows - 1, .ColIndex("��ɫ"), .rows - 1, .ColIndex("��ɫ")) = IIf(rsTmp!���� = 0, &H80000005, rsTmp!��ɫ)
            rsTmp.MoveNext
        Loop
        
        vsfBatch.Enabled = IsHavePrivs(mstrPrivs, "���ù�������")
        If vsfBatch.Enabled = False Then
            Label2.Caption = Label2.Caption & "(��Ȩ�޽����޸�)"
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnEdit = False
End Sub

Private Sub lvwPRI_DblClick()
    Dim strIDS As String
    Dim strReturn As String
    
    strReturn = ReturnSelectedPri(0, strIDS)
    
    If mblnPri Then
        With Me.vsfPri
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("��������") Then
                .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
            End If
        End With
    Else
        With Me.vsfVolume
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("��������") Then
                .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
            End If
        End With
    End If
End Sub

Private Sub lvwPRI_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With lvwPRI
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        
        Item.Selected = True
        If Mid(Item.Text, 1, 2) = "����" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub lvwPRI_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strIDS As String
    Dim strReturn As String

    If KeyCode = vbKeyReturn Then
        strReturn = ReturnSelectedPri(1, strIDS)
        
        If mblnPri Then
            With Me.vsfPri
                .TextMatrix(mintRow, mintCol) = strReturn
                If mintCol = .ColIndex("��������") Then
                    .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
                End If
            End With
        Else
            With Me.vsfVolume
                .TextMatrix(mintRow, mintCol) = strReturn
                If mintCol = .ColIndex("��������") Then
                    .TextMatrix(mintRow, .ColIndex("����id")) = strIDS
                End If
            End With
        End If
    End If
End Sub






Private Function ReturnSelectedPri(ByVal intType As Integer, ByRef strIDS As String) As String
    'intType:0-˫���б�ʱ��1-�б��а��س�ʱ
    Dim n As Integer
    Dim strReturn As String
    
    With lvwPRI
        If .SelectedItem Is Nothing Then Exit Function
        
        strReturn = .SelectedItem.Text
        strIDS = Mid(.SelectedItem.Key, 2)
        
'        '���ѡ����ȫѡ������ȡ����ѡ����
'        If .ListItems(1).Checked Then
'            strReturn = .ListItems(1).Text
'            ReturnSelectedPri = strReturn
'            picPRI.Visible = False
'            Exit Function
'        End If
'
'        For n = 1 To .ListItems.Count
'            If .ListItems(n).Checked Then
'                strReturn = IIf(strReturn = "", .ListItems(n).Text, strReturn & "," & .ListItems(n).Text)
'                strIDS = IIf(strIDS = "", Mid(.ListItems(n).Key, 2), strIDS & "," & Mid(.ListItems(n).Key, 2))
'            End If
'        Next
'
'        If intType = 0 Then
'            '�����ǰ˫����ѡ��δ��ѡ�ϣ�����ǰ˫����ѡ��Ҳ���뵽�༭����
'            If .SelectedItem.Checked = False Then
'                .SelectedItem.Checked = True
'                strReturn = IIf(strReturn = "", .SelectedItem.Text, strReturn & "," & .SelectedItem.Text)
'                strIDS = IIf(strIDS = "", Mid(.ListItems(n).Key, 2), strIDS & "," & Mid(.ListItems(n).Key, 2))
'            End If
'
'            If .ListItems(1).Checked Then
'                strReturn = .ListItems(1).Text
'                ReturnSelectedPri = strReturn
'                Exit Function
'            End If
'        End If
        
        picPRI.Visible = False
        
        CmdOK.Enabled = True
        CmdCancel.Enabled = True
        ReturnSelectedPri = strReturn
        mblnEdit = True
    End With
End Function

Private Sub picPRI_Resize()
    On Error Resume Next
    
    With lvwPRI
        .Top = 0
        .Left = 0
        .Width = picPRI.Width
        .Height = picPRI.Height - 200 - cmdNO.Height
    End With
    
    With cmdNO
        .Top = picPRI.Height - .Height - 50
        .Left = picPRI.Width - .Width - 50
    End With
    
    With cmdYes
        .Top = cmdNO.Top
        .Left = cmdNO.Left - .Width - 100
    End With
End Sub



Private Sub sstMain_Click(PreviousTab As Integer)
    Dim i As Integer
    
    If PreviousTab = 5 Then
        Me.vsfVolume.Row = Me.vsfVolume.rows - 1
        Me.vsfVolume.Col = Me.vsfVolume.ColIndex("��������")
    End If
End Sub







Private Sub LoadPRI()

    Set mRsDept = DeptSendWork_Get��������
    
    Set mRsType = DeptSendWork_Get��ҩ����
    
    Set mRsPC = DeptSendWork_GetƵ��
    
End Sub


Private Sub updDeff_DownClick()
    If Me.txtDeff.Text <> "0" Then
        Me.txtDeff.Text = Me.txtDeff.Text - 1
    End If
End Sub

Private Sub updDeff_UpClick()
    If Me.txtDeff.Text <> "24" Then
        Me.txtDeff.Text = Me.txtDeff.Text + 1
    End If
End Sub

Private Sub vsfBatch_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBatch
        Select Case Col
            Case .ColIndex("����ʱ�俪ʼ"), .ColIndex("����ʱ�����"), .ColIndex("��ҩʱ�俪ʼ"), .ColIndex("��ҩʱ�����")
                If .TextMatrix(Row, Col) = "" Then Exit Sub
                
                If IsDate(.TextMatrix(Row, Col)) = False Then
                    MsgBox "��¼��ʱ���ʽ������12:59����9:20�ȡ�", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = ""
                    Exit Sub
                End If
                
                If Col = .ColIndex("����ʱ�俪ʼ") And .TextMatrix(Row, .ColIndex("����ʱ�����")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("����ʱ�俪ʼ"))) >= CDate(.TextMatrix(Row, .ColIndex("����ʱ�����"))) Then
                        MsgBox "��ʼʱ�����С�ڽ���ʱ�䣬���������á�", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("����ʱ�����") And .TextMatrix(Row, .ColIndex("����ʱ�俪ʼ")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("����ʱ�����"))) <= CDate(.TextMatrix(Row, .ColIndex("����ʱ�俪ʼ"))) Then
                        MsgBox "����ʱ�������ڿ�ʼʱ�䣬���������á�", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("��ҩʱ�俪ʼ") And .TextMatrix(Row, .ColIndex("��ҩʱ�����")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("��ҩʱ�俪ʼ"))) >= CDate(.TextMatrix(Row, .ColIndex("��ҩʱ�����"))) Then
                        MsgBox "��ʼʱ�����С�ڽ���ʱ�䣬���������á�", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("��ҩʱ�����") And .TextMatrix(Row, .ColIndex("��ҩʱ�俪ʼ")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("��ҩʱ�����"))) <= CDate(.TextMatrix(Row, .ColIndex("��ҩʱ�俪ʼ"))) Then
                        MsgBox "����ʱ�������ڿ�ʼʱ�䣬���������á�", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfBatch_DblClick()
    With vsfBatch
        If .Row < 2 Then Exit Sub
        If (.Col <> .ColIndex("���") And .Col <> .ColIndex("����")) And .Col <> .ColIndex("��ɫ") Then Exit Sub
        If (.MouseRow <> .Row Or .MouseCol <> .Col) And .Col <> .ColIndex("��ɫ") Then Exit Sub
        
        If .Col <> .ColIndex("��ɫ") Then
            If .TextMatrix(.Row, .Col) = "��" Then
                If .TextMatrix(.Row, .ColIndex("����")) = "0#" And .Col = .ColIndex("����") Then
                    MsgBox "0������Ϊ�������Σ��޷�����Ϊ�������á�״̬��"
                Else
                    .TextMatrix(.Row, .Col) = ""
                End If
            Else
                .TextMatrix(.Row, .Col) = "��"
            End If
            
            If .Col = .ColIndex("����") Then
                If .TextMatrix(.Row, .Col) = "" Then
                    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &HE0E0E0
                Else
                    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &H80000005
                End If
            End If
        
        Else
            On Error GoTo errHandle
            cmdialog.CancelError = True
            cmdialog.ShowColor
            .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = cmdialog.Color
            
errHandle:
        End If
    End With
End Sub


Private Sub vsfBatch_EnterCell()
    With vsfBatch
        If .Row < 2 Then Exit Sub
        .Editable = flexEDNone
        
        If .Col = .ColIndex("����ʱ�俪ʼ") Or .Col = .ColIndex("����ʱ�����") Or .Col = .ColIndex("��ҩʱ�俪ʼ") Or .Col = .ColIndex("��ҩʱ�����") Or .Col = .ColIndex("ҩƷ����") Then
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


Private Sub vsfBatch_KeyPress(KeyAscii As Integer)
    With vsfBatch
        If KeyAscii = 13 Then
            If .Col < .Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row < .rows - 1 Then
                    .Row = .Row + 1
                    .Col = .ColIndex("����ʱ�俪ʼ")
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfBatch_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfBatch
        Select Case Col
            Case .ColIndex("����ʱ�俪ʼ"), .ColIndex("����ʱ�����"), .ColIndex("��ҩʱ�俪ʼ"), .ColIndex("��ҩʱ�����")
                If InStr("1234567890:" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(":") Then
                    If InStr(.EditText, ":") <> 0 Then
                        KeyAscii = 0
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfDept_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row <> 1 Then Cancel = True
End Sub

Private Sub vsfDept_EnterCell()
    If mblnEdit Then
        If MsgBox("�뱣�����õ����ȼ����л����Һ����������ȼ����ý�ʧЧ���Ƿ��л���", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))))
            mblnEdit = False
            
        End If
    Else
        If Me.vsfDept.Row > 1 Then
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("����id"))))
        End If
    End If
    
End Sub

Private Sub vsfDept_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
    With Me.vsfDept
        If KeyAscii <> 13 Or .TextMatrix(1, .ColIndex("��������")) = "" Or .Row <> 1 Then Exit Sub
        
        For intRow = 2 To .rows - 1
            If .TextMatrix(intRow, .ColIndex("����")) = UCase(.TextMatrix(1, .ColIndex("��������"))) Then
                .Row = intRow
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub vsfNoMedi_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strKey As String
    Dim StrCode As String
    
    If KeyCode = 13 Then
        vRect = zlControl.GetControlRect(vsfNoMedi.hWnd)
        dblLeft = vRect.Left + vsfNoMedi.CellLeft
        dblTop = vRect.Top + vsfNoMedi.CellTop + vsfNoMedi.CellHeight + 3200
        
        With vsfNoMedi
            If Col = .ColIndex("ҩƷ���������") Then
                strKey = Trim(.EditText)
                If strKey = "" Then Exit Sub
                
                If IsNumeric(strKey) Then
                    '������
                    StrCode = " d.���� like [1] "
                ElseIf zlCommFun.IsCharAlpha(strKey) Then
                    '����ĸ
                    StrCode = " n.���� Like [1] "
                ElseIf zlCommFun.IsCharChinese(strKey) Then
                    '������
                    StrCode = " d.���� like [1] "
                Else
                    StrCode = " (n.���� Like [1] Or d.���� Like [1] Or n.���� Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'��' || d.���� || '��' || d.���� || '(' || d.��� || ')' As ͨ����" & vbNewLine & _
                    " From ҩƷ��� T, �շ���ĿĿ¼ D, �շ���Ŀ���� N" & vbNewLine & _
                    " Where t.ҩƷid = d.Id And t.ҩƷid = n.�շ�ϸĿid And D.��� In ('5', '6') And" & StrCode & vbNewLine & _
                    " And (d.����ʱ�� Is Null Or To_Char(d.����ʱ��, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '��' || d.���� || '��' || d.���� || '(' || d.��� || ')'"
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ���������", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, IIf(gstrMatchMethod = 0, "", "%") & UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .rows - 1
                        If rsRecord!Id = Val(.TextMatrix(i, .ColIndex("ҩƷID"))) Then
                            MsgBox rsRecord!ͨ���� & "�Ѿ�¼�룬������ѡ��", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("ҩƷID")) = rsRecord!Id
                    .TextMatrix(.Row, .ColIndex("ҩƷ���������")) = rsRecord!ͨ����
                    .EditText = rsRecord!ͨ����
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfNoMedi_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfPRI_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    mblnPri = True
    mintRow = Row
    mintCol = Col
    With Me.picPRI
        .Visible = True
    
        .Height = vsfPri.Height
        .Top = sstMain.Top + vsfPri.Top
        .Left = sstMain.Left + vsfPri.Left
        .Width = vsfPri.Width
    End With
            
    Select Case Col
        Case vsfPri.ColIndex("��������")
            With Me.lvwPRI
                .ListItems.Clear
                .ListItems.Add , "_" & 0, "���п���", 1, 1
                mRsDept.MoveFirst
                Do While Not mRsDept.EOF
                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!����, 1, 1
                    mRsDept.MoveNext
                Loop
                .ListItems.Add , "_00", "��������", 1, 1
            End With
        Case vsfPri.ColIndex("��ҩ����")
            With Me.lvwPRI
                .ListItems.Clear
                If mRsType.RecordCount > 0 Then mRsType.MoveFirst
                Do While Not mRsType.EOF
                    .ListItems.Add , "_" & mRsType!����, mRsType!����, 1, 1
                    mRsType.MoveNext
                Loop
                 .ListItems.Add , "_00", "��������", 1, 1
            End With
        Case vsfPri.ColIndex("Ƶ��")
            With Me.lvwPRI
                .ListItems.Clear
                .ListItems.Add , "_" & 0, "����Ƶ��", 1, 1
                mRsPC.MoveFirst
                Do While Not mRsPC.EOF
                    .ListItems.Add , "_" & mRsPC!����, mRsPC!���� & "(" & mRsPC!Ӣ������ & ")", 1, 1
                    mRsPC.MoveNext
                Loop
                .ListItems.Add , "_00", "����Ƶ��", 1, 1
            End With
    End Select
End Sub

Private Sub vsfPrint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If Me.vsfPrint.rows = 2 Then
            Me.vsfPrint.TextMatrix(vsfPrint.Row, vsfPrint.ColIndex("ҩƷid")) = ""
            Me.vsfPrint.TextMatrix(vsfPrint.Row, vsfPrint.ColIndex("ҩƷ���������")) = ""
        Else
            Me.vsfPrint.RemoveItem vsfPrint.Row
        End If
        
    End If
End Sub

Private Sub vsfPrint_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strKey As String
    Dim StrCode As String
    
    If KeyCode = 13 Then
        vRect = zlControl.GetControlRect(vsfPrint.hWnd)
        dblLeft = vRect.Left + vsfPrint.CellLeft
        dblTop = vRect.Top + vsfPrint.CellTop + vsfPrint.CellHeight + 3200
        
        With vsfPrint
            If Col = .ColIndex("ҩƷ���������") Then
                strKey = Trim(.EditText)
                If strKey = "" Then Exit Sub
                
                If IsNumeric(strKey) Then
                    '������
                    StrCode = " d.���� like [1] "
                ElseIf zlCommFun.IsCharAlpha(strKey) Then
                    '����ĸ
                    StrCode = " n.���� Like [1] "
                ElseIf zlCommFun.IsCharChinese(strKey) Then
                    '������
                    StrCode = " d.���� like [1] "
                Else
                    StrCode = " (n.���� Like [1] Or d.���� Like [1] Or n.���� Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'��' || d.���� || '��' || d.���� || '(' || d.��� || ')' As ͨ����" & vbNewLine & _
                    " From ҩƷ��� T, �շ���ĿĿ¼ D, �շ���Ŀ���� N" & vbNewLine & _
                    " Where t.ҩƷid = d.Id And t.ҩƷid = n.�շ�ϸĿid And D.��� In ('5', '6') And" & StrCode & vbNewLine & _
                    " And (d.����ʱ�� Is Null Or To_Char(d.����ʱ��, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '��' || d.���� || '��' || d.���� || '(' || d.��� || ')'"
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ���������", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, IIf(gstrMatchMethod = 0, "", "%") & UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .rows - 1
                        If rsRecord!Id = Val(.TextMatrix(i, .ColIndex("ҩƷID"))) Then
                            MsgBox rsRecord!ͨ���� & "�Ѿ�¼�룬������ѡ��", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("ҩƷID")) = rsRecord!Id
                    .TextMatrix(.Row, .ColIndex("ҩƷ���������")) = rsRecord!ͨ����
                    .EditText = rsRecord!ͨ����
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfPrint_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub vsfVolume_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = Me.vsfVolume.ColIndex("����") Then
        If Not IsNumeric(vsfVolume.TextMatrix(Row, Col)) Then
            MsgBox "������¼�����֣�", vbInformation + vbOKOnly, gstrSysName
            vsfVolume.Col = vsfVolume.ColIndex("����")
        End If
    End If
End Sub

Private Sub vsfVolume_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str���� As String
    Dim i As Integer
    
    If Col <> vsfVolume.ColIndex("��ҩ����") Then Exit Sub
    With Me.vsfBatch
        If .rows > 2 Then
            For i = 2 To .rows - 1
                If .TextMatrix(i, .ColIndex("����")) <> "" And .TextMatrix(i, .ColIndex("����")) <> "" Then
                    str���� = IIf(str���� = "", "", str���� & "|") & .TextMatrix(i, .ColIndex("����"))
                End If
            Next
        End If
        If str���� <> "" Then Me.vsfVolume.ColComboList(vsfVolume.ColIndex("��ҩ����")) = str����
    End With
End Sub

Private Sub vsfVolume_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    mblnPri = False
'    mintRow = Row
'
'    mintCol = Col
''    With Me.picPRI
'        .Visible = True
'        .Height = vsfVolume.Height
'        .Top = sstMain.Top + vsfPri.Top
'        .Left = sstMain.Left + vsfVolume.Left
'        .Width = vsfVolume.Width
'    End With
'
'    With vsfVolume
'        If Col = .ColIndex("��������") Then
'            With Me.lvwPRI
'                .ListItems.Clear
'                .ListItems.Add , "_" & 0, "���п���", 1, 1
'                mRsDept.MoveFirst
'                Do While Not mRsDept.EOF
'                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!����, 1, 1
'                    mRsDept.MoveNext
'                Loop
'                .ListItems.Add , "_00", "��������", 1, 1
'            End With
'        End If
'    End With

    mblnPri = False
    mintRow = vsfVolume.Row
    mintCol = vsfVolume.Col

    With Me.lvwPRI
        .ListItems.Clear
        .ListItems.Add , "_" & 0, "���п���", 1, 1
        mRsDept.MoveFirst
        Do While Not mRsDept.EOF
            If vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) <> "" Then
                If mRsDept!���� = UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) Or mRsDept!��ʼ��� = UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) Or vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) = mRsDept!���� Then
                    vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) = mRsDept!����
                    vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.ColIndex("����id")) = mRsDept!Id
                    Exit Sub

                ElseIf InStr(1, mRsDept!��ʼ���, UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col))) > 0 Or InStr(1, mRsDept!����, UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col))) > 0 Or InStr(1, mRsDept!����, vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) > 0 Then
                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!����, 1, 1
                End If
            Else
                .ListItems.Add , "_" & mRsDept!Id, mRsDept!����, 1, 1
            End If
            mRsDept.MoveNext
        Loop
        
        If .ListItems.count = 1 Then
            .ListItems.Clear
            MsgBox "������ļ���û����֮ƥ��Ŀ��ң�������¼�룡"
            Exit Sub
        End If
        
        .ListItems.Add , "_00", "��������", 1, 1
    End With
    

    With Me.picPRI
        .Visible = True
        .Height = vsfVolume.Height
        .Top = sstMain.Top + vsfPri.Top
        .Left = sstMain.Left + vsfVolume.Left
        .Width = vsfVolume.Width
    End With
End Sub

Private Sub loadVolume()
    Dim rstemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select ����id,��������,����,��ҩ���� from ������������ where ��������ID=[1] order by ����id"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������������", Me.CboStore.ItemData(Me.CboStore.ListIndex))
    
    i = 1
    With Me.vsfVolume
        .RowHeight(0) = 250
        .rows = 1
        .rows = IIf(rstemp.RecordCount = 0, 1, rstemp.RecordCount) + 1
        Do While Not rstemp.EOF
            .RowHeight(i) = 250
            .TextMatrix(i, .ColIndex("����id")) = rstemp!����ID
            .TextMatrix(i, .ColIndex("��������")) = rstemp!��������
            .TextMatrix(i, .ColIndex("��ҩ����")) = zlStr.NVL(rstemp!��ҩ����)
            .TextMatrix(i, .ColIndex("����")) = rstemp!����
            i = i + 1
            rstemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfVolume_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> 13 Then Exit Sub
'
'    With Me.vsfVolume
'        If .Row = .rows - 1 Then
'            If .Col = .Cols - 1 Then
'                Exit Sub
'            Else
'                .Col = .Col + 1
'            End If
'        Else
'            If .Col = .Cols - 1 Then
'                .Row = .Row + 1
'                .Col = .ColIndex("��������")
'            Else
'                .Col = .Col + 1
'            End If
'        End If
'    End With
End Sub

Private Sub vsfVolume_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    With Me.vsfVolume
        If .Row = .rows - 1 Then
            If .Col = .Cols - 1 Then
                Exit Sub
            Else
                .Col = .Col + 1
            End If
        Else
            If .Col = .Cols - 1 Then
                .Row = .Row + 1
                .Col = .ColIndex("��������")
            Else
                .Col = .Col + 1
            End If
        End If
    End With
End Sub

Private Sub vsfVolume_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfVolume
        If Col = .ColIndex("����") Then
            If InStr("1234567890-." & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End With

End Sub

Private Sub LoadDept()
    Dim i As Integer
    
    i = 1
    vsfDept.rows = mRsDept.RecordCount + 2
    Do While Not mRsDept.EOF
        With Me.vsfDept
            i = i + 1
            .TextMatrix(i, .ColIndex("���")) = i - 1
            .TextMatrix(i, .ColIndex("����id")) = mRsDept!Id
            .TextMatrix(i, .ColIndex("��������")) = mRsDept!����
            .TextMatrix(i, .ColIndex("����")) = mRsDept!����
        End With
        mRsDept.MoveNext
    Loop
End Sub

Private Sub vsfPrint_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfPrint.ColIndex("ҩƷ���������") Then Cancel = True
End Sub

Private Sub vsfNoMedi_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfNoMedi.ColIndex("ҩƷ���������") Then Cancel = True
End Sub

Private Sub vsfNoMedi_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = 46 Then
        If Me.vsfNoMedi.rows = 2 Then
            Me.vsfNoMedi.TextMatrix(vsfNoMedi.Row, vsfNoMedi.ColIndex("ҩƷid")) = ""
            Me.vsfNoMedi.TextMatrix(vsfNoMedi.Row, vsfNoMedi.ColIndex("ҩƷ���������")) = ""
        Else
            Me.vsfNoMedi.RemoveItem vsfNoMedi.Row
        End If
    End If
End Sub
