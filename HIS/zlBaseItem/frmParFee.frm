VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmParFee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ò�������"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   Icon            =   "frmParFee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8602.637
   ScaleMode       =   0  'User
   ScaleWidth      =   11910
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   587
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11910
      TabIndex        =   119
      Top             =   7980
      Width           =   11910
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   130
         Top             =   145
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   124
         Top             =   145
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   122
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   11400
         TabIndex        =   121
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   10245
         TabIndex        =   120
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   131
         Top             =   165
         Width           =   4215
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   129
         Top             =   168
         Width           =   1095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&S)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   123
         Top             =   168
         Width           =   1095
      End
   End
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   7980
      Left            =   0
      ScaleHeight     =   7980
      ScaleWidth      =   2415
      TabIndex        =   117
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   6060
         Left            =   2260
         MousePointer    =   9  'Size W E
         ScaleHeight     =   6060
         ScaleWidth      =   45
         TabIndex        =   118
         Top             =   120
         Width           =   45
      End
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   30
         ScaleHeight     =   6375
         ScaleWidth      =   2370
         TabIndex        =   125
         Top             =   0
         Width           =   2370
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5490
            Left            =   0
            TabIndex        =   126
            Top             =   720
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   9684
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager imgFunc 
            Left            =   1800
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParFee.frx":6852
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   127
            Top             =   0
            Width           =   2200
            _Version        =   589884
            _ExtentX        =   3881
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   7005
         Left            =   0
         TabIndex        =   128
         Top             =   0
         Width           =   2400
         _Version        =   589884
         _ExtentX        =   4233
         _ExtentY        =   12356
         _StockProps     =   64
      End
      Begin XtremeCommandBars.ImageManager imgType 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmParFee.frx":10A52
      End
   End
   Begin TabDlg.SSTab stabDesign 
      Height          =   8205
      Left            =   2340
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   -15
      Visible         =   0   'False
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   14473
      _Version        =   393216
      Style           =   1
      Tabs            =   18
      Tab             =   11
      TabsPerRow      =   12
      TabHeight       =   520
      TabCaption(0)   =   "�Һ�"
      TabPicture(0)   =   "frmParFee.frx":1A13E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picPar(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ԥ��"
      TabPicture(1)   =   "frmParFee.frx":1A15A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPar(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "һ��ͨ"
      TabPicture(2)   =   "frmParFee.frx":1A176
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPar(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "ҽ�ƿ�"
      TabPicture(3)   =   "frmParFee.frx":1A192
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picPar(8)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "����"
      TabPicture(4)   =   "frmParFee.frx":1A1AE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picPar(9)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "���ﻮ��"
      TabPicture(5)   =   "frmParFee.frx":1A1CA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picPar(10)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "�����շ�"
      TabPicture(6)   =   "frmParFee.frx":1A1E6
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "picPar(11)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "�������"
      TabPicture(7)   =   "frmParFee.frx":1A202
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "picPar(12)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "������"
      TabPicture(8)   =   "frmParFee.frx":1A21E
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "picPar(13)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "סԺ����"
      TabPicture(9)   =   "frmParFee.frx":1A23A
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "picPar(14)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "���˽���"
      TabPicture(10)  =   "frmParFee.frx":1A256
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "picPar(15)"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "����"
      TabPicture(11)  =   "frmParFee.frx":1A272
      Tab(11).ControlEnabled=   -1  'True
      Tab(11).Control(0)=   "picPar(0)"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "������"
      TabPicture(12)  =   "frmParFee.frx":1A28E
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "picPar(16)"
      Tab(12).Control(0).Enabled=   0   'False
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "ҽ������"
      TabPicture(13)  =   "frmParFee.frx":1A2AA
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "picPar(17)"
      Tab(13).Control(0).Enabled=   0   'False
      Tab(13).ControlCount=   1
      TabCaption(14)  =   "���û���"
      TabPicture(14)  =   "frmParFee.frx":1A2C6
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "picPar(5)"
      Tab(14).ControlCount=   1
      TabCaption(15)  =   "����Ա����"
      TabPicture(15)  =   "frmParFee.frx":1A2E2
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "picPar(2)"
      Tab(15).ControlCount=   1
      TabCaption(16)  =   "���ʱ���"
      TabPicture(16)  =   "frmParFee.frx":1A2FE
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "picPar(3)"
      Tab(16).ControlCount=   1
      TabCaption(17)  =   "�Զ�����"
      TabPicture(17)  =   "frmParFee.frx":1A31A
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "picPar(4)"
      Tab(17).Control(0).Enabled=   0   'False
      Tab(17).ControlCount=   1
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7185
         Index           =   4
         Left            =   -75240
         ScaleHeight     =   7155
         ScaleWidth      =   9585
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   630
         Width           =   9615
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   15
            ItemData        =   "frmParFee.frx":1A336
            Left            =   1305
            List            =   "frmParFee.frx":1A338
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   6315
            Width           =   2880
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "�Լ۸���ߵĻ���ȼ�Ϊ��׼"
            Height          =   255
            Index           =   1
            Left            =   5520
            TabIndex        =   10
            Top             =   5940
            Width           =   2670
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "�����һ�λ���ȼ�Ϊ��׼"
            Height          =   255
            Index           =   0
            Left            =   2775
            TabIndex        =   9
            Top             =   5970
            Value           =   -1  'True
            Width           =   2625
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������Զ��Ʒ�(��ʾ�Ƿ��Զ��޸���һ�����ڼ���Զ����ü������ݡ�)"
            Height          =   285
            Index           =   12
            Left            =   120
            TabIndex        =   7
            Top             =   5715
            Width           =   6510
         End
         Begin VB.TextBox txtDateInput 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   270
            Left            =   1560
            TabIndex        =   6
            Top             =   1560
            Visible         =   0   'False
            Width           =   1380
         End
         Begin ZL9BillEdit.BillEdit Bill 
            Height          =   5235
            Index           =   0
            Left            =   4965
            TabIndex        =   5
            Top             =   420
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   9234
            CellAlignment   =   9
            Text            =   ""
            TextMatrix0     =   ""
            MaxDate         =   2958465
            MinDate         =   -53688
            Value           =   36395
            Cols            =   3
            RowHeight0      =   315
            RowHeightMin    =   315
            ColWidth0       =   1005
            BackColor       =   -2147483643
            BackColorBkg    =   -2147483643
            BackColorSel    =   10249818
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            ForeColorSel    =   -2147483634
            GridColor       =   -2147483630
            ColAlignment0   =   9
            ListIndex       =   -1
            CellBackColor   =   -2147483643
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshAutoCalc 
            Height          =   5235
            Left            =   90
            TabIndex        =   3
            Top             =   420
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   9234
            _Version        =   393216
            Cols            =   3
            RowHeightMin    =   315
            BackColorBkg    =   -2147483643
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin VB.Frame fraAutoCharge 
            Caption         =   "�Զ�����ģʽ"
            Height          =   975
            Left            =   90
            TabIndex        =   756
            Top             =   6375
            Width           =   9360
            Begin VB.CheckBox chk 
               Caption         =   "���������ģʽ (ָ�԰���Ϊ���㵥λ,������Ժ��1��,���������,�����Ժ���첻�����,���������)"
               Height          =   225
               Index           =   43
               Left            =   135
               TabIndex        =   12
               Top             =   405
               Width           =   8775
            End
            Begin VB.Label lblAutoChargeNM 
               AutoSize        =   -1  'True
               Caption         =   "�Զ����ʹ���˵����"
               Height          =   180
               Left            =   120
               TabIndex        =   757
               Top             =   330
               Width           =   1620
            End
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            Caption         =   "ͬ�첻ͬ����ȼ��Ļ���Ѽ���"
            Height          =   180
            Left            =   90
            TabIndex        =   8
            Top             =   6000
            Width           =   2520
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������ָ�����ý����Զ�����"
            Height          =   180
            Index           =   13
            Left            =   5070
            TabIndex        =   4
            Top             =   120
            Width           =   2520
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Դ�λ�ѻ���ѽ����Զ�����"
            Height          =   180
            Index           =   12
            Left            =   165
            TabIndex        =   2
            Top             =   135
            Width           =   2520
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7095
         Index           =   2
         Left            =   -74805
         ScaleHeight     =   7065
         ScaleWidth      =   8970
         TabIndex        =   13
         Top             =   705
         Visible         =   0   'False
         Width           =   9000
         Begin VB.CommandButton cmdOperate 
            Caption         =   "����(&A)"
            CausesValidation=   0   'False
            Height          =   350
            Index           =   0
            Left            =   8310
            TabIndex        =   16
            Top             =   405
            Width           =   1100
         End
         Begin VB.CommandButton cmdOperate 
            Caption         =   "�޸�(&M)"
            CausesValidation=   0   'False
            Height          =   350
            Index           =   1
            Left            =   8310
            TabIndex        =   17
            Top             =   885
            Width           =   1100
         End
         Begin VB.CommandButton cmdOperate 
            Caption         =   "ɾ��(&D)"
            CausesValidation=   0   'False
            Height          =   350
            Index           =   2
            Left            =   8310
            TabIndex        =   18
            Top             =   1365
            Width           =   1100
         End
         Begin VB.CommandButton cmdOperate 
            Caption         =   "���(&L)"
            CausesValidation=   0   'False
            Height          =   350
            Index           =   3
            Left            =   8310
            TabIndex        =   19
            Top             =   1845
            Width           =   1100
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2070
            MaxLength       =   12
            TabIndex        =   24
            Top             =   6705
            Width           =   1350
         End
         Begin MSComctlLib.ListView lvw 
            Height          =   6255
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   8100
            _ExtentX        =   14288
            _ExtentY        =   11033
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "������"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "��������"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "��ʷ����"
               Object.Width           =   2187
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "����������˵���"
               Object.Width           =   2893
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "�������"
               Object.Width           =   2187
            EndProperty
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ա�Բ�ͬ���ݵĲ���Ȩ�ޣ���Ե��ݵ���ʷ��������������˽�������"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   6120
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ʷ���������ѽ�"
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   6765
            Width           =   1980
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7335
         Index           =   5
         Left            =   -74850
         ScaleHeight     =   7305
         ScaleWidth      =   9060
         TabIndex        =   32
         Top             =   690
         Width           =   9090
         Begin VB.CommandButton cmd���ѷ������� 
            Caption         =   "ȫ��"
            Height          =   300
            Index           =   1
            Left            =   5880
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   3360
            Width           =   900
         End
         Begin VB.CommandButton cmd���ѷ������� 
            Caption         =   "ȫѡ"
            Height          =   300
            Index           =   0
            Left            =   4950
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   3360
            Width           =   900
         End
         Begin VB.CommandButton cmdҽ���������� 
            Caption         =   "ȫ��"
            Height          =   300
            Index           =   1
            Left            =   8280
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   3360
            Width           =   900
         End
         Begin VB.CommandButton cmdҽ���������� 
            Caption         =   "ȫѡ"
            Height          =   300
            Index           =   0
            Left            =   7350
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   3360
            Width           =   900
         End
         Begin VB.ListBox lst 
            Height          =   2790
            Index           =   1
            Left            =   4950
            Style           =   1  'Checkbox
            TabIndex        =   49
            Top             =   510
            Width           =   2220
         End
         Begin VB.Frame fra�ض��շ���Ŀ 
            Caption         =   " �ض��շ���Ŀ "
            Height          =   2265
            Left            =   360
            TabIndex        =   33
            Top             =   240
            Width           =   3495
            Begin VB.CommandButton cmdSelect 
               Caption         =   "��"
               Height          =   240
               Index           =   1
               Left            =   2775
               TabIndex        =   643
               TabStop         =   0   'False
               Top             =   780
               Width           =   255
            End
            Begin VB.CommandButton cmdSelect 
               Caption         =   "��"
               Height          =   240
               Index           =   0
               Left            =   2775
               TabIndex        =   642
               TabStop         =   0   'False
               Top             =   315
               Width           =   255
            End
            Begin VB.CommandButton cmdSelect 
               Caption         =   "��"
               Height          =   240
               Index           =   3
               Left            =   2775
               TabIndex        =   641
               TabStop         =   0   'False
               Top             =   1245
               Width           =   255
            End
            Begin VB.CommandButton cmdSelect 
               Caption         =   "��"
               Height          =   240
               Index           =   4
               Left            =   2775
               TabIndex        =   640
               TabStop         =   0   'False
               Top             =   1710
               Width           =   255
            End
            Begin VB.TextBox txtCmd 
               Height          =   300
               Index           =   0
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   285
               Width           =   1710
            End
            Begin VB.TextBox txtCmd 
               Height          =   300
               Index           =   1
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   750
               Width           =   1710
            End
            Begin VB.TextBox txtCmd 
               Height          =   300
               Index           =   3
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   1215
               Width           =   1710
            End
            Begin VB.TextBox txtCmd 
               Height          =   300
               Index           =   4
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   1680
               Width           =   1710
            End
            Begin VB.Label lbl 
               Caption         =   "������"
               Height          =   225
               Index           =   7
               Left            =   630
               TabIndex        =   36
               Top             =   795
               Width           =   585
            End
            Begin VB.Label lbl 
               Caption         =   "������"
               Height          =   225
               Index           =   6
               Left            =   630
               TabIndex        =   34
               Top             =   315
               Width           =   585
            End
            Begin VB.Label lbl 
               Caption         =   "��ͨ���÷�"
               Height          =   225
               Index           =   18
               Left            =   285
               TabIndex        =   38
               Top             =   1245
               Width           =   930
            End
            Begin VB.Label lbl 
               Caption         =   "�������÷�"
               Height          =   225
               Index           =   56
               Left            =   285
               TabIndex        =   40
               Top             =   1725
               Width           =   930
            End
         End
         Begin VB.Frame fraƱ�� 
            Caption         =   "Ʊ�ݺ������"
            Height          =   3855
            Left            =   360
            TabIndex        =   42
            Top             =   2880
            Width           =   3495
            Begin VB.TextBox txtUD 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Index           =   4
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   44
               Text            =   "7"
               Top             =   480
               Width           =   300
            End
            Begin VB.CheckBox chk 
               Caption         =   "�ϸ����"
               Height          =   285
               Index           =   13
               Left            =   1860
               TabIndex        =   46
               Top             =   495
               Width           =   1020
            End
            Begin MSComCtl2.UpDown ud 
               Height          =   300
               Index           =   4
               Left            =   1440
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   480
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   7
               BuddyControl    =   "txtUD(4)"
               BuddyDispid     =   196647
               BuddyIndex      =   4
               OrigLeft        =   1350
               OrigTop         =   240
               OrigRight       =   1590
               OrigBottom      =   540
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComctlLib.ListView lvw 
               Height          =   2805
               Index           =   0
               Left            =   240
               TabIndex        =   47
               Top             =   855
               Width           =   2985
               _ExtentX        =   5265
               _ExtentY        =   4948
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Ʊ������"
                  Object.Width           =   1765
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "���볤��"
                  Object.Width           =   1588
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "�ϸ����"
                  Object.Width           =   1588
               EndProperty
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "���볤��"
               Height          =   180
               Index           =   19
               Left            =   240
               TabIndex        =   43
               Top             =   540
               Width           =   720
            End
         End
         Begin VB.ListBox lst 
            Height          =   2160
            Index           =   3
            Left            =   4920
            Style           =   1  'Checkbox
            TabIndex        =   57
            Top             =   4440
            Width           =   1665
         End
         Begin VB.ListBox lst 
            Height          =   2790
            Index           =   0
            Left            =   7245
            Style           =   1  'Checkbox
            TabIndex        =   53
            Top             =   495
            Width           =   2220
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ҽ���������÷�������"
            Height          =   180
            Index           =   20
            Left            =   7350
            TabIndex        =   52
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���Ѳ������÷�������"
            Height          =   180
            Index           =   21
            Left            =   4950
            TabIndex        =   48
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ˢ��Ҫ��������"
            Height          =   180
            Index           =   41
            Left            =   4920
            TabIndex        =   56
            Top             =   4080
            Width           =   1260
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7035
         Index           =   17
         Left            =   -74895
         ScaleHeight     =   7005
         ScaleWidth      =   9015
         TabIndex        =   626
         TabStop         =   0   'False
         Top             =   705
         Visible         =   0   'False
         Width           =   9045
         Begin VB.Frame fra 
            Caption         =   "���ʺ�ҩ��ʽ����"
            Height          =   870
            Index           =   10
            Left            =   135
            TabIndex        =   636
            Top             =   3315
            Width           =   4530
            Begin VB.OptionButton optSendDrugFF 
               Caption         =   "ѡ���Ƿ�ҩ"
               Height          =   180
               Index           =   2
               Left            =   2895
               TabIndex        =   639
               Top             =   450
               Width           =   1470
            End
            Begin VB.OptionButton optSendDrugFF 
               Caption         =   "����ҩ"
               Height          =   180
               Index           =   0
               Left            =   210
               TabIndex        =   637
               Top             =   450
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton optSendDrugFF 
               Caption         =   "�Զ���ҩ"
               Height          =   180
               Index           =   1
               Left            =   1515
               TabIndex        =   638
               Top             =   450
               Width           =   1470
            End
         End
         Begin VB.Frame fra 
            Caption         =   "ҩƷ��ʾ��λ"
            Height          =   810
            Index           =   9
            Left            =   135
            TabIndex        =   633
            Top             =   2325
            Width           =   4530
            Begin VB.OptionButton optDrugUnitFF 
               Caption         =   "�ۼ۵�λ"
               Height          =   180
               Index           =   0
               Left            =   195
               TabIndex        =   634
               Top             =   405
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton optDrugUnitFF 
               Caption         =   "����/סԺ��λ"
               Height          =   180
               Index           =   1
               Left            =   1515
               TabIndex        =   635
               Top             =   405
               Width           =   1470
            End
         End
         Begin VB.Frame fra 
            Height          =   2010
            Index           =   8
            Left            =   135
            TabIndex        =   627
            Top             =   165
            Width           =   4530
            Begin VB.CheckBox chk 
               Caption         =   "��������ֻ�������ڲ�����"
               Height          =   195
               Index           =   187
               Left            =   195
               TabIndex        =   632
               Top             =   1545
               Width           =   2865
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʾ����ҩ�����"
               Height          =   195
               Index           =   166
               Left            =   195
               TabIndex        =   630
               Top             =   930
               Width           =   1770
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʾ����ҩ����"
               Height          =   195
               Index           =   167
               Left            =   195
               TabIndex        =   631
               Top             =   1230
               Width           =   1770
            End
            Begin VB.CheckBox chk 
               Caption         =   "���������������"
               Height          =   195
               Index           =   165
               Left            =   195
               TabIndex        =   628
               Top             =   345
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ҩ�������븶��"
               Height          =   195
               Index           =   164
               Left            =   210
               TabIndex        =   629
               Top             =   630
               Value           =   1  'Checked
               Width           =   1740
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6705
         Index           =   16
         Left            =   -74865
         ScaleHeight     =   6675
         ScaleWidth      =   9165
         TabIndex        =   608
         TabStop         =   0   'False
         Top             =   765
         Visible         =   0   'False
         Width           =   9195
         Begin VB.CheckBox chk 
            Caption         =   "Ԥ�����ʰ������סԺ�ֱ�����"
            Height          =   375
            Index           =   198
            Left            =   435
            TabIndex        =   786
            Top             =   5355
            Width           =   3060
         End
         Begin VB.Frame fraSplit 
            Height          =   75
            Index           =   2
            Left            =   1410
            TabIndex        =   784
            Top             =   5070
            Width           =   3810
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������ӡ��"
            Height          =   375
            Index           =   162
            Left            =   435
            TabIndex        =   625
            Top             =   4500
            Width           =   3060
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������ӡ��"
            Height          =   375
            Index           =   161
            Left            =   435
            TabIndex        =   624
            Top             =   4200
            Width           =   2265
         End
         Begin VB.Frame fraSplit 
            Height          =   75
            Index           =   1
            Left            =   1170
            TabIndex        =   623
            Top             =   3945
            Width           =   4065
         End
         Begin VB.Frame fraSplit 
            Height          =   75
            Index           =   0
            Left            =   1185
            TabIndex        =   620
            Top             =   2970
            Width           =   4065
         End
         Begin VB.Frame fraSplit 
            Height          =   75
            Index           =   7
            Left            =   1185
            TabIndex        =   610
            Top             =   435
            Width           =   4065
         End
         Begin VB.Frame fraPrintModeDraw 
            Caption         =   "���ý����õ���ӡ��ʽ"
            Height          =   810
            Left            =   435
            TabIndex        =   615
            Top             =   1845
            Width           =   4620
            Begin VB.OptionButton optPrintModeDraw 
               Caption         =   "����ӡ(&1)"
               Height          =   300
               Index           =   0
               Left            =   180
               TabIndex        =   616
               Top             =   315
               Width           =   1230
            End
            Begin VB.OptionButton optPrintModeDraw 
               Caption         =   "�Զ���ӡ"
               Height          =   300
               Index           =   1
               Left            =   1455
               TabIndex        =   617
               Top             =   315
               Value           =   -1  'True
               Width           =   1305
            End
            Begin VB.OptionButton optPrintModeDraw 
               Caption         =   "ѡ���Ƿ��ӡ"
               Height          =   300
               Index           =   2
               Left            =   2880
               TabIndex        =   618
               Top             =   315
               Width           =   1650
            End
         End
         Begin VB.Frame fraPrintModeSJ 
            Caption         =   "�տ��վݴ�ӡ��ʽ"
            Height          =   840
            Left            =   435
            TabIndex        =   611
            Top             =   765
            Width           =   4620
            Begin VB.OptionButton optPrintModeSJ 
               Caption         =   "ѡ���Ƿ��ӡ"
               Height          =   180
               Index           =   2
               Left            =   2880
               TabIndex        =   614
               Top             =   420
               Width           =   1710
            End
            Begin VB.OptionButton optPrintModeSJ 
               Caption         =   "����ӡ"
               Height          =   180
               Index           =   0
               Left            =   180
               TabIndex        =   612
               Top             =   420
               Width           =   1185
            End
            Begin VB.OptionButton optPrintModeSJ 
               Caption         =   "�Զ���ӡ"
               Height          =   180
               Index           =   1
               Left            =   1455
               TabIndex        =   613
               Top             =   420
               Value           =   -1  'True
               Width           =   1365
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "����Ʊ��ʱ,�������ǩ��ȷ��."
            Height          =   180
            Index           =   160
            Left            =   435
            TabIndex        =   621
            Top             =   3315
            Width           =   3720
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            Caption         =   "�շ�Ա���˹���"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   785
            Top             =   5010
            Width           =   1260
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            Caption         =   "��Ա������"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   622
            Top             =   3885
            Width           =   1080
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            Caption         =   "Ʊ��ʹ�ü��"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   619
            Top             =   2910
            Width           =   1080
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            Caption         =   "�շѲ�����"
            Height          =   180
            Index           =   3
            Left            =   135
            TabIndex        =   609
            Top             =   375
            Width           =   1080
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7785
         Index           =   0
         Left            =   -180
         ScaleHeight     =   7755
         ScaleWidth      =   9570
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   690
         Visible         =   0   'False
         Width           =   9600
         Begin VB.CheckBox chk 
            Caption         =   "ָ�����ϲ���ʱ����ʾ�޿������"
            Height          =   195
            Index           =   204
            Left            =   390
            TabIndex        =   80
            Top             =   3735
            Width           =   3060
         End
         Begin VB.CheckBox chk 
            Caption         =   "������Һ�ģʽ"
            Height          =   195
            Index           =   290
            Left            =   390
            TabIndex        =   75
            Top             =   2445
            Width           =   3120
         End
         Begin VB.CheckBox chk 
            Caption         =   "ͬһ���ֻ֤�ܶ�Ӧһ����������"
            Height          =   195
            Index           =   28
            Left            =   390
            TabIndex        =   76
            Top             =   2715
            Width           =   3120
         End
         Begin MSComCtl2.DTPicker dtpRegistTime 
            Height          =   300
            Left            =   1935
            TabIndex        =   68
            Top             =   1320
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "HH:mm:ss"
            Format          =   167903235
            UpDown          =   -1  'True
            CurrentDate     =   42804
         End
         Begin VB.Frame fra 
            Caption         =   "�ҺŹ���"
            Height          =   1200
            Index           =   11
            Left            =   4470
            TabIndex        =   677
            Top             =   6510
            Width           =   4995
            Begin VB.CheckBox chk 
               Caption         =   "ԤԼ�ŶӰ�ʱ����ʾ"
               Height          =   195
               Index           =   186
               Left            =   180
               TabIndex        =   116
               Top             =   885
               Width           =   2280
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   180
               Index           =   24
               Left            =   2085
               TabIndex        =   115
               Text            =   "0"
               Top             =   600
               Width           =   360
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   180
               Index           =   4
               Left            =   2085
               TabIndex        =   113
               Text            =   "0"
               Top             =   315
               Width           =   360
            End
            Begin VB.CheckBox chk 
               Caption         =   "ר�Һ�ͬһ�����޹�    ����"
               Height          =   180
               Index           =   178
               Left            =   180
               TabIndex        =   112
               Top             =   315
               Width           =   2775
            End
            Begin VB.CheckBox chk 
               Caption         =   "ר�Һ�ͬһ������Լ    ����"
               Height          =   180
               Index           =   179
               Left            =   180
               TabIndex        =   114
               Top             =   600
               Width           =   3135
            End
            Begin VB.Line lnEpr 
               Index           =   2
               X1              =   2070
               X2              =   2460
               Y1              =   795
               Y2              =   795
            End
            Begin VB.Line lnEpr 
               Index           =   0
               X1              =   2070
               X2              =   2460
               Y1              =   510
               Y2              =   510
            End
         End
         Begin VB.Frame fra 
            Caption         =   "�ѻ�ҽ�����㷽ʽ"
            Height          =   2160
            Index           =   5
            Left            =   360
            TabIndex        =   676
            Top             =   5550
            Width           =   3795
            Begin VB.ListBox lst 
               Height          =   1740
               Index           =   5
               Left            =   165
               Style           =   1  'Checkbox
               TabIndex        =   88
               Top             =   300
               Width           =   3420
            End
         End
         Begin VB.Frame fra 
            Caption         =   "���ѿ�����"
            Height          =   1260
            Index           =   7
            Left            =   4470
            TabIndex        =   98
            Top             =   1980
            Width           =   4995
            Begin VB.CheckBox chk 
               Caption         =   "�����˷�ʱ���ѿ���Ҫˢ����֤"
               Height          =   180
               Index           =   59
               Left            =   180
               TabIndex        =   101
               Top             =   930
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "���ѿ�ˢ�������붨λ�������"
               Height          =   180
               Index           =   193
               Left            =   180
               TabIndex        =   100
               Top             =   600
               Width           =   2970
            End
            Begin VB.CheckBox chk 
               Caption         =   "�ɿ��������ӡ�ɿ"
               Height          =   180
               Index           =   163
               Left            =   180
               TabIndex        =   99
               Top             =   270
               Width           =   2415
            End
         End
         Begin VB.Frame fra 
            Caption         =   "�������תסԺ���"
            Height          =   1470
            Index           =   4
            Left            =   4470
            TabIndex        =   107
            Top             =   4950
            Width           =   4995
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   16
               ItemData        =   "frmParFee.frx":1A33A
               Left            =   1620
               List            =   "frmParFee.frx":1A33C
               Style           =   2  'Dropdown List
               TabIndex        =   111
               Top             =   1110
               Width           =   2115
            End
            Begin VB.CheckBox chk 
               Caption         =   "����������Ԥ������һ���Դ�ӡ"
               Height          =   180
               Index           =   197
               Left            =   180
               TabIndex        =   110
               Top             =   845
               Width           =   2865
            End
            Begin VB.CheckBox chk 
               Caption         =   "��Ժ������������תסԺ"
               Height          =   180
               Index           =   183
               Left            =   180
               TabIndex        =   109
               Top             =   580
               Width           =   2685
            End
            Begin VB.CheckBox chk 
               Caption         =   "����תסԺ���������"
               Height          =   180
               Index           =   159
               Left            =   180
               TabIndex        =   108
               Top             =   315
               Width           =   2100
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Ԥ��Ʊ�ݴ�ӡ��ʽ"
               Height          =   180
               Index           =   3
               Left            =   150
               TabIndex        =   787
               Top             =   1170
               Width           =   1440
            End
         End
         Begin VB.Frame fra 
            Caption         =   "ִ�еǼǹ���"
            Height          =   1530
            Index           =   3
            Left            =   4470
            TabIndex        =   102
            Top             =   3345
            Width           =   4995
            Begin VB.Frame fraRegPrint 
               Caption         =   "ִ�еǼǵ���ӡ��ʽ"
               Height          =   765
               Left            =   135
               TabIndex        =   678
               Top             =   630
               Width           =   3585
               Begin VB.OptionButton optRegPrint 
                  Caption         =   "����ӡ"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   104
                  Top             =   330
                  Width           =   915
               End
               Begin VB.OptionButton optRegPrint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   105
                  Top             =   330
                  Width           =   1050
               End
               Begin VB.OptionButton optRegPrint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   255
                  Index           =   2
                  Left            =   2160
                  TabIndex        =   106
                  Top             =   330
                  Width           =   1380
               End
               Begin VB.CommandButton cmdRegPrint 
                  Caption         =   "ִ�еǼǵ���ӡ����"
                  Height          =   350
                  Left            =   4230
                  TabIndex        =   679
                  Top             =   255
                  Width           =   1860
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʾҽ�����͵ĵ���"
               Height          =   195
               Index           =   158
               Left            =   165
               TabIndex        =   103
               Top             =   315
               Width           =   2100
            End
         End
         Begin VB.Frame fra 
            Caption         =   "�����շѡ����ۡ�����ʱ����"
            Height          =   870
            Index           =   6
            Left            =   360
            TabIndex        =   83
            Top             =   4575
            Width           =   3795
            Begin VB.CheckBox chk 
               Caption         =   "��������"
               Height          =   210
               Index           =   7
               Left            =   510
               TabIndex        =   84
               Top             =   270
               Value           =   1  'Checked
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һŵ���"
               Height          =   210
               Index           =   10
               Left            =   2040
               TabIndex        =   87
               Top             =   540
               Value           =   1  'Checked
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "���˱�ʶ"
               Height          =   180
               Index           =   8
               Left            =   2040
               TabIndex        =   85
               Top             =   270
               Value           =   1  'Checked
               Width           =   1035
            End
            Begin VB.CheckBox chk 
               Caption         =   "ˢ���￨"
               Height          =   210
               Index           =   9
               Left            =   510
               TabIndex        =   86
               Top             =   540
               Value           =   1  'Checked
               Width           =   1020
            End
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   10
            Left            =   2340
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   62
            Text            =   "1"
            Top             =   585
            Width           =   520
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   73
            Text            =   "5"
            Top             =   2060
            Width           =   930
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������Ŀ��λ��������"
            Height          =   195
            Index           =   56
            Left            =   390
            TabIndex        =   78
            Top             =   3195
            Width           =   2640
         End
         Begin VB.CheckBox chk 
            Caption         =   "������Ŀ���ܼ����ۿ۶�"
            Height          =   195
            Index           =   39
            Left            =   390
            TabIndex        =   79
            Top             =   3435
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������Ŀʱ�������"
            Height          =   195
            Index           =   25
            Left            =   390
            TabIndex        =   77
            Top             =   2955
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   65
            Text            =   "15"
            Top             =   960
            Width           =   930
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   70
            Text            =   "0"
            Top             =   1700
            Width           =   930
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   2340
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   59
            Text            =   "1"
            Top             =   240
            Width           =   520
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            ItemData        =   "frmParFee.frx":1A33E
            Left            =   1530
            List            =   "frmParFee.frx":1A340
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   4100
            Width           =   2625
         End
         Begin VB.Frame fra 
            Caption         =   "�㳮�������"
            Height          =   1680
            Index           =   15
            Left            =   4470
            TabIndex        =   89
            Top             =   180
            Width           =   4995
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   17
               Left            =   3075
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   360
               Width           =   1695
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   14
               Left            =   765
               Style           =   2  'Dropdown List
               TabIndex        =   95
               Top             =   1230
               Width           =   1695
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   13
               Left            =   765
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Top             =   780
               Width           =   1695
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   12
               Left            =   765
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "���ѿ�"
               Height          =   180
               Index           =   4
               Left            =   2520
               TabIndex        =   96
               Top             =   405
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "����"
               Height          =   180
               Index           =   38
               Left            =   360
               TabIndex        =   94
               Top             =   1275
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "�շ�"
               Height          =   180
               Index           =   37
               Left            =   360
               TabIndex        =   92
               Top             =   855
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "�Һ�"
               Height          =   180
               Index           =   15
               Left            =   360
               TabIndex        =   90
               Top             =   420
               Width           =   360
            End
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   5
            Left            =   2865
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1700
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(5)"
            BuddyDispid     =   196647
            BuddyIndex      =   5
            OrigLeft        =   2625
            OrigTop         =   1560
            OrigRight       =   2880
            OrigBottom      =   1860
            Max             =   4
            Min             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   1
            Left            =   2865
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(1)"
            BuddyDispid     =   196647
            BuddyIndex      =   1
            OrigLeft        =   2625
            OrigTop         =   120
            OrigRight       =   2880
            OrigBottom      =   420
            Max             =   7
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   0
            Left            =   2865
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   960
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   15
            BuddyControl    =   "txtUD(0)"
            BuddyDispid     =   196647
            BuddyIndex      =   0
            OrigLeft        =   2625
            OrigTop         =   1200
            OrigRight       =   2880
            OrigBottom      =   1500
            Max             =   365
            Min             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   6
            Left            =   2865
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   2060
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(6)"
            BuddyDispid     =   196647
            BuddyIndex      =   6
            OrigLeft        =   2625
            OrigTop         =   1920
            OrigRight       =   2880
            OrigBottom      =   2220
            Max             =   5
            Min             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   10
            Left            =   2865
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   585
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(10)"
            BuddyDispid     =   196647
            BuddyIndex      =   10
            OrigLeft        =   2625
            OrigTop         =   465
            OrigRight       =   2880
            OrigBottom      =   765
            Max             =   7
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��Դ����ʱ��"
            Height          =   180
            Index           =   2
            Left            =   750
            TabIndex        =   67
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����Һŵ���Ч������"
            Height          =   180
            Index           =   49
            Left            =   390
            TabIndex        =   61
            Top             =   630
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���õ��۱���λ��"
            Height          =   180
            Index           =   35
            Left            =   390
            TabIndex        =   72
            Top             =   2110
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Һ�����ԤԼ����"
            Height          =   180
            Index           =   30
            Left            =   390
            TabIndex        =   64
            Top             =   1020
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ý���λ��"
            Height          =   180
            Index           =   28
            Left            =   390
            TabIndex        =   69
            Top             =   1750
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ͨ�Һŵ���Ч������"
            Height          =   180
            Index           =   16
            Left            =   390
            TabIndex        =   58
            Top             =   300
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������˷�ʽ"
            Height          =   180
            Index           =   52
            Left            =   390
            TabIndex        =   81
            Top             =   4160
            Width           =   1080
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7215
         Index           =   13
         Left            =   -74895
         ScaleHeight     =   7185
         ScaleWidth      =   9180
         TabIndex        =   555
         TabStop         =   0   'False
         Top             =   795
         Visible         =   0   'False
         Width           =   9210
         Begin VB.Frame fraƱ�ݸ�ʽ 
            Caption         =   "�˷�Ʊ�ݸ�ʽ"
            Height          =   1545
            Index           =   6
            Left            =   240
            TabIndex        =   579
            Top             =   5970
            Width           =   5910
            Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
               Height          =   1215
               Index           =   6
               Left            =   120
               TabIndex        =   580
               Top             =   240
               Width           =   5715
               _cx             =   10081
               _cy             =   2143
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
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   8421504
               GridColorFixed  =   8421504
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
               Rows            =   3
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmParFee.frx":1A342
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
               ExplorerBar     =   2
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
         Begin VB.Frame fraƱ�ݸ�ʽ 
            Caption         =   "�շ�Ʊ�ݸ�ʽ"
            Height          =   1545
            Index           =   2
            Left            =   240
            TabIndex        =   577
            Top             =   4335
            Width           =   5910
            Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
               Height          =   1215
               Index           =   2
               Left            =   120
               TabIndex        =   578
               Top             =   240
               Width           =   5715
               _cx             =   10081
               _cy             =   2143
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
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   8421504
               GridColorFixed  =   8421504
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
               Rows            =   3
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmParFee.frx":1A3D0
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
               ExplorerBar     =   2
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
         Begin VB.Frame fraSupplementaryPrint 
            Caption         =   "�����嵥��ӡ��ʽ"
            Height          =   795
            Left            =   255
            TabIndex        =   573
            Top             =   3435
            Width           =   5910
            Begin VB.OptionButton optSupplementaryPrint 
               Caption         =   "����ӡ"
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   574
               Top             =   390
               Value           =   -1  'True
               Width           =   900
            End
            Begin VB.OptionButton optSupplementaryPrint 
               Caption         =   "ѡ���Ƿ��ӡ"
               Height          =   180
               Index           =   2
               Left            =   2505
               TabIndex        =   576
               Top             =   390
               Width           =   1455
            End
            Begin VB.OptionButton optSupplementaryPrint 
               Caption         =   "�Զ���ӡ"
               Height          =   180
               Index           =   1
               Left            =   1275
               TabIndex        =   575
               Top             =   390
               Width           =   1065
            End
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   12
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   563
            Text            =   "3"
            Top             =   1005
            Width           =   375
         End
         Begin VB.Frame fra���㷽ʽ 
            Caption         =   "֧���շѽ��㷽ʽ"
            Height          =   6705
            Left            =   6750
            TabIndex        =   581
            Top             =   270
            Width           =   2295
            Begin VB.ListBox lst 
               Height          =   6360
               Index           =   2
               ItemData        =   "frmParFee.frx":1A45E
               Left            =   165
               List            =   "frmParFee.frx":1A460
               Style           =   1  'Checkbox
               TabIndex        =   582
               Top             =   270
               Width           =   2010
            End
         End
         Begin VB.Frame fraSupplementaryMode 
            Caption         =   "ҩƷ��ҩ���˷ѷ�ʽ"
            Height          =   795
            Left            =   255
            TabIndex        =   565
            Top             =   1440
            Width           =   5910
            Begin VB.OptionButton optDrugSupplementary 
               Caption         =   "����"
               Height          =   180
               Index           =   2
               Left            =   2505
               TabIndex        =   568
               Top             =   420
               Width           =   690
            End
            Begin VB.OptionButton optDrugSupplementary 
               Caption         =   "��ֹ"
               Height          =   180
               Index           =   1
               Left            =   1470
               TabIndex        =   567
               Top             =   420
               Width           =   690
            End
            Begin VB.OptionButton optDrugSupplementary 
               Caption         =   "�����"
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   566
               Top             =   405
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   11
            Left            =   1275
            Locked          =   -1  'True
            TabIndex        =   560
            Text            =   "10"
            Top             =   585
            Width           =   480
         End
         Begin VB.Frame fra 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   18
            Left            =   2865
            TabIndex        =   558
            Top             =   465
            Width           =   285
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   19
            Left            =   2865
            MaxLength       =   3
            TabIndex        =   557
            Text            =   "0"
            Top             =   270
            Width           =   285
         End
         Begin VB.Frame fra��λ 
            Caption         =   " ҩƷ��λ "
            Height          =   795
            Index           =   3
            Left            =   255
            TabIndex        =   569
            Top             =   2415
            Width           =   5910
            Begin VB.OptionButton optSupplementaryUnit 
               Caption         =   "����(��סԺ)��λ"
               Height          =   180
               Index           =   1
               Left            =   2505
               TabIndex        =   572
               Top             =   405
               Width           =   1770
            End
            Begin VB.OptionButton optSupplementaryUnit 
               Caption         =   "�ۼ۵�λ"
               Height          =   180
               Index           =   0
               Left            =   1275
               TabIndex        =   571
               Top             =   405
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.Label lbl��λ 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������ʱ��"
               Height          =   180
               Index           =   3
               Left            =   300
               TabIndex        =   570
               Top             =   405
               Width           =   900
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ͨ������������ģ������    ���ڵĲ�����Ϣ"
            Height          =   195
            Index           =   136
            Left            =   240
            TabIndex        =   556
            Top             =   270
            Width           =   4260
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   11
            Left            =   1740
            TabIndex        =   561
            TabStop         =   0   'False
            Top             =   585
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   10
            BuddyControl    =   "txtUD(11)"
            BuddyDispid     =   196647
            BuddyIndex      =   11
            OrigLeft        =   1740
            OrigTop         =   585
            OrigRight       =   1995
            OrigBottom      =   885
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chk 
            Caption         =   "Ʊ��ʣ��         ��ʱ��ʼ�����շ�Ա"
            Height          =   285
            Index           =   137
            Left            =   240
            TabIndex        =   559
            Top             =   585
            Width           =   3450
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   12
            Left            =   3510
            TabIndex        =   564
            TabStop         =   0   'False
            Top             =   1005
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   3
            BuddyControl    =   "txtUD(12)"
            BuddyDispid     =   196647
            BuddyIndex      =   12
            OrigLeft        =   1605
            OrigTop         =   4860
            OrigRight       =   1860
            OrigBottom      =   5160
            Max             =   100
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblVaildDays 
            Caption         =   "�ɽ��б��ղ������ķ�����Ч����"
            Height          =   225
            Left            =   240
            TabIndex        =   562
            Top             =   1035
            Width           =   2895
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7530
         Index           =   12
         Left            =   -74880
         ScaleHeight     =   7500
         ScaleWidth      =   9660
         TabIndex        =   360
         TabStop         =   0   'False
         Top             =   810
         Visible         =   0   'False
         Width           =   9690
         Begin VB.Frame fraNormal 
            Height          =   1725
            Index           =   0
            Left            =   285
            TabIndex        =   361
            Top             =   120
            Width           =   4575
            Begin VB.CheckBox chk 
               Caption         =   "����¼������ʹ�õĿ�����"
               Height          =   195
               Index           =   195
               Left            =   180
               TabIndex        =   369
               Top             =   1440
               Width           =   2490
            End
            Begin VB.Frame fraLineDays 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Index           =   2
               Left            =   2820
               TabIndex        =   367
               Top             =   1050
               Width           =   285
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   16
               Left            =   2835
               MaxLength       =   3
               TabIndex        =   366
               Text            =   "0"
               Top             =   840
               Width           =   285
            End
            Begin VB.CheckBox chk 
               Caption         =   "ֻ���Һ�Լ��λ����"
               Height          =   195
               Index           =   132
               Left            =   180
               TabIndex        =   368
               Top             =   1155
               Width           =   2100
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ҩ���븶��"
               Height          =   195
               Index           =   71
               Left            =   180
               TabIndex        =   362
               Top             =   285
               Value           =   1  'Checked
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����������"
               Height          =   195
               Index           =   72
               Left            =   2640
               TabIndex        =   363
               Top             =   285
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����˺���ʿ"
               Height          =   195
               Index           =   77
               Left            =   180
               TabIndex        =   364
               Top             =   570
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ͨ������������ģ������    ���ڵĲ�����Ϣ"
               Height          =   195
               Index           =   87
               Left            =   180
               TabIndex        =   365
               Top             =   855
               Width           =   4260
            End
         End
         Begin VB.Frame fraPrintBill 
            Caption         =   "���ݴ�ӡ"
            Height          =   1305
            Left            =   270
            TabIndex        =   374
            Top             =   2940
            Width           =   4575
            Begin VB.CheckBox chk 
               Caption         =   "���ʱ��ӡ���ʵ���"
               Height          =   195
               Index           =   135
               Left            =   285
               TabIndex        =   378
               Top             =   930
               Width           =   2505
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ʱ��ӡ���ۼ��ʵ���"
               Height          =   195
               Index           =   134
               Left            =   285
               TabIndex        =   376
               Top             =   630
               Width           =   2670
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ʱ��ӡ���ʵ���"
               Height          =   195
               Index           =   133
               Left            =   285
               TabIndex        =   375
               Top             =   330
               Width           =   2220
            End
         End
         Begin VB.Frame fra�����ʾ 
            Caption         =   "�����ʾ"
            Height          =   1305
            Index           =   2
            Left            =   270
            TabIndex        =   379
            Top             =   4455
            Width           =   4575
            Begin VB.OptionButton opt���ʿ����ʾ��ʽ 
               Caption         =   "����ʾ����"
               Height          =   180
               Index           =   1
               Left            =   2760
               TabIndex        =   384
               Top             =   945
               Width           =   1215
            End
            Begin VB.OptionButton opt���ʿ����ʾ��ʽ 
               Caption         =   "��ʾ�����"
               Height          =   180
               Index           =   0
               Left            =   1365
               TabIndex        =   383
               Top             =   945
               Width           =   1290
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʾ����ҩ�����"
               Height          =   195
               Index           =   83
               Left            =   150
               TabIndex        =   380
               Top             =   375
               Width           =   1770
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʾ����ҩ����"
               Height          =   195
               Index           =   82
               Left            =   2250
               TabIndex        =   381
               Top             =   390
               Width           =   1770
            End
            Begin VB.Line lnSplit 
               BorderColor     =   &H00FFFFFF&
               Index           =   5
               X1              =   15
               X2              =   4545
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Line lnSplit 
               BorderColor     =   &H80000000&
               Index           =   4
               X1              =   15
               X2              =   4545
               Y1              =   705
               Y2              =   705
            End
            Begin VB.Label lbl�����ʾ��ʽ 
               AutoSize        =   -1  'True
               Caption         =   "�����ʾ��ʽ"
               Height          =   180
               Index           =   2
               Left            =   150
               TabIndex        =   382
               Top             =   945
               Width           =   1080
            End
         End
         Begin VB.Frame fra��λ 
            Caption         =   " ҩƷ��λ "
            Height          =   735
            Index           =   2
            Left            =   285
            TabIndex        =   370
            Top             =   2025
            Width           =   4575
            Begin VB.OptionButton opt���ʵ�λ 
               Caption         =   "����(��סԺ)��λ"
               Height          =   180
               Index           =   1
               Left            =   2085
               TabIndex        =   373
               Top             =   375
               Width           =   1755
            End
            Begin VB.OptionButton opt���ʵ�λ 
               Caption         =   "�ۼ۵�λ"
               Height          =   180
               Index           =   0
               Left            =   1020
               TabIndex        =   372
               Top             =   375
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.Label lbl��λ 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ��"
               Height          =   180
               Index           =   2
               Left            =   15
               TabIndex        =   371
               Top             =   360
               Width           =   975
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7530
         Index           =   10
         Left            =   -74925
         ScaleHeight     =   7500
         ScaleWidth      =   10065
         TabIndex        =   358
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   10095
         Begin VB.Frame fraNormal 
            Caption         =   "���ݿ���"
            Height          =   2910
            Index           =   1
            Left            =   285
            TabIndex        =   385
            Top             =   240
            Width           =   4680
            Begin VB.CheckBox chk 
               Caption         =   "����¼������ʹ�õĿ�����"
               Height          =   195
               Index           =   194
               Left            =   240
               TabIndex        =   400
               Top             =   2640
               Width           =   2490
            End
            Begin VB.CheckBox chk 
               Caption         =   "סԺ���˰������շ�"
               Height          =   270
               Index           =   168
               Left            =   240
               TabIndex        =   392
               Top             =   1294
               Width           =   1920
            End
            Begin VB.TextBox txtUD 
               ForeColor       =   &H80000012&
               Height          =   285
               Index           =   7
               Left            =   1365
               MaxLength       =   3
               TabIndex        =   398
               Text            =   "0"
               Top             =   2310
               Width           =   600
            End
            Begin VB.TextBox txt 
               ForeColor       =   &H80000012&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   12
               Left            =   1365
               MaxLength       =   12
               TabIndex        =   396
               Text            =   "0.00"
               Top             =   1935
               Width           =   1335
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   14
               Left            =   2895
               MaxLength       =   3
               TabIndex        =   394
               Text            =   "0"
               Top             =   1635
               Width           =   285
            End
            Begin VB.Frame fraLineDays 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Index           =   0
               Left            =   2895
               TabIndex        =   401
               Top             =   1815
               Width           =   285
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʹ��ȱʡ������"
               Height          =   195
               Index           =   110
               Left            =   240
               TabIndex        =   388
               Top             =   672
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "����Ҫ���뿪����"
               Height          =   195
               Index           =   111
               Left            =   2550
               TabIndex        =   389
               Top             =   672
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "ȱʡ��������"
               Height          =   195
               Index           =   112
               Left            =   240
               TabIndex        =   390
               Top             =   1014
               Width           =   1620
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ҩ���븶��"
               Height          =   195
               Index           =   69
               Left            =   240
               TabIndex        =   386
               Top             =   330
               Value           =   1  'Checked
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����������"
               Height          =   195
               Index           =   74
               Left            =   2550
               TabIndex        =   387
               Top             =   330
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����˺���ʿ"
               Height          =   195
               Index           =   75
               Left            =   2550
               TabIndex        =   391
               Top             =   1014
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ͨ������������ģ������    ���ڵĲ�����Ϣ"
               Height          =   195
               Index           =   85
               Left            =   240
               TabIndex        =   393
               Top             =   1650
               Width           =   4260
            End
            Begin MSComCtl2.UpDown ud 
               Height          =   270
               Index           =   7
               Left            =   1950
               TabIndex        =   399
               TabStop         =   0   'False
               Top             =   2310
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   476
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtUD(7)"
               BuddyDispid     =   196647
               BuddyIndex      =   7
               OrigLeft        =   5265
               OrigTop         =   555
               OrigRight       =   5505
               OrigBottom      =   825
               Max             =   32767
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label lblDay 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ȡ�����۳���           ��δ����Ļ��۵�"
               Height          =   180
               Left            =   240
               TabIndex        =   397
               Top             =   2355
               Width           =   3510
            End
            Begin VB.Label lblMax 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���������"
               Height          =   180
               Left            =   240
               TabIndex        =   395
               Top             =   1995
               Width           =   1080
            End
         End
         Begin VB.Frame fra����֪ͨ����ӡ 
            Caption         =   "����֪ͨ����ӡ"
            Height          =   1290
            Left            =   5565
            TabIndex        =   420
            Top             =   240
            Width           =   3720
            Begin VB.OptionButton optPrintRequisition 
               Caption         =   "�Զ���ӡ"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   422
               Top             =   585
               Value           =   -1  'True
               Width           =   1260
            End
            Begin VB.OptionButton optPrintRequisition 
               Caption         =   "����ӡ"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   421
               Top             =   330
               Width           =   1020
            End
            Begin VB.OptionButton optPrintRequisition 
               Caption         =   "ѡ���Ƿ��ӡ"
               Height          =   180
               Index           =   2
               Left            =   240
               TabIndex        =   423
               Top             =   855
               Width           =   1500
            End
         End
         Begin VB.Frame fra 
            Caption         =   "��������ʾ��ʽ"
            Height          =   1290
            Index           =   1
            Left            =   5565
            TabIndex        =   424
            Top             =   1665
            Width           =   3720
            Begin VB.OptionButton optBillTotalShow 
               Caption         =   "���վݷ�Ŀ��ʾ����ϼ�"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   425
               Top             =   345
               Value           =   -1  'True
               Width           =   2280
            End
            Begin VB.OptionButton optBillTotalShow 
               Caption         =   "��������Ŀ��ʾ����ϼ�"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   426
               Top             =   615
               Width           =   2280
            End
            Begin VB.OptionButton optBillTotalShow 
               Caption         =   "�����ݷ��������ʾ"
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   427
               Top             =   870
               Width           =   2280
            End
         End
         Begin VB.Frame fraBillInputItem 
            Caption         =   "����ʱҪ�������Ŀ"
            Height          =   1050
            Index           =   0
            Left            =   285
            TabIndex        =   412
            Top             =   5730
            Width           =   4680
            Begin VB.CheckBox chk 
               Caption         =   "ҽ�Ƹ��ʽ"
               Height          =   210
               Index           =   95
               Left            =   2940
               TabIndex        =   419
               Top             =   675
               Value           =   1  'Checked
               Width           =   1380
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Ա�"
               Height          =   210
               Index           =   88
               Left            =   240
               TabIndex        =   413
               Top             =   360
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "����"
               Height          =   210
               Index           =   91
               Left            =   2940
               TabIndex        =   415
               Top             =   360
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "�ѱ�"
               Height          =   210
               Index           =   92
               Left            =   3810
               TabIndex        =   416
               Top             =   360
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Ƿ�Ӱ�"
               Height          =   210
               Index           =   89
               Left            =   1500
               TabIndex        =   414
               Top             =   360
               Value           =   1  'Checked
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "��������"
               Height          =   210
               Index           =   93
               Left            =   240
               TabIndex        =   417
               Top             =   675
               Value           =   1  'Checked
               Width           =   1020
            End
            Begin VB.CheckBox chk 
               Caption         =   "������"
               Height          =   210
               Index           =   94
               Left            =   1500
               TabIndex        =   418
               Top             =   675
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.Frame fra�����ʾ 
            Caption         =   "�����ʾ"
            Height          =   1275
            Index           =   0
            Left            =   285
            TabIndex        =   406
            Top             =   4260
            Width           =   4680
            Begin VB.OptionButton opt���ۿ����ʾ��ʽ 
               Caption         =   "����ʾ����"
               Height          =   180
               Index           =   1
               Left            =   2835
               TabIndex        =   411
               Top             =   900
               Width           =   1215
            End
            Begin VB.OptionButton opt���ۿ����ʾ��ʽ 
               Caption         =   "��ʾ�����"
               Height          =   180
               Index           =   0
               Left            =   1440
               TabIndex        =   410
               Top             =   900
               Width           =   1290
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʾ����ҩ�����"
               Height          =   195
               Index           =   78
               Left            =   240
               TabIndex        =   407
               Top             =   375
               Width           =   1770
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʾ����ҩ����"
               Height          =   195
               Index           =   79
               Left            =   2355
               TabIndex        =   408
               Top             =   375
               Width           =   1770
            End
            Begin VB.Line lnSplit 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   15
               X2              =   4625
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Line lnSplit 
               BorderColor     =   &H80000000&
               Index           =   1
               X1              =   15
               X2              =   4640
               Y1              =   705
               Y2              =   705
            End
            Begin VB.Label lbl�����ʾ��ʽ 
               AutoSize        =   -1  'True
               Caption         =   "�����ʾ��ʽ"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   409
               Top             =   900
               Width           =   1080
            End
         End
         Begin VB.Frame fra��λ 
            Caption         =   " ҩƷ��λ "
            Height          =   780
            Index           =   0
            Left            =   285
            TabIndex        =   402
            Top             =   3315
            Width           =   4680
            Begin VB.OptionButton opt���۵�λ 
               Caption         =   "����(��סԺ)��λ"
               Height          =   180
               Index           =   1
               Left            =   2250
               TabIndex        =   405
               Top             =   405
               Width           =   1740
            End
            Begin VB.OptionButton opt���۵�λ 
               Caption         =   "�ۼ۵�λ"
               Height          =   180
               Index           =   0
               Left            =   1050
               TabIndex        =   404
               Top             =   405
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.Label lbl��λ 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ��"
               Height          =   180
               Index           =   0
               Left            =   15
               TabIndex        =   403
               Top             =   405
               Width           =   975
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6810
         Index           =   9
         Left            =   -74925
         ScaleHeight     =   6780
         ScaleWidth      =   9000
         TabIndex        =   327
         TabStop         =   0   'False
         Top             =   780
         Width           =   9030
         Begin VB.CheckBox chk 
            Caption         =   "����̨ǩ����ʼ�Ŷ�"
            Height          =   330
            Index           =   65
            Left            =   4890
            TabIndex        =   352
            Top             =   210
            Width           =   1935
         End
         Begin VB.Frame fra����̨ǩ���Ŷ� 
            Height          =   7245
            Left            =   4740
            TabIndex        =   351
            Top             =   240
            Width           =   4635
            Begin VB.CommandButton cmdDepClearAll 
               Caption         =   "ȫ��"
               Height          =   350
               Left            =   3300
               TabIndex        =   355
               Top             =   6420
               Width           =   1100
            End
            Begin VB.CommandButton cmdDepSelectAll 
               Caption         =   "ȫѡ"
               Height          =   350
               Left            =   2100
               TabIndex        =   354
               Top             =   6420
               Width           =   1100
            End
            Begin VB.CheckBox chk 
               Caption         =   "�ٴ�ǩ���������Ŷ�"
               Height          =   330
               Index           =   177
               Left            =   2520
               TabIndex        =   357
               Top             =   6870
               Width           =   1935
            End
            Begin VB.CheckBox chk 
               Caption         =   "���ﲡ���������Ŷ�"
               Height          =   330
               Index           =   171
               Left            =   210
               TabIndex        =   356
               Top             =   6870
               Width           =   1935
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfTriageQueuingDep 
               Height          =   5985
               Left            =   150
               TabIndex        =   353
               Top             =   300
               Width           =   4305
               _cx             =   7594
               _cy             =   10557
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
               BackColorSel    =   16772055
               ForeColorSel    =   0
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   12698049
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   255
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmParFee.frx":1A462
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   1
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
               Editable        =   2
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
         End
         Begin VB.Frame fra���� 
            Caption         =   "���ﲡ������ʽ"
            Height          =   1260
            Left            =   210
            TabIndex        =   335
            Top             =   2040
            Width           =   4350
            Begin VB.OptionButton optTriageSort 
               Caption         =   "���ұ���,����,���ݺ�"
               Height          =   210
               Index           =   0
               Left            =   435
               TabIndex        =   336
               Top             =   330
               Value           =   -1  'True
               Width           =   2280
            End
            Begin VB.OptionButton optTriageSort 
               Caption         =   "���ұ���,����,�Һ�ʱ��"
               Height          =   210
               Index           =   1
               Left            =   435
               TabIndex        =   337
               Top             =   600
               Width           =   2280
            End
            Begin VB.OptionButton optTriageSort 
               Caption         =   "���ұ���,����,����ʱ��,�Ǽ�ʱ��"
               Height          =   210
               Index           =   2
               Left            =   435
               TabIndex        =   338
               Top             =   885
               Width           =   3555
            End
         End
         Begin VB.Frame fra�Ŷӵ� 
            Caption         =   "�Ŷӵ���ӡ"
            Height          =   825
            Left            =   210
            TabIndex        =   343
            Top             =   4515
            Width           =   4350
            Begin VB.OptionButton optTriagePrintMode 
               Caption         =   "����ӡ"
               Height          =   195
               Index           =   0
               Left            =   435
               TabIndex        =   344
               Top             =   375
               Value           =   -1  'True
               Width           =   990
            End
            Begin VB.OptionButton optTriagePrintMode 
               Caption         =   "�Զ���ӡ"
               Height          =   195
               Index           =   1
               Left            =   1455
               TabIndex        =   345
               Top             =   375
               Width           =   1125
            End
            Begin VB.OptionButton optTriagePrintMode 
               Caption         =   "��ʾѡ���ӡ"
               Height          =   195
               Index           =   2
               Left            =   2700
               TabIndex        =   346
               Top             =   375
               Width           =   1455
            End
         End
         Begin VB.Frame fra���� 
            Caption         =   "�����ӡ"
            Height          =   855
            Left            =   210
            TabIndex        =   339
            Top             =   3495
            Width           =   4350
            Begin VB.OptionButton optTriageBarcodePrintMode 
               Caption         =   "����ӡ"
               Height          =   195
               Index           =   0
               Left            =   435
               TabIndex        =   340
               Top             =   360
               Value           =   -1  'True
               Width           =   990
            End
            Begin VB.OptionButton optTriageBarcodePrintMode 
               Caption         =   "�Զ���ӡ"
               Height          =   195
               Index           =   1
               Left            =   1455
               TabIndex        =   341
               Top             =   360
               Width           =   1170
            End
            Begin VB.OptionButton optTriageBarcodePrintMode 
               Caption         =   "��ʾѡ���ӡ"
               Height          =   195
               Index           =   2
               Left            =   2700
               TabIndex        =   342
               Top             =   360
               Width           =   1395
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҽ������æʱ�������"
            Height          =   300
            Index           =   68
            Left            =   210
            TabIndex        =   333
            Top             =   1140
            Width           =   2460
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   11
            Left            =   960
            MaxLength       =   2
            TabIndex        =   332
            Text            =   "0"
            Top             =   585
            Width           =   435
         End
         Begin VB.CheckBox chk 
            Caption         =   "ԤԼ�ҺŽ������"
            Height          =   270
            Index           =   66
            Left            =   210
            TabIndex        =   334
            Top             =   1485
            Width           =   1905
         End
         Begin VB.Frame fra�Ŷӽк� 
            Caption         =   "�Ŷӽк�ģʽ"
            Height          =   1875
            Left            =   210
            TabIndex        =   347
            Top             =   5610
            Width           =   4350
            Begin VB.OptionButton optTriageQueuingMode 
               Caption         =   "��ֹȫԺ�Ŷӽк�"
               Height          =   240
               Index           =   0
               Left            =   435
               TabIndex        =   348
               Top             =   390
               Width           =   1770
            End
            Begin VB.OptionButton optTriageQueuingMode 
               Caption         =   "����̨������л�ҽ����������"
               Height          =   240
               Index           =   1
               Left            =   435
               TabIndex        =   350
               Top             =   1005
               Width           =   3045
            End
            Begin VB.OptionButton optTriageQueuingMode 
               Caption         =   "�ȷ������,��ҽ�����о���"
               Height          =   240
               Index           =   2
               Left            =   435
               TabIndex        =   349
               Top             =   705
               Width           =   2625
            End
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   329
            Text            =   "0"
            Top             =   210
            Width           =   420
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   3
            Left            =   1380
            TabIndex        =   330
            TabStop         =   0   'False
            Top             =   210
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtUD(3)"
            BuddyDispid     =   196647
            BuddyIndex      =   3
            OrigLeft        =   1635
            OrigTop         =   195
            OrigRight       =   1890
            OrigBottom      =   495
            Max             =   7
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl��ǰ���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ǰ      Сʱ����"
            Height          =   180
            Left            =   525
            TabIndex        =   331
            Top             =   615
            Width           =   1620
         End
         Begin VB.Label lbl��Ч���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Զ�ˢ��         ���ڵĹҺŲ���"
            Height          =   180
            Left            =   180
            TabIndex        =   328
            Top             =   270
            Width           =   3045
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7470
         Index           =   8
         Left            =   -74850
         ScaleHeight     =   7440
         ScaleWidth      =   9060
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   690
         Visible         =   0   'False
         Width           =   9090
         Begin VB.CheckBox chk 
            Caption         =   "������󶨿�ʱ�Զ����ɲ��������"
            Height          =   180
            Index           =   201
            Left            =   150
            TabIndex        =   167
            Top             =   1620
            Width           =   3225
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ȡ������"
            Height          =   180
            Index           =   199
            Left            =   150
            TabIndex        =   166
            Top             =   1290
            Width           =   2535
         End
         Begin VB.Frame fraƱ�ݸ�ʽ 
            Caption         =   "�շ�Ʊ�ݸ�ʽ"
            Height          =   1455
            Index           =   9
            Left            =   150
            TabIndex        =   766
            Top             =   5160
            Width           =   4155
            Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
               Height          =   1125
               Index           =   9
               Left            =   90
               TabIndex        =   185
               Top             =   255
               Width           =   3975
               _cx             =   7011
               _cy             =   1984
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
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   8421504
               GridColorFixed  =   8421504
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
               Rows            =   3
               Cols            =   2
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmParFee.frx":1A4EC
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
               ExplorerBar     =   2
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
         Begin VB.CheckBox chk 
            Caption         =   "����ʹ�������շ�ҽ���վ�"
            Height          =   180
            Index           =   192
            Left            =   150
            TabIndex        =   165
            Top             =   960
            Width           =   2535
         End
         Begin VB.Frame fraƱ�ݸ�ʽ 
            Height          =   120
            Index           =   8
            Left            =   1575
            TabIndex        =   680
            Top             =   6645
            Width           =   7935
         End
         Begin VB.Frame fraPrintMode_SendCard 
            Caption         =   "�����Ͱ󶨿�ƾ�ݴ�ӡ��ʽ"
            Height          =   1440
            Left            =   150
            TabIndex        =   168
            Top             =   1935
            Width           =   4155
            Begin VB.OptionButton optPrintMode_SendCard 
               Caption         =   "�Զ���ӡ"
               Height          =   180
               Index           =   1
               Left            =   300
               TabIndex        =   170
               Top             =   660
               Width           =   1020
            End
            Begin VB.OptionButton optPrintMode_SendCard 
               Caption         =   "ѡ���Ƿ��ӡ"
               Height          =   180
               Index           =   2
               Left            =   300
               TabIndex        =   173
               Top             =   975
               Width           =   1380
            End
            Begin VB.OptionButton optPrintMode_SendCard 
               Caption         =   "����ӡ"
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   169
               Top             =   345
               Value           =   -1  'True
               Width           =   900
            End
         End
         Begin VB.Frame fraShortLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   2790
            TabIndex        =   176
            Top             =   810
            Width           =   285
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   2790
            MaxLength       =   3
            TabIndex        =   164
            Text            =   "0"
            Top             =   630
            Width           =   285
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����Լ��˷�ʽ��ȡ"
            Height          =   180
            Index           =   14
            Left            =   150
            TabIndex        =   160
            Top             =   285
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.Frame fra�˿���ʽ 
            Caption         =   "�˿���ʽ����"
            Height          =   1695
            Left            =   150
            TabIndex        =   174
            Top             =   3420
            Width           =   4155
            Begin VB.OptionButton optDelCardMode 
               Caption         =   "���뵥�ݺ��˿���ˢ���˿�"
               Height          =   180
               Index           =   3
               Left            =   300
               TabIndex        =   183
               Top             =   1305
               Width           =   2460
            End
            Begin VB.OptionButton optDelCardMode 
               Caption         =   "���뵥�ݺź��ˢ���˿�"
               Height          =   180
               Index           =   2
               Left            =   315
               TabIndex        =   182
               Top             =   1005
               Width           =   2460
            End
            Begin VB.OptionButton optDelCardMode 
               Caption         =   "����ˢ���˿�"
               Height          =   180
               Index           =   1
               Left            =   315
               TabIndex        =   177
               Top             =   690
               Width           =   1740
            End
            Begin VB.OptionButton optDelCardMode 
               Caption         =   "������ˢ����֤"
               Height          =   180
               Index           =   0
               Left            =   315
               TabIndex        =   175
               Top             =   375
               Value           =   -1  'True
               Width           =   1740
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsInputItemSet 
            Height          =   6060
            Index           =   0
            Left            =   4830
            TabIndex        =   186
            Top             =   540
            Width           =   4380
            _cx             =   7726
            _cy             =   10689
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
            BackColor       =   -2147483634
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483634
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483643
            FloodColor      =   192
            SheetBorder     =   -2147483637
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParFee.frx":1A54D
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
         Begin VB.CheckBox chk 
            Caption         =   "����ͨ������������ģ������    ���ڵĲ���"
            Height          =   195
            Index           =   15
            Left            =   150
            TabIndex        =   162
            Top             =   630
            Width           =   4260
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   975
            Index           =   8
            Left            =   210
            TabIndex        =   187
            Top             =   6975
            Width           =   8925
            _cx             =   15743
            _cy             =   1720
            Appearance      =   2
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
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
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParFee.frx":1A5F1
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
            ExplorerBar     =   2
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
         Begin VB.Label lblCardDeposit 
            AutoSize        =   -1  'True
            Caption         =   "Ԥ��Ʊ�ݴ�ӡ����"
            Height          =   180
            Left            =   105
            TabIndex        =   681
            Top             =   6675
            Width           =   1440
         End
         Begin VB.Label lblInputSendCardSet 
            AutoSize        =   -1  'True
            Caption         =   "���������"
            Height          =   180
            Left            =   4800
            TabIndex        =   184
            Top             =   285
            Width           =   900
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7860
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   7830
         ScaleWidth      =   9210
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   660
         Visible         =   0   'False
         Width           =   9240
         Begin VB.Frame fra 
            Caption         =   "�����˷�ˢ������(�˻�Ԥ���)"
            Height          =   1035
            Index           =   13
            Left            =   4770
            TabIndex        =   783
            Top             =   6600
            Width           =   4320
            Begin VB.OptionButton optBrushCard 
               Caption         =   "��ֹˢ��"
               Height          =   180
               Index           =   10
               Left            =   240
               TabIndex        =   666
               Top             =   330
               Width           =   1035
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "����ˢ����֤"
               Height          =   180
               Index           =   11
               Left            =   1560
               TabIndex        =   667
               Top             =   360
               Value           =   -1  'True
               Width           =   1425
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "ҽ�ƿ����������룬����ˢ����֤"
               Height          =   180
               Index           =   12
               Left            =   240
               TabIndex        =   668
               Top             =   720
               Width           =   3045
            End
         End
         Begin VB.Frame fra 
            Caption         =   "��������ˢ������(ʹ��Ԥ���)"
            Height          =   1395
            Index           =   12
            Left            =   270
            TabIndex        =   782
            Top             =   6360
            Width           =   4215
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   26
               Left            =   480
               MaxLength       =   8
               TabIndex        =   792
               Text            =   "0"
               Top             =   1020
               Width           =   795
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "_________Ԫ������֧��"
               Height          =   180
               Index           =   3
               Left            =   210
               TabIndex        =   791
               Top             =   1080
               Width           =   3045
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "ҽ�ƿ����������룬����ˢ����֤"
               Height          =   180
               Index           =   2
               Left            =   210
               TabIndex        =   661
               Top             =   720
               Width           =   3045
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "����ˢ����֤"
               Height          =   180
               Index           =   1
               Left            =   1530
               TabIndex        =   660
               Top             =   375
               Value           =   -1  'True
               Width           =   1425
            End
            Begin VB.OptionButton optBrushCard 
               Caption         =   "��ֹˢ��"
               Height          =   180
               Index           =   0
               Left            =   210
               TabIndex        =   659
               Top             =   375
               Width           =   1035
            End
         End
         Begin VB.Frame fra 
            Caption         =   "�����շѿ���"
            Height          =   1095
            Index           =   14
            Left            =   270
            TabIndex        =   651
            Top             =   1860
            Width           =   4215
            Begin VB.CheckBox chk 
               Caption         =   "��Ŀִ��ǰ�������շѻ��ȼ������"
               Height          =   210
               Index           =   67
               Left            =   180
               TabIndex        =   652
               Top             =   360
               Width           =   3540
            End
            Begin VB.CheckBox chk 
               Caption         =   "��Ŀ�����������շѻ�������"
               Height          =   210
               Index           =   90
               Left            =   180
               TabIndex        =   653
               Top             =   675
               Width           =   3120
            End
         End
         Begin VB.Frame fraCharge 
            Caption         =   "����ȷ�����շ�Ʊ������"
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   270
            TabIndex        =   649
            Top             =   3015
            Width           =   4215
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   8
               ItemData        =   "frmParFee.frx":1A67F
               Left            =   1335
               List            =   "frmParFee.frx":1A681
               Style           =   2  'Dropdown List
               TabIndex        =   654
               Top             =   390
               Width           =   2460
            End
            Begin VB.Label lblPrintFormat 
               AutoSize        =   -1  'True
               Caption         =   "Ʊ�ݴ�ӡ��ʽ"
               Height          =   180
               Left            =   150
               TabIndex        =   650
               Top             =   450
               Width           =   1080
            End
         End
         Begin VB.Frame fraSetMoneyMode 
            Caption         =   "�����շѡ�����ˢ��ȱʡ����������������"
            Height          =   1260
            Left            =   270
            TabIndex        =   648
            Top             =   5010
            Width           =   4215
            Begin VB.OptionButton optSetMoneyMode 
               Caption         =   "��ȱʡˢ�����"
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   656
               Top             =   345
               Value           =   -1  'True
               Width           =   2640
            End
            Begin VB.OptionButton optSetMoneyMode 
               Caption         =   "ȱʡˢ������ҽ��������"
               Height          =   210
               Index           =   2
               Left            =   210
               TabIndex        =   658
               Top             =   975
               Width           =   3120
            End
            Begin VB.OptionButton optSetMoneyMode 
               Caption         =   "ȱʡˢ������ҽ���������"
               Height          =   210
               Index           =   1
               Left            =   210
               TabIndex        =   657
               Top             =   660
               Width           =   2760
            End
         End
         Begin VB.Frame fra 
            Caption         =   "�����ӿ�ҩ������"
            Height          =   4530
            Index           =   16
            Left            =   4770
            TabIndex        =   647
            Top             =   1860
            Width           =   4320
            Begin TabDlg.SSTab stabDrug 
               Height          =   4170
               Left            =   120
               TabIndex        =   662
               Top             =   270
               Width           =   4110
               _ExtentX        =   7250
               _ExtentY        =   7355
               _Version        =   393216
               Style           =   1
               TabHeight       =   520
               TabCaption(0)   =   "��ҩ��"
               TabPicture(0)   =   "frmParFee.frx":1A683
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "vsfDrugStore(0)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "��ҩ��"
               TabPicture(1)   =   "frmParFee.frx":1A69F
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "vsfDrugStore(1)"
               Tab(1).ControlCount=   1
               TabCaption(2)   =   "��ҩ��"
               TabPicture(2)   =   "frmParFee.frx":1A6BB
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "vsfDrugStore(2)"
               Tab(2).ControlCount=   1
               Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
                  Height          =   3780
                  Index           =   0
                  Left            =   60
                  TabIndex        =   663
                  Top             =   330
                  Width           =   4020
                  _cx             =   7091
                  _cy             =   6667
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
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   0
                  AllowSelection  =   0   'False
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   4
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmParFee.frx":1A6D7
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
               Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
                  Height          =   3780
                  Index           =   1
                  Left            =   -74940
                  TabIndex        =   664
                  Top             =   330
                  Width           =   3990
                  _cx             =   7038
                  _cy             =   6667
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
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   0
                  AllowSelection  =   0   'False
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   4
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmParFee.frx":1A74B
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
               Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
                  Height          =   3780
                  Index           =   2
                  Left            =   -74940
                  TabIndex        =   665
                  Top             =   330
                  Width           =   3990
                  _cx             =   7038
                  _cy             =   6667
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
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   0
                  AllowSelection  =   0   'False
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   4
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmParFee.frx":1A7BF
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
            End
         End
         Begin VB.Frame fraRecored 
            Caption         =   "���Ѽ�����˺����Ʊ������"
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   270
            TabIndex        =   171
            Top             =   4005
            Width           =   4215
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   9
               Left            =   1335
               Style           =   2  'Dropdown List
               TabIndex        =   655
               Top             =   345
               Width           =   2460
            End
            Begin VB.Label lblRecordPrint 
               AutoSize        =   -1  'True
               Caption         =   "Ʊ�ݴ�ӡ��ʽ"
               Height          =   180
               Left            =   210
               TabIndex        =   172
               Top             =   405
               Width           =   1080
            End
         End
         Begin VB.CommandButton cmdOneCard 
            Height          =   345
            Index           =   0
            Left            =   8925
            Picture         =   "frmParFee.frx":1A833
            Style           =   1  'Graphical
            TabIndex        =   159
            Top             =   405
            Width           =   345
         End
         Begin VB.CommandButton cmdOneCard 
            Enabled         =   0   'False
            Height          =   345
            Index           =   1
            Left            =   8925
            Picture         =   "frmParFee.frx":1ADBD
            Style           =   1  'Graphical
            TabIndex        =   161
            Top             =   795
            Width           =   345
         End
         Begin VB.CommandButton cmdOneCard 
            Enabled         =   0   'False
            Height          =   345
            Index           =   2
            Left            =   8925
            Picture         =   "frmParFee.frx":1B347
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   1200
            Width           =   345
         End
         Begin MSComctlLib.ListView lvw 
            Height          =   1305
            Index           =   3
            Left            =   240
            TabIndex        =   157
            Top             =   390
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   2302
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "NO"
               Text            =   "���"
               Object.Width           =   970
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "Name"
               Text            =   "����"
               Object.Width           =   7410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   "PayType"
               Text            =   "���㷽ʽ"
               Object.Width           =   2998
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "OrgCode"
               Text            =   "ҽԺ����"
               Object.Width           =   1677
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "Enable"
               Text            =   "����"
               Object.Width           =   970
            EndProperty
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "һ��ͨ�ӿ�(�ϰ�)"
            Height          =   180
            Index           =   45
            Left            =   285
            TabIndex        =   156
            Top             =   120
            Width           =   1440
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   8115
         Index           =   6
         Left            =   -74865
         ScaleHeight     =   8085
         ScaleWidth      =   8985
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   615
         Width           =   9015
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   975
            Index           =   0
            Left            =   255
            TabIndex        =   324
            TabStop         =   0   'False
            Top             =   60
            Width           =   5025
            _Version        =   589884
            _ExtentX        =   8864
            _ExtentY        =   1720
            _StockProps     =   64
         End
         Begin VB.PictureBox picOtherRegister 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   7050
            Left            =   345
            ScaleHeight     =   7020
            ScaleWidth      =   8970
            TabIndex        =   294
            TabStop         =   0   'False
            Top             =   0
            Width           =   9000
            Begin VB.Frame fraOrder 
               Caption         =   "ҽ��վ�Һ��������"
               Height          =   2520
               Index           =   1
               Left            =   4050
               TabIndex        =   320
               Top             =   4080
               Visible         =   0   'False
               Width           =   4410
               Begin VB.CommandButton cmdStationRegOrder 
                  Caption         =   "��"
                  Height          =   510
                  Index           =   0
                  Left            =   3720
                  TabIndex        =   322
                  Top             =   825
                  Width           =   375
               End
               Begin VB.CommandButton cmdStationRegOrder 
                  Caption         =   "��"
                  Height          =   510
                  Index           =   1
                  Left            =   3720
                  TabIndex        =   323
                  Top             =   1425
                  Width           =   375
               End
               Begin VSFlex8Ctl.VSFlexGrid vsStationRegSort 
                  Height          =   1995
                  Left            =   240
                  TabIndex        =   321
                  Top             =   360
                  Width           =   3330
                  _cx             =   5874
                  _cy             =   3519
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
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483634
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   2
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   2
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   6
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   300
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":1B8D1
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   4
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
                  ExplorerBar     =   8
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
            Begin VB.CheckBox chk 
               Caption         =   "ҽ��վԤԼ�������Ұ���"
               Height          =   330
               Index           =   184
               Left            =   90
               TabIndex        =   752
               Top             =   1980
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "ҽ��վ��δ����ҽ���ĺű�ʱ��������ҽ��"
               Height          =   195
               Index           =   169
               Left            =   90
               TabIndex        =   303
               Top             =   2655
               Width           =   3870
            End
            Begin VB.Frame fraSlip 
               Caption         =   "�Һ�ƾ����ӡ"
               Height          =   765
               Left            =   4050
               TabIndex        =   316
               Top             =   3165
               Width           =   4410
               Begin VB.OptionButton optPrintSlip 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   318
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.OptionButton optPrintSlip 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   317
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   900
               End
               Begin VB.OptionButton optPrintSlip 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   2670
                  TabIndex        =   319
                  Top             =   360
                  Width           =   1380
               End
            End
            Begin VB.Frame fraAppoint 
               Caption         =   "ԤԼ�Һŵ���ӡ"
               Height          =   765
               Left            =   4050
               TabIndex        =   312
               Top             =   2220
               Width           =   4410
               Begin VB.OptionButton optPrintAppoint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   2670
                  TabIndex        =   315
                  Top             =   375
                  Width           =   1380
               End
               Begin VB.OptionButton optPrintAppoint 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   313
                  Top             =   375
                  Value           =   -1  'True
                  Width           =   900
               End
               Begin VB.OptionButton optPrintAppoint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   314
                  Top             =   375
                  Width           =   1020
               End
            End
            Begin VB.Frame fraInvoice 
               Caption         =   "�Һ�Ʊ�ݴ�ӡ"
               Height          =   735
               Left            =   4050
               TabIndex        =   308
               Top             =   1380
               Width           =   4410
               Begin VB.OptionButton optPrintFact 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   2670
                  TabIndex        =   311
                  Top             =   390
                  Width           =   1380
               End
               Begin VB.OptionButton optPrintFact 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   309
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   900
               End
               Begin VB.OptionButton optPrintFact 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   310
                  Top             =   390
                  Width           =   1020
               End
            End
            Begin VB.Frame fraRegistMode 
               Caption         =   "�Һŷ�֧��ģʽ"
               Height          =   1155
               Left            =   4050
               TabIndex        =   304
               Top             =   150
               Width           =   4410
               Begin VB.OptionButton optRegist 
                  Caption         =   "����֧���򴰿�֧��ģʽ"
                  Height          =   225
                  Index           =   2
                  Left            =   300
                  TabIndex        =   307
                  Top             =   675
                  Width           =   3195
               End
               Begin VB.OptionButton optRegist 
                  Caption         =   "����֧��ģʽ"
                  Height          =   225
                  Index           =   1
                  Left            =   1920
                  TabIndex        =   306
                  Top             =   360
                  Width           =   1425
               End
               Begin VB.OptionButton optRegist 
                  Caption         =   "����֧��ģʽ"
                  Height          =   225
                  Index           =   0
                  Left            =   300
                  TabIndex        =   305
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1425
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һű���ˢ��"
               Height          =   330
               Index           =   102
               Left            =   90
               TabIndex        =   295
               Top             =   135
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һ�ʱ�ɷ�����ʹ��Ԥ����"
               Height          =   330
               Index           =   103
               Left            =   90
               TabIndex        =   296
               Top             =   450
               Width           =   2850
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   180
               Index           =   106
               Left            =   1815
               TabIndex        =   300
               Text            =   "7"
               Top             =   1425
               Width           =   360
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һ�ԤԼʱ�տ�"
               Height          =   330
               Index           =   108
               Left            =   90
               TabIndex        =   302
               Top             =   2310
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "ҽ��վ�ҺŰ������Ұ���"
               Height          =   330
               Index           =   107
               Left            =   90
               TabIndex        =   301
               Top             =   1680
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "������ѡ��"
               Height          =   330
               Index           =   105
               Left            =   90
               TabIndex        =   298
               Top             =   1080
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "����סԺ���˽��йҺ�"
               Height          =   330
               Index           =   104
               Left            =   90
               TabIndex        =   297
               Top             =   765
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "��������ģ������     ���ڵĲ���"
               Height          =   180
               Index           =   106
               Left            =   90
               TabIndex        =   299
               Top             =   1425
               Width           =   3135
            End
            Begin VB.Line lnEpr 
               Index           =   1
               X1              =   1800
               X2              =   2190
               Y1              =   1620
               Y2              =   1620
            End
         End
         Begin VB.PictureBox picRegist 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6705
            Left            =   15
            ScaleHeight     =   6675
            ScaleWidth      =   9165
            TabIndex        =   206
            TabStop         =   0   'False
            Top             =   30
            Width           =   9195
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   25
               Left            =   2280
               MaxLength       =   2
               TabIndex        =   234
               Text            =   "0"
               Top             =   6660
               Width           =   660
            End
            Begin VB.Frame fra�ɿ�������� 
               Caption         =   "�ҺŽɿ��������"
               Height          =   1050
               Left            =   4335
               TabIndex        =   778
               Top             =   3705
               Width           =   4500
               Begin VB.OptionButton optMoneyControl 
                  Caption         =   "�����п���"
                  Height          =   180
                  Index           =   0
                  Left            =   225
                  TabIndex        =   781
                  Top             =   255
                  Width           =   1260
               End
               Begin VB.OptionButton optMoneyControl 
                  Caption         =   "����ɿ���֮��Ž������ιҺ��շ�"
                  Height          =   180
                  Index           =   1
                  Left            =   225
                  TabIndex        =   780
                  Top             =   510
                  Width           =   3825
               End
               Begin VB.OptionButton optMoneyControl 
                  Caption         =   "��������ɿ���"
                  Height          =   180
                  Index           =   2
                  Left            =   225
                  TabIndex        =   779
                  Top             =   780
                  Width           =   3825
               End
            End
            Begin VB.PictureBox pic��ǰ��ɫ 
               BackColor       =   &H00000000&
               Height          =   270
               Left            =   2010
               ScaleHeight     =   210
               ScaleWidth      =   210
               TabIndex        =   758
               Top             =   5565
               Width           =   270
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һź�ԤԼʱ��ֹ��������"
               Height          =   300
               Index           =   190
               Left            =   135
               TabIndex        =   755
               Top             =   5280
               Width           =   2700
            End
            Begin VB.CheckBox chk 
               Caption         =   "�ƻ��Ű�ģʽ�������"
               Height          =   300
               Index           =   185
               Left            =   135
               TabIndex        =   753
               Top             =   4995
               Width           =   2700
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ͬ�ƹҺ��������ڼ���"
               Height          =   210
               Index           =   175
               Left            =   405
               TabIndex        =   226
               Top             =   4455
               Width           =   3735
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   23
               Left            =   1890
               MaxLength       =   2
               TabIndex        =   225
               Text            =   "0"
               Top             =   4185
               Width           =   660
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   20
               Left            =   2100
               MaxLength       =   2
               TabIndex        =   228
               Text            =   "0"
               Top             =   4755
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "��δ����ҽ���ĺű�ʱ��������ҽ��"
               Height          =   195
               Index           =   27
               Left            =   135
               TabIndex        =   212
               Top             =   1230
               Width           =   4080
            End
            Begin VB.Frame fraLine 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Left            =   840
               TabIndex        =   268
               Top             =   345
               Width           =   285
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   180
               Index           =   2
               Left            =   840
               MaxLength       =   2
               TabIndex        =   208
               Text            =   "5"
               Top             =   150
               Width           =   285
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һű��뽨����ʱ�Զ��������������"
               Height          =   195
               Index           =   23
               Left            =   135
               TabIndex        =   209
               Top             =   450
               Width           =   3360
            End
            Begin VB.CheckBox chk 
               Caption         =   "�������˹ҺŴ�Ϊ���۵�"
               Height          =   240
               Index           =   24
               Left            =   135
               TabIndex        =   211
               Top             =   960
               Width           =   4005
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ʹ��Ԥ����ɷ�"
               Height          =   195
               Index           =   30
               Left            =   135
               TabIndex        =   213
               Top             =   1485
               Width           =   2340
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   180
               Index           =   3
               Left            =   2025
               MaxLength       =   3
               TabIndex        =   215
               Text            =   "0"
               Top             =   1740
               Width           =   285
            End
            Begin VB.Frame fraLine2 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Left            =   2055
               TabIndex        =   267
               Top             =   1935
               Width           =   285
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һŷ���Ϊ��ʱҲ��ӡƱ��"
               Height          =   255
               Index           =   32
               Left            =   135
               TabIndex        =   216
               Top             =   1980
               Width           =   2460
            End
            Begin VB.Frame fraInput 
               Caption         =   "Ҫ��������"
               Height          =   1050
               Left            =   4335
               TabIndex        =   235
               Top             =   135
               Width           =   4485
               Begin VB.CheckBox chk 
                  Caption         =   "��ϵ�绰"
                  Height          =   195
                  Index           =   20
                  Left            =   1440
                  TabIndex        =   243
                  Top             =   765
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����"
                  Height          =   195
                  Index           =   33
                  Left            =   240
                  TabIndex        =   236
                  Top             =   285
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�Ա�"
                  Height          =   195
                  Index           =   34
                  Left            =   1440
                  TabIndex        =   237
                  Top             =   285
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����"
                  Height          =   195
                  Index           =   35
                  Left            =   2745
                  TabIndex        =   238
                  Top             =   285
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�ѱ�"
                  Height          =   195
                  Index           =   38
                  Left            =   240
                  TabIndex        =   239
                  Top             =   525
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���㷽ʽ"
                  Height          =   195
                  Index           =   40
                  Left            =   2745
                  TabIndex        =   241
                  Top             =   525
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���ʽ"
                  Height          =   195
                  Index           =   37
                  Left            =   1440
                  TabIndex        =   240
                  Top             =   525
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��ͥ��ַ"
                  Height          =   195
                  Index           =   36
                  Left            =   240
                  TabIndex        =   242
                  Top             =   765
                  Width           =   1020
               End
            End
            Begin VB.Frame fraRegistBillMode 
               Caption         =   "�Һ�Ʊ��"
               Height          =   510
               Left            =   4335
               TabIndex        =   244
               Top             =   1260
               Width           =   4500
               Begin VB.OptionButton optRegistPrintMode 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   2775
                  TabIndex        =   247
                  Top             =   240
                  Width           =   1380
               End
               Begin VB.OptionButton optRegistPrintMode 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   245
                  Top             =   240
                  Width           =   900
               End
               Begin VB.OptionButton optRegistPrintMode 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   246
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1020
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һź��ӡ������ǩ"
               Height          =   255
               Index           =   45
               Left            =   135
               TabIndex        =   217
               Top             =   2250
               Width           =   1980
            End
            Begin VB.CheckBox chk 
               Caption         =   "����סԺ���˹Һ�"
               Height          =   255
               Index           =   49
               Left            =   135
               TabIndex        =   218
               Top             =   2520
               Width           =   1845
            End
            Begin VB.ComboBox cbo 
               ForeColor       =   &H80000012&
               Height          =   300
               Index           =   2
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   230
               Top             =   5910
               Width           =   1800
            End
            Begin VB.CheckBox chk 
               Caption         =   "�������Ա���ѡ��ҺŰ������"
               Height          =   255
               Index           =   51
               Left            =   135
               TabIndex        =   219
               Top             =   2790
               Width           =   3090
            End
            Begin VB.Frame fraBarCodePrint 
               Caption         =   "���������ӡ"
               ForeColor       =   &H00000000&
               Height          =   510
               Left            =   4335
               TabIndex        =   248
               Top             =   1845
               Width           =   4500
               Begin VB.OptionButton optBarCodePrint 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   249
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   900
               End
               Begin VB.OptionButton optBarCodePrint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   250
                  Top             =   225
                  Width           =   1020
               End
               Begin VB.OptionButton optBarCodePrint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   2775
                  TabIndex        =   251
                  Top             =   225
                  Width           =   1380
               End
            End
            Begin VB.Frame fraClearMZInfor 
               Caption         =   "�˺�����������Ϣ(�Һ���Ч�����ڵĲ���)"
               Height          =   615
               Left            =   4335
               TabIndex        =   256
               Top             =   3030
               Width           =   4500
               Begin VB.OptionButton optRegistClearMzInfor 
                  Caption         =   "��ʾ���"
                  Height          =   180
                  Index           =   2
                  Left            =   2775
                  TabIndex        =   259
                  Top             =   285
                  Width           =   1110
               End
               Begin VB.OptionButton optRegistClearMzInfor 
                  Caption         =   "�Զ����"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   258
                  Top             =   285
                  Width           =   1110
               End
               Begin VB.OptionButton optRegistClearMzInfor 
                  Caption         =   "�����"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   257
                  Top             =   285
                  Width           =   1110
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ͥ��ַ��������"
               Height          =   255
               Index           =   54
               Left            =   135
               TabIndex        =   220
               Top             =   3075
               Value           =   1  'Checked
               Width           =   1980
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   10
               Left            =   240
               MaxLength       =   2
               TabIndex        =   231
               Text            =   "0"
               Top             =   6285
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����������Һ�"
               Height          =   300
               Index           =   60
               Left            =   135
               TabIndex        =   222
               Top             =   3600
               Width           =   2550
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʱ�κű��ϸ�ʱ�ιҺ�"
               Height          =   300
               Index           =   61
               Left            =   135
               TabIndex        =   221
               Top             =   3330
               Width           =   2550
            End
            Begin VB.Frame fraSlipPrint 
               Caption         =   "�Һ�ƾ��"
               ForeColor       =   &H00000000&
               Height          =   510
               Left            =   4335
               TabIndex        =   252
               Top             =   2445
               Width           =   4500
               Begin VB.OptionButton optSlipPrint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   2775
                  TabIndex        =   255
                  Top             =   240
                  Width           =   1380
               End
               Begin VB.OptionButton optSlipPrint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   254
                  Top             =   240
                  Width           =   1020
               End
               Begin VB.OptionButton optSlipPrint 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   253
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   900
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һ�ʱĬ�Ϲ�ѡ������ѡ��"
               Height          =   300
               Index           =   62
               Left            =   135
               TabIndex        =   223
               Top             =   3870
               Width           =   2700
            End
            Begin VB.CheckBox chk 
               Caption         =   "�������Ч�Լ��"
               Height          =   255
               Index           =   63
               Left            =   135
               TabIndex        =   210
               Top             =   690
               Width           =   2325
            End
            Begin VB.Frame fraRegistCards 
               Caption         =   "�Һŷ������"
               Height          =   2145
               Left            =   4335
               TabIndex        =   260
               Top             =   4845
               Width           =   4500
               Begin VB.CheckBox chk 
                  Caption         =   "������Һŷ�һ����(���򿨷Ѵ�Ϊ���۵�)"
                  Height          =   195
                  Index           =   26
                  Left            =   165
                  TabIndex        =   261
                  Top             =   315
                  Width           =   3960
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�����²����Զ�������ʱ����"
                  Height          =   195
                  Index           =   29
                  Left            =   165
                  TabIndex        =   262
                  Top             =   555
                  Width           =   3345
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�˺Ų��˿�ʱ�ش�Ʊ��"
                  Height          =   195
                  Index           =   44
                  Left            =   165
                  TabIndex        =   263
                  Top             =   825
                  Width           =   2400
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����������������Ϣ�ǼǴ���"
                  Height          =   195
                  Index           =   42
                  Left            =   165
                  TabIndex        =   264
                  Top             =   1110
                  Width           =   3345
               End
               Begin VB.CheckBox chk 
                  Caption         =   "ɨ�����֤ǩԼ"
                  Height          =   195
                  Index           =   58
                  Left            =   165
                  TabIndex        =   265
                  Top             =   1395
                  Width           =   3345
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���ϸ���ƿ�ʱʼ��Ϊ����"
                  Height          =   195
                  Index           =   64
                  Left            =   165
                  TabIndex        =   266
                  Top             =   1695
                  Width           =   3345
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "ÿ��     �����Զ�ˢ�¹ҺŰ��ű�"
               Height          =   195
               Index           =   22
               Left            =   135
               TabIndex        =   207
               Top             =   150
               Width           =   3480
            End
            Begin VB.CheckBox chk 
               Caption         =   "����������ģ������    ���ڵĲ���"
               Height          =   195
               Index           =   31
               Left            =   135
               TabIndex        =   214
               Top             =   1740
               Width           =   3300
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ͬһ�����޹�        ����"
               Height          =   210
               Index           =   174
               Left            =   135
               TabIndex        =   224
               Top             =   4170
               Width           =   3735
            End
            Begin VB.CheckBox chk 
               Caption         =   "ͬһ��������ܹҺ�        ������"
               Height          =   210
               Index           =   176
               Left            =   135
               TabIndex        =   227
               Top             =   4740
               Width           =   3705
            End
            Begin VB.CheckBox chk 
               Caption         =   "ͬһ����ͬһ��Դ�޹�         ����"
               Height          =   210
               Index           =   203
               Left            =   135
               TabIndex        =   233
               Top             =   6660
               Width           =   3465
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   2280
               X2              =   2880
               Y1              =   6900
               Y2              =   6900
            End
            Begin VB.Label lblColor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�°���ǰ�ҺŰ�����ɫ"
               Height          =   180
               Left            =   135
               TabIndex        =   759
               Top             =   5610
               Width           =   1800
            End
            Begin VB.Line Line1 
               Index           =   10
               X1              =   1845
               X2              =   2580
               Y1              =   4380
               Y2              =   4380
            End
            Begin VB.Line Line1 
               Index           =   9
               X1              =   2055
               X2              =   2775
               Y1              =   4950
               Y2              =   4950
            End
            Begin VB.Label lblSortMode 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ʽ"
               Height          =   180
               Left            =   135
               TabIndex        =   229
               Top             =   5970
               Width           =   720
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   240
               X2              =   960
               Y1              =   6525
               Y2              =   6525
            End
            Begin VB.Label lblGuardian 
               AutoSize        =   -1  'True
               Caption         =   "�����±���¼��໤��"
               Height          =   180
               Left            =   960
               TabIndex        =   232
               Top             =   6285
               Width           =   1800
            End
         End
         Begin VB.PictureBox picRegistPlan 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6810
            Left            =   840
            ScaleHeight     =   6780
            ScaleWidth      =   8595
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   30
            Width           =   8625
            Begin VB.Frame fraNewPaln 
               BorderStyle     =   0  'None
               Height          =   3855
               Left            =   150
               TabIndex        =   682
               Top             =   1560
               Visible         =   0   'False
               Width           =   5145
               Begin VB.ComboBox cbo 
                  Height          =   300
                  Index           =   18
                  Left            =   2160
                  Style           =   2  'Dropdown List
                  TabIndex        =   195
                  Top             =   930
                  Width           =   1695
               End
               Begin VB.ComboBox cbo 
                  Height          =   300
                  Index           =   19
                  Left            =   2160
                  Style           =   2  'Dropdown List
                  TabIndex        =   196
                  Top             =   1290
                  Width           =   1695
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����ҽ��ְ�񼶱���"
                  Height          =   195
                  Index           =   182
                  Left            =   2940
                  TabIndex        =   193
                  Top             =   30
                  Width           =   2115
               End
               Begin VB.CheckBox chk 
                  Caption         =   "������ҽ��ͬ������ԤԼ�Һŵ�"
                  Height          =   195
                  Index           =   180
                  Left            =   0
                  TabIndex        =   194
                  Top             =   420
                  Width           =   2835
               End
               Begin VB.Frame fraVisitTablePrintMode 
                  Caption         =   "������ӡ��ʽ"
                  Height          =   735
                  Left            =   0
                  TabIndex        =   685
                  Top             =   3150
                  Width           =   5055
                  Begin VB.OptionButton optVisitTablePrintMode 
                     Caption         =   "����ӡ"
                     Height          =   180
                     Index           =   0
                     Left            =   300
                     TabIndex        =   203
                     Top             =   360
                     Value           =   -1  'True
                     Width           =   855
                  End
                  Begin VB.OptionButton optVisitTablePrintMode 
                     Caption         =   "�Զ���ӡ"
                     Height          =   180
                     Index           =   1
                     Left            =   1680
                     TabIndex        =   204
                     Top             =   360
                     Width           =   1035
                  End
                  Begin VB.OptionButton optVisitTablePrintMode 
                     Caption         =   "ѡ���Ƿ��ӡ"
                     Height          =   180
                     Index           =   2
                     Left            =   3090
                     TabIndex        =   205
                     Top             =   360
                     Width           =   1395
                  End
               End
               Begin VB.Frame fraPrintMode 
                  Caption         =   "ԤԼ�嵥��ӡ��ʽ"
                  Height          =   1305
                  Left            =   2940
                  TabIndex        =   684
                  Top             =   1710
                  Width           =   2115
                  Begin VB.OptionButton optPrintMode 
                     Caption         =   "ѡ���Ƿ��ӡ"
                     Height          =   180
                     Index           =   2
                     Left            =   300
                     TabIndex        =   202
                     Top             =   930
                     Width           =   1395
                  End
                  Begin VB.OptionButton optPrintMode 
                     Caption         =   "�Զ���ӡ"
                     Height          =   180
                     Index           =   1
                     Left            =   300
                     TabIndex        =   201
                     Top             =   615
                     Width           =   1035
                  End
                  Begin VB.OptionButton optPrintMode 
                     Caption         =   "����ӡ"
                     Height          =   180
                     Index           =   0
                     Left            =   300
                     TabIndex        =   200
                     Top             =   300
                     Value           =   -1  'True
                     Width           =   855
                  End
               End
               Begin VB.Frame fraToExcelMode 
                  Caption         =   "ԤԼ�嵥���Ʒ�ʽ"
                  Height          =   1305
                  Left            =   0
                  TabIndex        =   683
                  Top             =   1710
                  Width           =   2715
                  Begin VB.OptionButton optToExcelMode 
                     Caption         =   "ѡ���Ƿ������Excel"
                     Height          =   225
                     Index           =   2
                     Left            =   300
                     TabIndex        =   199
                     Top             =   930
                     Width           =   2025
                  End
                  Begin VB.OptionButton optToExcelMode 
                     Caption         =   "�Զ������Excel"
                     Height          =   225
                     Index           =   1
                     Left            =   300
                     TabIndex        =   198
                     Top             =   615
                     Width           =   1665
                  End
                  Begin VB.OptionButton optToExcelMode 
                     Caption         =   "�������Excel"
                     Height          =   225
                     Index           =   0
                     Left            =   300
                     TabIndex        =   197
                     Top             =   300
                     Value           =   -1  'True
                     Width           =   1485
                  End
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�����ڶ�Ժ��ҽ�����йҺŰ���"
                  Height          =   180
                  Index           =   181
                  Left            =   0
                  TabIndex        =   192
                  Top             =   30
                  Width           =   2820
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "��δ����վ��ĺ�Դ�����                   ���г��ﰲ��"
                  Height          =   180
                  Index           =   5
                  Left            =   0
                  TabIndex        =   790
                  Top             =   990
                  Width           =   4950
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "����ʱ��Դ����ıȽϷ�ʽ"
                  Height          =   180
                  Index           =   8
                  Left            =   0
                  TabIndex        =   789
                  Top             =   1350
                  Width           =   2160
               End
            End
            Begin VB.Frame fraRegistPlanMode 
               Caption         =   "�Ű�ģʽ"
               Height          =   690
               Index           =   5
               Left            =   75
               TabIndex        =   189
               Top             =   135
               Width           =   8400
               Begin VB.OptionButton optRegistPlanMode 
                  Caption         =   "�ƻ��Ű�ģʽ"
                  Height          =   285
                  Index           =   0
                  Left            =   255
                  TabIndex        =   178
                  Top             =   240
                  Width           =   1590
               End
               Begin VB.OptionButton optRegistPlanMode 
                  Caption         =   "������Ű�ģʽ"
                  Height          =   285
                  Index           =   1
                  Left            =   1935
                  TabIndex        =   179
                  Top             =   240
                  Width           =   1980
               End
               Begin MSComCtl2.DTPicker dtpRegistPlanMode 
                  Height          =   300
                  Left            =   4710
                  TabIndex        =   180
                  Top             =   225
                  Width           =   2070
                  _ExtentX        =   3651
                  _ExtentY        =   529
                  _Version        =   393216
                  CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
                  Format          =   166789123
                  CurrentDate     =   42481
               End
               Begin VB.CommandButton cmdInstantActive 
                  Caption         =   "�������ó�����Ű�ģʽ"
                  Height          =   480
                  Left            =   6900
                  TabIndex        =   181
                  Top             =   135
                  Width           =   1380
               End
               Begin VB.Label lblRegistPlanMode 
                  AutoSize        =   -1  'True
                  Caption         =   "��������"
                  Height          =   180
                  Left            =   3945
                  TabIndex        =   675
                  Top             =   285
                  Width           =   720
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ԤԼ�Һŵ���ֹɾ������"
               Height          =   240
               Index           =   21
               Left            =   150
               TabIndex        =   191
               Top             =   1245
               Width           =   2655
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����ڶ�Ժ��ҽ�����йҺŰ���"
               Height          =   270
               Index           =   17
               Left            =   150
               TabIndex        =   190
               Top             =   945
               Width           =   2820
            End
         End
         Begin VB.PictureBox picԤԼ 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   7245
            Left            =   645
            ScaleHeight     =   7215
            ScaleWidth      =   9105
            TabIndex        =   269
            TabStop         =   0   'False
            Top             =   -60
            Width           =   9135
            Begin VB.ComboBox cbo 
               ForeColor       =   &H80000012&
               Height          =   300
               Index           =   10
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   750
               Top             =   2865
               Width           =   1350
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   195
               Index           =   22
               Left            =   1815
               MaxLength       =   2
               TabIndex        =   274
               Text            =   "0"
               Top             =   705
               Width           =   660
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   195
               Index           =   21
               Left            =   2025
               MaxLength       =   2
               TabIndex        =   272
               Text            =   "0"
               Top             =   450
               Width           =   660
            End
            Begin VB.Frame fraReceiveMode 
               Caption         =   "ԤԼ����ģʽ"
               Height          =   615
               Left            =   120
               TabIndex        =   644
               Top             =   3990
               Width           =   6210
               Begin VB.OptionButton optReceiveMode 
                  Caption         =   "��ԤԼ����"
                  Height          =   255
                  Index           =   1
                  Left            =   2715
                  TabIndex        =   646
                  Top             =   255
                  Width           =   3165
               End
               Begin VB.OptionButton optReceiveMode 
                  Caption         =   "ԤԼ���վ���"
                  Height          =   255
                  Index           =   0
                  Left            =   285
                  TabIndex        =   645
                  Top             =   255
                  Width           =   2130
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "ԤԼ��ʾ���к���"
               Height          =   195
               Index           =   46
               Left            =   75
               TabIndex        =   270
               Top             =   195
               Width           =   1935
            End
            Begin VB.CheckBox chk 
               Caption         =   "�Һŷ�����ԤԼ����ʱ��Ϊ׼!"
               Height          =   210
               Index           =   48
               Left            =   75
               TabIndex        =   275
               Top             =   975
               Width           =   3375
            End
            Begin VB.CheckBox chk 
               Caption         =   "ԤԼʱ�����������"
               Height          =   285
               Index           =   50
               Left            =   75
               TabIndex        =   276
               Top             =   1185
               Width           =   2655
            End
            Begin VB.Frame fra 
               BorderStyle     =   0  'None
               Caption         =   "Ĭ������ؿ���"
               Height          =   1350
               Index           =   0
               Left            =   135
               TabIndex        =   289
               Top             =   4545
               Width           =   8850
               Begin VB.ComboBox cbo 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  Index           =   11
                  Left            =   2955
                  Style           =   2  'Dropdown List
                  TabIndex        =   754
                  Top             =   555
                  Width           =   780
               End
               Begin VB.Frame fraSplitregister 
                  Height          =   45
                  Left            =   1335
                  TabIndex        =   325
                  Top             =   270
                  Width           =   8475
               End
               Begin VB.TextBox txt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   6
                  Left            =   1335
                  MaxLength       =   4
                  TabIndex        =   293
                  Text            =   "0"
                  Top             =   945
                  Width           =   660
               End
               Begin VB.TextBox txt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   5
                  Left            =   3750
                  MaxLength       =   4
                  TabIndex        =   291
                  Text            =   "0"
                  Top             =   615
                  Width           =   660
               End
               Begin VB.Label lblAvailabilityTimes 
                  AutoSize        =   -1  'True
                  Caption         =   "ԤԼ��Чʱ�䣺ԤԼ����ԤԼʱ��                  ����δ���յ�ΪʧԼ��"
                  Height          =   180
                  Left            =   225
                  TabIndex        =   290
                  Top             =   615
                  Width           =   6120
               End
               Begin VB.Label lblRegisterCtl 
                  AutoSize        =   -1  'True
                  Caption         =   "��������ؿ���"
                  Height          =   180
                  Left            =   15
                  TabIndex        =   326
                  Top             =   210
                  Width           =   1260
               End
               Begin VB.Label lblBreakAnAppointmentNums 
                  AutoSize        =   -1  'True
                  Caption         =   "����ԤԼʧԼ         ���Զ����������"
                  Height          =   180
                  Left            =   210
                  TabIndex        =   292
                  Top             =   945
                  Width           =   3330
               End
               Begin VB.Line Line1 
                  Index           =   0
                  X1              =   1350
                  X2              =   2070
                  Y1              =   1185
                  Y2              =   1185
               End
               Begin VB.Line Line1 
                  Index           =   2
                  X1              =   3780
                  X2              =   4500
                  Y1              =   855
                  Y2              =   855
               End
            End
            Begin VB.TextBox txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   195
               Index           =   7
               Left            =   660
               MaxLength       =   2
               TabIndex        =   279
               Text            =   "0"
               Top             =   1800
               Width           =   660
            End
            Begin VB.CheckBox chk 
               Caption         =   "�˺����:N����ȡ��ԤԼ��Ҫͨ�����"
               Height          =   210
               Index           =   52
               Left            =   3150
               TabIndex        =   280
               Top             =   1785
               Width           =   3735
            End
            Begin VB.CheckBox chk 
               Caption         =   "ԤԼʧԼ���ڹҺ�"
               Height          =   210
               Index           =   53
               Left            =   75
               TabIndex        =   277
               Top             =   1500
               Width           =   2055
            End
            Begin VB.TextBox txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   8
               Left            =   1275
               MaxLength       =   4
               TabIndex        =   288
               Text            =   "0"
               Top             =   2505
               Width           =   540
            End
            Begin VB.TextBox txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   9
               Left            =   1185
               MaxLength       =   4
               TabIndex        =   286
               Text            =   "0"
               Top             =   2175
               Width           =   540
            End
            Begin VB.Frame fraBespeak 
               Caption         =   "ԤԼ�Һŵ���ӡ"
               Height          =   615
               Left            =   120
               TabIndex        =   281
               Top             =   3270
               Width           =   6210
               Begin VB.OptionButton optPrintBespeak 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   2715
                  TabIndex        =   284
                  Top             =   255
                  Width           =   1380
               End
               Begin VB.OptionButton optPrintBespeak 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   285
                  TabIndex        =   282
                  Top             =   255
                  Width           =   900
               End
               Begin VB.OptionButton optPrintBespeak 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   1395
                  TabIndex        =   283
                  Top             =   255
                  Value           =   -1  'True
                  Width           =   1020
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "ͬһ���������ԤԼ        ������"
               Height          =   210
               Index           =   47
               Left            =   75
               TabIndex        =   271
               Top             =   450
               Width           =   3705
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ͬһ������Լ        ����"
               Height          =   210
               Index           =   172
               Left            =   75
               TabIndex        =   273
               Top             =   705
               Width           =   3735
            End
            Begin VB.Label lblAppStyle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ȱʡԤԼ��ʽ"
               Height          =   180
               Left            =   75
               TabIndex        =   751
               Top             =   2925
               Width           =   1080
            End
            Begin VB.Line Line1 
               Index           =   8
               X1              =   1785
               X2              =   2505
               Y1              =   915
               Y2              =   915
            End
            Begin VB.Line Line1 
               Index           =   7
               X1              =   1980
               X2              =   2700
               Y1              =   660
               Y2              =   660
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   660
               X2              =   1380
               Y1              =   2010
               Y2              =   2010
            End
            Begin VB.Label lblCancelBespeak 
               AutoSize        =   -1  'True
               Caption         =   "ԤԼ��         ���ڲ���ȡ��ԤԼ"
               Height          =   180
               Left            =   75
               TabIndex        =   278
               Top             =   1800
               Width           =   2790
            End
            Begin VB.Label lblBespeakMinTime 
               AutoSize        =   -1  'True
               Caption         =   "ԤԼ����ʱ��        ���ӣ�ָԤԼʱ���������ʱ�̵���С���"
               Height          =   180
               Left            =   75
               TabIndex        =   287
               Top             =   2520
               Width           =   5220
            End
            Begin VB.Line Line1 
               Index           =   4
               X1              =   1185
               X2              =   1905
               Y1              =   2745
               Y2              =   2745
            End
            Begin VB.Line Line1 
               Index           =   5
               X1              =   1185
               X2              =   1905
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Label lblBespeakDefaultDays 
               AutoSize        =   -1  'True
               Caption         =   "ԤԼȱʡ����        ��"
               Height          =   180
               Left            =   75
               TabIndex        =   285
               Top             =   2160
               Width           =   1980
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7395
         Index           =   7
         Left            =   -74850
         ScaleHeight     =   7365
         ScaleWidth      =   9060
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   1020
         Width           =   9090
         Begin VB.CheckBox chk 
            Caption         =   "Ԥ�����վ����ʾ"
            Height          =   300
            Index           =   202
            Left            =   285
            TabIndex        =   148
            Top             =   2790
            Width           =   1785
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ֹ��Ժ���˽�����Ԥ��"
            Height          =   300
            Index           =   57
            Left            =   285
            TabIndex        =   141
            Top             =   1425
            Width           =   2880
         End
         Begin VB.CheckBox chk 
            Caption         =   "������Ժ���˽���סԺ����˿�"
            Height          =   300
            Index           =   188
            Left            =   285
            TabIndex        =   147
            Top             =   2505
            Width           =   2925
         End
         Begin VB.Frame fraƱ�ݸ�ʽ 
            Height          =   120
            Index           =   4
            Left            =   1740
            TabIndex        =   672
            Top             =   5685
            Width           =   7770
         End
         Begin VB.CheckBox chk 
            Caption         =   "��סԺԤ��ˢ����֤"
            Height          =   300
            Index           =   6
            Left            =   285
            TabIndex        =   143
            Top             =   1935
            Width           =   2340
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ͨ������ģ�����Ҳ���"
            Height          =   300
            Index           =   5
            Left            =   285
            TabIndex        =   142
            Top             =   1680
            Value           =   1  'Checked
            Width           =   3840
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����Ժ���˽�סԺԤ��"
            Height          =   300
            Index           =   4
            Left            =   285
            TabIndex        =   140
            Top             =   1155
            Width           =   2340
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ժ����δ��Ʋ�׼��Ԥ��"
            Height          =   300
            Index           =   2
            Left            =   285
            TabIndex        =   137
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1290
            TabIndex        =   145
            Text            =   "10"
            Top             =   2220
            Width           =   510
         End
         Begin VB.Frame fraƱ�ݸ�ʽ 
            Height          =   120
            Index           =   0
            Left            =   1755
            TabIndex        =   138
            Top             =   4200
            Width           =   7770
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ԥ�������������Ϣ"
            Height          =   300
            Index           =   0
            Left            =   285
            TabIndex        =   135
            Top             =   105
            Width           =   2340
         End
         Begin VB.Frame fra�˿����� 
            Caption         =   "�˿�����"
            Height          =   930
            Left            =   285
            TabIndex        =   150
            Top             =   3120
            Width           =   4050
            Begin VB.OptionButton optDepsoitDelSet 
               Caption         =   "����ʱ�����˿�"
               Height          =   285
               Index           =   0
               Left            =   495
               TabIndex        =   151
               Top             =   270
               Width           =   2625
            End
            Begin VB.OptionButton optDepsoitDelSet 
               Caption         =   "����ʱ��ֹ�˿�"
               Height          =   285
               Index           =   1
               Left            =   495
               TabIndex        =   152
               Top             =   555
               Value           =   -1  'True
               Width           =   2220
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "������Ĳ��˵Ľɿ����"
            Height          =   300
            Index           =   3
            Left            =   285
            TabIndex        =   139
            Top             =   885
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ֻ��ʾ��ʣ�����ʷ�ɿ�"
            Height          =   300
            Index           =   1
            Left            =   285
            TabIndex        =   136
            Top             =   345
            Width           =   3120
         End
         Begin VSFlex8Ctl.VSFlexGrid vs���� 
            Height          =   3840
            Left            =   4620
            TabIndex        =   153
            Top             =   195
            Width           =   4590
            _cx             =   8096
            _cy             =   6773
            Appearance      =   2
            BorderStyle     =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483628
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
            Rows            =   10
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParFee.frx":1B989
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
            ShowComboButton =   0
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
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   2
            Left            =   1770
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   2220
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   10
            BuddyControl    =   "txtUD(2)"
            BuddyDispid     =   196647
            BuddyIndex      =   2
            OrigLeft        =   1755
            OrigTop         =   2505
            OrigRight       =   2010
            OrigBottom      =   2805
            Max             =   1000
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chk 
            Caption         =   "Ʊ��ʣ��         ��ʱ��ʼ�����շ�Ա"
            Height          =   300
            Index           =   11
            Left            =   285
            TabIndex        =   144
            Top             =   2235
            Width           =   3450
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1050
            Index           =   0
            Left            =   315
            TabIndex        =   154
            Top             =   4530
            Width           =   8940
            _cx             =   15769
            _cy             =   1852
            Appearance      =   2
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
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
            Rows            =   3
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParFee.frx":1B9EA
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
            ExplorerBar     =   2
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
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1125
            Index           =   4
            Left            =   315
            TabIndex        =   673
            Top             =   6015
            Width           =   8940
            _cx             =   15769
            _cy             =   1984
            Appearance      =   2
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
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
            Rows            =   3
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParFee.frx":1BA78
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
            ExplorerBar     =   2
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
         Begin VB.Label lblDepositPrintRedSet 
            AutoSize        =   -1  'True
            Caption         =   "Ԥ����Ʊ��ӡ����"
            Height          =   180
            Left            =   285
            TabIndex        =   674
            Top             =   5715
            Width           =   1440
         End
         Begin VB.Label lblDepositPrintSet 
            AutoSize        =   -1  'True
            Caption         =   "Ԥ��Ʊ�ݴ�ӡ����"
            Height          =   180
            Left            =   285
            TabIndex        =   149
            Top             =   4230
            Width           =   1440
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   9810
         Index           =   11
         Left            =   -74640
         ScaleHeight     =   9780
         ScaleWidth      =   11655
         TabIndex        =   359
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   11685
         Begin VB.PictureBox picChargePg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   7620
            Index           =   0
            Left            =   60
            ScaleHeight     =   7590
            ScaleWidth      =   8445
            TabIndex        =   429
            TabStop         =   0   'False
            Top             =   -960
            Width           =   8475
            Begin VB.Frame fra�˷�ȱʡ��ʽ 
               Caption         =   "�˷�ȱʡ��ʽ"
               Height          =   1275
               Left            =   5190
               TabIndex        =   497
               Top             =   6210
               Width           =   4215
               Begin VSFlex8Ctl.VSFlexGrid vsfDelFeeDefaultType 
                  Height          =   1035
                  Left            =   60
                  TabIndex        =   498
                  Top             =   210
                  Width           =   4095
                  _cx             =   7223
                  _cy             =   1826
                  Appearance      =   2
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
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   0
                  HighLight       =   1
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"frmParFee.frx":1BB06
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
            End
            Begin VB.Frame fraNormal 
               Height          =   5085
               Index           =   2
               Left            =   165
               TabIndex        =   430
               Top             =   90
               Width           =   4800
               Begin VB.CheckBox chk 
                  Caption         =   "����¼������ʹ�õĿ�����"
                  Height          =   195
                  Index           =   196
                  Left            =   2280
                  TabIndex        =   460
                  Top             =   4560
                  Width           =   2490
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�൥�ݷֵ��ݽ���ʱ��ֻ��ҽ������ɹ��ĵ����շ�"
                  Height          =   195
                  Index           =   170
                  Left            =   240
                  TabIndex        =   461
                  Top             =   4830
                  Width           =   4500
               End
               Begin VB.CheckBox chk 
                  Caption         =   "δ�Һ�ʱ�Զ������շ���Ŀ"
                  Height          =   195
                  Index           =   119
                  Left            =   240
                  TabIndex        =   449
                  Top             =   3195
                  Width           =   2460
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H00C00000&
                  Height          =   270
                  Index           =   17
                  Left            =   2700
                  Locked          =   -1  'True
                  TabIndex        =   450
                  Top             =   3157
                  Width           =   1575
               End
               Begin VB.CommandButton cmdAddedItem 
                  Caption         =   "��"
                  Height          =   280
                  Left            =   4290
                  TabIndex        =   451
                  TabStop         =   0   'False
                  Top             =   3152
                  Width           =   280
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�����˷���������"
                  Height          =   195
                  Index           =   16
                  Left            =   240
                  TabIndex        =   459
                  Top             =   4560
                  Width           =   1920
               End
               Begin VB.ComboBox cbo 
                  Height          =   300
                  Index           =   3
                  Left            =   1845
                  Style           =   2  'Dropdown List
                  TabIndex        =   456
                  Top             =   3840
                  Width           =   1170
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000012&
                  Height          =   180
                  Index           =   18
                  Left            =   1680
                  MaxLength       =   2
                  TabIndex        =   448
                  Text            =   "1"
                  Top             =   2940
                  Width           =   285
               End
               Begin VB.Frame fraLineDays 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   15
                  Index           =   3
                  Left            =   1650
                  TabIndex        =   500
                  Top             =   3120
                  Width           =   285
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Index           =   15
                  Left            =   2895
                  MaxLength       =   3
                  TabIndex        =   446
                  Text            =   "0"
                  Top             =   2655
                  Width           =   285
               End
               Begin VB.Frame fraLineDays 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   15
                  Index           =   1
                  Left            =   2865
                  TabIndex        =   499
                  Top             =   2835
                  Width           =   285
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��ҩ���븶��"
                  Height          =   195
                  Index           =   70
                  Left            =   240
                  TabIndex        =   431
                  Top             =   195
                  Value           =   1  'Checked
                  Width           =   1380
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�����������"
                  Height          =   195
                  Index           =   73
                  Left            =   2820
                  TabIndex        =   432
                  Top             =   195
                  Width           =   1380
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�����˺���ʿ"
                  Height          =   195
                  Index           =   76
                  Left            =   240
                  TabIndex        =   433
                  Top             =   480
                  Width           =   1380
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��ʾ�տ��ۼ�"
                  Height          =   195
                  Index           =   116
                  Left            =   2820
                  TabIndex        =   434
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   1380
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��ȡ���۵��շ�ʱ���Ƥ�Խ��"
                  Height          =   195
                  Index           =   117
                  Left            =   240
                  TabIndex        =   444
                  Top             =   2325
                  Width           =   2820
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����ʹ��Ԥ����ɷ�"
                  Height          =   195
                  Index           =   118
                  Left            =   240
                  TabIndex        =   435
                  Top             =   780
                  Width           =   2040
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��ȡ���۵��������ɿ�"
                  Height          =   300
                  Index           =   127
                  Left            =   240
                  TabIndex        =   437
                  Top             =   1065
                  Width           =   2160
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�շ�ʱ����ͬʱ������ŵ���"
                  Height          =   195
                  Index           =   120
                  Left            =   240
                  TabIndex        =   443
                  Top             =   2040
                  Width           =   3000
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�շ�ʱ��鲡�˹Һſ���"
                  Height          =   195
                  Index           =   123
                  Left            =   240
                  TabIndex        =   441
                  Top             =   1755
                  Width           =   2295
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���������۵�ѡ�񴰿�"
                  Height          =   195
                  Index           =   122
                  Left            =   240
                  TabIndex        =   439
                  Top             =   1455
                  Width           =   2160
               End
               Begin VB.TextBox txtUD 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Index           =   8
                  Left            =   1260
                  TabIndex        =   453
                  Text            =   "10"
                  Top             =   3480
                  Width           =   495
               End
               Begin VB.CheckBox chk 
                  Caption         =   "סԺ���˰������շ�"
                  Height          =   270
                  Index           =   126
                  Left            =   2820
                  TabIndex        =   442
                  Top             =   1710
                  Width           =   1920
               End
               Begin VB.CheckBox chk 
                  Caption         =   "ȱʡ��������"
                  Height          =   300
                  Index           =   115
                  Left            =   2820
                  TabIndex        =   438
                  Top             =   1035
                  Width           =   1620
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����Ҫ���뿪����"
                  Height          =   195
                  Index           =   114
                  Left            =   2820
                  TabIndex        =   440
                  Top             =   1425
                  Width           =   1740
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��ʹ��ȱʡ������"
                  Height          =   195
                  Index           =   113
                  Left            =   2820
                  TabIndex        =   436
                  Top             =   765
                  Width           =   1740
               End
               Begin VB.TextBox txt 
                  ForeColor       =   &H80000012&
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Index           =   13
                  Left            =   1380
                  MaxLength       =   12
                  TabIndex        =   458
                  Text            =   "0.00"
                  Top             =   4200
                  Width           =   1335
               End
               Begin MSComCtl2.UpDown ud 
                  Height          =   300
                  Index           =   8
                  Left            =   1755
                  TabIndex        =   454
                  Top             =   3480
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   10
                  BuddyControl    =   "txtUD(8)"
                  BuddyDispid     =   196647
                  BuddyIndex      =   8
                  OrigLeft        =   6345
                  OrigTop         =   3420
                  OrigRight       =   6600
                  OrigBottom      =   3705
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�Զ���Ѱ����    ���ڵĻ��۵���"
                  Height          =   195
                  Index           =   121
                  Left            =   240
                  TabIndex        =   447
                  Top             =   2910
                  Width           =   3000
               End
               Begin VB.CheckBox chk 
                  Caption         =   "Ʊ��ʣ��         ��ʱ��ʼ�����շ�Ա"
                  Height          =   285
                  Index           =   125
                  Left            =   240
                  TabIndex        =   452
                  Top             =   3495
                  Width           =   3450
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�շ���ϸ�Զ���              ��ϵ���"
                  Height          =   195
                  Index           =   124
                  Left            =   240
                  TabIndex        =   455
                  Top             =   3900
                  Width           =   3570
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����ͨ������������ģ������    ���ڵĲ�����Ϣ"
                  Height          =   195
                  Index           =   86
                  Left            =   240
                  TabIndex        =   445
                  Top             =   2625
                  Width           =   4260
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���������"
                  Height          =   210
                  Left            =   240
                  TabIndex        =   457
                  Top             =   4245
                  Width           =   1080
               End
            End
            Begin VB.Frame fra��λ 
               Caption         =   " ҩƷ��λ "
               Height          =   810
               Index           =   1
               Left            =   165
               TabIndex        =   462
               Top             =   5325
               Width           =   4785
               Begin VB.OptionButton opt�շѵ�λ 
                  Caption         =   "�ۼ۵�λ"
                  Height          =   180
                  Index           =   0
                  Left            =   1365
                  TabIndex        =   464
                  Top             =   405
                  Value           =   -1  'True
                  Width           =   1020
               End
               Begin VB.OptionButton opt�շѵ�λ 
                  Caption         =   "����(��סԺ)��λ"
                  Height          =   180
                  Index           =   1
                  Left            =   2505
                  TabIndex        =   465
                  Top             =   405
                  Width           =   1770
               End
               Begin VB.Label lbl��λ 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�շ�ʱ��"
                  Height          =   180
                  Index           =   1
                  Left            =   300
                  TabIndex        =   463
                  Top             =   405
                  Width           =   975
               End
            End
            Begin VB.Frame fra 
               Caption         =   "��������ʾ��ʽ"
               Height          =   1260
               Index           =   2
               Left            =   5190
               TabIndex        =   474
               Top             =   90
               Width           =   4215
               Begin VB.OptionButton optChargeBillTotalShow 
                  Caption         =   "�����ݷ��������ʾ"
                  Height          =   195
                  Index           =   2
                  Left            =   300
                  TabIndex        =   477
                  Top             =   900
                  Width           =   2280
               End
               Begin VB.OptionButton optChargeBillTotalShow 
                  Caption         =   "��������Ŀ��ʾ����ϼ�"
                  Height          =   195
                  Index           =   1
                  Left            =   300
                  TabIndex        =   476
                  Top             =   645
                  Width           =   2280
               End
               Begin VB.OptionButton optChargeBillTotalShow 
                  Caption         =   "���վݷ�Ŀ��ʾ����ϼ�"
                  Height          =   195
                  Index           =   0
                  Left            =   300
                  TabIndex        =   475
                  Top             =   375
                  Value           =   -1  'True
                  Width           =   2280
               End
            End
            Begin VB.Frame fraBillInputItem 
               Caption         =   "�շ�ʱҪ�������Ŀ"
               Height          =   1065
               Index           =   1
               Left            =   165
               TabIndex        =   466
               Top             =   6240
               Width           =   4785
               Begin VB.CheckBox chk 
                  Caption         =   "������"
                  Height          =   210
                  Index           =   109
                  Left            =   1530
                  TabIndex        =   472
                  Top             =   675
                  Value           =   1  'Checked
                  Width           =   840
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��������"
                  Height          =   210
                  Index           =   101
                  Left            =   240
                  TabIndex        =   471
                  Top             =   675
                  Value           =   1  'Checked
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�Ƿ�Ӱ�"
                  Height          =   210
                  Index           =   100
                  Left            =   1530
                  TabIndex        =   468
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   1020
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�ѱ�"
                  Height          =   210
                  Index           =   99
                  Left            =   3810
                  TabIndex        =   470
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����"
                  Height          =   210
                  Index           =   98
                  Left            =   2805
                  TabIndex        =   469
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�Ա�"
                  Height          =   210
                  Index           =   97
                  Left            =   240
                  TabIndex        =   467
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   660
               End
               Begin VB.CheckBox chk 
                  Caption         =   "ҽ�Ƹ��ʽ"
                  Height          =   210
                  Index           =   96
                  Left            =   2805
                  TabIndex        =   473
                  Top             =   675
                  Value           =   1  'Checked
                  Width           =   1380
               End
            End
            Begin VB.Frame fra�����ʾ 
               Caption         =   "�����ʾ"
               Height          =   1320
               Index           =   1
               Left            =   5190
               TabIndex        =   478
               Top             =   1410
               Width           =   4215
               Begin VB.CheckBox chk 
                  Caption         =   "��ʾ����ҩ����"
                  Height          =   195
                  Index           =   81
                  Left            =   2250
                  TabIndex        =   480
                  Top             =   375
                  Width           =   1770
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��ʾ����ҩ�����"
                  Height          =   195
                  Index           =   80
                  Left            =   300
                  TabIndex        =   479
                  Top             =   375
                  Width           =   1770
               End
               Begin VB.OptionButton opt�շѿ����ʾ��ʽ 
                  Caption         =   "��ʾ�����"
                  Height          =   180
                  Index           =   0
                  Left            =   1455
                  TabIndex        =   482
                  Top             =   915
                  Width           =   1290
               End
               Begin VB.OptionButton opt�շѿ����ʾ��ʽ 
                  Caption         =   "����ʾ����"
                  Height          =   180
                  Index           =   1
                  Left            =   2760
                  TabIndex        =   483
                  Top             =   915
                  Width           =   1215
               End
               Begin VB.Label lbl�����ʾ��ʽ 
                  AutoSize        =   -1  'True
                  Caption         =   "�����ʾ��ʽ"
                  Height          =   180
                  Index           =   1
                  Left            =   300
                  TabIndex        =   481
                  Top             =   915
                  Width           =   1080
               End
               Begin VB.Line lnSplit 
                  BorderColor     =   &H80000000&
                  Index           =   3
                  X1              =   15
                  X2              =   4180
                  Y1              =   705
                  Y2              =   705
               End
               Begin VB.Line lnSplit 
                  BorderColor     =   &H00FFFFFF&
                  Index           =   2
                  X1              =   15
                  X2              =   4180
                  Y1              =   720
                  Y2              =   720
               End
            End
            Begin VB.Frame fraRegPrompt 
               Caption         =   "δ�ҺŲ����շ�"
               Height          =   810
               Left            =   5190
               TabIndex        =   484
               Top             =   2790
               Width           =   4215
               Begin VB.OptionButton optChargeRegPrompt 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   485
                  Top             =   405
                  Value           =   -1  'True
                  Width           =   795
               End
               Begin VB.OptionButton optChargeRegPrompt 
                  Caption         =   "��ֹ"
                  Height          =   180
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   487
                  Top             =   405
                  Width           =   795
               End
               Begin VB.OptionButton optChargeRegPrompt 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   486
                  Top             =   405
                  Width           =   795
               End
            End
            Begin VB.Frame fra 
               Caption         =   "ҩƷ��ҩ���˷ѷ�ʽ"
               Height          =   810
               Index           =   17
               Left            =   5190
               TabIndex        =   488
               Top             =   3675
               Width           =   4215
               Begin VB.OptionButton optDrug 
                  Caption         =   "�����"
                  Height          =   180
                  Index           =   0
                  Left            =   300
                  TabIndex        =   489
                  Top             =   420
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton optDrug 
                  Caption         =   "��ֹ"
                  Height          =   180
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   490
                  Top             =   420
                  Width           =   690
               End
               Begin VB.OptionButton optDrug 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   491
                  Top             =   420
                  Width           =   690
               End
            End
            Begin VB.Frame fra�ɿ���� 
               Caption         =   "�ɿ����������"
               Height          =   1605
               Left            =   5190
               TabIndex        =   492
               Top             =   4545
               Width           =   4215
               Begin VB.OptionButton opt�ɿ� 
                  Caption         =   $"frmParFee.frx":1BB6C
                  Height          =   285
                  Index           =   2
                  Left            =   300
                  TabIndex        =   494
                  Top             =   600
                  Width           =   2655
               End
               Begin VB.OptionButton opt�ɿ� 
                  Caption         =   $"frmParFee.frx":1BB8A
                  Height          =   285
                  Index           =   0
                  Left            =   300
                  TabIndex        =   493
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   3780
               End
               Begin VB.OptionButton opt�ɿ� 
                  Caption         =   "�շ�ʱ���ಡ���ۼ�"
                  Height          =   285
                  Index           =   1
                  Left            =   300
                  TabIndex        =   495
                  Top             =   870
                  Width           =   2715
               End
               Begin VB.OptionButton opt�ɿ� 
                  Caption         =   "�շ�ʱ���������ۼ�"
                  Height          =   285
                  Index           =   3
                  Left            =   300
                  TabIndex        =   496
                  Top             =   1155
                  Width           =   2715
               End
            End
         End
         Begin VB.PictureBox picChargePg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   7455
            Index           =   1
            Left            =   450
            ScaleHeight     =   7425
            ScaleWidth      =   9105
            TabIndex        =   501
            TabStop         =   0   'False
            Top             =   240
            Width           =   9135
            Begin VB.Frame fraƱ�ݸ�ʽ 
               Caption         =   "�˷�Ʊ�ݸ�ʽ"
               Height          =   1455
               Index           =   5
               Left            =   180
               TabIndex        =   542
               Top             =   5850
               Width           =   6555
               Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
                  Height          =   1125
                  Index           =   5
                  Left            =   90
                  TabIndex        =   521
                  Top             =   255
                  Width           =   6375
                  _cx             =   11245
                  _cy             =   1984
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
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   8421504
                  GridColorFixed  =   8421504
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
                  Rows            =   3
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":1BBA8
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
                  ExplorerBar     =   2
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
            Begin VB.CheckBox chk 
               Caption         =   "�����˲���Ʊ�ݲ����ݽ����������Ʊ"
               Height          =   180
               Index           =   173
               Left            =   3195
               TabIndex        =   671
               Top             =   4290
               Width           =   3540
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   4
               ItemData        =   "frmParFee.frx":1BC3E
               Left            =   1470
               List            =   "frmParFee.frx":1BC40
               Style           =   2  'Dropdown List
               TabIndex        =   502
               Top             =   255
               Width           =   3015
            End
            Begin VB.Frame fraBillSplitRule 
               Caption         =   "Ʊ�ݷ������ "
               Height          =   3900
               Left            =   195
               TabIndex        =   503
               Top             =   330
               Width           =   6555
               Begin VB.PictureBox picRuleBack 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000008&
                  Height          =   2580
                  Index           =   0
                  Left            =   60
                  ScaleHeight     =   2550
                  ScaleWidth      =   6000
                  TabIndex        =   525
                  TabStop         =   0   'False
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   6030
                  Begin VB.CheckBox chk 
                     Caption         =   "�����շ�ʱ�Զ����չ�����"
                     Height          =   195
                     Index           =   130
                     Left            =   330
                     TabIndex        =   526
                     Top             =   855
                     Width           =   2460
                  End
                  Begin VB.CheckBox chk 
                     Caption         =   "��첡��ÿ�ŵ��ݷֱ��ӡ(�ò���ͬʱӰ�칤������������)"
                     Height          =   195
                     Index           =   189
                     Left            =   630
                     TabIndex        =   528
                     Top             =   300
                     Width           =   5190
                  End
                  Begin VB.CheckBox chk 
                     Caption         =   "�����շ�ÿ�ŵ��ݷֱ��ӡ(�ò���ͬʱӰ�칤������������)"
                     Height          =   195
                     Index           =   128
                     Left            =   345
                     TabIndex        =   527
                     Top             =   75
                     Width           =   5160
                  End
                  Begin VB.CheckBox chk 
                     Caption         =   "�շ�ÿ�δ�ӡֻ��һ��Ʊ��(�ò���ͬʱӰ�칤������������)"
                     Height          =   195
                     Index           =   129
                     Left            =   330
                     TabIndex        =   529
                     Top             =   585
                     Width           =   5160
                  End
                  Begin VB.Frame fraActuallyPrint 
                     Height          =   1695
                     Left            =   105
                     TabIndex        =   530
                     Top             =   855
                     Width           =   5850
                     Begin VB.OptionButton optBillMode 
                        Caption         =   "��ӡ�վݷ�Ŀ"
                        Height          =   255
                        Index           =   0
                        Left            =   2505
                        TabIndex        =   533
                        Top             =   795
                        Value           =   -1  'True
                        Width           =   1575
                     End
                     Begin VB.OptionButton optBillMode 
                        Caption         =   "��ӡ�շ���Ŀ"
                        Height          =   255
                        Index           =   1
                        Left            =   4065
                        TabIndex        =   534
                        Top             =   795
                        Width           =   1455
                     End
                     Begin VB.CheckBox chk 
                        Caption         =   "��ִ�п��ҷֱ��ӡ"
                        Height          =   195
                        Index           =   131
                        Left            =   200
                        TabIndex        =   532
                        Top             =   825
                        Width           =   1980
                     End
                     Begin VB.TextBox txtUD 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Index           =   9
                        Left            =   1305
                        Locked          =   -1  'True
                        TabIndex        =   536
                        Text            =   "3"
                        Top             =   1140
                        Width           =   405
                     End
                     Begin MSComCtl2.UpDown ud 
                        Height          =   300
                        Index           =   9
                        Left            =   1680
                        TabIndex        =   537
                        TabStop         =   0   'False
                        Top             =   1140
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   529
                        _Version        =   393216
                        Value           =   3
                        BuddyControl    =   "txtUD(9)"
                        BuddyDispid     =   196647
                        BuddyIndex      =   9
                        OrigLeft        =   1620
                        OrigTop         =   1140
                        OrigRight       =   1875
                        OrigBottom      =   1440
                        Max             =   100
                        Min             =   3
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   65547
                        Enabled         =   -1  'True
                     End
                     Begin VB.Label lblRows 
                        AutoSize        =   -1  'True
                        Caption         =   "�շ��վ��д�"
                        Height          =   180
                        Left            =   150
                        TabIndex        =   535
                        Top             =   1200
                        Width           =   1080
                     End
                     Begin VB.Label lbl 
                        Caption         =   "������������Ʊ����������,Ʊ�����������¹������.��ʵ�ʴ�ӡ������Ʊ������Դ��Ʊ����ƾ���,������߲�һ��,��������������׼ȷ."
                        Height          =   495
                        Index           =   25
                        Left            =   120
                        TabIndex        =   531
                        Top             =   285
                        Width           =   5655
                     End
                  End
               End
               Begin VB.PictureBox picRuleBack 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  ForeColor       =   &H80000008&
                  Height          =   3525
                  Index           =   1
                  Left            =   105
                  ScaleHeight     =   3495
                  ScaleWidth      =   6345
                  TabIndex        =   504
                  TabStop         =   0   'False
                  Top             =   285
                  Visible         =   0   'False
                  Width           =   6375
                  Begin VB.Frame fraRuleSystem 
                     Height          =   3345
                     Left            =   150
                     TabIndex        =   505
                     Top             =   75
                     Width           =   6150
                     Begin VB.OptionButton optRuleTotal 
                        Caption         =   "��ִ�п��ҷ������"
                        Height          =   240
                        Index           =   2
                        Left            =   2985
                        TabIndex        =   523
                        Top             =   2265
                        Width           =   2025
                     End
                     Begin VB.OptionButton optRuleTotal 
                        Caption         =   "��ҳ��ӡ����"
                        Height          =   240
                        Index           =   1
                        Left            =   1425
                        TabIndex        =   522
                        Top             =   2265
                        Width           =   1440
                     End
                     Begin VB.OptionButton optRuleTotal 
                        Caption         =   "������"
                        Height          =   240
                        Index           =   0
                        Left            =   330
                        TabIndex        =   520
                        Top             =   2265
                        Value           =   -1  'True
                        Width           =   1005
                     End
                     Begin VB.TextBox txtBillRuleNum 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Index           =   2
                        Left            =   2430
                        Locked          =   -1  'True
                        TabIndex        =   518
                        Text            =   "3"
                        Top             =   1875
                        Width           =   315
                     End
                     Begin VB.TextBox txtBillRuleNum 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Index           =   1
                        Left            =   2430
                        Locked          =   -1  'True
                        TabIndex        =   514
                        Text            =   "3"
                        Top             =   1530
                        Width           =   315
                     End
                     Begin VB.TextBox txtBillRuleNum 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Index           =   0
                        Left            =   2430
                        Locked          =   -1  'True
                        TabIndex        =   510
                        Text            =   "3"
                        Top             =   1155
                        Width           =   315
                     End
                     Begin VB.CheckBox chkBillRule 
                        Caption         =   "4.���շ�ϸĿ��ҳ"
                        Height          =   180
                        Index           =   3
                        Left            =   285
                        TabIndex        =   516
                        Top             =   1920
                        Width           =   1770
                     End
                     Begin VB.CheckBox chkBillRule 
                        Caption         =   "3.���վݷ�Ŀ��ҳ"
                        Height          =   180
                        Index           =   2
                        Left            =   270
                        TabIndex        =   512
                        Top             =   1575
                        Width           =   1770
                     End
                     Begin VB.CheckBox chkBillRule 
                        Caption         =   "2.��ִ�п��ҷ�ҳ"
                        Height          =   180
                        Index           =   1
                        Left            =   270
                        TabIndex        =   508
                        Top             =   1215
                        Width           =   1770
                     End
                     Begin VB.CheckBox chkBillRule 
                        Caption         =   "1.�����ݷ�ҳ"
                        Height          =   225
                        Index           =   0
                        Left            =   270
                        TabIndex        =   507
                        Top             =   870
                        Width           =   1635
                     End
                     Begin MSComCtl2.UpDown updBillRuleNum 
                        Height          =   300
                        Index           =   0
                        Left            =   2760
                        TabIndex        =   511
                        TabStop         =   0   'False
                        Top             =   1155
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   529
                        _Version        =   393216
                        Value           =   1
                        BuddyControl    =   "txtBillRuleNum(0)"
                        BuddyDispid     =   196806
                        BuddyIndex      =   0
                        OrigLeft        =   4440
                        OrigTop         =   825
                        OrigRight       =   4695
                        OrigBottom      =   1125
                        Max             =   100
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   65547
                        Enabled         =   -1  'True
                     End
                     Begin MSComCtl2.UpDown updBillRuleNum 
                        Height          =   300
                        Index           =   1
                        Left            =   2760
                        TabIndex        =   515
                        TabStop         =   0   'False
                        Top             =   1530
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   529
                        _Version        =   393216
                        Value           =   4
                        BuddyControl    =   "txtBillRuleNum(1)"
                        BuddyDispid     =   196806
                        BuddyIndex      =   1
                        OrigLeft        =   4440
                        OrigTop         =   825
                        OrigRight       =   4695
                        OrigBottom      =   1125
                        Max             =   100
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   65547
                        Enabled         =   -1  'True
                     End
                     Begin MSComCtl2.UpDown updBillRuleNum 
                        Height          =   300
                        Index           =   2
                        Left            =   2760
                        TabIndex        =   519
                        TabStop         =   0   'False
                        Top             =   1875
                        Width           =   255
                        _ExtentX        =   450
                        _ExtentY        =   529
                        _Version        =   393216
                        Value           =   20
                        BuddyControl    =   "txtBillRuleNum(2)"
                        BuddyDispid     =   196806
                        BuddyIndex      =   2
                        OrigLeft        =   4440
                        OrigTop         =   825
                        OrigRight       =   4695
                        OrigBottom      =   1125
                        Max             =   100
                        SyncBuddy       =   -1  'True
                        BuddyProperty   =   65547
                        Enabled         =   -1  'True
                     End
                     Begin VB.Label lblBillRuleNum 
                        AutoSize        =   -1  'True
                        Caption         =   "��ÿ����   ���շ�ϸĿ��һҳ"
                        Height          =   180
                        Index           =   2
                        Left            =   2025
                        TabIndex        =   517
                        Top             =   1935
                        Width           =   2430
                     End
                     Begin VB.Label lblBillRuleNum 
                        AutoSize        =   -1  'True
                        Caption         =   "��ÿ����   ���վݷ�Ŀ��һҳ"
                        Height          =   180
                        Index           =   1
                        Left            =   2025
                        TabIndex        =   513
                        Top             =   1590
                        Width           =   2430
                     End
                     Begin VB.Label lblBillRuleNum 
                        AutoSize        =   -1  'True
                        Caption         =   "��ÿ����   ��ִ�п��ҷ�һҳ"
                        Height          =   180
                        Index           =   0
                        Left            =   2025
                        TabIndex        =   509
                        Top             =   1215
                        Width           =   2430
                     End
                     Begin VB.Label lblInfor 
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        BorderStyle     =   1  'Fixed Single
                        ForeColor       =   &H80000008&
                        Height          =   540
                        Left            =   75
                        TabIndex        =   524
                        Top             =   2700
                        Width           =   6030
                     End
                     Begin VB.Label lblRuleSystem 
                        Caption         =   "������������Ʊ����������,Ʊ�����������¹������.��ʵ�ʴ�ӡ�������շѻ��۵�����,����ֹ�¼����õ���,�����ѵ��������㽫��׼ȷ."
                        Height          =   585
                        Left            =   180
                        TabIndex        =   506
                        Top             =   330
                        Width           =   5730
                     End
                  End
               End
               Begin VB.PictureBox picRuleBack 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000008&
                  Height          =   1050
                  Index           =   2
                  Left            =   150
                  ScaleHeight     =   1020
                  ScaleWidth      =   6210
                  TabIndex        =   538
                  TabStop         =   0   'False
                  Top             =   330
                  Visible         =   0   'False
                  Width           =   6240
                  Begin VB.Label lblCustomInfor 
                     Caption         =   $"frmParFee.frx":1BC42
                     Height          =   570
                     Left            =   60
                     TabIndex        =   539
                     Top             =   330
                     Width           =   6165
                  End
               End
            End
            Begin VB.Frame fraƱ�ݸ�ʽ 
               Caption         =   "�շ�Ʊ�ݸ�ʽ"
               Height          =   1455
               Index           =   1
               Left            =   180
               TabIndex        =   540
               Top             =   4305
               Width           =   6555
               Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
                  Height          =   1125
                  Index           =   1
                  Left            =   90
                  TabIndex        =   541
                  Top             =   255
                  Width           =   6375
                  _cx             =   11245
                  _cy             =   1984
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
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   8421504
                  GridColorFixed  =   8421504
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
                  Rows            =   3
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":1BCE1
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
                  ExplorerBar     =   2
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
            Begin VB.Frame fraFeeList 
               Caption         =   "�շѺ�����嵥"
               Height          =   1290
               Left            =   7020
               TabIndex        =   543
               Top             =   330
               Width           =   2205
               Begin VB.OptionButton optFeeListPrint 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   345
                  TabIndex        =   544
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1020
               End
               Begin VB.OptionButton optFeeListPrint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   345
                  TabIndex        =   546
                  Top             =   915
                  Width           =   1455
               End
               Begin VB.OptionButton optFeeListPrint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   345
                  TabIndex        =   545
                  Top             =   630
                  Width           =   1065
               End
            End
            Begin VB.Frame fraFeeExe 
               Caption         =   "�շ�ִ�е�"
               Height          =   1290
               Left            =   7020
               TabIndex        =   551
               Top             =   3180
               Width           =   2205
               Begin VB.OptionButton optChargeExeBillPrint 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   345
                  TabIndex        =   552
                  Top             =   420
                  Value           =   -1  'True
                  Width           =   1020
               End
               Begin VB.OptionButton optChargeExeBillPrint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   345
                  TabIndex        =   554
                  Top             =   930
                  Width           =   1455
               End
               Begin VB.OptionButton optChargeExeBillPrint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   345
                  TabIndex        =   553
                  Top             =   660
                  Width           =   1065
               End
            End
            Begin VB.Frame fraRefundReceipt 
               Caption         =   "�˷ѻص�����"
               Height          =   1290
               Left            =   7020
               TabIndex        =   547
               Top             =   1740
               Width           =   2205
               Begin VB.OptionButton optDelFeeRefundPrint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   360
                  TabIndex        =   549
                  Top             =   630
                  Width           =   1065
               End
               Begin VB.OptionButton optDelFeeRefundPrint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   360
                  TabIndex        =   550
                  Top             =   900
                  Width           =   1455
               End
               Begin VB.OptionButton optDelFeeRefundPrint 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   360
                  TabIndex        =   548
                  Top             =   375
                  Value           =   -1  'True
                  Width           =   1020
               End
            End
         End
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   975
            Index           =   1
            Left            =   150
            TabIndex        =   428
            TabStop         =   0   'False
            Top             =   105
            Width           =   5025
            _Version        =   589884
            _ExtentX        =   8864
            _ExtentY        =   1720
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7230
         Index           =   14
         Left            =   -74910
         ScaleHeight     =   7200
         ScaleWidth      =   9165
         TabIndex        =   583
         TabStop         =   0   'False
         Top             =   750
         Visible         =   0   'False
         Width           =   9195
         Begin VB.ComboBox cbo 
            ForeColor       =   &H80000012&
            Height          =   300
            Index           =   6
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   607
            Top             =   5655
            Width           =   1770
         End
         Begin VB.Frame fraPrint 
            Caption         =   "���ݴ�ӡ"
            Height          =   1815
            Left            =   195
            TabIndex        =   600
            Top             =   3720
            Width           =   4950
            Begin VB.CheckBox chk 
               Caption         =   "ҽ�������м��ʺ��ӡ����"
               Height          =   195
               Index           =   140
               Left            =   150
               TabIndex        =   603
               Top             =   840
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "���ҷ�ɢ�����м��ʺ��ӡ����"
               Height          =   240
               Index           =   139
               Left            =   150
               TabIndex        =   602
               Top             =   570
               Width           =   2820
            End
            Begin VB.CheckBox chk 
               Caption         =   "סԺ���ʹ����м��ʺ��ӡ����"
               Height          =   270
               Index           =   138
               Left            =   150
               TabIndex        =   601
               Top             =   285
               Width           =   2835
            End
            Begin VB.CheckBox chk 
               Caption         =   "���ۺ��ӡ���ۼ��ʵ�"
               Height          =   195
               Index           =   141
               Left            =   150
               TabIndex        =   604
               Top             =   1095
               Width           =   2325
            End
            Begin VB.CheckBox chk 
               Caption         =   "���۵���˺��ӡ���ʵ�"
               Height          =   195
               Index           =   142
               Left            =   150
               TabIndex        =   605
               Top             =   1365
               Width           =   2850
            End
         End
         Begin VB.Frame fraJZUnit 
            Caption         =   " ҩƷ��λ "
            Height          =   930
            Left            =   195
            TabIndex        =   596
            Top             =   2640
            Width           =   4950
            Begin VB.OptionButton optJZDrugUnit 
               Caption         =   "סԺ��λ"
               Height          =   180
               Index           =   1
               Left            =   2985
               TabIndex        =   599
               Top             =   435
               Width           =   1020
            End
            Begin VB.OptionButton optJZDrugUnit 
               Caption         =   "�ۼ۵�λ"
               Height          =   180
               Index           =   0
               Left            =   1410
               TabIndex        =   598
               Top             =   435
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.Label lbl��λ 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ��"
               Height          =   180
               Index           =   4
               Left            =   480
               TabIndex        =   597
               Top             =   435
               Width           =   720
            End
         End
         Begin VB.Frame fra 
            Height          =   2295
            Index           =   19
            Left            =   195
            TabIndex        =   584
            Top             =   195
            Width           =   4950
            Begin VB.CheckBox chk 
               Caption         =   "����δ��ƽ�ֹ���˲���"
               Height          =   210
               Index           =   84
               Left            =   2460
               TabIndex        =   590
               Top             =   880
               Width           =   2340
            End
            Begin VB.CheckBox chk 
               Caption         =   "�������뿪����"
               Height          =   210
               Index           =   18
               Left            =   2460
               TabIndex        =   586
               Top             =   262
               Width           =   1590
            End
            Begin VB.CheckBox chk 
               Caption         =   "���������������ҵĿ�����"
               Height          =   210
               Index           =   19
               Left            =   2460
               TabIndex        =   588
               Top             =   571
               Width           =   2460
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   0
               Left            =   1470
               Style           =   2  'Dropdown List
               TabIndex        =   595
               Top             =   1800
               Width           =   1140
            End
            Begin VB.CheckBox chk 
               Caption         =   "Ƿ��ʱ������Ϊ���۵�"
               Height          =   195
               Index           =   146
               Left            =   195
               TabIndex        =   593
               Top             =   1498
               Width           =   2400
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ҩ�������븶��"
               Height          =   195
               Index           =   143
               Left            =   195
               TabIndex        =   585
               Top             =   270
               Value           =   1  'Checked
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "�������а�����ʿ"
               Height          =   195
               Index           =   145
               Left            =   195
               TabIndex        =   589
               Top             =   884
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "���������������"
               Height          =   195
               Index           =   144
               Left            =   195
               TabIndex        =   587
               Top             =   577
               Width           =   1740
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʾ����ҩ����"
               Height          =   195
               Index           =   148
               Left            =   2460
               TabIndex        =   592
               Top             =   1191
               Width           =   1850
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ʾ����ҩ�����"
               Height          =   195
               Index           =   147
               Left            =   195
               TabIndex        =   591
               Top             =   1191
               Width           =   1845
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "�ѽ���ʵ�����"
               Height          =   180
               Index           =   1
               Left            =   195
               TabIndex        =   594
               Top             =   1860
               Width           =   1260
            End
         End
         Begin VB.Label lbl��ҩ 
            AutoSize        =   -1  'True
            Caption         =   "����֮��"
            Height          =   180
            Left            =   240
            TabIndex        =   606
            Top             =   5715
            Width           =   720
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7680
         Index           =   3
         Left            =   -74850
         ScaleHeight     =   7650
         ScaleWidth      =   9030
         TabIndex        =   377
         Top             =   105
         Visible         =   0   'False
         Width           =   9060
         Begin VB.OptionButton optInExseCharge 
            Caption         =   "����סԺ���۷���"
            Height          =   180
            Index           =   1
            Left            =   4725
            TabIndex        =   29
            Top             =   7395
            Width           =   1860
         End
         Begin VB.OptionButton optInExseCharge 
            Caption         =   "������סԺ���۷���"
            Height          =   180
            Index           =   0
            Left            =   6690
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   7395
            Width           =   2070
         End
         Begin VB.Frame fra 
            Height          =   45
            Index           =   20
            Left            =   -15
            TabIndex        =   669
            Top             =   7260
            Width           =   10000
         End
         Begin VB.CommandButton cmdWarnDel 
            Caption         =   "ɾ����������(&D)"
            Height          =   350
            Left            =   7845
            TabIndex        =   26
            Top             =   6825
            Width           =   1590
         End
         Begin VB.CommandButton cmdWarnNew 
            Caption         =   "���ӱ�������(&A)"
            Height          =   350
            Left            =   7845
            TabIndex        =   25
            Top             =   6465
            Width           =   1590
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����סԺ���ʱ����������۷���"
            Height          =   255
            Index           =   41
            Left            =   90
            TabIndex        =   28
            Top             =   7365
            Width           =   3090
         End
         Begin VB.ListBox lst��� 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   2130
            Left            =   2895
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   975
            Visible         =   0   'False
            Width           =   1530
         End
         Begin ZL9BillEdit.BillEdit Bill 
            Height          =   5370
            Index           =   1
            Left            =   210
            TabIndex        =   22
            Top             =   765
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   9472
            CellAlignment   =   9
            Text            =   ""
            TextMatrix0     =   ""
            MaxDate         =   2958465
            MinDate         =   -53688
            Value           =   36395
            Cols            =   3
            RowHeight0      =   315
            RowHeightMin    =   315
            ColWidth0       =   1005
            BackColor       =   -2147483643
            BackColorBkg    =   -2147483643
            BackColorSel    =   10249818
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            ForeColorSel    =   -2147483634
            GridColor       =   -2147483630
            ColAlignment0   =   9
            ListIndex       =   -1
            CellBackColor   =   -2147483643
         End
         Begin MSComctlLib.TabStrip tab���� 
            Height          =   5955
            Left            =   90
            TabIndex        =   21
            Top             =   405
            Width           =   9345
            _ExtentX        =   16484
            _ExtentY        =   10504
            HotTracking     =   -1  'True
            TabMinWidth     =   0
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "��ͨ����"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lblInExseCharge 
            AutoSize        =   -1  'True
            Caption         =   "סԺ���ʱ���"
            Height          =   180
            Left            =   3555
            TabIndex        =   670
            Top             =   7395
            Width           =   1080
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmParFee.frx":1BDAF
            Height          =   555
            Left            =   90
            TabIndex        =   31
            Top             =   6495
            Width           =   7740
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��������:ÿ�ַ������������������߼�������ʽ����� zl_PatiWarnScheme �������ʹ��"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   14
            Left            =   210
            TabIndex        =   27
            Top             =   120
            Width           =   7200
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   6795
         Index           =   15
         Left            =   -74760
         ScaleHeight     =   6765
         ScaleWidth      =   8940
         TabIndex        =   686
         TabStop         =   0   'False
         Top             =   720
         Width           =   8970
         Begin VB.PictureBox picSettlePar 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   8415
            Index           =   0
            Left            =   825
            ScaleHeight     =   8385
            ScaleWidth      =   9270
            TabIndex        =   688
            TabStop         =   0   'False
            Top             =   75
            Visible         =   0   'False
            Width           =   9300
            Begin VB.Frame fraOrder 
               Caption         =   "ȱʡ��Ԥ������"
               Height          =   2520
               Index           =   0
               Left            =   180
               TabIndex        =   744
               Top             =   4830
               Width           =   9075
               Begin VB.CommandButton cmdDepositUp 
                  Caption         =   "��"
                  Height          =   510
                  Left            =   8280
                  TabIndex        =   749
                  Top             =   1065
                  Width           =   375
               End
               Begin VB.CommandButton cmdDepositDown 
                  Caption         =   "��"
                  Height          =   510
                  Left            =   8280
                  TabIndex        =   748
                  Top             =   1665
                  Width           =   375
               End
               Begin VB.OptionButton optOrder 
                  Caption         =   "�����������г�Ԥ��"
                  Height          =   300
                  Index           =   1
                  Left            =   135
                  TabIndex        =   746
                  Top             =   495
                  Width           =   3585
               End
               Begin VB.OptionButton optOrder 
                  Caption         =   "���ɿ�ʱ���Ⱥ��Ԥ��"
                  Height          =   300
                  Index           =   0
                  Left            =   135
                  TabIndex        =   745
                  Top             =   225
                  Width           =   3645
               End
               Begin VSFlex8Ctl.VSFlexGrid vsDepositSort 
                  Height          =   1635
                  Left            =   165
                  TabIndex        =   747
                  Top             =   810
                  Width           =   8010
                  _cx             =   14129
                  _cy             =   2884
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
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483634
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483632
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   2
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   2
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   5
                  Cols            =   5
                  FixedRows       =   2
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   300
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":1BE90
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   4
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
                  ExplorerBar     =   8
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
            Begin VB.Frame fraBalanceFeeDate 
               Caption         =   "���ʷ����ڼ�����"
               ForeColor       =   &H00000000&
               Height          =   930
               Left            =   5205
               TabIndex        =   719
               Top             =   2145
               Width           =   2025
               Begin VB.OptionButton optBalanceTime 
                  Caption         =   "���Ǽ�ʱ��"
                  Height          =   195
                  Index           =   0
                  Left            =   195
                  TabIndex        =   721
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   1320
               End
               Begin VB.OptionButton optBalanceTime 
                  Caption         =   "������ʱ��"
                  Height          =   195
                  Index           =   1
                  Left            =   195
                  TabIndex        =   720
                  Top             =   600
                  Width           =   1320
               End
            End
            Begin VB.Frame frabalance 
               Caption         =   "��Ժ���ʴ��տ���"
               Height          =   900
               Index           =   0
               Left            =   5205
               TabIndex        =   716
               Top             =   120
               Width           =   2025
               Begin VB.OptionButton optBalanceDSCheck 
                  Caption         =   "��ֹ"
                  Height          =   195
                  Index           =   0
                  Left            =   195
                  TabIndex        =   718
                  Top             =   285
                  Width           =   870
               End
               Begin VB.OptionButton optBalanceDSCheck 
                  Caption         =   "��ʾ"
                  Height          =   195
                  Index           =   1
                  Left            =   195
                  TabIndex        =   717
                  Top             =   585
                  Value           =   -1  'True
                  Width           =   870
               End
            End
            Begin VB.Frame fraBalaceBlood 
               Caption         =   "����ʱ��Ѫ�Ѽ��"
               Height          =   900
               Left            =   5205
               TabIndex        =   713
               Top             =   1125
               Width           =   2025
               Begin VB.OptionButton optBlood 
                  Caption         =   "��鲢��ʾ"
                  Height          =   210
                  Index           =   1
                  Left            =   195
                  TabIndex        =   715
                  Top             =   570
                  Width           =   1305
               End
               Begin VB.OptionButton optBlood 
                  Caption         =   "�����"
                  Height          =   210
                  Index           =   0
                  Left            =   195
                  TabIndex        =   714
                  Top             =   300
                  Value           =   -1  'True
                  Width           =   945
               End
            End
            Begin VB.Frame fraMzDepositDefaultUse 
               Caption         =   "����Ԥ��ȱʡʹ�÷�ʽ"
               Height          =   1275
               Left            =   5205
               TabIndex        =   709
               Top             =   3120
               Width           =   4065
               Begin VB.OptionButton optMzDeposit 
                  Caption         =   "�����ʽ��ʹ��Ԥ��"
                  Height          =   350
                  Index           =   1
                  Left            =   195
                  TabIndex        =   712
                  Top             =   540
                  Width           =   2256
               End
               Begin VB.OptionButton optMzDeposit 
                  Caption         =   "��ʹ��Ԥ����"
                  Height          =   350
                  Index           =   0
                  Left            =   195
                  TabIndex        =   711
                  Top             =   255
                  Width           =   1524
               End
               Begin VB.OptionButton optMzDeposit 
                  Caption         =   "ʹ��ʣ������Ԥ����"
                  Height          =   350
                  Index           =   2
                  Left            =   195
                  TabIndex        =   710
                  Top             =   840
                  Value           =   -1  'True
                  Width           =   2028
               End
            End
            Begin VB.Frame fraOwnFeeType 
               Caption         =   "�����Ƚ��Է����"
               Height          =   2940
               Left            =   7335
               TabIndex        =   707
               Top             =   150
               Width           =   1950
               Begin VB.ListBox lst 
                  Height          =   2580
                  Index           =   4
                  Left            =   75
                  Style           =   1  'Checkbox
                  TabIndex        =   708
                  Top             =   315
                  Width           =   1770
               End
            End
            Begin VB.Frame frabalance 
               Caption         =   "���ʽɿ����"
               Height          =   1245
               Index           =   3
               Left            =   180
               TabIndex        =   703
               Top             =   3495
               Width           =   4800
               Begin VB.OptionButton optBalancePayin 
                  Caption         =   "�����нɿ����"
                  Height          =   180
                  Index           =   0
                  Left            =   255
                  TabIndex        =   706
                  Top             =   330
                  Value           =   -1  'True
                  Width           =   1770
               End
               Begin VB.OptionButton optBalancePayin 
                  Caption         =   "������ȡ�ֽ�ʱ,��������ɿ�"
                  Height          =   180
                  Index           =   1
                  Left            =   255
                  TabIndex        =   705
                  Top             =   615
                  Width           =   2835
               End
               Begin VB.OptionButton optBalancePayin 
                  Caption         =   "���������ۼ�"
                  Height          =   180
                  Index           =   2
                  Left            =   255
                  TabIndex        =   704
                  Top             =   900
                  Width           =   1470
               End
            End
            Begin VB.Frame frabalance 
               Height          =   2970
               Index           =   4
               Left            =   180
               TabIndex        =   690
               Top             =   435
               Width           =   4800
               Begin VB.CheckBox chk 
                  Caption         =   "�����������˿�Ĭ������"
                  Height          =   225
                  Index           =   200
                  Left            =   120
                  TabIndex        =   788
                  Top             =   2265
                  Width           =   3195
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�Էѷ��ý���ȱʡʹ��Ԥ��"
                  Height          =   255
                  Index           =   157
                  Left            =   120
                  TabIndex        =   701
                  Top             =   1660
                  Width           =   2640
               End
               Begin VB.CheckBox chk 
                  Caption         =   "ҽ���Ƚ��Էѷ��ò���ӡ����Ʊ��"
                  Height          =   210
                  Index           =   156
                  Left            =   120
                  TabIndex        =   700
                  Top             =   1367
                  Width           =   3105
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���ʼ�鲡���������"
                  Height          =   255
                  Index           =   155
                  Left            =   120
                  TabIndex        =   699
                  Top             =   751
                  Width           =   2190
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���ʺ����������Ϣ"
                  Height          =   225
                  Index           =   154
                  Left            =   2490
                  TabIndex        =   698
                  Top             =   458
                  Width           =   2175
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��;����ȱʡ��Ԥ����"
                  Height          =   195
                  Index           =   153
                  Left            =   120
                  TabIndex        =   697
                  Top             =   473
                  Width           =   2160
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��ʹ��ָ��סԺ������Ԥ����"
                  Height          =   195
                  Index           =   152
                  Left            =   120
                  TabIndex        =   696
                  Top             =   1089
                  Width           =   2760
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�Բ��˵�����ý��н���"
                  Height          =   195
                  Index           =   151
                  Left            =   2490
                  TabIndex        =   695
                  Top             =   195
                  Width           =   2280
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���˳�Ժ���ʺ��Զ���Ժ"
                  Height          =   195
                  Index           =   150
                  Left            =   120
                  TabIndex        =   694
                  Top             =   195
                  Width           =   2280
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��Լ��λ����ÿλ���˷ֱ��ӡƱ��"
                  Height          =   225
                  Index           =   149
                  Left            =   120
                  TabIndex        =   693
                  Top             =   1998
                  Width           =   3195
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��Ժ���˲������Ժ����"
                  Height          =   210
                  Index           =   55
                  Left            =   2490
                  TabIndex        =   692
                  Top             =   773
                  Width           =   2280
               End
               Begin VB.ComboBox cbo 
                  Height          =   300
                  Index           =   5
                  ItemData        =   "frmParFee.frx":1BFB6
                  Left            =   1320
                  List            =   "frmParFee.frx":1BFB8
                  Style           =   2  'Dropdown List
                  TabIndex        =   691
                  Top             =   2565
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "δ�󵥾ݽ���"
                  Height          =   180
                  Index           =   24
                  Left            =   120
                  TabIndex        =   702
                  Top             =   2625
                  Width           =   1080
               End
            End
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   7
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   689
               Top             =   135
               Width           =   2550
            End
            Begin VB.Label lblUnit 
               AutoSize        =   -1  'True
               Caption         =   "��Լ��λ����ʹ��                             ��Ʊ��"
               Height          =   180
               Left            =   210
               TabIndex        =   722
               Top             =   195
               Width           =   4590
            End
         End
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   975
            Index           =   2
            Left            =   0
            TabIndex        =   687
            TabStop         =   0   'False
            Top             =   0
            Width           =   5025
            _Version        =   589884
            _ExtentX        =   8864
            _ExtentY        =   1720
            _StockProps     =   64
         End
         Begin VB.PictureBox picSettlePar 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   7275
            Index           =   2
            Left            =   -360
            ScaleHeight     =   7245
            ScaleWidth      =   9780
            TabIndex        =   760
            TabStop         =   0   'False
            Top             =   3720
            Visible         =   0   'False
            Width           =   9810
            Begin VB.Frame fraColor 
               Caption         =   "���ʽɿ����������ɫ"
               Height          =   2085
               Left            =   6270
               TabIndex        =   767
               Top             =   3885
               Width           =   3285
               Begin VB.PictureBox pic�ɿ����˿�ɫ 
                  BackColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2415
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   772
                  Top             =   1665
                  Width           =   705
               End
               Begin VB.PictureBox pic��ǰ����δ��ɫ 
                  BackColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2415
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   771
                  Top             =   915
                  Width           =   705
               End
               Begin VB.PictureBox pic�ɿ����ɿ�ɫ 
                  BackColor       =   &H00FF0000&
                  Height          =   300
                  Left            =   2415
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   770
                  Top             =   1305
                  Width           =   705
               End
               Begin VB.PictureBox pic��ǰ����δ��ɫ 
                  BackColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2415
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   769
                  Top             =   555
                  Width           =   705
               End
               Begin VB.PictureBox pic�Ը��ϼ�ɫ 
                  BackColor       =   &H00FF0000&
                  Height          =   300
                  Left            =   1635
                  ScaleHeight     =   240
                  ScaleWidth      =   645
                  TabIndex        =   768
                  Top             =   225
                  Width           =   705
               End
               Begin VB.Label lbl����Color 
                  AutoSize        =   -1  'True
                  Caption         =   "�˿���ɫ         "
                  Height          =   180
                  Index           =   4
                  Left            =   1665
                  TabIndex        =   777
                  Top             =   1725
                  Width           =   1530
               End
               Begin VB.Label lbl����Color 
                  AutoSize        =   -1  'True
                  Caption         =   "δ����ɫ         "
                  Height          =   180
                  Index           =   2
                  Left            =   1665
                  TabIndex        =   776
                  Top             =   975
                  Width           =   1530
               End
               Begin VB.Label lbl����Color 
                  AutoSize        =   -1  'True
                  Caption         =   "�ɿ���������ɫ:�ɿ���ɫ         "
                  Height          =   180
                  Index           =   3
                  Left            =   315
                  TabIndex        =   775
                  Top             =   1365
                  Width           =   2880
               End
               Begin VB.Label lbl����Color 
                  AutoSize        =   -1  'True
                  Caption         =   "��ǰ����������ɫ:δ����ɫ         "
                  Height          =   180
                  Index           =   0
                  Left            =   135
                  TabIndex        =   774
                  Top             =   600
                  Width           =   3060
               End
               Begin VB.Label lbl����Color 
                  AutoSize        =   -1  'True
                  Caption         =   "�Ը��ϼ�������ɫ"
                  Height          =   180
                  Index           =   1
                  Left            =   135
                  TabIndex        =   773
                  Top             =   270
                  Width           =   1440
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "���˶�ν��ʵ���������������"
               Height          =   195
               Index           =   191
               Left            =   2010
               TabIndex        =   765
               Top             =   3615
               Width           =   4125
            End
            Begin VB.PictureBox picDisplay 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   3240
               Index           =   0
               Left            =   1590
               Picture         =   "frmParFee.frx":1BFBA
               ScaleHeight     =   3210
               ScaleWidth      =   4515
               TabIndex        =   763
               Top             =   3975
               Width           =   4545
            End
            Begin VB.OptionButton opt������ 
               Caption         =   "���ʽɿ���"
               Height          =   525
               Index           =   1
               Left            =   75
               TabIndex        =   762
               Top             =   5333
               Width           =   1515
            End
            Begin VB.OptionButton opt������ 
               Caption         =   "��ͳ���ʷ��"
               Height          =   525
               Index           =   0
               Left            =   75
               TabIndex        =   761
               Top             =   1560
               Width           =   1455
            End
            Begin VB.PictureBox picDisplay 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   3240
               Index           =   1
               Left            =   1590
               Picture         =   "frmParFee.frx":22690
               ScaleHeight     =   3210
               ScaleWidth      =   4515
               TabIndex        =   764
               Top             =   195
               Width           =   4545
            End
         End
         Begin VB.PictureBox picSettlePar 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6570
            Index           =   1
            Left            =   4500
            ScaleHeight     =   6540
            ScaleWidth      =   9780
            TabIndex        =   723
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   9810
            Begin VB.Frame fraƱ�ݸ�ʽ 
               Caption         =   "���ʺ�Ʊ��ӡ����"
               Height          =   1620
               Index           =   7
               Left            =   4875
               TabIndex        =   742
               Top             =   2235
               Width           =   4515
               Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
                  Height          =   1275
                  Index           =   7
                  Left            =   105
                  TabIndex        =   743
                  Top             =   255
                  Width           =   4320
                  _cx             =   7620
                  _cy             =   2249
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
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   8421504
                  GridColorFixed  =   8421504
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
                  Rows            =   3
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":27BC0
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
                  ExplorerBar     =   2
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
            Begin VB.Frame fraƱ�ݸ�ʽ 
               Caption         =   "����Ʊ�ݴ�ӡ����"
               Height          =   1740
               Index           =   3
               Left            =   4875
               TabIndex        =   740
               Top             =   240
               Width           =   4515
               Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
                  Height          =   1350
                  Index           =   3
                  Left            =   105
                  TabIndex        =   741
                  Top             =   255
                  Width           =   4305
                  _cx             =   7594
                  _cy             =   2381
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
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   8421504
                  GridColorFixed  =   8421504
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
                  Rows            =   3
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParFee.frx":27C2D
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
                  ExplorerBar     =   2
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
            Begin VB.Frame frabalance 
               Caption         =   "�����˿��վݴ�ӡ��ʽ"
               Height          =   510
               Index           =   1
               Left            =   75
               TabIndex        =   736
               Top             =   2235
               Width           =   4725
               Begin VB.OptionButton optDelBalancePrint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   1500
                  TabIndex        =   739
                  Top             =   240
                  Width           =   1065
               End
               Begin VB.OptionButton optDelBalancePrint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   2850
                  TabIndex        =   738
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.OptionButton optDelBalancePrint 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   255
                  TabIndex        =   737
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   900
               End
            End
            Begin VB.Frame fraOwnFee 
               Caption         =   "�Էѷ����嵥��ӡ"
               Height          =   480
               Left            =   75
               TabIndex        =   732
               Top             =   240
               Width           =   4725
               Begin VB.OptionButton optOwnFee 
                  Caption         =   "����ӡ"
                  Height          =   255
                  Index           =   0
                  Left            =   285
                  TabIndex        =   735
                  Top             =   195
                  Width           =   1050
               End
               Begin VB.OptionButton optOwnFee 
                  Caption         =   "�Զ���ӡ"
                  Height          =   255
                  Index           =   1
                  Left            =   1500
                  TabIndex        =   734
                  Top             =   195
                  Width           =   1050
               End
               Begin VB.OptionButton optOwnFee 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   255
                  Index           =   2
                  Left            =   2850
                  TabIndex        =   733
                  Top             =   195
                  Width           =   1395
               End
            End
            Begin VB.Frame frabalance 
               Caption         =   "���ʷ�����ϸ��ӡ��ʽ"
               Height          =   495
               Index           =   2
               Left            =   75
               TabIndex        =   728
               Top             =   885
               Width           =   4725
               Begin VB.OptionButton optBalanceFeeListPrint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   180
                  Index           =   2
                  Left            =   1500
                  TabIndex        =   731
                  Top             =   225
                  Width           =   1065
               End
               Begin VB.OptionButton optBalanceFeeListPrint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   180
                  Index           =   1
                  Left            =   2850
                  TabIndex        =   730
                  Top             =   225
                  Width           =   1455
               End
               Begin VB.OptionButton optBalanceFeeListPrint 
                  Caption         =   "����ӡ"
                  Height          =   180
                  Index           =   0
                  Left            =   255
                  TabIndex        =   729
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   900
               End
            End
            Begin VB.Frame fraDeposit 
               Caption         =   "Ԥ��Ʊ�ݴ�ӡ"
               Height          =   510
               Left            =   75
               TabIndex        =   724
               Top             =   1545
               Width           =   4725
               Begin VB.OptionButton optBalanceDepositPrint 
                  Caption         =   "����ӡ"
                  Height          =   255
                  Index           =   0
                  Left            =   255
                  TabIndex        =   727
                  Top             =   210
                  Width           =   1110
               End
               Begin VB.OptionButton optBalanceDepositPrint 
                  Caption         =   "�Զ���ӡ"
                  Height          =   255
                  Index           =   1
                  Left            =   1500
                  TabIndex        =   726
                  Top             =   210
                  Width           =   1335
               End
               Begin VB.OptionButton optBalanceDepositPrint 
                  Caption         =   "ѡ���Ƿ��ӡ"
                  Height          =   255
                  Index           =   2
                  Left            =   2850
                  TabIndex        =   725
                  Top             =   210
                  Width           =   1395
               End
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmParFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset '������ؼ���Ӧ��¼����ͬһ���������ܶ�Ӧһ�����ؼ���
Private marrFunc(2) As String
Private mlngPreFind As Long
Private mblnInstantActive As Boolean
Private mblnNotChange As Boolean
Private mblnExistPrintData As Boolean '�շѴ��ڴ�ӡ����
Private mintԭƱ�ݷ������ As Integer
Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk
    chk_��Ժ���˲�׼��Ժ���� = 55
    
    chk_���뿪���� = 18
    chk_���ƿ����� = 19
    chk_δ��ƽ�ֹ���� = 84
    
    chk_�������� = 7
    chk_����ID = 8
    chk_ˢ���￨ = 9
    chk_�Һŵ��� = 10
    
    chk_��Ժ������������תסԺ = 183
    chk_�������תסԺԤ��Ʊ�ݴ�ӡ���� = 197
    
    chk_ר�ҺŹҺ����� = 178
    chk_ר�Һ�ԤԼ���� = 179
    chk_������Һ�ģʽ = 290
    
    chk_ԤԼ�ŶӰ�ʱ�� = 186
        
    chk_�շ���Ŀ��λ�������� = 56
    chk_���������շ���� = 25
    chk_�����˷��������� = 16
    chk_������Ŀ���ܼ����ۿ� = 39
    chk_ָ�����ϲ���ʱ����ʾ�޿������ = 204
    
    'һ��ͨ
    chk_��Ŀִ��ǰ�����շѻ���� = 67
    chk_��Ŀ�����������շѻ������� = 90
    
    
    chk_Ʊ�ſ��� = 13
    
    chk_���ʱ����������۷��� = 41
    
    chk_�Զ����� = 12
    chk_���������ģʽ = 43
    
    chk_�Һű���ˢ�� = 102
    chk_����ʹ��Ԥ�� = 103
    chk_����סԺ���˹Һ� = 104
    chk_������ѡ�� = 105
    chk_����ģ������ = 106
    chk_�ҺŰ������Ұ��� = 107
    chk_ԤԼ�������Ұ��� = 184
    chk_ԤԼʱ�տ� = 108
    chk_ҽ��վ����ҽ�� = 169
    
    chk_����_���������������� = 187
    
    
    'Ԥ�������
    chk_��Ԥ�������������Ϣ = 0
    chk_ֻ��ʾ��ʣ�����ʷ�ɿ� = 1
    chk_��Ժ����δ��Ʋ�׼��Ԥ�� = 2
    chk_������Ĳ��˵Ľɿ���� = 3
    chk_�����Ժ���˽�סԺԤ�� = 4
    chk_��ֹ��Ժ���˽�����Ԥ�� = 57
    chk_����ͨ������ģ�����Ҳ��� = 5
    chk_��סԺԤ��ˢ����֤ = 6
    chk_Ʊ��ʣ��N�����Ѳ���Ա = 11
    chk_������Ժ��������˿� = 188
    chk_Ԥ�����վ����ʾ = 202
    
    'ҽ�ƿ�����
    chk_�����Լ��˷�ʽ��ȡ = 14
    chk_��������ģ������ = 15
    chk_����ʹ�������շ�ҽ���վ� = 192
    chk_��ȡ������ = 199
    chk_ҽ�ƿ�_�Զ���������� = 201
    
    '�ҺŰ���
    chk_ֻ��ҽ��ҽ�����йҺŰ��� = 17
    chk_����ԤԼ�Һŵ���ֹɾ������ = 21
    '�ҺŹ���
    chk_�Һ�_�Զ�ˢ�¹ҺŰ��� = 22
    chk_�Һ�_�Զ���������� = 23
    chk_�Һ�_��Ϊ���۵� = 24
    chk_�Һ�_������Һ�һ���� = 26
    chk_�Һ�_��ҽ����������ҽ�� = 27
    chk_�Һ�_�����Զ������²��� = 29
    chk_�Һ�_����ʹ��Ԥ����ɷ� = 30
    chk_�Һ�_����ģ������ = 31
    chk_�Һ�_����ô�ӡƱ�� = 32
    chk_�Һ�_����_���� = 33
    chk_�Һ�_����_�Ա� = 34
    chk_�Һ�_����_���� = 35
    chk_�Һ�_����_��ͥ��ַ = 36
    chk_�Һ�_����_���ʽ = 37
    chk_�Һ�_����_�ѱ� = 38
    chk_�Һ�_����_���㷽ʽ = 40
    chk_�Һ�_����_��ϵ�绰 = 20
    chk_�Һ�_�����������˵ǼǴ��� = 42
    chk_�Һ�_�˺Ų��˿��ش�Ʊ�� = 44
    chk_�Һ�_�Һź��ӡ������ǩ = 45
    chk_�Һ�_ԤԼ��ʾ���кű� = 46
    chk_�Һ�_ԤԼ����ȷ���Һŷ� = 48
    chk_�Һ�_����סԺ���˹Һ� = 49
    chk_�Һ�_ԤԼ����������� = 50
    chk_�Һ�_������ѡ�� = 51
    chk_�Һ�_N�����˺������ = 52
    chk_�Һ�_ԤԼʧԼ���ڹҺ� = 53
    chk_�Һ�_��ͥ��ַ�������� = 54
    chk_�Һ�_ɨ�����֤ǩԼ = 58
    chk_�Һ�_�����������Һ� = 60
    chk_�Һ�_�ϸ�ʱ�ιҺ� = 61
    chk_�Һ�_Ĭ�Ϲ�ѡ����ѡ�� = 62
    chk_�Һ�_�������Ч�Լ�� = 63
    chk_�Һ�_���ϸ����Ϊ���� = 64
    chk_�Һ�_����ԤԼ������ = 47
    chk_�Һ�_����ͬ����ԼN���� = 172
    chk_�Һ�_���˹Һſ������� = 176
    chk_�Һ�_����ͬ���޹�N���� = 174
    chk_�Һ�_����ͬ���޹�N����_���� = 175
    chk_�Һ�_�ƻ��Ű�ģʽ������� = 185
    chk_�Һ�_��ֹ�������� = 190
    chk_�Һ�_ͬһ���ֻ֤�ܶ�Ӧһ���������� = 28
    chk_�Һ�_����ͬһ��Դ�޹�N���� = 203
    
    '����
    chk_����_����̨ǩ����ʼ�Ŷ� = 65
    chk_����_ԤԼ�ҺŽ������ = 66
    chk_����_����æ������� = 68
    chk_����_����ǩ�� = 171
    chk_����_����ǩ�� = 177
    
    '�ٴ����ﰲ��
    chk_���ﰲ��_����ҽ�������� = 182
    chk_���ﰲ��_������ҽ��ͬ������ԤԼ�Һŵ� = 180
    chk_���ﰲ��_ֻ����ѡԺ��ҽ�� = 181
    
    '���ﻮ��
    chk_����_����������ҩ���� = 69
    chk_����_�������������� = 74
    chk_����_���￪���˺���ʿ = 75
    chk_����_������ʾ����ҩ����� = 78
    chk_����_������ʾ����ҩ���� = 79
    chk_����_��������ģ������ = 85
    chk_����_��������_�Ա� = 88
    chk_����_��������_�Ƿ�Ӱ� = 89
    chk_����_��������_���� = 91
    chk_����_��������_�ѱ� = 92
    chk_����_��������_�������� = 93
    chk_����_��������_������ = 94
    chk_����_��������_ҽ�Ƹ��� = 95
    chk_����_���ﲻȱʡ������ = 110
    chk_����_�������Ҫ���뿪���� = 111
    chk_����_����ȱʡ�������� = 112
    chk_����_סԺ���˰������շ� = 168
    chk_����_����¼������ʹ�õĿ����� = 194
    
    
    '�����շ�
    chk_�շ�_����������ҩ���� = 70
    chk_�շ�_�������������� = 73
    chk_�շ�_���￪���˺���ʿ = 76
    chk_�շ�_������ʾ����ҩ����� = 80
    chk_�շ�_������ʾ����ҩ���� = 81
    chk_�շ�_��������ģ������ = 86
    chk_�շ�_��������_�Ա� = 97
    chk_�շ�_��������_�Ƿ�Ӱ� = 100
    chk_�շ�_��������_���� = 98
    chk_�շ�_��������_�ѱ� = 99
    chk_�շ�_��������_�������� = 101
    chk_�շ�_��������_������ = 109
    chk_�շ�_��������_ҽ�Ƹ��� = 96
    chk_�շ�_���ﲻȱʡ������ = 113
    chk_�շ�_�������Ҫ���뿪���� = 114
    chk_�շ�_����ȱʡ�������� = 115
    chk_�շ�_������ʾ�տ��ۼ� = 116
    chk_�շ�_�Ữ�۵����Ƥ�Խ�� = 117
    chk_�շ�_����ʹ��Ԥ����ɷ� = 118
    chk_�շ�_δ�Һ��Զ����չҺŷ� = 119
    chk_�շ�_����������ŵ��� = 120
    chk_�շ�_��Ѱ���۵��� = 121
    chk_�շ�_���������۵�ѡ�񴰿� = 122
    chk_�շ�_��鲡�˹Һſ��� = 123
    chk_�շ�_�Զ���ϵ��� = 124
    chk_�շ�_Ʊ��ʣ��X�ſ�ʼ���� = 125
    chk_�շ�_סԺ�������շ� = 126
    chk_�շ�_��ȡ���������ɿ� = 127
    chk_�շ�_���ŵ����շѷֱ��ӡ = 128
    chk_�շ�_ÿ��ֻ��һ��Ʊ�� = 129
    chk_�շ�_�Զ����չ����� = 130
    chk_�շ�_Ʊ�����ɷ�ʽ = 131
    chk_�շ�_ֻ��ҽ������ɹ������շ� = 170
    chk_�շ�_�����˲���Ʊ�ݲ��ִ��� = 173
    chk_�շ�_��첡�˰����ݷֱ��ӡ = 189
    chk_�շ�_����¼������ʹ�õĿ����� = 196
    
    '�������
    chk_����_����������ҩ���� = 71
    chk_����_�������������� = 72
    chk_����_���￪���˺���ʿ = 77
    chk_����_������ʾ����ҩ����� = 82
    chk_����_������ʾ����ҩ���� = 83
    chk_����_��������ģ������ = 87
    chk_����_ֻ���Һ�Լ��λ���� = 132
    chk_����_���ʴ�ӡ = 133
    chk_����_���۴�ӡ = 134
    chk_����_��˴�ӡ = 135
    chk_����_����¼������ʹ�õĿ����� = 195
    
    '������
    chk_������_����ģ������ = 136
    chk_������_Ʊ��ʣ��X�ſ�ʼ���� = 137
    'סԺ�������
    chk_סԺ����_���ʴ�ӡ = 138
    chk_��ɢ����_���ʴ�ӡ = 139
    chk_ҽ������_���ʴ�ӡ = 140
    chk_���ʲ���_���۴�ӡ = 141
    chk_���ʲ���_��˴�ӡ = 142
    chk_���ʲ���_��ҩ���븶�� = 143
    chk_���ʲ���_�����˰�����ʿ = 145
    chk_���ʲ���_����������� = 144
    chk_���ʲ���_Ƿ�ѱ��滮�۵� = 146
    chk_���ʲ���_��ʾ����ҩ����� = 147
    chk_���ʲ���_��ʾ����ҩ���� = 148

    
    '���˽��ʹ���
    chk_����_��Լ��λ�����˴�ӡ = 149
    chk_����_��Ժ���ʺ��Զ���Ժ = 150
    chk_����_��������� = 151
    chk_����_ʹ��ָ������Ԥ���� = 152
    chk_����_��;����ȱʡ��Ԥ���� = 153
    chk_����_���ʺ����������Ϣ = 154
    chk_����_���ʼ�鲡��������� = 155
    chk_����_�Էѷ��ò���ӡ����Ʊ�� = 156
    chk_����_�Էѷ���ȱʡʹ��Ԥ�� = 157
    chk_����_���˶�ν��ʵ��������������� = 191
    chk_����_�����������˿���� = 200
    
    'ִ�еǼǹ���
    chk_ִ�еǼ�_��ʾҽ�����͵ĵ��� = 158   '��������
    
    'ִ�еǼǹ���
    chk_�������_����תסԺ����� = 159   '��������
    'Ʊ��ʹ�ü��
    chk_Ʊ�ݼ��_����Ʊ��ǩ��ȷ�� = 160
    '��Ա������
    chk_���_�����ӡ = 161
    chk_���_�����ӡ = 162
    '���ѿ�����
    chk_���ѿ�_�ɿ��ӡ = 163
    chk_���ѿ�_ˢ�������붨λ������� = 193
    chk_���ѿ�_���ѿ��˷�ˢ������ = 59
    
    'ҽ�����ѹ���
    chk_����_��ҩ���븶�� = 164
    chk_����_����������� = 165
    chk_����_��ʾ����ҩ����� = 166
    chk_����_��ʾ����ҩ���� = 167
    
    '�շ�Ա���˹���
    chk_����_Ԥ�����ʰ������סԺ�ֱ����� = 198
    
End Enum


Private Enum constCbo
       
    cbo_�ѽᵥ�� = 0
    cbo_������˷�ʽ = 1
    cbo_�Һ�_ȱʡ����ʽ = 2
    cbo_�շ�_�Զ���ϵ��� = 3
    cbo_�շ�_Ʊ�ݷ������ = 4
    cbo_δ�󵥾ݽ��� = 5
    cbo_���ʲ���_���ʺ�ҩ = 6
    cbo_����_��Լ��λ���ʴ�ӡ = 7
    cbo_һ��ͨ_��ƱƱ�ݸ�ʽ = 8
    cbo_һ��ͨ_����Ʊ�ݸ�ʽ = 9
    
    cbo_�ٴ�����_ȫԺͨ�ú�Դ����վ�� = 18
    cbo_�ٴ�����_����ȽϷ�ʽ = 19
    cbo_�Һ�_ȱʡԤԼ��ʽ = 10
    cbo_�Һ�_ԤԼ��Чʱ�� = 11
    
    cbo_�Һ���Ǯ���� = 12
    cbo_�շ���Ǯ���� = 13
    cbo_������Ǯ���� = 14
    cbo_���ѿ���Ǯ���� = 17
    cbo_�Զ�����ģʽ = 15
    cbo_�������תסԺԤ����Ʊ��ʽ = 16
End Enum

Private Enum constUpDown
    ud_�Һŵ� = 1
    ud_����Һŵ� = 10
    
    ud_�Һ�ԤԼ���� = 0
    
    ud_���ý���λ�� = 5
    ud_���õ��۱���λ�� = 6
    
    ud_���볤�� = 4
    
    'Ԥ�����
    ud_Ʊ������ = 2
    
    '����
    ud_������Ч���� = 3
    
    '����
    ud_����_ȡ��N��Ļ��۵� = 7
    
    '�շ�
    ud_�շ�_Ʊ������ = 8
    ud_�շ�_�շ��վ����д� = 9
    
    '������
    ud_������_Ʊ������ = 11
    ud_������_��Ч���� = 12
    
    
End Enum

Private Enum constOpt

    'Ԥ�������:�˿�����
    opt_����ʱ�����˿� = 0
    opt_����ʱ��ֹ�˿� = 1

End Enum

Private Enum constBill
    bill_�Զ����� = 0
    bill_���ʱ��� = 1
End Enum

Private Enum constLvw
    lvw_Ʊ�� = 0
    lvw_���� = 1
    lvw_һ��ͨ = 3
End Enum

Private Enum constListBox
    lst_ҽ������ = 0
    lst_���Ѳ��� = 1
    lst_ˢ������ = 3
    lst_������_���㷽ʽ = 2
    lst_����_�Էѷ������ = 4
    lst_�ѻ�ҽ�����㷽ʽ = 5
End Enum

Private Enum constTxt
    txt_��������� = 0
    txt_������������ = 106
    txt_ר�ҺŹҺ����� = 4
    txt_ר�Һ�ԤԼ���� = 24
    txt_����ģ���������� = 1
    txt_�Һ�_ˢ��ʱ�� = 2
    txt_�Һ�_������������ = 3
    txt_�Һ�_ԤԼ��Чʱ�� = 5
    txt_�Һ�_ԤԼʧЧ���� = 6
    txt_�Һ�_N���ڲ���ȡ��ԤԼ�� = 7
    txt_�Һ�_ԤԼ����ʱ��_���� = 8
    txt_�Һ�_ԤԼ����ʱ��_���� = 9
    txt_�Һ�_N����������໤�� = 10
    txt_����_��ǰNСʱ���� = 11
    
    txt_����_���ﵥ������� = 12
    txt_����_��������ģ���������� = 14
    txt_�շ�_���ﵥ������� = 13
    txt_�շ�_��������ģ���������� = 15
    txt_����_��������ģ���������� = 16
    txt_�շ�_�������չҺŷ� = 17
    txt_�շ�_��Ѱ���۵������� = 18
    txt_������_����ģ���������� = 19
    
    txt_�Һ�_����ͬ���޹�N���� = 23
    txt_�Һ�_���˹Һſ������� = 20
    txt_�Һ�_����ͬ����ԼN���� = 22
    txt_�Һ�_����ԤԼ������ = 21
    txt_�Һ�_����ͬһ��Դ�޹�N���� = 25
    
    txt_����֧�� = 26
    
End Enum
Private Enum constVsGridBill
    vsGrid_Ԥ��Ʊ�ݸ�ʽ = 0
    vsGrid_�շ�Ʊ�ݸ�ʽ = 1
    vsGrid_������Ʊ�ݸ�ʽ = 2
    vsGrid_����Ʊ�ݸ�ʽ = 3
    vsGrid_Ԥ����Ʊ��ʽ = 4
    vsGrid_�˷�Ʊ�ݸ�ʽ = 5
    vsGrid_�������˷�Ʊ�ݸ�ʽ = 6
    vsGrid_���ʺ�Ʊ��ʽ = 7
    vsGrid_����Ԥ��Ʊ�ݸ�ʽ = 8
    vsGrid_ҽ�ƿ��վݸ�ʽ = 9
End Enum
Private Enum constVsGridInputItemSet
    vsGrid_�������������� = 0
End Enum

Private Enum constTbPage
    Pg_�Һ�ҵ�� = 0
    Pg_�����շ� = 1
    Pg_����ҵ�� = 2
End Enum
Private Enum constTbPageItemID
    Pg_�Һ�_���� = 100
    Pg_�Һ�_�Һ� = 101
    Pg_�Һ�_ԤԼ = 102
    Pg_�Һ�_���� = 103
    Pg_�շ�_���ݿ��� = 204
    Pg_�շ�_Ʊ�ݿ��� = 205
    Pg_����_���ʲ��� = 300
    Pg_����_Ʊ�ݿ��� = 301
    Pg_����_������ = 302
End Enum


'�Զ����������б��浱ǰ�к���
Private mintCurRow As Integer
Private mintCurCol As Integer
Private mblnJRaiseByDate As Boolean     '�жϴ�λ����Ŀ��������Ŀ�Ƿ��յ���
Private mblnHRaiseByDate As Boolean     '�жϻ�������Ŀ��������Ŀ�Ƿ��յ���
Private mstrDel���ò��� As String           '��¼���ʱ�����ɾ�������ò�������

Private mintColumn As Integer '

Private mrsWarn As ADODB.Recordset
Private mrs��� As ADODB.Recordset
Private mrsBillUseType As ADODB.Recordset
Private mblnOK As Boolean


Private Sub chkBillRule_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkBillRule_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chkBillRule, Index, mrsPar)
End Sub

 

Private Sub cmdDepClearAll_Click()
    Dim i As Integer
    If chk(chk_����_����̨ǩ����ʼ�Ŷ�).value = 0 Then Exit Sub
    With vsfTriageQueuingDep
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("����")) = 0
            If .RowData(i) <> .TextMatrix(i, .ColIndex("����")) Then
                .Cell(flexcpForeColor, i, .ColIndex("����")) = vbRed
            Else
                .Cell(flexcpForeColor, i, .ColIndex("����")) = &H80000008
            End If
        Next
    End With
End Sub

Private Sub cmdDepSelectAll_Click()
    Dim i As Integer
    If chk(chk_����_����̨ǩ����ʼ�Ŷ�).value = 0 Then Exit Sub
    With vsfTriageQueuingDep
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("����")) = 1
            If .RowData(i) <> .TextMatrix(i, .ColIndex("����")) Then
                .Cell(flexcpForeColor, i, .ColIndex("����")) = vbRed
            Else
                .Cell(flexcpForeColor, i, .ColIndex("����")) = &H80000008
            End If
        Next
    End With
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub


Private Sub cmdPrintSet_Click(Index As Integer)
    Select Case Index
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1802", Me)
    Case 2
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me)
    End Select
End Sub

Private Sub cmdInstantActive_Click()
    Dim datTime As Date, rsCheck As ADODB.Recordset
    Dim strSQL As String, strValue As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    strSQL = "Select 1 From �ٴ������ Where ����ʱ�� Is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        MsgBox "�������κ��ٴ���������,�����л�Ϊ������Ű�ģʽ!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    datTime = zlDatabase.Currentdate
    If MsgBox("ע��:" & vbCrLf & "��Ҫ�����л���������Ű�ģʽ,�Ѿ����ڵ�ԤԼ��¼�����Զ���������Ӧ���µĳ�����Ű���,����������Ҫһ��ʱ��,�����ĵȴ�" & _
                vbCrLf & "��ȷ���Ѿ����ռƻ��Ű����������˳�����Ű�����,���ԤԼ��¼û���ҵ���Ӧ�ĳ�����Ű�,�������ý���ʧ��!" & _
                vbCrLf & "�Ƿ��������ó�����Ű�ģʽ?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    strSQL = "Zl_�����Һ�_Turn(To_Date('" & Format(datTime, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')" & ")"
    Me.MousePointer = 11
    zlDatabase.ExecuteProcedure strSQL, "����ԤԼ��¼��������Ű�"
    Me.MousePointer = 0
    
    mblnInstantActive = True
    optRegistPlanMode(0).value = 0
    optRegistPlanMode(1).value = 1
    optRegistPlanMode(0).Enabled = False
    optRegistPlanMode(1).Enabled = False
    dtpRegistPlanMode.Enabled = False
    dtpRegistPlanMode.value = datTime
    dtpRegistPlanMode.Enabled = False
    cmdInstantActive.Enabled = False
    mblnInstantActive = False
    
    strValue = "1|" & Format(dtpRegistPlanMode.value, "yyyy-mm-dd hh:mm:ss")
    zlDatabase.SetPara "�Һ��Ű�ģʽ", strValue, glngSys
    
    fraNewPaln.Visible = True
    chk(chk_ֻ��ҽ��ҽ�����йҺŰ���).Visible = False
    chk(chk_����ԤԼ�Һŵ���ֹɾ������).Visible = False
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdStationRegOrder_Click(Index As Integer)
    Dim strValue As String, i As Integer
    With vsStationRegSort
        If Index = 0 Then
            If .Row <= 1 Then Exit Sub
            .RowPosition(.Row) = .Row - 1
            .Row = .Row - 1
        Else
            If .Row >= .Rows - 1 Then Exit Sub
            .RowPosition(.Row) = .Row + 1
            .Row = .Row + 1
        End If
    End With
    With vsStationRegSort
        For i = 1 To 5
            strValue = strValue & "|" & .TextMatrix(i, .ColIndex("�����ֶ�")) & "," & IIF(.TextMatrix(i, .ColIndex("�Ƿ�����")) = -1, 1, 0)
        Next i
        strValue = Mid(strValue, 2)
    End With
    Call SetParChange(vsStationRegSort, 0, mrsPar, True, strValue)
End Sub

Private Sub dtpRegistTime_LostFocus()
    Call SetParChange(dtpRegistTime, 0, mrsPar, True, Format(dtpRegistTime.value, "HH:mm:ss"))
End Sub

Private Sub dtpRegistTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(dtpRegistTime, 0, mrsPar)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "��ʼ�ɹ�" Then
        '��֪��Ϊ�Σ�ʹ��Դ���������½�ҳǩ��ȱʡ�󲻻ᴥ��Form_Activate�¼���������ǻᴥ����
        'Ϊ����Դ����������Ҳ�ܴ���Form_Activate�¼�����ҳǩȱʡ������Form_Activate�¼���
        tbPage(Pg_�Һ�ҵ��).Item(1).Selected = True   'ȱʡΪ��һ��,��Ҫ�ǳ���ԹҺŴ��ڽ�������
        tbPage(Pg_�����շ�).Item(0).Selected = True
        tbPage(Pg_����ҵ��).Item(0).Selected = True
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    Dim objPic As PictureBox
    
    For Each objPic In picPar
        Set objPic.Container = Me
    Next
    
    strCategory = "��������,������Ŀ"
    
    'ͼ����,TaskPanelItem��ID(ͬʱҲ�ǲ�������Picture�ؼ������),TaskPanelItem�ı���;......
    marrFunc(0) = "100,0,���ù�������"
    marrFunc(0) = marrFunc(0) & ";108,1,һ��ͨҵ��"
    marrFunc(0) = marrFunc(0) & ";112,8,ҽ�ƿ�ҵ��"
    marrFunc(0) = marrFunc(0) & ";107,7,Ԥ����ҵ��"
    marrFunc(0) = marrFunc(0) & ";106,6,���˹Һ�ҵ��"
    marrFunc(0) = marrFunc(0) & ";115,9,����̨ҵ��"
    marrFunc(0) = marrFunc(0) & ";101,10,���ﻮ�۹���"
    marrFunc(0) = marrFunc(0) & ";113,11,�����շѹ���"
    marrFunc(0) = marrFunc(0) & ";109,12,������ʹ���"
    marrFunc(0) = marrFunc(0) & ";114,13,���ղ������"
    marrFunc(0) = marrFunc(0) & ";110,14,סԺ����ҵ��"
    marrFunc(0) = marrFunc(0) & ";115,15,���˽��ʹ���"
    marrFunc(0) = marrFunc(0) & ";111,16,������ҵ��"
    marrFunc(0) = marrFunc(0) & ";116,17,ҽ�����ѹ���"
    
    marrFunc(1) = "105,5,���û�������;102,2,����Ա����;103,3,���ʱ���;104,4,�����Զ�����"
    
    '1.��ʼ���������һ�������б�,ȱʡѡ�е�һ��
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.��ʼ���������Ķ��������б�,ȱʡѡ�е�һ��
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    mblnOK = False
    Me.Tag = "��ʼ�ɹ�"
End Sub

Private Sub lblAvailabilityTimes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_�Һ�_ԤԼ��Чʱ��, mrsPar)
End Sub

Private Sub lblBespeakDefaultDays_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_�Һ�_ԤԼ����ʱ��_����, mrsPar)
End Sub
Private Sub lblBespeakMinTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_�Һ�_ԤԼ����ʱ��_����, mrsPar)
End Sub

 

Private Sub lblBreakAnAppointmentNums_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_�Һ�_ԤԼʧЧ����, mrsPar)
End Sub

Private Sub lblCancelBespeak_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_�Һ�_N���ڲ���ȡ��ԤԼ��, mrsPar)
End Sub

Private Sub lblDeptNums_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_�Һ�_����ԤԼ������, mrsPar)
End Sub


Private Sub lblGuardian_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_�Һ�_N����������໤��, mrsPar)
End Sub
Private Sub optBalanceDepositPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalanceDepositPrint, Index, mrsPar)
End Sub

Private Sub optBalanceDepositPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBalanceDepositPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalanceDepositPrint, Index, mrsPar)
End Sub
Private Sub optBalanceDSCheck_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalanceDSCheck, Index, mrsPar)
End Sub

Private Sub optBalanceDSCheck_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBalanceDSCheck_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalanceDSCheck, Index, mrsPar)
End Sub

Private Sub optBalancePayin_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalancePayin, Index, mrsPar)
End Sub

Private Sub optBalancePayin_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBalancePayin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalancePayin, Index, mrsPar)
End Sub
Private Sub optBalanceTime_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalanceTime, Index, mrsPar)
End Sub

Private Sub optBalanceTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBalanceTime_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalanceTime, Index, mrsPar)
End Sub


Private Sub optBillMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBillMode, Index, mrsPar)
End Sub

Private Sub optBillMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBillMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBillMode, Index, mrsPar)
End Sub

 
 
Private Sub optBalanceFeeListPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBalanceFeeListPrint, Index, mrsPar)
End Sub

Private Sub optBalanceFeeListPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub optBalanceFeeListPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBalanceFeeListPrint, Index, mrsPar)
End Sub


Private Sub optBlood_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBlood, Index, mrsPar)
End Sub

Private Sub optBlood_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub optBlood_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBlood, Index, mrsPar)
End Sub


Private Sub optBrushCard_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Dim strValue As String
    txt(txt_����֧��).Enabled = optBrushCard(3).value = True
    If optBrushCard(1).value Then
        strValue = 1
    ElseIf optBrushCard(2).value Then
        strValue = 2
    ElseIf optBrushCard(3).value Then
        If Val(txt(txt_����֧��).Text) = 0 Then Exit Sub
        strValue = -1 * Val(txt(txt_����֧��).Text)
    Else
        strValue = 0
    End If
    strValue = strValue & "|" & IIF(optBrushCard(11).value, 1, IIF(optBrushCard(12).value, 2, 0))
    Call SetParChange(optBrushCard, Index, mrsPar, True, strValue)
End Sub

Private Sub optBrushCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBrushCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBrushCard, Index, mrsPar)
End Sub

Private Sub optChargeExeBillPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optChargeExeBillPrint, Index, mrsPar)
End Sub

Private Sub optChargeExeBillPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    'zlCommFun.PressKey vbKeyTab
End Sub
Private Sub optChargeExeBillPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optChargeExeBillPrint, Index, mrsPar)
End Sub
Private Sub optChargeRegPrompt_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optChargeRegPrompt, Index, mrsPar)
End Sub

Private Sub optChargeRegPrompt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optChargeRegPrompt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optChargeRegPrompt, Index, mrsPar)
End Sub

 
Private Sub optDelBalancePrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDelBalancePrint, Index, mrsPar)
End Sub

Private Sub optDelBalancePrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDelBalancePrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDelBalancePrint, Index, mrsPar)
End Sub

Private Sub optDelFeeRefundPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDelFeeRefundPrint, Index, mrsPar)
End Sub

Private Sub optDelFeeRefundPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDelFeeRefundPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDelFeeRefundPrint, Index, mrsPar)
End Sub

 

Private Sub optDrugSupplementary_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDrugSupplementary, Index, mrsPar)
End Sub

Private Sub optDrugSupplementary_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDrugSupplementary_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDrugSupplementary, Index, mrsPar)
End Sub
Private Sub optDrugUnitFF_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDrugUnitFF, Index, mrsPar)
End Sub

Private Sub optDrugUnitFF_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDrugUnitFF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDrugUnitFF, Index, mrsPar)
End Sub

Private Sub optFeeListPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optFeeListPrint, Index, mrsPar)
End Sub

Private Sub optFeeListPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optFeeListPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optFeeListPrint, Index, mrsPar)
End Sub

Private Sub optJZDrugUnit_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optJZDrugUnit, Index, mrsPar)
End Sub

Private Sub optJZDrugUnit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optJZDrugUnit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optJZDrugUnit, Index, mrsPar)
End Sub
Private Sub optMzDeposit_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optMzDeposit, Index, mrsPar)
End Sub

Private Sub optMzDeposit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optMzDeposit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optMzDeposit, Index, mrsPar)
End Sub

Private Sub optOrder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optOrder, Index, mrsPar)
End Sub

Private Sub opt������_Click(Index As Integer)
    Dim i As Integer
    If opt������(0).value = True Then
        chk(chk_����_���˶�ν��ʵ���������������).Enabled = True
        fraColor.Enabled = False
        For i = 0 To 4
            lbl����Color(i).Enabled = False
        Next i
        Call SetParChange(opt������, Index, mrsPar, True, "0")
    Else
        chk(chk_����_���˶�ν��ʵ���������������).Enabled = False
        fraColor.Enabled = True
        For i = 0 To 4
            lbl����Color(i).Enabled = True
        Next i
        Call SetParChange(opt������, Index, mrsPar, True, "1")
    End If
End Sub

Private Sub opt������_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt������, Index, mrsPar)
End Sub

Private Sub optOrder_Click(Index As Integer)
    Dim strValue As String, i As Integer
    Dim intIndex As Integer
    If optOrder(0).value = True Then
        vsDepositSort.Enabled = False
        cmdDepositDown.Enabled = False
        cmdDepositUp.Enabled = False
        Call SetParChange(optOrder, Index, mrsPar, True, "0")
    Else
        vsDepositSort.Enabled = True
        cmdDepositDown.Enabled = True
        cmdDepositUp.Enabled = True
        strValue = "1|"
        With vsDepositSort
            For i = 2 To 4
                If Abs(Val(.TextMatrix(i, 2))) = 1 Then intIndex = 0
                If Abs(Val(.TextMatrix(i, 3))) = 1 Then intIndex = 1
                If Abs(Val(.TextMatrix(i, 4))) = 1 Then intIndex = 2
                If i <> 4 Then
                    strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex & ","
                Else
                    strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex
                End If
            Next i
        End With
        Call SetParChange(optOrder, Index, mrsPar, True, strValue)
    End If
End Sub

Private Sub optOwnFee_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optOwnFee, Index, mrsPar)
End Sub

Private Sub optOwnFee_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optOwnFee_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optOwnFee, Index, mrsPar)
End Sub

Private Sub optPrintMode_SendCard_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintMode_SendCard, Index, mrsPar)
End Sub

Private Sub optPrintMode_SendCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintMode_SendCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintMode_SendCard, Index, mrsPar)
End Sub
Private Sub optPrintModeDraw_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintModeDraw, Index, mrsPar)
End Sub

Private Sub optPrintModeDraw_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintModeDraw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintModeDraw, Index, mrsPar)
End Sub
Private Sub optPrintModeSJ_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintModeSJ, Index, mrsPar)
End Sub

Private Sub optPrintModeSJ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintModeSJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintModeSJ, Index, mrsPar)
End Sub

Private Sub optPrintRequisition_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintRequisition, Index, mrsPar)
End Sub

Private Sub optPrintRequisition_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintRequisition_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintRequisition, Index, mrsPar)
End Sub


Private Sub optRegistClearMzInfor_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optRegistClearMzInfor, Index, mrsPar)
End Sub

Private Sub optRegistClearMzInfor_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRegistClearMzInfor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRegistClearMzInfor, Index, mrsPar)
End Sub

Private Sub optDelCardMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDelCardMode, Index, mrsPar)
End Sub

Private Sub optDelCardMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDelCardMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDelCardMode, Index, mrsPar)
End Sub
 
Private Sub optRegistPrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optRegistPrintMode, Index, mrsPar)
End Sub

Private Sub optRegPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optRegPrint, Index, mrsPar)
End Sub

Private Sub optRegistPrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRegPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRegistPrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRegistPrintMode, Index, mrsPar)
End Sub

Private Sub optRegPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRegPrint, Index, mrsPar)
End Sub
 

Private Sub optRuleTotal_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SaveBillRuleChange
End Sub

Private Sub optRuleTotal_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRuleTotal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRuleTotal, Index, mrsPar)
End Sub
Private Sub optSendDrugFF_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSendDrugFF, Index, mrsPar)
End Sub

Private Sub optSendDrugFF_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
End Sub

Private Sub optSendDrugFF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSendDrugFF, Index, mrsPar)
End Sub

Private Sub optSetMoneyMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSetMoneyMode, Index, mrsPar)
End Sub

Private Sub optSetMoneyMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optSetMoneyMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSetMoneyMode, Index, mrsPar)
End Sub

Private Sub optSlipPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSlipPrint, Index, mrsPar)
End Sub

Private Sub optSlipPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optSlipPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSlipPrint, Index, mrsPar)
End Sub

Private Sub optMoneyControl_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optMoneyControl, Index, mrsPar)
End Sub

Private Sub optMoneyControl_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optMoneyControl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optMoneyControl, Index, mrsPar)
End Sub

Private Sub optBarCodePrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBarCodePrint, Index, mrsPar)
End Sub

Private Sub optBarCodePrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBarCodePrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBarCodePrint, Index, mrsPar)
End Sub


Private Sub optPrintBespeak_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintBespeak, Index, mrsPar)
End Sub

Private Sub optPrintBespeak_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintBespeak_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintBespeak, Index, mrsPar)
End Sub

Private Sub optReceiveMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optReceiveMode, Index, mrsPar)
End Sub

Private Sub optReceiveMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optReceiveMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optReceiveMode, Index, mrsPar)
End Sub


Private Sub optSupplementaryPrint_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSupplementaryPrint, Index, mrsPar)
End Sub

Private Sub optSupplementaryPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optSupplementaryPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSupplementaryPrint, Index, mrsPar)
End Sub


Private Sub optSupplementaryUnit_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optSupplementaryUnit, Index, mrsPar)
End Sub

Private Sub optSupplementaryUnit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optSupplementaryUnit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optSupplementaryUnit, Index, mrsPar)
End Sub



Private Sub optTriageBarcodePrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optTriageBarcodePrintMode, Index, mrsPar)
End Sub

Private Sub optTriageBarcodePrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optTriageBarcodePrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optTriageBarcodePrintMode, Index, mrsPar)
End Sub

Private Sub optTriageQueuingMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optTriageQueuingMode, Index, mrsPar)
End Sub

Private Sub optTriageQueuingMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Index = 1 Then
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optTriageQueuingMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optTriageQueuingMode, Index, mrsPar)
End Sub
Private Sub optTriageSort_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optTriageSort, Index, mrsPar)
End Sub

Private Sub optTriageSort_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optTriageSort_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optTriageSort, Index, mrsPar)
End Sub
Private Sub optTriagePrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optTriagePrintMode, Index, mrsPar)
End Sub

Private Sub optTriagePrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optTriagePrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optTriagePrintMode, Index, mrsPar)
End Sub
Private Sub opt���۵�λ_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt���۵�λ, Index, mrsPar)
End Sub

Private Sub opt���۵�λ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt���۵�λ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt���۵�λ, Index, mrsPar)
End Sub

Private Sub optBillTotalShow_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBillTotalShow, Index, mrsPar)
End Sub

Private Sub optBillTotalShow_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    
    'zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBillTotalShow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBillTotalShow, Index, mrsPar)
End Sub
Private Sub optChargeBillTotalShow_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optChargeBillTotalShow, Index, mrsPar)
End Sub

Private Sub optChargeBillTotalShow_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optChargeBillTotalShow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optChargeBillTotalShow, Index, mrsPar)
End Sub
Private Sub optDrug_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDrug, Index, mrsPar)
End Sub
Private Sub optDrug_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDrug_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDrug, Index, mrsPar)
End Sub

Private Sub opt�ɿ�_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt�ɿ�, Index, mrsPar)
End Sub

Private Sub opt�ɿ�_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    tbPage(Pg_�����շ�).Item(1).Selected = True
    'zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt�ɿ�_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt�ɿ�, Index, mrsPar)
End Sub

Private Sub opt�շѿ����ʾ��ʽ_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt�շѿ����ʾ��ʽ, Index, mrsPar)
End Sub

Private Sub opt�շѿ����ʾ��ʽ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt�շѿ����ʾ��ʽ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt�շѿ����ʾ��ʽ, Index, mrsPar)
End Sub
Private Sub opt���ʿ����ʾ��ʽ_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt���ʿ����ʾ��ʽ, Index, mrsPar)
End Sub

Private Sub opt���ʿ����ʾ��ʽ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    'zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt���ʿ����ʾ��ʽ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt���ʿ����ʾ��ʽ, Index, mrsPar)
End Sub

Private Sub opt���ۿ����ʾ��ʽ_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt���ۿ����ʾ��ʽ, Index, mrsPar)
End Sub

Private Sub opt���ۿ����ʾ��ʽ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt���ۿ����ʾ��ʽ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt���ۿ����ʾ��ʽ, Index, mrsPar)
End Sub


Private Sub opt�շѵ�λ_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt�շѵ�λ, Index, mrsPar)
End Sub

Private Sub opt�շѵ�λ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt�շѵ�λ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt�շѵ�λ, Index, mrsPar)
End Sub

Private Sub opt���ʵ�λ_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(opt���ʵ�λ, Index, mrsPar)
End Sub

Private Sub opt���ʵ�λ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt���ʵ�λ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt���ʵ�λ, Index, mrsPar)
End Sub

Private Sub picDisplay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(picDisplay, Index, mrsPar)
End Sub

Private Sub picPar_Resize(Index As Integer)
    If Index = 6 Then
        'index=6:�Һ�ҵ��
        '�Һ�ҵ��
        With tbPage(Pg_�Һ�ҵ��)
            .Left = picPar(Index).ScaleLeft
            .Top = picPar(Index).ScaleTop
            .Height = picPar(Index).ScaleHeight
            .Width = picPar(Index).ScaleWidth
        End With
    End If
    If Index = 11 Then
        With tbPage(Pg_�����շ�)
            .Left = picPar(Index).ScaleLeft
            .Top = picPar(Index).ScaleTop
            .Height = picPar(Index).ScaleHeight
            .Width = picPar(Index).ScaleWidth
        End With
    End If
    If Index = 15 Then
        With tbPage(Pg_����ҵ��)
            .Left = picPar(Index).ScaleLeft
            .Top = picPar(Index).ScaleTop
            .Height = picPar(Index).ScaleHeight
            .Width = picPar(Index).ScaleWidth
        End With
    End If
End Sub

Private Sub tbPage_SelectedChanged(Index As Integer, ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Me.Visible Then Exit Sub
    Dim objTemp As Object
    Select Case Index
    Case Pg_�Һ�ҵ��
        With tbPage(Index)
            Select Case Val(.Selected.Tag)
            Case Pg_�Һ�_����
                If optRegistPlanMode(0).value = True Then
                    Set objTemp = chk(chk_ֻ��ҽ��ҽ�����йҺŰ���)
                Else
                    Set objTemp = chk(chk_���ﰲ��_ֻ����ѡԺ��ҽ��)
                End If
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case Pg_�Һ�_�Һ�
                Set objTemp = chk(chk_�Һ�_�Զ�ˢ�¹ҺŰ���)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case Pg_�Һ�_����
                Set objTemp = chk(chk_�Һű���ˢ��)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case Pg_�Һ�_ԤԼ
                Set objTemp = chk(chk_�Һ�_ԤԼ��ʾ���кű�)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            End Select
        End With
    Case Pg_�����շ�
        With tbPage(Index)
            Select Case Val(.Selected.Tag)
            Case Pg_�շ�_���ݿ���
               Set objTemp = chk(chk_�շ�_����������ҩ����)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            Case Pg_�շ�_Ʊ�ݿ���
                Set objTemp = cbo(cbo_�շ�_Ʊ�ݷ������)
                If objTemp.Enabled And objTemp.Visible Then
                    objTemp.SetFocus
                End If
            End Select
        End With
    Case Else
    End Select
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim i As Long
    
    For i = 0 To picPar.UBound
        picPar(i).Visible = (i = Item.ID)
    Next
    
    lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("���ʱ���", marrFunc) Or Item.ID = GetFuncID("�����Զ�����", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOK.Left - lblPrompt.Left - 120
    mlngPreFind = 1
    
    tplFunc.Tag = Item.ID   '���ڻ�ȡ��ǰѡ�е�TaskPanelItem
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picVbar.Left < 1500 Then picVbar.Left = 1500
    If picVbar.Left > Me.ScaleWidth - 3000 Then picVbar.Left = Me.ScaleWidth - 3000
    picVbar.Top = 0
    
    picFunc.Width = picVbar.Left + picVbar.Width
    
    For i = 0 To picPar.UBound
        picPar(i).Top = Me.ScaleTop
        picPar(i).Left = picFunc.Left + picFunc.ScaleWidth
        picPar(i).Width = Me.ScaleWidth - picPar(i).Left
        picPar(i).Height = Me.ScaleHeight - PicBottom.ScaleHeight
    Next
End Sub

Private Sub mshAutoCalc_GotFocus()
    If lblLocate(txt_Dept).Tag <> "mshAutoCalc" Then
        lblLocate(txt_Dept).Tag = "mshAutoCalc"
        mlngPreFind = 1
    End If
End Sub

Private Sub Bill_GotFocus(Index As Integer)
    If Index = bill_�Զ����� Then
        If lblLocate(txt_Dept).Tag <> "Bill" Then
            lblLocate(txt_Dept).Tag = "Bill"
            mlngPreFind = 1
        End If
    End If
End Sub

Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = PicBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
End Sub


Private Sub picFunc_Resize()
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    scbFunc.Height = picFunc.ScaleHeight
    
    picVbar.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub


Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIF(picVbar.Left + X < 2000, 2000, picVbar.Left + X)
        Call Form_Resize
    End If
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID�Ǵ�1��ʼ�ģ���ΪͬʱΪͼ����ţ�,�����Ǵ�0��ʼ
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub


Public Sub LocateFuncItem(ByVal lngFunc As Long)
'���ܣ�����IDѡ��һ���Ͷ�������
    Dim i As Long, j As Long, lngId As Long
    Dim arrTmp As Variant
    Dim n As Long
    
    For i = 0 To UBound(marrFunc)
        arrTmp = Split(marrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            lngId = Split(arrTmp(j), ",")(1)
            If lngFunc = lngId Then
                tplFunc.Tag = lngId
                Set scbFunc.Selected = scbFunc(i)
                
                For n = 1 To tplFunc.Groups(1).Items.Count
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).ID = lngId
                Next
            End If
        Next
    Next
End Sub

Private Sub InitData()
'���ܣ���ʼ������ؼ�,��ȡ����������

    '1.��ʼ������
    
    mlngPreFind = 1
    mblnJRaiseByDate = IsRaiseByDate("J")
    mblnHRaiseByDate = IsRaiseByDate("H")
        
    Call InitSystemPara
    
    
    
    '2.��ʼ������ؼ�
    Call InitEnv
        
    Call Load��������
    Call LoadOneCard
    Call Load����
    Call Load���ݲ���
    
    Call LoadOther
    
    RestoreFlexState mshAutoCalc, App.ProductName & "\" & Me.Name
    RestoreFlexState Bill(bill_�Զ�����), App.ProductName & "\" & Me.Name & bill_�Զ�����
    RestoreFlexState Bill(bill_���ʱ���), App.ProductName & "\" & Me.Name & bill_���ʱ���
    
    
    '3.����ϵͳ����
    Call LoadPar
    
    
End Sub

Private Sub LoadPar()
'���ܣ���ȡ�����ز���������ؼ�
    Dim strValue As String, strTmp As String
    Dim i As Long, n As Long, blnFind As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String, rsTemp As ADODB.Recordset
    Dim arrObj As Variant  '�������ģ��1,������1,�ؼ�����1,ģ��2,������2,�ؼ�����2,......
    Dim varData As Variant, varTemp As Variant
    Dim strArr() As String, strStyle() As String
    Dim intTemp As Integer
    
    Set rsTmp = GetPar(mrsPar, "9000,1103,1107,1110,1111,1113,1114,1120,1121,1122,1124,1133,1134,1135,1150,1137,1142,1143,1151,1500,1501,1502,1503,1504,1506,1257")
     '1.����CheckBox�����
    strTmp = "0:7:" & chk_�Զ����� & _
            ",0:31:" & chk_��Ժ���˲�׼��Ժ���� & _
            ",0:52:" & chk_���뿪���� & _
            ",0:53:" & chk_���ƿ����� & _
            ",0:72:" & chk_���������շ���� & _
            ",0:93:" & chk_������Ŀ���ܼ����ۿ� & _
            ",0:270:" & chk_ԤԼ�ŶӰ�ʱ�� & _
            ",0:98:" & chk_���ʱ����������۷��� & _
            ",0:276:" & chk_���ѿ�_ˢ�������붨λ������� & _
            ",0:282:" & chk_���ѿ�_���ѿ��˷�ˢ������ & _
            ",0:279:" & chk_�Һ�_ͬһ���ֻ֤�ܶ�Ӧһ���������� & _
            ",0:290:" & chk_������Һ�ģʽ & _
            ",0:316:" & chk_ָ�����ϲ���ʱ����ʾ�޿������
            
    strTmp = strTmp & _
            ",0:100:" & chk_���������ģʽ & _
            ",0:144:" & chk_�շ���Ŀ��λ�������� & _
            ",0:151:" & chk_�����˷��������� & _
            ",0:163:" & chk_��Ŀִ��ǰ�����շѻ���� & _
            ",0:232:" & chk_��Ŀ�����������շѻ������� & _
            ",0:215:" & chk_δ��ƽ�ֹ���� & _
            ",0:283:" & chk_�������תסԺԤ��Ʊ�ݴ�ӡ����
    
    strTmp = strTmp & _
            "," & p��������ģ�� & ":�Һű���ˢ��:" & chk_�Һű���ˢ�� & _
            "," & p��������ģ�� & ":����ʹ��Ԥ��:" & chk_����ʹ��Ԥ�� & _
            "," & p��������ģ�� & ":����סԺ���˹Һ�:" & chk_����סԺ���˹Һ� & _
            "," & p��������ģ�� & ":������ѡ��:" & chk_������ѡ�� & _
            "," & p��������ģ�� & ":ԤԼʱ�տ�:" & chk_ԤԼʱ�տ� & _
            "," & p��������ģ�� & ":����ģ������:" & chk_����ģ������ & _
            "," & p��������ģ�� & ":����ҽ��:" & chk_ҽ��վ����ҽ��
    
    'Ԥ�����Check����
    strTmp = strTmp & _
        "," & pԤ������� & ":���������Ľɿ:" & chk_ֻ��ʾ��ʣ�����ʷ�ɿ� & _
        "," & pԤ������� & ":������Ľɿ����:" & chk_������Ĳ��˵Ľɿ���� & _
        "," & pԤ������� & ":��Ԥ���������Ϣ:" & chk_��Ԥ�������������Ϣ & _
        "," & pԤ������� & ":����δ��Ʋ�׼��Ԥ��:" & chk_��Ժ����δ��Ʋ�׼��Ԥ�� & _
        "," & pԤ������� & ":�����Ժ���˽�סԺԤ��:" & chk_�����Ժ���˽�סԺԤ�� & _
        "," & pԤ������� & ":����ģ������:" & chk_����ͨ������ģ�����Ҳ��� & _
        "," & pԤ������� & ":סԺ��Ԥ����֤:" & chk_��סԺԤ��ˢ����֤ & _
        "," & pԤ������� & ":������Ժ��������˿�:" & chk_������Ժ��������˿� & _
        "," & pԤ������� & ":��ֹ��Ժ���˽�����Ԥ��:" & chk_��ֹ��Ժ���˽�����Ԥ�� & _
        "," & pԤ������� & ":Ԥ�����վ����ʾ:" & chk_Ԥ�����վ����ʾ

    'ҽ�ƿ����Check����
    strTmp = strTmp & _
        "," & pҽ�ƿ����� & ":���Ѽ���:" & chk_�����Լ��˷�ʽ��ȡ & _
        "," & pҽ�ƿ����� & ":����ģ������:" & chk_��������ģ������ & _
        "," & pҽ�ƿ����� & ":����ʹ�������շ�ҽ���վ�:" & chk_����ʹ�������շ�ҽ���վ� & _
        "," & pҽ�ƿ����� & ":��ȡ������:" & chk_��ȡ������ & _
        "," & pҽ�ƿ����� & ":�Զ������:" & chk_ҽ�ƿ�_�Զ���������� & _
        ""
    
    '�ҺŰ������Check����
    strTmp = strTmp & _
    "," & p�ҺŰ��� & ":ֻ����ѡԺ��ҽ��:" & chk_ֻ��ҽ��ҽ�����йҺŰ��� & _
    "," & p�ҺŰ��� & ":ԤԼ�����ڽ�ֹɾ��:" & chk_����ԤԼ�Һŵ���ֹɾ������ & _
    ""
    
    '�ҺŹ������Check����
    strTmp = strTmp & _
        "," & p�ҺŹ��� & ":�Զ������:" & chk_�Һ�_�Զ���������� & _
        "," & p�ҺŹ��� & ":��Ϊ���۵�:" & chk_�Һ�_��Ϊ���۵� & _
        "," & p�ҺŹ��� & ":��ȡ����:" & chk_�Һ�_������Һ�һ���� & _
        "," & p�ҺŹ��� & ":����ҽ��:" & chk_�Һ�_��ҽ����������ҽ�� & _
        "," & p�ҺŹ��� & ":�Զ���������:" & chk_�Һ�_�����Զ������²��� & _
        "," & p�ҺŹ��� & ":����ʹ��Ԥ����:" & chk_�Һ�_����ʹ��Ԥ����ɷ� & _
        "," & p�ҺŹ��� & ":����ģ������:" & chk_�Һ�_����ģ������ & _
        "," & p�ҺŹ��� & ":����ô�ӡ:" & chk_�Һ�_����ô�ӡƱ�� & _
        "," & p�ҺŹ��� & ":��������:" & chk_�Һ�_����_���� & _
        "," & p�ҺŹ��� & ":�����Ա�:" & chk_�Һ�_����_�Ա� & _
        "," & p�ҺŹ��� & ":��������:" & chk_�Һ�_����_���� & _
        "," & p�ҺŹ��� & ":�����ͥ��ַ:" & chk_�Һ�_����_��ͥ��ַ & _
        "," & p�ҺŹ��� & ":���븶�ʽ:" & chk_�Һ�_����_���ʽ & _
        "," & p�ҺŹ��� & ":����ѱ�:" & chk_�Һ�_����_�ѱ� & _
        "," & p�ҺŹ��� & ":������㷽ʽ:" & chk_�Һ�_����_���㷽ʽ & _
        "," & p�ҺŹ��� & ":������������:" & chk_�Һ�_�����������˵ǼǴ��� & _
        "," & p�ҺŹ��� & ":�˷��ش�:" & chk_�Һ�_�˺Ų��˿��ش�Ʊ�� & _
        "," & p�ҺŹ��� & ":��ӡ������ǩ:" & chk_�Һ�_�Һź��ӡ������ǩ & _
        "," & p�ҺŹ��� & ":ԤԼ��ʾ���кű�:" & chk_�Һ�_ԤԼ��ʾ���кű� & _
        "," & p�ҺŹ��� & ":ԤԼ����ȷ���Һŷ�:" & chk_�Һ�_ԤԼ����ȷ���Һŷ� & _
        "," & p�ҺŹ��� & ":����סԺ���˹Һ�:" & chk_�Һ�_����סԺ���˹Һ� & _
        ""
    
    strTmp = strTmp & _
        "," & p�ҺŹ��� & ":ԤԼ�����������:" & chk_�Һ�_ԤԼ����������� & _
        "," & p�ҺŹ��� & ":������ѡ��:" & chk_�Һ�_������ѡ�� & _
        "," & p�ҺŹ��� & ":�˺����:" & chk_�Һ�_N�����˺������ & _
        "," & p�ҺŹ��� & ":ʧԼ���ڹҺ�:" & chk_�Һ�_ԤԼʧԼ���ڹҺ� & _
        "," & p�ҺŹ��� & ":��ͥ��ַ���뷽ʽ:" & chk_�Һ�_��ͥ��ַ�������� & _
        "," & p�ҺŹ��� & ":ɨ�����֤ǩԼ:" & chk_�Һ�_ɨ�����֤ǩԼ & _
        "," & p�ҺŹ��� & ":�ƻ��Ű�Һ�Ĭ�Ͻ���:" & chk_�Һ�_�ƻ��Ű�ģʽ������� & _
        "," & p�ҺŹ��� & ":�����������Һ�:" & chk_�Һ�_�����������Һ� & _
        "," & p�ҺŹ��� & ":�ϸ�ʱ�ιҺ�:" & chk_�Һ�_�ϸ�ʱ�ιҺ� & _
        "," & p�ҺŹ��� & ":Ĭ�Ϲ�����:" & chk_�Һ�_Ĭ�Ϲ�ѡ����ѡ�� & _
        "," & p�ҺŹ��� & ":�������Ч�Լ��:" & chk_�Һ�_�������Ч�Լ�� & _
        "," & p�ҺŹ��� & ":���ϸ����ʱʼ�շ���:" & chk_�Һ�_���ϸ����Ϊ���� & _
        "," & p�ҺŹ��� & ":������ϵ�绰:" & chk_�Һ�_����_��ϵ�绰 & _
        "," & p�ҺŹ��� & ":��ֹ��������:" & chk_�Һ�_��ֹ�������� & _
        ""
    '����
        
    strTmp = strTmp & _
        "," & p������� & ":����̨ǩ���Ŷ�:" & chk_����_����̨ǩ����ʼ�Ŷ� & _
        "," & p������� & ":ԤԼ���ɶ���:" & chk_����_ԤԼ�ҺŽ������ & _
        "," & p������� & ":����æʱ�������:" & chk_����_����æ������� & _
        "," & p������� & ":���ﲡ���������Ŷ�:" & chk_����_����ǩ�� & _
        "," & p������� & ":�ٴ�ǩ���������Ŷ�:" & chk_����_����ǩ�� & _
        ""
        
    '�ٴ����ﰲ��
    strTmp = strTmp & _
        "," & p�ٴ����ﰲ�� & ":����ҽ��������:" & chk_���ﰲ��_����ҽ�������� & _
        "," & p�ٴ����ﰲ�� & ":������ҽ��ͬ������ԤԼ�Һŵ�:" & chk_���ﰲ��_������ҽ��ͬ������ԤԼ�Һŵ� & _
        "," & p�ٴ����ﰲ�� & ":ֻ����ѡԺ��ҽ��:" & chk_���ﰲ��_ֻ����ѡԺ��ҽ��
    
    '���ﻮ��
    strTmp = strTmp & _
        "," & p���ﻮ�۹��� & ":��ҩ����:" & chk_����_����������ҩ���� & _
        "," & p���ﻮ�۹��� & ":�������:" & chk_����_�������������� & _
        "," & p���ﻮ�۹��� & ":��ʾ��ʿ:" & chk_����_���￪���˺���ʿ & _
        "," & p���ﻮ�۹��� & ":��ʾ����ҩ�����:" & chk_����_������ʾ����ҩ����� & _
        "," & p���ﻮ�۹��� & ":��ʾ����ҩ����:" & chk_����_������ʾ����ҩ���� & _
        "," & p���ﻮ�۹��� & ":����ģ������:" & chk_����_��������ģ������ & _
        "," & p���ﻮ�۹��� & ":�Ա�:" & chk_����_��������_�Ա� & _
        "," & p���ﻮ�۹��� & ":����:" & chk_����_��������_���� & _
        "," & p���ﻮ�۹��� & ":�ѱ�:" & chk_����_��������_�ѱ� & _
        "," & p���ﻮ�۹��� & ":ҽ�Ƹ���:" & chk_����_��������_ҽ�Ƹ��� & _
        "," & p���ﻮ�۹��� & ":�Ӱ�:" & chk_����_��������_�Ƿ�Ӱ� & _
        "," & p���ﻮ�۹��� & ":��������:" & chk_����_��������_�������� & _
        "," & p���ﻮ�۹��� & ":������:" & chk_����_��������_������ & _
        "," & p���ﻮ�۹��� & ":��ʹ��ȱʡ������:" & chk_����_���ﲻȱʡ������ & _
        "," & p���ﻮ�۹��� & ":����Ҫ���뿪����:" & chk_����_�������Ҫ���뿪���� & _
        "," & p���ﻮ�۹��� & ":ȱʡ��������:" & chk_����_����ȱʡ�������� & _
        "," & p���ﻮ�۹��� & ":סԺ���˰������շ�:" & chk_����_סԺ���˰������շ� & _
        "," & p���ﻮ�۹��� & ":����¼������ʹ�õĿ�����:" & chk_����_����¼������ʹ�õĿ�����
   

    '�����շ�
    strTmp = strTmp & _
        "," & p�����շѹ��� & ":��ҩ����:" & chk_�շ�_����������ҩ���� & _
        "," & p�����շѹ��� & ":�������:" & chk_�շ�_�������������� & _
        "," & p�����շѹ��� & ":��ʾ��ʿ:" & chk_�շ�_���￪���˺���ʿ & _
        "," & p�����շѹ��� & ":��ʾ����ҩ�����:" & chk_�շ�_������ʾ����ҩ����� & _
        "," & p�����շѹ��� & ":��ʾ����ҩ����:" & chk_�շ�_������ʾ����ҩ���� & _
        "," & p�����շѹ��� & ":����ģ������:" & chk_�շ�_��������ģ������ & _
        "," & p�����շѹ��� & ":�Ա�:" & chk_�շ�_��������_�Ա� & _
        "," & p�����շѹ��� & ":����:" & chk_�շ�_��������_���� & _
        "," & p�����շѹ��� & ":�ѱ�:" & chk_�շ�_��������_�ѱ� & _
        "," & p�����շѹ��� & ":ҽ�Ƹ���:" & chk_�շ�_��������_ҽ�Ƹ��� & _
        "," & p�����շѹ��� & ":�Ӱ�:" & chk_�շ�_��������_�Ƿ�Ӱ� & _
        "," & p�����շѹ��� & ":��������:" & chk_�շ�_��������_�������� & _
        "," & p�����շѹ��� & ":������:" & chk_�շ�_��������_������ & _
        "," & p�����շѹ��� & ":��ʹ��ȱʡ������:" & chk_�շ�_���ﲻȱʡ������ & _
        "," & p�����շѹ��� & ":����Ҫ���뿪����:" & chk_�շ�_�������Ҫ���뿪���� & _
        "," & p�����շѹ��� & ":ȱʡ��������:" & chk_�շ�_����ȱʡ�������� & _
        "," & p�����շѹ��� & ":��ʾ�ۼ�:" & chk_�շ�_������ʾ�տ��ۼ� & _
        "," & p�����շѹ��� & ":���Ƥ�Խ��:" & chk_�շ�_�Ữ�۵����Ƥ�Խ�� & _
        "," & p�����շѹ��� & ":����ʹ��Ԥ����:" & chk_�շ�_����ʹ��Ԥ����ɷ� & _
        "," & p�����շѹ��� & ":�൥���շ�:" & chk_�շ�_����������ŵ��� & _
        "," & p�����շѹ��� & ":��Ѱ���۵���:" & chk_�շ�_��Ѱ���۵��� & _
        "," & p�����շѹ��� & ":���������۵�ѡ��:" & chk_�շ�_���������۵�ѡ�񴰿� & _
        "," & p�����շѹ��� & ":��鲡�˹Һſ���:" & chk_�շ�_��鲡�˹Һſ���

  strTmp = strTmp & _
        "," & p�����շѹ��� & ":סԺ���˰������շ�:" & chk_�շ�_סԺ�������շ� & _
        "," & p�����շѹ��� & ":��ȡ���ۺ������ɿ�:" & chk_�շ�_��ȡ���������ɿ� & _
        "," & p�����շѹ��� & ":�վݼ��չ�����:" & chk_�շ�_�Զ����չ����� & _
        "," & p�����շѹ��� & ":���ŵ����շѷֱ��ӡ:" & chk_�շ�_���ŵ����շѷֱ��ӡ & _
        "," & p�����շѹ��� & ":�շ�ÿ��ֻ��һ��Ʊ��:" & chk_�շ�_ÿ��ֻ��һ��Ʊ�� & _
        "," & p�����շѹ��� & ":ֻ��ҽ������ɹ������շ�:" & chk_�շ�_ֻ��ҽ������ɹ������շ� & _
        "," & p�����շѹ��� & ":�����˲���Ʊ�����ֽ������:" & chk_�շ�_�����˲���Ʊ�ݲ��ִ��� & _
        "," & p�����շѹ��� & ":��첡�˷ֵ��ݴ�ӡ:" & chk_�շ�_��첡�˰����ݷֱ��ӡ & _
        "," & p�����շѹ��� & ":����¼������ʹ�õĿ�����:" & chk_�շ�_����¼������ʹ�õĿ�����

    '�������
    strTmp = strTmp & _
        "," & p������ʹ��� & ":ֻ���Һ�Լ��λ����:" & chk_����_ֻ���Һ�Լ��λ���� & _
        "," & p������ʹ��� & ":��ҩ����:" & chk_����_����������ҩ���� & _
        "," & p������ʹ��� & ":�������:" & chk_����_�������������� & _
        "," & p������ʹ��� & ":��ʾ��ʿ:" & chk_����_���￪���˺���ʿ & _
        "," & p������ʹ��� & ":���ʴ�ӡ:" & chk_����_���ʴ�ӡ & _
        "," & p������ʹ��� & ":���۴�ӡ:" & chk_����_���۴�ӡ & _
        "," & p������ʹ��� & ":��˴�ӡ:" & chk_����_��˴�ӡ & _
        "," & p������ʹ��� & ":��ʾ����ҩ�����:" & chk_����_������ʾ����ҩ����� & _
        "," & p������ʹ��� & ":��ʾ����ҩ����:" & chk_����_������ʾ����ҩ���� & _
        "," & p������ʹ��� & ":����ģ������:" & chk_����_��������ģ������ & _
        "," & p������ʹ��� & ":����¼������ʹ�õĿ�����:" & chk_����_����¼������ʹ�õĿ�����
      
    'סԺ����ҵ��
    strTmp = strTmp & _
        "," & pסԺ���ʹ��� & ":���ʴ�ӡ:" & chk_סԺ����_���ʴ�ӡ & _
        "," & p���ҷ�ɢ���� & ":���ʴ�ӡ:" & chk_��ɢ����_���ʴ�ӡ & _
        "," & pҽ�����Ҽ��� & ":���ʴ�ӡ:" & chk_ҽ������_���ʴ�ӡ & _
        "," & pסԺ���ʲ��� & ":��ҩ����:" & chk_���ʲ���_��ҩ���븶�� & _
        "," & pסԺ���ʲ��� & ":�������:" & chk_���ʲ���_����������� & _
        "," & pסԺ���ʲ��� & ":��ʾ��ʿ:" & chk_���ʲ���_�����˰�����ʿ & _
        "," & pסԺ���ʲ��� & ":������Ϊ���۵�:" & chk_���ʲ���_Ƿ�ѱ��滮�۵� & _
        "," & pסԺ���ʲ��� & ":��ʾ����ҩ�����:" & chk_���ʲ���_��ʾ����ҩ����� & _
        "," & pסԺ���ʲ��� & ":��ʾ����ҩ����:" & chk_���ʲ���_��ʾ����ҩ���� & _
        "," & pסԺ���ʲ��� & ":���۴�ӡ:" & chk_���ʲ���_���۴�ӡ & _
        "," & pסԺ���ʲ��� & ":��˴�ӡ:" & chk_���ʲ���_��˴�ӡ & _
        ""
    
    '���˽���
    strTmp = strTmp & _
        "," & p���˽��ʹ��� & ":��Լ��λ�����˴�ӡ:" & chk_����_��Լ��λ�����˴�ӡ & _
        "," & p���˽��ʹ��� & ":��Ժ���˽��ʺ��Զ���Ժ:" & chk_����_��Ժ���ʺ��Զ���Ժ & _
        "," & p���˽��ʹ��� & ":���������:" & chk_����_��������� & _
        "," & p���˽��ʹ��� & ":����ָ��Ԥ����:" & chk_����_ʹ��ָ������Ԥ���� & _
        "," & p���˽��ʹ��� & ":��;������Ԥ��:" & chk_����_��;����ȱʡ��Ԥ���� & _
        "," & p���˽��ʹ��� & ":���ʺ������Ϣ:" & chk_����_���ʺ����������Ϣ & _
        "," & p���˽��ʹ��� & ":���ʼ�鲡������:" & chk_����_���ʼ�鲡��������� & _
        "," & p���˽��ʹ��� & ":�Ƚ��Էѷ��ò���ӡ����Ʊ��:" & chk_����_�Էѷ��ò���ӡ����Ʊ�� & _
        "," & p���˽��ʹ��� & ":�Է�ȱʡʹ��Ԥ��:" & chk_����_�Էѷ���ȱʡʹ��Ԥ�� & _
        "," & p���˽��ʹ��� & ":��Ժ������������תסԺ:" & chk_��Ժ������������תסԺ & _
        "," & p���˽��ʹ��� & ":���˶�ν��ʵ���������������:" & chk_����_���˶�ν��ʵ��������������� & _
        "," & p���˽��ʹ��� & ":�����������˿����:" & chk_����_�����������˿���� & _
        ""

    'ִ�еǼǹ���
    strTmp = strTmp & _
        "," & pִ�еǼǹ��� & ":ҽ��ҽ������:" & chk_ִ�еǼ�_��ʾҽ�����͵ĵ��� & _
        ""
        
    '������˹���
    strTmp = strTmp & _
        "," & p������˹��� & ":����תסԺ�����:" & chk_�������_����תסԺ����� & _
        ""
    'Ʊ��ʹ�ü��
    strTmp = strTmp & _
        "," & pƱ��ʹ�ü�� & ":����Ʊ��ǩ��ȷ��:" & chk_Ʊ�ݼ��_����Ʊ��ǩ��ȷ�� & _
        ""
    '��Ա������
    strTmp = strTmp & _
        "," & p��Ա������ & ":�����ӡ:" & chk_���_�����ӡ & _
        "," & p��Ա������ & ":�����ӡ:" & chk_���_�����ӡ & _
        ""
    
    '���ѿ�����
    strTmp = strTmp & _
        "," & p���ѿ����� & ":�ɿ��ӡ:" & chk_���ѿ�_�ɿ��ӡ & _
        ""
    
    'ҽ�����ѹ���
    strTmp = strTmp & _
        "," & pҽ�����ѹ��� & ":��ҩ���븶��:" & chk_����_��ҩ���븶�� & _
        "," & pҽ�����ѹ��� & ":�����������:" & chk_����_����������� & _
        "," & pҽ�����ѹ��� & ":��ʾ����ҩ�����:" & chk_����_��ʾ����ҩ����� & _
        "," & pҽ�����ѹ��� & ":��ʾ����ҩ����:" & chk_����_��ʾ����ҩ���� & _
        "," & pҽ�����ѹ��� & ":��������ֻ�����ڲ�����:" & chk_����_���������������� & _
        ""
        
    '�շ����˹���
    strTmp = strTmp & _
        "," & p�շ����ʹ��� & ":Ԥ�����ʰ������סԺ�ֱ�����:" & chk_����_Ԥ�����ʰ������סԺ�ֱ����� & _
        ""
    
    Call SetParToControl(strTmp, mrsPar, chk)
    Call LoadTriageQueuingDep
    Call SetTriageQueuingEnalbe(chk(chk_����_����̨ǩ����ʼ�Ŷ�).value)
    
    chk(chk_����_ֻ���Һ�Լ��λ����).Enabled = chk(chk_����_��������ģ������).value = 1
    
    With vsBillFormat(vsGrid_�շ�Ʊ�ݸ�ʽ)
        .ColHidden(.ColIndex("�����˲���Ʊ�ݸ�ʽ")) = chk(chk_�շ�_�����˲���Ʊ�ݲ��ִ���).value <> 1
    End With
    chk(chk_�շ�_��첡�˰����ݷֱ��ӡ).Enabled = (chk(chk_�շ�_���ŵ����շѷֱ��ӡ).value = vbChecked)


    '������ز���
    '2.����ComboBox�����
    strTmp = "0:23:" & cbo_�ѽᵥ�� & _
            ",0:58:" & cbo_δ�󵥾ݽ���
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    strTmp = "0:185:" & cbo_������˷�ʽ & _
            ",0:284:" & cbo_�������תסԺԤ����Ʊ��ʽ
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '�ٴ����ﰲ��
    strTmp = p�ٴ����ﰲ�� & ":��������ȽϷ�ʽ:" & cbo_�ٴ�����_����ȽϷ�ʽ
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '�Һ����
    strTmp = p�ҺŹ��� & ":ȱʡ����ʽ:" & cbo_�Һ�_ȱʡ����ʽ
        
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    'סԺ���ʲ���
    strTmp = pסԺ���ʲ��� & ":���ʺ�ҩ:" & cbo_���ʲ���_���ʺ�ҩ
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    
    '���˽��ʹ���
    strTmp = p���˽��ʹ��� & ":��Լ��λ���ʴ�ӡ:" & cbo_����_��Լ��λ���ʴ�ӡ
    Call SetParToControl(strTmp, mrsPar, cbo, 3)  '3-���ı�ֱ�ӱȽ�
    
    'һ��ͨ���Ѳ���
    strTmp = pһ��ͨ���Ѳ��� & ":�շ��վݸ�ʽ:" & cbo_һ��ͨ_��ƱƱ�ݸ�ʽ
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    strTmp = pһ��ͨ���Ѳ��� & ":����վݸ�ʽ:" & cbo_һ��ͨ_����Ʊ�ݸ�ʽ
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '3.����UpDown�����
    strTmp = "0:9:" & ud_���ý���λ�� & _
            ",0:66:" & ud_�Һ�ԤԼ���� & _
            ",0:157:" & ud_���õ��۱���λ��
    
    strTmp = strTmp & "," & p������� & ":������Ч����:" & ud_������Ч����
    '����
    strTmp = strTmp & "," & p���ﻮ�۹��� & ":ȡ�����۵�:" & ud_����_ȡ��N��Ļ��۵�
    
    '�շ�
    strTmp = strTmp & "," & p�����շѹ��� & ":�շ��վ����д�:" & ud_�շ�_�շ��վ����д�
    '������
    strTmp = strTmp & "," & p���ﲹ���� & ":��������Ч����:" & ud_������_��Ч����
    
    Call SetParToControl(strTmp, mrsPar, ud) 'mrsPar�洢�Ŀؼ�����txtUD


    '4.����TextBox�����
    
    strTmp = "" & p��������ģ�� & ":������������:" & txt_������������
    
    '   ҽ�ƿ�����
    strTmp = strTmp & "," & pҽ�ƿ����� & ":������������:" & txt_����ģ����������
    txt(txt_����ģ����������).Enabled = chk(chk_��������ģ������).value = 1
    
    '   �ҺŰ���
    
    '   �ҺŹ���
    strTmp = strTmp & "," & p�ҺŹ��� & ":�Զ�ˢ�¼��:" & txt_�Һ�_ˢ��ʱ��
    strTmp = strTmp & "," & p�ҺŹ��� & ":������������:" & txt_�Һ�_������������
    strTmp = strTmp & "," & p�ҺŹ��� & ":ԤԼʧԼ����:" & txt_�Һ�_ԤԼʧЧ����
    strTmp = strTmp & "," & p�ҺŹ��� & ":N���ڲ���ȡ��ԤԼ��:" & txt_�Һ�_N���ڲ���ȡ��ԤԼ��
    strTmp = strTmp & "," & p�ҺŹ��� & ":N�����±���¼��໤��:" & txt_�Һ�_N����������໤��
    
    '�������
    strTmp = strTmp & "," & p������� & ":��ǰNСʱ����:" & txt_����_��ǰNСʱ����
    
    '���ﻮ��
    strTmp = strTmp & "," & p���ﻮ�۹��� & ":�����:" & txt_����_���ﵥ�������
    strTmp = strTmp & "," & p���ﻮ�۹��� & ":������������:" & txt_����_��������ģ����������
    '�����շ�
    strTmp = strTmp & "," & p�����շѹ��� & ":�����:" & txt_�շ�_���ﵥ�������
    strTmp = strTmp & "," & p�����շѹ��� & ":������������:" & txt_�շ�_��������ģ����������
    strTmp = strTmp & "," & p�����շѹ��� & ":��Ѱ��������:" & txt_�շ�_��Ѱ���۵�������
    
    '�������
    strTmp = strTmp & "," & p������ʹ��� & ":������������:" & txt_����_��������ģ����������
    
    Call SetParToControl(strTmp, mrsPar, txt)
    chk(chk_�Һ�_�Զ�ˢ�¹ҺŰ���).value = IIF(Val(txt(txt_�Һ�_ˢ��ʱ��)) > 0, 1, 0)
    txt(txt_�Һ�_ˢ��ʱ��).Enabled = Val(txt(txt_�Һ�_ˢ��ʱ��).Text) > 0
    txt(txt_�Һ�_������������).Enabled = chk(chk_�Һ�_����ģ������).value = 1
    
    txt(txt_����_��������ģ����������).Enabled = chk(chk_����_��������ģ������).value = 1
    txt(txt_�շ�_��������ģ����������).Enabled = chk(chk_�շ�_��������ģ������).value = 1
    txt(txt_�շ�_��Ѱ���۵�������).Enabled = chk(chk_�շ�_��Ѱ���۵���).value = 1
    txt(txt_����_��������ģ����������).Enabled = chk(chk_����_��������ģ������).value = 1
    
    '5.����ListBox�����
    strTmp = ""
    'Call SetParToControl(strTmp, mrsPar, lst)

    '6.����OptionButton�����
    arrObj = Array(0, 160, opt����)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p��������ģ��, "�Һ�ģʽ", optRegist)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p��������ģ��, "�Һŷ�Ʊ��ӡ��ʽ", optPrintFact)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p��������ģ��, "�Һ�ƾ����ӡ��ʽ", optPrintSlip)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p��������ģ��, "ԤԼ�Һŵ���ӡ��ʽ", optPrintAppoint)
    Call SetParToControl("", mrsPar, arrObj)
    
    'Ԥ������
    arrObj = Array(pԤ�������, "�˿��ֹ��ʽ", optDepsoitDelSet)
    Call SetParToControl("", mrsPar, arrObj)
    
    'һ��ͨ���ѹ���
    arrObj = Array(pһ��ͨ���Ѳ���, "ˢ��ȱʡ������", optSetMoneyMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    'ҽ�ƿ�����
    arrObj = Array(pҽ�ƿ�����, "�˿�ˢ��", optDelCardMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(pҽ�ƿ�����, "������ӡ��ʽ", optPrintMode_SendCard)
    Call SetParToControl("", mrsPar, arrObj)
    
    '�ҺŹ���
    arrObj = Array(p�ҺŹ���, "�Һŷ�Ʊ��ӡ��ʽ", optRegistPrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�ҺŹ���, "�˺����������Ϣ", optRegistClearMzInfor)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�ҺŹ���, "���������ӡ��ʽ", optBarCodePrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�ҺŹ���, "ԤԼ�Һŵ���ӡ��ʽ", optPrintBespeak)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�ҺŹ���, "ԤԼ����ģʽ", optReceiveMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�ҺŹ���, "�Һ�ƾ����ӡ��ʽ", optSlipPrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�ҺŹ���, "�ҺŽɿ��������", optMoneyControl)
    Call SetParToControl("", mrsPar, arrObj)
    
    '����̨
    arrObj = Array(p�������, "�Ŷӽк�ģʽ", optTriageQueuingMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�������, "�Ŷӵ���ӡ", optTriagePrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�������, "��������ʽ", optTriageSort)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�������, "�����ӡ��ʽ", optTriageBarcodePrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    '�ٴ����ﰲ��
    arrObj = Array(p�ٴ����ﰲ��, "ԤԼ�嵥���Ʒ�ʽ", optToExcelMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�ٴ����ﰲ��, "ԤԼ�嵥��ӡ��ʽ", optPrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�ٴ����ﰲ��, "������ӡ��ʽ", optVisitTablePrintMode)
    Call SetParToControl("", mrsPar, arrObj)
    
    '���ﻮ��
    arrObj = Array(p���ﻮ�۹���, "ҩƷ��λ", opt���۵�λ)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���ﻮ�۹���, "�����ʾ��ʽ", opt���ۿ����ʾ��ʽ)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���ﻮ�۹���, "����ϼƷ�ʽ", optBillTotalShow)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���ﻮ�۹���, "����֪ͨ����ӡ��ʽ", optPrintRequisition)
    Call SetParToControl("", mrsPar, arrObj)
    
     
    
    '�����շ�
    arrObj = Array(p�����շѹ���, "ҩƷ��λ", opt�շѵ�λ)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�����շѹ���, "�����ʾ��ʽ", opt�շѿ����ʾ��ʽ)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�����շѹ���, "����ϼƷ�ʽ", optChargeBillTotalShow)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�����շѹ���, "�շѽɿ��������", opt�ɿ�)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�����շѹ���, "δ�ҺŲ����շ�", optChargeRegPrompt)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�����շѹ���, "�շ��嵥��ӡ��ʽ", optFeeListPrint)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�����շѹ���, "ҩƷ��ҩ�˷ѷ�ʽ", optDrug)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�����շѹ���, "�˷ѻص���ӡ��ʽ", optDelFeeRefundPrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�����շѹ���, "�շ�ִ�е���ӡ��ʽ", optChargeExeBillPrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    
    '�������
    arrObj = Array(p������ʹ���, "ҩƷ��λ", opt���ʵ�λ)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p������ʹ���, "�����ʾ��ʽ", opt���ʿ����ʾ��ʽ)
    Call SetParToControl("", mrsPar, arrObj)
    
    '������
    arrObj = Array(p���ﲹ����, "ҩƷ��λ��ʾ", optSupplementaryUnit)   '��ʾ��λ
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���ﲹ����, "�����嵥��ӡ��ʽ", optSupplementaryPrint)   '�����嵥��ӡ��ʽ
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���ﲹ����, "ҩƷ��ҩ�˷ѷ�ʽ", optDrugSupplementary)   'ҩƷ��ҩ���˷ѷ�ʽ
    Call SetParToControl("", mrsPar, arrObj)
    
    'סԺ����ҵ��
    arrObj = Array(pסԺ���ʲ���, "����ҩƷ��λ", optJZDrugUnit)    '��ʾ��λ
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(pסԺ���ʲ���, "���ʱ�����������סԺ���۷���", optInExseCharge)
    Call SetParToControl("", mrsPar, arrObj)
    
    '���˽��ʹ���
    arrObj = Array(p���˽��ʹ���, "���ʷ���ʱ��", optBalanceTime) '���ʷ��ð�����(�Ǽǻ���ʱ��)ʱ�����
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���˽��ʹ���, "���ʼ����տ���", optBalanceDSCheck) '���ʼ����տ���
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���˽��ʹ���, "�˿��վݴ�ӡ", optDelBalancePrint) '�����˿��վݴ�ӡ��ʽ
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���˽��ʹ���, "����ʱ��Ѫ�Ѽ��", optBlood)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���˽��ʹ���, "������ϸ��ӡ", optBalanceFeeListPrint) '������ϸ��ӡ
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p���˽��ʹ���, "���ʽɿ��������", optBalancePayin) '�ɿ����
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���˽��ʹ���, "����Ԥ��ȱʡʹ�÷�ʽ", optMzDeposit) '����Ԥ��ȱʡ��ӡ��ʽ
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���˽��ʹ���, "�Էѷ��ô�ӡ��ʽ", optOwnFee) '�Է��嵥��ӡ��ʽ
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p���˽��ʹ���, "Ԥ��Ʊ�ݴ�ӡ��ʽ", optBalanceDepositPrint) 'Ԥ�����ӡ��ʽ
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(p�շѲ�����, "�տ��վݴ�ӡ��ʽ", optPrintModeSJ)
    Call SetParToControl("", mrsPar, arrObj)
    arrObj = Array(p�շѲ�����, "���ý����õ���ӡ��ʽ", optPrintModeDraw)
    Call SetParToControl("", mrsPar, arrObj)
    
    'ִ�еǼǹ���
    arrObj = Array(pִ�еǼǹ���, "ִ�еǼǵ���ӡ��ʽ", optRegPrint)
    Call SetParToControl("", mrsPar, arrObj)

    'ҽ�����ѹ���
    arrObj = Array(pҽ�����ѹ���, "ҩƷ��λ", optDrugUnitFF)
    Call SetParToControl("", mrsPar, arrObj)
    
    arrObj = Array(pҽ�����ѹ���, "���ʺ�ҩ", optSendDrugFF)
    Call SetParToControl("", mrsPar, arrObj)

    
    '7.����ϵͳ����
    rsTmp.Filter = "ģ��=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
    
        Select Case rsTmp!������
        Case 28 'һ��ͨ����ˢ������
            strTmp = NVL(strValue, "1|0")
            If InStr(strTmp, "|") = 0 Then strTmp = "1|0"
            
            intTemp = Val(Split(strTmp, "|")(0))
            If intTemp > 2 Then intTemp = 1
            If intTemp < 0 Then txt(txt_����֧��).Text = -1 * intTemp: intTemp = 3
            optBrushCard(intTemp).value = True
            txt(txt_����֧��).Enabled = intTemp = 3
            intTemp = Val(Split(strTmp, "|")(1))
            If intTemp < 0 Or intTemp > 2 Then intTemp = 0
            optBrushCard(10 + intTemp).value = True
            
            Call SetParRelations(Array(optBrushCard(0), optBrushCard(1), optBrushCard(2), optBrushCard(3), _
                optBrushCard(10), optBrushCard(11), optBrushCard(12)), mrsPar, rsTmp!������)
        Case 14    '��Ǯ����
            strTmp = IIF(IsNull(strValue), "0000", strValue)
            n = Val(Mid(strTmp, 1, 1))
            For i = 0 To cbo(cbo_�Һ���Ǯ����).ListCount
                If Val(Split(cbo(cbo_�Һ���Ǯ����).List(i) & "-", "-")(0)) = n Then cbo(cbo_�Һ���Ǯ����).ListIndex = i: Exit For
            Next
            cbo(cbo_�շ���Ǯ����).ListIndex = Val(Mid(strTmp, 2, 1))
            cbo(cbo_������Ǯ����).ListIndex = Val(Mid(strTmp, 3, 1))
            cbo(cbo_���ѿ���Ǯ����).ListIndex = Val(Mid(strTmp, 4, 1))
        
            Call SetParRelation(cbo, cbo_�Һ���Ǯ����, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_�շ���Ǯ����, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_������Ǯ����, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_���ѿ���Ǯ����, mrsPar)
        Case 278    '�Զ�����ģʽ
            n = Val(strValue)
            With cbo(cbo_�Զ�����ģʽ)
                .ListIndex = -1
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = n Then .ListIndex = i: Exit For
                Next
                If .ListIndex < 0 And n = 2 Then
                    .AddItem "2-��ʱ�Ե����"
                    .ItemData(.NewIndex) = 2: .ListIndex = .NewIndex
                ElseIf .ListIndex < 0 Then
                    .ListIndex = 0
                End If
                chk(chk_���������ģʽ).Visible = InStr(1, ",0,2,", "," & .ItemData(.ListIndex) & ",") > 0
                opt����(0).Enabled = InStr(1, ",0,2,", "," & .ItemData(.ListIndex) & ",") > 0
                opt����(1).Enabled = InStr(1, ",0,2,", "," & .ItemData(.ListIndex) & ",") > 0
                
                lblAutoChargeNM.Visible = .ItemData(.ListIndex) = 1
            End With
            Call SetParRelation(cbo, cbo_�Զ�����ģʽ, mrsPar, rsTmp!������)

        Case 17    '�������뷽ʽ���ֱ�Ϊ���������￨���Һŵ�������ID
            strTmp = NVL(strValue, "1111")
            chk(chk_��������).value = IIF(Val(Mid(strTmp, 1, 1)) = 0, 0, 1)
            chk(chk_ˢ���￨).value = IIF(Val(Mid(strTmp, 2, 1)) = 0, 0, 1)
            chk(chk_�Һŵ���).value = IIF(Val(Mid(strTmp, 3, 1)) = 0, 0, 1)
            chk(chk_����ID).value = IIF(Val(Mid(strTmp, 4, 1)) = 0, 0, 1)
            
            Call SetParRelation(chk, chk_��������, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_ˢ���￨, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_�Һŵ���, mrsPar)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_����ID, mrsPar)
            
        Case 20    '��ʾ����Ʊ�ݵĺ��볤�ȣ���λ�ֱ�Ϊ1-�շ�,2-Ԥ��,3-����,4-�Һ�
            strTmp = IIF(strValue = "", "7|7|7|7", strValue)
            lvw(lvw_Ʊ��).ListItems("C1").SubItems(1) = Split(strTmp, "|")(0)
            lvw(lvw_Ʊ��).ListItems("C2").SubItems(1) = Split(strTmp, "|")(1)
            lvw(lvw_Ʊ��).ListItems("C3").SubItems(1) = Split(strTmp, "|")(2)
            lvw(lvw_Ʊ��).ListItems("C4").SubItems(1) = Split(strTmp, "|")(3)
            
            
            varData = Array(lvw(lvw_Ʊ��), txtUD(ud_���볤��))
            Call SetParRelations(varData, rsTmp, Val(NVL(rsTmp!������)))
            Call lvw_ItemClick(lvw_Ʊ��, lvw(lvw_Ʊ��).SelectedItem)
            
            'Call SetParRelation(lvw, lvw_Ʊ��, mrsPar, rsTmp!������)
        Case 21  '�Һ���Ч����
            '��ͨ��
            ud(ud_�Һŵ�).value = IIF(Left(strValue, 1) = 0, 1, Left(strValue, 1))
            '�����
            ud(ud_����Һŵ�).value = IIF(Mid(strValue, 2, 1) = 0, 1, Mid(strValue, 2, 1))
        
            Call SetParRelation(txtUD, ud_�Һŵ�, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(txtUD, ud_����Һŵ�, mrsPar)
            
        Case 24    '��ʾ�Ƿ��ϸ���ƹ����Ʊ�ݵ�ʹ�ã���λ�ֱ�Ϊ1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
            strTmp = NVL(strValue, "1111")
            lvw(lvw_Ʊ��).ListItems("C1").SubItems(2) = IIF(Mid(strTmp, 1, 1) = "1", "��", "")
            lvw(lvw_Ʊ��).ListItems("C2").SubItems(2) = IIF(Mid(strTmp, 2, 1) = "1", "��", "")
            lvw(lvw_Ʊ��).ListItems("C3").SubItems(2) = IIF(Mid(strTmp, 3, 1) = "1", "��", "")
            lvw(lvw_Ʊ��).ListItems("C4").SubItems(2) = IIF(Mid(strTmp, 4, 1) = "1", "��", "")
            
            lvw(lvw_Ʊ��).ListItems("C1").Selected = True
            Call lvw_ItemClick(lvw_Ʊ��, lvw(lvw_Ʊ��).SelectedItem)
            
            Call SetParRelation(chk, chk_Ʊ�ſ���, mrsPar, rsTmp!������)
            
        Case 41    'ҽ���������÷�������
            SetListByText lst(lst_ҽ������), Replace(strValue, "|", ",")
            Call SetParRelation(lst, lst_ҽ������, mrsPar, rsTmp!������)
            
        Case 42    '���Ѳ������÷�������
            SetListByText lst(lst_���Ѳ���), Replace(strValue, "|", ",")
            Call SetParRelation(lst, lst_���Ѳ���, mrsPar, rsTmp!������)
            
        Case 46    'ˢ��Ҫ����������
            With lst(lst_ˢ������)
                For i = 1 To Len(NVL(strValue))
                    If Mid(strValue, i, 1) = "1" And i - 1 <= .ListCount - 1 Then
                        .Selected(i - 1) = True
                    End If
                Next
            End With
            Call SetParRelation(lst, lst_ˢ������, mrsPar, rsTmp!������)
            
        Case 60    '���ʷ���������ѽ��
            txt(txt_���������).Text = strValue
            Call txt_Validate(txt_���������, False)
            Call SetParRelation(txt, txt_���������, mrsPar, rsTmp!������)
        Case 98   '���ʱ����������۷���
            If Val(strValue) = 1 Then
                lblInExseCharge.Enabled = True
                optInExseCharge(0).Enabled = True
                optInExseCharge(1).Enabled = True
            Else
                lblInExseCharge.Enabled = False
                optInExseCharge(0).Enabled = False
                optInExseCharge(1).Enabled = False
            End If
        Case 256    '�Һ��Ű�ģʽ
            fraNewPaln.Move chk(chk_ֻ��ҽ��ҽ�����йҺŰ���).Left, chk(chk_ֻ��ҽ��ҽ�����йҺŰ���).Top
            If Val(Split(strValue & "|", "|")(0)) = 0 Then
                mblnNotChange = True
                optRegistPlanMode(0).value = 1
                optRegistPlanMode(1).value = 0
                mblnNotChange = False
                dtpRegistPlanMode.Enabled = False
                strSQL = "Select Max(����ʱ��) As ���ʱ�� From ���˹Һż�¼ Where ��¼״̬ =1 And ����ʱ�� >= [1] And ����ʱ�� Is Not Null"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, zlDatabase.Currentdate)
                If IsNull(rsTemp!���ʱ��) Then
                    dtpRegistPlanMode.value = zlDatabase.Currentdate
                Else
                    dtpRegistPlanMode.value = CDate(rsTemp!���ʱ��)
                End If
                cmdInstantActive.Enabled = True
                
                fraNewPaln.Visible = False
                chk(chk_ֻ��ҽ��ҽ�����йҺŰ���).Visible = True
                chk(chk_����ԤԼ�Һŵ���ֹɾ������).Visible = True
            Else
                mblnNotChange = True
                optRegistPlanMode(0).value = 0
                optRegistPlanMode(1).value = 1
                mblnNotChange = False
                dtpRegistPlanMode.Enabled = True
                If Split(strValue & "|", "|")(1) = "" Then
                    strSQL = "Select Max(����ʱ��) As ���ʱ�� From ���˹Һż�¼ Where ��¼״̬ =1 And ����ʱ�� >= [1] And ����ʱ�� Is Not Null"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, zlDatabase.Currentdate)
                    If IsNull(rsTemp!���ʱ��) Then
                        dtpRegistPlanMode.value = zlDatabase.Currentdate
                    Else
                        dtpRegistPlanMode.value = CDate(rsTemp!���ʱ��)
                    End If
                Else
                    dtpRegistPlanMode.value = CDate(Split(strValue & "|", "|")(1))
                End If
                cmdInstantActive.Enabled = False
                
                fraNewPaln.Visible = True
                chk(chk_ֻ��ҽ��ҽ�����йҺŰ���).Visible = False
                chk(chk_����ԤԼ�Һŵ���ֹɾ������).Visible = False
            End If
            strSQL = "Select 1 From ���˹Һż�¼ Where ��¼״̬ =1 And �����¼ID Is Not Null And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If Not rsTemp.EOF Then
                dtpRegistPlanMode.Enabled = False
                dtpRegistPlanMode.ToolTipText = "�Ѿ����ڳ�����Ű��µĹҺż�¼,�������޸�����ʱ��!"
            End If
            Call SetParRelations(Array(optRegistPlanMode(0), optRegistPlanMode(1), dtpRegistPlanMode), mrsPar, rsTmp!������)
        Case 261
            strValue = Replace(strValue, " ", "")
            With lst(lst_�ѻ�ҽ�����㷽ʽ)
                For i = 0 To .ListCount - 1
                    If InStr("|" & strValue & "|", "|" & zlCommFun.GetNeedName(.List(i), "-") & "|") > 0 Then
                        .Selected(i) = True
                    End If
                Next
            End With
            Call SetParRelation(lst, lst_�ѻ�ҽ�����㷽ʽ, mrsPar, rsTmp!������)
        Case 263
            txt(txt_ר�ҺŹҺ�����).Text = Val(strValue)
            chk(chk_ר�ҺŹҺ�����).value = IIF(Val(strValue) <> 0, 1, 0)
            If chk(chk_ר�ҺŹҺ�����).value = 1 Then
                txt(txt_ר�ҺŹҺ�����).Enabled = True
            Else
                txt(txt_ר�ҺŹҺ�����).Enabled = False
                txt(txt_ר�ҺŹҺ�����).Text = ""
            End If
            Call SetParRelations(Array(txt(txt_ר�ҺŹҺ�����), chk(chk_ר�ҺŹҺ�����)), mrsPar, rsTmp!������)
        Case 264
            txt(txt_ר�Һ�ԤԼ����).Text = Val(strValue)
            chk(chk_ר�Һ�ԤԼ����).value = IIF(Val(strValue) <> 0, 1, 0)
            If chk(chk_ר�Һ�ԤԼ����).value = 1 Then
                txt(txt_ר�Һ�ԤԼ����).Enabled = True
            Else
                txt(txt_ר�Һ�ԤԼ����).Enabled = False
                txt(txt_ר�Һ�ԤԼ����).Text = ""
            End If
            Call SetParRelations(Array(txt(txt_ר�Һ�ԤԼ����), chk(chk_ר�Һ�ԤԼ����)), mrsPar, rsTmp!������)
        Case 277 '��Դ����ʱ��
            dtpRegistTime.value = strValue
            Call SetParRelations(Array(dtpRegistTime), mrsPar, rsTmp!������)
        End Select
        rsTmp.MoveNext
    Loop
    
    
    '8.����ģ�����
    rsTmp.Filter = "ģ��=" & p��������ģ��
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "����ģ������" '15
            If Val(strValue) = 0 Then
                txt(txt_������������).Enabled = False
            Else
                txt(txt_������������).Enabled = True
            End If
        Case "�������Ұ���"
            strValue = strValue & "|"
            chk(chk_�ҺŰ������Ұ���).value = Val(Split(strValue, "|")(0))
            chk(chk_ԤԼ�������Ұ���).value = Val(Split(strValue, "|")(1))
            Call SetParRelations(Array(chk(chk_�ҺŰ������Ұ���), chk(chk_ԤԼ�������Ұ���)), rsTmp, CStr(NVL(rsTmp!������)), p��������ģ��)
        Case "ҽ��վ�Һ��������"
            fraOrder(1).Visible = True
            Call LoadStationRegOrder(strValue)
            Call SetParRelation(vsStationRegSort, 0, mrsPar, CStr(NVL(rsTmp!������)), p��������ģ��)
        End Select
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "ģ��=" & pԤ�������
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա"
            If strValue = "" Then strValue = "0|10"
            varData = Split(strValue & "|", "|")
            chk(chk_Ʊ��ʣ��N�����Ѳ���Ա).value = IIF(Val(varData(0)) = 1, 1, 0)
            txtUD(ud_Ʊ������).Text = Val(varData(1))
            ud(ud_Ʊ������).value = Val(varData(1))
            txtUD(ud_Ʊ������).Enabled = Val(varData(0)) = 1
            ud(ud_Ʊ������).Enabled = Val(varData(0)) = 1
            
            
            varData = Array(chk(chk_Ʊ��ʣ��N�����Ѳ���Ա), txtUD(ud_Ʊ������), ud(ud_Ʊ������))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!������)), pԤ�������)
            
        Case "���տ�����"
            Call SetParRelation(vs����, 0, mrsPar, CStr(NVL(rsTmp!������)), pԤ�������)
            Call Load���տ�(strValue)
        End Select
        rsTmp.MoveNext
    Loop
    Call LoadԤ��Ʊ�ݸ�ʽ(rsTmp)
    Call LoadԤ����Ʊ��ʽ(rsTmp)
    
    Call SetDrugStore
    
    rsTmp.Filter = "ģ��=" & pһ��ͨ���Ѳ���
    Call Loadҩ��(rsTmp)
    
    'ҽ�ƿ�����
    rsTmp.Filter = "ģ��=" & pҽ�ƿ�����
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "���������"
            If strValue = "" Then strValue = "���ڵ�ַ�ʱ�|ҽ����|�ѱ�|����|ҽ�Ƹ���|ѧ��|����֤��|�����ص�|������λ|��λ�绰|��λ�ʱ�|��λ�ʻ�|��λ������|��ϵ�����֤��|��ϵ�˵�ַ|��ϵ������|��ϵ�˵绰|��ϵ�˹�ϵ"
            Call LoadInputItem(vsGrid_��������������, strValue)
            Call SetParRelation(vsInputItemSet, vsGrid_��������������, mrsPar, CStr(NVL(rsTmp!������)), pҽ�ƿ�����)
        Case 1
            
        End Select
        rsTmp.MoveNext
    Loop
    
    Call Load����Ԥ��Ʊ�ݸ�ʽ(rsTmp)
    Call Loadҽ�ƿ�Ʊ�ݸ�ʽ(rsTmp)
    
    '�ٴ����ﰲ��
    rsTmp.Filter = "ģ��=" & p�ٴ����ﰲ��
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "δ����վ��ĺ�Դ��ά��վ��"
            With cbo(cbo_�ٴ�����_ȫԺͨ�ú�Դ����վ��)
                For i = 0 To .ListCount - 1
                    If zlStr.NeedCode(.List(i), "-") = strValue Then .ListIndex = i: Exit For
                Next
            End With
            Call SetParRelations(Array(cbo(cbo_�ٴ�����_ȫԺͨ�ú�Դ����վ��)), rsTmp, CStr(NVL(rsTmp!������)), p�ٴ����ﰲ��)
        End Select
        rsTmp.MoveNext
    Loop
    
    '�ҺŹ���
    rsTmp.Filter = "ģ��=" & p�ҺŹ���
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "ԤԼ����ʱ��"
            varData = Split(strValue & "|", "|")
            txt(txt_�Һ�_ԤԼ����ʱ��_����) = Val(varData(1))
            txt(txt_�Һ�_ԤԼ����ʱ��_����) = Val(varData(0))
            Call SetParRelations(Array(txt(txt_�Һ�_ԤԼ����ʱ��_����), txt(txt_�Һ�_ԤԼ����ʱ��_����)), rsTmp, CStr(NVL(rsTmp!������)), p�ҺŹ���)
        Case "���˹Һſ�������"
            txt(txt_�Һ�_���˹Һſ�������).Text = Val(strValue)
            chk(chk_�Һ�_���˹Һſ�������).value = IIF(Val(strValue) <> 0, 1, 0)
            Call SetParRelations(Array(txt(txt_�Һ�_���˹Һſ�������), chk(chk_�Һ�_���˹Һſ�������)), rsTmp, CStr(NVL(rsTmp!������)), p�ҺŹ���)
        Case "����ԤԼ������"
            txt(txt_�Һ�_����ԤԼ������).Text = Val(strValue)
            chk(chk_�Һ�_����ԤԼ������).value = IIF(Val(strValue) <> 0, 1, 0)
            Call SetParRelations(Array(txt(txt_�Һ�_����ԤԼ������), chk(chk_�Һ�_����ԤԼ������)), rsTmp, CStr(NVL(rsTmp!������)), p�ҺŹ���)
        Case "����ͬ����ԼN����"
            txt(txt_�Һ�_����ͬ����ԼN����).Text = Val(strValue)
            chk(chk_�Һ�_����ͬ����ԼN����).value = IIF(Val(strValue) <> 0, 1, 0)
            If chk(chk_�Һ�_����ͬ����ԼN����).value = 1 Then
                txt(txt_�Һ�_����ͬ����ԼN����).Enabled = True
            Else
                txt(txt_�Һ�_����ͬ����ԼN����).Enabled = False
                txt(txt_�Һ�_����ͬ����ԼN����).Text = ""
            End If
            Call SetParRelations(Array(txt(txt_�Һ�_����ͬ����ԼN����), chk(chk_�Һ�_����ͬ����ԼN����)), rsTmp, CStr(NVL(rsTmp!������)), p�ҺŹ���)
        Case "����ͬ���޹�N����"
            varData = Split(strValue & "|", "|")
            txt(txt_�Һ�_����ͬ���޹�N����).Text = Val(varData(0))
            chk(chk_�Һ�_����ͬ���޹�N����).value = IIF(Val(varData(0)) <> 0, 1, 0)
            If chk(chk_�Һ�_����ͬ���޹�N����).value = 0 Then
                chk(chk_�Һ�_����ͬ���޹�N����_����).value = 0
                chk(chk_�Һ�_����ͬ���޹�N����_����).Enabled = False
                txt(txt_�Һ�_����ͬ���޹�N����).Enabled = False
                txt(txt_�Һ�_����ͬ���޹�N����).Text = ""
            Else
                txt(txt_�Һ�_����ͬ���޹�N����).Enabled = True
                chk(chk_�Һ�_����ͬ���޹�N����_����).value = IIF(Val(varData(1)) <> 0, 1, 0)
            End If
            Call SetParRelations(Array(txt(txt_�Һ�_����ͬ���޹�N����), chk(chk_�Һ�_����ͬ���޹�N����), chk(chk_�Һ�_����ͬ���޹�N����_����)), rsTmp, CStr(NVL(rsTmp!������)), p�ҺŹ���)
        Case "ȱʡԤԼ��ʽ"
            With cbo(cbo_�Һ�_ȱʡԤԼ��ʽ)
                If .ListCount > 0 Then
                    For i = 0 To .ListCount - 1
                        If zlCommFun.GetNeedName(.List(i), "-") = strValue Then
                            .ListIndex = i: Exit For
                        End If
                    Next i
                    If .ListIndex < 0 Then .ListIndex = 0
                End If
            End With
            Call SetParRelations(Array(cbo(cbo_�Һ�_ȱʡԤԼ��ʽ)), rsTmp, CStr(NVL(rsTmp!������)), p�ҺŹ���)
        Case "ԤԼ��Чʱ��"
            If Val(strValue) >= 0 Then
                '��ǰ
                cbo(cbo_�Һ�_ԤԼ��Чʱ��).ListIndex = 0
            Else
                '�Ӻ�
                cbo(cbo_�Һ�_ԤԼ��Чʱ��).ListIndex = 1
            End If
            txt(txt_�Һ�_ԤԼ��Чʱ��).Text = Abs(strValue)
            Call SetParRelations(Array(cbo(cbo_�Һ�_ԤԼ��Чʱ��), txt(txt_�Һ�_ԤԼ��Чʱ��)), rsTmp, CStr(NVL(rsTmp!������)), p�ҺŹ���)
        Case "��ǰ�Һ���ɫ"
            strValue = Replace(strValue, " ", "")
            If strValue = "" Then strValue = "0"
            pic��ǰ��ɫ.BackColor = strValue
            Call SetParRelations(Array(pic��ǰ��ɫ), rsTmp, CStr(NVL(rsTmp!������)), p�ҺŹ���)
        Case "����ͬһ��Դ�޹�N����"
            txt(txt_�Һ�_����ͬһ��Դ�޹�N����).Text = Val(strValue)
            chk(chk_�Һ�_����ͬһ��Դ�޹�N����).value = IIF(Val(strValue) <> 0, 1, 0)
            Call SetParRelations(Array(txt(txt_�Һ�_����ͬһ��Դ�޹�N����), chk(chk_�Һ�_����ͬһ��Դ�޹�N����)), rsTmp, CStr(NVL(rsTmp!������)), p�ҺŹ���)
        End Select
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "ģ��=" & p�����շѹ���
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "�Զ����չҺŷ�"
            If InStr(1, strValue, ";") > 0 Then
                chk(chk_�շ�_δ�Һ��Զ����չҺŷ�).value = 1   '�����click�¼�,���ȼ����շ����
                txt(txt_�շ�_�������չҺŷ�).Tag = Split(strValue, ";")(0)
                txt(txt_�շ�_�������չҺŷ�).Text = Split(strValue, ";")(1)
            Else
                chk(chk_�շ�_δ�Һ��Զ����չҺŷ�).value = 0
                txt(txt_�շ�_�������չҺŷ�).Text = ""
                txt(txt_�շ�_�������չҺŷ�).Tag = ""
            End If
            Call SetParRelations(Array(txt(txt_�շ�_�������չҺŷ�), chk(chk_�շ�_δ�Һ��Զ����չҺŷ�)), rsTmp, CStr(NVL(rsTmp!������)), p�����շѹ���)
        Case "�Զ���ϵ���"
            If Val(strValue) <> 0 Then
                chk(chk_�շ�_�Զ���ϵ���).value = 1
                cbo(cbo_�շ�_�Զ���ϵ���).ListIndex = IIF(Val(strValue) = 1, 0, 1)
                cbo(cbo_�շ�_�Զ���ϵ���).Enabled = True
            Else
                chk(chk_�շ�_�Զ���ϵ���).value = 0
                cbo(cbo_�շ�_�Զ���ϵ���).ListIndex = 0
                cbo(cbo_�շ�_�Զ���ϵ���).Enabled = False
            End If
            Call SetParRelations(Array(cbo(cbo_�շ�_�Զ���ϵ���), chk(chk_�շ�_�Զ���ϵ���)), rsTmp, CStr(NVL(rsTmp!������)), p�����շѹ���)
        Case "Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա"
            If strValue = "" Then strValue = "0|10"
            varData = Split(strValue & "|", "|")
            chk(chk_�շ�_Ʊ��ʣ��X�ſ�ʼ����).value = IIF(Val(varData(0)) = 1, 1, 0)
            txtUD(ud_�շ�_Ʊ������).Text = Val(varData(1))
            ud(ud_�շ�_Ʊ������).value = Val(varData(1))
            txtUD(ud_�շ�_Ʊ������).Enabled = Val(varData(0)) = 1
            ud(ud_�շ�_Ʊ������).Enabled = Val(varData(0)) = 1
            
            varData = Array(chk(chk_�շ�_Ʊ��ʣ��X�ſ�ʼ����), txtUD(ud_�շ�_Ʊ������), ud(ud_�շ�_Ʊ������))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!������)), p�����շѹ���)
        Case "Ʊ�ݷ������"
            '���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ;�շ�ϸĿ(����);��������(0-������;1-��ҳ����(����1ҳ����),2-�������(ѡ����ϸʱ��Ч)).
            If strValue = "" Then strValue = "0||0;0;0;0;0;0"
            varData = Split(strValue & "||", "||")
            If Val(varData(0)) = 0 Then cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex = 0
            If Val(varData(0)) = 1 Then cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex = 1
            If Val(varData(0)) = 2 Then cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex = 2
            If cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex < 0 Then cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex = 0
            
            mintԭƱ�ݷ������ = cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex
            Call SetBillRuleParaLocale
            
            varTemp = Split(varData(1) & ";;;;", ";")
            
            varData = Array(cbo(cbo_�շ�_Ʊ�ݷ������), chkBillRule(0), chkBillRule(1), chkBillRule(2), chkBillRule(3), optRuleTotal(0), optRuleTotal(1), optRuleTotal(2), _
                            lblBillRuleNum(0), updBillRuleNum(0), txtBillRuleNum(0), lblBillRuleNum(1), updBillRuleNum(1), lblBillRuleNum(2), txtBillRuleNum(1), updBillRuleNum(2), txtBillRuleNum(2))
            
'            Call SetCtlsEnabled(varData, Not mblnExistPrintData)
            
            '2.����Ԥ���������Ʊ��
            '2.1�����ݷ�
            i = Val(varTemp(0))
            chkBillRule(0).value = IIF(i = 1, 1, 0)
            '2.2��ִ�п��ҷ�
            i = Val(varTemp(1))
            chkBillRule(1).value = IIF(i >= 1, 1, 0)
            updBillRuleNum(0).value = IIF(i < 0 Or i > 100, 0, i)
            txtBillRuleNum(0).Text = updBillRuleNum(0).value
            txtBillRuleNum(0).Tag = IIF(updBillRuleNum(0).value = 0, 1, updBillRuleNum(0).value)
            '2.3 ���վݷ�Ŀ
            i = Val(varTemp(2))
            chkBillRule(2).value = IIF(i >= 1, 1, 0)
            updBillRuleNum(1).value = IIF(i < 0 Or i > 100, 0, i)
            txtBillRuleNum(1).Text = updBillRuleNum(1).value
            txtBillRuleNum(1).Tag = IIF(updBillRuleNum(1).value = 0, 1, updBillRuleNum(1).value)
            '2.4 ���շ�ϸĿ(�ȴ����շ�ϸĿ����Ȼ�ᴥ��Click�¼�������ҳ����ִ��Ϊ����
            i = Val(varTemp(3))
            chkBillRule(3).value = IIF(i >= 1, 1, 0)
            updBillRuleNum(2).value = IIF(i < 0 Or i > 100, 0, i)
            txtBillRuleNum(2).Text = updBillRuleNum(2).value
            txtBillRuleNum(2).Tag = IIF(updBillRuleNum(2).value = 0, 20, updBillRuleNum(2).value)
            '2.5 �������
            i = Val(varTemp(4)): i = IIF(i > 3 Or i < 0, 0, i)
            optRuleTotal(i).value = True
            
            varData = Array(cbo(cbo_�շ�_Ʊ�ݷ������), chkBillRule(0), chkBillRule(1), chkBillRule(2), chkBillRule(3), optRuleTotal(0), optRuleTotal(1), optRuleTotal(2), _
                             updBillRuleNum(0), txtBillRuleNum(0), updBillRuleNum(1), txtBillRuleNum(1), updBillRuleNum(2), txtBillRuleNum(2))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!������)), p�����շѹ���)
            Call ShowRuleInfor
            
        Case "�շ�Ʊ�����ɷ�ʽ"
            i = Val(strValue)
            chk(chk_�շ�_Ʊ�����ɷ�ʽ).value = IIF(i >= 10, 1, 0)
            optBillMode(i Mod 10).value = True
            varData = Array(chk(chk_�շ�_Ʊ�����ɷ�ʽ), optBillMode)
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!������)), p�����շѹ���)
        
        End Select
        rsTmp.MoveNext
    Loop
    Call Load�շ�Ʊ�ݸ�ʽ(rsTmp)
    Call Load�˷�Ʊ�ݸ�ʽ(rsTmp)
    Call LoadDelFeeDefaultType
    
    rsTmp.Filter = "ģ��=" & p���ﲹ����
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "����ģ�����ҷ�ʽ"
            varData = Split(strValue & "|", "|")
            '���ñ�־|����
            chk(chk_������_����ģ������).value = IIF(Val(varData(0)) = 1, 1, 0)
            txt(txt_������_����ģ����������).Text = Val(varData(1))
            txt(txt_������_����ģ����������).Enabled = chk(chk_������_����ģ������).value = 1
            varData = Array(chk(chk_������_����ģ������), txt(txt_������_����ģ����������))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!������)), p���ﲹ����)
        Case "Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա"
            If strValue = "" Then strValue = "0|10"
            varData = Split(strValue & "|", "|")
            chk(chk_������_Ʊ��ʣ��X�ſ�ʼ����).value = IIF(Val(varData(0)) = 1, 1, 0)
            txtUD(ud_������_Ʊ������).Text = Val(varData(1))
            ud(ud_������_Ʊ������).value = Val(varData(1))
            txtUD(ud_������_Ʊ������).Enabled = Val(varData(0)) = 1
            ud(ud_������_Ʊ������).Enabled = Val(varData(0)) = 1
            
            varData = Array(chk(chk_������_Ʊ��ʣ��X�ſ�ʼ����), txtUD(ud_������_Ʊ������), ud(ud_������_Ʊ������))
            Call SetParRelations(varData, rsTmp, CStr(NVL(rsTmp!������)), p���ﲹ����)
        Case "����������շѽ��㷽ʽ"
            SetListByText lst(lst_������_���㷽ʽ), Replace(strValue, "|", ",")
            Call SetParRelation(lst, lst_������_���㷽ʽ, mrsPar, CStr(NVL(rsTmp!������)), p���ﲹ����)
        End Select
        rsTmp.MoveNext
    Loop
    Call Load������Ʊ�ݸ�ʽ(rsTmp)
    Call Load�������˷�Ʊ�ݸ�ʽ(rsTmp)
    
    rsTmp.Filter = "ģ��=" & p���˽��ʹ���
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        Case "����ǰ�Ƚ��Էѷ���"
            strValue = Replace(strValue, " ", "")
            With lst(lst_����_�Էѷ������)
                For i = 0 To .ListCount - 1
                    If InStr("," & strValue & ",", "," & Chr(.ItemData(i)) & ",") > 0 Then
                        .Selected(i) = True
                    End If
                Next
            End With
            Call SetParRelation(lst, lst_����_�Էѷ������, mrsPar, CStr(NVL(rsTmp!������)), p���˽��ʹ���)
        Case "�Ը��ϼ�������ɫ"
            strValue = Replace(strValue, " ", "")
            If strValue = "" Then strValue = "16711680"
            pic�Ը��ϼ�ɫ.BackColor = strValue
            Call SetParRelations(Array(pic�Ը��ϼ�ɫ), rsTmp, CStr(NVL(rsTmp!������)), p���˽��ʹ���)
        Case "��ǰ����������ɫ"
            strValue = Replace(strValue, " ", "")
            If strValue = "" Then strValue = "255|255"
            pic��ǰ����δ��ɫ.BackColor = Mid(strValue, 1, InStr(strValue, "|") - 1)
            pic��ǰ����δ��ɫ.BackColor = Mid(strValue, InStr(strValue, "|") + 1)
            Call SetParRelations(Array(pic��ǰ����δ��ɫ, pic��ǰ����δ��ɫ), rsTmp, CStr(NVL(rsTmp!������)), p���˽��ʹ���)
        Case "�ɿ�������ɫ"
            strValue = Replace(strValue, " ", "")
            If strValue = "" Then strValue = "16711680|255"
            pic�ɿ����ɿ�ɫ.BackColor = Mid(strValue, 1, InStr(strValue, "|") - 1)
            pic�ɿ����˿�ɫ.BackColor = Mid(strValue, InStr(strValue, "|") + 1)
            Call SetParRelations(Array(pic�ɿ����ɿ�ɫ, pic�ɿ����˿�ɫ), rsTmp, CStr(NVL(rsTmp!������)), p���˽��ʹ���)
        Case "��Ԥ��ȱʡ˳��"
            strValue = Replace(strValue, " ", "")
            strArr = Split(strValue & "|", "|")
            vsDepositSort.MergeRow(0) = True
            vsDepositSort.MergeCol(0) = True
            vsDepositSort.MergeCol(1) = True
            vsDepositSort.MergeCol(2) = True
            If Val(strArr(0)) = 0 Then
                optOrder(0).value = True
                vsDepositSort.Enabled = False
                cmdDepositDown.Enabled = False
                cmdDepositUp.Enabled = False
            Else
                strStyle = Split(strArr(1), ",")
                For i = 0 To UBound(strStyle)
                    With vsDepositSort
                        .TextMatrix(i + 2, .ColIndex("�������")) = Split(strStyle(i), ":")(0)
                        intTemp = Val(Split(strStyle(i), ":")(1))
                        Select Case intTemp
                            Case 0
                                .TextMatrix(i + 2, 2) = -1
                                .TextMatrix(i + 2, 3) = 0
                                .TextMatrix(i + 2, 4) = 0
                            Case 1
                                .TextMatrix(i + 2, 2) = 0
                                .TextMatrix(i + 2, 3) = -1
                                .TextMatrix(i + 2, 4) = 0
                            Case 2
                                .TextMatrix(i + 2, 2) = 0
                                .TextMatrix(i + 2, 3) = 0
                                .TextMatrix(i + 2, 4) = -1
                        End Select
                    End With
                Next i
                optOrder(1).value = True
                vsDepositSort.Enabled = True
                cmdDepositDown.Enabled = True
                cmdDepositUp.Enabled = True
            End If
            Call SetParRelations(Array(optOrder(0), optOrder(1), vsDepositSort, cmdDepositDown, cmdDepositUp), rsTmp, CStr(NVL(rsTmp!������)), p���˽��ʹ���)
        Case "���ʽ�����"
            If Val(strValue) = 1 Then
                opt������(0).value = False
                opt������(1).value = True
                chk(chk_����_���˶�ν��ʵ���������������).Enabled = False
                fraColor.Enabled = True
                For i = 0 To 4
                    lbl����Color(i).Enabled = True
                Next i
            Else
                opt������(0).value = True
                opt������(1).value = False
                chk(chk_����_���˶�ν��ʵ���������������).Enabled = True
                fraColor.Enabled = False
                For i = 0 To 4
                    lbl����Color(i).Enabled = False
                Next i
            End If
            Call SetParRelations(Array(opt������(0), opt������(1), picDisplay(0), picDisplay(1)), rsTmp, CStr(NVL(rsTmp!������)), p���˽��ʹ���)
        End Select
        rsTmp.MoveNext
    Loop
    Call Load����Ʊ�ݸ�ʽ(rsTmp)
    Call Load���ʺ�Ʊ��ʽ(rsTmp)
End Sub


Private Sub cmdDepositUp_Click()
    Dim strValue As String, i As Integer, intIndex As Integer
    With vsDepositSort
        If .Row <= 2 Then Exit Sub
        .RowPosition(.Row) = .Row - 1
        .Row = .Row - 1
    End With
    strValue = "1|"
    With vsDepositSort
        For i = 2 To 4
            If Abs(Val(.TextMatrix(i, 2))) = 1 Then intIndex = 0
            If Abs(Val(.TextMatrix(i, 3))) = 1 Then intIndex = 1
            If Abs(Val(.TextMatrix(i, 4))) = 1 Then intIndex = 2
            If i <> 4 Then
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex & ","
            Else
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex
            End If
        Next i
    End With
    Call SetParChange(optOrder, 1, mrsPar, True, strValue)
End Sub

Private Sub cmdDepositDown_Click()
    Dim strValue As String, i As Integer, intIndex As Integer
    With vsDepositSort
        If .Row >= .Rows - 1 Then Exit Sub
        .RowPosition(.Row) = .Row + 1
        .Row = .Row + 1
    End With
    strValue = "1|"
    With vsDepositSort
        For i = 2 To 4
            If Abs(Val(.TextMatrix(i, 2))) = 1 Then intIndex = 0
            If Abs(Val(.TextMatrix(i, 3))) = 1 Then intIndex = 1
            If Abs(Val(.TextMatrix(i, 4))) = 1 Then intIndex = 2
            If i <> 4 Then
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex & ","
            Else
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex
            End If
        Next i
    End With
    Call SetParChange(optOrder, 1, mrsPar, True, strValue)
End Sub

Private Sub SetCtlsEnabled(ByVal varDara As Variant, blnEnabled)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enabled����
    '����:���˺�
    '����:2015-06-17 17:46:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngIndex As Long, blnNotClear As Boolean
    On Error GoTo ErrHandle
    For i = 0 To UBound(varDara)
        varDara(i).Enabled = blnEnabled
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub LoadInputItem(ByVal intIndex As Integer, ByVal strValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '���:intIndex-����ֵ
    '     strValue-ȱʡ����ֵ��������,��ʽ:������Ŀ,��ֹ¼��,����Ƿ�����,������|....
    '����:���˺�
    '����:2015-06-11 17:32:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim intRow As Integer, i As Integer
    
    On Error GoTo ErrHandle
    varData = Split(strValue, "|")
    With vsInputItemSet(intIndex)
        .redraw = flexRDNone
        .Clear 1
        If strValue = "" Then .Rows = 2: Exit Sub
        .Rows = 2: intRow = 1
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",,,,", ",")
            If varTemp(0) <> "" Then
                .TextMatrix(intRow, .ColIndex("������Ŀ")) = varTemp(0)
                .TextMatrix(intRow, .ColIndex("��ֹ¼��")) = IIF(Val(varTemp(1)) = 1, "��", "")
                .TextMatrix(intRow, .ColIndex("������")) = IIF(Val(varTemp(2)) = 1, "��", "")
                .TextMatrix(intRow, .ColIndex("������")) = IIF(Val(varTemp(3)) = 1, "��", "")
                If .TextMatrix(intRow, .ColIndex("��ֹ¼��")) = "��" Then
                    .Cell(flexcpBackColor, intRow, .ColIndex("������"), intRow, .ColIndex("������")) = &H8000000F
                ElseIf .TextMatrix(intRow, .ColIndex("������")) = "��" _
                    Or .TextMatrix(intRow, .ColIndex("������")) = "��" Then
                    .Cell(flexcpBackColor, intRow, .ColIndex("��ֹ¼��")) = &H8000000F
                End If
                .Rows = .Rows + 1: intRow = intRow + 1
            End If
        Next
        If .Rows > 2 And Trim(.TextMatrix(.Rows - 1, .ColIndex("������Ŀ"))) = "" Then
            .Rows = .Rows - 1
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandle:
    vsInputItemSet(intIndex).redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
'���ܣ���ʼ������ؼ������ػ�������
    Dim strTmp As String, rsTmp As ADODB.Recordset
    Dim i As Long, rsTemp As ADODB.Recordset

    cbo(cbo_�ѽᵥ��).AddItem "0-����"
    cbo(cbo_�ѽᵥ��).AddItem "1-��ʾ"
    cbo(cbo_�ѽᵥ��).AddItem "2-��ֹ"
    cbo(cbo_�ѽᵥ��).ListIndex = 0
        

    '6-�ֱ���������:34519
    strTmp = "0-������|1-�ֱ���������|2-�ֱҲ�����ȡ|3-�ֱ������ȡ|4-�ֱ������������˫|5-�Ǳ��������塢�������|6-�ֱ���������"
    For i = 0 To UBound(Split(strTmp, "|"))
        '�ҺŲ�֧�������������˫,��Һ���ʹ��ҽ���Ľ����������̴���ֱ�,Oracle��û�������������˫����
        If i <> 4 Then cbo(cbo_�Һ���Ǯ����).AddItem Split(strTmp, "|")(i)
        cbo(cbo_�շ���Ǯ����).AddItem Split(strTmp, "|")(i)
        cbo(cbo_������Ǯ����).AddItem Split(strTmp, "|")(i)
        cbo(cbo_���ѿ���Ǯ����).AddItem Split(strTmp, "|")(i)
    Next
    cbo(cbo_�Һ���Ǯ����).ListIndex = 0
    cbo(cbo_�շ���Ǯ����).ListIndex = 0
    cbo(cbo_������Ǯ����).ListIndex = 0
    cbo(cbo_���ѿ���Ǯ����).ListIndex = 0
    zlControl.CboSetWidth cbo(cbo_�Һ���Ǯ����).hwnd, 2300
    zlControl.CboSetWidth cbo(cbo_�շ���Ǯ����).hwnd, 2300
    zlControl.CboSetWidth cbo(cbo_������Ǯ����).hwnd, 2300
    zlControl.CboSetWidth cbo(cbo_���ѿ���Ǯ����).hwnd, 2300
    
    '�Զ�����
    With cbo(cbo_�Զ�����ģʽ)
        .AddItem "0-��׼����ģʽ": .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        .AddItem "1-��������Ƭ������ģʽ": .ItemData(.NewIndex) = 1
    End With
    lblAutoChargeNM.Visible = False
    lblAutoChargeNM.Caption = "" & "" & _
        "1.��λ:  ���벻�Ƴ�" & vbCrLf & _
        " 2.������������:  ��Ժ���찴һ�����,��Ժ��������12��֮ǰ����죬12��֮����һ��" & vbCrLf & _
        " 3.������;������(ת�ƣ�ת�������ȼ��䶯��),12����ǰ����ת�����Ϊ׼;12���Ժ���ת������Ϊ׼"
    '������˷�ʽ:49501
    With cbo(cbo_������˷�ʽ)
        .Clear
        .AddItem "0-δ��˲��������": .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        .AddItem "1-���ʱ����������ú�ҽ��": .ItemData(.NewIndex) = 1
    End With
        
    cbo(cbo_δ�󵥾ݽ���).AddItem "0-�����"
    cbo(cbo_δ�󵥾ݽ���).AddItem "1-��鲢��ʾ"
    cbo(cbo_δ�󵥾ݽ���).AddItem "2-��鲢��ֹ"
    cbo(cbo_δ�󵥾ݽ���).ListIndex = 0
    zlControl.CboSetWidth cbo(cbo_δ�󵥾ݽ���).hwnd, 2000
    
    '�ٴ����ﰲ��
    With cbo(cbo_�ٴ�����_ȫԺͨ�ú�Դ����վ��)
        .Clear
        .AddItem ""
        strTmp = "Select Distinct b.���,b.���� From ���ű� A,Zlnodelist B Where a.վ��=b.��� Order By b.���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption)
        Do While Not rsTmp.EOF
            .AddItem rsTmp!��� & "-" & rsTmp!����
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    With cbo(cbo_�ٴ�����_����ȽϷ�ʽ)
        .Clear
        .AddItem "0-���ַ��Ƚ�":  .ItemData(.NewIndex) = 0
        .AddItem "1-����ֵ�Ƚ�": .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    zlControl.CboSetWidth cbo(cbo_�ٴ�����_����ȽϷ�ʽ).hwnd, cbo(cbo_�ٴ�����_����ȽϷ�ʽ).Width * 4 / 3
    
    '�Һ����
    With cbo(cbo_�Һ�_ȱʡ����ʽ)
        .Clear
        .AddItem "0.�ű�"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
        .AddItem "1.����-��Ŀ"
        .ItemData(.NewIndex) = 1
        .AddItem "2.����"
        .ItemData(.NewIndex) = 2
    End With
    
    With cbo(cbo_�Һ�_ȱʡԤԼ��ʽ)
        .Clear
        strTmp = "Select ����,���� From ԤԼ��ʽ"
        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption)
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Loop
    End With
    
    With cbo(cbo_�Һ�_ԤԼ��Чʱ��)
        .Clear
        .AddItem "��ǰ"
        .AddItem "�Ӻ�"
    End With
    
    'Ʊ������
    lvw(lvw_Ʊ��).ListItems.Add , "C1", "�շ��վ�"
    lvw(lvw_Ʊ��).ListItems.Add , "C2", "Ԥ���վ�"
    lvw(lvw_Ʊ��).ListItems.Add , "C3", "�����վ�"
    lvw(lvw_Ʊ��).ListItems.Add , "C4", "�Һ��վ�"
    
    'ˢ��Ҫ����������ĳ���
    With lst(lst_ˢ������)
        .AddItem "����Һ�"
        .AddItem "���ﻮ��"
        .AddItem "�����շ�"
        .AddItem "�������"
        .AddItem "��Ժ�Ǽ�"
        .AddItem "סԺ����"
        .AddItem "���˽���"
        .AddItem "����Ԥ����"
        .AddItem "���鼼ʦվ"
        .AddItem "Ӱ��ҽ��վ"
        .ListIndex = 0
    End With
    
    strTmp = "����|1400|1,��λ��|630|4,��������|1000|1,�����|630|4,��������|1000|1,��λ��ԭʼ��������|0|4,�����ԭʼ��������|0|4"
    Call zlControl.MshSetFormat(mshAutoCalc, strTmp, Me.Caption)
    With mshAutoCalc
        .ColAlignmentFixed(0) = 1
        
        '�б�����ж���
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
    End With
    
    strTmp = "����|1400|1,�շ�ϸĿID|0|1,�շ���Ŀ|1000|4,���㷽ʽ|1000|1,��������|1000|1"
    Call zlControl.MshSetFormat(Bill(bill_�Զ�����), strTmp, Me.Caption)
    With Bill(bill_�Զ�����)
        .ColData(0) = 3 '�����еĿɲ�������
        .ColData(1) = 5
        .ColData(2) = 1
        .ColData(3) = 0
        .ColData(4) = 4
        
        .PrimaryCol = 0
        .Active = True
    End With
    
    strTmp = "����|1500|1,��������|1000|1,����ֵ|800|7,������ʽ1|1300|1,������ʽ2|1300|1,������ʽ3|1300|1,�߿�����|1000|7,�߿��׼|1000|7"
    Call zlControl.MshSetFormat(Bill(bill_���ʱ���), strTmp, Me.Caption)
    With Bill(bill_���ʱ���)
        .ColData(0) = 3
        .ColData(1) = 0
        .ColData(2) = 4
        .ColData(3) = 1
        .ColData(4) = 1
        .ColData(5) = 1
        .ColData(6) = 4
        .ColData(7) = 4
        
        .PrimaryCol = 0
        .Active = True
    End With
    With cbo(cbo_�շ�_�Զ���ϵ���)
        .Clear
        .AddItem "�շ����": .ListIndex = 0
        .AddItem "ִ�п���"
    End With
    Call InitPage(Pg_�Һ�ҵ��)  '��ʼҳ��
    Call InitPage(Pg_�����շ�)  '��ʼҳ��
    Call InitPage(Pg_����ҵ��)  '��ʼҳ��
    
    mblnExistPrintData = GetPrintListHaveData
    With cbo(cbo_�շ�_Ʊ�ݷ������)
        .Clear
        .AddItem "1-����ʵ�ʴ�ӡ����Ʊ��"
        .ListIndex = .NewIndex
        .AddItem "2-����Ԥ���������Ʊ��"
        .AddItem "3-�����Զ���������Ʊ��"
        .Enabled = Not mblnExistPrintData
    End With
    Call InitBillRuleCtrl
    Call SetBillRuleParaLocale
    
    With cbo(cbo_���ʲ���_���ʺ�ҩ)
        .AddItem "0-����ҩ": .ListIndex = .NewIndex
        .AddItem "1-�Զ���ҩ"
        .AddItem "2-��ʾ��ҩ"
    End With
    
    With cbo(cbo_����_��Լ��λ���ʴ�ӡ)
        .Clear
        If GetBillUseTypeRec(rsTemp) Then
            rsTemp.Filter = "ID<>0"
            If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!����)
                rsTemp.MoveNext
            Loop
        End If
    End With
    Call Load������㷽ʽ
    Call Load�Էѷ������
    Call Load�ѻ�ҽ�����㷽ʽ
    Call Loadһ��ͨƱ�ݸ�ʽ
    Call Load�������תסԺԤ��Ʊ��ʽ
End Sub

Private Sub pic��ǰ��ɫ_Click()
    Dim strColor As String
    dlgColor.Color = pic��ǰ��ɫ.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic��ǰ��ɫ.BackColor = strColor
    Call SetParChange(pic��ǰ��ɫ, 0, mrsPar, True, strColor)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK Then
        mrsPar.Filter = "(�޸�״̬=1 ANd ErrType =Null) OR  (�޸�״̬=1 And ErrType=" & PET_ֵ���� & ")"
        If mrsPar.RecordCount > 0 Or mshAutoCalc.Tag = "���޸�" Or Bill(bill_�Զ�����).Tag = "���޸�" _
            Or Bill(bill_���ʱ���).Tag = "���޸�" Or lvw(lvw_����).Tag = "���޸�" Or fra�ض��շ���Ŀ.Tag = "���޸�" Then
            
            If MsgBox("�����޸Ĳ��ֲ����������������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    
    SaveFlexState Bill(bill_�Զ�����), App.ProductName & "\" & Me.Name & bill_�Զ�����
    SaveFlexState Bill(bill_���ʱ���), App.ProductName & "\" & Me.Name & bill_���ʱ���
    SaveFlexState mshAutoCalc, App.ProductName & "\" & Me.Name
    
    Set mrsWarn = Nothing
    Set mrs��� = Nothing
    Set mrsPar = Nothing
    Set mrsBillUseType = Nothing
End Sub

Private Sub cmdOK_Click()
    If ValidateData() = False Then Exit Sub

    Call Save�Զ��Ƽ���Ŀ
    Call Save���ʱ�����
    
    Call Save���ݲ���
    Call Save�շ��ض���Ŀ
    If SaveƱ�ݷ������ = False Then Exit Sub
    Call SaveTriageQueuingDep
    If SavePar(mrsPar, Me) = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If lst���.Visible Then
            lst���.Visible = False
            Bill(bill_���ʱ���).SetFocus
        End If
    End If
End Sub



Private Sub cbo_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    
    If Not Me.Visible Then Exit Sub
    
    Select Case Index
    Case cbo_�Һ���Ǯ����, cbo_�շ���Ǯ����, cbo_������Ǯ����, cbo_���ѿ���Ǯ����
        blnValue = True
        strValue = Split(cbo(cbo_�Һ���Ǯ����).Text & "-", "-")(0) & cbo(cbo_�շ���Ǯ����).ListIndex & _
            cbo(cbo_������Ǯ����).ListIndex & cbo(cbo_���ѿ���Ǯ����).ListIndex
        Call SetParChange(cbo, cbo_�Һ���Ǯ����, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_�շ���Ǯ����, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_������Ǯ����, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_���ѿ���Ǯ����, mrsPar, blnValue, strValue)
        Exit Sub
    Case cbo_�Զ�����ģʽ
        blnValue = True
        With cbo(cbo_�Զ�����ģʽ)
            If .ListIndex >= 0 Then
               strValue = .ItemData(.ListIndex)
            Else
                strValue = "0"
            End If
        End With
        Call SetParChange(cbo, cbo_�Զ�����ģʽ, mrsPar, blnValue, strValue)
        chk(chk_���������ģʽ).Visible = InStr(1, ",0,2,", "," & strValue & ",") > 0
        opt����(0).Enabled = InStr(1, ",0,2,", "," & strValue & ",") > 0
        opt����(1).Enabled = InStr(1, ",0,2,", "," & strValue & ",") > 0
        lblAutoChargeNM.Visible = Val(strValue) = 1

    Case cbo_�շ�_�Զ���ϵ���
        blnValue = True
        strValue = "0"
        If chk(chk_�շ�_�Զ���ϵ���).value = 1 Then
            strValue = cbo(cbo_�շ�_�Զ���ϵ���).ListIndex + 1
        End If
        Call SetParChange(chk, chk_�շ�_�Զ���ϵ���, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_�շ�_�Զ���ϵ���, mrsPar, blnValue, strValue)
        Exit Sub
    Case cbo_�շ�_Ʊ�ݷ������
        Call SetBillRuleParaLocale
        Call SaveBillRuleChange
        Exit Sub
    Case cbo_����_��Լ��λ���ʴ�ӡ
        Call SetParChange(cbo, cbo_����_��Լ��λ���ʴ�ӡ, mrsPar, True, Trim(cbo(Index).Text))
        lblUnit.ForeColor = cbo(Index).ForeColor
        Exit Sub
    Case cbo_�Һ�_ȱʡԤԼ��ʽ
        strValue = zlCommFun.GetNeedName(cbo(cbo_�Һ�_ȱʡԤԼ��ʽ).Text, "-")
        Call SetParChange(cbo, cbo_�Һ�_ȱʡԤԼ��ʽ, mrsPar, True, strValue)
        Exit Sub
    Case cbo_�Һ�_ԤԼ��Чʱ��
        blnValue = True
        strValue = IIF(cbo(cbo_�Һ�_ԤԼ��Чʱ��).ListIndex = 0, 1, -1) * Val(txt(txt_�Һ�_ԤԼ��Чʱ��))
        Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
        txt(txt_�Һ�_ԤԼ��Чʱ��).ForeColor = cbo(Index).ForeColor
        lblAvailabilityTimes.ForeColor = cbo(Index).ForeColor
        Exit Sub
    Case cbo_�ٴ�����_ȫԺͨ�ú�Դ����վ��
        strValue = zlStr.NeedCode(cbo(Index).Text, "-")
        Call SetParChange(cbo, Index, mrsPar, True, strValue)
        Exit Sub
    End Select
    Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
End Sub

Private Sub SaveBillRuleChange()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ز��������ĸı�
    '����:���˺�
    '����:2015-06-18 10:22:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValue As Boolean, strValue As String, intBillRull As Integer
    
    On Error GoTo ErrHandle
    If Not Me.Visible Then Exit Sub
    
    With cbo(cbo_�շ�_Ʊ�ݷ������)
        intBillRull = IIF(.ListIndex < 0, 0, .ListIndex)
    End With
    
    strValue = intBillRull & "||"
    '�ֵ���
    strValue = strValue & IIF(chkBillRule(0).value = 1, 1, 0)
    'ִ�п���
    strValue = strValue & ";" & IIF(chkBillRule(1).value = 1, Val(txtBillRuleNum(0).Text), 0)
    '�վݷ�Ŀ
    strValue = strValue & ";" & IIF(chkBillRule(2).value = 1, Val(txtBillRuleNum(1).Text), 0)
    '�շ�ϸĿ
    strValue = strValue & ";" & IIF(chkBillRule(3).value = 1, Val(txtBillRuleNum(2).Text), 0)
    '��������
    strValue = strValue & ";" & IIF(optRuleTotal(0).value, 0, IIF(optRuleTotal(1).value, 1, 2))
    blnValue = True
    Call SetParChange(cbo, cbo_�շ�_Ʊ�ݷ������, mrsPar, blnValue, strValue)
    Call SetParChange(chkBillRule, 0, mrsPar, blnValue, strValue)
    Call SetParChange(chkBillRule, 1, mrsPar, blnValue, strValue)
    Call SetParChange(chkBillRule, 2, mrsPar, blnValue, strValue)
    Call SetParChange(chkBillRule, 3, mrsPar, blnValue, strValue)
    Call SetParChange(txtBillRuleNum, 0, mrsPar, blnValue, strValue)
    Call SetParChange(txtBillRuleNum, 1, mrsPar, blnValue, strValue)
    Call SetParChange(txtBillRuleNum, 2, mrsPar, blnValue, strValue)
    Call SetParChange(optRuleTotal, 0, mrsPar, blnValue, strValue)
    optRuleTotal(0).ForeColor = optRuleTotal(0).ForeColor
    optRuleTotal(1).ForeColor = optRuleTotal(0).ForeColor
    optRuleTotal(2).ForeColor = optRuleTotal(0).ForeColor
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case chk_�Һ�_�Զ�ˢ�¹ҺŰ���
        Call SetParTip(txt, txt_�Һ�_ˢ��ʱ��, mrsPar)
    Case chk_ר�ҺŹҺ�����
        Call SetParTip(txt, txt_ר�ҺŹҺ�����, mrsPar)
    Case chk_ר�Һ�ԤԼ����
        Call SetParTip(txt, txt_ר�Һ�ԤԼ����, mrsPar)
    Case Else
        Call SetParTip(chk, Index, mrsPar)
    End Select
End Sub

Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lst, Index, mrsPar)
End Sub

Private Sub lvw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lvw, Index, mrsPar)
End Sub

Private Sub opt����_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt����, Index, mrsPar)
    End If
End Sub

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt����, Index, mrsPar)
End Sub

Private Sub optRegist_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optRegist, Index, mrsPar)
    End If
End Sub

Private Sub optRegist_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRegist_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optRegist, Index, mrsPar)
End Sub

Private Sub optInExseCharge_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optInExseCharge, Index, mrsPar)
    End If
End Sub

Private Sub optInExseCharge_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optInExseCharge_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optInExseCharge, Index, mrsPar)
End Sub

Private Sub optPrintFact_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optPrintFact, Index, mrsPar)
    End If
End Sub

Private Sub optPrintFact_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintFact_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintFact, Index, mrsPar)
End Sub

Private Sub optPrintSlip_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optPrintSlip, Index, mrsPar)
    End If
End Sub

Private Sub optPrintSlip_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Index = 2 Then
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus: Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub


Private Sub pic�Ը��ϼ�ɫ_Click()
    Dim strColor As String
    dlgColor.Color = pic�Ը��ϼ�ɫ.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic�Ը��ϼ�ɫ.BackColor = strColor
    Call SetParChange(pic�Ը��ϼ�ɫ, 0, mrsPar, True, strColor)
End Sub

Private Sub pic�ɿ����ɿ�ɫ_Click()
    Dim strColor As String
    dlgColor.Color = pic�ɿ����ɿ�ɫ.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic�ɿ����ɿ�ɫ.BackColor = strColor
    strColor = strColor & "|" & pic�ɿ����˿�ɫ.BackColor
    Call SetParChange(pic�ɿ����ɿ�ɫ, 0, mrsPar, True, strColor)
End Sub

Private Sub pic�ɿ����˿�ɫ_Click()
    Dim strColor As String
    dlgColor.Color = pic�ɿ����˿�ɫ.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic�ɿ����˿�ɫ.BackColor = strColor
    strColor = pic�ɿ����ɿ�ɫ.BackColor & "|" & strColor
    Call SetParChange(pic�ɿ����˿�ɫ, 0, mrsPar, True, strColor)
End Sub

Private Sub pic��ǰ����δ��ɫ_Click()
    Dim strColor As String
    dlgColor.Color = pic��ǰ����δ��ɫ.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic��ǰ����δ��ɫ.BackColor = strColor
    strColor = strColor & "|" & pic��ǰ����δ��ɫ.BackColor
    Call SetParChange(pic��ǰ����δ��ɫ, 0, mrsPar, True, strColor)
End Sub

Private Sub pic��ǰ����δ��ɫ_Click()
    Dim strColor As String
    dlgColor.Color = pic��ǰ����δ��ɫ.BackColor
    dlgColor.ShowColor
    strColor = dlgColor.Color
    pic��ǰ����δ��ɫ.BackColor = strColor
    strColor = pic��ǰ����δ��ɫ.BackColor & "|" & strColor
    Call SetParChange(pic��ǰ����δ��ɫ, 0, mrsPar, True, strColor)
End Sub

Private Sub pic�Ը��ϼ�ɫ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic�Ը��ϼ�ɫ, 0, mrsPar)
End Sub

Private Sub pic�ɿ����ɿ�ɫ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic�ɿ����ɿ�ɫ, 0, mrsPar)
End Sub

Private Sub pic�ɿ����˿�ɫ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic�ɿ����˿�ɫ, 0, mrsPar)
End Sub

Private Sub pic��ǰ����δ��ɫ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic��ǰ����δ��ɫ, 0, mrsPar)
End Sub

Private Sub pic��ǰ����δ��ɫ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic��ǰ����δ��ɫ, 0, mrsPar)
End Sub

Private Sub optPrintSlip_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintSlip, Index, mrsPar)
End Sub

Private Sub optPrintAppoint_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(optPrintAppoint, Index, mrsPar)
    End If
End Sub

Private Sub optPrintAppoint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintAppoint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintAppoint, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    If mblnNotChange Then Exit Sub
    If Not Me.Visible Then Exit Sub
    
    Select Case Index
    Case txt_���������
        blnValue = True
        strValue = IIF(Val(txt(Index).Text) = 0, "", Val(txt(Index).Text))
    Case txt_����ģ����������
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_�Һ�_ˢ��ʱ��
        strValue = Val(txt(Index).Text)
        Call SetParChange(txt, Index, mrsPar, True, strValue)
        chk(chk_�Һ�_�Զ�ˢ�¹ҺŰ���).ForeColor = txt(Index).ForeColor
        Exit Sub
    Case txt_�Һ�_������������
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_�Һ�_ԤԼ����ʱ��_����
        blnValue = True
        strValue = Val(txt(txt_�Һ�_ԤԼ����ʱ��_����).Text) & "|" & Val(txt(Index).Text)
    Case txt_�Һ�_���˹Һſ�������
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_�Һ�_����ԤԼ������
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_�Һ�_����ͬ����ԼN����
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_ר�ҺŹҺ�����
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_ר�Һ�ԤԼ����
        blnValue = True
        strValue = Val(txt(Index).Text)
    Case txt_�Һ�_����ͬ���޹�N����
        blnValue = True
        strValue = Val(txt(Index).Text) & "|" & IIF(chk(chk_�Һ�_����ͬ���޹�N����_����).value = 1, "1", "0")
    Case txt_�Һ�_ԤԼ����ʱ��_����
        blnValue = True
        strValue = Val(txt(Index).Text) & "|" & Val(txt(txt_�Һ�_ԤԼ����ʱ��_����).Text)
    Case txt_�Һ�_N���ڲ���ȡ��ԤԼ��
        If Val(txt(Index).Text) = 0 Then
            chk(chk_�Һ�_N�����˺������).Caption = "������ȡ��ԤԼ��Ҫͨ�����"
        Else
            chk(chk_�Һ�_N�����˺������).Caption = "��" & txt(Index).Text & "����ȡ��ԤԼ��Ҫͨ�����"
        End If
    Case txt_�շ�_�������չҺŷ�
        blnValue = True: strValue = ""
        cmdAddedItem.Tag = ""
        If chk(chk_�շ�_δ�Һ��Զ����չҺŷ�).value = 1 Then strValue = cmdAddedItem.Tag & ";" & txt(Index).Text
        Call SetParChange(txt, Index, mrsPar, blnValue, strValue)
        Call SetParChange(chk, chk_�շ�_δ�Һ��Զ����չҺŷ�, mrsPar, blnValue, strValue)
        Exit Sub
    Case txt_������_����ģ����������
        strValue = IIF(chk(chk_������_����ģ������).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txt(txt_������_����ģ����������).Text)
        Call SetParChange(chk, chk_������_����ģ������, mrsPar, True, strValue)
        Call SetParChange(txt, txt_������_����ģ����������, mrsPar, True, strValue)
        Exit Sub
    Case txt_�Һ�_ԤԼ��Чʱ��
        blnValue = True
        strValue = IIF(cbo(cbo_�Һ�_ԤԼ��Чʱ��).ListIndex = 0, 1, -1) * Val(txt(txt_�Һ�_ԤԼ��Чʱ��))
    Case txt_�Һ�_����ͬһ��Դ�޹�N����
        blnValue = True
        strValue = Val(txt(Index).Text)
    End Select
    
    Call SetParChange(txt, Index, mrsPar, blnValue, strValue)
    
    '���ñ�ǩ��ɫ
    Select Case Index
    Case txt_�Һ�_N���ڲ���ȡ��ԤԼ��
        lblCancelBespeak.ForeColor = txt(Index).ForeColor
    Case txt_�Һ�_ԤԼ����ʱ��_����
        lblBespeakDefaultDays.ForeColor = txt(Index).ForeColor
    Case txt_�Һ�_ԤԼ����ʱ��_����
        lblBespeakMinTime.ForeColor = txt(Index).ForeColor
    Case txt_�Һ�_ԤԼ����ʱ��_����
        lblBespeakMinTime.ForeColor = txt(Index).ForeColor
    Case txt_�Һ�_ԤԼ��Чʱ��
        lblAvailabilityTimes.ForeColor = txt(Index).ForeColor
        cbo(cbo_�Һ�_ԤԼ��Чʱ��).ForeColor = txt(Index).ForeColor
    Case txt_�Һ�_ԤԼʧЧ����
        lblBreakAnAppointmentNums.ForeColor = txt(Index).ForeColor
    Case txt_�Һ�_N����������໤��
        lblGuardian.ForeColor = txt(Index).ForeColor
    Case Else
    End Select
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Select Case Index
    Case txt_����_���ﵥ�������, txt_�շ�_���ﵥ�������
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
    End Select
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub

Private Sub txtBillRuleNum_Change(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SaveBillRuleChange
End Sub

Private Sub txtBillRuleNum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtBillRuleNum_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtBillRuleNum, Index, mrsPar)
End Sub

Private Sub txtUD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtUD, Index, mrsPar)
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub

Private Sub dtpRegistPlanMode_Validate(Cancel As Boolean)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strValue As String
    If mblnInstantActive Then Exit Sub
    strSQL = "Select 1 From ���˹Һż�¼ Where ����ʱ�� > [1] And ��¼״̬=1 And �����¼ID Is Null And Rownum <2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(dtpRegistPlanMode.value))
    If Not rsTmp.EOF Then
        MsgBox "��������֮����ڼƻ��Ű�ģʽ�ĹҺŻ���ԤԼ��¼,�������������!", vbInformation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    strValue = "1|" & Format(dtpRegistPlanMode.value, "yyyy-mm-dd hh:mm:ss")
    Call SetParChange(optRegistPlanMode, 0, mrsPar, True, strValue)
End Sub

Private Sub optRegistPlanMode_Click(Index As Integer)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strValue As String
    If mblnInstantActive Then Exit Sub
    If mblnNotChange Then Exit Sub
    If optRegistPlanMode(0).value = 1 Or optRegistPlanMode(0).value = True Then
        strSQL = "Select 1 From ���˹Һż�¼ Where �����¼ID Is Not Null And ��¼״̬=1 And Rownum < 2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            MsgBox "�Ѿ����ڳ�����Ű�ģʽ�µĹҺż�¼,�������л��ؼƻ��Ű�ģʽ!", vbInformation, gstrSysName
            mblnNotChange = True
            optRegistPlanMode(1).value = 1
            mblnNotChange = False
            Exit Sub
        End If
        strValue = 0
        Call SetParChange(optRegistPlanMode, 0, mrsPar, True, strValue)
        
        fraNewPaln.Visible = False
        chk(chk_ֻ��ҽ��ҽ�����йҺŰ���).Visible = True
        chk(chk_����ԤԼ�Һŵ���ֹɾ������).Visible = True
    Else
        '������Ű�ģʽ
        strSQL = "Select 1 From �ٴ������ Where ����ʱ�� Is Not Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTmp.EOF Then
            MsgBox "�������κ��ٴ���������,�����л�Ϊ������Ű�ģʽ!", vbInformation, gstrSysName
            mblnNotChange = True
            optRegistPlanMode(0).value = 1
            mblnNotChange = False
            Exit Sub
        End If
        dtpRegistPlanMode.Enabled = True
        strValue = "1|" & Format(dtpRegistPlanMode.value, "yyyy-mm-dd hh:mm:ss")
        Call SetParChange(optRegistPlanMode, 0, mrsPar, True, strValue)
        
        fraNewPaln.Visible = True
        chk(chk_ֻ��ҽ��ҽ�����йҺŰ���).Visible = False
        chk(chk_����ԤԼ�Һŵ���ֹɾ������).Visible = False
    End If
End Sub

Private Sub Save���ʱ�����()
    Dim strTmp As String
    Dim i As Integer
    Dim strArr
    Dim str���ò��� As String
    
    If Bill(bill_���ʱ���).Tag = "���޸�" Then
    
        '�ȴ���ɾ�������ò��˼��ʱ���
        On Error GoTo ErrHandle
        If mstrDel���ò��� <> "" Then
            mstrDel���ò��� = mstrDel���ò��� & ";"
            strArr = Split(mstrDel���ò���, ";")
            For i = 0 To UBound(strArr) - 1
                If strArr(i) <> "" Then
                    str���ò��� = strArr(i)
                    strTmp = str���ò��� & "|"
                    gstrSQL = "zl_���ʱ�����_Modify('" & strTmp & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
            Next
        End If
        
        '�����ò��˷�������
        mrsWarn.Filter = 0
        For i = 1 To tab����.Tabs.Count
            strTmp = ""
            str���ò��� = tab����.Tabs.Item(i).Caption
            
            mrsWarn.Filter = "���ò���='" & str���ò��� & "'"
            Do While Not mrsWarn.EOF
                strTmp = strTmp & NVL(mrsWarn!����ID) & "," & mrsWarn!�������� & "," & _
                    mrsWarn!����ֵ & "," & NVL(mrsWarn!������־1) & "," & NVL(mrsWarn!������־2) & "," & NVL(mrsWarn!������־3) & "," & NVL(mrsWarn!�߿�����) & "," & NVL(mrsWarn!�߿��׼) & ","
                mrsWarn.MoveNext
            Loop
            
            strTmp = str���ò��� & "|" & strTmp
            
            gstrSQL = "zl_���ʱ�����_Modify('" & strTmp & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
        
        Bill(bill_���ʱ���).Tag = ""
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save�Զ��Ƽ���Ŀ()
    Dim str����ID As String
    Dim strϸĿID As String
    Dim str�����־ As String
    Dim str�������� As String
    Dim lngTemp As Long, i As Long, blnTrans As Boolean
    
    On Error GoTo ErrHandle
    If mshAutoCalc.Tag = "���޸�" Or Bill(bill_�Զ�����).Tag = "���޸�" Then
    
        gcnOracle.BeginTrans: blnTrans = True
        gstrSQL = "Zl_�Զ��Ƽ���Ŀ_Delete"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ɾ���Զ��Ƽ���Ŀ")
        
        '����λ
        For i = 1 To mshAutoCalc.Rows - 1
            lngTemp = mshAutoCalc.RowData(i)
            If lngTemp <> 0 Then
                If mshAutoCalc.TextMatrix(i, 1) <> "" Then
                    str����ID = str����ID & lngTemp & ","
                    strϸĿID = strϸĿID & ","
                    str�����־ = str�����־ & "1,"
                    str�������� = str�������� & mshAutoCalc.TextMatrix(i, 2) & ","
                End If
                If mshAutoCalc.TextMatrix(i, 3) <> "" Then
                    str����ID = str����ID & lngTemp & ","
                    strϸĿID = strϸĿID & ","
                    str�����־ = str�����־ & "2,"
                    str�������� = str�������� & mshAutoCalc.TextMatrix(i, 4) & ","
                End If
            End If
            If (i Mod 100) = 0 Or i >= mshAutoCalc.Rows - 1 Then
                gstrSQL = "zl_�Զ��Ƽ���Ŀ_Modify('" & str����ID & "','" & strϸĿID & "','" & str�����־ & "','" & str�������� & "' )"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                str����ID = ""
                strϸĿID = ""
                str�����־ = ""
                str�������� = ""
            End If
        Next
        '������
        For i = 1 To Bill(bill_�Զ�����).Rows - 1
            lngTemp = Bill(bill_�Զ�����).RowData(i)
            If lngTemp <> 0 And Bill(bill_�Զ�����).TextMatrix(i, 1) <> "" Then
                If Bill(bill_�Զ�����).TextMatrix(i, 1) <> "" Then
                    str����ID = str����ID & lngTemp & ","
                    strϸĿID = strϸĿID & Bill(bill_�Զ�����).TextMatrix(i, 1) & ","
                    str�����־ = str�����־ & Switch(Left(Bill(bill_�Զ�����).TextMatrix(i, 3), 1) = "1", "6", Left(Bill(bill_�Զ�����).TextMatrix(i, 3), 1) = "2", "8", True, "7") & ","
                    str�������� = str�������� & Bill(bill_�Զ�����).TextMatrix(i, 4) & ","
                End If
            End If
            If (i Mod 100) = 0 Or i >= Bill(bill_�Զ�����).Rows - 1 Then
                gstrSQL = "zl_�Զ��Ƽ���Ŀ_Modify('" & str����ID & "','" & strϸĿID & "','" & str�����־ & "','" & str�������� & "' )"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                str����ID = ""
                strϸĿID = ""
                str�����־ = ""
                str�������� = ""
            End If
        Next
        gcnOracle.CommitTrans: blnTrans = False
        mshAutoCalc.Tag = ""
        Bill(bill_�Զ�����).Tag = ":"
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub


Private Sub Save���ݲ���()
    Dim lst As ListItem
    Dim i As Integer, blnTrans As Boolean
    
    '����ɾ����ǰ�����е��ݲ���
    On Error GoTo ErrHandle
    
    If lvw(lvw_����).Tag = "���޸�" Then
        gcnOracle.BeginTrans: blnTrans = True
        gstrSQL = "zl_���ݲ�������_Delete"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        '�������µ�
        For Each lst In lvw(lvw_����).ListItems
            gstrSQL = "zl_���ݲ�������_Insert(" & lst.Tag & "," & lst.ListSubItems(1).Tag & _
                        "," & lst.SubItems(2) & "," & IIF(lst.SubItems(3) = "��", 1, 0) & "," & IIF(lst.SubItems(4) = "", "NULL", lst.SubItems(4)) & " )"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        lvw(lvw_����).Tag = ""
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Save�շ��ض���Ŀ()
    Dim strTmp As String
    
    If fra�ض��շ���Ŀ.Tag = "���޸�" Then
        '����Բ������б���
        On Error GoTo ErrHandle
        If txtCmd(0).Text <> "" Then
            strTmp = "������," & txtCmd(0).Tag & ","
        End If
        If txtCmd(1).Text <> "" Then
            strTmp = strTmp & "������," & txtCmd(1).Tag & ","
        End If
        
        If txtCmd(3).Text <> "" Then
            strTmp = strTmp & "��ͨ���÷�," & txtCmd(3).Tag & ","
        End If
        
        If txtCmd(4).Text <> "" Then
            strTmp = strTmp & "�������÷�," & txtCmd(4).Tag & ","
        End If
        
        If strTmp <> "" Then
            gstrSQL = "zl_�շ��ض���Ŀ_Modify('" & strTmp & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
        
        fra�ض��շ���Ŀ.Tag = ""
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Check���ʱ���() As Boolean
    Dim lngRow As Long, lngTemp As Long
    Dim lngCol1 As Long, lngCol2 As Long
    Dim arr���() As String
        
    With Bill(bill_���ʱ���)
        For lngRow = 1 To .Rows - 2
            If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                For lngTemp = lngRow + 1 To .Rows - 1
                    If .TextMatrix(lngRow, 0) = .TextMatrix(lngTemp, 0) And .TextMatrix(lngTemp, 2) <> "" Then
                        MsgBox "������" & .TextMatrix(lngTemp, 0) & "�����ֶ�Ρ�", vbExclamation, gstrSysName
                        .Row = lngTemp: .Col = 0: .SetFocus: Exit Function
                    End If
                Next
                '���˺� ����: 34770   ����:2010-12-21 10:54:02
                If Val(.TextMatrix(lngRow, 6)) > 999999999 Or Val(.TextMatrix(lngRow, 6)) < 0 Then
                    MsgBox "������" & .TextMatrix(lngRow, 0) & "���еĴ߿�������������(Ӧ����0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 6: .SetFocus: Exit Function
                End If
                If Val(.TextMatrix(lngRow, 7)) > 999999999 Or Val(.TextMatrix(lngRow, 7)) < 0 Then
                    MsgBox "������" & .TextMatrix(lngRow, 0) & "���еĴ߿��׼����(Ӧ����0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 7: .SetFocus: Exit Function
                End If
                
            End If
        Next
        
        '���ͬһ������ͬ������ʽ������Ƿ�һ����û�����û��ظ�
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                If Trim(.TextMatrix(lngRow, 3)) = "" And Trim(.TextMatrix(lngRow, 4)) = "" And Trim(.TextMatrix(lngRow, 5)) = "" Then
                    MsgBox "������" & .TextMatrix(lngRow, 0) & "��δ����Ҫ�������շ����", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                End If
                If (.TextMatrix(lngRow, 3) = "�������" And (Trim(.TextMatrix(lngRow, 4)) <> "" Or Trim(.TextMatrix(lngRow, 5)) <> "")) _
                    Or (.TextMatrix(lngRow, 4) = "�������" And (Trim(.TextMatrix(lngRow, 3)) <> "" Or Trim(.TextMatrix(lngRow, 5)) <> "")) _
                    Or (.TextMatrix(lngRow, 5) = "�������" And (Trim(.TextMatrix(lngRow, 4)) <> "" Or Trim(.TextMatrix(lngRow, 3)) <> "")) Then
                    
                    MsgBox "������" & .TextMatrix(lngRow, 0) & "����ͬ�ı�����ʽ������ͬ���շ����", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                End If
                If .TextMatrix(lngRow, 3) <> "�������" And Trim(.TextMatrix(lngRow, 4)) <> "�������" And Trim(.TextMatrix(lngRow, 5)) <> "�������" Then
                    For lngCol1 = 3 To 5
                        If Trim(.TextMatrix(lngRow, lngCol1)) <> "" Then
                            For lngCol2 = 3 To 5
                                If lngCol1 <> lngCol2 Then
                                    arr��� = Split(.TextMatrix(lngRow, lngCol1), ",")
                                    For lngTemp = 0 To UBound(arr���)
                                        If InStr("," & .TextMatrix(lngRow, lngCol2) & ",", "," & arr���(lngTemp) & ",") > 0 Then
                                            MsgBox "������" & .TextMatrix(lngRow, 0) & "����ͬ�ı�����ʽ������ͬ���շ����", vbExclamation, gstrSysName
                                            .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Next
    End With
    
    Check���ʱ��� = True
End Function

Private Function ValidateData() As Boolean
    Dim lngRow As Long, lngTmp As Long
    Dim lngIndex As Long, strTmp As String
    Dim i As Integer
    
    
    '����Զ�������Ŀ�Ƿ��ظ�
    With Bill(bill_�Զ�����)
        For lngRow = 1 To .Rows - 2
            If .RowData(lngRow) > 0 And .TextMatrix(lngRow, 1) <> "" Then
                For lngTmp = lngRow + 1 To .Rows - 1
                    If .RowData(lngRow) = .RowData(lngTmp) And .TextMatrix(lngRow, 1) = .TextMatrix(lngTmp, 1) Then
                        MsgBox "����Ϊ��" & .TextMatrix(lngTmp, 0) & "�����շ�ϸĿΪ��" & _
                            .TextMatrix(lngTmp, 2) & "��" & vbCrLf & "������ϳ��ֶ�Ρ�", vbExclamation, gstrSysName
                        .Row = lngTmp
                        .Col = 0
                        .SetFocus
                        Exit Function
                    End If
                Next
            End If
        Next
    End With
    
    '����Զ�������Ŀ����������
    With Bill(bill_�Զ�����)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 And .TextMatrix(lngRow, 1) <> "" Then
                If Not IsDate(.TextMatrix(lngRow, 4)) Then
                    MsgBox "�Զ�������Ŀ����������δ���û����ڸ�ʽ����ȷ��", vbInformation, gstrSysName
                    .Row = lngRow
                    .Col = 4
                    .SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
   
    
    If CheckParChanged(txtUD, ud_���ý���λ��, mrsPar) Then
        If MsgBox("���ѵ����˷��ý���С��λ�����ܻ�����С���������Ƿ������", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
         
            Exit Function
        End If
    End If
    
    If CheckParChanged(txtUD, ud_���õ��۱���λ��, mrsPar) Then
        If MsgBox("���ѵ����˷��õ��۱���С��λ�����ܻ�����С���������Ƿ������", vbYesNo + vbQuestion, gstrSysName) = vbNo Then

            Exit Function
        End If
    End If
    If VsaliedData_�շ� = False Then Exit Function
    
    ValidateData = True
End Function
Private Function VsaliedData_�շ�() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����շ���ز������õĺϷ���
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-06-18 11:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
    
    If cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex = 1 Then
        If chkBillRule(0).value = 0 And chkBillRule(1).value = 0 And chkBillRule(2).value = 0 And chkBillRule(3).value = 0 Then
            MsgBox "ע��:" & vbCrLf & "    Ʊ�ݺŷ�����򰴡�" & cbo(cbo_�շ�_Ʊ�ݷ������).Text & "���ı�������һ�ַ������,����!", vbInformation + vbOKOnly
           ' stab.Tab = 3
            If chkBillRule(0).Enabled And chkBillRule(0).Visible Then chkBillRule(0).SetFocus
            Exit Function
        End If
    End If
    VsaliedData_�շ� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function VsaliedData_����() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ز������õĺϷ���
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-06-18 11:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
    With cbo(cbo_����_��Լ��λ���ʴ�ӡ)
        If .ListIndex < 0 Then
            MsgBox "ע��:" & vbCrLf & _
                   "    ��δѡ���Լ��λ����ʱ��ʹ�õĺ���Ʊ��!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With
    VsaliedData_���� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub tab����_Click()
    Dim lngRow As Long
    
    mrsWarn.Filter = "���ò���='" & tab����.SelectedItem.Caption & "'"
    
    With Bill(bill_���ʱ���)
        If mrsWarn.RecordCount = 0 Then
            .ClearBill
            mlngPreFind = 1
            .Rows = 2: .Row = 1: .Col = 1
            .RowData(1) = 0
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
            .TextMatrix(1, 6) = ""
            .TextMatrix(1, 7) = ""
        Else
            .ClearBill
            mlngPreFind = 1
            .Rows = mrsWarn.RecordCount + 1: .Row = 1: .Col = 1
            lngRow = 1
            Do Until mrsWarn.EOF
                .RowData(lngRow) = NVL(mrsWarn!����ID, 0)
                .TextMatrix(lngRow, 0) = IIF(IsNull(mrsWarn!����ID), "*����*", mrsWarn!������ & "-" & mrsWarn!������)
                .TextMatrix(lngRow, 1) = IIF(mrsWarn!�������� = 1, "1-�ۼƷ���", "2-ÿ�շ���")
                .TextMatrix(lngRow, 2) = Format(mrsWarn!����ֵ, "##########0.00;-##########0.00;0.00;0.00")
                
                .TextMatrix(lngRow, 3) = Get������ƴ�(NVL(mrsWarn!������־1), mrs���)
                .TextMatrix(lngRow, 4) = Get������ƴ�(NVL(mrsWarn!������־2), mrs���)
                .TextMatrix(lngRow, 5) = Get������ƴ�(NVL(mrsWarn!������־3), mrs���)
                .TextMatrix(lngRow, 6) = Format(mrsWarn!�߿�����, "###0.00;-###0.00;0.00;0.00")
                .TextMatrix(lngRow, 7) = Format(mrsWarn!�߿��׼, "###0.00;-###0.00;0.00;0.00")
                
                lngRow = lngRow + 1
                mrsWarn.MoveNext
            Loop
        End If
    End With
End Sub

Private Sub pic��ǰ��ɫ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(pic��ǰ��ɫ, 0, mrsPar)
End Sub

Private Function IsRecord(ByVal strFind As String) As Boolean
'����:�������������Ƿ�����Ч�����ݿ��б�ļ�¼
'����:strFind SQL��������
'����ֵ:��Ч����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    
    rsTemp.CursorLocation = adUseClient
    IsRecord = False
    If InStr(strFind, "'") > 0 Then
        MsgBox "�����˷Ƿ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    gstrSQL = "select distinct A.����,A.����,A.���,A.���㵥λ ,A.id from �շ�ϸĿ A,�շѱ��� B,�շ���� C " & _
         " where A.ID=B.�շ�ϸĿID and A.�Ƿ��� <> 1 and A.ĩ��=1 and  A.���=C.���� and  (A.���� like [1] or B.���� like [2] " & _
         " or  upper(B.����) like [2]) and " & Where����ʱ��("A")
          
    With Bill(bill_�Զ�����)
        If .TextMatrix(.Row, 3) <> "2-����һ��" Then
            gstrSQL = gstrSQL & " and C.���� Not In('4','5','6','7') "
        End If
    End With
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFind & "%", "%" & UCase(strFind) & "%")
    
    If rsTemp.RecordCount < 1 Then Exit Function
    If rsTemp.RecordCount > 1 Then
        gstrSQL = ""
        gstrSQL = frmSelCurr.ShowCurrSel(Me, rsTemp, "����,1000,0,2;����,1800,0,1;���,2300,0,2;���㵥λ,1000,0,2;id,0,0,2", -1, "ѡ���շ�ϸĿ")
        If gstrSQL = "" Then
            Exit Function
        End If
        If Bill(bill_�Զ�����).TextMatrix(Bill(bill_�Զ�����).Row, 3) <> "2-����һ��" Then
            If Not IsRaiseByDate(Val(Split(gstrSQL, ";")(4))) Then
                MsgBox "��Ŀ[" & Split(gstrSQL, ";")(1) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۡ�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        With Bill(bill_�Զ�����)
            .TextMatrix(.Row, 1) = Split(gstrSQL, ";")(4) ' rsTemp("ID")
            .TextMatrix(.Row, 2) = Split(gstrSQL, ";")(1) 'rsTemp("����")
            If .TextMatrix(.Row, 3) = "" Then
                .TextMatrix(.Row, 3) = "0-��������"
            End If
        End With
    Else
        rsTemp.MoveFirst
        If Bill(bill_�Զ�����).TextMatrix(Bill(bill_�Զ�����).Row, 3) <> "2-����һ��" Then
            If Not IsRaiseByDate(Val(rsTemp!ID)) Then
                MsgBox "��Ŀ[" & rsTemp!���� & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۡ�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        With Bill(bill_�Զ�����)
            .TextMatrix(.Row, 1) = rsTemp("ID")
            .TextMatrix(.Row, 2) = rsTemp("����")
            If .TextMatrix(.Row, 3) = "" Then
                .TextMatrix(.Row, 3) = "0-��������"
            End If
        End With
    End If
    IsRecord = True
End Function

Private Sub bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim rsTmp As New ADODB.Recordset
    Dim lmX As Integer
    Dim lmY As Integer
    Dim strTmp As String
    
    With Bill(Index)
        If Index = bill_�Զ����� Then
            If .Col = 4 And KeyCode = vbKeyReturn Then
                If .Text <> "" And Not IsDate(.Text) Then
                    If Not IsDate(Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)) Then
                        .Text = ""
                        MsgBox "��������ȷ�����ڸ�ʽ(yyyy-mm-dd����yyyymmdd)��", vbInformation, gstrSysName
                    Else
                        .Text = Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)
                    End If
                    .TextMatrix(.Row, .Col) = .Text
                End If
            End If
                
            If .Col = 2 Then
                '�շ�ϸĿ��ֻ����س���
                If KeyCode = vbKeyDelete Then .Tag = "���޸�": Exit Sub   '118682
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TxtVisible = False Then
                    If .TextMatrix(.Row, 2) = "" Then
                        '����һ���ؼ�
                        zlCommFun.PressKey vbKeyTab
                    End If
                Else
                    'ѡ���շ�ϸĿ
                    If IsRecord(.Text) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = .TextMatrix(.Row, 2)
                    If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "0-��������"
                    
                End If
            End If
        End If
        
        If Index = bill_���ʱ��� Then
            If .Col = 2 Then
                '����ֵ��ֻ����س���
                If KeyCode = vbKeyDelete Then .Tag = "���޸�": Exit Sub     '118682
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TxtVisible = False Then
                    If .TextMatrix(.Row, 2) = "" Then
                        '����һ���ؼ�
                        zlCommFun.PressKey vbKeyTab
                    End If
                Else
                    '�ж�����ĺϷ���
                    .Text = Format(.Text, "##########0.00;-##########0.00;0.00;0,00")
                    
                End If
            ElseIf .Col = 3 Then
                '��ֹ���뱨�����
                If KeyCode <> vbKeyReturn And KeyCode <> vbKeyDelete Then KeyCode = 0: Cancel = True
            ElseIf .Col = 6 Or .Col = 6 Then
                .Text = Format(.Text, "###0.00;-###.00;0.00;0,00")
                
            End If
        End If
        
        .Tag = "���޸�"
    End With
End Sub

Private Sub SetDrugStore()
    Dim lngType As Long, strTmp As String, arrTmp As Variant
    Dim i As Long, j As Long, lngRow As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strTmp = "'��ҩ��','��ҩ��','��ҩ��'"
    Set rsTmp = GetDepartments(strTmp, "1,2,3")
    
    With vsfDrugStore(0)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            lngRow = 1
            rsTmp.Filter = "��������='��ҩ��'"
            If rsTmp.RecordCount > 0 Then
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(lngRow, 0) = 0
                    .TextMatrix(lngRow, 1) = rsTmp!����
                    .TextMatrix(lngRow, 2) = "�Զ�����"
                    .RowData(lngRow) = Val(rsTmp!ID)
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Next
            End If
        End If
    End With
    
    With vsfDrugStore(1)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            lngRow = 1
            rsTmp.Filter = "��������='��ҩ��'"
            If rsTmp.RecordCount > 0 Then
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(lngRow, 0) = 0
                    .TextMatrix(lngRow, 1) = rsTmp!����
                    .TextMatrix(lngRow, 2) = "�Զ�����"
                    .RowData(lngRow) = Val(rsTmp!ID)
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Next
            End If
        End If
    End With
    
    With vsfDrugStore(2)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            lngRow = 1
            rsTmp.Filter = "��������='��ҩ��'"
            If rsTmp.RecordCount > 0 Then
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(lngRow, 0) = 0
                    .TextMatrix(lngRow, 1) = rsTmp!����
                    .TextMatrix(lngRow, 2) = "�Զ�����"
                    .RowData(lngRow) = Val(rsTmp!ID)
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Next
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

Private Sub bill_KeyPress(Index As Integer, KeyAscii As Integer)
    With Bill(Index)
        If Index = bill_�Զ����� Then
            If .Col = 4 Then
                .TxtCheck = True
                .TextMask = "0123456789-"
            Else
                .TxtCheck = False
            End If
            If .Col = 3 Then
                Select Case KeyAscii
                    Case Asc(" ")
                        '�л������־
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "0"
                                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                        MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                                .TextMatrix(.Row, .Col) = "1-������"
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-����һ��"
                            Case Else
                                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                                    If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                        MsgBox "ҩƷ�����������Զ����㷽ʽ���ܸı䡣", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                        MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                                .TextMatrix(.Row, .Col) = "0-��������"
                        End Select
                        
                    Case vbKey0
                        If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                            If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                MsgBox "ҩƷ�����������Զ����㷽ʽ���ܸı䡣", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        .TextMatrix(.Row, .Col) = "0-��������"
                        
                    Case vbKey1
                        If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                            If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                MsgBox "ҩƷ�����������Զ��������Ͳ��ܸı䡣", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        .TextMatrix(.Row, .Col) = "1-������"
                        
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-����һ��"
                        
                End Select
            End If
        
        ElseIf Index = bill_���ʱ��� Then
            .TxtCheck = False
            If .Col = 1 Then
                
                '�л���������
                Select Case KeyAscii
                    Case Asc(" ")
                        '�л������־
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-ÿ�շ���"
                            Case Else
                                .TextMatrix(.Row, .Col) = "1-�ۼƷ���"
                        End Select
                        
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-�ۼƷ���"
                        
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-ÿ�շ���"
                        
                End Select
                If InStr(.TextMatrix(.Row, 1), "ÿ�շ���") > 0 Then
                    .TextMatrix(.Row, 4) = ""  'ÿ�շ����ޱ�����ʽ2
                End If
            ElseIf InStr(1, "267", .Col) > 0 Then
                    .TxtCheck = True
                    .TextMask = "0123456789-"
                    .MaxLength = 10
            End If
        End If
        
        .Tag = "���޸�"
    End With

End Sub

Private Sub Set���ѡ��(str��� As String)
'���ܣ���������"���,����..."�Ĵ������б��ѡ�����
    Dim i As Integer, j As Integer
    Dim arr���() As String
    
    For i = 0 To lst���.ListCount - 1
        lst���.Selected(i) = False
    Next
    
    If Trim(str���) = "" Then
        Exit Sub
    ElseIf str��� = "�������" Then
        For i = 0 To lst���.ListCount - 1
            lst���.Selected(i) = (i = 0)
        Next
    Else
        lst���.Selected(0) = False
        arr��� = Split(str���, ",")
        For i = 0 To UBound(arr���)
            For j = 1 To lst���.ListCount - 1
                If lst���.List(j) = arr���(i) Then
                    lst���.Selected(j) = True: Exit For
                End If
            Next
        Next
    End If
    
    For i = 0 To lst���.ListCount - 1
        If lst���.Selected(i) Then
            lst���.TopIndex = i: Exit For
        End If
    Next
End Sub

Private Sub bill_CommandClick(Index As Integer)
'ͨ����ťѡ���շ�ϸĿ
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim rsTmp As New ADODB.Recordset
    
    If Index = bill_���ʱ��� Then
        With Bill(Index)
            Call Set���ѡ��(.TextMatrix(.Row, .Col))
            
            lst���.Left = .Left + .MsfObj.CellLeft
            If .Top + .MsfObj.CellTop + .MsfObj.CellHeight + lst���.Height <= .Container.Height Then
                lst���.Top = .Top + .MsfObj.CellTop + .MsfObj.CellHeight
            Else
                lst���.Top = .Top + .MsfObj.CellTop - lst���.Height - 30
            End If
            lst���.Width = .MsfObj.CellWidth
            lst���.ZOrder
            lst���.Visible = True
            lst���.SetFocus
        End With
    End If
    
    If Index = bill_�Զ����� Then
        With Bill(bill_�Զ�����)
            If .TextMatrix(.Row, 3) <> "2-����һ��" Then
                blnRe = frmChargeListSel.ShowTree(strID, str����, False)
            Else
                blnRe = frmChargeListSel.ShowTree(strID, str����, True)
            End If
            If blnRe And strID <> "" Then
                If .TextMatrix(.Row, 3) <> "2-����һ��" Then
                    If Not IsRaiseByDate(strID) Then
                        MsgBox "��Ŀ[" & str���� & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۡ�", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .SetFocus
                .TextMatrix(.Row, 1) = strID
                .TextMatrix(.Row, 2) = str����
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "0-��������"
            End If
        End With
    End If
    Bill(Index).Tag = "���޸�"
End Sub

Private Sub bill_cboKeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With Bill(Index)
        If .ListIndex < 0 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            .RowData(.Row) = .ItemData(.ListIndex)

            If Index = bill_���ʱ��� Then
                If .TextMatrix(.Row, 1) = "" Then .TextMatrix(.Row, 1) = "1-�ۼƷ���"
            End If
            
            Bill(Index).Tag = "���޸�"
        End If
    End With
End Sub

Private Sub bill_DblClick(Index As Integer, Cancel As Boolean)
'�������һ�еı仯
With Bill(Index)
    If .MouseRow = 0 Then Exit Sub
    
    If Index = bill_�Զ����� Then
        If .MouseCol <> 3 Then Exit Sub
        Select Case Left(.TextMatrix(.Row, .Col), 1)
            Case "0"
                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                        MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "�����������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(.Row, .Col) = "1-������"
            Case "1"
                .TextMatrix(.Row, .Col) = "2-����һ��"
            Case Else
                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                    If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                        MsgBox "ҩƷ�����������Զ����㷽ʽ���ܸı䡣", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                        MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "�����������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(.Row, .Col) = "0-��������"
        End Select
    ElseIf Index = bill_���ʱ��� Then
        If .MouseCol <> .Cols - 1 And .MouseCol <> 1 Then Exit Sub
        If .Col = 1 Then
            .TextMatrix(.Row, 1) = IIF(Left(.TextMatrix(.Row, 1), 1) = "1", "2-ÿ�շ���", "1-�ۼƷ���")
            If InStr(.TextMatrix(.Row, 1), "ÿ�շ���") > 0 Then
                .TextMatrix(.Row, 4) = ""  'ÿ�շ����ޱ�����ʽ2
                
                'Ϊ��ÿ�շ��á�ʱ�ж�һ�½���Ϊ����
                If IsNumeric(.TextMatrix(.Row, 2)) Then
                    If Val(.TextMatrix(.Row, 2)) < 0 Then
                        .TextMatrix(.Row, 2) = "0.00"
                    End If
                Else
                    .TextMatrix(.Row, 2) = "0.00"
                End If
            End If
        End If
    End If
    .Tag = "���޸�"
End With
    
End Sub


Private Sub lst���_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 And lst���.Selected(Item) Then
        For i = 1 To lst���.ListCount - 1
            lst���.Selected(i) = False
        Next
    ElseIf Item > 0 And lst���.Selected(Item) Then
        lst���.Selected(0) = False
    End If
End Sub

Private Sub lst���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lst���_Validate(False)
        Call Form_KeyDown(vbKeyEscape, 0)
    End If
End Sub

Private Sub lst���_LostFocus()
    lst���.Visible = False
End Sub

Private Sub lst���_Validate(Cancel As Boolean)
    Dim objGrid As Object, i As Integer
    
    Set objGrid = Bill(bill_���ʱ���)
    
    With objGrid
        .TextMatrix(.Row, .Col) = Get���ѡ��
        If .TextMatrix(.Row, .Col) = "�������" Then
            For i = 3 To 5
                If i <> .Col Then .TextMatrix(.Row, i) = " "
            Next
        End If
    End With
    
End Sub


Private Function Get���ѡ��() As String
'���ܣ��������ѡ���ѡ��������������"���,����..."�Ĵ�
    Dim i As Integer, strTmp As String
    
    If lst���.Selected(0) Then
        Get���ѡ�� = "�������"
    Else
        For i = 1 To lst���.ListCount - 1
            If lst���.Selected(i) Then
                strTmp = strTmp & "," & lst���.List(i)
            End If
        Next
        Get���ѡ�� = Mid(strTmp, 2)
        If Get���ѡ�� = "" Then Get���ѡ�� = " " 'Ϊ���ܻس�������
    End If
End Function

Private Function Get������ƴ�(str��� As String, rs��� As ADODB.Recordset) As String
'���ܣ�������"CDEFG"�����ת��Ϊ����"���,����..."��
    Dim i As Integer, strTmp As String
    
    If str��� = "" Then
        Get������ƴ� = " " 'Ϊ���ܰ��س�������
        Exit Function
    End If
    
    If str��� = "-" Then
        Get������ƴ� = "�������"
        Exit Function
    End If
    
    For i = 1 To Len(str���)
        rs���.Filter = "����='" & Mid(str���, i, 1) & "'"
        If Not rs���.EOF Then strTmp = strTmp & "," & rs���!���
    Next
    Get������ƴ� = Mid(strTmp, 2)
End Function

Private Function Get�����봮(str��� As String) As String
'���ܣ���������"���,����"�Ĵ���������"CDEFG"�Ĵ�
    Dim i As Integer, j As Integer
    Dim arr���() As String, strTmp As String
    
    If Trim(str���) = "" Then Exit Function
    
    If str��� = "�������" Then
        Get�����봮 = "-"
    Else
        arr��� = Split(str���, ",")
        For i = 0 To UBound(arr���)
            For j = 1 To lst���.ListCount - 1
                If lst���.List(j) = arr���(i) Then
                    strTmp = strTmp & Chr(lst���.ItemData(j))
                    Exit For
                End If
            Next
        Next
        Get�����봮 = strTmp
    End If
End Function


Private Sub bill_AfterAddRow(Index As Integer, Row As Long)
    If Index = bill_���ʱ��� Then
        With Bill(Index)
            .TextMatrix(Row, 3) = " "
            .TextMatrix(Row, 4) = " "
            .TextMatrix(Row, 5) = " "
            .TextMatrix(Row, 6) = ""
            .TextMatrix(Row, 7) = ""
        End With
    End If
    
    If Index = bill_�Զ����� Then
        With Bill(Index)
            .TextMatrix(Row, 3) = "0-��������"
            .TextMatrix(Row, 4) = Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd")
        End With
    End If
    
    Bill(Index).Tag = "���޸�"
End Sub

Private Sub Bill_EditKeyPress(Index As Integer, KeyAscii As Integer)
    With Bill(Index)
        If Index = bill_�Զ����� Then
            If .Col = 4 Then
                .TxtCheck = True
                .TextMask = "0123456789-"
            End If
        End If
    End With
End Sub

Private Sub bill_EnterCell(Index As Integer, Row As Long, Col As Long)
    '��ֹ���뱨�����
    With Bill(Index)
        If Index = bill_���ʱ��� And .Col >= 3 Then
            If .Col = 6 Or .Col = 7 Then
                .TxtEnable = True
            Else
                .TxtEnable = False
            End If
        Else
            .TxtEnable = True
        End If
        
        If Index = bill_���ʱ��� And .Col = 4 Then  '������ʽ2
            If InStr(.TextMatrix(.Row, 1), "ÿ�շ���") > 0 Then
                .ColData(4) = 5 'ÿ�շ��ò��ܱ༭������ʽ2
            Else
                .ColData(4) = 1
            End If
        End If
        If Index = bill_���ʱ��� Then
            Select Case .Col
            Case 6, 7
                .ColData(.Col) = 4
            Case Else
            End Select
        End If
    End With
    
End Sub

Private Sub bill_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With Bill(Index)
        If Index = bill_���ʱ��� And .MouseCol >= 3 And .MouseRow > 0 Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub bill_Validate(Index As Integer, Cancel As Boolean)
    Dim lngRow As Long
    
    If Index = bill_���ʱ��� Then
        
        If MouseInRect(cmdCancel.hwnd) Then Exit Sub
        
        '�����ʱ�������
        If Not Check���ʱ��� Then Cancel = True: Exit Sub
        
        '�ռ����ʱ�������
        With mrsWarn
            .Filter = "���ò���='" & tab����.SelectedItem.Caption & "'"
            Do While Not .EOF
                .Delete
                .Update
                .MoveNext
            Loop
            .Filter = 0
        End With
        
        With Bill(bill_���ʱ���)
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                    mrsWarn.AddNew
                    mrsWarn!���ò��� = tab����.SelectedItem.Caption
                    
                    If .RowData(lngRow) <> 0 Then
                        mrsWarn!����ID = .RowData(lngRow)
                        mrsWarn!������ = Split(.TextMatrix(lngRow, 0), "-")(0)
                        mrsWarn!������ = Split(.TextMatrix(lngRow, 0), "-")(1)
                    End If
                    
                    mrsWarn!�������� = CInt(Left(.TextMatrix(lngRow, 1), 1))
                    mrsWarn!����ֵ = CCur(.TextMatrix(lngRow, 2))
                    
                    mrsWarn!������־1 = Get�����봮(.TextMatrix(lngRow, 3))
                    mrsWarn!������־2 = Get�����봮(.TextMatrix(lngRow, 4))
                    mrsWarn!������־3 = Get�����봮(.TextMatrix(lngRow, 5))
                    
                    mrsWarn!�߿����� = Round(Val(.TextMatrix(lngRow, 6)), 2)
                    mrsWarn!�߿��׼ = Round(Val(.TextMatrix(lngRow, 7)), 2)
                    
                    mrsWarn.Update
                End If
            Next
        End With
    End If
End Sub


Private Sub cmdOneCard_Click(Index As Integer)
    
    Select Case Index
        Case 0
            frmOneCard.mbytInFun = 0
            Call frmOneCard.ShowMe(Me)
            Call LoadOneCard
        Case 1
            If lvw(lvw_һ��ͨ).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_һ��ͨ).SelectedItem
                frmOneCard.mbytInFun = 1
                Call frmOneCard.ShowMe(Me, Mid(.Key, 2), .SubItems(1), .SubItems(2), .SubItems(3), IIF(.SubItems(4) = "����:��׼һ��ͨ", 2, IIF(.SubItems(4) = "����:���漰�ۿ�", 1, 0)))
                Call LoadOneCard
            End With
        Case 2
            If lvw(lvw_һ��ͨ).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_һ��ͨ).SelectedItem
                If MsgBox("��ȷʵҪɾ����" & .SubItems(1) & "����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    Call frmOneCard.DelOneCardRec(Val(Mid(.Key, 2)))
                    Call LoadOneCard
                End If
            End With
    End Select
End Sub


Private Sub cmdWarnDel_Click()
    If tab����.SelectedItem.Caption = "��ͨ����" Then
        MsgBox """" & tab����.SelectedItem.Caption & """��������������ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("ȷʵҪɾ��""" & tab����.SelectedItem.Caption & """����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    With mrsWarn
        .Filter = "���ò���='" & tab����.SelectedItem.Caption & "'"
        
        '��¼ɾ�������ò�������
        If InStr(1, mstrDel���ò���, tab����.SelectedItem.Caption) = 0 Then
            mstrDel���ò��� = IIF(mstrDel���ò��� = "", "", mstrDel���ò��� & ";") & tab����.SelectedItem.Caption
        End If
        
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    tab����.Tabs.Remove tab����.SelectedItem.Index
    tab����.Tabs(1).Selected = True
    
    Bill(bill_���ʱ���).Tag = "���޸�"
End Sub

Private Sub cmdWarnNew_Click()
    Dim strName As String, strCopy As String
    Dim strSchemes As String, i As Integer
    Dim rsCopy As ADODB.Recordset
    
    For i = 1 To tab����.Tabs.Count
        strSchemes = strSchemes & "," & tab����.Tabs(i).Caption
    Next
    
    strName = frmWarnEdit.ShowMe(Me, Mid(strSchemes, 2), strCopy)
    If strName = "" Then Exit Sub
    
    '��������
    Set rsCopy = mrsWarn.Clone
    rsCopy.Filter = "���ò���='" & strCopy & "'"
    Do While Not rsCopy.EOF
        mrsWarn.AddNew
        mrsWarn!���ò��� = strName
        mrsWarn!����ID = rsCopy!����ID
        mrsWarn!������ = rsCopy!������
        mrsWarn!������ = rsCopy!������
        mrsWarn!�������� = rsCopy!��������
        mrsWarn!����ֵ = rsCopy!����ֵ
        mrsWarn!������־1 = rsCopy!������־1
        mrsWarn!������־2 = rsCopy!������־2
        mrsWarn!������־3 = rsCopy!������־3
        mrsWarn!�߿����� = rsCopy!�߿�����
        mrsWarn!�߿��׼ = rsCopy!�߿��׼
        mrsWarn.Update
        rsCopy.MoveNext
    Loop
    
    tab����.Tabs.Add , , strName
    tab����.Tabs(tab����.Tabs.Count).Selected = True
    
    Bill(bill_���ʱ���).Tag = "���޸�"
End Sub



Private Function LoadOneCard() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim ObjItem As ListItem
    
    On Error GoTo errH
    
    lvw(lvw_һ��ͨ).ListItems.Clear
    
    strSQL = "Select ���,����,���㷽ʽ,ҽԺ����,���� From һ��ͨĿ¼ Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        Set ObjItem = lvw(lvw_һ��ͨ).ListItems.Add(, "_" & rsTmp!���, rsTmp!���)
        ObjItem.SubItems(1) = NVL(rsTmp!����)
        ObjItem.SubItems(2) = NVL(rsTmp!���㷽ʽ)
        ObjItem.SubItems(3) = NVL(rsTmp!ҽԺ����)
        ObjItem.SubItems(4) = IIF(NVL(rsTmp!����, 0) = 2, "����:��׼һ��ͨ", IIF(NVL(rsTmp!����, 0) = 1, "����:���漰�ۿ�", "ͣ��"))
        rsTmp.MoveNext
    Loop
    
    If Not lvw(lvw_һ��ͨ).SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw_һ��ͨ, lvw(lvw_һ��ͨ).SelectedItem)
    End If
    LoadOneCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsRaiseByDate(ByVal strID As String) As Boolean
    '�жϸ��շ���Ŀ�Ƿ��ǰ��յ���
    '����True-�ǰ�������
    '����False-���ǰ������
    'strID='J' -��λ��Ŀ
    'strID='H' -������Ŀ
    'strID=���� -����ָ������Ŀ
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    If strID = "J" Then
        strSQL = "Select ID" & _
              " From �շѼ�Ŀ " & _
              " Where Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And ִ������ <> Trunc(ִ������, 'dd') And " & _
              " �շ�ϸĿid In " & _
              " (Select ID " & _
              " From �շ���ĿĿ¼ " & _
              " Where ��� = [1] " & _
              " Union All " & _
              " Select ����id From �շѴ�����Ŀ Where ����id In (Select ID From �շ���ĿĿ¼ Where ��� = [1])) "
    ElseIf strID = "H" Then
            strSQL = "Select ID" & _
              " From �շѼ�Ŀ " & _
              " Where Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And ִ������ <> Trunc(ִ������, 'dd') And " & _
              " �շ�ϸĿid In " & _
              " (Select ID " & _
              " From �շ���ĿĿ¼ " & _
              " Where ��� = [1] " & _
              " Union All " & _
              " Select ����id From �շѴ�����Ŀ Where ����id In (Select ID From �շ���ĿĿ¼ Where ��� = [1])) "
    ElseIf Val(strID) <> 0 Then
        strSQL = "Select Id" & _
                " From �շѼ�Ŀ " & _
                " Where Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate " & _
                " And ִ������<>trunc(ִ������,'dd') And (�շ�ϸĿid = [2] or �շ�ϸĿid in (Select ����id From �շѴ�����Ŀ Where ����id = [2])) "
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID, Val(strID))
    
    IsRaiseByDate = Not (rs.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOperate_Click(Index As Integer)
    Dim str���� As String, str��ԱID As String, str���� As String
    Dim lng���� As Long, lng���� As Long, bln�޸����� As Boolean
    Dim dbl������� As Double
    Dim lst As ListItem
    
    
    Select Case Index
        Case 0 '����
            If frmBillPrivilege.�༭Ȩ��(str����, str��ԱID, str����, lng����, lng����, bln�޸�����, dbl�������, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_����).ListItems
                If lst.Tag = str��ԱID And lst.ListSubItems(1).Tag = lng���� Then
                    MsgBox "���������Ĳ��������Ѿ����ڡ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        Case 1 '�޸�
            If lvw(lvw_����).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_����).SelectedItem
                str���� = .Text
                str���� = .SubItems(1)
                lng���� = Val(.SubItems(2))
                bln�޸����� = (.SubItems(3) = "��")
                dbl������� = Val(.SubItems(4))
                str��ԱID = .Tag
                lng���� = .ListSubItems(1).Tag
            End With
            If frmBillPrivilege.�༭Ȩ��(str����, str��ԱID, str����, lng����, lng����, bln�޸�����, dbl�������, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_����).ListItems
                If Not lst Is lvw(lvw_����).SelectedItem Then
                    If lst.Tag = str��ԱID And lst.ListSubItems(1).Tag = lng���� Then
                        MsgBox "���θı�Ĳ��������Ѿ����ڡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            Next
            
        Case 2 'ɾ��
            If lvw(lvw_����).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_����).SelectedItem
                If MsgBox("��ȷʵҪɾ����" & .Text & "���ԡ�" & .SubItems(1) & "���Ĳ������ƣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                
                lvw(lvw_����).ListItems.Remove .Index
            End With
        Case 3 '���
            If MsgBox("��ȷʵҪɾ�����еĲ������ƣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            
            lvw(lvw_����).ListItems.Clear
    End Select
    
    lvw(lvw_����).Tag = "���޸�"
    
    If Index = 0 Or Index = 1 Then
        If Index = 0 Then
            Set lst = lvw(lvw_����).ListItems.Add(, , str����)
            lst.Selected = True
            lst.EnsureVisible
        Else
            Set lst = lvw(lvw_����).SelectedItem
            lst.Text = str����
        End If
        lst.SubItems(1) = str����
        lst.SubItems(2) = lng����
        lst.SubItems(3) = IIF(bln�޸����� = True, "��", "��")
        lst.SubItems(4) = IIF(Val(dbl�������) = 0, "", Format(Val(dbl�������), "0.00"))
        lst.Tag = str��ԱID
        lst.ListSubItems(1).Tag = lng����
    End If
    
End Sub

Private Sub lvw_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     If Index = lvw_���� Then
        If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
            lvw(lvw_����).SortOrder = IIF(lvw(lvw_����).SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            mintColumn = ColumnHeader.Index - 1
            lvw(lvw_����).SortKey = mintColumn
            lvw(lvw_����).SortOrder = lvwAscending
        End If
     End If
     
End Sub

Private Sub lvw_DblClick(Index As Integer)
    If Index = lvw_���� Then
        Call cmdOperate_Click(1)
    End If
End Sub

Private Sub lvw_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If Me.Visible Then Call SetParChange(lvw, Index, mrsPar)
    
    Dim itemTemp As MSComctlLib.ListItem
    For Each itemTemp In lvw(Index).ListItems
        If Not itemTemp Is Item Then
            itemTemp.Checked = False
        End If
    Next
End Sub

Private Sub lvw_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim lngԭֵ As Long, blnValue As Boolean, strValue As String
    If Index = lvw_Ʊ�� Then
        lngԭֵ = Val(Item.SubItems(1))
        ud(ud_���볤��).Max = 20
        '�������ֵʱ�������Ѿ��������б��е�ֵ
        ud(ud_���볤��).value = IIF(lngԭֵ = 0, 7, lngԭֵ)
        chk(chk_Ʊ�ſ���).value = IIF(Item.SubItems(2) = "��", 1, 0)
        Exit Sub
    ElseIf Index = lvw_һ��ͨ Then
        cmdOneCard(1).Enabled = Item.Text <> ""
        cmdOneCard(2).Enabled = cmdOneCard(1).Enabled
    End If
    
    If Not Me.Visible Then Exit Sub
    Call SetParChange(lvw, Index, mrsPar, blnValue, strValue)
    
    
End Sub

Private Sub lvw_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = lvw_Ʊ�� Then
        If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    ElseIf Index = lvw_���� Then
        If KeyAscii = vbKeyReturn Then Call cmdOperate_Click(1)
    End If
End Sub

Private Sub Load���ݲ���()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, str���� As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select A.��ԱID,B.����,A.����,A.ʱ������,A.���˵���,A.������� from ���ݲ������� A,��Ա�� B where A.��ԱID=B.ID"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    lvw(lvw_����).ListItems.Clear
    Do Until rsTemp.EOF
        Set lst = lvw(lvw_����).ListItems.Add(, , rsTemp("����"))
        
        str���� = Switch(rsTemp("����") = 1, "�Һŵ���", rsTemp("����") = 2, "�շѵ�", rsTemp("����") = 3, "���۵�", rsTemp("����") = 4, "�������", _
                       rsTemp("����") = 5, "סԺ����", rsTemp("����") = 6, "Ԥ����", rsTemp("����") = 7, "���ʵ���", rsTemp("����") = 8, "���￨", rsTemp("����") = 9, "����")
        lst.SubItems(1) = str����
        lst.SubItems(2) = rsTemp("ʱ������")
        lst.SubItems(3) = IIF(rsTemp("���˵���") = 1, "��", "��")
        lst.SubItems(4) = IIF(IsNull(rsTemp("�������")), "", Format(rsTemp("�������"), "0.00"))
        lst.Tag = rsTemp("��ԱID")
        lst.ListSubItems(1).Tag = rsTemp("����")
        
        rsTemp.MoveNext
    Loop
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Load����()
    Dim rs���� As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    gstrSQL = "select A.ID,A.����,A.���� " & _
               " from  ��������˵�� b,���ű� a " & _
               " where B.������� in(1,2,3) And B.��������='����' and  b.����ID=a.ID and " & _
               Where����ʱ��("A") & " order by ����"
    Call zlDatabase.OpenRecordset(rs����, gstrSQL, Me.Caption)
    
    Bill(bill_�Զ�����).Clear
    Bill(bill_���ʱ���).Clear
    
    If rs����.RecordCount > 0 Then
        mshAutoCalc.Rows = rs����.RecordCount + 1
        lngRow = 1
        Do Until rs����.EOF
            Bill(bill_�Զ�����).AddItem rs����("����") & "-" & rs����("����")
            Bill(bill_�Զ�����).ItemData(Bill(bill_�Զ�����).NewIndex) = rs����("ID")
            Bill(bill_���ʱ���).AddItem rs����("����") & "-" & rs����("����")
            Bill(bill_���ʱ���).ItemData(Bill(bill_�Զ�����).NewIndex) = rs����("ID")
            mshAutoCalc.TextMatrix(lngRow, 0) = rs����("����") & "-" & rs����("����")
            mshAutoCalc.RowData(lngRow) = rs����("ID")
            lngRow = lngRow + 1
            rs����.MoveNext
        Loop
        Bill(bill_�Զ�����).ListIndex = 0
    End If
    Bill(bill_���ʱ���).AddItem "*����*"
    Bill(bill_���ʱ���).ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load��������()
'���ܣ���ʼ����������
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle

    gstrSQL = "Select ����,���� From �������� Order by ����"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    lst(lst_ҽ������).Clear
    lst(lst_���Ѳ���).Clear
    Do Until rsTemp.EOF
        lst(lst_ҽ������).AddItem rsTemp("����") & "." & rsTemp("����")
        lst(lst_���Ѳ���).AddItem rsTemp("����") & "." & rsTemp("����")
        
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDepartments(ByVal str���� As String, _
    ByVal str������� As String, _
    Optional ByVal bln������Ա���� As Boolean = False, _
    Optional ByVal blnCheckվ�� As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ʵĲ����б�
    '���:str����='�ٴ�','����','��ҩ��',...,����Ϊ��
    '     str�������:��,����:��1,3
    '     bln������Ա����-����Ա����������
    '����:
    '����:
    '����:���˺�
    '����:2009-10-12 09:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    
    str���� = Replace(str����, "'", "")
    If str���� <> "" Then
        If InStr(1, str����, ",") > 0 Then
            strSQL = " And Instr(','||[1]||',',','||B.��������||',')>0"
        Else
            strSQL = " And B.�������� = [1]"
        End If
    End If
    If bln������Ա���� Then strSQL = strSQL & "  And A.id=C.����ID and C.��Աid =[3]"
    
    strSQL = _
        " Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.������� " & _
        " From ���ű� A,��������˵�� B " & IIF(bln������Ա����, ",������Ա C", "") & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And Instr(',' || [2]|| ',',',' || B.������� || ',')>0 " & strSQL & _
         IIF(blnCheckվ��, " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
        " Order by A.����"
    Set GetDepartments = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str����, str�������, glngUserId)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadOther()
'�������ĳ�ʼ������
    Dim rsTemp As New ADODB.Recordset
    Dim lngMaxRow As Long, lngRow As Long, lng��λ As Long
    Dim lngTmp As Long, i As Long
    Dim strobjTemp As String, strWorkTemp As String
    Dim blnHave As Boolean, strCoding As String
    
    
    '�շ��ض���Ŀ
    On Error GoTo ErrHandle
    gstrSQL = "select a.�ض���Ŀ ,c.ID,c.����  " & _
            " from �շ��ض���Ŀ a,�շ�ϸĿ c " & _
            " where a.�շ�ϸĿID =c.id"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("�ض���Ŀ")
            Case "������"
                txtCmd(0).Tag = rsTemp("ID")
                txtCmd(0).Text = rsTemp("����")
            Case "������"
                txtCmd(1).Tag = rsTemp("ID")
                txtCmd(1).Text = rsTemp("����")
            Case "��ͨ���÷�"
                txtCmd(3).Tag = rsTemp("ID")
                txtCmd(3).Text = rsTemp("����")
            Case "�������÷�"
                txtCmd(4).Tag = rsTemp("ID")
                txtCmd(4).Text = rsTemp("����")
        End Select
        rsTemp.MoveNext
    Loop
    
    '�����Զ����ʳ���
    gstrSQL = "select A.����ID,B.����,b.���� as ���� ,a.�շ�ϸĿID,c.���� as �շ�ϸĿ ,a.�����־,a.�������� " & _
            " from �Զ��Ƽ���Ŀ A,���ű� B,�շ�ϸĿ C " & _
            " where A.����ID= B.id and A.�շ�ϸĿID =C.id(+) " & _
            " order by b.���� "
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With Bill(bill_�Զ�����)
        lngRow = 1
        Do Until rsTemp.EOF
            If IsNull(rsTemp("�շ�ϸĿID")) Then
                '��λ�ѻ����
                For lngTmp = 1 To mshAutoCalc.Rows - 1
                    If mshAutoCalc.RowData(lngTmp) = rsTemp("����ID") Then
                        If rsTemp("�����־") = 1 Then
                            '��λ��
                            mshAutoCalc.TextMatrix(lngTmp, 1) = "��"
                            mshAutoCalc.TextMatrix(lngTmp, 2) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                            mshAutoCalc.TextMatrix(lngTmp, 5) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                        Else
                            '�����
                            mshAutoCalc.TextMatrix(lngTmp, 3) = "��"
                            mshAutoCalc.TextMatrix(lngTmp, 4) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                            mshAutoCalc.TextMatrix(lngTmp, 6) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                        End If
                    End If
                Next
            Else
                '��������
                .Rows = lngRow + 1
                .RowData(lngRow) = rsTemp("����ID")
                .TextMatrix(lngRow, 0) = rsTemp("����") & "-" & rsTemp("����")
                .TextMatrix(lngRow, 1) = rsTemp("�շ�ϸĿID")
                .TextMatrix(lngRow, 2) = rsTemp("�շ�ϸĿ")
                .TextMatrix(lngRow, 3) = Switch(rsTemp("�����־") = 6, "1-������", rsTemp("�����־") = 8, "2-����һ��", True, "0-��������")
                .TextMatrix(lngRow, 4) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                lngRow = lngRow + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    
    '���ʱ������
    gstrSQL = "Select ����,��� From �շ���� Order by ����"
    Set mrs��� = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrs���, gstrSQL, Me.Caption)
    
    lst���.Clear
    lst���.AddItem "�������"
    Do While Not mrs���.EOF
        lst���.AddItem mrs���!���
        lst���.ItemData(lst���.NewIndex) = Asc(mrs���!����)
        mrs���.MoveNext
    Loop
    
    '�������ʱ�����
    Set mrsWarn = New ADODB.Recordset
    mrsWarn.Fields.Append "����ID", adBigInt, , adFldIsNullable
    mrsWarn.Fields.Append "������", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "������", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "���ò���", adVarChar, 100
    mrsWarn.Fields.Append "��������", adSmallInt
    mrsWarn.Fields.Append "����ֵ", adCurrency
    mrsWarn.Fields.Append "������־1", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "������־2", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "������־3", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "�߿�����", adCurrency
    mrsWarn.Fields.Append "�߿��׼", adCurrency
    
    mrsWarn.CursorLocation = adUseClient
    mrsWarn.LockType = adLockOptimistic
    mrsWarn.CursorType = adOpenStatic
    mrsWarn.Open
    
    gstrSQL = "" & _
    "   Select a.����ID,B.����,b.���� as ����,a.���ò���,nvl(a.��������,1) as ��������, " & _
    "               a.����ֵ,a.������־1,a.������־2,a.������־3,A.�߿�����,a.�߿��׼ " & _
    "   From ���ʱ����� a,���ű� b " & _
    "   Where a.����ID= b.id(+)  " & _
    "   Order by Decode(a.���ò���,'��ͨ����',1,'ҽ������',2,3),a.���ò���,B.���� Desc"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    strCoding = ",��ͨ����" '������һ����ͨ����
    Do Until rsTemp.EOF
        mrsWarn.AddNew
        mrsWarn!����ID = rsTemp!����ID
        mrsWarn!������ = rsTemp!����
        mrsWarn!������ = rsTemp!����
        mrsWarn!���ò��� = rsTemp!���ò���
        mrsWarn!�������� = rsTemp!��������
        mrsWarn!����ֵ = rsTemp!����ֵ
        mrsWarn!������־1 = rsTemp!������־1
        mrsWarn!������־2 = rsTemp!������־2
        mrsWarn!������־3 = rsTemp!������־3
        mrsWarn!�߿����� = Val(NVL(rsTemp!�߿�����))
        mrsWarn!�߿��׼ = Val(NVL(rsTemp!�߿��׼))
        mrsWarn.Update
        
        If InStr(strCoding & ",", "," & rsTemp!���ò��� & ",") = 0 Then
            strCoding = strCoding & "," & rsTemp!���ò���
        End If
        rsTemp.MoveNext
    Loop
    strCoding = Mid(strCoding, 2)
    tab����.Tabs.Clear
    For i = 0 To UBound(Split(strCoding, ","))
        tab����.Tabs.Add , , Split(strCoding, ",")(i)
    Next
    tab����.Tabs(1).Selected = True '֮ǰ���ἤ��Click�¼�,��Ϊ����
   
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle

    '�����Ѻ͹������޶�Ϊ������Ŀ
    strSQL = "select id,����,����,���㵥λ,˵�� from �շ���ĿĿ¼ where ���='Z' and nvl(�Ƿ���,0)=0"

    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        If IsNumeric(txtCmd(Index).Tag) = False Then txtCmd(Index).Tag = 0
        strSQL = frmSelCurr.ShowCurrSel(Me, rsTmp, "id,0,0,2;���,1000,0,2;����,1800,0,1;��λ,800,0,2;˵��,2300,0,2", -1, "������Ŀѡ��", , CStr(txtCmd(Index).Tag), 0, 3)
        If strSQL <> "" Then
            txtCmd(Index).Tag = CLng(Split(strSQL, ";")(0))
            txtCmd(Index).Text = Trim(Split(strSQL, ";")(2))
            txtCmd(Index).SetFocus
            
            fra�ض��շ���Ŀ.Tag = "���޸�"
        End If
    Else
        MsgBox "���κ���Ŀ���ã�", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtLocate_Change(Index As Integer)
    If Index = txt_Dept Then
        mlngPreFind = 1
    ElseIf Index = txt_Par Then
        txtLocate(Index).Tag = ""
    End If
End Sub

Private Sub txtLocate_GotFocus(Index As Integer)
    txtLocate(Index).SelStart = 0
    txtLocate(Index).SelLength = Len(txtLocate(Index).Text)
End Sub

Private Sub txtLocate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim strFind As String
        
        If Trim(txtLocate(Index).Text) = "" Then Exit Sub
        strFind = UCase(Trim(txtLocate(Index).Text))
        
        Select Case Index
        Case txt_Par
            Call LocatePar(txtLocate(Index), Me)
        Case txt_Dept
            If mshAutoCalc.Visible Or Bill(bill_�Զ�����).Visible Then
                If lblLocate(txt_Dept).Tag = "mshAutoCalc" Or lblLocate(txt_Dept).Tag = "" Then
                    Call LocateDept(strFind, mshAutoCalc)
                Else
                    Call LocateDept(strFind, Bill(bill_�Զ�����))
                End If
            ElseIf Bill(bill_���ʱ���).Visible Then
                Call LocateDept(strFind, Bill(bill_���ʱ���))
                
            End If
        End Select
    End If
End Sub

Private Sub LocateDept(ByVal strFind As String, ByRef objBill As Object)
'���ܣ����ҿ���
    Dim i As Long
    Dim strCode As String, strName As String
    
    With objBill
        For i = mlngPreFind To .Rows - 1
            '0��Ϊ������
            strCode = Split(.TextMatrix(i, 0), "-")(0)
            strName = Split(.TextMatrix(i, 0), "-")(1)
            
            If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                objBill.SetFocus
                .Row = i: .Col = 1
                .TopRow = i
                Exit For
            End If
        Next
        If i < .Rows Then
            mlngPreFind = i + 1
        Else
            If mlngPreFind = 1 Then
                MsgBox "û���ҵ�ƥ��Ŀ��ң�������������ݡ�", vbInformation, Me.Caption
                txtLocate(txt_Dept).SetFocus
            Else
                MsgBox "ȫ�������ˣ�����û���ˡ�", vbInformation, Me.Caption
                mlngPreFind = 1
            End If
        End If
    End With
End Sub

Private Sub ud_Change(Index As Integer)
    Dim strValue As String
    If Not Me.Visible Then Exit Sub
    '��̬�ı�Ʊ�ų���
    If Index = ud_���볤�� Then
        lvw(lvw_Ʊ��).SelectedItem.SubItems(1) = ud(ud_���볤��).value
        strValue = GetBillLenSet
        Call SetParChange(lvw, lvw_Ʊ��, mrsPar, True, strValue)
    End If
End Sub

Private Sub txtUD_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtUD(Index).Text) > ud(Index).Max Or Val(txtUD(Index).Text) < ud(Index).Min Then
        txtUD(Index).Text = ud(Index).value
    End If
End Sub

Private Sub txtUD_Change(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    
    If Not Me.Visible Then Exit Sub
    
    Select Case Index
    Case ud_�Һŵ�, ud_����Һŵ�
        blnValue = True
        strValue = txtUD(ud_�Һŵ�).Text & txtUD(ud_����Һŵ�).Text
    Case ud_���볤��
        strValue = GetBillLenSet
        Call SetParChange(lvw, lvw_Ʊ��, mrsPar, True, strValue)
        Exit Sub
    Case ud_Ʊ������
        strValue = IIF(chk(chk_Ʊ��ʣ��N�����Ѳ���Ա).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(Index).Text)
        Call SetParChange(chk, chk_Ʊ��ʣ��N�����Ѳ���Ա, mrsPar, True, strValue)
        txtUD(Index).ForeColor = chk(chk_Ʊ��ʣ��N�����Ѳ���Ա).ForeColor
        Exit Sub
    Case ud_�շ�_Ʊ������
        strValue = IIF(chk(chk_�շ�_Ʊ��ʣ��X�ſ�ʼ����).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(Index).Text)
        Call SetParChange(chk, chk_�շ�_Ʊ��ʣ��X�ſ�ʼ����, mrsPar, True, strValue)
        txtUD(Index).ForeColor = chk(chk_�շ�_Ʊ��ʣ��X�ſ�ʼ����).ForeColor
        Exit Sub
    Case ud_������_Ʊ������
        strValue = IIF(chk(chk_������_Ʊ��ʣ��X�ſ�ʼ����).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(Index).Text)
        Call SetParChange(chk, chk_������_Ʊ��ʣ��X�ſ�ʼ����, mrsPar, True, strValue)
        txtUD(Index).ForeColor = chk(chk_������_Ʊ��ʣ��X�ſ�ʼ����).ForeColor
        Exit Sub
    Case ud_������_Ʊ������
        strValue = Val(txtUD(Index).Text): blnValue = True
    End Select
    
    Call SetParChange(txtUD, Index, mrsPar, blnValue, strValue)
 
End Sub

Private Sub txtUD_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtUD(Index))
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Select Case Index
        Case txt_�Һ�_ԤԼʧЧ����
            With tbPage(Pg_�Һ�ҵ��)
                .Item(3).Selected = True
            End With
        Case Else
            Call zlCommFun.PressKey(vbKeyTab)
        End Select
    ElseIf KeyAscii = Asc(gstrParSplit1) Or KeyAscii = Asc(gstrParSplit2) Then
        KeyAscii = 0
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Dim strValue As String
    If Val(txt(txt_���������).Text) = 0 Then txt(txt_���������).Text = ""
    If Index = txt_����֧�� Then
        If Val(txt(Index).Text) = 0 Then Exit Sub
        strValue = -1 * Val(txt(Index).Text)
        strValue = strValue & "|" & IIF(optBrushCard(11).value, 1, IIF(optBrushCard(12).value, 2, 0))
        Call SetParChange(optBrushCard, 3, mrsPar, True, strValue)
    End If
End Sub

Private Sub txtDateInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With txtDateInput
            If Not IsDate(.Text) Then
                If Not IsDate(Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)) Then
                    MsgBox "��������ȷ�����ڸ�ʽ(yyyy-mm-dd����yyyymmdd)��", vbInformation, gstrSysName
                    Exit Sub
                Else
                    .Text = Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)
                End If
            End If
            mshAutoCalc.TextMatrix(mintCurRow, mintCurCol) = .Text
            mshAutoCalc.Tag = "���޸�"
            .Visible = False
        End With
    End If
End Sub

Private Sub txtDateInput_LostFocus()
    txtDateInput.Text = ""
    txtDateInput.Visible = False
End Sub

Private Sub mshAutoCalc_Click()
    With Me.mshAutoCalc
        If .Row > 0 And (.Col = 2 Or .Col = 4) And .TextMatrix(.Row, IIF(.Col = 2, 1, 3)) <> "" Then
            mintCurRow = .Row
            mintCurCol = .Col
            txtDateInput.Move (.Left + .CellLeft - 10), (.Top + .CellTop - 10), .CellWidth, .CellHeight
            If .TextMatrix(.Row, .Col) <> "" Then
                txtDateInput.Text = .TextMatrix(.Row, .Col)
            End If
            txtDateInput.Visible = True
            txtDateInput.SetFocus
        End If
    End With
End Sub

Private Sub mshAutoCalc_DblClick()
    With mshAutoCalc
        If .MouseRow > 0 And .MouseCol > 0 And .RowData(.MouseRow) <> 0 Then
            If .Col = 1 Or .Col = 3 Then
                If .Col = 1 And Not mblnJRaiseByDate Then
                    MsgBox "��λ����Ŀ���������Ŀ�ļ۸�������ǰ�����ִ�еģ����顣", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                If .Col = 3 And Not mblnHRaiseByDate Then
                    MsgBox "��������Ŀ���������Ŀ�ļ۸�������ǰ�����ִ�еģ����顣", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                .Text = IIF(.Text = "", "��", "")
                .TextMatrix(.Row, IIF(.Col = 1, 2, 4)) = IIF(.Text = "", "", IIF(.TextMatrix(.Row, IIF(.Col = 1, 5, 6)) = "", Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd"), .TextMatrix(.Row, IIF(.Col = 1, 5, 6))))
                
                .Tag = "���޸�"
            End If
        End If
    End With
End Sub

Private Sub mshAutoCalc_KeyPress(KeyAscii As Integer)
    With mshAutoCalc
        If KeyAscii = vbKeyReturn Then
            If .Col = 1 Then
                .Col = 2
            ElseIf .Col = 4 Then
                If .Row = .Rows - 1 Then
                    Bill(bill_�Զ�����).SetFocus
                Else
                    .Row = .Row + 1
                    .Col = 1
                    If .Row - .TopRow > 8 Then .TopRow = .Row - 8
                End If
            End If
        ElseIf KeyAscii = Asc(" ") Then
            If .Row > 0 And (.Col = 1 Or .Col = 3) And .RowData(.Row) <> 0 Then
                If .Col = 1 And Not mblnJRaiseByDate Then
                    MsgBox "��λ����Ŀ���������Ŀ�ļ۸�������ǰ�����ִ�еģ����顣", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                If .Col = 3 And Not mblnHRaiseByDate Then
                    MsgBox "��������Ŀ���������Ŀ�ļ۸�������ǰ�����ִ�еģ����顣", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                .Text = IIF(.Text = "", "��", "")
                .TextMatrix(.Row, IIF(.Col = 1, 2, 4)) = IIF(.Text = "", "", IIF(.TextMatrix(.Row, IIF(.Col = 1, 5, 6)) = "", Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd"), .TextMatrix(.Row, IIF(.Col = 1, 5, 6))))
                
            End If
        Else
            If .Row > 0 And (.Col = 2 Or .Col = 4) And .TextMatrix(.Row, IIF(.Col = 2, 1, 3)) <> "" Then
                mintCurRow = .Row
                mintCurCol = .Col
                txtDateInput.Move (.Left + .CellLeft - 10), (.Top + .CellTop - 10), .CellWidth, .CellHeight
                If .TextMatrix(.Row, .Col) <> "" Then
                    txtDateInput.Text = .TextMatrix(.Row, .Col)
                End If
                txtDateInput.Visible = True
                txtDateInput.SetFocus
            End If
        End If
        .Tag = "���޸�"
    End With
End Sub

Private Sub txtCmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txtCmd(Index).Tag = ""
        txtCmd(Index).Text = ""
        
        fra�ض��շ���Ŀ.Tag = "���޸�"
    End If
End Sub

Private Sub txtCmd_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    ElseIf KeyAscii = Asc("*") Then
        Call cmdSelect_Click(Index)
    End If
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case cbo_���ʲ���_���ʺ�ҩ, cbo_һ��ͨ_����Ʊ�ݸ�ʽ
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case chk_����ԤԼ�Һŵ���ֹɾ������
        With tbPage(Pg_�Һ�ҵ��)
            .Item(1).Selected = True
        End With
    Case chk_�Һ�_���ϸ����Ϊ����
        With tbPage(Pg_�Һ�ҵ��)
            .Item(2).Selected = True
        End With
    Case chk_���_�����ӡ
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
        
End Sub


Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    Dim blnValue As Boolean, strValue As String
    Dim i As Long
    If Not Me.Visible Then Exit Sub
    Select Case Index
    Case lst_ҽ������, lst_���Ѳ���
        blnValue = True
        strValue = Replace(Replace(GetTextFromList(lst(Index)), "'", ""), ",", "|")
    Case lst_������_���㷽ʽ
        blnValue = True
        strValue = Replace(Replace(GetTextFromList(lst(Index)), "'", ""), ",", "|")
    Case lst_ˢ������
        blnValue = True
        With lst(lst_ˢ������)
            For i = 0 To .ListCount - 1
                strValue = strValue & IIF(.Selected(i), 1, 0)
            Next
        End With
    Case lst_����_�Էѷ������
        strValue = ""
        With lst(lst_����_�Էѷ������)
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    strValue = strValue & "," & Chr(.ItemData(i))
                End If
            Next
        End With
        If strValue <> "" Then strValue = Mid(strValue, 2)
        blnValue = True
    Case lst_�ѻ�ҽ�����㷽ʽ
        strValue = ""
        With lst(lst_�ѻ�ҽ�����㷽ʽ)
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    strValue = strValue & "|" & zlCommFun.GetNeedName(.List(i), "-")
                End If
            Next
        End With
        If strValue <> "" Then strValue = Mid(strValue, 2)
        blnValue = True
    End Select
    Call SetParChange(lst, Index, mrsPar, blnValue, strValue)
End Sub

Private Sub Load�ѻ�ҽ�����㷽ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Էѷ������
    '����:���˺�
    '����:2015-06-24 15:34:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrHandle
    
    strSQL = "Select ����,���� From ���㷽ʽ Where ���� = 2 And Nvl(Ӧ�տ�,0)=0 And Nvl(Ӧ����,0)=0 Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With lst(lst_�ѻ�ҽ�����㷽ʽ)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!���� & "-" & rsTemp!����
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case lst_������_���㷽ʽ
         If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    Case lst_ˢ������
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub chk_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    Dim intBillRull As Boolean
    If mblnNotChange Then Exit Sub
    
    If Not Me.Visible Then Exit Sub

    Select Case Index
    Case chk_���������շ����
        If chk(Index).value = 1 Then
            If chk(chk_�շ���Ŀ��λ��������).value = 1 Then chk(chk_�շ���Ŀ��λ��������).value = 0
        End If
    Case chk_�շ���Ŀ��λ��������
        If chk(Index).value = 1 Then
            If chk(chk_���������շ����).value = 1 Then chk(chk_���������շ����).value = 0
        End If
    Case chk_Ʊ�ſ���
        lvw(lvw_Ʊ��).SelectedItem.SubItems(2) = IIF(chk(Index).value = 1, "��", "")
        
        strValue = GetBillCtlSet: blnValue = True
        Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
        
    Case chk_��������, chk_ˢ���￨, chk_�Һŵ���, chk_����ID
        strValue = chk(chk_��������).value & chk(chk_ˢ���￨).value & chk(chk_�Һŵ���).value & chk(chk_����ID).value
        blnValue = True
        Call SetParChange(chk, chk_��������, mrsPar, blnValue, strValue)
        Call SetParChange(chk, chk_ˢ���￨, mrsPar, blnValue, strValue)
        Call SetParChange(chk, chk_�Һŵ���, mrsPar, blnValue, strValue)
        Call SetParChange(chk, chk_����ID, mrsPar, blnValue, strValue)
    Case chk_����ģ������
        If chk(chk_����ģ������).value = 1 Then
            txt(txt_������������).Enabled = True
        Else
            txt(txt_������������).Enabled = False
        End If
    Case chk_��������ģ������
        txt(txt_����ģ����������).Enabled = chk(Index).value = 1
    
    Case chk_Ʊ��ʣ��N�����Ѳ���Ա
        strValue = IIF(chk(Index).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(ud_Ʊ������).Text)
        txtUD(ud_Ʊ������).Enabled = chk(Index).value = 1
        ud(ud_Ʊ������).Enabled = chk(Index).value = 1
        
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        txtUD(ud_Ʊ������).ForeColor = chk(Index).ForeColor
        Exit Sub
    Case chk_�Һ�_�Զ�ˢ�¹ҺŰ���
        txt(txt_�Һ�_ˢ��ʱ��).Enabled = chk(Index).value = 1
        If chk(Index).value <> 1 Then txt(txt_�Һ�_ˢ��ʱ��).Text = 0
        strValue = Val(txt(txt_�Һ�_ˢ��ʱ��).Text)
        Call SetParChange(txt, txt_�Һ�_ˢ��ʱ��, mrsPar, True, strValue)
        chk(Index).ForeColor = txt(txt_�Һ�_ˢ��ʱ��).ForeColor
        Exit Sub
    Case chk_�Һ�_����ģ������
        txt(txt_�Һ�_������������).Enabled = chk(Index).value = 1
    Case chk_�Һ�_���˹Һſ�������
        txt(txt_�Һ�_���˹Һſ�������).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_�Һ�_���˹Һſ�������).Text
        Else
            strValue = "0"
        End If
    Case chk_�Һ�_����ͬ����ԼN����
        txt(txt_�Һ�_����ͬ����ԼN����).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_�Һ�_����ͬ����ԼN����).Text
        Else
            strValue = "0"
        End If
    Case chk_ר�ҺŹҺ�����
        txt(txt_ר�ҺŹҺ�����).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_ר�ҺŹҺ�����).Text
        Else
            txt(txt_ר�ҺŹҺ�����).Text = ""
            strValue = "0"
        End If
    Case chk_ר�Һ�ԤԼ����
        txt(txt_ר�Һ�ԤԼ����).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_ר�Һ�ԤԼ����).Text
        Else
            txt(txt_ר�Һ�ԤԼ����).Text = ""
            strValue = "0"
        End If
    Case chk_�Һ�_����ԤԼ������
        txt(txt_�Һ�_����ԤԼ������).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_�Һ�_����ԤԼ������).Text
        Else
            strValue = "0"
        End If
    Case chk_�Һ�_����ͬ���޹�N����
        txt(txt_�Һ�_����ͬ���޹�N����).Enabled = chk(Index).value = 1
        chk(chk_�Һ�_����ͬ���޹�N����_����).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_�Һ�_����ͬ���޹�N����).Text & "|" & IIF(chk(chk_�Һ�_����ͬ���޹�N����_����).value = 1, "1", "0")
        Else
            strValue = "0|0"
        End If
    Case chk_�Һ�_����ͬ���޹�N����_����
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_�Һ�_����ͬ���޹�N����).Text & "|1"
        Else
            strValue = "0|0"
        End If
    Case chk_�ҺŰ������Ұ���, chk_ԤԼ�������Ұ���
        blnValue = True
        strValue = chk(chk_�ҺŰ������Ұ���).value & "|" & chk(chk_ԤԼ�������Ұ���).value
    Case chk_�շ�_δ�Һ��Զ����չҺŷ�
        blnValue = True
        strValue = ""
        
        If chk(Index).value = 1 Then
            If txt(txt_�շ�_�������չҺŷ�).Text = "" And Me.Visible Then
                Call cmdAddedItem_Click: Exit Sub
            End If
            strValue = cmdAddedItem.Tag & ";" & txt(txt_�շ�_�������չҺŷ�).Text
        Else
            mblnNotChange = True
            txt(txt_�շ�_�������չҺŷ�).Text = "": cmdAddedItem.Tag = ""
            mblnNotChange = False
        End If
        Call SetParChange(txt, txt_�շ�_�������չҺŷ�, mrsPar, blnValue, strValue)
        Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
        Exit Sub
    Case chk_�շ�_�Զ���ϵ���
        blnValue = True
        strValue = "0"
        If chk(Index).value = 1 Then
            strValue = cbo(cbo_�շ�_�Զ���ϵ���).ListIndex + 1
        End If
        cbo(cbo_�շ�_�Զ���ϵ���).Enabled = chk(Index).value = 1
        Call SetParChange(chk, chk_�շ�_�Զ���ϵ���, mrsPar, blnValue, strValue)
        Call SetParChange(cbo, cbo_�շ�_�Զ���ϵ���, mrsPar, blnValue, strValue)
        Exit Sub
    Case chk_�շ�_Ʊ��ʣ��X�ſ�ʼ����
        strValue = IIF(chk(Index).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(ud_�շ�_Ʊ������).Text)
        txtUD(ud_�շ�_Ʊ������).Enabled = chk(Index).value = 1
        ud(ud_�շ�_Ʊ������).Enabled = chk(Index).value = 1
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        txtUD(ud_�շ�_Ʊ������).ForeColor = chk(Index).ForeColor
        Exit Sub
        
    Case chk_�շ�_Ʊ�����ɷ�ʽ
        intBillRull = IIF(cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex < 0, 0, cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex)
        If intBillRull <> 0 Then Exit Sub
        strValue = CStr(IIF(optBillMode(1).value, 1, 0) + Val(chk(chk_�շ�_Ʊ�����ɷ�ʽ).value) * 10)
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        Call SetParChange(optBillMode, 0, mrsPar, True, strValue)
        Exit Sub
    Case chk_�շ�_���ŵ����շѷֱ��ӡ
        chk(chk_�շ�_��첡�˰����ݷֱ��ӡ).Enabled = (chk(Index).value = vbChecked)
    Case chk_����_��������ģ������
        chk(chk_����_ֻ���Һ�Լ��λ����).Enabled = chk(chk_����_��������ģ������).value = 1
        txt(txt_����_��������ģ����������).Enabled = chk(Index).value = 1
    Case chk_����_��������ģ������
        txt(txt_����_��������ģ����������).Enabled = chk(Index).value = 1
    Case chk_�շ�_��������ģ������
        txt(txt_�շ�_��������ģ����������).Enabled = chk(Index).value = 1
    Case chk_�շ�_��Ѱ���۵���
        txt(txt_�շ�_��Ѱ���۵�������).Enabled = chk(Index).value = 1
    Case chk_������_����ģ������
        txt(txt_������_����ģ����������).Enabled = chk(Index).value = 1
        strValue = IIF(chk(chk_������_����ģ������).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txt(txt_������_����ģ����������).Text)
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        Call SetParChange(txt, txt_������_����ģ����������, mrsPar, True, strValue)
        Exit Sub
    Case chk_������_Ʊ��ʣ��X�ſ�ʼ����
        strValue = IIF(chk(Index).value = 1, 1, 0)
        strValue = strValue & "|" & Val(txtUD(ud_������_Ʊ������).Text)
        txtUD(ud_������_Ʊ������).Enabled = chk(Index).value = 1
        ud(ud_������_Ʊ������).Enabled = chk(Index).value = 1
        Call SetParChange(chk, Index, mrsPar, True, strValue)
        txtUD(ud_������_Ʊ������).ForeColor = chk(Index).ForeColor
        Exit Sub
    Case chk_���ʱ����������۷���
        strValue = IIF(chk(Index).value = 1, 1, 0)
        If Val(strValue) = 1 Then
            lblInExseCharge.Enabled = True
            optInExseCharge(0).Enabled = True
            optInExseCharge(1).Enabled = True
        Else
            lblInExseCharge.Enabled = False
            optInExseCharge(0).Enabled = False
            optInExseCharge(1).Enabled = False
        End If
    Case chk_�շ�_�����˲���Ʊ�ݲ��ִ���
        With vsBillFormat(vsGrid_�շ�Ʊ�ݸ�ʽ)
            .ColHidden(.ColIndex("�����˲���Ʊ�ݸ�ʽ")) = chk(Index).value <> 1
        End With
    Case chk_����_����̨ǩ����ʼ�Ŷ�
        Call LoadTriageQueuingDep
        Call SetTriageQueuingEnalbe(chk(Index).value)
    Case chk_�Һ�_����ͬһ��Դ�޹�N����
        txt(txt_�Һ�_����ͬһ��Դ�޹�N����).Enabled = chk(Index).value = 1
        blnValue = True
        If chk(Index).value = 1 Then
            strValue = txt(txt_�Һ�_����ͬһ��Դ�޹�N����).Text
        Else
            strValue = "0"
        End If
    End Select
    Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
End Sub

Public Function IsDrugOrStuff(ByVal strID As String) As Boolean
    '�ж��Ƿ�ΪҩƷ���
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "Select id From �շ�ϸĿ Where ��� In('4','5','6','7') and id=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
    
    IsDrugOrStuff = rs.RecordCount > 0
    rs.Close
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub cmd���ѷ�������_Click(Index As Integer)
    Dim i As Long
    
    With lst(lst_���Ѳ���)
        For i = 0 To .ListCount - 1
            .Selected(i) = Index = 0    '������lst_ItemCheck�¼�
        Next
    End With
End Sub

Private Sub cmdҽ����������_Click(Index As Integer)
    Dim i As Long
    
    With lst(lst_ҽ������)
        For i = 0 To .ListCount - 1
            .Selected(i) = Index = 0    '������lst_ItemCheck�¼�
        Next
    End With
End Sub

Private Sub vsBillFormat_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index = vsGrid_Ԥ��Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("Ԥ����ӡ��ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
               
            Case .ColIndex("Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    If Index = vsGrid_Ԥ����Ʊ��ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("�˿��ӡ��ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
               
            Case .ColIndex("Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    If Index = vsGrid_�շ�Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("�շѴ�ӡ��ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("�շ�Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case .ColIndex("�����˲���Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 2), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_�˷�Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("�շѴ�ӡ��ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("�շ�Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_������Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("�շѴ�ӡ��ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_�������˷�Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("�շѴ�ӡ��ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_����Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("���ʺ��ӡ��ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_���ʺ�Ʊ��ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("���Ϻ��ӡ��ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
            Case .ColIndex("Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
       
    If Index = vsGrid_����Ԥ��Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("Ԥ����ӡ��ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 0), CStr(Col))
               
            Case .ColIndex("Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_ҽ�ƿ��վݸ�ʽ Then
        With vsBillFormat(Index)
            Select Case Col
            Case .ColIndex("Ʊ�ݸ�ʽ")
                Call SetParChange(vsBillFormat, Index, mrsPar, True, GetBillFormat(Index, 1), CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
End Sub

Private Sub vsBillFormat_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
  If Index = vsGrid_Ԥ��Ʊ�ݸ�ʽ Then
    zl_vsGrid_Para_Save pԤ�������, vsBillFormat(Index), Me.Name, "Ԥ����Ʊ��ӡ��ʽ", False, False
    Exit Sub
  End If
  If Index = vsGrid_Ԥ����Ʊ��ʽ Then
    zl_vsGrid_Para_Save pԤ�������, vsBillFormat(Index), Me.Name, "Ԥ���˿��ӡ��ʽ", False, False
    Exit Sub
  End If
  If Index = vsGrid_�շ�Ʊ�ݸ�ʽ Then
    zl_vsGrid_Para_Save p�����շѹ���, vsBillFormat(Index), Me.Name, "�շ�Ʊ�ݸ�ʽ", False, False
    Exit Sub
  End If
  If Index = vsGrid_�˷�Ʊ�ݸ�ʽ Then
    zl_vsGrid_Para_Save p�����շѹ���, vsBillFormat(Index), Me.Name, "�˷�Ʊ�ݸ�ʽ", False, False
    Exit Sub
  End If
  If Index = vsGrid_������Ʊ�ݸ�ʽ Then
    zl_vsGrid_Para_Save p���ﲹ����, vsBillFormat(Index), Me.Name, "������Ʊ�ݸ�ʽ", False, False
    Exit Sub
  End If
  If Index = vsGrid_�������˷�Ʊ�ݸ�ʽ Then
    zl_vsGrid_Para_Save p���ﲹ����, vsBillFormat(Index), Me.Name, "�������˷�Ʊ�ݸ�ʽ", False, False
    Exit Sub
  End If
  
  If Index = vsGrid_����Ʊ�ݸ�ʽ Then
    zl_vsGrid_Para_Save p���˽��ʹ���, vsBillFormat(Index), Me.Name, "����Ʊ�ݸ�ʽ", False, False
    Exit Sub
  End If
  
  If Index = vsGrid_���ʺ�Ʊ��ʽ Then
    zl_vsGrid_Para_Save p���˽��ʹ���, vsBillFormat(Index), Me.Name, "���ʺ�Ʊ��ʽ", False, False
    Exit Sub
  End If
  
  If Index = vsGrid_����Ԥ��Ʊ�ݸ�ʽ Then
    zl_vsGrid_Para_Save pҽ�ƿ�����, vsBillFormat(Index), Me.Name, "����Ԥ����Ʊ��ʽ", False, False
    Exit Sub
  End If
  
  If Index = vsGrid_ҽ�ƿ��վݸ�ʽ Then
    zl_vsGrid_Para_Save pҽ�ƿ�����, vsBillFormat(Index), Me.Name, "ҽ�ƿ��վݸ�ʽ", False, False
    Exit Sub
  End If
  
End Sub

Private Sub vsBillFormat_KeyPress(Index As Integer, KeyAscii As Integer)
    With vsBillFormat(Index)
        Select Case Index
        Case vsGrid_Ԥ��Ʊ�ݸ�ʽ, vsGrid_�շ�Ʊ�ݸ�ʽ, vsGrid_�˷�Ʊ�ݸ�ʽ, vsGrid_������Ʊ�ݸ�ʽ, vsGrid_�������˷�Ʊ�ݸ�ʽ, _
            vsGrid_����Ʊ�ݸ�ʽ, vsGrid_Ԥ����Ʊ��ʽ, vsGrid_���ʺ�Ʊ��ʽ, vsGrid_����Ԥ��Ʊ�ݸ�ʽ, vsGrid_ҽ�ƿ��վݸ�ʽ
           If KeyAscii <> vbKeyReturn Then Exit Sub
            KeyAscii = 0
            If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                If Index = vsGrid_�շ�Ʊ�ݸ�ʽ _
                    Or Index = vsGrid_������Ʊ�ݸ�ʽ _
                    Or Index = vsGrid_�˷�Ʊ�ݸ�ʽ _
                    Or vsGrid_�������˷�Ʊ�ݸ�ʽ Then
                    zlCommFun.PressKey vbKeyTab
                Else
                    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
                End If
               Exit Sub
            End If
            zlVsMoveGridCell vsBillFormat(Index), 1, .Cols - 1
        Case Else
        End Select
    End With
End Sub

Private Sub vsBillFormat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = vsGrid_Ԥ��Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("Ԥ����ӡ��ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    If Index = vsGrid_Ԥ����Ʊ��ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("�˿��ӡ��ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_�շ�Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("�շѴ�ӡ��ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("�շ�Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("�����˲���Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    If Index = vsGrid_�˷�Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("�շѴ�ӡ��ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("�շ�Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    

    If Index = vsGrid_������Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("�շѴ�ӡ��ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_�������˷�Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("�շѴ�ӡ��ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_����Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("���ʺ��ӡ��ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_���ʺ�Ʊ��ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("���Ϻ��ӡ��ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_����Ԥ��Ʊ�ݸ�ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("Ԥ����ӡ��ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    
    If Index = vsGrid_ҽ�ƿ��վݸ�ʽ Then
        With vsBillFormat(Index)
            If .MouseCol = .ColIndex("Ʊ�ݸ�ʽ") Then
                Call SetParTip(vsBillFormat, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
End Sub

Private Sub Load���ʺ�Ʊ��ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ��ʺ�Ʊ��ʽ
    '����:���˺�
    '����:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_���ʺ�Ʊ��ʽ
    
    rsPara.Filter = "ģ��=" & p���˽��ʹ��� & " And ������='���Ϸ�Ʊ��ӡ��ʽ'"
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!����ֵ)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "���Ϸ�Ʊ��ӡ��ʽ", p���˽��ʹ���, "", vsBillFormat(intIndex).ColIndex("���Ϻ��ӡ��ʽ"))
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("���Ϻ��ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    rsTemp.Filter = "ID<>0"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���Ϻ��ӡ��ʽ")) = "0-����ӡƱ��"
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(NVL(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("���Ϻ��ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next

            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p���˽��ʹ���, vsBillFormat(intIndex), Me.Name, "���ʺ�Ʊ��ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load����Ԥ��Ʊ�ݸ�ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط���Ԥ��Ʊ�ݸ�ʽ
    '����:���ϴ�
    '����:2016/9/23 15:35:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_����Ԥ��Ʊ�ݸ�ʽ
    
    rsPara.Filter = "ģ��=" & pҽ�ƿ����� & " And ������='Ԥ����Ʊ��ʽ'"     'Ԥ����Ʊ��ʽ
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "Ԥ����Ʊ��ʽ", pҽ�ƿ�����, "", vsBillFormat(intIndex).ColIndex("Ʊ�ݸ�ʽ"))
    
    rsPara.Filter = "ģ��=" & pҽ�ƿ����� & " And ������='Ԥ����Ʊ��ӡ��ʽ'"    'Ԥ����Ʊ��ӡ��ʽ
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "Ԥ����Ʊ��ӡ��ʽ", pҽ�ƿ�����, "", vsBillFormat(intIndex).ColIndex("Ԥ����ӡ��ʽ"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
 
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
  '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsBillFormat(intIndex)
        .Rows = 3
        .TextMatrix(1, 0) = "����Ԥ��"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "סԺԤ��"
        .Cell(flexcpData, 2, 0) = 2
        
        .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = "0"
        .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = "0"
        
        .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = 1 ' IIF(intType = 5, 0, 1)
        .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = 1 'IIF(intType1 = 5, 0, 1)
        .Editable = flexEDKbdMouse
    End With
    
    
    With vsBillFormat(intIndex)
        .Clear 1: .Rows = 3
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
        Next
    End With
    zl_vsGrid_Para_Restore pҽ�ƿ�����, vsBillFormat(intIndex), Me.Name, "����Ԥ����Ʊ��ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub LoadԤ����Ʊ��ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ����Ʊ��ʽ
    '����:���˺�
    '����:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_Ԥ����Ʊ��ʽ
    
    rsPara.Filter = "ģ��=" & pԤ������� & " And ������='�˿Ʊ��ʽ'"    '�˿Ʊ��ʽ
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�˿Ʊ��ʽ", pԤ�������, "", vsBillFormat(intIndex).ColIndex("Ʊ�ݸ�ʽ"))
    
    rsPara.Filter = "ģ��=" & pԤ������� & " And ������='Ԥ���˿��ӡ��ʽ'"    'Ԥ���˿��ӡ��ʽ
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "Ԥ���˿��ӡ��ʽ", pԤ�������, "", vsBillFormat(intIndex).ColIndex("�˿��ӡ��ʽ"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1103_1"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�˿��ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
 
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsBillFormat(intIndex)
        .TextMatrix(1, 0) = "����Ԥ��"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "סԺԤ��"
        .Cell(flexcpData, 2, 0) = 2
        .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = "0"
        .ColData(.ColIndex("�˿��ӡ��ʽ")) = "0"
        '.ForeColor = &H80000008:  .ForeColorFixed = &H80000008
'        Select Case intType
'        Case 1, 3, 5, 15
             .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = 1 ' IIF(intType = 5, 0, 1)
'        End Select
'        Select Case intType1
'        Case 1, 3, 5, 15
             .ColData(.ColIndex("�˿��ӡ��ʽ")) = 1 'IIF(intType1 = 5, 0, 1)
'        End Select
'        If (Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Or _
'            Val(.ColData(.ColIndex("Ԥ����ӡ��ʽ"))) = 1) Then
'            .Editable = flexEDKbdMouse
'        Else
            .Editable = flexEDKbdMouse
'        End If
    End With
    
    With vsBillFormat(intIndex)
        .Clear 1: .Rows = 3
        For lngRow = 1 To .Cols - 1
            .TextMatrix(lngRow, .ColIndex("�˿��ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("�˿��ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
        Next
    End With
    zl_vsGrid_Para_Restore pԤ�������, vsBillFormat(intIndex), Me.Name, "Ԥ���˿��ӡ��ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsDepositSort_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String, i As Integer, intIndex As Integer
    With vsDepositSort
        Select Case Col
            Case 2
                If Val(.TextMatrix(Row, Col)) = 0 Then
                    .TextMatrix(Row, Col) = -1
                Else
                    .TextMatrix(Row, 3) = 0
                    .TextMatrix(Row, 4) = 0
                End If
            Case 3
                If Val(.TextMatrix(Row, Col)) = 0 Then
                    .TextMatrix(Row, Col) = -1
                Else
                    .TextMatrix(Row, 2) = 0
                    .TextMatrix(Row, 4) = 0
                End If
            Case 4
                If Val(.TextMatrix(Row, Col)) = 0 Then
                    .TextMatrix(Row, Col) = -1
                Else
                    .TextMatrix(Row, 2) = 0
                    .TextMatrix(Row, 3) = 0
                End If
        End Select
        strValue = "1|"
        For i = 2 To 4
            If Abs(Val(.TextMatrix(i, 2))) = 1 Then intIndex = 0
            If Abs(Val(.TextMatrix(i, 3))) = 1 Then intIndex = 1
            If Abs(Val(.TextMatrix(i, 4))) = 1 Then intIndex = 2
            If i <> 4 Then
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex & ","
            Else
                strValue = strValue & .TextMatrix(i, 1) & ":" & intIndex
            End If
        Next i
    End With
    Call SetParChange(optOrder, 1, mrsPar, True, strValue)
End Sub

Private Sub vsDepositSort_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Or Col = 1 Then Cancel = True
End Sub

Private Sub vsfDelFeeDefaultType_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String, i As Integer
    
    With vsfDelFeeDefaultType
        For i = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(i, .ColIndex("ȱʡ����")))) = 1 Then
                strValue = strValue & ";" & .TextMatrix(i, .ColIndex("�տʽ"))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    
    Call SetParChange(vsfDelFeeDefaultType, 0, mrsPar, True, strValue)
End Sub

Private Sub vsfDelFeeDefaultType_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save p�����շѹ���, vsfDelFeeDefaultType, Me.Name, "�˷�ȱʡ��ʽ", False, False
End Sub

Private Sub vsfDelFeeDefaultType_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    KeyAscii = 0
    With vsfDelFeeDefaultType
        If .Row = .Rows - 1 And .Col = .Cols - 1 Then
            zlCommFun.PressKey vbKeyTab
        Else
            zlVsMoveGridCell vsfDelFeeDefaultType, 1, .Cols - 1
        End If
    End With
End Sub

Private Sub vsfDelFeeDefaultType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsfDelFeeDefaultType, 0, mrsPar)
End Sub

Private Sub vsfDrugStore_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore(Index)
        Select Case Col
        Case .ColIndex("ȱʡ"), .ColIndex("����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) = 1
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick(Index As Integer)
    Dim strTmp As String, i As Long
    With vsfDrugStore(Index)
        If Not (.Row > 0 And .Col = 1) Then Exit Sub
        If .Cell(flexcpData, .Row, .ColIndex("ȱʡ")) = 1 Then Exit Sub
'        .TextMatrix(.Row, .Col) = IIF(Val(.TextMatrix(.Row, .Col)) = 0, 1, 0)
        Call SetDrugStockDeFault(.Row, Index)
    End With
End Sub

Private Sub SetDrugStockDeFault(ByVal lngRow As Long, ByVal Index As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩ����ȱʡֵ
    '���:lngRow-ָ����
    '����:���˺�
    '����:2009-09-02 14:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngȱʡ As Long, strType As String
    With vsfDrugStore(Index)
        lngȱʡ = Abs(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ"))))
        If lngȱʡ = 1 Then
            For i = 1 To .Rows - 1
                If i <> lngRow Then
                    .TextMatrix(i, .ColIndex("ȱʡ")) = 0
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfDrugStore_EnterCell(Index As Integer)
    Dim rsTmp As ADODB.Recordset, strList As String
    With vsfDrugStore(Index)
        If .Row > 0 Then
            If .Col = .ColIndex("����") Then
                Set rsTmp = Read��ҩ����(.RowData(.Row))
                strList = "�Զ�����|" & .BuildComboList(rsTmp, "����")
                .ColComboList(.Col) = strList
            Else
                .ColComboList(.Col) = ""
            End If
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Function Read��ҩ����(lngId As Long) As ADODB.Recordset
'���ܣ���ȡָ��ҩ���ķ�ҩ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select ���� From ��ҩ���� Where ҩ��ID=[1] Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngId)
    Set Read��ҩ���� = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsfDrugStore_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        With vsfDrugStore(Index)
            If .MouseCol = .ColIndex("ȱʡ") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("����") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    If Index = 1 Then
        With vsfDrugStore(Index)
            If .MouseCol = .ColIndex("ȱʡ") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("����") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
    If Index = 2 Then
        With vsfDrugStore(Index)
            If .MouseCol = .ColIndex("ȱʡ") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            ElseIf .MouseCol = .ColIndex("����") Then
                Call SetParTip(vsfDrugStore, Index, mrsPar, , , .MouseCol)
            End If
        End With
        Exit Sub
    End If
End Sub
 

Private Sub vsfDrugStore_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer, strWindow As String
    Dim blnHave As Boolean
    If Index = 0 Then
        With vsfDrugStore(Index)
            Select Case Col
            Case .ColIndex("ȱʡ")
                blnHave = False
                For i = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(i, .ColIndex("ȱʡ")))) = 1 Then blnHave = True
                Next i
                If blnHave = True Then
                    If Abs(Val(.TextMatrix(Row, .ColIndex("ȱʡ")))) = 1 Then
                        Call SetDrugStockDeFault(Row, Index)
                        Call SetParChange(vsfDrugStore, Index, mrsPar, True, .RowData(Row), CStr(Col))
                    End If
                Else
                    Call SetParChange(vsfDrugStore, Index, mrsPar, True, "", CStr(Col))
                End If
            Case .ColIndex("����")
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("����")) <> "�Զ�����" And .TextMatrix(i, .ColIndex("����")) <> "" Then
                        strWindow = strWindow & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("����"))
                    End If
                Next i
                If strWindow <> "" Then strWindow = Mid(strWindow, 2)
                Call SetParChange(vsfDrugStore, Index, mrsPar, True, strWindow, CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = 1 Then
        With vsfDrugStore(Index)
            Select Case Col
            Case .ColIndex("ȱʡ")
                If Abs(Val(.TextMatrix(Row, .ColIndex("ȱʡ")))) = 1 Then
                    Call SetDrugStockDeFault(Row, Index)
                    Call SetParChange(vsfDrugStore, Index, mrsPar, True, .RowData(Row), CStr(Col))
                End If
            Case .ColIndex("����")
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("����")) <> "�Զ�����" And .TextMatrix(i, .ColIndex("����")) <> "" Then
                        strWindow = strWindow & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("����"))
                    End If
                Next i
                If strWindow <> "" Then strWindow = Mid(strWindow, 2)
                Call SetParChange(vsfDrugStore, Index, mrsPar, True, strWindow, CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If Index = 2 Then
        With vsfDrugStore(Index)
            Select Case Col
            Case .ColIndex("ȱʡ")
                If Abs(Val(.TextMatrix(Row, .ColIndex("ȱʡ")))) = 1 Then
                    Call SetDrugStockDeFault(Row, Index)
                    Call SetParChange(vsfDrugStore, Index, mrsPar, True, .RowData(Row), CStr(Col))
                End If
            Case .ColIndex("����")
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("����")) <> "�Զ�����" And .TextMatrix(i, .ColIndex("����")) <> "" Then
                        strWindow = strWindow & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("����"))
                    End If
                Next i
                If strWindow <> "" Then strWindow = Mid(strWindow, 2)
                Call SetParChange(vsfDrugStore, Index, mrsPar, True, strWindow, CStr(Col))
            Case Else
            End Select
        End With
        Exit Sub
    End If
End Sub

Private Sub vsfTriageQueuingDep_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim l As Long
    Dim strTmp As String

    With vsfTriageQueuingDep
        If .TextMatrix(Row, .ColIndex("ID")) = "" Then Exit Sub
        If .RowData(Row) <> .TextMatrix(Row, .ColIndex("����")) Then
            .Cell(flexcpForeColor, Row, .ColIndex("����")) = vbRed
        Else
            .Cell(flexcpForeColor, Row, .ColIndex("����")) = &H80000008
        End If
    End With
End Sub

Private Sub vsfTriageQueuingDep_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfTriageQueuingDep
        If .ColKey(Col) <> "����" Then Cancel = 1
    End With
End Sub

Private Sub vsfTriageQueuingDep_DblClick()
    With vsfTriageQueuingDep
        If .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        .TextMatrix(.RowSel, .ColIndex("����")) = IIF(.TextMatrix(.RowSel, .ColIndex("����")) <> "0", "0", "1")
        Call vsfTriageQueuingDep_AfterEdit(.RowSel, .ColSel)
    End With
End Sub

Private Sub vsInputItemSet_DblClick(Index As Integer)
    Call SetInputItemValue(Index)
End Sub
Private Sub vsInputItemSet_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsInputItemSet(Index)
            Select Case Index
            Case vsGrid_��������������
                If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                   If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
                   Exit Sub
                End If
                
                zlVsMoveGridCell vsInputItemSet(Index), 1, .Cols - 1
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call SetInputItemValue(Index)
End Sub
Private Sub vsInputItemSet_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsInputItemSet, Index, mrsPar)
End Sub

Private Sub vsStationRegSort_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String, i As Integer
    With vsStationRegSort
        If Col <> .ColIndex("�Ƿ�����") Then Exit Sub
        For i = 1 To 5
            strValue = strValue & "|" & .TextMatrix(i, .ColIndex("�����ֶ�")) & "," & IIF(.TextMatrix(i, .ColIndex("�Ƿ�����")) = -1, 1, 0)
        Next i
        strValue = Mid(strValue, 2)
    End With
    Call SetParChange(vsStationRegSort, 0, mrsPar, True, strValue)
End Sub

'----------------------------------------------------
Private Sub vs����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vs����
        Select Case Col
        Case .ColIndex("�̶����")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, .Col)), "###0.00;-###0.00;;")
            Call SetParChange(vs����, 0, mrsPar, True, Get���ս��)
        Case .ColIndex("ѡ��")
        Case Else
        End Select
    End With
End Sub
Private Sub vs����_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs����
        Select Case Col
        Case .ColIndex("�̶����")
            Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vs����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs����
        If .Col >= .ColIndex("�̶����") And .Row = .Rows - 1 Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        If .Row < .Rows - 1 Then
           .Row = .Row + 1
        End If
    End With
End Sub

Private Sub vs����_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '�༭����
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs����
        Select Case Col
        Case .ColIndex("�̶����")
                If Row < .Rows - 1 Then
                    .Col = Col: .Row = .Row + 1
                Else
                    zlCommFun.PressKey vbKeyTab
                End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vs����_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vs����
        Select Case .Col
            Case .ColIndex("�̶����")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    If KeyAscii = Asc(".") Then
                        If InStr(1, .EditText, ".") = 0 Then
                            Exit Sub
                        End If
                    End If
                    KeyAscii = 0
                End If
            Case Else
        End Select
    End With
End Sub
Private Sub Load���տ�(ByVal strValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ش��տ�
    '���:strValue-���տ�,��ʽ:���㷽ʽ:���|���㷽ʽ:���....
    '����:���˺�
    '����:2011-07-19 15:13:59
    '����:  34705
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, varData As Variant, varTemp As Variant, j As Long, strTmp As String
    
      
    On Error GoTo ErrHandle
    '���㷽ʽ
    strSQL = _
    " Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
    " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
    " Where A.Ӧ�ó���='Ԥ����' And B.����=A.���㷽ʽ And Nvl(B.����,1)=5" & _
    " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    '���㷽ʽ:���|���㷽ʽ:���....
    varData = Split(strValue, "|")
    If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
    With vs����
        .Tag = "1": .Editable = IIF(rsTmp.RecordCount = 0, flexEDNone, flexEDKbdMouse): i = 1
        .Rows = IIF(rsTmp.RecordCount = 0, 1, rsTmp.RecordCount) + 1
        Do While rsTmp.EOF = False
            .TextMatrix(i, .ColIndex("���տ���")) = NVL(rsTmp!����)
            For j = 0 To UBound(varData)
                varTemp = Split(varData(j) & ":", ":")
                If NVL(rsTmp!����) = varTemp(0) Then
                    .TextMatrix(i, .ColIndex("�̶����")) = Format(Val(varTemp(1)), "###0.00;-###0.00;;")
                    Exit For
                End If
            Next
            i = i + 1
            rsTmp.MoveNext
        Loop
        
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function Get���ս��() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ս�����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-06-09 16:25:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTmp  As String
    On Error GoTo ErrHandle
    With vs����
        strTmp = ""
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("�̶����"))) <> 0 And Trim(.TextMatrix(i, .ColIndex("���տ���"))) <> "" Then
                strTmp = strTmp & "|" & Trim(.TextMatrix(i, .ColIndex("���տ���"))) & ":" & Val(.TextMatrix(i, .ColIndex("�̶����")))
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    End With
    Get���ս�� = strTmp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub vs����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vs����, 0, mrsPar)
End Sub
Private Sub SetParRelations(ByRef arrObj As Variant, ByRef rsPar As ADODB.Recordset, _
                        Optional ByVal varPar As Variant, Optional ByVal lngModule As Long, _
                        Optional ByVal strObjTag As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù�������ֵ
    '���:arrObj-��������
    '     varPar-������(����)�������(�ַ�)
    '     lngModule-ģ���
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-06-09 17:56:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngIndex As Long, blnNotClear As Boolean
    
    For i = 0 To UBound(arrObj)
        If i <> 0 Then
            Call zlDatabase.zlInsertCurrRowData(rsPar, mrsPar, "")
            varPar = 0: lngModule = 0
        End If
        lngIndex = 0: blnNotClear = False
        If GetControlIndex(arrObj(i)) >= 0 Then lngIndex = arrObj(i).Index: blnNotClear = True
        Call SetParRelation(arrObj(i), lngIndex, mrsPar, varPar, lngModule, , , blnNotClear)
    Next
End Sub

Private Function GetControlIndex(ByVal obj As Object) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ؼ�������ֵ
    '����:-1��ʾδ��ȡ�������������ǿؼ�����),����Ϊ����ֵ
    '����:���˺�
    '����:2015-06-10 15:48:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    GetControlIndex = obj.Index
    Exit Function
ErrHand:
    GetControlIndex = -1
End Function

Private Sub optDepsoitDelSet_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optDepsoitDelSet, Index, mrsPar)
End Sub

Private Sub optDepsoitDelSet_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDepsoitDelSet_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDepsoitDelSet, Index, mrsPar)
End Sub

Private Sub LoadԤ��Ʊ�ݸ�ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ��Ʊ�ݸ�ʽ
    '����:���˺�
    '����:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_Ԥ��Ʊ�ݸ�ʽ
    
    rsPara.Filter = "ģ��=" & pԤ������� & " And ������='Ԥ����Ʊ��ʽ'"    'Ԥ����Ʊ��ʽ
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "Ԥ����Ʊ��ʽ", pԤ�������, "", vsBillFormat(intIndex).ColIndex("Ʊ�ݸ�ʽ"))
    
    rsPara.Filter = "ģ��=" & pԤ������� & " And ������='Ԥ����Ʊ��ӡ��ʽ'"    'Ԥ����Ʊ��ӡ��ʽ
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "Ԥ����Ʊ��ӡ��ʽ", pԤ�������, "", vsBillFormat(intIndex).ColIndex("Ԥ����ӡ��ʽ"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
 
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsBillFormat(intIndex)
        .TextMatrix(1, 0) = "����Ԥ��"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "סԺԤ��"
        .Cell(flexcpData, 2, 0) = 2
        .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = "0"
        .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = "0"
        '.ForeColor = &H80000008:  .ForeColorFixed = &H80000008
'        Select Case intType
'        Case 1, 3, 5, 15
             .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = 1 ' IIF(intType = 5, 0, 1)
'        End Select
'        Select Case intType1
'        Case 1, 3, 5, 15
             .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = 1 'IIF(intType1 = 5, 0, 1)
'        End Select
'        If (Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Or _
'            Val(.ColData(.ColIndex("Ԥ����ӡ��ʽ"))) = 1) Then
'            .Editable = flexEDKbdMouse
'        Else
            .Editable = flexEDKbdMouse
'        End If
    End With
    
    With vsBillFormat(intIndex)
        .Clear 1: .Rows = 3
        For lngRow = 1 To .Cols - 1
            .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
        Next
    End With
    zl_vsGrid_Para_Restore pԤ�������, vsBillFormat(intIndex), Me.Name, "Ԥ����Ʊ��ӡ��ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetDrugStockEdit(ByVal Index As Integer, ByVal strType As String, ByVal intType As Integer, ByVal lngEditCol As Long, Optional strMachValue As String = "", Optional strDefaultValue As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩ���ı༭����
    '���:strType-���
    '     intType-���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    '     lngEditCol-���Ƶı༭��
    '����:
    '����:
    '����:���˺�
    '����:2009-09-02 14:53:10
    '����:25132
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnSetDefault As Boolean '������ȱʡֵ��,�����������ȱʡֵ
    Dim lngEditForColor As Long, blnAllowEdit As Boolean, bytLockEdit As Integer '1-����,0-������
    
    '���˺�:���ڿ��ܲ���Ȩ�޷������,���,����ͳһ��������,��Ҫ����ĳһ����:
    With vsfDrugStore(Index)
        blnSetDefault = False: blnAllowEdit = True
        bytLockEdit = 0
        If InStr(1, ",1,3,15,", "," & intType & ",") > 0 Then
'            lngEditForColor = IIF(blnAllowEdit, vbBlue, &H8000000C)
            bytLockEdit = IIF(blnAllowEdit, 0, 1)
        ElseIf intType = 5 Then
'            lngEditForColor = vbBlue
        Else
'            lngEditForColor = &H80000008
        End If
        
        For i = 1 To .Rows - 1
            If lngEditCol = .ColIndex("ȱʡ") Then
                '����ҩ��
                If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                    .TextMatrix(i, .ColIndex("ȱʡ")) = IIF(Val(strMachValue) > 0, 1, 0)
                    blnSetDefault = True
                End If
'                 .Cell(flexcpForeColor, i, .ColIndex("ȱʡ")) = lngEditForColor
'                 .Cell(flexcpForeColor, i, .ColIndex("ҩ��")) = lngEditForColor:
            Else
                If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                    .TextMatrix(i, lngEditCol) = strDefaultValue
                End If
                '���ô���
'                 .Cell(flexcpForeColor, i, .ColIndex("����")) = lngEditForColor
            End If
            .Cell(flexcpData, i, lngEditCol) = bytLockEdit
        Next
    End With
End Sub

Private Sub Loadҩ��(ByVal rsPara As ADODB.Recordset)
    Dim rsTemp As ADODB.Recordset, k As Integer
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long, arrWindow As Variant
    Dim intType As Integer
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    intType = 1
    rsPara.Filter = "ģ��=" & pһ��ͨ���Ѳ��� & " And ������='ȱʡ��ҩ��'"    'ȱʡ��ҩ��
    Call SetParRelation(vsfDrugStore, 0, mrsPar, "ȱʡ��ҩ��", pһ��ͨ���Ѳ���, "", vsfDrugStore(0).ColIndex("ȱʡ"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!����ֵ)
    If Val(strTmp) > 0 Then
        Call SetDrugStockEdit(0, "��ҩ��", intType, vsfDrugStore(0).ColIndex("ȱʡ"), Val(strTmp))
    Else
        Call SetDrugStockEdit(0, "��ҩ��", intType, vsfDrugStore(0).ColIndex("ȱʡ"), "")
    End If
    
    rsPara.Filter = "ģ��=" & pһ��ͨ���Ѳ��� & " And ������='ȱʡ��ҩ��'"    'ȱʡ��ҩ��
    Call SetParRelation(vsfDrugStore, 1, mrsPar, "ȱʡ��ҩ��", pһ��ͨ���Ѳ���, "", vsfDrugStore(1).ColIndex("ȱʡ"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!����ֵ)
    If Val(strTmp) > 0 Then
        Call SetDrugStockEdit(1, "��ҩ��", intType, vsfDrugStore(1).ColIndex("ȱʡ"), Val(strTmp))
    Else
        Call SetDrugStockEdit(1, "��ҩ��", intType, vsfDrugStore(1).ColIndex("ȱʡ"), "")
    End If
    
    rsPara.Filter = "ģ��=" & pһ��ͨ���Ѳ��� & " And ������='ȱʡ��ҩ��'"    'ȱʡ��ҩ��
    Call SetParRelation(vsfDrugStore, 2, mrsPar, "ȱʡ��ҩ��", pһ��ͨ���Ѳ���, "", vsfDrugStore(2).ColIndex("ȱʡ"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!����ֵ)
    If Val(strTmp) > 0 Then
        Call SetDrugStockEdit(2, "��ҩ��", intType, vsfDrugStore(2).ColIndex("ȱʡ"), Val(strTmp))
    Else
        Call SetDrugStockEdit(2, "��ҩ��", intType, vsfDrugStore(2).ColIndex("ȱʡ"), "")
    End If
    
    rsPara.Filter = "ģ��=" & pһ��ͨ���Ѳ��� & " And ������='��ҩ������'"    '��ҩ������
    Call SetParRelation(vsfDrugStore, 0, mrsPar, "��ҩ������", pһ��ͨ���Ѳ���, "", vsfDrugStore(0).ColIndex("����"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!����ֵ)
    If strTmp <> "" Then
        arrWindow = Split(strTmp, ",")
        For k = 0 To UBound(arrWindow)
            If arrWindow(k) <> "" Then
                Call SetDrugStockEdit(0, "��ҩ��", intType, vsfDrugStore(0).ColIndex("����"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
            End If
        Next
    Else
        Call SetDrugStockEdit(0, "��ҩ��", intType, vsfDrugStore(0).ColIndex("����"), "")
    End If
    
    rsPara.Filter = "ģ��=" & pһ��ͨ���Ѳ��� & " And ������='��ҩ������'"    '��ҩ������
    Call SetParRelation(vsfDrugStore, 1, mrsPar, "��ҩ������", pһ��ͨ���Ѳ���, "", vsfDrugStore(1).ColIndex("����"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!����ֵ)
    If strTmp <> "" Then
        arrWindow = Split(strTmp, ",")
        For k = 0 To UBound(arrWindow)
            If arrWindow(k) <> "" Then
                Call SetDrugStockEdit(1, "��ҩ��", intType, vsfDrugStore(1).ColIndex("����"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
            End If
        Next
    Else
        Call SetDrugStockEdit(1, "��ҩ��", intType, vsfDrugStore(1).ColIndex("����"), "")
    End If
    
    rsPara.Filter = "ģ��=" & pһ��ͨ���Ѳ��� & " And ������='��ҩ������'"    '��ҩ������
    Call SetParRelation(vsfDrugStore, 2, mrsPar, "��ҩ������", pһ��ͨ���Ѳ���, "", vsfDrugStore(2).ColIndex("����"))
    If Not rsPara.EOF Then strTmp = NVL(rsPara!����ֵ)
    If strTmp <> "" Then
        arrWindow = Split(strTmp, ",")
        For k = 0 To UBound(arrWindow)
            If arrWindow(k) <> "" Then
                Call SetDrugStockEdit(2, "��ҩ��", intType, vsfDrugStore(2).ColIndex("����"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
            End If
        Next
    Else
        Call SetDrugStockEdit(2, "��ҩ��", intType, vsfDrugStore(2).ColIndex("����"), "")
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load�������תסԺԤ��Ʊ��ʽ()
    '����:�����������תסԺԤ��Ʊ��ʽ
    Dim strReport As String, strBillFormat As String, strBillFormat1 As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    With cbo(cbo_�������תסԺԤ����Ʊ��ʽ)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!���) & "-" & NVL(rsTemp!˵��)
            .ItemData(.NewIndex) = Val(NVL(rsTemp!���))
            rsTemp.MoveNext
        Loop
        If .ListCount <> 0 Then .ListIndex = 0
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Loadһ��ͨƱ�ݸ�ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ��ͨ����Ʊ�ݸ�ʽ��
    '����:���˺�
    '����:2015-06-29 13:57:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strReport As String, strBillFormat As String, strBillFormat1 As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    strReport = "ZL" & glngSys \ 100 & "_BILL_1151"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    With cbo(cbo_һ��ͨ_��ƱƱ�ݸ�ʽ)
        .Clear: cbo(cbo_һ��ͨ_����Ʊ�ݸ�ʽ).Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!���) & "-" & NVL(rsTemp!˵��)
            .ItemData(.NewIndex) = Val(NVL(rsTemp!���))
            cbo(cbo_һ��ͨ_����Ʊ�ݸ�ʽ).AddItem NVL(rsTemp!���) & "-" & NVL(rsTemp!˵��)
            cbo(cbo_һ��ͨ_����Ʊ�ݸ�ʽ).ItemData(cbo(cbo_һ��ͨ_����Ʊ�ݸ�ʽ).NewIndex) = Val(NVL(rsTemp!���))
            rsTemp.MoveNext
        Loop
        If .ListCount <> 0 Then .ListIndex = 0
        If cbo(cbo_һ��ͨ_����Ʊ�ݸ�ʽ).ListCount <> 0 Then cbo(cbo_һ��ͨ_����Ʊ�ݸ�ʽ).ListIndex = 0
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub

Private Sub Load�շ�Ʊ�ݸ�ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�Ʊ�ݸ�ʽ
    '����:���˺�
    '����:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varPatiData As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    Dim strPatiBillFormat As String '�����˲���Ʊ�ݸ�ʽ
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_�շ�Ʊ�ݸ�ʽ
    
    rsPara.Filter = "ģ��=" & p�����շѹ��� & " And ������='�շѷ�Ʊ��ʽ'"    '�շѷ�Ʊ��ʽ
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!����ֵ)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�շѷ�Ʊ��ʽ", p�����շѹ���, "", vsBillFormat(intIndex).ColIndex("�շ�Ʊ�ݸ�ʽ"))
    
    rsPara.Filter = "ģ��=" & p�����շѹ��� & " And ������='�����˲���Ʊ��ʽ'"    '�����˲���Ʊ��ʽ
    If Not rsPara.EOF Then strPatiBillFormat = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�����˲���Ʊ��ʽ", p�����շѹ���, "", vsBillFormat(intIndex).ColIndex("�����˲���Ʊ�ݸ�ʽ"))
    
    
    rsPara.Filter = "ģ��=" & p�����շѹ��� & " And ������='�շѷ�Ʊ��ӡ��ʽ'"    '�շѷ�Ʊ��ӡ��ʽ
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!����ֵ)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�շѷ�Ʊ��ӡ��ʽ", p�����շѹ���, "", vsBillFormat(intIndex).ColIndex("�շѴ�ӡ��ʽ"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")
    varPatiData = Split(strPatiBillFormat, "|")
    
    strReport = "ZL" & glngSys \ 100 & "_BILL_1121_1"
    Set rsTemp = zlGetBillFormatRec(strReport)
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("�շ�Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�����˲���Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
 
 
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    rsTemp.Filter = "ID<>0"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsBillFormat(vsGrid_�շ�Ʊ�ݸ�ʽ)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = "0"
            
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varPatiData)
                varTemp = Split(varPatiData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("�����˲���Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p�����շѹ���, vsBillFormat(intIndex), Me.Name, "�շ�Ʊ�ݸ�ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load�˷�Ʊ�ݸ�ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����˷�Ʊ�ݸ�ʽ
    '����:Ƚ����
    '����:2016-06-1
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varPatiData As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_�˷�Ʊ�ݸ�ʽ
    
    rsPara.Filter = "ģ��=" & p�����շѹ��� & " And ������='�˷ѷ�Ʊ��ʽ'"    '�˷ѷ�Ʊ��ʽ
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�˷ѷ�Ʊ��ʽ", p�����շѹ���, "", vsBillFormat(intIndex).ColIndex("�շ�Ʊ�ݸ�ʽ"))
    
    rsPara.Filter = "ģ��=" & p�����շѹ��� & " And ������='�˷ѷ�Ʊ��ӡ��ʽ'"    '�˷ѷ�Ʊ��ӡ��ʽ
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�˷ѷ�Ʊ��ӡ��ʽ", p�����շѹ���, "", vsBillFormat(intIndex).ColIndex("�շѴ�ӡ��ʽ"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")
    strReport = "ZL" & glngSys \ 100 & "_BILL_1121_7"
    Set rsTemp = zlGetBillFormatRec(strReport)
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("�շ�Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
 
 
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    rsTemp.Filter = "ID<>0"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = "0"
            
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p�����շѹ���, vsBillFormat(intIndex), Me.Name, "�˷�Ʊ�ݸ�ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub Load������Ʊ�ݸ�ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ʊ�ݸ�ʽ
    '����:���˺�
    '����:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_������Ʊ�ݸ�ʽ
    
    rsPara.Filter = "ģ��=" & p���ﲹ���� & " And ������='�շѷ�Ʊ��ʽ'"     '�շѷ�Ʊ��ʽ
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!����ֵ)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�շѷ�Ʊ��ʽ", p���ﲹ����, "", vsBillFormat(intIndex).ColIndex("Ʊ�ݸ�ʽ"))
    rsPara.Filter = "ģ��=" & p���ﲹ���� & " And ������='�շѷ�Ʊ��ӡ��ʽ'"    '�շѷ�Ʊ��ӡ��ʽ
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!����ֵ)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�շѷ�Ʊ��ӡ��ʽ", p���ﲹ����, "", vsBillFormat(intIndex).ColIndex("�շѴ�ӡ��ʽ"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1124"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    If zlStartFactUseType(1) Then
        rsTemp.Filter = "ID<>0"
    End If
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p���ﲹ����, vsBillFormat(intIndex), Me.Name, "������Ʊ�ݸ�ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Loadҽ�ƿ�Ʊ�ݸ�ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ʊ�ݸ�ʽ
    '����:���˺�
    '����:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_ҽ�ƿ��վݸ�ʽ
    
    rsPara.Filter = "ģ��=" & pҽ�ƿ����� & " And ������='ҽ�ƿ��վݸ�ʽ'"      '�շѷ�Ʊ��ʽ
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!����ֵ)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "ҽ�ƿ��վݸ�ʽ", pҽ�ƿ�����, "", vsBillFormat(intIndex).ColIndex("Ʊ�ݸ�ʽ"))
    
    varData = Split(strBillFormat, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1107"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
    End With
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = 3
        lngRow = 1
        Dim j As Integer
        For i = 1 To 2
            .TextMatrix(lngRow, .ColIndex("��������")) = IIF(i = 1, "����", "�󶨿�")
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = IIF(i = 1, Val(varData(0)), Val(varData(1)))
            lngRow = lngRow + 1
        Next
    End With
    zl_vsGrid_Para_Restore pҽ�ƿ�����, vsBillFormat(intIndex), Me.Name, "ҽ�ƿ��վݸ�ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load�������˷�Ʊ�ݸ�ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز������˷�Ʊ�ݸ�ʽ
    '����:Ƚ����
    '����:2016-06-1
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_�������˷�Ʊ�ݸ�ʽ
    
    rsPara.Filter = "ģ��=" & p���ﲹ���� & " And ������='�˷ѷ�Ʊ��ʽ'"     '�˷ѷ�Ʊ��ʽ
    If Not rsPara.EOF Then strBillFormat = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�˷ѷ�Ʊ��ʽ", p���ﲹ����, "", vsBillFormat(intIndex).ColIndex("Ʊ�ݸ�ʽ"))
    
    rsPara.Filter = "ģ��=" & p���ﲹ���� & " And ������='�˷ѷ�Ʊ��ӡ��ʽ'"    '�˷ѷ�Ʊ��ӡ��ʽ
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!����ֵ)
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "�˷ѷ�Ʊ��ӡ��ʽ", p���ﲹ����, "", vsBillFormat(intIndex).ColIndex("�շѴ�ӡ��ʽ"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")

    strReport = "ZL" & glngSys \ 100 & "_BILL_1124_3"
    Set rsTemp = zlGetBillFormatRec(strReport)
    
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    If zlStartFactUseType(1) Then
        rsTemp.Filter = "ID<>0"
    End If
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p���ﲹ����, vsBillFormat(intIndex), Me.Name, "�������˷�Ʊ�ݸ�ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Load����Ʊ�ݸ�ʽ(ByVal rsPara As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ���Ʊ�ݸ�ʽ
    '����:���˺�
    '����:2015-06-10 11:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String, strBillFormat As String, strPrintMode As String
    Dim varData As Variant, varType As Variant, varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, intIndex As Integer, i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    intIndex = vsGrid_����Ʊ�ݸ�ʽ
    
    rsPara.Filter = "ģ��=" & p���˽��ʹ��� & " And ������='���˽��ʴ�ӡ'"
    If Not rsPara.EOF Then strPrintMode = NVL(rsPara!����ֵ)
    
    Call SetParRelation(vsBillFormat, intIndex, mrsPar, "���˽��ʴ�ӡ", p���˽��ʹ���, "", vsBillFormat(intIndex).ColIndex("���ʺ��ӡ��ʽ"))
    
    
    varData = Split(strBillFormat, "|"): varType = Split(strPrintMode, "|")
    
    With vsBillFormat(intIndex)
        .Clear 1
        .ColComboList(.ColIndex("���ʺ��ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    
    If GetBillUseTypeRec(rsTemp) = False Then Exit Sub
    
    rsTemp.Filter = "ID<>0"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
    With vsBillFormat(intIndex)
        .Editable = flexEDKbdMouse
        .Clear 1
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���ʺ��ӡ��ʽ")) = "0-����ӡƱ��"
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(NVL(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("���ʺ��ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next

            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    rsTemp.Filter = 0
    zl_vsGrid_Para_Restore p���˽��ʹ���, vsBillFormat(intIndex), Me.Name, "����Ʊ�ݸ�ʽ", False, False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlReadBillFormat(ByVal ReportCode As String) As ADODB.Recordset
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ������Ĵ�ӡ��ʽ
    '���:ReportCode-��������
    '����:�����ӡ��ʽ�ļ�¼��
    '����:���ϴ�
    '����:2014-10-20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo ErrHandle
    
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���='" & ReportCode & "'  " & _
    "   Order by  ���"
    Set zlReadBillFormat = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetBillFormat(ByVal intIndex As Integer, ByVal intType As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݲ�������
    '���:intIndex-Ʊ�ݴ�ӡ��ʽ����
    '     intType-0-��ȡƱ�ݴ�ӡ��ʽ;1-��ȡƱ�ݸ�ʽ;2-��ȡ�����˲���Ʊ�ݸ�ʽ
    '����:����Ʊ�ݴ�ӡ��ʽ���ӡ��ʽ
    '����:���˺�
    '����:2015-06-10 14:01:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrintMode As String, strPrintFormat As String, strPatiPrintFormat As String
    Dim i As Long
    
    On Error GoTo ErrHandle
    strPrintFormat = "": strPrintMode = ""
    Select Case intIndex
    Case vsGrid_Ԥ��Ʊ�ݸ�ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("Ԥ����ӡ��ʽ")), 1))
                End If
            Next
        End With
    Case vsGrid_Ԥ����Ʊ��ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("�˿��ӡ��ʽ")), 1))
                End If
            Next
        End With
    Case vsGrid_�շ�Ʊ�ݸ�ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                '80943,Ƚ����,2014-12-18,Ʊ��δʹ�á��շ����ʱ�����������շ����Ϊ�յĴ�ӡ��ʽ��Ʊ�ݸ�ʽ
                'If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�շ�Ʊ�ݸ�ʽ")))
                    strPatiPrintFormat = strPatiPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�����˲���Ʊ�ݸ�ʽ")))
                    strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("�շѴ�ӡ��ʽ")), 1))
                'End If
            Next
        End With
    Case vsGrid_������Ʊ�ݸ�ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                '80943,Ƚ����,2014-12-18,Ʊ��δʹ�á��շ����ʱ�����������շ����Ϊ�յĴ�ӡ��ʽ��Ʊ�ݸ�ʽ
                'If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("�շѴ�ӡ��ʽ")), 1))
                'End If
            Next
        End With
    Case vsGrid_�˷�Ʊ�ݸ�ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�շ�Ʊ�ݸ�ʽ")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("�շѴ�ӡ��ʽ")), 1))
            Next
        End With
    Case vsGrid_�������˷�Ʊ�ݸ�ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("�շѴ�ӡ��ʽ")), 1))
            Next
        End With
    
    Case vsGrid_����Ʊ�ݸ�ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                    'strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                    strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("���ʺ��ӡ��ʽ")), 1))
                End If
            Next
        End With
        
    Case vsGrid_���ʺ�Ʊ��ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                    'strPrintFormat = strPrintFormat & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                    strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("���Ϻ��ӡ��ʽ")), 1))
                End If
            Next
        End With
    Case vsGrid_����Ԥ��Ʊ�ݸ�ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("Ԥ����ӡ��ʽ")), 1))
                End If
            Next
        End With
    Case vsGrid_ҽ�ƿ��վݸ�ʽ
        With vsBillFormat(intIndex)
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("��������"))) <> "" Then
                    strPrintFormat = strPrintFormat & "|" & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                End If
            Next
        End With
    End Select
    If strPrintFormat <> "" Then strPrintFormat = Mid(strPrintFormat, 2)
    If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
    If strPatiPrintFormat <> "" Then strPatiPrintFormat = Mid(strPatiPrintFormat, 2)
    '0-��ȡƱ�ݴ�ӡ��ʽ;1-��ȡƱ�ݸ�ʽ;2-��ȡ�����˲���Ʊ�ݸ�ʽ
    If intType = 2 Then GetBillFormat = strPatiPrintFormat: Exit Function
    GetBillFormat = IIF(intType = 0, strPrintMode, strPrintFormat)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub SetInputItemValue(ByVal intIndex As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ�ǰ��Ŀ�����ֵ
    '���:intIndex-����ؼ����������ֵ
    '����:���˺�
    '����:2015-06-11 17:58:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
       
    On Error GoTo ErrHandle
    With vsInputItemSet(intIndex)
        Select Case .Col
        Case .ColIndex("��ֹ¼��")
            .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = IIF(.TextMatrix(.Row, .ColIndex("��ֹ¼��")) = "", "��", "")
            If .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = "��" Then
                .TextMatrix(.Row, .ColIndex("������")) = ""
                .TextMatrix(.Row, .ColIndex("������")) = ""
                .Cell(flexcpBackColor, .Row, .ColIndex("������"), .Row, .ColIndex("������")) = &H8000000F
            Else
                .Cell(flexcpBackColor, .Row, .ColIndex("������"), .Row, .ColIndex("������")) = &H8000000E
            End If
            .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000E
        Case .ColIndex("������")
        
            .TextMatrix(.Row, .ColIndex("������")) = IIF(.TextMatrix(.Row, .ColIndex("������")) = "", "��", "")
            If .TextMatrix(.Row, .ColIndex("������")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = ""
                .TextMatrix(.Row, .ColIndex("������")) = "��"
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000F
                .Cell(flexcpBackColor, .Row, .ColIndex("������")) = &H8000000E
            ElseIf .TextMatrix(.Row, .ColIndex("������")) = "��" Then
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000F
                .Cell(flexcpBackColor, .Row, .ColIndex("������")) = &H8000000E
            Else
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��"), .Row, .ColIndex("������")) = &H8000000E
            End If
             .Cell(flexcpBackColor, .Row, .ColIndex("������")) = &H8000000E
        Case .ColIndex("������")
            .TextMatrix(.Row, .ColIndex("������")) = IIF(.TextMatrix(.Row, .ColIndex("������")) = "", "��", "")
             .Cell(flexcpBackColor, .Row, .ColIndex("������")) = &H8000000E
            If .TextMatrix(.Row, .ColIndex("������")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = ""
                
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000F
            ElseIf .TextMatrix(.Row, .ColIndex("������")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = ""
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000F
            Else
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��"), .Row, .ColIndex("������")) = &H8000000E
            End If
        End Select
    End With
    Call SetParChange(vsInputItemSet, intIndex, mrsPar, True, GetInputItemSetValue(intIndex))
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function GetInputItemSetValue(ByVal intIndex As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������ֵ
    '���:intIndex-�ؼ�����
    '����:���������õ�ֵ,��ʽ:������,�Ƿ����,����Ƿ�����,�Ƿ������|....
    '����:���˺�
    '����:2015-06-11 18:10:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strTmp As String
    On Error GoTo ErrHandle
        
    With vsInputItemSet(intIndex)
        For i = 1 To .Rows - 1
            strTmp = strTmp & "|" & .TextMatrix(i, .ColIndex("������Ŀ"))
            strTmp = strTmp & "," & IIF(.TextMatrix(i, .ColIndex("��ֹ¼��")) = "��", 1, 0)
            strTmp = strTmp & "," & IIF(.TextMatrix(i, .ColIndex("������")) = "��", 1, 0)
            strTmp = strTmp & "," & IIF(.TextMatrix(i, .ColIndex("������")) = "��", 1, 0)
        Next
    End With
    GetInputItemSetValue = Mid(strTmp, 2)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitPage(ByVal intIndex As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '���:intIndex-ҳ��ؼ����������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-06-15 16:04:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo ErrHand:
     
    If intIndex = Pg_�Һ�ҵ�� Then
        With tbPage(intIndex)
            picRegistPlan.BorderStyle = 0
            picRegist.BorderStyle = 0
            picԤԼ.BorderStyle = 0
            picOtherRegister.BorderStyle = 0
            
            Set ObjItem = .InsertItem(Pg_�Һ�_����, "����", picRegistPlan.hwnd, 0)
            ObjItem.Tag = Pg_�Һ�_����
            Set ObjItem = .InsertItem(Pg_�Һ�_�Һ�, "�ҺŴ���", picRegist.hwnd, 0)
            ObjItem.Tag = Pg_�Һ�_�Һ�
            
            Set ObjItem = .InsertItem(Pg_�Һ�_ԤԼ, "ԤԼ����", picԤԼ.hwnd, 0)
            ObjItem.Tag = Pg_�Һ�_ԤԼ
            Set ObjItem = .InsertItem(Pg_�Һ�_����, "����(ҽ��վ/����̨��)", picOtherRegister.hwnd, 0)
            ObjItem.Tag = Pg_�Һ�_����
            
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
        End With
        Exit Sub
    End If
    If intIndex = Pg_����ҵ�� Then
        With tbPage(intIndex)
            picSettlePar(1).BorderStyle = 0
            picSettlePar(0).BorderStyle = 0
            picSettlePar(2).BorderStyle = 0
            picSettlePar(1).BackColor = &H8000000F
            picSettlePar(0).BackColor = &H8000000F
            picSettlePar(2).BackColor = &H8000000F
        
            Set ObjItem = .InsertItem(Pg_����_���ʲ���, "���ʲ���", picSettlePar(0).hwnd, 0)
            ObjItem.Tag = Pg_����_���ʲ���
            Set ObjItem = .InsertItem(Pg_����_Ʊ�ݿ���, "Ʊ�ݿ���", picSettlePar(1).hwnd, 0)
            ObjItem.Tag = Pg_����_Ʊ�ݿ���
            Set ObjItem = .InsertItem(Pg_����_������, "������", picSettlePar(2).hwnd, 0)
            ObjItem.Tag = Pg_����_������
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
        End With
        Exit Sub
    End If
    If intIndex = Pg_�����շ� Then
      With tbPage(intIndex)
            picChargePg(1).BorderStyle = 0
            picChargePg(0).BorderStyle = 0
            picChargePg(1).BackColor = &H8000000F
            picChargePg(0).BackColor = &H8000000F
            
            Set ObjItem = .InsertItem(Pg_�շ�_���ݿ���, "���ݿ���", picChargePg(0).hwnd, 0)
            ObjItem.Tag = Pg_�շ�_���ݿ���
            Set ObjItem = .InsertItem(Pg_�շ�_Ʊ�ݿ���, "Ʊ�ݿ���", picChargePg(1).hwnd, 0)
            ObjItem.Tag = Pg_�շ�_Ʊ�ݿ���
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
        End With
        Exit Sub
    End If
    
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub cmdAddedItem_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strValue As String
    
    mblnNotChange = True
    strSQL = "Select ID, ����, ����, ���㵥λ, ˵��" & vbNewLine & _
            "From �շ���ĿĿ¼" & vbNewLine & _
            "Where ��� = 'Z' And Nvl(�Ƿ���, 0) = 0 And ������� In(1,3)" & vbNewLine & _
            "Order By ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "(����)�շ���Ŀ")
    If Not rsTmp Is Nothing Then
        txt(txt_�շ�_�������չҺŷ�).Text = NVL(rsTmp!����)
        cmdAddedItem.Tag = NVL(rsTmp!ID)
        If chk(chk_�շ�_δ�Һ��Զ����չҺŷ�).value = 0 Then chk(chk_�շ�_δ�Һ��Զ����չҺŷ�).value = 1
    End If
    strValue = cmdAddedItem.Tag & ";" & txt(txt_�շ�_�������չҺŷ�).Text
    Call SetParChange(txt, txt_�շ�_�������չҺŷ�, mrsPar, True, strValue)
    Call SetParChange(chk, chk_�շ�_δ�Һ��Զ����չҺŷ�, mrsPar, True, strValue)
    
    mblnNotChange = False
    Exit Sub
errH:
    mblnNotChange = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPrintListHaveData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݴ�ӡ��ϸ�Ƿ�������
    '����:�����ݷ���true,���򷵻�False
    '����:���˺�
    '����:2013-05-17 14:24:40
    '˵��:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHandle
    strSQL = "Select 1 From Ʊ�ݴ�ӡ��ϸ where Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    GetPrintListHaveData = rsTemp.RecordCount >= 1
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetBillRuleParaLocale()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݷ�������λ��
    '����:���˺�
    '����:2013-03-26 15:43:12
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intIndex As Integer
    On Error GoTo ErrHandle
    
    intIndex = cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex
    
    If intIndex < 0 Or intIndex > 2 Then intIndex = 0
    
    For i = 0 To 2
        picRuleBack(i).Visible = intIndex = i
    Next
    
    If intIndex = 0 Then
        '����ʵ�ʴ�ӡ����Ʊ��
        '��������
        Set chk(chk_�շ�_�Զ����չ�����).Container = picRuleBack(0)
        '������
        With chk(chk_�շ�_�Զ����չ�����)
            .Top = fraActuallyPrint.Top
            .Left = chk(chk_�շ�_ÿ��ֻ��һ��Ʊ��).Left
            .TabIndex = chk(chk_�շ�_ÿ��ֻ��һ��Ʊ��).TabIndex + 1
        End With
     End If

    If intIndex = 1 Then
        '����Ԥ���������Ʊ��
        '��������
        Set chk(chk_�շ�_�Զ����չ�����).Container = picRuleBack(1)
        '������
       With chk(chk_�շ�_�Զ����չ�����)
            .Top = fraRuleSystem.Top
            .Left = fraRuleSystem.Left + 100
            .TabIndex = chkBillRule(0).TabIndex - 1
       End With
    End If
    If intIndex = 2 Then
        '�����û��Զ��������
        '��������
        Set chk(chk_�շ�_�Զ����չ�����).Container = picRuleBack(2)
        '������
       With chk(chk_�շ�_�Զ����չ�����)
            .Top = lblCustomInfor.Top - .Height - 50
            .Left = lblCustomInfor.Left
            .TabIndex = cbo(intIndex).TabIndex + 1
        End With
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub chkBillRule_Click(Index As Integer)
    If Me.Visible = False Then Exit Sub
    
    If Index <> 0 And chkBillRule(Index).value = 1 Then
        If Val(txtBillRuleNum(Index - 1).Text) = 0 Then
            updBillRuleNum(Index - 1).value = Val(txtBillRuleNum(Index - 1).Tag)    '�ָ�ȱʡֵ
        End If
    End If
    Call SetBillRuleEnable
    Call ShowRuleInfor
    If Not optRuleTotal(2).Visible Then
         If optRuleTotal(2).value Then optRuleTotal(0).value = True
    End If
    Call SaveBillRuleChange
End Sub


Private Sub ShowRuleInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾƱ�ŵķ������
    '����:���˺�
    '����:2013-03-26 14:14:08
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, i As Integer
    Dim strName As String
    
    On Error GoTo ErrHandle
    strInfor = ""
    If chkBillRule(0).value = 1 Then
       strInfor = strInfor & "+ NO"
    End If
    For i = 1 To 3
        If chkBillRule(i).value = 1 Then
            strName = Switch(i = 1, "ִ�п���", i = 2, "�վݷ�Ŀ", True, "�վ�ϸĿ")
            strInfor = strInfor & "+" & strName & "(" & txtBillRuleNum(i - 1).Text & ")"
        End If
    Next
    If strInfor <> "" Then strInfor = Mid(strInfor, 2)
    lblInfor.Caption = strInfor
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetBillRuleEnable()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݷ������,������Ӧ�ؼ���Enabled����
    '����:���˺�
    '����:2013-03-26 17:55:47
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer, blnEnable As Boolean
    On Error GoTo ErrHandle
    '��������(0-������;1-��ҳ����(����1ҳ����),2-�������(ѡ����ϸʱ��Ч))
    '1.�������:�з�����ϸʱ,ͬʱ���ڹ�ѡִ�п��һ����վݷ�Ŀ�򰴵��ݲŻ���ڷ��������,�Ż���ڻ�����
    blnEnable = chkBillRule(3).Enabled And chkBillRule(3).value = 1 And (chkBillRule(2).value = 1 Or chkBillRule(1).value = 1 Or chkBillRule(0).value = 1)
    optRuleTotal(2).Visible = blnEnable
    optRuleTotal(2).Enabled = blnEnable
    '2.��ҳ����:���������óɻ��ܶ�
    optRuleTotal(1).Enabled = chkBillRule(3).Enabled
    optRuleTotal(0).Enabled = chkBillRule(3).Enabled
    
    '���÷����������
    If chkBillRule(0).value = 1 Then
        optRuleTotal(2).Caption = "�����ݺŷ������"
    ElseIf chkBillRule(1).value = 1 Then
        optRuleTotal(2).Caption = "��ִ�п��ҷ������"
    ElseIf chkBillRule(3).value = 1 Then
        optRuleTotal(2).Caption = "���վݷ�Ŀ�������"
    ElseIf chkBillRule(3).value = 1 Then
        optRuleTotal(2).Caption = "��������������"
    End If
    For intIndex = 1 To 3
        txtBillRuleNum(intIndex - 1).Enabled = chkBillRule(intIndex).value = 1 And chkBillRule(intIndex).Enabled
        updBillRuleNum(intIndex - 1).Enabled = txtBillRuleNum(intIndex - 1).Enabled
        lblBillRuleNum(intIndex - 1).Enabled = txtBillRuleNum(intIndex - 1).Enabled
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function SaveƱ�ݷ������() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݷ���������仯
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-06-18 12:13:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intList As Integer
    On Error GoTo ErrHandle
    intList = cbo(cbo_�շ�_Ʊ�ݷ������).ListIndex
    
    If mintԭƱ�ݷ������ <> intList And intList <> 0 And mintԭƱ�ݷ������ <= 0 Then
       '�����ǰ�л�����ģʽ,��Ҫ��Ʊ�ݴ�ӡ��ʽ��¼����,�Ա����ش�򲿷��˷�ʱ���л�ǰ��Ʊ�ݸ�ʽ��ӡ
       Call zlDatabase.ExecuteProcedure("Zl_Update_Bill_Printformat(" & glngSys & ")", Me.Caption)
    End If
    SaveƱ�ݷ������ = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub updBillRuleNum_Change(Index As Integer)
    If Me.Visible = False Then Exit Sub
    
    If updBillRuleNum(Index).value = 0 Then
        chkBillRule(Index + 1).value = 0
    End If
     Call SaveBillRuleChange
End Sub
Private Sub InitBillRuleCtrl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ʊ�ݹ������ؿؼ�
    '����:���˺�
    '����:2015-06-19 14:43:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    On Error GoTo ErrHandle
    For i = 0 To picRuleBack.UBound
        picRuleBack(i).BorderStyle = 0
    Next
    cbo(cbo_�շ�_Ʊ�ݷ������).ZOrder 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function zlStartFactUseType(ByVal bytBillType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�ʹ����ʹ������
    '���:bytBillType-Ʊ��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-10 16:11:47
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo ErrHandle
    strSQL = "Select  1 as ���� From Ʊ�����ü�¼ where Ʊ��=[1] and nvl(ʹ�����,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���Ʊ���Ƿ�������ʹ������", bytBillType)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Load������㷽ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����㷽ʽ
    '����:���˺�
    '����:2015-06-23 14:14:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrHandle
    '�ų����ѿ���Ӧ�Ľ��㷽ʽ(����8)���Լ�δ����һ��ͨ����һ��ͨ���㷽ʽ(����7)
    strSQL = _
        "Select Distinct b.����, b.����" & vbNewLine & _
        "From ���㷽ʽӦ�� A, ���㷽ʽ B" & vbNewLine & _
        "Where a.���㷽ʽ = b.���� And a.Ӧ�ó��� In ('�Һ�', '�շ�')" & vbNewLine & _
        "     And Instr(',3,4,', ',' || b.���� || ',') = 0" & vbNewLine & _
        "     And (b.���� <> 7 Or b.���� = 7 And Exists (Select 1 From һ��ͨĿ¼ Where ���㷽ʽ = b.���� And ���� = 1))" & vbNewLine & _
        "     And (b.���� <> 8 Or b.���� = 8 And Not Exists (Select 1 From ���ѿ����Ŀ¼ Where ���㷽ʽ = b.����))" & vbNewLine & _
        "Order By LPad(����, 3, ' ')"
    strSQL = _
        "Select ����, ����" & vbNewLine & _
        "From (" & strSQL & ")" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select '00', '��Ԥ���' From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ղ������")
    
    With lst(lst_������_���㷽ʽ)
        .Clear
        Do Until rsTemp.EOF
            .AddItem NVL(rsTemp!����)
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetBillUseTypeRec(ByRef rsUseType As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ��ʹ�����
    '����:rsUseType-ʹ�����
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-06-24 10:35:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    If Not mrsBillUseType Is Nothing Then
        If mrsBillUseType.State = 1 Then
            Set rsUseType = mrsBillUseType: GetBillUseTypeRec = True: Exit Function
        End If
    End If
    strSQL = "" & _
    "   Select rowNum as ID,����, ���� From Ʊ��ʹ�����" & _
    "   Union All" & _
    "   Select 0 as ID, '', '' From Dual " & _
    "   Order By ����"
    Set mrsBillUseType = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set rsUseType = mrsBillUseType
    GetBillUseTypeRec = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Load�Էѷ������()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Էѷ������
    '����:���˺�
    '����:2015-06-24 15:34:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrHandle
    
    strSQL = "" & _
    "   Select ����,���� as ��� " & _
    "   From �շ���Ŀ��� " & _
    "   Where ���� <> '1'" & _
    "   Order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With lst(lst_����_�Էѷ������)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!���
            .ItemData(.NewIndex) = Asc(rsTemp!����)
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetBillLenSet() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݳ�������
    '����:����Ʊ�ݳ�������,��ʽ:1-�շ�,2-Ԥ��,3-����,4-�Һ�
    '����:���˺�
    '����:2015-07-28 17:07:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    strTemp = strTemp & lvw(lvw_Ʊ��).ListItems("C1").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_Ʊ��).ListItems("C2").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_Ʊ��).ListItems("C3").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_Ʊ��).ListItems("C4").SubItems(1)
    GetBillLenSet = strTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetBillCtlSet() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݿ�������
    '����:����Ʊ�ݿ�������:��ʽ:1111 �ֱ�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    '����:���˺�
    '����:2015-07-28 17:07:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    On Error GoTo ErrHandle

    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C1").SubItems(2) = "��", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C2").SubItems(2) = "��", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C3").SubItems(2) = "��", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C4").SubItems(2) = "��", "1", "0")
    GetBillCtlSet = strTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub optPrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optPrintMode, Index, mrsPar)
End Sub

Private Sub optPrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrintMode, Index, mrsPar)
End Sub


Private Sub optToExcelMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optToExcelMode, Index, mrsPar)
End Sub

Private Sub optToExcelMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optToExcelMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optToExcelMode, Index, mrsPar)
End Sub

Private Sub optVisitTablePrintMode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optVisitTablePrintMode, Index, mrsPar)
End Sub

Private Sub optVisitTablePrintMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optVisitTablePrintMode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optVisitTablePrintMode, Index, mrsPar)
End Sub

Private Sub LoadDelFeeDefaultType()
    '����ȱʡ�˷ѷ�ʽ
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Long, strValue As String
    
    Call SetParRelation(vsfDelFeeDefaultType, 0, mrsPar, "���������˷�ȱʡ��ʽ", p�����շѹ���)
    mrsPar.Filter = "ģ��=" & p�����շѹ��� & " And ������='���������˷�ȱʡ��ʽ'"
    If Not mrsPar.EOF Then strValue = NVL(mrsPar!����ֵ)
    
    strSQL = _
        "Select ����" & vbNewLine & _
        "From ���㷽ʽ A, ���㷽ʽӦ�� B" & vbNewLine & _
        "Where a.���� = b.���㷽ʽ And b.Ӧ�ó��� = '�շ�' And a.���� = 2 And Nvl(a.Ӧ����, 0) = 0 Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsfDelFeeDefaultType
        .Clear 1
        .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("�տʽ")) = NVL(rsTemp!����)
            If InStr(strValue, NVL(rsTemp!����)) > 0 Then
                .TextMatrix(lngRow, .ColIndex("ȱʡ����")) = 1
            End If
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    zl_vsGrid_Para_Restore p�����շѹ���, vsfDelFeeDefaultType, Me.Name, "�˷�ȱʡ��ʽ", False, False
End Sub

Private Sub SaveTriageQueuingDep()
'����:���沿�ŷ���̨ǩ���ŶӲ���
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    With vsfTriageQueuingDep
        For i = 1 To .Rows - 1
            strSQL = "zl_Parameters_Update('����̨ǩ���Ŷ�','" & IIF(.TextMatrix(i, .ColIndex("����")) = "0", "0", "1") & "'," & _
                                          glngSys & "," & p������� & ",1," & Val(.TextMatrix(i, .ColIndex("ID"))) & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "�������������̨ǩ���Ŷӡ�")
            Call zlDatabase.ClearParaCache
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTriageQueuingDep()
'���ܣ����ط���̨ǩ���ŶӲ������ÿ���
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHandle
    
    mrsPar.Filter = "ģ��=" & p������� & " And ������='����̨ǩ���Ŷ�'"
    If mrsPar.RecordCount <= 0 Then Exit Sub
        
    With vsfTriageQueuingDep
        If chk(chk_����_����̨ǩ����ʼ�Ŷ�).value = 0 Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("����")) = 0
            Next
            Exit Sub
        End If
        .Rows = 1
        strSQL = " Select a.Id, a.����, a.���� As ����, Nvl(b.����ֵ, 1) As ����" & vbNewLine & _
                 " From ���ű� A," & vbNewLine & _
                 "     (Select b.����id, b.����ֵ" & vbNewLine & _
                 "       From zlParameters A, Zldeptparas B" & vbNewLine & _
                 "       Where a.Id = b.����id And a.ϵͳ = 100 And a.ģ�� = 1113 And a.������ = '����̨ǩ���Ŷ�') B, ��������˵�� C" & vbNewLine & _
                 " Where a.Id = b.����id(+) And a.Id = c.����id And c.�������� = '�ٴ�'" & vbNewLine & _
                 " Order By a.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����̨ǩ���Ŷ����ÿ���")
        
        Do While rsTemp.EOF = False
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, .ColIndex("ID")) = NVL(rsTemp!ID)
            .TextMatrix(i, .ColIndex("����")) = NVL(rsTemp!����)
            .TextMatrix(i, .ColIndex("����")) = NVL(rsTemp!����)
            .TextMatrix(i, .ColIndex("����")) = NVL(rsTemp!����)
            .RowData(i) = NVL(rsTemp!����)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End With
    
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetTriageQueuingEnalbe(Optional ByVal blnEnable As Boolean = False)
    '����:���� "����̨ǩ����ʼ�Ŷ�"������ؿؼ���Enable����
    On Error GoTo ErrHandle
    With vsfTriageQueuingDep
        .Cell(flexcpForeColor, 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = IIF(blnEnable = True, &H80000008, &H8000000C)
        .Enabled = blnEnable
    End With
    cmdDepSelectAll.Enabled = blnEnable
    cmdDepSelectAll.Visible = blnEnable
    cmdDepClearAll.Enabled = blnEnable
    cmdDepClearAll.Visible = blnEnable

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadStationRegOrder(ByVal strOrder As String)
    '����:���ز���"ҽ��վ�Һ��������"��vsStationRegSort�б���
    Dim varOrder As Variant, varData As Variant
    Dim i As Integer
    
    If strOrder = "" Then Exit Sub
    varOrder = Split(strOrder, "|")
    With vsStationRegSort
        For i = 0 To UBound(varOrder)
            varData = Split(varOrder(i), ",")
            If varData(0) <> "" Then
                .TextMatrix(i + 1, .ColIndex("�����ֶ�")) = varData(0)
                .TextMatrix(i + 1, .ColIndex("�Ƿ�����")) = IIF(varData(1) = 0, 0, -1)
            End If
        Next
    End With
End Sub
