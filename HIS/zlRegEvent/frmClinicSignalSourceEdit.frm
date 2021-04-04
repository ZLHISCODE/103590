VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "ZLIDKIND.OCX"
Begin VB.Form frmClinicSignalSourceEdit 
   Caption         =   "出诊号码设置"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicSignalSourceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picDetailedList 
      BorderStyle     =   0  'None
      Height          =   6645
      Left            =   5070
      ScaleHeight     =   6645
      ScaleWidth      =   6990
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   360
      Width           =   6990
      Begin zl9RegEvent.ClinicPlanDetailPages CPDPages 
         Height          =   10620
         Left            =   330
         TabIndex        =   35
         Top             =   390
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   18733
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin VB.PictureBox picWorkTimeList 
      BackColor       =   &H00FFEFE3&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   165
      ScaleHeight     =   2400
      ScaleWidth      =   3195
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6600
      Width           =   3195
      Begin MSComctlLib.ListView lvwWorkTime 
         Height          =   1035
         Left            =   255
         TabIndex        =   34
         Top             =   435
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   1826
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16773091
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "时间段"
            Object.Width           =   9596
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "开始时间"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "终止时间"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgWork 
         Height          =   240
         Left            =   30
         Picture         =   "frmClinicSignalSourceEdit.frx":6852
         Top             =   60
         Width           =   240
      End
      Begin VB.Label lblCalendbarTittle 
         BackStyle       =   0  'Transparent
         Caption         =   "上班时段"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   255
         TabIndex        =   47
         Top             =   90
         Width           =   810
      End
      Begin VB.Shape shpWorkLine 
         BackColor       =   &H00FFEFE3&
         BorderColor     =   &H80000003&
         Height          =   915
         Left            =   30
         Top             =   30
         Width           =   3150
      End
   End
   Begin VB.PictureBox picBaseInfor 
      BackColor       =   &H00FFEFE3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5745
      Left            =   45
      ScaleHeight     =   5745
      ScaleWidth      =   4740
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   390
      Width           =   4740
      Begin VB.Frame fraBaseInfor 
         BackColor       =   &H00FFEFE3&
         BorderStyle     =   0  'None
         Height          =   5355
         Left            =   30
         TabIndex        =   49
         Top             =   375
         Width           =   4650
         Begin VB.Frame fraApplyAgeRange 
            BackColor       =   &H00FFEFE3&
            Caption         =   "适用年龄段"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   150
            TabIndex        =   52
            Top             =   4290
            Width           =   4455
            Begin VB.ComboBox cboAgeUnit 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   645
               Width           =   570
            End
            Begin VB.ComboBox cboAgeUnit 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   3660
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   255
               Width           =   570
            End
            Begin VB.ComboBox cboAgeUnit 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   2280
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   255
               Width           =   570
            End
            Begin VB.ComboBox cboAgeUnit 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   3270
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   645
               Width           =   570
            End
            Begin VB.TextBox txtAgeRange 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   450
               TabIndex        =   29
               Text            =   "100"
               Top             =   645
               Width           =   630
            End
            Begin VB.TextBox txtAgeRange 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   3030
               TabIndex        =   26
               Text            =   "100"
               Top             =   255
               Width           =   630
            End
            Begin VB.TextBox txtAgeRange 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   1650
               TabIndex        =   24
               Text            =   "20"
               Top             =   255
               Width           =   630
            End
            Begin VB.TextBox txtAgeRange 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   2640
               TabIndex        =   32
               Text            =   "20"
               Top             =   645
               Width           =   630
            End
            Begin VB.OptionButton optApplyAgeRange 
               BackColor       =   &H00FFEFE3&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   2370
               TabIndex        =   31
               Top             =   675
               Width           =   240
            End
            Begin VB.OptionButton optApplyAgeRange 
               BackColor       =   &H00FFEFE3&
               Caption         =   "不限制"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   22
               Top             =   285
               Value           =   -1  'True
               Width           =   885
            End
            Begin VB.OptionButton optApplyAgeRange 
               BackColor       =   &H00FFEFE3&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   180
               TabIndex        =   28
               Top             =   675
               Width           =   225
            End
            Begin VB.OptionButton optApplyAgeRange 
               BackColor       =   &H00FFEFE3&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   1380
               TabIndex        =   23
               Top             =   285
               Width           =   225
            End
            Begin VB.Label lblAgeRange 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEFE3&
               Caption         =   "至"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   2850
               TabIndex        =   54
               Top             =   315
               Width           =   180
            End
            Begin VB.Label lblAgeRange 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEFE3&
               Caption         =   "以上"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   1650
               TabIndex        =   55
               Top             =   705
               Width           =   360
            End
            Begin VB.Label lblAgeRange 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEFE3&
               Caption         =   "以下"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   3840
               TabIndex        =   53
               Top             =   705
               Width           =   360
            End
         End
         Begin VB.ComboBox cbo号类 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3225
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   0
            Width           =   1380
         End
         Begin VB.TextBox txt号码 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   885
            TabIndex        =   0
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox cbo科室 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   885
            TabIndex        =   2
            Top             =   405
            Width           =   3705
         End
         Begin VB.Frame fra节假日 
            BackColor       =   &H00FFEFE3&
            Caption         =   "节假日控制方式"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   150
            TabIndex        =   46
            Top             =   3375
            Width           =   4455
            Begin VB.OptionButton opt节假日 
               BackColor       =   &H00FFEFE3&
               Caption         =   "受节假日设置控制"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   150
               TabIndex        =   21
               Top             =   510
               Width           =   1785
            End
            Begin VB.OptionButton opt节假日 
               BackColor       =   &H00FFEFE3&
               Caption         =   "不上班"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   150
               TabIndex        =   18
               Top             =   240
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton opt节假日 
               BackColor       =   &H00FFEFE3&
               Caption         =   "允许预约"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   2775
               TabIndex        =   20
               Top             =   240
               Width           =   1035
            End
            Begin VB.OptionButton opt节假日 
               BackColor       =   &H00FFEFE3&
               Caption         =   "禁止预约"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   1515
               TabIndex        =   19
               Top             =   240
               Width           =   1050
            End
         End
         Begin VB.Frame fra排班方式 
            BackColor       =   &H00FFEFE3&
            Caption         =   "排班方式"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   150
            TabIndex        =   45
            Top             =   2655
            Width           =   4455
            Begin VB.OptionButton opt排班方式 
               BackColor       =   &H00FFEFE3&
               Caption         =   "固定排班"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   120
               TabIndex        =   15
               Top             =   300
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.OptionButton opt排班方式 
               BackColor       =   &H00FFEFE3&
               Caption         =   "月排班"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   1515
               TabIndex        =   16
               Top             =   300
               Width           =   930
            End
            Begin VB.OptionButton opt排班方式 
               BackColor       =   &H00FFEFE3&
               Caption         =   "周排班"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   2760
               TabIndex        =   17
               Top             =   300
               Width           =   945
            End
         End
         Begin VB.TextBox txt出诊频次 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   885
            TabIndex        =   6
            Text            =   "10"
            Top             =   1620
            Width           =   390
         End
         Begin VB.TextBox txt预约天数 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3345
            TabIndex        =   8
            Text            =   "0"
            Top             =   1620
            Width           =   390
         End
         Begin VB.ComboBox cboDoctor 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1605
            TabIndex        =   4
            Top             =   795
            Width           =   2985
         End
         Begin VB.ComboBox cbo收费项目 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1185
            Width           =   3705
         End
         Begin MSComCtl2.UpDown upd出诊频次 
            Height          =   315
            Left            =   1245
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1620
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txt出诊频次"
            BuddyDispid     =   196630
            OrigLeft        =   1350
            OrigTop         =   1058
            OrigRight       =   1605
            OrigBottom      =   1343
            Max             =   1000
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upd预约天数 
            Height          =   315
            Left            =   3735
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1620
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txt预约天数"
            BuddyDispid     =   196631
            OrigLeft        =   3960
            OrigTop         =   1065
            OrigRight       =   4215
            OrigBottom      =   1350
            Max             =   1000
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin zlIDKind.IDKindNew idkDoctor 
            Height          =   300
            Left            =   885
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   795
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   529
            ShowSortName    =   0   'False
            Appearance      =   2
            IDKindStr       =   "内|院内医生|0|0|0|0|0||0|0|0;外|院外医生|0|0|0|0|0||0|0|0"
            CaptionAlignment=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   9
            FontName        =   "宋体"
            IDKind          =   -1
            DefaultCardType =   "0"
            NotAutoAppendKind=   -1  'True
            BackColor       =   16773091
         End
         Begin VB.Frame fraCheck 
            BackColor       =   &H00FFEFE3&
            BorderStyle     =   0  'None
            Height          =   630
            Left            =   30
            TabIndex        =   50
            Top             =   1995
            Width           =   4575
            Begin VB.ComboBox cboApplySex 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3750
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   7
               Width           =   795
            End
            Begin VB.CheckBox chk病案 
               BackColor       =   &H00FFEFE3&
               Caption         =   "挂号时必须建病案"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   105
               TabIndex        =   10
               Top             =   30
               Width           =   1785
            End
            Begin VB.CheckBox chk临床排班 
               BackColor       =   &H00FFEFE3&
               Caption         =   "允许临床科室排班"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2730
               TabIndex        =   14
               Top             =   330
               Width           =   1740
            End
            Begin VB.CheckBox chk节假日换休 
               BackColor       =   &H00FFEFE3&
               Caption         =   "启用节假日换休控制"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   105
               TabIndex        =   13
               Top             =   330
               Width           =   1950
            End
            Begin VB.CheckBox chkApplySex 
               BackColor       =   &H00FFEFE3&
               Caption         =   "适用性别"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   2730
               TabIndex        =   11
               Top             =   52
               Width           =   1065
            End
         End
         Begin VB.Label lblDoctor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医    生"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   105
            TabIndex        =   51
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lbl号类 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "号类"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2790
            TabIndex        =   40
            Top             =   60
            Width           =   360
         End
         Begin VB.Label lbl号码 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "号    码"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   105
            TabIndex        =   39
            Top             =   60
            Width           =   720
         End
         Begin VB.Label lbl预约天数单位 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEFE3&
            Caption         =   "可预约        (天)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2775
            TabIndex        =   44
            Top             =   1680
            Width           =   1620
         End
         Begin VB.Label lbl出诊频次 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出诊频次        (分钟/人次) "
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   105
            TabIndex        =   43
            Top             =   1680
            Width           =   2520
         End
         Begin VB.Label lbl收费项目 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "项    目"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   105
            TabIndex        =   42
            Top             =   1215
            Width           =   720
         End
         Begin VB.Label lbl科室 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "科    室"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   105
            TabIndex        =   41
            Top             =   465
            Width           =   720
         End
      End
      Begin VB.Image imgBase 
         Height          =   240
         Left            =   30
         Picture         =   "frmClinicSignalSourceEdit.frx":6DDC
         Top             =   60
         Width           =   240
      End
      Begin VB.Shape shpBaseLine 
         BackColor       =   &H00FFEFE3&
         BorderColor     =   &H80000003&
         Height          =   585
         Left            =   15
         Top             =   30
         Width           =   480
      End
      Begin VB.Label lblSourceTittle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号源基本信息设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   48
         Top             =   105
         Width           =   1560
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   615
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmClinicSignalSourceEdit.frx":7366
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmClinicSignalSourceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As G_Enum_Fun '0-查看,1-添加,2-调整,3-删除
Private mlngModule As Long
Private mstrPrivs As String
Private mlng号源Id As Long
Private mdtCurDate As Date
Private mblnOk As Boolean
Private mrsDoctor As ADODB.Recordset, mrs科室 As ADODB.Recordset
Private mbln院内医生 As Boolean '当前选中的是否院内医生
Private mblnCboClick As Boolean     '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
Private Enum Pancel_Index
    Pan_BaseInforList = 1001
    Pan_WorkTimeList = 1002
    Pan_DetailList = 1004
End Enum
Private mblnNotCheck As Boolean
Private mobj所有分诊诊室集 As 分诊诊室集
Private mobj所有合作单位 As 合作单位控制集
Private mlngPre科室ID As Long
Private mstr号类 As String
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mobjPubPatient As Object
Private mstrAddNewItem As String
Private mstrNodeNo As String '当前选择部门站点,104620

Private mlngOldFeeItemID As Long '记录原始收费项目ID，用以判断是否调整收费项目
Private mblnUpdateFeeItem As Boolean '是否同步调整未发布的出诊安排的收费项目

Public Function ShowMe(frmParent As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal bytFun As G_Enum_Fun, Optional ByVal lng号源Id As Long, _
    Optional ByRef strAddNewItem As String) As Boolean
    '程序入口
    '出参：
    '   strAddNewItem:新增号源号码
    mlngModule = lngModule: mstrPrivs = strPrivs
    mbytFun = bytFun: mlng号源Id = lng号源Id
    mdtCurDate = zlDatabase.Currentdate
    mstrAddNewItem = ""
    
    Err.Clear: On Error Resume Next
    mblnOk = False
    Me.Show 1, frmParent
    
    If mblnOk Then strAddNewItem = mstrAddNewItem
    ShowMe = mblnOk
End Function

Private Function InitData() As Boolean
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandle
    mlngPre科室ID = 0
    If Val(zlDatabase.GetPara("只允许选院内医生", glngSys, mlngModule, "0")) = 1 Then
        idkDoctor.IDkindStr = "医生|医生|0|0|0|0|0||0|0|0"
        idkDoctor.ToolTipText = "只能选院内建档医生"
    Else
        idkDoctor.IDkindStr = "内|院内医生|0|0|0|0|0||0|0|0;外|院外医生|0|0|0|0|0||0|0|0"
        idkDoctor.ToolTipText = "除了可以选择院内医生外，还可以输入外援医生"
    End If
    
    Set mobj所有分诊诊室集 = GetVisitRoomsObjects(GetDoctorRooms(0))
    Set mobj所有合作单位 = GetUnitsObjects(GetUnitAll())
    
    '上班时段
    If mbytFun = Fun_Update Or mbytFun = Fun_Add Then
        mblnNotCheck = True
        Call LoadWorkTimes(mstrNodeNo, cbo号类.Text)
        mblnNotCheck = False
    End If
    
    '性别
    strSQL = "Select 编码, 名称, 简码, 缺省标志 From 性别 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboApplySex.Clear
    Do While Not rsTemp.EOF
        cboApplySex.AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
        If Val(Nvl(rsTemp!缺省标志, 0)) = 1 Then cboApplySex.ListIndex = cboApplySex.NewIndex
        rsTemp.MoveNext
    Loop
    If cboApplySex.ListIndex < 0 And cboApplySex.ListCount > 0 Then cboApplySex.ListIndex = 0
    
    '年龄单位
    For i = 0 To cboAgeUnit.UBound
        cboAgeUnit(i).AddItem "岁"
        cboAgeUnit(i).AddItem "月"
        cboAgeUnit(i).AddItem "天"
        cboAgeUnit(i).ListIndex = 0
    Next
    
    If mbytFun = Fun_View Or mbytFun = Fun_Delete Then InitData = True: Exit Function
    
    '号类
    strSQL = "Select 编码,名称,简码 From 号类 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo号类.Clear
    Do While Not rsTemp.EOF
        cbo号类.AddItem Nvl(rsTemp!名称)
        rsTemp.MoveNext
    Loop
    If cbo号类.ListIndex < 0 And cbo号类.ListCount > 0 Then cbo号类.ListIndex = 0
    '项目
    strSQL = "Select ID,名称 From 收费项目目录 " & _
        " Where 类别='1' And (Sysdate Between 建档时间 And 撤档时间 Or 建档时间<Sysdate And 撤档时间 Is Null)" & _
        " And (站点='" & gstrNodeNo & "' Or 站点 is Null) " & _
        " Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有可用的挂号项目信息，请先到挂号项目设置中初始！", vbInformation, gstrSysName
        Exit Function
    End If
    cbo收费项目.Clear
    Do While Not rsTemp.EOF
        cbo收费项目.AddItem rsTemp!名称
        cbo收费项目.ItemData(cbo收费项目.NewIndex) = Val(Nvl(rsTemp!id))
        rsTemp.MoveNext
    Loop
    '科室
    Set mrs科室 = GetDepartments("'临床'", "1,3", zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False)
    If mrs科室.RecordCount = 0 Then
        MsgBox "你不具备可用的临床科室信息，请先到部门管理中进行设置！", vbInformation, gstrSysName
        Exit Function
    End If
    cbo科室.Clear
    Do While Not mrs科室.EOF
        cbo科室.AddItem mrs科室!名称
        cbo科室.ItemData(cbo科室.NewIndex) = Val(Nvl(mrs科室!id))
        mrs科室.MoveNext
    Loop
    
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub LoadWorkTimes(ByVal str站点 As String, ByVal str号类 As String)
    '根据选择号类动态加载上班时段
    Dim rsWorkTime As ADODB.Recordset
    Dim strWorkTims As String, objListItem As ListItem
    
    Set rsWorkTime = GetWorkTimes(str站点, str号类)
    rsWorkTime.Sort = "站点 Desc,号类 Desc"
    With lvwWorkTime
        .ListItems.Clear
        If rsWorkTime.RecordCount > 0 Then rsWorkTime.MoveFirst
        strWorkTims = ""
        Do While Not rsWorkTime.EOF
            If InStr(1, strWorkTims & ",", "," & Nvl(rsWorkTime!时间段) & ",") = 0 Then
                Set objListItem = .ListItems.Add(, "K" & Nvl(rsWorkTime!时间段), Nvl(rsWorkTime!时间段) & _
                    "(" & Format(Nvl(rsWorkTime!开始时间), "hh:mm") & "-" & Format(Nvl(rsWorkTime!终止时间), "hh:mm") & ")")
                objListItem.Tag = Nvl(rsWorkTime!时间段)
                objListItem.SubItems(1) = Nvl(rsWorkTime!开始时间)
                objListItem.SubItems(2) = Nvl(rsWorkTime!终止时间)
                strWorkTims = strWorkTims & "," & Nvl(rsWorkTime!时间段)
            End If
            rsWorkTime.MoveNext
        Loop
    End With
End Sub

Private Function CheckExistsPlan(ByVal lng号源Id As Long) As Boolean
    '检查当前号源是否存在安排数据
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandle
    If lng号源Id = 0 Then Exit Function
    strSQL = "Select 1 From 临床出诊安排" & vbNewLine & _
            " Where 号源ID = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前号源是否存在安排数据", lng号源Id)
    CheckExistsPlan = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckExistsNotPublishPlan(ByVal lng号源Id As Long) As Boolean
    '检查当前号源是否存在未发布的月/周安排数据
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandle
    If lng号源Id = 0 Then Exit Function
    strSQL = "Select 1" & vbNewLine & _
            " From 临床出诊安排 A, 临床出诊表 B" & vbNewLine & _
            " Where a.出诊id = b.Id And a.号源id = [1] And Nvl(b.排班方式, 0) In (1, 2)" & vbNewLine & _
            "       And b.发布时间 Is Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前号源是否存在安排数据", lng号源Id)
    CheckExistsNotPublishPlan = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '返回:加载成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-03-23 11:54:49
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rs号源 As ADODB.Recordset
    Dim rs诊室 As ADODB.Recordset
    Dim i As Long, j As Long, blnExitFor As Boolean
    Dim obj出诊记录集 As 出诊记录集
    Dim ObjItem  As ListItem
    Dim obj出诊记录 As 出诊记录, objListItem As ListItem
    Dim lng预约天数 As Long, strTemp As String
    
    Err = 0: On Error GoTo errHandle
    Me.Caption = Choose(mbytFun + 1, "查看", "新增", "修改", "删除") & "号源"
    
    lng预约天数 = zlDatabase.GetPara(66, glngSys, , 15)
    If mbytFun = Fun_Add Then
        '自动填入号码
        txt号码.Text = GetMaxLocalCode("临床出诊号源", "号码")
        txt预约天数.Text = lng预约天数
        
        Set obj出诊记录集 = New 出诊记录集
        obj出诊记录集.出诊日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        cboDoctor.ListIndex = -1
        mblnNotCheck = True
        
        If mblnOk Then  '第二次保存时，才清除
            For Each ObjItem In lvwWorkTime.ListItems
                ObjItem.Checked = False
            Next
        End If
        mblnNotCheck = False
        Call CPDPages.LoadData(obj出诊记录集, mobj所有分诊诊室集, mobj所有合作单位)
        mblnChange = False
        LoadData = True: Exit Function
    End If
    
    strSQL = "" & _
            " Select A.ID, A.号类, A.号码, A.科室id, A.项目id,A.医生ID, A.医生姓名 As 医生, A.预约天数, A.出诊频次," & vbNewLine & _
            "        Nvl(A.是否建病案, 0) As 是否建病案," & vbNewLine & _
            "        Nvl(A.假日控制状态, 0) As 假日控制状态," & vbNewLine & _
            "        Nvl(A.是否临床排班, 0) As 是否临床排班,Nvl(排班方式, 0) As 排班方式," & vbNewLine & _
            "        Nvl(A.是否假日换休, 0) As 是否假日换休,A.撤档时间,nvl(A.是否删除,0) as 是否删除, " & _
            "        B.名称 as 科室名称,C.名称 as 收费项目名称, a.适用性别, a.适用年龄段, b.站点" & vbNewLine & _
            " From 临床出诊号源 A,部门表 B,收费项目目录 C" & vbNewLine & _
            " Where A.ID = [1] and a.科室ID=B.id and A.项目ID=C.ID "
    Set rs号源 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng号源Id)
    If rs号源.EOF Then
        MsgBox "当前号源不存在，可能已被他人删除，请重新刷新数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    txt号码.Text = Nvl(rs号源!号码)
    mstrNodeNo = Nvl(rs号源!站点)
    mblnNotCheck = True
    mstr号类 = Nvl(rs号源!号类)
    zlControl.CboSetText cbo号类, Nvl(rs号源!号类)
    If cbo号类.ListIndex = -1 Then cbo号类.AddItem Nvl(rs号源!号类): cbo号类.ListIndex = cbo号类.NewIndex
    mblnNotCheck = False
    
    zlControl.CboLocate cbo科室, Nvl(rs号源!科室ID), True
    If cbo科室.ListIndex = -1 Then
        cbo科室.AddItem Nvl(rs号源!科室名称): cbo科室.ItemData(cbo科室.NewIndex) = Val(Nvl(rs号源!科室ID))
        cbo科室.ListIndex = cbo科室.NewIndex
    End If
    
    mlngOldFeeItemID = Val(Nvl(rs号源!项目ID))
    zlControl.CboLocate cbo收费项目, Nvl(rs号源!项目ID), True
    If cbo收费项目.ListIndex = -1 Then
        cbo收费项目.AddItem Nvl(rs号源!收费项目名称): cbo收费项目.ItemData(cbo收费项目.NewIndex) = Val(Nvl(rs号源!项目ID))
        cbo收费项目.ListIndex = cbo收费项目.NewIndex
    End If
    If Nvl(rs号源!医生) <> "" Then
        If Val(Nvl(rs号源!医生ID)) = 0 Then
            cboDoctor.AddItem Nvl(rs号源!医生): cboDoctor.ListIndex = cboDoctor.NewIndex
            idkDoctor.IDKind = idkDoctor.ListCount
        Else
            zlControl.CboLocate cboDoctor, Val(Nvl(rs号源!医生ID)), True
            If cboDoctor.ListIndex = -1 Then
                cboDoctor.AddItem Nvl(rs号源!医生): cboDoctor.ItemData(cboDoctor.NewIndex) = Val(Nvl(rs号源!医生ID))
                cboDoctor.ListIndex = cboDoctor.NewIndex
            End If
        End If
    Else
        cboDoctor.ListIndex = -1
    End If
    txt出诊频次.Text = Val(Nvl(rs号源!出诊频次))
    txt预约天数.Text = IIf(Val(Nvl(rs号源!预约天数)) = 0, lng预约天数, Val(Nvl(rs号源!预约天数)))
    
    chkApplySex.Value = IIf(Nvl(rs号源!适用性别) = "", vbUnchecked, vbChecked)
    If Nvl(rs号源!适用性别) <> "" Then
        zlControl.CboLocate cboApplySex, Nvl(rs号源!适用性别)
        If cboApplySex.ListIndex = -1 Then cboApplySex.AddItem Nvl(rs号源!适用性别): cboApplySex.ListIndex = cboApplySex.NewIndex
    End If
    
    chk临床排班.Value = Val(Nvl(rs号源!是否临床排班))
    chk节假日换休.Value = Val(Nvl(rs号源!是否假日换休))
    chk病案.Value = Val(Nvl(rs号源!是否建病案))
    
    opt排班方式(Val(Nvl(rs号源!排班方式))).Value = True
    opt节假日(Val(Nvl(rs号源!假日控制状态))).Value = True
    
    strTemp = Nvl(rs号源!适用年龄段) '格式:开始年龄~终止年龄，用~分隔
    If InStr(strTemp, "~") = 0 Then
        optApplyAgeRange(0).Value = True
    Else
        If Split(strTemp, "~")(0) = "" Then
            optApplyAgeRange(3).Value = True
            Call LoadOldData(Split(strTemp, "~")(1), txtAgeRange(3), cboAgeUnit(3))
            txtAgeRange(3).Width = IIf(cboAgeUnit(3).Visible, 630, 1200)
        ElseIf Split(strTemp, "~")(1) = "" Then
            optApplyAgeRange(2).Value = True
            Call LoadOldData(Split(strTemp, "~")(0), txtAgeRange(2), cboAgeUnit(2))
            txtAgeRange(2).Width = IIf(cboAgeUnit(2).Visible, 630, 1200)
        Else
            optApplyAgeRange(1).Value = True
            Call LoadOldData(Split(strTemp, "~")(0), txtAgeRange(0), cboAgeUnit(0))
            txtAgeRange(0).Width = IIf(cboAgeUnit(0).Visible, 630, 1200)
            Call LoadOldData(Split(strTemp, "~")(1), txtAgeRange(1), cboAgeUnit(1))
            txtAgeRange(1).Width = IIf(cboAgeUnit(1).Visible, 630, 1200)
        End If
    End If
     
    Set obj出诊记录集 = GetClinicRecordFromSignalSource(mlng号源Id)
    If mbytFun = Fun_Update Then Call LoadWorkTimes(mstrNodeNo, Nvl(rs号源!号类))
    With lvwWorkTime
        mblnNotCheck = True
        For Each obj出诊记录 In obj出诊记录集
            Set objListItem = Nothing
            Err = 0: On Error Resume Next
            Set objListItem = .ListItems("K" & obj出诊记录.时间段)
            If Err <> 0 Then
                Set objListItem = .ListItems.Add(, "K" & obj出诊记录.时间段, obj出诊记录.时间段 & _
                    "(" & Format(obj出诊记录.开始时间, "hh:mm") & "-" & Format(obj出诊记录.终止时间, "hh:mm") & ")")
                objListItem.Tag = obj出诊记录.上班时段.时间段
                objListItem.SubItems(1) = obj出诊记录.上班时段.开始时间
                objListItem.SubItems(2) = obj出诊记录.上班时段.结束时间
            End If
            Err = 0: On Error GoTo 0
            If Not objListItem Is Nothing Then objListItem.Checked = True
       Next
       mblnNotCheck = False
    End With
    
    Call CPDPages.LoadData(obj出诊记录集, IIf(mbytFun = Fun_Update, mobj所有分诊诊室集, Nothing), mobj所有合作单位, True)
    
    '控制编辑状态
    If mbytFun = Fun_Delete Or mbytFun = Fun_View Then
        Call SetEnabled(Me.Controls, False)
        CPDPages.EditMode(-1) = ED_RegistPlan_View
    Else
        txt号码.Enabled = False
        
        '当前号源存在安排数据时不允许调整科室、医生、项目
        If CheckExistsPlan(mlng号源Id) Then
            cbo科室.Enabled = False
            idkDoctor.Enabled = False
            cboDoctor.Enabled = False
            '如果是固定安排，则不允许调整收费项目了
            cbo收费项目.Enabled = IIf(opt排班方式(0).Value, False, True)
        End If
    End If
    
    mblnChange = False
    LoadData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboAgeUnit_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboAgeUnit_LostFocus(index As Integer)
    If txtAgeRange(index).Enabled = False Or txtAgeRange(index).Locked Then Exit Sub
    If Trim(txtAgeRange(index).Text) <> "" Then
        If mobjPubPatient Is Nothing Then Exit Sub
        If mobjPubPatient.CheckPatiAge(Trim(txtAgeRange(index).Text) & cboAgeUnit(index).Text) = False Then
            If txtAgeRange(index).Visible And txtAgeRange(index).Enabled And Not txtAgeRange(index).Locked Then
                txtAgeRange(index).SetFocus: Exit Sub
            End If
        End If
    End If
End Sub

Private Sub cboApplySex_Click()
    mblnChange = True
End Sub

Private Sub cboApplySex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboDoctor_Click()
    mblnChange = True
End Sub

Private Sub cboDoctor_LostFocus()
    CPDPages.医生姓名 = cboDoctor.Text
End Sub

Private Sub cbo号类_Change()
    mblnChange = True
End Sub

Private Sub cbo号类_Click()
    On Error GoTo errHandle

    If mstr号类 = cbo号类.Text Or mblnNotCheck Then Exit Sub
    mstr号类 = cbo号类.Text
    mblnChange = True
    
    '号类改变，需要重新提取上班时间段
    Call LoadWorkTimes(mstrNodeNo, cbo号类.Text)
    '重新加载数据
    CPDPages.LoadData New 出诊记录集, mobj所有分诊诊室集, mobj所有合作单位
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo号类_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo收费项目_Click()
    mblnChange = True
End Sub

Private Sub cbo收费项目_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboDoctor_GotFocus()
    zlControl.TxtSelAll cboDoctor
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim obj出诊记录集 As 出诊记录集
    
    Select Case Control.id
    Case conMenu_Edit_Save    '保存数据
        If SaveData = False Then Exit Sub
    Case conMenu_File_Exit   '退出
        Set obj出诊记录集 = CPDPages.Get出诊记录集
        
        If Not obj出诊记录集 Is Nothing Then
            mblnChange = mblnChange Or obj出诊记录集.是否修改
        End If
        
        If mblnChange Then
             If MsgBox("号源信息已经发生改变，但您还未保存，您是否真的要退出吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        Set obj出诊记录集 = Nothing
        Unload Me: Exit Sub
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case conMenu_Edit_Save    '保存数据
        Control.Visible = mbytFun <> Fun_View
    Case conMenu_File_Exit   '退出
    End Select
End Sub

Private Sub chkApplySex_Click()
    mblnChange = True
    cboApplySex.Enabled = chkApplySex.Value = vbChecked
End Sub

Private Sub chkApplySex_GotFocus()
    chkApplySex.BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub chkApplySex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkApplySex_LostFocus()
    chkApplySex.BackColor = fraBaseInfor.BackColor
End Sub

Private Sub chk病案_Click()
    mblnChange = True
End Sub

Private Sub chk病案_GotFocus()
    chk病案.BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub chk病案_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

 

Private Sub chk病案_LostFocus()
    chk病案.BackColor = fraBaseInfor.BackColor
End Sub

Private Sub chk节假日换休_Click()
    mblnChange = True
End Sub

Private Sub chk节假日换休_GotFocus()
       chk节假日换休.BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub chk节假日换休_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk节假日换休_LostFocus()
    chk节假日换休.BackColor = fraBaseInfor.BackColor
End Sub

Private Sub chk临床排班_Click()
    mblnChange = True
End Sub

Private Sub chk临床排班_GotFocus()
    chk临床排班.BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub chk临床排班_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

  

Private Sub chk临床排班_LostFocus()
      chk临床排班.BackColor = fraBaseInfor.BackColor
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    
    If txt号码.Enabled And txt号码.Visible And txt号码.Text = "" Then
        txt号码.SetFocus
    ElseIf cbo号类.Enabled And cbo号类.Visible Then
        cbo号类.SetFocus
    Else
        If picBaseInfor.Enabled And picBaseInfor.Visible Then picBaseInfor.SetFocus
        zlCommFun.PressKey vbKeyTab
    End If
    
    lvwWorkTime.View = lvwReport
    lvwWorkTime.View = lvwList
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandle
    mblnFirst = True
    
    mstrNodeNo = ""
    mlngOldFeeItemID = 0
    mblnUpdateFeeItem = False
    
    RestoreWinState Me, App.ProductName
    Call DefCommandBars '初始化菜单
    Call InitPanel
    
    If CreatePublicPatient = False Then Unload Me: Exit Sub
    If InitData = False Then Unload Me: Exit Sub
    If LoadData = False Then Unload Me: Exit Sub
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobj所有分诊诊室集 = Nothing
    Set mobj所有合作单位 = Nothing
    Set mobjPubPatient = Nothing
    
    SaveWinState Me, App.ProductName
End Sub

Private Sub idkDoctor_ItemClick(index As Integer, objCard As zlIDKind.Card)

    mbln院内医生 = index = 1
    If mbln院内医生 Then
        idkDoctor.ToolTipText = "只能选院内建档医生"
    Else
        idkDoctor.ToolTipText = "除了可以选择院内医生外，还可以输入外援医生"
    End If
End Sub

Private Sub idkDoctor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
 


 
Private Sub lvwWorkTime_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If mblnNotCheck Then Exit Sub
    If Item.Checked Then
        If CheckWorkTimeSelValied(Item.Text, Item.SubItems(1), Item.SubItems(2)) = False Then Item.Checked = False: Exit Sub
        '重新加载数据
        Call ReLoadDetialData
    Else
        Call ReLoadDetialData
    End If
    
    mblnChange = True
End Sub

Private Sub lvwWorkTime_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        If IsCheckSelWorkTime = False Then
           ' If MsgBox("未设置缺省的上班时间段，你是否需要保存当前号源设置?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If SaveData = False Then Exit Sub
            Exit Sub
         End If
    End If
    zlCommFun.PressKey vbKeyTab
End Sub
Private Function IsCheckSelWorkTime() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在工作时间选择
    '入参:
    '返回:如果有，则返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2016-03-31 11:45:03
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ObjItem As ListItem
    On Error GoTo errHandle
    For Each ObjItem In lvwWorkTime.ListItems
        If ObjItem.Checked Then IsCheckSelWorkTime = True: Exit Function
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub optApplyAgeRange_Click(index As Integer)
    mblnChange = True
    txtAgeRange(0).Enabled = False: cboAgeUnit(0).Enabled = False
    txtAgeRange(1).Enabled = False: cboAgeUnit(1).Enabled = False
    txtAgeRange(2).Enabled = False: cboAgeUnit(2).Enabled = False
    txtAgeRange(3).Enabled = False: cboAgeUnit(3).Enabled = False
    
    If optApplyAgeRange(index).Value Then
        Select Case index
        Case 1
            txtAgeRange(0).Enabled = True: cboAgeUnit(0).Enabled = True
            txtAgeRange(1).Enabled = True: cboAgeUnit(1).Enabled = True
        Case 2
            txtAgeRange(2).Enabled = True: cboAgeUnit(2).Enabled = True
        Case 3
            txtAgeRange(3).Enabled = True: cboAgeUnit(3).Enabled = True
        End Select
    End If
End Sub

Private Sub optApplyAgeRange_GotFocus(index As Integer)
    optApplyAgeRange(index).BackColor = GCTRL_SELBACK_COLOR
    Select Case index
    Case 1
        lblAgeRange(0).BackColor = GCTRL_SELBACK_COLOR
    Case 2
        lblAgeRange(1).BackColor = GCTRL_SELBACK_COLOR
    Case 3
        lblAgeRange(2).BackColor = GCTRL_SELBACK_COLOR
    End Select
End Sub

Private Sub optApplyAgeRange_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optApplyAgeRange_LostFocus(index As Integer)
    optApplyAgeRange(index).BackColor = fraBaseInfor.BackColor
    Select Case index
    Case 1
        lblAgeRange(0).BackColor = fraBaseInfor.BackColor
    Case 2
        lblAgeRange(1).BackColor = fraBaseInfor.BackColor
    Case 3
        lblAgeRange(2).BackColor = fraBaseInfor.BackColor
    End Select
End Sub

Private Sub opt节假日_Click(index As Integer)
    mblnChange = True
End Sub

Private Sub opt节假日_GotFocus(index As Integer)
     opt节假日(index).BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub opt节假日_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub opt节假日_LostFocus(index As Integer)
    opt节假日(index).BackColor = fraBaseInfor.BackColor
End Sub

Private Sub opt排班方式_Click(index As Integer)
    mblnChange = True
    If index = 0 Then '固定排班
        If CheckExistsPlan(mlng号源Id) Then
            cbo收费项目.Enabled = False
            '如果收费项目被改变了则恢复
            If mlngOldFeeItemID <> 0 Then
                If cbo收费项目.ItemData(cbo收费项目.ListIndex) <> mlngOldFeeItemID Then
                    zlControl.CboLocate cbo收费项目, mlngOldFeeItemID, True
                End If
            End If
        End If
    Else '按月/周排班
        cbo收费项目.Enabled = True
    End If
End Sub

Private Sub opt排班方式_GotFocus(index As Integer)
     opt排班方式(index).BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub opt排班方式_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub opt排班方式_LostFocus(index As Integer)
     opt排班方式(index).BackColor = fraBaseInfor.BackColor
End Sub

Private Sub txtAgeRange_Change(index As Integer)
    mblnChange = True
End Sub

Private Sub txtAgeRange_GotFocus(index As Integer)
    zlControl.TxtSelAll txtAgeRange(index)
End Sub

Private Sub txtAgeRange_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cboAgeUnit(index).Visible = False And IsNumeric(txtAgeRange(index).Text) Then
            Call txtAgeRange_Validate(index, False)
            If cboAgeUnit(index).Visible And cboAgeUnit(index).Enabled Then cboAgeUnit(index).SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txtAgeRange(index).Text) And cboAgeUnit(index).Visible Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        '仅仅限制几个 指定的特殊的字符
        If InStr("~·！@#￥%……&*（）——-+=|、？、。，~`!#$%^&*()-_=+|\/?<>,/<>", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAgeRange_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtAgeRange(index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtAgeRange(index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtAgeRange_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtAgeRange(index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtAgeRange_Validate(index As Integer, Cancel As Boolean)
    Dim strBirth As String
    
    txtAgeRange(index).Text = Trim(txtAgeRange(index).Text)
    If Not IsNumeric(txtAgeRange(index).Text) And Trim(txtAgeRange(index).Text) <> "" Then
        cboAgeUnit(index).ListIndex = -1
        cboAgeUnit(index).Visible = False
        txtAgeRange(index).Width = 1200
    ElseIf cboAgeUnit(index).Visible = False Then
        cboAgeUnit(index).ListIndex = 0
        cboAgeUnit(index).Visible = True
        txtAgeRange(index).Width = 630
    End If
    
    If txtAgeRange(index).Visible And Trim(txtAgeRange(index).Text <> "") Then
        If mobjPubPatient Is Nothing Then Exit Sub
        If mobjPubPatient.CheckPatiAge(Trim(txtAgeRange(index).Text) & IIf(cboAgeUnit(index).Visible, cboAgeUnit(index).Text, "")) = False Then
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Sub txt出诊频次_Change()
    mblnChange = True
    CPDPages.诊疗频次 = Val(txt出诊频次.Text)
End Sub

Private Sub txt出诊频次_GotFocus()
    zlControl.TxtSelAll txt出诊频次
End Sub

Private Sub txt出诊频次_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt号码_Change()
    mblnChange = True
End Sub

Private Sub txt号码_GotFocus()
    zlControl.TxtSelAll txt号码
End Sub

Private Sub txt号码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt预约天数_Change()
    mblnChange = True
End Sub

Private Sub txt预约天数_GotFocus()
    zlControl.TxtSelAll txt预约天数
End Sub

Private Sub txt预约天数_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub
 

Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    
    Err = 0: On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    If cboDoctor.ListIndex <> -1 Or mbln院内医生 = False Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If Trim(cboDoctor.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If mrsDoctor Is Nothing Then Exit Sub
    
    If zlPersonSelect(Me, mlngModule, cboDoctor, mrsDoctor, Trim(cboDoctor.Text), True, "") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboDoctor_Validate(Cancel As Boolean)
    Err = 0: On Error GoTo errHandle
    If mbln院内医生 Then
        If cboDoctor.ListIndex < 0 Then cboDoctor.Text = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo科室_Click()
    Err = 0: On Error GoTo errHandle
    mblnCboClick = True
    If cbo科室.ListIndex = -1 Then Exit Sub
    If mlngPre科室ID = cbo科室.ItemData(cbo科室.ListIndex) Then Exit Sub
    mlngPre科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    If Not mrs科室 Is Nothing Then
        mrs科室.Filter = "ID=" & mlngPre科室ID
        If Not mrs科室.EOF Then
            If mstrNodeNo <> Nvl(mrs科室!站点) Then
                '站点发生改变
                mstrNodeNo = Nvl(mrs科室!站点)
                '站点改变，需要重新提取上班时间段
                Call LoadWorkTimes(mstrNodeNo, cbo号类.Text)
                CPDPages.LoadData New 出诊记录集, mobj所有分诊诊室集, mobj所有合作单位
            End If
        End If
        mrs科室.Filter = ""
    End If
    Call LoadDoctor
    
    '所有分诊诊室发生改变了
    Set mobj所有分诊诊室集 = GetVisitRoomsObjects(GetDoctorRooms(mlngPre科室ID))
    '重新加载数据
    Call ReLoadDetialData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadDoctor()
    Err = 0: On Error GoTo errHandle
    Set mrsDoctor = GetDoctor(Val(cbo科室.ItemData(cbo科室.ListIndex)), "")
    cboDoctor.Clear
    Do While Not mrsDoctor.EOF
        cboDoctor.AddItem mrsDoctor!姓名
        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!id
        mrsDoctor.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo科室_GotFocus()
    zlControl.TxtSelAll cbo科室
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Err = 0: On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo科室.Text = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If cbo科室.ListIndex >= 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        
    mblnCboClick = True
    If Select科室(Me, mlngModule, mrs科室, cbo科室, cbo科室.Text) = True Then
        mblnCboClick = False
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If cbo科室.Enabled And cbo科室.Visible Then cbo科室.SetFocus
    
    mblnCboClick = False
    zlControl.TxtSelAll cbo科室
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo科室_Validate(Cancel As Boolean)
    '如果在cbo的keypress事件中用了弹出列表的的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
    'cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
    If Not mblnCboClick Then cbo科室_Click
    If cbo科室.ListIndex < 0 Then cbo科室.Text = ""
    mblnCboClick = False
End Sub

Private Function GetAge(txtAge As TextBox, cbo年龄单位 As ComboBox) As String
    '获取年龄
    On Error GoTo errHandler
    If IsNumeric(Trim(txtAge.Text)) Then
        GetAge = Trim(txtAge.Text) & cbo年龄单位.Text
    Else
        GetAge = Trim(txtAge.Text)
    End If
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '入参:
    '返回:数据保存成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-03-23 10:54:20
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str诊室 As String, i  As Integer, j As Integer
    Dim strSQL As String, cllPro As New Collection
    Dim lngDoctor As Long, obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    Dim lng号源Id As Long, lng限制ID As Long, str临床出诊限制 As String
    Dim str诊室IDs As String, lng诊室ID As Long
    Dim cllNums As Collection, intNum As Integer, strTemp As String
    Dim cllControl As Collection, intControl As Integer
    Dim lngCount As Long
    
    If mbytFun = Fun_View Then Unload Me: Exit Function
    
    If IsValied() = False Then Exit Function
    
    Err = 0: On Error GoTo errHandler
    If cboDoctor.ListIndex <> -1 And mbln院内医生 Then lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
    Set cllPro = New Collection
    
    lng号源Id = mlng号源Id
    If mbytFun = Fun_Add Then lng号源Id = zlDatabase.GetNextId("临床出诊号源")
 
    'Zl_临床出诊号源_Modify(
    strSQL = "Zl_临床出诊号源_Modify("
    '  操作类型_In     Number,
    '  --0-新增，1-修改
    strSQL = strSQL & "" & IIf(mbytFun = Fun_Add, 0, 1) & ","
    '  Id_In           临床出诊号源.Id%Type,
    strSQL = strSQL & "" & lng号源Id & ","
    '  号类_In         临床出诊号源.号类%Type := Null,
    strSQL = strSQL & "'" & cbo号类.Text & "',"
    '  号码_In         临床出诊号源.号码%Type := Null,
    strSQL = strSQL & "'" & txt号码.Text & "',"
    '  科室id_In       临床出诊号源.科室id%Type := 0,
    strSQL = strSQL & "" & cbo科室.ItemData(cbo科室.ListIndex) & ","
    '  项目id_In       临床出诊号源.项目id%Type := 0,
    strSQL = strSQL & "" & cbo收费项目.ItemData(cbo收费项目.ListIndex) & ","
    '  医生id_In       临床出诊号源.医生id%Type := 0,
    strSQL = strSQL & "" & ZVal(lngDoctor) & ","
    '  医生姓名_In     临床出诊号源.医生姓名%Type := Null,
    strSQL = strSQL & "'" & cboDoctor.Text & "',"
    '  是否建病案_In   临床出诊号源.是否建病案%Type := 0,
    strSQL = strSQL & "" & chk病案.Value & ","
    '  预约天数_In     临床出诊号源.预约天数%Type := 0,
    strSQL = strSQL & "" & ZVal(Val(txt预约天数.Text)) & ","
    '  出诊频次_In     临床出诊号源.出诊频次%Type := 0,
    strSQL = strSQL & "" & Val(txt出诊频次.Text) & ","
    '  假日控制状态_In 临床出诊号源.假日控制状态%Type := 0,
    strSQL = strSQL & "" & GetSelectedIndex(opt节假日) & ","
    '  是否假日换休_In 临床出诊号源.是否假日换休%Type := 0,
    strSQL = strSQL & "" & chk节假日换休.Value & ","
    '  是否临床排班_In 临床出诊号源.是否临床排班%Type := 0,
    strSQL = strSQL & "" & chk临床排班.Value & ","
    '  排班方式_In     临床出诊号源.排班方式%Type := 0,
    strSQL = strSQL & "" & GetSelectedIndex(opt排班方式) & ","
    '  适用性别_In     临床出诊号源.适用性别%Type := Null,
    strSQL = strSQL & "'" & IIf(chkApplySex.Value = vbChecked, zlCommFun.GetNeedName(cboApplySex.Text), "") & "',"
    '  适用年龄段_In   临床出诊号源.适用年龄段%Type := Null --格式:开始年龄~终止年龄，用~分隔
    strTemp = ""
    Select Case GetSelectedIndex(optApplyAgeRange)
    Case 1
        strTemp = GetAge(txtAgeRange(0), cboAgeUnit(0)) & "~" & GetAge(txtAgeRange(1), cboAgeUnit(1))
    Case 2
        strTemp = GetAge(txtAgeRange(2), cboAgeUnit(2)) & "~"
    Case 3
        strTemp = "~" & GetAge(txtAgeRange(3), cboAgeUnit(3))
    End Select
    strSQL = strSQL & "'" & strTemp & "',"
    '  更新出诊表_In   Number := 0--操作类型_In=1时，号源收费项目改变后，是否同步调整未发布的按月/周安排的出诊表
    strSQL = strSQL & "" & IIf(mblnUpdateFeeItem, 1, 0) & ")"
    zlAddArray cllPro, strSQL
    
    
    Set obj出诊记录集 = CPDPages.Get出诊记录集
    If obj出诊记录集.Count = 0 Then
        '删除已有出诊号源限制
        strSQL = "Zl_临床出诊号源限制_Modify(Null, " & lng号源Id & ", Null, Null, Null, " & _
                "Null, Null, Null, Null, Null, Null, Null, Null, Null, -1)"
        zlAddArray cllPro, strSQL
    Else
        lngCount = 1
        For Each obj出诊记录 In obj出诊记录集
            str诊室IDs = GetRoomIDs(obj出诊记录.安排门诊诊室集)
            Call GetNumstoCollenct(obj出诊记录.号序信息集, cllNums)
            Call GetCtontroltoCollenct(obj出诊记录.合作单位控制集, cllControl)
        
            lng限制ID = zlDatabase.GetNextId("临床出诊号源限制")
            '插入号源限制
            '    Zl_临床出诊号源限制_Modify
            str临床出诊限制 = "Zl_临床出诊号源限制_Modify("
            '      Id_In           临床出诊号源限制.Id%Type,
            str临床出诊限制 = str临床出诊限制 & "" & lng限制ID & ","
            '      号源id_In       临床出诊号源限制.号源id%Type,
            str临床出诊限制 = str临床出诊限制 & "" & lng号源Id & ","
            '      上班时段_In     临床出诊号源限制.上班时段%Type,
            str临床出诊限制 = str临床出诊限制 & "'" & obj出诊记录.时间段 & "',"
            '      限号数_In       临床出诊号源限制.限号数%Type,
            str临床出诊限制 = str临床出诊限制 & "" & obj出诊记录.限号数 & ","
            '      限约数_In       临床出诊号源限制.限约数%Type,
            str临床出诊限制 = str临床出诊限制 & "" & obj出诊记录.限约数 & ","
            '      是否序号控制_In 临床出诊号源限制.是否序号控制%Type,
            str临床出诊限制 = str临床出诊限制 & "" & IIf(obj出诊记录.是否序号控制, 1, 0) & ","
            '      是否分时段_In   临床出诊号源限制.是否分时段%Type,
            str临床出诊限制 = str临床出诊限制 & "" & IIf(obj出诊记录.是否分时段, 1, 0) & ","
            '      预约控制_In     临床出诊号源限制.预约控制%Type,
            str临床出诊限制 = str临床出诊限制 & "" & obj出诊记录.预约控制 & ","
            '      是否独占_In     临床出诊号源限制.是否独占%Type,
            str临床出诊限制 = str临床出诊限制 & "" & IIf(obj出诊记录.是否独占, 1, 0) & ","
            '      分诊方式_In     临床出诊号源限制.分诊方式%Type,
            str临床出诊限制 = str临床出诊限制 & "" & obj出诊记录.分诊方式 & ","
            '      诊室id_In       临床出诊号源限制.诊室id%Type,
            lng诊室ID = Val(Split(str诊室IDs & ",,", ",")(0))
             
            str临床出诊限制 = str临床出诊限制 & "" & IIf(obj出诊记录.分诊方式 = 1 And lng诊室ID <> 0, lng诊室ID, "NULL") & ","
            strSQL = ""
            intNum = 1: intControl = 1
            Do While True
                strSQL = str临床出诊限制
                '      号源诊室_In     Varchar2 := Null,
                '      --格式:诊室id1,诊室id2,....
                strSQL = strSQL & "" & IIf(str诊室IDs = "", "Null", "'" & str诊室IDs & "'") & ","
                str诊室IDs = ""
                strTemp = "NULL"
                If intNum <= cllNums.Count Then
                    strTemp = "'" & cllNums(intNum) & "'"
                End If
                '      号源时段_In     Varchar2 := Null,
                '      --格式:序号,开始时间,终止时间,数量,是否预约|...
                strSQL = strSQL & strTemp & ","
                '      号源控制_In     Varchar2 := Null,
                '      --格式:类型,性质,名称,控制方式,序号,数量|
                strTemp = "NULL"
                If intControl <= cllControl.Count Then
                    strTemp = "'" & cllControl(intControl) & "'"
                End If
                strSQL = strSQL & strTemp & ","
                '      删除号源限制_In Number:=0 1-插入数据前，先删除号源限制,0-不删除数据，直接插入,-1-仅删除号源限制,不插入数据
                strSQL = strSQL & IIf(lngCount = 1 And intNum = 1 And intControl = 1, 1, 0) & ")"
                zlAddArray cllPro, strSQL
                If intNum >= cllNums.Count And intControl >= cllControl.Count Then Exit Do
                intNum = intNum + 1
                intControl = intControl + 1
            Loop
            lngCount = lngCount + 1
        Next
    End If
    
    Err = 0: On Error GoTo ErrRollback:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    
    Err = 0: On Error GoTo errHandler
    mblnOk = True: SaveData = True
    
    If mbytFun <> Fun_Add Then Unload Me: Exit Function
    mstrAddNewItem = txt号码.Text
    '清除界面信息
    Call LoadData
    If cbo号类.Enabled And cbo号类.Visible Then cbo号类.SetFocus
    Exit Function
ErrRollback:
    gcnOracle.RollbackTrans
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSelectedIndex(ByVal OptionButtons As Object) As Integer
    '获取单选按钮组的选中项的索引
    Dim i As Integer
    
    For i = OptionButtons.LBound To OptionButtons.UBound
        If OptionButtons(i).Value Then
            GetSelectedIndex = i: Exit For
        End If
    Next
End Function
  
Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据合法性检查
    '返回:数据合法返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2016-03-23 11:37:11
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, rsTemp As ADODB.Recordset
    Dim intCount As Integer, strSQL As String
    Dim lng科室ID As Long, lng项目id As Long
    Dim lng医生ID As Long, str医生 As String
    Dim strBirthBefore As String, strBirthAfter As String
    
    Err = 0: On Error GoTo errHandle
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    '完整性检查
    If Trim(txt号码.Text) = "" Then
        MsgBox "号码不能为空！", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(txt号码): Exit Function
    End If
    If cbo号类.ListIndex = -1 Then
        MsgBox "号类不能为空！", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(cbo号类): Exit Function
    End If
    If cbo科室.ListIndex = -1 Then
        MsgBox "科室不能为空！", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(cbo科室): Exit Function
    End If
    If cbo收费项目.ListIndex = -1 Then
        MsgBox "挂号项目不能为空！", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(cbo收费项目): Exit Function
    End If
    
    '适用年龄检查
    If mobjPubPatient Is Nothing Then Exit Function
    Select Case GetSelectedIndex(optApplyAgeRange)
    Case 1
        If mobjPubPatient.CheckPatiAge(GetAge(txtAgeRange(0), cboAgeUnit(0))) = False Then
            Call zlControl.ControlSetFocus(txtAgeRange(0)): Exit Function
        End If
        If mobjPubPatient.CheckPatiAge(GetAge(txtAgeRange(1), cboAgeUnit(1))) = False Then
            Call zlControl.ControlSetFocus(txtAgeRange(1)): Exit Function
        End If
        If mobjPubPatient.ReCalcBirthDay(GetAge(txtAgeRange(0), cboAgeUnit(0)), strBirthBefore) = False Then Exit Function
        If mobjPubPatient.ReCalcBirthDay(GetAge(txtAgeRange(1), cboAgeUnit(1)), strBirthAfter) = False Then Exit Function
        If DateDiff("s", strBirthBefore, strBirthAfter) >= 0 Then
            MsgBox "适用年龄段的最大年龄必须大于最小年龄！", vbInformation, gstrSysName
            Call zlControl.ControlSetFocus(txtAgeRange(1)): Exit Function
        End If
    Case 2
        If mobjPubPatient.CheckPatiAge(GetAge(txtAgeRange(2), cboAgeUnit(2))) = False Then
            Call zlControl.ControlSetFocus(txtAgeRange(2)): Exit Function
        End If
    Case 3
        If mobjPubPatient.CheckPatiAge(GetAge(txtAgeRange(3), cboAgeUnit(3))) = False Then
            Call zlControl.ControlSetFocus(txtAgeRange(3)): Exit Function
        End If
    End Select
    
    If mbln院内医生 Then
        If cboDoctor.ListIndex < 0 And cboDoctor.Text <> "" Then
            MsgBox "你选择的医生不存在，请重新输入医生！", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Call zlControl.ControlSetFocus(cboDoctor): Exit Function
        End If
    End If
    If mbytFun = Fun_Add Then
        If CheckExist(Trim(txt号码.Text)) Then
            MsgBox "号码 " & Trim(txt号码.Text) & " 已存在，请重新输入！", vbInformation, gstrSysName
            Call zlControl.ControlSetFocus(txt号码): Exit Function
        End If
    End If
    
    '检查同一科室，同一级别及同一医生不能存在多个号源
    lng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    lng项目id = cbo收费项目.ItemData(cbo收费项目.ListIndex)
    str医生 = cboDoctor.Text
    If Not mbln院内医生 Then
         lng医生ID = 0
    Else
        If cboDoctor.ListIndex >= 0 Then
            lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
        End If
    End If
    If Not mbln院内医生 And str医生 <> "" Then
        strSQL = "Select 号码  From 临床出诊号源 where 科室ID=[1] and 医生ID is null  and 医生姓名 =[3] and 项目ID=[4] and  nvl(是否删除,0)=0 and 号码<>[5]"
    ElseIf lng医生ID = 0 Then
        strSQL = "Select 号码  From 临床出诊号源 where 科室ID=[1] and 医生ID is null  and 医生姓名 is null  And 项目ID=[4] and  nvl(是否删除,0)=0 and 号码<>[5]"
    Else
        strSQL = "Select 号码  From 临床出诊号源 where 科室ID=[1] and 医生ID=[2]   And 项目ID=[4] and  nvl(是否删除,0)=0 and 号码<>[5]"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID, lng医生ID, str医生, lng项目id, CStr(Trim(txt号码.Text)))
    If Not rsTemp.EOF Then
        MsgBox cbo科室.Text & " " & IIf(str医生 = "", "", "的医生 " & str医生 & " ") & _
            "已经存在收费项目为 " & cbo收费项目.Text & " 的号源【" & Nvl(rsTemp!号码) & "】，" & _
            "您不能" & IIf(mbytFun = Fun_Add, "再增加此号源", "修改为此号源") & "！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '时间段检查
    If CPDPages.IsValied() = False Then
        Exit Function
    End If
    
    If (mbytFun = Fun_Update And mlngOldFeeItemID <> cbo收费项目.ItemData(cbo收费项目.ListIndex)) Then
        '收费项目改变了，检查是否存在未发布的出诊表
        If GetSelectedIndex(opt排班方式) <> 0 Then '按月/周排班
            If CheckExistsNotPublishPlan(mlng号源Id) Then
                mblnUpdateFeeItem = _
                    MsgBox("    您修改了号源的收费项目，同时该号源存在未发布的安排，" & _
                           "是否对这些未发布的安排的收费项目进行同步调整？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
            End If
        End If
    End If
    
    '检查是否设置了缺省的时间段
    Dim obj临床出诊记录集 As 出诊记录集
    Set obj临床出诊记录集 = CPDPages.Get出诊记录集
    If obj临床出诊记录集.Count = 0 Then
        If MsgBox("你还未设置缺省的工作时间，是否继续保存?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
           If lvwWorkTime.Enabled And lvwWorkTime.Visible Then lvwWorkTime.SetFocus
           Exit Function
        End If
    End If
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckExist(ByVal str号码 As String) As Boolean
    '检查号码是否已存在
    Dim rs号源 As ADODB.Recordset, strSQL As String
    
    Err = 0: On Error GoTo errHandle
    strSQL = "Select 1 From 临床出诊号源 Where 号码='" & str号码 & "'"
    Set rs号源 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    CheckExist = Not rs号源.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub picBaseInfor_Resize()
    Err = 0: On Error Resume Next
    With picBaseInfor
        shpBaseLine.Top = .ScaleTop
        shpBaseLine.Width = .ScaleWidth
        shpBaseLine.Left = .ScaleLeft
        shpBaseLine.Height = .ScaleHeight
        
        fraBaseInfor.Left = .ScaleLeft
        fraBaseInfor.Top = lblSourceTittle.Top + lblSourceTittle.Height + 100
        fraBaseInfor.Width = .ScaleWidth - fraBaseInfor.Left - 50
        fraBaseInfor.Height = .ScaleHeight - fraBaseInfor.Top - 50
    End With
End Sub
 
Private Sub picDetailedList_Resize()
    Err = 0: On Error Resume Next
    With picDetailedList
        CPDPages.Left = .ScaleLeft
        CPDPages.Top = .ScaleTop
        CPDPages.Width = .ScaleWidth - CPDPages.Left - CPDPages.Left * 2
        CPDPages.Height = .ScaleHeight - CPDPages.Top - CPDPages.Top * 2
    End With
End Sub

Private Sub picWorkTimeList_Resize()
    Err = 0: On Error Resume Next
    With picWorkTimeList
        shpWorkLine.Top = .ScaleTop
        shpWorkLine.Width = .ScaleWidth
        shpWorkLine.Left = .ScaleLeft
        shpWorkLine.Height = .ScaleHeight
        lvwWorkTime.Left = lblCalendbarTittle.Left
        lvwWorkTime.Top = lblCalendbarTittle.Top + lblCalendbarTittle.Height + 50
        lvwWorkTime.Width = .ScaleWidth - lvwWorkTime.Left - 50
        lvwWorkTime.Height = .ScaleHeight - lvwWorkTime.Top - 50
    End With
End Sub
Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case Pancel_Index.Pan_BaseInforList
        Item.Handle = picBaseInfor.Hwnd
    Case Pancel_Index.Pan_WorkTimeList
        Item.Handle = picWorkTimeList.Hwnd
    Case Pancel_Index.Pan_DetailList
        Item.Handle = picDetailedList.Hwnd
    End Select
End Sub
Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化Docking控件
    '编制:刘兴洪
    '日期:2016-01-08 14:34:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, sngHeight As Single
    Dim strReg As String
    Dim panThis As Pane, panLeft As Pane
    
    On Error GoTo Errhand
    dkpMain.SetCommandBars cbsThis
    sngWidth = picBaseInfor.Width / Screen.TwipsPerPixelX
    sngHeight = picBaseInfor.Height / Screen.TwipsPerPixelY
    Set panLeft = dkpMain.CreatePane(Pancel_Index.Pan_BaseInforList, sngWidth, sngHeight, DockTopOf, Nothing)
    panLeft.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panLeft.Title = "": panLeft.Tag = Pancel_Index.Pan_BaseInforList
    panLeft.Handle = picBaseInfor.Hwnd
    
    panLeft.MinTrackSize.Height = sngHeight
    panLeft.MaxTrackSize.Height = sngHeight
    panLeft.MaxTrackSize.Width = sngWidth
    panLeft.MinTrackSize.Width = sngWidth * 2 / 3
    
    
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_DetailList, sngWidth, 300, DockRightOf, panLeft)
    panThis.Title = ""
    panThis.Tag = Pancel_Index.Pan_DetailList
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picDetailedList.Hwnd
    
    Set panThis = dkpMain.CreatePane(Pancel_Index.Pan_WorkTimeList, sngWidth, 300, DockBottomOf, panLeft)
    panThis.Title = "上班时间"
    panThis.Tag = Pancel_Index.Pan_WorkTimeList
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picWorkTimeList.Hwnd
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    Call picBaseInfor_Resize
    'zlRestoreDockPanceToReg Me, dkpMan, "区域"
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function DefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '入参:
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-03-23 10:50:45
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo Errhand:
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    
    '菜单定义
    cbsThis.DeleteAll
    
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    
    With cbrToolBar.Controls

         Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存")
        cbrControl.flags = xtpFlagRightAlign
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出    ")
        cbrControl.flags = xtpFlagRightAlign
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        
    End With

    DefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCtontroltoCollenct(ByVal obj合作单位控制集 As 合作单位控制集, ByRef cllCtontrols As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取控制信息集给集合
    '入参:obj号序信息集-号序信息集
    '出参:cllCtontrols-返回号控制信息集
    '             每个项不得超过4000个字符,格式为:类型,性质,名称,控制方式,序号,数量|
    '返回:获取成功,返回true,否则返回Fasle
    '编制:刘兴洪
    '日期:2016-03-24 11:06:09
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String, obj合作单位控制 As 合作单位控制
    Dim strTemp As String, obj号序信息 As 号序信息
    Dim strNums As String
    On Error GoTo errHandle
    
    Set cllCtontrols = New Collection
    For Each obj合作单位控制 In obj合作单位控制集
        strTemp = obj合作单位控制.类型
        strTemp = strTemp & "," & 1  '目前只有预约
        strTemp = strTemp & "," & Replace(obj合作单位控制.合作单位名称, ",", "")
        strTemp = strTemp & "," & obj合作单位控制.预约控制方式
        strNums = ""
        For Each obj号序信息 In obj合作单位控制.号序信息集
            strNums = obj号序信息.序号 & ","
            strNums = strNums & obj号序信息.数量
            strNums = strTemp & "," & strNums
            
            If zlCommFun.ActualLen(strData & "|" & strNums) >= 4000 Then
                cllCtontrols.Add Mid(strData, 2)
                strData = ""
            End If
            
            strData = strData & "|" & strNums
            strNums = ""
        Next
    Next
    If strData <> "" Then cllCtontrols.Add Mid(strData, 2)
    GetCtontroltoCollenct = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetNumstoCollenct(ByVal obj号序信息集 As 号序信息集, ByRef cllNums As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取号序信息集给集
    '入参:obj号序信息集-号序信息集
    '出参:cllNums-返回号序信息集给集合
    '             每个项不得超过4000个字符,格式为:序号,开始时间,终止时间,数量,是否预约|...
    '返回:获取成功,返回true,否则返回Fasle
    '编制:刘兴洪
    '日期:2016-03-24 11:06:09
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String, obj号序信息 As 号序信息
    Dim strTemp As String
    Dim dtStart As Date, dtEndDate As Date
    On Error GoTo errHandle
    Set cllNums = New Collection
    For Each obj号序信息 In obj号序信息集
        strTemp = obj号序信息.序号
        strTemp = strTemp & "," & Format(obj号序信息.开始时间, "yyyy-mm-dd HH:MM:SS")
        strTemp = strTemp & "," & Format(obj号序信息.终止时间, "yyyy-mm-dd HH:MM:SS")
        strTemp = strTemp & "," & obj号序信息.数量
        strTemp = strTemp & "," & IIf(obj号序信息.是否预约, 1, 0)
        If zlCommFun.ActualLen(strData & "|" & strTemp) >= 4000 Then
            
            cllNums.Add Mid(strData, 2)
            strData = ""
        End If
        strData = strData & "|" & strTemp
    Next
    If strData <> "" Then
         cllNums.Add Mid(strData, 2)
    End If
    
    GetNumstoCollenct = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetRoomIDs(ByVal obj分诊诊室集 As 分诊诊室集) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取诊室ID,多个用逗号分隔
    '返回:返回诊室ID,多个用逗号分隔
    '编制:刘兴洪
    '日期:2016-03-24 10:48:06
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj分诊诊室 As 分诊诊室
    Dim strIDs As String
    On Error GoTo errHandle
    If obj分诊诊室集 Is Nothing Then Exit Function
    For Each obj分诊诊室 In obj分诊诊室集
        strIDs = strIDs & "," & obj分诊诊室.诊室ID
    Next
    If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    GetRoomIDs = strIDs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckWorkTimeSelValied(str时间段名 As String, _
    ByVal strStartTime As String, ByVal strEndTime As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查工作时间选择的数据合法性
    '入参:objItem-当前选中的接点
    '返回:数据合法，返回True,否则返回False
    '编制:刘兴洪
    '日期:2016-03-24 14:12:52
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsWorkTime As ADODB.Recordset, objListItem As ListItem
    Dim dtCurdate As Date, i As Long
    Dim dtStart As Date, dtEnd As Date
    Dim dtStartTemp As Date, dtEndTemp As Date
    
    On Error GoTo errHandle
    
    If LpadTime(strStartTime, strEndTime, dtStart, dtEnd) = False Then Exit Function    '格式出错
    
   
    With lvwWorkTime
        For i = 1 To .ListItems.Count
            Set objListItem = .ListItems(i)
            
            If objListItem.Checked And objListItem.Text <> str时间段名 Then
                If LpadTime(objListItem.SubItems(1), objListItem.SubItems(2), dtStartTemp, dtEndTemp) = False Then Exit Function    '格式出错
                If (dtStart >= dtStartTemp And dtStart <= dtEndTemp) Or (dtEnd >= dtStartTemp And dtEnd <= dtEndTemp) Then
                    '选择的时段不能有交差
             '       MsgBox "当前的上班时间(" & str时间段名 & ") 不能与已经选择的上班时间交叉(" & objListItem.Text & ")！", vbInformation + vbOKOnly, gstrSysName
             '       Exit Function
                End If
            End If
        Next
    End With
    CheckWorkTimeSelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function GetClinicRecord(ByVal obj出诊记录集 As 出诊记录集, ByVal str时间段 As String) As 出诊记录
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据时间段来获取对应的出诊记录集
    '入参:obj出诊记录集-出诊记录集
    '     str时间段-时间段
    '返回:出诊记录对象
    '编制:刘兴洪
    '日期:2016-03-24 15:37:50
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj出诊记录 As 出诊记录
    If obj出诊记录集 Is Nothing Then Exit Function
    
    On Error GoTo errHandle
    For Each obj出诊记录 In obj出诊记录集
        If obj出诊记录.时间段 = str时间段 Then
            Set GetClinicRecord = obj出诊记录.Clone: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetClinicRecordIndex(ByVal obj出诊记录集 As 出诊记录集, ByVal str时间段 As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据时间段来获取对应的出诊记录集的索引
    '入参:obj出诊记录集-出诊记录集
    '     str时间段-时间段
    '返回:出诊记录集的索引,未找到返回-1
    '编制:刘兴洪
    '日期:2016-03-24 15:37:50
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj出诊记录 As 出诊记录, i As Long, intIndex As Integer
    If obj出诊记录集 Is Nothing Then Exit Function
    
    On Error GoTo errHandle
    intIndex = -1
    For i = 1 To obj出诊记录集.Count
        If obj出诊记录集(i).时间段 = str时间段 Then
            intIndex = i: Exit For
        End If
    Next
    GetClinicRecordIndex = intIndex
    Exit Function
errHandle:
    intIndex = -1
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetBuildClinicRecord(ByVal str时间段 As String, _
    obj出诊记录集 As 出诊记录集) As 出诊记录
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据上班时间段来构建出诊记录对象
    '入参:str时间段-上班时间段名称
    '返回:返回出诊记录集对象
    '编制:刘兴洪
    '日期:2016-03-24 15:53:43
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj出诊记录 As New 出诊记录
    Dim obj上班时段 As 上班时段
    
    On Error GoTo errHandle
    Set obj上班时段 = GetWorkTimeRange(str时间段, mstrNodeNo, cbo号类.Text)
    With obj出诊记录
        .记录ID = 0
        .出诊日期 = Format(obj上班时段.开始时间, "yyyy-mm-dd")
        
        Set .安排门诊诊室集 = New 分诊诊室集
        .安排门诊诊室集.医生姓名 = cboDoctor.Text
        '缺省分诊诊室
        If obj出诊记录集.Count > 0 Then
            .安排门诊诊室集.分诊方式 = obj出诊记录集(1).安排门诊诊室集.分诊方式
            Set .安排门诊诊室集 = obj出诊记录集(1).安排门诊诊室集.Clone
        End If
        .分诊方式 = .安排门诊诊室集.分诊方式
        
        Set .号序信息集 = New 号序信息集
        .号序信息集.出诊频次 = Val(txt出诊频次.Text)
        Set .合作单位控制集 = New 合作单位控制集
        
        Set .上班时段 = obj上班时段
        .时间段 = str时间段
        .开始时间 = obj上班时段.开始时间
        .终止时间 = obj上班时段.结束时间
        
        .是否分时段 = 0
        .是否序号控制 = 0
        .是否独占 = 0
        .替诊医生 = ""
        .限号数 = 0
        .限约数 = 0
        .已挂数 = 0
        .已约数 = 0
        .预约控制 = 0
    End With
    Set GetBuildClinicRecord = obj出诊记录
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ReLoadDetialData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新加载明细数据
    '编制:刘兴洪
    '日期:2016-03-24 14:11:19
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    Dim ObjItem As ListItem, intIndex As Integer
    
    Err = 0: On Error GoTo errHandle
    Set obj出诊记录集 = CPDPages.Get出诊记录集
    If obj出诊记录集 Is Nothing Then Set obj出诊记录集 = New 出诊记录集
    With lvwWorkTime
        For Each ObjItem In .ListItems
            intIndex = GetClinicRecordIndex(obj出诊记录集, ObjItem.Tag)
            If ObjItem.Checked Then
                If intIndex = -1 Then
                    Set obj出诊记录 = GetBuildClinicRecord(ObjItem.Tag, obj出诊记录集)
                   obj出诊记录集.AddItem obj出诊记录, "K" & obj出诊记录.时间段
                End If
            ElseIf intIndex > 0 Then
                '存在，需要删除
                obj出诊记录集.Remove intIndex
            End If
        Next
    End With
    If obj出诊记录集.出诊日期 = "" Then
        obj出诊记录集.出诊日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    End If
    
    CPDPages.LoadData obj出诊记录集, mobj所有分诊诊室集, mobj所有合作单位
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CreatePublicPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建zlPublicPatient部件
    '返回:创建成功,返回True,否则返回False
    '编制:冉俊明
    '日期:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "病人信息公共部件（zlPublicPatient）创建失败！", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) = False Then
            MsgBox "病人信息公共部件（zlPublicPatient）初始化失败！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreatePublicPatient = True
End Function

Private Function GetDepartments(ByVal str性质 As String, _
    ByVal str服务对象 As String, _
    Optional ByVal bln仅操作员部门 As Boolean = False, _
    Optional ByVal blnCheck站点 As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定性质的部门列表
    '入参:str性质='临床','护理','中药房',...,允许为空
    '     str服务对象:以,分离:如1,3
    '     bln仅操作员部门-操作员的所属部门
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-10-12 09:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    
    str性质 = Replace(str性质, "'", "")
    If str性质 <> "" Then
        If InStr(1, str性质, ",") > 0 Then
            strSQL = " And Instr(','||[1]||',',','||B.工作性质||',')>0"
        Else
            strSQL = " And B.工作性质 = [1]"
        End If
    End If
    If bln仅操作员部门 Then strSQL = strSQL & "  And A.id=C.部门ID and C.人员id =[3]"
    
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称,A.简码,B.工作性质,B.服务对象,a.站点 " & _
        " From 部门表 A,部门性质说明 B " & IIf(bln仅操作员部门, ",部门人员 C", "") & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And Instr(',' || [2]|| ',',',' || B.服务对象 || ',')>0 " & strSQL & _
         IIf(blnCheck站点, " And Nvl(Nvl(a.站点,[5]),Nvl([4],'-')) = Nvl([4],'-')", "") & _
        " Order by A.编码"
    Set GetDepartments = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str性质, str服务对象, _
        UserInfo.id, gstrNodeNo, gVisitPlan_ModulePara.str号源维护站点)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

