VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "ZLIDKIND.OCX"
Begin VB.Form frmClinicSignalSourceEdit 
   Caption         =   "�����������"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   2  '��Ļ����
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
            Name            =   "����"
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
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ʱ���"
            Object.Width           =   9596
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��ʼʱ��"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��ֹʱ��"
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
         Caption         =   "�ϰ�ʱ��"
         BeginProperty Font 
            Name            =   "����"
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
         Name            =   "����"
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
            Caption         =   "���������"
            BeginProperty Font 
               Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "������"
               BeginProperty Font 
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "��"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "����"
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
         Begin VB.ComboBox cbo���� 
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.TextBox txt���� 
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.ComboBox cbo���� 
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Frame fra�ڼ��� 
            BackColor       =   &H00FFEFE3&
            Caption         =   "�ڼ��տ��Ʒ�ʽ"
            BeginProperty Font 
               Name            =   "����"
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
            Begin VB.OptionButton opt�ڼ��� 
               BackColor       =   &H00FFEFE3&
               Caption         =   "�ܽڼ������ÿ���"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.OptionButton opt�ڼ��� 
               BackColor       =   &H00FFEFE3&
               Caption         =   "���ϰ�"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.OptionButton opt�ڼ��� 
               BackColor       =   &H00FFEFE3&
               Caption         =   "����ԤԼ"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.OptionButton opt�ڼ��� 
               BackColor       =   &H00FFEFE3&
               Caption         =   "��ֹԤԼ"
               BeginProperty Font 
                  Name            =   "����"
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
         Begin VB.Frame fra�Ű෽ʽ 
            BackColor       =   &H00FFEFE3&
            Caption         =   "�Ű෽ʽ"
            BeginProperty Font 
               Name            =   "����"
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
            Begin VB.OptionButton opt�Ű෽ʽ 
               BackColor       =   &H00FFEFE3&
               Caption         =   "�̶��Ű�"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.OptionButton opt�Ű෽ʽ 
               BackColor       =   &H00FFEFE3&
               Caption         =   "���Ű�"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.OptionButton opt�Ű෽ʽ 
               BackColor       =   &H00FFEFE3&
               Caption         =   "���Ű�"
               BeginProperty Font 
                  Name            =   "����"
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
         Begin VB.TextBox txt����Ƶ�� 
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.TextBox txtԤԼ���� 
            BeginProperty Font 
               Name            =   "����"
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
               Name            =   "����"
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
         Begin VB.ComboBox cbo�շ���Ŀ 
            BeginProperty Font 
               Name            =   "����"
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
         Begin MSComCtl2.UpDown upd����Ƶ�� 
            Height          =   315
            Left            =   1245
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1620
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txt����Ƶ��"
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
         Begin MSComCtl2.UpDown updԤԼ���� 
            Height          =   315
            Left            =   3735
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1620
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtԤԼ����"
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
            IDKindStr       =   "��|Ժ��ҽ��|0|0|0|0|0||0|0|0;��|Ժ��ҽ��|0|0|0|0|0||0|0|0"
            CaptionAlignment=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   9
            FontName        =   "����"
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
                  Name            =   "����"
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
            Begin VB.CheckBox chk���� 
               BackColor       =   &H00FFEFE3&
               Caption         =   "�Һ�ʱ���뽨����"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.CheckBox chk�ٴ��Ű� 
               BackColor       =   &H00FFEFE3&
               Caption         =   "�����ٴ������Ű�"
               BeginProperty Font 
                  Name            =   "����"
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
            Begin VB.CheckBox chk�ڼ��ջ��� 
               BackColor       =   &H00FFEFE3&
               Caption         =   "���ýڼ��ջ��ݿ���"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "�����Ա�"
               BeginProperty Font 
                  Name            =   "����"
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
            Caption         =   "ҽ    ��"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    ��"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lblԤԼ������λ 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEFE3&
            Caption         =   "��ԤԼ        (��)"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl����Ƶ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����Ƶ��        (����/�˴�) "
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl�շ���Ŀ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    Ŀ"
            BeginProperty Font 
               Name            =   "����"
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
         Begin VB.Label lbl���� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    ��"
            BeginProperty Font 
               Name            =   "����"
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
         Caption         =   "��Դ������Ϣ����"
         BeginProperty Font 
            Name            =   "����"
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
Private mbytFun As G_Enum_Fun '0-�鿴,1-���,2-����,3-ɾ��
Private mlngModule As Long
Private mstrPrivs As String
Private mlng��ԴId As Long
Private mdtCurDate As Date
Private mblnOk As Boolean
Private mrsDoctor As ADODB.Recordset, mrs���� As ADODB.Recordset
Private mblnԺ��ҽ�� As Boolean '��ǰѡ�е��Ƿ�Ժ��ҽ��
Private mblnCboClick As Boolean     '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
Private Enum Pancel_Index
    Pan_BaseInforList = 1001
    Pan_WorkTimeList = 1002
    Pan_DetailList = 1004
End Enum
Private mblnNotCheck As Boolean
Private mobj���з������Ҽ� As �������Ҽ�
Private mobj���к�����λ As ������λ���Ƽ�
Private mlngPre����ID As Long
Private mstr���� As String
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mobjPubPatient As Object
Private mstrAddNewItem As String
Private mstrNodeNo As String '��ǰѡ����վ��,104620

Private mlngOldFeeItemID As Long '��¼ԭʼ�շ���ĿID�������ж��Ƿ�����շ���Ŀ
Private mblnUpdateFeeItem As Boolean '�Ƿ�ͬ������δ�����ĳ��ﰲ�ŵ��շ���Ŀ

Public Function ShowMe(frmParent As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal bytFun As G_Enum_Fun, Optional ByVal lng��ԴId As Long, _
    Optional ByRef strAddNewItem As String) As Boolean
    '�������
    '���Σ�
    '   strAddNewItem:������Դ����
    mlngModule = lngModule: mstrPrivs = strPrivs
    mbytFun = bytFun: mlng��ԴId = lng��ԴId
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
    mlngPre����ID = 0
    If Val(zlDatabase.GetPara("ֻ����ѡԺ��ҽ��", glngSys, mlngModule, "0")) = 1 Then
        idkDoctor.IDkindStr = "ҽ��|ҽ��|0|0|0|0|0||0|0|0"
        idkDoctor.ToolTipText = "ֻ��ѡԺ�ڽ���ҽ��"
    Else
        idkDoctor.IDkindStr = "��|Ժ��ҽ��|0|0|0|0|0||0|0|0;��|Ժ��ҽ��|0|0|0|0|0||0|0|0"
        idkDoctor.ToolTipText = "���˿���ѡ��Ժ��ҽ���⣬������������Ԯҽ��"
    End If
    
    Set mobj���з������Ҽ� = GetVisitRoomsObjects(GetDoctorRooms(0))
    Set mobj���к�����λ = GetUnitsObjects(GetUnitAll())
    
    '�ϰ�ʱ��
    If mbytFun = Fun_Update Or mbytFun = Fun_Add Then
        mblnNotCheck = True
        Call LoadWorkTimes(mstrNodeNo, cbo����.Text)
        mblnNotCheck = False
    End If
    
    '�Ա�
    strSQL = "Select ����, ����, ����, ȱʡ��־ From �Ա� Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboApplySex.Clear
    Do While Not rsTemp.EOF
        cboApplySex.AddItem Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
        If Val(Nvl(rsTemp!ȱʡ��־, 0)) = 1 Then cboApplySex.ListIndex = cboApplySex.NewIndex
        rsTemp.MoveNext
    Loop
    If cboApplySex.ListIndex < 0 And cboApplySex.ListCount > 0 Then cboApplySex.ListIndex = 0
    
    '���䵥λ
    For i = 0 To cboAgeUnit.UBound
        cboAgeUnit(i).AddItem "��"
        cboAgeUnit(i).AddItem "��"
        cboAgeUnit(i).AddItem "��"
        cboAgeUnit(i).ListIndex = 0
    Next
    
    If mbytFun = Fun_View Or mbytFun = Fun_Delete Then InitData = True: Exit Function
    
    '����
    strSQL = "Select ����,����,���� From ���� Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo����.Clear
    Do While Not rsTemp.EOF
        cbo����.AddItem Nvl(rsTemp!����)
        rsTemp.MoveNext
    Loop
    If cbo����.ListIndex < 0 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
    '��Ŀ
    strSQL = "Select ID,���� From �շ���ĿĿ¼ " & _
        " Where ���='1' And (Sysdate Between ����ʱ�� And ����ʱ�� Or ����ʱ��<Sysdate And ����ʱ�� Is Null)" & _
        " And (վ��='" & gstrNodeNo & "' Or վ�� is Null) " & _
        " Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п��õĹҺ���Ŀ��Ϣ�����ȵ��Һ���Ŀ�����г�ʼ��", vbInformation, gstrSysName
        Exit Function
    End If
    cbo�շ���Ŀ.Clear
    Do While Not rsTemp.EOF
        cbo�շ���Ŀ.AddItem rsTemp!����
        cbo�շ���Ŀ.ItemData(cbo�շ���Ŀ.NewIndex) = Val(Nvl(rsTemp!id))
        rsTemp.MoveNext
    Loop
    '����
    Set mrs���� = GetDepartments("'�ٴ�'", "1,3", zlStr.IsHavePrivs(mstrPrivs, "���п���") = False)
    If mrs����.RecordCount = 0 Then
        MsgBox "�㲻�߱����õ��ٴ�������Ϣ�����ȵ����Ź����н������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    cbo����.Clear
    Do While Not mrs����.EOF
        cbo����.AddItem mrs����!����
        cbo����.ItemData(cbo����.NewIndex) = Val(Nvl(mrs����!id))
        mrs����.MoveNext
    Loop
    
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub LoadWorkTimes(ByVal strվ�� As String, ByVal str���� As String)
    '����ѡ����ද̬�����ϰ�ʱ��
    Dim rsWorkTime As ADODB.Recordset
    Dim strWorkTims As String, objListItem As ListItem
    
    Set rsWorkTime = GetWorkTimes(strվ��, str����)
    rsWorkTime.Sort = "վ�� Desc,���� Desc"
    With lvwWorkTime
        .ListItems.Clear
        If rsWorkTime.RecordCount > 0 Then rsWorkTime.MoveFirst
        strWorkTims = ""
        Do While Not rsWorkTime.EOF
            If InStr(1, strWorkTims & ",", "," & Nvl(rsWorkTime!ʱ���) & ",") = 0 Then
                Set objListItem = .ListItems.Add(, "K" & Nvl(rsWorkTime!ʱ���), Nvl(rsWorkTime!ʱ���) & _
                    "(" & Format(Nvl(rsWorkTime!��ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsWorkTime!��ֹʱ��), "hh:mm") & ")")
                objListItem.Tag = Nvl(rsWorkTime!ʱ���)
                objListItem.SubItems(1) = Nvl(rsWorkTime!��ʼʱ��)
                objListItem.SubItems(2) = Nvl(rsWorkTime!��ֹʱ��)
                strWorkTims = strWorkTims & "," & Nvl(rsWorkTime!ʱ���)
            End If
            rsWorkTime.MoveNext
        Loop
    End With
End Sub

Private Function CheckExistsPlan(ByVal lng��ԴId As Long) As Boolean
    '��鵱ǰ��Դ�Ƿ���ڰ�������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandle
    If lng��ԴId = 0 Then Exit Function
    strSQL = "Select 1 From �ٴ����ﰲ��" & vbNewLine & _
            " Where ��ԴID = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ��Դ�Ƿ���ڰ�������", lng��ԴId)
    CheckExistsPlan = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckExistsNotPublishPlan(ByVal lng��ԴId As Long) As Boolean
    '��鵱ǰ��Դ�Ƿ����δ��������/�ܰ�������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandle
    If lng��ԴId = 0 Then Exit Function
    strSQL = "Select 1" & vbNewLine & _
            " From �ٴ����ﰲ�� A, �ٴ������ B" & vbNewLine & _
            " Where a.����id = b.Id And a.��Դid = [1] And Nvl(b.�Ű෽ʽ, 0) In (1, 2)" & vbNewLine & _
            "       And b.����ʱ�� Is Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ��Դ�Ƿ���ڰ�������", lng��ԴId)
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
    '����:��������
    '����:���سɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-03-23 11:54:49
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rs��Դ As ADODB.Recordset
    Dim rs���� As ADODB.Recordset
    Dim i As Long, j As Long, blnExitFor As Boolean
    Dim obj�����¼�� As �����¼��
    Dim ObjItem  As ListItem
    Dim obj�����¼ As �����¼, objListItem As ListItem
    Dim lngԤԼ���� As Long, strTemp As String
    
    Err = 0: On Error GoTo errHandle
    Me.Caption = Choose(mbytFun + 1, "�鿴", "����", "�޸�", "ɾ��") & "��Դ"
    
    lngԤԼ���� = zlDatabase.GetPara(66, glngSys, , 15)
    If mbytFun = Fun_Add Then
        '�Զ��������
        txt����.Text = GetMaxLocalCode("�ٴ������Դ", "����")
        txtԤԼ����.Text = lngԤԼ����
        
        Set obj�����¼�� = New �����¼��
        obj�����¼��.�������� = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        cboDoctor.ListIndex = -1
        mblnNotCheck = True
        
        If mblnOk Then  '�ڶ��α���ʱ�������
            For Each ObjItem In lvwWorkTime.ListItems
                ObjItem.Checked = False
            Next
        End If
        mblnNotCheck = False
        Call CPDPages.LoadData(obj�����¼��, mobj���з������Ҽ�, mobj���к�����λ)
        mblnChange = False
        LoadData = True: Exit Function
    End If
    
    strSQL = "" & _
            " Select A.ID, A.����, A.����, A.����id, A.��Ŀid,A.ҽ��ID, A.ҽ������ As ҽ��, A.ԤԼ����, A.����Ƶ��," & vbNewLine & _
            "        Nvl(A.�Ƿ񽨲���, 0) As �Ƿ񽨲���," & vbNewLine & _
            "        Nvl(A.���տ���״̬, 0) As ���տ���״̬," & vbNewLine & _
            "        Nvl(A.�Ƿ��ٴ��Ű�, 0) As �Ƿ��ٴ��Ű�,Nvl(�Ű෽ʽ, 0) As �Ű෽ʽ," & vbNewLine & _
            "        Nvl(A.�Ƿ���ջ���, 0) As �Ƿ���ջ���,A.����ʱ��,nvl(A.�Ƿ�ɾ��,0) as �Ƿ�ɾ��, " & _
            "        B.���� as ��������,C.���� as �շ���Ŀ����, a.�����Ա�, a.���������, b.վ��" & vbNewLine & _
            " From �ٴ������Դ A,���ű� B,�շ���ĿĿ¼ C" & vbNewLine & _
            " Where A.ID = [1] and a.����ID=B.id and A.��ĿID=C.ID "
    Set rs��Դ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ԴId)
    If rs��Դ.EOF Then
        MsgBox "��ǰ��Դ�����ڣ������ѱ�����ɾ����������ˢ�����ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    txt����.Text = Nvl(rs��Դ!����)
    mstrNodeNo = Nvl(rs��Դ!վ��)
    mblnNotCheck = True
    mstr���� = Nvl(rs��Դ!����)
    zlControl.CboSetText cbo����, Nvl(rs��Դ!����)
    If cbo����.ListIndex = -1 Then cbo����.AddItem Nvl(rs��Դ!����): cbo����.ListIndex = cbo����.NewIndex
    mblnNotCheck = False
    
    zlControl.CboLocate cbo����, Nvl(rs��Դ!����ID), True
    If cbo����.ListIndex = -1 Then
        cbo����.AddItem Nvl(rs��Դ!��������): cbo����.ItemData(cbo����.NewIndex) = Val(Nvl(rs��Դ!����ID))
        cbo����.ListIndex = cbo����.NewIndex
    End If
    
    mlngOldFeeItemID = Val(Nvl(rs��Դ!��ĿID))
    zlControl.CboLocate cbo�շ���Ŀ, Nvl(rs��Դ!��ĿID), True
    If cbo�շ���Ŀ.ListIndex = -1 Then
        cbo�շ���Ŀ.AddItem Nvl(rs��Դ!�շ���Ŀ����): cbo�շ���Ŀ.ItemData(cbo�շ���Ŀ.NewIndex) = Val(Nvl(rs��Դ!��ĿID))
        cbo�շ���Ŀ.ListIndex = cbo�շ���Ŀ.NewIndex
    End If
    If Nvl(rs��Դ!ҽ��) <> "" Then
        If Val(Nvl(rs��Դ!ҽ��ID)) = 0 Then
            cboDoctor.AddItem Nvl(rs��Դ!ҽ��): cboDoctor.ListIndex = cboDoctor.NewIndex
            idkDoctor.IDKind = idkDoctor.ListCount
        Else
            zlControl.CboLocate cboDoctor, Val(Nvl(rs��Դ!ҽ��ID)), True
            If cboDoctor.ListIndex = -1 Then
                cboDoctor.AddItem Nvl(rs��Դ!ҽ��): cboDoctor.ItemData(cboDoctor.NewIndex) = Val(Nvl(rs��Դ!ҽ��ID))
                cboDoctor.ListIndex = cboDoctor.NewIndex
            End If
        End If
    Else
        cboDoctor.ListIndex = -1
    End If
    txt����Ƶ��.Text = Val(Nvl(rs��Դ!����Ƶ��))
    txtԤԼ����.Text = IIf(Val(Nvl(rs��Դ!ԤԼ����)) = 0, lngԤԼ����, Val(Nvl(rs��Դ!ԤԼ����)))
    
    chkApplySex.Value = IIf(Nvl(rs��Դ!�����Ա�) = "", vbUnchecked, vbChecked)
    If Nvl(rs��Դ!�����Ա�) <> "" Then
        zlControl.CboLocate cboApplySex, Nvl(rs��Դ!�����Ա�)
        If cboApplySex.ListIndex = -1 Then cboApplySex.AddItem Nvl(rs��Դ!�����Ա�): cboApplySex.ListIndex = cboApplySex.NewIndex
    End If
    
    chk�ٴ��Ű�.Value = Val(Nvl(rs��Դ!�Ƿ��ٴ��Ű�))
    chk�ڼ��ջ���.Value = Val(Nvl(rs��Դ!�Ƿ���ջ���))
    chk����.Value = Val(Nvl(rs��Դ!�Ƿ񽨲���))
    
    opt�Ű෽ʽ(Val(Nvl(rs��Դ!�Ű෽ʽ))).Value = True
    opt�ڼ���(Val(Nvl(rs��Դ!���տ���״̬))).Value = True
    
    strTemp = Nvl(rs��Դ!���������) '��ʽ:��ʼ����~��ֹ���䣬��~�ָ�
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
     
    Set obj�����¼�� = GetClinicRecordFromSignalSource(mlng��ԴId)
    If mbytFun = Fun_Update Then Call LoadWorkTimes(mstrNodeNo, Nvl(rs��Դ!����))
    With lvwWorkTime
        mblnNotCheck = True
        For Each obj�����¼ In obj�����¼��
            Set objListItem = Nothing
            Err = 0: On Error Resume Next
            Set objListItem = .ListItems("K" & obj�����¼.ʱ���)
            If Err <> 0 Then
                Set objListItem = .ListItems.Add(, "K" & obj�����¼.ʱ���, obj�����¼.ʱ��� & _
                    "(" & Format(obj�����¼.��ʼʱ��, "hh:mm") & "-" & Format(obj�����¼.��ֹʱ��, "hh:mm") & ")")
                objListItem.Tag = obj�����¼.�ϰ�ʱ��.ʱ���
                objListItem.SubItems(1) = obj�����¼.�ϰ�ʱ��.��ʼʱ��
                objListItem.SubItems(2) = obj�����¼.�ϰ�ʱ��.����ʱ��
            End If
            Err = 0: On Error GoTo 0
            If Not objListItem Is Nothing Then objListItem.Checked = True
       Next
       mblnNotCheck = False
    End With
    
    Call CPDPages.LoadData(obj�����¼��, IIf(mbytFun = Fun_Update, mobj���з������Ҽ�, Nothing), mobj���к�����λ, True)
    
    '���Ʊ༭״̬
    If mbytFun = Fun_Delete Or mbytFun = Fun_View Then
        Call SetEnabled(Me.Controls, False)
        CPDPages.EditMode(-1) = ED_RegistPlan_View
    Else
        txt����.Enabled = False
        
        '��ǰ��Դ���ڰ�������ʱ������������ҡ�ҽ������Ŀ
        If CheckExistsPlan(mlng��ԴId) Then
            cbo����.Enabled = False
            idkDoctor.Enabled = False
            cboDoctor.Enabled = False
            '����ǹ̶����ţ�����������շ���Ŀ��
            cbo�շ���Ŀ.Enabled = IIf(opt�Ű෽ʽ(0).Value, False, True)
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
    CPDPages.ҽ������ = cboDoctor.Text
End Sub

Private Sub cbo����_Change()
    mblnChange = True
End Sub

Private Sub cbo����_Click()
    On Error GoTo errHandle

    If mstr���� = cbo����.Text Or mblnNotCheck Then Exit Sub
    mstr���� = cbo����.Text
    mblnChange = True
    
    '����ı䣬��Ҫ������ȡ�ϰ�ʱ���
    Call LoadWorkTimes(mstrNodeNo, cbo����.Text)
    '���¼�������
    CPDPages.LoadData New �����¼��, mobj���з������Ҽ�, mobj���к�����λ
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�շ���Ŀ_Click()
    mblnChange = True
End Sub

Private Sub cbo�շ���Ŀ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboDoctor_GotFocus()
    zlControl.TxtSelAll cboDoctor
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim obj�����¼�� As �����¼��
    
    Select Case Control.id
    Case conMenu_Edit_Save    '��������
        If SaveData = False Then Exit Sub
    Case conMenu_File_Exit   '�˳�
        Set obj�����¼�� = CPDPages.Get�����¼��
        
        If Not obj�����¼�� Is Nothing Then
            mblnChange = mblnChange Or obj�����¼��.�Ƿ��޸�
        End If
        
        If mblnChange Then
             If MsgBox("��Դ��Ϣ�Ѿ������ı䣬������δ���棬���Ƿ����Ҫ�˳���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        Set obj�����¼�� = Nothing
        Unload Me: Exit Sub
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case conMenu_Edit_Save    '��������
        Control.Visible = mbytFun <> Fun_View
    Case conMenu_File_Exit   '�˳�
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

Private Sub chk����_Click()
    mblnChange = True
End Sub

Private Sub chk����_GotFocus()
    chk����.BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

 

Private Sub chk����_LostFocus()
    chk����.BackColor = fraBaseInfor.BackColor
End Sub

Private Sub chk�ڼ��ջ���_Click()
    mblnChange = True
End Sub

Private Sub chk�ڼ��ջ���_GotFocus()
       chk�ڼ��ջ���.BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub chk�ڼ��ջ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�ڼ��ջ���_LostFocus()
    chk�ڼ��ջ���.BackColor = fraBaseInfor.BackColor
End Sub

Private Sub chk�ٴ��Ű�_Click()
    mblnChange = True
End Sub

Private Sub chk�ٴ��Ű�_GotFocus()
    chk�ٴ��Ű�.BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub chk�ٴ��Ű�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

  

Private Sub chk�ٴ��Ű�_LostFocus()
      chk�ٴ��Ű�.BackColor = fraBaseInfor.BackColor
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    
    If txt����.Enabled And txt����.Visible And txt����.Text = "" Then
        txt����.SetFocus
    ElseIf cbo����.Enabled And cbo����.Visible Then
        cbo����.SetFocus
    Else
        If picBaseInfor.Enabled And picBaseInfor.Visible Then picBaseInfor.SetFocus
        zlCommFun.PressKey vbKeyTab
    End If
    
    lvwWorkTime.View = lvwReport
    lvwWorkTime.View = lvwList
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandle
    mblnFirst = True
    
    mstrNodeNo = ""
    mlngOldFeeItemID = 0
    mblnUpdateFeeItem = False
    
    RestoreWinState Me, App.ProductName
    Call DefCommandBars '��ʼ���˵�
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
    Set mobj���з������Ҽ� = Nothing
    Set mobj���к�����λ = Nothing
    Set mobjPubPatient = Nothing
    
    SaveWinState Me, App.ProductName
End Sub

Private Sub idkDoctor_ItemClick(index As Integer, objCard As zlIDKind.Card)

    mblnԺ��ҽ�� = index = 1
    If mblnԺ��ҽ�� Then
        idkDoctor.ToolTipText = "ֻ��ѡԺ�ڽ���ҽ��"
    Else
        idkDoctor.ToolTipText = "���˿���ѡ��Ժ��ҽ���⣬������������Ԯҽ��"
    End If
End Sub

Private Sub idkDoctor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
 


 
Private Sub lvwWorkTime_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If mblnNotCheck Then Exit Sub
    If Item.Checked Then
        If CheckWorkTimeSelValied(Item.Text, Item.SubItems(1), Item.SubItems(2)) = False Then Item.Checked = False: Exit Sub
        '���¼�������
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
           ' If MsgBox("δ����ȱʡ���ϰ�ʱ��Σ����Ƿ���Ҫ���浱ǰ��Դ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If SaveData = False Then Exit Sub
            Exit Sub
         End If
    End If
    zlCommFun.PressKey vbKeyTab
End Sub
Private Function IsCheckSelWorkTime() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ���ڹ���ʱ��ѡ��
    '���:
    '����:����У��򷵻�true,���򷵻�Flase
    '����:���˺�
    '����:2016-03-31 11:45:03
    '˵����
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

Private Sub opt�ڼ���_Click(index As Integer)
    mblnChange = True
End Sub

Private Sub opt�ڼ���_GotFocus(index As Integer)
     opt�ڼ���(index).BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub opt�ڼ���_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub opt�ڼ���_LostFocus(index As Integer)
    opt�ڼ���(index).BackColor = fraBaseInfor.BackColor
End Sub

Private Sub opt�Ű෽ʽ_Click(index As Integer)
    mblnChange = True
    If index = 0 Then '�̶��Ű�
        If CheckExistsPlan(mlng��ԴId) Then
            cbo�շ���Ŀ.Enabled = False
            '����շ���Ŀ���ı�����ָ�
            If mlngOldFeeItemID <> 0 Then
                If cbo�շ���Ŀ.ItemData(cbo�շ���Ŀ.ListIndex) <> mlngOldFeeItemID Then
                    zlControl.CboLocate cbo�շ���Ŀ, mlngOldFeeItemID, True
                End If
            End If
        End If
    Else '����/���Ű�
        cbo�շ���Ŀ.Enabled = True
    End If
End Sub

Private Sub opt�Ű෽ʽ_GotFocus(index As Integer)
     opt�Ű෽ʽ(index).BackColor = GCTRL_SELBACK_COLOR
End Sub

Private Sub opt�Ű෽ʽ_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub opt�Ű෽ʽ_LostFocus(index As Integer)
     opt�Ű෽ʽ(index).BackColor = fraBaseInfor.BackColor
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
        '�������Ƽ��� ָ����������ַ�
        If InStr("~����@#��%����&*��������-+=|����������~`!#$%^&*()-_=+|\/?<>,/<>", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
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

Private Sub txt����Ƶ��_Change()
    mblnChange = True
    CPDPages.����Ƶ�� = Val(txt����Ƶ��.Text)
End Sub

Private Sub txt����Ƶ��_GotFocus()
    zlControl.TxtSelAll txt����Ƶ��
End Sub

Private Sub txt����Ƶ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt����_Change()
    mblnChange = True
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtԤԼ����_Change()
    mblnChange = True
End Sub

Private Sub txtԤԼ����_GotFocus()
    zlControl.TxtSelAll txtԤԼ����
End Sub

Private Sub txtԤԼ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub
 

Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    
    Err = 0: On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    If cboDoctor.ListIndex <> -1 Or mblnԺ��ҽ�� = False Then
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
    If mblnԺ��ҽ�� Then
        If cboDoctor.ListIndex < 0 Then cboDoctor.Text = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo����_Click()
    Err = 0: On Error GoTo errHandle
    mblnCboClick = True
    If cbo����.ListIndex = -1 Then Exit Sub
    If mlngPre����ID = cbo����.ItemData(cbo����.ListIndex) Then Exit Sub
    mlngPre����ID = cbo����.ItemData(cbo����.ListIndex)
    If Not mrs���� Is Nothing Then
        mrs����.Filter = "ID=" & mlngPre����ID
        If Not mrs����.EOF Then
            If mstrNodeNo <> Nvl(mrs����!վ��) Then
                'վ�㷢���ı�
                mstrNodeNo = Nvl(mrs����!վ��)
                'վ��ı䣬��Ҫ������ȡ�ϰ�ʱ���
                Call LoadWorkTimes(mstrNodeNo, cbo����.Text)
                CPDPages.LoadData New �����¼��, mobj���з������Ҽ�, mobj���к�����λ
            End If
        End If
        mrs����.Filter = ""
    End If
    Call LoadDoctor
    
    '���з������ҷ����ı���
    Set mobj���з������Ҽ� = GetVisitRoomsObjects(GetDoctorRooms(mlngPre����ID))
    '���¼�������
    Call ReLoadDetialData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadDoctor()
    Err = 0: On Error GoTo errHandle
    Set mrsDoctor = GetDoctor(Val(cbo����.ItemData(cbo����.ListIndex)), "")
    cboDoctor.Clear
    Do While Not mrsDoctor.EOF
        cboDoctor.AddItem mrsDoctor!����
        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!id
        mrsDoctor.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo����_GotFocus()
    zlControl.TxtSelAll cbo����
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Err = 0: On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo����.Text = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If cbo����.ListIndex >= 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        
    mblnCboClick = True
    If Select����(Me, mlngModule, mrs����, cbo����, cbo����.Text) = True Then
        mblnCboClick = False
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If cbo����.Enabled And cbo����.Visible Then cbo����.SetFocus
    
    mblnCboClick = False
    zlControl.TxtSelAll cbo����
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
    '�����cbo��keypress�¼������˵����б�ĵ�API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
    'cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
    If Not mblnCboClick Then cbo����_Click
    If cbo����.ListIndex < 0 Then cbo����.Text = ""
    mblnCboClick = False
End Sub

Private Function GetAge(txtAge As TextBox, cbo���䵥λ As ComboBox) As String
    '��ȡ����
    On Error GoTo errHandler
    If IsNumeric(Trim(txtAge.Text)) Then
        GetAge = Trim(txtAge.Text) & cbo���䵥λ.Text
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
    '����:���ݱ���
    '���:
    '����:���ݱ���ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-03-23 10:54:20
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String, i  As Integer, j As Integer
    Dim strSQL As String, cllPro As New Collection
    Dim lngDoctor As Long, obj�����¼�� As �����¼��, obj�����¼ As �����¼
    Dim lng��ԴId As Long, lng����ID As Long, str�ٴ��������� As String
    Dim str����IDs As String, lng����ID As Long
    Dim cllNums As Collection, intNum As Integer, strTemp As String
    Dim cllControl As Collection, intControl As Integer
    Dim lngCount As Long
    
    If mbytFun = Fun_View Then Unload Me: Exit Function
    
    If IsValied() = False Then Exit Function
    
    Err = 0: On Error GoTo errHandler
    If cboDoctor.ListIndex <> -1 And mblnԺ��ҽ�� Then lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
    Set cllPro = New Collection
    
    lng��ԴId = mlng��ԴId
    If mbytFun = Fun_Add Then lng��ԴId = zlDatabase.GetNextId("�ٴ������Դ")
 
    'Zl_�ٴ������Դ_Modify(
    strSQL = "Zl_�ٴ������Դ_Modify("
    '  ��������_In     Number,
    '  --0-������1-�޸�
    strSQL = strSQL & "" & IIf(mbytFun = Fun_Add, 0, 1) & ","
    '  Id_In           �ٴ������Դ.Id%Type,
    strSQL = strSQL & "" & lng��ԴId & ","
    '  ����_In         �ٴ������Դ.����%Type := Null,
    strSQL = strSQL & "'" & cbo����.Text & "',"
    '  ����_In         �ٴ������Դ.����%Type := Null,
    strSQL = strSQL & "'" & txt����.Text & "',"
    '  ����id_In       �ٴ������Դ.����id%Type := 0,
    strSQL = strSQL & "" & cbo����.ItemData(cbo����.ListIndex) & ","
    '  ��Ŀid_In       �ٴ������Դ.��Ŀid%Type := 0,
    strSQL = strSQL & "" & cbo�շ���Ŀ.ItemData(cbo�շ���Ŀ.ListIndex) & ","
    '  ҽ��id_In       �ٴ������Դ.ҽ��id%Type := 0,
    strSQL = strSQL & "" & ZVal(lngDoctor) & ","
    '  ҽ������_In     �ٴ������Դ.ҽ������%Type := Null,
    strSQL = strSQL & "'" & cboDoctor.Text & "',"
    '  �Ƿ񽨲���_In   �ٴ������Դ.�Ƿ񽨲���%Type := 0,
    strSQL = strSQL & "" & chk����.Value & ","
    '  ԤԼ����_In     �ٴ������Դ.ԤԼ����%Type := 0,
    strSQL = strSQL & "" & ZVal(Val(txtԤԼ����.Text)) & ","
    '  ����Ƶ��_In     �ٴ������Դ.����Ƶ��%Type := 0,
    strSQL = strSQL & "" & Val(txt����Ƶ��.Text) & ","
    '  ���տ���״̬_In �ٴ������Դ.���տ���״̬%Type := 0,
    strSQL = strSQL & "" & GetSelectedIndex(opt�ڼ���) & ","
    '  �Ƿ���ջ���_In �ٴ������Դ.�Ƿ���ջ���%Type := 0,
    strSQL = strSQL & "" & chk�ڼ��ջ���.Value & ","
    '  �Ƿ��ٴ��Ű�_In �ٴ������Դ.�Ƿ��ٴ��Ű�%Type := 0,
    strSQL = strSQL & "" & chk�ٴ��Ű�.Value & ","
    '  �Ű෽ʽ_In     �ٴ������Դ.�Ű෽ʽ%Type := 0,
    strSQL = strSQL & "" & GetSelectedIndex(opt�Ű෽ʽ) & ","
    '  �����Ա�_In     �ٴ������Դ.�����Ա�%Type := Null,
    strSQL = strSQL & "'" & IIf(chkApplySex.Value = vbChecked, zlCommFun.GetNeedName(cboApplySex.Text), "") & "',"
    '  ���������_In   �ٴ������Դ.���������%Type := Null --��ʽ:��ʼ����~��ֹ���䣬��~�ָ�
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
    '  ���³����_In   Number := 0--��������_In=1ʱ����Դ�շ���Ŀ�ı���Ƿ�ͬ������δ�����İ���/�ܰ��ŵĳ����
    strSQL = strSQL & "" & IIf(mblnUpdateFeeItem, 1, 0) & ")"
    zlAddArray cllPro, strSQL
    
    
    Set obj�����¼�� = CPDPages.Get�����¼��
    If obj�����¼��.Count = 0 Then
        'ɾ�����г����Դ����
        strSQL = "Zl_�ٴ������Դ����_Modify(Null, " & lng��ԴId & ", Null, Null, Null, " & _
                "Null, Null, Null, Null, Null, Null, Null, Null, Null, -1)"
        zlAddArray cllPro, strSQL
    Else
        lngCount = 1
        For Each obj�����¼ In obj�����¼��
            str����IDs = GetRoomIDs(obj�����¼.�����������Ҽ�)
            Call GetNumstoCollenct(obj�����¼.������Ϣ��, cllNums)
            Call GetCtontroltoCollenct(obj�����¼.������λ���Ƽ�, cllControl)
        
            lng����ID = zlDatabase.GetNextId("�ٴ������Դ����")
            '�����Դ����
            '    Zl_�ٴ������Դ����_Modify
            str�ٴ��������� = "Zl_�ٴ������Դ����_Modify("
            '      Id_In           �ٴ������Դ����.Id%Type,
            str�ٴ��������� = str�ٴ��������� & "" & lng����ID & ","
            '      ��Դid_In       �ٴ������Դ����.��Դid%Type,
            str�ٴ��������� = str�ٴ��������� & "" & lng��ԴId & ","
            '      �ϰ�ʱ��_In     �ٴ������Դ����.�ϰ�ʱ��%Type,
            str�ٴ��������� = str�ٴ��������� & "'" & obj�����¼.ʱ��� & "',"
            '      �޺���_In       �ٴ������Դ����.�޺���%Type,
            str�ٴ��������� = str�ٴ��������� & "" & obj�����¼.�޺��� & ","
            '      ��Լ��_In       �ٴ������Դ����.��Լ��%Type,
            str�ٴ��������� = str�ٴ��������� & "" & obj�����¼.��Լ�� & ","
            '      �Ƿ���ſ���_In �ٴ������Դ����.�Ƿ���ſ���%Type,
            str�ٴ��������� = str�ٴ��������� & "" & IIf(obj�����¼.�Ƿ���ſ���, 1, 0) & ","
            '      �Ƿ��ʱ��_In   �ٴ������Դ����.�Ƿ��ʱ��%Type,
            str�ٴ��������� = str�ٴ��������� & "" & IIf(obj�����¼.�Ƿ��ʱ��, 1, 0) & ","
            '      ԤԼ����_In     �ٴ������Դ����.ԤԼ����%Type,
            str�ٴ��������� = str�ٴ��������� & "" & obj�����¼.ԤԼ���� & ","
            '      �Ƿ��ռ_In     �ٴ������Դ����.�Ƿ��ռ%Type,
            str�ٴ��������� = str�ٴ��������� & "" & IIf(obj�����¼.�Ƿ��ռ, 1, 0) & ","
            '      ���﷽ʽ_In     �ٴ������Դ����.���﷽ʽ%Type,
            str�ٴ��������� = str�ٴ��������� & "" & obj�����¼.���﷽ʽ & ","
            '      ����id_In       �ٴ������Դ����.����id%Type,
            lng����ID = Val(Split(str����IDs & ",,", ",")(0))
             
            str�ٴ��������� = str�ٴ��������� & "" & IIf(obj�����¼.���﷽ʽ = 1 And lng����ID <> 0, lng����ID, "NULL") & ","
            strSQL = ""
            intNum = 1: intControl = 1
            Do While True
                strSQL = str�ٴ���������
                '      ��Դ����_In     Varchar2 := Null,
                '      --��ʽ:����id1,����id2,....
                strSQL = strSQL & "" & IIf(str����IDs = "", "Null", "'" & str����IDs & "'") & ","
                str����IDs = ""
                strTemp = "NULL"
                If intNum <= cllNums.Count Then
                    strTemp = "'" & cllNums(intNum) & "'"
                End If
                '      ��Դʱ��_In     Varchar2 := Null,
                '      --��ʽ:���,��ʼʱ��,��ֹʱ��,����,�Ƿ�ԤԼ|...
                strSQL = strSQL & strTemp & ","
                '      ��Դ����_In     Varchar2 := Null,
                '      --��ʽ:����,����,����,���Ʒ�ʽ,���,����|
                strTemp = "NULL"
                If intControl <= cllControl.Count Then
                    strTemp = "'" & cllControl(intControl) & "'"
                End If
                strSQL = strSQL & strTemp & ","
                '      ɾ����Դ����_In Number:=0 1-��������ǰ����ɾ����Դ����,0-��ɾ�����ݣ�ֱ�Ӳ���,-1-��ɾ����Դ����,����������
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
    mstrAddNewItem = txt����.Text
    '���������Ϣ
    Call LoadData
    If cbo����.Enabled And cbo����.Visible Then cbo����.SetFocus
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
    '��ȡ��ѡ��ť���ѡ���������
    Dim i As Integer
    
    For i = OptionButtons.LBound To OptionButtons.UBound
        If OptionButtons(i).Value Then
            GetSelectedIndex = i: Exit For
        End If
    Next
End Function
  
Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݺϷ��Լ��
    '����:���ݺϷ�����true,���򷵻�Flase
    '����:���˺�
    '����:2016-03-23 11:37:11
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, rsTemp As ADODB.Recordset
    Dim intCount As Integer, strSQL As String
    Dim lng����ID As Long, lng��Ŀid As Long
    Dim lngҽ��ID As Long, strҽ�� As String
    Dim strBirthBefore As String, strBirthAfter As String
    
    Err = 0: On Error GoTo errHandle
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    '�����Լ��
    If Trim(txt����.Text) = "" Then
        MsgBox "���벻��Ϊ�գ�", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(txt����): Exit Function
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "���಻��Ϊ�գ�", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(cbo����): Exit Function
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "���Ҳ���Ϊ�գ�", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(cbo����): Exit Function
    End If
    If cbo�շ���Ŀ.ListIndex = -1 Then
        MsgBox "�Һ���Ŀ����Ϊ�գ�", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(cbo�շ���Ŀ): Exit Function
    End If
    
    '����������
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
            MsgBox "��������ε����������������С���䣡", vbInformation, gstrSysName
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
    
    If mblnԺ��ҽ�� Then
        If cboDoctor.ListIndex < 0 And cboDoctor.Text <> "" Then
            MsgBox "��ѡ���ҽ�������ڣ�����������ҽ����", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
            Call zlControl.ControlSetFocus(cboDoctor): Exit Function
        End If
    End If
    If mbytFun = Fun_Add Then
        If CheckExist(Trim(txt����.Text)) Then
            MsgBox "���� " & Trim(txt����.Text) & " �Ѵ��ڣ����������룡", vbInformation, gstrSysName
            Call zlControl.ControlSetFocus(txt����): Exit Function
        End If
    End If
    
    '���ͬһ���ң�ͬһ����ͬһҽ�����ܴ��ڶ����Դ
    lng����ID = cbo����.ItemData(cbo����.ListIndex)
    lng��Ŀid = cbo�շ���Ŀ.ItemData(cbo�շ���Ŀ.ListIndex)
    strҽ�� = cboDoctor.Text
    If Not mblnԺ��ҽ�� Then
         lngҽ��ID = 0
    Else
        If cboDoctor.ListIndex >= 0 Then
            lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
        End If
    End If
    If Not mblnԺ��ҽ�� And strҽ�� <> "" Then
        strSQL = "Select ����  From �ٴ������Դ where ����ID=[1] and ҽ��ID is null  and ҽ������ =[3] and ��ĿID=[4] and  nvl(�Ƿ�ɾ��,0)=0 and ����<>[5]"
    ElseIf lngҽ��ID = 0 Then
        strSQL = "Select ����  From �ٴ������Դ where ����ID=[1] and ҽ��ID is null  and ҽ������ is null  And ��ĿID=[4] and  nvl(�Ƿ�ɾ��,0)=0 and ����<>[5]"
    Else
        strSQL = "Select ����  From �ٴ������Դ where ����ID=[1] and ҽ��ID=[2]   And ��ĿID=[4] and  nvl(�Ƿ�ɾ��,0)=0 and ����<>[5]"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lngҽ��ID, strҽ��, lng��Ŀid, CStr(Trim(txt����.Text)))
    If Not rsTemp.EOF Then
        MsgBox cbo����.Text & " " & IIf(strҽ�� = "", "", "��ҽ�� " & strҽ�� & " ") & _
            "�Ѿ������շ���ĿΪ " & cbo�շ���Ŀ.Text & " �ĺ�Դ��" & Nvl(rsTemp!����) & "����" & _
            "������" & IIf(mbytFun = Fun_Add, "�����Ӵ˺�Դ", "�޸�Ϊ�˺�Դ") & "��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ʱ��μ��
    If CPDPages.IsValied() = False Then
        Exit Function
    End If
    
    If (mbytFun = Fun_Update And mlngOldFeeItemID <> cbo�շ���Ŀ.ItemData(cbo�շ���Ŀ.ListIndex)) Then
        '�շ���Ŀ�ı��ˣ�����Ƿ����δ�����ĳ����
        If GetSelectedIndex(opt�Ű෽ʽ) <> 0 Then '����/���Ű�
            If CheckExistsNotPublishPlan(mlng��ԴId) Then
                mblnUpdateFeeItem = _
                    MsgBox("    ���޸��˺�Դ���շ���Ŀ��ͬʱ�ú�Դ����δ�����İ��ţ�" & _
                           "�Ƿ����Щδ�����İ��ŵ��շ���Ŀ����ͬ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
            End If
        End If
    End If
    
    '����Ƿ�������ȱʡ��ʱ���
    Dim obj�ٴ������¼�� As �����¼��
    Set obj�ٴ������¼�� = CPDPages.Get�����¼��
    If obj�ٴ������¼��.Count = 0 Then
        If MsgBox("�㻹δ����ȱʡ�Ĺ���ʱ�䣬�Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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

Private Function CheckExist(ByVal str���� As String) As Boolean
    '�������Ƿ��Ѵ���
    Dim rs��Դ As ADODB.Recordset, strSQL As String
    
    Err = 0: On Error GoTo errHandle
    strSQL = "Select 1 From �ٴ������Դ Where ����='" & str���� & "'"
    Set rs��Դ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    CheckExist = Not rs��Դ.EOF
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
    '����:��ʼ��Docking�ؼ�
    '����:���˺�
    '����:2016-01-08 14:34:56
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
    panThis.Title = "�ϰ�ʱ��"
    panThis.Tag = Pancel_Index.Pan_WorkTimeList
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picWorkTimeList.Hwnd
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    Call picBaseInfor_Resize
    'zlRestoreDockPanceToReg Me, dkpMan, "����"
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function DefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '���:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-03-23 10:50:45
    '˵����
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    
    '�˵�����
    cbsThis.DeleteAll
    
    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    
    With cbrToolBar.Controls

         Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        cbrControl.flags = xtpFlagRightAlign
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�    ")
        cbrControl.flags = xtpFlagRightAlign
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    '�����
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

Private Function GetCtontroltoCollenct(ByVal obj������λ���Ƽ� As ������λ���Ƽ�, ByRef cllCtontrols As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ��������
    '���:obj������Ϣ��-������Ϣ��
    '����:cllCtontrols-���غſ�����Ϣ��
    '             ÿ����ó���4000���ַ�,��ʽΪ:����,����,����,���Ʒ�ʽ,���,����|
    '����:��ȡ�ɹ�,����true,���򷵻�Fasle
    '����:���˺�
    '����:2016-03-24 11:06:09
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String, obj������λ���� As ������λ����
    Dim strTemp As String, obj������Ϣ As ������Ϣ
    Dim strNums As String
    On Error GoTo errHandle
    
    Set cllCtontrols = New Collection
    For Each obj������λ���� In obj������λ���Ƽ�
        strTemp = obj������λ����.����
        strTemp = strTemp & "," & 1  'Ŀǰֻ��ԤԼ
        strTemp = strTemp & "," & Replace(obj������λ����.������λ����, ",", "")
        strTemp = strTemp & "," & obj������λ����.ԤԼ���Ʒ�ʽ
        strNums = ""
        For Each obj������Ϣ In obj������λ����.������Ϣ��
            strNums = obj������Ϣ.��� & ","
            strNums = strNums & obj������Ϣ.����
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
Private Function GetNumstoCollenct(ByVal obj������Ϣ�� As ������Ϣ��, ByRef cllNums As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ������
    '���:obj������Ϣ��-������Ϣ��
    '����:cllNums-���غ�����Ϣ��������
    '             ÿ����ó���4000���ַ�,��ʽΪ:���,��ʼʱ��,��ֹʱ��,����,�Ƿ�ԤԼ|...
    '����:��ȡ�ɹ�,����true,���򷵻�Fasle
    '����:���˺�
    '����:2016-03-24 11:06:09
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String, obj������Ϣ As ������Ϣ
    Dim strTemp As String
    Dim dtStart As Date, dtEndDate As Date
    On Error GoTo errHandle
    Set cllNums = New Collection
    For Each obj������Ϣ In obj������Ϣ��
        strTemp = obj������Ϣ.���
        strTemp = strTemp & "," & Format(obj������Ϣ.��ʼʱ��, "yyyy-mm-dd HH:MM:SS")
        strTemp = strTemp & "," & Format(obj������Ϣ.��ֹʱ��, "yyyy-mm-dd HH:MM:SS")
        strTemp = strTemp & "," & obj������Ϣ.����
        strTemp = strTemp & "," & IIf(obj������Ϣ.�Ƿ�ԤԼ, 1, 0)
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
Private Function GetRoomIDs(ByVal obj�������Ҽ� As �������Ҽ�) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ID,����ö��ŷָ�
    '����:��������ID,����ö��ŷָ�
    '����:���˺�
    '����:2016-03-24 10:48:06
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj�������� As ��������
    Dim strIDs As String
    On Error GoTo errHandle
    If obj�������Ҽ� Is Nothing Then Exit Function
    For Each obj�������� In obj�������Ҽ�
        strIDs = strIDs & "," & obj��������.����ID
    Next
    If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    GetRoomIDs = strIDs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckWorkTimeSelValied(strʱ����� As String, _
    ByVal strStartTime As String, ByVal strEndTime As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鹤��ʱ��ѡ������ݺϷ���
    '���:objItem-��ǰѡ�еĽӵ�
    '����:���ݺϷ�������True,���򷵻�False
    '����:���˺�
    '����:2016-03-24 14:12:52
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsWorkTime As ADODB.Recordset, objListItem As ListItem
    Dim dtCurdate As Date, i As Long
    Dim dtStart As Date, dtEnd As Date
    Dim dtStartTemp As Date, dtEndTemp As Date
    
    On Error GoTo errHandle
    
    If LpadTime(strStartTime, strEndTime, dtStart, dtEnd) = False Then Exit Function    '��ʽ����
    
   
    With lvwWorkTime
        For i = 1 To .ListItems.Count
            Set objListItem = .ListItems(i)
            
            If objListItem.Checked And objListItem.Text <> strʱ����� Then
                If LpadTime(objListItem.SubItems(1), objListItem.SubItems(2), dtStartTemp, dtEndTemp) = False Then Exit Function    '��ʽ����
                If (dtStart >= dtStartTemp And dtStart <= dtEndTemp) Or (dtEnd >= dtStartTemp And dtEnd <= dtEndTemp) Then
                    'ѡ���ʱ�β����н���
             '       MsgBox "��ǰ���ϰ�ʱ��(" & strʱ����� & ") �������Ѿ�ѡ����ϰ�ʱ�佻��(" & objListItem.Text & ")��", vbInformation + vbOKOnly, gstrSysName
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
Private Function GetClinicRecord(ByVal obj�����¼�� As �����¼��, ByVal strʱ��� As String) As �����¼
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ�������ȡ��Ӧ�ĳ����¼��
    '���:obj�����¼��-�����¼��
    '     strʱ���-ʱ���
    '����:�����¼����
    '����:���˺�
    '����:2016-03-24 15:37:50
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj�����¼ As �����¼
    If obj�����¼�� Is Nothing Then Exit Function
    
    On Error GoTo errHandle
    For Each obj�����¼ In obj�����¼��
        If obj�����¼.ʱ��� = strʱ��� Then
            Set GetClinicRecord = obj�����¼.Clone: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetClinicRecordIndex(ByVal obj�����¼�� As �����¼��, ByVal strʱ��� As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ�������ȡ��Ӧ�ĳ����¼��������
    '���:obj�����¼��-�����¼��
    '     strʱ���-ʱ���
    '����:�����¼��������,δ�ҵ�����-1
    '����:���˺�
    '����:2016-03-24 15:37:50
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj�����¼ As �����¼, i As Long, intIndex As Integer
    If obj�����¼�� Is Nothing Then Exit Function
    
    On Error GoTo errHandle
    intIndex = -1
    For i = 1 To obj�����¼��.Count
        If obj�����¼��(i).ʱ��� = strʱ��� Then
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

Private Function GetBuildClinicRecord(ByVal strʱ��� As String, _
    obj�����¼�� As �����¼��) As �����¼
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ϰ�ʱ��������������¼����
    '���:strʱ���-�ϰ�ʱ�������
    '����:���س����¼������
    '����:���˺�
    '����:2016-03-24 15:53:43
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj�����¼ As New �����¼
    Dim obj�ϰ�ʱ�� As �ϰ�ʱ��
    
    On Error GoTo errHandle
    Set obj�ϰ�ʱ�� = GetWorkTimeRange(strʱ���, mstrNodeNo, cbo����.Text)
    With obj�����¼
        .��¼ID = 0
        .�������� = Format(obj�ϰ�ʱ��.��ʼʱ��, "yyyy-mm-dd")
        
        Set .�����������Ҽ� = New �������Ҽ�
        .�����������Ҽ�.ҽ������ = cboDoctor.Text
        'ȱʡ��������
        If obj�����¼��.Count > 0 Then
            .�����������Ҽ�.���﷽ʽ = obj�����¼��(1).�����������Ҽ�.���﷽ʽ
            Set .�����������Ҽ� = obj�����¼��(1).�����������Ҽ�.Clone
        End If
        .���﷽ʽ = .�����������Ҽ�.���﷽ʽ
        
        Set .������Ϣ�� = New ������Ϣ��
        .������Ϣ��.����Ƶ�� = Val(txt����Ƶ��.Text)
        Set .������λ���Ƽ� = New ������λ���Ƽ�
        
        Set .�ϰ�ʱ�� = obj�ϰ�ʱ��
        .ʱ��� = strʱ���
        .��ʼʱ�� = obj�ϰ�ʱ��.��ʼʱ��
        .��ֹʱ�� = obj�ϰ�ʱ��.����ʱ��
        
        .�Ƿ��ʱ�� = 0
        .�Ƿ���ſ��� = 0
        .�Ƿ��ռ = 0
        .����ҽ�� = ""
        .�޺��� = 0
        .��Լ�� = 0
        .�ѹ��� = 0
        .��Լ�� = 0
        .ԤԼ���� = 0
    End With
    Set GetBuildClinicRecord = obj�����¼
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ReLoadDetialData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼�����ϸ����
    '����:���˺�
    '����:2016-03-24 14:11:19
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj�����¼�� As �����¼��, obj�����¼ As �����¼
    Dim ObjItem As ListItem, intIndex As Integer
    
    Err = 0: On Error GoTo errHandle
    Set obj�����¼�� = CPDPages.Get�����¼��
    If obj�����¼�� Is Nothing Then Set obj�����¼�� = New �����¼��
    With lvwWorkTime
        For Each ObjItem In .ListItems
            intIndex = GetClinicRecordIndex(obj�����¼��, ObjItem.Tag)
            If ObjItem.Checked Then
                If intIndex = -1 Then
                    Set obj�����¼ = GetBuildClinicRecord(ObjItem.Tag, obj�����¼��)
                   obj�����¼��.AddItem obj�����¼, "K" & obj�����¼.ʱ���
                End If
            ElseIf intIndex > 0 Then
                '���ڣ���Ҫɾ��
                obj�����¼��.Remove intIndex
            End If
        Next
    End With
    If obj�����¼��.�������� = "" Then
        obj�����¼��.�������� = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    End If
    
    CPDPages.LoadData obj�����¼��, mobj���з������Ҽ�, mobj���к�����λ
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CreatePublicPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����zlPublicPatient����
    '����:�����ɹ�,����True,���򷵻�False
    '����:Ƚ����
    '����:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "������Ϣ����������zlPublicPatient������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) = False Then
            MsgBox "������Ϣ����������zlPublicPatient����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreatePublicPatient = True
End Function

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
        " Select Distinct A.ID,A.����,A.����,A.����,B.��������,B.�������,a.վ�� " & _
        " From ���ű� A,��������˵�� B " & IIf(bln������Ա����, ",������Ա C", "") & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And Instr(',' || [2]|| ',',',' || B.������� || ',')>0 " & strSQL & _
         IIf(blnCheckվ��, " And Nvl(Nvl(a.վ��,[5]),Nvl([4],'-')) = Nvl([4],'-')", "") & _
        " Order by A.����"
    Set GetDepartments = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����, str�������, _
        UserInfo.id, gstrNodeNo, gVisitPlan_ModulePara.str��Դά��վ��)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

