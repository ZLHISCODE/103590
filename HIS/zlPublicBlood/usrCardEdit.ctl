VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl usrCardEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   11610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13080
   ScaleHeight     =   11610
   ScaleWidth      =   13080
   Begin VB.PictureBox Pic4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10020
      Left            =   240
      ScaleHeight     =   10020
      ScaleWidth      =   9975
      TabIndex        =   62
      Top             =   1200
      Width           =   9975
      Begin VB.ComboBox cbo2 
         Height          =   300
         Left            =   10440
         TabIndex        =   99
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.VScrollBar VS1 
         Height          =   615
         Left            =   9960
         TabIndex        =   102
         Top             =   4440
         Width           =   255
      End
      Begin VB.HScrollBar HS1 
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   9840
         Width           =   615
      End
      Begin VB.ComboBox Cbo1 
         Height          =   300
         Left            =   10440
         TabIndex        =   100
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9900
         Left            =   90
         ScaleHeight     =   9870
         ScaleWidth      =   9825
         TabIndex        =   63
         Top             =   210
         Width           =   9855
         Begin VB.PictureBox PicHave 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2790
            ScaleHeight     =   255
            ScaleWidth      =   870
            TabIndex        =   116
            Top             =   2505
            Width           =   870
            Begin VB.ComboBox cboHave 
               Height          =   300
               Left            =   -30
               TabIndex        =   19
               Top             =   -45
               Width           =   945
            End
         End
         Begin VB.PictureBox PicDW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   8955
            ScaleHeight     =   240
            ScaleWidth      =   855
            TabIndex        =   113
            Top             =   1755
            Width           =   855
            Begin VB.ComboBox CboDW 
               Height          =   300
               Left            =   -30
               TabIndex        =   114
               Text            =   "ml"
               Top             =   -45
               Width           =   930
            End
         End
         Begin RichTextLib.RichTextBox TXT33 
            Height          =   1935
            Left            =   2685
            TabIndex        =   49
            Top             =   7455
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   3413
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"usrCardEdit.ctx":0000
         End
         Begin RichTextLib.RichTextBox TXT32 
            Height          =   855
            Left            =   3000
            TabIndex        =   48
            Top             =   6390
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   1508
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"usrCardEdit.ctx":009D
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
         Begin VB.PictureBox pic3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9560
            Picture         =   "usrCardEdit.ctx":013A
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   53
            Top             =   1320
            Width           =   255
         End
         Begin VB.PictureBox Pic2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   -2520
            ScaleHeight     =   255
            ScaleWidth      =   285
            TabIndex        =   67
            Top             =   -600
            Width           =   320
         End
         Begin VB.TextBox TXT21 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   7665
            TabIndex        =   11
            Top             =   1350
            Width           =   1830
         End
         Begin VB.TextBox TXT11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   0
            Left            =   2520
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox TXT11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   1
            Left            =   5280
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox TXT11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   2
            Left            =   7920
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox TXT11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   3
            Left            =   2520
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox TXT11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   4
            Left            =   5280
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox TXT11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   5
            Left            =   7920
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox TXT11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   6
            Left            =   2520
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   840
            Width           =   6675
         End
         Begin VB.TextBox TXT21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   0
            Left            =   4560
            MaxLength       =   3
            TabIndex        =   16
            Top             =   2130
            Width           =   615
         End
         Begin VB.TextBox TXT21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   1
            Left            =   8085
            MaxLength       =   4
            TabIndex        =   17
            Top             =   2130
            Width           =   495
         End
         Begin VB.TextBox TXT21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   2
            Left            =   8925
            MaxLength       =   4
            TabIndex        =   18
            Top             =   2130
            Width           =   615
         End
         Begin VB.TextBox TXT21 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   3
            Left            =   2760
            TabIndex        =   12
            Top             =   1755
            Width           =   3720
         End
         Begin VB.TextBox TXT21 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   4
            Left            =   7665
            MaxLength       =   6
            TabIndex        =   13
            Top             =   1770
            Width           =   1215
         End
         Begin VB.OptionButton Opt32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   7920
            TabIndex        =   23
            Top             =   2925
            Width           =   735
         End
         Begin VB.OptionButton Opt32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   8715
            TabIndex        =   24
            Top             =   2925
            Width           =   855
         End
         Begin VB.TextBox TXT31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   22
            Top             =   2955
            Width           =   735
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   25
            Top             =   3330
            Width           =   855
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3840
            TabIndex        =   26
            Top             =   3330
            Width           =   975
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   27
            Top             =   3330
            Width           =   1215
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���β���ʪ������"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   6720
            TabIndex        =   28
            Top             =   3330
            Width           =   1935
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ݿ�"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   2760
            TabIndex        =   29
            Top             =   3690
            Width           =   855
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ƥ����Ѫ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   3840
            TabIndex        =   30
            Top             =   3690
            Width           =   1095
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�˿���Ѫ��ֹ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   5160
            TabIndex        =   31
            Top             =   3690
            Width           =   1455
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ȴ���Ѫ����ĭ��̵"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   6720
            TabIndex        =   32
            Top             =   3690
            Width           =   2175
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   2760
            TabIndex        =   33
            Top             =   4050
            Width           =   855
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ݡ����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   3840
            TabIndex        =   34
            Top             =   4050
            Width           =   855
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������ŭ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   5160
            TabIndex        =   35
            Top             =   4050
            Width           =   1335
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ɫ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   6720
            TabIndex        =   36
            Top             =   4050
            Width           =   1095
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ƾ�"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   2760
            TabIndex        =   37
            Top             =   4410
            Width           =   855
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ʹ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   3840
            TabIndex        =   38
            Top             =   4410
            Width           =   855
         End
         Begin VB.CheckBox Chk31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   5160
            TabIndex        =   39
            Top             =   4410
            Width           =   855
         End
         Begin VB.CheckBox Chk32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���ȷ�Ӧ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   40
            Top             =   4860
            Width           =   1095
         End
         Begin VB.CheckBox Chk32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������Ӧ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   41
            Top             =   4860
            Width           =   1095
         End
         Begin VB.CheckBox Chk32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������Ѫ��Ӧ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   42
            Top             =   4860
            Width           =   1455
         End
         Begin VB.CheckBox Chk32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   6720
            TabIndex        =   43
            Top             =   4860
            Width           =   855
         End
         Begin VB.CheckBox Chk33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "1.����ֹͣ��Ѫ�����־���ͨ·��ͬʱ�۲�ʣ��Ѫ���"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   44
            Top             =   5325
            Width           =   6255
         End
         Begin VB.CheckBox Chk33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "2.�ɻ���Ѫ������ʣ��Ѫ(��ú�Ѫ��һ��)����Ѫ�Ƽ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   45
            Top             =   5565
            Width           =   6255
         End
         Begin VB.CheckBox Chk33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "3.��ȡ��Ӧ���һ�����ͼ�"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   46
            Top             =   5805
            Width           =   6255
         End
         Begin VB.CheckBox Chk33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "4.��֢����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   47
            Top             =   6045
            Width           =   6255
         End
         Begin VB.TextBox TXT41 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   0
            Left            =   2760
            TabIndex        =   50
            Top             =   9615
            Width           =   1455
         End
         Begin VB.TextBox TXT41 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   1
            Left            =   5520
            MaxLength       =   8
            TabIndex        =   51
            Top             =   9615
            Width           =   1575
         End
         Begin VB.TextBox TXT41 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   2
            Left            =   8370
            TabIndex        =   52
            Top             =   9600
            Width           =   1455
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2760
            ScaleHeight     =   300
            ScaleWidth      =   2295
            TabIndex        =   66
            Top             =   2925
            Width           =   2295
            Begin VB.OptionButton Opt31 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��Ѫ�ڼ�"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton Opt31 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��Ѫ��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   21
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2760
            ScaleHeight     =   285
            ScaleWidth      =   3735
            TabIndex        =   65
            Top             =   1240
            Width           =   3735
            Begin VB.OptionButton Opt22 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "һ������"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   8
               Top             =   45
               Width           =   1095
            End
            Begin VB.OptionButton Opt22 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��������"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   9
               Top             =   45
               Width           =   1095
            End
            Begin VB.OptionButton Opt22 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "�޹�ϵ"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   2400
               TabIndex        =   10
               Top             =   45
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2760
            ScaleHeight     =   330
            ScaleWidth      =   1455
            TabIndex        =   64
            Top             =   2040
            Width           =   1455
            Begin VB.OptionButton Opt21 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   14
               Top             =   75
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton Opt21 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   15
               Top             =   75
               Width           =   615
            End
         End
         Begin VB.Label lbl��Ѫ������ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   9240
            TabIndex        =   117
            Top             =   735
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Line Line28 
            X1              =   2715
            X2              =   3405
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line Line27 
            X1              =   8910
            X2              =   9585
            Y1              =   2340
            Y2              =   2340
         End
         Begin VB.Line Line26 
            X1              =   8040
            X2              =   8655
            Y1              =   2340
            Y2              =   2340
         End
         Begin VB.Line Line25 
            X1              =   5220
            X2              =   4515
            Y1              =   2340
            Y2              =   2340
         End
         Begin VB.Line Line24 
            X1              =   2745
            X2              =   6375
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line Line23 
            X1              =   7605
            X2              =   9555
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Line Line22 
            X1              =   7620
            X2              =   8895
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line Line21 
            X1              =   7560
            X2              =   7560
            Y1              =   2835
            Y2              =   3240
         End
         Begin VB.Line Line20 
            X1              =   6600
            X2              =   6600
            Y1              =   2835
            Y2              =   3240
         End
         Begin VB.Label lblHave 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������Ѫ��Ӧ"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1515
            TabIndex        =   115
            Top             =   2520
            Width           =   1080
         End
         Begin VB.Line Line19 
            X1              =   1470
            X2              =   10110
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   10560
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   10560
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   10560
            Y1              =   9510
            Y2              =   9510
         End
         Begin VB.Line Line4 
            X1              =   1440
            X2              =   1440
            Y1              =   0
            Y2              =   10200
         End
         Begin VB.Label lbl1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������Ϣ"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   360
            TabIndex        =   98
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lbl2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��Ѫ���"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   360
            TabIndex        =   97
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lbl3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��Ѫ������Ӧ������"
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   240
            TabIndex        =   96
            Top             =   5880
            Width           =   975
         End
         Begin VB.Label lbl4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ǩ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   9600
            Width           =   855
         End
         Begin VB.Label lbl11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��    ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   94
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lbl11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��    ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   93
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lbl11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��    ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   7080
            TabIndex        =   92
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lbl11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ס Ժ ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   91
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lbl11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ѫ    ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   4440
            TabIndex        =   90
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lbl11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "RH(D)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   7080
            TabIndex        =   89
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lbl11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ٴ����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   1680
            TabIndex        =   88
            Top             =   840
            Width           =   855
         End
         Begin VB.Line Line5 
            X1              =   1440
            X2              =   10080
            Y1              =   1665
            Y2              =   1665
         End
         Begin VB.Line Line6 
            X1              =   1440
            X2              =   10080
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Line Line7 
            X1              =   2640
            X2              =   2640
            Y1              =   1200
            Y2              =   10200
         End
         Begin VB.Label lbl21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������Ѫʷ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   87
            Top             =   2145
            Width           =   975
         End
         Begin VB.Label lbl21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ʷ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   6720
            TabIndex        =   86
            Top             =   2130
            Width           =   735
         End
         Begin VB.Line Line8 
            X1              =   7560
            X2              =   7560
            Y1              =   1200
            Y2              =   2385
         End
         Begin VB.Line Line9 
            X1              =   6600
            X2              =   6600
            Y1              =   1200
            Y2              =   2415
         End
         Begin VB.Label lbl21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��Ѫ��Ŀ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   85
            Top             =   1755
            Width           =   975
         End
         Begin VB.Label lbl21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   6600
            TabIndex        =   84
            Top             =   1755
            Width           =   975
         End
         Begin VB.Label lbl21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��Ѫ������Ѫ�ߵĹ�ϵ"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   1560
            TabIndex        =   83
            Top             =   1245
            Width           =   975
         End
         Begin VB.Label lbl21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ѫ�����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   6720
            TabIndex        =   82
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lbl211 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   81
            Top             =   2130
            Width           =   255
         End
         Begin VB.Label lbl211 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   7800
            TabIndex        =   80
            Top             =   2115
            Width           =   375
         End
         Begin VB.Label lbl211 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   8640
            TabIndex        =   79
            Top             =   2115
            Width           =   375
         End
         Begin VB.Line Line10 
            X1              =   1440
            X2              =   10080
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line11 
            X1              =   1440
            X2              =   10080
            Y1              =   4770
            Y2              =   4770
         End
         Begin VB.Line Line12 
            X1              =   1440
            X2              =   10080
            Y1              =   5205
            Y2              =   5205
         End
         Begin VB.Label lbl31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ʱ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   78
            Top             =   2940
            Width           =   975
         End
         Begin VB.Label lbl31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ת��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   6720
            TabIndex        =   77
            Top             =   2940
            Width           =   735
         End
         Begin VB.Label lbl31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��״������"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   76
            Top             =   3900
            Width           =   975
         End
         Begin VB.Label lbl31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   75
            Top             =   4890
            Width           =   975
         End
         Begin VB.Label lbl31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ٴ������ʩ"
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   4
            Left            =   1560
            TabIndex        =   74
            Top             =   6045
            Width           =   975
         End
         Begin VB.Label lbl31 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��Ѫ�ƴ����ʩ"
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   5
            Left            =   1560
            TabIndex        =   73
            Top             =   8325
            Width           =   975
         End
         Begin VB.Line Line13 
            X1              =   1440
            X2              =   10080
            Y1              =   7350
            Y2              =   7350
         End
         Begin VB.Label lbl31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "(h/d)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   5880
            TabIndex        =   72
            Top             =   2955
            Width           =   615
         End
         Begin VB.Line Line14 
            X1              =   5160
            X2              =   5900
            Y1              =   3165
            Y2              =   3165
         End
         Begin VB.Label lbl31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   7
            Left            =   2760
            TabIndex        =   71
            Top             =   6405
            Width           =   255
         End
         Begin VB.Label lbl41 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��ʿ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   70
            Top             =   9600
            Width           =   975
         End
         Begin VB.Label lbl41 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ҽʦ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   69
            Top             =   9600
            Width           =   855
         End
         Begin VB.Label lbl41 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��Ѫ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   7320
            TabIndex        =   68
            Top             =   9600
            Width           =   855
         End
         Begin VB.Line Line15 
            X1              =   4320
            X2              =   4320
            Y1              =   9525
            Y2              =   10005
         End
         Begin VB.Line Line16 
            X1              =   5400
            X2              =   5400
            Y1              =   9525
            Y2              =   10005
         End
         Begin VB.Line Line17 
            X1              =   7200
            X2              =   7200
            Y1              =   9525
            Y2              =   10005
         End
         Begin VB.Line Line18 
            X1              =   8280
            X2              =   8280
            Y1              =   9525
            Y2              =   10005
         End
      End
   End
   Begin VB.PictureBox pictop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      ScaleHeight     =   1050
      ScaleWidth      =   12075
      TabIndex        =   60
      Top             =   0
      Width           =   12075
      Begin VB.PictureBox pic5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   0
         Left            =   360
         ScaleHeight     =   735
         ScaleWidth      =   3855
         TabIndex        =   61
         Top             =   120
         Width           =   3855
         Begin VB.PictureBox pic51 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   2160
            ScaleHeight     =   495
            ScaleWidth      =   975
            TabIndex        =   109
            Top             =   120
            Width           =   975
            Begin VB.Label lbl5 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "lbl5"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   111
               Top             =   0
               Width           =   495
            End
         End
         Begin MSMask.MaskEdBox Msk51 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   103
            Top             =   120
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483646
            MaxLength       =   10
            Format          =   "YYYY-MM-DD"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker DTP5 
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   107
            Top             =   120
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   125239299
            CurrentDate     =   42677
         End
         Begin MSMask.MaskEdBox Msk52 
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   108
            Top             =   240
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483646
            MaxLength       =   8
            Format          =   "HH:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
      End
   End
   Begin VB.PictureBox Picright 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10095
      Left            =   11040
      ScaleHeight     =   10095
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
      Begin VB.VScrollBar vsbRight 
         Height          =   1695
         Left            =   960
         TabIndex        =   118
         Top             =   2400
         Width           =   255
      End
      Begin VB.PictureBox Pic6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   0
         Left            =   240
         ScaleHeight     =   1935
         ScaleWidth      =   1575
         TabIndex        =   59
         Top             =   360
         Width           =   1575
         Begin VB.PictureBox pic61 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   360
            ScaleHeight     =   495
            ScaleWidth      =   735
            TabIndex        =   110
            Top             =   720
            Width           =   735
            Begin VB.Label lbl6 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "lbl6"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   112
               Top             =   0
               Width           =   375
            End
         End
         Begin MSMask.MaskEdBox Msk61 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   104
            Top             =   0
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483646
            MaxLength       =   10
            Format          =   "YYYY-MM-DD"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker DTP6 
            Height          =   375
            Index           =   0
            Left            =   840
            TabIndex        =   105
            Top             =   0
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   125239299
            CurrentDate     =   42677
         End
         Begin MSMask.MaskEdBox Msk62 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   106
            Top             =   360
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   -2147483646
            MaxLength       =   8
            Format          =   "HH:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame Fra1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "��ǩ״̬"
         ForeColor       =   &H00000000&
         Height          =   1815
         Left            =   120
         TabIndex        =   54
         Top             =   7320
         Width           =   1455
         Begin VB.Label lbl��ǩ״̬ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "�����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   760
            Width           =   1215
         End
         Begin VB.Label lbl��ǩ״̬ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            Caption         =   "ҽ�����ύ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lbl��ǩ״̬ 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "��Ѫ���ύ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   1400
            Width           =   1215
         End
         Begin VB.Label lbl��ǩ״̬ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��ǩ״̬"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "usrCardEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'�������ݺ󣬲���������ݶ�ȥ�޸Ļ�ɾ���������ݣ��ͻ���ɴ�������⣬�������Ű�ɾ�����ݺ��漸��ҳ�������ȫ����ǰ�ƶ�
Option Explicit
 
'Implements clsBloodEdit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function HideCaret Lib "User32.dll" (ByVal hWnd As Long) As Boolean
Private Declare Function DestroyCaret Lib "user32" () As Long
'״̬����
Public Enum Status
    ȱʡ = -1 'Ĭ��״̬
    ���� = 0
    ���� = 1
    �޸� = 2
    ɾ�� = 3
End Enum

Public Enum Position '��ǩ��ʾ��ʲôλ��
    PosiTop = 0
    Posiright = 1
End Enum
Public Enum ��״������
    ���� = 0
    ��� = 1
    �������� = 2
    ���β���ʪ������ = 3
    �ݿ� = 4
    Ƥ����Ѫ = 5
    �˿���Ѫ��ֹ = 6
    �ȴ���Ѫ����ĭ��̵ = 7
    ���� = 8
    ݡ���� = 9
    ������ŭ�� = 10
    ����ɫ�� = 11
    �ƾ� = 12
    ����ʹ = 13
    ������״ = 14
End Enum

Public Enum ��� '�� ��״������ ͬ��
    ���ȷ�Ӧ = 0
    ������Ӧ = 1
    ������Ѫ��Ӧ = 2
    ������� = 3
End Enum

Public Enum �ٴ������ʩ
    ����ֹͣ��Ѫ = 0
    ѪҺ����Ѫ�� = 1
    �����ͼ� = 2
    ��֢���� = 3
End Enum
'ȱʡ����ֵ:
Const m_def_TabsPosition = PosiTop

'���Ա���:
Private m_TabsPosition As Position

'�ؼ�����
Private mPicTabs As PictureBox
Private mTXTbox As TextBox
Private mButtion As CommandButton
Private mDataChanged As Boolean
'ȫ�ֱ���
Private mlngSum As Long                    '�洢tab�ĸ���
Private mstrST As Status                   '�洢��ǰ�����������桢ɾ����
Private mlngSelNum As Long                 'ѡ��tab��index
Private mlng����ID As Long
Private mlng��ҳid As Long
Private mlng������Դ As Long
Private mlng�շ�ID As Long
Private mgcnCpOracle As ADODB.Connection
Private mRsBR As ADODB.Recordset           '��Ų��˵Ļ�����Ϣ
Private mRsFY As ADODB.Recordset           '�����Ѫ��Ӧ��¼��Ϣ
Private mrsXD As ADODB.Recordset           '���Ѫ����ţ�ѪҺ���ƣ��в��������Ϣ
Private mlng״̬ As Long                   '����˼����״̬��Ϣ
Private mlng�׶� As Long                   '��ʾ�û��Ľ׶���ҽ�������׶κ���Ѫ�Ʋ����׶�
Private mblnStart As Boolean               '�������ʼ��ȫ�ֱ���
Private mobjfrm As Object
Private mblnAddPage As Boolean             '����ҳ���־
Private mblnCancel As Boolean              'ȡ����־
Private mlngģ��� As Long
Private mblnAddNew As Boolean
Private marrFilter                         '��������
Private mstrFilter As String               '�������ݴ�
Private mblnHaveData As Boolean            '�û������ݴ��뼴����Ա����ʱ��mblnHaveData=true,û����Ա������һֱΪfalse
Private mstr�Һŵ� As String
Private mbln��Ѫ������ As Boolean           '�Ƿ�����Ѫ������������
Private mbln��Ѫ������Ȩ�� As Boolean       '��Ѫ���Ƿ�������Ȩ��
'Event Clear() '��������ϳ��û���������������ݣ�����ɾ����
'�������ڽ�ͼ��api

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''��ǩ��λ'''''''''''''''''''''''''''''''''
Public Function BloodLocation(ByVal lng�շ�ID As Long) As Boolean
    Dim lngi As Long
    If mRsFY Is Nothing Then Exit Function
    If mRsFY.State = adStateClosed Then BloodLocation = False: Exit Function
    'mRsFY.Filter = "�շ�id = '" & mRsFY!�շ�ID & "'"
    If mRsFY.RecordCount = 0 Then BloodLocation = False: Exit Function
    mRsFY.MoveFirst
    '��λ�����ҵ�������
    For lngi = 0 To mRsFY.RecordCount - 1 '���ܹ���ѯ���������ʱ
        If lng�շ�ID = mRsFY!�շ�ID & "" Then
            If m_TabsPosition = PosiTop Then
                pic5_GotFocus (lngi)
            ElseIf m_TabsPosition = Posiright Then
                Pic6_GotFocus (lngi)
            End If
            BloodLocation = True
            Exit For
        End If
        mRsFY.MoveNext
    Next
    mRsFY.MoveFirst
End Function
'''''''''''''������������ɾ�ģ����棬��ӡ�Ȳ���'''''''''''''''''''''''''''''''''
Public Function ShowSave() As Boolean
    '���ܣ�����
    '������
    '���أ�
    If mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then Exit Function '��Ѫ��û������Ȩ��ʱ�����ܱ���   ���˸�   mlng״̬ = 0 And
    If cbo2.Text = "" Then MsgBox "Ѫ�����Ϊ�գ����ܱ��棡", vbInformation, gstrSysName: Exit Function
    Call ExecuteCommand("��������") '��������Ҫ���޸ı������������
    Call ExecuteCommand("�ؼ�״̬") '���ݲ�ͬ�Ĳ�������ɾ�ģ����ı�ؼ���״̬
    ShowSave = True
End Function
Public Function ShowDelete() As Boolean
    '���ܣ�ɾ��
    '������
    '���أ�
    If (mlng�׶� = 1 And mlng״̬ = 1) Or (mlng�׶� = 2 And mlng״̬ = 2) Then Exit Function '��������ݲ���ɾ��
    If MsgBox("�Ƿ�ɾ���ü�¼��", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Function
    Call ExecuteCommand("ɾ������") '���״̬Ϊɾ��������ɾ���������
    ShowDelete = True
End Function

Public Sub showPrintSet() '��ӡ����
    Call ReportPrintSet(gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm)
End Sub

Public Sub ShowPrint(id As Long)  '��ӡ
    '���ܣ���ӡ��Ԥ����
    '������id-1:��ӡ 2:Ԥ��
    '���أ�
    If id = 1 Then '��ӡ
        ReportOpen gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm, "�շ�id=" & Val(cbo2.Text), 2
    ElseIf id = 2 Then '��ӡԤ��
        ReportOpen gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm, "�շ�id=" & Val(cbo2.Text), 1
    End If
End Sub
Public Sub ShowPrintList(id As Long)
    '��ʾ��ӡ�б����԰����û��������ӡ����
    Dim strPrint As String
    Dim ArrPrint
    Dim lngFilter As Long
    Dim lngi As Long
    Dim strSelBloodid As String
    If mRsFY Is Nothing Then MsgBox "δѡ�в��˻��޲�����Ϣ��", vbInformation, gstrSysName: Exit Sub
    If mRsFY.RecordCount = 0 Then MsgBox "�޸ò��˵ķ�Ӧ��¼��", vbInformation, gstrSysName: Exit Sub
    
    strSelBloodid = cbo2.Text & ""
    
    strPrint = frmbloodReactionPrint.BloodPrintList(mlng����ID, mlng������Դ, mlng��ҳid, mstrFilter, mlng�׶�, strSelBloodid)
    
    If strPrint = "" Then Exit Sub 'û��ѡ��Ҫ��ӡ���������˳�
    ArrPrint = Split(strPrint, ";")
    For lngi = 0 To UBound(ArrPrint)
        If id = 1 Then '��ӡ
            ReportOpen gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm, "�շ�id=" & Val(ArrPrint(lngi)), 2
        ElseIf id = 2 Then '��ӡԤ��
            ReportOpen gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm, "�շ�id=" & Val(ArrPrint(lngi)), 1
        End If
    Next
End Sub

Public Sub ShowCancel()
    '���ܣ�ȡ����ǰ�������ỹԭΪ��ǰ��״̬
    '������
    '���أ�
'    Call ExecuteCommand("ˢ��ȫ������")
    Dim Arrtag
    
    mblnCancel = True
    Call ExecuteCommand("�ؼ�״̬")
    If mstrST = ���� And mlngSelNum <> 0 Then
        If m_TabsPosition = PosiTop Then
            Unload lbl5(mlngSelNum)
            Unload DTP5(mlngSelNum)
            Unload Msk51(mlngSelNum)
            Unload Msk52(mlngSelNum)
            Unload pic51(mlngSelNum)
            Unload Pic5(mlngSelNum)
        Else
            Unload lbl6(mlngSelNum)
            Unload DTP6(mlngSelNum)
            Unload Msk61(mlngSelNum)
            Unload Msk62(mlngSelNum)
            Unload pic61(mlngSelNum)
            Unload Pic6(mlngSelNum)
        End If
        mlngSum = mlngSum - 1
        If mlngSelNum > mlngSum Then 'ѡ�е�ѡ���������ҳ���ϣ����ɾ��������ҳ�棬��ôҲҪ����ѡ��״̬����Ȼ�����
            mlngSelNum = mlngSum
        End If
        mstrST = ȱʡ
        mblnAddPage = False
    Else
        If m_TabsPosition = PosiTop Then
            Msk51(mlngSelNum).Visible = False
            Msk52(mlngSelNum).Visible = False
        Else
            Msk61(mlngSelNum).Visible = False
            Msk62(mlngSelNum).Visible = False
        End If
        mstrST = ȱʡ
        mblnAddPage = False
    End If
    Call ExecuteCommand("ˢ������")
    If m_TabsPosition = PosiTop Then
        Arrtag = Split(lbl5(mlngSelNum).Tag, ":")
        lbl5(mlngSelNum).Tag = Arrtag(0) & ":" & mstrST & ":0"
        If mstrST <> ���� Then
            mDataChanged = False
        End If
        mblnCancel = False
        Pic5(mlngSelNum).SetFocus
    ElseIf m_TabsPosition = Posiright Then
        Arrtag = Split(lbl6(mlngSelNum).Tag, ":")
        lbl6(mlngSelNum).Tag = Arrtag(0) & ":" & mstrST & ":0"
        If mstrST <> ���� Then
            mDataChanged = False
        End If
        mblnCancel = False
        Pic6(mlngSelNum).SetFocus
    End If
    
End Sub

Public Function SubmitData() As Boolean
    '���ܣ��ύ����
    '������
    '���أ�
    Dim SelNum As Long '����ѡ��ҳ�棬��ֹ�ύҳ���ҳ�治��ѡ�С�
    If mstrST = ���� Then Exit Function '����ҳ�治���ύ��ֻ������ҳ�汣���������ύ
    SelNum = mlngSelNum
    Call ExecuteCommand("�ύ����")
    mlngSelNum = SelNum
    If m_TabsPosition = PosiTop Then '�ύ���ݺ�mblngoptchange=true
        pic5_GotFocus (mlngSelNum)
    ElseIf m_TabsPosition = Posiright Then
        Pic6_GotFocus (mlngSelNum)
    End If
    SubmitData = True
End Function

Public Sub ShowModify()
    '���ܣ��޸�
    '������
    '���أ�
    Dim Arrtag
    Dim blnDtpVisible As Boolean
    
    If mlng�׶� = 2 Then
        blnDtpVisible = lbl��Ѫ������.Visible
    Else
        blnDtpVisible = True
    End If
    
    If m_TabsPosition = PosiTop Then '
        Arrtag = Split(lbl5(mlngSelNum).Tag, ":")
        lbl5(mlngSelNum).Tag = Arrtag(0) & ":2:1"
        mstrST = �޸�
        DTP5(mlngSelNum).Visible = blnDtpVisible
    Else
        Arrtag = Split(lbl6(mlngSelNum).Tag, ":")
        lbl6(mlngSelNum).Tag = Arrtag(0) & ":2:1"
        mstrST = �޸�
        DTP6(mlngSelNum).Visible = blnDtpVisible
    End If
    
    Call ExecuteCommand("�ؼ�״̬") '���ݲ�ͬ�Ĳ�������ɾ�ģ����ı�ؼ���״̬
    mDataChanged = True
    If Not UserControl.ActiveControl Is Nothing Then
        If UserControl.ActiveControl.name = "TXT21" Then
            If UserControl.ActiveControl.Index = 5 Then Call TXT21_GotFocus(5)
        End If
    End If
    If mbln��Ѫ������ = True Then
        TXT41(1).SelStart = 0
        TXT41(1).SelLength = Len(TXT41(1).Text)
        TXT41(1).SetFocus
    End If
End Sub

Public Sub AddPage()
    '���ܣ�����ҳ��
    '������
    '���أ�
    If Not mRsFY Is Nothing Then
        If mRsFY.RecordCount = 0 Then
            mDataChanged = True
            mblnAddPage = True
            Call ExecuteCommand("�ؼ�״̬") '
            If Not UserControl.ActiveControl Is Nothing Then
                If UserControl.ActiveControl.name = "TXT21" Then
                    If UserControl.ActiveControl.Index = 5 Then Call TXT21_GotFocus(5)
                End If
            End If
            Exit Sub
        End If
    End If
    If mblnAddPage = True Then Exit Sub
    Call ExecuteCommand("����ҳ��") '
    mDataChanged = True
    mblnAddPage = True
End Sub

Public Sub ShowClear()
    '���ҳ�����ݣ�ͬʱ��ճ���ʼ�ؼ�������пؼ�
    Dim lngi As Long
    
    Clear
    For lngi = 0 To TXT11.Count - 1
        TXT11(lngi).Text = ""
    Next
    lbl��Ѫ������.Visible = False
    If m_TabsPosition = Posiright Then
        If Pic6.Count > 1 Then
            For lngi = 1 To Pic6.Count - 1
                Unload DTP6(lngi)
                Unload lbl6(lngi)
                Unload Msk61(lngi)
                Unload Msk62(lngi)
                Unload pic61(lngi)
                Unload Pic6(lngi)
            Next
        End If
    Else
        If Pic5.Count > 1 Then
            For lngi = 1 To Pic5.Count - 1
                Unload DTP5(lngi)
                Unload lbl5(lngi)
                Unload Msk51(lngi)
                Unload Msk52(lngi)
                Unload pic51(lngi)
                Unload Pic5(lngi)
            Next
        End If
    End If
    Set mRsBR = Nothing
    Set mRsFY = Nothing
    mstrST = ȱʡ
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Public Sub ShowUntread()
    '���ܣ�����
    '��������
    '���أ���
    If mstrST = ���� Then Exit Sub '����ҳ�治�ܻ��ˣ�ֻ������ҳ�汣�����������
    Dim SelNum As Long
    SelNum = mlngSelNum
    Call ExecuteCommand("��������")
    mlngSelNum = SelNum
'    Call ExecuteCommand("ˢ������")
    If m_TabsPosition = PosiTop Then 'ˢ��ȫ�����ݺ�mblngoptchange=true
        pic5_GotFocus (mlngSelNum)
    ElseIf m_TabsPosition = Posiright Then
        Pic6_GotFocus (mlngSelNum)
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitEdit()
    '���ܣ���ʼ��ҳ�棬��ָ�����˵���Ϣ��ʾ�ڽ����ϣ�Ĭ�ϲ��ɱ༭
    '����������ID���˵����id���ڲ����û��������Ϣ��
    '      ҽ��ID��ѯҽ�������Ϣ
    '����
    mstrST = ȱʡ '��ʼ��״̬��Ϣ
    mlngSum = 0
    mlng״̬ = 0
    mlngSelNum = 0
    mDataChanged = False
    pic3.Visible = False
    CboDW.Clear
    CboDW.AddItem "ml"
    CboDW.AddItem "U"
    CboDW.AddItem "������"
    CboDW.ListIndex = 0
    cboHave.Clear
    cboHave.AddItem "��"
    cboHave.AddItem "��"
    Call ExecuteCommand("�ؼ�����")
    Call ExecuteCommand("��ʼ�ؼ�")
    
End Sub

Public Sub showInfor(lng����ID As Long, lng������Դ As Long, lng��ҳid As Long, lng�׶� As Long, cnMain As ADODB.Connection, objfrmMain As Object, lngģ��� As Long, _
                Optional strFilter As String = "", Optional bln��Ѫ������Ȩ�� As Boolean = False, Optional ByVal lng�շ�ID As Long)
    '���ܣ����ݲ���id����ҳid��ʾ���˵�����
    '������lng����id-���˵�id�ţ�lng������Դ-סԺ2������1��lng��ҳid-���˵���ҳid�����id��lng�׶�-ҽ���׶�1����Ѫ�ƽ׶�2��
    '    : cnMain���ݿ����ӣ�objfrm-�����壬lngϵͳ��-�������ϵͳ��
    '    : Filter�ǹ����������飬����������˿��ң�����ʱ�䣬��д�ˣ��ύ״̬��
    '    ��bln��Ѫ������Ȩ��-true��Ѫ��������Ȩ�� false-��Ѫ��������Ȩ��
    '���أ�
    Dim lngi As Long
    Dim strSQL As String
    Dim rsSQL As ADODB.Recordset
    Dim blnLocated As Boolean
'    ReDim mArrFilter(0 To 3)
    
    mlng����ID = lng����ID
    mlng��ҳid = lng��ҳid '��������סԺ����ʱ����ҳid��סԺ���˵���ҳid�����������ﲡ��ʱ����ҳid�ǲ��˵ľ���id������Ϊ�˷��㹲����mlng��ҳid��һ����
    mlng�׶� = lng�׶�
    mlngģ��� = lngģ���
    mlng������Դ = lng������Դ
    mbln��Ѫ������Ȩ�� = bln��Ѫ������Ȩ��
    mlng�շ�ID = lng�շ�ID
    
    mstrFilter = strFilter
    If strFilter <> "" Then '����й��������򽫹���������������
        marrFilter = Split(strFilter, "|")
    Else '���û�й����������ض������飬�������е���������Ϊ��
        ReDim marrFilter(0 To 4)
    End If
    
    Set mobjfrm = objfrmMain
    Set mgcnCpOracle = cnMain
    
    
    If zlGetComLib = False Then MsgBox "��ȡ����ʧ�ܣ�", vbInformation, gstrSysName: Exit Sub
    mblnAddPage = False '��ʼ�����ҳ����Ϣ
    mblnHaveData = True
    HS1.Visible = False
    VS1.Visible = False
    
    '��ȡ�Һŵ���
    mstr�Һŵ� = ""
    If mlng������Դ <> 2 Then
        strSQL = " select no from ���˹Һż�¼ where id=[1] "
        Set rsSQL = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", mlng��ҳid)
        If rsSQL.RecordCount > 0 Then
            mstr�Һŵ� = rsSQL.Fields("no")
        End If
    End If
    
    Call ExecuteCommand("��տؼ�")
    Call gobjControl.PicShowFlat(pic3, 1)
'    Call ExecuteCommand("�ؼ�����")
    Call ExecuteCommand("��ʼ������Ϣ") '1����ѯ��������ݷ������ݼ��У�����������ؼ�������Դ
    Call ExecuteCommand("��ȡ��Ӧ��¼")
    If mRsFY.BOF = False Then
        For lngi = 0 To mRsFY.RecordCount - 1 '���ܹ���ѯ���������ʱ
            mlngSum = lngi
            mlng״̬ = Val(mRsFY.Fields(16).Value & "")
            Call ExecuteCommand("��ʼ�ؼ�")
            cbo2.Text = mRsFY.Fields(0).Value & ""
            
            mRsFY.MoveNext
        Next
        If mlng�շ�ID = 0 Then
        mlngSelNum = mlngSum
        Else
            If Not BloodLocation(mlng�շ�ID) Then
                AddPage
                blnLocated = False
            Else
                blnLocated = True
            End If
        End If
        mRsFY.MoveFirst
    Else
        mlngSum = 0
        mlng״̬ = 0
        Call ExecuteCommand("��ʼ�ؼ�")
        Call ExecuteCommand("���˻�������")
        mstrST = ����
    End If
    
'    Call ExecuteCommand("ˢ������")
    Call ExecuteCommand("�ؼ�״̬") '4��Ĭ�����пؼ����ɱ༭
    mblnStart = True
    If mlng�շ�ID <> 0 Then
        If Not blnLocated Then
            mstrST = ����
            Call ExecuteCommand("�ؼ�״̬")
            Call ExecuteCommand("���˻�������")
            '��дABO��RH
            If mRsBR.Fields("ABO").Value & "" = "" Then
                Set rsSQL = GetPatientOtherInfo(mlng����ID, "ABO")
                If rsSQL.BOF = False Then TXT11(4).Text = rsSQL("��Ϣֵ").Value
            Else
            TXT11(4).Text = mRsBR.Fields("ABO").Value
            End If
            If mRsBR.Fields("RH").Value & "" = "" Then
                Set rsSQL = GetPatientOtherInfo(mlng����ID, "RH")
                If rsSQL.BOF = False Then TXT11(5).Text = rsSQL("��Ϣֵ").Value
            Else
                TXT11(5).Text = mRsBR.Fields("RH").Value
            End If
            Call pic3_MouseDown(1, 0, 0, 0)
            mlng�շ�ID = 0
            vsbRight.Visible = False
            If Pic6.Count > 19 Then
                vsbRight.Max = Pic6(Pic6.UBound).Height + Pic6(Pic6.UBound).Top - Fra1.Top
                vsbRight.Max = (Pic6(2).Top - Pic6(1).Top) * (Pic6.UBound - 18)
                vsbRight.Visible = True
                vsbRight.Value = vsbRight.Max
            End If
        End If
    Else
    If m_TabsPosition = PosiTop Then
        pic5_GotFocus (mlngSelNum)
    ElseIf m_TabsPosition = Posiright Then
        Pic6_GotFocus (mlngSelNum)
        End If
    End If
    If TXT11(1).Text = "��" Then
        lbl21(1).ForeColor = &H80000000
        lbl211(1).ForeColor = &H80000000
        Line26.BorderColor = &H80000000
        lbl211(2).ForeColor = &H80000000
        Line27.BorderColor = &H80000000
        TXT21(1).Enabled = False
        TXT21(2).Enabled = False
        TXT21(1).locked = True
        TXT21(2).locked = True
    Else
        lbl21(1).ForeColor = vbBlack
        lbl211(1).ForeColor = vbBlack
        Line26.BorderColor = vbBlack
        lbl211(2).ForeColor = vbBlack
        Line27.BorderColor = vbBlack
        TXT21(1).Enabled = True
        TXT21(2).Enabled = True
        TXT21(1).locked = False
        TXT21(2).locked = False
    End If
End Sub


'���Ի�ȡ��ֵ
Public Property Get BloodID() As Long
    BloodID = Val(cbo2.Text)
End Property

Public Property Get TabsPosition() As Position
    TabsPosition = m_TabsPosition
    Call ExecuteCommand("�ؼ�����")
    Call ExecuteCommand("��ʼ�ؼ�")
End Property

Public Property Let TabsPosition(ByVal NewTabsPosition As Position)
    m_TabsPosition = NewTabsPosition
    PropertyChanged "TabsPosition"
    Call ExecuteCommand("�ؼ�����")
    Call ExecuteCommand("��ʼ�ؼ�")
End Property
'��ȡ��ǰ���˵���Ѫ��Ӧ������
Public Property Get lngFYCount() As Long
    If Not mRsFY Is Nothing Then
        lngFYCount = mRsFY.RecordCount
    Else
        lngFYCount = 0
    End If
End Property

Public Property Get ��Ѫ������() As Boolean
    ��Ѫ������ = mbln��Ѫ������
End Property

Public Property Get ������Ѫ��Ӧ() As Boolean
    ������Ѫ��Ӧ = IIf(cboHave.Text = "��", True, False)
End Property

Public Property Get Doctor() As String
    Doctor = TXT41(1).Text
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mDataChanged
End Property

Public Property Let DataChanged(ByVal DataChanged As Boolean)
    mDataChanged = DataChanged
End Property

Public Property Get lng״̬() As Long
    lng״̬ = mlng״̬
End Property

Public Property Get strST() As Status
    strST = mstrST
End Property

Public Property Get blnAddPage() As Boolean
    blnAddPage = mblnAddPage
End Property

Private Function CanbeChange(lng�׶� As Long, lng״̬ As Long, st���� As Status) As Boolean
    '���ܣ����Ը��ݽ׶κ�״̬�����Ŀؼ���״̬������ͨ������������ж�һЩ�����Ƿ����
    '������st�������ֽ׶�������ʲô���������޸Ļ���ɾ��
    '���أ�
    Dim lngi As Long
    Dim str�׶����ж� As String
    Dim bln���޷�Ӧ As Boolean
    str�׶����ж� = lng�׶� & ":" & lng״̬
    Select Case st����
        Case �޸�:
                Select Case str�׶����ж�
                    Case "1:0":
                        For lngi = 0 To TXT21.Count - 1
                            TXT21(lngi).locked = False
                        Next
                        TXT31.locked = False
                        TXT32.locked = False
                        TXT21(5).locked = False
                    Case "2:1":
                        TXT33.locked = False
                    Case "2:0" '���˸�   ��Ѫ������ҳ����޸�
                        If mbln��Ѫ������Ȩ�� = True Then
                            For lngi = 0 To TXT21.Count - 1
                                TXT21(lngi).locked = False
                            Next
                            TXT31.locked = False
                            TXT32.locked = False
                            TXT33.locked = False
    '                        TXT21(5).Locked = False
                        End If
                End Select
                '���Բ�����д����ʷ
                If TXT11(1).Text = "��" Then
                    lbl21(1).ForeColor = &H80000000
                    lbl211(1).ForeColor = &H80000000
                    Line26.BorderColor = &H80000000
                    lbl211(2).ForeColor = &H80000000
                    Line27.BorderColor = &H80000000
                    TXT21(1).Enabled = False
                    TXT21(2).Enabled = False
                    TXT21(1).locked = True
                    TXT21(2).locked = True
                Else
                    lbl21(1).ForeColor = vbBlack
                    lbl211(1).ForeColor = vbBlack
                    Line26.BorderColor = vbBlack
                    lbl211(2).ForeColor = vbBlack
                    Line27.BorderColor = vbBlack
                    TXT21(1).Enabled = True
                    TXT21(2).Enabled = True
                    TXT21(1).locked = False
                    TXT21(2).locked = False
                End If
        Case ����:
            '������Ȼ��ҽ������������������ֻ��ҽ����д����ʹ��
            For lngi = 0 To TXT21.Count - 1
                TXT21(lngi).locked = Not mDataChanged
            Next
            '���Բ�����д����ʷ
            If TXT11(1).Text = "��" Then
                lbl21(1).ForeColor = &H80000000
                lbl211(1).ForeColor = &H80000000
                Line26.BorderColor = &H80000000
                lbl211(2).ForeColor = &H80000000
                Line27.BorderColor = &H80000000
                TXT21(1).Enabled = False
                TXT21(2).Enabled = False
                TXT21(1).locked = True
                TXT21(2).locked = True
            Else
                lbl21(1).ForeColor = vbBlack
                lbl211(1).ForeColor = vbBlack
                Line26.BorderColor = vbBlack
                lbl211(2).ForeColor = vbBlack
                Line27.BorderColor = vbBlack
                TXT21(1).Enabled = True
                TXT21(2).Enabled = True
                TXT21(1).locked = False
                TXT21(2).locked = False
            End If
            TXT31.locked = Not mDataChanged
            TXT32.locked = Not mDataChanged
            TXT21(5).locked = Not mDataChanged
            If mbln��Ѫ������Ȩ�� = True Then '���˸ģ��������Ѫ������Ȩ�������ֱ��������ҳ�������Ѫ�ƴ����ʩ��
                TXT33.locked = Not mDataChanged
            End If
    End Select
    
    bln���޷�Ӧ = IIf(cboHave.Text = "��", True, False)
    '�����Ƿ�����Ѫ��Ӧ��ѡ�������Ѫ��Ӧ���ֿؼ�
    If cboHave.Text <> "��" Then
        Opt31(0).Value = False
        Opt31(1).Value = False
        Opt31(0).Tag = 2
        Opt31(1).Tag = 2
        Opt32(0).Value = False
        Opt32(1).Value = False
        Opt32(0).Tag = 2
        Opt32(1).Tag = 2
    End If
    Opt31(0).Enabled = bln���޷�Ӧ
    Opt31(1).Enabled = bln���޷�Ӧ
    TXT31.Enabled = bln���޷�Ӧ
    lbl31(6).ForeColor = IIf(bln���޷�Ӧ = False, &H80000011, &H80000008)
    Opt32(0).Enabled = bln���޷�Ӧ
    Opt32(1).Enabled = bln���޷�Ӧ
    For lngi = 0 To Chk31.Count - 1
        Chk31(lngi).Enabled = bln���޷�Ӧ
    Next
    For lngi = 0 To Chk32.Count - 1
        Chk32(lngi).Enabled = bln���޷�Ӧ
    Next
    For lngi = 0 To Chk33.Count - 1
        Chk33(lngi).Enabled = bln���޷�Ӧ
    Next
    lbl31(7).ForeColor = IIf(bln���޷�Ӧ = False, &H80000011, &H80000008)
    TXT32.Enabled = bln���޷�Ӧ
    TXT33.Enabled = bln���޷�Ӧ
End Function
Private Sub Clear()
    '���ܣ����ҳ���ϵĵ���������
    '����
    '����
    Dim lngi As Long

    For lngi = 0 To TXT21.Count - 1
        TXT21(lngi).Text = ""
    Next
    
    cboHave.Text = ""
    TXT31.Text = ""
    TXT32.Text = ""
    TXT33.Text = ""
    
    For lngi = 0 To TXT41.Count - 1
        TXT41(lngi).Text = ""
    Next
    Opt21(0).Value = True
    Opt21(0).Tag = 0
    Opt21(1).Tag = 0
    
    Opt22(2).Value = True
    Opt22(0).Tag = 2
    Opt22(1).Tag = 2
    Opt22(2).Tag = 2
    
    Opt31(0).Value = False
    Opt31(0).Tag = 0
    Opt31(1).Value = False
    Opt31(1).Tag = 0
    
    Opt32(0).Value = False
    Opt32(0).Tag = 0
    Opt32(1).Value = False
    Opt32(1).Tag = 0
    
    For lngi = 0 To Chk31.Count - 1
        Chk31(lngi).Value = Unchecked
        Chk31(lngi).Tag = 0
    Next
    For lngi = 0 To Chk32.Count - 1
        Chk32(lngi).Value = Unchecked
        Chk32(lngi).Tag = 0
    Next
    For lngi = 0 To Chk33.Count - 1
        Chk33(lngi).Value = Unchecked
        Chk33(lngi).Tag = 0
    Next
End Sub

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim lngi As Long
    Dim lngj As Long
    Dim lngk As Long
    Dim strSqlBR As String
    Dim strSqlFY As String
    Dim StrSplit
    Dim StrSqlSAD As String
    Dim rsSAD As New ADODB.Recordset
    Dim strXD As String '���Ѫ����ţ��в���¼����Ѫ��¼�ȵ�sql���
    Dim blnLoad As Boolean '�ж��Ƿ���ؿؼ�
    Dim rsTmp As New Recordset

    On Error GoTo Error

    Call SQLRecord(rsSAD)
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "�ؼ�����"
            pic4.Left = UserControl.ScaleLeft
            pic1.Left = pic4.ScaleLeft
            pic1.Top = pic4.ScaleTop
            
            If m_TabsPosition = PosiTop Then
                pic4.Top = pictop.Top + pictop.Height + 20
                pic4.Width = UserControl.ScaleWidth
                pic4.Height = IIf(UserControl.ScaleHeight - pictop.Height < 0, 0, UserControl.ScaleHeight - pictop.Height)
                pic1.Left = IIf(pic4.ScaleWidth < pic1.Width, pic4.ScaleLeft, (pic4.ScaleWidth - pic1.Width) / 2)
                pictop.Visible = True
                pictop.Left = pic1.Left
                If Pic5.Count >= 8 Then
                    pictop.Height = ((Pic5.Count - 1) \ 8 + 1) * (Pic5(Pic5.Count - 1).Height + 20)
                Else
                pictop.Height = 440
                End If
                pictop.Top = UserControl.ScaleTop
                pictop.Width = pic4.Width - pictop.Left
                For lngi = 0 To Pic5.Count - 1
                    If lngi < 8 Then
                    Pic5(lngi).Move pictop.ScaleLeft + lngi * 1265, pictop.ScaleTop + 20, 1215, 400
                    Else
                        Pic5(lngi).Move Pic5(lngi - 8).Left, Pic5(lngi - 8).Top + Pic5(lngi - 8).Height + 20, 1215, 400
                    End If
                    Pic5(lngi).Visible = True
                Next
                Picright.Visible = False
                For lngi = 0 To Pic6.Count - 1
                    Pic6(lngi).Visible = False
                Next
                
            Else
                pic4.Top = UserControl.ScaleTop
                Picright.Visible = True
                Picright.Top = pic4.Top

                pic4.Height = UserControl.ScaleHeight
                
                Picright.Height = pic4.Height
                Picright.Width = 1480
                
                pic4.Width = IIf(UserControl.ScaleWidth - Picright.Width < 0, 0, UserControl.ScaleWidth - Picright)
                If pic4.Left + pic4.Width + 50 > pic4.Left + pic1.Width + 50 Then
                    Picright.Left = pic4.Left + pic1.Width + 50
                    Picright.ZOrder 0
                Else
                    Picright.Left = pic4.Left + pic4.Width + 50
                End If
                pictop.Visible = False
                For lngi = 0 To Pic5.Count - 1
                    Pic5(lngi).Visible = False
                Next
                '�˶δ����ˢ���Ҳ��ǩλ�ã���ǰλ���޷��䶯�����������˹�������ˢ�»ᵼ�¶�λ��ʧ�����Ҵ�ˢ�����岻��������ʱ�����Ҳ��ǩˢ��
'                For lngi = 0 To Pic6.Count - 1
'                    Pic6(lngi).Move Picright.ScaleLeft, Picright.ScaleTop + lngi * 450, 1215, 400
'                    Pic6(lngi).Visible = True
'                Next
                
                '��ʾ��ǩ״̬
                Fra1.Visible = True
                Fra1.Move Picright.ScaleLeft, Picright.ScaleTop + Picright.ScaleHeight - 1500, 1220, 1455
                
                lbl��ǩ״̬(0).Move 30, 540, Fra1.Width - 60, 255
                lbl��ǩ״̬(1).Move 30, 810, Fra1.Width - 60, 255
                lbl��ǩ״̬(2).Move 30, 1080, Fra1.Width - 60, 255
                lbl��ǩ״̬(3).Move 30, 270, Fra1.Width - 60, 255
            
                vsbRight.Visible = False
                If Pic6.Count > 19 Then
                    vsbRight.Max = (Pic6(2).Top - Pic6(1).Top) * (Pic6.UBound - 18)
                    vsbRight.Move Picright.Width - vsbRight.Width, 0, vsbRight.Width, Picright.Height - 50
                    vsbRight.SmallChange = 450
                    vsbRight.LargeChange = 450
                    vsbRight.Visible = True
                End If
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "��ʼ�ؼ�"

            If m_TabsPosition = PosiTop Then
                For lngi = 0 To Pic5.Count - 1
                    If lngi = mlngSum Then
                        blnLoad = True
                    End If
                Next
                If Not blnLoad And mlngSum > 0 Then '��mlngsum>0�ҵ�ǰ�ؼ�����ѡ�пؼ�ʱ�ż�����ؿؼ��������Ǳ����Ѽ��ؿؼ��ظ�����
                    Load Pic5(mlngSum)
                    Load DTP5(mlngSum)
                    Load Msk51(mlngSum)
                    Load Msk52(mlngSum)
                    Load pic51(mlngSum)
                    Load lbl5(mlngSum)
                End If
                SetParent pic51(mlngSum).hWnd, Pic5(mlngSum).hWnd
                Set lbl5(mlngSum).Container = pic51(mlngSum) '����ǩ����������
                SetParent DTP5(mlngSum).hWnd, Pic5(mlngSum).hWnd
                SetParent Msk51(mlngSum).hWnd, Pic5(mlngSum).hWnd
                SetParent Msk52(mlngSum).hWnd, Pic5(mlngSum).hWnd
                Pic5(mlngSum).Move pictop.ScaleLeft + mlngSum * 1265, pictop.ScaleTop + 20, 1215, 400
                Pic5(mlngSum).Visible = True
                
                Msk51(mlngSum).Move Pic5(mlngSum).ScaleLeft, Pic5(mlngSum).ScaleTop, 975, 255
                Msk52(mlngSum).Move Msk51(mlngSum).Left, Msk51(mlngSum).Top + Msk51(mlngSum).Height - 70, Msk51(mlngSum).Width, Msk51(mlngSum).Height
                Msk51(mlngSum).Visible = False
                Msk52(mlngSum).Visible = False
                Msk51(mlngSum).Text = Format(Now, "YYYY-MM-DD")
                Msk52(mlngSum).Text = Format(Now, "HH:mm:ss")
                
                pic51(mlngSum).Move Msk51(mlngSum).Left, Msk51(mlngSum).Top, Msk51(mlngSum).Width, Msk51(mlngSum).Height * 2 - 70
                pic51(mlngSum).Visible = True
                lbl5(mlngSum).Move pic51(mlngSum).ScaleLeft, pic51(mlngSum).ScaleTop, pic51(mlngSum).ScaleWidth, pic51(mlngSum).ScaleHeight
                lbl5(mlngSum).Visible = True
                DTP5(mlngSum).Move Msk51(mlngSum).Left - 50, Msk51(mlngSum).Top - 50, Msk51(mlngSum).Width + 320, Msk51(mlngSum).Height * 2 - 10
                DTP5(mlngSum).Visible = False
                DTP5(mlngSum).Value = Now
                pic51(mlngSum).ZOrder 0
                lbl5(mlngSum).Tag = mlng״̬ & ":-1:0"   '��ʶ����ҳ���״̬��Ϣ����������ؼ�״̬,��ʽ����״̬:����:mdatachanged��
                mstrST = ȱʡ
            Else
                For lngi = 0 To Pic6.Count - 1
                    If lngi = mlngSum Then
                        blnLoad = True
                    End If
                Next
                If Not blnLoad And mlngSum > 0 Then '��mlngsum>0�ҵ�ǰ�ؼ�����ѡ�пؼ�ʱ�ż�����ؿؼ��������Ǳ����Ѽ��ؿؼ��ظ�����
                    Load Pic6(mlngSum)
                    Load DTP6(mlngSum)
                    Load Msk61(mlngSum) '''''''
                    Load Msk62(mlngSum)
                    Load pic61(mlngSum)
                    Load lbl6(mlngSum)
                End If
                SetParent pic61(mlngSum).hWnd, Pic6(mlngSum).hWnd
                Set lbl6(mlngSum).Container = pic61(mlngSum) '����ǩ����������
                SetParent DTP6(mlngSum).hWnd, Pic6(mlngSum).hWnd
                SetParent Msk61(mlngSum).hWnd, Pic6(mlngSum).hWnd
                SetParent Msk62(mlngSum).hWnd, Pic6(mlngSum).hWnd
                Pic6(mlngSum).Move Picright.ScaleLeft, Picright.ScaleTop + mlngSum * 450, 1215, 400
                Pic6(mlngSum).Visible = True
                
                Msk61(mlngSum).Move Pic6(mlngSum).ScaleLeft, Pic6(mlngSum).ScaleTop, 975, 255
                Msk62(mlngSum).Move Msk61(mlngSum).Left, Msk61(mlngSum).Top + Msk61(mlngSum).Height - 70, Msk61(mlngSum).Width, Msk61(mlngSum).Height
                Msk61(mlngSum).Visible = False
                Msk62(mlngSum).Visible = False
                Msk61(mlngSum).Text = Format(Now, "YYYY-MM-DD")
                Msk62(mlngSum).Text = Format(Now, "HH:mm:ss")
'
                pic61(mlngSum).Move Msk61(mlngSum).Left, Msk61(mlngSum).Top, Msk61(mlngSum).Width, Msk61(mlngSum).Height * 2 - 70
                pic61(mlngSum).Visible = True
                lbl6(mlngSum).Move pic61(mlngSum).ScaleLeft, pic61(mlngSum).ScaleTop, pic61(mlngSum).ScaleWidth, pic61(mlngSum).ScaleHeight
                lbl6(mlngSum).Visible = True
                DTP6(mlngSum).Move Msk61(mlngSum).Left - 50, Msk61(mlngSum).Top - 50, Msk61(mlngSum).Width + 320, Msk61(mlngSum).Height * 2 - 10
                DTP6(mlngSum).Visible = False
                DTP6(mlngSum).Value = Now
                pic61(mlngSum).ZOrder 0
                lbl6(mlngSum).Tag = mlng״̬ & ":-1:0"
                mstrST = ȱʡ
            End If
            '����״̬��ʾ��ͬ��tab��ɫ
            If m_TabsPosition = PosiTop Then
                If mlng״̬ = 0 Then '����״̬�����ı���Ӧ�Ŀؼ�����ɫ
                    pic51(mlngSum).BackColor = &H80000002
                    Pic5(mlngSum).BackColor = &H80000002
                ElseIf mlng״̬ = 1 Then
                    pic51(mlngSum).BackColor = &HC0FFC0
                    Pic5(mlngSum).BackColor = &HC0FFC0
                ElseIf mlng״̬ = 2 Then
                    pic51(mlngSum).BackColor = &H80000000
                    Pic5(mlngSum).BackColor = &H80000000
                End If
            ElseIf m_TabsPosition = Posiright Then

                If mlng״̬ = 0 Then '����״̬�����ı���Ӧ�Ŀؼ�����ɫ
                    pic61(mlngSum).BackColor = &H80000002
                    Pic6(mlngSum).BackColor = &H80000002
                ElseIf mlng״̬ = 1 Then
                    pic61(mlngSum).BackColor = &HC0FFC0
                    Pic6(mlngSum).BackColor = &HC0FFC0
                ElseIf mlng״̬ = 2 Then
                    pic61(mlngSum).BackColor = &H80000000
                    Pic6(mlngSum).BackColor = &H80000000
                End If
            End If
            
            mlngSelNum = mlngSum
            If m_TabsPosition = PosiTop Then
                If mlngSum > 0 Then
'                   pic5(mlngSum).SetFocus'�޷��۽������쳣�������������ȥ��
                Else
                   lbl5(mlngSum).Caption = Format(Now, "YYYY-MM-DD HH:mm:ss")
                End If
            Else
                If mlngSum > 0 Then
'                   pic6(mlngSum).SetFocus
                Else
                   lbl6(mlngSum).Caption = Format(Now, "YYYY-MM-DD HH:mm:ss")
                End If
            End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "��տؼ�"
            'ж�س���һ��ѡ��������ѡ�����һ����Ϊ����ת������ʱʹ��
            If m_TabsPosition = PosiTop Then
                For lngi = 0 To Pic5.Count - 1
                    If lngi = 0 Then
                        Clear
                    Else
                        Unload DTP5(lngi)
                        Unload lbl5(lngi)
                        Unload Msk51(lngi)
                        Unload Msk52(lngi)
                        Unload pic51(lngi)
                        Unload Pic5(lngi)
                    End If
                Next
            Else
                For lngi = 0 To Pic6.Count - 1
                    If lngi = 0 Then
                        Clear
                    Else
                        Unload DTP6(lngi)
                        Unload lbl6(lngi)
                        Unload Msk61(lngi)
                        Unload Msk62(lngi)
                        Unload pic61(lngi)
                        Unload Pic6(lngi)
                    End If
                Next
            End If
            Call ExecuteCommand("�ؼ�����")
        Case "��ʼ������Ϣ"
            Dim sqlFilter As String
            '����   2017��2��8��
            strSqlBR = " Select b.����id, a.סԺ�� As סԺ��, b.����, b.�Ա�, b.����, b.Id As ҽ��id, b.ҽ������, b.ҽ������, d.Abo, d.Rh, d.Ѫ�����, d.Id As �շ�id, h.�������,c.ִ�в���id " & _
                       " From ������ҳ a, ����ҽ����¼ b, ѪҺ��Ѫ��¼ c, ѪҺ�շ���¼ d, " & _
                       "      (Select g.ҽ��id, f.������� " & _
                       "        From ������ϼ�¼ f, �������ҽ�� g " & _
                       "        Where f.Id = g.���id And f.����id = [1] And f.��ҳid = [2]) h " & _
                       " Where d.���� = 6 And d.�䷢id = c.Id And c.����id = b.Id and c.��¼����=1 And b.������� = 'K' And h.ҽ��id(+) = b.Id And Mod(d.��¼״̬, 3) = 1 And " & _
                       "       d.����� Is not Null And b.������� = 'K' And a.����id(+) = b.����id And a.��ҳid(+) = b.��ҳid And b.����id = [1] "
            If mlng������Դ = 2 Then 'סԺ����
                strSqlBR = strSqlBR & " And b.��ҳid =[2]"
                Set mRsBR = gobjDatabase.OpenSQLRecord(strSqlBR, "������Ϣ", mlng����ID, mlng��ҳid)
            Else '���ﲡ��,
                strSqlBR = strSqlBR & " and b.�Һŵ� =[3]"
                Set mRsBR = gobjDatabase.OpenSQLRecord(strSqlBR, "������Ϣ", mlng����ID, mlng��ҳid, mstr�Һŵ�)
            End If
            cbo1.Clear
            cbo2.Clear
            If mRsBR.RecordCount > 0 Then
                For lngj = 0 To mRsBR.RecordCount - 1
                    cbo1.AddItem mRsBR.Fields("Ѫ�����").Value   'Ѫ�����
                    cbo2.AddItem mRsBR.Fields("�շ�id").Value '�շ�id
                    mRsBR.MoveNext
                Next
                mRsBR.MoveFirst
            End If
            Call ExecuteCommand("�ؼ�����")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "��ȡ��Ӧ��¼"
            Dim lngFilter
            Dim ArrTime
            Dim str�Ƿ��� As String
            Dim lng����id As Long
            Dim str��ʼʱ�� As String
            Dim str����ʱ�� As String
            Dim int��Ѫ��Ӧ As Integer  '��Ϊ0����Ϊ1������Ϊ2
            
            If mstrFilter <> "" Then
                int��Ѫ��Ӧ = Int(marrFilter(4))
                lngFilter = marrFilter(3) '�ύ״̬
                str�Ƿ��� = marrFilter(2) '��¼��
                lng����id = marrFilter(0) '����id
                str��ʼʱ�� = Split(marrFilter(1), "'")(0)
                str����ʱ�� = Split(marrFilter(1), "'")(1)
            Else
                lngFilter = 0
                str�Ƿ��� = ""
                lng����id = 0
                str��ʼʱ�� = Now
                str����ʱ�� = Now
            End If
            '��ȡ���˵���Ѫ��Ӧ���ݣ���Ҫ�Ǵ���Ѫ��Ӧ��¼�л�ȡ
            'ȥ������ǰ�Ĳ��Ź���������û��Ҫ�ڷ�Ӧ�����
            strSqlFY = " Select distinct d.�շ�id, Nvl(d.��Ӧʱ��,d.��¼ʱ��) ��Ӧʱ��, d.��Ѫʷ, d.��Ѫ����, d.����ʷ, d.��Ѫ��Ŀ, d.������, d.�����߹�ϵ, d.����ʱ��, d.ת��, d.������Ӧ, d.��Ӧ���, d.���Ҵ����ʶ, d.���Ҵ����ʩ," & _
                       " d.��¼�� , d.��¼ʱ��, d.״̬, d.Ѫ�⴦���ʩ, d.ȷ����, d.ȷ��ʱ��,decode(d.������Ѫ��Ӧ,0,'',1,'��',2,'��') as ������Ѫ��Ӧ,decode(d.�Ƿ���Ѫ������,1,1,0) as �Ƿ���Ѫ������ " & _
                       " From ����ҽ����¼ a, ѪҺ��Ѫ��¼ b, ѪҺ�շ���¼ c, ��Ѫ��Ӧ��¼ d" & _
                       " Where d.�շ�id = c.id  And c.�䷢id = b.ID and mod(c.��¼״̬,3)=1 and c.����� is not null And b.����id = a.ID and b.��¼����=1 and a.�������='K' And a.����ID = [1] "
            If mlng������Դ = 2 Then 'סԺ����
                If lngFilter = 0 And mlng�׶� = 1 Then 'ȫ������,ҽ���׶�
                    strSqlFY = strSqlFY & "and a.��ҳid=[2] "
                ElseIf lngFilter = 0 And mlng�׶� = 2 Then 'ȫ������,��Ѫ�ƽ׶� �������Ѫ����������Ȩ�޵�����ѯ������
                    strSqlFY = strSqlFY & "and a.��ҳid=[2]  and (d.״̬<>0 or d.�Ƿ���Ѫ������=1 )"
                ElseIf lngFilter = 1 And mlng�׶� = 1 Then 'δ�ύ����,ҽ��
                    strSqlFY = strSqlFY & "and a.��ҳid=[2] and d.״̬=0 "
                ElseIf lngFilter = 1 And mlng�׶� = 2 Then 'δ�ύ����,��Ѫ��
                    strSqlFY = strSqlFY & "and a.��ҳid=[2] and (d.״̬<>2 and d.�Ƿ���Ѫ������=1 Or d.״̬=1)"
                ElseIf lngFilter = 2 And mlng�׶� = 1 Then '���ύ���ݣ�ҽ��
                    strSqlFY = strSqlFY & "and a.��ҳid=[2] and d.״̬ <>0 "
                ElseIf lngFilter = 2 And mlng�׶� = 2 Then '���ύ���ݣ���Ѫ��
                    strSqlFY = strSqlFY & "and a.��ҳid=[2] and d.״̬=2 "
                End If
            Else
                If lngFilter = 0 And mlng�׶� = 1 Then 'ȫ������,ҽ���׶�
                    strSqlFY = strSqlFY & "and a.�Һŵ�=[7] "
                ElseIf lngFilter = 0 And mlng�׶� = 2 Then 'ȫ������,��Ѫ�ƽ׶�
                    strSqlFY = strSqlFY & "and a.�Һŵ�=[7]  and (d.״̬<>0 or d.�Ƿ���Ѫ������=1 )"
                ElseIf lngFilter = 1 And mlng�׶� = 1 Then 'δ�ύ����,ҽ��
                    strSqlFY = strSqlFY & "and a.�Һŵ�=[7] and d.״̬=0 "
                ElseIf lngFilter = 1 And mlng�׶� = 2 Then 'δ�ύ����,��Ѫ��
                    strSqlFY = strSqlFY & "and a.�Һŵ�=[7] and (d.״̬<>2 and d.�Ƿ���Ѫ������=1 Or d.״̬=1)"
                ElseIf lngFilter = 2 And mlng�׶� = 1 Then '���ύ���ݣ�ҽ��
                    strSqlFY = strSqlFY & "and a.�Һŵ�=[7] and d.״̬ <>0 "
                ElseIf lngFilter = 2 And mlng�׶� = 2 Then '���ύ���ݣ���Ѫ��
                    strSqlFY = strSqlFY & "and a.�Һŵ�=[7] and d.״̬=2 "
                End If
            End If
            
            If marrFilter(2) <> "" Then
                strSqlFY = strSqlFY & IIf(mlng�׶� = 2, " And (d.ȷ����=[3] or d.�Ƿ���Ѫ������ =1 And d.��¼��=[3]) ", " and d.��¼��=[3] ")
            End If
            
            '��Ѫ��Ҫ��ʱ������޷�Ӧ���˷�Ӧ��¼
            If mlng�׶� = 2 Then
                If int��Ѫ��Ӧ = 0 Then
                    strSqlFY = strSqlFY & " and d.������Ѫ��Ӧ = 2 "
                ElseIf int��Ѫ��Ӧ = 1 Then
                    strSqlFY = strSqlFY & " and d.������Ѫ��Ӧ = 1 "
                ElseIf int��Ѫ��Ӧ = 3 Then
                    strSqlFY = strSqlFY & " and d.������Ѫ��Ӧ = 0 "
                End If
                
                If str��ʼʱ�� <> "" And str����ʱ�� <> "" Then
                    strSqlFY = strSqlFY & " and d.��Ӧʱ�� Between [5] and [6] "
                End If
            End If
                strSqlFY = strSqlFY & " order by Nvl(d.��Ӧʱ��,d.��¼ʱ��) "
            Set mRsFY = gobjDatabase.OpenSQLRecord(strSqlFY, "������Ѫ��Ӧ��¼", mlng����ID, mlng��ҳid, str�Ƿ���, lng����id, CDate(str��ʼʱ��), CDate(str����ʱ��), mstr�Һŵ�)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "�ύ����"
            Dim lng�շ�ID As Long, lng�ⷿid As Long
            Dim lng����ID As Long, lng����ID As Long
            lng�շ�ID = Val(cbo2.Text)
            If cbo2.Text = "" Then MsgBox "δѡ��Ҫ�ύ����˵���Ѫ��Ӧ��¼!", vbInformation, gstrSysName: ExecuteCommand = False: Exit Function
            
            StrSqlSAD = "Zl_��Ѫ��Ӧ��¼_Submit(" & lng�շ�ID & "," & mlng�׶� & "," & mlng״̬ & ",'" & UserInfo.���� & "'," & IIf(mbln��Ѫ������ = True, 1, 0) & ")" '���˸�
            Call SQLRecordAdd(rsSAD, StrSqlSAD)
            If mlng�׶� = 1 And cboHave.Text = "��" Then
                lng�ⷿid = Val(mRsBR!ִ�в���ID)
                If mlng������Դ = 2 Then
                    StrSqlSAD = "select ��Ժ����ID,��ǰ����ID from ������ҳ where ����id = [1] and ��ҳid = [2] "
                    Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "��ȡ�ⷿid", mlng����ID, mlng��ҳid)
                    If Not rsTmp.EOF Then lng����ID = Val(rsTmp!��ǰ����id): lng����ID = Val(rsTmp!��Ժ����ID)
                Else
                    StrSqlSAD = "select ִ�в���id from ���˹Һż�¼ where ����id = [1] and id = [2] "
                    Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "��ȡ�ⷿid", mlng����ID, mlng��ҳid)
                    If Not rsTmp.EOF Then lng����ID = Val(rsTmp!ִ�в���ID): lng����ID = Val(rsTmp!ִ�в���ID)
                End If
                StrSqlSAD = "Zl_ҵ����Ϣ�嵥_Insert(" & mlng����ID & "," & mlng��ҳid & ","  '����id ����id
                StrSqlSAD = StrSqlSAD & Val(lng����ID) & ","     '�������id
                StrSqlSAD = StrSqlSAD & Val(lng����ID) & ","      '���ﲡ��id
                StrSqlSAD = StrSqlSAD & mlng������Դ & ","                                      '������Դ
                StrSqlSAD = StrSqlSAD & "'���µ���Ѫ��Ӧ��Ҫ����','"             '��Ϣ����
                StrSqlSAD = StrSqlSAD & IIf(Val(mlng������Դ) = 1, "0000", "0000") & "','ZLHIS_BLOOD_008',"     ' ���ѳ���, ���ͱ���
                StrSqlSAD = StrSqlSAD & "'" & lng�շ�ID & "',"                      'ҵ���ʶ���շ�id��
                StrSqlSAD = StrSqlSAD & "1,0,NULL,'" & Val(lng�ⷿid) & "',NULL)"
                Call SQLRecordAdd(rsSAD, StrSqlSAD)
            End If
            Call SQLRecordExecute(rsSAD)
            Call ExecuteCommand("ˢ��ȫ������")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "��������"
            '����ָ��id�����ݣ�delete���̰����˻��˵Ĳ������������ö�л�洢����
            StrSqlSAD = "Zl_��Ѫ��Ӧ��¼_Submit(" & Val(cbo2.Text) & "," & mlng�׶� & "," & mlng״̬ & ",'" & UserInfo.���� & "'," & IIf(mbln��Ѫ������ = True, 1, 0) & ")" '���˸�
            Call SQLRecordAdd(rsSAD, StrSqlSAD)
            '���ø��շ�id����ϢΪ�Ѷ�
            If mlng�׶� = 1 Then
                If mlng������Դ = 2 Then
                    StrSqlSAD = "select ��Ժ����ID ����id from ������ҳ where ����id = [1] and ��ҳid = [2] "
                Else
                    StrSqlSAD = "select ִ�в���id ����id from ���˹Һż�¼ where ����id = [1] and id = [2] "
                End If
                Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "��ȡ�ⷿid", mlng����ID, mlng��ҳid)
                StrSqlSAD = "Zl_ҵ����Ϣ�嵥_Read(" & mlng����ID & "," & mlng��ҳid & ",'ZLHIS_BLOOD_008',"
                StrSqlSAD = StrSqlSAD & IIf(mlng������Դ = 2, 2, 1) & ",'" & UserInfo.���� & "'," & Val(Nvl(rsTmp!����ID & "")) & ",NULL,"
                StrSqlSAD = StrSqlSAD & "NULL," & Val(cbo2.Text) & ")"
                Call SQLRecordAdd(rsSAD, StrSqlSAD)
            End If
            Call SQLRecordExecute(rsSAD)
            Call ExecuteCommand("ˢ��ȫ������")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "����ҳ��"
             
            Call Clear
            mlngSum = mlngSum + 1
            mlng״̬ = 0 '����������Ĭ��Ϊ���ύ
            Call ExecuteCommand("��ʼ�ؼ�")
            mDataChanged = True
            If m_TabsPosition = PosiTop Then '������״̬��Ϊ����
                lbl5(mlngSelNum).Tag = "0:0:1" 'tag�����ŵĸ�ʽΪ"״̬������"������ҳ��Ĭ����״̬Ϊ0������������,datachange��true
                mstrST = ����
            Else
                lbl6(mlngSelNum).Tag = "0:0:1"
                mstrST = ����
            End If
            Call ExecuteCommand("�ؼ�״̬")
            If m_TabsPosition = PosiTop Then '������״̬��Ϊ����
                pic5_GotFocus (mlngSelNum) '.SetFocus
            Else
                Pic6_GotFocus (mlngSelNum) '(mlngSelNum).SetFocus
            End If
            Call ExecuteCommand("�ؼ�����")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "�ؼ�״̬"
            '���ݲ����޸ĸ����ؼ���״̬
            pic3.Visible = False
            For lngi = 0 To TXT11.Count - 1
                TXT11(lngi).locked = True
            Next
            For lngi = 0 To TXT21.Count - 1
                TXT21(lngi).locked = True
            Next
            TXT31.locked = True
            TXT32.locked = True
            TXT33.locked = True
            For lngi = 0 To TXT41.Count - 1
                TXT41(lngi).locked = True
            Next
            
            If (mstrST = ���� Or mstrST = �޸�) And mbln��Ѫ������ = True Then
                TXT41(1).locked = False
            End If
            Call CanbeChange(mlng�׶�, mlng״̬, mstrST)
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "��������"
            '�������ݷ�Ϊ�޸ı�����������棬������һ��sql�����а���������
            Dim int��Ѫʷ As Integer
            Dim str����ʷ As String
            Dim int�����߹�ϵ As Integer
            Dim str����ʱ�� As String
            Dim intת�� As Integer
            Dim str������Ӧ  As String
            Dim str��Ӧ��� As String
            Dim str���Ҵ����ʶ As String
            Dim lngselnum As Long
            Dim str��Ӧʱ�� As String
            Dim str��¼�� As String
            Dim strȷ���� As String
            Dim str��Ѫ�� As String
            Dim lng������Ѫ��Ӧ As Long
            Dim lng��Ѫ������ As Long '���˸�
            '�����ݽ��д�����������¼�����ݿ�
            
            If mlng�׶� = 2 Then
                StrSqlSAD = "select id from ҵ����Ϣ�嵥 where ����id = [1] and ����id = [2] and ҵ���ʶ = [3] and �Ƿ����� = 0 "
                Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "��ѯ��Ϣ��¼", mlng����ID, mlng��ҳid, Val(cbo2.Text))
                If rsTmp.RecordCount <> 0 Then
                    StrSqlSAD = "select �ⷿid ����id from ѪҺ�շ���¼ where id = [1] "
                    Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "��ȡ�ⷿid", Val(cbo2.Text))
                    StrSqlSAD = "Zl_ҵ����Ϣ�嵥_Read(" & mlng����ID & "," & mlng��ҳid & ",'ZLHIS_BLOOD_008',"
                    StrSqlSAD = StrSqlSAD & "5,'" & UserInfo.���� & "'," & Val(Nvl(rsTmp!����ID & "")) & ",NULL,"
                    StrSqlSAD = StrSqlSAD & "NULL,NULL)"
                    Call SQLRecordAdd(rsSAD, StrSqlSAD)
                End If
            End If
            If TXT11(1).Text = "��" Then
                str����ʷ = ":"
            Else
                str����ʷ = TXT21(1).Text & ":" & TXT21(2).Text
            End If
            
            int��Ѫʷ = Val(Opt21(0).Tag)
            int�����߹�ϵ = Val(Opt22(0).Tag) + 1 '���ܹ�ϵ
            
            intת�� = Val(Opt32(0).Tag) + 1 'ת�飺1-���� 2-���� 3-��
            '����Ѫ�����м��
            If Right(TXT21(4).Text, 1) = "." Then
                If Len(TXT21(4).Text) > 5 Then
                    MsgBox "���������ܴ���" & IIf(CboDW.Text = "U", "5U!", "1000ml!"), vbInformation, gstrSysName
                    TXT21(4).SetFocus
                    ExecuteCommand = False
                    Exit Function
                Else
                    str��Ѫ�� = TXT21(4).Text & "0" & CboDW.Text
                End If
            Else
                str��Ѫ�� = TXT21(4).Text & CboDW.Text
            End If
            If (CboDW.Text = "U" And Val(TXT21(4).Text) > 5) Then
                MsgBox "���������ܴ���5U!", vbInformation, gstrSysName
                TXT21(4).SetFocus
                ExecuteCommand = False
                Exit Function
            End If
            If (CboDW.Text = "ml" And Val(TXT21(4).Text) > 1000) Then
                MsgBox "���������ܴ���1000ml!", vbInformation, gstrSysName
                TXT21(4).SetFocus
                ExecuteCommand = False
                Exit Function
            End If
            If (CboDW.Text = "������" And Val(TXT21(4).Text) > 5) Then
                MsgBox "���������ܴ���5��������!", vbInformation, gstrSysName
                TXT21(4).SetFocus
                ExecuteCommand = False
                Exit Function
            End If

            If Opt31(0).Value = True Then
                str����ʱ�� = "��Ѫ�ڼ�"
            ElseIf Opt31(1).Value = True Then
                If TXT31.Text <> "" Then
                    str����ʱ�� = IIf(Right(TXT31.Text, 1) = "/", Left(TXT31.Text, Len(TXT31.Text) - 1), TXT31.Text) 'ȥ�����Ҳ��/
                    str����ʱ�� = IIf(Left(TXT31.Text, 1) = "/", Right(TXT31.Text, Len(TXT31.Text) - 1), TXT31.Text) 'ȥ��������/
                Else
                    str����ʱ�� = ""
                End If
            Else
                str����ʱ�� = "��"
            End If
            
            str������Ӧ = ""
            For lngi = 0 To Chk31.Count - 1 '��"����"���ⱻѡ������
                If Chk31(lngi).Value = Checked Then
                    str������Ӧ = str������Ӧ & Chk31(lngi).Caption & "," '���˸� 2017��6��8�� �������ݲ�����
                End If
            Next
            If str������Ӧ <> "" Then
                str������Ӧ = Left(str������Ӧ, Len(str������Ӧ) - 1)
            End If
            
'            If Chk31(������״).Value = Checked Then
'                str������Ӧ = str������Ӧ & "99:"
'            End If

            str��Ӧ��� = ""
            For lngi = 0 To Chk32.Count - 1 '��"����"���ⱻѡ������
                If Chk32(lngi).Value = Checked Then
                    str��Ӧ��� = str��Ӧ��� & Chk32(lngi).Caption & ","
                End If
            Next
            If str��Ӧ��� <> "" Then
                str��Ӧ��� = Left(str��Ӧ���, Len(str��Ӧ���) - 1)
            End If
            
'            If Chk32(�������).Value = Checked Then
'                str��Ӧ��� = str��Ӧ��� & "99:"
'            End If
            
            str���Ҵ����ʶ = ""
            For lngi = 0 To Chk33.Count - 1
                If Chk33(lngi).Value = Checked Then
                    str���Ҵ����ʶ = str���Ҵ����ʶ & Split(Chk33(lngi).Caption, ".")(1) & ","
                End If
            Next
            If str���Ҵ����ʶ <> "" Then
                str���Ҵ����ʶ = Left(str���Ҵ����ʶ, Len(str���Ҵ����ʶ) - 1)
            End If
            
            
            If m_TabsPosition = PosiTop Then
                lbl5(mlngSelNum).Caption = Msk51(mlngSelNum) & " " & Msk52(mlngSelNum)
                If IsDate(lbl5(mlngSelNum).Caption) = False Then MsgBox "ʱ���ʽ����!��˶�", vbInformation, gstrSysName: Exit Function '��ʱ���ʽ�����ж�
                str��Ӧʱ�� = lbl5(mlngSelNum).Caption
            ElseIf m_TabsPosition = Posiright Then
                lbl6(mlngSelNum).Caption = Msk61(mlngSelNum) & " " & Msk62(mlngSelNum)
                If IsDate(lbl6(mlngSelNum).Caption) = False Then MsgBox "ʱ���ʽ����!��˶�", vbInformation, gstrSysName: Exit Function
                str��Ӧʱ�� = lbl6(mlngSelNum).Caption
            End If
            
            If mlng�׶� = 1 Then
                str��¼�� = IIf(TXT41(1).Text = "", UserInfo.����, TXT41(1).Text) '�������ʱ�û�û��¼��ҽ����������Ĭ�ϵ�ǰ�û�����
                strȷ���� = TXT41(2).Text
            ElseIf mlng�׶� = 2 Then
                If mbln��Ѫ������Ȩ�� = True And mstrST = ���� Then
                    str��¼�� = IIf(TXT41(1).Text = "", UserInfo.����, TXT41(1).Text)
                    strȷ���� = TXT41(2).Text
                Else
                    str��¼�� = TXT41(1).Text
                    strȷ���� = UserInfo.����
                End If
            End If

            lng������Ѫ��Ӧ = IIf(cboHave.Text = "", 0, IIf(cboHave.Text = "��", 1, 2))
            lng��Ѫ������ = IIf((mlng�׶� = 2 And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = True And mlng�׶� = 2), 1, 0) '   ���˸�
            '�����Ѫ��Ӧ��ʩ���ٴ������ʩ�����Ƿ��������Ҫ��
            If gobjCommFun.StrIsValid(TXT32.Text, 500) = False Then ExecuteCommand = False: Exit Function
            If gobjCommFun.StrIsValid(TXT33.Text, 500) = False Then ExecuteCommand = False: Exit Function
            '���˸�
            If cbo2.Text <> "" And TXT21(5).Text <> "" Then
                StrSqlSAD = "Zl_��Ѫ��Ӧ��¼_Insert(" & Val(cbo2.Text) & ",to_date('" & str��Ӧʱ�� & "','yyyy-mm-dd hh24:mi:ss')," & int��Ѫʷ & "," & Val(TXT21(0).Text) & ",'" & str����ʷ & "','" & TXT21(3).Text & "','" & str��Ѫ�� & "'," & int�����߹�ϵ & _
                             ",'" & str����ʱ�� & "'," & intת�� & ",'" & str������Ӧ & "','" & str��Ӧ��� & "','" & str���Ҵ����ʶ & "','" & TXT32.Text & "','" & str��¼�� & _
                             "'," & mlng״̬ & ",'" & TXT33.Text & "','" & strȷ���� & "'," & lng������Ѫ��Ӧ & "," & lng��Ѫ������ & ")"
                Call SQLRecordAdd(rsSAD, StrSqlSAD)
                ExecuteCommand = SQLRecordExecute(rsSAD)
            Else
                ShowCancel '��������ʱ�������ձ��������������������û�йؼ�����Ѫ����ţ��շ�id�ȣ����ֱ�������Ч�ģ���ʱ����Զ�ɾ��ҳ�棬һ��Ӱ����������
                ExecuteCommand = True
                Exit Function
            End If
            lngselnum = mlngSelNum
            
            If mstrST = ���� Then
                mblnAddPage = False
            End If
            Call ExecuteCommand("ˢ��ȫ������")

            If m_TabsPosition = PosiTop Then '�������ݺ�Ҫ�۽����������ݵ�ѡ��ϣ�ʹ��pic5.setfocusû��Ч������������ֱ�ӵ���pic5_gotFocus
                pic5_GotFocus (lngselnum)
            ElseIf m_TabsPosition = Posiright Then
                Pic6_GotFocus (lngselnum)
            End If
            mDataChanged = False
            
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "ˢ��ȫ������"
            'ˢ������ҳ��
            Call ExecuteCommand("��ʼ������Ϣ") '1����ѯ��������ݷ������ݼ��У�����������ؼ�������Դ
            Call ExecuteCommand("��ȡ��Ӧ��¼")
            If mRsFY.BOF = False Then
                For lngi = 0 To mRsFY.RecordCount - 1 '���ܹ���ѯ���������ʱ
                    mlngSum = lngi
                    mlng״̬ = mRsFY.Fields(16).Value
                    Call ExecuteCommand("��ʼ�ؼ�") '2����������Դ�е����ݳ�ʼ�������ؼ�,����Ҫ��Ӷ��ٸ�ѡ���ѡ��е�������ʲô��ѡ��Ǹ���ʱ���������Ҫע�⣬���û��������ô���ٻ���Ҫ��һ��ѡ�������Ϊ��
                    cbo2.Text = mRsFY.Fields(0).Value
                    mRsFY.MoveNext
                Next
                
                mlngSelNum = mlngSum
                mRsFY.MoveFirst
            Else
                mlngSum = 0
                mlng״̬ = 0
                Call ExecuteCommand("��ʼ�ؼ�")
            End If
            Call Clear 'ɾ�����ҳ���������ҳ�����Ϣ
            Call ExecuteCommand("ˢ������") '3������Ӧʱ�������Ͷ�뵽ѡ��У�ѡ���ѡ��ѡ�Ĭ��Ϊѡ�1����ʾ����
            Call ExecuteCommand("�ؼ�״̬") '4��Ĭ�����пؼ����ɱ༭
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Case "���˻�������"
            '��ͨ��"��ʼ������Ϣ"���ܲ�ѯ������ʱ��˵���ò���û����Ѫ��¼����ʱ��ֻ��Ҫ��ʾ���˵Ļ�����Ϣ��
            Dim strSqlBasic As String
            Dim rsBasic As ADODB.Recordset
            For lngi = 0 To TXT11.Count - 1
                TXT11(lngi).Text = ""
            Next
            If mlng������Դ = 2 Then
            strSqlBasic = " select a.����,a.�Ա�,a.����,a.סԺ�� || '' as סԺ�� from ������ҳ a where a.����id=[1] and a.��ҳid=[2] "
            Else
                strSqlBasic = "select b.����,b.�Ա�,b.����,'' as סԺ�� from ���˹Һż�¼ b where b.����id=[1] and b.id=[2]"
            End If
            Set rsBasic = gobjDatabase.OpenSQLRecord(strSqlBasic, "���˻�����Ϣ", mlng����ID, mlng��ҳid)
            
            If rsBasic.RecordCount > 0 Then
                TXT11(0).Text = rsBasic.Fields("����").Value & ""
                TXT11(1).Text = rsBasic.Fields("�Ա�").Value & ""
                TXT11(2).Text = rsBasic.Fields("����").Value & ""
                TXT11(3).Text = rsBasic.Fields("סԺ��").Value & ""
            End If
        Case "ɾ������"
            If cbo2.Text = "" Then MsgBox "δѡ��Ӧ��¼!", vbInformation, gstrSysName: Exit Function
'            If Not mRsFY Is Nothing Then
'                If mRsFY.RecordCount = 0 Then Exit Function
'            End If
            'ɾ��ָ��id������
            StrSqlSAD = "Zl_��Ѫ��Ӧ��¼_delete(" & Val(cbo2.Text) & "," & mlng�׶� & "," & mlng״̬ & ")"
            Call SQLRecordAdd(rsSAD, StrSqlSAD)
            Call SQLRecordExecute(rsSAD)
            
            If mlngSum > 0 Then
                If m_TabsPosition = PosiTop Then
                    Unload lbl5(mlngSum)
                    Unload DTP5(mlngSum)
                    Unload Msk51(mlngSum)
                    Unload Msk52(mlngSum)
                    Unload pic51(mlngSum)
                    Unload Pic5(mlngSum)
                Else
                    Unload lbl6(mlngSum)
                    Unload DTP6(mlngSum)
                    Unload Msk61(mlngSum)
                    Unload Msk62(mlngSum)
                    Unload pic61(mlngSum)
                    Unload Pic6(mlngSum)
                End If
                
                mlngSum = mlngSum - 1
                If mblnAddPage = True Then
                    mblnAddPage = False
                End If
                Call ExecuteCommand("ˢ��ȫ������")
            Else
                Call ExecuteCommand("ˢ��ȫ������")
            End If
            
            If m_TabsPosition = PosiTop Then
                pic5_GotFocus (mlngSum) 'pic5(mlngSum).SetFocus
            ElseIf m_TabsPosition = Posiright Then
                Pic6_GotFocus (mlngSum) 'Pic6(mlngSum).SetFocus
            End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "ˢ������"
            Dim strABO As String
            Dim strRH As String
            Dim strZD As String
            '�����ͬtabʱ������ˢ�²���ʾ��ҳ����,��ʵ���ǽ����ݼ��ж�Ӧ������Ͷ�ŵ�ҳ����
            If mblnHaveData = False Then ExecuteCommand = False: Exit Function '��Ѫ��Ӧ��¼������ʱ
            If mRsFY Is Nothing Then ExecuteCommand = False: Exit Function
            ClearTag
            '��ȡ���˵�abo��rh
            Set rsTmp = GetPatientOtherInfo(mlng����ID, "ABO")
            If rsTmp.BOF = False Then strABO = rsTmp("��Ϣֵ").Value
            Set rsTmp = GetPatientOtherInfo(mlng����ID, "RH")
            If rsTmp.BOF = False Then strRH = rsTmp("��Ϣֵ").Value
            cboHave.Text = "" '���cbohave�е�����
            If mRsFY.RecordCount = 0 Then '����������Ѫ��Ӧ��¼ʱ����ʾ���˵Ļ�����Ϣ������״̬��Ϊ�������Է����û����
                mlng״̬ = 0 '���û�в�ѯ����Ӧ��¼��������Ĭ���Ǵ��ύ������
                If mRsBR.EOF = False Then '2��3��4��1��8��9
                     TXT11(0).Text = mRsBR.Fields("����").Value & ""
                     TXT11(1).Text = mRsBR.Fields("�Ա�").Value & ""
                     TXT11(2).Text = mRsBR.Fields("����").Value & ""
                     TXT11(3).Text = mRsBR.Fields("סԺ��").Value & ""
                     TXT11(4).Text = IIf(mRsBR.Fields("ABO").Value & "" = "", strABO, mRsBR.Fields("ABO").Value & "")
                     TXT11(5).Text = IIf(mRsBR.Fields("RH").Value & "" = "", strRH, mRsBR.Fields("RH").Value & "")
                End If
'                mDataChanged = True '���� 2017��6��22��16:17:41  ҽ������վ����Ѫ��Ӧ�������û����Ѫ��Ӧ��¼ʱ���޷������Ѫ��Ӧ��
                If m_TabsPosition = PosiTop Then '��û����Ѫ��Ӧ��¼�򽫳�ʼҳ���״̬��Ϊ�������������������ݶ�ɾ�����������޷�Ӧ��¼�����
                    mstrST = ����
                    lbl5(mlngSelNum).Tag = "0:" & mstrST & ":0"
                Else
                    mstrST = ����
                    lbl6(mlngSelNum).Tag = "0:" & mstrST & ":0"
                End If
                
            Else '�з�Ӧ��¼ʱ
                With mRsFY
                    For lngi = 0 To .RecordCount - 1 '��������
                        If .EOF Then Exit For
                        If m_TabsPosition = PosiTop Then
                            lbl5(lngi).Caption = Format(.Fields("��Ӧʱ��").Value & "", "YYYY-MM-DD HH:mm:ss")
                            Msk51(lngi).Text = Format(Nvl(.Fields("��Ӧʱ��").Value, "____-__-__"), "YYYY-MM-DD")
                            Msk52(lngi).Text = Format(Nvl(.Fields("��Ӧʱ��").Value, "__:__:__"), "HH:mm:ss")
                            Msk51(lngi).Tag = Val(.Fields("�Ƿ���Ѫ������").Value & "") '���˸�
                        Else
                            lbl6(lngi).Caption = Format(.Fields("��Ӧʱ��").Value & "", "YYYY-MM-DD HH:mm:ss")
                            Msk61(lngi).Text = Format(Nvl(.Fields("��Ӧʱ��").Value, "____-__-__"), "YYYY-MM-DD")
                            Msk62(lngi).Text = Format(Nvl(.Fields("��Ӧʱ��").Value, "__:__:__"), "HH:mm:ss")
                            Msk61(lngi).Tag = Val(.Fields("�Ƿ���Ѫ������").Value & "")
                        End If
                        If mlngSelNum > 0 Then
                            If lngi = mlngSelNum - 1 Then 'Ϊ��������ҳ�����ʹҳ��������ʾ
                                Call Clear
                            End If
                        End If
                        If lngi = mlngSelNum Then 'ѡ��ѡ������
                            If mRsBR.BOF = False Then
                                 TXT11(0).Text = mRsBR.Fields("����").Value & ""  '����
                                 TXT11(1).Text = mRsBR.Fields("�Ա�").Value & "" '�Ա�
                                 TXT11(2).Text = mRsBR.Fields("����").Value & "" '����
                                 TXT11(3).Text = mRsBR.Fields("סԺ��").Value & "" 'סԺ��
                                 TXT11(4).Text = IIf(mRsBR.Fields("ABO").Value & "" = "", strABO, mRsBR.Fields("ABO").Value & "")
                                 TXT11(5).Text = IIf(mRsBR.Fields("RH").Value & "" = "", strRH, mRsBR.Fields("RH").Value & "")
                            End If
                            '����  2017��2��8��
                            strXD = " Select e.Id, b.����, e.Ѫ�����, c.������Ѫʷ, c.�в���� " & _
                                    " From ����ҽ����¼ a, ������ĿĿ¼ b, ѪҺ��� g,��Ѫ�����¼ c, ѪҺ��Ѫ��¼ d, ѪҺ�շ���¼ e " & _
                                    " Where e.�䷢id = d.Id And d.����id = a.Id and g.Ʒ��id = b.Id AND g.���id = e.ѪҺid AND " & _
                                    " d.��¼����=1 And c.ҽ��id(+) =  a.Id " & _
                                    " And a.���id Is Null and Mod(e.��¼״̬, 3) = 1 And e.����� Is not Null and e.id=[1] "
                            Set mrsXD = gobjDatabase.OpenSQLRecord(strXD, "��ѯѪ����ŵ�", Val(.Fields("�շ�id").Value))

                            If mrsXD.BOF = False Then
                                TXT21(5).Text = mrsXD.Fields("Ѫ�����").Value
                                cbo2.Text = mrsXD.Fields("id").Value
                                TXT21(3).Text = mrsXD.Fields("����").Value '��Ѫ��Ŀ
                            End If
                            
                            TXT21(0).Text = IIf(.Fields("��Ѫ����").Value = 0, "", .Fields("��Ѫ����").Value) '��Ѫ����
                            '�в����
                            If .Fields("����ʷ").Value & "" <> "" Then
                                TXT21(1).Text = Split(.Fields("����ʷ").Value, ":")(0) ' ��
                                TXT21(2).Text = Split(.Fields("����ʷ").Value, ":")(1)  '��
                            End If
                            
                            If InStr(.Fields("������").Value & "", "������") > 0 Then
                                TXT21(4).Text = Split(.Fields("������").Value & "", "������")(0)
                                CboDW.ListIndex = 2
                            ElseIf InStr(.Fields("������").Value & "", "U") > 0 Then
                                TXT21(4).Text = Split(.Fields("������").Value & "", "U")(0)
                                CboDW.ListIndex = 1
                            ElseIf InStr(UCase(.Fields("������").Value & ""), "ML") > 0 Then
                                TXT21(4).Text = Split(UCase(.Fields("������").Value & ""), "ML")(0)
                                CboDW.ListIndex = 0
                            End If
                            Opt21(0).Tag = Val(.Fields("��Ѫʷ").Value & "")
                            Opt21(Val(Opt21(0).Tag)).Value = True
                            Opt22(0).Tag = Val(Nvl(.Fields("�����߹�ϵ").Value, 1)) - 1
                            Opt22(Val(Opt22(0).Tag)).Value = True
                            If Val(.Fields("ת��").Value & "") < 3 Then
                                Opt32(0).Tag = Val(Nvl(.Fields("ת��").Value, 1)) - 1
                                Opt32(Val(Opt32(0).Tag)).Value = True
                            Else
                                Opt32(0).Tag = Val(.Fields("ת��").Value & "") - 1
                            End If
                            
                            cboHave.Text = .Fields("������Ѫ��Ӧ").Value & ""
                            
                            If .Fields("����ʱ��").Value & "" = "��Ѫ�ڼ�" Then '����ʱ��
                                Opt31(0).Tag = 0
                                Opt31(0).Value = True
                            ElseIf .Fields("����ʱ��").Value & "" = "��" Then
                                Opt31(0).Tag = 2
                            Else
                                Opt31(0).Tag = 1
                                Opt31(1).Value = True
                                TXT31.Text = .Fields(8).Value & ""  '����ʱ��
                            End If
                            
                            For lngk = 0 To Chk31.Count - 1 'ˢ������ʱ���Ƚ��ؼ���ѡ��״̬��ԭ
                                Chk31(lngk).Value = Unchecked
                            Next
                            If IsNull(.Fields("������Ӧ").Value) = False Then
                                StrSplit = Split(.Fields("������Ӧ").Value, ",")
                                For lngk = 0 To Chk31.Count - 1 '������Ӧ
                                    For lngj = 0 To UBound(StrSplit)
                                        If StrSplit(lngj) = Chk31(lngk).Caption Then '���� 2017��6��8�� ��Ӧ��¼�����������ݵ�����
                                            Chk31(lngk).Value = Checked
                                            Chk31(lngk).Tag = 1
                                        End If
                                    Next
                                Next
                            End If
                            
                            For lngk = 0 To Chk32.Count - 1
                                Chk32(lngk).Value = Unchecked
                            Next
                            If IsNull(.Fields("��Ӧ���").Value) = False Then
                                StrSplit = Split(.Fields("��Ӧ���").Value, ",")
                                For lngk = 0 To Chk32.Count - 1 '��Ӧ���
                                    For lngj = 0 To UBound(StrSplit)
                                        If StrSplit(lngj) = Chk32(lngk).Caption Then
                                            Chk32(lngk).Value = Checked
                                            Chk32(lngk).Tag = 1
                                        End If
                                    Next
                                Next
                            End If

                            For lngk = 0 To Chk33.Count - 1
                                Chk33(lngk).Value = Unchecked
                            Next
                            If IsNull(.Fields("���Ҵ����ʶ").Value) = False Then
                                StrSplit = Split(.Fields("���Ҵ����ʶ").Value, ",")
                                For lngk = 0 To Chk33.Count - 1 ''���Ҵ����ʶ
                                    For lngj = 0 To UBound(StrSplit)
                                        If StrSplit(lngj) = Split(Chk33(lngk).Caption, ".")(1) Then
                                            Chk33(lngk).Value = Checked
                                            Chk33(lngk).Tag = 1
                                        End If
                                    Next
                                Next
                            End If

                            mlng״̬ = Val(.Fields("״̬").Value & "") '��ȡ��Ӧ��¼��״̬
                            TXT32.Text = .Fields("���Ҵ����ʩ").Value & "" '���Ҵ����ʩ
                            TXT33.Text = .Fields("Ѫ�⴦���ʩ").Value & "" 'Ѫ�⴦���ʩ
                            TXT41(0).Text = GetZXR(.Fields("�շ�id")) & "" '��ʿ(ִ����)
                            TXT41(1).Text = .Fields("��¼��").Value & "" '��¼��
                            TXT41(2).Text = .Fields("ȷ����").Value & ""  'ȷ����
                            '��д��Ѫ����Ӧ��ҽ���������Ϣ
                            strZD = "select c.���� from ѪҺ��Ѫ��¼ a,ѪҺ�շ���¼ b,����ҽ������ c where c.ҽ��id=a.����id and c.��Ŀ='���뵥���' and a.id=b.�䷢id and b.id=[1]"
                            Set rsTmp = gobjDatabase.OpenSQLRecord(strZD, "��ѯ��ϼ�¼", Val(.Fields("�շ�id").Value))
                            If rsTmp.EOF = False Then
                                TXT11(6).Text = rsTmp.Fields("����").Value & ""
                            Else
                                TXT11(6).Text = ""
                            End If
                            mRsBR.MoveFirst
                        End If
                        .MoveNext
                    Next
                    .MoveFirst
                End With
                Call ExecuteCommand("�ؼ�����")
            End If
        End Select
    Next
    
    ExecuteCommand = True
    Exit Function
Error:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    ExecuteCommand = False
End Function
Private Sub cbo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub CboDW_Click()
    '������Ѫ��
    If CboDW.Text = "U" And Val(TXT21(4).Text) > 5 Then
        MsgBox "���������ܴ���5U!", vbInformation, gstrSysName
    End If
    If CboDW.Text = "ml" And Val(TXT21(4).Text) > 1000 Then
        MsgBox "���������ܴ���1000ml!", vbInformation, gstrSysName
    End If
End Sub

Private Sub CboDW_DropDown()
    '�������޸Ĺ����п����޸ĵ�λ�������������
    If mstrST = ���� Or mstrST = �޸� Then
        CboDW.locked = False
    Else
        CboDW.locked = True
    End If
    If mlng�׶� = 2 Then CboDW.locked = True '��Ѫ�ƽ׶β������޸ĵ�λ
End Sub

Private Sub cboHave_Click()
    '��շ�Ӧ��������
    Dim lngi As Long
'    Opt31(0).Value = True
'    Opt32(0).Value = True
    TXT31.Text = ""
    TXT32.Text = ""
    TXT33.Text = ""
    For lngi = 0 To Chk31.Count - 1
        Chk31(lngi).Value = Unchecked
    Next
    For lngi = 0 To Chk32.Count - 1
        Chk32(lngi).Value = Unchecked
    Next
    For lngi = 0 To Chk33.Count - 1
        Chk33(lngi).Value = Unchecked
    Next
    '�ı�ؼ�״̬
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub cboHave_DropDown()
    '�������޸Ĺ����п����޸ĵ�λ�������������
    If (mstrST = ���� And mDataChanged = True) Or mstrST = �޸� Then
        cboHave.locked = False
    Else
        cboHave.locked = True
    End If
    '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� �������޸�������Ѫ��Ӧ   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then cboHave.locked = True
End Sub

Private Sub cboHave_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
    gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub Chk31_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And mstrST <> ���� And mstrST <> �޸� Then KeyCode = 0: Exit Sub '�ų��������޸�����Ĳ���
    '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� ��������chk31�ؼ�������   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then KeyCode = 0: Exit Sub
'    If mstrST = �޸� And mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then KeyCode = 0: Exit Sub '�ų���Ѫ�ƽ׶�û������Ȩ��ʱ���޸Ĳ��� ���˸�
End Sub

Private Sub Chk31_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub ClearTag()
    Dim lngi As Long
    For lngi = 0 To Chk31.Count - 1
        Chk31(lngi).Tag = 0
    Next
    For lngi = 0 To Chk32.Count - 1
        Chk32(lngi).Tag = 0
    Next
    For lngi = 0 To Chk33.Count - 1
        Chk33(lngi).Tag = 0
    Next
End Sub

Private Sub Chk31_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '���޸�״̬��ֻ��mlng�׶�=1��mlng״̬=1ʱ�ſ����޸�,����״̬Ҳ�����޸ģ�����״̬���������޸�
    If Button = 2 Then Exit Sub '����Ҽ������ǿ��Ա��checkbox��״̬�������������Ҽ��������¼�
    '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� ��������chk31�ؼ�������   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then
'    If mstrST = �޸� And mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then '��Ѫ�Ƶ��chk31ʱ,���û����Ѫ������Ȩ���򣬲��������޸�   ���˸�
        Chk31(Index).Value = Val(Chk31(Index).Tag)
        Exit Sub
    End If
    If mstrST = �޸� Or mstrST = ���� Then
        Chk31(Index).Tag = Chk31(Index).Value
    Else
        Chk31(Index).Value = Val(Chk31(Index).Tag)
    End If
End Sub

Private Sub Chk32_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And mstrST <> ���� And mstrST <> �޸� Then KeyCode = 0: Exit Sub '�ų��������޸�����Ĳ���
    If mstrST = �޸� And mlng�׶� = 2 Then KeyCode = 0: Exit Sub '�ų���Ѫ�ƽ׶ε��޸Ĳ���
End Sub

Private Sub Chk32_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Chk32_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '���޸�״̬��ֻ��mlng�׶�=1��mlng״̬=1ʱ�ſ����޸�,����״̬Ҳ�����޸ģ�����״̬���������޸�
    If Button = 2 Then Exit Sub ''����Ҽ������ǿ��Ա��checkbox��״̬�������������Ҽ��������¼�
    '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� ��������chk32�ؼ�������   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then
'    If mstrST = �޸� And mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then '��Ѫ�Ƶ��chk32ʱ,���û����Ѫ������Ȩ���򣬲��������޸�   ���˸�
        Chk32(Index).Value = Val(Chk32(Index).Tag)
        Exit Sub
    End If
    If mstrST = �޸� Or mstrST = ���� Then
        Chk32(Index).Tag = Chk32(Index).Value
    Else
        Chk32(Index).Value = Val(Chk32(Index).Tag)
    End If
End Sub

Private Sub Chk33_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And mstrST <> ���� And mstrST <> �޸� Then KeyCode = 0: Exit Sub '�ų��������޸�����Ĳ���
'    If mstrST = �޸� And mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then KeyCode = 0: Exit Sub '�ų���Ѫ�ƽ׶���û������Ȩ���µ��޸Ĳ���  ���˸�
        '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� ��������chk33�ؼ�������   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then KeyCode = 0: Exit Sub
End Sub

Private Sub Chk33_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Chk33_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '���޸�״̬��ֻ��mlng�׶�=1��mlng״̬=1ʱ�ſ����޸�,����״̬Ҳ�����޸ģ�����״̬���������޸�
    If Button = 2 Then Exit Sub '����Ҽ������ǿ��Ա��checkbox��״̬�������������Ҽ��������¼�
    '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� ��������chk33�ؼ�������   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then
'    If mstrST = �޸� And mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then   '��Ѫ�Ƶ��chk33ʱ�����û������Ȩ�����������޸�  ���˸�
        Chk33(Index).Value = Val(Chk33(Index).Tag)
        Exit Sub
    End If
    If mstrST = �޸� Or mstrST = ���� Then
        Chk33(Index).Tag = Chk33(Index).Value
    Else
        Chk33(Index).Value = Val(Chk33(Index).Tag)
    End If
End Sub

Private Sub CboDW_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
    gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub vsbRight_Change()
    vsbRight_Scroll
End Sub
Private Sub vsbRight_Scroll()
    Dim i As Integer
    For i = 0 To Pic6.Count - 1
        Pic6(i).Top = Picright.ScaleTop + i * (450) - vsbRight.Value
    Next
End Sub
Private Sub lbl5_DblClick(Index As Integer)
    If (mstrST = ���� And mDataChanged = True) Or (mstrST = �޸� And mlng�׶� = 1 And mlng״̬ = 0) Or (mstrST = �޸� And mlng�׶� = 2 And mlng״̬ = 0 And mbln��Ѫ������Ȩ�� = True) Then '���˸�  ������Ȩ�޵���Ѫ�ƿ��Խ����޸Ĳ�����
'    If mstrST <> ���� Then Exit Sub
        Msk52(Index).Visible = True
        Msk51(Index).Visible = True
        Msk51(Index).Text = Format(Split(lbl5(Index).Caption, " ")(0), "YYYY-MM-DD")
        Msk52(Index).Text = Format(Split(lbl5(Index).Caption, " ")(1), "HH:mm:ss")
        Msk51(Index).ZOrder 0
        Msk52(Index).ZOrder 0
    End If
End Sub

Private Sub lbl6_DblClick(Index As Integer)
    If (mstrST = ���� And mDataChanged = True) Or (mstrST = �޸� And mlng�׶� = 1 And mlng״̬ = 0) Or (mstrST = �޸� And mlng�׶� = 2 And mlng״̬ = 0 And mbln��Ѫ������Ȩ�� = True) Then '���˸�  ������Ȩ�޵���Ѫ�ƿ��Խ����޸Ĳ����� '��������ҽ���׶�δ�ύ�������Ǵ����޸Ĳ���ʱ�����޸ķ�Ӧʱ��
    '    If mstrST <> ���� Then Exit Sub
        Msk62(Index).Visible = True
        Msk61(Index).Visible = True
        Msk61(Index).Text = Format(Split(lbl6(Index).Caption, " ")(0), "YYYY-MM-DD")
        Msk62(Index).Text = Format(Split(lbl6(Index).Caption, " ")(1), "HH:mm:ss")
        Msk61(Index).ZOrder 0
        Msk62(Index).ZOrder 0
    End If
End Sub

Private Sub pic3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim strXD As String
    Dim rsXD As Recordset
    Dim lngID As Long
    Dim str���� As String
    Dim strѪ����� As String
    Dim str�в���� As String
    Dim str������Ѫʷ As String
    Dim lngIndex As Long
    Dim str�����Ϣ As String
    Dim rsTmp As Recordset
    lngIndex = 5
    
    If Val(mlng����ID) = 0 Then MsgBox "�޲�����Ϣ��", vbInformation, gstrSysName: Exit Sub
    If (mstrST = �޸� And mlng�׶� = 1 And mlng״̬ = 0) Or (mstrST = ���� And mDataChanged = True) Then  'ֻ�е��û�Ϊҽ�����׶�Ϊ�޸Ļ�����ʱ������˫��txt21����
        pic3.Visible = True
        Call GetSecondUserName(TXT21(lngIndex), 1, mlng����ID, mlng��ҳid, mlng������Դ, lngID, str����, strѪ�����, str������Ѫʷ, str�в����)
        TXT21(lngIndex).Text = strѪ�����
        TXT21(3).Text = str����
        cbo1.Text = strѪ�����
        If lngID <> 0 Then
            cbo2.Text = lngID
        End If
        If str������Ѫʷ = "��" Then '������Ѫʷ
            Opt21(0).Value = True
            TXT21(0).Text = ""
        ElseIf str������Ѫʷ = "��" Then
            Opt21(1).Value = True
            TXT21(0).Text = ""
        End If
        If InStr(1, str�в����, "/") <= 0 Then '�в����
            TXT21(1).Text = ""
            TXT21(2).Text = ""
        Else
            TXT21(1).Text = IIf(Split(str�в����, "/")(0) = "" And Split(str�в����, "/")(1) <> "", 1, Split(str�в����, "/")(0))
            TXT21(2).Text = Split(str�в����, "/")(1) & ""
        End If
        TXT41(0).Text = GetZXR(lngID) '��ʿ(ִ����)
'        If Val(lngID) > 0 Then
'            '��д��Ѫ����Ӧ��ҽ���������Ϣ
'            str�����Ϣ = "select c.���� from ѪҺ��Ѫ��¼ a,ѪҺ�շ���¼ b,����ҽ������ c where c.ҽ��id=a.����id and c.��Ŀ='���뵥���' and a.id=b.�䷢id and b.id=[1]"
'            Set rsTmp = gobjDatabase.OpenSQLRecord(str�����Ϣ, "��ѯ��ϼ�¼", Val(lngID))
'            If rsTmp.EOF = False Then TXT11(6).Text = rsTmp.Fields("����").Value & ""
'        End If
    End If
End Sub
Private Sub DTP5_LostFocus(Index As Integer)
    Pic5(Index).BorderStyle = 0
End Sub

Private Sub DTP6_LostFocus(Index As Integer)
    Pic6(Index).BorderStyle = 0
End Sub

Private Sub lbl5_Click(Index As Integer)
    Dim lngi As Long
    For lngi = 0 To Msk51.Count - 1
        Msk52(lngi).Visible = False
        Msk51(lngi).Visible = False
    Next
     '����ҳ��Ҫ�޸�ʱ��ʱ��˫��֮ǰ����Ӧ����ʱ�䣬����ҳ��������գ��������һ���ж��Խ����������
    If mlngSelNum = Index And mstrST = ���� Then Exit Sub
    Pic5(Index).SetFocus
End Sub

Private Sub lbl6_Click(Index As Integer)
    Dim lngi As Long
    For lngi = 0 To Msk61.Count - 1
        Msk62(lngi).Visible = False
        Msk61(lngi).Visible = False
    Next
    If mlngSelNum = Index And mstrST = ���� Then Exit Sub
    Pic6(Index).SetFocus
End Sub

Private Sub Opt21_Click(Index As Integer)
    Dim lngTag As Long
    lngTag = Val(Opt21(0).Tag)
    '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� ��������opt21�ؼ�������   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then
'    If mstrST = �޸� And mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then '���˸�
        Opt21(lngTag).Value = True
        Exit Sub
    End If
    If mstrST = �޸� Or (mstrST = ���� And mDataChanged = True) Then
        Opt21(0).Tag = Index
    Else
        Opt21(lngTag).Value = True
    End If
    If Opt21(0).Value = True Then TXT21(0).Text = "" '��Ѫ����ѡ���ޣ��Զ���txt21(0)�����ݸ�Ϊ""
End Sub

Private Sub Opt21_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Opt22_Click(Index As Integer)
    'opt22���޸ķ�ʽ��opt21\31\32������ͬ
    Dim lngTag As Long
    lngTag = Val(Opt22(0).Tag)
    '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� ��������Opt22�ؼ�������   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then
'    If mstrST = �޸� And mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then '���˸�
        Opt22(lngTag).Value = True
        Exit Sub
    End If
    If mstrST = �޸� Or (mstrST = ���� And mDataChanged = True) Then
        Opt22(0).Tag = Index
    Else
        Opt22(lngTag).Value = True
    End If
End Sub

Private Sub Opt22_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Opt31_Click(Index As Integer)
    Dim lngTag As Long
    lngTag = Val(Opt31(0).Tag)
    '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� ��������Opt31�ؼ�������   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then
'    If mstrST = �޸� And mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then '���˸�
        If lngTag = 2 Then
            Opt31(0).Value = False
            Opt31(1).Value = False
        Else
            Opt31(lngTag).Value = True
        End If
        Exit Sub
    End If
    If mstrST = �޸� Or mstrST = ���� Then
        Opt31(0).Tag = Index
    ElseIf lngTag = 2 Then
        Opt31(0).Value = False
        Opt31(1).Value = False
        Exit Sub
    Else
        Opt31(lngTag).Value = True
    End If
    If Opt31(0).Value = True Then TXT31.Text = ""
End Sub

Private Sub Opt31_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Opt32_Click(Index As Integer)
    Dim lngTag As Long
    lngTag = Val(Opt32(0).Tag)
    '��Ѫ�ƽ׶�  ��������ʱû������Ȩ�� ���� �޸�����ʱ������Ѫ���������� ��������Opt32�ؼ�������   ���˸�
    If mlng�׶� = 2 And ((mbln��Ѫ������Ȩ�� = False And mstrST = ����) Or (mstrST = �޸� And lbl��Ѫ������.Visible = False)) Then
'    If mstrST = �޸� And mlng�׶� = 2 And mbln��Ѫ������Ȩ�� = False Then '���˸�
        If lngTag = 2 Then
            Opt32(0).Value = False
            Opt32(1).Value = False
        Else
            Opt32(lngTag).Value = True
        End If
        Exit Sub
    End If
    If mstrST = �޸� Or mstrST = ���� Then
        Opt32(0).Tag = Index
    ElseIf lngTag = 2 Then
        Opt32(0).Value = False
        Opt32(1).Value = False
        Exit Sub
    Else
        Opt32(lngTag).Value = True
    End If
End Sub

Private Sub Opt32_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Pic4_Resize()
    '���ù�����,���Թ�������λ�ý��е���
    On Error GoTo Errorhand
    VS1.Visible = False
    HS1.Visible = False
    If pic4.ScaleWidth < pic1.Width Then '����pic4�ĸ߶�С��pic1�ĸ߶ȵ����
        VS1.Left = pic4.ScaleWidth - VS1.Width
        VS1.Top = pic4.ScaleTop
        VS1.Height = pic4.ScaleHeight
    
        HS1.Left = pic4.ScaleLeft
        HS1.Top = pic4.ScaleHeight - HS1.Height

        If pic4.ScaleWidth <= pic1.ScaleWidth Then
            If pic4.ScaleWidth - VS1.Width > 0 Then
                HS1.Width = pic4.ScaleWidth - VS1.Width
            Else
                HS1.Width = pic4.ScaleWidth
            End If
        Else
            If pic4.ScaleWidth - VS1.Width > 0 Then
                HS1.Width = pic1.ScaleWidth - VS1.Width + 30
            Else
                HS1.Width = pic1.ScaleWidth
            End If
        End If
        
        HS1.Min = 1
        HS1.Max = pic1.Width - pic4.ScaleWidth + VS1.Width
        HS1.SmallChange = 300
        HS1.LargeChange = 300
        
        VS1.Min = 1
        VS1.Max = pic1.Height - pic4.ScaleHeight + HS1.Height
        VS1.SmallChange = 300
        VS1.LargeChange = 300
        
        VS1.Visible = True
        HS1.Visible = True
    End If
    If pic4.ScaleWidth >= pic1.Width And pic4.ScaleHeight < pic1.Height Then 'pic4�ĸ߶�С��pic1�ĸ߶ȵ����
        VS1.Left = pic1.Width + (pic4.Width - pic1.Width) / 2 'pic1.Left + pic1.Width - VS1.Width
        VS1.Top = pic4.ScaleTop
        VS1.Height = pic4.ScaleHeight
    
        HS1.Left = (pic4.Width - pic1.Width) / 2
        HS1.Top = pic4.ScaleHeight - HS1.Height
        If pic4.ScaleWidth <= pic1.ScaleWidth Then
            If pic4.ScaleWidth - VS1.Width > 0 Then
                HS1.Width = pic4.ScaleWidth - VS1.Width
            Else
                HS1.Width = pic4.ScaleWidth
            End If
        Else
            If pic4.ScaleWidth - VS1.Width > 0 Then
                HS1.Width = pic1.ScaleWidth
            Else
                HS1.Width = pic1.ScaleWidth
            End If
        End If
        
        HS1.Min = 1
        HS1.Max = VS1.Width
        HS1.SmallChange = 300
        HS1.LargeChange = 300
        
        VS1.Min = 1
        VS1.Max = pic1.Height - pic4.ScaleHeight + HS1.Height
        VS1.SmallChange = 300
        VS1.LargeChange = 300
        
        VS1.Visible = True
        HS1.Visible = False
    End If
Errorhand:
End Sub

Private Sub pic5_Click(Index As Integer)
    Pic5(Index).SetFocus
End Sub

Private Sub pic6_Click(Index As Integer)
    Pic6(Index).SetFocus
End Sub

Private Sub TXT11_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub TXT11_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(TXT11(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(TXT11(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub TXT11_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(TXT11(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub TXT21_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(TXT21(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(TXT21(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub TXT21_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(TXT21(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub TXT21_Validate(Index As Integer, Cancel As Boolean)
    '������Ѫ��
    If CboDW.Text = "U" And Val(TXT21(4).Text) > 5 Then
        MsgBox "���������ܴ���5U!", vbInformation, gstrSysName
    End If
    If CboDW.Text = "ml" And Val(TXT21(4).Text) > 1000 Then
        MsgBox "���������ܴ���1000ml!", vbInformation, gstrSysName
    End If
    If CboDW.Text = "������" And Val(TXT21(4).Text) > 5 Then
        MsgBox "���������ܴ���5��������!", vbInformation, gstrSysName
    End If
    If Right(TXT21(4).Text, 1) = "." And Len(TXT21(4).Text) < 6 Then TXT21(4).Text = TXT21(4).Text & "0"
End Sub

Private Sub TXT31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '����27554 by lesfeng 2010-01-19 lngTXTProc �޸�ΪglngTXTProc
        glngTXTProc = GetWindowLong(TXT31.hWnd, GWL_WNDPROC)
        Call SetWindowLong(TXT31.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub TXT31_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(TXT31.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub TXT32_Validate(Cancel As Boolean)
    TXT32.SelStart = 0
    TXT32.SelLength = Len(TXT32.Text)
    TXT32.SelBold = False
    TXT32.SelFontName = "����"
    TXT32.SelFontSize = 9
End Sub

Private Sub TXT33_Validate(Cancel As Boolean)
    TXT33.SelStart = 0
    TXT33.SelLength = Len(TXT33.Text)
    TXT33.SelBold = False
    TXT33.SelFontName = "����"
    TXT33.SelFontSize = 9
End Sub

Private Sub TXT41_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '����27554 by lesfeng 2010-01-19 lngTXTProc �޸�ΪglngTXTProc
        glngTXTProc = GetWindowLong(TXT41(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(TXT41(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub TXT41_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(TXT41(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub TXT21_GotFocus(Index As Integer)
    If mlng�׶� = 2 Then
        If mbln��Ѫ������Ȩ�� = False Or mstrST = �޸� And lbl��Ѫ������.Visible = False Then Exit Sub   '��Ѫ�ƽ׶�û����Ѫ������Ȩ���򲻻�����ѡ��Ѫ���İ�ť   ���˸�
    End If
    'ֻ������״̬������ѡ��Ѫ����ť
    If (mstrST = ���� Or mstrST = �޸�) And Index = 5 And mDataChanged = True Then   '(mstrST = �޸� And mlng�׶� = 1 And mlng״̬ = 0) Or
        pic3.Visible = True
        Call gobjControl.PicShowFlat(pic3, 1)
    End If
    If Index = 4 And mstrST = �޸� Then
        If InStr(1, LCase(TXT21(4).Text), "ml") > 0 Then TXT21(4).Text = Left(TXT21(4).Text, Len(TXT21(4).Text) - 2) '������ȥ�������ml��λ
    End If
End Sub

Private Sub TXT21_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 4:
            If Not (KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 46 Or (KeyAscii >= 48 And KeyAscii <= 57)) Then KeyAscii = 0: Exit Sub  '��Ѫ���ĳ��Ȳ�����6����ֻ�����롰.���س���backspace
            If InStr(TXT21(4).Text, ".") > 0 And KeyAscii = 46 Then KeyAscii = 0: Exit Sub '����С���������£������ظ�����С����
            If TXT21(4).Text = "" And KeyAscii = 46 Then KeyAscii = 0: Exit Sub '��û����������ǰ���������롰.��С����
        Case 3, 5:
            If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub 'Ѫ����ź���Ѫ��Ŀ����������س������������
        Case 0, 1, 2
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8) Then KeyAscii = 0: Exit Sub '������벻������Ҳ���ǻس�Ҳ����backspace���˳�
    End Select
    
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
    If Opt21(0).Value = True And Index = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TXT21_LostFocus(Index As Integer)
    pic3.Visible = False
End Sub

Private Sub TXT31_Change()
    If InStr(1, TXT31.Text, "/") > 0 Then
        If Val(Mid(TXT31.Text, 1, InStr(1, TXT31.Text, "/"))) > 24 Then
            TXT31.Text = "24" & Mid(TXT31.Text, InStr(1, TXT31.Text, "/"), Len(TXT31.Text) - InStr(1, TXT31.Text, "/") + 1)
        End If
    Else
        If Val(TXT31.Text) > 24 Then '��ֻ����Сʱʱ�����ֻ������24����������24�����ֶ�Ϊ��Ϊ24
            TXT31.Text = 24
        End If
    End If
End Sub

Private Sub TXT31_KeyPress(KeyAscii As Integer)
    'ֻ���������ֺ�/���лس���backspace
    If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8) Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
    
    If Opt31(0).Tag <> 1 Then '����ѡ������Ѫ�󣬷����޷���������
        KeyAscii = 0
    End If
    If InStr(1, TXT31.Text, "/") > 0 And KeyAscii = 47 Then KeyAscii = 0 'ֻ������һ��/
    
End Sub

Private Sub TXT32_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0 '�������ݲ����е�����
End Sub

Private Sub TXT33_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0 '�������ݲ����е�����
End Sub

Private Sub TXT41_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        If Index = 1 And (mstrST = ���� Or mstrST = �޸�) Then
            TXT41(1).Text = FindHS(TXT41(1).Text)
        End If
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub UserControl_Initialize()
    InitEdit
End Sub

Public Sub UserControl_Resize()
    On Error GoTo Errorhand
    Call ExecuteCommand("�ؼ�����")
Errorhand:
End Sub

Private Sub UserControl_Terminate()
    Set mRsBR = Nothing
    Set mRsFY = Nothing
    mDataChanged = False
    mblnAddPage = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '���ܣ���������û�������������
    Call PropBag.WriteProperty("TabsPosition", m_TabsPosition, m_def_TabsPosition)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_TabsPosition = PropBag.ReadProperty("TabsPosition", m_def_TabsPosition)
End Sub

Private Sub pic5_GotFocus(Index As Integer)
    '
    Dim lngIndex As Long
    Dim Arrtag
    On Error GoTo Errorhand
'    Call DestroyCaret'ȥ�������˸
    For lngIndex = 0 To lbl5.Count - 1
        DTP5(lngIndex).Visible = False
        Pic5(lngIndex).BorderStyle = 0
    Next
    Pic5(Index).BorderStyle = 1

    mlngSelNum = Index
    Call ExecuteCommand("ˢ������")
    If TXT21(5).Text = "" Then
        cbo2.Text = ""
    Else
        For lngIndex = 0 To cbo1.ListCount - 1 '��һ����Ϊ�����շ�id��Ѫ����ű���һ�£��Է������ݵ���ɾ��
            If cbo1.List(lngIndex) = TXT21(5).Text Then
                cbo2.Text = cbo2.List(lngIndex) & ""
            End If
        Next
    End If
    '��dtp����ʾ��һ���ĵ�����ֻ��δ��¼�����ݲŻ���ʾdtp�ؼ�
    Arrtag = Split(lbl5(mlngSelNum).Tag, ":")
    mstrST = Val(Arrtag(1))
    mlng״̬ = Val(Arrtag(0))
    mDataChanged = Val(Arrtag(2)) = 1
    If mstrST = ���� Then
        DTP5(Index).Visible = mDataChanged
        lbl5(mlngSelNum).Caption = Msk51(mlngSelNum).Text & " " & Msk52(mlngSelNum).Text
        Call Clear
    End If
    Call ExecuteCommand("�ؼ�״̬")
    Call ExecuteCommand("ˢ������")
    lbl��Ѫ������.Visible = IIf(Val(Msk51(mlngSelNum).Tag) = 1, True, False) '���˸�   ������Ѫ�������ֶ�����ʾ����������Ѫ��������־
    mbln��Ѫ������ = lbl��Ѫ������.Visible
Errorhand:
End Sub

Private Sub Pic6_GotFocus(Index As Integer)
    'lbl6��ѡ��ʱ
    Dim lngIndex As Long
    Dim Arrtag
    Dim lngVsfRight As Long
'    Call DestroyCaret
    For lngIndex = 0 To lbl6.Count - 1
        DTP6(lngIndex).Visible = False
        Pic6(lngIndex).BorderStyle = 0
    Next
    Pic6(Index).BorderStyle = 1
    mlngSelNum = Index
    'mlngSum = mlngSelNum
    lngVsfRight = vsbRight.Value
    vsbRight.Value = 0
    Call ExecuteCommand("ˢ������")
    If TXT21(5).Text = "" Then
        cbo2.Text = ""
    Else
        For lngIndex = 0 To cbo1.ListCount - 1 '��һ����Ϊ�����շ�id��Ѫ����ű���һ�£��Է������ݵ���ɾ��
            If cbo1.List(lngIndex) = TXT21(5).Text Then
                cbo2.Text = cbo2.List(lngIndex)
            End If
        Next
    End If

    '��dtp����ʾ��һ���ĵ�����ֻ��δ��¼�����ݲŻ���ʾdtp�ؼ�
    Arrtag = Split(lbl6(mlngSelNum).Tag, ":")
    mstrST = Val(Arrtag(1))
    mlng״̬ = Val(Arrtag(0))
    mDataChanged = Val(Arrtag(2)) = 1
    If mstrST = ���� Then
        DTP6(Index).Visible = mDataChanged
        lbl6(mlngSelNum).Caption = Msk61(mlngSelNum).Text & " " & Msk62(mlngSelNum).Text
        Call Clear
    End If
    Call ExecuteCommand("�ؼ�״̬")
    Call ExecuteCommand("ˢ������") '�����ٴ�ˢ��������Ϊ�˱���option�ؼ��޷���������⣬���ֻ������һ��ˢ�����ݻ���ֻ������һ��ˢ�����ݣ����޷�������option�ؼ���ˢ�£�����ԭ������ǿؼ���������⣬��Ϊˢ������ֻ�Ƕ�option�������ؼ���ֵ��һ���������δ������������
    If Pic6(Index).Top >= 18 * 450 + lngVsfRight Or Pic6(Index).Top < 0 Then lngVsfRight = (Index - 18) * 450
    If lngVsfRight < 0 Then lngVsfRight = 0
    vsbRight.Value = lngVsfRight
    lbl��Ѫ������.Visible = IIf(Val(Msk61(mlngSelNum).Tag) = 1, True, False) '���˸�   ������Ѫ�������ֶ�����ʾ����������Ѫ��������־
    mbln��Ѫ������ = lbl��Ѫ������.Visible
End Sub
Private Sub DTP5_CloseUp(Index As Integer)
    'dtp5ѡ��ʱ���
'    Pic5(Index).SetFocus
    Msk51(Index).Text = Format(DTP5(Index).Value, "YYYY-MM-DD")
    If IsDate(Msk52(Index).Text) = False Then
        Msk52(Index).Text = Format(Now, "HH:mm:ss")
    End If
    lbl5(Index).Caption = Msk51(Index) & " " & Msk52(Index)
    
End Sub
Private Sub DTP6_CloseUp(Index As Integer)
    'dtp6ѡ��ʱ���
'    pic6(Index).SetFocus
    Msk61(Index).Text = Format(DTP6(Index).Value, "YYYY-MM-DD")
    If IsDate(Msk62(Index).Text) = False Then
        Msk62(Index).Text = Format(Now, "HH:mm:ss")
    End If
    lbl6(Index).Caption = Msk61(Index) & " " & Msk62(Index)
End Sub

Private Function GetSecondUserName(ByVal objControl As TextBox, ByVal lngDeptID As Long, ByVal lng����ID As Long, lng��ҳid As Long, lng������Դ As Long, lngID As Long, str���� As String, strѪ����� As String, str������Ѫʷ As String, str�в���� As String) As Boolean
    '���ܣ�������ز�������ѯ���ݣ�����ʾ��һ�������ϣ��ҿ��Թ涨�������ʾλ�ú�ģʽ��ע����ѯ�����ݱ���Ҫ��id
    
    Dim rsUser As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim vPoint As POINTAPI, blnCancel As Boolean

    On Error GoTo ErrHand
    vPoint = GetCoordPos(UserControl.hWnd, TXT21(5).Left + pic4.Left + pic1.Left, TXT21(5).Top + pic4.Top + pic1.Top)
    '����
    strSQL = " Select id,ѪҺ����,Ѫ�����,ҽ����,decode(������Ѫʷ,0,'��','��') as ������Ѫʷ, �в����, ѪҺ״̬ " & _
            " From (SELECT e.Id AS Id, b.���� AS ѪҺ����, e.Ѫ����� AS Ѫ�����, a.Id AS ҽ����, c.������Ѫʷ AS ������Ѫʷ, c.�в����," & vbNewLine & _
            "       Decode(f.�շ�id, NULL, 1, f.����״̬) AS ����״̬," & vbNewLine & _
            "       Decode(Nvl(f.ִ��״̬, 0), 0, '�ѽ���', 1, '����ִ��', 2, '���ִ��', 3, 'ִֹͣ��') ѪҺ״̬" & vbNewLine & _
            "FROM ��Ѫ�����¼ c, ѪҺ���ͼ�¼ f, ������ĿĿ¼ b, ѪҺ��� g, ѪҺ�շ���¼ e, ѪҺ��Ѫ��¼ d, ����ҽ����¼ a" & vbNewLine & _
            "WHERE c.ҽ��id(+) = a.Id AND a.���id IS NULL AND NOT EXISTS" & vbNewLine & _
            " (SELECT 1 FROM ��Ѫ��Ӧ��¼ h WHERE e.Id = h.�շ�id" & IIf(mstrST = �޸�, " And h.�շ�ID<>[3]", " ") & ") AND e.Id = f.�շ�id(+) AND g.Ʒ��id = b.Id AND g.���id = e.ѪҺid AND" & vbNewLine & _
            "      e.����� IS NOT NULL AND MOD(e.��¼״̬, 3) = 1 AND e.�䷢id = d.Id AND d.��¼���� = 1 AND d.����id = a.Id AND a.������� = 'K' AND" & vbNewLine & _
            "      a.����id = [1] "
    If mlng�շ�ID = 0 Then
    If lng������Դ = 2 Then
        strSQL = strSQL & " And a.��ҳid = [2] ) Where ����״̬ in(1,3)"
        Set rsUser = gobjDatabase.ShowSQLSelect(mobjfrm, strSQL, 0, "��Ѫ��Ӧ", False, "", "", False, False, True, vPoint.X, vPoint.Y, TXT21(5).Height, blnCancel, False, True, lng����ID, lng��ҳid, Val(cbo2.Text))
    Else
        strSQL = strSQL & " And a.�Һŵ� = [2] ) Where ����״̬ in(1,3)"
        Set rsUser = gobjDatabase.ShowSQLSelect(mobjfrm, strSQL, 0, "��Ѫ��Ӧ", False, "", "", False, False, True, vPoint.X, vPoint.Y, TXT21(5).Height, blnCancel, False, True, lng����ID, mstr�Һŵ�, Val(cbo2.Text))
        End If
    Else
        strSQL = strSQL & " and e.id = [4] "
        If lng������Դ = 2 Then
            strSQL = strSQL & " And a.��ҳid = [2] )" & IIf(gbln���պ����ִ�� = True, " Where ����״̬ in(1,3)", "")
            Set rsUser = gobjDatabase.OpenSQLRecord(strSQL, "��Ѫ��Ӧ", lng����ID, lng��ҳid, Val(cbo2.Text), mlng�շ�ID)
        Else
            strSQL = strSQL & " And a.�Һŵ� = [2] ) " & IIf(gbln���պ����ִ�� = True, " Where ����״̬ in(1,3)", "")
            Set rsUser = gobjDatabase.OpenSQLRecord(strSQL, "��Ѫ��Ӧ", lng����ID, mstr�Һŵ�, Val(cbo2.Text), mlng�շ�ID)
        End If
        mlng�շ�ID = 0
    End If
    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then Exit Function
            lngID = Val(rsUser!id)
            str���� = Nvl(rsUser!ѪҺ����) & ""
            str������Ѫʷ = rsUser!������Ѫʷ & ""
            str�в���� = rsUser!�в���� & ""
            strѪ����� = Nvl(rsUser!Ѫ�����)
            GetSecondUserName = True
            
            strSQL = "select c.���� from ѪҺ��Ѫ��¼ a,ѪҺ�շ���¼ b,����ҽ������ c where c.ҽ��id=a.����id and c.��Ŀ='���뵥���' and a.id=b.�䷢id and b.id=[1]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ѯ��ϼ�¼", lngID)
            If rsTmp.EOF = False Then
                TXT11(6).Text = rsTmp.Fields("����").Value & ""
            Else
                TXT11(6).Text = ""
            End If
        End If
    ElseIf blnCancel = False Then
        MsgBox "����Ѫ��¼��", vbInformation, gstrSysName
    End If

    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function
Public Property Get hWnd() As Long
    hWnd = Me.hWnd
End Property

Private Sub VS1_Change()
    VS1_Scroll
End Sub

Private Sub VS1_Scroll()
    pic1.Top = -VS1.Value
End Sub
Private Sub HS1_Change()
    HS1_Scroll
End Sub

Private Sub HS1_Scroll()
    pic1.Left = -HS1.Value
End Sub

Private Function GetZXR(lng�շ�ID As Long) As String
    '���ܣ�ͨ���շ�id��ȡִ����
    '������lng�շ�id-�շ�id
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    strSQL = " Select ִ���� ��ʼִ���� From ѪҺִ�м�¼ where �շ�id=[1] and ��¼����=1 and ���=0"
    Set rsUser = gobjDatabase.OpenSQLRecord(strSQL, "��ѯִ����", lng�շ�ID)
    If rsUser.EOF = False Then GetZXR = rsUser.Fields("��ʼִ����").Value & "": Exit Function
    GetZXR = ""
End Function

Private Function FindHS(strName As String) As String
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    
    vPoint = GetCoordPos(UserControl.hWnd, TXT41(1).Left + pic4.Left + pic1.Left, TXT41(1).Top + pic1.Top + pic4.Top)
    strSQL = " Select distinct a.Id ,a.���, a.����,a.���� " & vbNewLine & _
             " From ��Ա�� a, ��Ա����˵�� b " & vbNewLine & _
             " Where a.Id = b.��Աid And b.��Ա���� = 'ҽ��' And (a.���� Like [1] or a.��� like [1] or a.���� like [1])"
    
    Set rsUser = gobjDatabase.ShowSQLSelect(mobjfrm, strSQL, 0, "��Ѫ��Ӧ", False, "", "", False, False, True, vPoint.X, vPoint.Y, TXT41(1).Height, blnCancel, False, True, strName & "%")

    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then FindHS = "": Exit Function
            FindHS = rsUser!����
            Exit Function
        End If
    End If
    FindHS = ""
    
End Function

