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
               Name            =   "宋体"
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
            Caption         =   "治愈"
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
            Caption         =   "死亡"
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
            Caption         =   "发热"
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
            Caption         =   "发绀"
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
            Caption         =   "呼吸困难"
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
            Caption         =   "两肺布满湿性罗音"
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
            Caption         =   "休克"
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
            Caption         =   "皮肤充血"
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
            Caption         =   "伤口渗血不止"
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
            Caption         =   "咳大量血性泡沫样痰"
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
            Caption         =   "寒颤"
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
            Caption         =   "荨麻诊"
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
            Caption         =   "颈静脉怒张"
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
            Caption         =   "酱油色尿"
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
            Caption         =   "黄疽"
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
            Caption         =   "腰背痛"
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
            Caption         =   "其他"
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
            Caption         =   "发热反应"
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
            Caption         =   "过敏反应"
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
            Caption         =   "急性溶血反应"
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
            Caption         =   "其他"
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
            Caption         =   "1.立即停止输血，保持静脉通路，同时观察剩余血外观"
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
            Caption         =   "2.采患者血及袋中剩余血(最好和血袋一起)送输血科检查"
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
            Caption         =   "3.留取反应后第一次尿送检"
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
            Caption         =   "4.对症处理"
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
               Caption         =   "输血期间"
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
               Caption         =   "输血后"
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
               Caption         =   "一级亲属"
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
               Caption         =   "二级亲属"
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
               Caption         =   "无关系"
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
               Caption         =   "无"
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
               Caption         =   "有"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   15
               Top             =   75
               Width           =   615
            End
         End
         Begin VB.Label lbl输血科新增 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "输"
            BeginProperty Font 
               Name            =   "宋体"
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
            Caption         =   "有无输血反应"
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
            Caption         =   "患者信息"
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
            Caption         =   "输血情况"
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
            Caption         =   "输血不良反应相关情况"
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
            Caption         =   "签名"
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
            Caption         =   "姓    名"
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
            Caption         =   "性    别"
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
            Caption         =   "年    龄"
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
            Caption         =   "住 院 号"
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
            Caption         =   "血    型"
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
            Caption         =   "临床诊断"
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
            Caption         =   "既往输血史"
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
            Caption         =   "妊娠史"
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
            Caption         =   "输血项目"
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
            Caption         =   "输入量"
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
            Caption         =   "献血者与受血者的关系"
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
            Caption         =   "血袋编号"
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
            Caption         =   "次"
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
            Caption         =   "孕:"
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
            Caption         =   "产:"
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
            Caption         =   "发生时间"
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
            Caption         =   "转归"
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
            Caption         =   "病状与体征"
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
            Caption         =   "诊断"
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
            Caption         =   "临床处理措施"
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
            Caption         =   "输血科处理措施"
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
            Caption         =   "处理描述"
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
            Caption         =   "护士"
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
            Caption         =   "经治医师"
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
            Caption         =   "输血科"
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
         Caption         =   "标签状态"
         ForeColor       =   &H00000000&
         Height          =   1815
         Left            =   120
         TabIndex        =   54
         Top             =   7320
         Width           =   1455
         Begin VB.Label lbl标签状态 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "已完成"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   760
            Width           =   1215
         End
         Begin VB.Label lbl标签状态 
            Appearance      =   0  'Flat
            BackColor       =   &H80000002&
            Caption         =   "医生待提交"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lbl标签状态 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "输血待提交"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   1400
            Width           =   1215
         End
         Begin VB.Label lbl标签状态 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "标签状态"
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

'新增数据后，不管这次数据而去修改或删除其他数据，就会造成错误的问题，可以试着把删除数据后面几个页面的数据全部往前移动
Option Explicit
 
'Implements clsBloodEdit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function HideCaret Lib "User32.dll" (ByVal hWnd As Long) As Boolean
Private Declare Function DestroyCaret Lib "user32" () As Long
'状态数据
Public Enum Status
    缺省 = -1 '默认状态
    新增 = 0
    保存 = 1
    修改 = 2
    删除 = 3
End Enum

Public Enum Position '标签显示在什么位置
    PosiTop = 0
    Posiright = 1
End Enum
Public Enum 病状与体征
    发热 = 0
    发绀 = 1
    呼吸困难 = 2
    两肺布满湿性罗音 = 3
    休克 = 4
    皮肤充血 = 5
    伤口渗血不止 = 6
    咳大量血性泡沫样痰 = 7
    寒颤 = 8
    荨麻诊 = 9
    颈静脉怒张 = 10
    酱油色尿 = 11
    黄疽 = 12
    腰背痛 = 13
    其他病状 = 14
End Enum

Public Enum 诊断 '与 病状与体征 同理
    发热反应 = 0
    过敏反应 = 1
    急性溶血反应 = 2
    其他诊断 = 3
End Enum

Public Enum 临床处理措施
    立即停止输血 = 0
    血液送输血科 = 1
    采尿送检 = 2
    对症处理 = 3
End Enum
'缺省属性值:
Const m_def_TabsPosition = PosiTop

'属性变量:
Private m_TabsPosition As Position

'控件引用
Private mPicTabs As PictureBox
Private mTXTbox As TextBox
Private mButtion As CommandButton
Private mDataChanged As Boolean
'全局变量
Private mlngSum As Long                    '存储tab的个数
Private mstrST As Status                   '存储当前操作名：保存、删除等
Private mlngSelNum As Long                 '选中tab的index
Private mlng病人ID As Long
Private mlng主页id As Long
Private mlng病人来源 As Long
Private mlng收发ID As Long
Private mgcnCpOracle As ADODB.Connection
Private mRsBR As ADODB.Recordset           '存放病人的基础信息
Private mRsFY As ADODB.Recordset           '存放输血反应记录信息
Private mrsXD As ADODB.Recordset           '存放血袋编号，血液名称，孕产情况等信息
Private mlng状态 As Long                   '顾名思义存放状态信息
Private mlng阶段 As Long                   '表示用户的阶段如医生操作阶段和输血科操作阶段
Private mblnStart As Boolean               '代表程序开始的全局变量
Private mobjfrm As Object
Private mblnAddPage As Boolean             '新增页面标志
Private mblnCancel As Boolean              '取消标志
Private mlng模块号 As Long
Private mblnAddNew As Boolean
Private marrFilter                         '过滤数组
Private mstrFilter As String               '过滤数据串
Private mblnHaveData As Boolean            '用户有数据传入即有人员数据时，mblnHaveData=true,没有人员数据则一直为false
Private mstr挂号单 As String
Private mbln输血科新增 As Boolean           '是否是输血科新增的数据
Private mbln输血科新增权限 As Boolean       '输血科是否有新增权限
'Event Clear() '清除界面上除用户数据外的所有数据，并非删除。
'这是用于截图的api

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''标签定位'''''''''''''''''''''''''''''''''
Public Function BloodLocation(ByVal lng收发ID As Long) As Boolean
    Dim lngi As Long
    If mRsFY Is Nothing Then Exit Function
    If mRsFY.State = adStateClosed Then BloodLocation = False: Exit Function
    'mRsFY.Filter = "收发id = '" & mRsFY!收发ID & "'"
    If mRsFY.RecordCount = 0 Then BloodLocation = False: Exit Function
    mRsFY.MoveFirst
    '定位到查找到的数据
    For lngi = 0 To mRsFY.RecordCount - 1 '当能够查询到相关数据时
        If lng收发ID = mRsFY!收发ID & "" Then
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
'''''''''''''公共方法，增删改，保存，打印等操作'''''''''''''''''''''''''''''''''
Public Function ShowSave() As Boolean
    '功能：保存
    '参数：
    '返回：
    If mlng阶段 = 2 And mbln输血科新增权限 = False Then Exit Function '输血科没有新增权限时，不能保存   余浪改   mlng状态 = 0 And
    If cbo2.Text = "" Then MsgBox "血袋编号为空，不能保存！", vbInformation, gstrSysName: Exit Function
    Call ExecuteCommand("保存数据") '保存数据要分修改保存和新增保存
    Call ExecuteCommand("控件状态") '根据不同的操作如增删改，来改变控件的状态
    ShowSave = True
End Function
Public Function ShowDelete() As Boolean
    '功能：删除
    '参数：
    '返回：
    If (mlng阶段 = 1 And mlng状态 = 1) Or (mlng阶段 = 2 And mlng状态 = 2) Then Exit Function '已审核数据不能删除
    If MsgBox("是否删除该记录？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Function
    Call ExecuteCommand("删除数据") '如果状态为删除数据则删除相关数据
    ShowDelete = True
End Function

Public Sub showPrintSet() '打印设置
    Call ReportPrintSet(gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm)
End Sub

Public Sub ShowPrint(id As Long)  '打印
    '功能：打印和预览，
    '参数：id-1:打印 2:预览
    '返回：
    If id = 1 Then '打印
        ReportOpen gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm, "收发id=" & Val(cbo2.Text), 2
    ElseIf id = 2 Then '打印预览
        ReportOpen gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm, "收发id=" & Val(cbo2.Text), 1
    End If
End Sub
Public Sub ShowPrintList(id As Long)
    '显示打印列表，可以按照用户的需求打印数据
    Dim strPrint As String
    Dim ArrPrint
    Dim lngFilter As Long
    Dim lngi As Long
    Dim strSelBloodid As String
    If mRsFY Is Nothing Then MsgBox "未选中病人或无病人信息！", vbInformation, gstrSysName: Exit Sub
    If mRsFY.RecordCount = 0 Then MsgBox "无该病人的反应记录！", vbInformation, gstrSysName: Exit Sub
    
    strSelBloodid = cbo2.Text & ""
    
    strPrint = frmbloodReactionPrint.BloodPrintList(mlng病人ID, mlng病人来源, mlng主页id, mstrFilter, mlng阶段, strSelBloodid)
    
    If strPrint = "" Then Exit Sub '没有选中要打印的数据则退出
    ArrPrint = Split(strPrint, ";")
    For lngi = 0 To UBound(ArrPrint)
        If id = 1 Then '打印
            ReportOpen gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm, "收发id=" & Val(ArrPrint(lngi)), 2
        ElseIf id = 2 Then '打印预览
            ReportOpen gcnOracle, 2200, "ZL22_BILL_1938", mobjfrm, "收发id=" & Val(ArrPrint(lngi)), 1
        End If
    Next
End Sub

Public Sub ShowCancel()
    '功能：取消当前操作，会还原为以前的状态
    '参数：
    '返回：
'    Call ExecuteCommand("刷新全部数据")
    Dim Arrtag
    
    mblnCancel = True
    Call ExecuteCommand("控件状态")
    If mstrST = 新增 And mlngSelNum <> 0 Then
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
        If mlngSelNum > mlngSum Then '选中的选项卡在新增的页面上，如果删除该新增页面，那么也要更改选中状态，不然会出错
            mlngSelNum = mlngSum
        End If
        mstrST = 缺省
        mblnAddPage = False
    Else
        If m_TabsPosition = PosiTop Then
            Msk51(mlngSelNum).Visible = False
            Msk52(mlngSelNum).Visible = False
        Else
            Msk61(mlngSelNum).Visible = False
            Msk62(mlngSelNum).Visible = False
        End If
        mstrST = 缺省
        mblnAddPage = False
    End If
    Call ExecuteCommand("刷新数据")
    If m_TabsPosition = PosiTop Then
        Arrtag = Split(lbl5(mlngSelNum).Tag, ":")
        lbl5(mlngSelNum).Tag = Arrtag(0) & ":" & mstrST & ":0"
        If mstrST <> 新增 Then
            mDataChanged = False
        End If
        mblnCancel = False
        Pic5(mlngSelNum).SetFocus
    ElseIf m_TabsPosition = Posiright Then
        Arrtag = Split(lbl6(mlngSelNum).Tag, ":")
        lbl6(mlngSelNum).Tag = Arrtag(0) & ":" & mstrST & ":0"
        If mstrST <> 新增 Then
            mDataChanged = False
        End If
        mblnCancel = False
        Pic6(mlngSelNum).SetFocus
    End If
    
End Sub

Public Function SubmitData() As Boolean
    '功能：提交数据
    '参数：
    '返回：
    Dim SelNum As Long '保留选中页面，防止提交页面后页面不能选中。
    If mstrST = 新增 Then Exit Function '新增页面不能提交，只有新增页面保存后才允许提交
    SelNum = mlngSelNum
    Call ExecuteCommand("提交数据")
    mlngSelNum = SelNum
    If m_TabsPosition = PosiTop Then '提交数据后mblngoptchange=true
        pic5_GotFocus (mlngSelNum)
    ElseIf m_TabsPosition = Posiright Then
        Pic6_GotFocus (mlngSelNum)
    End If
    SubmitData = True
End Function

Public Sub ShowModify()
    '功能：修改
    '参数：
    '返回：
    Dim Arrtag
    Dim blnDtpVisible As Boolean
    
    If mlng阶段 = 2 Then
        blnDtpVisible = lbl输血科新增.Visible
    Else
        blnDtpVisible = True
    End If
    
    If m_TabsPosition = PosiTop Then '
        Arrtag = Split(lbl5(mlngSelNum).Tag, ":")
        lbl5(mlngSelNum).Tag = Arrtag(0) & ":2:1"
        mstrST = 修改
        DTP5(mlngSelNum).Visible = blnDtpVisible
    Else
        Arrtag = Split(lbl6(mlngSelNum).Tag, ":")
        lbl6(mlngSelNum).Tag = Arrtag(0) & ":2:1"
        mstrST = 修改
        DTP6(mlngSelNum).Visible = blnDtpVisible
    End If
    
    Call ExecuteCommand("控件状态") '根据不同的操作如增删改，来改变控件的状态
    mDataChanged = True
    If Not UserControl.ActiveControl Is Nothing Then
        If UserControl.ActiveControl.name = "TXT21" Then
            If UserControl.ActiveControl.Index = 5 Then Call TXT21_GotFocus(5)
        End If
    End If
    If mbln输血科新增 = True Then
        TXT41(1).SelStart = 0
        TXT41(1).SelLength = Len(TXT41(1).Text)
        TXT41(1).SetFocus
    End If
End Sub

Public Sub AddPage()
    '功能：新增页面
    '参数：
    '返回：
    If Not mRsFY Is Nothing Then
        If mRsFY.RecordCount = 0 Then
            mDataChanged = True
            mblnAddPage = True
            Call ExecuteCommand("控件状态") '
            If Not UserControl.ActiveControl Is Nothing Then
                If UserControl.ActiveControl.name = "TXT21" Then
                    If UserControl.ActiveControl.Index = 5 Then Call TXT21_GotFocus(5)
                End If
            End If
            Exit Sub
        End If
    End If
    If mblnAddPage = True Then Exit Sub
    Call ExecuteCommand("新增页面") '
    mDataChanged = True
    mblnAddPage = True
End Sub

Public Sub ShowClear()
    '清空页面数据，同时清空除初始控件外的所有控件
    Dim lngi As Long
    
    Clear
    For lngi = 0 To TXT11.Count - 1
        TXT11(lngi).Text = ""
    Next
    lbl输血科新增.Visible = False
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
    mstrST = 缺省
    Call ExecuteCommand("控件状态")
End Sub

Public Sub ShowUntread()
    '功能：回退
    '参数：无
    '返回：无
    If mstrST = 新增 Then Exit Sub '新增页面不能回退，只有新增页面保存后才允许回退
    Dim SelNum As Long
    SelNum = mlngSelNum
    Call ExecuteCommand("回退数据")
    mlngSelNum = SelNum
'    Call ExecuteCommand("刷新数据")
    If m_TabsPosition = PosiTop Then '刷新全部数据后mblngoptchange=true
        pic5_GotFocus (mlngSelNum)
    ElseIf m_TabsPosition = Posiright Then
        Pic6_GotFocus (mlngSelNum)
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitEdit()
    '功能：初始化页面，将指定病人的信息显示在界面上，默认不可编辑
    '参数：病人ID病人的相关id用于查找用户的相关信息。
    '      医嘱ID查询医嘱相关信息
    '返回
    mstrST = 缺省 '初始化状态信息
    mlngSum = 0
    mlng状态 = 0
    mlngSelNum = 0
    mDataChanged = False
    pic3.Visible = False
    CboDW.Clear
    CboDW.AddItem "ml"
    CboDW.AddItem "U"
    CboDW.AddItem "治疗量"
    CboDW.ListIndex = 0
    cboHave.Clear
    cboHave.AddItem "有"
    cboHave.AddItem "无"
    Call ExecuteCommand("控件布局")
    Call ExecuteCommand("初始控件")
    
End Sub

Public Sub showInfor(lng病人ID As Long, lng病人来源 As Long, lng主页id As Long, lng阶段 As Long, cnMain As ADODB.Connection, objfrmMain As Object, lng模块号 As Long, _
                Optional strFilter As String = "", Optional bln输血科新增权限 As Boolean = False, Optional ByVal lng收发ID As Long)
    '功能：根据病人id和主页id显示病人的数据
    '参数：lng病人id-病人的id号，lng病人来源-住院2、门诊1，lng主页id-病人的主页id或就诊id，lng阶段-医生阶段1、输血科阶段2，
    '    : cnMain数据库连接，objfrm-主窗体，lng系统号-主窗体的系统号
    '    : Filter是过滤条件数组，里面包含过滤科室，过滤时间，填写人，提交状态。
    '    ：bln输血科新增权限-true输血科有新增权限 false-输血科无新增权限
    '返回：
    Dim lngi As Long
    Dim strSQL As String
    Dim rsSQL As ADODB.Recordset
    Dim blnLocated As Boolean
'    ReDim mArrFilter(0 To 3)
    
    mlng病人ID = lng病人ID
    mlng主页id = lng主页id '当病人是住院病人时，主页id是住院病人的主页id，病人是门诊病人时，主页id是病人的就诊id，这里为了方便共用了mlng主页id这一变量
    mlng阶段 = lng阶段
    mlng模块号 = lng模块号
    mlng病人来源 = lng病人来源
    mbln输血科新增权限 = bln输血科新增权限
    mlng收发ID = lng收发ID
    
    mstrFilter = strFilter
    If strFilter <> "" Then '如果有过滤条件则将过滤条件放入数组
        marrFilter = Split(strFilter, "|")
    Else '如果没有过滤条件则重定义数组，但数组中的所有数据为空
        ReDim marrFilter(0 To 4)
    End If
    
    Set mobjfrm = objfrmMain
    Set mgcnCpOracle = cnMain
    
    
    If zlGetComLib = False Then MsgBox "获取对象失败！", vbInformation, gstrSysName: Exit Sub
    mblnAddPage = False '初始化添加页面信息
    mblnHaveData = True
    HS1.Visible = False
    VS1.Visible = False
    
    '获取挂号单号
    mstr挂号单 = ""
    If mlng病人来源 <> 2 Then
        strSQL = " select no from 病人挂号记录 where id=[1] "
        Set rsSQL = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", mlng主页id)
        If rsSQL.RecordCount > 0 Then
            mstr挂号单 = rsSQL.Fields("no")
        End If
    End If
    
    Call ExecuteCommand("清空控件")
    Call gobjControl.PicShowFlat(pic3, 1)
'    Call ExecuteCommand("控件布局")
    Call ExecuteCommand("初始病人信息") '1、查询到相关数据放入数据集中，这就是整个控件的数据源
    Call ExecuteCommand("获取反应记录")
    If mRsFY.BOF = False Then
        For lngi = 0 To mRsFY.RecordCount - 1 '当能够查询到相关数据时
            mlngSum = lngi
            mlng状态 = Val(mRsFY.Fields(16).Value & "")
            Call ExecuteCommand("初始控件")
            cbo2.Text = mRsFY.Fields(0).Value & ""
            
            mRsFY.MoveNext
        Next
        If mlng收发ID = 0 Then
        mlngSelNum = mlngSum
        Else
            If Not BloodLocation(mlng收发ID) Then
                AddPage
                blnLocated = False
            Else
                blnLocated = True
            End If
        End If
        mRsFY.MoveFirst
    Else
        mlngSum = 0
        mlng状态 = 0
        Call ExecuteCommand("初始控件")
        Call ExecuteCommand("病人基础数据")
        mstrST = 新增
    End If
    
'    Call ExecuteCommand("刷新数据")
    Call ExecuteCommand("控件状态") '4、默认所有控件不可编辑
    mblnStart = True
    If mlng收发ID <> 0 Then
        If Not blnLocated Then
            mstrST = 新增
            Call ExecuteCommand("控件状态")
            Call ExecuteCommand("病人基础数据")
            '填写ABO和RH
            If mRsBR.Fields("ABO").Value & "" = "" Then
                Set rsSQL = GetPatientOtherInfo(mlng病人ID, "ABO")
                If rsSQL.BOF = False Then TXT11(4).Text = rsSQL("信息值").Value
            Else
            TXT11(4).Text = mRsBR.Fields("ABO").Value
            End If
            If mRsBR.Fields("RH").Value & "" = "" Then
                Set rsSQL = GetPatientOtherInfo(mlng病人ID, "RH")
                If rsSQL.BOF = False Then TXT11(5).Text = rsSQL("信息值").Value
            Else
                TXT11(5).Text = mRsBR.Fields("RH").Value
            End If
            Call pic3_MouseDown(1, 0, 0, 0)
            mlng收发ID = 0
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
    If TXT11(1).Text = "男" Then
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


'属性获取赋值
Public Property Get BloodID() As Long
    BloodID = Val(cbo2.Text)
End Property

Public Property Get TabsPosition() As Position
    TabsPosition = m_TabsPosition
    Call ExecuteCommand("控件布局")
    Call ExecuteCommand("初始控件")
End Property

Public Property Let TabsPosition(ByVal NewTabsPosition As Position)
    m_TabsPosition = NewTabsPosition
    PropertyChanged "TabsPosition"
    Call ExecuteCommand("控件布局")
    Call ExecuteCommand("初始控件")
End Property
'获取当前病人的输血反应的条数
Public Property Get lngFYCount() As Long
    If Not mRsFY Is Nothing Then
        lngFYCount = mRsFY.RecordCount
    Else
        lngFYCount = 0
    End If
End Property

Public Property Get 输血科新增() As Boolean
    输血科新增 = mbln输血科新增
End Property

Public Property Get 有无输血反应() As Boolean
    有无输血反应 = IIf(cboHave.Text = "有", True, False)
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

Public Property Get lng状态() As Long
    lng状态 = mlng状态
End Property

Public Property Get strST() As Status
    strST = mstrST
End Property

Public Property Get blnAddPage() As Boolean
    blnAddPage = mblnAddPage
End Property

Private Function CanbeChange(lng阶段 As Long, lng状态 As Long, st操作 As Status) As Boolean
    '功能：可以根据阶段和状态来更改控件的状态还可以通过这个函数来判断一些操作是否可行
    '参数：st操作即现阶段是在做什么：新增，修改还是删除
    '返回：
    Dim lngi As Long
    Dim str阶段性判断 As String
    Dim bln有无反应 As Boolean
    str阶段性判断 = lng阶段 & ":" & lng状态
    Select Case st操作
        Case 修改:
                Select Case str阶段性判断
                    Case "1:0":
                        For lngi = 0 To TXT21.Count - 1
                            TXT21(lngi).locked = False
                        Next
                        TXT31.locked = False
                        TXT32.locked = False
                        TXT21(5).locked = False
                    Case "2:1":
                        TXT33.locked = False
                    Case "2:0" '余浪改   输血科新增页面的修改
                        If mbln输血科新增权限 = True Then
                            For lngi = 0 To TXT21.Count - 1
                                TXT21(lngi).locked = False
                            Next
                            TXT31.locked = False
                            TXT32.locked = False
                            TXT33.locked = False
    '                        TXT21(5).Locked = False
                        End If
                End Select
                '男性不能填写妊娠史
                If TXT11(1).Text = "男" Then
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
        Case 新增:
            '新增必然是医生的新增操作，所以只让医生填写部分使能
            For lngi = 0 To TXT21.Count - 1
                TXT21(lngi).locked = Not mDataChanged
            Next
            '男性不能填写妊娠史
            If TXT11(1).Text = "男" Then
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
            If mbln输血科新增权限 = True Then '余浪改，如果有输血科新增权限则可以直接在新增页面添加输血科处理措施。
                TXT33.locked = Not mDataChanged
            End If
    End Select
    
    bln有无反应 = IIf(cboHave.Text = "有", True, False)
    '根据是否有输血反应的选项，控制输血反应部分控件
    If cboHave.Text <> "有" Then
        Opt31(0).Value = False
        Opt31(1).Value = False
        Opt31(0).Tag = 2
        Opt31(1).Tag = 2
        Opt32(0).Value = False
        Opt32(1).Value = False
        Opt32(0).Tag = 2
        Opt32(1).Tag = 2
    End If
    Opt31(0).Enabled = bln有无反应
    Opt31(1).Enabled = bln有无反应
    TXT31.Enabled = bln有无反应
    lbl31(6).ForeColor = IIf(bln有无反应 = False, &H80000011, &H80000008)
    Opt32(0).Enabled = bln有无反应
    Opt32(1).Enabled = bln有无反应
    For lngi = 0 To Chk31.Count - 1
        Chk31(lngi).Enabled = bln有无反应
    Next
    For lngi = 0 To Chk32.Count - 1
        Chk32(lngi).Enabled = bln有无反应
    Next
    For lngi = 0 To Chk33.Count - 1
        Chk33(lngi).Enabled = bln有无反应
    Next
    lbl31(7).ForeColor = IIf(bln有无反应 = False, &H80000011, &H80000008)
    TXT32.Enabled = bln有无反应
    TXT33.Enabled = bln有无反应
End Function
Private Sub Clear()
    '功能：清空页面上的的所有数据
    '参数
    '返回
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
    Dim strXD As String '存放血袋编号，孕产记录，输血记录等的sql语句
    Dim blnLoad As Boolean '判断是否加载控件
    Dim rsTmp As New Recordset

    On Error GoTo Error

    Call SQLRecord(rsSAD)
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "控件布局"
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
                '此段代码会刷新右侧标签位置，以前位置无法变动，而后新增了滚动条，刷新会导致定位丢失，并且此刷新意义不明，故暂时屏蔽右侧标签刷新
'                For lngi = 0 To Pic6.Count - 1
'                    Pic6(lngi).Move Picright.ScaleLeft, Picright.ScaleTop + lngi * 450, 1215, 400
'                    Pic6(lngi).Visible = True
'                Next
                
                '显示标签状态
                Fra1.Visible = True
                Fra1.Move Picright.ScaleLeft, Picright.ScaleTop + Picright.ScaleHeight - 1500, 1220, 1455
                
                lbl标签状态(0).Move 30, 540, Fra1.Width - 60, 255
                lbl标签状态(1).Move 30, 810, Fra1.Width - 60, 255
                lbl标签状态(2).Move 30, 1080, Fra1.Width - 60, 255
                lbl标签状态(3).Move 30, 270, Fra1.Width - 60, 255
            
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
        Case "初始控件"

            If m_TabsPosition = PosiTop Then
                For lngi = 0 To Pic5.Count - 1
                    If lngi = mlngSum Then
                        blnLoad = True
                    End If
                Next
                If Not blnLoad And mlngSum > 0 Then '当mlngsum>0且当前控件不是选中控件时才加载相关控件，这里是避免已加载控件重复加载
                    Load Pic5(mlngSum)
                    Load DTP5(mlngSum)
                    Load Msk51(mlngSum)
                    Load Msk52(mlngSum)
                    Load pic51(mlngSum)
                    Load lbl5(mlngSum)
                End If
                SetParent pic51(mlngSum).hWnd, Pic5(mlngSum).hWnd
                Set lbl5(mlngSum).Container = pic51(mlngSum) '将标签放在容器里
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
                lbl5(mlngSum).Tag = mlng状态 & ":-1:0"   '标识各个页面的状态信息，方便调整控件状态,格式：（状态:操作:mdatachanged）
                mstrST = 缺省
            Else
                For lngi = 0 To Pic6.Count - 1
                    If lngi = mlngSum Then
                        blnLoad = True
                    End If
                Next
                If Not blnLoad And mlngSum > 0 Then '当mlngsum>0且当前控件不是选中控件时才加载相关控件，这里是避免已加载控件重复加载
                    Load Pic6(mlngSum)
                    Load DTP6(mlngSum)
                    Load Msk61(mlngSum) '''''''
                    Load Msk62(mlngSum)
                    Load pic61(mlngSum)
                    Load lbl6(mlngSum)
                End If
                SetParent pic61(mlngSum).hWnd, Pic6(mlngSum).hWnd
                Set lbl6(mlngSum).Container = pic61(mlngSum) '将标签放在容器里
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
                lbl6(mlngSum).Tag = mlng状态 & ":-1:0"
                mstrST = 缺省
            End If
            '根据状态显示不同的tab颜色
            If m_TabsPosition = PosiTop Then
                If mlng状态 = 0 Then '根据状态，来改变相应的控件的颜色
                    pic51(mlngSum).BackColor = &H80000002
                    Pic5(mlngSum).BackColor = &H80000002
                ElseIf mlng状态 = 1 Then
                    pic51(mlngSum).BackColor = &HC0FFC0
                    Pic5(mlngSum).BackColor = &HC0FFC0
                ElseIf mlng状态 = 2 Then
                    pic51(mlngSum).BackColor = &H80000000
                    Pic5(mlngSum).BackColor = &H80000000
                End If
            ElseIf m_TabsPosition = Posiright Then

                If mlng状态 = 0 Then '根据状态，来改变相应的控件的颜色
                    pic61(mlngSum).BackColor = &H80000002
                    Pic6(mlngSum).BackColor = &H80000002
                ElseIf mlng状态 = 1 Then
                    pic61(mlngSum).BackColor = &HC0FFC0
                    Pic6(mlngSum).BackColor = &HC0FFC0
                ElseIf mlng状态 = 2 Then
                    pic61(mlngSum).BackColor = &H80000000
                    Pic6(mlngSum).BackColor = &H80000000
                End If
            End If
            
            mlngSelNum = mlngSum
            If m_TabsPosition = PosiTop Then
                If mlngSum > 0 Then
'                   pic5(mlngSum).SetFocus'无法聚焦，会异常跳出，所以这儿去掉
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
        Case "清空控件"
            '卸载除第一个选项卡外的所有选项卡，这一步是为了在转换病人时使用
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
            Call ExecuteCommand("控件布局")
        Case "初始病人信息"
            Dim sqlFilter As String
            '余浪   2017年2月8日
            strSqlBR = " Select b.病人id, a.住院号 As 住院号, b.姓名, b.性别, b.年龄, b.Id As 医嘱id, b.医嘱内容, b.医生嘱托, d.Abo, d.Rh, d.血袋编号, d.Id As 收发id, h.诊断描述,c.执行部门id " & _
                       " From 病案主页 a, 病人医嘱记录 b, 血液配血记录 c, 血液收发记录 d, " & _
                       "      (Select g.医嘱id, f.诊断描述 " & _
                       "        From 病人诊断记录 f, 病人诊断医嘱 g " & _
                       "        Where f.Id = g.诊断id And f.病人id = [1] And f.主页id = [2]) h " & _
                       " Where d.单据 = 6 And d.配发id = c.Id And c.申请id = b.Id and c.记录性质=1 And b.诊疗类别 = 'K' And h.医嘱id(+) = b.Id And Mod(d.记录状态, 3) = 1 And " & _
                       "       d.审核人 Is not Null And b.诊疗类别 = 'K' And a.病人id(+) = b.病人id And a.主页id(+) = b.主页id And b.病人id = [1] "
            If mlng病人来源 = 2 Then '住院病人
                strSqlBR = strSqlBR & " And b.主页id =[2]"
                Set mRsBR = gobjDatabase.OpenSQLRecord(strSqlBR, "病人信息", mlng病人ID, mlng主页id)
            Else '门诊病人,
                strSqlBR = strSqlBR & " and b.挂号单 =[3]"
                Set mRsBR = gobjDatabase.OpenSQLRecord(strSqlBR, "病人信息", mlng病人ID, mlng主页id, mstr挂号单)
            End If
            cbo1.Clear
            cbo2.Clear
            If mRsBR.RecordCount > 0 Then
                For lngj = 0 To mRsBR.RecordCount - 1
                    cbo1.AddItem mRsBR.Fields("血袋编号").Value   '血袋编号
                    cbo2.AddItem mRsBR.Fields("收发id").Value '收发id
                    mRsBR.MoveNext
                Next
                mRsBR.MoveFirst
            End If
            Call ExecuteCommand("控件布局")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "获取反应记录"
            Dim lngFilter
            Dim ArrTime
            Dim str是否本人 As String
            Dim lng部门id As Long
            Dim str开始时间 As String
            Dim str结束时间 As String
            Dim int输血反应 As Integer  '无为0，有为1，所有为2
            
            If mstrFilter <> "" Then
                int输血反应 = Int(marrFilter(4))
                lngFilter = marrFilter(3) '提交状态
                str是否本人 = marrFilter(2) '记录人
                lng部门id = marrFilter(0) '部门id
                str开始时间 = Split(marrFilter(1), "'")(0)
                str结束时间 = Split(marrFilter(1), "'")(1)
            Else
                lngFilter = 0
                str是否本人 = ""
                lng部门id = 0
                str开始时间 = Now
                str结束时间 = Now
            End If
            '获取病人的输血反应数据，主要是从输血反应记录中获取
            '去掉了以前的部门过滤条件，没必要在反应里面加
            strSqlFY = " Select distinct d.收发id, Nvl(d.反应时间,d.记录时间) 反应时间, d.输血史, d.输血次数, d.妊娠史, d.输血项目, d.输入量, d.献受者关系, d.发生时机, d.转归, d.不良反应, d.反应诊断, d.科室处理标识, d.科室处理措施," & _
                       " d.记录人 , d.记录时间, d.状态, d.血库处理措施, d.确认人, d.确认时间,decode(d.有无输血反应,0,'',1,'有',2,'无') as 有无输血反应,decode(d.是否输血科新增,1,1,0) as 是否输血科新增 " & _
                       " From 病人医嘱记录 a, 血液配血记录 b, 血液收发记录 c, 输血反应记录 d" & _
                       " Where d.收发id = c.id  And c.配发id = b.ID and mod(c.记录状态,3)=1 and c.审核人 is not null And b.申请id = a.ID and b.记录性质=1 and a.诊疗类别='K' And a.病人ID = [1] "
            If mlng病人来源 = 2 Then '住院病人
                If lngFilter = 0 And mlng阶段 = 1 Then '全部数据,医生阶段
                    strSqlFY = strSqlFY & "and a.主页id=[2] "
                ElseIf lngFilter = 0 And mlng阶段 = 2 Then '全部数据,输血科阶段 会根据输血科有无新增权限调整查询条件。
                    strSqlFY = strSqlFY & "and a.主页id=[2]  and (d.状态<>0 or d.是否输血科新增=1 )"
                ElseIf lngFilter = 1 And mlng阶段 = 1 Then '未提交数据,医生
                    strSqlFY = strSqlFY & "and a.主页id=[2] and d.状态=0 "
                ElseIf lngFilter = 1 And mlng阶段 = 2 Then '未提交数据,输血科
                    strSqlFY = strSqlFY & "and a.主页id=[2] and (d.状态<>2 and d.是否输血科新增=1 Or d.状态=1)"
                ElseIf lngFilter = 2 And mlng阶段 = 1 Then '已提交数据，医生
                    strSqlFY = strSqlFY & "and a.主页id=[2] and d.状态 <>0 "
                ElseIf lngFilter = 2 And mlng阶段 = 2 Then '已提交数据，输血科
                    strSqlFY = strSqlFY & "and a.主页id=[2] and d.状态=2 "
                End If
            Else
                If lngFilter = 0 And mlng阶段 = 1 Then '全部数据,医生阶段
                    strSqlFY = strSqlFY & "and a.挂号单=[7] "
                ElseIf lngFilter = 0 And mlng阶段 = 2 Then '全部数据,输血科阶段
                    strSqlFY = strSqlFY & "and a.挂号单=[7]  and (d.状态<>0 or d.是否输血科新增=1 )"
                ElseIf lngFilter = 1 And mlng阶段 = 1 Then '未提交数据,医生
                    strSqlFY = strSqlFY & "and a.挂号单=[7] and d.状态=0 "
                ElseIf lngFilter = 1 And mlng阶段 = 2 Then '未提交数据,输血科
                    strSqlFY = strSqlFY & "and a.挂号单=[7] and (d.状态<>2 and d.是否输血科新增=1 Or d.状态=1)"
                ElseIf lngFilter = 2 And mlng阶段 = 1 Then '已提交数据，医生
                    strSqlFY = strSqlFY & "and a.挂号单=[7] and d.状态 <>0 "
                ElseIf lngFilter = 2 And mlng阶段 = 2 Then '已提交数据，输血科
                    strSqlFY = strSqlFY & "and a.挂号单=[7] and d.状态=2 "
                End If
            End If
            
            If marrFilter(2) <> "" Then
                strSqlFY = strSqlFY & IIf(mlng阶段 = 2, " And (d.确认人=[3] or d.是否输血科新增 =1 And d.记录人=[3]) ", " and d.记录人=[3] ")
            End If
            
            '输血科要按时间和有无反应过滤反应记录
            If mlng阶段 = 2 Then
                If int输血反应 = 0 Then
                    strSqlFY = strSqlFY & " and d.有无输血反应 = 2 "
                ElseIf int输血反应 = 1 Then
                    strSqlFY = strSqlFY & " and d.有无输血反应 = 1 "
                ElseIf int输血反应 = 3 Then
                    strSqlFY = strSqlFY & " and d.有无输血反应 = 0 "
                End If
                
                If str开始时间 <> "" And str结束时间 <> "" Then
                    strSqlFY = strSqlFY & " and d.反应时间 Between [5] and [6] "
                End If
            End If
                strSqlFY = strSqlFY & " order by Nvl(d.反应时间,d.记录时间) "
            Set mRsFY = gobjDatabase.OpenSQLRecord(strSqlFY, "病人输血反应记录", mlng病人ID, mlng主页id, str是否本人, lng部门id, CDate(str开始时间), CDate(str结束时间), mstr挂号单)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "提交数据"
            Dim lng收发ID As Long, lng库房id As Long
            Dim lng病区ID As Long, lng科室ID As Long
            lng收发ID = Val(cbo2.Text)
            If cbo2.Text = "" Then MsgBox "未选中要提交或回退的输血反应记录!", vbInformation, gstrSysName: ExecuteCommand = False: Exit Function
            
            StrSqlSAD = "Zl_输血反应记录_Submit(" & lng收发ID & "," & mlng阶段 & "," & mlng状态 & ",'" & UserInfo.姓名 & "'," & IIf(mbln输血科新增 = True, 1, 0) & ")" '余浪改
            Call SQLRecordAdd(rsSAD, StrSqlSAD)
            If mlng阶段 = 1 And cboHave.Text = "有" Then
                lng库房id = Val(mRsBR!执行部门ID)
                If mlng病人来源 = 2 Then
                    StrSqlSAD = "select 出院科室ID,当前病区ID from 病案主页 where 病人id = [1] and 主页id = [2] "
                    Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "获取库房id", mlng病人ID, mlng主页id)
                    If Not rsTmp.EOF Then lng病区ID = Val(rsTmp!当前病区id): lng科室ID = Val(rsTmp!出院科室ID)
                Else
                    StrSqlSAD = "select 执行部门id from 病人挂号记录 where 病人id = [1] and id = [2] "
                    Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "获取库房id", mlng病人ID, mlng主页id)
                    If Not rsTmp.EOF Then lng病区ID = Val(rsTmp!执行部门ID): lng科室ID = Val(rsTmp!执行部门ID)
                End If
                StrSqlSAD = "Zl_业务消息清单_Insert(" & mlng病人ID & "," & mlng主页id & ","  '病人id 就诊id
                StrSqlSAD = StrSqlSAD & Val(lng科室ID) & ","     '就诊科室id
                StrSqlSAD = StrSqlSAD & Val(lng病区ID) & ","      '就诊病区id
                StrSqlSAD = StrSqlSAD & mlng病人来源 & ","                                      '病人来源
                StrSqlSAD = StrSqlSAD & "'有新的输血反应需要处理。','"             '消息内容
                StrSqlSAD = StrSqlSAD & IIf(Val(mlng病人来源) = 1, "0000", "0000") & "','ZLHIS_BLOOD_008',"     ' 提醒场合, 类型编码
                StrSqlSAD = StrSqlSAD & "'" & lng收发ID & "',"                      '业务标识（收发id）
                StrSqlSAD = StrSqlSAD & "1,0,NULL,'" & Val(lng库房id) & "',NULL)"
                Call SQLRecordAdd(rsSAD, StrSqlSAD)
            End If
            Call SQLRecordExecute(rsSAD)
            Call ExecuteCommand("刷新全部数据")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "回退数据"
            '回退指定id的数据，delete过程包括了回退的操作，这样不用多谢存储过程
            StrSqlSAD = "Zl_输血反应记录_Submit(" & Val(cbo2.Text) & "," & mlng阶段 & "," & mlng状态 & ",'" & UserInfo.姓名 & "'," & IIf(mbln输血科新增 = True, 1, 0) & ")" '余浪改
            Call SQLRecordAdd(rsSAD, StrSqlSAD)
            '设置该收发id的消息为已读
            If mlng阶段 = 1 Then
                If mlng病人来源 = 2 Then
                    StrSqlSAD = "select 出院科室ID 部门id from 病案主页 where 病人id = [1] and 主页id = [2] "
                Else
                    StrSqlSAD = "select 执行部门id 部门id from 病人挂号记录 where 病人id = [1] and id = [2] "
                End If
                Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "获取库房id", mlng病人ID, mlng主页id)
                StrSqlSAD = "Zl_业务消息清单_Read(" & mlng病人ID & "," & mlng主页id & ",'ZLHIS_BLOOD_008',"
                StrSqlSAD = StrSqlSAD & IIf(mlng病人来源 = 2, 2, 1) & ",'" & UserInfo.姓名 & "'," & Val(Nvl(rsTmp!部门ID & "")) & ",NULL,"
                StrSqlSAD = StrSqlSAD & "NULL," & Val(cbo2.Text) & ")"
                Call SQLRecordAdd(rsSAD, StrSqlSAD)
            End If
            Call SQLRecordExecute(rsSAD)
            Call ExecuteCommand("刷新全部数据")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "新增页面"
             
            Call Clear
            mlngSum = mlngSum + 1
            mlng状态 = 0 '新增的数据默认为待提交
            Call ExecuteCommand("初始控件")
            mDataChanged = True
            If m_TabsPosition = PosiTop Then '新增后将状态改为新增
                lbl5(mlngSelNum).Tag = "0:0:1" 'tag里面存放的格式为"状态：操作"，新增页面默认是状态为0，操作是新增,datachange是true
                mstrST = 新增
            Else
                lbl6(mlngSelNum).Tag = "0:0:1"
                mstrST = 新增
            End If
            Call ExecuteCommand("控件状态")
            If m_TabsPosition = PosiTop Then '新增后将状态改为新增
                pic5_GotFocus (mlngSelNum) '.SetFocus
            Else
                Pic6_GotFocus (mlngSelNum) '(mlngSelNum).SetFocus
            End If
            Call ExecuteCommand("控件布局")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "控件状态"
            '根据操作修改各个控件的状态
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
            
            If (mstrST = 新增 Or mstrST = 修改) And mbln输血科新增 = True Then
                TXT41(1).locked = False
            End If
            Call CanbeChange(mlng阶段, mlng状态, mstrST)
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "保存数据"
            '保存数据分为修改保存和新增保存，可以在一个sql函数中包含这两种
            Dim int输血史 As Integer
            Dim str妊娠史 As String
            Dim int献受者关系 As Integer
            Dim str发生时机 As String
            Dim int转归 As Integer
            Dim str不良反应  As String
            Dim str反应诊断 As String
            Dim str科室处理标识 As String
            Dim lngselnum As Long
            Dim str反应时间 As String
            Dim str记录人 As String
            Dim str确认人 As String
            Dim str输血量 As String
            Dim lng有无输血反应 As Long
            Dim lng输血科新增 As Long '余浪改
            '对数据进行处理，方便数据录入数据库
            
            If mlng阶段 = 2 Then
                StrSqlSAD = "select id from 业务消息清单 where 病人id = [1] and 就诊id = [2] and 业务标识 = [3] and 是否已阅 = 0 "
                Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "查询消息记录", mlng病人ID, mlng主页id, Val(cbo2.Text))
                If rsTmp.RecordCount <> 0 Then
                    StrSqlSAD = "select 库房id 部门id from 血液收发记录 where id = [1] "
                    Set rsTmp = gobjDatabase.OpenSQLRecord(StrSqlSAD, "获取库房id", Val(cbo2.Text))
                    StrSqlSAD = "Zl_业务消息清单_Read(" & mlng病人ID & "," & mlng主页id & ",'ZLHIS_BLOOD_008',"
                    StrSqlSAD = StrSqlSAD & "5,'" & UserInfo.姓名 & "'," & Val(Nvl(rsTmp!部门ID & "")) & ",NULL,"
                    StrSqlSAD = StrSqlSAD & "NULL,NULL)"
                    Call SQLRecordAdd(rsSAD, StrSqlSAD)
                End If
            End If
            If TXT11(1).Text = "男" Then
                str妊娠史 = ":"
            Else
                str妊娠史 = TXT21(1).Text & ":" & TXT21(2).Text
            End If
            
            int输血史 = Val(Opt21(0).Tag)
            int献受者关系 = Val(Opt22(0).Tag) + 1 '献受关系
            
            int转归 = Val(Opt32(0).Tag) + 1 '转归：1-治愈 2-死亡 3-无
            '对输血量进行检查
            If Right(TXT21(4).Text, 1) = "." Then
                If Len(TXT21(4).Text) > 5 Then
                    MsgBox "输入量不能大于" & IIf(CboDW.Text = "U", "5U!", "1000ml!"), vbInformation, gstrSysName
                    TXT21(4).SetFocus
                    ExecuteCommand = False
                    Exit Function
                Else
                    str输血量 = TXT21(4).Text & "0" & CboDW.Text
                End If
            Else
                str输血量 = TXT21(4).Text & CboDW.Text
            End If
            If (CboDW.Text = "U" And Val(TXT21(4).Text) > 5) Then
                MsgBox "输入量不能大于5U!", vbInformation, gstrSysName
                TXT21(4).SetFocus
                ExecuteCommand = False
                Exit Function
            End If
            If (CboDW.Text = "ml" And Val(TXT21(4).Text) > 1000) Then
                MsgBox "输入量不能大于1000ml!", vbInformation, gstrSysName
                TXT21(4).SetFocus
                ExecuteCommand = False
                Exit Function
            End If
            If (CboDW.Text = "治疗量" And Val(TXT21(4).Text) > 5) Then
                MsgBox "输入量不能大于5个治疗量!", vbInformation, gstrSysName
                TXT21(4).SetFocus
                ExecuteCommand = False
                Exit Function
            End If

            If Opt31(0).Value = True Then
                str发生时机 = "输血期间"
            ElseIf Opt31(1).Value = True Then
                If TXT31.Text <> "" Then
                    str发生时机 = IIf(Right(TXT31.Text, 1) = "/", Left(TXT31.Text, Len(TXT31.Text) - 1), TXT31.Text) '去掉最右侧的/
                    str发生时机 = IIf(Left(TXT31.Text, 1) = "/", Right(TXT31.Text, Len(TXT31.Text) - 1), TXT31.Text) '去掉最左侧的/
                Else
                    str发生时机 = ""
                End If
            Else
                str发生时机 = "无"
            End If
            
            str不良反应 = ""
            For lngi = 0 To Chk31.Count - 1 '除"其他"以外被选择的情况
                If Chk31(lngi).Value = Checked Then
                    str不良反应 = str不良反应 & Chk31(lngi).Caption & "," '余浪改 2017年6月8日 和老数据不兼容
                End If
            Next
            If str不良反应 <> "" Then
                str不良反应 = Left(str不良反应, Len(str不良反应) - 1)
            End If
            
'            If Chk31(其他病状).Value = Checked Then
'                str不良反应 = str不良反应 & "99:"
'            End If

            str反应诊断 = ""
            For lngi = 0 To Chk32.Count - 1 '除"其他"以外被选择的情况
                If Chk32(lngi).Value = Checked Then
                    str反应诊断 = str反应诊断 & Chk32(lngi).Caption & ","
                End If
            Next
            If str反应诊断 <> "" Then
                str反应诊断 = Left(str反应诊断, Len(str反应诊断) - 1)
            End If
            
'            If Chk32(其他诊断).Value = Checked Then
'                str反应诊断 = str反应诊断 & "99:"
'            End If
            
            str科室处理标识 = ""
            For lngi = 0 To Chk33.Count - 1
                If Chk33(lngi).Value = Checked Then
                    str科室处理标识 = str科室处理标识 & Split(Chk33(lngi).Caption, ".")(1) & ","
                End If
            Next
            If str科室处理标识 <> "" Then
                str科室处理标识 = Left(str科室处理标识, Len(str科室处理标识) - 1)
            End If
            
            
            If m_TabsPosition = PosiTop Then
                lbl5(mlngSelNum).Caption = Msk51(mlngSelNum) & " " & Msk52(mlngSelNum)
                If IsDate(lbl5(mlngSelNum).Caption) = False Then MsgBox "时间格式错误!请核对", vbInformation, gstrSysName: Exit Function '对时间格式进行判断
                str反应时间 = lbl5(mlngSelNum).Caption
            ElseIf m_TabsPosition = Posiright Then
                lbl6(mlngSelNum).Caption = Msk61(mlngSelNum) & " " & Msk62(mlngSelNum)
                If IsDate(lbl6(mlngSelNum).Caption) = False Then MsgBox "时间格式错误!请核对", vbInformation, gstrSysName: Exit Function
                str反应时间 = lbl6(mlngSelNum).Caption
            End If
            
            If mlng阶段 = 1 Then
                str记录人 = IIf(TXT41(1).Text = "", UserInfo.姓名, TXT41(1).Text) '如果保存时用户没有录入医生姓名，则默认当前用户姓名
                str确认人 = TXT41(2).Text
            ElseIf mlng阶段 = 2 Then
                If mbln输血科新增权限 = True And mstrST = 新增 Then
                    str记录人 = IIf(TXT41(1).Text = "", UserInfo.姓名, TXT41(1).Text)
                    str确认人 = TXT41(2).Text
                Else
                    str记录人 = TXT41(1).Text
                    str确认人 = UserInfo.姓名
                End If
            End If

            lng有无输血反应 = IIf(cboHave.Text = "", 0, IIf(cboHave.Text = "有", 1, 2))
            lng输血科新增 = IIf((mlng阶段 = 2 And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = True And mlng阶段 = 2), 1, 0) '   余浪改
            '检查输血反应措施和临床处理措施，看是否符合字数要求
            If gobjCommFun.StrIsValid(TXT32.Text, 500) = False Then ExecuteCommand = False: Exit Function
            If gobjCommFun.StrIsValid(TXT33.Text, 500) = False Then ExecuteCommand = False: Exit Function
            '余浪改
            If cbo2.Text <> "" And TXT21(5).Text <> "" Then
                StrSqlSAD = "Zl_输血反应记录_Insert(" & Val(cbo2.Text) & ",to_date('" & str反应时间 & "','yyyy-mm-dd hh24:mi:ss')," & int输血史 & "," & Val(TXT21(0).Text) & ",'" & str妊娠史 & "','" & TXT21(3).Text & "','" & str输血量 & "'," & int献受者关系 & _
                             ",'" & str发生时机 & "'," & int转归 & ",'" & str不良反应 & "','" & str反应诊断 & "','" & str科室处理标识 & "','" & TXT32.Text & "','" & str记录人 & _
                             "'," & mlng状态 & ",'" & TXT33.Text & "','" & str确认人 & "'," & lng有无输血反应 & "," & lng输血科新增 & ")"
                Call SQLRecordAdd(rsSAD, StrSqlSAD)
                ExecuteCommand = SQLRecordExecute(rsSAD)
            Else
                ShowCancel '保存数据时会遇到空保存的情况，即保存的数据没有关键数据血袋编号，收发id等，这种保存是无效的，这时便会自动删除页面，一面影响其他操作
                ExecuteCommand = True
                Exit Function
            End If
            lngselnum = mlngSelNum
            
            If mstrST = 新增 Then
                mblnAddPage = False
            End If
            Call ExecuteCommand("刷新全部数据")

            If m_TabsPosition = PosiTop Then '保存数据后要聚焦到保存数据的选项卡上，使用pic5.setfocus没有效果，所以这里直接调用pic5_gotFocus
                pic5_GotFocus (lngselnum)
            ElseIf m_TabsPosition = Posiright Then
                Pic6_GotFocus (lngselnum)
            End If
            mDataChanged = False
            
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "刷新全部数据"
            '刷新整个页面
            Call ExecuteCommand("初始病人信息") '1、查询到相关数据放入数据集中，这就是整个控件的数据源
            Call ExecuteCommand("获取反应记录")
            If mRsFY.BOF = False Then
                For lngi = 0 To mRsFY.RecordCount - 1 '当能够查询到相关数据时
                    mlngSum = lngi
                    mlng状态 = mRsFY.Fields(16).Value
                    Call ExecuteCommand("初始控件") '2、根据数据源中的数据初始化各个控件,即看要添加多少个选项卡，选项卡中的数据有什么，选项卡是根据时间来的这个要注意，如果没有数据那么至少还是要有一个选项卡，内容为空
                    cbo2.Text = mRsFY.Fields(0).Value
                    mRsFY.MoveNext
                Next
                
                mlngSelNum = mlngSum
                mRsFY.MoveFirst
            Else
                mlngSum = 0
                mlng状态 = 0
                Call ExecuteCommand("初始控件")
            End If
            Call Clear '删除相关页面后清除相关页面的信息
            Call ExecuteCommand("刷新数据") '3、将对应时间的数据投入到选项卡中，选项卡有选中选项，默认为选项卡1中显示数据
            Call ExecuteCommand("控件状态") '4、默认所有控件不可编辑
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Case "病人基础数据"
            '当通过"初始病人信息"不能查询到数据时，说明该病人没有输血记录，这时候只需要显示病人的基础信息。
            Dim strSqlBasic As String
            Dim rsBasic As ADODB.Recordset
            For lngi = 0 To TXT11.Count - 1
                TXT11(lngi).Text = ""
            Next
            If mlng病人来源 = 2 Then
            strSqlBasic = " select a.姓名,a.性别,a.年龄,a.住院号 || '' as 住院号 from 病案主页 a where a.病人id=[1] and a.主页id=[2] "
            Else
                strSqlBasic = "select b.姓名,b.性别,b.年龄,'' as 住院号 from 病人挂号记录 b where b.病人id=[1] and b.id=[2]"
            End If
            Set rsBasic = gobjDatabase.OpenSQLRecord(strSqlBasic, "病人基础信息", mlng病人ID, mlng主页id)
            
            If rsBasic.RecordCount > 0 Then
                TXT11(0).Text = rsBasic.Fields("姓名").Value & ""
                TXT11(1).Text = rsBasic.Fields("性别").Value & ""
                TXT11(2).Text = rsBasic.Fields("年龄").Value & ""
                TXT11(3).Text = rsBasic.Fields("住院号").Value & ""
            End If
        Case "删除数据"
            If cbo2.Text = "" Then MsgBox "未选择反应记录!", vbInformation, gstrSysName: Exit Function
'            If Not mRsFY Is Nothing Then
'                If mRsFY.RecordCount = 0 Then Exit Function
'            End If
            '删除指定id的数据
            StrSqlSAD = "Zl_输血反应记录_delete(" & Val(cbo2.Text) & "," & mlng阶段 & "," & mlng状态 & ")"
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
                Call ExecuteCommand("刷新全部数据")
            Else
                Call ExecuteCommand("刷新全部数据")
            End If
            
            If m_TabsPosition = PosiTop Then
                pic5_GotFocus (mlngSum) 'pic5(mlngSum).SetFocus
            ElseIf m_TabsPosition = Posiright Then
                Pic6_GotFocus (mlngSum) 'Pic6(mlngSum).SetFocus
            End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case "刷新数据"
            Dim strABO As String
            Dim strRH As String
            Dim strZD As String
            '点击不同tab时，数据刷新并显示在页面上,其实就是将数据集中对应的数据投放到页面上
            If mblnHaveData = False Then ExecuteCommand = False: Exit Function '输血反应记录不存在时
            If mRsFY Is Nothing Then ExecuteCommand = False: Exit Function
            ClearTag
            '读取病人的abo和rh
            Set rsTmp = GetPatientOtherInfo(mlng病人ID, "ABO")
            If rsTmp.BOF = False Then strABO = rsTmp("信息值").Value
            Set rsTmp = GetPatientOtherInfo(mlng病人ID, "RH")
            If rsTmp.BOF = False Then strRH = rsTmp("信息值").Value
            cboHave.Text = "" '清空cbohave中的数据
            If mRsFY.RecordCount = 0 Then '当不存在输血反应记录时，显示病人的基础信息，并将状态改为新增，以方便用户添加
                mlng状态 = 0 '如果没有查询到反应记录的数据则默认是待提交的数据
                If mRsBR.EOF = False Then '2，3，4，1，8，9
                     TXT11(0).Text = mRsBR.Fields("姓名").Value & ""
                     TXT11(1).Text = mRsBR.Fields("性别").Value & ""
                     TXT11(2).Text = mRsBR.Fields("年龄").Value & ""
                     TXT11(3).Text = mRsBR.Fields("住院号").Value & ""
                     TXT11(4).Text = IIf(mRsBR.Fields("ABO").Value & "" = "", strABO, mRsBR.Fields("ABO").Value & "")
                     TXT11(5).Text = IIf(mRsBR.Fields("RH").Value & "" = "", strRH, mRsBR.Fields("RH").Value & "")
                End If
'                mDataChanged = True '余浪 2017年6月22日16:17:41  医生工作站的输血反应单，如果没有输血反应记录时，无法添加输血反应。
                If m_TabsPosition = PosiTop Then '若没有输血反应记录则将初始页面的状态改为新增，适用于所有数据都删除掉，或者无反应记录的情况
                    mstrST = 新增
                    lbl5(mlngSelNum).Tag = "0:" & mstrST & ":0"
                Else
                    mstrST = 新增
                    lbl6(mlngSelNum).Tag = "0:" & mstrST & ":0"
                End If
                
            Else '有反应记录时
                With mRsFY
                    For lngi = 0 To .RecordCount - 1 '数据条数
                        If .EOF Then Exit For
                        If m_TabsPosition = PosiTop Then
                            lbl5(lngi).Caption = Format(.Fields("反应时间").Value & "", "YYYY-MM-DD HH:mm:ss")
                            Msk51(lngi).Text = Format(Nvl(.Fields("反应时间").Value, "____-__-__"), "YYYY-MM-DD")
                            Msk52(lngi).Text = Format(Nvl(.Fields("反应时间").Value, "__:__:__"), "HH:mm:ss")
                            Msk51(lngi).Tag = Val(.Fields("是否输血科新增").Value & "") '余浪改
                        Else
                            lbl6(lngi).Caption = Format(.Fields("反应时间").Value & "", "YYYY-MM-DD HH:mm:ss")
                            Msk61(lngi).Text = Format(Nvl(.Fields("反应时间").Value, "____-__-__"), "YYYY-MM-DD")
                            Msk62(lngi).Text = Format(Nvl(.Fields("反应时间").Value, "__:__:__"), "HH:mm:ss")
                            Msk61(lngi).Tag = Val(.Fields("是否输血科新增").Value & "")
                        End If
                        If mlngSelNum > 0 Then
                            If lngi = mlngSelNum - 1 Then '为了在新增页面后能使页面正常显示
                                Call Clear
                            End If
                        End If
                        If lngi = mlngSelNum Then '选中选单数据
                            If mRsBR.BOF = False Then
                                 TXT11(0).Text = mRsBR.Fields("姓名").Value & ""  '姓名
                                 TXT11(1).Text = mRsBR.Fields("性别").Value & "" '性别
                                 TXT11(2).Text = mRsBR.Fields("年龄").Value & "" '年龄
                                 TXT11(3).Text = mRsBR.Fields("住院号").Value & "" '住院号
                                 TXT11(4).Text = IIf(mRsBR.Fields("ABO").Value & "" = "", strABO, mRsBR.Fields("ABO").Value & "")
                                 TXT11(5).Text = IIf(mRsBR.Fields("RH").Value & "" = "", strRH, mRsBR.Fields("RH").Value & "")
                            End If
                            '余浪  2017年2月8日
                            strXD = " Select e.Id, b.名称, e.血袋编号, c.即往输血史, c.孕产情况 " & _
                                    " From 病人医嘱记录 a, 诊疗项目目录 b, 血液规格 g,输血申请记录 c, 血液配血记录 d, 血液收发记录 e " & _
                                    " Where e.配发id = d.Id And d.申请id = a.Id and g.品种id = b.Id AND g.规格id = e.血液id AND " & _
                                    " d.记录性质=1 And c.医嘱id(+) =  a.Id " & _
                                    " And a.相关id Is Null and Mod(e.记录状态, 3) = 1 And e.审核人 Is not Null and e.id=[1] "
                            Set mrsXD = gobjDatabase.OpenSQLRecord(strXD, "查询血袋编号等", Val(.Fields("收发id").Value))

                            If mrsXD.BOF = False Then
                                TXT21(5).Text = mrsXD.Fields("血袋编号").Value
                                cbo2.Text = mrsXD.Fields("id").Value
                                TXT21(3).Text = mrsXD.Fields("名称").Value '输血项目
                            End If
                            
                            TXT21(0).Text = IIf(.Fields("输血次数").Value = 0, "", .Fields("输血次数").Value) '输血次数
                            '孕产情况
                            If .Fields("妊娠史").Value & "" <> "" Then
                                TXT21(1).Text = Split(.Fields("妊娠史").Value, ":")(0) ' 孕
                                TXT21(2).Text = Split(.Fields("妊娠史").Value, ":")(1)  '产
                            End If
                            
                            If InStr(.Fields("输入量").Value & "", "治疗量") > 0 Then
                                TXT21(4).Text = Split(.Fields("输入量").Value & "", "治疗量")(0)
                                CboDW.ListIndex = 2
                            ElseIf InStr(.Fields("输入量").Value & "", "U") > 0 Then
                                TXT21(4).Text = Split(.Fields("输入量").Value & "", "U")(0)
                                CboDW.ListIndex = 1
                            ElseIf InStr(UCase(.Fields("输入量").Value & ""), "ML") > 0 Then
                                TXT21(4).Text = Split(UCase(.Fields("输入量").Value & ""), "ML")(0)
                                CboDW.ListIndex = 0
                            End If
                            Opt21(0).Tag = Val(.Fields("输血史").Value & "")
                            Opt21(Val(Opt21(0).Tag)).Value = True
                            Opt22(0).Tag = Val(Nvl(.Fields("献受者关系").Value, 1)) - 1
                            Opt22(Val(Opt22(0).Tag)).Value = True
                            If Val(.Fields("转归").Value & "") < 3 Then
                                Opt32(0).Tag = Val(Nvl(.Fields("转归").Value, 1)) - 1
                                Opt32(Val(Opt32(0).Tag)).Value = True
                            Else
                                Opt32(0).Tag = Val(.Fields("转归").Value & "") - 1
                            End If
                            
                            cboHave.Text = .Fields("有无输血反应").Value & ""
                            
                            If .Fields("发生时机").Value & "" = "输血期间" Then '发生时机
                                Opt31(0).Tag = 0
                                Opt31(0).Value = True
                            ElseIf .Fields("发生时机").Value & "" = "无" Then
                                Opt31(0).Tag = 2
                            Else
                                Opt31(0).Tag = 1
                                Opt31(1).Value = True
                                TXT31.Text = .Fields(8).Value & ""  '发生时机
                            End If
                            
                            For lngk = 0 To Chk31.Count - 1 '刷新数据时首先将控件的选中状态还原
                                Chk31(lngk).Value = Unchecked
                            Next
                            If IsNull(.Fields("不良反应").Value) = False Then
                                StrSplit = Split(.Fields("不良反应").Value, ",")
                                For lngk = 0 To Chk31.Count - 1 '不良反应
                                    For lngj = 0 To UBound(StrSplit)
                                        If StrSplit(lngj) = Chk31(lngk).Caption Then '余浪 2017年6月8日 反应记录不兼容老数据的问题
                                            Chk31(lngk).Value = Checked
                                            Chk31(lngk).Tag = 1
                                        End If
                                    Next
                                Next
                            End If
                            
                            For lngk = 0 To Chk32.Count - 1
                                Chk32(lngk).Value = Unchecked
                            Next
                            If IsNull(.Fields("反应诊断").Value) = False Then
                                StrSplit = Split(.Fields("反应诊断").Value, ",")
                                For lngk = 0 To Chk32.Count - 1 '反应诊断
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
                            If IsNull(.Fields("科室处理标识").Value) = False Then
                                StrSplit = Split(.Fields("科室处理标识").Value, ",")
                                For lngk = 0 To Chk33.Count - 1 ''科室处理标识
                                    For lngj = 0 To UBound(StrSplit)
                                        If StrSplit(lngj) = Split(Chk33(lngk).Caption, ".")(1) Then
                                            Chk33(lngk).Value = Checked
                                            Chk33(lngk).Tag = 1
                                        End If
                                    Next
                                Next
                            End If

                            mlng状态 = Val(.Fields("状态").Value & "") '提取反应记录的状态
                            TXT32.Text = .Fields("科室处理措施").Value & "" '科室处理措施
                            TXT33.Text = .Fields("血库处理措施").Value & "" '血库处理措施
                            TXT41(0).Text = GetZXR(.Fields("收发id")) & "" '护士(执行人)
                            TXT41(1).Text = .Fields("记录人").Value & "" '记录人
                            TXT41(2).Text = .Fields("确认人").Value & ""  '确认人
                            '填写该血袋对应的医嘱的诊断信息
                            strZD = "select c.内容 from 血液配血记录 a,血液收发记录 b,病人医嘱附件 c where c.医嘱id=a.申请id and c.项目='申请单诊断' and a.id=b.配发id and b.id=[1]"
                            Set rsTmp = gobjDatabase.OpenSQLRecord(strZD, "查询诊断记录", Val(.Fields("收发id").Value))
                            If rsTmp.EOF = False Then
                                TXT11(6).Text = rsTmp.Fields("内容").Value & ""
                            Else
                                TXT11(6).Text = ""
                            End If
                            mRsBR.MoveFirst
                        End If
                        .MoveNext
                    Next
                    .MoveFirst
                End With
                Call ExecuteCommand("控件布局")
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
    '限制输血量
    If CboDW.Text = "U" And Val(TXT21(4).Text) > 5 Then
        MsgBox "输入量不能大于5U!", vbInformation, gstrSysName
    End If
    If CboDW.Text = "ml" And Val(TXT21(4).Text) > 1000 Then
        MsgBox "输入量不能大于1000ml!", vbInformation, gstrSysName
    End If
End Sub

Private Sub CboDW_DropDown()
    '新增和修改功能中可以修改单位，其他情况不行
    If mstrST = 新增 Or mstrST = 修改 Then
        CboDW.locked = False
    Else
        CboDW.locked = True
    End If
    If mlng阶段 = 2 Then CboDW.locked = True '输血科阶段不允许修改单位
End Sub

Private Sub cboHave_Click()
    '清空反应部分内容
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
    '改变控件状态
    Call ExecuteCommand("控件状态")
End Sub

Private Sub cboHave_DropDown()
    '新增和修改功能中可以修改单位，其他情况不行
    If (mstrST = 新增 And mDataChanged = True) Or mstrST = 修改 Then
        cboHave.locked = False
    Else
        cboHave.locked = True
    End If
    '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许修改有无输血反应   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then cboHave.locked = True
End Sub

Private Sub cboHave_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
    gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub Chk31_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And mstrST <> 新增 And mstrST <> 修改 Then KeyCode = 0: Exit Sub '排除新增和修改其外的操作
    '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许变更chk31控件的内容   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then KeyCode = 0: Exit Sub
'    If mstrST = 修改 And mlng阶段 = 2 And mbln输血科新增权限 = False Then KeyCode = 0: Exit Sub '排除输血科阶段没有新增权限时的修改操作 余浪改
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
    '在修改状态，只有mlng阶段=1和mlng状态=1时才可以修改,新增状态也可以修改，其他状态都不允许修改
    If Button = 2 Then Exit Sub '点击右键，还是可以变更checkbox的状态，所以这里点击右键就跳出事件
    '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许变更chk31控件的内容   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then
'    If mstrST = 修改 And mlng阶段 = 2 And mbln输血科新增权限 = False Then '输血科点击chk31时,如果没有输血科新增权限则，不允许其修改   余浪改
        Chk31(Index).Value = Val(Chk31(Index).Tag)
        Exit Sub
    End If
    If mstrST = 修改 Or mstrST = 新增 Then
        Chk31(Index).Tag = Chk31(Index).Value
    Else
        Chk31(Index).Value = Val(Chk31(Index).Tag)
    End If
End Sub

Private Sub Chk32_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And mstrST <> 新增 And mstrST <> 修改 Then KeyCode = 0: Exit Sub '排除新增和修改其外的操作
    If mstrST = 修改 And mlng阶段 = 2 Then KeyCode = 0: Exit Sub '排除输血科阶段的修改操作
End Sub

Private Sub Chk32_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Chk32_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '在修改状态，只有mlng阶段=1和mlng状态=1时才可以修改,新增状态也可以修改，其他状态都不允许修改
    If Button = 2 Then Exit Sub ''点击右键，还是可以变更checkbox的状态，所以这里点击右键就跳出事件
    '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许变更chk32控件的内容   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then
'    If mstrST = 修改 And mlng阶段 = 2 And mbln输血科新增权限 = False Then '输血科点击chk32时,如果没有输血科新增权限则，不允许其修改   余浪改
        Chk32(Index).Value = Val(Chk32(Index).Tag)
        Exit Sub
    End If
    If mstrST = 修改 Or mstrST = 新增 Then
        Chk32(Index).Tag = Chk32(Index).Value
    Else
        Chk32(Index).Value = Val(Chk32(Index).Tag)
    End If
End Sub

Private Sub Chk33_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And mstrST <> 新增 And mstrST <> 修改 Then KeyCode = 0: Exit Sub '排除新增和修改其外的操作
'    If mstrST = 修改 And mlng阶段 = 2 And mbln输血科新增权限 = False Then KeyCode = 0: Exit Sub '排除输血科阶段在没有新增权限下的修改操作  余浪改
        '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许变更chk33控件的内容   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then KeyCode = 0: Exit Sub
End Sub

Private Sub Chk33_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Chk33_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '在修改状态，只有mlng阶段=1和mlng状态=1时才可以修改,新增状态也可以修改，其他状态都不允许修改
    If Button = 2 Then Exit Sub '点击右键，还是可以变更checkbox的状态，所以这里点击右键就跳出事件
    '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许变更chk33控件的内容   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then
'    If mstrST = 修改 And mlng阶段 = 2 And mbln输血科新增权限 = False Then   '输血科点击chk33时，如果没有新增权限则不允许其修改  余浪改
        Chk33(Index).Value = Val(Chk33(Index).Tag)
        Exit Sub
    End If
    If mstrST = 修改 Or mstrST = 新增 Then
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
    If (mstrST = 新增 And mDataChanged = True) Or (mstrST = 修改 And mlng阶段 = 1 And mlng状态 = 0) Or (mstrST = 修改 And mlng阶段 = 2 And mlng状态 = 0 And mbln输血科新增权限 = True) Then '余浪改  有新增权限的输血科可以进行修改操作。
'    If mstrST <> 新增 Then Exit Sub
        Msk52(Index).Visible = True
        Msk51(Index).Visible = True
        Msk51(Index).Text = Format(Split(lbl5(Index).Caption, " ")(0), "YYYY-MM-DD")
        Msk52(Index).Text = Format(Split(lbl5(Index).Caption, " ")(1), "HH:mm:ss")
        Msk51(Index).ZOrder 0
        Msk52(Index).ZOrder 0
    End If
End Sub

Private Sub lbl6_DblClick(Index As Integer)
    If (mstrST = 新增 And mDataChanged = True) Or (mstrST = 修改 And mlng阶段 = 1 And mlng状态 = 0) Or (mstrST = 修改 And mlng阶段 = 2 And mlng状态 = 0 And mbln输血科新增权限 = True) Then '余浪改  有新增权限的输血科可以进行修改操作。 '新增或者医生阶段未提交数据且是处于修改操作时允许修改反应时间
    '    If mstrST <> 新增 Then Exit Sub
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
    Dim str名称 As String
    Dim str血袋编号 As String
    Dim str孕产情况 As String
    Dim str即往输血史 As String
    Dim lngIndex As Long
    Dim str诊断信息 As String
    Dim rsTmp As Recordset
    lngIndex = 5
    
    If Val(mlng病人ID) = 0 Then MsgBox "无病人信息！", vbInformation, gstrSysName: Exit Sub
    If (mstrST = 修改 And mlng阶段 = 1 And mlng状态 = 0) Or (mstrST = 新增 And mDataChanged = True) Then  '只有当用户为医生，阶段为修改或新增时才允许双击txt21操作
        pic3.Visible = True
        Call GetSecondUserName(TXT21(lngIndex), 1, mlng病人ID, mlng主页id, mlng病人来源, lngID, str名称, str血袋编号, str即往输血史, str孕产情况)
        TXT21(lngIndex).Text = str血袋编号
        TXT21(3).Text = str名称
        cbo1.Text = str血袋编号
        If lngID <> 0 Then
            cbo2.Text = lngID
        End If
        If str即往输血史 = "无" Then '既往输血史
            Opt21(0).Value = True
            TXT21(0).Text = ""
        ElseIf str即往输血史 = "有" Then
            Opt21(1).Value = True
            TXT21(0).Text = ""
        End If
        If InStr(1, str孕产情况, "/") <= 0 Then '孕产情况
            TXT21(1).Text = ""
            TXT21(2).Text = ""
        Else
            TXT21(1).Text = IIf(Split(str孕产情况, "/")(0) = "" And Split(str孕产情况, "/")(1) <> "", 1, Split(str孕产情况, "/")(0))
            TXT21(2).Text = Split(str孕产情况, "/")(1) & ""
        End If
        TXT41(0).Text = GetZXR(lngID) '护士(执行人)
'        If Val(lngID) > 0 Then
'            '填写该血袋对应的医嘱的诊断信息
'            str诊断信息 = "select c.内容 from 血液配血记录 a,血液收发记录 b,病人医嘱附件 c where c.医嘱id=a.申请id and c.项目='申请单诊断' and a.id=b.配发id and b.id=[1]"
'            Set rsTmp = gobjDatabase.OpenSQLRecord(str诊断信息, "查询诊断记录", Val(lngID))
'            If rsTmp.EOF = False Then TXT11(6).Text = rsTmp.Fields("内容").Value & ""
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
     '新增页面要修改时间时，双击之前会响应单机时间，导致页面内容清空，这里加上一个判断以解决此类问题
    If mlngSelNum = Index And mstrST = 新增 Then Exit Sub
    Pic5(Index).SetFocus
End Sub

Private Sub lbl6_Click(Index As Integer)
    Dim lngi As Long
    For lngi = 0 To Msk61.Count - 1
        Msk62(lngi).Visible = False
        Msk61(lngi).Visible = False
    Next
    If mlngSelNum = Index And mstrST = 新增 Then Exit Sub
    Pic6(Index).SetFocus
End Sub

Private Sub Opt21_Click(Index As Integer)
    Dim lngTag As Long
    lngTag = Val(Opt21(0).Tag)
    '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许变更opt21控件的内容   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then
'    If mstrST = 修改 And mlng阶段 = 2 And mbln输血科新增权限 = False Then '余浪改
        Opt21(lngTag).Value = True
        Exit Sub
    End If
    If mstrST = 修改 Or (mstrST = 新增 And mDataChanged = True) Then
        Opt21(0).Tag = Index
    Else
        Opt21(lngTag).Value = True
    End If
    If Opt21(0).Value = True Then TXT21(0).Text = "" '输血次数选择无，自动将txt21(0)的内容改为""
End Sub

Private Sub Opt21_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Opt22_Click(Index As Integer)
    'opt22的修改方式和opt21\31\32有所不同
    Dim lngTag As Long
    lngTag = Val(Opt22(0).Tag)
    '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许变更Opt22控件的内容   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then
'    If mstrST = 修改 And mlng阶段 = 2 And mbln输血科新增权限 = False Then '余浪改
        Opt22(lngTag).Value = True
        Exit Sub
    End If
    If mstrST = 修改 Or (mstrST = 新增 And mDataChanged = True) Then
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
    '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许变更Opt31控件的内容   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then
'    If mstrST = 修改 And mlng阶段 = 2 And mbln输血科新增权限 = False Then '余浪改
        If lngTag = 2 Then
            Opt31(0).Value = False
            Opt31(1).Value = False
        Else
            Opt31(lngTag).Value = True
        End If
        Exit Sub
    End If
    If mstrST = 修改 Or mstrST = 新增 Then
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
    '输血科阶段  新增数据时没有新增权限 或者 修改数据时不是输血科新增数据 则不允许变更Opt32控件的内容   余浪改
    If mlng阶段 = 2 And ((mbln输血科新增权限 = False And mstrST = 新增) Or (mstrST = 修改 And lbl输血科新增.Visible = False)) Then
'    If mstrST = 修改 And mlng阶段 = 2 And mbln输血科新增权限 = False Then '余浪改
        If lngTag = 2 Then
            Opt32(0).Value = False
            Opt32(1).Value = False
        Else
            Opt32(lngTag).Value = True
        End If
        Exit Sub
    End If
    If mstrST = 修改 Or mstrST = 新增 Then
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
    '设置滚动条,并对滚动条的位置进行调整
    On Error GoTo Errorhand
    VS1.Visible = False
    HS1.Visible = False
    If pic4.ScaleWidth < pic1.Width Then '除了pic4的高度小于pic1的高度的情况
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
    If pic4.ScaleWidth >= pic1.Width And pic4.ScaleHeight < pic1.Height Then 'pic4的高度小于pic1的高度的情况
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
    '限制输血量
    If CboDW.Text = "U" And Val(TXT21(4).Text) > 5 Then
        MsgBox "输入量不能大于5U!", vbInformation, gstrSysName
    End If
    If CboDW.Text = "ml" And Val(TXT21(4).Text) > 1000 Then
        MsgBox "输入量不能大于1000ml!", vbInformation, gstrSysName
    End If
    If CboDW.Text = "治疗量" And Val(TXT21(4).Text) > 5 Then
        MsgBox "输入量不能大于5个治疗量!", vbInformation, gstrSysName
    End If
    If Right(TXT21(4).Text, 1) = "." And Len(TXT21(4).Text) < 6 Then TXT21(4).Text = TXT21(4).Text & "0"
End Sub

Private Sub TXT31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '问题27554 by lesfeng 2010-01-19 lngTXTProc 修改为glngTXTProc
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
    TXT32.SelFontName = "宋体"
    TXT32.SelFontSize = 9
End Sub

Private Sub TXT33_Validate(Cancel As Boolean)
    TXT33.SelStart = 0
    TXT33.SelLength = Len(TXT33.Text)
    TXT33.SelBold = False
    TXT33.SelFontName = "宋体"
    TXT33.SelFontSize = 9
End Sub

Private Sub TXT41_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '问题27554 by lesfeng 2010-01-19 lngTXTProc 修改为glngTXTProc
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
    If mlng阶段 = 2 Then
        If mbln输血科新增权限 = False Or mstrST = 修改 And lbl输血科新增.Visible = False Then Exit Sub   '输血科阶段没有输血科新增权限则不会跳出选择血袋的按钮   余浪改
    End If
    '只有新增状态会跳出选择血袋按钮
    If (mstrST = 新增 Or mstrST = 修改) And Index = 5 And mDataChanged = True Then   '(mstrST = 修改 And mlng阶段 = 1 And mlng状态 = 0) Or
        pic3.Visible = True
        Call gobjControl.PicShowFlat(pic3, 1)
    End If
    If Index = 4 And mstrST = 修改 Then
        If InStr(1, LCase(TXT21(4).Text), "ml") > 0 Then TXT21(4).Text = Left(TXT21(4).Text, Len(TXT21(4).Text) - 2) '输入量去掉后面的ml单位
    End If
End Sub

Private Sub TXT21_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 4:
            If Not (KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 46 Or (KeyAscii >= 48 And KeyAscii <= 57)) Then KeyAscii = 0: Exit Sub  '输血量的长度不超过6，且只能输入“.”回车和backspace
            If InStr(TXT21(4).Text, ".") > 0 And KeyAscii = 46 Then KeyAscii = 0: Exit Sub '在有小数点的情况下，不能重复输入小数点
            If TXT21(4).Text = "" And KeyAscii = 46 Then KeyAscii = 0: Exit Sub '在没有输入数字前，不能输入“.”小数点
        Case 3, 5:
            If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub '血袋编号和输血项目不能输入除回车键以外的数据
        Case 0, 1, 2
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8) Then KeyAscii = 0: Exit Sub '如果输入不是数字也不是回车也不是backspace则退出
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
        If Val(TXT31.Text) > 24 Then '当只输入小时时，最大只能输入24，其他大于24的数字都为变为24
            TXT31.Text = 24
        End If
    End If
End Sub

Private Sub TXT31_KeyPress(KeyAscii As Integer)
    '只能输入数字和/还有回车和backspace
    If Not ((KeyAscii >= 47 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8) Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        gobjCommFun.PressKey vbKeyTab
    End If
    
    If Opt31(0).Tag <> 1 Then '除非选择了输血后，否则无法输入内容
        KeyAscii = 0
    End If
    If InStr(1, TXT31.Text, "/") > 0 And KeyAscii = 47 Then KeyAscii = 0 '只能输入一个/
    
End Sub

Private Sub TXT32_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0 '输入内容不能有单引号
End Sub

Private Sub TXT33_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0 '输入内容不能有单引号
End Sub

Private Sub TXT41_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        If Index = 1 And (mstrST = 新增 Or mstrST = 修改) Then
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
    Call ExecuteCommand("控件布局")
Errorhand:
End Sub

Private Sub UserControl_Terminate()
    Set mRsBR = Nothing
    Set mRsFY = Nothing
    mDataChanged = False
    mblnAddPage = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '功能：保存相关用户定义属性数据
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
'    Call DestroyCaret'去除光标闪烁
    For lngIndex = 0 To lbl5.Count - 1
        DTP5(lngIndex).Visible = False
        Pic5(lngIndex).BorderStyle = 0
    Next
    Pic5(Index).BorderStyle = 1

    mlngSelNum = Index
    Call ExecuteCommand("刷新数据")
    If TXT21(5).Text = "" Then
        cbo2.Text = ""
    Else
        For lngIndex = 0 To cbo1.ListCount - 1 '这一步是为了让收发id和血袋编号保持一致，以方便数据的增删改
            If cbo1.List(lngIndex) = TXT21(5).Text Then
                cbo2.Text = cbo2.List(lngIndex) & ""
            End If
        Next
    End If
    '对dtp的显示做一定的调整，只有未记录的数据才会显示dtp控件
    Arrtag = Split(lbl5(mlngSelNum).Tag, ":")
    mstrST = Val(Arrtag(1))
    mlng状态 = Val(Arrtag(0))
    mDataChanged = Val(Arrtag(2)) = 1
    If mstrST = 新增 Then
        DTP5(Index).Visible = mDataChanged
        lbl5(mlngSelNum).Caption = Msk51(mlngSelNum).Text & " " & Msk52(mlngSelNum).Text
        Call Clear
    End If
    Call ExecuteCommand("控件状态")
    Call ExecuteCommand("刷新数据")
    lbl输血科新增.Visible = IIf(Val(Msk51(mlngSelNum).Tag) = 1, True, False) '余浪改   根据输血科新增字段来显示或者隐藏输血科新增标志
    mbln输血科新增 = lbl输血科新增.Visible
Errorhand:
End Sub

Private Sub Pic6_GotFocus(Index As Integer)
    'lbl6被选中时
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
    Call ExecuteCommand("刷新数据")
    If TXT21(5).Text = "" Then
        cbo2.Text = ""
    Else
        For lngIndex = 0 To cbo1.ListCount - 1 '这一步是为了让收发id和血袋编号保持一致，以方便数据的增删改
            If cbo1.List(lngIndex) = TXT21(5).Text Then
                cbo2.Text = cbo2.List(lngIndex)
            End If
        Next
    End If

    '对dtp的显示做一定的调整，只有未记录的数据才会显示dtp控件
    Arrtag = Split(lbl6(mlngSelNum).Tag, ":")
    mstrST = Val(Arrtag(1))
    mlng状态 = Val(Arrtag(0))
    mDataChanged = Val(Arrtag(2)) = 1
    If mstrST = 新增 Then
        DTP6(Index).Visible = mDataChanged
        lbl6(mlngSelNum).Caption = Msk61(mlngSelNum).Text & " " & Msk62(mlngSelNum).Text
        Call Clear
    End If
    Call ExecuteCommand("控件状态")
    Call ExecuteCommand("刷新数据") '这里再次刷新数据是为了避免option控件无法变更的问题，如果只有上面一个刷新数据或者只有下面一个刷新数据，都无法做到多option控件的刷新，根本原因可能是控件本身的问题，因为刷新数据只是对option和其他控件的值的一个变更，并未做其他操作。
    If Pic6(Index).Top >= 18 * 450 + lngVsfRight Or Pic6(Index).Top < 0 Then lngVsfRight = (Index - 18) * 450
    If lngVsfRight < 0 Then lngVsfRight = 0
    vsbRight.Value = lngVsfRight
    lbl输血科新增.Visible = IIf(Val(Msk61(mlngSelNum).Tag) = 1, True, False) '余浪改   根据输血科新增字段来显示或者隐藏输血科新增标志
    mbln输血科新增 = lbl输血科新增.Visible
End Sub
Private Sub DTP5_CloseUp(Index As Integer)
    'dtp5选择时间后
'    Pic5(Index).SetFocus
    Msk51(Index).Text = Format(DTP5(Index).Value, "YYYY-MM-DD")
    If IsDate(Msk52(Index).Text) = False Then
        Msk52(Index).Text = Format(Now, "HH:mm:ss")
    End If
    lbl5(Index).Caption = Msk51(Index) & " " & Msk52(Index)
    
End Sub
Private Sub DTP6_CloseUp(Index As Integer)
    'dtp6选择时间后
'    pic6(Index).SetFocus
    Msk61(Index).Text = Format(DTP6(Index).Value, "YYYY-MM-DD")
    If IsDate(Msk62(Index).Text) = False Then
        Msk62(Index).Text = Format(Now, "HH:mm:ss")
    End If
    lbl6(Index).Caption = Msk61(Index) & " " & Msk62(Index)
End Sub

Private Function GetSecondUserName(ByVal objControl As TextBox, ByVal lngDeptID As Long, ByVal lng病人ID As Long, lng主页id As Long, lng病人来源 As Long, lngID As Long, str名称 As String, str血袋编号 As String, str即往输血史 As String, str孕产情况 As String) As Boolean
    '功能：根据相关查找语句查询数据，并显示在一个窗体上，且可以规定窗体的显示位置和模式，注：查询的数据必须要有id
    
    Dim rsUser As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim vPoint As POINTAPI, blnCancel As Boolean

    On Error GoTo ErrHand
    vPoint = GetCoordPos(UserControl.hWnd, TXT21(5).Left + pic4.Left + pic1.Left, TXT21(5).Top + pic4.Top + pic1.Top)
    '余浪
    strSQL = " Select id,血液名称,血袋编号,医嘱号,decode(即往输血史,0,'无','有') as 即往输血史, 孕产情况, 血液状态 " & _
            " From (SELECT e.Id AS Id, b.名称 AS 血液名称, e.血袋编号 AS 血袋编号, a.Id AS 医嘱号, c.即往输血史 AS 即往输血史, c.孕产情况," & vbNewLine & _
            "       Decode(f.收发id, NULL, 1, f.接收状态) AS 接收状态," & vbNewLine & _
            "       Decode(Nvl(f.执行状态, 0), 0, '已接收', 1, '正在执行', 2, '完成执行', 3, '停止执行') 血液状态" & vbNewLine & _
            "FROM 输血申请记录 c, 血液发送记录 f, 诊疗项目目录 b, 血液规格 g, 血液收发记录 e, 血液配血记录 d, 病人医嘱记录 a" & vbNewLine & _
            "WHERE c.医嘱id(+) = a.Id AND a.相关id IS NULL AND NOT EXISTS" & vbNewLine & _
            " (SELECT 1 FROM 输血反应记录 h WHERE e.Id = h.收发id" & IIf(mstrST = 修改, " And h.收发ID<>[3]", " ") & ") AND e.Id = f.收发id(+) AND g.品种id = b.Id AND g.规格id = e.血液id AND" & vbNewLine & _
            "      e.审核人 IS NOT NULL AND MOD(e.记录状态, 3) = 1 AND e.配发id = d.Id AND d.记录性质 = 1 AND d.申请id = a.Id AND a.诊疗类别 = 'K' AND" & vbNewLine & _
            "      a.病人id = [1] "
    If mlng收发ID = 0 Then
    If lng病人来源 = 2 Then
        strSQL = strSQL & " And a.主页id = [2] ) Where 接收状态 in(1,3)"
        Set rsUser = gobjDatabase.ShowSQLSelect(mobjfrm, strSQL, 0, "输血反应", False, "", "", False, False, True, vPoint.X, vPoint.Y, TXT21(5).Height, blnCancel, False, True, lng病人ID, lng主页id, Val(cbo2.Text))
    Else
        strSQL = strSQL & " And a.挂号单 = [2] ) Where 接收状态 in(1,3)"
        Set rsUser = gobjDatabase.ShowSQLSelect(mobjfrm, strSQL, 0, "输血反应", False, "", "", False, False, True, vPoint.X, vPoint.Y, TXT21(5).Height, blnCancel, False, True, lng病人ID, mstr挂号单, Val(cbo2.Text))
        End If
    Else
        strSQL = strSQL & " and e.id = [4] "
        If lng病人来源 = 2 Then
            strSQL = strSQL & " And a.主页id = [2] )" & IIf(gbln接收后才能执行 = True, " Where 接收状态 in(1,3)", "")
            Set rsUser = gobjDatabase.OpenSQLRecord(strSQL, "输血反应", lng病人ID, lng主页id, Val(cbo2.Text), mlng收发ID)
        Else
            strSQL = strSQL & " And a.挂号单 = [2] ) " & IIf(gbln接收后才能执行 = True, " Where 接收状态 in(1,3)", "")
            Set rsUser = gobjDatabase.OpenSQLRecord(strSQL, "输血反应", lng病人ID, mstr挂号单, Val(cbo2.Text), mlng收发ID)
        End If
        mlng收发ID = 0
    End If
    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then Exit Function
            lngID = Val(rsUser!id)
            str名称 = Nvl(rsUser!血液名称) & ""
            str即往输血史 = rsUser!即往输血史 & ""
            str孕产情况 = rsUser!孕产情况 & ""
            str血袋编号 = Nvl(rsUser!血袋编号)
            GetSecondUserName = True
            
            strSQL = "select c.内容 from 血液配血记录 a,血液收发记录 b,病人医嘱附件 c where c.医嘱id=a.申请id and c.项目='申请单诊断' and a.id=b.配发id and b.id=[1]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "查询诊断记录", lngID)
            If rsTmp.EOF = False Then
                TXT11(6).Text = rsTmp.Fields("内容").Value & ""
            Else
                TXT11(6).Text = ""
            End If
        End If
    ElseIf blnCancel = False Then
        MsgBox "无输血记录！", vbInformation, gstrSysName
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

Private Function GetZXR(lng收发ID As Long) As String
    '功能：通过收发id提取执行人
    '参数：lng收发id-收发id
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    strSQL = " Select 执行人 开始执行人 From 血液执行记录 where 收发id=[1] and 记录性质=1 and 序号=0"
    Set rsUser = gobjDatabase.OpenSQLRecord(strSQL, "查询执行人", lng收发ID)
    If rsUser.EOF = False Then GetZXR = rsUser.Fields("开始执行人").Value & "": Exit Function
    GetZXR = ""
End Function

Private Function FindHS(strName As String) As String
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    
    vPoint = GetCoordPos(UserControl.hWnd, TXT41(1).Left + pic4.Left + pic1.Left, TXT41(1).Top + pic1.Top + pic4.Top)
    strSQL = " Select distinct a.Id ,a.编号, a.姓名,a.简码 " & vbNewLine & _
             " From 人员表 a, 人员性质说明 b " & vbNewLine & _
             " Where a.Id = b.人员id And b.人员性质 = '医生' And (a.姓名 Like [1] or a.编号 like [1] or a.简码 like [1])"
    
    Set rsUser = gobjDatabase.ShowSQLSelect(mobjfrm, strSQL, 0, "输血反应", False, "", "", False, False, True, vPoint.X, vPoint.Y, TXT41(1).Height, blnCancel, False, True, strName & "%")

    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then FindHS = "": Exit Function
            FindHS = rsUser!姓名
            Exit Function
        End If
    End If
    FindHS = ""
    
End Function

