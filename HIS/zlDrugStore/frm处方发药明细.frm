VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm������ҩ��ϸ 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picRecipt 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6375
      ScaleWidth      =   11775
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.PictureBox picRecInfo 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   1515
         Left            =   0
         ScaleHeight     =   1515
         ScaleWidth      =   10755
         TabIndex        =   24
         Top             =   0
         Width           =   10755
         Begin VB.PictureBox picRecipeColor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   460
            Left            =   120
            ScaleHeight     =   465
            ScaleWidth      =   1095
            TabIndex        =   27
            Top             =   50
            Width           =   1095
            Begin VB.Label lblRecipeType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ͨ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   240
               TabIndex        =   28
               Top             =   105
               Width           =   600
            End
         End
         Begin VB.ComboBox TxtNo 
            Height          =   300
            ItemData        =   "frm������ҩ��ϸ.frx":0000
            Left            =   8085
            List            =   "frm������ҩ��ϸ.frx":0002
            TabIndex        =   26
            Top             =   280
            Width           =   1965
         End
         Begin VB.TextBox txt������� 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   2400
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   715
            Width           =   7695
         End
         Begin VB.Label LblTel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "15310625533"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   6000
            TabIndex        =   52
            Tag             =   "���䣺"
            Top             =   45
            Width           =   1155
         End
         Begin VB.Label LblTel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�绰��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   5400
            TabIndex        =   51
            Tag             =   "���䣺"
            Top             =   45
            Width           =   585
         End
         Begin VB.Label lblNotice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩƷ˵����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Left            =   1440
            TabIndex        =   50
            Tag             =   "��ҩ�巨:"
            Top             =   1140
            Width           =   1365
         End
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ٴ���ϣ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Left            =   1440
            TabIndex        =   49
            Tag             =   "�ٴ���ϣ�"
            Top             =   715
            Width           =   975
         End
         Begin VB.Label Lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���ң�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   48
            Tag             =   "���ң�"
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Lbl�Ա� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   3120
            TabIndex        =   47
            Tag             =   "�Ա�"
            Top             =   45
            Width           =   585
         End
         Begin VB.Label Lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���䣺"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   4200
            TabIndex        =   46
            Tag             =   "���䣺"
            Top             =   45
            Width           =   585
         End
         Begin VB.Label LblסԺ�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ʶ�ţ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   5400
            TabIndex        =   45
            Tag             =   "��ʶ�ţ�"
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���ţ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   3120
            TabIndex        =   44
            Tag             =   "���ţ�"
            Top             =   330
            Width           =   585
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   43
            Tag             =   "������"
            Top             =   45
            Width           =   585
         End
         Begin VB.Label Lbl�շ�Ա 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�շ�Ա��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   9165
            TabIndex        =   42
            Tag             =   "�շ�Ա:"
            Top             =   45
            Width           =   780
         End
         Begin VB.Label LblNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���ݺ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Left            =   7320
            TabIndex        =   41
            Top             =   330
            Width           =   585
         End
         Begin VB.Label lblҩ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҩ��"
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
            Height          =   180
            Left            =   120
            TabIndex        =   40
            Top             =   675
            Width           =   390
         End
         Begin VB.Label lbl���￨�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���￨��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   7320
            TabIndex        =   39
            Tag             =   "���￨��"
            Top             =   45
            Width           =   780
         End
         Begin VB.Label Lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ٿ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   1980
            TabIndex        =   38
            Top             =   330
            Width           =   585
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
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
            Index           =   1
            Left            =   1980
            TabIndex        =   37
            Top             =   45
            Width           =   585
         End
         Begin VB.Label Lbl�Ա� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   3720
            TabIndex        =   36
            Tag             =   "�Ա�"
            Top             =   45
            Width           =   195
         End
         Begin VB.Label Lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "33"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   3720
            TabIndex        =   35
            Tag             =   "���ţ�"
            Top             =   330
            Width           =   210
         End
         Begin VB.Label Lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "22��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   4800
            TabIndex        =   34
            Tag             =   "���䣺"
            Top             =   45
            Width           =   405
         End
         Begin VB.Label LblסԺ�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1234567"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   6120
            TabIndex        =   33
            Tag             =   "��ʶ�ţ�"
            Top             =   330
            Width           =   735
         End
         Begin VB.Label lbl���￨�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "123456789"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   8040
            TabIndex        =   32
            Tag             =   "���￨��"
            Top             =   45
            Width           =   945
         End
         Begin VB.Label Lbl�շ�Ա 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���ʥ "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   9960
            TabIndex        =   31
            Tag             =   "�շ�Ա:"
            Top             =   45
            Width           =   690
         End
         Begin VB.Label LblWeight 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "55kg"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   4800
            TabIndex        =   30
            Tag             =   "���أ�"
            Top             =   330
            Width           =   420
         End
         Begin VB.Label LblWeight 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���أ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   4200
            TabIndex        =   29
            Tag             =   "���ţ�"
            Top             =   330
            Width           =   585
         End
      End
      Begin VB.PictureBox picMark1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   10920
         ScaleHeight     =   855
         ScaleWidth      =   855
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         Begin VB.PictureBox picMark2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   695
            Left            =   80
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   19
            Top             =   80
            Width           =   695
            Begin VB.Label lblMark 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   24
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   600
               Left            =   100
               TabIndex        =   20
               Top             =   0
               Width           =   615
            End
         End
      End
      Begin VB.PictureBox picHscSend 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   4575
         TabIndex        =   16
         Tag             =   "0"
         Top             =   3480
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label lblDiag 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "����ҩ�������Ϣ"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1200
            TabIndex        =   17
            Top             =   30
            Width           =   2400
         End
         Begin VB.Image imgDown 
            Height          =   240
            Left            =   0
            Picture         =   "frm������ҩ��ϸ.frx":0004
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgUp 
            Height          =   240
            Left            =   0
            Picture         =   "frm������ҩ��ϸ.frx":0346
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frm������ҩ��ϸ.frx":0688
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfNoList 
         Height          =   1005
         Left            =   7560
         TabIndex        =   12
         Top             =   2760
         Visible         =   0   'False
         Width           =   2820
         _cx             =   4974
         _cy             =   1773
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm������ҩ��ϸ.frx":0BD6
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
         ExplorerBar     =   1
         PicturesOver    =   -1  'True
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   1296
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
         FormatString    =   $"frm������ҩ��ϸ.frx":0D2E
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   1800
         Left            =   3840
         TabIndex        =   9
         Top             =   2040
         Width           =   3720
         _cx             =   6562
         _cy             =   3175
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
         BackColorSel    =   16769992
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   315
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm������ҩ��ϸ.frx":0D7C
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
      Begin VB.PictureBox picProcess 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   10935
         TabIndex        =   4
         Top             =   5400
         Width           =   10935
         Begin VB.CommandButton cmdSendByNoTake 
            Caption         =   "����δȡҩ��ҩ(&T)"
            Height          =   350
            Left            =   7320
            TabIndex        =   23
            ToolTipText     =   "�ȼ���F2"
            Top             =   360
            Width           =   1815
         End
         Begin VB.ComboBox cbo��ҩ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   810
            TabIndex        =   22
            Text            =   "cbo��ҩ��"
            Top             =   400
            Width           =   1815
         End
         Begin VB.ComboBox cbo����ҽ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   810
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   30
            Width           =   1815
         End
         Begin VB.ComboBox cbo�˲��� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3480
            TabIndex        =   13
            Text            =   "cbo�˲���"
            Top             =   400
            Width           =   1695
         End
         Begin VB.CommandButton CmdSend 
            Caption         =   "��ҩ(&S)"
            Height          =   350
            Left            =   9360
            TabIndex        =   6
            ToolTipText     =   "�ȼ���F2"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Chkȫ�� 
            Appearance      =   0  'Flat
            Caption         =   "ȫ��"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5400
            TabIndex        =   5
            Top             =   445
            Value           =   1  'Checked
            Width           =   765
         End
         Begin VB.Label lbl�˲��� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�˲���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2880
            TabIndex        =   14
            Top             =   460
            Width           =   540
         End
         Begin VB.Label Lbl����ҽ�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   8
            Top             =   85
            Width           =   720
         End
         Begin VB.Label Lbl��ҩ�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   7
            Top             =   460
            Width           =   540
         End
      End
      Begin VB.TextBox txt��ҩ���� 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   3840
         Width           =   9975
      End
      Begin VB.PictureBox picRecInfo_CM 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   7455
         TabIndex        =   1
         Top             =   4800
         Width           =   7455
         Begin VB.Label lblԭʼ���� 
            AutoSize        =   -1  'True
            Caption         =   "ԭʼ������"
            Height          =   180
            Left            =   0
            TabIndex        =   3
            Tag             =   "ԭʼ����:"
            Top             =   60
            Width           =   900
         End
         Begin VB.Label lbl��ҩ�巨 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ�巨��"
            Height          =   180
            Left            =   1830
            TabIndex        =   2
            Tag             =   "��ҩ�巨:"
            Top             =   60
            Width           =   900
         End
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1920
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":0DF1
            Key             =   "��ӡ11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":118B
            Key             =   "��ǰ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":79ED
            Key             =   "ָʾ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":E24F
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":E7E9
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":EB83
            Key             =   "��־"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":EF1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":F2B7
            Key             =   "ͼ��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":F651
            Key             =   "ѡ��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":10063
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":168C5
            Key             =   "δ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":1D127
            Key             =   "�ڼ�"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":23989
            Key             =   "�Ѽ�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":2A1EB
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":30A4D
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":30DE7
            Key             =   "����_ѡ��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":31181
            Key             =   "�ײ�"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":379E3
            Key             =   "����"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":3E245
            Key             =   "��Ƭ"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":44AA7
            Key             =   "����"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":4B309
            Key             =   "ָ��"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":51B6B
            Key             =   "���"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":583CD
            Key             =   "������ʽ"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":5EC2F
            Key             =   "�����ļ�"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":65491
            Key             =   "����"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":6BCF3
            Key             =   "�շ�"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":6C705
            Key             =   "���"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":72F67
            Key             =   "����"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":797C9
            Key             =   "ȷ��"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8002B
            Key             =   "��ʼ"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8688D
            Key             =   "����"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8D0EF
            Key             =   "����"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8D489
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8D823
            Key             =   "�����ܼ�"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8DBBD
            Key             =   "ȫ���ܼ�"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8DF57
            Key             =   "�ܼ�"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8E2F1
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8ED03
            Key             =   "�Ѿ���ӡ"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":8F715
            Key             =   "ҩƷ"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":95F77
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgCheck 
      Left            =   1200
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":96511
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":96AAB
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ��ϸ.frx":97045
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   6240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frm������ҩ��ϸ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�������˵�
Private Const conMenu_Tool_ShowPlug = 101           '���ò����������ҩ
Private mInt�ɲ��� As Integer

Private mlngMode As Long

'����
Private Type Type_ShowListCondition
    intListType As Integer                          '0-����ҩ;1-����ҩ;2-����ҩ;3-��ҩ
    bln��ʾ���� As Boolean
    intShowPass As Integer
    bln��ʾ��С��λ As Boolean
    bln��ʾ���̵��� As Boolean
    bln��ʾ���� As Boolean
    blnУ�鴦�� As Boolean                            '�Ƿ���ҪУ�鴦��
    lngҩ��ID As Long
    blnҽ������ As Boolean
    str��ҩ�� As String
    str�˲��� As String
    bln�Ƿ���Ҫ��ҩ���� As Boolean
    bln�Զ���ҩ As Boolean
    int����� As Integer
    bln����ģʽ As Boolean
    int�����ʾ As Integer                          '�����ʾ��ʽ��0-��ʾӦ�ս��,1-��ʾʵ�ս��,2-��ʾӦ�պ�ʵ�ս��
    blnȡҩȷ�� As Boolean          '�Ƿ����ò���ʵ��ȡҩȷ��ģʽ��0-�����ã�1-����
    bln������� As Boolean
    bln����˲��˺���ҩ����ͬ As Boolean
    bln��ʾԭ���� As Boolean
End Type
Private mcondition As Type_ShowListCondition

Private mbln��ҩ���� As Boolean
Private mstrDosUser As String
Private mstr�˲��� As String

Private mblnAllowClick As Boolean                        '����ִ��Click�¼�
Public mblnInput As Boolean                             '�Ƿ���ͨ��¼�뷽ʽ�����˴���
Private mblnδȡҩ��ҩ As Boolean                       'δȡҩ��ҩģʽ

Private mstrPrivs As String
Private mstrRecipeInfo As String                         '������Ϣ������;������;��¼����;�����־;ҩ��ID

'�б�����
Private Enum mListType
    ��ҩȷ�� = 0
    ����ҩ = 1
    ����ҩ = 2
    ����ҩ = 3
    ��ʱδ�� = 4
    ��ҩ = 5
End Enum

'�������ͣ���ͨ�����ơ������������һ������
Private Enum ��������
    ��ͨ = 0
    ���� = 1
    ���� = 2
    ���� = 3
    ��һ = 4
    ���� = 5
End Enum

'�û�����Ĵ�����ɫ����ע���ȡ���ַ�������;�ָ�
Private mstrUserRecipeColor As String

Private mrsDetail As ADODB.Recordset

Private Const mlng��ɫ As Long = &HC000C0
Private mblnResize As Boolean

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private mstrOracleMoneyForamt As String                 'ORACLE�н���ʽ
Private mstrVBMoneyForamt As String                     'VB�н���ʽ

Private mstrUnallowSetColHide As String         '�������������ص���
Private mstrUnallowShow As String                   '��������ʾ����

Private mblnAllBack As Boolean

Private Const mconIntCol���� = 68
Private mIntCol��ǰ�� As Integer
Private mIntCol˳��� As Integer
Private mIntCol����� As Integer
Private mIntColҩƷ���� As Integer
Private mintColƤ�Խ�� As Integer
Private mIntCol������ As Integer
Private mIntColӢ���� As Integer
Private mIntCol�䷽���� As Integer
Private mintcol��� As Integer
Private mintcol��� As Integer
Private mintcol���� As Integer
Private mintcolЧ�� As Integer
Private mIntColId As Integer
Private mintcolҩƷid As Integer
Private mintcol���� As Integer
Private mintcol��λ As Integer
Private mIntCol���� As Integer
Private mIntCol���� As Integer
Private mintcol���� As Integer
Private mIntCol��� As Integer
Private mIntColʵ�ս�� As Integer
Private mIntCol���� As Integer
Private mIntCol���� As Integer
Private mIntCol�÷� As Integer
Private mIntColƵ�� As Integer
Private mIntCol��ҩĿ�� As Integer
Private mIntCol����˵�� As Integer
Private mIntColҽ������ As Integer
Private mIntCol�ѱ� As Integer
Private mIntCol����� As Integer
Private mIntCol��λ As Integer
Private mIntCol������ As Integer
Private mIntCol׼���� As Integer
Private mIntCol׼������ As Integer
Private mIntCol׼����С As Integer
Private mIntCol��ҩ�� As Integer
Private mIntCol��ҩ���� As Integer
Private mIntCol��λ�� As Integer
Private mIntCol��ҩ��С As Integer
Private mIntCol��λС As Integer
Private mIntCol���� As Integer
Private mIntCol������ As Integer
Private mIntCol��Ч�� As Integer
Private mIntCol�²��� As Integer
Private mIntCol��ע As Integer
Private mIntColҽ��id As Integer
Private mIntColʵ������ As Integer
Private mIntCol��װ As Integer
Private mIntCol���� As Integer
Private mIntColNO As Integer
Private mIntCol�����־ As Integer
Private mIntCol��¼���� As Integer
Private mIntCol�ⷿID As Integer
Private mIntCol��ҩ���� As Integer
Private mIntCol���id As Integer
Private mIntCol����ҽ�� As Integer
Private mIntColƵ�ʼ�� As Integer
Private mIntCol�����λ As Integer
Private mIntColҽ����־ As Integer
Private mIntCol��ʼʱ�� As Integer
Private mIntCol����ʱ�� As Integer
Private mIntColƵ�ʴ��� As Integer
Private mIntCol���� As Integer
Private mIntCol����� As Integer
Private mIntColסԺ�� As Integer
Private mIntCol����ҩƷ˵�� As Integer
Private mintcol������ As Integer
Private mintcolԭ���� As Integer
Public Sub CmdProcess()
    If CmdSend.Enabled Then CmdSend_Click
End Sub

Public Sub FormClear()
    Me.lbl����(1).Caption = ""
    Me.Lbl����(1).Caption = ""
    Me.Lbl����(1).Caption = ""
    Me.Lbl����(1).Caption = ""
    Me.Lbl����(1).Caption = ""
    Me.lbl���￨��(1).Caption = ""
    Me.Lbl�շ�Ա(1).Caption = ""
    Me.Lbl�Ա�(1).Caption = ""
    Me.LblסԺ��(1).Caption = ""
    Me.LblWeight(1).Caption = ""
    LblTel(1).Caption = ""
    
    Me.txt�������.Text = ""
    Me.txt�������.Tag = ""
    
    txtNo.Clear
    txtNo.Tag = ""
    Lblҩ��.Caption = ""
    
    vsfList.rows = 1
    vsfList.rows = 2
    
    Me.lblԭʼ����.Caption = Me.lblԭʼ����.Tag
    Me.lbl��ҩ�巨.Caption = Me.lbl��ҩ�巨.Tag
    lblNotice.Caption = "����ҩƷ˵����"
    
    CmdSend.Enabled = False
End Sub

Public Function GetDetailList() As VSFlexGrid
    Set GetDetailList = vsfList
End Function

Private Sub GetRecipeByNO()
    If mblnAllowClick = False Then Exit Sub
    If txtNo.ListIndex = -1 Then Exit Sub
    
    If mcondition.intListType <> mListType.��ҩ Then
        If zlStr.IsHavePrivs(mstrPrivs, "������ҩ���Ĵ���") = True Then
            If mcondition.lngҩ��ID <> Val(Split(txtNo.Tag, "|")(0)) Then
                If Val(Split(txtNo.Tag, "|")(5)) <> 1 Then
                    MsgBox "[" & Mid(txtNo.Text, 1, 8) & "]�Ѿ����й���ҩ����,���ܽ��д����������뵽" & Split(txtNo.Tag, "|")(1) & "ȡҩ��", vbInformation + vbOKOnly, gstrSysName
                    DoEvents
                    txtNo.Clear
                    txtNo.Text = ""
                    txtNo.Tag = ""
                    Lblҩ��.Caption = ""
                    txtNo.SetFocus
                    Exit Sub
                End If
                If CDate(Format(Split(txtNo.Tag, "|")(4), "yyyy-MM-dd")) < CDate(Format(Sys.Currentdate, "yyyy-MM-dd")) - 30 Then
                    MsgBox "[" & Mid(txtNo.Text, 1, 8) & "]����" & Split(txtNo.Tag, "|")(1) & "30�����ڵĵĴ���,���ܽ��д����������뵽" & Split(txtNo.Tag, "|")(1) & "ȡҩ��", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
                
                If MsgBox("[" & Mid(txtNo.Text, 1, 8) & "]��" & Split(txtNo.Tag, "|")(1) & "�Ĵ������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    DoEvents
                    txtNo.Clear
                    txtNo.Text = ""
                    txtNo.Tag = ""
                    Lblҩ��.Caption = ""
                    txtNo.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    
    If (frmҩƷ������ҩNew.Dtp��ʼʱ�� > CDate(Split(txtNo.Tag, "|")(4)) Or frmҩƷ������ҩNew.cboʱ�䷶Χ.ListIndex <> 3 Or (Split(txtNo.Tag, "|")(5) <> 1 And frmҩƷ������ҩNew.Chk��ʾ��ҩ��������.Value <> 1)) And Not mcondition.bln����ģʽ And mcondition.lngҩ��ID = Val(Split(txtNo.Tag, "|")(0)) Then
        If zlStr.IsHavePrivs(mstrPrivs, "�����ѯ����ʱ�䷶Χ����") And Not (Not zlStr.IsHavePrivs(mstrPrivs, "�޸Ĺ�������") And mcondition.intListType = mListType.��ҩ) Then
            frmҩƷ������ҩNew.cboʱ�䷶Χ.ListIndex = 3
            If frmҩƷ������ҩNew.Dtp��ʼʱ�� > CDate(Split(txtNo.Tag, "|")(4)) Then frmҩƷ������ҩNew.Dtp��ʼʱ�� = CDate(Split(txtNo.Tag, "|")(4))
            If frmҩƷ������ҩNew.Chk��ʾ��ҩ��������.Value <> 1 And Split(txtNo.Tag, "|")(5) <> 1 And mcondition.intListType <> mListType.��ҩ Then frmҩƷ������ҩNew.Chk��ʾ��ҩ��������.Value = 1
            frmҩƷ������ҩNew.GetCondition
            If mcondition.intListType = mListType.��ҩ Then
                frmҩƷ������ҩNew.RefreshList_Return Mid(txtNo.Text, 1, 8), True
            ElseIf mcondition.intListType = mListType.����ҩ Then
                frmҩƷ������ҩNew.RefreshList_Send Mid(txtNo.Text, 1, 8), True
            ElseIf mcondition.intListType = mListType.��ʱδ�� Then
                frmҩƷ������ҩNew.RefreshList_OverTime Mid(txtNo.Text, 1, 8), True
            Else
                frmҩƷ������ҩNew.RefreshList_Dosage Mid(txtNo.Text, 1, 8), True
            End If
        End If
    End If
    
    frmҩƷ������ҩNew.FindListRow 1, Mid(txtNo.Text, 1, 8), Mid(txtNo.Text, 11)

    DoEvents

    If mcondition.intListType = mListType.��ҩ Then
        frmҩƷ������ҩNew.RefreshDetail_Return Val(txtNo.ItemData(txtNo.ListIndex)), Mid(txtNo.Text, 1, 8), "", 1, Val(Split(txtNo.Tag, "|")(2)), Val(Split(txtNo.Tag, "|")(3)), True
    Else
        frmҩƷ������ҩNew.RefreshDetail_Send Val(Split(txtNo.Tag, "|")(0)), Val(txtNo.ItemData(txtNo.ListIndex)), Mid(txtNo.Text, 1, 8), Val(Split(txtNo.Tag, "|")(2)), Val(Split(txtNo.Tag, "|")(3))
    End If

    If CmdSend.Enabled = True Then
        CmdSend.SetFocus
        mblnInput = True
    End If
End Sub

Public Function Get����ҽ��() As String
    If cbo����ҽ��.ListIndex = -1 Then
        Get����ҽ�� = ""
    ElseIf InStr(cbo����ҽ��.Text, "-") > 0 Then
        Get����ҽ�� = Mid(cbo����ҽ��.Text, InStr(cbo����ҽ��.Text, "-") + 1)
    Else
        Get����ҽ�� = cbo����ҽ��.Text
    End If
End Function

Public Function Get��ҩ��() As String
    If Cbo��ҩ��.ListIndex = -1 Then
        Get��ҩ�� = ""
    ElseIf InStr(Cbo��ҩ��.Text, "-") > 0 Then
        Get��ҩ�� = Mid(Cbo��ҩ��.Text, InStr(Cbo��ҩ��.Text, "-") + 1)
    Else
        Get��ҩ�� = Cbo��ҩ��.Text
    End If
End Function
Public Function Get�˲���() As String
    If cbo�˲���.ListIndex = -1 Then
        Get�˲��� = ""
    ElseIf InStr(cbo�˲���.Text, "-") > 0 Then
        Get�˲��� = Mid(cbo�˲���.Text, InStr(cbo�˲���.Text, "-") + 1)
    Else
        Get�˲��� = cbo�˲���.Text
    End If
End Function


Private Sub Loadҽ��()
    Dim rsData As ADODB.Recordset
    
    Set rsData = RecipeSendWork_Getҽ��
    
    Me.cbo����ҽ��.Clear
    cbo����ҽ��.AddItem ""
    Do While Not rsData.EOF
        cbo����ҽ��.AddItem rsData!ҽ��
        rsData.MoveNext
    Loop
    cbo����ҽ��.ListIndex = 0
End Sub

Private Sub SetCmdSendPrivs(ByVal int����� As Integer)
    'Ȩ�޿���
    Select Case mcondition.intListType
    Case mListType.��ҩȷ��
       '��ҩȷ��
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "��ҩȷ��")
    Case mListType.����ҩ
        '��ҩ
        CmdSend.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "��ҩ") And mcondition.bln�Զ���ҩ = False And (mcondition.bln������� = False Or (mcondition.bln������� = True And int����� = 1)))
    Case mListType.����ҩ
        'ȡ��
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "��ҩ")
    Case mListType.����ҩ, mListType.��ʱδ��
        '��ҩ
        CmdSend.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "��ҩ") And (mcondition.bln������� = False Or (mcondition.bln������� = True And int����� = 1)))
    Case mListType.��ҩ
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "��ҩ")
        If mcondition.intListType = mListType.��ҩ And mInt�ɲ��� = 1 And zlStr.IsHavePrivs(mstrPrivs, "��ҩ") Then
            CmdSend.Enabled = True
        Else
            CmdSend.Enabled = False
        End If
    End Select
End Sub

Public Sub SetParams()
    Dim bln�Ƿ���ҩȷ�� As Boolean
    
    With mcondition
        .bln��ʾ���� = (Val(zldatabase.GetPara("��ʾ����", glngSys, 1341)) = 1)
        .intShowPass = gintPass
        .bln��ʾ��С��λ = (Val(zldatabase.GetPara("��ʾ��С��λ", glngSys, 1341)) = 1)
        .blnУ�鴦�� = IsInString(gstrprivs, "У�鴦��", ";")
        .blnҽ������ = (gtype_UserSysParms.P68_����ҩ�������Ϻ���ҩ = 0)
        .str��ҩ�� = zldatabase.GetPara("��ҩ��", glngSys, 1341)
        .bln�Զ���ҩ = (Val(zldatabase.GetPara("�Զ���ҩ", glngSys, 1341)) = 1)
        .bln����ģʽ = (Val(GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "���涨λ", 0)) = 1)
        .int�����ʾ = Val(zldatabase.GetPara("�����ʾ��ʽ", glngSys, 1341, 0))
        .blnȡҩȷ�� = (Val(zldatabase.GetPara("���ò���ʵ��ȡҩȷ��ģʽ", glngSys, 1341, 0)) = 1)
        .bln������� = ((gtype_UserSysParms.P240_ҩ��������� = 1 Or gtype_UserSysParms.P240_ҩ��������� = 3) And gtype_UserSysParms.P241_�������ʱ�� = 2)
        .str�˲��� = zldatabase.GetPara("�˲���", glngSys, 1341)
        .bln����˲��˺���ҩ����ͬ = (Val(zldatabase.GetPara("����˲��˺���ҩ����ͬ", glngSys, 1341, 0)) = 1)
        
        If mcondition.str��ҩ�� = "|��ǰ����Ա|" Then
            mstrDosUser = gstrUserName
        Else
            mstrDosUser = mcondition.str��ҩ��
        End If
        
        If mcondition.str�˲��� = "|��ǰ����Ա|" Then
            mstr�˲��� = gstrUserName
        Else
            mstr�˲��� = mcondition.str�˲���
        End If
    
        If .lngҩ��ID <> Val(zldatabase.GetPara("��ҩҩ��", glngSys, 1341)) Then
            .lngҩ��ID = Val(zldatabase.GetPara("��ҩҩ��", glngSys, 1341))
            .bln�Ƿ���Ҫ��ҩ���� = RecipeSendWork_DispensingMedi(.lngҩ��ID, bln�Ƿ���ҩȷ��)
            Call Load��ҩ��(.lngҩ��ID)
            Call GetDrugDigit(.lngҩ��ID, "ҩƷ������ҩ", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
            
            '�����ʾ����ȡ���þ���
            mintMoneyDigit = Val(zldatabase.GetPara("���ý���λ��", glngSys, 0))
            
            .int����� = MediWork_GetCheckStockRule(.lngҩ��ID)
        End If
        
        .bln��ʾԭ���� = Is��ҩ�ⷿ(.lngҩ��ID)
        
        If .intListType = mListType.��ҩ Then
            Cbo��ҩ��.Enabled = False
            cbo�˲���.Enabled = False
        Else
            If zlStr.IsHavePrivs(mstrPrivs, "��ҩ") = True Then
                If .bln�Զ���ҩ = False Then
                    Cbo��ҩ��.Enabled = True
                Else
                    Cbo��ҩ��.Enabled = False
                End If
            Else
                Cbo��ҩ��.Enabled = False
            End If
            cbo�˲���.Enabled = True
        End If
        
        Call Load�˲���(.lngҩ��ID)
        
        cmdSendByNoTake.Visible = (.blnȡҩȷ�� And .intListType = mListType.����ҩ)
        
        Call Load��ҩ��(.lngҩ��ID)
        Call Load�˲���(.lngҩ��ID)
    
    End With
    
    
    
    mstrUserRecipeColor = zldatabase.GetPara("������ɫ", glngSys, 1341)
    If mstrUserRecipeColor = "" Then mstrUserRecipeColor = GetDefaultRecipeColor
End Sub

Public Sub ShowList(ByVal intType As Integer, ByVal bln��ʾ���̵��� As Boolean)
    Dim i As Integer
    
    With mcondition
        If .intListType <> intType Then
            SaveListColState
            .intListType = intType
            
            If .intListType = mListType.����ҩ Or .intListType = mListType.��ʱδ�� Then
                cbo����ҽ��.Enabled = True
            Else
                cbo����ҽ��.Enabled = False
            End If
            
            If .intListType <> mListType.��ҩ Then
                For i = 0 To Cbo��ҩ��.ListCount - 1
                    If mstrDosUser = Cbo��ҩ��.List(i) Then
                        Cbo��ҩ��.ListIndex = i
                        Exit For
                    End If
                Next
                
                For i = 0 To cbo�˲���.ListCount - 1
                    If mstr�˲��� = cbo�˲���.List(i) Then
                        cbo�˲���.ListIndex = i
                        Exit For
                    End If
                Next
            End If
            
            If .intListType = mListType.��ҩ Then
                Cbo��ҩ��.Enabled = False
                cbo�˲���.Enabled = False
            Else
                If zlStr.IsHavePrivs(mstrPrivs, "��ҩ") = True Then
                    If .bln�Զ���ҩ = False Then
                        Cbo��ҩ��.Enabled = True
                    Else
                        Cbo��ҩ��.Enabled = False
                    End If
                Else
                    Cbo��ҩ��.Enabled = False
                End If
                cbo�˲���.Enabled = True
            End If
        End If
        .bln��ʾ���̵��� = bln��ʾ���̵���
    End With
   
    Call SetComandBars(intType)
    
    InitList mcondition.intListType
    
    Call InitColSelList(intType)
    
    Select Case mcondition.intListType
        Case mListType.��ҩȷ��
            Me.cbo����ҽ��.Enabled = False
            Me.Cbo��ҩ��.Enabled = False
            Me.cbo�˲���.Enabled = False
            CmdSend.Caption = "��ҩȷ��(&O)"
            picRecipeColor.Visible = True
        Case mListType.����ҩ
            CmdSend.Caption = "��ҩ(&V)"
            picRecipeColor.Visible = True
        Case mListType.����ҩ
            CmdSend.Caption = "ȡ����ҩ(&C)"
            picRecipeColor.Visible = True
        Case mListType.����ҩ, mListType.��ʱδ��
            CmdSend.Caption = "��ҩ(&S)"
            picRecipeColor.Visible = True
        Case mListType.��ҩ
            CmdSend.Caption = "��ҩ(&R)"
            picRecipeColor.Visible = True
    End Select
    
    cmdSendByNoTake.Visible = (mcondition.blnȡҩȷ�� And mcondition.intListType = mListType.����ҩ)
    Chkȫ��.Enabled = (mcondition.intListType = mListType.��ҩ)
    
    SetCmdSendPrivs mcondition.bln�������
    
    DoEvents
    Call Form_Resize
End Sub

Private Sub cbo�˲���_Click()
    Dim i As Integer
    
    If mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ�� Then
        mstr�˲��� = Me.cbo�˲���.Text
    End If
End Sub

Private Sub cbo�˲���_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstemp As ADODB.Recordset
    Dim i As Integer
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo��ҩ��.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    vRect = zlControl.GetControlRect(cbo�˲���.hWnd) '��ȡλ��
    gstrSQL = "Select ID, ���, ����, ����" & _
               " From ��Ա��" & _
               " Where ID In (Select ��Աid From ������Ա Where ����id = [1]) And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) And" & _
                     " (��� Like [2] Or ���� Like [2] Or ���� Like [2])"
    
    Set rstemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "��ѯ��Ա", False, "", "", False, False, _
                        True, vRect.Left, vRect.Top, 500, blnCancel, False, True, mcondition.lngҩ��ID, IIf(gstrMatchMethod = 0, "%", "") & UCase(cbo�˲���.Text) & "%")
        
    If rstemp Is Nothing Then
        Exit Sub
    End If
    For i = 0 To cbo�˲���.ListCount
        If Mid(cbo�˲���.List(i), InStr(1, cbo�˲���.List(i), "-") + 1) = rstemp!���� Then
            cbo�˲���.ListIndex = i
        End If
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo��ҩ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstemp As ADODB.Recordset
    Dim i As Integer
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo��ҩ��.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    vRect = zlControl.GetControlRect(Cbo��ҩ��.hWnd) '��ȡλ��
    gstrSQL = "Select id, ���,����,����" & _
              "  From ��Ա��" & _
               " Where ID In (Select Distinct ��Աid" & _
                            " From ��Ա����˵�� " & _
                            " Where ��Ա���� = 'ҩ����ҩ��' And ��Աid In (Select ��Աid From ������Ա Where ����id = [1])) And" & _
                     " (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) and  (��� like [2] or ���� like [2] or ���� like [2])"
    
    Set rstemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "��ѯ��Ա", False, "", "", False, False, _
                        True, vRect.Left, vRect.Top, 500, blnCancel, False, True, mcondition.lngҩ��ID, IIf(gstrMatchMethod = 0, "%", "") & UCase(Cbo��ҩ��.Text) & "%")
        
    If rstemp Is Nothing Then
        Exit Sub
    End If
    For i = 0 To Cbo��ҩ��.ListCount
        If Mid(Cbo��ҩ��.List(i), InStr(1, Cbo��ҩ��.List(i), "-") + 1) = rstemp!���� Then
            Cbo��ҩ��.ListIndex = i
        End If
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub cbo��ҩ��_Click()
    If mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ�� Then
        mstrDosUser = Me.Cbo��ҩ��.Text
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Dim Int���� As Integer
    Dim strNo As String
    Dim strsql As String
    Dim rsTmp As Recordset
    Dim lngPatiID As Long
    Dim str�Һŵ� As String
    Dim lng��ҳID As Long
    Dim lngCurrAdviceID As Long
    
    If vsfList.Row = 0 Then Exit Sub
    If vsfList.Row = vsfList.rows - 1 Then Exit Sub
    
    Int���� = vsfList.TextMatrix(vsfList.Row, mIntCol����)
    strNo = vsfList.TextMatrix(vsfList.Row, mIntColNO)
    lngCurrAdviceID = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ҽ��id")))
    
    
    '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�������˳�
            strsql = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.����id,0 ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,������ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] "
            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, strNo, Int����)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!����ID
            str�Һŵ� = NVL(rsTmp!�Һŵ�)
            lng��ҳID = rsTmp!��ҳid
            
    Select Case Control.Id
        Case conMenu_Tool_ShowPlug
            '���ܣ��Բ��˹���ʷ/����״̬���й���
            Call gobjPass.zlPassCmdAlleyManage_YF(mlngMode, lngPatiID, lng��ҳID, str�Һŵ�)
        '�����˵���PASS����
        Case mconMenu_PASS * 10# To mconMenu_PASS * 10# + 99
            Call gobjPass.zlPassCommandBarExe_YF(mlngMode, Control.Id - (mconMenu_PASS * 10#), lngPatiID, lng��ҳID, str�Һŵ�, lngCurrAdviceID)
    End Select
End Sub


Private Function AdviceCheckWarn(ByVal Int���� As Integer, ByVal strNo As String, ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'���ܣ�����Passϵͳ��ع���
'������lngCmd=
'        0-�������PASS�˵�״̬
'        21-����״̬/����ʷ����(ֻ��)
'      lngRow=��ǰҩƷҽ�����кţ�lngCmd=0ʱ��Ҫ
'���أ����PASS�˵�ʱ������>=0��ʾ���Ե����˵�,��������-1
'˵������ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New ADODB.Recordset
    Dim strҩƷ As String, str�÷� As String, lngҩƷid As Long, str������λ As String, strƵ�� As String
    Dim strsql As String, i As Long, k As Long
    Dim lngPatiID As Long
    Dim lngPassPati As Long
    Dim lng��ҳID As Long
    Dim str�Һŵ� As String
    Dim lngCount As Long
    Dim blnDo As Boolean
    

    AdviceCheckWarn = -1

    On Error GoTo errH
    Screen.MousePointer = 11

    If strNo = "" Then Exit Function

    '����PASS����״̬
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�������˳�
    strsql = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] " & _
        " Union All " & _
        " Select distinct B.����id,0 ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,������ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, strNo, Int����)

    If rsTmp.RecordCount = 0 Then
        rsTmp.Close
        Exit Function
    End If

    lngPatiID = rsTmp!����ID
    str�Һŵ� = zlStr.NVL(rsTmp!�Һŵ�)
    lng��ҳID = rsTmp!��ҳid
    

    '���벡�˾�����Ϣ(PASS��Ҫ�Ļ�������,ͬһ���˿ɲ��ظ�����)
    '-------------------------------------------------------------
    If lngPatiID <> lngPassPati Then
        If str�Һŵ� <> "" Then               '���ﲡ��
            strsql = "Select ����ID,Count(Distinct Trunc(�Ǽ�ʱ��)) as ������� From ���˹Һż�¼ Where ��¼����=1 And ��¼״̬=1 And ����ID=[1] Group by ����ID"
            strsql = "Select D.�������,A.����,A.�Ա�,A.��������," & _
                " C.���� as ������,C.���� as ������,E.��� as ҽ����,E.���� as ҽ����" & _
                " From ������Ϣ A,���˹Һż�¼ B,���ű� C,(" & strsql & ") D,��Ա�� E" & _
                " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And A.����ID=D.����ID" & _
                " And B.ִ����=E.����(+) And A.����ID=[1] And B.NO=[2]"
            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lngPatiID, str�Һŵ�)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

            Call PassSetPatientInfo(lngPatiID, rsTmp!�������, rsTmp!����, zlStr.NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), zlStr.NVL(rsTmp!ҽ����) & "/" & zlStr.NVL(rsTmp!ҽ����), ""), "")
        Else                                    'סԺ����
            strsql = _
                " Select A.����,A.�Ա�,A.��������,B.��Ժ����,B.��Ժ����," & _
                " C.���� as ������,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
                " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
                " Where A.����ID=B.����ID And A.��ҳid=b.��ҳid And B.��Ժ����ID=C.ID" & _
                " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[2]"
            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lngPatiID, lng��ҳID)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

            Call PassSetPatientInfo(lngPatiID, lng��ҳID, rsTmp!����, zlStr.NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), zlStr.NVL(rsTmp!ҽ����) & "/" & zlStr.NVL(rsTmp!ҽ����), ""), _
                IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy-MM-dd")))
        End If
        lngPassPati = lngPatiID
    End If
    
    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        'ȡҩƷ����
        strҩƷ = vsfList.TextMatrix(lngRow, mIntColҩƷ����)
        lngҩƷid = vsfList.TextMatrix(lngRow, mintcolҩƷid)
        str������λ = Mid(vsfList.TextMatrix(lngRow, mIntCol����), InStr(vsfList.TextMatrix(lngRow, mIntCol����), "(") + 1)
        If InStr(str������λ, ")") > 0 Then str������λ = Replace(str������λ, ")", "")
        'ȡҩƷ��ҩ;��
        str�÷� = vsfList.TextMatrix(lngRow, mIntCol�÷�)

        If InStr(strҩƷ, " ") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, " ") - 1)
        If InStr(strҩƷ, "]") > 0 Then strҩƷ = Mid(strҩƷ, InStr(strҩƷ, "]") + 1, Len(strҩƷ) - InStr(strҩƷ, "]"))
        '�����ѯҩƷ��Ϣ
        Call PassSetQueryDrug(lngҩƷid, strҩƷ, str������λ, str�÷�)

        '���ò˵�����״̬
        Call SetPassMenuState

        AdviceCheckWarn = 1 '��ʾ���Ե����˵�

        Screen.MousePointer = 0: Exit Function
    ElseIf lngCmd = 6 Then
        Call PassSetWarnDrug(Val(vsfList.TextMatrix(lngRow, mIntCol���id))) '��ҩ����(�Ѿ����ҽ��Ψһ��)
    Else
        With Me.vsfList
            '��ҩ��˻���ҩ�о�
            lngCount = 0
            strҩƷ = "": str�÷� = "": strƵ�� = ""
            i = 1
            If .TextMatrix(i, mIntCol����ҽ��) <> "" Then
                strsql = "select ��� from ��Ա�� where ����=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strsql, "", .TextMatrix(i, mIntCol����ҽ��))
            End If
            For i = .FixedRows To .rows - 1
                blnDo = Val(.TextMatrix(i, mIntColҽ��id)) <> 0 And Val(.TextMatrix(i, mintcolҩƷid)) <> 0
                If blnDo Then
                    'ȡҩƷ����
                    strҩƷ = .TextMatrix(i, mIntColҩƷ����)
                    If InStr(strҩƷ, "]") > 0 Then strҩƷ = Mid(strҩƷ, InStr(strҩƷ, "]") + 1, Len(strҩƷ) - InStr(strҩƷ, "]"))
                    
                    'ȡҩƷ��ҩ;��
                    str�÷� = .TextMatrix(i, mIntCol�÷�)
                    
                    'ȡ��ҩƵ��(��/��),��Ϊ������������
                    If .TextMatrix(i, mIntCol�����λ) = "��" Then
                        strƵ�� = .TextMatrix(i, mIntColƵ�ʴ���) & "/" & .TextMatrix(i, mIntColƵ�ʼ��)
                    ElseIf .TextMatrix(i, mIntCol�����λ) = "��" Then
                        strƵ�� = .TextMatrix(i, mIntColƵ�ʴ���) & "/7"
                    ElseIf .TextMatrix(i, mIntCol�����λ) = "Сʱ" Then
                        If Val(.TextMatrix(i, mIntColƵ�ʼ��)) <= 24 Then
                            strƵ�� = Format(24 / Val(.TextMatrix(i, mIntColƵ�ʼ��)) * Val(.TextMatrix(i, mIntColƵ�ʴ���)), "0") & "/1"
                        Else
                            strƵ�� = Val(.TextMatrix(i, mIntColƵ�ʴ���)) & "/" & Format(Val(.TextMatrix(i, mIntColƵ�ʼ��)) / 24, "0")
                        End If
                    ElseIf .TextMatrix(i, mIntCol�����λ) = "����" Then
                        strƵ�� = Format((24 * 60) / Val(.TextMatrix(i, mIntColƵ�ʼ��)) * Val(.TextMatrix(i, mIntColƵ�ʴ���)), "0") & "/1"
                    End If
                    
'                    MsgBox "ҽ��id��" & .TextMatrix(i, mIntColҽ��id) & "��ҩƷid:" & .TextMatrix(i, mIntColҩƷID) & ";ҩƷ��" & strҩƷ & "��������" & Mid(.TextMatrix(i, mIntCol����), 1, InStr(1, .TextMatrix(i, mIntCol����), "(") - 1) & ";" & _
'                            "��λ��" & Mid(.TextMatrix(i, mIntCol����), InStr(1, .TextMatrix(i, mIntCol����), "(") + 1, InStr(1, .TextMatrix(i, mIntCol����), ")") - InStr(1, .TextMatrix(i, mIntCol����), "(") - 1) & ";��ҩƵ��" & strƵ�� & ";" & _
'                            "��ʼʱ�䣺" & Format(.TextMatrix(i, mIntCol��ʼʱ��), "yyyy-MM-dd") & ";����ʱ�䣺" & Format(.TextMatrix(i, mIntCol����ʱ��), "yyyy-MM-dd") & ";�÷���" & str�÷� & _
'                            "���id��" & .TextMatrix(i, mIntCol���id) & ";ҽ����־��" & .TextMatrix(i, mIntColҽ����־) & ";ҽ�������" & rsTmp!��� & "\" & .TextMatrix(i, mIntCol����ҽ��)
                    '����ҽ����Ϣ
                    Call PassSetRecipeInfo(.TextMatrix(i, mIntColҽ��id), .TextMatrix(i, mintcolҩƷid), strҩƷ, _
                        Mid(.TextMatrix(i, mIntCol����), 1, InStr(1, .TextMatrix(i, mIntCol����), "(") - 1), Mid(.TextMatrix(i, mIntCol����), InStr(1, .TextMatrix(i, mIntCol����), "(") + 1, InStr(1, .TextMatrix(i, mIntCol����), ")") - InStr(1, .TextMatrix(i, mIntCol����), "(") - 1), strƵ��, _
                        Format(.TextMatrix(i, mIntCol��ʼʱ��), "yyyy-MM-dd"), Format(.TextMatrix(i, mIntCol����ʱ��), "yyyy-MM-dd"), str�÷�, _
                        .TextMatrix(i, mIntCol���id), .TextMatrix(i, mIntColҽ����־), rsTmp!��� & "\" & .TextMatrix(i, mIntCol����ҽ��))
                    lngCount = lngCount + 1
                End If
            Next
            
            '�޿�����ҩƷ
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And lngCount = 0 Then
                Screen.MousePointer = 0: Exit Function
            End If
        End With
    End If

    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    Screen.MousePointer = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Private Sub SetPassMenuState()
    '���ܣ�����Pass�˵�����״̬
    'Pass
    Dim objPopup As CommandBarControl

    ''''һ���˵�
    'ҩ���ٴ���Ϣ�ο�
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CPRRes") = 1

    'ҩƷ˵����
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 1, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("Directions") = 1

    '�й�ҩ��
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 2, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("Chp") = 1

    '������ҩ����
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 3, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CPERes") = 1

    '����ֵ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 4, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CheckRes") = 1

    'ר����Ϣ
'    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 5, , True)
'    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("") = 1

    'ҽҩ��Ϣ����
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 6, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MEDInfo") = 1

    'ҩƷ�����Ϣ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 7, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MATCH-DRUG") = 1

    '��ҩ;�������Ϣ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 8, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MATCH-ROUTE") = 1

    'ҽԺҩƷ��Ϣ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 9, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("HisDrugInfo") = 1
    
    
    ''''ר����Ϣ�����˵�
    'ҩ��-ҩ���໥����
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DDIM") = 1
    
    'ҩ��-ʳ���໥ʹ��
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 1, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DFIM") = 1
    
    '����ע�����������
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 2, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MatchRes") = 1
    
    '����ע�����������
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 3, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("TriessRes") = 1
    
    '����֢
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 4, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DDCM") = 1
    
    '������
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 5, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("SIDE") = 1
    
    '��������ҩ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 6, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("GERI") = 1
    
    '��ͯ��ҩ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 7, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("PEDI") = 1
    
    '��������ҩ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 8, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("PREG") = 1
    
    '��������ҩ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 9, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("LACT") = 1
End Sub
Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    If cbsMain.count > 1 Then
        Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

        picRecipt.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    End If
End Sub


Sub InitList(ByVal intType As Integer)
    Dim i As Integer
    Dim n As Integer
    Dim str������ As String
    Dim arr������
    Dim bln��������Ч As Boolean
    
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntCol��ǰ�� = 0
    mIntCol˳��� = 1
    mIntCol����� = 2
    mIntColҩƷ���� = 3
    mintColƤ�Խ�� = 4
    mIntCol������ = 5
    mIntColӢ���� = 6
    mIntCol�䷽���� = 7
    mintcol��� = 8
    mintcol��� = 9
    mintcol������ = 10
    mintcolԭ���� = 11
    mintcol���� = 12
    mintcolЧ�� = 13
    mintcol��λ = 14
    mIntCol���� = 15
    mIntCol���� = 16
    mintcol���� = 17
    mIntCol��� = 18
    mIntColʵ�ս�� = 19
    mIntCol���� = 20
    mIntCol���� = 21
    mIntCol�÷� = 22
    mIntColƵ�� = 23
    mIntCol����˵�� = 24
    mIntCol��ҩĿ�� = 25
    mIntColҽ������ = 26
    mIntCol�ѱ� = 27
    mIntCol����� = 28
    mIntCol��λ = 29
    mIntCol������ = 30
    mIntCol׼���� = 31
    mIntCol׼������ = 32
    mIntCol׼����С = 33
    mIntCol��ҩ�� = 34
    mIntCol��ҩ���� = 35
    mIntCol��λ�� = 36
    mIntCol��ҩ��С = 37
    mIntCol��λС = 38
    mIntCol��ע = 39
    '--------------������Ϊ���ɼ�--------------
    mIntCol���� = 40
    mIntCol������ = 41
    mIntCol��Ч�� = 42
    mIntCol�²��� = 43
    mIntColҽ��id = 44
    mIntColʵ������ = 45
    mIntCol��װ = 46
    mIntCol���� = 47
    mIntColNO = 48
    mIntCol�����־ = 49
    mIntCol��¼���� = 50
    mIntCol�ⷿID = 51
    mIntCol��ҩ���� = 52
    mIntCol���id = 53
    mIntCol����ҽ�� = 54
    mIntColƵ�ʼ�� = 55
    mIntCol�����λ = 56
    mIntColҽ����־ = 57
    mIntCol��ʼʱ�� = 58
    mIntCol����ʱ�� = 59
    mIntColƵ�ʴ��� = 60
    mIntCol���� = 61
    mIntCol����� = 62
    mIntColסԺ�� = 63
    mIntColId = 64
    mintcolҩƷid = 65
    mintcol���� = 66
    mIntCol����ҩƷ˵�� = 67
    
    
    '�ָ��û��Զ�����˳��
    str������ = LoadListColState
    If str������ <> "" Then
        arr������ = Split(str������, "|")
        If UBound(arr������) + 1 <> mconIntCol���� Then
            str������ = ""
        Else
            For n = 0 To UBound(arr������)
                If Split(arr������(n), ",")(0) = "" Then
                    bln��������Ч = True
                    Exit For
                End If
            Next
            
            If bln��������Ч = True Then
                str������ = ""
            Else
                For n = 0 To UBound(arr������)
                    SetColumnValue Split(arr������(n), ",")(0), n
                Next
            End If
        End If
    End If
    
    '��ʼ��δ��ҩ�嵥
    With vsfList
        .Redraw = flexRDNone
        
        .rows = 1
        .rows = 2
        .Cols = mconIntCol����
        
        .Cell(flexcpPicture, 1, mIntCol��ǰ��, 1, mIntCol��ǰ��) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol��ǰ��, .rows - 1, mIntCol��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList, mIntCol��ǰ��, "", 250, flexAlignCenterCenter, "��ǰ��"
        VsfGridColFormat vsfList, mIntCol˳���, "���", 450, flexAlignRightCenter, "˳���"
        
        If IsInString(gstrprivs, "������ҩ���", ";") And Not gobjPass Is Nothing Then
            VsfGridColFormat vsfList, mIntCol�����, "��", 280, flexAlignCenterCenter, "�����"
        Else
            VsfGridColFormat vsfList, mIntCol�����, "��", 0, flexAlignCenterCenter, "�����"
        End If
        
        VsfGridColFormat vsfList, mIntColҩƷ����, "ҩƷ����", 2500, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfList, mintColƤ�Խ��, "", 400, flexAlignLeftCenter, "Ƥ�Խ��"
        VsfGridColFormat vsfList, mIntCol������, "������", 2000, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfList, mIntColӢ����, "Ӣ����", 2000, flexAlignLeftCenter, "Ӣ����"
        VsfGridColFormat vsfList, mIntCol�䷽����, "�䷽����", 1800, flexAlignLeftCenter, "�䷽����"
        VsfGridColFormat vsfList, mintcol���, "���", 0, flexAlignCenterCenter, "���"
        VsfGridColFormat vsfList, mintcol���, "���", 1500, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfList, mintcol����, "����", 1500, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList, mintcolЧ��, "Ч��", 1500, flexAlignLeftCenter, "Ч��"
        VsfGridColFormat vsfList, mIntColId, "Id", 0, flexAlignCenterCenter, "Id"
        VsfGridColFormat vsfList, mintcolҩƷid, "ҩƷID", 0, flexAlignCenterCenter, "ҩƷID"
        
        VsfGridColFormat vsfList, mintcol����, "����", 0, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList, mintcol��λ, "��λ", IIf(intType = mListType.��ҩ And mcondition.bln��ʾ��С��λ, 0, 500), flexAlignCenterCenter, "��λ"
        VsfGridColFormat vsfList, mIntCol����, "����", 1000, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList, mIntCol����, "����", IIf(mcondition.bln��ʾ����, 800, 0), flexAlignRightCenter, "����"
        VsfGridColFormat vsfList, mintcol����, "����", 1200, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList, mIntCol���, "Ӧ�ս��", IIf(mcondition.int�����ʾ = 1, 0, 1000), flexAlignRightCenter, "Ӧ�ս��"
        VsfGridColFormat vsfList, mIntColʵ�ս��, "ʵ�ս��", IIf(mcondition.int�����ʾ = 0, 0, 1000), flexAlignRightCenter, "ʵ�ս��"
        VsfGridColFormat vsfList, mIntCol����, "����", 1200, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList, mIntCol����, "����", 1200, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList, mIntCol�÷�, "�÷�", 1500, flexAlignLeftCenter, "�÷�"
        VsfGridColFormat vsfList, mIntColƵ��, "Ƶ��", 1500, flexAlignLeftCenter, "Ƶ��"
        VsfGridColFormat vsfList, mIntCol����˵��, "����˵��", 1500, flexAlignLeftCenter, "����˵��"
        VsfGridColFormat vsfList, mIntCol��ҩĿ��, "��ҩĿ��", 0, flexAlignLeftCenter, "��ҩĿ��"
        
       
        VsfGridColFormat vsfList, mIntColҽ������, "ҽ������", IIf(intType = mListType.��ҩ, 0, 1500), flexAlignLeftCenter, "ҽ������"
        VsfGridColFormat vsfList, mIntCol�ѱ�, "�ѱ�", 1000, flexAlignLeftCenter, "�ѱ�"
        VsfGridColFormat vsfList, mIntCol�����, "�����", IIf(intType = mListType.��ҩ, 0, 1200), flexAlignRightCenter, "�����"
        VsfGridColFormat vsfList, mIntCol��λ, "�ⷿ��λ", IIf(intType = mListType.��ҩ, 0, 1200), flexAlignLeftCenter, "�ⷿ��λ"
        VsfGridColFormat vsfList, mIntCol������, "������", IIf(intType = mListType.��ҩ, 1200, 0), flexAlignRightCenter, "������"
        VsfGridColFormat vsfList, mIntCol׼����, "׼����", IIf(intType = mListType.��ҩ, 1200, 0), flexAlignRightCenter, "׼����"
        VsfGridColFormat vsfList, mIntCol׼������, "׼������", 0, flexAlignCenterCenter, "׼������"
        VsfGridColFormat vsfList, mIntCol׼����С, "׼����С", 0, flexAlignCenterCenter, "׼����С"
        VsfGridColFormat vsfList, mIntCol��ҩ��, "��ҩ��", IIf(intType = mListType.��ҩ And mcondition.bln��ʾ��С��λ = False, 1200, 0), flexAlignRightCenter, "��ҩ��"
        VsfGridColFormat vsfList, mIntCol��ҩ����, "��ҩ��(���װ)", IIf(intType = mListType.��ҩ And mcondition.bln��ʾ��С��λ = True, 1500, 0), flexAlignRightCenter, "��ҩ��(���װ)"
        
        VsfGridColFormat vsfList, mIntCol��λ��, "��λ(��)", IIf(intType = mListType.��ҩ And mcondition.bln��ʾ��С��λ = True, 500, 0), flexAlignCenterCenter, "��λ(��)"
        VsfGridColFormat vsfList, mIntCol��ҩ��С, "��ҩ��(С��װ)", IIf(intType = mListType.��ҩ And mcondition.bln��ʾ��С��λ = True, 1500, 0), flexAlignRightCenter, "��ҩ��(С��װ)"
        VsfGridColFormat vsfList, mIntCol��λС, "��λ(С)", IIf(intType = mListType.��ҩ And mcondition.bln��ʾ��С��λ = True, 500, 0), flexAlignCenterCenter, "��λ(С)"
        VsfGridColFormat vsfList, mIntCol����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList, mIntCol������, "������", 0, flexAlignCenterCenter, "������"
        VsfGridColFormat vsfList, mIntCol��Ч��, "��Ч��", 0, flexAlignCenterCenter, "��Ч��"
        VsfGridColFormat vsfList, mIntCol�²���, "�²���", 0, flexAlignCenterCenter, "�²���"
        VsfGridColFormat vsfList, mIntCol��ע, "��ע", 1200, flexAlignLeftCenter, "��ע"
        VsfGridColFormat vsfList, mintcol������, "������", 1200, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfList, mintcolԭ����, "ԭ����", 1200, flexAlignLeftCenter, "ԭ����"
        VsfGridColFormat vsfList, mIntColҽ��id, "ҽ��id", 0, flexAlignCenterCenter, "ҽ��id"
        VsfGridColFormat vsfList, mIntColʵ������, "ʵ������", 0, flexAlignCenterCenter, "ʵ������"
        
        VsfGridColFormat vsfList, mIntCol��װ, "��װ", 0, flexAlignLeftCenter, "��װ"
        VsfGridColFormat vsfList, mIntCol����, "����", 0, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList, mIntColNO, "No", 0, flexAlignLeftCenter, "NO"
        
        VsfGridColFormat vsfList, mIntCol�����־, "�����־", 0, flexAlignLeftCenter, "�����־"
        VsfGridColFormat vsfList, mIntCol��¼����, "��¼����", 0, flexAlignLeftCenter, "��¼����"
        VsfGridColFormat vsfList, mIntCol�ⷿID, "�ⷿID", 0, flexAlignLeftCenter, "�ⷿID"
        VsfGridColFormat vsfList, mIntCol��ҩ����, "��ҩ����", 0, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfList, mIntCol���id, "���id", 0, flexAlignLeftCenter, "���id"
        VsfGridColFormat vsfList, mIntCol����ҽ��, "����ҽ��", 0, flexAlignLeftCenter, "����ҽ��"
        VsfGridColFormat vsfList, mIntColƵ�ʼ��, "Ƶ�ʼ��", 0, flexAlignLeftCenter, "Ƶ�ʼ��"
        VsfGridColFormat vsfList, mIntCol�����λ, "�����λ", 0, flexAlignLeftCenter, "�����λ"
        VsfGridColFormat vsfList, mIntColҽ����־, "ҽ����־", 0, flexAlignLeftCenter, "ҽ����־"
        VsfGridColFormat vsfList, mIntCol��ʼʱ��, "��ʼʱ��", 0, flexAlignLeftCenter, "��ʼʱ��"
        VsfGridColFormat vsfList, mIntCol����ʱ��, "����ʱ��", 0, flexAlignLeftCenter, "����ʱ��"
        VsfGridColFormat vsfList, mIntColƵ�ʴ���, "Ƶ�ʴ���", 0, flexAlignLeftCenter, "Ƶ�ʴ���"
        VsfGridColFormat vsfList, mIntCol����, "����", 0, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList, mIntCol�����, "�����", 0, flexAlignLeftCenter, "�����"
        VsfGridColFormat vsfList, mIntColסԺ��, "סԺ��", 0, flexAlignLeftCenter, "סԺ��"
        VsfGridColFormat vsfList, mIntCol����ҩƷ˵��, "����ҩƷ˵��", 0, flexAlignLeftCenter, "����ҩƷ˵��"
        
        mstrUnallowShow = "��ǰ��;���;Ƥ�Խ��;Id;ҩƷID;����;��ҩĿ��;����;������;��Ч��;�²���;ҽ��id;ʵ������;��װ;����;NO;�����־;��¼����;�ⷿID;��ҩ����;���id;����ҽ��;Ƶ�ʼ��;�����λ;ҽ����־;��ʼʱ��;����ʱ��;Ƶ�ʴ���;����;�����;סԺ��;����ҩƷ˵��"
        If mcondition.int�����ʾ = 0 Then mstrUnallowShow = mstrUnallowShow & ";ʵ�ս��"
        If mcondition.int�����ʾ = 1 Then mstrUnallowShow = mstrUnallowShow & ";Ӧ�ս��"
        
        If intType <> mListType.��ҩ Then
            mstrUnallowSetColHide = "ҩƷ����;����"
            mstrUnallowShow = mstrUnallowShow & ";" & "��ҩ��(���װ);��ҩ��(С��װ);������;׼����;׼������;׼����С;��ҩ��;��ҩ��(���װ);��λ(��);��ҩ��(С��װ);��λ(С)"
        Else
            mstrUnallowShow = mstrUnallowShow & ";" & "ҽ������;�����;�ⷿ��λ"
            If mcondition.bln��ʾ��С��λ Then
                mstrUnallowSetColHide = "ҩƷ����;����;������;׼����;��ҩ��(���װ);��ҩ��(С��װ)"
                mstrUnallowShow = mstrUnallowShow & ";" & "��λ;��ҩ��;׼������;׼����С"
            Else
                mstrUnallowSetColHide = "ҩƷ����;����;������;׼����;��ҩ��"
                mstrUnallowShow = mstrUnallowShow & ";" & "׼������;׼����С;��ҩ��(���װ);��ҩ��(С��װ);��λ(��);��λ(С)"
            End If
        End If
        
        If mcondition.intShowPass <> 0 Or Not IsInString(gstrprivs, "������ҩ���", ";") Then mstrUnallowShow = mstrUnallowShow & ";" & "�����"
        If mcondition.bln��ʾ���� = False Then mstrUnallowShow = mstrUnallowShow & ";" & "����"
        
        '�ָ��Զ����п���������������ʾ���У�
        If str������ <> "" Then
            arr������ = Split(str������, "|")
            For n = 0 To UBound(arr������)
                If IsInString(mstrUnallowShow, Split(arr������(n), ",")(0), ";") = False Then
                    For i = 0 To vsfList.Cols - 1
                        If Split(arr������(n), ",")(0) = vsfList.ColKey(i) Then
                            vsfList.ColWidth(i) = Val(Split(arr������(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If intType = mListType.��ҩ And mcondition.bln��ʾ��С��λ = True Then
            VsfGridColFormat vsfList, mIntCol��ҩ��, "��ҩ��", 0, flexAlignRightCenter, "��ҩ��"
            VsfGridColFormat vsfList, mIntCol��ҩ����, "��ҩ��(���װ)", 1500, flexAlignRightCenter, "��ҩ��(���װ)"
            
            VsfGridColFormat vsfList, mIntCol��λ��, "��λ(��)", 500, flexAlignCenterCenter, "��λ(��)"
            VsfGridColFormat vsfList, mIntCol��ҩ��С, "��ҩ��(С��װ)", 1500, flexAlignRightCenter, "��ҩ��(С��װ)"
            VsfGridColFormat vsfList, mIntCol��λС, "��λ(С)", 500, flexAlignCenterCenter, "��λ(С)"
        End If
        
        'ֻ����ҩ��ⷿ����ʾ"ԭ����"��
        If mcondition.bln��ʾԭ���� = False Then VsfGridColFormat vsfList, mintcolԭ����, "ԭ����", 0, flexAlignLeftCenter, "ԭ����"
        
        '������������
        .Select 0, 0, .rows - 1, .Cols - 1
        .CellBorder &H9D9D9D, 1, 1, 1, 1, 1, 1
        
        .Select 0, mIntColҩƷ����, vsfList.rows - 1, mintColƤ�Խ��
        .CellBorder &H9D9D9D, -1, -1, -1, -1, 0, 1
        
        .RowSel = 1
        
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub SaveListColState()
    Dim str������ As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    
    If vsfList.Cols <> mconIntCol���� Then Exit Sub
    
    Select Case mcondition.intListType
        Case mListType.��ҩȷ��
            strType = "��ҩȷ��"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.��ʱδ��
            strType = "��ʱδ��"
        Case mListType.��ҩ
            strType = "��ҩ"
    End Select
    
    With vsfList
        For i = 0 To .Cols - 1
            If vsfList.ColKey(i) = "" Then
                MsgBox "AA"
            End If
            str������ = IIf(str������ = "", "", str������ & "|") & vsfList.ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList), strType, str������)
End Sub

Private Function LoadListColState() As String
    Dim str������ As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    
    Select Case mcondition.intListType
        Case mListType.��ҩȷ��
            strType = "��ҩȷ��"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.��ʱδ��
            strType = "��ʱδ��"
        Case mListType.��ҩ
            strType = "��ҩ"
    End Select
    
    LoadListColState = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList), strType, "")
End Function

Public Function RefreshList(ByVal RecData As ADODB.Recordset, Optional ByVal strWeight As String, Optional ByVal int�ɲ��� As Integer = 0, Optional ByVal int�Ŷ�״̬ As Integer, Optional ByVal int����� As Integer) As Boolean
    Dim dblӦ�ս��, dblʵ�ս�� As Double
    Dim IntLocate As Integer
    Dim str����Ա As String
    Dim dbl������ As Double
    Dim str������λ As String
    Dim lng������ As Long
    Dim dblС���� As Double
    Dim int���� As Integer
    Dim intRow As Integer
    Dim strDiag As String
    Dim strSum As String
    Dim i As Integer
    Dim blnƤ�� As Boolean
    Dim bln��ҩ���� As Boolean
    Dim dateCurrent As Date
    Dim dblʵ������ As Double
    
    dateCurrent = Sys.Currentdate
    
    CmdSend.Enabled = False
    mInt�ɲ��� = int�ɲ���
    
    If Chkȫ��.Enabled = True Then Chkȫ��.Value = 1
    
    mcondition.bln��ʾ���̵��� = (Val(GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "��ʾ��ҩ���̵���", 1)) = 1)
    
    Set mrsDetail = RecData
    
    SaveListColState
    InitList mcondition.intListType
    
    Lbl��ҩ��.Caption = IIf(mcondition.intListType = mListType.��ҩ, IIf(int�ɲ��� <> 3, "��ҩ��", "��ҩ��"), "��ҩ��")
    
    RefreshList = False

    dblӦ�ս�� = 0
    dblʵ�ս�� = 0
    txtNo.Clear
    
    '�������ԭʼ��¼������ʾ�У���������׼������
    vsfList.ColWidth(mIntCol������) = 0
    vsfList.ColWidth(mIntCol׼����) = 0
    vsfList.ColWidth(mIntCol��ҩ��) = 0
    If mcondition.intListType = mListType.��ҩ And int�ɲ��� = 1 Then
        vsfList.ColWidth(mIntCol������) = 1000
        vsfList.ColWidth(mIntCol׼����) = 1000
        vsfList.ColWidth(mIntCol��ҩ��) = IIf(mcondition.bln��ʾ��С��λ = False, 1000, 0)
    End If
    
    '��䵥������
    With mrsDetail
        If .EOF Then
            Call FormClear
        Else
            If !��¼���� = 1 Or (!��¼���� = 2 And (!�����־ = 1 Or !�����־ = 4)) Then
                int���� = 1
            Else
                int���� = 2
            End If
            
            'ȷ������δȡҩ��ť����ʾ״̬
            cmdSendByNoTake.Visible = (mcondition.blnȡҩȷ�� And mcondition.intListType = mListType.����ҩ And int���� = 1)
                
            '����ͷ
            Me.lbl����(1).Caption = IIf(IsNull(!����), "", !����)
            Me.lbl����(1).ForeColor = zldatabase.GetPatiColor(IIf(IsNull(!��������), "", !��������))
            
            If mcondition.intListType <> mListType.��ҩ Then
                If zlStr.NVL(!����ģʽ, 0) = 1 Then
                    Me.picMark1.Visible = True
                Else
                    Me.picMark1.Visible = False
                End If
            Else
                Me.picMark1.Visible = False
            End If
            
            Me.Lbl����(1).Caption = IIf(IsNull(!����), "", !����)
            If !���� = 8 Then Me.Lbl����(1).Caption = ""
            Me.cbo����ҽ��.ListIndex = 0
            If (mcondition.blnУ�鴦�� = False) And zlStr.IsHavePrivs(gstrprivs, "ҽ����ѯ") Then
                str����Ա = IIf(IsNull(!������), "", !������)
            Else
                If mcondition.intListType = mListType.��ҩ And mcondition.blnУ�鴦�� = True Then
                    str����Ա = IIf(IsNull(!������), "", !������)
                Else
                    str����Ա = ""
                End If
            End If
            If str����Ա <> "" Then
                '��λҽ��
                For IntLocate = 1 To cbo����ҽ��.ListCount
                    If Mid(cbo����ҽ��.List(IntLocate), InStr(1, cbo����ҽ��.List(IntLocate), "-") + 1) = str����Ա Then
                        cbo����ҽ��.ListIndex = IntLocate
                        Exit For
                    End If
                Next
            End If
            
            cbo����ҽ��.Enabled = ((mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ��) And mcondition.blnУ�鴦�� = True And cbo����ҽ��.ListIndex = 0)
            
            Me.Lbl����(1).Caption = IIf(IsNull(!����), "", !����)
            Me.Lbl����(1).Caption = IIf(IsNull(!����), "", !����)
            Me.lbl���￨��(1).Caption = IIf(IsNull(!���￨��), "", !���￨��)
            Me.LblTel(1).Caption = IIf(IsNull(!��ϵ�˵绰), "", !��ϵ�˵绰)
            
            If mcondition.intListType = mListType.��ҩ Then
                If IIf(IsNull(!��ҩ��), "", !��ҩ��) <> "" Then
                    Me.Cbo��ҩ�� = IIf(IsNull(!��ҩ��), "", !��ҩ��)
                End If
            Else
            
                If IIf(IsNull(!��ҩ��), "", !��ҩ��) <> "" Then
                    Me.Cbo��ҩ�� = IIf(IsNull(!��ҩ��), "", !��ҩ��)
                End If
            End If
            
            
            If mcondition.intListType = mListType.��ҩ Then
                If IIf(IsNull(!�˲���), "", !�˲���) <> "" Then
                    Me.cbo�˲��� = IIf(IsNull(!�˲���), "", !�˲���)
                End If
            End If
            
            Me.Lbl�շ�Ա(1).Caption = IIf(IsNull(!����Ա����), "", !����Ա����)
            Me.Lbl�Ա�(1).Caption = IIf(IsNull(!�Ա�), "", !�Ա�)
            Me.LblסԺ��(1).Caption = IIf(IsNull(!סԺ��), "", !סԺ��)
            
'            If mcondition.intListType <> mListType.��ҩ Then
                picRecipeColor.BackColor = Val(Split(mstrUserRecipeColor, ";")(Val(!��������)))
                lblRecipeType.Caption = Split(gconstrRecipeType, ";")(Val(!��������))
'            Else
'                picRecipeColor.BackColor = &HFFFFFF
'                lblRecipeType.Caption = "����"
'            End If
            '82922,��ʾ����
            Me.LblWeight(1).Caption = IIf(IsNumeric(strWeight), strWeight & "kg", strWeight)
            
            '�����Ϣ
            txt�������.Text = ""
            txt�������.Tag = ""
            txt�������.Height = 180
            
            Call picRecInfo_Resize
            Call picRecipt_Resize
            
            txtNo.AddItem !NO & "--" & !����
            txtNo.ItemData(txtNo.NewIndex) = !����
            txtNo.Tag = !ҩ��ID & "|" & !ҩ�� & "|" & !�����־ & "|" & !��¼���� & "|" & !�������� & "|" & !״̬
            Lblҩ��.Caption = !ҩ��
            
            If �ж��Ƿ���ҩ����(!ҩ��ID, !����, !NO) Then
                bln��ҩ���� = True
            End If

            mblnAllowClick = False
            txtNo.ListIndex = 0
            mblnAllowClick = True

            '�Ƿ���ʾ������ҩ
            If (mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ��) And mcondition.intShowPass = 1 And IsInString(gstrprivs, "������ҩ���", ";") Then
                Dim cbrControl As CommandBarControl
                
'                Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, conMenu_Tool_ShowPlug, , True)
'                If Not cbrControl Is Nothing Then cbrControl.Enabled = Check�Ƿ����ҽ��(Val(!����), !NO)
            End If
            
            vsfList.rows = 1
            vsfList.Redraw = False
            
            Do While Not .EOF
                '60022�����Ϣ��ʾ
                strDiag = RecipeSendWork_GetDiagnosis(int����, IIf(int���� = 1, Val(!ҽ��id), Val(!����ID)), Val(!��ҳid), IIf(bln��ҩ����, 1, 2))
                '����סԺҽ������վ��ҽ�����͵������շ�,���������ϢΪ�յ����
                If int���� = 1 And !��Ժ = 1 And strDiag = "" Then
                    int���� = 2
                    strDiag = RecipeSendWork_GetDiagnosis(int����, IIf(int���� = 1, Val(!ҽ��id), Val(!����ID)), Val(!��ҳid), IIf(bln��ҩ����, 1, 2))
                End If
                
                If strDiag <> "" Then
                    strDiag = strDiag & "|"
                    For i = 0 To UBound(Split(strDiag, "|"))
                        If Split(strDiag, "|")(i) <> "" Then
                            If InStr(1, txt�������.Text & " ��", "��" & Split(strDiag, "|")(i) & " ��") < 1 Then
                                txt�������.Text = IIf(txt�������.Text = "", " ��", txt�������.Text & " ��") & Split(strDiag, "|")(i)
                                txt�������.Tag = IIf(txt�������.Tag = "", "�� ", txt�������.Tag & vbCrLf & "�� ") & Split(strDiag, "|")(i)
                            End If
                        End If
                    Next
                End If
            
                intRow = intRow + 1
                vsfList.rows = intRow + 1

                vsfList.TextMatrix(intRow, mIntCol˳���) = intRow
                If Val(!��ΣҩƷ) <> 0 Then
                    vsfList.Cell(flexcpPicture, intRow, mIntColҩƷ����) = Me.ImgList.ListImages(40).Picture
                    vsfList.Cell(flexcpPictureAlignment, intRow, mIntColҩƷ����) = flexPicAlignLeftCenter
                Else
                    If Val(!������) <> 0 Then
                        vsfList.Cell(flexcpPicture, intRow, mIntColҩƷ����) = Me.ImgList.ListImages(39).Picture
                        vsfList.Cell(flexcpPictureAlignment, intRow, mIntColҩƷ����) = flexPicAlignLeftCenter
                    End If
                End If
                vsfList.TextMatrix(intRow, mIntColҩƷ����) = !Ʒ��
                
                If Not bln��ҩ���� And !�Ƿ�Ƥ�� = 1 Then
                    vsfList.TextMatrix(intRow, mintColƤ�Խ��) = GetƤ�Խ��(!����ID, !ҩ��ID, dateCurrent, !����ʱ��)
                    If vsfList.TextMatrix(intRow, mintColƤ�Խ��) <> "" Then
                        blnƤ�� = True
                    End If
                End If
                
                vsfList.TextMatrix(intRow, mIntCol������) = IIf(IsNull(!������), "", !������)
                vsfList.TextMatrix(intRow, mIntColӢ����) = IIf(IsNull(!Ӣ����), "", !Ӣ����)
                vsfList.TextMatrix(intRow, mIntCol�䷽����) = IIf(IsNull(!�䷽����), "", !�䷽����)
                vsfList.TextMatrix(intRow, mintcol���) = !���
                vsfList.TextMatrix(intRow, mintcol���) = IIf(IsNull(!���), "", !���)
                vsfList.TextMatrix(intRow, mintcol����) = IIf(IsNull(!����), "", !����)
                vsfList.TextMatrix(intRow, mintcolЧ��) = IIf(IsNull(!Ч��), "", !Ч��)
                vsfList.TextMatrix(intRow, mIntColId) = !�շ�ID
                vsfList.TextMatrix(intRow, mintcolҩƷid) = !ҩƷID
                vsfList.TextMatrix(intRow, mintcol����) = !����
                vsfList.TextMatrix(intRow, mintcol��λ) = IIf(IsNull(!��λ), "", !��λ)
                vsfList.TextMatrix(intRow, mIntCol����) = Format(!����, "#0." & String(mintPriceDigit, "0"))
                vsfList.TextMatrix(intRow, mIntCol����) = Format(!����, "#####0;-#####0; ;")
                vsfList.TextMatrix(intRow, mIntCol����) = !����
                vsfList.TextMatrix(intRow, mIntColNO) = !NO
                vsfList.TextMatrix(intRow, mIntColסԺ��) = zlStr.NVL(!סԺ��)
                vsfList.TextMatrix(intRow, mIntCol�����) = zlStr.NVL(!�����)
                vsfList.TextMatrix(intRow, mIntCol����ҩƷ˵��) = zlStr.NVL(!����ҩƷ˵��)
                vsfList.TextMatrix(intRow, mintcol������) = zlStr.NVL(!����)
                vsfList.TextMatrix(intRow, mintcolԭ����) = zlStr.NVL(!ԭ����)
                
                If mcondition.bln��ʾ��С��λ = True Then
                    '����С��װ��ʾ����
                    lng������ = Int(!����)
                    If !�ۼ۵�λ = !��λ Then
                        '�ۼ۵�λ�����ﵥλ������ͬ�������ֱ�������ﵥλ�����赥λ�Ļ���
                        vsfList.TextMatrix(intRow, mintcol����) = !���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                    Else
                        '�ۼ۵�λ�����ﵥλ��ͬ����Ҫ���л���
                        If !ʵ������ = 0 Then
                            dblʵ������ = !С��λ����
                        Else
                            dblʵ������ = !ʵ������
                        End If
                        
                        If dblʵ������ < 0 Then
                            lng������ = -Int(Abs(dblʵ������) / !��װ)
                        Else
                            lng������ = Int(dblʵ������ / !��װ)
                        End If

                        If lng������ = 0 Then
                            '��������С��1ʱ�����ۼ۵�λ��ʵ������ֱ����ʾ
                            vsfList.TextMatrix(intRow, mintcol����) = Abs(dblʵ������) & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                        Else
                            '������������1ʱ����Ҫ�����Ƿ�����ͷ�����������:1��3Ƭ
                            If (dblʵ������ / !��װ) = lng������ Then
                                'û����ͷ
                                vsfList.TextMatrix(intRow, mintcol����) = Abs(lng������) & IIf(IsNull(!��λ), "", !��λ)
                            Else
                                '����ͷ
                                vsfList.TextMatrix(intRow, mintcol����) = Abs(lng������) & IIf(IsNull(!��λ), "", !��λ) & Abs((dblʵ������ - (lng������ * !��װ))) & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                            End If
                        End If
                    End If
                    
                    If !���� < 0 Then
                        vsfList.TextMatrix(intRow, mintcol����) = "���ˣ�" & vsfList.TextMatrix(intRow, mintcol����)
                    End If
                    vsfList.TextMatrix(intRow, mIntCol��װ) = Val(!��װ)
                Else
                    vsfList.TextMatrix(intRow, mintcol����) = zlStr.FormatEx(!����, 5)
                End If
                
                vsfList.TextMatrix(intRow, mIntCol���) = zlStr.FormatEx(Val(!���۽��), mintMoneyDigit, , True)
                vsfList.TextMatrix(intRow, mIntColʵ�ս��) = zlStr.FormatEx(Val(!ʵ�ս��), mintMoneyDigit, , True)
                vsfList.TextMatrix(intRow, mIntCol����) = !���� & !���㵥λ
                
                dbl������ = dbl������ + !����
                str������λ = !���㵥λ
                vsfList.TextMatrix(intRow, mIntColƵ��) = IIf(IsNull(!Ƶ��), "", !Ƶ��)
                vsfList.TextMatrix(intRow, mIntCol��ҩĿ��) = zlStr.NVL(!��ҩĿ��)
                If Not IsNull(!����) Then
                    vsfList.TextMatrix(intRow, mIntCol����) = zlStr.FormatEx(!����, mintNumberDigit) & "(" & zlStr.NVL(!���㵥λ) & ")"
                End If
                vsfList.TextMatrix(intRow, mIntCol�÷�) = zlStr.NVL(!�÷�)
                vsfList.TextMatrix(intRow, mIntCol�����־) = Val(!�����־)
                vsfList.TextMatrix(intRow, mIntCol��¼����) = Val(!��¼����)
                vsfList.TextMatrix(intRow, mIntCol��ҩ����) = zlStr.NVL(!��ҩ����)
                vsfList.TextMatrix(intRow, mIntCol���id) = Val(!���id)
                vsfList.TextMatrix(intRow, mIntCol����ҽ��) = zlStr.NVL(!����ҽ��)
                vsfList.TextMatrix(intRow, mIntColƵ�ʼ��) = zlStr.NVL(!Ƶ�ʼ��)
                vsfList.TextMatrix(intRow, mIntCol�����λ) = zlStr.NVL(!�����λ)
                vsfList.TextMatrix(intRow, mIntColҽ����־) = zlStr.NVL(!ҽ����־)
                vsfList.TextMatrix(intRow, mIntCol��ʼʱ��) = zlStr.NVL(!��ʼʱ��)
                vsfList.TextMatrix(intRow, mIntCol����ʱ��) = zlStr.NVL(!����ʱ��)
                vsfList.TextMatrix(intRow, mIntColƵ�ʴ���) = zlStr.NVL(!Ƶ�ʴ���)
                vsfList.TextMatrix(intRow, mIntCol����˵��) = zlStr.NVL(!����˵��)
                
                If mcondition.intListType = mListType.��ҩ Then
                    vsfList.TextMatrix(intRow, mIntCol��װ) = Val(!��װ)
                    If mcondition.bln��ʾ��С��λ = True Then
                        '����С��װ��ʾ�������ֱ�������������׼����������ҩ����
                        '����������׼����������ʾģʽΪ"���װ����+���װ��λ+С��װ����+�ۼ۵�λ"����ҩ����������ʾ����ֻ��ʾ��ֵ
                        lng������ = Int(!��������)
                        If !�ۼ۵�λ = !��λ Or lng������ = !�������� Then
                            vsfList.TextMatrix(intRow, mIntCol������) = lng������ & IIf(IsNull(!��λ), "", !��λ)
                        Else
                            
'                            lng������ = Int(!С��λ�������� / !��װ)
                            If lng������ = 0 Then
                                vsfList.TextMatrix(intRow, mIntCol������) = !С��λ������ & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                            Else
                                vsfList.TextMatrix(intRow, mIntCol������) = lng������ & IIf(IsNull(!��λ), "", !��λ) & (!С��λ������ Mod !��װ) & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                            End If
                        End If
                        
                        lng������ = Int(!׼����)
                        If !�ۼ۵�λ = !��λ Or lng������ = !׼���� Then
                            vsfList.TextMatrix(intRow, mIntCol׼����) = lng������ & IIf(IsNull(!��λ), "", !��λ)
                        Else
'                            lng������ = Int(!С��λ׼���� / !��װ)
                            If lng������ = 0 Then
                                vsfList.TextMatrix(intRow, mIntCol׼����) = !С��λ׼���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                            Else
                                vsfList.TextMatrix(intRow, mIntCol׼����) = lng������ & IIf(IsNull(!��λ), "", !��λ) & (!С��λ׼���� Mod !��װ) & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                            End If
                        End If
                        
                        lng������ = Int(!׼����)
                        If !�ۼ۵�λ = !��λ Then
                            vsfList.TextMatrix(intRow, mIntCol׼����С) = zlStr.FormatEx(lng������, mintNumberDigit)
                        ElseIf lng������ = !׼���� Then
                            vsfList.TextMatrix(intRow, mIntCol׼������) = zlStr.FormatEx(lng������, mintNumberDigit)
                            vsfList.TextMatrix(intRow, mIntCol׼����С) = zlStr.FormatEx(0, mintNumberDigit)
                        Else
'                            dblС���� = (Val(!׼����) - lng������) * !��װ
                            If lng������ = 0 Then
                                vsfList.TextMatrix(intRow, mIntCol׼����С) = zlStr.FormatEx(!С��λ׼����, mintNumberDigit)
                            Else
                                vsfList.TextMatrix(intRow, mIntCol׼������) = zlStr.FormatEx(lng������, mintNumberDigit)
                                vsfList.TextMatrix(intRow, mIntCol׼����С) = zlStr.FormatEx((!С��λ׼���� Mod !��װ), mintNumberDigit)
                            End If
                        End If
                        
                        vsfList.TextMatrix(intRow, mIntCol��ҩ��) = zlStr.FormatEx(!׼����, mintNumberDigit)
                        vsfList.TextMatrix(intRow, mIntCol��ҩ����) = vsfList.TextMatrix(intRow, mIntCol׼������)
                        vsfList.TextMatrix(intRow, mIntCol��ҩ��С) = vsfList.TextMatrix(intRow, mIntCol׼����С)
                        vsfList.TextMatrix(intRow, mIntCol��λ��) = IIf(IsNull(!��λ), "", !��λ)
                        vsfList.TextMatrix(intRow, mIntCol��λС) = IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                    Else
                        vsfList.TextMatrix(intRow, mIntCol������) = zlStr.FormatEx(!��������, mintNumberDigit)
                        vsfList.TextMatrix(intRow, mIntCol׼����) = zlStr.FormatEx(!׼����, mintNumberDigit)
                        vsfList.TextMatrix(intRow, mIntCol��ҩ��) = zlStr.FormatEx(!׼����, mintNumberDigit)
                    End If
                
                    vsfList.TextMatrix(intRow, mIntColʵ������) = !ʵ������
                Else
                    If mcondition.bln��ʾ��С��λ = True Then
                        '����С��װ��ʾ����
                        lng������ = Int(!�����)
                        If !�ۼ۵�λ = !��λ Or lng������ = !����� Then
                            vsfList.TextMatrix(intRow, mIntCol�����) = lng������ & IIf(IsNull(!��λ), "", !��λ)
                        Else
                            lng������ = Int(!���ʵ������ / !��װ)
                            If lng������ = 0 Then
                                vsfList.TextMatrix(intRow, mIntCol�����) = !���ʵ������ & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                            Else
                                vsfList.TextMatrix(intRow, mIntCol�����) = lng������ & IIf(IsNull(!��λ), "", !��λ) & (!���ʵ������ Mod !��װ) & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                            End If
                        End If
                    Else
                        vsfList.TextMatrix(intRow, mIntCol�����) = zlStr.FormatEx(zlStr.NVL(!�����, 0), mintNumberDigit)
                    End If
                
                    vsfList.TextMatrix(intRow, mIntCol��λ) = zlStr.NVL(!�ⷿ��λ)
                    vsfList.TextMatrix(intRow, mIntColҽ������) = zlStr.NVL(!ҽ������)
                    vsfList.TextMatrix(intRow, mIntColҽ��id) = zlStr.NVL(!ҽ��id)
                    
                    If Not gobjPass Is Nothing Then
                        vsfList.Cell(flexcpPicture, intRow, mIntCol�����, intRow, mIntCol�����) = gobjPass.zlPassSetWarnLight_YF(Val(!�����))
                        vsfList.Cell(flexcpPictureAlignment, intRow, mIntCol�����, intRow, mIntCol�����) = flexPicAlignCenterCenter
                    End If
                    
                    vsfList.TextMatrix(intRow, mIntCol����) = Val(!�����)
'                    '�������ڲ���PASS
'                    vsfList.Cell(flexcpPicture, intRow, mIntCol�����, intRow, mIntCol�����) = frmPublic.imgPass.ListImages(Val(!�����) + 2).Picture
'                    vsfList.Cell(flexcpPictureAlignment, intRow, mIntCol�����, intRow, mIntCol�����) = flexPicAlignCenterCenter
                End If
                
                vsfList.TextMatrix(intRow, mIntCol����) = IIf(IsNull(!����), 0, !����)
                vsfList.TextMatrix(intRow, mIntCol������) = ""
                vsfList.TextMatrix(intRow, mIntCol��Ч��) = ""
                vsfList.TextMatrix(intRow, mIntCol�²���) = ""
                vsfList.TextMatrix(intRow, mIntCol��ע) = ""
                vsfList.TextMatrix(intRow, mIntCol�ѱ�) = IIf(IsNull(!�ѱ�), "", !�ѱ�)
                
                dblӦ�ս�� = dblӦ�ս�� + Val(!���۽��)
                dblʵ�ս�� = dblʵ�ս�� + Val(!ʵ�ս��)
                
                '�Ե��ڿ�����޵�ҩƷ��ɫ
                vsfList.Redraw = flexRDNone
                If !������� = 0 Then
                    vsfList.Cell(flexcpForeColor, intRow, 1, intRow, vsfList.Cols - 1) = mlng��ɫ
                Else
                    vsfList.Cell(flexcpForeColor, intRow, 1, intRow, vsfList.Cols - 1) = vbBlack
                End If
                            
                '����ҩƷ������ʾ
                If InStr(";����ҩ;����ҩ;����I��;����II��;", zlStr.NVL(!�������)) > 0 And zlStr.NVL(!�������) <> "" Then
                    vsfList.Cell(flexcpFontBold, intRow, mIntColҩƷ����, intRow, mIntColҩƷ����) = True
                End If
            
                .MoveNext
            Loop
        End If
        
        '��ҩ��������ʾ
        If mcondition.intListType = mListType.��ҩ Then
            If mcondition.bln��ʾ��С��λ = True Then
                vsfList.Cell(flexcpFontBold, 1, mIntCol��ҩ����, intRow, mIntCol��ҩ����) = True
                vsfList.Cell(flexcpFontBold, 1, mIntCol��ҩ��С, intRow, mIntCol��ҩ��С) = True
            Else
                vsfList.Cell(flexcpFontBold, 1, mIntCol��ҩ��, intRow, mIntCol��ҩ��) = True
            End If
        End If
        
        '���հ�����ʾ���ϼ�
        intRow = intRow + 1
        vsfList.rows = intRow + 1

        If mcondition.int�����ʾ = 1 Then
            strSum = "ʵ�ս�" & Format(dblʵ�ս��, mstrVBMoneyForamt) & "Ԫ" & "(" & zlStr.ChineseMoney(dblʵ�ս��) & ")"
        ElseIf mcondition.int�����ʾ = 2 Then
            strSum = "Ӧ�ս�" & Format(dblӦ�ս��, mstrVBMoneyForamt) & "Ԫ" & "  ʵ�ս�" & Format(dblʵ�ս��, mstrVBMoneyForamt) & "Ԫ" & "(" & zlStr.ChineseMoney(dblʵ�ս��) & ")"
        Else
            strSum = "Ӧ�ս�" & Format(dblӦ�ս��, mstrVBMoneyForamt) & "Ԫ" & "(" & zlStr.ChineseMoney(dblӦ�ս��) & ")"
        End If
        
        If mcondition.bln��ʾ���� And mbln��ҩ���� Then
            strSum = strSum & "  ��������" & dbl������ & str������λ
        End If
        
        vsfList.Cell(flexcpText, intRow, mIntCol˳���, intRow, vsfList.Cols - 1) = strSum
        vsfList.Cell(flexcpAlignment, intRow, mIntCol˳���, intRow, vsfList.Cols - 1) = flexAlignLeftCenter
        vsfList.Cell(flexcpFontBold, intRow, mIntCol˳���, intRow, vsfList.Cols - 1) = True
        
        vsfList.MergeCells = flexMergeRestrictRows
        vsfList.MergeRow(vsfList.rows - 1) = True
        
        '��ҩ����
        picRecInfo_CM.Visible = False
        If txtNo.ListIndex <> -1 Then
            If txtNo.Tag <> "" Then
                If bln��ҩ���� Then
                    picRecInfo_CM.Visible = True
                    Call ��ҩ�����ر���(Val(Split(txtNo.Tag, "|")(0)), Val(txtNo.ItemData(txtNo.ListIndex)), Mid(txtNo.Text, 1, 8), Val(Split(txtNo.Tag, "|")(2)), Val(Split(txtNo.Tag, "|")(3)))
                    vsfList.ColWidth(mIntCol����) = 1200
                Else
                    vsfList.ColWidth(mIntCol����) = 0
                End If
            End If
        End If
        Call picRecipt_Resize
        
        '������������
        vsfList.Select 0, 0, vsfList.rows - 1, vsfList.Cols - 1
        vsfList.CellBorder &H9D9D9D, 1, 1, 1, 1, 1, 1
        
        '�����Ƥ�ԣ�����Ƥ�Խ��
        If blnƤ�� = True Then
            vsfList.ColWidth(mintColƤ�Խ��) = 800
            vsfList.Select 0, mIntColҩƷ����, vsfList.rows - 1, mintColƤ�Խ��
            vsfList.CellBorder &H9D9D9D, -1, -1, -1, -1, 0, 1
            
            For i = 1 To vsfList.rows - 1
                If vsfList.TextMatrix(i, mintColƤ�Խ��) = "(+)" Then
                    vsfList.Cell(flexcpForeColor, i, mintColƤ�Խ��, i, mintColƤ�Խ��) = vbRed
                ElseIf vsfList.TextMatrix(i, mintColƤ�Խ��) = "(-)" Then
                    vsfList.Cell(flexcpForeColor, i, mintColƤ�Խ��, i, mintColƤ�Խ��) = vbBlue
                Else
                    vsfList.Cell(flexcpForeColor, i, mintColƤ�Խ��, i, mintColƤ�Խ��) = &H80000008
                End If
            Next
        Else
            vsfList.ColWidth(mintColƤ�Խ��) = 0
        End If
        
        vsfList.Row = vsfList.rows - 1
        
        vsfList.Redraw = flexRDDirect
    End With
    
    Form_Resize
    Call picProcess_Resize
    Call InitColSelList(mcondition.intListType)
    RefreshList = True
    

    SetCmdSendPrivs int�����
    
    If Me.CmdSend.Caption = "��ҩȷ��(&O)" And int�Ŷ�״̬ = 1 Then
        Me.CmdSend.Caption = "ȡ��ȷ��(&C)"
    ElseIf Me.CmdSend.Caption = "ȡ��ȷ��(&C)" And int�Ŷ�״̬ = 0 Then
        Me.CmdSend.Caption = "��ҩȷ��(&O)"
    End If
End Function

Private Function �ж��Ƿ���ҩ����(ByVal lngNOҩ��id As Long, ByVal BillType As Integer, ByVal BillNo As String) As Boolean
    'ͨ��ҩƷid�ж��Ƿ�����ҩ
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim lngҩ��ID As Long
    Dim blnMoved As Boolean
    
    On Error GoTo errHandle
    
    lngҩ��ID = lngNOҩ��id
    If lngNOҩ��id = 0 Then lngҩ��ID = mcondition.lngҩ��ID
    
    strsql = "Select a.��� as ��� From �շ���ĿĿ¼ a ,ҩƷ�շ���¼ b Where b.ҩƷid=a.Id And b.����=[2] and b.No=[1] And (b.��¼״̬=1 Or Mod(b.��¼״̬,3)=0) and (b.�ⷿID+0=[3] OR b.�ⷿID IS NULL) " _
   
    '�������ת������ֱ�ӴӺ󱸱�����ȡ����
    blnMoved = Sys.IsMovedByNO("ҩƷ�շ���¼", BillNo, " ���� = ", BillType)
    If blnMoved Then
        gstrSQL = Replace(gstrSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
    End If
    
    Set rs = zldatabase.OpenSQLRecord(strsql, Me.Caption & "[�ж��Ƿ���ҩ����]", BillNo, BillType, lngҩ��ID)
    
    mbln��ҩ���� = IIf(rs!��� = 7, True, False)
    rs.Close
    
    �ж��Ƿ���ҩ���� = mbln��ҩ����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ��ҩ�����ر���(ByVal lngNOҩ��id As Long, ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer)
    '��ҩ������ʾԭʼ��������ҩ�巨
    Dim rs As New ADODB.Recordset
    Dim lngҩ��ID As Long
    
    On Error GoTo errHandle
    lngҩ��ID = lngNOҩ��id
    If lngNOҩ��id = 0 Then lngҩ��ID = mcondition.lngҩ��ID

    gstrSQL = "Select a.���,b.���� From ҩƷ�շ���¼ a ,������ü�¼ b Where a.����id=b.Id " _
        & " And a.����=[2] And a.No=[1] " _
        & " And (a.��¼״̬=1 Or Mod(a.��¼״̬,3)=0) and (a.�ⷿID+0=[3] OR a.�ⷿID IS NULL) "
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ҩ�����ر���]", BillNo, BillStyle, lngҩ��ID)
    
    lblԭʼ����.Caption = lblԭʼ����.Tag & CStr(IIf(IsNull(rs!����), 1, rs!����))
    lbl��ҩ�巨.Caption = lbl��ҩ�巨.Tag & IIf(IsNull(rs!���), "", rs!���)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function GetRecord(Optional ByRef int�ɲ��� As Integer) As ADODB.Recordset
    int�ɲ��� = mInt�ɲ���
    Set GetRecord = mrsDetail
End Function

Private Sub GetMoneyFormat()
    Dim n As Integer
    Dim strOracleTmp As String
    Dim strVbTmp As String
    
    strOracleTmp = "999999990."
    strVbTmp = "########0."
    For n = 1 To mintMoneyDigit
        strOracleTmp = strOracleTmp & "0"
        strVbTmp = strVbTmp & "0"
    Next
    
    mstrOracleMoneyForamt = strOracleTmp
    mstrVBMoneyForamt = strVbTmp
    
End Sub

Private Sub Chkȫ��_Click()
    Dim intRow As Integer
    Dim lng������ As Long
    Dim dblС���� As Double
    
    If mcondition.intListType <> mListType.��ҩ Then Exit Sub
    
    If Not Chkȫ��.Enabled Then Exit Sub
    With vsfList
        For intRow = 1 To .rows - 2
            If mcondition.bln��ʾ��С��λ = True Then
                If Chkȫ��.Value = 1 Then
                    .TextMatrix(intRow, mIntCol��ҩ����) = .TextMatrix(intRow, mIntCol׼������)
                    .TextMatrix(intRow, mIntCol��ҩ��С) = .TextMatrix(intRow, mIntCol׼����С)
                    
                    .TextMatrix(intRow, mIntCol��ҩ��) = zlStr.FormatEx(Val(.TextMatrix(intRow, mIntColʵ������)) / Val(.TextMatrix(intRow, mIntCol��װ)), 5)
                Else
                    .TextMatrix(intRow, mIntCol��ҩ��) = ""
                    .TextMatrix(intRow, mIntCol��ҩ����) = ""
                    .TextMatrix(intRow, mIntCol��ҩ��С) = ""
                End If
            Else
                .TextMatrix(intRow, mIntCol��ҩ��) = IIf(Chkȫ��.Value = 1, .TextMatrix(intRow, mIntCol׼����), "")
            End If
        Next
        mblnAllBack = (Chkȫ��.Value = 1)
    End With
End Sub

Private Sub CmdSend_Click()
    Dim blnmsg As Boolean
    
    If (mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ��) And mcondition.bln����˲��˺���ҩ����ͬ = False Then
        If Me.cbo�˲���.Text = Me.Cbo��ҩ��.Text Then
            If MsgBox("��ǰ��ҩ�����ĺ˲��˺���ҩ����ͬ���Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            Else
                blnmsg = True
            End If
        End If
        
        If InStr(1, Me.Cbo��ҩ��.Text, "-") < 1 And Not blnmsg Then
            If Mid(Me.cbo�˲���.Text, InStr(1, Me.cbo�˲���.Text, "-") + 1) = Me.Cbo��ҩ��.Text Then
                If MsgBox("��ǰ��ҩ�����ĺ˲��˺���ҩ����ͬ���Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If mcondition.intListType = mListType.��ҩ Then
        If frmҩƷ������ҩNew.RecipeWork(mcondition.intListType, False, vsfList) = False Then
            RefreshList mrsDetail
        End If
    Else
        If mcondition.lngҩ��ID <> Val(Split(txtNo.Tag, "|")(0)) Then
            If CDate(Format(Split(txtNo.Tag, "|")(4), "yyyy-MM-dd")) <> CDate(Format(Sys.Currentdate, "yyyy-MM-dd")) Then
                If MsgBox("        �����ǵ��쵥�ݣ���ɾ�������������»��ܣ�" & vbCrLf & "����Ѿ����˱���Ŀ�����Ҫ���³������Ƿ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        
        Call FormClear
        
        If frmҩƷ������ҩNew.RecipeWork(mcondition.intListType, mblnInput, vsfList, mblnδȡҩ��ҩ) = False Then
            RefreshList mrsDetail
        End If
        
        mblnδȡҩ��ҩ = False
    End If
    
    If mblnInput = True Then
        txtNo.SetFocus
        Call zlControl.TxtSelAll(txtNo)
        mblnInput = False
    End If
End Sub

Private Sub cmdSendByNoTake_Click()
    mblnδȡҩ��ҩ = True
    Call CmdSend_Click
End Sub


Private Sub imgDown_Click()
    imgDown.Visible = False
    imgUp.Visible = True
    
    picRecipt_Resize
End Sub
Private Sub imgUp_Click()
    imgDown.Visible = True
    imgUp.Visible = False
    
    picRecipt_Resize
End Sub

Private Sub lblDiag_Click()
    If Me.imgDown.Visible Then
        imgDown_Click
    Else
        imgUp_Click
    End If
End Sub

Private Sub picHscSend_Click()
    If Me.imgDown.Visible Then
        imgDown_Click
    Else
        imgUp_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveListColState
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    On Error Resume Next
    
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList.ColHidden(.RowData(i)) Or vsfList.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = fraColSel.Top + fraColSel.Height
                If .Top + .Height > Me.ScaleHeight - picRecipt.Top - vsfList.Top Then
                    .Height = Me.ScaleHeight - .Top - picRecipt.Top - vsfList.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                
                .Left = fraColSel.Left
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub InitColSelList(ByVal intListType As Integer)
    Dim i As Integer
    
    With vsfColSel
        .Tag = intListType
        
        .rows = .FixedRows
        For i = 1 To vsfList.Cols - 1
            '���ڲ�������ʾ�б���в��ܼ�����ѡ���б�
            If IsInString(mstrUnallowShow, vsfList.ColKey(i), ";") = False Then
                If (mcondition.bln��ʾԭ���� And vsfList.ColKey(i) = "ԭ����") Or vsfList.ColKey(i) <> "ԭ����" Then
                    .rows = .rows + 1
                    .TextMatrix(.rows - 1, 1) = vsfList.TextMatrix(0, i)
                    .RowData(.rows - 1) = i
                End If
                
                '�п�Ϊ�ջ������ص�������Ϊ����ѡ
                If Not (vsfList.ColWidth(i) = 0 Or vsfList.ColHidden(i)) Then
                    .TextMatrix(.rows - 1, 0) = 0
                End If
                
                'ָ����������Ϊ������������
                If IsInString(mstrUnallowSetColHide, vsfList.ColKey(i), ";") = True Then
                    .Cell(flexcpForeColor, .rows - 1, 1) = .BackColorFixed
                End If
            End If
        Next
    End With
End Sub
Private Sub Form_Load()
    mblnAllowClick = True
    
    mstrPrivs = gstrprivs
    
    mlngMode = glngModul
    
    Lblҩ��.Caption = ""
    
    'ȡ���λ��
    mintMoneyDigit = Val(zldatabase.GetPara("���ý���λ��", glngSys, 0))
    
    '���ý���ʽ
    Call GetMoneyFormat
    
    Call Loadҽ��
    
    Call SetParams

    Call InitComandBars
    
    Call FormClear
    
    picRecInfo_CM.BackColor = &H8000000F
    picProcess.BackColor = &H8000000F
    picRecInfo.BackColor = &H8000000F
End Sub

Private Sub Load��ҩ��(ByVal lngҩ��ID As Long)
    '��ҩ��
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = " Select ����||'-'||���� As ����,���� As ���� From ��Ա��  Where ID in " & _
             " (Select Distinct ��ԱID From ��Ա����˵�� Where ��Ա����='ҩ����ҩ��' " & _
             " And ��ԱID IN (Select ��ԱID From ������Ա Where ����ID=[1]))" & _
             " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��ҩ��", lngҩ��ID)
    
    With rsData
        Me.Cbo��ҩ��.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            Cbo��ҩ��.AddItem !����
            
            If InStr(1, mstrDosUser, "-") > 0 Then
                mstrDosUser = Mid(mstrDosUser, InStr(1, mstrDosUser, "-") + 1)
            End If
            
            If mstrDosUser = !���� Then
                intIndex = .AbsolutePosition - 1
            End If

            .MoveNext
        Loop

        Cbo��ҩ��.Enabled = Not Cbo��ҩ��.ListCount = 0

        If intIndex <> -1 Then Cbo��ҩ��.ListIndex = intIndex
        mstrDosUser = Me.Cbo��ҩ��.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load�˲���(ByVal lngҩ��ID As Long)
    '�˲���
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = " Select ����||'-'||���� As ����,���� As ���� From ��Ա��  Where ID in " & _
             " (Select Distinct ��ԱID From ��Ա����˵�� Where ��Ա����='ҩ����ҩ��' " & _
             " And ��ԱID IN (Select ��ԱID From ������Ա Where ����ID=[1]))" & _
             " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
             
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�����", lngҩ��ID)
    
    With rsData
        Me.cbo�˲���.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            cbo�˲���.AddItem !����
            
            If InStr(1, mstr�˲���, "-") > 0 Then
                mstr�˲��� = Mid(mstr�˲���, InStr(1, mstr�˲���, "-") + 1)
            End If
            
            If mstr�˲��� = !���� Then
                intIndex = .AbsolutePosition - 1
            End If

            .MoveNext
        Loop

        cbo�˲���.Enabled = Not cbo�˲���.ListCount = 0

        If intIndex <> -1 Then cbo�˲���.ListIndex = intIndex
        mstr�˲��� = Me.cbo�˲���.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If cbsMain.count = 1 Then
        picRecipt.Move 0, 0, Me.Width, Me.Height
    End If
End Sub


Private Sub picProcess_Resize()
    On Error Resume Next
    
    With CmdSend
        .Left = picProcess.Width - .Width - 100
    End With
    
    With cmdSendByNoTake
        .Left = CmdSend.Left - .Width - 100
    End With
    
    With Chkȫ��
        .Left = IIf(cmdSendByNoTake.Visible, cmdSendByNoTake.Left, CmdSend.Left) - .Width - 100
    End With
End Sub


Private Sub picRecInfo_Resize()
    Dim objTmp As Object
    
    On Error Resume Next

    With txt�������
        .Width = picRecInfo.Width - .Left - 50
    End With
    
    With lblNotice
        .Top = txt�������.Top + txt�������.Height + 200
    End With

    
    With picRecInfo
        .Height = lblNotice.Top + lblNotice.Height + 100
    End With
End Sub


Private Sub picRecipt_Resize()
    On Error Resume Next
    
    With picRecInfo
        .Top = 0
        .Left = 0
        .Width = picRecipt.Width
    End With
    
    With picProcess
        .Top = picRecipt.Height - .Height
        .Left = 0
        .Width = picRecipt.Width
    End With
        
    With picRecInfo_CM
        If .Visible Then
            .Top = picProcess.Top - .Height - 100
            .Left = 0
            .Width = picRecipt.Width
        End If
    End With
    
    With vsfList
        .Top = picRecInfo.Top + picRecInfo.Height + 100
        .Left = 0
        .Width = picRecipt.Width
        .Height = IIf(picRecInfo_CM.Visible, picRecInfo_CM.Top, picProcess.Top) - picRecInfo.Height - 100 - IIf(Me.picHscSend.Visible, Me.picHscSend.Height, 0) - IIf(imgDown.Visible, Me.txt��ҩ����.Height, 0)
    End With
    
    With picHscSend
        .Top = Me.vsfList.Top + Me.vsfList.Height - IIf(Me.picHscSend.Visible, 0, Me.picHscSend.Height)
        .Left = Me.vsfList.Left
        .Width = Me.vsfList.Width - 20
    End With
    
    lblDiag.Left = (picHscSend.Width - lblDiag.Width) / 2
    
    If imgDown.Visible Then
        With Me.txt��ҩ����
            .Visible = True
            .Top = Me.picHscSend.Top + Me.picHscSend.Height
            .Left = Me.picHscSend.Left
            .Width = Me.picHscSend.Width
        End With
    Else
        txt��ҩ����.Visible = False
    End If
    
    With fraColSel
        .Left = vsfList.Left + vsfList.ColWidth(0) - .Width - 30
        .Top = vsfList.Top + (vsfList.RowHeight(0) - .Height) / 2 + 30
        .ZOrder
    End With
End Sub

Private Sub InitComandBars()
    Dim cbrControl As CommandBarControl
    Dim objCmdBar As CommandBar
    Dim lngCount As Integer
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
'        .SetIconSize False, 24, 24
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.ActiveMenuBar.Visible = False
    Me.cbsMain.AddImageList Me.imgCheck
End Sub

Private Sub SetComandBars(ByVal intListType As Integer)
    Dim cbrControl As CommandBarControl
    Dim cbrControlSub As CommandBarControl
    Dim objCmdBar As CommandBar
    Dim lngCount As Integer
    Dim objMenu As CommandBarPopup
        
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    If (intListType <> mListType.����ҩ And intListType <> mListType.��ʱδ��) Then Exit Sub
    
    Select Case intListType
        Case mListType.����ҩ, mListType.��ʱδ��
            '���ù������˵�
            If Not gobjPass Is Nothing And IsInString(gstrprivs, "������ҩ���", ";") Then
'                Set objCmdBar = cbsMain.Add("����", xtpBarTop)
'                objCmdBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
'                objCmdBar.ModifyStyle XTP_CBRS_GRIPPER, 0
'                objCmdBar.ContextMenuPresent = False
'
'                Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowPlug, "����ʷ/����״̬")
'                cbrControl.BeginGroup = True
'                cbrControl.ToolTipText = "��ʾ����ʾ����ʷ/����״̬"
'                cbrControl.Style = xtpButtonIconAndCaption
'                cbrControl.IconId = 3

                If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, cbsMain, conMenu_Tool_ShowPlug, 3)
            End If
            
'            ���õ����˵���PASS
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_PASS, "PASS��&P)", 1, False)
            objMenu.Id = mconMenu_PASS
'            If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, cbsMain, mconMenu_PASS, 1)
'            With objMenu.CommandBar.Controls
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 0, "ҩ���ٴ���Ϣ�ο�(&C)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 1, "ҩƷ˵����(&D)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 2, "�й�ҩ��(&N)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 3, "������ҩ����(&S)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 4, "����ֵ(&T)")
'
'                Set cbrControl = .Add(xtpControlPopup, mconMenu_PASS_Item + 5, "ר����Ϣ(&P)")
'                cbrControl.BeginGroup = True
'
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 0, "ҩ��-ҩ���໥����(&D)", -1, False)
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 1, "ҩ��-ʳ���໥����(&F)", -1, False)
'
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 2, "����ע�������(&M)", -1, False)
'                cbrControlSub.BeginGroup = True
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 3, "����ע�������(&T)", -1, False)
'
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 4, "����֢(&C)", -1, False)
'                cbrControlSub.BeginGroup = True
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 5, "������(&S)", -1, False)
'
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 6, "��������ҩ(&G)", -1, False)
'                cbrControlSub.BeginGroup = True
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 7, "��ͯ��ҩ(&P)", -1, False)
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 8, "��������ҩ(&E)", -1, False)
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 9, "��������ҩ(&L)", -1, False)
'
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 6, "ҽҩ��Ϣ����(&I)")
'                cbrControl.BeginGroup = True
'
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 7, "ҩƷ�����Ϣ(&M)")
'                cbrControl.BeginGroup = True
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 8, "��ҩ;�������Ϣ(&R)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 9, "ҽԺҩƷ��Ϣ(&F)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 10, "����(&W)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 11, "���(&V)")
'            End With

    End Select
End Sub


Public Sub SetFontSize(ByVal intFont As Integer)
    With vsfList
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("��") + 100
        .RowHeightMax = TextHeight("��") + 100
        .Refresh
    End With
End Sub

Private Function Check�Ƿ����ҽ��(ByVal Int���� As Integer, ByVal strNo As String) As Boolean
    '�ж���סԺ�������ﲡ�ˣ�����ҽ����¼
    Dim rsData As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[1] And A.no=[2] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡҽ����¼", Int����, strNo)
    
    Check�Ƿ����ҽ�� = Not rsData.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub TxtNo_Click()
    GetRecipeByNO
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNo As String
    Dim rsData As ADODB.Recordset
    Dim rstemp As ADODB.Recordset
    Dim strTmp As String
    Dim ArrTmp
    Dim blnExit As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(txtNo.Text) = "" Then Exit Sub
    
    If Len(txtNo.Text) > 8 Or InStr(1, txtNo.Text, "-") > 0 Then
        If vsfList.rows > 1 Then
            If vsfList.TextMatrix(1, mIntColҩƷ����) <> "" Then
                If CmdSend.Enabled = True Then
                    CmdSend.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    
    strNo = GetFullNO(Trim(txtNo.Text), 13)
    
    Set rsData = frmҩƷ������ҩNew.GetRecipeByNO(strNo)
    
    DoEvents
    If rsData Is Nothing Then
        If mcondition.intListType <> mListType.��ҩ Then
            MsgBox zlStr.FormatString("�ô�����[1]�������ڣ������Ѿ���������ɾ���������������룡", strNo), vbInformation, gstrSysName
        Else
            MsgBox zlStr.FormatString("δ�ҵ�ָ��������[1]����ô����Ѿ���ҩ�����������룡", strNo), vbInformation, gstrSysName
        End If
        
        DoEvents
        txtNo.Text = ""
        txtNo.SetFocus
        Exit Sub
    ElseIf rsData.RecordCount = 0 Then
        '[��ҩ]��ǩ������ĵ��ݺ������[����ҩ]]��[����ҩ]�д��ڣ�����ʾ
'        Set rsData = frmҩƷ������ҩNew.GetRecipeByNO(strNo, 1)
'        rsData.Filter = "������� = Null"
        Set rstemp = frmҩƷ������ҩNew.GetRecipeByNO(strNo, 1)
        If rstemp Is Nothing Then
            MsgBox zlStr.FormatString("�ô�����[1]�������ڣ������Ѿ���������ɾ���������������룡", strNo), vbInformation, gstrSysName
            Exit Sub
        End If
        If rstemp.RecordCount = 0 Then
            MsgBox zlStr.FormatString("�ô�����[1]�������ڣ������Ѿ���������ɾ���������������룡", strNo), vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        If mcondition.intListType <> mListType.��ҩ Then
            If rsData.EOF Then
                MsgBox zlStr.FormatString("�ô�����[1]�������ڣ������Ѿ���������ɾ���������������룡", strNo), vbInformation, gstrSysName
                blnExit = True
            Else
                rsData.Sort = "�������,��¼״̬ desc"   'ȡ���һ�ε�����
                If mcondition.intListType = mListType.����ҩ Then
                    rsData.Filter = "������� = Null and ��ҩ���� = null"
                    If rsData.RecordCount = 0 Then
'                        If IsNull(rsData!��ҩ����) = False Then
'                            MsgBox zlStr.FormatString("�ô�����[1]���Ѿ���ҩȷ�ϣ����������룡", strNo), vbInformation, gstrSysName
'                            blnExit = True
'                        End If

                        '��ѯ[����ҩ]��[��ҩ]���Ƿ���ڸõ��ݡ�
                        rsData.Filter = "������� = Null And ��ҩ���� <> null"
                        If rsData.RecordCount > 0 Then
                            MsgBox zlStr.FormatString("�ô�����[1]���Ѿ���ҩȷ�ϣ����������룡", strNo), vbInformation, gstrSysName
                            blnExit = True
                        Else
                            rsData.Filter = "������� <> Null"
                            If rsData.RecordCount > 0 Then
                                MsgBox zlStr.FormatString("�ô�����[1]���Ѿ���ҩ�����������룡", strNo), vbInformation, gstrSysName
                                blnExit = True
                            End If
                        End If
                        
                    End If
                ElseIf mcondition.intListType = mListType.����ҩ Then
                    If mcondition.bln�Ƿ���Ҫ��ҩ���� Then
                        rsData.Filter = "������� = Null And ��ҩ���� <> null"
                    Else
                        rsData.Filter = "������� = Null"
                    End If
                    If rsData.RecordCount = 0 Then
'                        If IsNull(rsData!�������) = False Then
'                            MsgBox zlStr.FormatString("�ô�����[1]���Ѿ���ҩ�����������룡", strNo), vbInformation, gstrSysName
'                            blnExit = True
'                        ElseIf IsNull(rsData!��ҩ����) And mcondition.bln�Ƿ���Ҫ��ҩ���� Then
'                            MsgBox zlStr.FormatString("�ô�����[1]��δ��ҩ��ɣ����������룡", strNo), vbInformation, gstrSysName
'                            blnExit = True
'                        End If

                        '��ѯ[����ҩ]��[��ҩ]���Ƿ���ڸõ��ݡ�
                        rsData.Filter = "������� = Null And ��ҩ���� = null"
                        If rsData.RecordCount > 0 And mcondition.bln�Ƿ���Ҫ��ҩ���� Then
                            MsgBox zlStr.FormatString("�ô�����[1]��δ��ҩ��ɣ����������룡", strNo), vbInformation, gstrSysName
                            blnExit = True
                        Else
                            rsData.Filter = "������� <> Null"
                            If rsData.RecordCount > 0 Then
                                MsgBox zlStr.FormatString("�ô�����[1]���Ѿ���ҩ�����������룡", strNo), vbInformation, gstrSysName
                                blnExit = True
                            End If
                        End If
                        
                    End If
                End If
            End If
            
            If blnExit Then
                DoEvents    '��ֹ���㶨λtxtNoʧЧ��ԭ����Ƕ�봰��������
                txtNo.SelStart = 1: txtNo.SelLength = Len(txtNo.Text)
                txtNo.SetFocus
                Exit Sub
            End If
            
            If mcondition.intListType = mListType.����ҩ Then
                '���˳�δ����ҩ���ļ�¼
                rsData.Filter = "������� = Null"
            ElseIf mcondition.intListType = mListType.����ҩ Then
                '���˳�δ����ҩ���ļ�¼
                 rsData.Filter = "������� = Null"
            End If
        End If
    End If
    
    If rsData.RecordCount > 1 Then
        With vsfNoList
            .rows = 2
            .Redraw = flexRDNone
            Do While Not rsData.EOF
                .TextMatrix(.rows - 1, .ColIndex("ҩ��")) = rsData!ҩ��
                .TextMatrix(.rows - 1, .ColIndex("����")) = rsData!����
                .TextMatrix(.rows - 1, .ColIndex("NO")) = rsData!NO
                .TextMatrix(.rows - 1, .ColIndex("����")) = IIf(IsNull(rsData!����), "", rsData!����)
                .TextMatrix(.rows - 1, .ColIndex("�ⷿID")) = rsData!ҩ��ID
                .TextMatrix(.rows - 1, .ColIndex("����")) = rsData!����
                .TextMatrix(.rows - 1, .ColIndex("��¼����")) = rsData!��¼����
                .TextMatrix(.rows - 1, .ColIndex("�����־")) = rsData!�����־
                .TextMatrix(.rows - 1, .ColIndex("��������")) = rsData!��������
                .TextMatrix(.rows - 1, .ColIndex("��¼״̬")) = rsData!��¼״̬
                .rows = .rows + 1
                rsData.MoveNext
            Loop
            .Redraw = flexRDDirect
            .Top = txtNo.Top + txtNo.Height + 50
            .Width = 4500
            .Height = 1300
            .Left = txtNo.Left - (.Width - txtNo.Width)
            .Visible = True
            .ZOrder 0
            DoEvents
            .SetFocus
        End With
    ElseIf rsData.RecordCount = 1 Then
        If CheckAndProcessBill(mcondition.intListType, rsData!����, rsData!NO, rsData!ҩ��) = False Then
            DoEvents
            txtNo.Clear
            txtNo.Text = ""
            txtNo.SetFocus
            Exit Sub
        End If
        
        txtNo.Clear
        
        Do While Not rsData.EOF
            txtNo.AddItem rsData!NO & "--" & rsData!����
            txtNo.ItemData(txtNo.NewIndex) = rsData!����
            txtNo.Tag = rsData!ҩ��ID & "|" & rsData!ҩ�� & "|" & rsData!��¼���� & "|" & rsData!�����־ & "|" & rsData!�������� & "|" & rsData!��¼״̬
            Lblҩ��.Caption = rsData!ҩ��
            rsData.MoveNext
        Loop
        
        If txtNo.ListCount = 0 Then Exit Sub
        
        txtNo.ListIndex = 0
    End If
End Sub

Private Function CheckAndProcessBill(ByVal intType As Integer, ByVal Int���� As Integer, ByVal strNo As String, ByVal strҩ�� As String) As Boolean
    '��鵥��
    Dim rsTmp As ADODB.Recordset
    
    '����Ƿ�����ѱ�־ͣ����ҩƷ��¼
    '��Ӧ���ڣ���ҩ������ҩ������ҩ
    '���������Ƿ���лָ���ҩȨ�ޣ�����ָ���־
    On Error GoTo errHandle
    If intType = mListType.����ҩ Or intType = mListType.����ҩ Or intType = mListType.����ҩ Or intType = mListType.��ʱδ�� Then
        gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And Nvl(��ҩ��ʽ, 0) = -1 And Rownum = 1 "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ����ѱ�־ͣ��ҩ��ҩƷ��¼", Int����, strNo)
        
        If Not rsTmp.EOF Then
            If zlStr.IsHavePrivs(mstrPrivs, "�ָ���ҩ") = True And (intType = mListType.����ҩ Or intType = mListType.����ҩ Or intType = mListType.����ҩ Or intType = mListType.��ʱδ��) Then
                If MsgBox("[" & strҩ�� & "]����" & strNo & "�����ѱ��Ϊ���ٷ�ҩ��ҩƷ���Ƿ�ȡ����־������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '�������ȡ����־��ǰ�ⷿ�Ŀ��ÿ���Ƿ��㹻
                    If CheckUnSendStock(Int����, strNo) = False Then
                        Exit Function
                    End If
                    
                    'ȡ����ҩ��־
                    CancelUnCheck Int����, strNo
                Else
                    Exit Function
                End If
            Else
                MsgBox "[" & strҩ�� & "]����" & strNo & "�����ѱ��Ϊ���ٷ�ҩ��ҩƷ����û����Ӧ��Ȩ�ޣ����ܼ�����ҩ��", vbInformation, gstrSysName
                txtNo.Text = ""
                Exit Function
            End If
        End If
    End If
    
    '����Ƿ�����ҩ
    '��Ӧ���ڣ�����ҩ
    If intType = mListType.����ҩ Then
        gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And (��¼״̬=1 or Mod(��¼״̬,3)=1) And ��ҩ���� Is Null And Rownum = 1"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ�δ��ҩ", Int����, strNo)
            
        If rsTmp.EOF Then
            MsgBox "[" & strҩ�� & "]����" & strNo & "�Ѿ���ҩ�ˣ����������룡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '����Ƿ�����ҩ
    '��Ӧ���ڣ�����ҩ������ҩ���Ƿ���Ҫ��ҩ���ڲ�����
    If intType = mListType.����ҩ Or ((intType = mListType.����ҩ Or intType = mListType.��ʱδ��) And mcondition.bln�Ƿ���Ҫ��ҩ���� = True) Then
        gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And (��¼״̬=1 or Mod(��¼״̬,3)=1) And ��ҩ���� Is Not Null And ������� Is Null And Rownum = 1"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ�����ҩ", Int����, strNo)
            
        If rsTmp.EOF Then
            If intType = mListType.����ҩ Then
                MsgBox "[" & strҩ�� & "]����" & strNo & "��δ��ҩ�����Ѿ�ȡ����ҩ�ˣ����������룡", vbInformation, gstrSysName
                Exit Function
            ElseIf (intType = mListType.����ҩ Or intType = mListType.��ʱδ��) And mcondition.bln�Ƿ���Ҫ��ҩ���� = True Then
                MsgBox "[" & strҩ�� & "]����" & strNo & "��δ��ҩ�����������룡", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '����Ƿ��ѷ�ҩ
    '��黷�ڣ���ҩ������ҩ������ҩ
    If intType = mListType.����ҩ Or intType = mListType.����ҩ Or intType = mListType.����ҩ Or intType = mListType.��ʱδ�� Then
        gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And (��¼״̬=1 or Mod(��¼״̬,3)=1) And ������� Is Null And Rownum = 1"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ�δ��ҩ", Int����, strNo)
        
        If rsTmp.EOF Then
            MsgBox "[" & strҩ�� & "]����" & strNo & "�Ѿ���ҩ�����������룡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '����Ƿ��Ѿ�������ҩ
    '��黷�ڣ�����ҩ����������ҩ��������ǰ���£�
    If (intType = mListType.����ҩ Or intType = mListType.��ʱδ��) And zlStr.IsHavePrivs(mstrPrivs, "������ҩ���Ĵ���") = True Then
        gstrSQL = " Select A.NO, A.����, A.ҩƷid, A.���, Sum(Nvl(A.����, 1) * A.ʵ������) �ѷ����� " & _
            " From ҩƷ�շ���¼ A " & _
            " Where A.����� Is Not Null And A.��¼״̬ <> 1 And A.NO = [2] And A.�ⷿid <> [3] And A.���� = [1] " & _
            " Group By A.NO, A.����, A.ҩƷid, A.��� Having Sum(Nvl(A.����, 1) * A.ʵ������) > 0"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ�δ��ҩ", Int����, strNo, mcondition.lngҩ��ID)
        
        If Not rsTmp.EOF Then
            MsgBox "[" & strҩ�� & "]����" & strNo & "�Ѿ�������ҩ�����ܴ���ҩ�����������룡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckAndProcessBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CancelUnCheck(ByVal Int���� As Integer, ByVal strNo As String)
    'ȡ�����ٷ�ҩ��־����ָ���ĵ�����ִ��
    Dim rsData As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim i As Integer
    Dim arrSql As Variant
    
    arrSql = Array()
    On Error GoTo errHandle
    gstrSQL = "Select ID From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And Nvl(��ҩ��ʽ, 0) = -1 And ������� Is Null "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�����ٷ�ҩ��־", Int����, strNo)
    
    If rsData.EOF Then Exit Sub
    
    Do While Not rsData.EOF
        gstrSQL = "Zl_����ҩ�������_Unchecked(" & Val(rsData!Id) & ",0)"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        rsData.MoveNext
    Loop
    
    gcnOracle.BeginTrans
    blnTrans = True
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption & "-ȡ�����")
        Next
    gcnOracle.CommitTrans
    blnTrans = False
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans

    MsgBox "��ʾ������ʧ�ܡ�"
    Call SaveErrLog
End Sub
Private Function CheckUnSendStock(ByVal Int���� As Integer, ByVal strNo As String) As Boolean
    'ȡ���ѱ��Ϊ���ٷ�ҩ�ı�־��Ӧ����¼�뵥�ݺŷ�ʽ����Ҫ��鵱ǰ�ⷿ�Ŀ��������Ƿ��㹻
    '����飺0-�����;1-���,��������;2-���,�����ֹ
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    
    If mcondition.int����� = 0 Then
        CheckUnSendStock = True
        Exit Function
    End If
    On Error GoTo errHandle
    gstrSQL = "Select '[' || C.���� || ']' || C.���� || ' ' || C.��� As ����, A.ʵ������ * A.����, Nvl(B.��������, 0) As �������� " & _
        " From ҩƷ�շ���¼ A, �շ���ĿĿ¼ C, " & _
        " (Select ҩƷid, Nvl(����, 0) As ����, Nvl(��������, 0) As �������� " & _
        " From ҩƷ��� " & _
        " Where ���� = 1 And �ⷿid + 0 = [3]) B " & _
        " Where A.ҩƷid = C.ID And A.���� = [1] And A.NO = [2] And Nvl(A.��ҩ��ʽ, 0) = -1 And A.ҩƷid = B.ҩƷid(+) " & _
        " And Nvl(A.����, 0) = B.����(+) And A.ʵ������ * A.���� > Nvl(B.��������, 0) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "�����ÿ��", Int����, strNo, mcondition.lngҩ��ID)

    With rsData
        Do While Not .EOF
            strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & !����
            .MoveNext
        Loop
    End With
    
    If strMsg <> "" Then
        If mcondition.int����� = 1 Then
            strMsg = "����ҩƷ�ڻָ���ҩ��Ǻ󣬵�ǰ�ⷿ�Ŀ����������㣬�Ƿ������ҩ��" & vbCrLf & strMsg
            
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            
        ElseIf mcondition.int����� = 2 Then
            strMsg = "����ҩƷ�ڻָ���ҩ��Ǻ󣬵�ǰ�ⷿ�Ŀ����������㣬���ܷ�ҩ��" & vbCrLf & strMsg
            
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckUnSendStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt�������_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetTip(txt�������, txt�������.Tag)
End Sub

Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsfColSel.RowData(Row)
        If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
            vsfList.ColWidth(lngCol) = vsfList.ColData(lngCol)
            vsfList.ColHidden(lngCol) = False
        Else
            vsfList.ColWidth(lngCol) = 0
            vsfList.ColHidden(lngCol) = True
        End If
    End If
End Sub
Private Sub vsfColSel_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfColSel
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub


Private Sub vsfColSel_LostFocus()
    vsfColSel.Visible = False
End Sub

Private Sub vsfColSel_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsfColSel.Cell(flexcpForeColor, Row, 1) = vsfColSel.BackColorFixed Then Cancel = True
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim blnUnValid As Boolean
    Dim dblCount As Double
    Dim dblSumCount As Double
    Dim rstemp As New ADODB.Recordset
    Dim dbl��ҩ�� As Double
    
    On Error GoTo errHandle
    With vsfList
        If mcondition.bln��ʾ��С��λ = True Then
            If Col <> mIntCol��ҩ���� And Col <> mIntCol��ҩ��С Then Exit Sub
        Else
            If Col <> mIntCol��ҩ�� Then Exit Sub
        End If
        
        blnUnValid = False
        dbl��ҩ�� = Val(.TextMatrix(Row, Col))
        
        If mcondition.bln��ʾ��С��λ = True Then
            If Col = mIntCol��ҩ���� Then
                dblSumCount = dbl��ҩ�� * Val(.TextMatrix(Row, mIntCol��װ)) + Val(.TextMatrix(Row, mIntCol��ҩ��С))
            Else
                dblSumCount = Val(.TextMatrix(Row, mIntCol��ҩ����)) * Val(.TextMatrix(Row, mIntCol��װ)) + dbl��ҩ��
            End If
        Else
            dblSumCount = dbl��ҩ��
        End If
        blnUnValid = Not ((Abs(dblSumCount) <= Abs(Val(.Tag))) And ((Val(dblSumCount) >= 0 And Val(.Tag) >= 0) Or (Val(dblSumCount) <= 0 And Val(.Tag) <= 0)))
        
        If blnUnValid Then
            If mcondition.bln��ʾ��С��λ = True Then
                If Col = mIntCol��ҩ���� Then
                    .TextMatrix(Row, Col) = Val(.TextMatrix(Row, mIntCol׼������))
                Else
                    .TextMatrix(Row, Col) = Val(.TextMatrix(Row, mIntCol׼����С))
                End If
            Else
                .TextMatrix(Row, Col) = Val(.Tag)
            End If
        End If
        
        '�ȼ���Ƿ���ҽ��������ҩƷ��¼
        '��������򲻹�
        '����ǣ����ϵͳ�����Ƿ�����δ����ҽ����ҩ�������������ҩ��Ϊ��
        '��������򲻹�
        dblCount = Val(.TextMatrix(Row, Col))
        If dblCount <> 0 And mcondition.blnҽ������ = False Then
            gstrSQL = "select ���� From ҩƷ�շ���¼ Where ID=[1] "
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ�������]", Val(vsfList.TextMatrix(vsfList.Row, mIntColId)))
            
            If (rstemp!���� Like "1*") Then       '����
                gstrSQL = "select B.ִ��״̬ from ����ҽ����¼ A,����ҽ������ B,������ü�¼ C where A.���id=B.ҽ��ID and A.id=C.ҽ����� and  C.ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                If Val(vsfList.TextMatrix(vsfList.Row, mIntCol��¼����)) = 1 Or (Val(vsfList.TextMatrix(vsfList.Row, mIntCol��¼����)) = 2 And (Val(vsfList.TextMatrix(vsfList.Row, mIntCol�����־)) = 1 Or Val(vsfList.TextMatrix(vsfList.Row, mIntCol�����־)) = 4)) Then
                Else
                    gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                End If
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[���ҽ���ĸ�ҩ;���Ƿ��Ѿ�ִ��]", Val(vsfList.TextMatrix(vsfList.Row, mIntColId)))
                
                If Not rstemp.EOF Then
                    If rstemp!ִ��״̬ = 0 Then
                        gstrSQL = "Select Nvl(ҽ�����,0) ҽ�����,Nvl(�����־,1) �����־ From ������ü�¼ Where ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                        If Val(vsfList.TextMatrix(vsfList.Row, mIntCol��¼����)) = 1 Or (Val(vsfList.TextMatrix(vsfList.Row, mIntCol��¼����)) = 2 And (Val(vsfList.TextMatrix(vsfList.Row, mIntCol�����־)) = 1 Or Val(vsfList.TextMatrix(vsfList.Row, mIntCol�����־)) = 4)) Then
                        Else
                            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                        End If
                        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ҽ��]", Val(vsfList.TextMatrix(vsfList.Row, mIntColId)))
                        
                        If Not rstemp.EOF Then
                            If (rstemp!�����־ = 1 Or rstemp!�����־ = 4) And rstemp!ҽ����� <> 0 Then
                                gstrSQL = "Select Nvl(��ҳid, 0) As ��ҳid, �Һŵ�, decode(ҽ��״̬,4,1,0) ���� From ����ҽ����¼ Where ������Դ=1  And ID=[1]"
                                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�жϸ�ҽ���Ƿ�����]", CLng(rstemp!ҽ�����))
                                
                                If Not rstemp.EOF Then
                                    If rstemp!��ҳid > 0 And IsNull(rstemp!�Һŵ�) Then
                                        '������ҳID����û�йҺŵ��Ĳ���ҽ���Ƿ����ϵ�����
                                    Else
                                        If rstemp!���� = 0 Then
                                            dblCount = 0
                                            MsgBox "�ñ�ҽ����δ���ϣ�������ҩ��", vbInformation, gstrSysName
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    gstrSQL = "Select Nvl(ҽ�����,0) ҽ�����,Nvl(�����־,1) �����־ From ������ü�¼ Where ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                    If Val(vsfList.TextMatrix(vsfList.Row, mIntCol��¼����)) = 1 Or (Val(vsfList.TextMatrix(vsfList.Row, mIntCol��¼����)) = 2 And (Val(vsfList.TextMatrix(vsfList.Row, mIntCol�����־)) = 1 Or Val(vsfList.TextMatrix(vsfList.Row, mIntCol�����־)) = 4)) Then
                    Else
                        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                    End If
                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ҽ��]", Val(vsfList.TextMatrix(vsfList.Row, mIntColId)))
                    
                    If Not rstemp.EOF Then
                        If (rstemp!�����־ = 1 Or rstemp!�����־ = 4) And rstemp!ҽ����� <> 0 Then
                            gstrSQL = "Select Nvl(��ҳid, 0) As ��ҳid, �Һŵ�, decode(ҽ��״̬,4,1,0) ���� From ����ҽ����¼ Where ������Դ=1  And ID=[1]"
                            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�жϸ�ҽ���Ƿ�����]", CLng(rstemp!ҽ�����))
                            
                            If Not rstemp.EOF Then
                                If rstemp!��ҳid > 0 And IsNull(rstemp!�Һŵ�) Then
                                    '������ҳID����û�йҺŵ��Ĳ���ҽ���Ƿ����ϵ�����
                                Else
                                    If rstemp!���� = 0 Then
                                        dblCount = 0
                                        MsgBox "�ñ�ҽ����δ���ϣ�������ҩ��", vbInformation, gstrSysName
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        .TextMatrix(Row, Col) = zlStr.FormatEx(dblCount, 5)
        
        If mcondition.bln��ʾ��С��λ = True Then
            If Col = mIntCol��ҩ���� Then
                .TextMatrix(Row, mIntCol��ҩ����) = zlStr.FormatEx(dblCount, 5)
            Else
                .TextMatrix(Row, mIntCol��ҩ��С) = zlStr.FormatEx(dblCount, 5)
            End If
            .TextMatrix(Row, mIntCol��ҩ��) = zlStr.FormatEx(dblSumCount, 5) / Val(.TextMatrix(Row, mIntCol��װ))
            
            If Val(.TextMatrix(Row, mIntCol��ҩ��)) <> Val(.TextMatrix(Row, mIntColʵ������)) / Val(.TextMatrix(Row, mIntCol��װ)) Then
                mblnAllBack = False
            End If
        Else
            .TextMatrix(Row, mIntCol��ҩ��) = zlStr.FormatEx(dblCount, 5)
            
            If Val(.TextMatrix(Row, mIntCol��ҩ��)) <> Val(.TextMatrix(Row, mIntCol׼����)) Then
                mblnAllBack = False
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Dim i As Integer
    
    '������ѡ���б�
    Call InitColSelList(mcondition.intListType)
    
    '������˳���
    For i = 0 To vsfList.Cols - 1
        Call SetColumnValue(vsfList.TextMatrix(0, i), i)
    Next
End Sub


Private Sub SetColumnValue(ByVal str���� As String, ByVal intValue As Integer)
    Select Case str����
        Case "��ҩĿ��"
            mIntCol��ҩĿ�� = intValue
        Case "�ⷿ��λ"
            mIntCol��λ = intValue
        Case "���"
            mintcol��� = intValue
        Case "�����"
            mIntCol����� = intValue
        Case "˳���"
            mIntCol˳��� = intValue
        Case "ҩƷ����"
            mIntColҩƷ���� = intValue
        Case "������"
            mIntCol������ = intValue
        Case "Ӣ����"
            mIntColӢ���� = intValue
        Case "�䷽����"
            mIntCol�䷽���� = intValue
        Case "���"
            mintcol��� = intValue
        Case "����"
            mintcol���� = intValue
        Case "��λ"
            mintcol��λ = intValue
        Case "����"
            mIntCol���� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "����"
            mintcol���� = intValue
        Case "���", "Ӧ�ս��"
            mIntCol��� = intValue
        Case "ʵ�ս��"
            mIntColʵ�ս�� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "�÷�"
            mIntCol�÷� = intValue
        Case "Ƶ��"
            mIntColƵ�� = intValue
        Case "ҽ������"
            mIntColҽ������ = intValue
        Case "�ѱ�"
            mIntCol�ѱ� = intValue
        Case "�����"
            mIntCol����� = intValue
        Case "�ⷿ��λ"
            mIntCol��λ = intValue
        Case "������"
            mIntCol������ = intValue
        Case "׼����"
            mIntCol׼���� = intValue
        Case "׼������"
            mIntCol׼������ = intValue
        Case "׼����С"
            mIntCol׼����С = intValue
        Case "��ҩ��"
            mIntCol��ҩ�� = intValue
        Case "��ҩ��(���װ)"
            mIntCol��ҩ���� = intValue
        Case "��λ(��)"
            mIntCol��λ�� = intValue
        Case "��ҩ��(С��װ)"
            mIntCol��ҩ��С = intValue
        Case "��λ(С)"
            mIntCol��λС = intValue
        Case "��ע"
            mIntCol��ע = intValue
        Case "��ҩ��(���װ)"
            mIntCol��ҩ���� = intValue
        Case "Ƥ�Խ��"
            mintColƤ�Խ�� = intValue
        Case "Ч��"
            mintcolЧ�� = intValue
        Case "�ѱ�"
            mIntCol�ѱ� = intValue
        Case "����˵��"
            mIntCol����˵�� = intValue
        Case "������"
            mintcol������ = intValue
        Case "ԭ����"
            mintcolԭ���� = intValue
    End Select
                   
End Sub

Private Sub vsfList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If Col = mintColƤ�Խ�� Then
        With vsfList
            If .ColWidth(mintColƤ�Խ��) > 400 Then
                .ColWidth(mIntColҩƷ����) = .ColWidth(mIntColҩƷ����) + (.ColWidth(mintColƤ�Խ��) - 400)
                .ColWidth(mintColƤ�Խ��) = 400
            Else
                .ColWidth(mIntColҩƷ����) = .ColWidth(mIntColҩƷ����) - (400 - .ColWidth(mintColƤ�Խ��))
                .ColWidth(mintColƤ�Խ��) = 400
            End If
        End With
    End If
End Sub

Private Sub vsfList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    '���ò����ƶ�����
    Select Case mcondition.intListType
        Case mListType.����ҩ, mListType.����ҩ, mListType.����ҩ, mListType.��ʱδ��, mListType.��ҩ, mListType.��ҩȷ��
            If Col = mIntColҩƷ���� Then
                Position = mIntColҩƷ����
            End If
            
            If Col = mintColƤ�Խ�� Then
                Position = mintColƤ�Խ��
            End If
            
            If Col = mIntCol˳��� Then
                Position = mIntCol˳���
            End If
        
            If Col = mIntCol����� Then
                Position = mIntCol�����
            End If
            
            If (Col <> mIntColҩƷ���� And Position = mIntColҩƷ����) Or (Col <> mintColƤ�Խ�� And Position = mintColƤ�Խ��) Or (Col <> mIntCol˳��� And Position = mIntCol˳���) Or (Col <> mIntCol����� And Position = mIntCol�����) Then
                Position = Col
            End If
    End Select
End Sub

Private Sub vsfList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '���ò��ܵ����п����
    Select Case mcondition.intListType
        Case mListType.����ҩ, mListType.����ҩ, mListType.����ҩ, mListType.��ʱδ��, mListType.��ҩ, mListType.��ҩȷ��
            If Col = mIntCol��ǰ�� Or Col = mIntCol˳��� Or Col = mIntCol����� Then Cancel = True
    End Select
End Sub

Private Sub vsfList_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    With vsfList
'        If Row = 0 Then Exit Sub
'        If mcondition.bln��ʾ��С��λ = True Then
'            If Col <> mIntCol��ҩ���� And Col <> mIntCol��ҩ��С Then Exit Sub
'            If Val(.TextMatrix(Row, Col)) > 0 Then
'                .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = &HBFC5FF
'            Else
'                .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = .BackColor
'            End If
'        Else
'            If Col <> mIntCol��ҩ�� Then Exit Sub
'            If Val(.TextMatrix(Row, Col)) > 0 Then
'                .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = &HBFC5FF
'            Else
'                .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = .BackColor
'            End If
'        End If
'    End With
End Sub

Private Sub vsfList_DblClick()
    Dim strID As String
    Dim strFlag As String
    
    If vsfList.Col = mIntCol����� Then
        If mcondition.intShowPass = 3 And IsInString(gstrprivs, "������ҩ���", ";") And IsNumeric(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("����"))) Then
            With vsfList
                If Val(.TextMatrix(.Row, mIntCol�����־)) = 1 Or (Val(.TextMatrix(.Row, mIntCol��¼����)) = 2 And Val(.TextMatrix(.Row, mIntCol�����־)) = 4) Then
                    strID = .TextMatrix(.Row, .ColIndex("�����"))
                    strFlag = "1"
                Else
                    strID = .TextMatrix(.Row, .ColIndex("סԺ��"))
                    strFlag = "2"
                End If
                If Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassQueryCheckResult_YF(mlngMode, strID, strFlag)
                End If
            End With
        End If
    End If
End Sub

Private Sub vsfList_EnterCell()
    With vsfList
        .Editable = flexEDNone
        
        Me.txt��ҩ����.Text = ""
        
        If .Row = 0 Then Exit Sub
        
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = Me.ImgList.ListImages(2).Picture
        If Val(.TextMatrix(.Row, mIntColId)) <> 0 Then Me.lblNotice.Caption = "����ҩƷ˵����" & .TextMatrix(.Row, mIntCol����ҩƷ˵��)
        
        
        If .TextMatrix(.Row, mIntCol��ҩĿ��) <> "" And Val(.TextMatrix(.Row, mIntColId)) <> 0 Then
            Me.picHscSend.Visible = True
'            Me.txt��ҩ����.Visible = True
            Me.txt��ҩ����.Text = "��ҩĿ�ģ�" & .TextMatrix(.Row, mIntCol��ҩĿ��) & vbCrLf & "��ҩ���ɣ�" & .TextMatrix(.Row, mIntCol��ҩ����)
            If Not mblnResize Then
                imgDown.Visible = False
                imgUp.Visible = True
            
                picRecipt_Resize
                mblnResize = True
            End If
        Else
            Me.picHscSend.Visible = False
            Me.txt��ҩ����.Visible = False
    
            If mblnResize Then
                imgDown.Visible = False
                imgUp.Visible = True
                picRecipt_Resize
                mblnResize = False
            End If
        End If
        
        If Val(.TextMatrix(.Row, mIntColId)) = 0 Then Exit Sub
        
        If mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ�� Then
            If Not gobjPass Is Nothing Then Call gobjPass.zlPassSetDrug_YF(.TextMatrix(.Row, mintcolҩƷid), .TextMatrix(.Row, mIntCol������))
        ElseIf mcondition.intListType = mListType.��ҩ Then
            If mcondition.bln��ʾ��С��λ = True Then
                If .Col <> mIntCol��ҩ���� And .Col <> mIntCol��ҩ��С Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol׼������)) = 0 And Val(.TextMatrix(.Row, mIntCol׼����С)) = 0 Then Exit Sub
                .Tag = Val(.TextMatrix(.Row, mIntCol׼������)) * Val(.TextMatrix(.Row, mIntCol��װ)) + Val(.TextMatrix(.Row, mIntCol׼����С))
                .Editable = flexEDKbdMouse
            Else
                If .Col <> mIntCol��ҩ�� Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol׼����)) = 0 Then Exit Sub
                .Tag = Val(.TextMatrix(.Row, mIntCol׼����))
                .Editable = flexEDKbdMouse
            End If
        End If
    End With
End Sub


Private Sub SetPassMenuButton(ByVal lngRow As Long)
    '����cmdAlley��ť״̬
    Dim cbrControl As CommandBarControl
    Dim rsData As ADODB.Recordset
    
    If mcondition.intShowPass <> 1 Or Not IsInString(gstrprivs, "������ҩ���", ";") Then Exit Sub
    
    '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�����Ͳ���ʾcmdAlley��ť
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfList.TextMatrix(lngRow, vsfList.ColIndex("NO")), Val(vsfList.TextMatrix(lngRow, vsfList.ColIndex("����"))))
    
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, conMenu_Tool_ShowPlug, , True)
    
    If rsData.RecordCount = 0 Then
        If Not cbrControl Is Nothing Then cbrControl.Enabled = False
    Else
        If Not cbrControl Is Nothing Then cbrControl.Enabled = True
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    
    If mcondition.intListType <> mListType.��ҩ Then Exit Sub
    
    With vsfList
        strKey = .EditText
        If Col = mIntCol��ҩ�� Or Col = mIntCol��ҩ���� Or Col = mIntCol��ҩ��С Then
            If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
                Exit Sub
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If .EditSelLength = Len(strKey) Then Exit Sub
                If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= mintNumberDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub


Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim Int���� As Integer
    Dim strNo As String
    Dim str����� As Integer
    Dim lngҽ��id As Long
    Dim strsql As String
    Dim rsTmp As Recordset
    Dim lngPatiID As Long
    Dim str�Һŵ� As String
    Dim lng��ҳID As Long
    
    If vsfList.Row = 0 Then Exit Sub
 
    If Button = 2 Then
        If Not gobjPass Is Nothing And IsInString(gstrprivs, "������ҩ���", ";") And (mcondition.intListType = mListType.����ҩ Or mcondition.intListType = mListType.��ʱδ��) And vsfList.Col = vsfList.ColIndex("�����") Then
            Int���� = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("����")))
            strNo = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("NO"))
            lngҽ��id = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ҽ��id")))
            str����� = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("����"))
            
            '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�������˳�
            strsql = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.����id,0 ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,������ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] "
            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, strNo, Int����)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!����ID
            str�Һŵ� = NVL(rsTmp!�Һŵ�)
            lng��ҳID = rsTmp!��ҳid
            
            
  
            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_PASS)
            
            Call gobjPass.zlPASSPopupCommandBars_YF(mlngMode, objPopup.CommandBar, mconMenu_PASS, lngPatiID, lng��ҳID, str�Һŵ�, str�����, lngҽ��id)
            
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub


Private Sub vsfNoList_DblClick()
    vsfNoList_KeyDown vbKeyReturn, 0
End Sub


Private Sub vsfNoList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        vsfNoList.Visible = False
        txtNo.SetFocus
        txtNo.Text = ""
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then
        With vsfNoList
            If .Row = 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("����")) = "" Then Exit Sub
            
            If CheckAndProcessBill(mcondition.intListType, Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), .TextMatrix(.Row, .ColIndex("ҩ��"))) = False Then
                DoEvents
                txtNo.Text = ""
                txtNo.SetFocus
                Exit Sub
            End If
            
            txtNo.Clear
        
            txtNo.AddItem .TextMatrix(.Row, .ColIndex("NO")) & "--" & .TextMatrix(.Row, .ColIndex("����"))
            txtNo.ItemData(txtNo.NewIndex) = .TextMatrix(.Row, .ColIndex("����"))
            txtNo.Tag = Val(.TextMatrix(.Row, .ColIndex("�ⷿID"))) & "|" & .TextMatrix(.Row, .ColIndex("ҩ��")) & "|" & .TextMatrix(.Row, .ColIndex("��¼����")) & "|" & .TextMatrix(.Row, .ColIndex("�����־")) & "|" & .TextMatrix(.Row, .ColIndex("��������")) & "|" & .TextMatrix(.Row, .ColIndex("��¼״̬"))
            Lblҩ��.Caption = .TextMatrix(.Row, .ColIndex("ҩ��"))
            
            txtNo.ListIndex = 0
            
            .Visible = False
        End With
    End If
End Sub


Private Sub vsfNoList_LostFocus()
    If vsfNoList.Visible Then
        vsfNoList.Visible = False
    End If
End Sub


