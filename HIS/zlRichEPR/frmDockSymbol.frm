VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockSymbol 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "��������"
   ClientHeight    =   8475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDockSymbol.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picData 
      Align           =   2  'Align Bottom
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
      Height          =   6105
      Left            =   0
      ScaleHeight     =   6105
      ScaleWidth      =   10695
      TabIndex        =   4
      Top             =   2370
      Width           =   10695
      Begin VB.TextBox txtSearch 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   7800
         TabIndex        =   72
         ToolTipText     =   "������ƴ��������Ķ�λ,��λ�ɹ���س����벡��"
         Top             =   75
         Width           =   2400
      End
      Begin VB.PictureBox picYJS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   120
         ScaleHeight     =   2130
         ScaleWidth      =   6000
         TabIndex        =   27
         ToolTipText     =   "˫���հ��������ɲ��빦��"
         Top             =   120
         Width           =   6000
         Begin VB.TextBox txtYJ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   1560
            TabIndex        =   31
            Top             =   915
            Width           =   1890
         End
         Begin VB.TextBox txtYJ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   810
            TabIndex        =   30
            Top             =   1125
            Width           =   720
         End
         Begin VB.TextBox txtYJ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   810
            TabIndex        =   29
            Top             =   720
            Width           =   720
         End
         Begin VB.TextBox txtYJ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   45
            TabIndex        =   28
            Top             =   915
            Width           =   720
         End
         Begin VB.Line Line1 
            X1              =   795
            X2              =   1530
            Y1              =   1050
            Y2              =   1050
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
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
            Left            =   60
            TabIndex        =   35
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ���о�����"
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
            Left            =   675
            TabIndex        =   34
            Top             =   480
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����������"
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
            Left            =   675
            TabIndex        =   33
            Tag             =   "�����������"
            Top             =   1455
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�վ�����/ĩ��ͣ������"
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
            Index           =   3
            Left            =   1560
            TabIndex        =   32
            Top             =   675
            Width           =   1890
         End
      End
      Begin VB.PictureBox picFree 
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
         Height          =   2130
         Left            =   120
         ScaleHeight     =   2130
         ScaleWidth      =   5520
         TabIndex        =   52
         Top             =   120
         Width           =   5520
         Begin VB.ComboBox cboGroup 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   0
            Width           =   2475
         End
         Begin VSFlex8Ctl.VSFlexGrid mfgFree 
            Height          =   3105
            Left            =   0
            TabIndex        =   63
            Top             =   360
            Width           =   4800
            _cx             =   8467
            _cy             =   5477
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   14.25
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
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   420
            RowHeightMax    =   420
            ColWidthMin     =   420
            ColWidthMax     =   420
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
         Begin VB.Label lblGroup 
            AutoSize        =   -1  'True
            Caption         =   "�ַ��Ӽ�(&K)"
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
            Left            =   0
            TabIndex        =   54
            Top             =   60
            Width           =   990
         End
      End
      Begin VB.PictureBox picRY 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   120
         ScaleHeight     =   2130
         ScaleWidth      =   6000
         TabIndex        =   5
         Tag             =   "������ע"
         ToolTipText     =   "˫���հ��������ɲ��빦��"
         Top             =   120
         Width           =   6000
         Begin VB.Frame fraLineRYH 
            BackColor       =   &H00000000&
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
            Height          =   30
            Left            =   240
            TabIndex        =   7
            Top             =   1515
            Width           =   4065
         End
         Begin VB.Frame fraLineRYV 
            BackColor       =   &H00000000&
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
            Height          =   1635
            Left            =   2280
            TabIndex        =   6
            Top             =   225
            Width           =   30
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRY 
            Height          =   675
            Left            =   240
            TabIndex        =   8
            Top             =   1185
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   1191
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   16
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   350
            BackColorBkg    =   16777215
            GridColor       =   12632256
            AllowBigSelection=   0   'False
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   16
         End
         Begin VB.Label lblRYLeft 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   15
            TabIndex        =   17
            Top             =   1440
            Width           =   180
         End
         Begin VB.Label lblRYRight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   4395
            TabIndex        =   16
            Top             =   1440
            Width           =   180
         End
         Begin VB.Label lblRYDn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
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
            Left            =   2100
            TabIndex        =   15
            Top             =   1905
            Width           =   360
         End
         Begin VB.Label lblRYUp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
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
            Left            =   2100
            TabIndex        =   14
            Top             =   45
            Width           =   360
         End
         Begin VB.Label lblRY 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmDockSymbol.frx":000C
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Index           =   0
            Left            =   2475
            TabIndex        =   13
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblRY 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmDockSymbol.frx":001E
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Index           =   1
            Left            =   2790
            TabIndex        =   12
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblRY 
            BackStyle       =   0  'Transparent
            Caption         =   "    �����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Index           =   2
            Left            =   3135
            TabIndex        =   11
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblRY 
            BackStyle       =   0  'Transparent
            Caption         =   "��һ��ĥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Index           =   3
            Left            =   3465
            TabIndex        =   10
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblRY 
            BackStyle       =   0  'Transparent
            Caption         =   "�ڶ���ĥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Index           =   4
            Left            =   3810
            TabIndex        =   9
            Top             =   255
            Width           =   165
         End
      End
      Begin VB.PictureBox picHY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   120
         ScaleHeight     =   2130
         ScaleWidth      =   5580
         TabIndex        =   36
         Tag             =   $"frmDockSymbol.frx":0032
         ToolTipText     =   "˫���հ��������ɲ��빦��"
         Top             =   120
         Width           =   5580
         Begin VB.Frame fraLineHYV 
            BackColor       =   &H00000000&
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
            Height          =   1635
            Left            =   2700
            TabIndex        =   38
            Top             =   210
            Width           =   30
         End
         Begin VB.Frame fraLineHYH 
            BackColor       =   &H00000000&
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
            Height          =   30
            Left            =   15
            TabIndex        =   37
            Top             =   1500
            Width           =   5505
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHY 
            Height          =   675
            Left            =   15
            TabIndex        =   39
            Top             =   1170
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   1191
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   16
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   350
            BackColorBkg    =   16777215
            GridColor       =   12632256
            AllowBigSelection=   0   'False
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   16
         End
         Begin VB.Label lblHYRight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   5280
            TabIndex        =   41
            Top             =   60
            Width           =   180
         End
         Begin VB.Label lblHYLeft 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   15
            TabIndex        =   40
            Top             =   75
            Width           =   180
         End
         Begin VB.Label lblHY 
            BackStyle       =   0  'Transparent
            Caption         =   "  ����ĥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   7
            Left            =   5265
            TabIndex        =   51
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblHY 
            BackStyle       =   0  'Transparent
            Caption         =   "  �ڶ�ĥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   6
            Left            =   4920
            TabIndex        =   50
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblHY 
            BackStyle       =   0  'Transparent
            Caption         =   "  ��һĥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   5
            Left            =   4575
            TabIndex        =   49
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblHY 
            BackStyle       =   0  'Transparent
            Caption         =   "�ڶ�ǰĥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   4
            Left            =   4230
            TabIndex        =   48
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblHY 
            BackStyle       =   0  'Transparent
            Caption         =   "��һǰĥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   3
            Left            =   3885
            TabIndex        =   47
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblHY 
            BackStyle       =   0  'Transparent
            Caption         =   "      ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   2
            Left            =   3555
            TabIndex        =   46
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblHY 
            BackStyle       =   0  'Transparent
            Caption         =   "    ������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   1
            Left            =   3210
            TabIndex        =   45
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblHY 
            BackStyle       =   0  'Transparent
            Caption         =   "    ������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   0
            Left            =   2865
            TabIndex        =   44
            Top             =   255
            Width           =   165
         End
         Begin VB.Label lblHYUp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
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
            Left            =   2520
            TabIndex        =   43
            Top             =   45
            Width           =   360
         End
         Begin VB.Label lblHYDn 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
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
            Left            =   2520
            TabIndex        =   42
            Top             =   1890
            Width           =   360
         End
      End
      Begin VB.PictureBox picSpot 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   120
         ScaleHeight     =   2130
         ScaleWidth      =   6000
         TabIndex        =   18
         ToolTipText     =   "˫���հ��������ɲ��빦��"
         Top             =   120
         Width           =   6000
         Begin VB.Line Line2 
            Index           =   0
            X1              =   960
            X2              =   960
            Y1              =   155
            Y2              =   1680
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   1764
            X2              =   194
            Y1              =   930
            Y2              =   930
         End
         Begin VB.Line Line7 
            Visible         =   0   'False
            X1              =   2535
            X2              =   3645
            Y1              =   435
            Y2              =   1545
         End
         Begin VB.Line Line8 
            Visible         =   0   'False
            X1              =   2520
            X2              =   3675
            Y1              =   1560
            Y2              =   405
         End
         Begin VB.Label lblPot 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   480
            TabIndex        =   26
            Top             =   1110
            Width           =   330
         End
         Begin VB.Label lblPot 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   1155
            TabIndex        =   25
            Top             =   1110
            Width           =   330
         End
         Begin VB.Label lblPot 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   480
            TabIndex        =   24
            Top             =   435
            Width           =   330
         End
         Begin VB.Label lblPot 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   1155
            TabIndex        =   23
            Top             =   435
            Width           =   330
         End
         Begin VB.Label lblPot 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   2520
            TabIndex        =   22
            Top             =   810
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lblPot 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   3375
            TabIndex        =   21
            Top             =   810
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lblPot 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   2910
            TabIndex        =   20
            Top             =   420
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lblPot 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   2910
            TabIndex        =   19
            Top             =   1230
            Visible         =   0   'False
            Width           =   330
         End
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   3480
         Top             =   3600
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
               Picture         =   "frmDockSymbol.frx":003F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockSymbol.frx":05D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockSymbol.frx":0B73
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockSymbol.frx":110D
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   735
         Left            =   120
         TabIndex        =   75
         Top             =   2400
         Width           =   6015
         _cx             =   10610
         _cy             =   1296
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   3
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
      Begin MSComctlLib.ImageList img16 
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
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockSymbol.frx":19E7
               Key             =   "Selected"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockSymbol.frx":1F81
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockSymbol.frx":251B
               Key             =   "ǩ��"
            EndProperty
         EndProperty
      End
      Begin XtremeCommandBars.CommandBars CommandBars 
         Left            =   1680
         Top             =   4320
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
      Begin VB.Shape shpSearch 
         BorderColor     =   &H00E09060&
         Height          =   270
         Left            =   7620
         Top             =   480
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblSearch 
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
         Height          =   195
         Left            =   6990
         TabIndex        =   73
         Top             =   153
         Width           =   600
      End
   End
   Begin VB.PictureBox picPre 
      Align           =   1  'Align Top
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
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10695
      TabIndex        =   1
      Top             =   1650
      Width           =   10695
      Begin VB.CheckBox chkLanguage 
         Caption         =   "Ӣ"
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
         Index           =   1
         Left            =   3675
         TabIndex        =   91
         Top             =   435
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkLanguage 
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
         Height          =   210
         Index           =   0
         Left            =   3150
         TabIndex        =   90
         Top             =   435
         Width           =   435
      End
      Begin VB.CheckBox chkCY 
         Caption         =   "Ӥ"
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
         Index           =   1
         Left            =   4740
         TabIndex        =   89
         Top             =   435
         Value           =   1  'Checked
         Width           =   450
      End
      Begin VB.CheckBox chkCY 
         Caption         =   "ĸ"
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
         Index           =   0
         Left            =   4215
         TabIndex        =   88
         Top             =   435
         Value           =   1  'Checked
         Width           =   450
      End
      Begin VB.CheckBox chkRem 
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
         Height          =   195
         Left            =   60
         TabIndex        =   87
         Top             =   465
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.CheckBox chkref 
         Caption         =   "�ο�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   86
         Top             =   465
         Width           =   660
      End
      Begin VB.PictureBox picPhase 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   5280
         ScaleHeight     =   330
         ScaleWidth      =   3570
         TabIndex        =   81
         Top             =   360
         Width           =   3570
         Begin VB.OptionButton optPhase 
            Caption         =   "����"
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   85
            Top             =   60
            Value           =   -1  'True
            Width           =   720
         End
         Begin VB.OptionButton optPhase 
            Caption         =   "����"
            Height          =   210
            Index           =   1
            Left            =   850
            TabIndex        =   84
            Top             =   60
            Width           =   720
         End
         Begin VB.OptionButton optPhase 
            Caption         =   "����"
            Height          =   210
            Index           =   2
            Left            =   1700
            TabIndex        =   83
            Top             =   60
            Width           =   720
         End
         Begin VB.OptionButton optPhase 
            Caption         =   "����"
            Height          =   210
            Index           =   3
            Left            =   2550
            TabIndex        =   82
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.ComboBox cboTimes 
         Height          =   330
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   15
         Width           =   3270
      End
      Begin VB.OptionButton optFormat 
         Caption         =   "�����ı�"
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
         Left            =   4200
         TabIndex        =   79
         Top             =   45
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton optFormat 
         Caption         =   "��ʽ�ı�"
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
         Left            =   3135
         TabIndex        =   78
         Top             =   45
         Width           =   1050
      End
      Begin VB.Frame fraSplit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   45
         TabIndex        =   64
         Top             =   690
         Width           =   3405
      End
      Begin VB.PictureBox picFormat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   2265
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   3
         Top             =   150
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdInsert 
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
         Height          =   350
         Left            =   60
         TabIndex        =   2
         Top             =   90
         Width           =   1200
      End
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
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
      Height          =   1650
      Left            =   0
      ScaleHeight     =   1650
      ScaleWidth      =   10695
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin VB.Frame fraType 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   45
         Index           =   3
         Left            =   0
         TabIndex        =   77
         Top             =   1575
         Width           =   4935
      End
      Begin VB.Frame fraType 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   45
         Index           =   2
         Left            =   0
         TabIndex        =   65
         Top             =   1182
         Width           =   4935
      End
      Begin VB.Frame fraType 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   45
         Index           =   1
         Left            =   0
         TabIndex        =   62
         Top             =   396
         Width           =   4935
      End
      Begin VB.Frame fraType 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   45
         Index           =   0
         Left            =   0
         TabIndex        =   61
         Top             =   789
         Width           =   4935
      End
      Begin VB.Label lblType 
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
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   76
         Top             =   1275
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "���ĵ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   74
         Top             =   870
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "̥��λ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   2625
         TabIndex        =   71
         ToolTipText     =   "������Ŀ��Դ����Ҫ���й��������������Ŀ"
         Top             =   870
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "�¾�ʷ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   2640
         TabIndex        =   70
         Top             =   1320
         Width           =   750
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00E09060&
         Height          =   270
         Left            =   4530
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblType 
         Caption         =   "��ѧ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   90
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   68
         Top             =   90
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "����ѡ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1800
         TabIndex        =   67
         Top             =   90
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "������ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2610
         TabIndex        =   66
         Top             =   90
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   60
         Top             =   492
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "����ҩ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   1800
         TabIndex        =   59
         Top             =   870
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "������ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   2610
         TabIndex        =   58
         Top             =   495
         Width           =   780
      End
      Begin VB.Label lblType 
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
         Height          =   225
         Index           =   6
         Left            =   1800
         TabIndex        =   57
         Top             =   495
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "��λ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   960
         TabIndex        =   56
         Top             =   495
         Width           =   780
      End
      Begin VB.Label lblType 
         Caption         =   "ҽѧ��λ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   960
         TabIndex        =   55
         Top             =   870
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmDockSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event GetPosFontSize()
Public Event InsertSymbol(strSymbol As String, intStrLen As Integer)
Public Event InsertPicSymbol(strInfor As String, picSy As StdPicture, strReturn As String)
Public Event InsertEPRDemo(lngEPRDemoID As Long)            '���뷶��
Public Event SetFouse()          '��txtSearch���»�ý���
Public mblnOk As Boolean         '����
Private mlngFileID As Long          '�����ļ�id
Private mlngPatient As Long         '����id���ڲ��˲����༭ʱ������ȷ������ʾ���Ƿ�����
Private mlngVisit As Long           '��ҳid��Һŵ�ID
Private mlngAdvice As Long          'ҽ��ID
Private mblnChk As Boolean         '��Ӣ���л�����
'�¾�ʷ������ʾ
Private Const YJ���� = "��������������������"
Private Const YJ��ĸ = "���˪�����ū۫񬩬�"
Private Const YJ����1 = _
        "�ͪϪѪӪժת٪۪ݪ�" & _
        "�����������" & _
        "��������������������" & _
        "��������������������" & _
        "�ǫɫ˫ͫϫѫӫի׫�" & _
        "�ݫ߫��������" & _
        "�������������������" & _
        "��������������������" & _
        "���ìŬǬɬˬͬϬѬ�"
Private Const YJ����2 = _
        "�����������ªĪƪȪ�" & _
        "�ΪЪҪԪ֪تڪܪު�" & _
        "������������" & _
        "��������������������" & _
        "�����������������«�" & _
        "�ȫʫ̫ΫЫҫԫ֫ث�" & _
        "�ޫ���������" & _
        "��������������������" & _
        "��������������������" & _
        "�¬ĬƬȬʬ̬άЬҬ�"
        
'������ע�ַ�
Private Const RY���� = "��������������������������������������������������"
Private Const RYС���� = "����������"
Private Const RYС��ĸ = "����������"
Private Const RY����� = "����������"
Private Const RY���ĸ = "����������"
Private Const RY����� = "����������"
Private Const RY���ĸ = "����������"
Private Const RY�ҷ��� = "����������"
Private Const RY�ҷ�ĸ = "����������"
'������ע�ַ�
Private Const HY���� = "��������������������������������������������������������������������������������������������������������������������������������"
Private Const HYС���� = "����������������"
Private Const HYС��ĸ = "����������������"
Private Const HY����� = "����������������"
Private Const HY���ĸ = "����������������"
Private Const HY����� = "����������������"
Private Const HY���ĸ = "����������������"
Private Const HY�ҷ��� = "����������������"
Private Const HY�ҷ�ĸ = "����������������"

'Word�������
Private Const CON������ As String = "�����������������U���E��F�����������o�p�q�r�s�t�u���C�򡪦����n������������񡲡���㡾��������硴����塸����顺�����v�w�x�y�z�{�������������A�@"
Private Const CON��λ���� As String = "����磤����꣥����H�����멈�T�L�M�N�Q�O�J�K�P����"
Private Const CON������� As String = "����������������������������������������������������������������������������¢âĢŢƢǢȢɢʢˢ̢͢΢ϢТѢҢӢԢբ֢עآ٢ڢۢܢݢޢߢ�������������"
Private Const CON��ѧ���� As String = "�֡ԡ٣��ܡݣ����ڡۡˡ��������£��ҡӡءޡġšơǡȡɡʡߡ�͡ΡϡСѡաס̨Q�R�P�ԩ������������N�S�S�R"
Private Const CON������� As String = "�����������졨������������������������I�G�����ߩh�i�l�m�j�k�|�}�~��ᨒ�ѡ��������I�J�L�K�ΨO���ܨM��"
Private Const CONҽѧ���� As String = "������������������"
Private Const CONҽѧ��λ As String = "��,�H,��,��,T,�O,��,��,��,��,��,��,��,��,��,��,��,��,��,��,g/L,mm/h,x10^6/L,x10^9/L,x10^12/L,��,��,ML,��/��,mmHg,��g,Bid,mmol/L,qd,Bw,IU/L,cm,mg,tid,mm,u/ml,ng/ml,��g/L,qW,umol/L,q8h"

'���ݱ�ע��ɫ
Private Const M_FLAGCOLOR = &HC0E0FF
'���ڵ��벡�˼�������
Private Enum mCol
    ��� = 0
    ѡ�� = 1
    ָ�� = 2
    ��� = 3
    ��־ = 4
    ��λ = 5
    �ο� = 6
    ������Դ = 7
    ���ʱ�� = 8
End Enum
'�°�LIS��Ҫ
Private Enum mcItem
    ����
    ����ID
    ������Դ
    ����ʱ��
    ������
    �����
    ���ʱ��
    ��Ŀ����
    �걾����
    Ӥ��
End Enum
Private Enum mcList
    ָ��
    ���
    ��λ
    ��־
    �ο�
    ���
    ��˽
    ����
    ������
    Ӣ����
End Enum
'�ڲ�����
Private mobjLis As Object
Private mstrInfor As String
Private mOldLisRs As ADODB.Recordset
Private mNewLisRs As ADODB.Recordset
Public mlFontSize As Long
Private Sub cboGroup_Click()
Dim intStart As Integer, i As Integer
    If Me.cboGroup.Visible = False Then Exit Sub
    If Me.ActiveControl.Name <> Me.cboGroup.Name Then Exit Sub
    
    intStart = 0
    For i = 0 To Me.cboGroup.ListIndex - 1
        intStart = intStart + Me.cboGroup.ItemData(i)
    Next
    
    With Me.mfgFree
        .Row = intStart \ .Cols
        .Col = intStart Mod .Cols
        .TopRow = .Row
        If .Visible Then .SetFocus
    End With
End Sub

Private Sub cboTimes_Click()
    Call FillLisItem
End Sub

Private Sub chkCY_Click(Index As Integer)
    If chkCY(0).Value = vbUnchecked And chkCY(1).Value = vbUnchecked Then '��ûѡ��
        chkCY(0).Value = vbChecked
    End If
    
    Call FilterLisItem
End Sub

Private Sub chkLanguage_Click(Index As Integer)
    If mblnChk = True Then Exit Sub
    mblnChk = True
    If Index = 0 Then
        chkLanguage(0).Value = 1: chkLanguage(1).Value = 0
    Else
        chkLanguage(0).Value = 0: chkLanguage(1).Value = 1
    End If
    mblnChk = False
End Sub

Private Sub cmdInsert_Click()
    If lblType(Val(shpSearch.Tag)).Caption = "������" Then
        Dim i As Integer, strGroup As String, strItem As String, strItems As String, strReturn As String
        With vsList
            If .Rows < 2 Then Exit Sub
            '�����ı���ָ����Ŀ˳����ɣ���ʽ�ı���һ���̶��������Ʊ��
            For i = 1 To .Rows - 1
                If .RowOutlineLevel(i) = 0 Then
Re:                 If strGroup <> "" And strGroup <> .Cell(flexcpData, i, mCol.ָ��) And strItems <> "" Then
                         If optFormat(1).Value = True Then
                            strReturn = strReturn & "��" & strGroup & ":" & Mid(strItems, 2)
                        Else
                            strReturn = strReturn & vbCrLf & vbCrLf & strGroup & ":" & vbCrLf & Mid(strItems, 3)
                        End If
                        strItems = ""
                    End If
                    If .Cell(flexcpData, i, mCol.���) <> "" Then
                        strGroup = "(" & Format(.Cell(flexcpData, i, mCol.���), "yyyy-mm-dd") & ")" & .Cell(flexcpData, i, mCol.ָ��)
                    Else
                        strGroup = .Cell(flexcpData, i, mCol.ָ��)
                    End If
                Else
                    If .Cell(flexcpData, i, mCol.ѡ��) = 1 Then
                        If optFormat(1).Value = True Then
                            strItem = ""
                            strItem = strItem & IIf(chkLanguage(0).Value = 1, Split(.Cell(flexcpData, i, mCol.ָ��), "|")(0), "")
                            strItem = strItem & IIf(chkLanguage(1).Value = 1 And chkLanguage(0).Value = 1, "(", "") '��ѡ����ʱӢ��������
                            strItem = strItem & IIf(chkLanguage(1).Value = 1, Split(.Cell(flexcpData, i, mCol.ָ��), "|")(1), "")
                            strItem = strItem & IIf(chkLanguage(1).Value = 1 And chkLanguage(0).Value = 1, ")", "") '��ѡ����ʱӢ��������
                            strItem = strItem & " " & .TextMatrix(i, mCol.���) & " " & .TextMatrix(i, mCol.��λ) & IIf(chkRem.Value = vbChecked, .TextMatrix(i, mCol.��־), "")
                            strItem = strItem & IIf(chkref.Value = vbChecked, " �ο�ֵ" & .TextMatrix(i, mCol.�ο�) & " " & .TextMatrix(i, mCol.��λ), "")
                            strItems = strItems & "��" & strItem
                        Else
                            strItem = ""
                            strItem = strItem & IIf(chkLanguage(0).Value = 1, Split(.Cell(flexcpData, i, mCol.ָ��), "|")(0), "")
                            strItem = strItem & IIf(chkLanguage(1).Value = 1 And chkLanguage(0).Value = 1, "(", "") '��ѡ����ʱӢ��������
                            strItem = strItem & IIf(chkLanguage(1).Value = 1, Split(.Cell(flexcpData, i, mCol.ָ��), "|")(1), "")
                            strItem = strItem & IIf(chkLanguage(1).Value = 1 And chkLanguage(0).Value = 1, ")", "") '��ѡ����ʱӢ��������
                            strItem = Rpad(strItem, 32)
                            strItem = strItem & Rpad(MidUni(.TextMatrix(i, mCol.���), 1, 8) & " " & MidUni(.TextMatrix(i, mCol.��λ), 1, 6) & IIf(chkRem.Value = vbChecked, .TextMatrix(i, mCol.��־), ""), 18)
                            strItem = strItem & Rpad(IIf(chkref.Value = vbChecked, "�ο�ֵ" & .TextMatrix(i, mCol.�ο�) & " " & MidUni(.TextMatrix(i, mCol.��λ), 1, 6), ""), 26)
                            strItems = strItems & vbCrLf & strItem
                        End If
                    End If
                    If i = .Rows - 1 Then GoTo Re
                End If
            Next
    
            .Cell(flexcpData, 1, mCol.ѡ��, .Rows - 1, mCol.ѡ��) = 0
            .Cell(flexcpData, 1, mCol.��־, .Rows - 1, mCol.��־) = 0
            Set .Cell(flexcpPicture, 1, mCol.ѡ��, .Rows - 1, mCol.ѡ��) = Nothing
            If strReturn = "" Then Exit Sub
            strReturn = IIf(optFormat(1).Value = True, Mid(strReturn, 2) & "��", Mid(strReturn, 3))
            RaiseEvent InsertSymbol(strReturn, Len(strReturn))
        End With
    Else
        If Not picFormat.Picture Is Nothing And mstrInfor <> "" Then
            RaiseEvent InsertPicSymbol(mstrInfor, picFormat.Image, picFormat.Tag)
            Set picFormat.Picture = Nothing
            cmdInsert.Enabled = False
        End If
    End If
End Sub



Private Sub Form_Load()
Dim i As Integer, j As Integer
    On Error Resume Next
    mlFontSize = 8
    
    '���б�׼�ַ�
    Dim aryFree(28, 1) As String
    aryFree(0, 0) = "����������": aryFree(0, 1) = " !" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    aryFree(1, 0) = "������-1������": aryFree(1, 1) = "��졧��������������������������¨�������������������������������������"
    aryFree(2, 0) = "������������": aryFree(2, 1) = "����"
    aryFree(3, 0) = "���������ַ�": aryFree(3, 1) = "�����@�A�B"
    aryFree(4, 0) = "����ϣ����": aryFree(4, 1) = "���������������������������������������������������¦æĦŦƦǦȦɦʦ˦̦ͦΦϦЦѦҦӦԦզ֦צ�"
    aryFree(5, 0) = "�������": aryFree(5, 1) = "�������������������������������������������������������������������ѧҧӧԧէ֧ا٧ڧۧܧݧާߧ�������������������"
    aryFree(6, 0) = "������": aryFree(6, 1) = "�\�C���D�����������E������F��"
    aryFree(7, 0) = "���ҷ���": aryFree(7, 1) = "�"
    aryFree(8, 0) = "������ĸ�ķ���": aryFree(8, 1) = "��G�H��Y"
    aryFree(9, 0) = "������ʽ": aryFree(9, 1) = "�����������������������������������������"
    aryFree(10, 0) = "��ͷ": aryFree(10, 1) = "���������I�J�K�L"
    aryFree(11, 0) = "��ѧ�����": aryFree(11, 1) = "�ʡǡƨM�̡ءިN�ϨO�Ρġšɡȡҡӡ�ߡáˡס֡ըP�١ԡܡݨR�ڡۨ��ѡͨS"
    aryFree(12, 0) = "���Ӽ����÷���": aryFree(12, 1) = "��"
    aryFree(13, 0) = "�����ŵ���ĸ����": aryFree(13, 1) = "�٢ڢۢܢݢޢߢ���ŢƢǢȢɢʢˢ̢͢΢ϢТѢҢӢԢբ֢עآ����������������������������������¢â�"
    aryFree(14, 0) = "�Ʊ��": aryFree(14, 1) = "�������������������������������������������������������������©éĩũƩǩȩɩʩ˩̩ͩΩϩЩѩҩөԩթ֩שة٩ک۩ܩݩީߩ����������������T�U�V�W�X�Y�Z�[�\�]�^�_�`�a�b�c�d�e�f�g�h�i�j�k�l�m�n�o�p�q�r�s�t�u�v�w"
    aryFree(15, 0) = "����Ԫ��": aryFree(15, 1) = "�x�y�z�{�|�}�~����������������������"
    aryFree(16, 0) = "����ͼ�η�": aryFree(16, 1) = "������������������񨍨�����"
    aryFree(17, 0) = "���Ӷ�����(ʾ�����)": aryFree(17, 1) = "�����"
    aryFree(18, 0) = "CJK���źͱ��": aryFree(18, 1) = "���������e���������������������������������������@�A�B�C�D�E�F�G�H"
    aryFree(19, 0) = "ƽ����": aryFree(19, 1) = "�������������������������������������������������������������������¤äĤŤƤǤȤɤʤˤ̤ͤΤϤФѤҤӤԤդ֤פؤ٤ڤۤܤݤޤߤ��������������������a�b�f�g"
    aryFree(20, 0) = "Ƭ����": aryFree(20, 1) = "�������������������������������������������������������������������¥åĥťƥǥȥɥʥ˥̥ͥΥϥХѥҥӥԥե֥ץ٥ڥۥܥݥޥߥ��������������������������`�c�d"
    aryFree(21, 0) = "ע��": aryFree(21, 1) = "�ŨƨǨȨɨʨ˨̨ͨΨϨШѨҨӨԨը֨רب٨ڨۨܨݨިߨ����������"
    aryFree(22, 0) = "�����ŵ�CJK��ĸ���·�": aryFree(22, 1) = "�����������Z�I"
    aryFree(23, 0) = "CJK�����ַ�": aryFree(23, 1) = "�J�K�L�M�N�O�P�Q�R�S�T"
    aryFree(24, 0) = "CJK������ʽ": aryFree(24, 1) = "�U����������������������h�i�j�k�l�m�n"
    aryFree(25, 0) = "Сд����": aryFree(25, 1) = "�o�p�q�r�s�t�u�v�w�x�y�z�{�|�}�~������������������"
    aryFree(26, 0) = "���м�ȫ���ַ�": aryFree(26, 1) = "��" & Chr(-23646) & "���磥���������������������������������������������������������£ãģţƣǣȣɣʣˣ̣ͣΣϣУѣңӣԣգ֣ףأ٣ڣۣܣݣޣߣ��������������������������������������������V���W��"
    aryFree(27, 0) = "�����ַ�": aryFree(27, 1) = "�ͪϪѪӪժת٪۪ݪߪ�������������������������������������������������ëǫɫ˫ͫϫѫӫի׫٫ݫ߫�������������������������������������������������ìŬǬɬˬͬϬѬӪ����������������������˪�����ū۫񬩬������������ªĪƪȪʪΪЪҪԪ֪تڪܪު�������������������������������������������������«īȫʫ̫ΫЫҫԫ֫ثګޫ�������������������������������������������������¬ĬƬȬʬ̬άЬҬ�"

    Dim intRow As Integer, intCol As Integer
    With Me.mfgFree
        For i = 0 To .Cols - 1
            .ColWidth(i) = 420
            .ColAlignment(i) = 4
        Next
        .ROWHEIGHT(0) = (.Height - 90) / 5
    End With
    
    intRow = 0: intCol = 0
    cboGroup.Clear
    For i = 0 To UBound(aryFree) - 1
        Me.cboGroup.AddItem aryFree(i, 0)
        Me.cboGroup.ItemData(Me.cboGroup.NewIndex) = Len(aryFree(i, 1))
        For j = 0 To Len(aryFree(i, 1)) - 1
            Me.mfgFree.TextMatrix(intRow, intCol) = Mid(aryFree(i, 1), j + 1, 1)
            intCol = intCol + 1
            If intCol = Me.mfgFree.Cols Then
                intRow = intRow + 1: intCol = 0
                If intRow >= Me.mfgFree.Rows - 1 Then
                    Me.mfgFree.Rows = Me.mfgFree.Rows + 1
                    Me.mfgFree.ROWHEIGHT(Me.mfgFree.Rows - 1) = Me.mfgFree.ROWHEIGHT(0)
                End If
            End If
        Next
    Next
    Me.cboGroup.ListIndex = 0
    
    Set CommandBars.Icons = zlCommFun.GetPubIcons
    'Ĭ��ѡ���ϴιر�ʱѡ�е�ҳ��
    i = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & "frmDockSymbol", "Selection", 2)
    Call lblType_Click(i)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    CommandBars.ActiveMenuBar.Visible = False
    picData.Move 0, picPre.Top + IIf(picPre.Visible, picPre.Height, 0), 100, Me.ScaleHeight - picTitle.Height - IIf(picPre.Visible, picPre.Height, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & "frmDockSymbol", "Selection", Val(shpSearch.Tag)
    Set mobjLis = Nothing
    Set mOldLisRs = Nothing
    Set mNewLisRs = Nothing
    img16.ListImages.Clear
    imgList.ListImages.Clear
    ImageList_Destroy img16.hImageList
    ImageList_Destroy imgList.hImageList
    Set picData.Picture = Nothing
    Set picFormat.Picture = Nothing

End Sub

Private Sub lblPot_Click(Index As Integer)
    If Index >= 4 Then
        lblPot(0) = "��": lblPot(1) = "��": lblPot(2) = "��": lblPot(3) = "��"
    Else
        lblPot(4) = "��": lblPot(5) = "��": lblPot(6) = "��": lblPot(7) = "��"
    End If
    
    If lblPot(Index).Caption = "��" Then
       lblPot(Index).Caption = "��"
    Else
       lblPot(Index).Caption = "��"
    End If
    
    If picSpot.Visible Then
        Call MakeSpotPic
    End If
End Sub
Public Property Get PicFontSize() As Long
    PicFontSize = mlFontSize
End Property

'################################
    '   ����DOCK�����б�
'################################
Public Function FillEPRDemos() As Long
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "Select l.Id, l.���, l.����,zlspellcode(l.����) As ����, Nvl(l.����,'δ����') as ����,l.˵��, l.ͨ�ü�" & vbNewLine & _
            "From ��������Ŀ¼ l, Table(Cast(f_Segment_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) u" & vbNewLine & _
            "Where l.�ļ�id = [1] And Nvl(l.����, 0) = [5] And l.Id = To_Number(u.����)"
        gstrSQL = gstrSQL & " And" & vbNewLine & _
            "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
            "      L.ͨ�ü� = 1 And" & vbNewLine & _
            "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User) Or" & vbNewLine & _
            "      L.ͨ�ü� = 2 And L.��Աid In (Select U.��Աid From �ϻ���Ա�� U Where U.�û��� = User)) order by l.ͨ�ü� desc, l.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmDockSymbol", mlngFileID, mlngPatient, mlngVisit, mlngAdvice, 0)
        'If rsTemp.EOF Then Exit Function
        Err = 0: On Error GoTo errHand
         With vsList
                .Visible = True
                .Clear
                .ToolTipText = ""
                .Rows = rsTemp.RecordCount + 1
                .Cols = 6
                .FixedRows = 1
                .FixedCols = 0
                .SelectionMode = flexSelectionByRow
                .ColAlignment(1) = flexAlignLeftCenter
                .TextMatrix(0, 0) = "���"
                .TextMatrix(0, 1) = "����"
                .TextMatrix(0, 2) = "��Χ"
                .TextMatrix(0, 3) = "ID"
                .TextMatrix(0, 4) = "����"
                .TextMatrix(0, 5) = "˵��"
                .ColWidth(0) = 600
                .ColWidth(1) = 2100
                .ColWidth(2) = 500
                .ColWidth(3) = 0
                .ColWidth(4) = 0
                .ColWidth(5) = 0
                .FontSize = 10
            'ѭ����ӵ�VsGridView��
            Do While Not rsTemp.EOF
                '0-ȫԺͨ��;1-����ͨ��;2-����ʹ��
                    .Cell(flexcpPicture, rsTemp.AbsolutePosition, 2) = imgList.ListImages(Val(rsTemp("ͨ�ü�").Value) + 1).Picture
                    .TextMatrix(rsTemp.AbsolutePosition, 3) = NVL(rsTemp("ID").Value)
                    .TextMatrix(rsTemp.AbsolutePosition, 0) = NVL(rsTemp("���").Value)
                    .TextMatrix(rsTemp.AbsolutePosition, 1) = NVL(rsTemp("����").Value)
                    .TextMatrix(rsTemp.AbsolutePosition, 4) = NVL(rsTemp("����").Value)
                    .TextMatrix(rsTemp.AbsolutePosition, 5) = NVL(rsTemp("˵��").Value)
                    .ROWHEIGHT(rsTemp.AbsolutePosition) = 300
                rsTemp.MoveNext
            Loop
        End With
    
    FillEPRDemos = rsTemp.RecordCount
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    FillEPRDemos = rsTemp.RecordCount
End Function
'################################
    '   ���ع���ҩ���б�
'################################
Public Function FillAllergyDrugs()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "select I.����,zlspellcode(I.����) as ���� from ������ĿĿ¼ I,�����÷����� Z " & _
    "Where i.ID = Z.��ĿID " & _
    "and Z.����=0 " & _
    "and I.��� in ('5', '6') " & _
    "and (I.����ʱ�� is null or I.����ʱ�� = to_date('3000-01-01', 'YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmDockSymbol")
    Err = 0: On Error GoTo errHand
    With Me.vsList
        .Clear
        .ToolTipText = ""
        .Cols = 2
        .SelectionMode = flexSelectionByRow
        .FixedRows = 0: .FixedCols = 0
        .RowHeightMin = 250
        .Rows = rsTemp.RecordCount
        .ColWidth(0) = 3000
        .ColWidth(1) = 0
        .ColAlignment(0) = flexAlignLeftCenter
        
        Do Until rsTemp.EOF
            .TextMatrix(rsTemp.AbsolutePosition - 1, 0) = rsTemp!����
            .TextMatrix(rsTemp.AbsolutePosition - 1, 1) = rsTemp!����
            .Cell(flexcpFontSize, rsTemp.AbsolutePosition - 1, 0) = 10
            .ROWHEIGHT(rsTemp.AbsolutePosition - 1) = 300
             rsTemp.MoveNext
        Loop
        If .Visible Then .SetFocus
    End With
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub InitLisvfg()
    With vsList
        .Clear
        .Tag = ""
        .ToolTipText = "˫����Ҫ������"
        .Cols = 9
        .Rows = 1
        .FixedRows = 1
        .FontSize = 10
        .MergeCells = flexMergeFree
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortAndMove
        .OutlineBar = flexOutlineBarCompleteLeaf
        .OutlineCol = 0
        .TextMatrix(0, mCol.ѡ��) = "ѡ"
        .TextMatrix(0, mCol.ָ��) = "ָ��"
        .TextMatrix(0, mCol.���) = "���"
        .TextMatrix(0, mCol.��λ) = "��λ"
        .TextMatrix(0, mCol.��־) = ""
        .TextMatrix(0, mCol.�ο�) = "�ο�"
        .ColWidth(mCol.���) = 200
        .ColWidth(mCol.ѡ��) = 300
        .ColWidth(mCol.ָ��) = 1800
        .ColWidth(mCol.���) = 600
        .ColWidth(mCol.��λ) = 600
        .ColWidth(mCol.��־) = 300
        .ColWidth(mCol.�ο�) = 1200
        .ColWidth(mCol.������Դ) = 0
        .ColWidth(mCol.���ʱ��) = 0
    End With
End Sub
Private Sub InitLisItem()
'��ʼ����񼰾����
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    If mlngPatient = 0 Then Exit Sub
    
    gstrSQL = "Select ��Դ,����ID,����ID,����ʱ��" & vbNewLine & _
            "From (" & vbNewLine & _
            "Select 2 ��Դ,����ID,to_char(��ҳID) ����ID,��Ժ���� ����ʱ�� from ������ҳ where ����ID=[1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select 1 ��Դ,����ID,NO ����ID,�Ǽ�ʱ�� ����ʱ�� from ���˹Һż�¼ where ����ID=[1] And ��¼����=1 and ��¼״̬=1)" & vbNewLine & _
            "Order by ����ʱ�� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˾����", mlngPatient)
    With cboTimes
        .Clear
        Do Until rsTemp.EOF
            If rsTemp!��Դ = 2 Then
                .AddItem "��" & rsTemp!����ID & "��סԺ  " & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & Space(200) & rsTemp!��Դ & "|" & rsTemp!����ID & "|" & rsTemp!����ID
            Else
                .AddItem "�������  " & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & Space(200) & rsTemp!��Դ & "|" & rsTemp!����ID & "|" & rsTemp!����ID
            End If
            rsTemp.MoveNext
        Loop
        If .ListCount > 0 Then
            .ListIndex = 0 '����cboTimes_Click �Ӷ����� FillLisItem
        Else
            Call FilterLisItem
        End If
    End With

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub FillLisItem()
'��ȡ���˼���ָ���¼
Dim rsTemp As New ADODB.Recordset, lngPatientID As Long, strPageId As String, intType As Integer, strAdvices As String
    On Error GoTo errHand
    Set mNewLisRs = Nothing
    Set mOldLisRs = Nothing
    
    intType = Split(Split(cboTimes.Text, Space(200))(1), "|")(0)
    lngPatientID = Split(Split(cboTimes.Text, Space(200))(1), "|")(1)
    strPageId = Split(Split(cboTimes.Text, Space(200))(1), "|")(2)
    
    If intType <> 1 Then
        '��ȡӤ����¼����ʾĸӤѡ��
        gstrSQL = "select ���,decode(Ӥ������,null,'Ӥ��'||���,Ӥ������)||' ����' ���� from ������������¼ where ����id = [1] And ��ҳid = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmDockSymbol", lngPatientID, CLng(strPageId))
        If rsTemp.EOF Then
            chkCY(0).Visible = False: chkCY(1).Visible = False
        Else
            chkCY(0).Visible = True: chkCY(1).Visible = True
        End If
    End If
    
    '�°�LIS
    If intType = 1 Then '�������
        gstrSQL = "Select Distinct ���id" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ������ B" & vbNewLine & _
                "Where a.����id = [1] And a.�Һŵ� = [2] And a.������� = 'C' And a.Id = b.ҽ��id And b.ִ��״̬ = 1 And Not Exists" & vbNewLine & _
                    "(Select 1 From ������Ŀ�ֲ� Where ҽ��id = a.Id)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����ҽ��", lngPatientID, strPageId)
    Else
        gstrSQL = "Select Distinct ���id" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ������ B" & vbNewLine & _
                "Where a.����id = [1] And a.��ҳid = [2] And a.������� = 'C' And a.Id = b.ҽ��id And b.ִ��״̬ = 1 And Not Exists" & vbNewLine & _
                    "(Select 1 From ������Ŀ�ֲ� Where ҽ��id = a.Id)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����ҽ��", lngPatientID, CLng(strPageId))
    End If
    Do Until rsTemp.EOF
        strAdvices = strAdvices & "," & rsTemp!���ID
        rsTemp.MoveNext
    Loop
    If strAdvices <> "" Then
        strAdvices = Mid(strAdvices, 2)
        Set rsTemp = GetLisItems(strAdvices)
        If Not rsTemp Is Nothing Then
            Set mNewLisRs = rsTemp
        End If
    End If
    
'    '�ϰ�LIS
'    If intType = 1 Then
'        gstrSQL = "Select  g.���� ҽ������, c.������ As ������Ŀ, d.��д, b.������, d.��λ, Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
'                " Trim(Replace(Replace(' ' || Zlgetreference(b.������Ŀid, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������, a.����id, a.����)," & vbNewLine & _
'                "                       ' .', '0.'), '��.', '��0.')) As �ο�, Decode(a.������Դ, 1, '����', 2, 'סԺ', 4, '���', '����') ������Դ, a.���ʱ��,0 Ӥ��" & vbNewLine & _
'                "From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ������Ŀ D, ������Ŀ�ֲ� F, ����ҽ����¼ E, ������ĿĿ¼ G, ����ҽ����¼ H" & vbNewLine & _
'                "Where a.����id = [1] And a.�Һŵ� = [2] And a.����id = e.����id And a.����� Is Not Null And a.Id = b.����걾id And" & vbNewLine & _
'                "      b.������Ŀid = c.Id And c.Id = d.������Ŀid And b.��¼���� = a.������ And a.Id = f.�걾id And f.��Ŀid = d.������Ŀid And f.ҽ��id = e.Id And h.���id = e.Id And" & vbNewLine & _
'                "      g.Id = h.������Ŀid" & vbNewLine & _
'                "Order By a.���ʱ�� Desc, e.ҽ������, b.�������, c.������"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��������", lngPatientID, strPageId)
'    Else
'        gstrSQL = "Select g.���� ҽ������, c.������ As ������Ŀ, d.��д, b.������, d.��λ," & vbNewLine & _
'                    "       Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
'                    "       Trim(Replace(Replace(' ' ||" & vbNewLine & _
'                    "                             Zlgetreference(b.������Ŀid, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������, a.����id, a.����), ' .'," & vbNewLine & _
'                    "                             '0.'), '��.', '��0.')) As �ο�, Decode(a.������Դ, 1, '����', 2, 'סԺ', 4, '���', '����') ������Դ, a.���ʱ��," & vbNewLine & _
'                    "       Nvl(a.Ӥ��, 0) Ӥ��" & vbNewLine & _
'                    "From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ������Ŀ D, ������Ŀ�ֲ� F, ����ҽ����¼ E, ������ĿĿ¼ G, ����ҽ����¼ H" & vbNewLine & _
'                    "Where a.����id = [1] And a.��ҳid = [2] And a.����id = e.����id And a.����� Is Not Null And a.Id = b.����걾id And b.������Ŀid = c.Id And" & vbNewLine & _
'                    "      c.Id = d.������Ŀid And b.��¼���� = a.������ And a.Id = f.�걾id And f.��Ŀid = d.������Ŀid And f.ҽ��id = e.Id And h.���id = e.Id And" & vbNewLine & _
'                    "      g.Id = h.������Ŀid" & vbNewLine & _
'                    "Order By a.���ʱ�� Desc, e.ҽ������, b.�������, c.������"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��������", lngPatientID, CLng(strPageId))
'    End If

    If intType = 1 Then
        gstrSQL = "Select Nvl(c.����, '�ֹ���Ŀ') ҽ������, e.������ As ������Ŀ, d.��д, b.������, d.��λ," & vbNewLine & _
                    "       Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
                    "       Trim(Replace(Replace(' ' ||" & vbNewLine & _
                    "                             Zlgetreference(b.������Ŀid, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������, a.����id, a.����), ' .'," & vbNewLine & _
                    "                             '0.'), '��.', '��0.')) As �ο�, Decode(a.������Դ, 1, '����', 2, 'סԺ', 4, '���', '����') ������Դ, a.���ʱ��," & vbNewLine & _
                    "       Nvl(a.Ӥ��, 0) Ӥ��" & vbNewLine & _
                    "From ����걾��¼ A, ������ͨ��� B, ������Ŀ D, ����������Ŀ E, ������ĿĿ¼ C" & vbNewLine & _
                    "Where a.����id = [1] And a.�Һŵ� = [2] And a.����� Is Not Null And a.Id = b.����걾id And b.������Ŀid = c.Id(+) And" & vbNewLine & _
                    "      b.������Ŀid = d.������Ŀid And b.������Ŀid = e.Id" & vbNewLine & _
                    "Order By a.���ʱ�� Desc, c.����, b.�������, e.������"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��������", lngPatientID, strPageId)
    Else
        gstrSQL = "Select Nvl(c.����, '�ֹ���Ŀ') ҽ������, e.������ As ������Ŀ, d.��д, b.������, d.��λ," & vbNewLine & _
                    "       Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
                    "       Trim(Replace(Replace(' ' ||" & vbNewLine & _
                    "                             Zlgetreference(b.������Ŀid, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������, a.����id, a.����), ' .'," & vbNewLine & _
                    "                             '0.'), '��.', '��0.')) As �ο�, Decode(a.������Դ, 1, '����', 2, 'סԺ', 4, '���', '����') ������Դ, a.���ʱ��," & vbNewLine & _
                    "       Nvl(a.Ӥ��, 0) Ӥ��" & vbNewLine & _
                    "From ����걾��¼ A, ������ͨ��� B, ������Ŀ D, ����������Ŀ E, ������ĿĿ¼ C" & vbNewLine & _
                    "Where a.����id = [1] And a.��ҳid = [2] And a.����� Is Not Null And a.Id = b.����걾id And b.������Ŀid = c.Id(+) And" & vbNewLine & _
                    "      b.������Ŀid = d.������Ŀid And b.������Ŀid = e.Id" & vbNewLine & _
                    "Order By a.���ʱ�� Desc, c.����, b.�������, e.������"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��������", lngPatientID, CLng(strPageId))
    End If
    
    Set mOldLisRs = rsTemp
    Call FilterLisItem
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function GetPhase() As String
Dim dCurTime As Date, sDate As Date, eDate As Date, Result As String
    dCurTime = zlDatabase.Currentdate
    Select Case True
        Case optPhase(0).Value '����
            sDate = dCurTime - Weekday(dCurTime, vbMonday) + 1
            eDate = dCurTime
            Result = "���ʱ��>=#" & Format(sDate, "yyyy-mm-dd 00:00:00") & "# And ���ʱ��<=#" & Format(eDate, "yyyy-mm-dd 23:59:59") & "#"
        Case optPhase(1).Value '����
            sDate = dCurTime - Weekday(dCurTime, vbMonday) + 1 - 7
            eDate = dCurTime - Weekday(dCurTime, vbMonday)
            Result = "���ʱ��>=#" & Format(sDate, "yyyy-mm-dd 00:00:00") & "# And ���ʱ��<=#" & Format(eDate, "yyyy-mm-dd 23:59:59") & "#"
        Case optPhase(2).Value '����
            sDate = dCurTime
            eDate = dCurTime
            Result = "���ʱ��>=#" & Format(sDate, "yyyy-mm-01 00:00:00") & "# And ���ʱ��<=#" & Format(eDate, "yyyy-mm-dd 23:59:59") & "#"
        Case optPhase(3).Value '����
            Result = "���ʱ��<#" & Format(dCurTime, "yyyy-mm-01 00:00:00") & "#"
    End Select
    
    If chkCY(0).Visible Then
        If chkCY(0).Value = vbChecked And chkCY(1).Value = vbUnchecked Then 'ֻѡ��"ĸ"
            Result = Result & " And Ӥ��=0"
        ElseIf chkCY(0).Value = vbUnchecked And chkCY(1).Value = vbChecked Then 'ֻѡ��"Ӥ"
            Result = Result & " And Ӥ��<>0"
        Else '��ѡ��
            Result = Result
        End If
    End If
    
    GetPhase = Result
End Function

Private Sub optPhase_Click(Index As Integer)
    FilterLisItem
End Sub
Private Sub FilterLisItem()
    InitLisvfg
    If Not mNewLisRs Is Nothing Then
        AddListItem mNewLisRs
    End If
    
    If Not mOldLisRs Is Nothing Then
        AddListItem mOldLisRs
    End If
End Sub
Private Sub AddListItem(ByVal rsItems As ADODB.Recordset)
Dim strGroup As String, strTmpG As String, strAdvice As String
    With vsList
        If rsItems Is Nothing Then Exit Sub
        If rsItems.State = adStateClosed Then Exit Sub
        rsItems.Filter = 0
        If rsItems.RecordCount = 0 Then Exit Sub
        
        rsItems.Filter = GetPhase()
        Do Until rsItems.EOF
            '�����ʽ ������Ŀ(���ʱ��)
            strTmpG = rsItems!ҽ������ & "(" & Format(rsItems!���ʱ��, "yyyy-MM-dd hh:mm") & ")" & IIf(rsItems!Ӥ�� = 0, "", "Ӥ" & rsItems!Ӥ��)
            '�ж��Ƿ����µķ��࣬���������ӷ���
            If strGroup <> strTmpG Then
                strGroup = strTmpG
                .AddItem ""
                .TextMatrix(.Rows - 1, mCol.ָ��) = strGroup
                .Cell(flexcpData, .Rows - 1, mCol.ָ��) = NVL(rsItems!ҽ������)
                .Cell(flexcpData, .Rows - 1, mCol.���) = Format(rsItems!���ʱ��, "yyyy-MM-dd hh:mm")
                .TextMatrix(.Rows - 1, mCol.���) = strGroup
                .TextMatrix(.Rows - 1, mCol.��λ) = strGroup
                .TextMatrix(.Rows - 1, mCol.��־) = strGroup
                .TextMatrix(.Rows - 1, mCol.�ο�) = strGroup
                .IsSubtotal(.Rows - 1) = True    '������ʾ
                .RowOutlineLevel(.Rows - 1) = 0     '�ڵ�
                .MergeRow(.Rows - 1) = True
            End If
            
            .AddItem ""
            .TextMatrix(.Rows - 1, mCol.ָ��) = NVL(rsItems!������Ŀ) & "(" & NVL(rsItems!��д) & ")"
            .Cell(flexcpData, .Rows - 1, mCol.ָ��) = NVL(rsItems!������Ŀ) & "|" & NVL(rsItems!��д)
            .TextMatrix(.Rows - 1, mCol.���) = NVL(rsItems!������)
            .TextMatrix(.Rows - 1, mCol.��λ) = Replace(NVL(rsItems!��λ), "��", "u")
            .TextMatrix(.Rows - 1, mCol.��־) = NVL(rsItems!��־)
            Select Case rsItems!��־
                Case "��"
                    .Cell(flexcpBackColor, .Rows - 1, mCol.��־, .Rows - 1, mCol.��־) = &H80FFFF
                Case "��"
                    .Cell(flexcpBackColor, .Rows - 1, mCol.��־, .Rows - 1, mCol.��־) = &H80C0FF
                Case "����", "����"
                    .Cell(flexcpBackColor, .Rows - 1, mCol.��־, .Rows - 1, mCol.��־) = &H40C0&
            End Select
            .TextMatrix(.Rows - 1, mCol.�ο�) = NVL(rsItems!�ο�)
            .IsSubtotal(.Rows - 1) = True   '������ʾ
            .RowOutlineLevel(.Rows - 1) = 1 '�ӽڵ�
            rsItems.MoveNext
        Loop
        
        If .Rows > 1 Then
            Dim i As Integer
            For i = 1 To .Rows - 1
                .GetNode(i).Expanded = False
            Next
            .Cell(flexcpPictureAlignment, 1, mCol.ѡ��, .Rows - 1, mCol.ѡ��) = flexPicAlignCenterCenter
            .Cell(flexcpAlignment, 1, mCol.ָ��, .Rows - 1, mCol.���ʱ��) = flexAlignLeftCenter
            .TopRow = 1
        End If
    End With
End Sub
Public Property Let PicFontSize(vData As Long)
    mlFontSize = vData
    picPre.Height = picFormat.Height + 200
    Call Form_Resize
End Property
Private Sub lblType_Click(Index As Integer)
    Dim strTemp As String, i As Integer, intRow As Integer, intCol As Integer
    Dim rsTemp As ADODB.Recordset
        On Error Resume Next
        If Index = Val(Me.shpSearch.Tag) Then Exit Sub
        picHY.Visible = False
        picRY.Visible = False
        picFree.Visible = False
        picSpot.Visible = False
        vsList.Visible = False
        picYJS.Visible = False
        picFormat.Visible = False
        picPre.Visible = False
        cmdInsert.Enabled = False
        fraSplit.Visible = False
        shpBorder.Visible = True
        shpBorder.Move lblType(Index).Left - Screen.TwipsPerPixelX, lblType(Index).Top - Screen.TwipsPerPixelX, lblType(Index).Width + Screen.TwipsPerPixelX * 2, lblType(Index).Height + Screen.TwipsPerPixelX * 2
        lblSearch.Visible = False
        txtSearch.Visible = False
        shpSearch.Visible = False
        vsList.Visible = False
        cboTimes.Visible = False
        chkRem.Visible = False
        chkref.Visible = False
        optFormat(0).Visible = False
        optFormat(1).Visible = False
        chkLanguage(0).Visible = False
        chkLanguage(1).Visible = False
        picPhase.Visible = False
        optPhase(0).Visible = False: optPhase(1).Visible = False: optPhase(2).Visible = False: optPhase(3).Visible = False
        chkCY(0).Visible = False: chkCY(1).Visible = False
        vsList.FixedRows = 0: vsList.FixedCols = 0
        vsList.MergeCells = flexMergeNever: vsList.ToolTipText = ""
        shpSearch.Tag = Index
        For i = 0 To lblType.UBound
            If i = Index Then
                lblType(i).FontBold = True
            Else
                lblType(i).FontBold = False
            End If
        Next
        
        Select Case lblType(Index).Caption
            Case "������"
                strTemp = CON������
            Case "��λ����"
                strTemp = CON��λ����
            Case "�������"
                strTemp = CON�������
            Case "��ѧ����"
                strTemp = CON��ѧ����
            Case "�������"
                strTemp = CON������� + CONҽѧ����
            Case "ҽѧ��λ"
                strTemp = CONҽѧ��λ
            Case "����ҩ��"
                strTemp = ""
                Call FillAllergyDrugs
            Case "���ĵ���"
                strTemp = ""
                Call FillEPRDemos
            Case "������ע"
                picHY.Visible = True
                picPre.Visible = True
                fraSplit.Visible = True
                picFormat.Visible = True
            Case "������ע"
                picRY.Visible = True
                picPre.Visible = True
                fraSplit.Visible = True
                picFormat.Visible = True
            Case "�¾�ʷ"
                picYJS.Visible = True
                picPre.Visible = True
                fraSplit.Visible = True
                picFormat.Visible = True
                txtYJ(0).SetFocus
            Case "̥��λ��"
                picPre.Visible = True
                fraSplit.Visible = True
                picFormat.Visible = True
                picSpot.Visible = True
            Case "����ѡ��"
                picFree.Visible = True
                mfgFree.SetFocus
            Case "������"
                picPre.Height = cmdInsert.Height + 350 + chkRem.Height + cboTimes.Height + optPhase(0).Height
                picPre.Visible = True
                cboTimes.Visible = True
                cmdInsert.Enabled = True
                fraSplit.Visible = True
                vsList.Visible = True
                chkRem.Visible = True
                chkref.Visible = True
                optFormat(0).Visible = True
                optFormat(1).Visible = True
                chkLanguage(0).Visible = True
                chkLanguage(1).Visible = True
                optPhase(0).Visible = True: optPhase(1).Visible = True: optPhase(2).Visible = True: optPhase(3).Visible = True
                picPhase.Visible = True
        End Select
        
        Select Case lblType(Index).Caption
            Case "������", "��λ����", "�������", "��ѧ����", "�������"
                vsList.Visible = True
                With vsList
                    .Clear
                    .FixedRows = 0
                    .Cols = 8
                    .SelectionMode = flexSelectionFree
                    .Rows = Len(strTemp) \ .Cols + 1
                    .Row = 0
                    .Col = 0
                    For i = 0 To Len(strTemp) - 1
                        intRow = i \ .Cols: intCol = i Mod .Cols
                        .TextMatrix(intRow, intCol) = Mid(strTemp, i + 1, 1)
                    Next
        
                    For i = 0 To .Rows - 1
                        .ROWHEIGHT(i) = 420
                    Next
                    For i = 0 To .Cols - 1
                        .ColAlignment(i) = 4
                        .ColWidth(i) = 420
                        .FontSize = 12
                    Next
                    If .Visible Then .SetFocus
                End With
            Case "ҽѧ��λ"
                 'Wordҽѧ��������
                vsList.Visible = True
                With vsList
                    .Clear
                    .FixedRows = 0
                    .Cols = 3
                    .SelectionMode = flexSelectionFree
                    .Rows = (UBound(Split(strTemp, ",")) + 1) \ .Cols + 1
                    
                        For i = 0 To UBound(Split(strTemp, ","))
                            intRow = i \ .Cols: intCol = i Mod .Cols
                            .TextMatrix(intRow, intCol) = Replace(Split(strTemp, ",")(i), "��", "u") '��ֹ���ַ����֣����±������滻Ϊu
                            .Cell(flexcpFontSize, intRow, intCol) = 10
                        Next
                    For i = 0 To .Rows - 1
                        .ROWHEIGHT(i) = 420
                    Next
                    For i = 0 To .Cols - 1
                        .ColAlignment(i) = 4
                        .ColWidth(i) = 1000
                    Next
                    If .Visible Then .SetFocus
                End With
            Case "����ҩ��", "���ĵ���"
                    vsList.Visible = True
                    txtSearch.Visible = True
                    txtSearch.ToolTipText = "������ƴ��������Ķ�λ,��λ�ɹ���س�����"
                    txtSearch.Text = ""
                    lblSearch.Visible = True
                    shpSearch.Visible = True
                    txtSearch.SetFocus
            Case "������"
                Call InitLisItem
        End Select
        Call Form_Resize
        Call picData_Resize
End Sub

Private Sub lblType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To lblType.UBound
        If Index = i Then
            lblType(i).FontUnderline = True
            lblType(i).ForeColor = RGB(0, 0, 128)
        Else
            lblType(i).FontUnderline = False
            lblType(i).ForeColor = RGB(0, 0, 0)
        End If
    Next
End Sub
Private Sub mfgFree_DblClick()
    If mfgFree.TextMatrix(mfgFree.Row, mfgFree.Col) <> "" Then
        RaiseEvent InsertSymbol(mfgFree.TextMatrix(mfgFree.Row, mfgFree.Col), 1)
    End If
End Sub

Private Sub mfgFree_RowColChange()
Dim intPoint As Integer, intStart As Integer
Dim i As Integer, j As Integer
    With Me.mfgFree
        intPoint = .Cols * .Row + .Col + 1
    End With
    intStart = 0
    For i = 0 To Me.cboGroup.ListCount - 1
        intStart = intStart + Me.cboGroup.ItemData(i)
        If intPoint <= intStart Then Me.cboGroup.ListIndex = i: Exit Sub
    Next
End Sub

Private Sub mshHY_Click()
    If mshHY.CellBackColor = vbWhite Then
        mshHY.CellBackColor = M_FLAGCOLOR
        mshHY.CellFontBold = True
        mshHY.CellFontUnderline = True
        mshHY.CellForeColor = vbBlue
    Else
        mshHY.CellBackColor = vbWhite
        mshHY.CellFontBold = False
        mshHY.CellFontUnderline = False
        mshHY.CellForeColor = mshHY.ForeColor
    End If
    Call MakeToothString(mshHY, 8)
    Call MakeToothPic(mshHY, 8)
End Sub

Private Sub mshRY_Click()
    If mshRY.CellBackColor = vbWhite Then
        mshRY.CellBackColor = M_FLAGCOLOR
        mshRY.CellFontBold = True
        mshRY.CellFontUnderline = True
        mshRY.CellForeColor = vbBlue
    Else
        mshRY.CellBackColor = vbWhite
        mshRY.CellFontBold = False
        mshRY.CellFontUnderline = False
        mshRY.CellForeColor = mshRY.ForeColor
    End If
    Call MakeToothString(mshRY, 5)
    Call MakeToothPic(mshRY, 5)
End Sub


Private Sub picData_Resize()
On Error Resume Next
    picHY.Move 0, 100, picData.Width, picData.Height
    picRY.Move 0, 100, picData.Width, picData.Height
    picRY.Move 0, 100, picData.Width, picData.Height
    picFree.Move 0, 100, picData.Width, picData.Height
    picSpot.Move 0, 100, picData.Width, picData.Height
    picYJS.Move 0, 100, picData.Width, picData.Height
    If lblType(Val(shpSearch.Tag)).Caption = "����ҩ��" Or lblType(Val(shpSearch.Tag)).Caption = "���ĵ���" Then
        lblSearch.Move 100, 50
        txtSearch.Move lblSearch.Width + 100, lblSearch.Top - 30
        shpSearch.Move txtSearch.Left - Screen.TwipsPerPixelX, txtSearch.Top - Screen.TwipsPerPixelY, txtSearch.Width + Screen.TwipsPerPixelX * 2, txtSearch.Height + Screen.TwipsPerPixelY * 2
        vsList.Move 0, txtSearch.Top + txtSearch.Height + 50, picData.Width, picData.Height - 500
    Else
        vsList.Move 0, 0, picData.Width, picData.Height - 100
    End If
    fraSplit.Move -15, IIf(picPre.Visible, picPre.Height - 30, -15), Me.ScaleWidth
    
    If lblType(Val(shpSearch.Tag)).Caption = "������" Then
        cmdInsert.Move 100, 50
        optFormat(0).Move cmdInsert.Left + cmdInsert.Width + 100, cmdInsert.Top + 50
        optFormat(1).Move optFormat(0).Left + optFormat(0).Width, cmdInsert.Top + 50
        chkLanguage(0).Move optFormat(0).Left + 10, cmdInsert.Top + cmdInsert.Height + 50
        chkLanguage(1).Move chkLanguage(0).Left - 10 + chkLanguage(0).Width + 100, cmdInsert.Top + cmdInsert.Height + 50
        chkCY(0).Move optFormat(1).Left + 10, cmdInsert.Top + cmdInsert.Height + 50
        chkCY(1).Move chkCY(0).Left + chkCY(0).Width + 80, cmdInsert.Top + cmdInsert.Height + 50
        chkRem.Move cmdInsert.Left, cmdInsert.Top + cmdInsert.Height + 50
        chkref.Move chkRem.Left + chkRem.Width + 10, cmdInsert.Height + 100
        cboTimes.Move cmdInsert.Left, chkRem.Top + chkRem.Height + 50
        picPhase.Move cmdInsert.Left, cboTimes.Top + cboTimes.Height + 50
    Else
        cmdInsert.Move 100, 100
    End If
End Sub
Private Sub picFree_Resize()
    On Error Resume Next
    cboGroup.Move lblGroup.Left + lblGroup.Width, lblGroup.Top
    mfgFree.Move 0, cboGroup.Top + cboGroup.Height + 50, picFree.Width, picFree.Height - cboGroup.Height - 50
    vsList.Move 0, cboGroup.Top + cboGroup.Height + 300, picFree.Width, picFree.Height - cboGroup.Height - 50
End Sub

Private Sub picHY_DblClick()
    If cmdInsert.Enabled Then Call cmdInsert_Click
End Sub

Private Sub picHY_Resize()
Dim i As Integer
    On Error Resume Next
    '������ע
    mshHY.Rows = 2: mshHY.Cols = 16
    mshHY.Height = mshHY.RowHeightMin * mshHY.Rows - 30
    mshHY.Width = 210 * mshHY.Cols + 30
    mshHY.Left = (mshHY.Container.Width - mshHY.Width) / 2
    For i = 0 To mshHY.Cols - 1
        mshHY.ColWidth(i) = 210
        mshHY.ColAlignment(i) = 4
        If i + 1 <= 8 Then
            mshHY.TextMatrix(0, i) = 8 - ((i + 1) Mod 9) + 1
            mshHY.TextMatrix(1, i) = 8 - ((i + 1) Mod 9) + 1
        Else
            mshHY.TextMatrix(0, i) = (i - 7) Mod 9
            mshHY.TextMatrix(1, i) = (i - 7) Mod 9
        End If
    Next
    fraLineHYH.Move mshHY.Left, mshHY.Top + (mshHY.Height - fraLineHYH.Height) / 2, mshHY.Width
    
    fraLineHYV.Left = mshHY.Left + mshHY.ColWidth(0) * (mshHY.Cols / 2)
    
    For i = 0 To 7
        lblHY(i).Left = fraLineHYV.Left + (mshHY.ColWidth(0) - lblHY(i).Width) / 2 + mshHY.ColWidth(0) * i
    Next
    
    lblHYUp.Move fraLineHYV.Left - lblHYUp.Width / 2, fraLineHYV.Top - lblHYUp.Height - 30
    lblHYDn.Move lblHYUp.Left, mshHY.Top + mshHY.Height + 60
    
    lblHYLeft.Move mshHY.Left, lblHYUp.Top
    lblHYRight.Move mshHY.Left + mshHY.Width - lblHYRight.Width, lblHYUp.Top
End Sub

Private Sub picRY_DblClick()
    If cmdInsert.Enabled Then Call cmdInsert_Click
End Sub

Private Sub picRY_Resize()
Dim i As Integer
    On Error Resume Next '������ע
    mshRY.Rows = 2: mshRY.Cols = 10
    mshRY.Height = mshRY.RowHeightMin * mshRY.Rows - 30
    mshRY.Width = 350 * mshRY.Cols - 60
    mshRY.Left = (mshRY.Container.Width - mshRY.Width) / 2
    
    mshRY.TextMatrix(0, 0) = "��"
    mshRY.TextMatrix(0, 1) = "��"
    mshRY.TextMatrix(0, 2) = "��"
    mshRY.TextMatrix(0, 3) = "��"
    mshRY.TextMatrix(0, 4) = "��"
    For i = 0 To mshRY.Cols - 1
        mshRY.ColWidth(i) = 350
        mshRY.ColAlignment(i) = 4
        
        If i >= 5 Then mshRY.TextMatrix(0, i) = mshRY.TextMatrix(0, mshRY.Cols - i - 1)
        mshRY.TextMatrix(1, i) = mshRY.TextMatrix(0, i)
    Next
    
    fraLineRYH.Move mshRY.Left, mshRY.Top + (mshRY.Height - fraLineRYH.Height) / 2, mshRY.Width
    fraLineRYV.Move mshRY.Left + mshRY.ColWidth(0) * (mshRY.Cols / 2)
    
    For i = 0 To 4
        lblRY(i).Left = fraLineRYV.Left + (mshRY.ColWidth(0) - lblRY(i).Width) / 2 + mshRY.ColWidth(0) * i
    Next
    
    lblRYUp.Move fraLineRYV.Left - lblRYUp.Width / 2, fraLineRYV.Top - lblRYUp.Height - 30
    lblRYDn.Move lblRYUp.Left, mshRY.Top + mshRY.Height + 60
    lblRYLeft.Move mshRY.Left, lblRYUp.Top
    lblRYRight.Move mshRY.Left + mshRY.Width - lblRYRight.Width, lblRYUp.Top
End Sub
Private Sub picSpot_DblClick()
If cmdInsert.Enabled Then Call cmdInsert_Click
End Sub

Private Sub picTitle_Resize()
    On Error Resume Next
    fraType(0).Move -15, fraType(0).Top, ScaleWidth
    fraType(1).Move -15, fraType(1).Top, ScaleWidth
    fraType(2).Move -15, fraType(2).Top, ScaleWidth
    fraType(3).Move -15, fraType(3).Top, ScaleWidth
End Sub
Private Function MakeToothPic(objMSH As MSHFlexGrid, bytCount As Byte) As StdPicture
'���ܣ����ݺ�����ע��������ʾ������ע��ͼƬ
'��ʽΪ������|���ݡ��¾�ʷ 1|ǰ�|����|��ĸ|���|�ֺ�; ���� 2(����)/3(����)|����|����|����|����|�ֺ�; ̥��λ�� 4|�Ϸ�|�·�|��|�ҷ�|�ֺ�
Dim intRow As Integer, intCol As Integer, i As Integer
Dim a As String, b As String, C As String, D As String 'A=����,B=����,C=����,D=����

    '��ABCD�ĸ�����ı�ע���,�����Ŀ�ʼ��ݺ�,��"37"
    RaiseEvent GetPosFontSize
    objMSH.Redraw = False
    intRow = objMSH.Row: intCol = objMSH.Col
    
    objMSH.Row = 0
    For i = 0 To bytCount - 1
        objMSH.Col = i
        If objMSH.CellBackColor = M_FLAGCOLOR Then a = a & objMSH.TextMatrix(0, i)
    Next
    For i = bytCount To bytCount * 2 - 1
        objMSH.Col = i
        If objMSH.CellBackColor = M_FLAGCOLOR Then b = b & objMSH.TextMatrix(0, i)
    Next
    
    objMSH.Row = 1
    For i = 0 To bytCount - 1
        objMSH.Col = i
        If objMSH.CellBackColor = M_FLAGCOLOR Then C = C & objMSH.TextMatrix(1, i)
    Next
    For i = bytCount To bytCount * 2 - 1
        objMSH.Col = i
        If objMSH.CellBackColor = M_FLAGCOLOR Then D = D & objMSH.TextMatrix(1, i)
    Next
    
    objMSH.Row = intRow: objMSH.Col = intCol
    objMSH.Redraw = True
    
    '���ݲ�ͬ�ĸ��������������ע
Dim r As RECT, pt As POINTAPI
Dim lAW As Long, lBW As Long, lCW As Long, lDW As Long
Dim lAH As Long, lBH As Long, lCH As Long, lDH As Long
    On Error Resume Next
    
    Set picFormat.Picture = Nothing: picFormat.Cls: picFormat.Width = "2400"
    picFormat.Font.Size = 8: picFormat.Refresh
    If a = "" And b = "" And C = "" And D = "" Then cmdInsert.Enabled = False: Exit Function
    '����������
    lAW = picFormat.TextWidth(a):   lAH = picFormat.TextHeight(a):      lBW = picFormat.TextWidth(b):       lBH = picFormat.TextHeight(b)
    lCW = picFormat.TextWidth(C):   lCH = picFormat.TextHeight(C):      lDW = picFormat.TextWidth(D):       lDH = picFormat.TextHeight(D)
    
    If a <> "" And b = "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        picFormat.Width = picFormat.ScaleX(lAW + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lAH + 1, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Bottom = r.Top + lAH: r.Left = 2: r.Right = r.Left + lAW
        DrawTextEx picFormat.hDC, a, -1, r, DT_CENTER, ByVal 0&         'д��
        MoveToEx picFormat.hDC, 4, lAH, pt  '����
        LineTo picFormat.hDC, lAW + 4, lAH
        MoveToEx picFormat.hDC, lAW + 4, 2, pt  '����
        LineTo picFormat.hDC, lAW + 4, lAH
    ElseIf a = "" And b <> "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        picFormat.Width = picFormat.ScaleX(lBW + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lBH + 1, vbPixels, vbTwips)
        picFormat.Refresh: picFormat.AutoRedraw = True
        
        r.Top = 0: r.Bottom = r.Top + lBH: r.Left = 5: r.Right = r.Left + lBW
        DrawTextEx picFormat.hDC, b, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, lBH, pt
        LineTo picFormat.hDC, lBW + 5, lBH
        MoveToEx picFormat.hDC, 2, 2, pt
        LineTo picFormat.hDC, 2, lBH
    ElseIf a = "" And b = "" And C <> "" And D = "" Then
        'ֻ�����±�ע
        picFormat.Width = picFormat.ScaleX(lCW + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lCH, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 2: r.Bottom = r.Top + lCH: r.Left = 2: r.Right = r.Left + lCW
        DrawTextEx picFormat.hDC, C, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, lCW + 5, 1
        MoveToEx picFormat.hDC, lCW + 4, 1, pt
        LineTo picFormat.hDC, lCW + 4, lCH + 4
    ElseIf a = "" And b = "" And C = "" And D <> "" Then
        'ֻ�����±�ע
        picFormat.Width = picFormat.ScaleX(lDW + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lDH, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 2: r.Bottom = r.Top + lDH: r.Left = 5: r.Right = r.Left + lDW
        DrawTextEx picFormat.hDC, D, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, lDW + 5, 1
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, 2, lDH + 4
    ElseIf a <> "" And b <> "" And C = "" And D = "" Then
        'ֻ���������б�ע
        picFormat.Width = picFormat.ScaleX(lAW + lBW + 9, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lAH + 1, vbPixels, vbTwips)
        picFormat.Refresh
        
         r.Bottom = r.Top + lAH: r.Left = 2: r.Right = r.Left + lAW
        DrawTextEx picFormat.hDC, a, -1, r, DT_CENTER, ByVal 0&  'д��
         r.Bottom = r.Top + lAH: r.Left = r.Right + 5: r.Right = r.Left + lBW
        DrawTextEx picFormat.hDC, b, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, lAH, pt
        LineTo picFormat.hDC, lAW + lBW + 7, lAH
        MoveToEx picFormat.hDC, lAW + 4, 2, pt
        LineTo picFormat.hDC, lAW + 4, lAH
    ElseIf a = "" And b = "" And C <> "" And D <> "" Then
        'ֻ���������б�ע
        picFormat.Width = picFormat.ScaleX(lCW + lDW + 9, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lCH + 1, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 2: r.Bottom = r.Top + lCH: r.Left = 2: r.Right = r.Left + lCW
        DrawTextEx picFormat.hDC, C, -1, r, DT_CENTER, ByVal 0&
        r.Top = 2: r.Bottom = r.Top + lCH: r.Left = r.Right + 5: r.Right = r.Left + lDW
        DrawTextEx picFormat.hDC, D, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, lCW + lDW + 7, 1
        MoveToEx picFormat.hDC, lCW + 4, 2, pt
        LineTo picFormat.hDC, lCW + 4, lCH + 3
    ElseIf a <> "" And b = "" And C <> "" And D = "" Then
        'ֻ�����������б�ע
        picFormat.Width = picFormat.ScaleX(IIf(lAW > lCW, lAW, lCW) + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lAH + lCH - 2, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 0: r.Bottom = r.Top + lAH: r.Left = 2: r.Right = r.Left + IIf(lAW > lCW, lAW, lCW)
        DrawTextEx picFormat.hDC, a, -1, r, DT_CENTER, ByVal 0&
        r.Top = r.Bottom: r.Bottom = r.Top + lCH: r.Left = 2: r.Right = r.Left + IIf(lAW > lCW, lAW, lCW)
        DrawTextEx picFormat.hDC, C, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, lAH - 1, pt
        LineTo picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, lAH - 1
        MoveToEx picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, 2, pt
        LineTo picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, lAH + lCH + 7
    ElseIf a = "" And b <> "" And C = "" And D <> "" Then
        'ֻ�����������б�ע
        picFormat.Width = picFormat.ScaleX(IIf(lBW > lDW, lBW, lDW) + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lBH + lDH - 2, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 0: r.Bottom = r.Top + lBH: r.Left = 3: r.Right = r.Left + IIf(lBW > lDW, lBW, lDW)
        DrawTextEx picFormat.hDC, b, -1, r, DT_CENTER, ByVal 0&
        r.Top = r.Bottom: r.Bottom = r.Top + lDH: r.Left = 3: r.Right = r.Left + IIf(lBW > lDW, lBW, lDW)
        DrawTextEx picFormat.hDC, D, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 3, lBH - 1, pt
        LineTo picFormat.hDC, IIf(lBW > lDW, lBW, lDW) + 4, lBH - 1
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, 2, lAH + lCH + 6
    Else
        '�������Ҷ��б�ע
        picFormat.Width = picFormat.ScaleX(IIf(lAW > lCW, lAW, lCW) + IIf(lBW > lDW, lBW, lDW) + 9, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(IIf(lAH > lBH, lAH, lBH) + IIf(lCH > lDH, lCH, lDH) - 2, vbPixels, vbTwips)
        picFormat.Refresh
        
        If a <> "" Then
            r.Bottom = lAH: r.Left = 2: r.Right = r.Left + IIf(lAW > lCW, lAW, lCW)
            DrawTextEx picFormat.hDC, a, -1, r, DT_CENTER, ByVal 0&
        End If
        If b <> "" Then
          r.Bottom = r.Top + lBH: r.Left = IIf(lAW > lCW, lAW, lCW) + 7: r.Right = r.Left + IIf(lBW > lDW, lBW, lDW)
            DrawTextEx picFormat.hDC, b, -1, r, DT_CENTER, ByVal 0&
        End If
        If C <> "" Then
            r.Top = IIf(lAH > lBH, lAH, lBH): r.Bottom = r.Top + lCH: r.Left = 2: r.Right = r.Left + IIf(lAW > lCW, lAW, lCW)
            DrawTextEx picFormat.hDC, C, -1, r, DT_CENTER, ByVal 0&
        End If
        If D <> "" Then
            r.Top = IIf(lAH > lBH, lAH, lBH): r.Bottom = r.Top + lDH: r.Left = IIf(lAW > lCW, lAW, lCW) + 7: r.Right = r.Left + IIf(lBW > lDW, lBW, lDW)
            DrawTextEx picFormat.hDC, D, -1, r, DT_CENTER, ByVal 0&
        End If
        
        MoveToEx picFormat.hDC, 2, IIf(lAH > lBH, lAH, lBH) - 1, pt
        LineTo picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + IIf(lBW > lDW, lBW, lDW) + 7, IIf(lAH > lBH, lAH, lBH) - 1
        MoveToEx picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, 2, pt
        LineTo picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, IIf(lAH > lBH, lAH, lBH) + IIf(lCH > lDH, lCH, lDH)
    End If
    cmdInsert.Enabled = True
    PicFontSize = mlFontSize '�����������ͼƬλ��
    Set picFormat.Picture = picFormat.Image
    mstrInfor = IIf(bytCount = 8, 2, 3) & "|" & a & "|" & b & "|" & C & "|" & D & "|" & mlFontSize
End Function



Private Function MakeToothString(objMSH As MSHFlexGrid, bytCount As Byte) As String
    '���ܣ����ݺ�����ע��������ʾ������ע�������ַ�����
    '������objMSH=������������ע���
    '      bytCount=����������
Dim byt���� As Byte, byt��ĸ As Byte, strTemp As String
Dim intRow As Integer, intCol As Integer
Dim i As Integer, j As Integer
Dim a As String, b As String, C As String, D As String 'A=����,B=����,C=����,D=����
Dim YC���� As String
Dim YCС���� As String, YCС��ĸ As String
Dim YC����� As String, YC���ĸ As String
Dim YC����� As String, YC���ĸ As String
Dim YC�ҷ��� As String, YC�ҷ�ĸ As String
        
    strTemp = ""
    If objMSH.Name = "mshHY" Then
        YC���� = HY����
        YCС���� = HYС����: YCС��ĸ = HYС��ĸ
        YC����� = HY�����: YC���ĸ = HY���ĸ
        YC����� = HY�����: YC���ĸ = HY���ĸ
        YC�ҷ��� = HY�ҷ���: YC�ҷ�ĸ = HY�ҷ�ĸ
    Else
        YC���� = RY����
        YCС���� = RYС����: YCС��ĸ = RYС��ĸ
        YC����� = RY�����: YC���ĸ = RY���ĸ
        YC����� = RY�����: YC���ĸ = RY���ĸ
        YC�ҷ��� = RY�ҷ���: YC�ҷ�ĸ = RY�ҷ�ĸ
    End If
            
    '��ABCD�ĸ�����ı�ע���,�����Ŀ�ʼ��ݺ�,��"37"
    objMSH.Redraw = False
    intRow = objMSH.Row: intCol = objMSH.Col
    
    objMSH.Row = 0
    For i = bytCount To 1 Step -1
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then a = a & bytCount + 1 - i
    Next
    For i = bytCount + 1 To bytCount * 2
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then b = b & i - bytCount
    Next
    
    objMSH.Row = 1
    For i = bytCount To 1 Step -1
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then C = C & bytCount + 1 - i
    Next
    For i = bytCount + 1 To bytCount * 2
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then D = D & i - bytCount
    Next
    
    objMSH.Row = intRow: objMSH.Col = intCol
    objMSH.Redraw = True
    
    '���ݲ�ͬ�ĸ��������������ע�����ַ���
    If a <> "" And b = "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        For i = Len(a) To 1 Step -1
            If i = 1 Then
                strTemp = strTemp & Mid(YC�����, CByte(Mid(a, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC�����, CByte(Mid(a, i, 1)), 1)
            End If
        Next
    ElseIf a = "" And b <> "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        For i = 1 To Len(b)
            If i = 1 Then
                strTemp = strTemp & Mid(YC�ҷ���, CByte(Mid(b, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC�����, CByte(Mid(b, i, 1)), 1)
            End If
        Next
    ElseIf a = "" And b = "" And C <> "" And D = "" Then
        'ֻ�����±�ע
        For i = Len(C) To 1 Step -1
            If i = 1 Then
                strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(C, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(C, i, 1)), 1)
            End If
        Next
    ElseIf a = "" And b = "" And C = "" And D <> "" Then
        'ֻ�����±�ע
        For i = 1 To Len(D)
            If i = 1 Then
                strTemp = strTemp & Mid(YC�ҷ�ĸ, CByte(Mid(D, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(D, i, 1)), 1)
            End If
        Next
    ElseIf a <> "" And b <> "" And C = "" And D = "" Then
        'ֻ���������б�ע
        For i = Len(a) To 1 Step -1
            strTemp = strTemp & Mid(YC�����, CByte(Mid(a, i, 1)), 1)
        Next
        strTemp = strTemp & "��"
        For i = 1 To Len(b)
            strTemp = strTemp & Mid(YC�����, CByte(Mid(b, i, 1)), 1)
        Next
    ElseIf a = "" And b = "" And C <> "" And D <> "" Then
        'ֻ���������б�ע
        For i = Len(C) To 1 Step -1
            strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(C, i, 1)), 1)
        Next
        strTemp = strTemp & "��"
        For i = 1 To Len(D)
            strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(D, i, 1)), 1)
        Next
    ElseIf a <> "" And b = "" And C = "" And D <> "" Then
        'ֻ�����������б�ע
        For i = Len(a) To 1 Step -1
            strTemp = strTemp & Mid(YCС����, CByte(Mid(a, i, 1)), 1)
        Next
        strTemp = strTemp & "��"
        For i = 1 To Len(D)
            strTemp = strTemp & Mid(YCС��ĸ, CByte(Mid(D, i, 1)), 1)
        Next
    ElseIf a = "" And b <> "" And C <> "" And D = "" Then
        'ֻ�����������б�ע
        For i = Len(C) To 1 Step -1
            strTemp = strTemp & Mid(YCС��ĸ, CByte(Mid(C, i, 1)), 1)
        Next
        strTemp = strTemp & "��"
        For i = 1 To Len(b)
            strTemp = strTemp & Mid(YCС����, CByte(Mid(b, i, 1)), 1)
        Next
    ElseIf Not (a = "" And b = "" And C = "" And D = "") Then
        '���¶��б�ע
        If a = "" And C = "" Then strTemp = "��"
        
        '����߷�����
        i = 1: j = 1 'i��ӦA,j��ӦC
        Do While i <= Len(a) Or j <= Len(C)
            byt���� = 0: byt��ĸ = 0
            If i <= Len(a) Then byt���� = Mid(a, i, 1)
            If j <= Len(C) Then byt��ĸ = Mid(C, j, 1)
            '���ݷ��ӷ�ĸ��һ�������������
            If byt���� <> 0 And byt��ĸ <> 0 Then
                strTemp = strTemp & Mid(YC����, (byt��ĸ - 1) * bytCount + byt����, 1)
            ElseIf byt���� <> 0 And byt��ĸ = 0 Then
                strTemp = strTemp & Mid(YCС����, byt����, 1)
            ElseIf byt���� = 0 And byt��ĸ <> 0 Then
                strTemp = strTemp & Mid(YCС��ĸ, byt��ĸ, 1)
            End If
            i = i + 1: j = j + 1
        Loop
        strTemp = StrReverse(strTemp)
        
        '���ӷ�
        If (a <> "" Or C <> "") And (b <> "" Or D <> "") Then
            strTemp = strTemp & "��"
        ElseIf b = "" And D = "" Then
            strTemp = strTemp & "��"
        End If
        
        '���ұ߷�����
        i = 1: j = 1 'i��ӦB,j��ӦD
        Do While i <= Len(b) Or j <= Len(D)
            byt���� = 0: byt��ĸ = 0
            If i <= Len(b) Then byt���� = Mid(b, i, 1)
            If j <= Len(D) Then byt��ĸ = Mid(D, j, 1)
            '���ݷ��ӷ�ĸ��һ�������������
            If byt���� <> 0 And byt��ĸ <> 0 Then
                strTemp = strTemp & Mid(YC����, (byt��ĸ - 1) * bytCount + byt����, 1)
            ElseIf byt���� <> 0 And byt��ĸ = 0 Then
                strTemp = strTemp & Mid(YCС����, byt����, 1)
            ElseIf byt���� = 0 And byt��ĸ <> 0 Then
                strTemp = strTemp & Mid(YCС��ĸ, byt��ĸ, 1)
            End If
            i = i + 1: j = j + 1
        Loop
    End If
    picFormat.Tag = strTemp
    MakeToothString = strTemp
End Function
Public Function HideSomeThing(ByVal bType As Byte)
    If bType = 1 Then '����ʱ�����¾�ʷ
        Dim i As Integer
        For i = 0 To lblType.UBound
            If lblType(i).Caption = "�¾�ʷ" Then lblType(i).Visible = False
            If lblType(i).Caption = "̥��λ��" Then lblType(i).Visible = False
        Next
    End If
End Function

Private Sub picYJS_DblClick()
    If cmdInsert.Enabled Then Call cmdInsert_Click
End Sub

Private Sub txtSearch_Change()
    Dim i As Integer, colName As Integer, colSpell As Integer
    txtSearch.Tag = ""
    If txtSearch.Text = "" Then Exit Sub
    Select Case lblType(Val(shpSearch.Tag)).Caption
            Case "���ĵ���"
                colName = 1: colSpell = 4
            Case "����ҩ��"
                colName = 0: colSpell = 1
    End Select
    With vsList
        For i = 0 To .Rows - 1
            If InStr(.TextMatrix(i, colSpell), UCase(txtSearch.Text)) > 0 Or InStr(.TextMatrix(i, colName), Trim(txtSearch.Text)) > 0 Or InStr(.TextMatrix(i, 0), Trim(txtSearch.Text)) > 0 Then
                .Row = i
                .TopRow = i
                txtSearch.Tag = "Selected " & i
                Exit Sub
            End If
        Next
        .Row = -1
    End With
End Sub


Private Sub txtSearch_GotFocus()
  RaiseEvent SetFouse
End Sub
Private Sub txtYJ_GotFocus(Index As Integer)
   RaiseEvent SetFouse
End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtSearch.Tag <> "" Then
        vsList_DblClick
    End If
End Sub

Private Sub txtYJ_Change(Index As Integer)
    If Visible Then
        Call MakeYJString
        Call MakeYJPic
    End If
End Sub
Private Sub txtYJ_DblClick(Index As Integer)
    txtYJ_Change Index
End Sub
Private Sub txtYJ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        txtYJ(Index + 1).SetFocus
    End If
End Sub

Private Sub txtYJ_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("|',", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Function MakeYJPic() As StdPicture
'��ʽΪ������|���ݡ��¾�ʷ 1|ǰ�|����|��ĸ|���|�ֺ�; ���� 2(����)/3(����)|����|����|����|����|�ֺ�; ̥��λ�� 4|�Ϸ�|�·�|��|�ҷ�|�ֺ�
Dim strB As String, strU As String, strD As String, strA As String, r As RECT, lPW As Long, lPH As Long, pt As POINTAPI
Dim lBW As Long, lBH As Long, lUW As Long, lUH As Long, lDW As Long, lDH As Long, lAW As Long, lAH As Long
    RaiseEvent GetPosFontSize
    mstrInfor = ""
    strB = txtYJ(0).Text:   strU = txtYJ(1).Text:   strD = txtYJ(2).Text:   strA = txtYJ(3).Text
    If strB <> "" And strU <> "" And strD <> "" And strA <> "" And lblType(13).Visible Then
        cmdInsert.Enabled = True
    Else
        cmdInsert.Enabled = False
    End If
    
    Set picFormat.Picture = Nothing:                picFormat.Cls: picFormat.Width = "2400"
    picFormat.FontSize = 8:        picFormat.Refresh
    
    
    lBW = picFormat.TextWidth(strB): lBH = picFormat.TextHeight(strB): lUW = picFormat.TextWidth(strU): lUH = picFormat.TextHeight(strU)
    lDW = picFormat.TextWidth(strD): lDH = picFormat.TextHeight(strB): lAW = picFormat.TextWidth(strA): lAH = picFormat.TextHeight(strA)
    lPW = lBW + IIf(lUW > lDW, lUW, lDW) + lAW + 8
    lPH = IIf(lBH > 0, lBH, IIf(lUH > 0, lUH, IIf(lDH > 0, lDH, IIf(lAH > 0, lAH, 30)))) * 2 - 5
    picFormat.Width = picFormat.ScaleX(lPW, vbPixels, vbTwips)
    picFormat.Height = picFormat.ScaleY(lPH, vbPixels, vbTwips)
    picFormat.Refresh
    
    If strB <> "" Then
        r.Top = (lPH - lBH) / 2: r.Bottom = r.Top + lBH: r.Left = 2: r.Right = r.Left + lBW
        DrawTextEx picFormat.hDC, strB, -1, r, DT_CENTER, ByVal 0&
    End If
    
    If strU <> "" Then
        r.Top = -1: r.Bottom = r.Top + lUH: r.Left = lBW + 4: r.Right = r.Left + IIf(lUW > lDW, lUW, lDW)
        DrawTextEx picFormat.hDC, strU, -1, r, DT_CENTER, ByVal 0&
    End If
    
    If strD <> "" Then
        r.Top = IIf(lUH > lDH, lUH, lDH) - 3: r.Bottom = r.Top + lDH: r.Left = lBW + 4: r.Right = r.Left + IIf(lUW > lDW, lUW, lDW)
        DrawTextEx picFormat.hDC, strD, -1, r, DT_CENTER, ByVal 0&
    End If
    
    If strA <> "" Then
        r.Top = (lPH - lAH) / 2: r.Bottom = r.Top + lAH: r.Left = lBW + IIf(lUW > lDW, lUW, lDW) + 7: r.Right = r.Left + lAW
        DrawTextEx picFormat.hDC, strA, -1, r, DT_CENTER, ByVal 0&
    End If
    
    MoveToEx picFormat.hDC, lBW + 2, (lPH) / 2, pt
    LineTo picFormat.hDC, lBW + IIf(lUW > lDW, lUW, lDW) + 6, (lPH) / 2
    
    Set picFormat.Picture = picFormat.Image
    mstrInfor = "1|" & strB & "|" & strU & "|" & strD & "|" & strA & "|" & mlFontSize
End Function

Private Function MakeYJString() As String
'���ܣ������¾�ʷ��д���������������ַ���ע��
    Dim str���� As String, str��ĸ As String
    Dim strTmp As String
    
    
    '��������֣��������Ҷ���
    '------------------------
    str���� = Right(Format(Int(Val(txtYJ(1).Text)), "00"), 2)
    str��ĸ = Right(Format(Int(Val(txtYJ(2).Text)), "00"), 2)
    
    '��10λ���ַ�
    If Val(Left(str��ĸ, 1)) <> 0 Or Val(Left(str����, 1)) <> 0 Then
        If Val(Left(str��ĸ, 1)) <> 0 And Val(Left(str����, 1)) <> 0 Then
            strTmp = Mid(YJ����1, (Val(Left(str��ĸ, 1)) - 1) * 10 + Val(Left(str����, 1)) + 1, 1)
        ElseIf Val(Left(str����, 1)) = 0 Then
            strTmp = Mid(YJ��ĸ, Val(Left(str��ĸ, 1)) + 1, 1)
        ElseIf Val(Left(str��ĸ, 1)) = 0 Then
            strTmp = Mid(YJ����, Val(Left(str����, 1)) + 1, 1)
        End If
    End If
        
    '���λ���ַ�
    strTmp = strTmp & Mid(YJ����2, Val(Right(str��ĸ, 1)) * 10 + Val(Right(str����, 1)) + 1, 1)
        
    '��������ַ�
    strTmp = txtYJ(0).Text & strTmp
    strTmp = strTmp & txtYJ(3).Text
    picFormat.Tag = strTmp
    MakeYJString = strTmp
End Function
Private Function MakeSpotPic() As StdPicture
'�� ��
'���ܣ�����ѡ������̥��λ��ͼƬ,��������Ӧ��Ϣ
'��ʽΪ������|���ݡ��¾�ʷ 1|ǰ�|����|��ĸ|���|�ֺ�; ���� 2(����)/3(����)|����|����|����|����|�ֺ�; ̥��λ�� 4|�Ϸ�|�·�|��|�ҷ�|�ֺ�
Dim lPW As Long, lPH As Long, r As RECT, pt As POINTAPI, intType As Integer, lsw As Long, lsh As Long
    RaiseEvent GetPosFontSize
    mstrInfor = ""
    Set picFormat.Picture = Nothing:                picFormat.Cls: picFormat.Width = "2400"
    picFormat.FontSize = 8:      picFormat.Refresh
    lsw = picFormat.TextWidth("��"): lsh = picFormat.TextHeight("��")
    If lblPot(0) = "��" Or lblPot(1) = "��" Or lblPot(2) = "��" Or lblPot(3) = "��" And lblType(11).Visible Then
        lPW = lsw * 2 + 3
        lPH = lsh * 2
        intType = 1
        cmdInsert.Enabled = True
    ElseIf lblPot(4) = "��" Or lblPot(5) = "��" Or lblPot(6) = "��" Or lblPot(7) = "��" And lblType(11).Visible Then
        lPW = lsw * 3 - 8
        lPH = lsh * 3 - 10
        intType = 2
        cmdInsert.Enabled = True
    Else
        cmdInsert.Enabled = False
        Exit Function
    End If
    picFormat.Width = picFormat.ScaleX(lPW, vbPixels, vbTwips)
    picFormat.Height = picFormat.ScaleY(lPH, vbPixels, vbTwips)
    picFormat.Refresh
    
Dim ba As Byte, bb As Byte, bc As Byte, bd As Byte, be As Byte, bf As Byte, bg As Byte, bh As Byte
    If lblPot(0) = "��" Then
        r.Top = 0: r.Bottom = r.Top + lsh: r.Left = 1: r.Right = r.Left + lsw: ba = 1
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(1) = "��" Then
        r.Top = 0: r.Bottom = r.Top + lsh: r.Left = lsw + 4: r.Right = r.Left + lsw: bb = 1
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(2) = "��" Then
        r.Top = lsh: r.Bottom = r.Top + lsh: r.Left = 1: r.Right = r.Left + lsw: bc = 1
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(3) = "��" Then
        r.Top = lsh: r.Bottom = r.Top + lsh: r.Left = lsw + 4: r.Right = r.Left + lsw: bd = 1
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(4) = "��" Then
        r.Top = -1: r.Bottom = r.Top + lsh: r.Left = lsw - 4: r.Right = r.Left + lsw: be = 2
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(5) = "��" Then
        r.Top = lPH - lsh + 2: r.Bottom = r.Top + lsh: r.Left = lsw - 3: r.Right = r.Left + lsw: bf = 2
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    If lblPot(6) = "��" Then
        r.Top = lsh - 4: r.Bottom = r.Top + lsh: r.Left = -1: r.Right = r.Left + lsw: bg = 2
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    If lblPot(7) = "��" Then
        r.Top = lsh - 4: r.Bottom = r.Top + lsh: r.Left = lPW - lsw + 2: r.Right = r.Left + lsw: bh = 2
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If intType = 1 Then
        MoveToEx picFormat.hDC, 0, lsh - 1, pt
        LineTo picFormat.hDC, lPW - 1, lsh - 1
        MoveToEx picFormat.hDC, lsw + 2, 0, pt
        LineTo picFormat.hDC, lsw + 2, lPH - 1
        mstrInfor = "4|" & ba & "|" & bb & "|" & bc & "|" & bd & "|" & mlFontSize
    ElseIf intType = 2 Then
        MoveToEx picFormat.hDC, 1, 2, pt
        LineTo picFormat.hDC, lPW - 1, lPH - 1
        MoveToEx picFormat.hDC, 1, lPH - 1, pt
        LineTo picFormat.hDC, lPW - 1, 1
        mstrInfor = "4|" & be & "|" & bf & "|" & bg & "|" & bh & "|" & mlFontSize
    End If
    picFormat.Tag = ""
    Set picFormat.Picture = picFormat.Image
End Function


Private Sub vsList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If lblType(Val(shpSearch.Tag)).Caption = "������" Then
        If Col = mCol.ѡ�� Or Col = mCol.��� Then Cancel = True
    End If
End Sub

Private Sub vsList_Click()
    If vsList.Row < 0 Then Exit Sub
    Select Case lblType(Val(shpSearch.Tag)).Caption
        Case "���ĵ���"
            Me.vsList.ToolTipText = Me.vsList.TextMatrix(vsList.Row, 5)
        Case "����ҩ��"
            Me.vsList.ToolTipText = Me.vsList.TextMatrix(vsList.Row, 0)
        Case "������"
            If vsList.MouseCol = mCol.ѡ�� Or vsList.MouseCol = mCol.��־ Then
                Call vsList_KeyDown(32, 0)
            End If
        Case Else
            Me.vsList.ToolTipText = ""
    End Select
    
End Sub

Private Sub vsList_DblClick()
     If vsList.Row < 0 Then Exit Sub
     Select Case lblType(Val(shpSearch.Tag)).Caption
        Case "���ĵ���"
            If Val(vsList.TextMatrix(vsList.Row, 3)) > 0 And vsList.Row > 0 Then
                RaiseEvent InsertEPRDemo(Val(vsList.TextMatrix(vsList.Row, 3)))
            End If
        Case "����ҩ��", "ҽѧ��λ"
            If vsList.TextMatrix(vsList.Row, vsList.Col) <> "" Then
                RaiseEvent InsertSymbol(vsList.TextMatrix(vsList.Row, vsList.Col), Len(vsList.TextMatrix(vsList.Row, vsList.Col)))
            End If
        Case "������", "��λ����", "�������", "��ѧ����", "�������"
            If vsList.TextMatrix(vsList.Row, vsList.Col) <> "" Then
                RaiseEvent InsertSymbol(vsList.TextMatrix(vsList.Row, vsList.Col), 1)
            End If
        Case "������"
            Call vsList_KeyDown(32, 0)
    End Select
End Sub
Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer, lsel As Long
    Select Case Control.ID
        Case 100
            If lblType(Val(shpSearch.Tag)).Caption = "������" Then
                cmdInsert_Click
            Else
                vsList_DblClick
            End If
        Case 101 'ѡ���쳣
            With vsList
                If .Row < 1 Then
                    MsgBox "����ѡ����Ҫ�ļ������ݣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                If .RowOutlineLevel(.Row) <> 0 Then
                '����ѭ�������ڵ�
                    For i = .Row To 0 Step -1
                        If .RowOutlineLevel(i) = 0 Then
                            .Row = i: Exit For
                        End If
                    Next
                End If
                
                For i = .Row To .Rows - 1
                    If .RowOutlineLevel(i) = 0 And i <> .Row Then Exit For '��һ��ҽ��
                    If i = .Row Then
                        lsel = .Cell(flexcpData, .Row, mCol.��־)
                        .Cell(flexcpData, i, mCol.��־) = IIf(lsel = 0, 1, 0)
                    Else
                        Set .Cell(flexcpPicture, i, mCol.ѡ��) = IIf(lsel = 0 And .TextMatrix(i, mCol.��־) <> "", img16.ListImages("Selected").Picture, Nothing)
                        .Cell(flexcpData, i, mCol.ѡ��) = IIf(lsel = 0 And .TextMatrix(i, mCol.��־) <> "", 1, 0)
                    End If
                Next
            End With
        Case 102 'ѡ������
            With vsList
                If .Row < 1 Then
                    MsgBox "����ѡ����Ҫ�ļ������ݣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                If .RowOutlineLevel(.Row) <> 0 Then
                '����ѭ�������ڵ�
                    For i = .Row To 0 Step -1
                        If .RowOutlineLevel(i) = 0 Then
                            .Row = i: Exit For
                        End If
                    Next
                End If
                
                For i = .Row To .Rows - 1
                    If .RowOutlineLevel(i) = 0 And i <> .Row Then Exit For '��һ��ҽ��
                    If i = .Row Then
                        lsel = .Cell(flexcpData, .Row, mCol.ѡ��)
                        .Cell(flexcpData, i, mCol.ѡ��) = IIf(lsel = 0, 1, 0)
                    Else
                        Set .Cell(flexcpPicture, i, mCol.ѡ��) = IIf(lsel = 0, img16.ListImages("Selected").Picture, Nothing)
                        .Cell(flexcpData, i, mCol.ѡ��) = IIf(lsel = 0, 1, 0)
                    End If
                Next
            End With
    End Select
End Sub

Private Sub vsList_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblType(Val(shpSearch.Tag)).Caption <> "������" Then Exit Sub
    With vsList
    If KeyCode = 32 Then
        Dim i As Integer, lsel As Long
        If .Row < 1 Then Exit Sub
        If .RowOutlineLevel(.Row) = 0 Then
            If .MouseCol = mCol.��־ Then
                If .GetNode(.Row).Expanded Then
                    For i = .Row To .Rows - 1
                        If .RowOutlineLevel(i) = 0 And i <> .Row Then Exit For '��һ��ҽ��
                        If i = .Row Then
                            lsel = .Cell(flexcpData, .Row, mCol.��־)
                            .Cell(flexcpData, i, mCol.��־) = IIf(lsel = 0, 1, 0)
                        Else
                            Set .Cell(flexcpPicture, i, mCol.ѡ��) = IIf(lsel = 0 And .TextMatrix(i, mCol.��־) <> "", img16.ListImages("Selected").Picture, Nothing)
                            .Cell(flexcpData, i, mCol.ѡ��) = IIf(lsel = 0 And .TextMatrix(i, mCol.��־) <> "", 1, 0)
                        End If
                    Next
                Else
                    .GetNode(.Row).Expanded = True
                End If
            ElseIf .MouseCol = mCol.ѡ�� Then
                If .GetNode(.Row).Expanded Then
                    For i = .Row To .Rows - 1
                        If .RowOutlineLevel(i) = 0 And i <> .Row Then Exit For '��һ��ҽ��
                        If i = .Row Then
                            lsel = .Cell(flexcpData, .Row, mCol.ѡ��)
                            .Cell(flexcpData, i, mCol.ѡ��) = IIf(lsel = 0, 1, 0)
                        Else
                            Set .Cell(flexcpPicture, i, mCol.ѡ��) = IIf(lsel = 0, img16.ListImages("Selected").Picture, Nothing)
                            .Cell(flexcpData, i, mCol.ѡ��) = IIf(lsel = 0, 1, 0)
                        End If
                    Next
                Else
                    .GetNode(.Row).Expanded = True
                End If
            Else
                .GetNode(.Row).Expanded = Not .GetNode(.Row).Expanded
            End If
        Else
            If .Cell(flexcpData, .Row, mCol.ѡ��) = 0 Then
                Set .Cell(flexcpPicture, .Row, mCol.ѡ��) = img16.ListImages("Selected").Picture
                .Cell(flexcpData, .Row, mCol.ѡ��) = 1
            Else
                .Cell(flexcpData, .Row, mCol.ѡ��) = 0
                Set .Cell(flexcpPicture, .Row, mCol.ѡ��) = Nothing
            End If
        End If
    ElseIf KeyCode = vbKeyLeft Then
        .GetNode(.Row).Expanded = False
    ElseIf KeyCode = vbKeyRight Then
        .GetNode(.Row).Expanded = True
    End If
    End With
End Sub

Private Sub vsList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngRow As Long, lngCol As Long
    If lblType(Val(shpSearch.Tag)).Caption <> "������" Then Exit Sub
    If mlngPatient = 0 Then Exit Sub
    
    With vsList
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow < 1 Then .MousePointer = flexDefault: Exit Sub
            
        If .MergeRow(lngRow) Then
            .ToolTipText = .Cell(flexcpData, lngRow, mCol.���)
        ElseIf lngCol = mCol.ѡ�� Then
            .ToolTipText = ""
        Else
            .ToolTipText = "˫����Ҫ������"
        End If
            
        If .GetNode(lngRow).Expanded And (lngCol = mCol.ѡ�� Or lngCol = mCol.��־) Then
            vsList.MousePointer = flexCustom
            Set vsList.MouseIcon = img16.ListImages("Selected").Picture
        Else
            vsList.MousePointer = flexDefault
        End If
    End With
End Sub

'����Ҽ��˵�
Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Me.vsList.Rows = 1 Then Exit Sub
     If Button = vbRightButton And Not Me.vsList.Row < 0 Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        Set Popup = CommandBars.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set Control = .Add(xtpControlButton, 100, "����(&I)")
            If lblType(Val(shpSearch.Tag)).Caption = "������" Then
                Set Control = .Add(xtpControlButton, 101, "ѡ���쳣(&S)")
                Set Control = .Add(xtpControlButton, 102, "ѡ�б�������(&A)")
            End If
            Popup.ShowPopup
        End With
    End If
End Sub
Public Function SetItems(lngFileID As Long, lngPatId As Long, lngVisit As Long, lngAdvice As Long)
    mlngFileID = lngFileID        '�ļ�id
    mlngPatient = lngPatId        '����id���ڲ��˲����༭ʱ������ȷ������ʾ���Ƿ�����
    mlngVisit = lngVisit          '��ҳid��Һŵ�ID
    mlngAdvice = lngAdvice
    If lblType(Val(shpSearch.Tag)).Caption = "���ĵ���" Then Call FillEPRDemos
    If lblType(Val(shpSearch.Tag)).Caption = "������" Then Call InitLisItem
    lblType(8).Enabled = IIf(mlngFileID = 0 And mlngPatient = 0 And mlngVisit = 0 And mlngAdvice = 0, False, True)
    lblType(12).Enabled = IIf(mlngFileID = 0 And mlngPatient = 0 And mlngVisit = 0 And mlngAdvice = 0, False, True)
End Function
Private Function GetLisItems(strAdvices As String) As ADODB.Recordset
Dim rsTemp As New ADODB.Recordset, strErr As String
Dim strContent As String, arrItems As Variant, arrItem As Variant, arrList As Variant, arrEle As Variant, i As Integer, l As Integer

    On Error GoTo errHand
    If strAdvices = "" Then Set GetLisItems = Nothing: Exit Function
    If mobjLis Is Nothing Then
        Set mobjLis = DynamicCreate("zl9LisInsideComm.clsLisInsideComm", False)
        If Not mobjLis Is Nothing Then
            If mobjLis.InitComponentsHIS(glngSys, 1070, gcnOracle, strErr) = False Then
                Set mobjLis = Nothing
            End If
        End If
    End If
    
    If mobjLis Is Nothing And strErr = "" Then
        Set GetLisItems = Nothing
    Else
        '--ʹ���°�LIS
        'mobjLisInsideComm.GetPatientSampleValue (lngPatientID)
        '����                   ��ȡָ���걾�Ľ��
        '����                   lngPatientID   ����ID
        '����
'             ����(1=��ͨ)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2>Ӥ�����<split2>
'            ָ��1<split4>������1<split4>��λ1<split4>�����־1<split4>�������1<split4>�������1<split4>��˽��Ŀ1<split4>ָ�����1<split4>������1<split4>Ӣ����1<split3>
'            ָ��2<split4>������2<split4>��λ2<split4>�����־2<split4>�������2<split4>�������2<split4>��˽��Ŀ2<split4>ָ�����2<split4>������2<split4>Ӣ����2<split3>
'            ָ��3<split4>������3<split4>��λ3<split4>�����־3<split4>�������3<split4>�������3<split4>��˽��Ŀ3<split4>ָ�����3<split4>������3<split4>Ӣ����3<split1>
'
'            ����(2=΢����)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2>
'            ϸ����1<split3>����1<split3>��ҩ����1<split3>
'            ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
'            ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split2>
'            ϸ����2<split3>����2<split3>��ҩ����2<split3>
'            ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
'            ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split1>
'
'            �ָ������������:
'            1.  ���ڷָ��걾,ʹ��"<split1>"�ָ�����ǰʹ��"|"
'            2.  ���ڷָ��걾��Ϣ,ʹ��"<split2>"�ָ�����ǰʹ��";"
'            3.  ���ڷָ��걾ָ����Ϣ,ʹ��"<split3>"�ָ�����ǰʹ��","
'            4.  ���ڷָ�ָ������Ϣ,ʹ��"<split4>"�ָ�����ǰʹ��"^"

        strContent = mobjLis.GetSampleValue(strAdvices)
        arrItems = Array() '��Ŀ�б�
        arrItem = Array() '������Ŀ��Ϣ
        arrEle = Array() 'ָ���嵥
        With rsTemp
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            With .Fields
                .Append "ҽ������", adVarChar, 1000, adFldIsNullable
                .Append "������Ŀ", adVarChar, 100, adFldIsNullable
                .Append "��д", adVarChar, 100, adFldIsNullable
                .Append "������", adVarChar, 1000, adFldIsNullable
                .Append "��λ", adVarChar, 100, adFldIsNullable
                .Append "��־", adVarChar, 100, adFldIsNullable
                .Append "�ο�", adVarChar, 100, adFldIsNullable
                .Append "������Դ", adVarChar, 100, adFldIsNullable
                .Append "���ʱ��", adDBTimeStamp, 100, adFldIsNullable
                .Append "Ӥ��", adInteger, 100, adFldIsNullable
            End With
            .Open
        End With
        
        If strContent <> "" Then
            arrItems = Split(strContent, "<split1>") '�����߷ָ��Ķ����Ŀ
            For i = 0 To UBound(arrItems)
                arrItem = Split(arrItems(i), "<split2>") '�Էֺŷָ�����Ŀ��Ϣ
                If arrItem(mcItem.����) = 1 And arrItem(mcItem.�����) <> "" Then 'ֻ������ͨ���鲢����˹��ģ�������΢����
                    arrList = Array()
                    arrList = Split(arrItem(UBound(arrItem)), "<split3>") '�Զ��ŷָ��Ķ��ָ��
                    For l = 0 To UBound(arrList)
                        arrEle = Split(arrList(l), "<split4>") 'ÿ��ָ������Ϣ��^�ָ�
                        With rsTemp
                            .AddNew
                            !ҽ������ = arrItem(mcItem.��Ŀ����)
                            !������Դ = Decode(Val(arrItem(mcItem.������Դ)), 1, "����", 2, "סԺ", 4, "���", "����")
                            !���ʱ�� = arrItem(mcItem.���ʱ��)
                            If UBound(arrEle) >= CLng(mcList.������) Then
                                !������Ŀ = arrEle(mcList.������)
                            Else
                                !������Ŀ = arrEle(mcList.ָ��)
                            End If
                            If UBound(arrEle) >= CLng(mcList.Ӣ����) Then
                                !��д = arrEle(mcList.Ӣ����)
                            Else
                                !��д = arrEle(mcList.����)
                            End If
                            !������ = arrEle(mcList.���)
                            !��λ = arrEle(mcList.��λ)
                            !��־ = arrEle(mcList.��־)
                            !�ο� = arrEle(mcList.�ο�)
                            If UBound(arrItem) > 9 Then
                                !Ӥ�� = CInt(Val(arrItem(mcItem.Ӥ��)))
                            Else
                                !Ӥ�� = 0
                            End If
                            .Update
                        End With
                    Next
                End If
            Next
            If Not rsTemp.EOF Then
                rsTemp.MoveFirst
            End If
        End If
        Set GetLisItems = rsTemp
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


