VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportPath 
   Caption         =   "�����׼·��"
   ClientHeight    =   10005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11730
   Icon            =   "frmImportPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   11730
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9375
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   3
         Left            =   7320
         Picture         =   "frmImportPath.frx":6852
         ScaleHeight     =   495
         ScaleWidth      =   2775
         TabIndex        =   81
         Top             =   3480
         Width           =   2775
      End
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   7320
         Picture         =   "frmImportPath.frx":7983
         ScaleHeight     =   495
         ScaleWidth      =   2295
         TabIndex        =   80
         Top             =   2760
         Width           =   2295
      End
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   7320
         Picture         =   "frmImportPath.frx":86B7
         ScaleHeight     =   495
         ScaleWidth      =   2775
         TabIndex        =   79
         Top             =   2040
         Width           =   2775
      End
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   7320
         Picture         =   "frmImportPath.frx":97DA
         ScaleHeight     =   495
         ScaleWidth      =   3135
         TabIndex        =   78
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Frame fraRuleDefine 
         Caption         =   "���½ṹ����"
         Height          =   3090
         Left            =   1080
         TabIndex        =   49
         Top             =   1080
         Width           =   9690
         Begin VB.CommandButton cmdSize 
            Caption         =   "ѡ��"
            Height          =   350
            Index           =   3
            Left            =   5400
            TabIndex        =   65
            Top             =   2520
            Width           =   550
         End
         Begin VB.CommandButton cmdSize 
            Caption         =   "ѡ��"
            Height          =   350
            Index           =   2
            Left            =   5400
            TabIndex        =   64
            Top             =   1680
            Width           =   550
         End
         Begin VB.CommandButton cmdSize 
            Caption         =   "ѡ��"
            Height          =   350
            Index           =   1
            Left            =   5400
            TabIndex        =   63
            Top             =   1035
            Width           =   550
         End
         Begin VB.CommandButton cmdSize 
            Caption         =   "ѡ��"
            Height          =   350
            Index           =   0
            Left            =   5400
            TabIndex        =   62
            Top             =   275
            Width           =   550
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "�Ӵ�"
            Height          =   255
            Index           =   3
            Left            =   4560
            TabIndex        =   61
            Tag             =   "0"
            Top             =   2550
            Width           =   735
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "�Ӵ�"
            Height          =   255
            Index           =   2
            Left            =   4560
            TabIndex        =   60
            Tag             =   "1"
            Top             =   1710
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "�Ӵ�"
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   59
            Tag             =   "0"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "�Ӵ�"
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   58
            Tag             =   "1"
            Top             =   323
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   7
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   57
            Tag             =   "12"
            Text            =   "12"
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   6
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   56
            Tag             =   "����"
            Text            =   "����"
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   5
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   55
            Tag             =   "14"
            Text            =   "14"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   4
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   54
            Tag             =   "����"
            Text            =   "����"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   3
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   53
            Tag             =   "16"
            Text            =   "16"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   2
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   52
            Tag             =   "����"
            Text            =   "����"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   1
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   51
            Tag             =   "18"
            Text            =   "18"
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   0
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   50
            Tag             =   "����"
            Text            =   "����"
            Top             =   300
            Width           =   1095
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9720
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9720
            Y1              =   2295
            Y2              =   2295
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9720
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9720
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��С"
            Height          =   180
            Index           =   10
            Left            =   3240
            TabIndex        =   77
            Top             =   2580
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   9
            Left            =   1200
            TabIndex        =   76
            Top             =   2580
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   3
            Left            =   480
            TabIndex        =   75
            Top             =   2580
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��С"
            Height          =   180
            Index           =   8
            Left            =   3240
            TabIndex        =   74
            Top             =   1740
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   5
            Left            =   1200
            TabIndex        =   73
            Top             =   1740
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   1740
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��С"
            Height          =   180
            Index           =   7
            Left            =   3240
            TabIndex        =   71
            Top             =   1140
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   6
            Left            =   1200
            TabIndex        =   70
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   69
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��С"
            Height          =   180
            Index           =   4
            Left            =   3240
            TabIndex        =   68
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   13
            Left            =   1200
            TabIndex        =   67
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "�����"
            Height          =   180
            Index           =   0
            Left            =   300
            TabIndex        =   66
            Top             =   360
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPathTable 
         Height          =   2205
         Left            =   1080
         TabIndex        =   17
         Top             =   6720
         Width           =   10305
         _cx             =   1963869953
         _cy             =   1963855674
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   20
         RowHeightMax    =   5000
         ColWidthMin     =   100
         ColWidthMax     =   12000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmImportPath.frx":B05C
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
      Begin VB.Label lblSet 
         Caption         =   "��׼·�����̵ĵ����ʽ���£�"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   27
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵��5���������ָ·�����ƣ�����������ָ�������ƻ�����ƣ�����������ָ·�����̵���Ŀ���ƣ�������ָ·���������ݡ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   19
         Left            =   1080
         TabIndex        =   26
         Top             =   5520
         Width           =   10575
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵��4���汾��ϢĬ��λ��·���������һ�β������ؼ���""��""��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   18
         Left            =   1080
         TabIndex        =   25
         Top             =   5160
         Width           =   5175
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵��2���������鶨����,���ܴ���������ȫ��ͬ�����������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   17
         Left            =   1080
         TabIndex        =   24
         Top             =   4530
         Width           =   5055
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵��3��·������������""�ٴ�·��""�ؼ��֣����������""�ٴ�·����""�ؼ��ʡ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   16
         Left            =   1080
         TabIndex        =   23
         Top             =   4860
         Width           =   6735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵��1�����½ṹ�Ľ�������������Ҫ��ʶ,Ȼ���Թؼ���ƥ�䡣"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   20
         Left            =   1080
         TabIndex        =   22
         Top             =   4200
         Width           =   6255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSet 
         Caption         =   "��ʾ��һ����Ԫ���ﲻ�ܳ��ֶ���Ļس�����һ����Ԫ���ﲻ�ܳ��ֶ��л����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   21
         Top             =   9120
         Width           =   7935
      End
      Begin VB.Label lblSet 
         Caption         =   "��׼·�����ĵ����ʽ���£�"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   20
         Top             =   6360
         Width           =   4095
      End
      Begin VB.Label lblTitle 
         Caption         =   "�ڶ��� �������ݹ������ʽ˵��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Tag             =   "�ڶ��� �������ݹ������ʽ˵��"
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵��6�����������������������Ҫ����ĸ�ʽ���óɱ�����ʽ�����ĸ�ʽ����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   25
         Left            =   1080
         TabIndex        =   18
         Top             =   5880
         Width           =   9495
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9495
      Index           =   4
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame fraProcess 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   480
         TabIndex        =   42
         Top             =   4200
         Width           =   10695
         Begin MSComctlLib.ProgressBar prgImp 
            Height          =   360
            Index           =   1
            Left            =   1080
            TabIndex        =   43
            Top             =   630
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar prgImp 
            Height          =   360
            Index           =   0
            Left            =   1080
            TabIndex        =   44
            Top             =   150
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblPrg 
            AutoSize        =   -1  'True
            Caption         =   "��ǰ����"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblPrg 
            AutoSize        =   -1  'True
            Caption         =   "�ܽ���"
            Height          =   180
            Index           =   1
            Left            =   420
            TabIndex        =   45
            Top             =   720
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsErrInfo 
         Height          =   6615
         Left            =   600
         TabIndex        =   47
         Top             =   960
         Visible         =   0   'False
         Width           =   10455
         _cx             =   2004371465
         _cy             =   2004364692
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmImportPath.frx":B0C8
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
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "���岽 ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   4
         Left            =   480
         TabIndex        =   48
         Tag             =   "���岽 ����"
         Top             =   360
         Width           =   1665
      End
   End
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton cmdBraw 
         Caption         =   "���"
         Height          =   350
         Left            =   9480
         TabIndex        =   9
         Top             =   960
         Width           =   1110
      End
      Begin VB.TextBox txtFile 
         Height          =   270
         Left            =   2880
         TabIndex        =   8
         Top             =   1005
         Width           =   6495
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "�ļ���"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   7
         Top             =   1005
         Width           =   855
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "�ļ�"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   6
         Top             =   1005
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame fraFloder 
         Caption         =   "����ļ���"
         Height          =   5025
         Left            =   2880
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   6495
         Begin VB.DriveListBox div 
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   276
            Width           =   6165
         End
         Begin VB.DirListBox dirFloder 
            Height          =   3870
            Left            =   96
            TabIndex        =   4
            Top             =   600
            Width           =   6165
         End
         Begin VB.CommandButton cmdPathOk 
            Caption         =   "ȷ��(&O)"
            Height          =   350
            Left            =   3840
            TabIndex        =   3
            Top             =   4560
            Width           =   1100
         End
         Begin VB.CommandButton cmdPathCancel 
            Caption         =   "ȡ��(&C)"
            Height          =   350
            Left            =   5160
            TabIndex        =   2
            Top             =   4560
            Width           =   1100
         End
      End
      Begin MSComDlg.CommonDialog dlgCom 
         Left            =   120
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "ע2���������ļ�ʱ��ʱ����ܻ�Ƚϳ��������ĵȴ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   11
         Left            =   960
         TabIndex        =   82
         Top             =   3960
         Width           =   4935
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         Caption         =   "��һ�� ѡ����Ҫ������ļ������ļ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Tag             =   "��һ�� ѡ����Ҫ������ļ������ļ���"
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵�����ļ�������������""-""ʱ���Զ������ļ���,��ʽ��""������-����ǰ׺-������ʼֵ-�汾"",�ָ���Ĭ��Ϊ"".""��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   21
         Left            =   960
         TabIndex        =   11
         Top             =   4320
         Width           =   7575
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "ע1��������ļ�����Ϊword�ĵ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   22
         Left            =   960
         TabIndex        =   10
         Top             =   3600
         Width           =   5175
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9495
      Index           =   2
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VSFlex8Ctl.VSFlexGrid vsDefineImp 
         Height          =   6375
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   11055
         _cx             =   1969703980
         _cy             =   1969695725
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmImportPath.frx":B151
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
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵�����ָ����Ϊ"".""��""-""��������ʼֵ��������Ȼ��,����·���Դ�Ϊ��������������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   23
         Left            =   240
         TabIndex        =   31
         Top             =   8160
         Width           =   6975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "������ ���õ��������ѡ�����ļ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   30
         Tag             =   "������ ���õ��������ѡ�����ļ�"
         Top             =   360
         Width           =   5460
      End
   End
   Begin VB.Frame frmImport 
      BorderStyle     =   0  'None
      Height          =   9495
      Index           =   3
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame fraProcess 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   480
         TabIndex        =   33
         Top             =   4320
         Width           =   10695
         Begin MSComctlLib.ProgressBar prgImp 
            Height          =   360
            Index           =   2
            Left            =   960
            TabIndex        =   34
            Top             =   570
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar prgImp 
            Height          =   360
            Index           =   3
            Left            =   960
            TabIndex        =   35
            Top             =   90
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblPrg 
            AutoSize        =   -1  'True
            Caption         =   "�ܽ���"
            Height          =   180
            Index           =   2
            Left            =   300
            TabIndex        =   37
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblPrg 
            AutoSize        =   -1  'True
            Caption         =   "��ǰ����"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   180
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAnalyse 
         Height          =   6720
         Left            =   480
         TabIndex        =   38
         Top             =   960
         Width           =   10605
         _cx             =   2004371730
         _cy             =   2004364877
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmImportPath.frx":B243
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
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵������Ҫ���ֵ���,��ѡ����Ӧ·�����е��롣"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   24
         Left            =   480
         TabIndex        =   40
         Top             =   8040
         Width           =   4335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "���Ĳ� ѡ��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   39
         Tag             =   "���Ĳ� ѡ��������"
         Top             =   360
         Width           =   3150
      End
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "��һ��(&N)"
      Height          =   350
      Index           =   1
      Left            =   9120
      TabIndex        =   15
      Top             =   9600
      Width           =   1110
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "�˳�(&X)"
      Height          =   350
      Index           =   2
      Left            =   10320
      TabIndex        =   14
      Top             =   9600
      Width           =   1110
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "��һ��(&P)"
      Height          =   350
      Index           =   0
      Left            =   7920
      TabIndex        =   13
      Top             =   9600
      Width           =   1110
   End
End
Attribute VB_Name = "frmImportPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintPage As Integer
Private mblnFile As Boolean
Private mlngSelFileCount As Long 'ѡ����ļ�����
Private mlngSelPathCount As Long 'ѡ���·������
Private mlngImpFileCount As Long '�Ѿ�������ļ�����
Private mlngImpPathCount As Long '�Ѿ������·������

'��������������������ɵ��ַ�������ʽΪ"������,�����С,�Ƿ�Ӵ�"
Private mstrFontStr����� As String
Private mstrFontStr�������� As String
Private mstrFontStrС���� As String
Private mstrFontStr���� As String
Private mlngStPathID As Long

Private Enum PageEnu
    PE_PathInput = 0 '�ļ�·������
    PE_DefineImp = 1 '�������ݹ���
    PE_AnaRules = 2 '��������
    PE_AnaResult = 3 '�������
    PE_ErrInfo = 4 '������Ϣ��ʾ
End Enum
Private Enum DefineImpCols
    DC_ѡ�� = 0
    DC_�ļ����� = 1
    DC_�������� = 2
    DC_�汾 = 3
    DC_����ǰ׺ = 4
    DC_�ָ��� = 5
    DC_������ʼֵ = 6
    DC_�ļ�·�� = 7
End Enum

Private Enum AnaCols
    AC_ѡ�� = 0
    AC_·������ = 1
    AC_���� = 2
    AC_�汾 = 3
    ac_���� = 4
    AC_���Ŀ�ʼ = 5
    AC_���Ľ��� = 6
    AC_���⿪ʼ = 7
    AC_�ļ�·�� = 8
End Enum

Private Enum ErrCols
    EC_�ļ���
    EC_·������
    EC_������Ϣ
End Enum

Private Sub cmdBraw_Click()
'���ܣ�ѡ���ļ����ļ���
    If optSelect(1).Value = True Then
        fraFloder.Visible = True
        '������ǰҳ��
        Call SetContolStat(PE_PathInput, Not fraFloder.Visible)
        cmdImport(0).Enabled = False
        cmdImport(0).Visible = False
    Else
        With dlgCom
            .FileName = ""
            .DialogTitle = "ѡ���ļ�"
            .FileName = ""
            .Filter = ".docx"
            .ShowOpen
            If .FileName <> "" Then
              txtFile.Text = .FileName
            End If
        End With
    End If
End Sub

Private Sub cmdImport_Click(Index As Integer)
    Dim blnCanNext As Boolean
    Dim intNextPage As Integer
    Dim intPrePage As Integer
    Dim str����ǰ׺ As String
    Dim str�ָ��� As String
    Dim str������ʼֵ As String
    Dim i As Long
    Dim strMsg As String
    
    Select Case Index
        Case 1
            If mintPage = 0 Then mblnFile = optSelect(0).Value
            
            blnCanNext = CheckStepNext(mintPage, mblnFile, intNextPage)
            
            If blnCanNext Then
                If mintPage = 2 Then
                    For i = 1 To vsDefineImp.Rows - 1
                        str����ǰ׺ = Trim(vsDefineImp.TextMatrix(i, DC_����ǰ׺))
                        str������ʼֵ = Trim(vsDefineImp.TextMatrix(i, DC_������ʼֵ))
                        str�ָ��� = Trim(vsDefineImp.TextMatrix(i, DC_�ָ���))
                        If Not zlCommFun.IsNumOrChar(str����ǰ׺) Then
                            strMsg = "�ڡ�" & i & "���б���ǰ׺ֻ�������֡���ĸ������ĸ�����ֵ���ϣ�"
                        ElseIf Not IsNumeric(str������ʼֵ) Then
                            strMsg = "�ڡ�" & i & "���б�����ʼֵֻ������Ȼ����"
                        End If
                        If InStr(".-", Trim(str�ָ���)) = 0 Then
                            strMsg = "�ڡ�" & i & "���зָ���ֻ���ǡ�.�����ߡ�-����"
                        Else
                            If Len(Trim(str�ָ���)) > 1 Then
                                strMsg = "�ڡ�" & i & "���зָ���ֻ������1���ַ����ȣ�"
                            End If
                        End If
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName
                            intNextPage = 2
                            frmImport(mintPage).Visible = False
                            frmImport(intNextPage).Visible = True
                            '���������������ݺ����
                            Call SetContolStat(intNextPage, False)
                            If intNextPage = PE_PathInput Or intNextPage = PE_ErrInfo Then
                                cmdImport(0).Enabled = False
                                cmdImport(0).Visible = False
                            End If
                            lblTitle(intNextPage).Caption = lblTitle(intNextPage).Tag
                            cmdImport(1).Caption = "��һ��(&N)"
                            If intNextPage = PE_ErrInfo Then
                                cmdImport(1).Caption = "����(&B)"
                            End If
                            mintPage = intNextPage
                            Call SetContolStat(intNextPage, True)
                            Exit Sub
                        End If
                    Next
                End If
                frmImport(mintPage).Visible = False
                frmImport(intNextPage).Visible = True
                '���������������ݺ����
                Call SetContolStat(intNextPage, False)
                If intNextPage = PE_PathInput Or intNextPage = PE_ErrInfo Then
                    cmdImport(0).Enabled = False
                    cmdImport(0).Visible = False
                End If
                lblTitle(intNextPage).Caption = lblTitle(intNextPage).Tag
                cmdImport(1).Caption = "��һ��(&N)"
                If intNextPage = PE_ErrInfo Then
                    cmdImport(1).Caption = "����(&B)"
                End If
                mintPage = intNextPage
                '������һҳ������,���ݼ��غ�������
                Select Case intNextPage
                    Case PE_PathInput '�����һҳ����
                        Call ClearPage(-1) '�������
                            Call SetContolStat(intNextPage, True)
                    Case PE_DefineImp
                        Call SetContolStat(intNextPage, True)
                        With vsPathTable
                            .Rows = 4
                            .Cols = 4
                            .TextMatrix(0, 0) = "ʱ��"
                            .TextMatrix(0, 1) = "סԺ��һ��"
                            .TextMatrix(0, 2) = "סԺ�ڶ���(������)"
                            .TextMatrix(0, 3) = "סԺ������(��Ժ�գ�"
                            .TextMatrix(1, 0) = "��Ҫ���ƹ���"
                            .TextMatrix(1, 1) = "ѯ�ʲ�ʷ������"
                            .TextMatrix(1, 2) = "����ۿ�������"
                            .TextMatrix(1, 3) = "����ۿ�������"
                            .TextMatrix(2, 0) = "�ص�ҽ��"
                            .TextMatrix(2, 1) = "����ҽ��:�ۿ�����������"
                            .TextMatrix(2, 2) = "����ҽ�������󣩣��ۿƶ���������"
                            .TextMatrix(2, 3) = "����ҽ�������󣩣��ۿƶ���������"
                            .TextMatrix(3, 0) = "��Ҫ������"
                            .TextMatrix(3, 1) = "ִ��ҽ���������������"
                            .TextMatrix(3, 2) = "ִ��ҽ���������������"
                            .TextMatrix(3, 3) = "ִ��ҽ���������������"
                        End With
                        Call SetVsStyle
                Case PE_AnaRules
                    Call LoadFileList(txtFile.Text, mblnFile)
                    Call SetContolStat(intNextPage, True)
                Case PE_AnaResult
                    If Not LoadAnalyseResult Then
                        intNextPage = 2
                        frmImport(mintPage).Visible = False
                        frmImport(intNextPage).Visible = True
                        '���������������ݺ����
                        Call SetContolStat(intNextPage, False)
                        If intNextPage = PE_PathInput Or intNextPage = PE_ErrInfo Then
                            cmdImport(0).Enabled = False
                            cmdImport(0).Visible = False
                        End If
                        lblTitle(intNextPage).Caption = lblTitle(intNextPage).Tag
                        cmdImport(1).Caption = "��һ��(&N)"
                        If intNextPage = PE_ErrInfo Then
                            cmdImport(1).Caption = "����(&B)"
                        End If
                        mintPage = intNextPage
                    End If
                    Call SetContolStat(intNextPage, True)
                Case PE_ErrInfo
                    If LoadPath Then
                        Call SetContolStat(intNextPage, True)
                        MsgBox "¼�����", vbInformation + vbOKOnly, Me.Caption
                    Else
                        Call SetContolStat(intNextPage, True)
                        MsgBox "¼��ʧ��", vbInformation + vbOKOnly, Me.Caption
                    End If
                End Select
            End If
        Case 0
            intPrePage = GetStepPre(mintPage)
            frmImport(mintPage).Visible = False
            frmImport(intPrePage).Visible = True
            cmdImport(0).Enabled = True
            cmdImport(0).Visible = True
            lblTitle(intPrePage).Caption = lblTitle(intPrePage).Tag
            cmdImport(1).Caption = "��һ��(&N)"
            If mintPage = PE_PathInput Then
                cmdImport(0).Enabled = False
                cmdImport(0).Visible = False
            End If
        
            '�����ǰ��������
            Select Case intPrePage
            Case PE_PathInput
                 Call ClearPage(-1)
            Case PE_DefineImp
                Call ClearPage(PE_AnaRules)
                Call ClearPage(PE_AnaResult)
                Call ClearPage(PE_ErrInfo)
            Case PE_AnaRules
                Call ClearPage(PE_AnaResult)
                Call ClearPage(PE_ErrInfo)
            Case PE_AnaResult
                Call ClearPage(PE_ErrInfo)
            End Select
            mintPage = intPrePage
        Case 2
            '���ܣ��˳�
            Unload Me
    End Select
End Sub

Private Function CheckStepNext(ByVal intPage As Integer, ByVal blnFile As Boolean, ByRef intNextPage As Integer) As Boolean
'���ܣ����е�ǰ����ļ�飬���Ƿ��ܽ�����һ������
'      intPage :��ǰ�ɼ�ҳ���index
'      blnfile  :��������ļ�����
'      intNextPage :��һҳ���Index
    Dim fileTemp As File, flrTemp As Folder, objfso As New FileSystemObject
    Dim strPath As String, strTest As String, strTmp As String
    Dim blnCanNext As Boolean, blnReback As Boolean
    Dim i As Long, lngRowCount As Long

    strPath = Trim(txtFile.Text)
    mlngSelFileCount = 0
    mlngSelPathCount = 0
    Select Case intPage
            Case PE_PathInput
                If blnFile Then
                    If objfso.FileExists(strPath) Then '�ļ�����
                        Set fileTemp = objfso.GetFile(strPath)
                        If (UCase(Right(fileTemp.Name, 4)) = ".DOC" Or UCase(Right(fileTemp.Name, 5)) = ".DOCX") And Mid(fileTemp.Name, 1, 2) <> "~$" Then
                            blnCanNext = True
                        Else
                            MsgBox "��������ļ����Ͳ��ǿ��Ե����Word�ļ�,����������", vbInformation, "ϵͳ��Ϣ"
                        End If
                    Else
                        MsgBox "��������ļ�������,����������", vbInformation, "ϵͳ��Ϣ"
                        Call txtFile.SetFocus
                    End If
                Else
                    If objfso.FolderExists(strPath) Then  '�ļ��д���
                        Set flrTemp = objfso.GetFolder(strPath)
                        For Each fileTemp In flrTemp.Files
                            If (UCase(Right(fileTemp.Name, 4)) = ".DOC" Or UCase(Right(fileTemp.Name, 5)) = ".DOCX") And Mid(fileTemp.Name, 1, 2) <> "~$" Then
                                blnCanNext = True
                                Exit For
                            End If
                        Next
                        If Not blnCanNext Then MsgBox "�ļ��в����ڿ��Ե����Word�ļ�,����������", vbInformation, "ϵͳ��Ϣ"
                    Else
                        MsgBox "��������ļ��в�����,����������", vbInformation, "ϵͳ��Ϣ"
                        Call txtFile.SetFocus
                    End If
                End If
            Case PE_DefineImp
                For i = chkBold.LBound To chkBold.UBound
                    If Trim(txtInfo(i * 2).Text) = "" Then
                        MsgBox "��������������", vbInformation, "ϵͳ��Ϣ"
                        txtInfo(i).SetFocus
                        Exit Function
                    End If
                    
                    strTmp = txtInfo(i * 2).Text & "," & txtInfo(i * 2 + 1).Text & "," & IIf(chkBold(i).Value = 1, 1, 0)
                    strTest = strTest & "|" & strTmp
                    lblInfo(i).Tag = strTmp
                Next
                
                For i = chkBold.LBound To chkBold.UBound
                    If HaveMoreStr(strTest, lblInfo(i).Tag) Then
                        MsgBox "�������ֻ������ͬ���ĵ��ṹ����", vbInformation, "ϵͳ��Ϣ"
                        Exit Function
                    End If
                Next
                blnCanNext = True
                '���������
                mstrFontStr����� = lblInfo(0).Tag
                mstrFontStr�������� = lblInfo(1).Tag
                mstrFontStrС���� = lblInfo(2).Tag
                mstrFontStr���� = lblInfo(3).Tag
            Case PE_AnaRules
                With vsDefineImp
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, DC_��������) = "" Then
                            MsgBox "�������Ʋ��ܴ��ļ��ļ��ڲ���ȡ,���ֹ������������", vbInformation, "ϵͳ��Ϣ"
                            vsDefineImp.SetFocus
                            vsDefineImp.Select i, DC_��������
                            Exit Function
                        End If

                        If .TextMatrix(i, DC_����ǰ׺) = "" Then
                            MsgBox "���벻�ܴ��ļ��ж�ȡ����ȷ���������", vbInformation, "ϵͳ��Ϣ"
                            vsDefineImp.SetFocus
                            vsDefineImp.Select i, DC_����ǰ׺
                            Exit Function
                        End If

                        If .TextMatrix(i, DC_�ָ���) = "" Then
                            MsgBox "���벻�ܴ��ļ��ж�ȡ����ȷ���������", vbInformation, "ϵͳ��Ϣ"
                            vsDefineImp.SetFocus
                            vsDefineImp.Select i, DC_�ָ���
                            Exit Function
                        End If

                        If .TextMatrix(i, DC_������ʼֵ) = "" Then
                            MsgBox "���벻�ܴ��ļ��ж�ȡ����ȷ���������", vbInformation, "ϵͳ��Ϣ"
                            vsDefineImp.SetFocus
                            vsDefineImp.Select i, DC_������ʼֵ
                            Exit Function
                        End If

                        If .TextMatrix(i, DC_ѡ��) = "-1" Then
                            mlngSelFileCount = mlngSelFileCount + 1
                        End If
                    Next i
                    blnCanNext = mlngSelFileCount <> 0
                    If Not blnCanNext Then
                        MsgBox "����δѡ��Ҫ������ļ�", vbInformation, "ϵͳ��Ϣ"
                    End If
                End With
            Case PE_AnaResult
                With vsAnalyse
                    lngRowCount = 0
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, AC_·������) <> "" Then
                            If .TextMatrix(i, AC_ѡ��) = "-1" Then
                                mlngSelPathCount = mlngSelPathCount + 1
                            End If
                            lngRowCount = lngRowCount + 1
                        End If
                    Next i
                End With
                blnCanNext = mlngSelPathCount <> 0
                blnReback = lngRowCount = 0
                If Not blnCanNext And Not blnReback Then
                    MsgBox "����δѡ��Ҫ�����·������ѡ��·�����е���", vbInformation, "ϵͳ��Ϣ"
                    Exit Function
                End If
            Case PE_ErrInfo
                blnCanNext = True
    End Select
    'ȷ����һҳ��ҳ��
    If intPage = PE_ErrInfo Then
        intNextPage = PE_PathInput
    Else
        If intPage = PE_AnaResult And blnReback Then 'û�н�����·��,������ҳ
            MsgBox "���������κ�·������鿴�����ļ����Ƿ���ȷ��������Ծ�����Ƿ���ȷ", vbInformation, "ϵͳ��Ϣ"
            intNextPage = PE_PathInput
        Else
            intNextPage = intPage + 1
        End If
    End If
    CheckStepNext = blnCanNext
End Function

Private Sub LoadFileList(ByVal strPath As String, ByVal blnFile As Boolean)
'���ܣ������ļ��б�
'   strPath:�ļ���·�����ļ�·��
'   blnFile:���ļ���ʽ����
    Dim fileTemp As File, flrTemp As Folder, objfso As New FileSystemObject
    Dim arrTmp As Variant
    Dim i As Long, lngRow  As Long
    Dim strTem As String
    Dim strFileFullName As String, str�ļ��� As String
    
    vsDefineImp.Rows = vsDefineImp.FixedRows
    If blnFile Then '�����ļ�����
        Set fileTemp = objfso.GetFile(strPath)
        str�ļ��� = fileTemp.Name
        strFileFullName = fileTemp.Path
        Call AddNewFileRow(strFileFullName, str�ļ���)
    Else
        Set flrTemp = objfso.GetFolder(strPath)
        For Each fileTemp In flrTemp.Files
            If (UCase(Right(fileTemp.Name, 4)) = ".DOC" Or UCase(Right(fileTemp.Name, 5)) = ".DOCX") And Mid(fileTemp.Name, 1, 2) <> "~$" Then '�Ǳ��ݷ������ļ�
                str�ļ��� = fileTemp.Name
                strFileFullName = fileTemp.Path
                Call AddNewFileRow(strFileFullName, str�ļ���)
                str�ļ��� = ""
                strFileFullName = ""
            End If
        Next
    End If
End Sub

Private Sub AddNewFileRow(ByVal strFileFullName As String, ByVal str�ļ��� As String)
'����:����µ�һ���ļ���Ϣ
    Dim lngRow As Long, arrTmp As Variant
    With vsDefineImp
        .Rows = .Rows + 1
        lngRow = .Rows - 1
        .Cell(flexcpChecked, lngRow, DC_ѡ��) = True
        .TextMatrix(lngRow, DC_�ļ�����) = str�ļ���
        .TextMatrix(lngRow, DC_�ļ�·��) = strFileFullName
        If InStr(str�ļ���, "-") > 0 Then
            arrTmp = Split(str�ļ���, "-")
            If UBound(arrTmp) >= 1 Then
                .TextMatrix(lngRow, DC_��������) = Trim(arrTmp(0))
                .TextMatrix(lngRow, DC_����ǰ׺) = Trim(arrTmp(1))
            End If
            
            If UBound(arrTmp) >= 2 Then
                .TextMatrix(lngRow, DC_������ʼֵ) = Trim(arrTmp(2))
            End If
            
            If UBound(arrTmp) >= 3 Then
                .TextMatrix(lngRow, DC_�汾) = Trim(arrTmp(3))
            End If
        End If
        .TextMatrix(lngRow, DC_�ָ���) = IIf(.colData(DC_�ָ���) = ".", .colData(DC_�ָ���), ".")
        .TextMatrix(lngRow, DC_��������) = IIf(.TextMatrix(lngRow, DC_��������) = "", .colData(DC_��������), .TextMatrix(lngRow, DC_��������))
        .TextMatrix(lngRow, DC_����ǰ׺) = IIf(.TextMatrix(lngRow, DC_����ǰ׺) = "", .colData(DC_����ǰ׺), .TextMatrix(lngRow, DC_����ǰ׺))
        .TextMatrix(lngRow, DC_������ʼֵ) = IIf(.TextMatrix(lngRow, DC_������ʼֵ) = "", .colData(DC_������ʼֵ), .TextMatrix(lngRow, DC_������ʼֵ))
        .TextMatrix(lngRow, DC_�汾) = IIf(.TextMatrix(lngRow, DC_�汾) = "", .colData(DC_�汾), .TextMatrix(lngRow, DC_�汾))
        
    End With
End Sub

Private Function LoadAnalyseResult() As Boolean

'���ܣ����ؽ������
    Dim i As Long, lngCount As Long
    vsAnalyse.Visible = False
    fraProcess(1).Visible = True
    With vsDefineImp
        prgImp(2).Max = mlngSelFileCount
        prgImp(2).Value = 0
        '�������
        vsAnalyse.Rows = .FixedRows
        For i = .FixedRows To .Rows - 1
            prgImp(2).Value = lngCount
            If .TextMatrix(i, DC_ѡ��) = "-1" Then
                lngCount = lngCount + 1
                If AnalyseDoc(.TextMatrix(i, DC_�ļ�·��), .TextMatrix(i, DC_�ļ�����), .TextMatrix(i, DC_��������), .TextMatrix(i, DC_�汾), .TextMatrix(i, DC_����ǰ׺) & .TextMatrix(i, DC_�ָ���), Val(.TextMatrix(i, DC_������ʼֵ))) = False Then
                    LoadAnalyseResult = False
                    Exit Function
                End If
            End If
        Next
        If vsAnalyse.Rows = .FixedRows Then
            vsAnalyse.Rows = vsAnalyse.Rows + 1
            LoadAnalyseResult = False
            Exit Function
        Else
            LoadAnalyseResult = True
        End If
    End With
    vsAnalyse.Visible = True
    fraProcess(1).Visible = False
End Function

Private Function AnalyseDoc(ByVal strFilePath As String, ByVal str�ļ����� As String, ByVal str���� As String, ByVal str�汾In As String, ByVal str���� As String, ByVal lng������ʼֵ As Long) As Boolean
'���ܣ������ĵ�,�����������Ӧ�ڱ����
'      str����:�������õĿ�������
'      str�汾In:�������õİ汾
'      str����:�������õı���ǰ׺&�ָ���
'      lng������ʼֵ:�������õı�����ʼֵ

    Dim objWord As Object, objWordApp As Object
    Dim rngFind As Object, rng�汾 As Object
    Dim i As Long, j As Long, BlnFind As Boolean, lngRow As Long
    Dim str��׼·��Name As String, str�汾 As String, str��������� As String
    Dim lngParCount As Long
    Dim blnFont As Boolean, blnNameKey As Boolean
    
    On Error GoTo errH
    
    Set objWordApp = CreateObject("Word.Application")
    If objWordApp Is Nothing Then
        MsgBox "Word.Application����ʧ�ܣ�"
        Exit Function
    End If
    '�Ƿ��ܴ���word�ļ�
    Set objWord = objWordApp.Documents.Open(strFilePath, False, True, , , , , , , , , False)
    If objWord Is Nothing Then
        MsgBox "�ļ�" & strFilePath & "�򿪲��ɹ�,����·����������", vbInformation, "ϵͳ��Ϣ"
        Exit Function
    End If

    With vsAnalyse
    
        lngParCount = objWord.Paragraphs.Count
        If lngParCount = 0 Then Exit Function
        
        prgImp(3).Max = lngParCount
        prgImp(3).Value = 0
        i = 1
        
        Do
            str��׼·��Name = ""
            str�汾 = ""
            '�ж��Ƿ��ҵ�����
            Set rngFind = objWord.Paragraphs(i).Range
            str��������� = rngFind.Font.Name & "," & rngFind.Font.Size & "," & IIf(rngFind.Font.Bold = -1, 1, 0)
            If str��������� = mstrFontStr����� Then
                blnFont = True
                str��׼·��Name = rngFind.Text
                If InStr(str��׼·��Name, "�ٴ�·��") > 0 Then
                    blnNameKey = True
                    str��׼·��Name = Trim(Replace(Replace(Replace(str��׼·��Name, " ", ""), Chr(13), ""), Chr(12), ""))
                    str�汾 = objWord.Paragraphs(i + 1).Range.Text
                    If InStr(str�汾, "��") > 0 Then
                        str�汾 = Trim(Replace(Replace(Replace(str�汾, "��", ""), "��", ""), Chr(13), ""))
                    Else
                        str�汾 = ""
                    End If
                    BlnFind = True
                Else
                    str��׼·��Name = ""
                End If
            End If
            '�ҵ����������
            If BlnFind Then
                lng������ʼֵ = lng������ʼֵ + 1
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                .TextMatrix(lngRow, AC_·������) = str��׼·��Name
                .TextMatrix(lngRow, AC_�汾) = IIf(str�汾 = "", str�汾In, str�汾)
                .TextMatrix(lngRow, AC_���Ŀ�ʼ) = i + 2
                If .TextMatrix(lngRow - 1, AC_���Ľ���) = "" And lngRow <> .FixedRows Then '��һ���Ѿ���ֵ�Ͳ����и�ֵ
                    .TextMatrix(lngRow - 1, AC_���Ľ���) = i - 1
                End If
                .TextMatrix(lngRow, AC_���⿪ʼ) = i
                .TextMatrix(lngRow, ac_����) = str���� & lng������ʼֵ
                .TextMatrix(lngRow, AC_����) = str����
                .TextMatrix(lngRow, AC_�ļ�·��) = strFilePath
                BlnFind = False
                prgImp(3).Value = i
            End If
            i = i + 1
        Loop While i < lngParCount
        
        .TextMatrix(lngRow, AC_���Ľ���) = lngParCount  '�޸����һ�е�����
        If str��׼·��Name = "" Then
            If blnFont Then
                If Not blnNameKey Then
                    If MsgBox("���롾" & str�ļ����� & "���ĵ��Ĵ���ⲻ�����ٴ�·�����ؼ��֣��Ƿ������)", vbYesNo, gstrSysName) = vbNo Then
                        AnalyseDoc = False
                        Exit Function
                    End If
                End If
            Else
                If MsgBox("���롾" & str�ļ����� & "���ĵ��Ĵ�������������õġ�" & mstrFontStr����� & "����һ��,�Ƿ������", vbYesNo, gstrSysName) = vbNo Then
                    AnalyseDoc = False
                    Exit Function
                End If
            End If
       End If
    End With
    
    Set objWord = Nothing
    Call objWordApp.Quit
    Set objWordApp = Nothing
    blnFont = False
    blnNameKey = False
    AnalyseDoc = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, "ϵͳ��Ϣ"
    If 0 = 1 Then
        Resume
    End If
    err.Clear
End Function

Private Sub ClearPage(ByVal intPage As Integer)
'���ܣ����ָ��ҳ�������
'       intPage:��ǰ�ɼ�ҳ���index,-1ʱ�����������������·����������������
    With vsDefineImp
        '�����һ�еĵ������
        .colData(DC_��������) = .TextMatrix(.FixedRows, DC_��������)
        .colData(DC_����ǰ׺) = .TextMatrix(.FixedRows, DC_����ǰ׺)
        .colData(DC_�ָ���) = .TextMatrix(.FixedRows, DC_�ָ���)
        .colData(DC_������ʼֵ) = .TextMatrix(.FixedRows, DC_������ʼֵ)
        If intPage = PE_DefineImp Or intPage = -1 Then
            '�������
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End If
    End With
    
    With vsAnalyse
        If intPage = PE_AnaResult Or intPage = -1 Then
            '�������
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End If
    End With
    
    With vsErrInfo
        If intPage = PE_ErrInfo Or intPage = -1 Then
            '�������
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End If
    End With
End Sub

Private Function GetStepPre(ByVal intPage As Integer) As Integer
'���ܣ����ݵ�ǰ�����ȡ��һҳ��
'      intPage :��ǰ�ɼ�ҳ���index
'���� :��һҳ���Index
    If intPage = PE_PathInput Then
        GetStepPre = intPage
    Else
        GetStepPre = intPage - 1
    End If
End Function

Private Sub cmdPathCancel_Click()
    fraFloder.Visible = False
    '������ǰҳ��
    Call SetContolStat(PE_PathInput, Not fraFloder.Visible)
    cmdImport(0).Enabled = False
    cmdImport(0).Visible = False
End Sub

Private Sub cmdPathOk_Click()
    txtFile.Text = dirFloder.List(dirFloder.ListIndex)
    fraFloder.Visible = False
    '������ǰҳ��
    Call SetContolStat(PE_PathInput, Not fraFloder.Visible)
    cmdImport(0).Enabled = False
    cmdImport(0).Visible = False
End Sub


Private Sub cmdSize_Click(Index As Integer)
    With dlgCom
        .Flags = &H80000 + &H100000
        .ShowFont
        If .FontSize <> 0 Then
            txtInfo(Index * 2 + 1).Text = .FontSize
        End If
        
        If .FontName <> "" Then
            txtInfo(Index * 2).Text = .FontName
        End If
        
        If .FontBold Then
            chkBold(Index).Value = 1
        End If
    End With
End Sub

Private Sub div_Change()
    dirFloder.Path = div.Drive
End Sub

Private Sub Form_Load()
    If mintPage = 0 Then cmdImport(0).Visible = False
End Sub
Private Sub SetVsStyle()
'���ܣ������������ñ����ĵ�Ԫ��ĸ߶�����,�Լ�������ɫ�ȣ��Լ���Ԫ��ĺϲ���

    Dim i As Long, j As Long
    Dim lngmaxHeight As Long
   On Error GoTo errH
    With vsPathTable
        If .Rows = 0 And .Cols = 0 Then Exit Sub
        '�޸ķ������ƣ��׶Σ�����Ӵ־���
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = 4 '����
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = &HE1FFE1
        
        .AutoResize = False
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1, False, 0) '�Զ�������С
        '���ý׶����壬��ɫ�����뷽ʽ
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = "ʱ��" Then
                .Cell(flexcpAlignment, i, 0, i, .Cols - 1) = 4
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False '���üӴ�ǰҪ������Ӵ�
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE1FFE1
            Else
                If .Cols > 1 Then
                    .Cell(flexcpAlignment, i, 1, i, .Cols - 1) = 0
                End If
            End If
        Next
        
        '��ȡͬһ����ߵĵ�Ԫ��߶ȸ�ֵ���и�
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                For j = 0 To .Cols - 1
                    If j = 0 Then
                        lngmaxHeight = ComputerLines(.TextMatrix(i, j))
                    Else
                        lngmaxHeight = IIf(lngmaxHeight > ComputerLines(.TextMatrix(i, j)), lngmaxHeight, ComputerLines(.TextMatrix(i, j)))
                    End If
                Next
                .RowHeight(i) = IIf(lngmaxHeight = 0, 5, lngmaxHeight) * Me.TextHeight("��") * 1.5
            Else
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = " " 'Ϊ�˺ϲ���Ԫ��
                Next
            End If
        Next
        '�ָ��е�Ԫ��ϲ����Լ��߿���ɫ����
        .MergeCells = flexMergeFree
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = " " Then
                Call .CellBorderRange(i, 0, i, .Cols - 1, &HFFFFFF, 1, 0, 1, 0, 1, 0)
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HFFFFFF
                .MergeRow(i) = True
            End If
        Next
        'ʵ�������϶��п�
        .FixedRows = 1
        Call .CellBorderRange(0, 0, 0, .Cols - 1, &H8000&, 0, 0, 1, 1, 1, 1)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetContolStat(ByVal intPage As Integer, ByVal blnInput As Boolean)
'����:����ָ��ҳ��ؼ�״̬
'   intPage:���ý���
    Dim i As Long
    
    cmdImport(1).Enabled = blnInput
    cmdImport(2).Enabled = blnInput
    cmdImport(0).Enabled = blnInput
    cmdImport(0).Visible = True
    Select Case intPage
        Case PE_PathInput
            optSelect(0).Enabled = blnInput
            optSelect(1).Enabled = blnInput
            txtFile.Enabled = blnInput
            cmdBraw.Enabled = blnInput
        Case PE_DefineImp
            vsPathTable.Enabled = blnInput
        Case PE_AnaRules
            vsDefineImp.Enabled = blnInput
        Case PE_AnaResult
            vsAnalyse.Enabled = blnInput
    End Select
End Sub

Private Function LoadPath() As Boolean
'���ܣ�����·������·���������ݿ�

    Dim i As Long, lngCount As Long
    Dim strFilePath As String
    
    vsErrInfo.Visible = False
    vsErrInfo.Rows = vsErrInfo.FixedRows
    mlngImpFileCount = 0
    mlngImpPathCount = 0
    fraProcess(0).Visible = True
    With vsAnalyse
        prgImp(1).Max = mlngSelPathCount
        prgImp(1).Value = 0
        For i = .FixedRows To .Rows - 1
            prgImp(1).Value = mlngImpPathCount
            If .TextMatrix(i, DC_ѡ��) = "-1" Then
               If ImpSelPathByFile(.TextMatrix(i, AC_�ļ�·��), .TextMatrix(i, DC_��������), .TextMatrix(i, DC_�汾), .TextMatrix(i, ac_����), , True, Val(.TextMatrix(i, AC_���Ľ���)), Val(.TextMatrix(i, AC_���⿪ʼ))) Then
                    If strFilePath <> .TextMatrix(i, AC_�ļ�·��) Then
                        strFilePath = .TextMatrix(i, AC_�ļ�·��)
                        mlngImpFileCount = mlngImpFileCount + 1
                    End If
                End If
            End If
        Next
    End With
    If vsErrInfo.Rows <> vsErrInfo.FixedRows Then
        vsErrInfo.Visible = True
        fraProcess(0).Visible = False
        LoadPath = False
    Else
        LoadPath = True
        prgImp(1).Value = prgImp(1).Max
        prgImp(0).Value = 0
    End If
End Function

Private Function ImpSelPathByFile(ByVal strFilePath As String, ByVal str���� As String, ByVal strVerSionIn As String, ByVal strCode As String, _
            Optional ByVal lngCodeStart As Long, Optional ByVal blnAna As Boolean, Optional ByVal lng���Ľ��� As Long, Optional ByVal lng���⿪ʼ As Long) As Boolean
'���ܣ�����ѡ����ļ�����·��
'      strFilePath:�ļ���ȫ������·��)
'      str����:�������õĿ�������
'      strVerSionIn:�������õİ汾
'      strCode:�������õı���ǰ׺&�ָ���
'      lngCodeStart:�������õı�����ʼֵ
'      blnAna:�Ƿ񾭹�����,������������·������

    Dim objWord As Object, objWordApp As Object
    Dim rngTitle As Object, rngText As Object, rngTable As Object, rngTotal As Object, rngTableTitle As Object, rngTmp As Object
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, h As Long, l As Long
    Dim lngRows As Long, lngCols As Long, lngParCont As Long, lngCurRow As Long
    Dim strStPathName As String, strVerSion As String, strCodeCur As String, strFontStr As String
    Dim lngPathMark As Long, lngCoursNo As Long, lngStPathID As Long, lng�׶���� As Long, lng������� As Long
    Dim strTableTitle As String, str������ As String, str�������� As String, strCoursContent As String, strDiseaseCodes As String, strOpeCode As String
    Dim str�׶����� As String, str�������� As String, strTableContent As String
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim strTCD As String, strTmp As String   '��ҽ��������
    Dim arrSql As Variant
    
    On Error GoTo errH
    
    Set objWordApp = CreateObject("Word.Application")
    If objWordApp Is Nothing Then
        MsgBox "Word.Application����ʧ�ܣ�"
        Exit Function
    End If
    '�Ƿ��ܴ���word�ļ�
    Set objWord = objWordApp.Documents.Open(strFilePath, False, True, , , , , , , , , False)
    If objWord Is Nothing Then
        Call err.Raise(200000, "�ļ��򿪲��ɹ�", "�ļ�" & strFilePath & "�򿪲��ɹ�,����·����������")
    End If
    
    If objWord.Paragraphs.Count = 0 Then
        Call err.Raise(200000, "�ļ�û�а�������", "�ļ�" & objWord.Name & "�������κ�����")
    End If
    Set rngTotal = objWord.Paragraphs(1).Range
    If blnAna Then
        Call rngTotal.SetRange(objWord.Paragraphs(lng���⿪ʼ).Range.Start, objWord.Paragraphs(lng���Ľ���).Range.End)
    Else
        Call rngTotal.SetRange(0, objWord.Paragraphs(objWord.Paragraphs.Count).Range.End)
    End If
    lngParCont = rngTotal.Paragraphs.Count
    If lngParCont = 0 Then
        Call err.Raise(200001, "·����������Ϣ", objWord.Name & "����ѡ·����������Ϣ")
    End If
    
    prgImp(0).Max = lngParCont
    prgImp(0).Value = 0
    i = 1
    arrSql = Array()
    Do
        '�ж��Ƿ��ҵ�����
        Set rngTitle = rngTotal.Paragraphs(i).Range
        prgImp(0).Value = i
        If Trim(Replace(Replace(Replace(rngTitle.Text, " ", ""), Chr(13), ""), Chr(12), "")) <> "" Then
            strFontStr = rngTitle.Font.Name & "," & rngTitle.Font.Size & "," & IIf(rngTitle.Font.Bold = -1, 1, 0)
            If strFontStr = mstrFontStr����� Then
                    If InStr(rngTitle.Text, "�ٴ�·��") > 0 Then
                        mlngImpPathCount = mlngImpPathCount + 1
                        '��ʼ����
                        strVerSion = ""
                        strStPathName = ""
                        strVerSion = ""
                        lngStPathID = 0
                        strCodeCur = ""
                        strSql = ""
                        '��ȡ·��������Ϣ
                        If Not blnAna Then
                            strCodeCur = strCode & lngCodeStart
                            lngCodeStart = lngCodeStart + 1
                        Else
                            strCodeCur = strCode
                        End If
                        strStPathName = Trim(Replace(Replace(Replace(rngTitle.Text, " ", ""), Chr(13), ""), Chr(12), ""))
                        For j = i + 1 To i + 10
                            If j <= rngTotal.Paragraphs.Count Then
                                Set rngText = rngTotal.Paragraphs(j).Range
                                strFontStr = rngText.Font.Name & "," & rngText.Font.Size & "," & IIf(rngText.Font.Bold = -1, 1, 0)
                                If strFontStr = mstrFontStr�������� Then i = j - 1: Exit For
                                If InStr(rngText.Text, "��") > 0 Then
                                    strVerSion = Trim(Replace(Replace(Replace(strVerSion, "��", ""), "��", ""), Chr(13), ""))
                                    i = j
                                    Exit For
                                End If
                            End If
                        Next
                        strVerSion = IIf(strVerSion = "", strVerSionIn, strVerSion)
                        
                        strSql = "select ID,��������,����,·������ from ��׼·��Ŀ¼ where ��������=[1] and ·������=[2] and �汾˵��=[3] and ����=[4]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str����, strStPathName, strVerSion, strCodeCur)
                        
                        If rsTmp.RecordCount = 0 Then
                            strSql = "Zl_��׼·��Ŀ¼_Insert(NULL,'" & str���� & "','" & strCodeCur & "','" & strStPathName & "','" & strVerSion & "',Null,Null)"
                            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                        Else
                            Call err.Raise(19999, "���ܲ����·��", "·����" & strStPathName & "���ܲ���,��顾��׼·��Ŀ¼�������Ƿ���ڡ���������, ����, ·������, �汾˵������ͬ������")
                        End If
                        
                        strSql = "select ID,��������,����,·������ from ��׼·��Ŀ¼ where ��������=[1] and ·������=[2] and �汾˵��=[3] and ����=[4]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str����, strStPathName, strVerSion, strCodeCur)

                        If rsTmp.RecordCount <> 0 Then
                            rsTmp.MoveFirst
                            lngStPathID = Val(rsTmp!ID & "")
                            mlngStPathID = lngStPathID
                        End If
                        '����������
                        lngPathMark = 1
                        lngCoursNo = 1
                    End If
            ElseIf strFontStr = mstrFontStr�������� Or InStr(rngTitle.Text, "�ٴ�·����") > 0 And InStr(rngTitle.Text, NumberToChar(lngPathMark) & "��") > 0 Then
                    '����Ǳ�,��·�����+1��lng·����Ǻ���Ϊ��1����׼·�����̣�2��·����1��3��·����2,....
                    If InStr(rngTitle.Text, "�ٴ�·����") > 0 Then
                        If InStr(rngTitle.Text, NumberToChar(lngPathMark) & "��") = 0 Then lngPathMark = lngPathMark + 1
                    End If
                    '�����
                    If lngPathMark > 1 Then
                        strSql = ""
                        str�������� = Trim(Replace(rngTitle.Text, Chr(13), ""))
                        str������ = Mid(str��������, InStr(str��������, NumberToStr(lngPathMark) & "��") + Len(NumberToStr(lngPathMark) & "��"))
                        str������ = Mid(str������, 1, InStr(str������, "�ٴ�·����") + 5)
                        '��ȡ�ôα�
                        For j = i + 1 To rngTotal.Paragraphs.Count
                            Set rngText = rngTotal.Paragraphs(j).Range
                            strFontStr = rngText.Font.Name & "," & rngText.Font.Size & "," & IIf(rngText.Font.Bold = -1, 1, 0)
                            
                            '��һ·��
                            If strFontStr = mstrFontStr����� Then
                                If InStr(rngText.Text, "�ٴ�·��") > 0 Then
                                    i = j - 1 '����굽��һ��·����ʼ��
                                    Exit For
                                End If
                            End If
                            '��һ��������������IF�ֿ���Ϊ�����Ч��)
                            If strFontStr = mstrFontStr�������� Then
                                If InStr(rngText.Text, "�ٴ�·����") > 0 Then
                                    i = j - 1 '����굽��һ��·������ʼ��
                                    Exit For
                                End If
                            End If
                        Next
                        Set rngTable = objWord.Range(rngTitle.End, rngText.Start)
                        '����ƶ���Χ�ڴ��ڱ���������ݽ���
                        If rngTable.Tables.Count <> 0 Then
                            Call rngText.SetRange(rngTable.Start, rngTable.Tables(1).Range.Start)
                            strTableTitle = ""
                            For k = 1 To rngText.Paragraphs.Count
                                Set rngTableTitle = rngText.Paragraphs(k).Range
                                If Trim(Replace(Replace(rngTableTitle.Text, Chr(13), ""), Chr(12), "")) <> "" Then
                                    strTableTitle = strTableTitle & rngTableTitle.Text
                                End If
                            Next
                            lng�׶���� = 1: lng������� = 1
                            '������ı�ͷ���� Zl_��׼·��Ŀ¼_Insertʱ,�Ѿ�����һ������ͷ����
                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = "Zl_��׼·����_Update(" & lngStPathID & "," & IIf(lngPathMark = 2, 1, 0) & ",'" & Trim(str������) & "','" & strTableTitle & "')"
                            
                            '���Ĭ������
                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = "Zl_��׼·����_ContentClear(" & lngStPathID & "," & Val(lngPathMark - 1) & ")"
         
                            
                            '��ȡ�������
                            For k = 1 To rngTable.Tables.Count
                                lngRows = rngTable.Tables(k).Rows.Count
                                lngCols = rngTable.Tables(k).Columns.Count
                                For m = 1 To rngTable.Tables(k).Columns.Count
                                    lng�׶���� = lng�׶���� + 1 '�׶α�ʶ��ʵ�ʽ׶����Ϊlng�׶����-j(��Ϊÿ����ĵ�һ�в������׶Σ�
                                    '�ϲ���Ԫ���ȡ
                                    str�׶����� = rngTable.Tables(k).Cell(1, m).Range.Text
                                    str�׶����� = Trim(Replace(str�׶�����, Chr(13) & Chr(7), ""))
                                    If m <> 1 Then
                                        For n = 2 To lngRows
                                            lng������� = n '�������1ʱ�������洢����ͷ
                                            str�������� = Trim(Replace(rngTable.Tables(k).Cell(n, 1).Range.Text, Chr(13) & Chr(7), ""))
                                            str�������� = Replace(Replace(str��������, " ", ""), Chr(13), "")
                                            If InStr(",��������¼,���λ�ʿǩ��,ҽʦǩ��,", "," & str�������� & ",") > 0 Then Exit For
                                            strTableContent = Trim(Replace(rngTable.Tables(k).Cell(n, m).Range.Text, Chr(13) & Chr(7), ""))
                                            If Len(Trim(strTableContent)) > 2000 Then
                                                Call err.Raise(19999, "·�����벻�ɹ�", "·����" & strStPathName & "û�в���ɹ�,��顾" & str�������� & "��-��" & str�׶����� & "���е������Ƿ񳬹���2000���ַ����ȣ�")
                                            End If
                                            '������е�·����Ŀ����
                                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                                            arrSql(UBound(arrSql)) = "Zl_��׼·����_ContentInsert(" & lngStPathID & "," & Val(lngPathMark - 1) & "," & _
                                                    lng������� & ",'" & str�������� & "','" & lng�׶���� - k & "','" & str�׶����� & "','" & strTableContent & "')"
                                        Next
                                    End If
                                Next
                            Next
                        End If
                        '��ʶ��һ��·���������
                        lngPathMark = lngPathMark + 1
                    End If
                ElseIf strFontStr = mstrFontStrС���� And lngPathMark = 1 Then
                    strSql = ""
                    strCoursContent = ""
    '                If lngCoursNo = 7 Then Stop
                    h = i
                    For j = i + 1 To rngTotal.Paragraphs.Count
                        Set rngText = rngTotal.Paragraphs(j).Range
                        strFontStr = rngText.Font.Name & "," & rngText.Font.Size & "," & IIf(rngText.Font.Bold = -1, 1, 0)
                        If strFontStr = lblInfo(2).Tag Or strFontStr = lblInfo(1).Tag Or strFontStr = lblInfo(0).Tag Or _
                            lngCoursNo > 6 And InStr(rngText.Text, "�ٴ�·����") > 0 And InStr(rngText.Text, NumberToChar(lngPathMark + 1) & "��") > 0 Then
                            i = j - 1 '����굽��һ��·��������Ŀ��ʼ��
                            Exit For
                        End If
                    Next
                    If h <> i And rngTotal.Paragraphs(h + 1).Range.Start <> rngTotal.Paragraphs(i).Range.End Then
                        If rngTmp Is Nothing Then Set rngTmp = rngTotal.Paragraphs(h + 1).Range
                        Call rngTmp.SetRange(rngTotal.Paragraphs(h + 1).Range.Start, rngTotal.Paragraphs(i).Range.End)
                        If rngTmp.Tables.Count <> 0 Then
                            If rngTotal.Paragraphs(h + 1).Range.Start <> rngTmp.Tables(1).Range.Start Then
                                Call rngText.SetRange(rngTotal.Paragraphs(h + 1).Range.Start, rngTmp.Tables(1).Range.Start)
                                If Trim(Replace(Replace(rngText.Text, Chr(13), ""), Chr(12), "")) <> "" Then '��س�
                                    strCoursContent = strCoursContent & rngText.Text
                                End If
                            End If
                            For j = 1 To rngTmp.Tables.Count
                                For m = 1 To rngTmp.Tables(j).Rows.Count
                                    For n = 1 To rngTmp.Tables(j).Columns.Count
                                        strCoursContent = strCoursContent & "        " & RPAD(rngTmp.Tables(j).Cell(m, n).Range.Text, " ", 15)
                                    Next
                                    strCoursContent = strCoursContent & vbNewLine
                                Next
                                If j < rngTmp.Tables.Count Then
                                    If rngTmp.Tables(j).Range.End <> rngTmp.Tables(j + 1).Range.Start Then
                                        Call rngText.SetRange(rngTmp.Tables(j).Range.End, rngTmp.Tables(j + 1).Range.Start)
                                        If Trim(Replace(Replace(rngText.Text, Chr(13), ""), Chr(12), "")) <> "" Then '��س�
                                            strCoursContent = strCoursContent & rngText.Text
                                        End If
                                    End If
                                Else
                                    If rngTmp.Tables(rngTmp.Tables.Count).Range.End <> rngTmp.End Then
                                        Call rngText.SetRange(rngTmp.Tables(rngTmp.Tables.Count).Range.End, rngTmp.End)
                                        If Trim(Replace(Replace(rngText.Text, Chr(13), ""), Chr(12), "")) <> "" Then '��س�
                                            strCoursContent = strCoursContent & rngText.Text
                                        End If
                                    End If
                                End If
                            Next
                            
                        Else
                            For j = 1 To rngTmp.Paragraphs.Count
                                Set rngText = rngTmp.Paragraphs(j).Range
                                If Trim(Replace(Replace(rngText.Text, Chr(13), ""), Chr(12), "")) <> "" Then '��س�
                                    strCoursContent = strCoursContent & rngText.Text
                                End If
                            Next
                        End If
                    End If
                    '�Կ�ʼ�����ö��������⴦��Ϊ�˻�ȡ��������
                    str�������� = Trim(Replace(rngTitle.Text, Chr(13), ""))
                    If Not (str�������� = "" And strCoursContent = "") Then
'                        If lngCoursNo = 11 Then Stop
                        If InStr(str��������, "���ö���") > 0 And lngCoursNo = 1 Or InStr(str��������, "����·����׼") > 0 And lngCoursNo = 2 And (strDiseaseCodes = "" Or strTCD = "" Or strOpeCode = "") Then
                            'strTCD ��ҽ��������
                            strTCD = "" 'Ϊ��������ʱ,��һ�ٴ�·����Ҫ���
                            If InStr(strCoursContent, "TCD") > 0 Then
                                strTmp = Replace(Replace(Mid(strCoursContent, InStr(strCoursContent, "TCD")), "��", ")"), "��", ":")
                                For l = 1 To UBound(Split(strCoursContent, "TCD"))
                                    strTmp = Mid(strTmp, InStr(strTmp, ":") + 1)
                                    If Len(strTmp) > 0 And InStr(strTmp, ")") > 0 Then
                                        strTCD = strTCD & "," & Mid(strTmp, 1, InStr(strTmp, ")") - 1)
                                    End If
                                Next
                                strTCD = Mid(strTCD, 2)
                            Else
                                strTCD = ""
                            End If
                            '��ȡ������������������
                            strDiseaseCodes = ""
                            If InStr(strCoursContent, "ICD-10") > 0 Then
                                strTmp = Replace(Replace(Mid(strCoursContent, InStr(strCoursContent, "ICD-10")), "��", ")"), "��", ":")
                                For l = 1 To UBound(Split(strCoursContent, "ICD-10"))
                                    strTmp = Mid(strTmp, InStr(strTmp, ":") + 1)
                                     If Len(strTmp) > 0 And InStr(strTmp, ")") > 0 Then
                                        strDiseaseCodes = strDiseaseCodes & "," & Mid(strTmp, 1, InStr(strTmp, ")") - 1)
                                    End If
                                Next
                                strDiseaseCodes = Mid(strDiseaseCodes, 2)
                                If Len(strDiseaseCodes) > 200 Then
                                    Call err.Raise(19999, "·�����벻�ɹ�", "·����" & strStPathName & "û�в���ɹ�,��顾" & str�������� & "���еļ�����������Ƿ񳬹���200���ַ����ȣ�")
                                End If
                            Else
                                strDiseaseCodes = ""
                            End If
                            strOpeCode = ""
                            If InStr(strCoursContent, "ICD-9") > 0 Then
                                strOpeCode = Replace(Replace(Mid(strCoursContent, InStr(strCoursContent, "ICD-9")), "��", ")"), "��", ":")
                                strOpeCode = Mid(strOpeCode, InStr(strOpeCode, ":") + 1)
                                If Len(strOpeCode) > 0 And InStr(strOpeCode, ")") > 0 Then
                                    strOpeCode = Mid(strOpeCode, 1, InStr(strOpeCode, ")") - 1)
                                    If Len(strOpeCode) > 100 Then
                                        Call err.Raise(19999, "·�����벻�ɹ�", "·����" & strStPathName & "û�в���ɹ�,��顾" & str�������� & "���е�������������Ƿ񳬹���100���ַ����ȣ�")
                                    End If
                                End If
                            Else
                                strOpeCode = ""
                            End If
                            ReDim Preserve arrSql(UBound(arrSql) + 1)
                            arrSql(UBound(arrSql)) = "Zl_��׼·������_Update(" & lngStPathID & ",'" & IIf(strTCD = "", "", strTCD & IIf(strDiseaseCodes = "", "", ",")) & strDiseaseCodes & "','" & strOpeCode & "')"
                        End If
                        If Len(Trim(strCoursContent)) > 4000 Then
                            Call err.Raise(19999, "·�����벻�ɹ�", "·����" & strStPathName & "û�в���ɹ�,��顾" & str�������� & "���������Ƿ񳬹���4000���ַ����ȣ�")
                        End If
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = "Zl_��׼·������_Insert(" & lngStPathID & "," & lngCoursNo & ",'" & Trim(str��������) & "','" & Trim(strCoursContent) & "')"
                        '��ʶ��һ��Ŀ�����
                        lngCoursNo = lngCoursNo + 1
                    End If
            End If
        End If
        i = i + 1
    Loop While i + 1 < lngParCont
    
    '�����ύ����
    For l = LBound(arrSql) To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(l)), Me.Caption)
    Next
    
    If vsErrInfo.Rows > 1 Then
        ImpSelPathByFile = False
    Else
        ImpSelPathByFile = True
    End If
    Set objWord = Nothing
    Call objWordApp.Quit(False)
    Set objWordApp = Nothing
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    If err.Number = 5941 Then
        If err.Number <> 0 Then
            err.Description = "·����" & strStPathName & "û�в���ɹ�,��顾" & Trim(str������) & "���еĵ�Ԫ�����Ƿ���ڶ��л��߶��е������"
        End If
    End If
    With vsErrInfo
        .Rows = .Rows + 1
        lngCurRow = .Rows - 1
        If Not objWord Is Nothing Then
            .TextMatrix(lngCurRow, EC_�ļ���) = objWord.Name
            .TextMatrix(lngCurRow, EC_·������) = strStPathName
            .TextMatrix(lngCurRow, EC_������Ϣ) = err.Description
        Else
            .TextMatrix(lngCurRow, EC_�ļ���) = objWord.Name
            .TextMatrix(lngCurRow, EC_·������) = err.Source
            .TextMatrix(lngCurRow, EC_������Ϣ) = err.Description
        End If
    End With
    err.Clear
    Set objWord = Nothing
    Call objWordApp.Quit
    Set objWordApp = Nothing
End Function

Private Function ComputerLines(ByVal strInput As String) As Long
'���ܣ����������ı��лس����ĸ���
'������  strInput   Ҫ����س������ַ���
'���أ�   �س����ĸ���

    Dim strTmp As String
    Dim Count  As Long, lngPos As Long, lngLen As Long
    
    lngPos = InStr(strInput, Chr(13))
    lngLen = Len(strInput)
    strTmp = strInput
    
    Do While lngPos <> 0
        If Trim(strTmp) = "" Then Exit Do
        If lngPos + 1 <= lngLen Then
            strTmp = Mid(strTmp, lngPos + 1)
            Count = Count + 1
            lngPos = InStr(strTmp, Chr(13))
            lngLen = Len(strTmp)
        End If
    Loop
    
    ComputerLines = Count + 2
    
End Function


Private Sub Form_Unload(Cancel As Integer)
    mintPage = 0
    mblnFile = False
    mlngSelFileCount = 0
    mlngSelPathCount = 0
    mlngImpFileCount = 0
    mlngImpPathCount = 0
    mstrFontStr����� = ""
    mstrFontStr�������� = ""
    mstrFontStrС���� = ""
    mstrFontStr���� = ""
    mlngStPathID = 0
End Sub
Private Sub vsAnalyse_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAnalyse
        If Not (Col = AC_ѡ�� And .TextMatrix(Row, AC_·������) <> "" And Row <> 0) Then
            Cancel = True
        End If
    End With
End Sub

Private Function HaveMoreStr(ByVal strSouce As String, ByVal strJudge As String)
'���ܣ��ж�strSouce�Ƿ�����������ϵ�strJudge
    If Len(strSouce) = Len(Replace(strSouce, strJudge, "")) + Len(strJudge) Then
        HaveMoreStr = False
    Else
        HaveMoreStr = True
    End If
End Function

Public Function ShowMe(frmParent As Object, ByRef lngId As Long) As Boolean
    Me.Show 1, frmParent
    lngId = mlngStPathID
    ShowMe = True
End Function

Private Sub vsDefineImp_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = DC_�ָ��� Then
        If Len(Trim(vsDefineImp.TextMatrix(Row, Col))) > 1 Then
            MsgBox "�ڡ�" & Row & "���зָ���ֻ����һλ��", vbInformation, gstrSysName
            zlControl.ControlSetFocus vsDefineImp
            Exit Sub
        End If
    End If
End Sub

Private Sub vsErrInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strMsg As String
    If Shift = 2 And KeyCode = vbKeyC Then
        Clipboard.Clear
        Debug.Print vsErrInfo.MouseRow & "-" & vsErrInfo.MouseCol
        If Not vsErrInfo.MouseRow < 0 And Not vsErrInfo.MouseCol < 0 Then
            strMsg = vsErrInfo.TextMatrix(vsErrInfo.MouseRow, vsErrInfo.MouseCol)
            Clipboard.SetText strMsg
        End If
    End If
End Sub
