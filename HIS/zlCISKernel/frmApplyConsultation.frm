VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmApplyConsultation 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������뵥"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   Icon            =   "frmApplyConsultation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   10500
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   6600
      TabIndex        =   61
      Top             =   2385
      Width           =   3645
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   930
         ScaleHeight     =   285
         ScaleWidth      =   2520
         TabIndex        =   62
         Top             =   15
         Width           =   2520
         Begin VB.ComboBox cboItem 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -45
            TabIndex        =   9
            Top             =   -25
            Width           =   2520
         End
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "������Ŀ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   0
         TabIndex        =   63
         Top             =   45
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   17
         X1              =   855
         X2              =   3105
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   16
      Left            =   8970
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   15
      Left            =   5325
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   14
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1500
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   13
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   2000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5235
      Width           =   8550
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   56
      Top             =   2835
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   75
         Width           =   5175
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   885
         X2              =   6285
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H8000000E&
         Caption         =   "����ҽԺ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   57
         Top             =   75
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   1
      Left            =   3150
      Picture         =   "frmApplyConsultation.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "�༭(F4)"
      Top             =   2415
      Width           =   285
   End
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   480
      TabIndex        =   55
      Top             =   3285
      Width           =   9510
      Begin VSFlex8Ctl.VSFlexGrid vsDetail 
         CausesValidation=   0   'False
         Height          =   1305
         Left            =   15
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         Width           =   9495
         _cx             =   16757
         _cy             =   2293
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
         BackColorSel    =   16444122
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmApplyConsultation.frx":6948
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame fraIdea 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   2300
      Left            =   -60
      TabIndex        =   47
      Top             =   7515
      Width           =   10700
      Begin VB.CommandButton cmdInfo 
         Caption         =   "��"
         Height          =   270
         Index           =   2
         Left            =   9765
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   1755
         Width           =   270
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "��"
         Height          =   270
         Index           =   1
         Left            =   6810
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   1725
         Width           =   270
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   5295
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1500
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   1470
         TabIndex        =   19
         Text            =   "2013-06-20 18:00"
         Top             =   1770
         Width           =   1695
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Index           =   18
         Left            =   1470
         MaxLength       =   4000
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   8565
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   8745
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1000
      End
      Begin VB.CommandButton cmdDate 
         Height          =   285
         Index           =   0
         Left            =   3165
         Picture         =   "frmApplyConsultation.frx":69D3
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "�༭(F4)"
         Top             =   1740
         Width           =   285
      End
      Begin VB.Line Line2 
         BorderStyle     =   2  'Dash
         X1              =   -30
         X2              =   11130
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "�������ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   15
         Left            =   165
         TabIndex        =   54
         Top             =   1770
         Width           =   1260
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   1410
         X2              =   3135
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   5205
         X2              =   6810
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   4350
         TabIndex        =   53
         Top             =   1770
         Width           =   840
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   8670
         X2              =   9750
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "����ҽʦ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   7815
         TabIndex        =   52
         Top             =   1770
         Width           =   840
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   525
         TabIndex        =   51
         Top             =   195
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
      Height          =   270
      Index           =   0
      Left            =   9720
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   4770
      Width           =   270
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4800
      Width           =   8265
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Index           =   17
      Left            =   1410
      MaxLength       =   4000
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5640
      Width           =   8565
   End
   Begin VB.PictureBox picNo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   8320
      ScaleHeight     =   345
      ScaleWidth      =   1680
      TabIndex        =   26
      Top             =   1110
      Visible         =   0   'False
      Width           =   1680
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   60
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   34
         X1              =   255
         X2              =   1695
         Y1              =   285
         Y2              =   285
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H8000000E&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   37
         Left            =   0
         TabIndex        =   28
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   405
      Index           =   0
      Left            =   1440
      TabIndex        =   25
      Top             =   2745
      Width           =   1575
      Begin VB.OptionButton optInfo 
         BackColor       =   &H8000000E&
         Caption         =   "Ժ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optInfo 
         BackColor       =   &H8000000E&
         Caption         =   "Ժ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1640
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2025
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2025
      Width           =   1500
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   3930
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1640
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1640
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   3930
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2025
      Width           =   1000
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   19
      Left            =   1425
      TabIndex        =   4
      Text            =   "2013-06-20 18:00"
      Top             =   2430
      Width           =   1725
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1640
      Width           =   1000
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Index           =   1
      Left            =   4785
      TabIndex        =   1
      Top             =   2310
      Width           =   1575
      Begin VB.OptionButton optInfo 
         BackColor       =   &H8000000E&
         Caption         =   "��ͨ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optInfo 
         BackColor       =   &H8000000E&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   10275
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   251330561
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VB.Line Line1 
      Index           =   33
      X1              =   1335
      X2              =   2925
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   36
      Left            =   480
      TabIndex        =   46
      Top             =   7125
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   8865
      X2              =   9990
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "����ҽʦ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   34
      Left            =   8010
      TabIndex        =   45
      Top             =   7125
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   5220
      X2              =   6330
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "����ҽʦ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   33
      Left            =   4335
      TabIndex        =   44
      Top             =   7125
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   1350
      X2              =   9990
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "����Ŀ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   17
      Left            =   450
      TabIndex        =   43
      Top             =   5235
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   1350
      X2              =   9705
      Y1              =   5025
      Y2              =   5025
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "����ժҪ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   13
      Left            =   450
      TabIndex        =   42
      Top             =   5640
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "�ٴ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   19
      Left            =   450
      TabIndex        =   41
      Top             =   4800
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   6405
      X2              =   8040
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   8895
      X2              =   10005
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   6405
      X2              =   7530
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   3825
      X2              =   4935
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "���ﷶΧ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   39
      Top             =   2880
      Width           =   855
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   1335
      X2              =   2430
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1335
      X2              =   2430
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   3825
      X2              =   4935
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4185
      X2              =   6015
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label lblHead 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "�������뵥"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4170
      TabIndex        =   38
      Top             =   750
      Width           =   1875
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   37
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   480
      TabIndex        =   36
      Top             =   2025
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   5550
      TabIndex        =   35
      Top             =   2025
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   34
      Top             =   1640
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ס Ժ ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2940
      TabIndex        =   33
      Top             =   2025
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   5550
      TabIndex        =   32
      Top             =   1635
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2940
      TabIndex        =   31
      Top             =   1635
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "����ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   35
      Left            =   480
      TabIndex        =   30
      Top             =   2430
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   1320
      X2              =   3135
      Y1              =   2685
      Y2              =   2685
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��    ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   8055
      TabIndex        =   29
      Top             =   1635
      Width           =   840
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   75
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmApplyConsultation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event FormUnload(Cancel As Integer)


Private Enum mCtlID '�����ϵĿؼ�����ֵ
    txtNO = 6
    txt���� = 2
    txt�Ա� = 3
    txt���� = 7
    txt���� = 0
    txt���� = 5
    txtסԺ�� = 1
    txt���� = 4
    
    txt����ʱ�� = 19
    cmd����ʱ�� = 1
    
    fra�������� = 1
    opt��ͨ = 3
    opt���� = 2
    
    fra������Ŀ = 3
    
    fra���ﷶΧ = 0
    optԺ�� = 1
    optԺ�� = 0
    
    fra����ҽԺ = 2
    txt����ҽԺ = 12
    
    txt�ٴ���� = 8
    cmd�ٴ���� = 0
    
    txt����Ŀ�� = 13
    txt����ժҪ = 17
    txt������� = 14
    txt����ҽʦ = 15
    txt����ҽʦ = 16
    txt������� = 18
    txt���ʱ�� = 9
    cmd���ʱ�� = 0
    txt������� = 10
    cmd������� = 1
    txt����ҽʦ = 11
    cmd����ҽʦ = 2
End Enum

Private Enum VSDETAIL_COL
    COL_������� = 0
    COL_�������
    COL_ҽ������
    COL_����ҽ��
    
    COL_�������ID
    COL_����ҽ��ID
End Enum

Private mobjVBA As Object
Private mobjScript As clsScript
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mobjReport As Object
Private mclsDiagEdit As zlMedRecPage.clsDiagEdit
Private mclsMipModule As zl9ComLib.clsMipModule '��Ϣƽ̨����
Private mbln��¼ As Boolean
Private mbln����� As Boolean

Private mstr�ϴ�ת��ʱ�� As String
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlngAdviceID As Long 'һ����������е�ĳһ��ҽ��
Private mlngNo As Long ' �������
Private mlng���˿���id As Long
Private mlng����ID As Long
Private mlng��������ID As Long
Private mstr���IDs As String
Private mintPState As Integer
Private mdatTurn As Date
Private mint���� As Integer
Private mintType As Integer '0��������1���޸ģ�2���鿴��3����д�������
Private mintPos As Integer
Private mstr��Ժʱ�� As String
Private mrsItem���� As ADODB.Recordset
Private mlng��ĿID As Long
Private mbln��дҪ�� As Boolean 'true-ֻдһ��
Private mblnReturn As Boolean
Private mbln��� As Boolean '�Ƿ���ʾ�������
Private mbln���Ѷ��� As Boolean
Private mlngִ�п���ID As Long '���ﲡ��ʱ��ִ�п���ID
Private mlng���ͺ� As Long
Private mstr�����Ժ��� As String
Private mstrժҪ As String 'ժҪ���� gclsInsure.GetItemInfo ��ȡ
Private mstr�ѱ� As String
Private mlngǰ��ID As Long
Private mbytBaby As Byte  'Ӥ�����
'ҽ���༭����Ŀ�Ƭ�ؼ���ѡ��״̬
Private mrsCard As ADODB.Recordset
Private mstr_Ctl_ҽ������ As String            '����Ŀ�ı���ΪĬ��

Public Function ShowMe(ByRef frmParent As Object, Optional ByRef lngҽ��ID As Long, Optional ByRef lngNo As Long, Optional ByVal intType As Integer = 2, _
    Optional ByVal intPos As Integer, Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal lng����id As Long, _
    Optional ByVal lng��������ID As Long, Optional ByVal lng����ID As Long, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, _
    Optional ByRef objMip As Object, Optional ByVal lng��Ŀid As Long, Optional ByVal rsCard As ADODB.Recordset, Optional ByVal lngǰ��ID As Long, Optional ByVal bytBaby As Byte) As Boolean
'���ܣ������ӿڷ���
'������intType �������� 0��������1���޸ģ�2���鿴��3����д���������Ĭ��Ϊ�鿴״̬2
'      strDefine ҽ�����ݸ�ʽ����intPState ����״̬��intPos����λ�� 0��ҽ��վ�����棬1��ҽ���������
'      lngNo �������
'      lng��ĿID-ҽ���༭�������������ĿID
'      rsCard��ҽ����Ƭ��ҽ��������Ϣ��ֻ����ҽ���༭�����ҽ�������ұߵ�������ťʱ�Żᴫ�롣
'�������ò���˵����lngNo������ʱ��Ϊ�����������������������䣻lngҽ��ID�������޸�ʱ��Ϊ����������������������
    mintType = intType
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mintPos = intPos
    mlngAdviceID = lngҽ��ID
    mlng���˿���id = lng����id
    mlng����ID = lng����ID
    mlng��������ID = lng��������ID
    mlngǰ��ID = lngǰ��ID
    mbytBaby = bytBaby
    mintPState = intPState
    mdatTurn = datTurn
    mlngNo = lngNo
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    mlng��ĿID = lng��Ŀid
    Set mrsCard = rsCard
    mblnOK = False
    On Error Resume Next
    
    If intPos = 0 And (intType = 0 Or intType = 1 Or intType = 3) Then
        Me.Show , frmParent
    Else
        Me.Show 1, frmParent
    End If
    
    On Error GoTo 0
    
    If mblnOK Then
        lngҽ��ID = mlngAdviceID
        lngNo = mlngNo
    End If
    
    ShowMe = mblnOK
    
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Asc("'") = KeyAscii Or Asc(";") = KeyAscii Or Asc("%") = KeyAscii Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("��ǰ���뵥�Ѿ������˵�����δ���棬�Ƿ�Ҫ�����˳���", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Me.Tag = IIF(mblnOK, "1", "")
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
    If Not mclsDiagEdit Is Nothing Then Set mclsDiagEdit = Nothing
 
    mbln��¼ = False
    mbln����� = False
    mstr��Ժʱ�� = ""
    mstr�ϴ�ת��ʱ�� = ""
    mint���� = 0
    Set mclsMipModule = Nothing
    RaiseEvent FormUnload(Cancel)
End Sub

Private Sub Form_Load()
    
    Call InitCommandBar
    
    mbln��дҪ�� = Val(zlDatabase.GetPara(237, glngSys)) = 1
    mstr�����Ժ��� = zlDatabase.GetPara("Ҫ��������Ժ���", glngSys, pסԺҽ���´�)
    Call InitVSDetail
    
    Call LoadItem����
    
    Call LoadData
    
    Call SetFormState
    
    mblnChange = False
End Sub

Private Sub SetFormState()
'���ܣ����ý���״̬
    Select Case mintType
    Case 0, 1
        Me.Height = 7900
        Call SetItemEditable(1, 1, 1, 1, 1, 1, 1, 1, 1)
        picNo.Visible = False
    Case 2
        Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1)
        picNo.Visible = True
    Case 3
        Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1, 1, 1, 1, 1)
        picNo.Visible = True
    Case 4
        Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1)
        picNo.Visible = True
    End Select
End Sub

Private Sub InitVSDetail()
'���ܣ���ʼ������������ϸ���
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant
    
    
    If mbln��дҪ�� Then
        strHead = "�������,1040,4;�������,4560,1;ҽ������,1650,1;����ҽ��,1650,1;�������ID;����ҽ��ID"
    Else
        strHead = "�������;�������,5600,1;ҽ������,1650,1;����ҽ��,1650,1;�������ID;����ҽ��ID"
    End If
    
    arrHead = Split(strHead, ";")
    
    With vsDetail
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .ColData(COL_ҽ������) = "    |סԺҽʦ|����ҽʦ|(��)����ҽʦ"
    End With
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub vsDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String
    If mintType = 2 Then Exit Sub
    If NewRow < 0 Then Exit Sub
    With vsDetail
        If (NewCol = COL_������� Or NewCol = COL_����ҽ��) Then
            If optInfo(optԺ��).value Then
                .Editable = flexEDKbdMouse
                .ComboList = "..."
                .FocusRect = flexFocusLight
            Else
                .Editable = flexEDKbdMouse
                .ComboList = ""
                .FocusRect = flexFocusLight
            End If
        ElseIf NewCol = COL_ҽ������ Then
            .Editable = flexEDKbdMouse
            .FocusRect = flexFocusSolid
            .ComboList = .ColData(COL_ҽ������)
        Else
            .Editable = flexEDNone
            .FocusRect = flexFocusNone
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call InputDepOrDoc(Row, Col)
End Sub

Private Sub vsDetail_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�ɾ��������
    Dim i As Long
    
    If mintType = 2 Then Exit Sub
    If KeyCode = vbKeyDelete Then
        With vsDetail
            If .TextMatrix(.Row, COL_�������) <> "" Then
                If MsgBox("Ҫɾ����ǰ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                .RowData(.Row) = 0
                For i = 0 To .Cols - 1
                    .TextMatrix(.Row, i) = ""
                    .Cell(flexcpData, .Row, i) = ""
                Next
                mblnChange = True
            End If
            If Not (.Rows = .FixedRows + 1 And .Row = .FixedRows) Then .RemoveItem .Row
        End With
    End If
    If KeyCode > 127 Then
        Call vsDetail_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsDetail_DblClick()
    Call vsDetail_KeyPress(32)
End Sub

Private Sub vsDetail_GotFocus()
    Call vsDetail_AfterRowColChange(vsDetail.Row, vsDetail.Col, vsDetail.Row, vsDetail.Col)
End Sub

Private Sub vsDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ��ǻس�ȷ�����༭�Ĵ���(����Text:=EditText,��ValidateEdit�¼��л�û��)
    If Not mblnReturn Then
        If Col = COL_������� Or Col = COL_����ҽ�� Then
            vsDetail.TextMatrix(Row, Col) = CStr(vsDetail.Cell(flexcpData, Row, Col))
            Call vsDetail_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
        ElseIf Col = COL_ҽ������ Then
            If Trim(vsDetail.TextMatrix(Row, Col)) <> "" Then
                vsDetail.RowData(Row) = 1
            End If
        End If
    End If
    If COL_������� = Col Then Call Set�������(Row)
    mblnChange = True
End Sub

Private Sub InputDepOrDoc(ByVal lngRow As Long, lngCol As Long, Optional ByVal strEditText As String)
'���ܣ������Ŀ��¼�룬 �������   ����ҽ�� �� strEditText Ϊ�ձ��ֱ�ӵ�İ�ť��
    Dim strSQL As String, strLike As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strLikeText As String
    Dim lng����id As String
    Dim strִҵ��� As String
    
    On Error GoTo errH
    
    If Not (lngCol = COL_������� Or lngCol = COL_����ҽ��) Then Exit Sub

    vPoint = zlControl.GetCoordPos(vsDetail.hwnd, vsDetail.CellLeft, vsDetail.CellTop)
    
    If lngCol = COL_������� Then
        strSQL = "select a.id,a.����,a.����,a.����" & _
            " from ���ű� a,��������˵�� b where a.id=b.����id and b.��������='�ٴ�' and b.������� in (2,3)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)"
    ElseIf lngCol = COL_����ҽ�� Then
    
        lng����id = Val(vsDetail.TextMatrix(lngRow, COL_�������ID))
        strִҵ��� = vsDetail.TextMatrix(lngRow, COL_ҽ������)
        
        If lng����id = 0 Then
            strSQL = "Select /*+ rule +*/ distinct a.Id, a.���, a.����, a.����" & vbNewLine & _
                "From ��Ա�� A, ������Ա B, ���ű� C, ��������˵�� D" & vbNewLine & _
                "Where a.Id = b.��Աid And b.����id = c.Id And c.Id = d.����id And d.�������� = '�ٴ�' And d.������� In (2, 3)" & vbNewLine & _
                "And (a.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or a.����ʱ�� Is Null) And (c.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or c.����ʱ�� Is Null)"
        Else
            strSQL = "select distinct a.id,a.���,a.����,a.���� from ��Ա�� a,������Ա b where a.id=b.��Աid And (a.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or a.����ʱ�� Is Null) and b.����id=[1]"
        End If
        
        If strִҵ��� = "סԺҽʦ" Then
            strSQL = strSQL & " and a.רҵ����ְ�� like '%ҽʦ'"
        ElseIf strִҵ��� = "����ҽʦ" Then
            strSQL = strSQL & " and a.רҵ����ְ��='����ҽʦ'"
        ElseIf strִҵ��� = "(��)����ҽʦ" Then
            strSQL = strSQL & " and a.רҵ����ְ�� like '%����ҽʦ'"
        End If
        
    End If
    
    '���ں�ģ���ҵ�������
    If strEditText <> "" Then
        '�Ż�
        strLike = gstrLike
        If Len(strEditText) < 2 Then strLike = ""
        
        If lngCol = COL_������� Then
            strLikeText = " And (A.���� Like [2] Or A.���� Like [3] or a.���� like [4]) order by a.����"
        Else
            strLikeText = " And (A.��� Like [2] Or A.���� Like [3] or a.���� like [4]) order by a.���"
        End If
    Else
        If lngCol = COL_������� Then
            strLikeText = " and a.id<>[5] order by a.����"
        Else
            strLikeText = "  order by a.���"
        End If
    End If
    
    strSQL = strSQL & strLikeText
    
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�������뵥", False, "", "", False, False, True, vPoint.X, vPoint.Y, vsDetail.CellHeight, blnCancel, False, True, _
         lng����id, UCase(strEditText) & "%", strLike & UCase(strEditText) & "%", strLike & UCase(strEditText) & "%", mlng��������ID)
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
        End If
        vsDetail.TextMatrix(lngRow, lngCol) = CStr(vsDetail.Cell(flexcpData, lngRow, lngCol))
        Call vsDetail_AfterRowColChange(-1, -1, lngRow, lngCol) '����ʹ��ť�ɼ�
        Exit Sub
    End If
    
    If lngCol = COL_������� Then
        Call Set�������(lngRow, rsTmp)
    ElseIf lngCol = COL_����ҽ�� Then
        Call Set����ҽ��(lngRow, rsTmp)
    End If
    vsDetail.RowData(lngRow) = 1
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Set�������(ByVal lngRow As Long)
'���ܣ���ϸ����� ������� �и�ѡ������
    Dim i As Long
    
    With vsDetail
        For i = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, i, COL_�������) = flexUnchecked
        Next
        .Cell(flexcpChecked, lngRow, COL_�������) = flexChecked
    End With
End Sub

Private Sub vsDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strTmp As String
    If Asc("'") = KeyAscii Or Asc(";") = KeyAscii Or Asc("%") = KeyAscii Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        mblnReturn = True '����ǰ��س�ȷ�ϱ༭
        If optInfo(optԺ��).value And (Col = COL_������� Or Col = COL_����ҽ��) Then
            KeyAscii = 0
            strTmp = vsDetail.EditText
            Call InputDepOrDoc(Row, Col, strTmp)
            vsDetail.EditText = vsDetail.TextMatrix(Row, Col)
            If vsDetail.TextMatrix(Row, Col) <> "" Then Call LocatedNextCell(Row, Col)
        End If
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDetail_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strTmp As String
    If optInfo(optԺ��).value And (Col = COL_������� Or Col = COL_����ҽ��) Then
        strTmp = vsDetail.EditText
        Call InputDepOrDoc(Row, Col, strTmp)
        If vsDetail.TextMatrix(Row, Col) <> "" Then
            mblnReturn = False
            Call LocatedNextCell(Row, Col)
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub Set�������(ByVal lngRow, ByVal rsTmp As ADODB.Recordset)
'���ܣ������������
    Dim i As Long
    
    With vsDetail
        '����ظ�����
        For i = .FixedRows To .Rows - 1
            If i <> lngRow And Val(.TextMatrix(i, COL_�������ID)) = Val(rsTmp!ID & "") Then
                MsgBox "�ÿ����Ѿ��������б��С�", vbInformation, gstrSysName
                .TextMatrix(lngRow, COL_�������) = CStr(.Cell(flexcpData, lngRow, COL_�������))
                Call vsDetail_AfterRowColChange(lngRow, COL_�������, lngRow, COL_�������) '����ʹ��ť�ɼ�
                .SetFocus
                Exit Sub
            End If
        Next
                   
        If Val(.TextMatrix(lngRow, COL_�������ID)) <> Val(rsTmp!ID & "") Then
            .TextMatrix(lngRow, COL_�������ID) = Val(rsTmp!ID & "")
            .TextMatrix(lngRow, COL_�������) = rsTmp!���� & ""
            .Cell(flexcpData, lngRow, COL_�������) = rsTmp!���� & ""
            
            .TextMatrix(lngRow, COL_����ҽ��) = ""
            .Cell(flexcpData, lngRow, COL_����ҽ��) = ""
            .TextMatrix(lngRow, COL_����ҽ��ID) = ""
        End If
        .TextMatrix(lngRow, COL_�������) = .Cell(flexcpData, lngRow, COL_�������)
        If .TextMatrix(lngRow, COL_�������) <> "" Then Call LocatedNextCell(lngRow, COL_�������)
    End With
    
End Sub

Private Sub Set����ҽ��(ByVal lngRow, ByVal rsTmp As ADODB.Recordset)
'���ܣ���������ҽ��
    With vsDetail
        If Val(.TextMatrix(lngRow, COL_����ҽ��ID)) <> Val(rsTmp!ID & "") Then
            .TextMatrix(lngRow, COL_����ҽ��ID) = Val(rsTmp!ID & "")
            .TextMatrix(lngRow, COL_����ҽ��) = rsTmp!���� & ""
            .Cell(flexcpData, lngRow, COL_����ҽ��) = rsTmp!���� & ""
        End If
        .TextMatrix(lngRow, COL_����ҽ��) = .Cell(flexcpData, lngRow, COL_����ҽ��)
        If .TextMatrix(lngRow, COL_����ҽ��) <> "" Then Call LocatedNextCell(lngRow, COL_����ҽ��)
    End With
End Sub

Private Sub LocatedNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ�������ϸ�������һ����Ԫ���λ����һ����Ԫ��
    Dim lngTmpRow As Long, lngTmpCol As Long
    Dim i As Long
    
    lngTmpRow = lngRow
    lngTmpCol = lngCol
    
    With vsDetail
        lngTmpCol = lngTmpCol + 1
        If lngTmpCol < .Cols Then
            If Not .ColHidden(lngTmpCol) Then
                .Col = lngTmpCol
                .Row = lngTmpRow
            Else
                Call LocatedNextCell(lngTmpRow, lngTmpCol)
            End If
        Else
            lngTmpCol = .FixedCols
            lngTmpRow = lngTmpRow + 1
            
            If lngTmpRow > .Rows - 1 Then
                i = .Rows - 1
                If Not (Trim(.TextMatrix(i, COL_�������)) = "" And Trim(.TextMatrix(i, COL_ҽ������)) = "" And Trim(.TextMatrix(i, COL_����ҽ��)) = "") Then
                    .Rows = .Rows + 1
                Else
                    lngTmpRow = lngTmpRow - 1
                    lngTmpCol = COL_����ҽ��
                End If
            End If
            
            .Col = lngTmpCol
            .Row = lngTmpRow
            
            If .ColHidden(lngTmpCol) Then Call LocatedNextCell(lngTmpRow, lngTmpCol)
        End If
        
        
        If .Rows - 1 = .Row Then
            i = .Rows - 1
            If Not (Trim(.TextMatrix(i, COL_�������)) = "" And Trim(.TextMatrix(i, COL_ҽ������)) = "" And Trim(.TextMatrix(i, COL_����ҽ��)) = "") Then
                .Rows = .Rows + 1
            End If
        End If
        
        Call .ShowCell(.Row, .Col)
        Call vsDetail_AfterRowColChange(.Row, .Col, .Row, .Col)
    End With
End Sub

Private Sub vsDetail_KeyPress(KeyAscii As Integer)
    If mintType = 2 Then Exit Sub
    With vsDetail
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call LocatedNextCell(.Row, .Col)
        Else
            If .Col = COL_������� Or .Col = COL_����ҽ�� Or COL_ҽ������ Then
                .Editable = flexEDKbdMouse
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDetail_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
                
                If .Col = COL_ҽ������ Then
                    .Editable = flexEDKbdMouse
                    .FocusRect = flexFocusSolid
                    .ComboList = .ColData(COL_ҽ������)
                End If
                
            ElseIf .Col = COL_������� Then
                Call Set�������(.Row)
            Else
                .Editable = flexEDNone
            End If
        End If
    End With
    mblnChange = True
End Sub

Private Sub vsDetail_LostFocus()
'ʧȥ�������հ���
    Dim i As Long
    With vsDetail
        For i = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(i, COL_�������) = "" Then .RemoveItem i
        Next
        .AddItem ""
    End With
End Sub

Private Sub vsDetail_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDetail.EditSelStart = 0
    vsDetail.EditSelLength = zlCommFun.ActualLen(vsDetail.EditText)
End Sub

Private Sub vsDetail_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ���ʼ�༭ʱ��Ϊû�а��»س�
    If mintType = 2 Then
        Cancel = True
        Exit Sub
    End If
    mblnReturn = False
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����")
        objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
        
        If mintType = 3 Or mintType = 4 Then
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MeetFinish, "��ɻ���(&F)")
            objControl.IconId = 225
            Set objControl = .Add(xtpControlButton, conMenu_Tool_MeetCancel, "ȡ�����(&C)")
            objControl.IconId = 3014
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_19", Me)
        Case conMenu_File_Preview: Call PrintApply(1)
        Case conMenu_File_Print: Call PrintApply(2)
        Case conMenu_Edit_Save  '����
            If CheckData = False Then Exit Sub
            If SaveData() Then
                mblnOK = True
            End If
        Case conMenu_Tool_MeetFinish
            If mblnChange Then
                If CheckData = False Then Exit Sub
                If SaveData() Then
                    mblnOK = True
                End If
            End If
            If ExecuteMeet(0) = True Then
                mblnOK = True
            End If
        Case conMenu_Tool_MeetCancel
            If ExecuteMeet(1) = True Then
                mblnOK = True
            End If
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'''''''''
    Dim blnVisible As Boolean
    
    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Save
            Control.Enabled = mblnChange
        Case conMenu_Tool_MeetFinish
            Control.Enabled = mintType = 3
        Case conMenu_Tool_MeetCancel
            Control.Enabled = mintType = 4
    End Select
    Control.Visible = blnVisible
End Sub

Private Sub cmdInfo_Click(Index As Integer)
    Select Case Index
    Case cmd�ٴ����
        Call GetPatiDiag
    Case cmd�������
        Call GetItem�������(1)
    Case cmd����ҽʦ
        Call GetItem����ҽʦ(1)
    End Select
End Sub

Private Sub GetPatiDiag()
'���ܣ���ȡ��ǰ���
    Dim str��� As String
    
    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, pסԺҽ��վ, mclsMipModule)
    End If
    
    If mclsDiagEdit.ShowDiagEdit(Me, mlngAdviceID, mlng����ID, mlng��ҳID, 2, mlng���˿���id, mstr���IDs, str���, 0, mlngAdviceID) Then
        txtInfo(txt�ٴ����).Text = str���
        Call SeekNextCtl
    End If
End Sub

Private Sub GetDefaultPatiDiag()
'���ܣ���ȡ����Ĭ�ϵ���ϣ�ҽ������ʱ����
'˵����ȡ��ԴΪ3��ҳ��д�ġ��������ҽ�ƣ�������ȡ��ҽ��Ժ��ϣ����Ϊ������ȡ��ҽ��Ժ���
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim bln��ҽ�� As Boolean, i As Long
    Dim strIDs As String, str��� As String
    
    On Error GoTo errH
    
    strSQL = "select a.id,a.������� as ����,a.������� as ���� from ������ϼ�¼ a where a.��¼��Դ=3 And NVL(A.�������,1) = 1 and a.����id=[1] and a.��ҳid=[2] order by a.�������,a.��ϴ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    If Not rsTmp.EOF Then
        rsTmp.Filter = "����>10" '������Ƿ��������ҽ���
        
        If Not rsTmp.EOF Then
            bln��ҽ�� = Sys.DeptHaveProperty(mlng��������ID, "��ҽ��")
            If Not bln��ҽ�� Then rsTmp.Filter = 0
        Else
            rsTmp.Filter = 0
        End If

        For i = 1 To rsTmp.RecordCount
            strIDs = strIDs & "," & rsTmp!ID
            str��� = str��� & "," & rsTmp!����
            rsTmp.MoveNext
        Next
        
        mstr���IDs = Mid(strIDs, 2)
        txtInfo(txt�ٴ����).Text = Mid(str���, 2)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SeekNextCtl() As Boolean
'���ܣ���λ����һ������Ŀؼ���
    Call zlCommFun.PressKey(vbKeyTab)
    SeekNextCtl = True
End Function

Private Sub LoadPatiInfo()
'���ܣ���ȡ���˻�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '��ȡ���������Ϣ
    If mbytBaby = 0 Then
        strSQL = "Select c.סԺ��,C.����,c.��ǰ���� as ����,C.�Ա�,C.����, B.���� As ����,c.��Ժ����, C.��Ժ���� As ��ǰ����,c.����,c.�ѱ�" & vbNewLine & _
            "From ���ű� B, ������ҳ C" & vbNewLine & _
            "Where C.��Ժ����id = B.Id And C.����id = [1] And C.��ҳid = [2]"
    Else
        strSQL = "Select c.סԺ��,Nvl(q.Ӥ������, c.����||'֮Ӥ'||q.���) as ����,null as ����,q.Ӥ���Ա� as �Ա�,Round(Decode(q.����ʱ��, Null, Sysdate, q.����ʱ��) - q.����ʱ��) || '��' as ����, B.���� As ����,c.��Ժ����, C.��Ժ���� As ��ǰ����,c.����,c.�ѱ�" & vbNewLine & _
            "From ���ű� B, ������ҳ C, ������������¼ Q" & vbNewLine & _
            "Where C.��Ժ����id = B.Id  And c.����id = q.����id And c.��ҳid = q.��ҳid And C.����id = [1] And C.��ҳid = [2] And q.��� = [3]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mbytBaby)
    
    If rsTmp.RecordCount > 0 Then
        txtInfo(txtסԺ��).Text = rsTmp!סԺ�� & ""
        txtInfo(txt����).Text = rsTmp!���� & ""
        txtInfo(txt�Ա�).Text = rsTmp!�Ա� & ""
        txtInfo(txt����).Text = rsTmp!���� & ""
        txtInfo(txt����).Text = rsTmp!��ǰ���� & ""
        txtInfo(txt����).Text = rsTmp!���� & ""
        txtInfo(txt����).Text = rsTmp!���� & ""
        mstr��Ժʱ�� = Format(rsTmp!��Ժ���� & "", "YYYY-MM-DD HH:mm")
        mint���� = Val(rsTmp!���� & "")
        mstr�ѱ� = rsTmp!�ѱ� & ""
    End If
    
    mstr�ϴ�ת��ʱ�� = Get�ϴ�ת������
    
    txtInfo(txt����ҽʦ).Text = UserInfo.����
    
    strSQL = "select ��Ϣֵ as ����ҽʦ from ������ҳ�ӱ� where ����id=[1] and ��ҳid=[2] and ��Ϣ��=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, "����ҽʦ")
    If Not rsTmp.EOF Then txtInfo(txt����ҽʦ).Text = rsTmp!����ҽʦ & ""
    
    If mintType = 0 Then Call GetDefaultPatiDiag
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadData()
'���ܣ����ز������ݺͲ���
    Dim strTmp As String, strSQL As String
    Dim rsTmp As ADODB.Recordset, rs���� As ADODB.Recordset, rs���� As ADODB.Recordset
    Dim strIDs As String, i As Integer, lngRow As Long
    Dim datCur As Date
    Dim str�������IDs As String
    Dim blnCanSave As Boolean
    Dim str��Чʱ�� As String
    Dim str������� As String
    
    Dim varArr As Variant
    
    On Error GoTo errH
    
    If Not mrsCard Is Nothing Then
        str�������IDs = mrsCard!�������IDs & ""
        blnCanSave = Val(mrsCard!�Ƿ񱣴� & "") = 1
        str��Чʱ�� = mrsCard!��Чʱ�� & ""
    End If
    datCur = zlDatabase.Currentdate
    txtInfo(txt���ʱ��).Text = ""
    If mintType > 0 Then
        If mlng����ID = 0 Or mlng��ҳID = 0 Or mlngNo = 0 Then
            strSQL = "select a.����ID,a.��ҳID,a.id,a.������� from ����ҽ����¼ a where a.id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID)
            mlngNo = Val(rsTmp!������� & "")
            mlng����ID = Val(rsTmp!����ID & "")
            mlng��ҳID = Val(rsTmp!��ҳID & "")
        End If
        
        If mintType = 3 Or mintType = 4 Then
            strSQL = "select a.���ͺ�,a.ִ�в���ID from ����ҽ������ a where a.ҽ��id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID)
            mlngִ�п���ID = Val(rsTmp!ִ�в���ID & "")
            mlng���ͺ� = Val(rsTmp!���ͺ� & "")
        End If
        
        txtInfo(txtNO).Text = mlngNo
        
        Call LoadPatiInfo
        
    Else
    
        Call LoadPatiInfo
        '����
        If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
            If mdatTurn <> CDate(0) Then datCur = mdatTurn - 1 / 24 / 60
            mbln��¼ = True
        End If
        
        txtInfo(txt����ʱ��).Text = Format(datCur, "YYYY-MM-DD HH:mm")
        
        txtInfo(txt����ҽʦ).Text = UserInfo.����
        
        strSQL = "select ���� as ������� from ���ű� where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��������ID)
        txtInfo(txt�������).Text = rsTmp!������� & ""
        
        Call GetDefaultPatiDiag
        
        If str�������IDs <> "" Then
            '���û������
            '���ؼ���Ĭ�ϵĻ����������
            varArr = Split(str�������IDs, ",")
            strSQL = "select ����,id from ���ű� where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�������IDs)
            With vsDetail
                .Rows = .FixedRows
                For i = 0 To UBound(varArr)
                    If Val(varArr(i)) <> 0 Then
                        .AddItem ""
                        lngRow = .Rows - 1
                        rsTmp.Filter = "id=" & Val(varArr(i))
                        .TextMatrix(lngRow, COL_�������) = rsTmp!���� & ""
                        .Cell(flexcpData, lngRow, COL_�������) = rsTmp!���� & ""
                        .TextMatrix(lngRow, COL_�������ID) = rsTmp!ID & ""
                    End If
                Next
            End With
        End If
        
        Exit Sub
    End If
    
    strSQL = "select a.id,to_char(a.��ʼִ��ʱ��,'yyyy-MM-dd hh24:mi') as ����ʱ��,a.������־,a.������ĿID,a.ִ�п���ID,a.��������ID,a.����ҽ��" & _
        " from ����ҽ����¼ a where a.�������=[1] order by a.���"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngNo)
    
    mlng��ĿID = Val(rsTmp!������ĿID & "")
    Call Cbo.Locate(cboItem, mlng��ĿID, True)
    mlng��������ID = Val(rsTmp!��������id & "")
    For i = 1 To rsTmp.RecordCount
        strIDs = strIDs & "," & Val(rsTmp!ID & "")
        rsTmp.MoveNext
    Next
    rsTmp.MoveFirst
    strIDs = Mid(strIDs, 2)
    strSQL = "select ҽ��id,��Ŀ,���� from ����ҽ������ where ҽ��id IN (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) order by ҽ��ID"
 
    Set rs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
    rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='��������ҽԺ'"
    If Not rs����.EOF Then
        If rs����!���� & "" <> "" Then
            fraInfo(fra����ҽԺ).Visible = True
            txtInfo(txt����ҽԺ).Text = rs����!���� & ""
        End If
    End If
    
    optInfo(optԺ��).value = rs����.EOF
    optInfo(optԺ��).value = Not rs����.EOF
    
    With vsDetail
        .Rows = .FixedRows + rsTmp.RecordCount
        lngRow = .FixedRows - 1
        
        For i = 1 To rsTmp.RecordCount
            lngRow = 1 + lngRow
            .RowData(lngRow) = 1
            
            rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='�����������'"
            .TextMatrix(lngRow, COL_�������) = rs����!���� & ""
            .Cell(flexcpData, lngRow, COL_�������) = rs����!���� & ""
            .TextMatrix(lngRow, COL_�������ID) = Val(rsTmp!ִ�п���ID & "")
            
            rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='����ҽ������'"
            
            .TextMatrix(lngRow, COL_ҽ������) = rs����!���� & ""
            
            rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='��������ҽ��'"
            .TextMatrix(lngRow, COL_����ҽ��) = rs����!���� & ""
            
            rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='����������'"
            If Not rs����.EOF Then
                If .TextMatrix(lngRow, COL_�������) = rs����!���� & "" Then
                    .Cell(flexcpChecked, lngRow, COL_�������) = flexChecked
                End If
                If mintType = 3 Then
                    mbln����� = Is�������(mlngAdviceID)
                End If
            Else
                .Cell(flexcpChecked, lngRow, COL_�������) = flexUnchecked
            End If
            
            If Val(rsTmp!ID & "") = mlngAdviceID Then
                optInfo(opt����).value = 1 = Val(rsTmp!������־ & "")
                rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='���ﷶΧ'"
                If Not rs����.EOF Then optInfo(optԺ��).value = rs����!���� & "" = "Ժ��"
                If optInfo(optԺ��).value Then
                    rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='��������ҽԺ'"
                    If Not rs����.EOF Then txtInfo(txt����ҽԺ).Text = rs����!���� & ""
                End If
                rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='����Ŀ��'"
                If Not rs����.EOF Then txtInfo(txt����Ŀ��).Text = rs����!���� & ""
                
                rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='����ժҪ'"
                If Not rs����.EOF Then txtInfo(txt����ժҪ).Text = rs����!���� & ""
                
                rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='�������'"
                If Not rs����.EOF Then txtInfo(txt�������).Text = rs����!���� & ""
                
                rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='�������ʱ��'"
                If Not rs����.EOF Then txtInfo(txt���ʱ��).Text = rs����!���� & ""
                
                rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='������ɿ���'"
                If Not rs����.EOF Then txtInfo(txt�������).Text = rs����!���� & ""
                
                rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='����ҽ��'"
                If Not rs����.EOF Then txtInfo(txt����ҽʦ).Text = rs����!���� & ""
                
                rs����.Filter = "ҽ��ID=" & Val(rsTmp!ID & "") & " and ��Ŀ='�������'"
                If Not rs����.EOF Then str������� = rs����!���� & ""
            
                txtInfo(txt����ʱ��).Text = rsTmp!����ʱ�� & ""
                txtInfo(txt����ҽʦ).Text = rsTmp!����ҽ�� & ""
                
                strSQL = "select ���� as ������� from ���ű� where id=[1]"
                Set rs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!��������id & ""))
                txtInfo(txt�������).Text = rs����!������� & ""
                
                If mintType = 3 Then
                    txtInfo(txt���ʱ��).Text = IIF(txtInfo(txt���ʱ��).Text = "", Format(datCur, "YYYY-MM-DD HH:mm"), txtInfo(txt���ʱ��).Text)
                    txtInfo(txt�������).Text = .TextMatrix(lngRow, COL_�������)
                    txtInfo(txt����ҽʦ).Text = UserInfo.����
                End If
            End If
            rsTmp.MoveNext
        Next
    End With
    
    '��ȡ���
    strTmp = ""
    mstr���IDs = GetAdviceDiag(mlngAdviceID, strTmp)
    txtInfo(txt�ٴ����).Text = strTmp
    If str������� <> "" Then
        txtInfo(txt�ٴ����).Text = str�������
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Is�������(ByVal lngҽ��ID As Long) As Boolean
'���ܣ���д�������ʱ�жϵ�ǰ�Ŀ����ǲ��Ǵ������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "select 1 from ����ҽ����¼ a,���ű� b where a.ִ�п���id=b.id and a.id=[1] and" & vbNewLine & _
        "exists (select 1 from ����ҽ������ c where c.ҽ��id =[1] and c.��Ŀ='����������'and b.����=c.����)"
        
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    
    If Not rsTmp.EOF Then
        Is������� = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadItem����()
'���ܣ����ػ�����Ŀ�������б����ֻ��һ����Ŀ�����������б�

    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
    
    strSQL = "Select id,����,ִ�п���,�Ƽ����� from ������ĿĿ¼ where ���='Z' and ��������='7' and (����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL) order by ����"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Set mrsItem���� = zlDatabase.CopyNewRec(rsTmp)
    
    If rsTmp.EOF Then
        MsgBox "δ�ҵ�������Ŀ�����ȵ�������Ŀ�����д�������Ŀ��", vbInformation, "��������"
        mblnChange = False
        Unload Me
        Exit Sub
    Else
        If rsTmp.RecordCount = 1 Then
            fraInfo(fra������Ŀ).Visible = False
            mlng��ĿID = Val(rsTmp!ID & "")
        Else
            fraInfo(fra������Ŀ).Visible = True
            With cboItem
                .Clear
                For i = 1 To rsTmp.RecordCount
                    .AddItem rsTmp!���� & ""
                    .ItemData(.ListCount - 1) = Val(rsTmp!ID & "")
                    
                    If Val(rsTmp!ID & "") = mlng��ĿID And mlng��ĿID <> 0 Then
                        .ListIndex = .ListCount - 1
                    End If
                    rsTmp.MoveNext
                Next
                If .ListIndex = -1 Then .ListIndex = 0
            End With
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboItem_Click()
    mlng��ĿID = Val(cboItem.ItemData(cboItem.ListIndex))
    If Visible Then mblnChange = True
End Sub

Private Sub cboItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SeekNextCtl
End Sub

Private Sub optInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call SeekNextCtl
End Sub

Private Sub GetItem�������(ByVal intType As Integer)
'���ܣ���ȡ�������
'������0 �ı��򰴻س���1 �㰴ť
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    Dim i As Integer, strDoctor As String
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(txt�������).Tag = txtInfo(txt�������).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(txt�������).Text = "" Then  '�൱�����������Ŀ
            txtInfo(txt�������).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
       
    strDoctor = txtInfo(txt����ҽʦ).Text
    strInput = Trim(UCase(txtInfo(txt�������).Text))    '�����ֵ����ǰ׺�ո�
    
    strSQL = "Select Distinct A.ID,A.����,A.���� as ����,A.���� From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) And a.Id = b.����id" & _
        IIF(intType = 0, " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])", "") & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) And B.��������='�ٴ�'" & _
        IIF(strDoctor <> "", " And a.id in (select x.����id from ������Ա X, ��Ա�� Y where x.��Աid=y.id and y.����=[3])", "") & _
        " and a.id<>[4] Order by A.����"
        
    vRect = zlControl.GetControlRect(txtInfo(txt�������).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(txt�������).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", strDoctor, mlng��������ID)
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("û���ҵ�ƥ��Ŀ���!", vbInformation, gstrSysName)
            txtInfo(txt�������).SetFocus
            zlControl.TxtSelAll txtInfo(txt�������)
            Exit Sub
        End If
    Else
        txtInfo(txt�������).Text = rsTmp!���� & ""
        txtInfo(txt�������).Tag = rsTmp!���� & ""
        txtInfo(txt�������).SetFocus
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItem����ҽʦ(ByVal intType As Integer)
'���ܣ���ȡ����ҽʦ
'������intType 0 ����ƥ�䣬1 �����ť
    Dim strInput As String, intIndex As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    intIndex = txt����ҽʦ
    
    If intType = 0 Then
        If txtInfo(intIndex).Tag = txtInfo(intIndex).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(intIndex).Text = "" Then '�൱�����������Ŀ
            txtInfo(intIndex).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
    
    strSQL = "Select A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        IIF(intType = 0, " And (A.��� Like [1] Or A.���� Like [2] Or A.���� Like [2])", "") & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.���"
    
    On Error GoTo errH
    
    strInput = Trim(UCase(txtInfo(intIndex).Text))
    vRect = zlControl.GetControlRect(txtInfo(intIndex).hwnd)
    
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҽ��", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(intIndex).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "δ�ҵ�ƥ�����Ա��", vbInformation, gstrSysName
        End If
        Call zlControl.TxtSelAll(txtInfo(intIndex))
        txtInfo(intIndex).SetFocus: Exit Sub
    Else
        txtInfo(intIndex).Text = rsTmp!���� & ""
        txtInfo(intIndex).Tag = rsTmp!���� & ""
        txtInfo(intIndex).SetFocus
        Call SeekNextCtl
        If Visible Then mblnChange = True
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
'�����¼�����ģ����
    If Asc("'") = KeyAscii Or Asc(";") = KeyAscii Or Asc("%") = KeyAscii Then
        KeyAscii = 0
    End If
    
    
    If KeyAscii = 13 Then
        Select Case Index
        Case txt�������
            Call GetItem�������(0)
        Case txt����ҽʦ
            Call GetItem����ҽʦ(0)
        End Select
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Select Case Index
        Case txt�������
            Call GetItem�������(1)
        Case txt����ҽʦ
            Call GetItem����ҽʦ(1)
        End Select
    End If
    
    Select Case Index
        Case txt����ʱ��, txt����Ŀ��, txt����ҽԺ, txt�ٴ����
            If KeyAscii = vbKeyReturn Then Call SeekNextCtl
    End Select
    
    Select Case Index
        Case txt�������, txt����Ŀ��, txt����ժҪ
            If KeyAscii = Asc("'") Then KeyAscii = 0
    End Select
    
End Sub

Private Sub cmdDate_Click(Index As Integer)
'���ܣ�ѡ������
    Dim lngIndex As Long
    
    If Index = cmd����ʱ�� Then
        lngIndex = txt����ʱ��
        
        dtpDate.Left = txtInfo(lngIndex).Left
        dtpDate.Top = cmdDate(Index).Top + cmdDate(Index).Height
        
    ElseIf Index = cmd���ʱ�� Then
    
        lngIndex = txt���ʱ��
        dtpDate.Left = cmdDate(Index).Left + cmdDate(Index).Width - dtpDate.Width
        dtpDate.Top = txtInfo(lngIndex).Top - dtpDate.Height + fraIdea.Top
    End If
    
    If IsDate(txtInfo(lngIndex).Text) Then
        dtpDate.value = CDate(txtInfo(lngIndex).Text)
    Else
        dtpDate.value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = lngIndex
    
    dtpDate.ZOrder 0
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
'���ںϷ��Լ��
    Dim strDate As String, intIndex As Integer
    
    intIndex = Val(dtpDate.Tag)
    
    If intIndex = txt����ʱ�� Then
        'ȡֵ
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '�ж�ʱ��Ϸ���
        If Not Check��ʼʱ��(strDate) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txtInfo(intIndex).Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txtInfo_Validate(intIndex, False) '��������
        txtInfo(intIndex).SetFocus
        If Visible Then mblnChange = True
    ElseIf intIndex = txt���ʱ�� Then
        'ȡֵ
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '�жϻ������ʱ��Ϸ���
        If Not Check���ʱ��(strDate, txtInfo(txt����ʱ��).Text) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txtInfo(intIndex).Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txtInfo_Validate(intIndex, False) '��������
        txtInfo(intIndex).SetFocus
        If Visible Then mblnChange = True
    End If
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    Dim intIndex As Integer
    
    If KeyAscii = vbKeyEscape Then
        intIndex = Val(dtpDate.Tag)
        If intIndex >= 0 Then txtInfo(intIndex).SetFocus
        dtpDate.Tag = ""
        dtpDate.Visible = False
    End If
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Select Case Index
        Case txt����ʱ��
            Call cmdDate_Click(cmd����ʱ��)
        Case txt���ʱ��
            Call cmdDate_Click(cmd���ʱ��)
        End Select
    End If
End Sub

Private Function Check��ʼʱ��(ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ��������Ŀ�ʼʱ���Ƿ�Ϸ�   �����棭����Чʱ��
'˵����
'1.��ʼʱ�䲻��С�ڲ��˵���Ժʱ��
'2.��ʼʱ�䲻��С�ڲ��˵�ת��ʱ��
'3.��ʼʱ�����С����ֹʱ��
'4.����¼��ʱ,��ʼʱ�䲻��С�ڵ�ǰʱ��֮ǰ30����(�Ӷ�������ɿ���ʱ����ڿ�ʼʱ��30����)
'5.��¼��ҽ����ʼʱ�䲻�ܴ��ڵ�ǰʱ�䣬ת�Ʋ�¼���ܴ���ת�ƿ�ʼʱ��
    Dim strInDate As String, blnOut As Boolean
        
    If Not IsDate(strStart) Then
        MsgBox "����Ļ���ʱ����Ч��", vbInformation, gstrSysName
        Exit Function
    End If

    strInDate = mstr��Ժʱ��
    If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
        strMsg = "ҽ���Ļ���ʱ�䲻��С�ڲ��˵���Ժʱ�� " & strInDate & " ��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    strInDate = ""
    If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
        strInDate = Format(mdatTurn, "yyyy-MM-dd HH:mm")
    ElseIf IsDate(mstr�ϴ�ת��ʱ��) Then
        strInDate = mstr�ϴ�ת��ʱ��
    End If
    If strInDate <> "" Then
        If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
            If Format(strStart, "yyyy-MM-dd HH:mm") >= strInDate Then
                strMsg = "ҽ���Ļ���ʱ��ӦС�ڲ���" & IIF(mintPState = ps���ת��, "ת��", IIF(mintPState = psԤ��, "Ԥ��Ժ", "��Ժ")) & "��ʱ�� " & strInDate & " ��"
                If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
                strMsg = "ҽ���Ļ���ʱ�䲻��С�ڲ��������ת��ʱ�� " & strInDate & " ��"
                If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
   
    Check��ʼʱ�� = True
End Function

Private Function Check���ʱ��(ByVal strDate As String, ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ����������ʱ��

    Dim strInDate As String, strDateType As String
    
    If Not IsDate(strDate) Then
        strMsg = "�����ʱ����Ч��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf IsDate(strStart) Then
        If Format(strDate, "yyyy-MM-dd HH:mm") < Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = "�������ʱ�䲻��С��ҽ������ʱ�䡣"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check���ʱ�� = True
End Function

Private Function Get�ϴ�ת������() As String
'���ܣ���ȡ�ϴ�ת��ʱ��
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select ��ʼʱ�� From ���˱䶯��¼ Where ��ʼʱ�� is Not NULL And ��ʼԭ��=3 And ����ID=[1] And ��ҳID=[2] Order by ��ʼʱ�� desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount > 0 Then Get�ϴ�ת������ = Format(rsTmp!��ʼʱ�� & "", "YYYY-MM-DD HH:mm")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub PrintApply(ByVal intType As Integer)
'���ܴ�ӡԤ�����뵥
'������intType:1-Ԥ����2-��ӡ
    '�ж������δ�������ȱ����ٴ�ӡ
    If mintType <> 2 Then
        If mblnChange Then
            If CheckData = False Then Exit Sub
            If SaveData() Then
                mblnOK = True
            End If
        Else
            '��������ã�����ҽ���Ƿ����
            If CheckData = False Then Exit Sub
        End If
    End If
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_19", Me, "ҽ��ID=" & mlngAdviceID, "�������=" & mlngNo, intType)
End Sub

Private Function CheckData() As Boolean
'���ܣ����������ȷ��
    Dim strIDs As String, strҽ������ As String, strMsg As String
    Dim lngTmp As Long, i As Integer
    Dim vMsg As VbMsgBoxResult
    Dim intCount As Integer
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim bln��ҽ As Boolean
    Dim str���� As String
    Dim strTmp As String
    Dim strTabAdvice As String
    Dim rsPrice As ADODB.Recordset
    
    If mintType < 2 Then
        'Call SeekNextControl  '�����ַ�ʽ�������71290
        '������������費ͬ�ؼ��Ľ��㣬ȷ��validata�¼���ִ�С�
        txtInfo(txt����ʱ��).SetFocus
        txtInfo(txt����Ŀ��).SetFocus
        
        '���ʱ��Ϸ���
        If Not Check��ʼʱ��(txtInfo(txt����ʱ��).Text) Then
        If txtInfo(txt����ʱ��).Enabled Then txtInfo(txt����ʱ��).SetFocus
            Exit Function
        End If
        
        '�ж��Ƿ��ǲ�¼ҽ��
        If DateDiff("n", CDate(txtInfo(txt����ʱ��).Text), CDate(zlDatabase.Currentdate)) > gint��¼��� _
            Or mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
            mbln��¼ = True
        Else
            mbln��¼ = False
        End If
        
        If optInfo(optԺ��).value And txtInfo(txt����ҽԺ).Text = "" Then
            MsgBox "��ѡ���Ժ��������д����ҽԺ��", vbInformation, Me.Caption
            txtInfo(txt����ҽԺ).SetFocus
            Exit Function
        End If
           
        '����������ȷ��
        If mbln��дҪ�� Then
             lngTmp = 0: strMsg = ""
             With vsDetail
                 For i = .FixedRows To .Rows - 1
                     If .Cell(flexcpChecked, i, COL_�������) = flexChecked Then
                         lngTmp = lngTmp + 1
                     End If
                     '���ֻ��һ������ʱ����Ĭ�Ϲ��ϲ���ʾ
                     If .TextMatrix(i, COL_�������) <> "" Then
                         intCount = intCount + 1
                     End If
                 Next
             End With
             
             If intCount = 1 Then
                 vsDetail.Cell(flexcpChecked, vsDetail.FixedRows, COL_�������) = flexChecked
             Else
                 If lngTmp = 0 Then
                     strMsg = "��ȷ�����������ң�"
                 ElseIf lngTmp > 1 Then
                     strMsg = "����������ֻ����һ����"
                 End If
             End If
             
             If strMsg <> "" Then
                 MsgBox strMsg, vbInformation, Me.Caption
                 vsDetail.SetFocus
                 Exit Function
             End If
        End If
        
        If txtInfo(txt����Ŀ��).Text = "" Then
            MsgBox "û����д����Ŀ�ġ�", vbInformation, Me.Caption
            txtInfo(txt����Ŀ��).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt����ժҪ).Text = "" Then
            MsgBox "û����д����ժҪ��", vbInformation, Me.Caption
            txtInfo(txt����ժҪ).SetFocus
            Exit Function
        End If
        
        '��ϼ��
        If InStr(mstr�����Ժ���, "Z") > 0 Then
            bln��ҽ = Sys.DeptHaveProperty(mlng���˿���id, "��ҽ��")
            str���� = IIF(bln��ҽ, "2,12", "2")
            If Not ExistsDiagNoses(mlng����ID, mlng��ҳID, str����) Then
                strMsg = "���˵���Ժ��ϻ�û�����룬�������벡�˵���Ժ������´�������롣"
            End If
            If strMsg <> "" Then
                MsgBox strMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '������
        strҽ������ = Get��Ŀ����
        
        strTmp = mlng��ĿID & "||2"
        mstrժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", strTmp)
            
        With vsDetail
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_�������ID)) <> 0 And optInfo(optԺ��).value Then
                    strIDs = strIDs & "," & mlng��ĿID & ":" & Val(.TextMatrix(i, COL_�������ID))
                    
                    strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as ��� From Dual"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", 2, mlng����ID, mlng��ҳID, mint����, 1, _
                         "Z", mlng��ĿID, mlng��������ID, UserInfo.����, Val(.TextMatrix(i, COL_�������ID)), 0, 0, 0, mstrժҪ)
                    
                    If Not rsTmp.EOF Then
                        strMsg = NVL(rsTmp!���)
                        If strMsg <> "" Then
                            Select Case Val(Split(strMsg, "|")(0))
                            Case 1 '��ʾ
                                If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    strMsg = "": Exit Function
                                End If
                            Case 2 '��ֹ
                                MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                                strMsg = "": Exit Function
                            End Select
                            strMsg = ""
                        End If
                    End If
                    
                    strTabAdvice = strTabAdvice & " Union ALL " & _
                        "select " & i & " as ID," & i & " as ���,-null as ���ID,'Z' as �������," & mlng��ĿID & " as ������ĿID," & _
                            mlng��ĿID & " as ������ĿID,1 As ����, 0 As ����,null as �걾��λ,null As ��鷽��," & _
                            "0 as ִ�б��,0 as �Ƽ�����, null As ��������," & Val(mrsItem����!ִ�п��� & "") & " As ִ������," & Val(.TextMatrix(i, COL_�������ID)) & " as ִ�п���id from dual"
 
                ElseIf .TextMatrix(i, COL_�������) <> "" Then
                    strIDs = strIDs & .TextMatrix(i, COL_�������)
                End If
            Next
            If strIDs = "" Then
                MsgBox "û����д������ҡ�", vbInformation, Me.Caption
                .SetFocus
                Exit Function
            End If
            If optInfo(optԺ��).value Then strIDs = ""
        End With
        
        strIDs = Mid(strIDs, 2)
     
        If gintҽ������ = 2 Then mbln���Ѷ��� = True
    
        strMsg = CheckAdviceInsure(mint����, mbln���Ѷ���, mlng����ID, 2, "", strIDs, strҽ������)
        
        If strMsg <> "" Then
            If gintҽ������ = 1 Then
                vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "Ҫ��������ҽ����", Me)
                If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                If vMsg = vbIgnore Then mbln���Ѷ��� = False
            ElseIf gintҽ������ = 2 Then
                MsgBox strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        strTabAdvice = Mid(strTabAdvice, 12)
        'ҽ���ܿ�ʵʱ���
        If mint���� <> 0 And strTabAdvice <> "" Then
            If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) Then
                If MakePriceRecord���뵥("52", mlng����ID, mlng��ҳID, strTabAdvice, strIDs, mstr�ѱ�, mlng��������ID, rsPrice) Then
                    If Not gclsInsure.CheckItem(mint����, 1, 0, rsPrice) Then
                        MsgBox "ҽ�������δͨ(ִ��Insure.CheckItem�ӿ�)�������´�Ļ������뵥���ܱ��档", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
    ElseIf mintType = 3 Then
        If Not Check���ʱ��(txtInfo(txt���ʱ��).Text, txtInfo(txt����ʱ��).Text) Then
            If txtInfo(txt���ʱ��).Enabled Then txtInfo(txt���ʱ��).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt�������).Text = "" Then
            MsgBox "û����д���������", vbInformation, Me.Caption
            txtInfo(txt�������).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt�������).Text = "" Then
            MsgBox "û����д������ҡ�", vbInformation, Me.Caption
            txtInfo(txt�������).SetFocus
            Exit Function
        End If
        
        If txtInfo(txt����ҽʦ).Text = "" Then
            MsgBox "û����д����ҽʦ��", vbInformation, Me.Caption
            txtInfo(txt����ҽʦ).SetFocus
            Exit Function
        End If
    End If
    CheckData = True
End Function

Private Function SaveData() As Boolean
'���ܣ���������
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim lngҽ��ID As Long, lngҽ����� As Long, lng������� As Long
    Dim i As Long, str���� As String, int���� As Integer, lngִ�п���ID As Long, lngҪ��ID As Long
    Dim int�Ƽ����� As Integer, intִ������ As Integer, int������� As Integer
    Dim datCur As Date, str����ʱ�� As String, strSQL As String, strSource As String
    Dim rsҪ�� As ADODB.Recordset, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    arrSQL = Array()
    
    strSQL = "Select b.Id As Ҫ��id, b.������ As ���� from ����������Ŀ B where b.������ in " & _
      " ('���ﷶΧ','��������ҽԺ','�����������','����ҽ������','����������','��������ҽ��'," & _
      " '�������','����Ŀ��','�������','�������ʱ��','������ɿ���','����ҽ��','����ժҪ')"
    Set rsҪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
    If mintType = 0 Or mintType = 1 Then
        
        txtInfo(txtNO).Text = mlngNo
        
        If mintType = 1 Then '������޸�ҽ��������Һ�����ҽ����ɵ�ǰ��
            txtInfo(txt����ҽʦ).Text = UserInfo.����
            strSQL = "select ���� as ������� from ���ű� where id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��������ID)
            txtInfo(txt�������).Text = rsTmp!������� & ""
        End If
        
        If mlngNo <> 0 Then
            strSQL = "select a.id from ����ҽ����¼ a where a.�������=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngNo)
            For i = 1 To rsTmp.RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & rsTmp!ID & ",1)"
                rsTmp.MoveNext
            Next
        End If
        
        If mbln��¼ Then
            int���� = 2
        Else
            If optInfo(opt����).value Then int���� = 1
        End If
        
        If mlngNo <> 0 Then
            lng������� = mlngNo
        Else
            lng������� = Get�������
            mlngNo = lng�������
        End If
        
        lngҽ����� = GetMaxAdviceNO(mlng����ID, mlng��ҳID, mbytBaby)
        
        mrsItem����.Filter = "ID=" & mlng��ĿID
        str���� = mrsItem����!���� & ""
        
        int�Ƽ����� = Val(mrsItem����!�Ƽ����� & "")
        intִ������ = Val(mrsItem����!ִ�п��� & "")
        
        datCur = zlDatabase.Currentdate
        str����ʱ�� = IIF(datCur > CDate(txtInfo(txt����ʱ��).Text), txtInfo(txt����ʱ��).Text, datCur)
        
        '��3����Ŀ����Ϊ��[ID],[���],[ִ�п���ID]
        strSource = "ZL_����ҽ����¼_Insert([0],NULL,[1],2," & mlng����ID & "," & mlng��ҳID & "," & mbytBaby & ",1,1,'Z'," & mlng��ĿID & ",NULL,NULL,NULL,1,'" & str���� & "',NULL," & _
            "NULL,'һ����',NULL,NULL,NULL,NULL," & ZVal(int�Ƽ�����) & ",[2]," & ZVal(intִ������) & "," & int���� & "," & _
            "To_Date('" & Format(txtInfo(txt����ʱ��).Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
            mlng���˿���id & "," & mlng��������ID & ",'" & UserInfo.���� & "'," & _
            "To_Date('" & Format(str����ʱ��, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
            "NULL," & ZVal(mlngǰ��ID) & ",NULL,0,NULL," & IIF(mstrժҪ = "", "null", "'" & mstrժҪ & "'") & ",'" & UserInfo.���� & "',Null,NULL,NULL,NULL," & lng������� & ")"
                
        With vsDetail
            '�������Ƶ�ȷ��
            str���� = ""
            If mbln��дҪ�� Then
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpChecked, i, COL_�������) = flexChecked Then
                        str���� = .TextMatrix(i, COL_�������)
                        Exit For
                    End If
                Next
            End If
            
            For i = .FixedRows To .Rows - 1
                lngִ�п���ID = Val(.TextMatrix(i, COL_�������ID))
                int������� = 0
                If optInfo(optԺ��).value Then lngִ�п���ID = mlng���˿���id
                If .TextMatrix(i, COL_�������) <> "" Then
                
                    lngҽ����� = lngҽ����� + 1
                    lngҽ��ID = zlDatabase.GetNextID("����ҽ����¼")
                    
                    strSQL = GetStrExcSQL(strSource, lngҽ��ID, lngҽ�����, ZVal(lngִ�п���ID))
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int������� = int������� + 1
                    lngҪ��ID = Get����Ҫ��ID(rsҪ��, "���ﷶΧ")
                    strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'���ﷶΧ',0," & int������� & "," & lngҪ��ID & ",'" & IIF(optInfo(optԺ��).value, "Ժ��", "Ժ��") & "',1)"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    If optInfo(optԺ��).value Then
                        int������� = int������� + 1
                        lngҪ��ID = Get����Ҫ��ID(rsҪ��, "��������ҽԺ")
                        strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'��������ҽԺ',0," & int������� & "," & lngҪ��ID & ",'" & txtInfo(txt����ҽԺ).Text & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    End If
                    
                    If str���� <> "" Then
                        int������� = int������� + 1
                        lngҪ��ID = Get����Ҫ��ID(rsҪ��, "����������")
                        strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'����������',0," & int������� & "," & lngҪ��ID & ",'" & str���� & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    End If
                    
                    int������� = int������� + 1
                    lngҪ��ID = Get����Ҫ��ID(rsҪ��, "�������")
                    strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'�������',0," & int������� & "," & lngҪ��ID & ",'" & txtInfo(txt�ٴ����).Text & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int������� = int������� + 1
                    lngҪ��ID = Get����Ҫ��ID(rsҪ��, "����Ŀ��")
                    strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'����Ŀ��',0," & int������� & "," & lngҪ��ID & ",'" & txtInfo(txt����Ŀ��).Text & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int������� = int������� + 1
                    lngҪ��ID = Get����Ҫ��ID(rsҪ��, "����ժҪ")
                    strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'����ժҪ',0," & int������� & "," & lngҪ��ID & ",'" & txtInfo(txt����ժҪ).Text & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int������� = int������� + 1
                    lngҪ��ID = Get����Ҫ��ID(rsҪ��, "�����������")
                    strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'�����������',0," & int������� & "," & lngҪ��ID & ",'" & .TextMatrix(i, COL_�������) & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int������� = int������� + 1
                    lngҪ��ID = Get����Ҫ��ID(rsҪ��, "����ҽ������")
                    strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'����ҽ������',0," & int������� & "," & lngҪ��ID & ",'" & .TextMatrix(i, COL_ҽ������) & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    int������� = int������� + 1
                    lngҪ��ID = Get����Ҫ��ID(rsҪ��, "��������ҽ��")
                    strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'��������ҽ��',0," & int������� & "," & lngҪ��ID & ",'" & .TextMatrix(i, COL_����ҽ��) & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    
                    '��Ϲ�����Ϣ
                    If mstr���IDs <> "" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Insert(" & lngҽ��ID & ",'" & mstr���IDs & "')"
                    End If
                End If
            Next
            mlngAdviceID = lngҽ��ID
        End With
    ElseIf mintType = 3 Then '��д�������
        int������� = 0
        lngҽ��ID = mlngAdviceID
        
        strSQL = "select a.��Ŀ,a.Ҫ��id,a.���� from ����ҽ������ a where a.ҽ��id=[1] order by a.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
        
        For i = 1 To rsTmp.RecordCount
            If rsTmp!��Ŀ & "" = "�������" Then Exit For
            int������� = int������� + 1
            strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'" & rsTmp!��Ŀ & "',0," & int������� & "," & rsTmp!Ҫ��ID & ",'" & rsTmp!���� & "'" & IIF(i = 1, ",1)", ")")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
            rsTmp.MoveNext
        Next
        
        int������� = int������� + 1
        lngҪ��ID = Get����Ҫ��ID(rsҪ��, "�������")
        strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'�������',0," & int������� & "," & lngҪ��ID & ",'" & txtInfo(txt�������).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        int������� = int������� + 1
        lngҪ��ID = Get����Ҫ��ID(rsҪ��, "�������ʱ��")
        strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'�������ʱ��',0," & int������� & "," & lngҪ��ID & ",'" & txtInfo(txt���ʱ��).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        int������� = int������� + 1
        lngҪ��ID = Get����Ҫ��ID(rsҪ��, "������ɿ���")
        strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'������ɿ���',0," & int������� & "," & lngҪ��ID & ",'" & txtInfo(txt�������).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        int������� = int������� + 1
        lngҪ��ID = Get����Ҫ��ID(rsҪ��, "����ҽ��")
        strSQL = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'����ҽ��',0," & int������� & "," & lngҪ��ID & ",'" & txtInfo(txt����ҽʦ).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
       
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    SaveData = True
    mblnChange = False
    
    If mintType = 0 Or mintType = 1 Then
        Call ZLHIS_CIS_001(mclsMipModule, mlng����ID, txtInfo(txt����).Text, txtInfo(txtסԺ��).Text, , 2, _
            mlng��ҳID, mlng����ID, , mlng���˿���id, "", , txtInfo(txt����).Text, _
            lngҽ��ID, IIF(int���� = 1, 1, 0), 1, "Z", "", UserInfo.����, Format(str����ʱ��, "yyyy-MM-dd HH:mm:ss"), mlng��������ID, "", , , "")
    End If
    
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetStrExcSQL(ByVal strSource As String, ParamArray arrInput() As Variant) As String
'���ܣ����� ZL_����ҽ����¼_Insert������䣬arrInput���� [ID],[���],[ִ�п���ID]
    Dim i As Integer, strTmp As String, strResult As String
    
    strResult = strSource
    For i = 0 To UBound(arrInput)
        strTmp = arrInput(i)
        strResult = Replace(strResult, "[" & i & "]", strTmp)
    Next
    GetStrExcSQL = strResult
End Function
 
Private Function Get��Ŀ����() As String
    mrsItem����.Filter = 0
    mrsItem����.Filter = "ID=" & mlng��ĿID
    Get��Ŀ���� = mrsItem����!���� & ""
End Function

Private Function Get����Ҫ��ID(ByRef rsIn As ADODB.Recordset, ByVal str���� As String) As Long
'���ܣ���ȡ����̶�Ҫ��
    Dim strSQL As String
    
    On Error GoTo errH
    
    rsIn.Filter = "����='" & str���� & "'"
    If Not rsIn.EOF Then
        Get����Ҫ��ID = Val(rsIn!Ҫ��ID & "")
    End If
  
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteMeet(ByVal intType As Integer) As Boolean
'���ܣ���ɶԵ�ǰ�����˻���
'������intType 0����ɻ��1��ȡ��Ϣ����
    Dim strSQL As String
    
    If MsgBox("ȷʵҪ" & IIF(intType = 1, "ȡ��", "") & "��ɶԸ�""" & txtInfo(txt����).Text & """�Ļ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    If intType = 0 Then
        strSQL = "ZL_����ҽ��ִ��_Finish(" & mlngAdviceID & "," & mlng���ͺ� & ",NULL,0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngִ�п���ID & ")"
    Else
        strSQL = "ZL_����ҽ��ִ��_Cancel(" & mlngAdviceID & "," & mlng���ͺ� & ",Null,0," & mlngִ�п���ID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    End If
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    mintType = IIF(mintType = 3, 4, 3)
    Call SetFormState
    ExecuteMeet = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get�������() As Long
'���ܣ���ȡ�������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSQL = "Select ����ҽ����¼_�������.Nextval as ������� From Dual"
    
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Get������� = Val(rsTmp!������� & "")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub optInfo_Click(Index As Integer)
    Select Case Index
    Case optԺ��
        fraInfo(fra����ҽԺ).Visible = False
        txtInfo(txt����ҽԺ).Text = ""
        vsDetail.Rows = vsDetail.FixedRows
        vsDetail.Rows = vsDetail.FixedRows + 1
    Case optԺ��
        fraInfo(fra����ҽԺ).Visible = True
        txtInfo(txt����ҽԺ).Locked = False
        txtInfo(txt����ҽԺ).TabStop = True
        txtInfo(txt����ҽԺ).Text = ""
        vsDetail.Rows = vsDetail.FixedRows
        vsDetail.Rows = vsDetail.FixedRows + 1
    End Select
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
'�Ϸ��Լ���ֵ�Ļָ�
    If mintType = 0 Then Exit Sub
    
    Select Case Index
    Case txt����ʱ��
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(Index).Text) Then
                    '�ָ���Ϊ�����ȱʡΪ��ʼʱ��
                    txtInfo(Index).Text = txtInfo(Index).Text
                End If
            End If
        Else
            '���ʱ��Ϸ���
            If Not Check���ʱ��(txtInfo(Index).Text, txtInfo(Index).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
        End If
        '�ж��Ƿ��ǲ�¼ҽ��
        If DateDiff("n", CDate(txtInfo(Index).Text), CDate(zlDatabase.Currentdate)) > gint��¼��� _
            Or mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
            mbln��¼ = True
        Else
            mbln��¼ = False
        End If
    Case txt���ʱ��
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(Index).Tag) Then
                    '�ָ���Ϊ�����
                    txtInfo(Index).Text = txtInfo(Index).Tag
                End If
            End If
        Else
            '���ʱ��Ϸ���
            If Not Check��ʼʱ��(txtInfo(Index).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
            
        End If
    
    Case txt�������, txt����ҽʦ
        If txtInfo(Index).Tag <> txtInfo(Index).Text And txtInfo(Index).Tag <> "" Then txtInfo(Index).Text = txtInfo(Index).Tag
    End Select
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub SetItemEditable(Optional int����ʱ�� As Integer, Optional int�������� As Integer, _
    Optional int������Ŀ As Integer, Optional int���ﷶΧ As Integer, Optional int����ҽԺ As Integer, Optional int������ϸ As Integer, Optional int�ٴ���� As Integer, _
    Optional int����Ŀ�� As Integer, Optional int����ժҪ As Integer, _
    Optional int������� As Integer, Optional int���ʱ�� As Integer, _
    Optional int������� As Integer, Optional int����ҽʦ As Integer)
'���ܣ�����ָ���༭��Ŀ���״̬
'������0-���ֲ���,-1-����,1-����
        
    If int����ʱ�� = 1 Then
        txtInfo(txt����ʱ��).Locked = False
        txtInfo(txt����ʱ��).TabStop = True
        txtInfo(txt����ʱ��).BackColor = vbWindowBackground
        cmdDate(cmd����ʱ��).Enabled = True
    ElseIf int����ʱ�� = -1 Then
        txtInfo(txt����ʱ��).Locked = True
        txtInfo(txt����ʱ��).TabStop = False
        txtInfo(txt����ʱ��).BackColor = vbButtonFace
        cmdDate(cmd����ʱ��).Enabled = False
    End If
    
    If int�������� = 1 Then
        fraInfo(fra��������).Enabled = True
    ElseIf int�������� = -1 Then
        fraInfo(fra��������).Enabled = False
    End If
    
    If int������Ŀ = 1 Then
        cboItem.Locked = False
        cboItem.TabStop = True
        cboItem.BackColor = vbWindowBackground
    ElseIf int������Ŀ = -1 Then
        cboItem.Locked = True
        cboItem.TabStop = False
        cboItem.BackColor = vbButtonFace
    End If
    
    If int���ﷶΧ = 1 Then
        fraInfo(fra���ﷶΧ).Enabled = True
    ElseIf int�������� = -1 Then
        fraInfo(fra���ﷶΧ).Enabled = False
    End If
    
    If int����ҽԺ = 1 Then
        txtInfo(txt����ҽԺ).Locked = False
        txtInfo(txt����ҽԺ).TabStop = True
        txtInfo(txt����ҽԺ).BackColor = vbWindowBackground
    ElseIf int����ҽԺ = -1 Then
        txtInfo(txt����ҽԺ).Locked = True
        txtInfo(txt����ҽԺ).TabStop = False
        txtInfo(txt����ҽԺ).BackColor = vbButtonFace
    End If
    
    If int������ϸ = 1 Then
        vsDetail.TabStop = True
        fraDetail.Enabled = True
    ElseIf int������ϸ = -1 Then
        vsDetail.TabStop = False
        vsDetail.Editable = flexEDNone
        fraDetail.Enabled = False
    End If
    
    If int�ٴ���� = 1 Then
        txtInfo(txt�ٴ����).Locked = False
        txtInfo(txt�ٴ����).TabStop = True
        txtInfo(txt�ٴ����).BackColor = vbWindowBackground
        cmdInfo(cmd�ٴ����).Enabled = True
    ElseIf int�ٴ���� = -1 Then
        txtInfo(txt�ٴ����).Locked = True
        txtInfo(txt�ٴ����).TabStop = False
        txtInfo(txt�ٴ����).BackColor = vbButtonFace
        cmdInfo(cmd�ٴ����).Enabled = False
    End If
    
    If int����Ŀ�� = 1 Then
        txtInfo(txt����Ŀ��).Locked = False
        txtInfo(txt����Ŀ��).TabStop = True
        txtInfo(txt����Ŀ��).BackColor = vbWindowBackground
    ElseIf int����Ŀ�� = -1 Then
        txtInfo(txt����Ŀ��).Locked = True
        txtInfo(txt����Ŀ��).TabStop = False
        txtInfo(txt����Ŀ��).BackColor = vbButtonFace
    End If
    
    
    If int������� = 1 Then
        txtInfo(txt�������).Locked = False
        txtInfo(txt�������).TabStop = True
        txtInfo(txt�������).BackColor = vbWindowBackground
    ElseIf int������� = -1 Then
        txtInfo(txt�������).Locked = True
        txtInfo(txt�������).TabStop = False
        txtInfo(txt�������).BackColor = vbButtonFace
    End If
    
    If int����ժҪ = 1 Then
        txtInfo(txt����ժҪ).Locked = False
        txtInfo(txt����ժҪ).TabStop = True
        txtInfo(txt����ժҪ).BackColor = vbWindowBackground
    ElseIf int����ժҪ = -1 Then
        txtInfo(txt����ժҪ).Locked = True
        txtInfo(txt����ժҪ).TabStop = False
        txtInfo(txt����ժҪ).BackColor = vbButtonFace
    End If
    
    
    If int���ʱ�� = 1 Then
        txtInfo(txt���ʱ��).Locked = False
        txtInfo(txt���ʱ��).TabStop = True
        txtInfo(txt���ʱ��).BackColor = vbWindowBackground
        cmdDate(cmd���ʱ��).Enabled = True
    ElseIf int���ʱ�� = -1 Then
        txtInfo(txt���ʱ��).Locked = True
        txtInfo(txt���ʱ��).TabStop = False
        txtInfo(txt���ʱ��).BackColor = vbButtonFace
        cmdDate(cmd���ʱ��).Enabled = False
    End If
    
    
    If int������� = 1 Then
        txtInfo(txt�������).Locked = False
        txtInfo(txt�������).TabStop = True
        txtInfo(txt�������).BackColor = vbWindowBackground
        cmdInfo(cmd�������).Enabled = True
    ElseIf int������� = -1 Then
        txtInfo(txt�������).Locked = True
        txtInfo(txt�������).TabStop = False
        txtInfo(txt�������).BackColor = vbButtonFace
        cmdInfo(cmd�������).Enabled = False
    End If
    
    
    If int����ҽʦ = 1 Then
        txtInfo(txt����ҽʦ).Locked = False
        txtInfo(txt����ҽʦ).TabStop = True
        txtInfo(txt����ҽʦ).BackColor = vbWindowBackground
        cmdInfo(cmd����ҽʦ).Enabled = True
    ElseIf int����ҽʦ = -1 Then
        txtInfo(txt����ҽʦ).Locked = True
        txtInfo(txt����ҽʦ).TabStop = False
        txtInfo(txt����ҽʦ).BackColor = vbButtonFace
        cmdInfo(cmd����ҽʦ).Enabled = False
    End If
     
End Sub

