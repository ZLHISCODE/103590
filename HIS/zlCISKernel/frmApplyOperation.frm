VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmApplyOperation 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������뵥"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10650
   Icon            =   "frmApplyOperation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   27
      Left            =   5490
      ScaleHeight     =   255
      ScaleWidth      =   1305
      TabIndex        =   72
      Top             =   2565
      Width           =   1300
      Begin VB.ComboBox cboInfo 
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
         Index           =   27
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   -25
         Width           =   1200
      End
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
      Height          =   255
      Index           =   17
      Left            =   5325
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   6885
      Width           =   1260
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
      Height          =   270
      Index           =   17
      Left            =   6705
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   6885
      Width           =   270
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8715
      ScaleHeight     =   330
      ScaleWidth      =   1770
      TabIndex        =   67
      Top             =   1080
      Width           =   1770
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   0
         Width           =   1330
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
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
         Index           =   0
         Left            =   0
         TabIndex        =   69
         Top             =   45
         Width           =   240
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   240
         X2              =   1680
         Y1              =   285
         Y2              =   285
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   13
      Left            =   5400
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   65
      Top             =   4080
      Width           =   2055
      Begin VB.ComboBox cboInfo 
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
         Index           =   13
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   -25
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   24
      Left            =   10080
      Picture         =   "frmApplyOperation.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   57
      TabStop         =   0   'False
      ToolTipText     =   "�༭(F4)"
      Top             =   10005
      Width           =   285
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
      Height          =   255
      Index           =   21
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   9570
      Width           =   1455
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
      Height          =   255
      Index           =   23
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   9570
      Width           =   1455
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
      Height          =   255
      Index           =   22
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   9570
      Width           =   1455
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
      Height          =   255
      Index           =   11
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3600
      Width           =   855
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
      Height          =   255
      Index           =   20
      Left            =   8445
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1380
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
      Height          =   270
      Index           =   20
      Left            =   9945
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   7440
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
      Height          =   255
      Index           =   19
      Left            =   4845
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   7440
      Width           =   1500
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
      Height          =   270
      Index           =   19
      Left            =   6405
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   7440
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
      Height          =   255
      Index           =   18
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   7440
      Width           =   1260
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
      Height          =   270
      Index           =   18
      Left            =   3165
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   7440
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
      Height          =   255
      Index           =   16
      Left            =   1695
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   6885
      Width           =   1260
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
      Height          =   270
      Index           =   16
      Left            =   3165
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   6885
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
      Height          =   255
      Index           =   14
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
      Height          =   270
      Index           =   14
      Left            =   3165
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   6360
      Width           =   270
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   12
      Left            =   1530
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   39
      Top             =   4080
      Width           =   2055
      Begin VB.ComboBox cboInfo 
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
         Index           =   12
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   -25
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
      Height          =   270
      Index           =   10
      Left            =   7845
      TabIndex        =   36
      ToolTipText     =   "ѡ��(*)"
      Top             =   3600
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
      Height          =   255
      Index           =   10
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3600
      Width           =   6060
   End
   Begin VB.CommandButton cmdDate 
      Height          =   285
      Index           =   8
      Left            =   3285
      Picture         =   "frmApplyOperation.frx":6948
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "�༭(F4)"
      Top             =   2640
      Width           =   285
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
      Height          =   255
      Index           =   8
      Left            =   1620
      TabIndex        =   18
      Text            =   "2013-06-20 18:00"
      Top             =   2655
      Width           =   1935
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
      Height          =   255
      Index           =   4
      Left            =   8745
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1545
      Width           =   1330
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
      Height          =   255
      Index           =   6
      Left            =   4365
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2010
      Width           =   1215
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
      Height          =   255
      Index           =   1
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1330
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
      Height          =   255
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1330
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
      Height          =   255
      Index           =   7
      Left            =   6885
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2010
      Width           =   1215
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
      Height          =   255
      Index           =   5
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1335
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
      Height          =   255
      Index           =   3
      Left            =   6450
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1665
      Width           =   1330
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
      Height          =   255
      Index           =   9
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3090
      Width           =   8175
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "��"
      Height          =   270
      Index           =   9
      Left            =   9960
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   3090
      Width           =   270
   End
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   10800
      TabIndex        =   17
      Top             =   5280
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
   Begin VSFlex8Ctl.VSFlexGrid vsOper 
      Height          =   1605
      Left            =   1605
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4560
      Width           =   8760
      _cx             =   15452
      _cy             =   2831
      Appearance      =   0
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
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmApplyOperation.frx":6A3E
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
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
      Height          =   255
      Index           =   24
      Left            =   8400
      TabIndex        =   34
      Text            =   "2013-06-20 18:00"
      Top             =   10020
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid vsOther 
      Height          =   1545
      Left            =   1560
      TabIndex        =   33
      Top             =   7845
      Width           =   8775
      _cx             =   15478
      _cy             =   2725
      Appearance      =   0
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
      SheetBorder     =   -2147483632
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmApplyOperation.frx":6A68
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   15
      Left            =   5280
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   43
      Top             =   6360
      Width           =   2175
      Begin VB.ComboBox cboInfo 
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
         Index           =   15
         Left            =   -25
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   -25
         Width           =   1995
      End
   End
   Begin VB.Line Line1 
      Index           =   27
      X1              =   5460
      X2              =   6765
      Y1              =   2835
      Y2              =   2835
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "�ٴ���������"
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
      Index           =   27
      Left            =   4140
      TabIndex        =   71
      Top             =   2595
      Width           =   1260
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   225
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   26
      Left            =   645
      TabIndex        =   66
      Top             =   7905
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   8400
      X2              =   10080
      Y1              =   10290
      Y2              =   10290
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   1560
      X2              =   3240
      Y1              =   9840
      Y2              =   9840
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
      Index           =   21
      Left            =   645
      TabIndex        =   64
      Top             =   9600
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��Чʱ��"
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
      Index           =   24
      Left            =   7485
      TabIndex        =   63
      Top             =   10020
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   8400
      X2              =   10080
      Y1              =   9840
      Y2              =   9840
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
      Index           =   23
      Left            =   7485
      TabIndex        =   62
      Top             =   9600
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   22
      X1              =   4920
      X2              =   6600
      Y1              =   9840
      Y2              =   9840
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
      Index           =   22
      Left            =   3960
      TabIndex        =   61
      Top             =   9600
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   11
      Left            =   8325
      TabIndex        =   56
      Top             =   3615
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   12
      Left            =   645
      TabIndex        =   55
      Top             =   4080
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   9285
      X2              =   10320
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   20
      Left            =   7455
      TabIndex        =   53
      Top             =   7440
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   8325
      X2              =   10200
      Y1              =   7710
      Y2              =   7710
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "�ڶ�����"
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
      Index           =   19
      Left            =   3960
      TabIndex        =   51
      Top             =   7440
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   4815
      X2              =   6720
      Y1              =   7710
      Y2              =   7710
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��һ����"
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
      Index           =   18
      Left            =   645
      TabIndex        =   49
      Top             =   7440
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   1605
      X2              =   3480
      Y1              =   7710
      Y2              =   7710
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "����ҽ������"
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
      Index           =   17
      Left            =   3960
      TabIndex        =   47
      Top             =   6900
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   5280
      X2              =   7200
      Y1              =   7155
      Y2              =   7155
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "����ҽ��"
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
      Index           =   16
      Left            =   645
      TabIndex        =   46
      Top             =   6900
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   1605
      X2              =   3480
      Y1              =   7155
      Y2              =   7155
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "����ִ�п���"
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
      Left            =   3960
      TabIndex        =   44
      Top             =   6360
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   5280
      X2              =   7200
      Y1              =   6630
      Y2              =   6630
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "������"
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
      Index           =   14
      Left            =   645
      TabIndex        =   42
      Top             =   6360
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   1605
      X2              =   3480
      Y1              =   6630
      Y2              =   6630
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��������"
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
      Index           =   10
      Left            =   585
      TabIndex        =   40
      Top             =   3600
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   1605
      X2              =   3525
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "����ʱ��"
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
      Index           =   8
      Left            =   585
      TabIndex        =   38
      Top             =   2655
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "��ǰ���"
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
      Index           =   9
      Left            =   645
      TabIndex        =   37
      Top             =   3135
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   1605
      X2              =   8160
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   5280
      X2              =   7200
      Y1              =   4350
      Y2              =   4350
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   13
      Left            =   4320
      TabIndex        =   35
      Top             =   4095
      Width           =   840
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   1605
      X2              =   3285
      Y1              =   2925
      Y2              =   2925
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   8700
      X2              =   9660
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   6765
      X2              =   8085
      Y1              =   1920
      Y2              =   1920
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
      Left            =   7816
      TabIndex        =   16
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   25
      Left            =   645
      TabIndex        =   15
      Top             =   4560
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
      Left            =   2925
      TabIndex        =   14
      Top             =   1680
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
      Left            =   5371
      TabIndex        =   13
      Top             =   1680
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
      Index           =   6
      Left            =   3285
      TabIndex        =   12
      Top             =   2040
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
      Index           =   1
      Left            =   480
      TabIndex        =   11
      Top             =   1640
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
      Left            =   5790
      TabIndex        =   10
      Top             =   2040
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
      Index           =   5
      Left            =   645
      TabIndex        =   9
      Top             =   2070
      Width           =   840
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
      Left            =   4605
      TabIndex        =   8
      Top             =   750
      Width           =   1875
   End
   Begin VB.Line linHead 
      X1              =   4590
      X2              =   6510
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   4245
      X2              =   5565
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1380
      X2              =   2640
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   4245
      X2              =   5565
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   6765
      X2              =   8085
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   1605
      X2              =   3045
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   1605
      X2              =   10200
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "frmApplyOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCtlID
    e_No = 0
    e_���� = 1
    e_�Ա� = 2
    e_���� = 3
    e_���� = 4
    e_���� = 5
    e_סԺ�� = 6
    e_���� = 7
    e_����ʱ�� = 8
    e_��ǰ��� = 9
    e_�������� = 10
    e_�������� = 11
    e_ִ�п��� = 12
    e_������� = 13
    e_������ = 14
    e_����ִ�п��� = 15
    e_����ҽ�� = 16
    e_����ҽ������ = 17
    e_��һ���� = 18
    e_�ڶ����� = 19
    e_�������� = 20
    e_������� = 21
    e_����ҽʦ = 22
    e_����ҽʦ = 23
    e_��Чʱ�� = 24
    e_�������� = 25
    e_�������� = 26
    e_�ٴ��������� = 27
End Enum

Private Enum mTableCol
    '�����������
    COL_���� = 0
    COL_�������� = 1
    COL_�Ƽ����� = 2
    COL_ִ������ = 3
    COL_ժҪ = 4
    COL_��Ҫʱ = 5
    
    '��չ��Ŀ
    COL_��Ŀ���� = 0
    col_���� = 1
End Enum

Private mobjVBA As Object
Private mobjScript As clsScript
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mobjReport As Object
Private mblnReturn As Boolean '�Ƿ��˻س�ȷ��

Private mintType As Integer '=0������=1�޸ģ�=2�鿴
Private mint���� As Integer '0 סԺҽ������վ��1 ����ҽ������վ��
Private mint���ó��� As Integer '���뵥���ó��ϣ�0��ҽ��վ���ã�1��ҽ���༭������á�Ϊ1ʱ�����û������ݼ��ؽ���ͱ���Ϊ�������ݡ�
Private mint�������� As Integer  '1-����,2-סԺ
Private mint������� As Integer '1-����,2-סԺ

Private mbln���Ѷ��� As Boolean
Private mint���� As Integer '��ǰ��������
Private mrsAppend As ADODB.Recordset
Private mbln��Ժҽ������ As Boolean '��Ժҽ�����뽨��
Private mstr�����ȼ� As String   '����ҽ���������ȼ�
Private mobjEmrInterface As Object           '�°没�����븽���ȡ����

Private mlng����ID As Long
Private mstr�Һŵ� As String
Private mlng��ҳID As Long
Private mlng�Һ�ID As Long
Private mlng����ID As Long
Private mlng���˿���id As Long '���˿���id/�Һ�ִ�п���id
Private mlng�������� As Long   '0-סԺ��1-����
Private mlng��������ID As Long
Private mstr�������� As String '�����������

Private mlng������ĿID As Long

Private mlng����ִ�п������� As Long
Private mlng����ִ�п���ID As Long

Private mlng������ĿID As Long
Private mlng����ִ�п������� As Long
Private mlng����ִ�п���id As Long

Private mstr������IDs As String
Private mintPState As Integer
Private mdatTurn As Date
Private mstr��Ժʱ�� As String
Private mstr�ϴ�ת��ʱ�� As String
Private mstrDefine As String
Private mlngUpdateAdvice As Long  '�����ҽ��ID����ҽ��ID
Private mstr���IDs As String  '��Ϲ���
Private mbln��¼ As Boolean
Private mclsDiagEdit As zlMedRecPage.clsDiagEdit
Private mclsMipModule As zl9ComLib.clsMipModule '��Ϣƽ̨����
Private mrsCard As ADODB.Recordset
Private mstrժҪ As String '������ҽ��ժҪ �� gclsInsure.GetItemInfo ��ȡ��lblInfo(e_����ִ�п���).Tag ��������Ŀ��ժҪ
Private mstr�ѱ� As String
Private mlngǰ��ID As Long
Private mbytBaby As Byte  'Ӥ�����

Public Function ShowMe(frmParent As Object, ByVal int���� As Integer, ByVal intType As Integer, ByVal lng����ID As Long, ByVal str����ID As String, ByVal lng�������� As Long, _
    Optional ByRef lngҽ��ID As Long, Optional ByVal lng����id As Long, Optional ByVal lng��������ID As Long, Optional ByVal strDefine As String, _
    Optional ByVal lng����ID As Long, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, Optional ByVal int���ó��� As Integer, _
    Optional ByRef objMip As Object, Optional ByVal lng��Ŀid As Long, Optional ByRef rsCard As ADODB.Recordset, Optional ByVal lngǰ��ID As Long, Optional ByVal bytBaby As Byte) As Boolean
'���ܣ������ӿ�
'������frmParent �������壻int���� 0 סԺҽ������վ��1 ����ҽ������վ�� lng�������� 0-סԺ��1-���lng����ID��
'      str����ID ���� int���� �жϣ���ҳid/�Һŵ���
'      intType ��������   0-������1-�޸ģ�2-�鿴,3-ҽ���༭���ã�lngҽ��ID �����ҽ��ID��
'      lng����ID  ���˿���id/�Һ�ִ�п���id��int���ó��� 0������վ���棬1�� ҽ���༭���棻 lng��������ID ��������id��
'      strDefine ҽ�����ݸ�ʽ����
'      lng����ID��intPState��datTurn סԺ���У�objMip ��Ϣ����������Ϣ���� סԺ���У�
'      lng��ĿID-ҽ���༭��������������Ŀ�������������ĿID
'      rsCard��ҽ����Ƭ��ҽ��������Ϣ��ֻ����ҽ���༭�����ҽ�������ұߵ�������ťʱ�Żᴫ�롣
    
    mint���� = int����
    
    If mint���� = 0 Then
        mlng��ҳID = Val(str����ID)
        mint�������� = 2
        mint������� = 2
    Else
        mstr�Һŵ� = str����ID
        mint�������� = 1
        mint������� = 1
    End If
    mint���ó��� = int���ó���
    mlng����ID = lng����ID
    mlng�������� = lng��������
    mlng���˿���id = lng����id
    mlng����ID = lng����ID
    mlng��������ID = lng��������ID
    mlngǰ��ID = lngǰ��ID
    mbytBaby = bytBaby
    mstrDefine = strDefine
    mlngUpdateAdvice = lngҽ��ID
    mintPState = intPState
    mintType = intType
    mdatTurn = datTurn
    mlng������ĿID = lng��Ŀid
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Set mrsCard = rsCard
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    If mblnOK Then lngҽ��ID = mlngUpdateAdvice
    
    ShowMe = mblnOK
    Set rsCard = mrsCard
End Function

Private Sub cboInfo_Click(Index As Integer)
    Dim blnCancel As Boolean, intIdx As Integer
    Dim strSQL As String, rsTmp As Recordset
    Dim vRect As RECT
    
    If Index = e_ִ�п��� Or Index = e_����ִ�п��� Then
        With cboInfo(Index)
            If .ItemData(.ListIndex) = -1 Then
                '����ִ�У�����ѡ��ִ�п���
                strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
                    " From ���ű� A,��������˵�� B" & _
                    " Where A.ID=B.����ID And B.������� IN(2,3)" & _
                    IIF(gstrNodeNo <> "", " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)", "") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " Order by A.����"
                vRect = zlControl.GetControlRect(cboInfo(Index).hwnd)
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "ִ�п���", , , , , , True, vRect.Left, vRect.Top, .Height, blnCancel, , True)
                If Not rsTmp Is Nothing Then
                    intIdx = Cbo.FindIndex(cboInfo(Index), rsTmp!ID)
                    If intIdx <> -1 Then
                        .ListIndex = intIdx
                    Else
                        .AddItem rsTmp!���� & "-" & rsTmp!����, .ListCount - 1
                        .ItemData(.NewIndex) = rsTmp!ID
                        .ListIndex = .NewIndex
                    End If
                    If .ListIndex >= 0 Then
                        .Tag = .ItemData(.ListIndex)
                    End If
                Else
                    If Not blnCancel Then
                        MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
                    End If
                    '�ָ������еĿ���(������Click)
                    If .Tag <> "" Then
                        intIdx = Cbo.FindIndex(cboInfo(Index), Val(.Tag))
                        Call Cbo.SetIndex(cboInfo(Index).hwnd, intIdx)
                    End If
                End If
            Else
                If Index = e_ִ�п��� Then
                    mlng����ִ�п���ID = cboInfo(Index).ItemData(cboInfo(Index).ListIndex)
                ElseIf Index = e_����ִ�п��� Then
                    mlng����ִ�п���id = cboInfo(Index).ItemData(cboInfo(Index).ListIndex)
                End If
            End If
        End With
    End If
    
    If Visible Then mblnChange = True
End Sub

Private Sub cboInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SeekNextCtl
End Sub

Private Sub cmdDate_Click(Index As Integer)
'���ܣ�ѡ������
    Dim lngIndex As Long
    
    If Index = e_����ʱ�� Then
        lngIndex = e_����ʱ��
        dtpDate.Left = txtInfo(lngIndex).Left
        dtpDate.Top = cmdDate(Index).Top + cmdDate(Index).Height
    ElseIf Index = e_��Чʱ�� Then
        lngIndex = e_��Чʱ��
        dtpDate.Left = cmdDate(Index).Left + cmdDate(Index).Width - dtpDate.Width
        dtpDate.Top = txtInfo(lngIndex).Top - dtpDate.Height
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
    If intIndex = e_����ʱ�� Then
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
    ElseIf intIndex = e_��Чʱ�� Then
        'ȡֵ
        If IsDate(txtInfo(intIndex).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(intIndex).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '�ж�ʱ��Ϸ���
        If Not Check����ʱ��(strDate, txtInfo(e_��Чʱ��).Text) Then
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

Private Sub cmdInfo_Click(Index As Integer)

    Select Case Index
    Case e_��ǰ���
        Call GetBeforeOperDiag
    Case e_��������
        Call GetItemOper(1)
    Case e_������
        Call GetItem����(1)
    Case e_����ҽ��
        Call GetItem����ҽ��(1)
    Case e_����ҽ������
        Call GetItem����ҽ������(1)
    Case e_��һ����
        Call GetItemDoctor(1, e_��һ����)
    Case e_�ڶ�����
        Call GetItemDoctor(1, e_�ڶ�����)
    Case e_��������
        Call GetItemDoctor(1, e_��������)
    End Select
End Sub

Private Sub GetBeforeOperDiag()
'���ܣ���ȡ��ǰ���
    Dim str��� As String
    Dim lng����ID As Long
    
    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mlng�������� = 1, 1260, 1261), mclsMipModule)
    End If
    lng����ID = IIF(mint���� = 0, mlng��ҳID, mlng�Һ�ID)
    If mclsDiagEdit.ShowDiagEdit(Me, mlngUpdateAdvice, mlng����ID, lng����ID, IIF(mlng�������� = 1, 1, 2), mlng���˿���id, mstr���IDs, str���, 0, mlngUpdateAdvice) Then
        txtInfo(e_��ǰ���).Text = str���
        Call SeekNextCtl
    End If
End Sub

Private Sub Loadȱʡ����()
'���ܣ�����ȱʡ������Ŀ��������Ŀ�����洫��ʱ
    Dim strSQL As String, rsTmp As Recordset
    
    strSQL = "Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ,A.ִ�п��� as ִ�п�������ID,a.�Ƽ����� as �Ƽ�����ID" & _
        " From ������ĿĿ¼ A where a.id=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng������ĿID)
    If Not rsTmp.EOF Then
        mlng������ĿID = Val(rsTmp!ID & "")
        Call Init���븽��
        mlng����ִ�п������� = Val(rsTmp!ִ�п�������ID & "")
        lblInfo(e_��������).Tag = Val(rsTmp!�Ƽ�����ID & "")
        txtInfo(e_��������).Text = GetMax�����ȼ�(mlng������ĿID)
        txtInfo(e_��������).Text = rsTmp!���� & ""
        txtInfo(e_��������).Tag = txtInfo(e_��������).Text
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItemOper(ByVal intType As Integer)
'���ܣ�ѡ��������Ŀ������������ �ؼ�����
'������intType =0 KeyPress���ã�=1 ������ť����
    Dim strSQL As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim lngTmp As Long
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ,A.ִ�п��� as ִ�п�������ID,a.�Ƽ����� as �Ƽ�����ID" & _
        " From ������ĿĿ¼ A,������Ŀ���� B" & _
        " Where A.ID=B.������ĿID And A.���='F' And A.������� IN(" & IIF(mint������� = 1, 1, 2) & ",3)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        IIF(intType = 0, " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2])", "") & _
        " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[4])" & _
        " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
        Decode(gbytCode, 0, " And B.���� IN([3],3)", 1, " And B.���� IN([3],3)", "") & _
        " Order by A.����"
    vRect = zlControl.GetControlRect(txtInfo(e_��������).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Me.Caption, False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(e_��������).Height, blnCancel, False, True, UCase(txtInfo(e_��������).Text) & "%", _
        gstrLike & UCase(txtInfo(e_��������).Text) & "%", gbytCode + 1, mlng���˿���id)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "δ�ҵ�ƥ�����Ŀ��", vbInformation, gstrSysName
        End If
        Call zlControl.TxtSelAll(txtInfo(e_��������))
        txtInfo(e_��������).SetFocus: Exit Sub
    Else
        blnDo = True
        '������������ˣ���������Ŀ���Ȩ
        If gbln������Ȩ���� Then
            If CheckUserEmpower(Val(rsTmp!ID & "")) = False Then
                blnDo = False
                MsgBox "��ǰ����Ա���߱�����""" & rsTmp!���� & """�Ŀ���Ȩ���������´", vbInformation, Me.Caption
                Call zlControl.TxtSelAll(txtInfo(e_��������))
                txtInfo(e_��������).SetFocus: Exit Sub
            End If
        End If
    End If
    
    If blnDo Then
        mlng������ĿID = Val(rsTmp!ID & "")
        Call Init���븽��
        mlng����ִ�п������� = Val(rsTmp!ִ�п�������ID & "")
        Call Set����ִ�п���(mlng����ִ�п�������)
        lblInfo(e_��������).Tag = Val(rsTmp!�Ƽ�����ID & "")
        txtInfo(e_��������).Text = GetMax�����ȼ�(mlng������ĿID)
        txtInfo(e_��������).Text = rsTmp!���� & ""
        txtInfo(e_��������).Tag = txtInfo(e_��������).Text
        txtInfo(e_��������).SetFocus
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

Private Sub GetItem����(ByVal intType As Integer)
'���ܣ�ѡ��������Ŀ�������� �ؼ�����
'������intType =0 KeyPress���ã�=1 ������ť����
    Dim strSQL As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim lngTmp As Long
    Dim strSQLItem As String, strLike As String
    
    On Error GoTo errH
    
    cboInfo(e_����ִ�п���).TabStop = True
    
    If intType = 1 Then
        '����������Ŀ
        strSQLItem = " From ������ĿĿ¼ A Where A.���='G' And A.������� IN([2],3) And A.ID<>[1]" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[3])" & _
                " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
                
        strSQL = "Select 0 as ĩ��,Max(Level) as ��ID,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ��ģ,NULL as ִ�п�������ID,null as �Ƽ�����ID" & _
            " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select ����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID" & _
            " Group by ID,�ϼ�ID,����,����"
            
        strSQL = strSQL & " Union ALL" & _
            " Select 1 as ĩ��,1 as ��ID,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ,A.ִ�п��� as ִ�п�������ID,a.�Ƽ����� as �Ƽ�����ID" & _
            strSQLItem & " Order By ĩ��,��ID Desc,����"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "������Ŀ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, mlng������ĿID, mint�������, mlng���˿���id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
            End If
            cboInfo(e_����ִ�п���).TabStop = False
            txtInfo(e_������).SetFocus: Exit Sub
        End If
    Else
        If txtInfo(e_������).Tag = txtInfo(e_������).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(e_������).Text = "" Then '�൱�����������Ŀ
            mlng������ĿID = 0
            Call cboInfo(e_����ִ�п���).Clear
            cboInfo(e_����ִ�п���).TabStop = False
            txtInfo(e_������).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
        
        '�Ż�
        strLike = gstrLike
        If Len(txtInfo(e_������).Text) < 2 Then strLike = ""
    
        '����������Ŀ
        strSQL = _
            " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ,A.ִ�п��� as ִ�п�������ID,a.�Ƽ����� as �Ƽ�����ID" & _
            " From ������ĿĿ¼ A,������Ŀ���� B" & _
            " Where A.ID=B.������ĿID And A.���='G' And A.������� IN([3],3)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]) And B.����=[4]" & _
            " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[5])" & _
            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
            " Order by A.����"
        vRect = zlControl.GetControlRect(txtInfo(e_������).hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������Ŀ", False, "", "", False, False, True, vRect.Left, vRect.Top, txtInfo(e_������).Height, blnCancel, False, True, _
            UCase(txtInfo(e_������).Text) & "%", strLike & UCase(txtInfo(e_������).Text) & "%", mint�������, gbytCode + 1, mlng���˿���id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
            End If
            txtInfo(e_������).Text = cmdInfo(e_������).Tag
            zlControl.TxtSelAll txtInfo(e_������)
            txtInfo(e_������).SetFocus
            Exit Sub
        End If
    End If
    
    mlng������ĿID = Val(rsTmp!ID & "")
    mlng����ִ�п������� = Val(rsTmp!ִ�п�������ID & "")
    Call Set����ִ�п���(mlng����ִ�п�������)
    
    lblInfo(e_������).Tag = Val(rsTmp!�Ƽ�����ID & "")
    txtInfo(e_������).Text = "[" & rsTmp!���� & "]" & rsTmp!����
    txtInfo(e_������).Tag = txtInfo(e_������).Text
    txtInfo(e_������).SetFocus
    Call SeekNextCtl
    If Visible Then mblnChange = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItem����ҽ��(ByVal intType As Integer)
'���ܣ���ȡ����ҽ����Ŀ
'������0 �ı��򰴻س���1 �㰴ť
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean, strTmp As String
    Dim blnDo As Boolean, str���� As String
    Dim lng����ID As Long, lng��Աid As Long
    Dim i As Integer
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(e_����ҽ��).Tag = txtInfo(e_����ҽ��).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(e_����ҽ��).Text = "" Then '�൱�����������Ŀ
            txtInfo(e_����ҽ��).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
            
    strInput = Trim(UCase(txtInfo(e_����ҽ��).Text))   '�����ֵ����ǰ׺�ո�
    
    strSQL = "Select A.ID,A.���,A.����,A.����,A.�����ȼ�" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        IIF(intType = 0, " And (A.��� Like [1] Or A.���� Like [2] Or A.���� Like [2])", "") & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.���"
    vRect = zlControl.GetControlRect(txtInfo(e_����ҽ��).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҽ��", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(e_����ҽ��).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
        
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            If mbln��Ժҽ������ Then
                Call MsgBox("û���ҵ�ƥ���ҽ��!", vbInformation, gstrSysName)
                blnDo = False
            Else
                If MsgBox("û���ҵ�ƥ���ҽ������ȷ��Ҫ����û�н�����Ա������ҽ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    blnDo = True
                    strTmp = strInput
                Else
                    blnDo = False
                End If
            End If
        Else
            Call MsgBox("û���ҵ�ƥ���ҽ��!", vbInformation, gstrSysName)
            blnDo = False
        End If
    Else
        blnDo = True
        txtInfo(e_����ҽ��).Text = rsTmp!���� & ""
        txtInfo(e_����ҽ��).Tag = rsTmp!���� & ""
        lng��Աid = rsTmp!ID
        If gbln�����ּ����� Then mstr�����ȼ� = rsTmp!�����ȼ� & ""
        txtInfo(e_����ҽ��).SetFocus
    End If
    
    If blnDo Then
        strSQL = "Select b.����,a.ȱʡ From ������Ա A, ���ű� B, ��������˵�� C" & _
            " Where a.����id = b.Id And b.Id = c.����id And c.�������� = '�ٴ�' And a.��Աid = [1]" & _
            " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) "

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Աid)
        If Not rsTmp.EOF Then
            txtInfo(e_����ҽ������).Text = rsTmp!���� & ""
            rsTmp.Filter = "ȱʡ=1"
            If Not rsTmp.EOF Then txtInfo(e_����ҽ������).Text = rsTmp!���� & ""
            
            txtInfo(e_����ҽ������).Tag = txtInfo(e_����ҽ������).Text
            txtInfo(e_����ҽ������).SetFocus
        End If
        Call SeekNextCtl
    Else
        txtInfo(e_����ҽ��).SetFocus
        Call zlControl.TxtSelAll(txtInfo(e_����ҽ��))
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItem����ҽ������(ByVal intType As Integer)
'���ܣ���ȡ����ҽ������
'������0 �ı��򰴻س���1 �㰴ť
    Dim strSQL As String, rsTmp As Recordset
    Dim strInput As String, vRect As RECT
    Dim blnCancel As Boolean
    Dim i As Integer, strDoctor As String
    
    On Error GoTo errH
    
    If intType = 0 Then
        If txtInfo(e_����ҽ������).Tag = txtInfo(e_����ҽ������).Text Then
            Call SeekNextCtl
            Exit Sub
        ElseIf txtInfo(e_����ҽ������).Text = "" Then  '�൱�����������Ŀ
            txtInfo(e_����ҽ������).Tag = ""
            Call SeekNextCtl
            Exit Sub
        End If
    End If
       
    strDoctor = txtInfo(e_����ҽ��).Text
    strInput = Trim(UCase(txtInfo(e_����ҽ������).Text))    '�����ֵ����ǰ׺�ո�
    
    strSQL = "Select Distinct A.ID,A.����,A.���� as ����,A.���� From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)  And a.Id = b.����id" & _
        IIF(intType = 0, " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])", "") & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) And B.��������='�ٴ�'" & _
        IIF(strDoctor <> "", " And a.id in (select x.����id from ������Ա X, ��Ա�� Y where x.��Աid=y.id and y.����=[3])", "") & _
        " Order by A.����"
        
    vRect = zlControl.GetControlRect(txtInfo(e_����ҽ������).hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtInfo(e_����ҽ������).Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", strDoctor)
    
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            Call MsgBox("û���ҵ�ƥ��Ŀ���!", vbInformation, gstrSysName)
            txtInfo(e_����ҽ������).SetFocus
            zlControl.TxtSelAll txtInfo(e_����ҽ������)
            Exit Sub
        End If
    Else
        txtInfo(e_����ҽ������).Text = rsTmp!���� & ""
        txtInfo(e_����ҽ������).Tag = rsTmp!���� & ""
        txtInfo(e_����ҽ������).SetFocus
        Call SeekNextCtl
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetItemDoctor(ByVal intType As Integer, ByVal intIndex As Integer)
'���ܣ���ȡҽ��
'������intType 0 ����ƥ�䣬1 �����ť��intIndex �ؼ�����
    Dim strInput As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    
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

Private Sub Set����ִ�п���(ByVal lngִ�п��� As Long, Optional ByVal lngִ�п���ID As Long)
'���ܣ�����ִ�п���
'������lngִ�п���-ִ�����ʣ�lngִ�п���ID=������룬���ʾ���ô�ִ�п���Ϊ��ǰִ�п���
    Dim lngTmp As Long
    
    With cboInfo(e_ִ�п���)
        .Enabled = True
        If lngִ�п��� = 5 Then
            .Clear:
            .AddItem "-"
            .ListIndex = 0
        Else
            If .ListIndex >= 0 And lngִ�п���ID = 0 Then
                lngTmp = .ItemData(.ListIndex)
            ElseIf lngִ�п���ID <> 0 Then
                lngTmp = lngִ�п���ID
            End If
            
            If lngTmp = 0 Then lngTmp = Get����ִ�п���ID(mlng����ID, mlng��ҳID, "F", mlng������ĿID, 0, lngִ�п���, mlng���˿���id, mlng��������ID, 1, IIF(mlng�������� = 1, 1, 2))
            
            Call Get����ִ�п���(mlng����ID, mlng��ҳID, cboInfo(e_ִ�п���), "F", mlng������ĿID, 0, _
                lngִ�п���, mlng���˿���id, mlng��������ID, lngTmp, 1, IIF(mlng�������� = 1, 1, 2))

            If lngִ�п���ID = 0 Then
                If .ListIndex = -1 And .ListCount = 1 Then
                    .ListIndex = 0
                Else
                     '����ж����ȡĬ�ϵ�ִ�п���
                    lngִ�п���ID = Get����ִ�п���ID(mlng����ID, mlng��ҳID, "F", mlng������ĿID, 0, _
                            lngִ�п���, mlng���˿���id, mlng��������ID, 1, IIF(mlng�������� = 1, 1, 2))
                End If
            End If

            If lngִ�п���ID <> 0 Then Call Cbo.Locate(cboInfo(e_ִ�п���), lngִ�п���ID, True)
         
        End If
        
        If .ListCount = 1 Then .Enabled = False
        
        If .ListIndex >= 0 Then .Tag = .ItemData(.ListIndex)
     
    End With
End Sub

Private Sub Set����ִ�п���(ByVal lngִ�п��� As Long, Optional ByVal lngִ�п���ID As Long)
'���ܣ���������ִ�п���
'������lngִ�п���-ִ�����ʣ�lngִ�п���ID=������룬���ʾ���ô�ִ�п���Ϊ��ǰִ�п���
    Dim lngTmp As Long
    
    With cboInfo(e_����ִ�п���)
        .Enabled = True
        If lngִ�п��� = 5 Then
            .Clear: .AddItem "-"
            .ListIndex = 0
        Else
            If .ListIndex >= 0 And lngִ�п���ID = 0 Then
                lngTmp = .ItemData(.ListIndex)
            ElseIf lngִ�п���ID <> 0 Then
                lngTmp = lngִ�п���ID
            End If
            
            If lngTmp = 0 Then lngTmp = Get����ִ�п���ID(mlng����ID, mlng��ҳID, "G", mlng������ĿID, 0, lngִ�п���, mlng���˿���id, mlng��������ID, 1, IIF(mlng�������� = 1, 1, 2))
            
            Call Get����ִ�п���(mlng����ID, mlng��ҳID, cboInfo(e_����ִ�п���), "G", mlng������ĿID, 0, _
                lngִ�п���, mlng���˿���id, mlng��������ID, lngTmp, 1, IIF(mlng�������� = 1, 1, 2))
                
            If lngִ�п���ID = 0 Then
                If .ListIndex = -1 And .ListCount = 1 Then
                    .ListIndex = 0
                Else
                     '����ж����ȡĬ�ϵ�ִ�п���
                    lngִ�п���ID = Get����ִ�п���ID(mlng����ID, mlng��ҳID, "G", mlng������ĿID, 0, _
                            lngִ�п���, mlng���˿���id, mlng��������ID, 1, IIF(mlng�������� = 1, 1, 2))
                End If
            End If
            
            If lngִ�п���ID <> 0 Then Call Cbo.Locate(cboInfo(e_����ִ�п���), lngִ�п���ID, True)
          
        End If
        
        If .ListCount = 1 Then cboInfo(e_����ִ�п���).Enabled = False
        
        If .ListIndex >= 0 Then .Tag = .ItemData(.ListIndex)
 
    End With
End Sub

Private Function GetItemAppend(ByVal lngҪ��ID As Long, ByVal str������ As String, ByVal str��Ŀ As String) As String
'���ܣ���ȡָ�������븽��ֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strText As String
    Dim arrItem As Variant, i As Long
    
    On Error GoTo errH
    
    If mint�������� = 1 Then
        '3.δȡ����δ��ӦҪ�صģ��Ӳ���֮ǰ�ѱ����ҽ������ȡ,�������д��Ϊ׼
        strSQL = " Select ���� From (" & _
            " Select B.���� From ����ҽ����¼ A,����ҽ������ B" & _
            " Where A.ID=B.ҽ��ID And A.����ID=[1] And Nvl(A.Ӥ��,0)=[4]" & _
            IIF(mint���� = 1, " And A.�Һŵ�=[2]", " And A.��ҳID=[3]") & _
            " And B.��Ŀ=[5] And B.���� is Not Null and nvl(a.ҽ��״̬,0)<>4 Order by A.����ʱ�� Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, mlng��ҳID, 0, str��Ŀ)
        If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
    End If
    
    '1.����ж�ӦҪ�أ���Ҫ����ȡ������ȡ
    If lngҪ��ID <> 0 And strText = "" Then
        '���ϰ棬���°�
        If mint���� = 1 Then
            '�ϰ���Ӳ���
            strSQL = "Select Zl_Replace_Element_Value(B.������,[1],A.ID,1) as ����" & _
                " From ���˹Һż�¼ A,����������Ŀ B Where A.NO=[2] And B.ID=[3] And a.��¼����=1 And a.��¼״̬=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, lngҪ��ID)
            If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
        Else
            '�ϰ���Ӳ���
            strSQL = "Select Zl_Replace_Element_Value(������,[1],[2],2) as ����" & _
                " From ����������Ŀ Where ID=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, lngҪ��ID)
            If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
        End If
        If strText = "" Then
            strText = GetItemAppendByEmr(str������)
        End If
    End If
    
    '2.�����ϣ���δ�������¼���������ȡ
    If str��Ŀ Like "*���" And strText = "" And txtInfo(e_��ǰ���).Text <> "" Then
        strText = txtInfo(e_��ǰ���).Text
    End If
    
    If mint�������� = 2 And strText = "" Then
        '3.δȡ����δ��ӦҪ�صģ��Ӳ���֮ǰ�ѱ����ҽ������ȡ,�������д��Ϊ׼
        strSQL = " Select ���� From (" & _
            " Select B.���� From ����ҽ����¼ A,����ҽ������ B" & _
            " Where A.ID=B.ҽ��ID And A.����ID=[1] And Nvl(A.Ӥ��,0)=[4]" & _
            IIF(mint���� = 1, " And A.�Һŵ�=[2]", " And A.��ҳID=[3]") & _
            " And B.��Ŀ=[5] And B.���� is Not Null and nvl(a.ҽ��״̬,0)<>4 Order by A.����ʱ�� Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, mlng��ҳID, 0, str��Ŀ)
        If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
    End If
    
    GetItemAppend = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetItemAppendByEmr(ByVal str������ As String) As String
'���ܣ���ȡָ�����˵�ָ������ڲ�����д����Ϣ�����磺���ߣ���ϵȡ��Ӳ����л�ȡ����ֵ
    Dim strText As String
    Dim intType As Integer
    Dim lng����ID As Long
    
    On Error Resume Next
    
    If mobjEmrInterface Is Nothing Then Set mobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
 
    If Not mobjEmrInterface Is Nothing Then
        If mint���� = 0 Then
            intType = 2
            lng����ID = mlng��ҳID
        Else
            intType = 1
            lng����ID = mlng�Һ�ID
        End If
        
        strText = mobjEmrInterface.GetOrderInspectInfoEx(intType, mlng����ID, lng����ID, str������)
        If err.Number <> 0 Then
            strText = mobjEmrInterface.GetOrderInspectInfo(mlng����ID, str������)
        End If
        
    End If
    
    err.Clear
    GetItemAppendByEmr = strText
End Function

Private Sub LoadPatiInfo()
'���ܣ���ȡ���˻�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
     
    '��ȡ���������Ϣ
    If mint���� = 0 Then
        If 0 = mbytBaby Then
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
            txtInfo(e_סԺ��).Text = rsTmp!סԺ�� & ""
            txtInfo(e_����).Text = rsTmp!���� & ""
            txtInfo(e_�Ա�).Text = rsTmp!�Ա� & ""
            txtInfo(e_����).Text = rsTmp!���� & ""
            txtInfo(e_����).Text = rsTmp!��ǰ���� & ""
            txtInfo(e_����).Text = rsTmp!���� & ""
            txtInfo(e_����).Text = rsTmp!���� & ""
            mstr��Ժʱ�� = Format(rsTmp!��Ժ���� & "", "YYYY-MM-DD HH:mm")
            mint���� = Val(rsTmp!���� & "")
            mstr�ѱ� = rsTmp!�ѱ� & ""
        End If
        
        mstr�ϴ�ת��ʱ�� = Get�ϴ�ת������
        
 
        strSQL = "select ��Ϣֵ as ����ҽʦ from ������ҳ�ӱ� where ����id=[1] and ��ҳid=[2] and ��Ϣ��=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, "����ҽʦ")
        If Not rsTmp.EOF Then txtInfo(e_����ҽʦ).Text = rsTmp!����ҽʦ & ""
        
    Else
        strSQL = "Select a.id, A.����,A.�Ա�,A.����,a.no,a.�����,a.����,b.���� as ����,a.ִ���� as ����ҽʦ,decode(a.����,1,'��','��') as ����,c.�ѱ�" & _
                " From ���˹Һż�¼ A,���ű� b,������Ϣ c Where a.����id=c.����id and A.NO=[1] And a.��¼����=1 And a.��¼״̬=1 And A.����ID+0=[2]" & _
                " and a.ִ�в���id=b.id"
            
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�, mlng����ID)
        
        If rsTmp.RecordCount > 0 Then
            lblInfo(e_סԺ��).Caption = "�� �� ��"
            txtInfo(e_סԺ��).Text = rsTmp!NO & ""
            
            txtInfo(e_����).Text = rsTmp!���� & ""
            txtInfo(e_�Ա�).Text = rsTmp!�Ա� & ""
  
            txtInfo(e_����).Text = rsTmp!���� & ""
            lblInfo(e_����).Caption = "�� �� ��"
            txtInfo(e_����).Text = rsTmp!����� & ""
            txtInfo(e_����).Text = rsTmp!���� & ""
            
            lblInfo(e_����).Caption = "��    ��"
            txtInfo(e_����).Text = rsTmp!���� & ""
            
            mlng�Һ�ID = Val(rsTmp!ID & "")
            mint���� = Val(rsTmp!���� & "")
            mstr�ѱ� = rsTmp!�ѱ� & ""
            txtInfo(e_����ҽʦ).Text = rsTmp!����ҽʦ & ""
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitDefault()
'���ܣ���ʼ��һЩ���У�ȱʡֵ�趨
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsTmpOther As ADODB.Recordset
    Dim datCur  As Date, str��ĿIDs As String, strTmp As String
    Dim blnCanSave As Boolean, str��Чʱ�� As String, str����ʱ�� As String
    Dim lngִ�п���ID As Long, lng����ִ��ID As Long
    Dim lng������ĿID As Long
    Dim int������� As Integer
    Dim arrTmp As Variant
    Dim i As Long
    Dim bln�ٴ� As Boolean
    
    On Error GoTo errH
    datCur = zlDatabase.Currentdate
 
    '������ȱʡֵ
    txtInfo(e_����ʱ��).Text = Format(datCur, "YYYY-MM-DD HH:mm")
    txtInfo(e_��Чʱ��).Text = txtInfo(e_����ʱ��).Text
    txtInfo(e_��Чʱ��).Tag = txtInfo(e_��Чʱ��).Text
    
    If mint���� = 0 Then
        If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
            If mdatTurn <> CDate(0) Then datCur = mdatTurn - 1 / 24 / 60
            mbln��¼ = True
        End If
    End If
    If mlng��������ID <> 0 Then
        mstr�������� = Sys.RowValue("���ű�", mlng��������ID, "����")
        txtInfo(e_�������).Text = mstr��������
    End If
    txtInfo(e_����ҽʦ).Text = UserInfo.����
    
    ' �ٴ���������  �ɼ���
    If mintType = 2 Then
        '�鿴������Ŀʱ�����ж�
        strSQL = "select 1 from ����ҽ������ a where a.ҽ��id=[1] and a.��Ŀ='�ٴ���������'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
        bln�ٴ� = Not rsTmp.EOF
    Else
        bln�ٴ� = IsOperateAgain()
    End If
    
    lblInfo(e_�ٴ���������).Visible = bln�ٴ�
    picInfo(e_�ٴ���������).Visible = bln�ٴ�
    Line1(e_�ٴ���������).Visible = bln�ٴ�
        
    If mintType = 0 Then '����������Ŀ
        lblInfo(e_��������).Visible = False
        vsOther.Visible = False
        picNo.Visible = False
        If mlng������ĿID <> 0 Then
            Call Loadȱʡ����
            If lngִ�п���ID <> 0 Then
                Call Set����ִ�п���(mlng����ִ�п�������, lngִ�п���ID)
            Else
                Call Set����ִ�п���(mlng����ִ�п�������)
            End If
        End If
        SetItemEditable 1, , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, , 1, 1
    ElseIf mintType = 2 Then '�鿴������Ŀ
        picNo.Visible = True
        Call InitFormContent
        SetItemEditable -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1
        mblnChange = False
        Exit Sub
    ElseIf mintType = 1 Then '�޸�
        picNo.Visible = False
        If mint���ó��� = 0 Then
            Call InitFormContent
        Else
            Call LoadDataFromCache
        End If
        SetItemEditable 1, , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, , 1, 1
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitFormContent()
'���ܣ������������ݼ��ؽ���ؼ�
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Integer, lng��������ID As Long

    On Error GoTo errH
    '��ȡ���
    mstr���IDs = GetAdviceDiag(mlngUpdateAdvice, strTmp)
    txtInfo(e_��ǰ���).Text = strTmp
    
    '�Ӹ����л�ȡ���������������Ը���Ϊ׼
    strSQL = "select ���� from ����ҽ������ where ҽ��ID=[1] and ��Ŀ='���뵥���'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
    If Not rsTmp.EOF Then
        txtInfo(e_��ǰ���).Text = rsTmp!���� & ""
    End If
    
    '����ҽ����Ϣ 1����������2������������3������
    strSQL = "select a.id,a.���id,b.id as ��Ŀid,b.���,a.�������,a.�걾��λ as ����ʱ��,b.����,b.����,a.ִ�п���id as ִ�п���id,c.���� as ���ұ���,a.�Ƽ�����," & vbNewLine & _
        "c.���� as ��������,b.�������� as ��ģ,a.��������id, nvl(a.�������,0) as �������,a.��ʼִ��ʱ�� as ��Чʱ��,a.ִ������,a.����ҽ��,a.ִ�б��" & vbNewLine & _
        "from ����ҽ����¼ a,������ĿĿ¼ b,���ű� c" & vbNewLine & _
        "where a.������Ŀid =b.id and a.ִ�п���id=c.id(+) and (a.���id=[1] or a.id=[1]) order by a.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
    
    '1��������
    rsTmp.Filter = "id=" & mlngUpdateAdvice
    If Not rsTmp.EOF Then
        mlng��������ID = Val(rsTmp!��������id & "")
        mlng������ĿID = Val(rsTmp!��ĿID & "")
        mlng����ִ�п������� = Val(rsTmp!ִ������ & "")
        mlng����ִ�п���ID = Val(rsTmp!ִ�п���ID & "")
        
        txtInfo(e_No).Text = rsTmp!������� & ""
        txtInfo(e_����ʱ��).Text = rsTmp!����ʱ��
        
        txtInfo(e_��������).Text = rsTmp!����
        txtInfo(e_��������).Tag = rsTmp!����
        
        txtInfo(e_��������).Text = GetMax�����ȼ�(rsTmp!��ĿID & "")
        txtInfo(e_��Чʱ��).Text = Format(rsTmp!��Чʱ��, "YYYY-MM-DD HH:MM")
		lblInfo(e_��������).Tag = Val(rsTmp!�Ƽ����� & "")
        With cboInfo(e_ִ�п���)
            .Clear
            .AddItem IIF(rsTmp!���ұ��� & "" <> "", rsTmp!���ұ��� & "-" & rsTmp!��������, "-")
            .ItemData(.NewIndex) = mlng����ִ�п���ID
            .AddItem "[����...]"
            .ItemData(.NewIndex) = -1
            .ListIndex = 0
            .Tag = mlng����ִ�п���ID
        End With
        cboInfo(e_�������).ListIndex = Val(rsTmp!������� & "")
        
        txtInfo(e_�������).Text = Sys.RowValue("���ű�", mlng��������ID, "����")
        txtInfo(e_����ҽʦ).Text = rsTmp!����ҽ�� & ""
    End If
    '2����������
    rsTmp.Filter = "id<>" & mlngUpdateAdvice & " and ���='F'"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With vsOper
                .Rows = rsTmp.RecordCount + 1
                .RowData(i) = Val(rsTmp!��ĿID & "")
                .TextMatrix(i, COL_����) = "[" & rsTmp!���� & "]" & rsTmp!����
                .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����)
                .TextMatrix(i, COL_�Ƽ�����) = 0
                .TextMatrix(i, COL_ִ������) = Val(rsTmp!ִ������ & "")
                .TextMatrix(i, COL_��������) = GetMax�����ȼ�(rsTmp!��ĿID & "")
                .TextMatrix(i, COL_��Ҫʱ) = IIF(Val(rsTmp!ִ�б�� & "") = 0, "", "-1")
                If mintType = 1 Then .AddItem ""
            End With
            rsTmp.MoveNext
        Next
    Else
        If mintType = 1 Then vsOper.Rows = 2
    End If
    '3������
    rsTmp.Filter = "���='G'"
    If Not rsTmp.EOF Then
        mlng������ĿID = Val(rsTmp!��ĿID & "")
        mlng����ִ�п������� = Val(rsTmp!ִ������ & "")
        mlng����ִ�п���id = Val(rsTmp!ִ�п���ID & "")
        
        txtInfo(e_������).Text = "[" & rsTmp!���� & "]" & rsTmp!����
        txtInfo(e_������).Tag = txtInfo(e_������).Text
        
        With cboInfo(e_����ִ�п���)
            .Clear
            .AddItem IIF(rsTmp!���ұ��� & "" <> "", rsTmp!���ұ��� & "-" & rsTmp!��������, "-")
            .ItemData(.NewIndex) = mlng����ִ�п���id
            .AddItem "[����...]"
            .ItemData(.NewIndex) = -1
            .ListIndex = 0
            .Tag = mlng����ִ�п���id
        End With
    End If

    Call Init���븽��(False)
    
    '�������븽��ֵ
    strSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��id =[1] Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)

    If rsTmp.EOF Then Exit Sub
    txtInfo(e_����ҽ��).Text = rsTmp!���� & ""
    txtInfo(e_����ҽ��).Tag = txtInfo(e_����ҽ��).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    txtInfo(e_����ҽ������).Text = rsTmp!���� & ""
    txtInfo(e_����ҽ������).Tag = txtInfo(e_����ҽ������).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    txtInfo(e_��һ����).Text = rsTmp!���� & ""
    txtInfo(e_��һ����).Tag = txtInfo(e_��һ����).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    txtInfo(e_�ڶ�����).Text = rsTmp!���� & ""
    txtInfo(e_�ڶ�����).Tag = txtInfo(e_�ڶ�����).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    txtInfo(e_��������).Text = rsTmp!���� & ""
    txtInfo(e_��������).Tag = txtInfo(e_��������).Text
    rsTmp.MoveNext
    
    If rsTmp.EOF Then Exit Sub
    If rsTmp!��Ŀ & "" = "�ٴ���������" Then
        If rsTmp!���� & "" = "�Ǽƻ�" Then
            cboInfo(e_�ٴ���������).ListIndex = 1
        End If
        rsTmp.MoveNext
    End If
    
    With vsOther
        For i = .FixedRows To .Rows - 1
            If rsTmp.EOF Then Exit Sub
            .TextMatrix(i, col_����) = rsTmp!���� & ""
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDataFromCache()
'���ܣ������������ݼ��ؽ���ؼ�
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Integer, lng��������ID As Long
    Dim varA1 As Variant, varA2 As Variant, arrTmp As Variant
    Dim str��ĿIDs As String, int������� As Integer
    Dim str��Чʱ�� As String, str����ʱ�� As String
    Dim lngִ�п���ID As Long, lng����ִ��ID As Long
    Dim str���� As String
    Dim blnDo As Boolean
    Dim str�����Ҫʱ As String
    On Error GoTo errH
    
    '���ػ�������
    If Not mrsCard Is Nothing Then
        If mrsCard.RecordCount > 0 Then
            mrsCard.MoveFirst
            blnDo = True
        End If
    End If
    
    If Not blnDo Then Exit Sub
     
    str��ĿIDs = mrsCard!��������ĿIDs & ""
    str�����Ҫʱ = mrsCard!��������Ҫʱ & ""
    mlng������ĿID = Val(mrsCard!��������ĿID & "")
    If mrsCard!��Чʱ�� & "" <> "" Then
        txtInfo(e_��Чʱ��).Text = mrsCard!��Чʱ��
        txtInfo(e_��Чʱ��).Tag = txtInfo(e_��Чʱ��).Text
    End If
    If mrsCard!����ʱ�� & "" <> "" Then
        txtInfo(e_����ʱ��).Text = mrsCard!����ʱ�� & ""
    End If
     
    lngִ�п���ID = Val(mrsCard!����ִ�п���ID & "")
    lng����ִ��ID = Val(mrsCard!����ִ�п���ID & "")
    str���� = mrsCard!���븽�� & ""
    
    '��ȡ���
    mstr���IDs = mrsCard!�ٴ����IDs & ""
    txtInfo(e_��ǰ���).Text = mrsCard!�ٴ�������� & ""
    
    mlng��������ID = Val(mrsCard!�������id & "")
    mstr�������� = Sys.RowValue("���ű�", mlng��������ID, "����")
    txtInfo(e_�������).Text = mstr��������
    mlng������ĿID = Val(mrsCard!��������ĿID & "")
    mlng����ִ�п���ID = Val(mrsCard!����ִ�п���ID & "")
    Set rsTmp = Get������Ŀ��¼(mlng������ĿID)
    mlng����ִ�п������� = Val(rsTmp!ִ�п��� & "")
    Call Set����ִ�п���(mlng����ִ�п�������, mlng����ִ�п���ID)
    txtInfo(e_��������).Text = rsTmp!����
    txtInfo(e_��������).Tag = rsTmp!����
    
    txtInfo(e_��������).Text = GetMax�����ȼ�(mlng������ĿID)
    cboInfo(e_�������).ListIndex = Val(mrsCard!������� & "")
 
    '2����������
    vsOper.Rows = 2
    If str��ĿIDs <> "" Then
        Set rsTmp = Get������Ŀ��¼ID(0, str��ĿIDs)
        With vsOper
            .Rows = rsTmp.RecordCount + 2
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!ID & "")
                .TextMatrix(i, COL_����) = "[" & rsTmp!���� & "]" & rsTmp!����
                .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����)
                .TextMatrix(i, COL_�Ƽ�����) = 0
                .TextMatrix(i, COL_ִ������) = Val(rsTmp!ִ�п��� & "")
                .TextMatrix(i, COL_��������) = rsTmp!�������� & ""
                .TextMatrix(i, COL_��Ҫʱ) = IIF(InStr(str�����Ҫʱ, .RowData(i) & ":1") > 0, -1, 0)
                rsTmp.MoveNext
            Next
        End With
    End If
    '����
    If Val(mrsCard!������ĿID & "") <> 0 Then
        Set rsTmp = Get������Ŀ��¼(Val(mrsCard!������ĿID & ""))
        mlng������ĿID = Val(rsTmp!ID & "")
        mlng����ִ�п������� = Val(rsTmp!ִ�п��� & "")
        mlng����ִ�п���id = lng����ִ��ID
        txtInfo(e_������).Text = "[" & rsTmp!���� & "]" & rsTmp!����
        txtInfo(e_������).Tag = txtInfo(e_������).Text
        mlng����ִ�п���id = Val(mrsCard!����ִ�п���ID & "")
        Call Set����ִ�п���(mlng����ִ�п�������, mlng����ִ�п���id)
    End If
   
    Call Init���븽��(False)
    
    '�������븽��ֵ
    strSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��id =[1] Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 0)
    If str���� <> "" Then
        Set rsTmp = zlDatabase.CopyNewRec(rsTmp, True)
        varA1 = Split(str����, "<Split1>")
        For i = 0 To UBound(varA1)
            strTmp = varA1(i)
            If InStr(strTmp, "<Split2>") > 0 Then
                varA2 = Split(strTmp, "<Split2>")
                rsTmp.AddNew Array("��Ŀ", "����"), Array(varA2(0), varA2(3))
            End If
        Next
    End If

'    rsTmp.MoveFirst
    rsTmp.Filter = "��Ŀ='����ҽ��'"
    If Not rsTmp.EOF Then
        txtInfo(e_����ҽ��).Text = rsTmp!���� & ""
        txtInfo(e_����ҽ��).Tag = txtInfo(e_����ҽ��).Text
    End If
    
    rsTmp.Filter = "��Ŀ='����ҽ������'"
    If Not rsTmp.EOF Then
        txtInfo(e_����ҽ������).Text = rsTmp!���� & ""
        txtInfo(e_����ҽ������).Tag = txtInfo(e_����ҽ������).Text
    End If
    
    rsTmp.Filter = "��Ŀ='��һ����'"
    If Not rsTmp.EOF Then
        txtInfo(e_��һ����).Text = rsTmp!���� & ""
        txtInfo(e_��һ����).Tag = txtInfo(e_��һ����).Text
    End If
    
    rsTmp.Filter = "��Ŀ='�ڶ�����'"
    If Not rsTmp.EOF Then
        txtInfo(e_�ڶ�����).Text = rsTmp!���� & ""
        txtInfo(e_�ڶ�����).Tag = txtInfo(e_�ڶ�����).Text
    End If
    
    rsTmp.Filter = "��Ŀ='��������'"
    If Not rsTmp.EOF Then
        txtInfo(e_��������).Text = rsTmp!���� & ""
        txtInfo(e_��������).Tag = txtInfo(e_��������).Text
    End If
    
    With vsOther
        For i = .FixedRows To .Rows - 1
            rsTmp.Filter = "��Ŀ='" & .TextMatrix(i, COL_��Ŀ����) & "'"
            If Not rsTmp.EOF Then
                .TextMatrix(i, col_����) = rsTmp!���� & ""
            End If
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check����() As Boolean
'���ܣ�������븽���еı�����Ŀ
    Dim i As Long
    
    On Error GoTo errH
    
    mrsAppend.Filter = 0
    
    If Not mrsAppend.EOF Then
        Check���� = False
        
        mrsAppend.Filter = "������='����ҽ��'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!���� & "") = 1 And txtInfo(e_����ҽ��).Text = "" Then
                MsgBox """����ҽ��""��ĿΪ������Ŀ��", vbInformation, Me.Caption
                txtInfo(e_����ҽ��).SetFocus
                Exit Function
            End If
        End If
        
        mrsAppend.Filter = "������='����ҽ������'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!���� & "") = 1 And txtInfo(e_����ҽ������).Text = "" Then
                MsgBox """����ҽ������""��ĿΪ������Ŀ��", vbInformation, Me.Caption
                txtInfo(e_����ҽ������).SetFocus
                Exit Function
            End If
        End If
        
        mrsAppend.Filter = "������='��һ����'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!���� & "") = 1 And txtInfo(e_��һ����).Text = "" Then
                MsgBox """��һ����""��ĿΪ������Ŀ��", vbInformation, Me.Caption
                txtInfo(e_��һ����).SetFocus
                Exit Function
            End If
        End If
        
        mrsAppend.Filter = "������='�ڶ�����'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!���� & "") = 1 And txtInfo(e_�ڶ�����).Text = "" Then
                MsgBox """�ڶ�����""��ĿΪ������Ŀ��", vbInformation, Me.Caption
                txtInfo(e_�ڶ�����).SetFocus
                Exit Function
            End If
        End If
        
        mrsAppend.Filter = "������='��������'"
        If Not mrsAppend.EOF Then
            If Val(mrsAppend!���� & "") = 1 And txtInfo(e_��������).Text = "" Then
                MsgBox """��������""��ĿΪ������Ŀ��", vbInformation, Me.Caption
                txtInfo(e_��������).SetFocus
                Exit Function
            End If
        End If
        
        If vsOther.Visible Then
            With vsOther
                For i = .FixedRows To .Rows - 1
                    If Val(.RowData(i)) <> 0 Then
                        mrsAppend.Filter = "���=" & Val(.RowData(i))
                        If Not mrsAppend.EOF Then
                            If Val(mrsAppend!���� & "") = 1 And .TextMatrix(i, col_����) = "" Then
                                MsgBox """" & mrsAppend!��Ŀ & """��ĿΪ������Ŀ��", vbInformation, Me.Caption
                                .Row = i
                                .Col = col_����
                                .SetFocus
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End With
        End If
    End If
    
    Check���� = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Init���븽��(Optional ByVal blnȡֵ As Boolean = True)
'���ܣ�������������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strFilter As String, i As Integer, intIndex As Integer
    Dim bln��չ As Boolean, strText As String
    
    vsOther.Visible = False
    lblInfo(e_��������).Visible = False
    
    strSQL = "Select C.���� as ���,C.��Ŀ,C.����,C.Ҫ��ID,C.����,D.������" & _
        " From ��������Ӧ�� A,�����ļ��б� B,�������ݸ��� C,����������Ŀ D" & _
        " Where A.������ĿID=[1] And A.Ӧ�ó���=[2]" & _
        " And A.�����ļ�ID=B.ID And B.����=7 And B.ID=C.�ļ�ID And c.Ҫ��id=d.id(+)" & _
        " Order by C.����"
    
    On Error GoTo errH
    
    Set mrsAppend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng������ĿID, mint�������)
    
    If Not mrsAppend.EOF Then
        lblInfo(e_����ҽ������).Tag = ""
        lblInfo(e_����ҽ��).Tag = ""
        lblInfo(e_��һ����).Tag = ""
        lblInfo(e_�ڶ�����).Tag = ""
        lblInfo(e_��������).Tag = ""
        
        mrsAppend.Filter = "������='����ҽ��'"
        
        mrsAppend.Sort = "���"
        If Not mrsAppend.EOF Then
            lblInfo(e_����ҽ��).Tag = Val(mrsAppend!��� & "")
            lblInfo(e_����ҽ��).ToolTipText = mrsAppend!��Ŀ & ""
            strFilter = strFilter & " And ���<>" & Val(mrsAppend!��� & "")
            If blnȡֵ Then
                txtInfo(e_����ҽ��).Text = GetItemAppend(Val(mrsAppend!Ҫ��ID & ""), mrsAppend!������ & "", mrsAppend!��Ŀ & "")
                txtInfo(e_����ҽ��).Tag = txtInfo(e_����ҽ��).Text
            End If
        End If
        
        mrsAppend.Filter = "������='����ҽ������'"
        mrsAppend.Sort = "���"
        
        If Not mrsAppend.EOF Then
            lblInfo(e_����ҽ������).Tag = Val(mrsAppend!��� & "")
            lblInfo(e_����ҽ������).ToolTipText = mrsAppend!��Ŀ & ""
            strFilter = strFilter & " And ���<>" & Val(mrsAppend!��� & "")
            
            If blnȡֵ Then
                txtInfo(e_����ҽ������).Text = GetItemAppend(Val(mrsAppend!Ҫ��ID & ""), mrsAppend!������ & "", mrsAppend!��Ŀ & "")
                txtInfo(e_����ҽ������).Tag = txtInfo(e_����ҽ������).Text
            End If
        End If
        
        mrsAppend.Filter = "������='����ҽ��'"
        mrsAppend.Sort = "���"
        If Not mrsAppend.EOF Then
            For i = 1 To mrsAppend.RecordCount
                If i = 1 Then
                    intIndex = e_��һ����
                ElseIf i = 2 Then
                    intIndex = e_�ڶ�����
                ElseIf i = 3 Then
                    intIndex = e_��������
                Else
                    Exit For
                End If
                
                lblInfo(intIndex).Tag = Val(mrsAppend!��� & "")
                lblInfo(intIndex).ToolTipText = mrsAppend!��Ŀ & ""
                
                strFilter = strFilter & " And ���<>" & Val(mrsAppend!��� & "")
                
                If blnȡֵ Then
                    txtInfo(intIndex).Text = GetItemAppend(Val(mrsAppend!Ҫ��ID & ""), mrsAppend!������ & "", mrsAppend!��Ŀ & "")
                    txtInfo(intIndex).Tag = txtInfo(intIndex).Text
                End If
                mrsAppend.MoveNext
            Next
        End If
        
        If strFilter <> "" Then
            strFilter = Mid(strFilter, 5)
            mrsAppend.Filter = strFilter
        Else
            mrsAppend.Filter = 0
        End If
        mrsAppend.Sort = "���"
        
        If Not mrsAppend.EOF Then
            '������չ���
            vsOther.Visible = True
            lblInfo(e_��������).Visible = True
            Me.Height = 10200
            i = mrsAppend.RecordCount
            With vsOther
                .Clear
                .Rows = IIF(i = 0, 2, i + 1)
                .Cols = 2
                .FixedRows = 1: .FixedCols = 1
                
                .ColAlignment(COL_��Ŀ����) = 1
                .FixedAlignment(COL_��Ŀ����) = 4
                .ColWidth(COL_��Ŀ����) = 2000
                .TextMatrix(0, COL_��Ŀ����) = "���븽����չ"
                
                .ColAlignment(col_����) = 1
                .FixedAlignment(col_����) = 4
                 .ColWidth(col_����) = 6200
                 .TextMatrix(0, col_����) = "����"
                 
                .Editable = flexEDKbdMouse
                
                For i = 1 To mrsAppend.RecordCount
                    .RowData(i) = Val(mrsAppend!��� & "")
                    .TextMatrix(i, COL_��Ŀ����) = mrsAppend!��Ŀ & ""
                    
                    If blnȡֵ Then .TextMatrix(i, col_����) = GetItemAppend(Val(mrsAppend!Ҫ��ID & ""), mrsAppend!������ & "", mrsAppend!��Ŀ & "")
                    
                    mrsAppend.MoveNext
                Next
                .Row = 1: .Col = 0
            End With
        Else
            Me.Height = 8500
        End If
    End If
    'û�а���ȡ�������̶�����
    '���û�а󶨸���ҳ�Ҫ��ID������ Tag �С�����ҽ����ͬһ��Ҫ�أ�ֻ��Ҫ�ж�  ����һ���֡� ����
    If lblInfo(e_����ҽ������).Tag = "" Or lblInfo(e_����ҽ��).Tag = "" Or lblInfo(e_��һ����).Tag = "" Then
        strSQL = "Select i.Id As Ҫ��id, i.������ From ����������Ŀ I, ������������ K" & _
            " Where i.����id = k.Id And k.���� = 1 And k.���� = '06' And i.������ In ('����ҽ��', '����ҽ��', '����ҽ������')"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        For i = 1 To rsTmp.RecordCount
            If rsTmp!������ & "" = "����ҽ��" And lblInfo(e_����ҽ��).Tag = "" Then
                lblInfo(e_����ҽ��).Tag = Val(rsTmp!Ҫ��ID & "")
                If txtInfo(e_����ҽ��).Text = "" Then
                    txtInfo(e_����ҽ��).Text = GetItemAppend(Val(rsTmp!Ҫ��ID & ""), "����ҽ��", "����ҽ��")
                    txtInfo(e_����ҽ��).Tag = txtInfo(e_����ҽ��).Text
                End If
            ElseIf rsTmp!������ & "" = "����ҽ������" And lblInfo(e_����ҽ������).Tag = "" Then
                lblInfo(e_����ҽ������).Tag = Val(rsTmp!Ҫ��ID & "")
                If txtInfo(e_����ҽ������).Text = "" Then
                    txtInfo(e_����ҽ������).Text = GetItemAppend(Val(rsTmp!Ҫ��ID & ""), "����ҽ������", "����ҽ������")
                    txtInfo(e_����ҽ������).Tag = txtInfo(e_����ҽ������).Text
                End If
            ElseIf rsTmp!������ & "" = "����ҽ��" And lblInfo(e_��һ����).Tag = "" Then
                lblInfo(e_��һ����).Tag = Val(rsTmp!Ҫ��ID & "")
            End If
            rsTmp.MoveNext
        Next
    End If
    Call Form_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Select Case Index
        Case e_����ʱ��
            Call cmdDate_Click(e_����ʱ��)
        Case e_��Чʱ��
            Call cmdDate_Click(e_��Чʱ��)
        End Select
    End If
End Sub

Private Sub vsOper_DblClick()
    Call vsOper_KeyPress(32)
End Sub

Private Sub vsOper_GotFocus()
    Call vsOper_AfterRowColChange(vsOper.Row, vsOper.Col, vsOper.Row, vsOper.Col)
End Sub

Private Sub vsOper_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�ɾ��������
    Dim i As Long
    If mintType = 2 Then Exit Sub
    If KeyCode = vbKeyDelete Then
        With vsOper
            If .RowData(.Row) <> 0 Then
                If MsgBox("Ҫɾ����ǰ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                .RowData(.Row) = 0
                For i = 0 To .Cols - 1
                    .TextMatrix(.Row, i) = ""
                    .Cell(flexcpData, .Row, i) = ""
                Next
                If Not (.Rows = .FixedRows + 1 And .Row = .FixedRows) Then .RemoveItem .Row
            End If
        End With
    End If
End Sub

Private Sub vsOper_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ��ǻس�ȷ�����༭�Ĵ���(����Text:=EditText,��ValidateEdit�¼��л�û��)
    If Not mblnReturn Then
        If Col = 0 Then
            vsOper.TextMatrix(Row, Col) = CStr(vsOper.Cell(flexcpData, Row, Col))
            Call vsOper_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
        End If
    End If
    mblnChange = True
End Sub

Private Sub vsOper_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'''''''''''''
    If mintType = 2 Then Exit Sub
    
    If NewCol = COL_���� Then
        vsOper.Editable = flexEDKbdMouse
        vsOper.ComboList = "..."
        vsOper.FocusRect = flexFocusLight
    ElseIf NewCol = COL_��Ҫʱ And vsOper.TextMatrix(NewRow, COL_����) <> "" Then
        vsOper.ComboList = ""
        vsOper.Editable = flexEDKbdMouse
    Else
        vsOper.ComboList = ""
        vsOper.Editable = flexEDNone
    End If
End Sub

Private Sub vsOper_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'���ܣ�ֱ�Ӵ򿪸���������Ŀѡ����
    Dim strSQLItem As String, strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim int�Ա� As Integer
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    int�Ա� = GetType�Ա�
 
    strSQLItem = " From ������ĿĿ¼ A , ������϶��� B, ��������Ŀ¼ C Where A.���='F' And A.ID<>-1*" & mlng������ĿID & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) And a.Id = b.����id(+) And b.����id = c.Id(+)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[4])" & _
            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
            " And A.������� IN([1],3) And Nvl(A.ִ��Ƶ��,0) IN(0,[2]) And Nvl(A.�����Ա�,0) IN(0,[3])"
    
    strSQL = "Select 0 as ĩ��,Max(Level) as ��ID,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ��������,NULL as ִ�п�������ID,null as �Ƽ�����ID" & _
        " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Start With ID In (Select a.����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID" & _
        " Group by ID,�ϼ�ID,����,����  Union ALL" & _
        " Select 1 as ĩ��,1 as ��ID,A.ID,a.����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,c.�������� As ��������,A.ִ�п��� as ִ�п�������ID,a.�Ƽ����� as �Ƽ�����ID" & _
        strSQLItem & " Order By ĩ��,��ID Desc,����"
        
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
        mint�������, 1, int�Ա�, mlng���˿���id)
        
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "δ�ҵ����õ�������Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
        End If
        vsOper.TextMatrix(Row, Col) = CStr(vsOper.Cell(flexcpData, Row, Col))
        Call vsOper_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
        vsOper.SetFocus
        Exit Sub
    End If
    
    Call Set��������(vsOper.Row, rsTmp)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsOper_KeyPress(KeyAscii As Integer)
    If mintType = 2 Then Exit Sub
    With vsOper
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call SeekNextCtl
        Else
            If .Col = 0 Then
                .Editable = flexEDKbdMouse
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsOper_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            ElseIf .Col = COL_��Ҫʱ And .TextMatrix(.Row, COL_����) <> "" Then
                .ComboList = ""
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End If
    End With
End Sub

Private Sub vsOper_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If mintType = 2 Then Exit Sub
    vsOper.EditSelStart = 0
    vsOper.EditSelLength = zlCommFun.ActualLen(vsOper.EditText)
End Sub

Private Sub vsOper_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ���ʼ�༭ʱ��Ϊû�а��»س�
    If mintType = 2 Then
        Cancel = True
        Exit Sub
    End If
    If Col = COL_��Ҫʱ And vsOper.TextMatrix(Row, COL_����) <> "" Then
        vsOper.ComboList = ""
        vsOper.Editable = flexEDKbdMouse
    End If
    mblnReturn = False
End Sub
    

Private Sub vsOper_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'���ܣ�����������������ȷ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, int�Ա� As Integer
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI, strLike As String
    
    On Error GoTo errH
    If mintType = 2 Then Exit Sub
    If KeyAscii = 13 Then
    
        mblnReturn = True '����ǰ��س�ȷ�ϱ༭
        
        KeyAscii = 0
        
        int�Ա� = GetType�Ա�
        
        '�Ż�
        strLike = gstrLike
        If Len(vsOper.EditText) < 2 Then strLike = ""
        
        strSQL = " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,d.�������� As ��������,A.ִ�п��� as ִ�п�������ID,a.�Ƽ����� as �Ƽ�����ID" & _
            " From ������ĿĿ¼ A,������Ŀ���� B, ������϶��� C, ��������Ŀ¼ D" & _
            " Where A.ID=B.������ĿID And A.���='F' And A.ID<>-1*[3]" & IIF(strLike = "", "", " And Rownum<=100") & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) And a.Id = c.����id(+) And c.����id = d.Id(+)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]) And B.����=[4]" & _
            " And A.������� IN([5],3) And Nvl(A.ִ��Ƶ��,0) IN(0,[6]) And Nvl(A.�����Ա�,0) IN(0,[7])" & _
            " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[8])" & _
            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
            " Order by A.����"
        vPoint = zlControl.GetCoordPos(vsOper.hwnd, vsOper.CellLeft, vsOper.CellTop)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, False, True, vPoint.X, vPoint.Y, vsOper.CellHeight, blnCancel, False, True, _
            UCase(vsOper.EditText) & "%", strLike & UCase(vsOper.EditText) & "%", mlng������ĿID, gbytCode + 1, mint�������, 1, int�Ա�, mlng���˿���id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
            End If
            vsOper.TextMatrix(Row, Col) = CStr(vsOper.Cell(flexcpData, Row, Col))
            Call vsOper_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
            vsOper.SetFocus
            Exit Sub
        End If
        Call Set��������(Row, rsTmp)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsOther_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If mintType = 2 Then Exit Sub
    If KeyAscii = 13 Then
        With vsOther
            If Row = .Rows - 1 Then
                Call SeekNextCtl
            Else
                .Row = Row + 1
                .Col = col_����
                .Refresh
            End If
        End With
    End If
End Sub

Private Sub vsOther_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mintType = 2 Then Cancel = True: Exit Sub
    mblnChange = True
End Sub

Private Function GetType�Ա�() As Integer
'���ܣ����ݵ�ǰ�����Ա��ȡ �����Ա�
    If txtInfo(e_�Ա�).Text Like "*��*" Then
        GetType�Ա� = 1
    ElseIf txtInfo(e_�Ա�).Text Like "*Ů*" Then
        GetType�Ա� = 2
    End If
End Function

Private Sub Set��������(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    Dim i As Integer
    Dim strMsg As String
    
    '��������
    With vsOper
        '����ظ�����
     
        If Val(rsInput!ID & "") = mlng������ĿID Then
            strMsg = "�ø���������������Ŀ��ͬ��"
        Else
            For i = .FixedRows To .Rows - 1
                If .RowData(i) = Val(rsInput!ID) Then
                    strMsg = "�ø��������Ѿ���������¼�롣"
                    Exit For
                End If
            Next
        End If
        
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            .TextMatrix(lngRow, COL_����) = CStr(.Cell(flexcpData, lngRow, COL_����))
            Call vsOper_AfterRowColChange(lngRow, COL_����, lngRow, COL_����) '����ʹ��ť�ɼ�
            .SetFocus
            Exit Sub
        End If
    
        .EditText = "[" & rsInput!���� & "]" & rsInput!���� '��������ֱ��ƥ��ʱ�б�Ҫ
        
        .RowData(lngRow) = Val(rsInput!ID)
        
        .TextMatrix(lngRow, COL_����) = "[" & rsInput!���� & "]" & rsInput!����
        .Cell(flexcpData, lngRow, COL_����) = .TextMatrix(lngRow, COL_����)
        
        .TextMatrix(lngRow, COL_��������) = NVL(rsInput!��������)
        
        .TextMatrix(lngRow, COL_ִ������) = Val(rsInput!ִ�п�������ID & "")
        .TextMatrix(lngRow, COL_�Ƽ�����) = Val(rsInput!�Ƽ�����ID & "")
        
        '��һ������
        If .RowData(.Rows - 1) <> 0 Then .AddItem ""
        .Row = .Rows - 1: .Col = COL_����
        Call .ShowCell(.Row, .Col)
    End With
    
    mblnChange = True
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
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " ����(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, " �����˳�(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " �˳�(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call ReportPrintSet(gcnOracle, glngSys, "ZL1_INSIDE_1254_18", Me)
        Case conMenu_File_Preview: Call PrintApply(1)
        Case conMenu_File_Print: Call PrintApply(2)
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit '����
            If CheckData = False Then Exit Sub
            If mint���ó��� = 0 Then
                mblnOK = SaveData
            Else
                mblnOK = SaveCacheData
            End If
            If Control.ID = conMenu_Edit_SaveExit Then
                Unload Me
            End If
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnVisible As Boolean
    
    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Save, conMenu_Edit_SaveExit
            Control.Enabled = mblnChange
        Case conMenu_File_PrintSet, conMenu_File_Preview, conMenu_File_Print
            blnVisible = mint���ó��� = 0
    End Select
    Control.Visible = blnVisible
End Sub


Private Sub Form_Load()
    
    mblnOK = False
    
    mbln���Ѷ��� = True
    If mint������� = 0 Then mint������� = 2 'ȱʡΪסԺ
 
    '��Ժҽ�����뽨��
    mbln��Ժҽ������ = Val(zlDatabase.GetPara(253, glngSys)) <> 0
        
    Call LoadPatiInfo
    
    Me.Height = 8500
    
    Call InitCommandBar
    
    With cboInfo(e_�������)
        .Clear
        .AddItem "����"
        .AddItem "����"
        .AddItem "����"
        .ListIndex = 0
    End With
    
    With cboInfo(e_�ٴ���������)
        .Clear
        .AddItem "�ƻ�"
        .AddItem "�Ǽƻ�"
        .ListIndex = 0
    End With
    
    With vsOper
        .Clear
        .Rows = 2
        .Cols = 6
        .FixedRows = 1: .FixedCols = 0
        
        .ColAlignment(COL_����) = 1
        .FixedAlignment(COL_����) = 4
        .ColWidth(COL_����) = 6000
        .TextMatrix(0, COL_����) = "��������"
        
        .ColAlignment(COL_��������) = 1
        .FixedAlignment(COL_��������) = 4
         .ColWidth(COL_��������) = 1500
         .TextMatrix(0, COL_��������) = "����"
         
         
        .ColAlignment(COL_��Ҫʱ) = 1
        .FixedAlignment(COL_��Ҫʱ) = 4
        .ColWidth(COL_��Ҫʱ) = 1000
        .TextMatrix(0, COL_��Ҫʱ) = "��Ҫʱ"
        .ColDataType(COL_��Ҫʱ) = flexDTBoolean
         
        .ColHidden(COL_�Ƽ�����) = True
        .ColHidden(COL_ִ������) = True
        .ColHidden(COL_ժҪ) = True
        
        .Editable = flexEDKbdMouse
        .Row = 1: .Col = 0
    End With
    Call InitDefault
    mblnChange = False
End Sub

Private Function SeekNextCtl() As Boolean
'���ܣ���λ����һ������Ŀؼ���
    Call zlCommFun.PressKey(vbKeyTab)
    SeekNextCtl = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("��ǰ���뵥�Ѿ������˵�����δ���棬�Ƿ�Ҫ�����˳���", vbQuestion + vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
    If Not mclsDiagEdit Is Nothing Then Set mclsDiagEdit = Nothing
    mlng������ĿID = 0
    mbln��¼ = False
    mstr��Ժʱ�� = ""
    mstr�ϴ�ת��ʱ�� = ""
    mint���� = 0
    Set mclsMipModule = Nothing
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub Form_Resize()
    Dim lngL As Long
    Dim lngT As Long
    Dim lngDY As Long
    Dim lngTmp As Long

    On Error Resume Next
    
    lngL = 480
    lngDY = 200
    
    lblHead.Top = 750
    linHead.Y1 = lblHead.Height + lblHead.Top + 20
    linHead.Y2 = linHead.Y1
    
    picNo.Top = linHead.Y1 + 30
    picNo.Left = Me.ScaleWidth - picNo.Width - 400
    Call SetFaceItemPos(e_No)
    
    
    lblInfo(e_����).Left = lngL
    lblInfo(e_����).Top = picNo.Top + picNo.Height + 150
    
        lblInfo(e_�Ա�).Left = lngL + 2695
        lblInfo(e_�Ա�).Top = lblInfo(e_����).Top
        
        lblInfo(e_����).Left = lngL + 5340
        lblInfo(e_����).Top = lblInfo(e_����).Top
        
        lblInfo(e_����).Left = lngL + 7785
        lblInfo(e_����).Top = lblInfo(e_����).Top
        
        txtInfo(e_����).Width = 800
        
        Call SetFaceItemPos(e_����)
        Call SetFaceItemPos(e_�Ա�)
        Call SetFaceItemPos(e_����)
        Call SetFaceItemPos(e_����)
    
    
    lblInfo(e_����).Left = lngL
    lblInfo(e_����).Top = Line1(e_����).Y1 + lngDY
    
        lblInfo(e_סԺ��).Left = lblInfo(e_�Ա�).Left
        lblInfo(e_סԺ��).Top = lblInfo(e_����).Top
        
        lblInfo(e_����).Left = lblInfo(e_����).Left
        lblInfo(e_����).Top = lblInfo(e_����).Top
        
        txtInfo(e_סԺ��).Width = txtInfo(e_����).Width
        txtInfo(e_����).Width = txtInfo(e_����).Width
        
        Call SetFaceItemPos(e_����)
        Call SetFaceItemPos(e_סԺ��)
        Call SetFaceItemPos(e_����)
        
        
    lblInfo(e_����ʱ��).Left = lngL - 30
    lblInfo(e_����ʱ��).Top = Line1(e_����).Y1 + lngDY
        Call SetFaceItemPos(e_����ʱ��)
        cmdDate(e_����ʱ��).Top = txtInfo(e_����ʱ��).Top - 30
        cmdDate(e_����ʱ��).Left = txtInfo(e_����ʱ��).Left + txtInfo(e_����ʱ��).Width
        
        lblInfo(e_�ٴ���������).Top = lblInfo(e_����ʱ��).Top
        lblInfo(e_�ٴ���������).Left = lblInfo(e_�������).Left + lblInfo(e_�������).Width - lblInfo(e_�ٴ���������).Width
        picInfo(e_�ٴ���������).Top = lblInfo(e_�ٴ���������).Top
        picInfo(e_�ٴ���������).Left = lblInfo(e_�ٴ���������).Left + lblInfo(e_�ٴ���������).Width + 160
        Line1(e_�ٴ���������).Y1 = picInfo(e_�ٴ���������).Top + picInfo(e_�ٴ���������).Height
        Line1(e_�ٴ���������).Y2 = Line1(e_�ٴ���������).Y1
        Line1(e_�ٴ���������).X1 = picInfo(e_�ٴ���������).Left - 100
        Line1(e_�ٴ���������).X2 = Line1(e_�ٴ���������).X1 + picInfo(e_�ٴ���������).Width
        
    lblInfo(e_��ǰ���).Left = lngL
    lblInfo(e_��ǰ���).Top = Line1(e_����ʱ��).Y1 + lngDY
        txtInfo(e_��ǰ���).Width = Me.ScaleWidth - txtInfo(e_��ǰ���).Left - 700
        Call SetFaceItemPos(e_��ǰ���)
        cmdInfo(e_��ǰ���).Top = txtInfo(e_��ǰ���).Top - 30
        cmdInfo(e_��ǰ���).Left = txtInfo(e_��ǰ���).Left + txtInfo(e_��ǰ���).Width
        
        
    lblInfo(e_��������).Left = lngL - 30
    lblInfo(e_��������).Top = Line1(e_��ǰ���).Y1 + lngDY
        Call SetFaceItemPos(e_��������)
        cmdInfo(e_��������).Top = txtInfo(e_��������).Top - 30
        cmdInfo(e_��������).Left = txtInfo(e_��������).Left + txtInfo(e_��������).Width
        
        lblInfo(e_��������).Left = lblInfo(e_����).Left
        lblInfo(e_��������).Top = lblInfo(e_��������).Top
        txtInfo(e_��������).Width = 800
        Call SetFaceItemPos(e_��������)
        
        
    lblInfo(e_ִ�п���).Left = lngL
    lblInfo(e_ִ�п���).Top = Line1(e_��������).Y1 + lngDY
        picInfo(e_ִ�п���).Left = lblInfo(e_ִ�п���).Left + lblInfo(e_ִ�п���).Width + 160
        picInfo(e_ִ�п���).Top = lblInfo(e_ִ�п���).Top
        picInfo(e_ִ�п���).Width = 2400
        picInfo(e_ִ�п���).Height = 250
        cboInfo(e_ִ�п���).Width = 2450
        
        Line1(e_ִ�п���).Y1 = picInfo(e_ִ�п���).Top + picInfo(e_ִ�п���).Height
        Line1(e_ִ�п���).Y2 = Line1(e_ִ�п���).Y1
        Line1(e_ִ�п���).X1 = picInfo(e_ִ�п���).Left - 100
        Line1(e_ִ�п���).X2 = Line1(e_ִ�п���).X1 + picInfo(e_ִ�п���).Width + 110
        
        lblInfo(e_�������).Left = lblInfo(e_����).Left
        lblInfo(e_�������).Top = lblInfo(e_ִ�п���).Top
        picInfo(e_�������).Left = lblInfo(e_�������).Left + lblInfo(e_�������).Width + 160
        picInfo(e_�������).Top = lblInfo(e_�������).Top
        picInfo(e_�������).Width = 1020
        picInfo(e_�������).Height = 250
        cboInfo(e_�������).Width = 1070
        Line1(e_�������).Y1 = picInfo(e_�������).Top + picInfo(e_�������).Height
        Line1(e_�������).Y2 = Line1(e_�������).Y1
        Line1(e_�������).X1 = picInfo(e_�������).Left - 100
        Line1(e_�������).X2 = Line1(e_�������).X1 + picInfo(e_�������).Width + 100
        
                
    lblInfo(e_��������).Left = lngL
    lblInfo(e_��������).Top = Line1(e_ִ�п���).Y1 + lngDY
        vsOper.Top = lblInfo(e_��������).Top
        vsOper.Left = picInfo(e_ִ�п���).Left
    
    
    lblInfo(e_������).Left = lngL
    lblInfo(e_������).Top = vsOper.Top + vsOper.Height + lngDY
        txtInfo(e_������).Width = 3800
        
        Call SetFaceItemPos(e_������)
        cmdInfo(e_������).Top = txtInfo(e_������).Top - 30
        cmdInfo(e_������).Left = txtInfo(e_������).Left + txtInfo(e_������).Width
        
        lblInfo(e_����ִ�п���).Left = lblInfo(e_����).Left
        lblInfo(e_����ִ�п���).Top = lblInfo(e_������).Top
        picInfo(e_����ִ�п���).Left = lblInfo(e_����ִ�п���).Left + lblInfo(e_����ִ�п���).Width + 160
        picInfo(e_����ִ�п���).Top = lblInfo(e_����ִ�п���).Top
        picInfo(e_����ִ�п���).Width = 2400
        picInfo(e_����ִ�п���).Height = 250
        cboInfo(e_����ִ�п���).Width = 2450
        
        Line1(e_����ִ�п���).Y1 = picInfo(e_����ִ�п���).Top + picInfo(e_����ִ�п���).Height + 10
        Line1(e_����ִ�п���).Y2 = Line1(e_����ִ�п���).Y1
        Line1(e_����ִ�п���).X1 = picInfo(e_����ִ�п���).Left - 100
        Line1(e_����ִ�п���).X2 = Line1(e_����ִ�п���).X1 + picInfo(e_����ִ�п���).Width + 100
        
    
    lblInfo(e_����ҽ��).Left = lngL
    lblInfo(e_����ҽ��).Top = Line1(e_������).Y1 + lngDY
        Call SetFaceItemPos(e_����ҽ��)
        cmdInfo(e_����ҽ��).Top = txtInfo(e_����ҽ��).Top - 30
        cmdInfo(e_����ҽ��).Left = txtInfo(e_����ҽ��).Left + txtInfo(e_����ҽ��).Width
            
        lblInfo(e_����ҽ������).Left = lblInfo(e_סԺ��).Left
        lblInfo(e_����ҽ������).Top = lblInfo(e_����ҽ��).Top
        Call SetFaceItemPos(e_����ҽ������)
        cmdInfo(e_����ҽ������).Top = txtInfo(e_����ҽ������).Top - 30
        cmdInfo(e_����ҽ������).Left = txtInfo(e_����ҽ������).Left + txtInfo(e_����ҽ������).Width
        
    lblInfo(e_��һ����).Left = lngL
    lblInfo(e_��һ����).Top = Line1(e_����ҽ��).Y1 + lngDY
        Call SetFaceItemPos(e_��һ����)
        cmdInfo(e_��һ����).Top = txtInfo(e_��һ����).Top - 30
        cmdInfo(e_��һ����).Left = txtInfo(e_��һ����).Left + txtInfo(e_��һ����).Width
        
        lblInfo(e_�ڶ�����).Left = lblInfo(e_סԺ��).Left
        lblInfo(e_�ڶ�����).Top = lblInfo(e_��һ����).Top
        txtInfo(e_�ڶ�����).Width = txtInfo(e_��һ����).Width
        Call SetFaceItemPos(e_�ڶ�����)
        cmdInfo(e_�ڶ�����).Top = txtInfo(e_�ڶ�����).Top - 30
        cmdInfo(e_�ڶ�����).Left = txtInfo(e_�ڶ�����).Left + txtInfo(e_�ڶ�����).Width
        
        lblInfo(e_��������).Left = lblInfo(e_����).Left
        lblInfo(e_��������).Top = lblInfo(e_��һ����).Top
        Call SetFaceItemPos(e_��������)
        cmdInfo(e_��������).Top = txtInfo(e_��������).Top - 30
        cmdInfo(e_��������).Left = txtInfo(e_��������).Left + txtInfo(e_��������).Width
        
    
    lblInfo(e_��������).Left = lngL
    lblInfo(e_��������).Top = Line1(e_��һ����).Y1 + lngDY
        vsOther.Top = lblInfo(e_��������).Top
        vsOther.Left = picInfo(e_ִ�п���).Left
    
    If lblInfo(e_��������).Visible Then
        lngTmp = vsOther.Top + vsOther.Height + lngDY
    Else
        lngTmp = Line1(e_��һ����).Y1 + lngDY
    End If
    
    lblInfo(e_�������).Left = lngL
    lblInfo(e_�������).Top = lngTmp 'vsOther.Top + vsOther.Height + lngDY
        Call SetFaceItemPos(e_�������)
        
        lblInfo(e_����ҽʦ).Top = lblInfo(e_�������).Top
        lblInfo(e_����ҽʦ).Left = 4200
        Call SetFaceItemPos(e_����ҽʦ)
        
        lblInfo(e_����ҽʦ).Top = lblInfo(e_�������).Top
        lblInfo(e_����ҽʦ).Left = 7800
        Call SetFaceItemPos(e_����ҽʦ)
        
        
    lblInfo(e_��Чʱ��).Top = Line1(e_�������).Y1 + lngDY
        lblInfo(e_��Чʱ��).Left = 7050
        Call SetFaceItemPos(e_��Чʱ��)
        cmdDate(e_��Чʱ��).Top = txtInfo(e_��Чʱ��).Top - 30
        cmdDate(e_��Чʱ��).Left = txtInfo(e_��Чʱ��).Left + txtInfo(e_��Чʱ��).Width
        
End Sub

Private Sub SetFaceItemPos(ByVal lngIndex As Long)
'���ܣ����ý��� ��ǩ���ı����»��ߵ�λ��
'������lngIndex �ؼ��±�����
    Dim lngL As Long
    Dim lngT As Long
    
    lngL = lblInfo(lngIndex).Left
    lngT = lblInfo(lngIndex).Top
    
    txtInfo(lngIndex).Top = lngT
    txtInfo(lngIndex).Left = lngL + lblInfo(lngIndex).Width + 160
    txtInfo(lngIndex).Height = 200
    Line1(lngIndex).Y1 = lngT + txtInfo(lngIndex).Height + 10
    Line1(lngIndex).Y2 = Line1(lngIndex).Y1
    
    Line1(lngIndex).X1 = txtInfo(lngIndex).Left - 100
    Line1(lngIndex).X2 = 110 + txtInfo(lngIndex).Width + Line1(lngIndex).X1

End Sub

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
    Call mobjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_18", Me, "ҽ��ID=" & mlngUpdateAdvice, "����ID=" & mlng������ĿID, intType)
End Sub

Private Function CheckData() As Boolean
'���ܣ����������ȷ��
    Dim strIDs As String, strҽ������ As String, strMsg As String
    Dim lngTmp As Long, i As Integer
    Dim vMsg As VbMsgBoxResult
    Dim strExtra As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngִ������ As Long
    Dim lngִ�п���ID As Long
    Dim lng����ID As Long
    Dim strTmp As String
    Dim j As Long
    Dim strTabAdvice As String
    Dim rsPrice As ADODB.Recordset
    
    Call Me.ValidateControls
    '����¼����������Ŀ
    If mlng������ĿID = 0 Then
        MsgBox "û��ȷ����������Ŀ��", vbInformation, Me.Caption
        If txtInfo(e_��������).Enabled Then txtInfo(e_��������).SetFocus
        Exit Function
    End If
    
    '���ִ�п���
    If cboInfo(e_ִ�п���).Text = "" Then
        MsgBox "û��ȷ��ִ�п��ҡ�", vbInformation, Me.Caption
        If cboInfo(e_ִ�п���).Enabled Then cboInfo(e_ִ�п���).SetFocus
        Exit Function
    End If
    mlng����ִ�п���ID = cboInfo(e_ִ�п���).ItemData(cboInfo(e_ִ�п���).ListIndex)
    '����
    If mlng������ĿID <> 0 Then
        If cboInfo(e_����ִ�п���).Text = "" Then
            MsgBox "û��ȷ������ִ�п��ҡ�", vbInformation, Me.Caption
            If cboInfo(e_����ִ�п���).Enabled Then cboInfo(e_����ִ�п���).SetFocus
            Exit Function
        End If
        mlng����ִ�п���id = cboInfo(e_����ִ�п���).ItemData(cboInfo(e_����ִ�п���).ListIndex)
    End If

    '���ʱ��Ϸ���
    If Not Check��ʼʱ��(txtInfo(e_��Чʱ��).Text) Then
        If txtInfo(e_��Чʱ��).Enabled Then txtInfo(e_��Чʱ��).SetFocus
        Exit Function
    End If
    
    If Not Check����ʱ��(txtInfo(e_����ʱ��).Text, txtInfo(e_��Чʱ��).Text) Then
        If txtInfo(e_����ʱ��).Enabled Then txtInfo(e_����ʱ��).SetFocus
        Exit Function
    End If
    If Not Check���� Then
        Exit Function
    End If
    
    '���������������Ȩ������������ҽʦִ��Ȩ
    If gbln������Ȩ���� And mint�������� = 2 Then
        If CheckDocEmpowerEx() = False Then
            If Not gbln�����ȼ����� Then
                MsgBox "����ҽ�����߱���������ִ��Ȩ���������´", vbInformation, "������Ȩ����"
                Exit Function
            Else
                MsgBox "����ҽ�����߱���������ִ��Ȩ��", vbInformation, "������Ȩ����"
            End If
        End If
    End If
    
    If mint���ó��� = 0 Then
        strTmp = mlng������ĿID & "||" & mint��������
        mstrժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", strTmp)
    
        strExtra = "||0||" & mlng������ĿID & "|| ||0"
        lng����ID = IIF(mint���� = 1, mlng�Һ�ID, mlng��ҳID)
        
        lngִ������ = IIF(cboInfo(e_ִ�п���).ItemData(cboInfo(e_ִ�п���).ListIndex) <= 0, 5, mlng����ִ�п�������)
        
        lngִ�п���ID = IIF(cboInfo(e_ִ�п���).ItemData(cboInfo(e_ִ�п���).ListIndex) <= 0, 0, cboInfo(e_ִ�п���).ItemData(cboInfo(e_ִ�п���).ListIndex))
    
        If Not AdviceCheck(lng����ID, "F", mlng������ĿID, lngִ�п���ID, lngִ������, mstrժҪ & strExtra) Then
            Exit Function
        End If
        j = 1
        strTabAdvice = "select " & j & " as ID," & j & " as ���,-null as ���ID,'F' as �������," & mlng������ĿID & " as ������ĿID," & _
            mlng������ĿID & " as ������ĿID,1 As ����, 0 As ����,null as �걾��λ,null As ��鷽��," & _
            "0 as ִ�б��,0 as �Ƽ�����, null As ��������," & lngִ������ & " As ִ������," & lngִ�п���ID & " as ִ�п���id from dual"
    
        '������
        strҽ������ = FormatAdviceContext
        
        strIDs = mlng������ĿID & ":" & lngִ�п���ID
        
        With vsOper
            For i = .FixedRows To .Rows - 1
                If .RowData(i) <> 0 Then
                    strIDs = strIDs & "," & .RowData(i) & ":" & lngִ�п���ID
                    .TextMatrix(i, COL_ժҪ) = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", CStr(.RowData(i)) & "||" & mint��������)
                    If Not AdviceCheck(lng����ID, "F", Val(.RowData(i)), lngִ�п���ID, Val(.TextMatrix(i, COL_ִ������)), .TextMatrix(i, COL_ժҪ) & strExtra) Then
                        Exit Function
                    End If
                    j = j + 1
                    strTabAdvice = strTabAdvice & " Union ALL " & _
                        "select " & j & " as ID," & j & " as ���,1 as ���ID,'F' as �������," & Val(.RowData(i)) & " as ������ĿID," & _
                             Val(.RowData(i)) & " as ������ĿID,1 As ����, 0 As ����,null as �걾��λ,null As ��鷽��," & _
                             Abs(Val(.TextMatrix(i, COL_��Ҫʱ))) & " as ִ�б��,0 as �Ƽ�����, 1 As ��������," & Val(.TextMatrix(i, COL_ִ������)) & " As ִ������," & lngִ�п���ID & " as ִ�п���id from dual"
                End If
            Next
        End With
        
        If mlng������ĿID <> 0 Then
        
            lngִ�п���ID = IIF(cboInfo(e_����ִ�п���).ItemData(cboInfo(e_����ִ�п���).ListIndex) <= 0, 0, cboInfo(e_����ִ�п���).ItemData(cboInfo(e_����ִ�п���).ListIndex))
            
            strIDs = strIDs & "," & mlng������ĿID & ":" & lngִ�п���ID
  
            lblInfo(e_����ִ�п���).Tag = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", CStr(mlng������ĿID) & "||" & mint��������)
            
            If Not AdviceCheck(lng����ID, "G", mlng������ĿID, lngִ�п���ID, lngִ������, lblInfo(e_����ִ�п���).Tag & strExtra) Then
                Exit Function
            End If
            
            j = j + 1
            strTabAdvice = strTabAdvice & " Union ALL " & _
                "select " & j & " as ID," & j & " as ���,1 as ���ID,'F' as �������," & mlng������ĿID & " as ������ĿID," & _
                    mlng������ĿID & " as ������ĿID,1 As ����, 0 As ����,null as �걾��λ,null As ��鷽��," & _
                    "0 as ִ�б��,0 as �Ƽ�����, null As ��������," & lngִ������ & " As ִ������," & lngִ�п���ID & " as ִ�п���id from dual"
        End If
        
        If gintҽ������ = 2 Then mbln���Ѷ��� = True
    
        strMsg = CheckAdviceInsure(mint����, mbln���Ѷ���, mlng����ID, IIF(mlng�������� = 0, 2, 1), "", strIDs, strҽ������)
        
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
        
        'ҽ���ܿ�ʵʱ���
        If mint���� <> 0 Then
            If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) Then
                If MakePriceRecord���뵥("4" & mint��������, mlng����ID, lng����ID, strTabAdvice, strIDs, mstr�ѱ�, mlng��������ID, rsPrice) Then
                    If Not gclsInsure.CheckItem(mint����, IIF(mint�������� = 1, 0, 1), 0, rsPrice) Then
                        MsgBox "ҽ�������δͨ(ִ��Insure.CheckItem�ӿ�)�������´���������뵥���ܱ��档", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
    End If
    CheckData = True
End Function

Private Function SaveCacheData() As Boolean
'���ܣ���������
    Dim strTmp As String
    Dim str��Ҫʱ As String
    
    Dim i As Long
    On Error GoTo errH
    If mrsCard Is Nothing Then
         Call InitCardRsOperate(mrsCard)
         mrsCard.AddNew
    End If
    
    mrsCard!�ٴ�������� = txtInfo(e_��ǰ���).Text
    mrsCard!�ٴ����IDs = mstr���IDs
    mrsCard!������� = cboInfo(e_�������).ListIndex
    mrsCard!��������ĿID = mlng������ĿID
    mrsCard!����ִ�п���ID = mlng����ִ�п���ID
    '��������
    With vsOper
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                strTmp = strTmp & "," & Val(.RowData(i))
                str��Ҫʱ = str��Ҫʱ & "," & Val(.RowData(i)) & ":" & Abs(Val(.TextMatrix(i, COL_��Ҫʱ)))
            End If
        Next
    End With
    mrsCard!��������ĿIDs = Mid(strTmp, 2)
    mrsCard!��������Ҫʱ = Mid(str��Ҫʱ, 2)
    mrsCard!������ĿID = mlng������ĿID
    mrsCard!����ִ�п���ID = mlng����ִ�п���ID
    mrsCard!����ִ�п���ID = mlng����ִ�п���id
    mrsCard!��Чʱ�� = txtInfo(e_��Чʱ��).Text
    mrsCard!����ʱ�� = txtInfo(e_����ʱ��).Text
    mrsCard!�������id = mlng��������ID
    
    strTmp = ""
    mrsAppend.Filter = 0
    mrsAppend.Filter = "��Ŀ='����ҽ��'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>����ҽ��<Split2>" & Val(mrsAppend!���� & "") & "<Split2>" & mrsAppend!Ҫ��ID & "<Split2>" & txtInfo(e_����ҽ��).Text
    Else
        strTmp = strTmp & "<Split1>����ҽ��<Split2>0<Split2>0<Split2>" & txtInfo(e_����ҽ��).Text
    End If
     
    mrsAppend.Filter = "��Ŀ='����ҽ������'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>����ҽ������<Split2>" & Val(mrsAppend!���� & "") & "<Split2>" & mrsAppend!Ҫ��ID & "<Split2>" & txtInfo(e_����ҽ������).Text
    Else
        strTmp = strTmp & "<Split1>����ҽ������<Split2>0<Split2>0<Split2>" & txtInfo(e_����ҽ������).Text
    End If
    
    mrsAppend.Filter = "��Ŀ='��һ����'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>��һ����<Split2>" & Val(mrsAppend!���� & "") & "<Split2>" & mrsAppend!Ҫ��ID & "<Split2>" & txtInfo(e_��һ����).Text
    Else
        strTmp = strTmp & "<Split1>��һ����<Split2>0<Split2>0<Split2>" & txtInfo(e_��һ����).Text
    End If
    
    mrsAppend.Filter = "��Ŀ='�ڶ�����'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>�ڶ�����<Split2>" & Val(mrsAppend!���� & "") & "<Split2>" & mrsAppend!Ҫ��ID & "<Split2>" & txtInfo(e_�ڶ�����).Text
    Else
        strTmp = strTmp & "<Split1>�ڶ�����<Split2>0<Split2>0<Split2>" & txtInfo(e_�ڶ�����).Text
    End If
    
    mrsAppend.Filter = "��Ŀ='��������'"
    If Not mrsAppend.EOF Then
        strTmp = strTmp & "<Split1>��������<Split2>" & Val(mrsAppend!���� & "") & "<Split2>" & mrsAppend!Ҫ��ID & "<Split2>" & txtInfo(e_��������).Text
    Else
        strTmp = strTmp & "<Split1>��������<Split2>0<Split2>0<Split2>" & txtInfo(e_��������).Text
    End If
    
    If lblInfo(e_�ٴ���������).Visible Then
        strTmp = strTmp & "<Split1>�ٴ���������<Split2>0<Split2>0<Split2>" & cboInfo(e_�ٴ���������).Text
    End If
    
    With vsOther
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col_����) <> "" Then
                mrsAppend.Filter = "��Ŀ='" & .TextMatrix(i, COL_��Ŀ����) & "'"
                If Not mrsAppend.EOF Then
                    strTmp = strTmp & "<Split1>" & .TextMatrix(i, COL_��Ŀ����) & "<Split2>" & Val(mrsAppend!���� & "") & "<Split2>" & mrsAppend!Ҫ��ID & "<Split2>" & .TextMatrix(i, col_����)
                End If
            End If
        Next
    End With
    
    If mstr���IDs <> "" Then
        strTmp = strTmp & "<Split1>���뵥���<Split2>0<Split2>0<Split2>" & txtInfo(e_��ǰ���).Text
    End If
    
    mrsCard!���븽�� = Mid(strTmp, 9)
    
    mrsCard.Update
    mblnChange = False
    SaveCacheData = True
    Exit Function
errH:
    If 2 = 1 Then
        Resume
    End If
    err.Clear
End Function

Private Function SaveData() As Boolean
'���ܣ���������
    Dim lngҽ��ID As Long, lng���ID As Long, lngҽ����� As Long, lng������� As Long
    Dim strҽ������ As String, str����ʱ�� As String, strTmp As String
    Dim lngִ�п���ID As Long, arrSQLTmp As Variant
    Dim strSQL As String, arrSQL As Variant, blnTrans As Boolean
    Dim datCur As Date, str��ҳID As String, str�Һŵ� As String
    Dim rsAffer As ADODB.Recordset, lngTmp As Long
    Dim i As Long, int���� As Integer
    Dim strSource As String 'SQLģ��
    Dim rsTmp As ADODB.Recordset
    Dim str���״̬ As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    arrSQL = Array()
    
    If mintType = 1 Then
        strSQL = "Zl_����ҽ����¼_Delete(" & mlngUpdateAdvice & ",1)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        If gbln����Ӱ����ϢϵͳԤԼ Then
            Set rsTmp = Nothing
            Set rsTmp = GetDataRISԤԼ(CStr(mlngUpdateAdvice))
            On Error Resume Next
            If Not rsTmp.EOF Then
                If 0 <> gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!ԤԼid & "")) Then
                    MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ����β���ɾ�����޸����Ѿ�ԤԼҽ����������Ӱ����Ϣϵͳ�ӿ�(HISSchedulingEx)ȡ��ϢԤԼδ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                End If
                rsTmp.MoveNext
            End If
            err.Clear: On Error GoTo errH
        End If
    End If
    
    lngҽ����� = GetMaxAdviceNO(mlng����ID, mlng��ҳID, mbytBaby, mstr�Һŵ�)
    If lng������� = 0 Then lng������� = Get�������
    mlngUpdateAdvice = zlDatabase.GetNextID("����ҽ����¼")
    
    If mint���� = 0 Then
        str��ҳID = mlng��ҳID
        str�Һŵ� = "NULL"
    Else
        str��ҳID = "NULL"
        str�Һŵ� = "'" & mstr�Һŵ� & "'"
    End If
    
    If cboInfo(e_�������).ListIndex = 1 Then int���� = 1
    '���ﳡ������ҽ���������
    str���״̬ = IIF(gbln�����ּ����� And int���� = 0 And mint�������� = 2, "1", "NULL")
    If str���״̬ = "1" Then
        If gbln������Ȩ���� Then
            If gbln�����ȼ����� Then
                If CheckDocEmpowerEx() Then
                    str���״̬ = "Null"
                End If
            End If
        Else
            If gbln�����ȼ����� Then
                If CheckDoc�����ȼ� Then
                    str���״̬ = "Null"
                End If
            End If
        End If
    End If
    datCur = zlDatabase.Currentdate
    str����ʱ�� = IIF(datCur > CDate(txtInfo(e_��Чʱ��).Text), txtInfo(e_��Чʱ��).Text, datCur)
    '��9����Ŀ����Ϊ��[ID],[���ID],[���],[�������],[������ĿID],[ҽ������],[�Ƽ�����],[ִ�п���ID],[ִ������],[ҽ������],[ִ�б��]
    strSource = "ZL_����ҽ����¼_Insert([0],[1],[2]," & mint�������� & "," & mlng����ID & "," & str��ҳID & "," & mbytBaby & ",1,1,'[3]',[4],NULL,NULL,NULL,1,'[5]',[9]," & _
        "'" & txtInfo(e_����ʱ��).Text & "','һ����',NULL,NULL,NULL,NULL,[6],[7],[8]," & IIF(mbln��¼, 2, int����) & "," & _
        "To_Date('" & Format(txtInfo(e_��Чʱ��).Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
        mlng���˿���id & "," & mlng��������ID & ",'" & UserInfo.���� & "'," & _
        "To_Date('" & Format(str����ʱ��, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
        str�Һŵ� & "," & ZVal(mlngǰ��ID) & ",NULL,[10],NULL,[11],'" & UserInfo.���� & "',Null,NULL,NULL," & str���״̬ & _
        "," & lng������� & ",NULL,NULL,NULL," & ZVal(cboInfo(e_�������).ListIndex) & ")"
    
    
    '��������Ŀ������ҽ����ҽ������ֻ������ҽ������
    lngҽ����� = lngҽ����� + 1    '����ҽ����¼.��ţ�����
    strҽ������ = FormatAdviceContext
    With cboInfo(e_ִ�п���)
        lngִ�п���ID = IIF(.ItemData(.ListIndex) <= 0, 0, .ItemData(.ListIndex))
    End With
    strTmp = "Null"
    strSQL = GetStrExcSQL(strSource, mlngUpdateAdvice, "NULL", lngҽ�����, "F", mlng������ĿID, strҽ������, Val(lblInfo(e_��������).Tag), ZVal(lngִ�п���ID), IIF(lngִ�п���ID <= 0, "5", mlng����ִ�п�������), "Null", 0, IIF(mstrժҪ = "", "null", "'" & mstrժҪ & "'"))
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    
    lng���ID = mlngUpdateAdvice
    
    '��������
    With vsOper
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                strҽ������ = .TextMatrix(i, COL_����)
                strҽ������ = Mid(strҽ������, InStr(strҽ������, "]") + 1)
                strTmp = IIF(.TextMatrix(i, COL_ժҪ) = "", "null", "'" & .TextMatrix(i, COL_ժҪ) & "'")
                lngҽ��ID = zlDatabase.GetNextID("����ҽ����¼")
                lngҽ����� = lngҽ����� + 1
                strSQL = GetStrExcSQL(strSource, lngҽ��ID, lng���ID, lngҽ�����, "F", Val(.RowData(i)), strҽ������, Val(.TextMatrix(i, COL_�Ƽ�����)), ZVal(lngִ�п���ID), Val(.TextMatrix(i, COL_ִ������)), "NULL", Abs(Val(.TextMatrix(i, COL_��Ҫʱ))), strTmp)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        Next
    End With
    
    '������Ŀ
    If mlng������ĿID <> 0 Then
        lngִ�п���ID = 0
        With cboInfo(e_����ִ�п���)
            lngִ�п���ID = IIF(.ItemData(.ListIndex) <= 0, 0, .ItemData(.ListIndex))
        End With
        
        strҽ������ = txtInfo(e_������).Text
        strҽ������ = Mid(strҽ������, InStr(strҽ������, "]") + 1)
        strTmp = IIF(lblInfo(e_����ִ�п���).Tag = "", "null", "'" & lblInfo(e_����ִ�п���).Tag & "'")
        lngҽ��ID = zlDatabase.GetNextID("����ҽ����¼")
        
        lngҽ����� = lngҽ����� + 1
        strSQL = GetStrExcSQL(strSource, lngҽ��ID, lng���ID, lngҽ�����, "G", mlng������ĿID, strҽ������, Val(lblInfo(e_������).Tag), ZVal(lngִ�п���ID), IIF(lngִ�п���ID <= 0, "5", mlng����ִ�п�������), "NULL", 0, strTmp)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
    
    '���븽��
    lngTmp = GetStrExcSQL���븽��(lng���ID, arrSQLTmp)
    For i = 0 To UBound(arrSQLTmp)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = CStr(arrSQLTmp(i))
    Next
    
    '��Ϲ�����Ϣ
    If mstr���IDs <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Insert(" & lng���ID & ",'" & mstr���IDs & "')"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lng���ID & ",'���뵥���',null," & lngTmp + 1 & ",null,'" & txtInfo(e_��ǰ���).Text & "',0)"
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If mint���� = 0 Then
        If gbln�����ּ����� And int���� = 0 Then
            Call ZLHIS_CIS_Audit("ZLHIS_CIS_028", mclsMipModule, mlng����ID, txtInfo(e_����).Text, txtInfo(e_סԺ��).Text, , IIF(mlng�������� = 1, 1, 2), _
                mlng��ҳID, mlng����ID, , mlng���˿���id, "", , txtInfo(e_����).Text, _
                lngҽ��ID, UserInfo.����, Format(str����ʱ��, "yyyy-MM-dd HH:mm:ss"), mlng��������ID, "", , "")
        Else
            Call ZLHIS_CIS_001(mclsMipModule, mlng����ID, txtInfo(e_����).Text, txtInfo(e_סԺ��).Text, , IIF(mlng�������� = 1, 1, 2), _
                mlng��ҳID, mlng����ID, , mlng���˿���id, "", , txtInfo(e_����).Text, _
                lngҽ��ID, IIF(int���� = 1, 1, 0), 1, "F", "", UserInfo.����, Format(str����ʱ��, "yyyy-MM-dd HH:mm:ss"), mlng��������ID, "", , , "")
        End If
    End If
    
    If mstr�������� = "" Then
        strSQL = "select ���� from ���ű� where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��������ID)
        mstr�������� = rsTmp!���� & ""
        txtInfo(e_�������).Text = mstr��������
        txtInfo(e_����ҽʦ).Text = UserInfo.����
    End If
    
    If gbln����Ӱ����ϢϵͳԤԼ Then
        On Error Resume Next
        Call gobjRis.HISScheduling(IIF(1 = mint����, 1, 2), lng���ID, mlng������ĿID)
        err.Clear
    End If
    
    mintType = 1
    SaveData = True
    mblnChange = False
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

Private Sub ReSetAdviceNo(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
'���ܣ���������ҽ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Count(*) as Num From (Select ���,Count(ID) From ����ҽ����¼ Where ����ID=[1] And ��ҳID=[2] Having Count(ID)>1 Group by ���)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    If rsTmp.EOF Then Exit Sub
    
    If NVL(rsTmp!Num, 0) = 0 Then Exit Sub
    
    strSQL = "ZL_����ҽ����¼_�������(NULL,NULL," & mlng����ID & "," & mlng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
 
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetAdviceNO(ByRef lngMinNo As Long, ByRef lngMaxNo As Long, ByRef lngNo As Long)
'���ܣ���ȡ��ǰ��ҽ����������С���,lngNo-�������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSQL = "select min(���) as ��С,max(���) as ���,max(�������) as ������� from ����ҽ����¼ where id=[1] or ���id=[1]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUpdateAdvice)
    
    lngMinNo = rsTmp!��С
    lngMaxNo = rsTmp!���
    lngNo = Val(rsTmp!������� & "")
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetAfferAdvice() As ADODB.Recordset
'���ܣ���ȡ��ǰ����ҽ��֮���ҽ��
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSQL = "select id as ҽ��id from ����ҽ����¼ where ����id=[1] and nvl(Ӥ��,0)=[2]" & _
        IIF(mint���� = 0, " and ��ҳid=[3]", " and �Һŵ�=[4]") & " order by ���"
    
    On Error GoTo errH
    
    Set GetAfferAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, 0, mlng��ҳID, mstr�Һŵ�)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function


Private Function GetStrExcSQL���븽��(ByVal lngҽ��ID As Long, ByRef arrSQL As Variant) As Long
'���ܣ��������븽��Ŀ�ִ��SQL������varSQL�У����û���������� false�� strItems ��ҽ���༭����ʱ
'���أ�true �У�false ��
    Dim intCount As Integer, i As Integer
    Dim lng����Ҫ��ID As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    arrSQL = Array()
    
    strTmp = txtInfo(e_����ҽ��).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    mrsAppend.Filter = "���=" & Val(lblInfo(e_����ҽ��).Tag)
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'" & mrsAppend!��Ŀ & "'," & mrsAppend!���� & "," & intCount & "," & mrsAppend!Ҫ��ID & "," & strTmp & ",1)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'����ҽ��',0," & intCount & "," & Val(lblInfo(e_����ҽ��).Tag) & "," & strTmp & ",1)"
    End If
    
    strTmp = txtInfo(e_����ҽ������).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    mrsAppend.Filter = "���=" & Val(lblInfo(e_����ҽ������).Tag)
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'" & mrsAppend!��Ŀ & "'," & mrsAppend!���� & "," & intCount & "," & mrsAppend!Ҫ��ID & "," & strTmp & ",0)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'����ҽ������',0," & intCount & "," & Val(lblInfo(e_����ҽ������).Tag) & "," & strTmp & ",0)"
    End If
 
    
    '��ȡ������Ҫ��ID-------------------
    mrsAppend.Filter = "���=" & Val(lblInfo(e_��һ����).Tag)
    If Not mrsAppend.EOF Then
        lng����Ҫ��ID = Val(mrsAppend!Ҫ��ID & "")
    Else
        lng����Ҫ��ID = Val(lblInfo(e_��һ����).Tag)
    End If
    
    strTmp = txtInfo(e_��һ����).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'" & mrsAppend!��Ŀ & "'," & mrsAppend!���� & "," & intCount & "," & lng����Ҫ��ID & "," & strTmp & ",0)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'��һ����',0," & intCount & "," & lng����Ҫ��ID & "," & strTmp & ",0)"
    End If
 
    '------------------------------
    strTmp = txtInfo(e_�ڶ�����).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    mrsAppend.Filter = "���=" & Val(lblInfo(e_�ڶ�����).Tag)
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'" & mrsAppend!��Ŀ & "'," & mrsAppend!���� & "," & intCount & "," & lng����Ҫ��ID & "," & strTmp & ",0)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'�ڶ�����',0," & intCount & "," & lng����Ҫ��ID & "," & strTmp & ",0)"
    End If
    
    strTmp = txtInfo(e_��������).Text
    strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
    mrsAppend.Filter = "���=" & Val(lblInfo(e_��������).Tag)
    intCount = intCount + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    If Not mrsAppend.EOF Then
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'" & mrsAppend!��Ŀ & "'," & mrsAppend!���� & "," & intCount & "," & lng����Ҫ��ID & "," & strTmp & ",0)"
    Else
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'��������',0," & intCount & "," & lng����Ҫ��ID & "," & strTmp & ",0)"
    End If
    
    If lblInfo(e_�ٴ���������).Visible Then
        intCount = intCount + 1
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'�ٴ���������',0," & intCount & ",null,'" & cboInfo(e_�ٴ���������).Text & "',0)"
    End If
    
    '����еĸ���
    With vsOther
        If .Visible Then
            For i = .FixedRows To .Rows - 1
                strTmp = .TextMatrix(i, col_����)
                strTmp = IIF(strTmp = "", "NULL", "'" & strTmp & "'")
                mrsAppend.Filter = "���=" & Val(.RowData(i))
                intCount = intCount + 1
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngҽ��ID & ",'" & mrsAppend!��Ŀ & "'," & mrsAppend!���� & "," & intCount & "," & ZVal(mrsAppend!Ҫ��ID & "") & "," & strTmp & ",0)"
            Next
        End If
    End With
    
    GetStrExcSQL���븽�� = intCount
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetStrExcSQL(ByVal strSource As String, ParamArray arrInput() As Variant) As String
'���ܣ����� ZL_����ҽ����¼_Insert������䣬arrInput���� [ID],[���ID],[���],[�������],[������ĿID],[ҽ������],[�Ƽ�����],[ִ�п���ID],[ִ������]
    Dim i As Integer
    Dim strTmp As String
    Dim strResult As String
    
    strResult = strSource
    For i = 0 To UBound(arrInput)
        strTmp = arrInput(i)
        strResult = Replace(strResult, "[" & i & "]", strTmp)
    Next
    GetStrExcSQL = strResult
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

Private Function GetMax�����ȼ�(ByVal str������Ŀ As String) As String
'���ܣ�ȡ�õ�ǰҽ�������������
'������str������Ŀ��������ĿID�ã��ָ���lng�����ȼ�������ߵ������ȼ�
    Dim strSQL As String, rsTmp As Recordset
    Dim str�����ȼ� As String, i As Integer
    
    On Error GoTo errH
    strSQL = "Select a.�������� From ��������Ŀ¼ A,������϶��� B Where a.ID=b.����ID And a.���='S' And instr([1], b.����id)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str������Ŀ)
    For i = 1 To rsTmp.RecordCount
        If Decode(rsTmp!�������� & "", "һ��", 1, "����", 2, "����", 3, "�ļ�", 4, 0) > Decode(str�����ȼ�, "һ��", 1, "����", 2, "����", 3, "�ļ�", 4, 0) Then
            str�����ȼ� = rsTmp!�������� & ""
        End If
        rsTmp.MoveNext
    Next
    GetMax�����ȼ� = str�����ȼ�
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get������Ŀ��¼ID(ByVal lngID As Long, Optional ByVal strIDs As String) As ADODB.Recordset
'���ܣ���ȡָ��������Ŀ��¼��ID
'������
    Dim strSQL As String
    
    strSQL = "Select " & vbNewLine & _
        " a.�������, a.վ��, a.���, a.����id, a.Id, a.����, a.����, a.�걾��λ, a.���㵥λ, a.���㷽ʽ, a.ִ��Ƶ��, a.�����Ա�, a.����Ӧ��, a.�����Ŀ, c.�������� as ��������, a.ִ�а���," & vbNewLine & _
        " a.ִ�п���, a.�������, a.�Ƽ�����, a.�ο�Ŀ¼id, a.��Աid, a.����ʱ��, a.����ʱ��, a.¼������, a.�Թܱ���, a.ִ�з���, a.ִ�б��" & vbNewLine & _
        "From ������ĿĿ¼ A, ������϶��� B, ��������Ŀ¼ C" & vbNewLine & _
        "Where a.Id = b.����id(+) And b.����id = c.Id(+) And a.ID"
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = strSQL & " IN (Select /*+cardinality(E,10)*/ Column_Value From Table(f_Num2list([1])) E)"
        Set Get������Ŀ��¼ID = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strIDs)
    Else
        strSQL = strSQL & " = [1]"
        Set Get������Ŀ��¼ID = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FormatAdviceContext() As String
'���ܣ�����ϵͳ������������ʽ��ҽ������
'������strBloodWay=��Ѫ;��,strAdvicePro=��Ѫ����
    Dim strReturn As String, strText As String, strField As String
    Dim str���� As String, str���� As String
    Dim i As Integer, strTmp As String
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
 
    strReturn = mstrDefine
    
    strTmp = txtInfo(e_������).Text
    strTmp = Mid(strTmp, InStr(strTmp, "]") + 1)
    str���� = strTmp
    
    With vsOper
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
            
                strTmp = .TextMatrix(i, COL_����)
                strTmp = Mid(strTmp, InStr(strTmp, "]") + 1)
                strTmp = strTmp & IIF(Abs(Val(.TextMatrix(i, COL_��Ҫʱ))) = 1, "(��Ҫʱ)", "")
                str���� = str���� & "," & strTmp
            End If
        Next
        str���� = Mid(str����, 2)
    End With
    
    If strReturn = "" Then
        strText = Format(txtInfo(e_����ʱ��).Text, "MM��dd��HH:mm")
        If str���� <> "" Then
            strText = strText & IIF(str���� <> "", " �� " & str���� & " ���� ", " �� ")
        End If
        strText = strText & txtInfo(e_��������).Text
        If str���� <> "" Then
            strText = strText & " �� " & str����
        End If
        strReturn = strText
    Else
        strText = strReturn
        If InStr(strText, "[����ʱ��]") > 0 Then
            strField = txtInfo(e_����ʱ��).Text
            strText = Replace(strText, "[����ʱ��]", """" & strField & """")
        End If
        
        If InStr(strText, "[��Ҫ����]") > 0 Then
            strField = txtInfo(e_��������).Text
            strText = Replace(strText, "[��Ҫ����]", """" & strField & """")
        End If
        
        If InStr(strText, "[��������]") > 0 Then
            strField = str����
            strText = Replace(strText, "[��������]", """" & strField & """")
        End If
        
        If InStr(strText, "[������]") > 0 Then
            strField = str����
            strText = Replace(strText, "[������]", """" & strField & """")
        End If
        
        strReturn = mobjVBA.Eval(strText)
    End If

    FormatAdviceContext = strReturn
End Function

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
        MsgBox "�����ҽ����ʼִ��ʱ����Ч��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mint���� = 0 Then
        strInDate = mstr��Ժʱ��
        If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
            strMsg = "ҽ������Чʱ�䲻��С�ڲ��˵���Ժʱ�� " & strInDate & " ��"
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
                    strMsg = "ҽ������Чʱ��ӦС�ڲ���" & IIF(mintPState = ps���ת��, "ת��", IIF(mintPState = psԤ��, "Ԥ��Ժ", "��Ժ")) & "��ʱ�� " & strInDate & " ��"
                    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If Format(strStart, "yyyy-MM-dd HH:mm") < strInDate Then
                    strMsg = "ҽ������Чʱ�䲻��С�ڲ��������ת��ʱ�� " & strInDate & " ��"
                    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    Check��ʼʱ�� = True
End Function

Private Function Check����ʱ��(ByVal strDate As String, ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ�����������������ʱ���Ƿ�Ϸ�  �������ϣ�������ʱ��
'˵����
'1.��Ѫʱ�䲻��С��ҽ���Ŀ�ʼʱ��
    Dim strInDate As String, strDateType As String
    
    If Not IsDate(strDate) Then
        strMsg = "���������ʱ����Ч��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf IsDate(strStart) Then
        If Format(strDate, "yyyy-MM-dd HH:mm") < Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = "����ʱ�䲻��С��ҽ����Чʱ�䡣"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check����ʱ�� = True
End Function

Private Sub txtInfo_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtInfo(Index)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
'�����¼�����ģ����
    If KeyAscii = 13 Then
        Select Case Index
        Case e_��������
            Call GetItemOper(0)
        Case e_������
            Call GetItem����(0)
        Case e_����ҽ��
            Call GetItem����ҽ��(0)
        Case e_����ҽ������
            Call GetItem����ҽ������(0)
        Case e_��һ����, e_�ڶ�����, e_��������
            Call GetItemDoctor(0, Index)
        End Select
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Select Case Index
        Case e_��������
            Call GetItemOper(1)
        Case e_������
            Call GetItem����(1)
        Case e_����ҽ��
            Call GetItem����ҽ��(1)
        Case e_����ҽ������
            Call GetItem����ҽ������(1)
        Case e_��һ����, e_�ڶ�����, e_��������
            Call GetItemDoctor(1, Index)
        End Select

    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
'�Ϸ��Լ���ֵ�Ļָ�
    If mintType = 2 Then Exit Sub
    
    Select Case Index
    Case e_����ʱ��
        If Not IsDate(txtInfo(Index).Text) Then
            If txtInfo(Index).Text <> "" Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            Else
                If IsDate(txtInfo(e_��Чʱ��).Text) Then
                    '�ָ���Ϊ�����ȱʡΪ��ʼʱ��
                    txtInfo(Index).Text = txtInfo(e_��Чʱ��).Text
                End If
            End If
        Else
            '���ʱ��Ϸ���
            If Not Check����ʱ��(txtInfo(Index).Text, txtInfo(e_��Чʱ��).Text) Then
                Cancel = True
                Call txtInfo_GotFocus(Index)
                Exit Sub
            End If
        End If
    Case e_��Чʱ��
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
            If mint���� = 0 Then
                '���ʱ��Ϸ���
                If Not Check��ʼʱ��(txtInfo(Index).Text) Then
                    Cancel = True
                    Call txtInfo_GotFocus(Index)
                    Exit Sub
                End If
                '�ж��Ƿ��ǲ�¼ҽ��
                If DateDiff("n", CDate(txtInfo(Index).Text), CDate(zlDatabase.Currentdate)) > gint��¼��� _
                    Or mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
                    mbln��¼ = True
                Else
                    mbln��¼ = False
                End If
            Else
                mbln��¼ = False
            End If
        End If
    
    Case e_��������, e_������, e_����ҽ��, e_����ҽ������, e_��һ����, e_�ڶ�����, e_��������
        If txtInfo(Index).Tag <> txtInfo(Index).Text And txtInfo(Index).Tag <> "" Then txtInfo(Index).Text = txtInfo(Index).Tag
         
    End Select
End Sub

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

Private Sub SetItemEditable(Optional int����ʱ�� As Integer, Optional int��ǰ��� As Integer, _
    Optional int�������� As Integer, Optional intִ�п��� As Integer, Optional int������� As Integer, _
    Optional int������ As Integer, Optional int����ִ�п��� As Integer, _
    Optional int����ҽ�� As Integer, Optional int����ҽ������ As Integer, _
    Optional int��һ���� As Integer, Optional int�ڶ����� As Integer, Optional int�������� As Integer, _
    Optional int��Чʱ�� As Integer, Optional int�������� As Integer, Optional int�������� As Integer, Optional int�ٴ��������� As Integer)
'���ܣ�����ָ���༭��Ŀ���״̬
'������0-���ֲ���,-1-����,1-����
        
    If int����ʱ�� = 1 Then
        txtInfo(e_����ʱ��).Locked = False
        txtInfo(e_����ʱ��).TabStop = True
        txtInfo(e_����ʱ��).BackColor = vbWindowBackground
        cmdDate(e_����ʱ��).Enabled = True
    ElseIf int����ʱ�� = -1 Then
        txtInfo(e_����ʱ��).Locked = True
        txtInfo(e_����ʱ��).TabStop = False
        txtInfo(e_����ʱ��).BackColor = vbButtonFace
        cmdDate(e_����ʱ��).Enabled = False
    End If
    
    If int��ǰ��� = 1 Then
        txtInfo(e_��ǰ���).Locked = False
        txtInfo(e_��ǰ���).TabStop = True
        txtInfo(e_��ǰ���).BackColor = vbWindowBackground
        cmdInfo(e_��ǰ���).Enabled = True
    ElseIf int��ǰ��� = -1 Then
        txtInfo(e_��ǰ���).Locked = True
        txtInfo(e_��ǰ���).TabStop = False
        txtInfo(e_��ǰ���).BackColor = vbButtonFace
        cmdInfo(e_��ǰ���).Enabled = False
    End If
    
    If int�������� = 1 Then
        txtInfo(e_��������).Locked = False
        txtInfo(e_��������).TabStop = True
        txtInfo(e_��������).BackColor = vbWindowBackground
        cmdInfo(e_��������).Enabled = True
    ElseIf int�������� = -1 Then
        txtInfo(e_��������).Locked = True
        txtInfo(e_��������).TabStop = False
        txtInfo(e_��������).BackColor = vbButtonFace
        cmdInfo(e_��������).Enabled = False
    End If
    
    If intִ�п��� = 1 Then
        cboInfo(e_ִ�п���).Locked = False
        cboInfo(e_ִ�п���).TabStop = True
        cboInfo(e_ִ�п���).BackColor = vbWindowBackground
    ElseIf intִ�п��� = -1 Then
        cboInfo(e_ִ�п���).Locked = True
        cboInfo(e_ִ�п���).TabStop = False
        cboInfo(e_ִ�п���).BackColor = vbButtonFace
    End If
    
    If int������� = 1 Then
        cboInfo(e_�������).Locked = False
        cboInfo(e_�������).TabStop = True
        cboInfo(e_�������).BackColor = vbWindowBackground
    ElseIf int������� = -1 Then
        cboInfo(e_�������).Locked = True
        cboInfo(e_�������).TabStop = False
        cboInfo(e_�������).BackColor = vbButtonFace
    End If
    
    If int������ = 1 Then
        txtInfo(e_������).Locked = False
        txtInfo(e_������).TabStop = True
        txtInfo(e_������).BackColor = vbWindowBackground
        cmdInfo(e_������).Enabled = True
    ElseIf int������ = -1 Then
        txtInfo(e_������).Locked = True
        txtInfo(e_������).TabStop = False
        txtInfo(e_������).BackColor = vbButtonFace
        cmdInfo(e_������).Enabled = False
    End If
    
    If int����ִ�п��� = 1 Then
        cboInfo(e_����ִ�п���).Locked = False
        cboInfo(e_����ִ�п���).TabStop = True
        cboInfo(e_����ִ�п���).BackColor = vbWindowBackground
    ElseIf int����ִ�п��� = -1 Then
        cboInfo(e_����ִ�п���).Locked = True
        cboInfo(e_����ִ�п���).TabStop = False
        cboInfo(e_����ִ�п���).BackColor = vbButtonFace
    End If
    
    If int����ҽ�� = 1 Then
        txtInfo(e_����ҽ��).Locked = False
        txtInfo(e_����ҽ��).TabStop = True
        txtInfo(e_����ҽ��).BackColor = vbWindowBackground
        cmdInfo(e_����ҽ��).Enabled = True
    ElseIf int����ҽ�� = -1 Then
        txtInfo(e_����ҽ��).Locked = True
        txtInfo(e_����ҽ��).TabStop = False
        txtInfo(e_����ҽ��).BackColor = vbButtonFace
        cmdInfo(e_����ҽ��).Enabled = False
    End If
    
    If int����ҽ������ = 1 Then
        txtInfo(e_����ҽ������).Locked = False
        txtInfo(e_����ҽ������).TabStop = True
        txtInfo(e_����ҽ������).BackColor = vbWindowBackground
        cmdInfo(e_����ҽ������).Enabled = True
    ElseIf int����ҽ������ = -1 Then
        txtInfo(e_����ҽ������).Locked = True
        txtInfo(e_����ҽ������).TabStop = False
        txtInfo(e_����ҽ������).BackColor = vbButtonFace
        cmdInfo(e_����ҽ������).Enabled = False
    End If
    
    If int��һ���� = 1 Then
        txtInfo(e_��һ����).Locked = False
        txtInfo(e_��һ����).TabStop = True
        txtInfo(e_��һ����).BackColor = vbWindowBackground
        cmdInfo(e_��һ����).Enabled = True
    ElseIf int��һ���� = -1 Then
        txtInfo(e_��һ����).Locked = True
        txtInfo(e_��һ����).TabStop = False
        txtInfo(e_��һ����).BackColor = vbButtonFace
        cmdInfo(e_��һ����).Enabled = False
    End If
    
    If int�ڶ����� = 1 Then
        txtInfo(e_�ڶ�����).Locked = False
        txtInfo(e_�ڶ�����).TabStop = True
        txtInfo(e_�ڶ�����).BackColor = vbWindowBackground
        cmdInfo(e_�ڶ�����).Enabled = True
    ElseIf int�ڶ����� = -1 Then
        txtInfo(e_�ڶ�����).Locked = True
        txtInfo(e_�ڶ�����).TabStop = False
        txtInfo(e_�ڶ�����).BackColor = vbButtonFace
        cmdInfo(e_�ڶ�����).Enabled = False
    End If
    
    If int�������� = 1 Then
        txtInfo(e_��������).Locked = False
        txtInfo(e_��������).TabStop = True
        txtInfo(e_��������).BackColor = vbWindowBackground
        cmdInfo(e_��������).Enabled = True
    ElseIf int�������� = -1 Then
        txtInfo(e_��������).Locked = True
        txtInfo(e_��������).TabStop = False
        txtInfo(e_��������).BackColor = vbButtonFace
        cmdInfo(e_��������).Enabled = False
    End If
    
    If int��Чʱ�� = 1 Then
        txtInfo(e_��Чʱ��).Locked = False
        txtInfo(e_��Чʱ��).TabStop = True
        txtInfo(e_��Чʱ��).BackColor = vbWindowBackground
        cmdDate(e_��Чʱ��).Enabled = True
    ElseIf int��Чʱ�� = -1 Then
        txtInfo(e_��Чʱ��).Locked = True
        txtInfo(e_��Чʱ��).TabStop = False
        txtInfo(e_��Чʱ��).BackColor = vbButtonFace
        cmdDate(e_��Чʱ��).Enabled = False
    End If
    
 
    If int�������� = 1 Then
        vsOther.TabStop = True
    ElseIf int�������� = -1 Then
        vsOther.TabStop = False
        vsOper.Editable = flexEDNone
    End If
    
    If int�������� = 1 Then
        vsOper.TabStop = True
    ElseIf int�������� = -1 Then
        vsOper.TabStop = False
        vsOper.Editable = flexEDNone
    End If
    
    
    If int�ٴ��������� = 1 Then
        cboInfo(e_�ٴ���������).Locked = False
        cboInfo(e_�ٴ���������).TabStop = True
        cboInfo(e_�ٴ���������).BackColor = vbWindowBackground
    ElseIf int�ٴ��������� = -1 Then
        cboInfo(e_�ٴ���������).Locked = True
        cboInfo(e_�ٴ���������).TabStop = False
        cboInfo(e_�ٴ���������).BackColor = vbButtonFace
    End If
End Sub

Private Function CheckDoc�����ȼ�() As Boolean
'���ܣ�������Ա������ҽ���Ƿ�ﵽ������Ŀ�ĵȼ�
'˵������Ա�ȼ��������ȼ� ͬʱ��ʱֵ�Ž����жϣ�������Ҫ���
    Dim strSQL As String, rsTmp As Recordset
    Dim strDoc As String
    Dim str�����ȼ� As String
    Dim str��Ա�ȼ� As String
    Dim blnTmp As Boolean
    
    On Error GoTo errH
    
    If txtInfo(e_����ҽ��).Text = "" Then
        str��Ա�ȼ� = UserInfo.�����ȼ�
    Else
        strDoc = txtInfo(e_����ҽ��).Text
        strSQL = "Select b.�����ȼ� From ��Ա�� B Where B.����=[1] and b.�����ȼ� is not null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDoc�����ȼ�", strDoc)
        If Not rsTmp.EOF Then str��Ա�ȼ� = rsTmp!�����ȼ� & ""
    End If
    
    If str��Ա�ȼ� <> "" Then
        strSQL = "select i.�������� as �����ȼ� from ��������Ŀ¼ I, ������϶��� R where r.����id=[1] and i.id=r.����id and i.�������� is not null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDoc�����ȼ�", mlng������ĿID)
        If Not rsTmp.EOF Then str�����ȼ� = rsTmp!�����ȼ� & ""
        
        If str��Ա�ȼ� <> "" And str�����ȼ� <> "" Then
            If Decode(str��Ա�ȼ�, "��", 1, "��", 2, "��", 3, "��", 4, "һ��", 1, "����", 2, "����", 3, "�ļ�", 4, 0) >= _
                Decode(str�����ȼ�, "��", 1, "��", 2, "��", 3, "��", 4, "һ��", 1, "����", 2, "����", 3, "�ļ�", 4, 0) Then
                blnTmp = True
            End If
        End If
    End If
    
    CheckDoc�����ȼ� = blnTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckDocEmpowerEx() As Boolean
'���ܣ�������Ա�Ƿ����������Ŀ��ִ��Ȩ
'������strAppend=��ǰ���븽�����д�����,��ʽΪ"��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."
    Dim strSQL As String, rsTmp As Recordset
    Dim strDoc As String
    
    On Error GoTo errH
 
    If txtInfo(e_����ҽ��).Text = "" Then
        strDoc = UserInfo.����
    Else
        strDoc = txtInfo(e_����ҽ��).Text
    End If
    strSQL = "Select Count(*) as Ȩ�� From ��Ա����Ȩ�� A,��Ա�� B Where A.��Աid = B.ID And B.����=[1] And A.������Ŀid = [2] And A.��¼���� = 2"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", strDoc, mlng������ĿID)
    CheckDocEmpowerEx = Val(rsTmp!Ȩ�� & "") > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceCheck(ByVal lng����ID As String, ByVal str������� As String, ByVal lng��Ŀid As Long, ByVal lngִ�п���ID As Long, ByVal lngִ������ As Long, ByVal strExtra As String) As Boolean
'���ܣ��������ݿⷽ��zl_AdviceCheck��ҽ����Ŀ���м��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String
    
    strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", mint��������, mlng����ID, lng����ID, mint����, 1, _
         str�������, lng��Ŀid, mlng��������ID, UserInfo.����, lngִ�п���ID, lngִ������, 0, 0, strExtra)
    
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
    AdviceCheck = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsOperateAgain() As Boolean
'���ܣ��ǲ����ٴ����������������Ч������ҽ������Ϊ�ٴ�����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    If mstr�Һŵ� = "" Then
        strSQL = "Select 1 From ����ҽ����¼ Where ����ID=[1]   And ��ҳid=[2] And Nvl(Ӥ��,0)=[3] and ҽ��״̬=8 and �������='F' and rownum<2"
    Else
        strSQL = "Select 1 From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[4] And Nvl(Ӥ��,0)=[3] and ҽ��״̬=8 and �������='F' and rownum<2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mbytBaby, mstr�Һŵ�)
    IsOperateAgain = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
