VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UsrPriceQuery 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10440
   KeyPreview      =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   10440
   Begin VB.PictureBox picKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   5055
      MouseIcon       =   "UsrPriceQuery.ctx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   3075
      ScaleWidth      =   5310
      TabIndex        =   25
      Top             =   645
      Width           =   5310
      Begin zl9NewQuery.ctlKeyBoard cmdKey 
         Height          =   3090
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5450
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   2310
      ScaleHeight     =   285
      ScaleWidth      =   2580
      TabIndex        =   44
      Top             =   6240
      Width           =   2580
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϼ�:10000000.00Ԫ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   45
         Top             =   30
         Width           =   2385
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   75
      MouseIcon       =   "UsrPriceQuery.ctx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   2460
      ScaleWidth      =   1845
      TabIndex        =   15
      Top             =   4875
      Width           =   1845
      Begin VB.CommandButton cmdBtn 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   0
         Left            =   30
         TabIndex        =   31
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   1
         Left            =   630
         TabIndex        =   30
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   2
         Left            =   1230
         TabIndex        =   29
         Top             =   45
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   12
         Left            =   1230
         TabIndex        =   24
         Top             =   1845
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   11
         Left            =   630
         TabIndex        =   23
         Top             =   1845
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   9
         Left            =   30
         TabIndex        =   22
         Top             =   1845
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   8
         Left            =   1230
         TabIndex        =   21
         Top             =   1245
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   7
         Left            =   630
         TabIndex        =   20
         Top             =   1245
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   6
         Left            =   30
         TabIndex        =   19
         Top             =   1245
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   5
         Left            =   1230
         TabIndex        =   18
         Top             =   645
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   4
         Left            =   630
         TabIndex        =   17
         Top             =   645
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   3
         Left            =   30
         TabIndex        =   16
         Top             =   645
         Width           =   600
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "ȷ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         Left            =   900
         TabIndex        =   28
         Top             =   1455
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   0
      Left            =   15
      ScaleHeight     =   930
      ScaleWidth      =   10425
      TabIndex        =   0
      Top             =   3915
      Width           =   10425
      Begin VB.Frame fra2 
         Caption         =   "Frame1"
         Height          =   960
         Left            =   4800
         TabIndex        =   41
         Top             =   -60
         Width           =   30
      End
      Begin VB.Frame fra 
         Height          =   30
         Left            =   5430
         TabIndex        =   33
         Top             =   435
         Width           =   4920
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   7065
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   555
         Width           =   1500
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   1
         Left            =   8640
         TabIndex        =   26
         Top             =   480
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   741
         Caption         =   "����������ѯ"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   14
         Left            =   3780
         TabIndex        =   42
         Top             =   30
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   741
         Caption         =   "�Ϸ�"
         BackColor       =   16777215
         ForeColor       =   12583104
         FontSize        =   10.5
         AutoSize        =   0   'False
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   15
         Left            =   3780
         TabIndex        =   43
         Top             =   480
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   741
         Caption         =   "�·�"
         BackColor       =   16777215
         ForeColor       =   12583104
         FontSize        =   10.5
         AutoSize        =   0   'False
         TextAligment    =   0
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ˵��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   4
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   3225
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�۸�����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   3
         Left            =   6615
         TabIndex        =   32
         Top             =   90
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ѯ����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   2
         Left            =   5415
         TabIndex        =   2
         Top             =   615
         Width           =   1575
      End
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   2640
   End
   Begin VB.Timer tmrInfo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   2685
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   75
      ScaleHeight     =   480
      ScaleWidth      =   7920
      TabIndex        =   12
      Top             =   7350
      Width           =   7920
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   495
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   37
         Top             =   90
         Width           =   645
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   12
         Left            =   5670
         TabIndex        =   13
         Top             =   15
         Width           =   540
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "�Ϸ�"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   13
         Left            =   6225
         TabIndex        =   14
         Top             =   15
         Width           =   540
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "�·�"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   0
         Left            =   1845
         TabIndex        =   38
         Top             =   30
         Width           =   540
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "ɾ��"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   3
         Left            =   1215
         TabIndex        =   39
         Top             =   30
         Width           =   540
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "���"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   0
         Left            =   45
         TabIndex        =   40
         Top             =   150
         Width           =   420
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   9540
      TabIndex        =   3
      Top             =   0
      Width           =   9540
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   4
         Left            =   45
         TabIndex        =   4
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "ҩ��"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   5
         Left            =   930
         TabIndex        =   5
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "����"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   6
         Left            =   1815
         TabIndex        =   6
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "���"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   7
         Left            =   2700
         TabIndex        =   7
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "����"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   8
         Left            =   3585
         TabIndex        =   8
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "����"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   9
         Left            =   4485
         TabIndex        =   9
         Top             =   15
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   741
         Caption         =   "��������"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   10
         Left            =   6660
         TabIndex        =   10
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "�Ϸ�"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   11
         Left            =   7560
         TabIndex        =   11
         Top             =   15
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
         Caption         =   "�·�"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   420
         Index           =   2
         Left            =   8625
         TabIndex        =   27
         Top             =   15
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   741
         Caption         =   "������ô��۸�?"
         BackColor       =   16777215
         FontSize        =   10.5
         TextAligment    =   0
      End
   End
   Begin MSComctlLib.ImageList ilsImage 
      Left            =   960
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":0614
            Key             =   "search"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":09AE
            Key             =   "hide"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":2AE8
            Key             =   "add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":2E82
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":321C
            Key             =   "remove"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":35B6
            Key             =   "select"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":3950
            Key             =   "unselect"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":3CEA
            Key             =   "down"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":4084
            Key             =   "up"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UsrPriceQuery.ctx":441E
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid msfResult 
      Height          =   3285
      Left            =   525
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   795
      Width           =   4740
      _cx             =   8361
      _cy             =   5794
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483648
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   16761024
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   345
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid msfCalc 
      Height          =   1230
      Left            =   2355
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4770
      Width           =   4740
      _cx             =   8361
      _cy             =   2170
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
      BackColorFixed  =   -2147483648
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   16761024
      TreeColor       =   -2147483643
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
End
Attribute VB_Name = "UsrPriceQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarCurPos1 As Long
Private mvarRows1 As Long

Private mvarCurPos2 As Long
Private mvarRows2 As Long

Private mblnNumber As Boolean

Private mvarStop As Long                '�û���ѯ��Ϣͣ�����
Private mvarScroll As Long

Private mvarRs As New ADODB.Recordset
Private mrsPrice As New ADODB.Recordset         'ʱ��ҩƷ�۸�ƽ���ۣ�
Private mrs����id As ADODB.Recordset
Private mblnUnSelect  As Boolean
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event ClickOK(ByVal strQuery As String, blnCancel As Boolean)

Private Enum mCol
    ����
    ����
    ���
    ����
    ��λ
    �۸�
    ָ���ۼ�
    ����
    ��ʶ����
    ��ʶ����
    ��������
    �۸�����
    ��Ŀ˵��
End Enum

Public Sub InitLoad()
    '��ʼ������
    
    Dim i As Long
    
    mvarCurPos1 = 1
    mvarCurPos2 = 1
    
    UsrCmd(10).Enabled = False
    UsrCmd(11).Enabled = False
    UsrCmd(12).Enabled = False
    UsrCmd(13).Enabled = False
    UsrCmd(14).ShowPicture = False
    UsrCmd(15).ShowPicture = False
    
    msfResult.Rows = 50
    For i = 0 To msfResult.Cols - 1
        msfResult.TextMatrix(1, i) = ""
    Next
    
    msfCalc.Rows = 2
    For i = 0 To msfCalc.Cols - 1
        msfCalc.TextMatrix(1, i) = ""
    Next
                
    Call DrawMsfHeader
    
    UsrCmd(0).Picture = ilsImage.ListImages("delete")
    UsrCmd(1).Picture = ilsImage.ListImages("hide")
    UsrCmd(2).Picture = ilsImage.ListImages("help")
    UsrCmd(3).Picture = ilsImage.ListImages("add")
    
    UsrCmd(11).Picture = ilsImage.ListImages("down")
    UsrCmd(10).Picture = ilsImage.ListImages("up")
    
    UsrCmd(13).Picture = ilsImage.ListImages("down")
    UsrCmd(12).Picture = ilsImage.ListImages("up")
    
    UsrCmd(3).ShowPicture = False
    UsrCmd(0).ShowPicture = False
    UsrCmd(13).ShowPicture = False
    UsrCmd(12).ShowPicture = False
    
    Dim blnHave As Boolean
    Dim strTmp As String
    
    strTmp = Trim(zlDatabase.GetPara("�۸���ʾ���", glngSys, 1536, "000000"))
    
    For i = 4 To 9
        If Val(Mid(strTmp, i - 3, 1)) = 1 Then
'        If Val(GetPara(UsrCmd(i).Caption)) = 1 Then
            UsrCmd(i).Picture = ilsImage.ListImages("select")
            UsrCmd(i).Tag = "1"
            blnHave = True
        Else
            UsrCmd(i).Picture = ilsImage.ListImages("unselect")
            UsrCmd(i).Tag = ""
        End If
    Next
    If blnHave = False Then
        For i = 4 To 9
            UsrCmd(i).Picture = ilsImage.ListImages("select")
            UsrCmd(i).Tag = "1"
        Next
    End If
    
    
    Dim varTmp As Variant
    Dim lngLoop As Long
    
    Set mvarRs = New ADODB.Recordset
    Set mrs����id = New ADODB.Recordset
    mrs����id.Fields.Append "����id", adVarChar, 30, adFldKeyColumn
    mrs����id.Open
    
    mblnUnSelect = False
    strTmp = ""
    strTmp = GetPara("������ʾ���շѷ���")
    If strTmp <> "" Then
        varTmp = Split(strTmp, ",")
        For lngLoop = 0 To UBound(varTmp)
            If CStr(varTmp(lngLoop)) <> "" Then
                
                mrs����id.AddNew
                
                If Left(CStr(varTmp(lngLoop)), 1) = "-" Then
                    mblnUnSelect = True
                    mrs����id("����id").Value = Mid(CStr(varTmp(lngLoop)), 2)
                Else
                    mrs����id("����id").Value = CStr(varTmp(lngLoop))
                End If
                
                
            End If
        Next
    End If
    
    txt(0).Text = ""
    txt(1).Text = ""
                
    UsrCmd(1).Caption = "����������ѯ"
    Call UserControl_Resize
    
    Call CalcMoney
    Call SearchItem("")
    
    tmrInfo.Enabled = True
    
    mvarStop = Val(GetPara("�۸��ѯͣ��ʱ��", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    mvarScroll = Val(GetPara("�۸��ѯ�������", "10"))
    mvarScroll = IIf(mvarScroll <= 0, 10, mvarScroll)
    
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    Dim intPos As Long
    
    Select Case cmdBtn(Index).Caption
    Case "ȷ��"
        If msfResult.TextMatrix(msfResult.Row, mCol.����) <> "" And Val(txt(0).Text) > 0 Then
            mvarCurPos2 = 1
            
            If msfCalc.Rows = 2 And msfCalc.TextMatrix(1, 0) = "" Then
                
            Else
                msfCalc.Rows = msfCalc.Rows + 1
            End If
                                                
            mvarRows2 = msfCalc.Rows
            
            msfCalc.TextMatrix(msfCalc.Rows - 1, 0) = msfResult.TextMatrix(msfResult.Row, mCol.����)
            msfCalc.TextMatrix(msfCalc.Rows - 1, 1) = Val(txt(0).Text)
            
            intPos = InStr(msfResult.TextMatrix(msfResult.Row, mCol.�۸�), "(ָ����)")
            
            If intPos > 0 Then
                msfCalc.TextMatrix(msfCalc.Rows - 1, 2) = Format(Val(txt(0).Text) * Val(msfResult.RowData(msfResult.Row)), "0.00")
            Else
                msfCalc.TextMatrix(msfCalc.Rows - 1, 2) = Format(Val(txt(0).Text) * Val(msfResult.TextMatrix(msfResult.Row, mCol.�۸�)), "0.00")
            End If
            
            Call CalcMoney
            Call EnablePageButton(msfCalc, mvarCurPos2, mvarRows2, UsrCmd(12), UsrCmd(13))
        ElseIf Val(txt(0).Text) <= 0 Then
            MsgBox "������������������������ӣ�", vbInformation, gstrSysName
        ElseIf msfResult.TextMatrix(msfResult.Row, 0) = "" Then
            MsgBox "û��ѡ��Ҫ��ӵ��շ���Ŀ��ǰû���շ���Ŀ��", vbInformation, gstrSysName
        End If

        EnterFocus msfResult
    Case "���"
        txt(0).Text = ""
'        msfCalc.SetFocus
        EnterFocus msfCalc
    Case Else
        txt(0).Text = txt(0).Text & Trim(cmdBtn(Index).Caption)
        'msfResult.SetFocus
        EnterFocus msfResult
    End Select
    
End Sub

Private Sub cmdBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cmdKey_CommandClick(Caption As String)
    Dim strTmp As String
    Dim blnCancel As Boolean

    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("�۸��ѯͣ��ʱ��", 30))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    Select Case Caption
    Case "ȷ��"
        strTmp = txt(1).Text
        tmrScroll.Enabled = False
        lbl(3).Caption = "�۸�����:"
        lbl(4).Caption = "��Ŀ˵��:"
        '�޸ı��2667
        '�ж�������ǲ���adminexitnewquery����û�ȡ����ֱ���˳�����
        RaiseEvent ClickOK(strTmp, blnCancel)
        If blnCancel = True Then Exit Sub
        
        Call SearchItem(strTmp)
        
        msfResult.Row = 1
        Call msfResult_RowColChange
        txt(1).Text = ""
        mvarInfo = Val(GetPara("�۸��ѯͣ��ʱ��", "30"))
        tmrInfo.Enabled = IIf(mvarInfo = 0, False, True)
    Case "���"
        txt(1).Text = ""
    Case Else
        txt(1).Text = txt(1).Text & Trim(Caption)
        txt(1).SelStart = Len(txt(1).Text & Trim(Caption))
    End Select
    'msfResult.SetFocus
    EnterFocus msfResult
End Sub

Private Sub cmdKey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub msfCalc_Click()
    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("�۸��ѯͣ��ʱ��", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
End Sub

Private Sub msfCalc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub msfResult_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call CalcAutoColWidth(msfResult, mCol.����)
    Call SaveFlexState(msfResult, App.ProductName)
End Sub

Private Sub msfResult_Click()
    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("�۸��ѯͣ��ʱ��", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    If UsrCmd(1).Caption = "����������ѯ" Then
        Call UsrCmd_CommandClick(1)
    End If
End Sub

Private Sub msfResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub msfResult_RowColChange()
    
    On Error Resume Next
    lbl(3).Caption = "�۸�����:" & msfResult.TextMatrix(msfResult.Row, mCol.�۸�����)
    lbl(4).Caption = msfResult.TextMatrix(msfResult.Row, mCol.��Ŀ˵��)
    If lbl(4).Caption = "" Then lbl(4).Caption = "��Ŀ˵��:"
'    lbl(4).Caption = "������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������" & _
'                        "������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������������������͸��������ԭ���µĿ���۲�׼ȷʱ�ĵ�������ֻ����һ��������"
'    lbl(4).Caption = "��Ŀ˵��:" & msfResult.TextMatrix(msfResult.Row, mCol.��Ŀ˵��)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picKey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub tmrInfo_Timer()
    If mvarStop > 0 Then
        mvarStop = mvarStop - 1
    Else
        tmrScroll.Enabled = True
    End If
End Sub

Private Sub tmrScroll_Timer()
    
    If mvarScroll > 0 Then
        mvarScroll = mvarScroll - 1
    Else
        If UsrCmd(11).Enabled Then
            If UsrCmd(1).Caption <> "����������ѯ" Then
                Call UsrCmd_CommandClick(1)
            End If
            Call UsrCmd_CommandClick(11)
        Else
            If UsrCmd(1).Caption <> "����������ѯ" Then
                Call UsrCmd_CommandClick(1)
            End If
            Call SearchItem("")
        End If
        mvarScroll = Val(GetPara("�۸��ѯ�������", "10"))
        mvarScroll = IIf(mvarScroll <= 0, 10, mvarScroll)
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If Index = 0 Then
        mblnNumber = True
    Else
        mblnNumber = False
    End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    tmrScroll.Enabled = False
    mvarStop = Val(GetPara("�۸��ѯͣ��ʱ��", "30"))
    mvarStop = IIf(mvarStop <= 0, 30, mvarStop)
    
    Select Case KeyCode
    Case vbKeyA
        txt(1).Text = txt(1).Text & "A"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyB
        txt(1).Text = txt(1).Text & "B"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyC
        txt(1).Text = txt(1).Text & "C"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyD
        txt(1).Text = txt(1).Text & "D"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyE
        txt(1).Text = txt(1).Text & "E"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyF
        txt(1).Text = txt(1).Text & "F"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyG
        txt(1).Text = txt(1).Text & "G"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyH
        txt(1).Text = txt(1).Text & "H"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyI
        txt(1).Text = txt(1).Text & "I"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyJ
        txt(1).Text = txt(1).Text & "J"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyK
        txt(1).Text = txt(1).Text & "K"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyL
        txt(1).Text = txt(1).Text & "L"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyM
        txt(1).Text = txt(1).Text & "M"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyN
        txt(1).Text = txt(1).Text & "N"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyO
        txt(1).Text = txt(1).Text & "O"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyP
        txt(1).Text = txt(1).Text & "P"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyQ
        txt(1).Text = txt(1).Text & "Q"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyR
        txt(1).Text = txt(1).Text & "R"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyS
        txt(1).Text = txt(1).Text & "S"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyT
        txt(1).Text = txt(1).Text & "T"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyU
        txt(1).Text = txt(1).Text & "U"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyV
        txt(1).Text = txt(1).Text & "V"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyW
        txt(1).Text = txt(1).Text & "W"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyX
        txt(1).Text = txt(1).Text & "X"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyY
        txt(1).Text = txt(1).Text & "Y"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKeyZ
        txt(1).Text = txt(1).Text & "Z"
        txt(1).SelStart = Len(txt(1).Text) + 1
    Case vbKey0, vbKeyNumpad0
        txt(0).Text = txt(0).Text & "0"
    Case vbKey1, vbKeyNumpad1
        txt(0).Text = txt(0).Text & "1"
    Case vbKey2, vbKeyNumpad2
        txt(0).Text = txt(0).Text & "2"
    Case vbKey3, vbKeyNumpad3
        txt(0).Text = txt(0).Text & "3"
    Case vbKey4, vbKeyNumpad4
        txt(0).Text = txt(0).Text & "4"
    Case vbKey5, vbKeyNumpad5
        txt(0).Text = txt(0).Text & "5"
    Case vbKey6, vbKeyNumpad6
        txt(0).Text = txt(0).Text & "6"
    Case vbKey7, vbKeyNumpad7
        txt(0).Text = txt(0).Text & "7"
    Case vbKey8, vbKeyNumpad8
        txt(0).Text = txt(0).Text & "8"
    Case vbKey9, vbKeyNumpad9
        txt(0).Text = txt(0).Text & "9"
    Case vbKeyReturn, vbKeySeparator
        If txt(0).Text <> "" Then
            Call cmdBtn_Click(10)
        Else
            Call cmdKey_CommandClick("ȷ��")
        End If
    Case vbKeyDecimal
        Call cmdBtn_Click(11)
    Case vbKeyDelete
        Call cmdKey_CommandClick("���")
        Call cmdBtn_Click(12)
    End Select
    
    If KeyCode <> 27 Then KeyCode = 0
    
    If KeyCode = 27 Then
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
    
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    Dim vHeight As Single
    
    If UsrCmd(1).Caption = "���ز�ѯ��" Then
        vHeight = picKey.Height + picBack(0).Height + 45
        txt(0).Visible = True
        txt(1).Visible = True
        lbl(0).Visible = True
        lbl(2).Visible = True
        UsrCmd(0).Visible = True
        UsrCmd(3).Visible = True
    Else
        vHeight = picBack(0).Height
        txt(0).Visible = False
        txt(1).Visible = False
        lbl(0).Visible = False
        lbl(2).Visible = False
        UsrCmd(0).Visible = False
        UsrCmd(3).Visible = False
    End If
    
    Call ResizeControl(picBack(1), 0, 0, UserControl.Width, picBack(1).Height)
    Call ResizeControl(msfResult, 0, picBack(1).Top + picBack(1).Height + 15, UserControl.Width - 15, UserControl.Height - picBack(1).Height - vHeight)
    Call ResizeControl(picBack(0), 0, msfResult.Top + msfResult.Height + 15, UserControl.Width, picBack(0).Height)

    
'
    Call ResizeControl(msfCalc, pic.Width, picBack(0).Top + picBack(0).Height + 15, UserControl.Width - picKey.Width - pic.Width, picKey.Height - picBack(2).Height - picBack(3).Height)
    Call ResizeControl(picBack(3), msfCalc.Left, msfCalc.Top + msfCalc.Height, msfCalc.Width, picBack(3).Height)
    Call ResizeControl(pic, 0, msfCalc.Top, pic.Width, pic.Height)
    
    Call ResizeControl(picKey, msfCalc.Left + msfCalc.Width + 30, msfCalc.Top, picKey.Width, picKey.Height)
    Call ResizeControl(picBack(2), pic.Left, picBack(3).Top + picBack(3).Height, msfCalc.Width + pic.Width, picBack(2).Height)
    
    
    
    UsrCmd(2).Left = picBack(1).ScaleWidth - UsrCmd(2).Width - 30
    UsrCmd(11).Left = UsrCmd(2).Left - UsrCmd(11).Width - 30
    UsrCmd(10).Left = UsrCmd(11).Left - UsrCmd(10).Width - 30

    UsrCmd(13).Left = picBack(2).ScaleWidth - UsrCmd(13).Width - 30
    UsrCmd(12).Left = UsrCmd(13).Left - UsrCmd(12).Width - 30

    lbl(2).Left = picKey.Left
    

    txt(0).Left = lbl(0).Left + lbl(0).Width
    txt(0).Top = lbl(0).Top - 60
    txt(1).Top = UsrCmd(1).Top
    lbl(2).Top = txt(1).Top + 60
    lbl(3).Left = lbl(2).Left

    UsrCmd(1).Left = picBack(0).ScaleWidth - UsrCmd(1).Width - 30
    txt(1).Left = UsrCmd(1).Left - txt(1).Width - 30
    lbl(2).Left = txt(1).Left - lbl(2).Width - 15
    
    
    fra.Move lbl(3).Left - 30, fra.Top, picKey.Width
    fra2.Move lbl(3).Left - 30, -90, fra2.Width, picBack(0).Height + 90
    
    lbl(4).Move 0, 0, fra.Left - UsrCmd(14).Width - 30 - 60, picBack(0).Height
    
    UsrCmd(14).Left = fra.Left - UsrCmd(14).Width - 60
    
    UsrCmd(15).Left = fra.Left - UsrCmd(15).Width - 60
    
End Sub

Private Sub DrawMsfHeader()

    msfResult.Cols = 0
        
    Call AddColumn(msfResult, "����", 1080, 1)
    Call AddColumn(msfResult, "����", 4020, 1)
    Call AddColumn(msfResult, "���", 2700, 1)
    Call AddColumn(msfResult, "����", 900, 1)
    Call AddColumn(msfResult, "��λ", 600, 1)
    Call AddColumn(msfResult, "�۸�", 1800, 7)
    Call AddColumn(msfResult, "ָ���ۼ�", 1800, 7)
    
    Call AddColumn(msfResult, "����", 2100, 1)
    Call AddColumn(msfResult, "��ʶ����", 1080, 1)
    Call AddColumn(msfResult, "��ʶ����", 1080, 1)
    
    Call AddColumn(msfResult, "�������", 1200, 1)
    Call AddColumn(msfResult, "�۸�����", 0, 1)
    Call AddColumn(msfResult, "��Ŀ˵��", 0, 1)
    Call AddColumn(msfResult, "", 1200, 1)
    
    Call RestoreFlexState(msfResult, App.ProductName)
    
    Dim strTmp As String
    
    strTmp = Trim(zlDatabase.GetPara("�۸���ʾ��Ϣ", glngSys, 1536, "0000011"))
    If Len(strTmp) = 6 Then strTmp = strTmp & "1"
    
    
    If Val(Mid(strTmp, 1, 1)) = 1 Then msfResult.ColHidden(mCol.��������) = True
    If Val(Mid(strTmp, 2, 1)) = 1 Then msfResult.ColHidden(mCol.����) = True
    If Val(Mid(strTmp, 3, 1)) = 1 Then msfResult.ColHidden(mCol.����) = True
    If Val(Mid(strTmp, 4, 1)) = 1 Then msfResult.ColHidden(mCol.��ʶ����) = True
    If Val(Mid(strTmp, 5, 1)) = 1 Then msfResult.ColHidden(mCol.��ʶ����) = True
    If Val(Mid(strTmp, 6, 1)) = 1 Then msfResult.ColHidden(mCol.ָ���ۼ�) = True
    If Val(Mid(strTmp, 7, 1)) = 1 Then msfResult.ColHidden(mCol.����) = True
    
    msfCalc.Cols = 0
    Call AddColumn(msfCalc, "����", 3030, 1)
    Call AddColumn(msfCalc, "����", 540, 7)
    Call AddColumn(msfCalc, "���", 810, 7)
    Call AddColumn(msfCalc, "", 15, 1)
    
    Call CalcAutoColWidth(msfResult, mCol.����)
    Call CalcAutoColWidth(msfCalc, 0)
           
End Sub


Private Sub UserControl_Show()
    cmdKey.KeyMode = 1
End Sub

Private Sub UsrCmd_CommandClick(Index As Integer)
    Dim i As Long
    
    If Index >= 4 And Index < 10 Then
        UsrCmd(Index).Picture = ilsImage.ListImages(IIf(UsrCmd(Index).Tag = "1", "unselect", "select"))
        UsrCmd(Index).Tag = IIf(UsrCmd(Index).Tag = "1", "", "1")
        DoEvents
        Call SearchItem(txt(1).Text)
        Exit Sub
    End If
    
    Select Case Index
    Case 0
        If msfCalc.Row < 1 Then Exit Sub
        If msfCalc.TextMatrix(msfCalc.Row, 0) <> "" Then
            mvarRows2 = mvarRows2 - 1
            mvarCurPos2 = 1
            If msfCalc.Rows <= 2 Then
                Call ClearSpecRowCol(msfCalc, 1, Array())
            Else
                msfCalc.RemoveItem msfCalc.Row
            End If
            'msfCalc.Rows = mvarRows2 + 10
            Call CalcMoney
            Call EnablePageButton(msfCalc, mvarCurPos2, mvarRows2, UsrCmd(12), UsrCmd(13))
        End If
    Case 1
        If UsrCmd(Index).Caption = "���ز�ѯ��" Then
            '���ز�ѯ��
            UsrCmd(Index).Caption = "����������ѯ"
        Else
            UsrCmd(Index).Caption = "���ز�ѯ��"
        End If
        Call UserControl_Resize
    Case 2
        Call frmHelp.ShowHelp(Me, -1, UserControl.Width, UserControl.Height)
    Case 3
        Call cmdBtn_Click(10)
    Case 10
        Call TurnToPage(msfResult, -1, mvarCurPos1)
        Call EnablePageButton(msfResult, mvarCurPos1, mvarRows1, UsrCmd(10), UsrCmd(11))
    Case 11             '��һҳ
        Call TurnToPage(msfResult, 1, mvarCurPos1)
        Call EnablePageButton(msfResult, mvarCurPos1, mvarRows1, UsrCmd(10), UsrCmd(11))
    Case 12
        Call TurnToPage(msfCalc, -1, mvarCurPos2)
        Call EnablePageButton(msfCalc, mvarCurPos2, mvarRows2, UsrCmd(12), UsrCmd(13))
    Case 13             '��һҳ
        Call TurnToPage(msfCalc, 1, mvarCurPos2)
        Call EnablePageButton(msfCalc, mvarCurPos2, mvarRows2, UsrCmd(12), UsrCmd(13))
    Case 14
        
        If (lbl(4).Top + lbl(4).Height) > picBack(0).Height Then lbl(4).Top = lbl(4).Top - 210
    Case 15
        If lbl(4).Top < 0 Then lbl(4).Top = lbl(4).Top + 210
    End Select
End Sub


Private Sub SearchItem(ByVal strKey As String)
    Dim strLike As String
    Dim strInput As String
    Dim strSort As String
    Dim i As Long
    Dim lngSvrRow As Long
    Dim sgl�۸� As Single
    Dim rs As New ADODB.Recordset
    Dim lngBkColor As Long
    Dim strTmp As String
    Dim blnAllow As Boolean
    
    On Error GoTo errHand
    
    lngBkColor = 15987699
    
    strTmp = GetPara("������ʾ���շ����")
    
    strSort = " ���='999' "
    If UsrCmd(4).Tag = "1" Then strSort = strSort & " OR ���='5' OR ���='6' OR ���='7'"
    If UsrCmd(5).Tag = "1" Then strSort = strSort & " OR ���='C'"
    If UsrCmd(6).Tag = "1" Then strSort = strSort & " OR ���='D'"
    If UsrCmd(7).Tag = "1" Then strSort = strSort & " OR ���='E'"
    If UsrCmd(8).Tag = "1" Then strSort = strSort & " OR ���='F'"
    
    If mrsPrice.State <> adStateOpen Then
        gstrSQL = "Select a.ҩƷid,Sum(ʵ�ʽ��)/Sum(ʵ������) As ���� from ҩƷ��� a,�շ���ĿĿ¼ b where a.ҩƷid=b.ID And Nvl(b.�Ƿ���,0)=1 And " & GetNodeCheckSQL("b.վ��") & " Group By a.ҩƷid Having Sum(ʵ������)<>0"
        Set mrsPrice = zlDatabase.OpenSQLRecord(gstrSQL, "�۸��ѯ")
    End If
    
    If UsrCmd(9).Tag = "1" Then
        '�����������������ָ����ǰ��ļ����������������
        strSort = strSort & " OR (���<>'5' AND ���<>'6' AND ���<>'7' AND ���<>'C' AND ���<>'D' AND ���<>'E' AND ���<>'F')"
    End If

    
    strSearchSQL = ""
    If strKey <> "" Then
        strInput = "%" & strKey & "%"
        strSearchSQL = " AND (Y.���� Like [1] OR Upper(Y.����) Like [1] OR Y.ID IN (SELECT �շ�ϸĿID From �շ���Ŀ���� WHERE UPPER(����) Like [1] OR Upper(����) Like [1]))"
    End If
            
    '�Ǳ�۵����շ���Ŀ,ҩƷ�����Ŀ����ʾ�۸�ΪҩƷָ�����ۼ�,���������Ŀ����ʾ
    gstrSQL = "" & _
        "Select A.���,A.����,A.����,A.���,A.����,A.��λ,A.����,A.��ʶ����,A.��ʶ����,A.�۸�����,A.��������,A.��Ŀ˵��,A.��������,A.�Ƿ����,A.�Ƿ���,A.�ּ�,A.ָ�����ۼ�,A.����ID,A.����ID,A.����ID,DECODE(�ּ�,0,NULL,DECODE(a.��������, 0, 1, NULL, 1, a.��������) * a.�ּ�) AS �۸� " & _
        "From ( " & _
        "Select ������� As ���, " & _
               "����, " & _
               "����, " & _
               "���,����, " & _
               "��λ,����,��ʶ����,��ʶ����, " & _
               "�۸�����, " & _
               "��������, " & _
               "��Ŀ˵��, " & _
               "��������, " & _
               "Decode(����id,0,0,1) As �Ƿ����, " & _
               "Decode(���, '5', 1, '6', 1, '7', 1, 0) *Decode(�Ƿ���, 1, 1, 0) As �Ƿ���, " & _
               "Decode(Decode(���, '5', 1, '6', 1, '7', 1, 0) * Decode(�Ƿ���, 1, 1, 0),1,ָ�����ۼ�,�ּ�) as �ּ�,ָ�����ۼ�,����id,����id,����id " & _
        "From ( " & _
        "Select X.����id,X.����id, " & _
               "Y.��� As �������, " & _
               "Decode(X.����id,0,Y.���,Z.���) As ���, " & _
               "Decode(X.����id,0,Y.����,Z.����) As ����, " & _
               "Decode(X.����id,0,Y.����,'  '||Z.����||X.����˵��) As ����, " & _
               "Decode(X.����id,0,Y.���,Z.���) As ���, " & _
               "Decode(X.����id,0,Y.���㵥λ,Z.���㵥λ) As ��λ, " & _
               "Decode(X.����id,0,Y.��ʶ����,Z.��ʶ����) As ��ʶ����,Decode(X.����id,0,Y.��ʶ����,Z.��ʶ����) As ��ʶ����, " & _
               "P.�ּ�,y.����,m.����, "
                       
    gstrSQL = gstrSQL & _
               "p.ָ�����ۼ�," & _
               "x.��������," & _
               "x.����˵��," & _
               "P.����˵�� AS �۸�����," & _
               "Decode(X.����id,0,Y.��������,Z.��������) As ��������," & _
               "Decode(X.����id,0,Y.˵��,Z.˵��) As ��Ŀ˵��," & _
               "Decode(X.����id, 0, Y.�Ƿ���, Z.�Ƿ���) As �Ƿ���,Decode(m.����id,Null,'P'||To_Char(y.����id),'K'||To_Char(m.����id)) As ����id " & _
        "From ( Select ID,ID As ����id,0 As ����id,'' As ����˵��,0 As �������� From �շ���ĿĿ¼ Where ������� In (1,2,3) And " & GetNodeCheckSQL("վ��") & " " & _
               "Union All " & _
               "Select ����id AS ID,����ID,����id,DECODE(���д���, 2, '[��', 1, '[��', '[��') ||to_char(��������) || ']' AS ����˵��,�������� From �շѴ�����Ŀ a,�շ���ĿĿ¼ b where b.id=a.����id and b.������� In (1,2,3) And " & GetNodeCheckSQL("b.վ��") & " " & _
              ") x, " & _
              "( Select k.�շ�ϸĿID,K.����˵��,K.�ּ�,Decode(t.ָ�����ۼ�,Null,s.ָ�����ۼ�,t.ָ�����ۼ�) As ָ�����ۼ� " & _
                "From ҩƷ��� t,�������� s," & _
                     "( Select �շ�ϸĿID,����˵��,Sum(�ּ�) As �ּ� " & _
                     "From �շѼ�Ŀ " & _
                     "Where (��ֹ���� is Null OR ��ֹ���� = TO_DATE('3000-01-01', 'YYYY-MM-DD')) Group By �շ�ϸĿID,����˵�� " & _
                     ") k " & _
                "Where k.�շ�ϸĿID=t.ҩƷid(+) And k.�շ�ϸĿID=s.����ID(+) " & _
              ") p, " & _
              "�շ���ĿĿ¼ y, " & _
              "�շ���ĿĿ¼ z, " & _
              "(Select t.ҩƷid,w.����id,p.ҩƷ���� As ���� From ҩƷ��� t,������ĿĿ¼ w,ҩƷ���� p Where t.ҩ��id=w.ID And p.ҩ��id=w.ID And " & GetNodeCheckSQL("w.վ��") & ") m " & _
        "Where x.����ID = y.ID and y.������� In (1,2,3) And " & GetNodeCheckSQL("y.վ��") & " And " & GetNodeCheckSQL("z.վ��") & " AND x.ID=p.�շ�ϸĿID(+) AND x.����id=z.ID(+) And y.ID=m.ҩƷid(+) "
    
    If strTmp <> "" Then
        gstrSQL = gstrSQL & " And y.��� In (" & strTmp & ")"
    End If
    
    gstrSQL = gstrSQL & _
              "AND (Y.����ʱ�� is null OR Y.����ʱ�� = TO_DATE('3000-01-01', 'YYYY-MM-DD')) " & _
              "AND (Z.����ʱ�� is null OR Z.����ʱ�� = TO_DATE('3000-01-01', 'YYYY-MM-DD')) " & strSearchSQL & _
        "Order By Y.����,X.����id " & _
        ")) a"

    mvarCurPos1 = 1
    mvarRows1 = 0
    UsrCmd(10).Enabled = False
    UsrCmd(11).Enabled = False
                
    i = 1
    msfResult.Rows = 2
    Call ClearSpecRowCol(msfResult, 1, Array())
    
    ShowFlatFlash "������ȡ�۸���Ϣ..."
    DoEvents
    
    If strKey = "" Then
        If mvarRs.State = adStateOpen Then
            mvarRs.Filter = ""
            mvarRs.Filter = strSort
        Else
            Set mvarRs = zlDatabase.OpenSQLRecord(gstrSQL, "�۸��ѯ")
            mvarRs.Filter = strSort
        End If
        Set rs = mvarRs
    Else
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "�۸��ѯ", strInput)
    End If
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            blnAllow = False
            mrs����id.Filter = ""
            If mrs����id.RecordCount > 0 And zlCommFun.Nvl(rs("����id").Value) <> "P" Then
                If mblnUnSelect Then
                    mrs����id.Filter = "����id='" & zlCommFun.Nvl(rs("����id").Value) & "'"
                    blnAllow = Not (mrs����id.RecordCount > 0)
                Else
                    mrs����id.Filter = "����id='" & zlCommFun.Nvl(rs("����id").Value) & "'"
                    blnAllow = (mrs����id.RecordCount > 0)
                End If
            Else
                blnAllow = True
            End If
            
            If blnAllow Then
                msfResult.TextMatrix(i, mCol.����) = IIf(IsNull(rs!����), "", rs!����)
                msfResult.TextMatrix(i, mCol.����) = IIf(IsNull(rs!����), "", rs!����)
                msfResult.TextMatrix(i, mCol.���) = IIf(IsNull(rs!���), "", rs!���)
                msfResult.TextMatrix(i, mCol.����) = IIf(IsNull(rs!����), "", rs!����)
                msfResult.TextMatrix(i, mCol.��λ) = IIf(IsNull(rs!��λ), "", rs!��λ)
                msfResult.TextMatrix(i, mCol.����) = IIf(IsNull(rs!����), "", rs!����)
                msfResult.TextMatrix(i, mCol.��ʶ����) = IIf(IsNull(rs!��ʶ����), "", rs!��ʶ����)
                msfResult.TextMatrix(i, mCol.��ʶ����) = IIf(IsNull(rs!��ʶ����), "", rs!��ʶ����)
                msfResult.TextMatrix(i, mCol.�۸�) = IIf(IsNull(rs!�۸�), "", Format(rs!�۸�, "0.00##"))
                msfResult.TextMatrix(i, mCol.ָ���ۼ�) = IIf(IsNull(rs!ָ�����ۼ�), "", Format(rs!ָ�����ۼ�, "0.00##"))
                msfResult.TextMatrix(i, mCol.��������) = IIf(IsNull(rs!��������), "", rs!��������)
                msfResult.TextMatrix(i, mCol.�۸�����) = IIf(IsNull(rs!�۸�����), "", rs!�۸�����)
                msfResult.TextMatrix(i, mCol.��Ŀ˵��) = IIf(IsNull(rs!��Ŀ˵��), "", rs!��Ŀ˵��)
                msfResult.RowData(i) = Val(msfResult.TextMatrix(i, mCol.�۸�))
                
                '����ʱ��ҩƷ�ļ۸�
                msfResult.Cell(flexcpData, i, mCol.�۸�, i, mCol.�۸�) = 0
                Select Case zlCommFun.Nvl(rs("���").Value)
                Case "4", "5", "6", "7"
                    'ҩƷ,����
                    
                    If zlCommFun.Nvl(rs("�Ƿ���").Value, 0) = 1 Then
                        msfResult.Cell(flexcpData, i, mCol.�۸�, i, mCol.�۸�) = 1
                        mrsPrice.Filter = ""
                        
                        If zlCommFun.Nvl(rs("����id").Value, 0) > 0 Then
                            mrsPrice.Filter = "ҩƷid=" & zlCommFun.Nvl(rs("����id").Value, 0)
                        Else
                            mrsPrice.Filter = "ҩƷid=" & zlCommFun.Nvl(rs("����id").Value, 0)
                        End If
                        
                        If mrsPrice.RecordCount > 0 Then
                            If zlCommFun.Nvl(mrsPrice("����").Value, 0) > 0 Then
                                msfResult.TextMatrix(i, mCol.�۸�) = Format(zlCommFun.Nvl(mrsPrice("����").Value, 0), "0.00##")
                                msfResult.Cell(flexcpData, i, mCol.�۸�, i, mCol.�۸�) = 0
                            End If
                        End If
                    End If
                End Select
            
            
                If zlCommFun.Nvl(rs("�Ƿ����"), 0) = 1 Then
                
                    '�ۼӼ۸�
                    sgl�۸� = sgl�۸� + Val(msfResult.TextMatrix(i, mCol.�۸�))
                    
                Else
                    If sgl�۸� > 0 Then
                        If InStr(msfResult.TextMatrix(lngSvrRow, mCol.�۸�), "(ָ����)") > 0 Then
                            msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = Left(msfResult.TextMatrix(lngSvrRow, mCol.�۸�), Len(msfResult.TextMatrix(lngSvrRow, mCol.�۸�)) - 5)
                            msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = Format(Val(msfResult.TextMatrix(lngSvrRow, mCol.�۸�)) + sgl�۸�, "0.00##")
                            If Val(msfResult.TextMatrix(lngSvrRow, mCol.�۸�)) = 0 Then
                                msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = ""
                            Else
                                msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = msfResult.TextMatrix(lngSvrRow, mCol.�۸�) & "(ָ����)"
                            End If
                        Else
                            msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = Format(Val(msfResult.TextMatrix(lngSvrRow, mCol.�۸�)) + sgl�۸�, "0.00##")
                        End If
                        
                        sgl�۸� = 0
                    End If
                    lngSvrRow = i
                    
                End If
            
                If i Mod 2 = 0 Then msfResult.Cell(flexcpBackColor, i, 0, i, msfResult.Cols - 1) = lngBkColor
    
                If msfResult.Cell(flexcpData, i, mCol.�۸�, i, mCol.�۸�) = 1 Then
                    msfResult.TextMatrix(i, mCol.�۸�) = msfResult.TextMatrix(i, mCol.�۸�) & "(ָ����)"
                End If
                
                i = i + 1
                msfResult.Rows = i + 1
            End If
            
            rs.MoveNext
            If msfResult.Rows = 30 Then DoEvents
        Loop
        
        If sgl�۸� > 0 Then
            If InStr(msfResult.TextMatrix(lngSvrRow, mCol.�۸�), "(ָ����)") > 0 Then
                msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = Left(msfResult.TextMatrix(lngSvrRow, mCol.�۸�), Len(msfResult.TextMatrix(lngSvrRow, mCol.�۸�)) - 5)
                msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = Format(Val(msfResult.TextMatrix(lngSvrRow, mCol.�۸�)) + sgl�۸�, "0.00##")
                If Val(msfResult.TextMatrix(lngSvrRow, mCol.�۸�)) = 0 Then
                    msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = ""
                Else
                    msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = msfResult.TextMatrix(lngSvrRow, mCol.�۸�) & "(ָ����)"
                End If
            Else
                msfResult.TextMatrix(lngSvrRow, mCol.�۸�) = Format(Val(msfResult.TextMatrix(lngSvrRow, mCol.�۸�)) + sgl�۸�, "0.00##")
            End If
            
            sgl�۸� = 0
        End If
        If msfResult.Rows > 2 Then msfResult.Rows = msfResult.Rows - 1
        mvarCurPos1 = 1
        mvarRows1 = msfResult.Rows - 1
    End If
        
    StopFlatFlash
    
    msfResult.Rows = msfResult.Rows + 50
    Call EnablePageButton(msfResult, mvarCurPos1, mvarRows1, UsrCmd(10), UsrCmd(11))
    Call msfResult_RowColChange
    
    Exit Sub
    
errHand:
    StopFlatFlash
End Sub

Private Sub CalcMoney()
    Dim i As Long
    Dim vTmp As Single

    vTmp = 0
    For i = 1 To msfCalc.Rows - 1
        vTmp = vTmp + Val(msfCalc.TextMatrix(i, 2))
    Next
    lbl(1).Caption = "�ϼ�:" & Format(vTmp, "0.00Ԫ")
End Sub

Private Sub UsrCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Let Enabled(ByVal vData As Boolean)
    UserControl.Enabled = vData
End Property

