VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPacsApplication 
   Caption         =   "�������"
   ClientHeight    =   10860
   ClientLeft      =   135
   ClientTop       =   495
   ClientWidth     =   14055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPacsApplication.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   612
      Left            =   6600
      TabIndex        =   52
      Top             =   120
      Width           =   5052
      _Version        =   589884
      _ExtentX        =   8911
      _ExtentY        =   1080
      _StockProps     =   64
   End
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   12450
      Top             =   675
   End
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   11820
      Top             =   660
   End
   Begin VSFlex8Ctl.VSFlexGrid vsMethod 
      Height          =   1815
      Left            =   7020
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6855
      Visible         =   0   'False
      Width           =   2055
      _cx             =   1993543209
      _cy             =   1993542785
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
      BackColor       =   12648447
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   12648447
      BackColorAlternate=   12648447
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPacsApplication.frx":1422
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   1
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
   Begin XtremeSuiteControls.TabControl tbcRequest 
      Height          =   336
      Left            =   0
      TabIndex        =   21
      Top             =   960
      Width           =   6420
      _Version        =   589884
      _ExtentX        =   11324
      _ExtentY        =   593
      _StockProps     =   64
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   11310
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   705
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   15
      ScaleHeight     =   795
      ScaleWidth      =   6285
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   30
      Width           =   6285
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   9900
      Left            =   -60
      ScaleHeight     =   9900
      ScaleWidth      =   13950
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   13950
      Begin zlCISKernel.ucScrollPanel uspRequestPage 
         Bindings        =   "frmPacsApplication.frx":145E
         Height          =   9915
         Left            =   75
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   13785
         _ExtentX        =   24289
         _ExtentY        =   17515
         UCBackColor     =   14737632
         Begin VB.PictureBox picPart 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3900
            Left            =   0
            ScaleHeight     =   3900
            ScaleWidth      =   13575
            TabIndex        =   5
            TabStop         =   0   'False
            Tag             =   "5330-105"
            Top             =   4425
            Visible         =   0   'False
            Width           =   13575
            Begin VB.CheckBox chkPriority 
               BackColor       =   &H00FFFFFF&
               Caption         =   "����ҽ��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   4575
               TabIndex        =   7
               Top             =   270
               Width           =   1185
            End
            Begin MSComCtl2.DTPicker dtpExeTime 
               Height          =   360
               Left            =   6855
               TabIndex        =   10
               Top             =   240
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   635
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy-MM-dd HH:mm"
               Format          =   251330563
               CurrentDate     =   41348.5555555556
            End
            Begin VB.ComboBox cbxExeRoom 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   10425
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   240
               Width           =   3015
            End
            Begin MSComctlLib.ImageList img16 
               Left            =   5415
               Top             =   1125
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
                     Picture         =   "frmPacsApplication.frx":1472
                     Key             =   "c0"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmPacsApplication.frx":1A0C
                     Key             =   "c1"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmPacsApplication.frx":1FA6
                     Key             =   "o0"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmPacsApplication.frx":2540
                     Key             =   "o1"
                  EndProperty
               EndProperty
            End
            Begin VB.TextBox txtFind 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1140
               TabIndex        =   6
               Top             =   240
               Width           =   3300
            End
            Begin VSFlex8Ctl.VSFlexGrid vfgList 
               Height          =   3135
               Left            =   5865
               TabIndex        =   12
               Top             =   675
               Width           =   7590
               _cx             =   13388
               _cy             =   5530
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
               BackColorSel    =   16769985
               ForeColorSel    =   0
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   2
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   7
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmPacsApplication.frx":2ADA
               ScrollTrack     =   -1  'True
               ScrollBars      =   2
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
               Begin VB.CommandButton cmd 
                  Caption         =   "��"
                  Height          =   240
                  Left            =   3435
                  TabIndex        =   8
                  TabStop         =   0   'False
                  ToolTipText     =   "ѡ����Ŀ(*)"
                  Top             =   1035
                  Visible         =   0   'False
                  Width           =   270
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid vfgRequestProject 
               Height          =   3150
               Left            =   135
               TabIndex        =   9
               Top             =   675
               Width           =   5640
               _cx             =   9948
               _cy             =   5556
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   0
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
               BackColorSel    =   16769985
               ForeColorSel    =   0
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   2
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   ""
               ScrollTrack     =   -1  'True
               ScrollBars      =   2
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
               FrozenRows      =   1
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
            Begin VB.Label labExeRoom 
               BackStyle       =   0  'Transparent
               Caption         =   "ִ�п��ң�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   9435
               TabIndex        =   44
               Top             =   300
               Width           =   1530
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "ִ��ʱ�䣺"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   5850
               TabIndex        =   42
               Top             =   300
               Width           =   1575
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "��Ŀ��λ��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   165
               TabIndex        =   39
               Top             =   300
               Width           =   1350
            End
            Begin VB.Label labPartArea 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��鲿λ����"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   4950
               TabIndex        =   22
               Top             =   975
               Visible         =   0   'False
               Width           =   1800
            End
         End
         Begin VB.PictureBox picBaseInf 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1185
            Left            =   0
            ScaleHeight     =   1185
            ScaleWidth      =   13635
            TabIndex        =   3
            TabStop         =   0   'False
            Tag             =   "90-0"
            Top             =   90
            Visible         =   0   'False
            Width           =   13635
            Begin VB.PictureBox picAuditing 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               FillColor       =   &H80000008&
               ForeColor       =   &H80000008&
               Height          =   525
               Left            =   150
               ScaleHeight     =   495
               ScaleWidth      =   1125
               TabIndex        =   48
               Top             =   105
               Visible         =   0   'False
               Width           =   1155
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "��У��"
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000C000&
                  Height          =   270
                  Left            =   90
                  TabIndex        =   49
                  Top             =   120
                  Width           =   915
               End
            End
            Begin VB.Label labConditionTag 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "��!!"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   12090
               TabIndex        =   47
               Top             =   120
               Width           =   675
            End
            Begin VB.Shape shpBase 
               BorderStyle     =   0  'Transparent
               FillStyle       =   0  'Solid
               Height          =   60
               Left            =   0
               Top             =   1095
               Width           =   13590
            End
            Begin VB.Label labInPatientValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "123456789"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   11490
               TabIndex        =   35
               Top             =   735
               Width           =   1410
            End
            Begin VB.Label labInPatient 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺ�ţ�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   10410
               TabIndex        =   34
               Top             =   735
               Width           =   1200
            End
            Begin VB.Label labOutPatientValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "123456789"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8190
               TabIndex        =   33
               Top             =   735
               Width           =   1410
            End
            Begin VB.Label labAgeValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "30��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5715
               TabIndex        =   32
               Top             =   735
               Width           =   585
            End
            Begin VB.Label labSexValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ů"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3735
               TabIndex        =   31
               Top             =   735
               Width           =   285
            End
            Begin VB.Label labNameValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1005
               TabIndex        =   30
               Top             =   735
               Width           =   855
            End
            Begin VB.Label labOutPatientNo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ţ�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7110
               TabIndex        =   29
               Top             =   735
               Width           =   1200
            End
            Begin VB.Label labAge 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���䣺"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4890
               TabIndex        =   28
               Top             =   735
               Width           =   855
            End
            Begin VB.Label labSex 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�Ա�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2985
               TabIndex        =   27
               Top             =   735
               Width           =   870
            End
            Begin VB.Label labName 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   165
               TabIndex        =   26
               Top             =   735
               Width           =   900
            End
            Begin VB.Label labTitle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CT������뵥"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   21.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   4845
               TabIndex        =   25
               Top             =   105
               Width           =   2730
            End
            Begin VB.Label labBaseArea 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������Ϣ����"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   2520
               TabIndex        =   24
               Top             =   405
               Visible         =   0   'False
               Width           =   1800
            End
         End
         Begin VB.PictureBox picInput 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3150
            Left            =   15
            ScaleHeight     =   3150
            ScaleWidth      =   13575
            TabIndex        =   2
            TabStop         =   0   'False
            Tag             =   "1865-90"
            Top             =   1275
            Visible         =   0   'False
            Width           =   13575
            Begin VB.CommandButton cmdInput 
               Appearance      =   0  'Flat
               Caption         =   "��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   12990
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   45
               Width           =   375
            End
            Begin VB.TextBox rtbInputPro 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   1680
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               TabStop         =   0   'False
               Text            =   "frmPacsApplication.frx":2BE6
               Top             =   15
               Visible         =   0   'False
               Width           =   11280
            End
            Begin VB.Label labMustPro 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   0
               Left            =   585
               TabIndex        =   41
               Top             =   345
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.Shape shpBackLine 
               BorderStyle     =   0  'Transparent
               FillStyle       =   0  'Solid
               Height          =   30
               Index           =   0
               Left            =   1680
               Top             =   405
               Visible         =   0   'False
               Width           =   11715
            End
            Begin VB.Shape shpInputPro 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00E0E0E0&
               FillStyle       =   0  'Solid
               Height          =   45
               Index           =   0
               Left            =   60
               Top             =   1395
               Visible         =   0   'False
               Width           =   13380
            End
            Begin VB.Label labInputPro 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "¼����Ŀ��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   45
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label labInputArea 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "¼����Ŀ����"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   4680
               TabIndex        =   23
               Top             =   1800
               Visible         =   0   'False
               Width           =   1800
            End
         End
         Begin VB.PictureBox picRequestInf 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1485
            Left            =   0
            ScaleHeight     =   1485
            ScaleWidth      =   13545
            TabIndex        =   13
            TabStop         =   0   'False
            Tag             =   "9405-135"
            Top             =   8340
            Visible         =   0   'False
            Width           =   13545
            Begin VB.TextBox txtCurStudyProject 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   15
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   90
               Width           =   13470
            End
            Begin VB.Label lblInsureInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽ����Ϣ��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   51
               Top             =   1080
               Width           =   1500
            End
            Begin VB.Label labPrice 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����---"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   315
               Left            =   30
               TabIndex        =   46
               Top             =   585
               Width           =   1200
            End
            Begin VB.Shape shpRequest 
               BorderStyle     =   0  'Transparent
               FillStyle       =   0  'Solid
               Height          =   60
               Left            =   15
               Top             =   435
               Width           =   13470
            End
            Begin VB.Label labRequestInfArea 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������Ϣ����"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   4860
               TabIndex        =   20
               Top             =   75
               Visible         =   0   'False
               Width           =   1800
            End
            Begin VB.Label labRequestDoct 
               BackStyle       =   0  'Transparent
               Caption         =   "�����ˣ�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4080
               TabIndex        =   19
               Top             =   570
               Width           =   1245
            End
            Begin VB.Label labRequestDoctValue 
               BackStyle       =   0  'Transparent
               Caption         =   "�ųԷ�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5130
               TabIndex        =   18
               Top             =   555
               Width           =   1365
            End
            Begin VB.Label labRequestRoom 
               BackStyle       =   0  'Transparent
               Caption         =   "������ң�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6465
               TabIndex        =   17
               Top             =   570
               Width           =   1530
            End
            Begin VB.Label labRequestRoomValue 
               BackStyle       =   0  'Transparent
               Caption         =   "���Ǻ����-"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   7800
               TabIndex        =   16
               Top             =   555
               Width           =   1755
            End
            Begin VB.Label labRequestTime 
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ�䣺"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   9525
               TabIndex        =   15
               Top             =   570
               Width           =   1575
            End
            Begin VB.Label labRequestTimeValue 
               BackStyle       =   0  'Transparent
               Caption         =   "2013-03-06 17:50"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10950
               TabIndex        =   14
               Top             =   555
               Width           =   2565
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   50
      Top             =   10500
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2381
            MinWidth        =   882
            Picture         =   "frmPacsApplication.frx":2BEF
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13600
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPacsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const M_STR_FIXEDELEMENT_DIAGNOSE As String = "������"

'�˵�id����
Private Enum TApplicationMenu
    amAppFile = 0
    amAppPreview = 1
    amAppPrint = 2
    
    amAppEdit = 3
    amAppSave = 4
    amAppDel = 5
    amAppExit = 6
    
    amAppKind = 7
    amAppSend = 8
    
    amAppType = 9
End Enum


'��Ŀ��������
Private Enum TProjectCol
    pcId = 0                'Id��
    pcName = 1              '��Ŀ������
    pcRoomType = 2          '����������
    pcMethod = 3            '���ַ��������
    pcExeRoom = 4           'ִ�п�����
    pcNormalCol = 5            '������
    pcBedCol = 6               '������
    pcOperCol = 7              '������
End Enum


'������Դ����
Private Enum TPatientFrom
    pfOutPatient = 1    '���ﲡ��
    pfInPatient = 2     'סԺ����
End Enum


Private mrsRequestPart As ADODB.Recordset       '��������뵥�����м�鲿λ������ѡ����Ŀʱ�������ݿ��ж�ȡ
          

'������Ϣ
Private Type TPatientInf
    lngID As Long                       '��ǰ����ID�����ⲿ����
    lngFrom As Long                     '������Դ�����ⲿ���� ,1��ʾ���2��ʾסԺ
    lngInsure As Long                   '�������࣬���ڶ�ҽ����Ŀ���м��ʱ����
        
    lngPageId As Long                   '��ҳId��סԺ���˲�����ҳId
    lngRegId  As Long                   '�Һ�id�����ﲡ�˲��йҺ�Id
    
    strConditionTag As String           '������ �硰Σ���� ��������

    strName As String                   '��������
    strSex As String                    '�����Ա�
    strRegNo As String                  '�Һŵ���
    strRegDate As String                '�Ǽ�ʱ��
    lngRoomId As Long                   '���˿���ID
    strAge As String                    '����
    strInNO As String                   'סԺ��
    strOutNo As String                  '�����
    strInHospitalDate As String         '��Ժʱ��
End Type


'ҽ����Ϣ
Private Type TDoctorInf
    lngID As Long
    str�û��� As String
    str���� As String
    lng����ID As Long
End Type


Private Type TRECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Type TPoint
    X As Long
    Y As Long
End Type

Private mCurPatientInf As TPatientInf
Private mCurDoctorInf As TDoctorInf

Public mobjAppDatas As New Scripting.Dictionary         '��������ҳ�������

Private mblnPageUpdateState As Boolean
Private mblnIsLoadRequestPage As Boolean        '�Ƿ��������ҳ
Public mblnIsSaveRequestPage As Boolean        '�Ƿ񱣴�������ҳ

Private mintBabyID As Integer                   'Ӥ�����
Private mlngCurDeptId As Long                   '��ǰ����Id
Private mlngUpdateAppNo As Long                 '��Ҫ���µ�ҽ���������
Private mstrDoctorName As String
Private mlngProjectId As Long                   '��Ҫ�Զ���λ��������Ŀ����ĿID
Private mlngRequestPageCount As Long
Private mblnIsRestoreTab As Boolean

Private mfrmPacsApplyWord As frmPacsApplyWord

Private mstrRequestAffixConfig As String        '���븽�����ã�������1:����,����,Ҫ��Id|������2:����,����,Ҫ��Id|������n:����,����,Ҫ��Id

Private mblnShowWord  As Boolean        '��ť�Ƿ���ʾ�ʾ���棬True-�ʾ���棬False-ģ�����

Private mobjLastControl As Object
Private mobjOwner As Object
Private mclsDiagEdit As zlMedRecPage.clsDiagEdit
Private mobjEmrInterface As Object           '�°没�����븽���ȡ����

'Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As TRECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As TPoint)
Private Declare Sub OffsetRect Lib "user32" (lpRect As TRECT, ByVal X As Long, ByVal Y As Long)


Public Sub InitComponents(ByVal lngDeptID As Long, objOwner As Object)
    mlngCurDeptId = lngDeptID
    Set mobjOwner = objOwner
    
    Call GetUserInfo
End Sub


Public Sub GetUserInfo()
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    If Not rsTmp.EOF Then
        mCurDoctorInf.lngID = rsTmp!ID
        mCurDoctorInf.lng����ID = IIF(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        mCurDoctorInf.str���� = IIF(IsNull(rsTmp!����), "", rsTmp!����)
        mCurDoctorInf.str�û��� = IIF(IsNull(rsTmp!�û���), "", rsTmp!�û���)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function ShowApplicationForm(ByVal lngPatientID As Long, _
                                    ByVal lngCallFrom As Long, _
                                    ByVal lngPatientRegId As Long, _
                                    ByVal lngPatientPageId As Long, _
                                    ByVal lngUpdateAppNoOrAdvId As Long, _
                                    ByRef objAppPages() As clsApplicationData, _
                                    Optional ByVal intBabyID As Integer = 0, _
                                    Optional ByVal blnEdit As Boolean = True, _
                                    Optional ByVal lngProjectId As Long = 0) As Boolean
    Dim curItem
    Dim i As Long
    
    ShowApplicationForm = False
    
    If lngPatientID <= 0 Then
        MsgBox "��ǰ����ID��Ч��IDֵΪ [" & lngPatientID & "]�������´������롣", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    If mlngCurDeptId <= 0 Then
        MsgBox "��ǰ����ID��Ч��IDֵΪ [" & mlngCurDeptId & "]�������´������롣", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    '�ͷ����뵥����
    Set mobjAppDatas = New Scripting.Dictionary
    
    'objAppPages��ֵʱ��Ϊ�޸����룬��ʱlngUpdateAppNoOrAdvId>0
    If SafeArrayGetDim(objAppPages) > 0 Then
        For i = 0 To UBound(objAppPages)
            mobjAppDatas.Add "_" & objAppPages(i).lngApplicationPageId & "_" & i, objAppPages(i)
        Next
    End If

    mblnPageUpdateState = lngUpdateAppNoOrAdvId = 1
    mblnIsLoadRequestPage = True
    mblnIsSaveRequestPage = False
    mlngProjectId = lngProjectId
    
    mCurPatientInf.lngFrom = lngCallFrom
    mCurPatientInf.lngID = lngPatientID
    mCurPatientInf.lngRegId = lngPatientRegId
    mCurPatientInf.lngPageId = lngPatientPageId
    mintBabyID = intBabyID
    
    mlngUpdateAppNo = lngUpdateAppNoOrAdvId
    
    picBaseInf.Enabled = blnEdit = True
    picInput.Enabled = blnEdit = True
    picPart.Enabled = blnEdit = True
    txtCurStudyProject.Enabled = blnEdit = True
    
    Me.Show 1, mobjOwner
    
    If Me.mobjAppDatas.Count > 0 Then
        ReDim objAppPages(Me.mobjAppDatas.Count - 1)
        
        i = 0
        For Each curItem In Me.mobjAppDatas.Items
            If Not curItem Is Nothing Then
                Set objAppPages(i) = curItem
                i = i + 1
            End If
        Next
    End If
    
    '�ͷ�����
    For i = mobjAppDatas.Count - 1 To 0 Step -1
        Call mobjAppDatas.Remove(mobjAppDatas.Keys(i))
        Set mobjAppDatas.Item(i) = Nothing
    Next i
    
    ShowApplicationForm = Me.mblnIsSaveRequestPage
    
End Function

Private Function GetCurRequestPageFormat(ByVal lngTabIndex As Long) As clsApplicationData
'��ȡ��ǰ����ҳ��ı����ʽ

    Dim objAppData As New clsApplicationData
    Dim strAffix As String
    Dim lngExeType As Long
    Dim objText As TextBox
    Dim strExeRoomInf As String
    Dim lngProjectRowIndex As Long
    Dim strProName As String
    Dim strElement As String
    Dim lngRequestPageIndex As Long
        
    Set GetCurRequestPageFormat = Nothing
    
    strAffix = ""
    
    For Each objText In rtbInputPro
        If objText.Index > 0 And objText.Visible Then
            strProName = GetInputProName(objText.Tag)
            strElement = GetInputProElement(objText.Tag)
            
            If strAffix <> "" Then strAffix = strAffix & "|"
            strAffix = strAffix & strProName & ":" & objText.Text
            
            If strElement = M_STR_FIXEDELEMENT_DIAGNOSE Then
                '�����ٴ���ϵ����Id
                objAppData.strDiagnoseId = objText.ToolTipText
            End If
        End If
    Next
    
    lngRequestPageIndex = tbcPage.Selected.Index
    lngProjectRowIndex = Val(vfgRequestProject.Tag)
    
    lngExeType = 0    'ִ������
    If Val(vfgRequestProject.Cell(flexcpData, lngProjectRowIndex, TProjectCol.pcBedCol)) = 1 Then
        lngExeType = 1
    ElseIf Val(vfgRequestProject.Cell(flexcpData, lngProjectRowIndex, TProjectCol.pcOperCol)) = 1 Then
        lngExeType = 2
    End If

    strExeRoomInf = vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcExeRoom)
    
    objAppData.strApplicationPageName = tbcPage.Item(lngRequestPageIndex).Caption                '���뵥����
    objAppData.lngApplicationPageId = GetRequestId(tbcPage.Item(lngRequestPageIndex).Tag)  '���뵥Id
    objAppData.strRequestTime = zlDatabase.Currentdate                                      '����ʱ��
    objAppData.strRequestAffixCfg = mstrRequestAffixConfig
    objAppData.blnIsModify = True
    
    objAppData.blnIsPriority = IIF(chkPriority.value <> 0, True, False)                 '�Ƿ����
    objAppData.lngProjectId = Val(vfgRequestProject.Cell(flexcpData, lngProjectRowIndex, TProjectCol.pcId))  '������ĿId
    objAppData.lngExeType = lngExeType                                                  'ִ������
    objAppData.strStartExeTime = dtpExeTime.value                                       'ִ��ʱ��
    objAppData.lngExeRoomId = Val(strExeRoomInf)                                        'ִ�п���Id
    objAppData.strExeRoomName = Replace(strExeRoomInf, Val(strExeRoomInf) & "-", "") 'ִ�п�������
    objAppData.lngExeRoomType = Val(vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcRoomType))
    objAppData.strPartMethod = Trim(vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcMethod))   '��λ����
    objAppData.strRequestAffix = strAffix                                               '���븽��(10)
    objAppData.lngRequestRoomId = mlngCurDeptId                                         '�������(11)
    objAppData.strRequestDoctor = UserInfo.����
    
    If mCurPatientInf.lngFrom = TPatientFrom.pfInPatient Then                           'ֻ��סԺҽ���Ŵ��ڲ�¼���
        objAppData.blnIsAdditionalRec = IIF(DateDiff("n", dtpExeTime.value, zlDatabase.Currentdate) >= gint��¼���, True, False)  '�Ƿ�¼ҽ��
    Else
        objAppData.blnIsAdditionalRec = False
    End If
    
    Set GetCurRequestPageFormat = objAppData
End Function

Private Function VerificationDataIsRight() As String
'��֤�����Ƿ���ȷ
    Dim objRich As TextBox
    Dim i As Long
    Dim aryPart() As String
    Dim lngProjectRowIndex As Long
    Dim strMethod As String
    
    VerificationDataIsRight = ""
    
    lngProjectRowIndex = Val(vfgRequestProject.Tag)
    
    '�жϱ�¼�ֶ���ȷ��
    For Each objRich In rtbInputPro
        If Not objRich Is Nothing Then
            '�жϸ������Ƿ�Ϊ��¼�����Ϊ��¼�������¼�����ݺ��������
            If labMustPro(objRich.Index).Visible Then
                If Trim(objRich.Text) = "" Then
                    VerificationDataIsRight = "��" & objRich.Tag & "������Ϊ�ա�"
                    objRich.SetFocus
                    
                    Exit Function
                End If
            End If
        End If
    Next
    
    
    '�жϼ�鲿λ��ȷ��
    If lngProjectRowIndex < 0 Then
        VerificationDataIsRight = "��ѡ�����������Ŀ��"
        vfgRequestProject.SetFocus
        
        Exit Function
    End If
    
    '�жϲ�λ������ȷ��
    If vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcMethod) = "" Then
        VerificationDataIsRight = "�����ü����Ŀ��" & vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcName) & "����Ӧ�ļ�鲿λ�ͷ�����"
        If vfgRequestProject.Row <> lngProjectRowIndex Then vfgRequestProject.Row = lngProjectRowIndex
        vfgList.SetFocus
        
        Exit Function
    End If
    
    '��֤��ѡ��λ�Ƿ������˼�鷽��
    strMethod = vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcMethod)
    
    '��ѡ���˼�鲿λ�������ж��Ƿ����ö�Ӧ�ļ�鷽�������δѡ����������б���
    If Trim(strMethod) <> "" Then
        aryPart = Split(vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcMethod), "|")
        For i = 0 To UBound(aryPart)
            If Split(aryPart(i), ";")(1) = "" Then
                VerificationDataIsRight = "��ԡ�" & vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcName) & "����Ŀ��" & _
                                            vfgList.TextMatrix(0, 1) & " ��" & Split(aryPart(i), ";")(0) & "����Ӧ��" & vfgList.TextMatrix(0, 2) & "����������"
                If vfgRequestProject.Row <> lngProjectRowIndex Then vfgRequestProject.Row = lngProjectRowIndex
                Call vfgList.SetFocus
                
                Exit Function
            End If
        Next i
    End If
    
    '�ж��Ƿ�����ִ�п���
    If vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcExeRoom) = "" Then
        VerificationDataIsRight = "�����ü����Ŀ��" & vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcName) & "����Ӧ�ļ����ҡ�"
        If vfgRequestProject.Row <> lngProjectRowIndex Then vfgRequestProject.Row = lngProjectRowIndex
        
        cbxExeRoom.SetFocus
        
        Exit Function
    End If
    
End Function


Private Sub ReadRequestPatientInf()
'�������뵥�Ļ�����Ϣ�粡�������Ա�,�������,�����˵�
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngDate As Long

    If mCurPatientInf.lngRegId <> 0 Then '��ѯ����
        strSQL = "select a.����,a.�Ա�,a.����,a.����,a.�����,a.סԺ��," & _
                " Nvl(Nvl(b.�������ID,Decode(b.ת��״̬,1,b.ת�����ID,NULL)),b.ִ�в���ID) as ���˿���ID,b.Id as �Һ�Id, b.No as �Һŵ�,b.�Ǽ�ʱ��,b.���� " & _
                " from ������Ϣ a, ���˹Һż�¼ b " & _
                " where a.����Id=b.����ID and a.����Id=[1] and  b.id=[2]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���뵥������Ϣ", mCurPatientInf.lngID, mCurPatientInf.lngRegId)
    Else    '��ѯסԺ
        strSQL = "select a.����,a.�Ա�,a.����,a.����,a.�����,a.סԺ��,a.��Ժʱ��, a.��ǰ����ID as ���˿���Id,b.��ҳId, b.��������,b.��ǰ���� from ������Ϣ a, ������ҳ b " & _
                " where a.����Id=b.����ID and a.����Id=[1] and b.��ҳId=[2] "
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���뵥������Ϣ", mCurPatientInf.lngID, mCurPatientInf.lngPageId)
    End If
    
    Call ClearRequestPatientInf
    
    If rsData.RecordCount > 0 Then
    
        mCurPatientInf.lngRoomId = Val(NVL(rsData!���˿���id))
        mCurPatientInf.strName = NVL(rsData!����)
        mCurPatientInf.strSex = NVL(rsData!�Ա�)
        mCurPatientInf.strAge = NVL(rsData!����)
        mCurPatientInf.lngInsure = Val(NVL(rsData!����))
        mCurPatientInf.strInNO = NVL(rsData!סԺ��)
        mCurPatientInf.strOutNo = NVL(rsData!�����)
        
        If mCurPatientInf.lngInsure <= 0 Then lblInsureInfo.Caption = "ҽ����Ϣ����ҽ������"
        
        If mCurPatientInf.lngRegId <> 0 Then
            mCurPatientInf.strRegNo = NVL(rsData!�Һŵ�)
            mCurPatientInf.strConditionTag = IIF(Val(NVL(rsData!����)) <> 0, "��!!", "")
            mCurPatientInf.strRegDate = NVL(rsData!�Ǽ�ʱ��)
        ElseIf mCurPatientInf.lngPageId <> 0 Then
            '�����סԺ���ˣ�����ݲ������ʴ����жϲ�����Դ
            mCurPatientInf.lngFrom = IIF(Val(NVL(rsData!��������)) = 1, 1, 2)
            mCurPatientInf.strConditionTag = Decode(Val(NVL(rsData!��ǰ����)), 9, "��!!", 10, "Σ!!", "")
            mCurPatientInf.strInHospitalDate = NVL(rsData!��Ժʱ��)
        Else
            '...
        End If
    End If
    
    If mintBabyID > 0 Then
        strSQL = "select Ӥ������,Ӥ���Ա�,����ʱ��,Round(Decode(����ʱ��,NULL,SysDate,����ʱ��)-����ʱ��) || '��' As Ӥ������ from ������������¼ where ����id=[1] and ��ҳid=[2] and ���=[3]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���뵥������Ϣ", mCurPatientInf.lngID, mCurPatientInf.lngPageId, mintBabyID)
         
        If rsData.RecordCount > 0 Then
            mCurPatientInf.strName = NVL(rsData!Ӥ������)
            mCurPatientInf.strSex = NVL(rsData!Ӥ���Ա�)
            mCurPatientInf.strAge = NVL(rsData!Ӥ������)
        End If
    End If
End Sub


Private Sub LoadRequestBaseInf()
'�������뵥���еĻ�����Ϣ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    picBaseInf.Visible = True
    
    labNameValue.Caption = mCurPatientInf.strName
    labSexValue.Caption = mCurPatientInf.strSex
    labAgeValue.Caption = mCurPatientInf.strAge
    labInPatientValue.Caption = mCurPatientInf.strInNO
    labOutPatientValue.Caption = mCurPatientInf.strOutNo
    labConditionTag.Caption = mCurPatientInf.strConditionTag
    
    If labRequestDoctValue.Caption = "" Then labRequestDoctValue.Caption = mCurDoctorInf.str����
    
    picRequestInf.Visible = True
    
    '��ȡ���������Ϣ
    strSQL = "select ���� from ���ű� where id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���������Ϣ", mlngCurDeptId)
    
    If rsData.RecordCount > 0 Then
        labRequestRoomValue.Caption = NVL(rsData!����)
    Else
        labRequestRoomValue.Caption = ""
    End If
    
    
End Sub

Private Function GetOrderInspectInfo(ByVal lng����ID As Long, ByVal strCondition As String, ByVal intType As Integer, ByVal lng����ID As Long) As String
'���ܣ���ȡָ�����˵�ָ������ڲ�����д����Ϣ�����磺���ߣ���ϵ�
    Dim strText As String
    On Error Resume Next
    If mobjEmrInterface Is Nothing Then
        Set mobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
    End If
    If Not mobjEmrInterface Is Nothing Then
        strText = mobjEmrInterface.GetOrderInspectInfoEx(intType, lng����ID, lng����ID, strCondition)
        If err.Number <> 0 Then
            strText = mobjEmrInterface.GetOrderInspectInfo(lng����ID, strCondition)
        End If
    End If
    GetOrderInspectInfo = strText
End Function

Public Function GetAppendItemValue(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, ByVal lngӤ�� As Long, ByVal str������ As String, ByVal str��Ŀ As String, ByVal lngҪ��ID As Long) As String
'���ܣ���ȡָ�������븽��ֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strText As String
    Dim intType As Integer
    Dim lng����ID As Long
     
    On Error GoTo errH
    
    If str�Һŵ� <> "" Then
        '�Ӳ���֮ǰ�ѱ����ҽ������ȡ,�������д��Ϊ׼
        strSQL = " Select ���� From (" & _
            " Select B.���� From ����ҽ����¼ A,����ҽ������ B" & _
            " Where A.ID=B.ҽ��ID And A.����ID=[1] And Nvl(A.Ӥ��,0)=[4]" & _
            IIF(str�Һŵ� <> "", " And A.�Һŵ�=[2]", " And Nvl(A.��ҳID,0)=[3]") & _
            " And B.��Ŀ=[5] And B.���� is Not Null and nvl(a.ҽ��״̬,0)<>4" & _
            " Order by A.����ʱ�� Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplication", lng����ID, str�Һŵ�, lng��ҳID, lngӤ��, str��Ŀ)
        If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
    End If
    
    'δ��ȡ��ֵ������ж�ӦҪ�أ���Ҫ����ȡ������ȡ
    If lngҪ��ID <> 0 And strText = "" Then
        '���ϰ棬���°�
        If str�Һŵ� <> "" Then '����
            strSQL = "Select Zl_Replace_Element_Value(B.������,[1],A.ID,1) as ����" & _
                " From ���˹Һż�¼ A,����������Ŀ B Where A.NO=[2] And B.ID=[3] And a.��¼����=1 And a.��¼״̬=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, str�Һŵ�, lngҪ��ID)
        Else
            strSQL = "Select Zl_Replace_Element_Value(������,[1],[2],2) as ���� From ����������Ŀ Where ID=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, lngҪ��ID)
        End If
        If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
        If strText = "" Then
            If str�Һŵ� <> "" Then
                strSQL = "select a.id From ���˹Һż�¼ A Where A.NO=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�Һŵ�)
                lng����ID = Val(rsTmp!ID & "")
                intType = 1
            Else
                lng����ID = lng��ҳID
                intType = 2
            End If
            strText = GetOrderInspectInfo(lng����ID, str������, intType, lng����ID)
        End If
    End If
    
    If str�Һŵ� = "" And strText = "" Then
        '�Ӳ���֮ǰ�ѱ����ҽ������ȡ,�������д��Ϊ׼
        strSQL = " Select ���� From (" & _
            " Select B.���� From ����ҽ����¼ A,����ҽ������ B" & _
            " Where A.ID=B.ҽ��ID And A.����ID=[1] And Nvl(A.Ӥ��,0)=[4]" & _
            IIF(str�Һŵ� <> "", " And A.�Һŵ�=[2]", " And Nvl(A.��ҳID,0)=[3]") & _
            " And B.��Ŀ=[5] And B.���� is Not Null and nvl(a.ҽ��״̬,0)<>4" & _
            " Order by A.����ʱ�� Desc) Where Rownum=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplication", lng����ID, str�Һŵ�, lng��ҳID, lngӤ��, str��Ŀ)
        If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
    End If
    
    GetAppendItemValue = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'ִ�в˵��¼�
On Error GoTo errHandle
    'ת�ƽ��㣬ʹ��dtpicker��change�¼�����������
    If tbcPage.Visible Then tbcPage.SetFocus
    
    Select Case Control.ID
        Case TApplicationMenu.amAppSave
            If Not SaveRequest(tbcRequest.Selected.Index) Then
                Exit Sub
            End If
            
        Case TApplicationMenu.amAppDel
            Call CancelRequest(tbcRequest.Selected.Index)
            
        Case TApplicationMenu.amAppType * 100 To TApplicationMenu.amAppType * 100 + 99
            Call SwitchRequestPageType(Control.Category, IIF(Len(Control.Parent.Title) > 0, Control.Parent.Title & "-", "") & Control.Caption)
            
        Case TApplicationMenu.amAppEdit
            Unload Me
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetRequestItemCount(ByVal lngRequestId As Long) As Long
    Dim i As Long
    Dim lngCount As Long
    
    lngCount = 0
    For i = 0 To tbcRequest.ItemCount - 1
        If GetRequestId(tbcRequest.Item(i).Tag) = lngRequestId Then
            lngCount = lngCount + 1
        End If
    Next i
    
    GetRequestItemCount = lngCount
End Function

Private Sub CancelRequest(ByVal lngRequestTabIndex As Long)
'��������
    Dim blnIsSelect As Boolean
    Dim i As Long
    
    If MsgBox("����������������ݽ����ָܻ����Ƿ������", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '��mobjAppDatas���Ƴ����������
    If mobjAppDatas.Exists("_" & CStr(tbcRequest.Item(lngRequestTabIndex).Tag)) Then
        Call mobjAppDatas.Remove("_" & tbcRequest.Item(lngRequestTabIndex).Tag)
    End If
    
    mblnPageUpdateState = False
    
    If GetRequestItemCount(GetRequestId(tbcRequest.Item(lngRequestTabIndex).Tag)) > 1 Then
        If lngRequestTabIndex > 0 Then
            '��ǰ�ƶ�����ҳ��
            tbcRequest.Item(lngRequestTabIndex).Tag = ""    '���⳷��ʱ��ʾ����
            Call tbcRequest.RemoveItem(lngRequestTabIndex)
        ElseIf lngRequestTabIndex < tbcRequest.ItemCount - 1 Then
            '����ƶ�����ҳ��
            tbcRequest.Item(lngRequestTabIndex).Tag = ""    '���⳷��ʱ��ʾ����
            Call tbcRequest.RemoveItem(lngRequestTabIndex)
        End If
        
        blnIsSelect = False
        For i = lngRequestTabIndex - 1 To 0 Step -1
            If tbcRequest.Item(i).Visible = True Then
                tbcRequest.Item(i).Selected = True
                blnIsSelect = True
                Exit For
            End If
        Next i
        
        If blnIsSelect = False Then
            For i = lngRequestTabIndex + 1 To tbcRequest.ItemCount - 1
                If tbcRequest.Item(i).Visible = True Then
                    tbcRequest.Item(i).Selected = True
                    Exit For
                End If
            Next i
        End If
        
        tbcRequest.Tag = tbcRequest.Selected.Index & "-" & tbcRequest.Selected.Caption
        
    Else
        tbcRequest.Item(lngRequestTabIndex).Caption = "����Ŀ"
        Call LoadRequestPage(GetRequestId(tbcRequest.Item(lngRequestTabIndex).Tag), tbcRequest.Item(lngRequestTabIndex).Tag)
    End If
End Sub


Private Function SaveRequest(ByVal lngRequestTabIndex As Long) As Boolean
    Dim strVerification As String
    Dim lngCurAdviceId As Long
    
    Dim strDataKey As String
    Dim objOldAppData As clsApplicationData
    Dim objCurAppData As clsApplicationData
    
    Dim strMsg As String
    Dim strCheckContext As String   '��Ҫ����ļ����Ŀ
    Dim lngMsgResult As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strCurProName As String
    
    SaveRequest = False
    
    If lngRequestTabIndex < 0 Then Exit Function
    
    strVerification = VerificationDataIsRight()
    If strVerification <> "" Then
        MsgBox strVerification, vbOKOnly, Me.Caption
        Exit Function
    End If

    If mCurPatientInf.lngFrom = TPatientFrom.pfOutPatient Then  '���ﲡ��
        
        If Trim(mCurPatientInf.strRegDate) <> "" Then
            If dtpExeTime.value < CDate(mCurPatientInf.strRegDate) Then
                Call MsgBox("ִ��ʱ�䲻�������ڲ��˵ĹҺ�ʱ�䡣", vbInformation + vbOKOnly, Me.Caption)
                dtpExeTime.SetFocus
                
                Exit Function
            End If
        End If
                
    Else    'סԺ����
        
        If Trim(mCurPatientInf.strInHospitalDate) <> "" Then
            If dtpExeTime.value < CDate(mCurPatientInf.strInHospitalDate) Then
                Call MsgBox("ִ��ʱ�䲻�������ڲ��˵���Ժʱ�䡣", vbInformation + vbOKOnly, Me.Caption)
                dtpExeTime.SetFocus
                
                Exit Function
            End If
        End If
        
        '�ж��Ƿ�Ϊ��¼ҽ��
        If DateDiff("n", dtpExeTime.value, zlDatabase.Currentdate) >= gint��¼��� And gint��¼��� > 0 Then
            If MsgBox("���Ŀ�ʼִ��ʱ�����ڵ�ǰʱ�䣬���Զ����ղ�¼��ʽ���д����Ƿ������", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                dtpExeTime.SetFocus
                Exit Function
            End If
            
            If chkPriority.value <> 0 Then
                Call MsgBox("��Ϊ��¼ҽ�����д���ʱ��������ѡ������ҽ������", vbInformation + vbOKOnly, Me.Caption)
                chkPriority.SetFocus
                
                Exit Function
            End If
        End If
    End If
    
    strCurProName = vfgRequestProject.Cell(flexcpData, Val(vfgRequestProject.Tag), TProjectCol.pcName)
    
    '�жϼ����Ŀ�Ƿ��Ѿ�����
    If tbcRequest.ItemCount > 1 Then
        For i = 0 To tbcRequest.ItemCount - 1
            If tbcRequest.Item(i).Caption = strCurProName And i <> lngRequestTabIndex Then
                Call MsgBox("�����Ŀ [" & strCurProName & "] ���ڱ��������д��ڣ������ظ����롣", vbInformation + vbOKOnly, Me.Caption)
                Exit Function
            End If
        Next i
    End If
    
    
    strDataKey = tbcRequest.Item(lngRequestTabIndex).Tag
    
    Set objCurAppData = GetCurRequestPageFormat(lngRequestTabIndex)
    
    'ҽ����Ŀ����
    strCheckContext = objCurAppData.lngProjectId & ":" & objCurAppData.lngExeRoomId
    strMsg = CheckAdviceInsure(mCurPatientInf.lngInsure, True, mCurPatientInf.lngID, mCurPatientInf.lngFrom, _
                                "", strCheckContext, "��ǰ��Ŀ")
                                
    If strMsg <> "" Then
        If gintҽ������ = 1 Then
            lngMsgResult = MsgBox(strMsg & vbCrLf & vbCrLf & "Ҫ��������ҽ����", vbYesNo, gstrSysName)
            If lngMsgResult = vbNo Then Exit Function
        ElseIf gintҽ������ = 2 Then
            Call MsgBox(strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档", vbInformation, gstrSysName)
            Exit Function
        End If
    End If
                                        
    
    If mobjAppDatas.Exists("_" & strDataKey) Then
        Set objOldAppData = mobjAppDatas.Item("_" & strDataKey)
        
        objCurAppData.lngUpdateAdviceId = objOldAppData.lngUpdateAdviceId
        objCurAppData.lngUpdateAppNo = objOldAppData.lngUpdateAppNo
        objCurAppData.blnAllowUpdate = objOldAppData.blnAllowUpdate
        
        '����Ѿ���mobjAppDatas�������������ݣ���ִ��ɾ�������±���
        Set mobjAppDatas.Item("_" & strDataKey) = Nothing
        Call mobjAppDatas.Remove("_" & strDataKey)
    End If
    
    Call mobjAppDatas.Add("_" & strDataKey, objCurAppData)
    
    
    Call SetRequestPageState(False)
    
    tbcRequest.Item(lngRequestTabIndex).Caption = strCurProName
    
    mblnIsSaveRequestPage = True
    
    If mCurPatientInf.lngInsure > 0 Then
        strSQL = "Select b.�շ���Ŀid From ������ĿĿ¼ a, �����շѹ�ϵ b Where a.id = b.������Ŀid  And a.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", vfgRequestProject.Cell(flexcpData, Val(vfgRequestProject.Tag), TProjectCol.pcId))
        
        If rsTemp.RecordCount > 0 Then
            lblInsureInfo.Caption = "ҽ����Ϣ��" & gclsInsure.GetItemInfo(mCurPatientInf.lngInsure, mCurPatientInf.lngID, NVL(rsTemp!�շ���ĿID), "", 0, "", TProjectCol.pcId & "||" & mCurPatientInf.lngFrom)
        End If
    End If
    
    SaveRequest = True
End Function


Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Select Case Control.ID
        Case TApplicationMenu.amAppSave
            Control.Enabled = mblnPageUpdateState
            
        Case TApplicationMenu.amAppType * 100 + 1 To TApplicationMenu.amAppType * 100 + 99
            If Not tbcPage.Visible Or tbcPage.ItemCount <= 0 Then Exit Sub
            
            '�����Ӧ���뵥�Ѿ����棬������ͼ��
            Control.IconId = IIF(mobjAppDatas.Exists("_" & Control.Category), 3558, 1)
            
        Case TApplicationMenu.amAppDel
            If Not tbcPage.Visible Or tbcPage.ItemCount <= 0 Then
                Control.Enabled = False
                Exit Sub
            End If
            
            Control.Enabled = mobjAppDatas.Exists("_" & CStr(tbcRequest.Selected.Tag)) And mlngUpdateAppNo <= 0
    End Select
Exit Sub
errHandle:
End Sub

Private Sub cbxExeRoom_Click()
On Error GoTo errHandle
    Dim lngProjectRowIndex As Long
    Dim lngExeRoomId As Long
    
    '��mblnIsLoadRequestPageΪtrueʱ����ʾ���ڼ���ִ�п�������
    If cbxExeRoom.ListIndex < 0 Or mblnIsLoadRequestPage Then Exit Sub
    
    lngProjectRowIndex = Val(vfgRequestProject.Tag)
    
    If lngProjectRowIndex < 0 And vfgRequestProject.RowSel < 0 Then Exit Sub
    
    If lngProjectRowIndex <> vfgRequestProject.RowSel Then
        Call SelectRequestProject(vfgRequestProject.RowSel)
        lngProjectRowIndex = vfgRequestProject.RowSel
    End If
    
    lngExeRoomId = Val(vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcExeRoom))
    
    If lngExeRoomId <> cbxExeRoom.ItemData(cbxExeRoom.ListIndex) Then
        vfgRequestProject.Cell(flexcpText, lngProjectRowIndex, TProjectCol.pcExeRoom) = cbxExeRoom.ItemData(cbxExeRoom.ListIndex) & "-" & Replace(cbxExeRoom.Text, Val(cbxExeRoom.Text) & "-", "")
        
        Call SetRequestPageState(True)
    End If
    
    Call ShowArrangeState
    Call ShowStudyProjectToTxt
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SetRequestPageState(ByVal blnModifyState As Boolean)
'�������������޸�״̬
    mblnPageUpdateState = blnModifyState
End Sub

Private Sub cbxExeRoom_DropDown()
On Error GoTo errHandle
    Call SendMessage(cbxExeRoom.hwnd, &H160, 200, 0)
errHandle:
End Sub

Private Sub SwitchRequestPageType(ByVal lngRequestPageId As Long, ByVal strRequestPageName As String)
On Error GoTo errHandle
    Dim i As Long
    Dim objControlItem As XtremeSuiteControls.TabControlItem
    Dim objPage As XtremeSuiteControls.TabControlItem
    
    Dim lngFindIndex As Long
    Dim lngHintResult As Long
    Dim lngItemAdviceTag As String
    Dim blnRefresh As Boolean
    Dim strName As String
    Dim strRequestName As String
    Dim blnIsNew As Boolean
    Dim lngPageIndex As Long
    Dim blnIsChangeRequestType As Boolean
    
 
    'ȥ���˵���Ŀ�ݼ�
    strName = Mid(strRequestPageName, 1, InStr(strRequestPageName, "(&") - 1)

    If mblnPageUpdateState Then
        strRequestName = Replace(tbcRequest.Tag, Val(tbcRequest.Tag) & "-", "")
        '���Ѿ��ı�����뵥���б�����ʾ
        lngHintResult = MsgBox("��" & strRequestName & "�������Ѿ��ı䣬�Ƿ񱣴棿", vbYesNoCancel, Me.Caption)
        
        If lngHintResult = vbYes Then
            '�������뵥
            If Not SaveRequest(Val(tbcRequest.Tag)) Then
                '�������ʧ�ܣ����˳�ҳ���л�
                Exit Sub
            End If
            
            Call SetRequestPageState(False)
            
        ElseIf lngHintResult = vbNo Then
            Call SetRequestPageState(False)
            
            '�����ϱ༭���еĻ�������
            If Not mclsDiagEdit Is Nothing Then
                Call mclsDiagEdit.DeleteApplyDiag(Val(tbcRequest.Item(Val(tbcRequest.Tag)).Tag))
            End If
        Else
            Exit Sub
        End If
    End If
    
    lngPageIndex = -1
    lngFindIndex = -1
    
    blnIsNew = False
    blnRefresh = False
    
    If tbcRequest.Selected Is Nothing Then
        blnIsChangeRequestType = True
    Else
        blnIsChangeRequestType = GetRequestId(tbcRequest.Selected.Tag) <> lngRequestPageId
    End If
    
    '��������ҳ������
    For i = 0 To tbcPage.ItemCount - 1
        If tbcPage.Item(i).Tag = lngRequestPageId Then
            lngPageIndex = i
            Exit For
        End If
    Next i
    
        '����tabҳ��ʾ
    If lngPageIndex >= 0 Then
       '�л�����������ҳ��
        If Not tbcPage.Selected Is Nothing Then
            If tbcPage.Selected.Index <> lngPageIndex Then
                tbcPage.Item(lngPageIndex).Selected = True
                Call tbcPage_SelectedChanged(tbcPage.Selected)
            End If
        End If
        
        
        For i = 0 To tbcRequest.ItemCount - 1
            If GetRequestId(tbcRequest.Item(i).Tag) = lngRequestPageId Then
                If mobjAppDatas.Exists("_" & CStr(tbcRequest.Item(i).Tag)) = False Then
                    '�����û�б�������룬��ֱ�ӽ����л�
                    blnIsNew = False
                    
                    If i <> tbcRequest.Selected.Index Then blnRefresh = True
                    
                    tbcRequest.Item(i).Selected = True
                    lngFindIndex = i
                    
                    Exit For
                End If
            End If
        Next i
    Else
        Set objPage = tbcPage.InsertItem(tbcPage.ItemCount, strName, picTab.hwnd, 0)
        
        objPage.Tag = CStr(lngRequestPageId)
        objPage.Selected = True
        
        Call tbcPage_SelectedChanged(objPage)
    End If
    
    '�������ҳ�����͸ı䣬����Ҫ������
    If blnIsChangeRequestType Then
        '�����ϱ༭���еĻ�������
        If Not mclsDiagEdit Is Nothing Then
            Call mclsDiagEdit.DeleteApplyDiag(GetRequestId(tbcRequest.Item(lngFindIndex).Tag))
        End If
    End If


    '���lngFindIndex ����-1��˵���Ѿ������˸�ҳ��
    If lngFindIndex < 0 Then
        '�����µ�tabҳ��
        Set objControlItem = tbcRequest.InsertItem(tbcRequest.ItemCount, "����Ŀ", picTab.hwnd, 0)
        
        '�����ظ����룬��tagֵΪ������ID_����ʱ�䡱
        objControlItem.Tag = CStr(lngRequestPageId) & "_" & Format(Now, "mmddhhmmss")
        objControlItem.Selected = True
        
        blnRefresh = True
    End If
    
    '��tab����Ϊһ��������ҳ�����͸ı�ʱ���������������뵥�������
    If tbcRequest.ItemCount = 1 Or blnRefresh Then
        Call tbcRequest_SelectedChanged(tbcRequest.Selected)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub chkPriority_Click()
On Error GoTo errHandle
    If mblnIsLoadRequestPage Then Exit Sub
    
    Call SetRequestPageState(True)
    
'    chkPriority.FontBold = IIF(chkPriority.value <> 0, True, False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdDiagnose_Click()
'����ϱ���
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim strPro As String
    Dim strCode As String
    Dim i As Long
    
    Call CloseDiagnoseCodeInput
    
    strPro = rtbInputPro(Val(cmdInput.Tag)).Text
    strCode = GetProCode(strPro, "1")
    
    '��ϱ���
    Set rsTemp = zlDatabase.ShowILLSelect(Me, "1", mCurPatientInf.lngRoomId, mCurPatientInf.strSex, True, False, strCode)
    
    If rsTemp Is Nothing Then
        Exit Sub
    End If
    
    If rsTemp.RecordCount <= 0 Then
        rtbInputPro(Val(cmdInput.Tag)).Text = ""
    Else
        i = 1
        strPro = ""
        
        While Not rsTemp.EOF
            strPro = strPro & i & "��" & NVL(rsTemp!����) & "  "
            i = i + 1
            
            Call rsTemp.MoveNext
        Wend
        
        rtbInputPro(Val(cmdInput.Tag)).Text = strPro
    End If

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExitCode_Click()
On Error GoTo errHandle
    Call CloseDiagnoseCodeInput
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function GetProCode(ByVal strProContext As String, ByVal strType As String) As String
'strProContext:��Ŀ��Ӧ��¼������
'strType:ѡ���������ͣ�D:��ʾ������������1:��ʾ��ϱ�����
'������Ŀ�Ķ�Ӧ����
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    Dim strPro As String
    Dim aryPro() As String
    Dim strCode As String
    
    strCode = ""
    strPro = strProContext
    
    '��ȡ�Ѿ�ѡ��ļ���������ϡ�,J02.901,M03.756,��
    If strPro <> "" Then
        For i = 1 To 9
            strPro = Replace(strPro, i & "��", "<#>")
        Next i
        
        aryPro() = Split(strPro, "<#>")
        
        strPro = ""
        For i = LBound(aryPro) To UBound(aryPro)
            If Trim(aryPro(i)) <> "" Then
                strPro = strPro & Trim(aryPro(i)) & ","
            End If
        Next i
        
        strSQL = "select ���� from " & IIF(strType = "D", "��������Ŀ¼", "�������Ŀ¼") & " a, " & _
                " (select column_value from Table(Cast(f_Str2List([1]) As zlTools.t_StrList))) b " & _
                " where a.����=b.column_value and ���=[2]"
                
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����", strPro, strType)
        
        While Not rsTemp.EOF
            strCode = strCode & "," & NVL(rsTemp!����)
            Call rsTemp.MoveNext
        Wend
        
        If strCode <> "" Then strCode = strCode & ","
    End If
    
    GetProCode = strCode
End Function


Private Sub cmdIcd10_Click()
'�򿪼�������
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim strPro As String
    Dim strCode As String
    Dim i As Long
    
    Call CloseDiagnoseCodeInput
    
    strPro = rtbInputPro(Val(cmdInput.Tag)).Text
    strCode = GetProCode(strPro, "D")
    
    'D-ICD-10��������
    Set rsTemp = zlDatabase.ShowILLSelect(Me, "D", mCurPatientInf.lngRoomId, mCurPatientInf.strSex, True, True, strCode)
    
    If rsTemp Is Nothing Then
        Exit Sub
    End If
    
    If rsTemp.RecordCount <= 0 Then
        rtbInputPro(Val(cmdInput.Tag)).Text = ""
    Else
        i = 1
        strPro = ""
        
        While Not rsTemp.EOF
            strPro = strPro & i & "��" & NVL(rsTemp!����) & "  "
            i = i + 1
            
            Call rsTemp.MoveNext
        Wend
        
        rtbInputPro(Val(cmdInput.Tag)).Text = strPro
    End If

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetRequestId(ByVal strTag As String) As Long
'��ȡ����ID
    GetRequestId = Val(Split(strTag & "_")(0))
End Function

Private Function OpenDiagnoseEdit(ByRef str���Id As String, ByRef str������� As String) As Boolean
'����ϱ༭��
    
    Dim lngApplicationPageId As Long
    Dim lngAdviceID As Long
    Dim objAppData As clsApplicationData
    
    OpenDiagnoseEdit = False
    
    If mclsDiagEdit Is Nothing Then
        Set mclsDiagEdit = New zlMedRecPage.clsDiagEdit
        Call mclsDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mCurPatientInf.lngFrom = TPatientFrom.pfOutPatient, 1260, 1261))
    End If
    
    lngApplicationPageId = GetRequestId(tbcRequest.Selected.Tag)
    
    '�жϵ�ǰ���뵥�Ƿ���mobjAppDatas�д���
    If mobjAppDatas.Exists("_" & tbcRequest.Selected.Tag) Then
        Set objAppData = mobjAppDatas("_" & tbcRequest.Selected.Tag)
        
        str���Id = objAppData.strDiagnoseId
        lngAdviceID = objAppData.lngUpdateAdviceId
    Else
        lngAdviceID = 0
    End If
    
    '��ʾ��ϱ༭������
    OpenDiagnoseEdit = mclsDiagEdit.ShowDiagEdit(Me, _
                                    lngApplicationPageId, _
                                    mCurPatientInf.lngID, _
                                    IIF(mCurPatientInf.lngFrom = 2, mCurPatientInf.lngPageId, mCurPatientInf.lngRegId), _
                                    mCurPatientInf.lngFrom, _
                                    mCurPatientInf.lngRoomId, _
                                    str���Id, _
                                    str�������, _
                                    1, _
                                    lngAdviceID)
                                    
End Function

Private Sub cmdInput_Click()
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim strPro As String
    Dim strElement As String
    Dim rct As TRECT
    Dim pointOffset As TPoint
    Dim str���IDs As String
    Dim str������� As String

    If mblnShowWord Then
        If mfrmPacsApplyWord Is Nothing Then
            Set mfrmPacsApplyWord = New frmPacsApplyWord
        End If
        rtbInputPro(Val(cmdInput.Tag)).SelText = rtbInputPro(Val(cmdInput.Tag)).SelText & mfrmPacsApplyWord.ShowPacsApplyWord(mlngCurDeptId, mCurDoctorInf.lngID, labInputPro(Val(cmdInput.Tag)).Caption, Me)
    Else
        strPro = GetInputProName(rtbInputPro(Val(cmdInput.Tag)).Tag)
        strElement = GetInputProElement(rtbInputPro(Val(cmdInput.Tag)).Tag)
        
        If strElement = M_STR_FIXEDELEMENT_DIAGNOSE Then
            '�����¼����
            If OpenDiagnoseEdit(str���IDs, str�������) Then
                rtbInputPro(Val(cmdInput.Tag)).Text = str�������
                rtbInputPro(Val(cmdInput.Tag)).ToolTipText = str���IDs
            End If
    
        Else
            strSQL = "select id, ģ�����,ģ������,ʹ�ô��� from ��������ģ�� where �����ļ�Id=" & GetRequestId(tbcRequest.Selected.Tag) & " and ���ݸ���='" & strPro & "' order by ʹ�ô��� desc "
            Set rsData = zlDatabase.ShowSelect(Me, strSQL, 0, strPro)
            
            '���û�б�ѡ����ֱ���˳�
            If rsData Is Nothing Then Exit Sub
            
            rtbInputPro(Val(cmdInput.Tag)).SelText = NVL(rsData!ģ������)
            
            strSQL = "zl_��������ģ��_Update(" & rsData!ID & ",'" & NVL(rsData!ģ�����) & "','" & NVL(rsData!ģ������) & "'," & Val(NVL(rsData!ʹ�ô���)) + 1 & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "����ģ��ʹ�ô���")
        End If
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
 

Private Sub dtpExeTime_Change()
    Dim lngProjectRowIndex As Long
    
    If mblnIsLoadRequestPage Then Exit Sub

'    lngProjectRowIndex = Val(vfgRequestProject.Tag)
'    If lngProjectRowIndex < 0 Then Exit Sub
'
'    If lngProjectRowIndex <> vfgRequestProject.RowSel Then Call SelectRequestProject(vfgRequestProject.RowSel)
    
    Call ShowArrangeState
    Call SetRequestPageState(True)
    
'    Debug.Print "dtpExeTime_Change:" & Time
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHandle
    Dim i As Long
    Dim objCon As Object
    
    If Not uspRequestPage.UCEnabled Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        
        Case vbKeyS
            If Shift <> 2 Then Exit Sub
            
            If mblnPageUpdateState And Not tbcRequest.Selected Is Nothing Then
                'ת�ƽ��㣬ʹ��dtpicker��change�¼�����������
                If tbcRequest.Visible Then tbcRequest.SetFocus
                
                '�������뵥
                Call SaveRequest(tbcRequest.Selected.Index)
            End If
        
        Case vbKeyLeft
            If Shift <> 2 Then Exit Sub
            If tbcRequest.Selected Is Nothing Then Exit Sub
            
            '�л�tabҳ��
            If tbcRequest.Selected.Index > 0 Then tbcRequest.Item(tbcRequest.Selected.Index - 1).Selected = True
        
        Case vbKeyRight
            If Shift <> 2 Then Exit Sub
            If tbcRequest.Selected Is Nothing Then Exit Sub
            
            '�л�tabҳ��
            If tbcRequest.Selected.Index < tbcRequest.ItemCount - 1 Then tbcRequest.Item(tbcRequest.Selected.Index + 1).Selected = True
            
        Case vbKeyReturn
'            zlCommFun.PressKey vbKeyTab
            If Shift <> 2 Then Exit Sub
            
            'ctrl+enter ִ�н����л�����
            Select Case UCase(TypeName(Screen.ActiveControl))
                Case "TEXTBOX"
                    If UCase(Screen.ActiveControl.Name) = "TXTFIND" Then
                        chkPriority.SetFocus
                    Else
                        For Each objCon In rtbInputPro
                            If objCon.Index >= Screen.ActiveControl.Index + 1 And objCon.Visible And objCon.TabStop Then
                                objCon.SetFocus
                                Exit Sub
                            End If
                        Next
                        
                        txtFind.SetFocus
                    End If
                    
                Case "VSFLEXGRID"
                    If UCase(Screen.ActiveControl.Name) = "VFGREQUESTPROJECT" Then
                        dtpExeTime.SetFocus
                    Else
                        If rtbInputPro.Count > 1 And picInput.Visible Then
                            rtbInputPro(1).SetFocus
                        Else
                            txtFind.SetFocus
                        End If
                    End If
                Case "DTPICKER"
                    cbxExeRoom.SetFocus
                    
                Case "COMBOBOX"
                    If UCase(Screen.ActiveControl.Name) = "CBXEXEROOM" Then
                        vfgList.SetFocus
                    Else
                        If rtbInputPro.Count > 1 And picInput.Visible Then
                            rtbInputPro(1).SetFocus
                        Else
                            txtFind.SetFocus
                        End If
                    End If
                    
                Case "CHECKBOX"
                    vfgRequestProject.SetFocus
                    
                Case Else
                    If rtbInputPro.Count > 1 And picInput.Visible Then
                        rtbInputPro(1).SetFocus
                    Else
                        txtFind.SetFocus
                    End If
            End Select
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Load()
    Dim cbrControl As CommandBarControl
    Dim i As Long

    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFaceElement
    Call FreeRequestInputControl
    
    Call ReadRequestPatientInf

    Call InitRequestTab
    Call InitCommandBars
    
    labRequestDoctValue.Caption = ""
    If mlngUpdateAppNo > 0 Then
        '����ҽ������ҳ��
        Call CfgUpdatePage(mlngUpdateAppNo)
        
        '�������뵥���еĻ�����Ϣ
        Call LoadRequestBaseInf
        
        Me.Caption = "������"
    Else
        '�����������Ͳ˵�
        If mlngProjectId <> 0 Then Set cbrControl = GetDefaultPage(mlngProjectId)
        If cbrControl Is Nothing Then
            Set cbrControl = cbrMain.FindControl(, TApplicationMenu.amAppType * 100 + 1, False, True)
        End If
        
        If cbrControl Is Nothing Then
            uspRequestPage.UCEnabled = False
        Else
            Call cbrMain_Execute(cbrControl)
            
            '�������뵥���еĻ�����Ϣ
            Call LoadRequestBaseInf
        End If
        
        Me.Caption = "�������"
    End If
    
    If mCurPatientInf.lngID <= 0 Then
        uspRequestPage.UCEnabled = False
    End If
End Sub

Private Function GetDefaultPage(ByVal lngProjectId As Long) As Object
'���ܣ���ȡʡҳ��
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngTmp As Long, i As Long
    Dim cbrControl As CommandBarControl
    
    strSQL = "Select Max(�����ļ�id) As PageId From ��������Ӧ�� Where ������Ŀid =[1] And Ӧ�ó���=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngProjectId, mCurPatientInf.lngFrom)
    
    lngTmp = Val(rsData!PageId & "")
 
    If lngTmp <> 0 Then
        For i = 1 To mlngRequestPageCount
            Set cbrControl = cbrMain.FindControl(, TApplicationMenu.amAppType * 100 + i, False, True)
            If Not cbrControl Is Nothing Then
                If Val(cbrControl.Category) = lngTmp Then
                    Exit For
                Else
                    Set cbrControl = Nothing
                End If
            End If
        Next
    End If
    Set GetDefaultPage = cbrControl
End Function

Private Function GetProjectName(ByVal lngProId As Long) As String
'��ȡ��Ŀ����
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetProjectName = ""
    
    strSQL = "Select ���� From ������ĿĿ¼ Where Id =[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��Ŀ����", lngProId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetProjectName = NVL(rsData!����)
End Function


Private Sub CfgUpdatePage(ByVal lngUpdateAppNo As Long)
'���ø���ҳ��
    Dim i As Integer
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strTbcTag As String
    Dim strInsertTags As String
    
    Dim objAppData As clsApplicationData
    Dim curTab As XtremeSuiteControls.TabControlItem
    Dim objCurPage As XtremeSuiteControls.TabControlItem
    
    Dim blnOldUpdateState As Boolean
    
    blnOldUpdateState = mblnPageUpdateState
    
    If mobjAppDatas.Count <= 0 Then Exit Sub
    
    mblnPageUpdateState = False
    
    strInsertTags = ""
    For i = 0 To mobjAppDatas.Count - 1
        Set objAppData = mobjAppDatas(mobjAppDatas.Keys(i))
        If objAppData Is Nothing Then Exit Sub
        
        labRequestTimeValue.Caption = Format(IIF(Trim(objAppData.strRequestTime) = "", Now, objAppData.strRequestTime), "yyyy-mm-dd hh:mm")
        Set curTab = tbcRequest.InsertItem(0, GetProjectName(objAppData.lngProjectId), picTab.hwnd, 0)
        
        curTab.Tag = CStr(objAppData.lngApplicationPageId) & "_" & i
        
        If InStr(strInsertTags, objAppData.lngApplicationPageId) <= 0 Then
            strInsertTags = strInsertTags & "," & objAppData.lngApplicationPageId
            
            Set objCurPage = tbcPage.InsertItem(0, objAppData.strApplicationPageName, picTab.hwnd, 0)
            objCurPage.Tag = CStr(objAppData.lngApplicationPageId)
        End If
        
        If strTbcTag = "" Then strTbcTag = curTab.Index & "-" & objAppData.strApplicationPageName
    Next
        
    If tbcPage.ItemCount > 0 Then
        tbcPage.Item(0).Selected = True
        Call tbcPage_SelectedChanged(tbcPage.Selected)
        
        tbcRequest.Item(0).Selected = True
        Call tbcRequest_SelectedChanged(tbcRequest.Selected)
    End If
    
    mblnPageUpdateState = blnOldUpdateState
End Sub


Private Sub InitFaceElement()
'��ʼ�����뵥������ص�Ԫ��
    Dim dtNow As Date
    
    dtNow = zlDatabase.Currentdate
    
    vfgRequestProject.Tag = -1
    vfgRequestProject.Rows = 0
    vfgRequestProject.Cols = 8
    
    vfgRequestProject.ColHidden(TProjectCol.pcRoomType) = True
    vfgRequestProject.ColHidden(TProjectCol.pcMethod) = True
    vfgRequestProject.ColHidden(TProjectCol.pcExeRoom) = True
    
    labRequestTimeValue.Caption = Format(dtNow, "yyyy-mm-dd hh:mm")
    dtpExeTime.value = dtNow
    
    tbcRequest.Tag = -1     '���浱ǰѡ�е�ҳ�棬�Ա���tabҳ�л����¼��У��ܹ��ж���һ�ε�ѡ��ҳ�棬Ĭ�ϲ�ѡ���κ�ҳ��
    
    picInput.Tag = 0        'Ϊ0ʱ����ʾ������ʾ���븽���¼��
End Sub


Private Sub InitRequestTab()
    Dim objfont As New StdFont
    
    objfont.Name = "����"
    objfont.Size = 12
    
    With tbcPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPageFlat ' xtpTabAppearancePropertyPageSelected ' xtpTabAppearancePropertyPage2003 'xtpTabAppearancePropertyPageFlat
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .PaintManager.HotTracking = True
        .PaintManager.ButtonMargin.Top = 11
        .PaintManager.ButtonMargin.Bottom = 11
        
        Set .PaintManager.Font = objfont
    End With
    
    With tbcRequest
        .PaintManager.Appearance = xtpTabAppearancePropertyPageFlat ' xtpTabAppearancePropertyPageSelected ' xtpTabAppearancePropertyPage2003 'xtpTabAppearancePropertyPageFlat
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = False
        .PaintManager.HotTracking = True
        .PaintManager.ButtonMargin.Top = 4
        .PaintManager.ButtonMargin.Bottom = 2
        
        Set .PaintManager.Font = objfont
    End With
End Sub


Private Sub LoadRequestPageKind(mnuParent As Object)
'���������뵥���
    Dim objParentMenu As CommandBarControl
    Dim objMenuBar As CommandBarControl
    Dim objControl As CommandBarControl
    
    Dim blnExist As Boolean
    Dim lngCount As Long
    Dim lngID As Long
    Dim arrName() As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim objControlItem As TabControlItem
    Dim i As Long
    Dim j As Long
    
    Set objParentMenu = mnuParent
    
    strSQL = "Select a.Id, a.���� From �����ļ��б� A, ��������Ӧ�� B Where a.Id = b.�����ļ�id And b.Ӧ�ó��� = [1] And a.���� = 7 And a.���� = '���' Group By a.Id, a.����, a.��� Order By a.���"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������뵥", mCurPatientInf.lngFrom)

    
    If rsData.RecordCount <= 0 Then
        mnuParent.Enabled = False
        Exit Sub
    End If
    mlngRequestPageCount = rsData.RecordCount
    i = 1
    lngID = 1
    
    With objParentMenu.CommandBar
        While Not rsData.EOF
            blnExist = False
            lngCount = 1
            '��������а�����-��,��-��ǰ�沿��Ϊ������࣬���沿��Ϊ���뵥����
            If InStr(NVL(rsData!����), "-") = 0 Then
                Set objMenuBar = .Controls.Add(xtpControlButton, TApplicationMenu.amAppType * 100 + lngID, _
                                                NVL(rsData!����) & "����(&" & Chr(IIF(48 + i > 57, 56 + i, 48 + i)) & ")", "")
                
                objMenuBar.Category = NVL(rsData!ID)
                objMenuBar.IconId = 1
                lngID = lngID + 1
                i = i + 1
            Else
                arrName = Split(NVL(rsData!����), "-")
                '�����������Ƿ��Ѵ���
                For j = 1 To .Controls.Count
                    If NVL(arrName(0)) = Mid(.Controls(j).Caption, 1, InStr(.Controls(j).Caption, "(&") - 1) Then
                        Set objMenuBar = .Controls(j)
                        lngCount = objMenuBar.CommandBar.Controls.Count + 1
                        blnExist = True
                        Exit For
                    End If
                Next
                
                '������಻����ʱ�����ӷ���
                'Ϊ�˱�֤���뵥��ID����һ�£�xtpControlButton, TApplicationMenu.amAppType * 100�����������ID��ΪxtpControlButtonPopup, TApplicationMenu.amAppType * 1000
                If Not blnExist Then
                    Set objMenuBar = .Controls.Add(xtpControlButtonPopup, TApplicationMenu.amAppType * 1000 + i, NVL(arrName(0)) & "(&" & Chr(IIF(48 + i > 57, 56 + i, 48 + i)) & ")", "")
                    objMenuBar.IconId = 1
                    i = i + 1
                End If
                
                '��ʾ����
                Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, TApplicationMenu.amAppType * 100 + lngID, NVL(arrName(1)) & "����(&" & Chr(IIF(48 + lngCount > 57, 56 + lngCount, 48 + lngCount)) & ")", "")
                lngID = lngID + 1
                '��������˷��࣬��¼�������ƣ����ڱ�����ʾ
                If Not blnExist Then
                    objControl.Parent.Title = NVL(arrName(0))
                End If
                objControl.Category = NVL(rsData!ID)
                objControl.IconId = 1
            End If
            
            
            Call rsData.MoveNext
        Wend
    End With

End Sub



Private Sub LoadRequestAffixInputCfg(ByVal lngRequestPageId As Long, Optional ByVal strAffixs As String = "")
'�������뵥����¼��
    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strProName As String
    Dim lngProOrder As Long
    Dim rsRequestContext As ADODB.Recordset
    Dim strInputValue As String
    Dim arrAffix() As String
    Dim strElement As String
    Dim i As Long
    
    labTitle.Caption = tbcPage.Selected.Caption
    
    strSQL = "select a.Ҫ��id,a.�ļ�ID,a.��Ŀ,a.����,a.����,Ҫ��Id,b.������ as Ҫ����, a.����,a.ֻ��,b.������  " & _
            " from �������ݸ��� a, ����������Ŀ b  " & _
            " where a.Ҫ��id=b.id(+) and a.�ļ�Id=[1] order by ���� "
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ¼����Ŀ", lngRequestPageId)
        
    If rsData.RecordCount <= 0 Then
        picInput.Visible = False
        picInput.Tag = 0
        
        Exit Sub
    End If

    arrAffix() = Split(strAffixs, "|")
    mstrRequestAffixConfig = ""
    
    While Not rsData.EOF
        '����¼����Ŀ��������
        strProName = NVL(rsData!��Ŀ)
        lngProOrder = Val(NVL(rsData!����))
        
        strInputValue = ""
        
        If mstrRequestAffixConfig <> "" Then mstrRequestAffixConfig = mstrRequestAffixConfig & "|"
        mstrRequestAffixConfig = mstrRequestAffixConfig & strProName & ":" & Val(NVL(rsData!����)) & "," & lngProOrder & "," & NVL(rsData!Ҫ��ID) & ","
        
        
        If strAffixs <> "" Then
            For i = 0 To UBound(arrAffix)
                If Split(arrAffix(i), ":")(0) = strProName Then
                    strSQL = arrAffix(i)
                    strInputValue = Replace(strSQL, Mid(strSQL, 1, InStr(strSQL, ":")), "")
                    Exit For
                End If
            Next i
        End If
        
        If Trim(strInputValue) = "" Then
            strInputValue = GetAppendItemValue(mCurPatientInf.lngID, mCurPatientInf.lngPageId, mCurPatientInf.strRegNo, mintBabyID, NVL(rsData!������), strProName, NVL(rsData!Ҫ��ID, 0))
            
            If strInputValue = "" Then
                strInputValue = NVL(rsData!����)
            End If
        End If
 
        '��ȡҪ������
        strElement = NVL(rsData!Ҫ����)
        
        Call SetInputControl(strProName, strInputValue, lngProOrder, Val(NVL(rsData!����)), _
                            strElement, IIF(rsData.RecordCount <= 4, True, False), IIF(Val(NVL(rsData!ֻ��)) = 1, True, False))

        Call rsData.MoveNext
    Wend
    
    picInput.Height = mobjLastControl.Top + mobjLastControl.Height + 120
    picInput.Visible = True
    picInput.Tag = 1
    
    If rtbInputPro.Count > 1 Then rtbInputPro(1).TabIndex = 0
End Sub



Private Sub LoadRequestPageProject(ByVal lngRequestPageId As Long, _
    Optional ByVal lngProjectId As Long = 0, _
    Optional ByVal lngExeRoomId As Long = 0, _
    Optional ByVal strExeRoomName As String = "", _
    Optional ByVal lngExeType As Long = 0, _
    Optional ByVal strPartMethod As String = "")
    
'�������뵥��Ӧ�ļ����Ŀ
'��Ҫ����Ŀ��Χ���й��ˣ��粡��ΪסԺ����ʱ��ֻ��������סԺ�ļ����Ŀ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim i As Long
    Dim strUseDeptIds As String
    Dim blnShowNormal As Boolean
    Dim lngSelIndex As Long
    
    picPart.Visible = True
    
    strUseDeptIds = ",||" & mCurPatientInf.lngRoomId & "||,||" & mlngCurDeptId & "||,"
    
    strSQL = "select distinct a.Id,����,a.����, a.ִ�п���,a.ִ�б�� from ������ĿĿ¼ a, ��������Ӧ�� b " & _
            " where a.Id=b.������ĿId " & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) " & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " & _
                    " And A.������� IN(" & IIF(mCurPatientInf.lngFrom = 3, "1,2,4", mCurPatientInf.lngFrom) & ",3) " & _
                    " and b.Ӧ�ó���=[3] " & _
                    " And Nvl(A.����Ӧ��,0)=1" & _
                    " And Nvl(A.�����Ա�,0) IN (" & IIF(mCurPatientInf.strSex Like "*��*", "1,0)", "2,0)") & _
                    " And Nvl(A.ִ��Ƶ��,0) IN(0,1) " & _
                    " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And Instr([2],',||'||����ID||'||,')>0)" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))" & _
                    " And �����ļ�Id=[1] " & _
            " order by ����"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���Ƶ��ݶ�Ӧ�����Ŀ", lngRequestPageId, strUseDeptIds, mCurPatientInf.lngFrom)
        
    vfgRequestProject.Rows = 0
    vfgRequestProject.Tag = -1
    txtCurStudyProject.Text = ""
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    vfgRequestProject.Rows = rsData.RecordCount
  
    lngSelIndex = -1
    i = 0
    blnShowNormal = False
    
    While Not rsData.EOF
        '�¿����ʱ�����������ĿID=mlngProjectId,���Զ���λ������Ŀ���޸ļ��ʱ�����������ĿID=lngProjectId,���Զ���λ������Ŀ��
        If Val(NVL(rsData!ID)) = lngProjectId Or Val(NVL(rsData!ID)) = mlngProjectId Then
            lngSelIndex = i
        End If
        
        Set vfgRequestProject.Cell(flexcpPicture, i, TProjectCol.pcId) = img16.ListImages("o0").Picture
        
        
        vfgRequestProject.ColWidth(TProjectCol.pcId) = 260
        
        vfgRequestProject.Cell(flexcpData, i, TProjectCol.pcId) = NVL(rsData!ID)
        vfgRequestProject.Cell(flexcpData, i, TProjectCol.pcName) = NVL(rsData!����)
        
        vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcName) = NVL(rsData!����) & "-" & NVL(rsData!����)
        vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcRoomType) = Val(NVL(rsData!ִ�п���))
        
        If Val(NVL(rsData!ִ�б��)) <> 0 Then
            '���ӳ��棬���ԣ�����ѡ����
            blnShowNormal = True
            
            vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcNormalCol) = "��"
            Set vfgRequestProject.Cell(flexcpPicture, i, TProjectCol.pcNormalCol) = img16.ListImages(IIF(lngExeType = 0, "o1", "o0")).Picture
            
            vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcBedCol) = "��"
            Set vfgRequestProject.Cell(flexcpPicture, i, TProjectCol.pcBedCol) = img16.ListImages(IIF(lngExeType = 1, "o1", "o0")).Picture
            
            vfgRequestProject.Cell(flexcpText, i, TProjectCol.pcOperCol) = "��"
            Set vfgRequestProject.Cell(flexcpPicture, i, TProjectCol.pcOperCol) = img16.ListImages(IIF(lngExeType = 2, "o1", "o0")).Picture
        End If
    
        i = i + 1
        
        Call rsData.MoveNext
    Wend
    
    If blnShowNormal Then
        vfgRequestProject.ColWidth(TProjectCol.pcName) = 3400
        
        vfgRequestProject.ColWidth(TProjectCol.pcNormalCol) = 550
        vfgRequestProject.ColWidth(TProjectCol.pcBedCol) = 550
        vfgRequestProject.ColWidth(TProjectCol.pcOperCol) = 550
        
        vfgRequestProject.ColHidden(TProjectCol.pcNormalCol) = False
        vfgRequestProject.ColHidden(TProjectCol.pcBedCol) = False
        vfgRequestProject.ColHidden(TProjectCol.pcOperCol) = False
    Else
        vfgRequestProject.ColHidden(TProjectCol.pcNormalCol) = True
        vfgRequestProject.ColHidden(TProjectCol.pcBedCol) = True
        vfgRequestProject.ColHidden(TProjectCol.pcOperCol) = True
    End If
    
    If lngSelIndex >= 0 Then
        Call SelectRequestProject(lngSelIndex, True)
        
        vfgRequestProject.Cell(flexcpText, lngSelIndex, TProjectCol.pcMethod) = strPartMethod
        vfgRequestProject.Cell(flexcpText, lngSelIndex, TProjectCol.pcExeRoom) = lngExeRoomId & "-" & strExeRoomName
        
        Call vfgRequestProject.ShowCell(lngSelIndex, 0)
        Call vfgRequestProject.Select(lngSelIndex, 0)
    Else
        If vfgRequestProject.Rows > 0 Then
            Call vfgRequestProject.ShowCell(0, 0)
            Call vfgRequestProject.Select(0, 0)
        End If
    End If
    
    If vfgRequestProject.RowSel > 0 Then vfgRequestProject.TopRow = vfgRequestProject.RowSel
End Sub


Public Function Get����Ա����ID(ByVal int������� As Integer) As Long
'���ܣ�ȡ����Ա���������ָ������Ĳ��ţ�ȱʡ��������
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    
    If blnNew Then
        strSQL = "Select Distinct B.����ID,Nvl(B.ȱʡ,0) as ȱʡ,C.������� From ������Ա B,��������˵�� C" & _
            " Where B.��ԱID = [1] And B.����ID=C.����ID" & _
            " Order by ȱʡ Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", mCurDoctorInf.lngID)
    End If
    rsTmp.Filter = "������� = 3 or ������� = " & int�������
    
    If Not rsTmp.EOF Then
        Get����Ա����ID = rsTmp!����ID
    Else
        Get����Ա����ID = mCurDoctorInf.lng����ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadRequestProjectExeRoom(ByVal lngProjectId As Long, ByVal lngProjectRoomType As Long, _
    ByVal lngCurExeDeptId As Long)
'[1]��������Դ��[2]�����˿���Id��[3]��������ĿId��[4]����ǰִ�п���Id��Ĭ��ִ�п���ID
'[5]������Ա����ID�� [6]������Id��[7]��������ҳID��[8]����������ID
    Dim strSQL As String
    Dim lng����Ա����ID As Long
    Dim rsData As ADODB.Recordset
    Dim i As Long
    Dim lngDefExeRoomId As Long
    Dim bytDay As Byte
    Dim bln�ϰల�� As Boolean
    
'    strDefExeRoomId = Get����ִ�п���ID(mCurPatientInf.lngID, mCurPatientInf.lngPageID, "D", , lngProjectRoomType, mCurPatientInf.lngRoomId, mlngCurDeptId, 0)
    
    
    cbxExeRoom.Clear

    lngDefExeRoomId = 0
    
    Select Case lngProjectRoomType
        Case 0 '0-��ִ�еĶ���
            Exit Sub
        Case 5  '5-Ժ��ִ��
            'ʹ��API���ټ���,��Ȼ�����е���
            AddComboItem cbxExeRoom.hwnd, CB_ADDSTRING, 0, "-"
            SetComboData cbxExeRoom.hwnd, CB_SETITEMDATA, i - 1, -1
            
            Call Cbo.SetIndex(cbxExeRoom.hwnd, 0)

            Exit Sub
        Case 1 '1-�������ڿ���
            strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([2],[4]) Order by ����"
        Case 2 '2-�������ڲ���
            If mCurPatientInf.lngFrom = 1 Then
                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([2],[4]) Order by ����"
            Else
                strSQL = _
                    " Select A.ID,A.����,A.����,A.����" & _
                    " From ���ű� A,������ҳ B" & _
                    " Where A.ID=B.��ǰ����ID And B.����ID=[6] And B.��ҳID=[7]" & _
                    " Union " & _
                    " Select ID,����,����,���� From ���ű� Where ID=[4]" & _
                    " Order by ����"
            End If
        Case 3 '3-����Ա���ڿ���
            strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([5],[4]) Order by ����"
            lng����Ա����ID = Get����Ա����ID(mCurPatientInf.lngFrom)
        Case 4 '4-ָ������
'            strSql = _
'                " Select Distinct A.ID,A.����,A.����,A.����" & _
'                " From ���ű� A,����ִ�п��� B,��������˵�� C" & _
'                " Where A.ID=B.ִ�п���ID And B.������ĿID=[3] And A.ID=C.����ID" & _
'                " And C.������� IN([1],3) And (B.������Դ is NULL Or B.������Դ=[1])" & _
'                " And (B.��������ID is NULL Or B.��������ID=[2])" & _
'                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
'                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
'                " Union Select ID,����,����,���� From ���ű� Where ID=[4]" & _
'                " Order by ����"

                If lngProjectId = 0 Then
                    If mCurPatientInf.lngFrom = 1 Then  '1��ʾ����
                        lngDefExeRoomId = mCurPatientInf.lngRoomId
                        strSQL = "select ID,����,����,���� From ���ű� where id=[10]"
                        
                    ElseIf mCurPatientInf.lngFrom = 2 Then  '2��ʾסԺ
                        lngDefExeRoomId = GetPatiUnitID(mCurPatientInf.lngID, mCurPatientInf.lngPageId)
                        strSQL = "select ID,����,����,���� From ���ű� where id=[10]"
                    End If
                    
                Else
                    'pacs�����Ŀ�����������������Ҫ�������ж��ϰ�ʱ��
                    bln�ϰల�� = Check�ϰల��(False)
                    
                    If Not bln�ϰల�� Then
                        strSQL = "Select Distinct c.Id, c.����,c.����,c.����, Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID,Decode(A.������Դ,Null,2,1) as ����" & _
                            " From ����ִ�п��� A,��������˵�� B,���ű� C" & _
                            " Where A.ִ�п���ID=B.����ID And A.������ĿID=[3]" & _
                            " And B.������� IN([1],3) And (A.������Դ is NULL Or A.������Դ=[1])" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And A.ִ�п���ID=C.ID And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " Order by ����" 'Ĭ�Ͽ�������
                    Else
                        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
                        strSQL = _
                            " Select Distinct d.Id, d.����,d.����,d.����, Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID,Decode(A.������Դ,Null,2,1) as ����" & _
                            " From ����ִ�п��� A,���Ű��� B,��������˵�� C,���ű� D" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.����=[9]" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.��ʼʱ��,'HH24:MI:SS') and To_Char(B.��ֹʱ��,'HH24:MI:SS') " & _
                            " And A.ִ�п���ID=C.����ID And C.������� IN([1],3) And (A.������Դ is NULL Or A.������Դ=[1])" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And A.ִ�п���ID=D.ID And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & _
                            " And (D.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� is NULL)" & _
                            " And A.������ĿID=[3]" & _
                            " Order by ����"
                    End If
                End If
        Case 6 '6-���������ڿ���
            strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([8],[4]) Order by ����"
    End Select
    
        

    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯִ�п���", mCurPatientInf.lngFrom, mCurPatientInf.lngRoomId, _
                                        lngProjectId, 0, mCurDoctorInf.lng����ID, _
                                        mCurPatientInf.lngID, mCurPatientInf.lngPageId, mlngCurDeptId, bytDay, lngDefExeRoomId)
                                        
    If rsData.RecordCount > 0 Then lngDefExeRoomId = rsData!ID
    
    '4��ʾָ��ִ�п���
    If lngProjectRoomType = 4 Then
        If Not rsData.EOF Then
            lngDefExeRoomId = rsData!ִ�п���ID
            rsData.Filter = "��������ID=" & mCurPatientInf.lngRoomId
            
            If rsData.EOF Then rsData.Filter = "ִ�п���ID=" & mCurPatientInf.lngRoomId
            If rsData.EOF And mCurPatientInf.lngFrom = 2 Then rsData.Filter = "ִ�п���ID=" & GetPatiUnitID(mCurPatientInf.lngID, mCurPatientInf.lngPageId)
            
            If Not rsData.EOF Then lngDefExeRoomId = rsData!ִ�п���ID
            
        ElseIf gblnָ��ҽ������������ִ�� Then
            If mCurPatientInf.lngFrom = 1 Then         '1��ʾ����
                lngDefExeRoomId = mCurPatientInf.lngRoomId
                strSQL = "select ID,����,����,���� From ���ű� where id=[10]"
                
            ElseIf mCurPatientInf.lngFrom = 2 Then     '2��ʾסԺ
                lngDefExeRoomId = GetPatiUnitID(mCurPatientInf.lngID, mCurPatientInf.lngPageId)
                strSQL = "select ID,����,����,���� From ���ű� where id=[10]"
                
            End If
            
            '���»�ȡ������Ϣ
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯִ�п���", mCurPatientInf.lngFrom, mCurPatientInf.lngRoomId, _
                                                lngProjectId, 0, mCurDoctorInf.lng����ID, _
                                                mCurPatientInf.lngID, mCurPatientInf.lngPageId, mlngCurDeptId, bytDay, lngDefExeRoomId)
            If rsData.RecordCount > 0 Then lngDefExeRoomId = rsData!ID
            
        End If
    End If
    
    rsData.Filter = ""

    For i = 1 To rsData.RecordCount
        'ʹ��API���ټ���,��Ȼ�����е���
        AddComboItem cbxExeRoom.hwnd, CB_ADDSTRING, 0, rsData!���� & "-" & rsData!����
        SetComboData cbxExeRoom.hwnd, CB_SETITEMDATA, i - 1, CLng(rsData!ID)
        
        '����Ĭ�ϵļ�����
        If lngDefExeRoomId = rsData!ID Then
            Call Cbo.SetIndex(cbxExeRoom.hwnd, i - 1)
        End If

        '���õ�ǰ�Ѿ�ѡ��ļ�����
        If lngCurExeDeptId = rsData!ID Then
            Call Cbo.SetIndex(cbxExeRoom.hwnd, i - 1)
        End If

        rsData.MoveNext
    Next
    
    'ֻ��ִ�п���Ϊ1ʱ��������Ĭ�Ͽ��ң�������Ҫ���û��ֶ�ѡ��
    If cbxExeRoom.ListCount = 1 And cbxExeRoom.ListIndex < 0 Then cbxExeRoom.ListIndex = 0
End Sub

Private Sub ClearRequestPatientInf()
'������뵥��صĲ��˲�����Ϣ
    mCurPatientInf.lngRoomId = 0
    mCurPatientInf.strName = ""
    mCurPatientInf.strSex = ""
    mCurPatientInf.strRegNo = ""
    mCurPatientInf.strAge = ""
    mCurPatientInf.strInNO = ""
    mCurPatientInf.strOutNo = ""
    mCurPatientInf.strConditionTag = ""
End Sub


Private Sub RestoreRequestPageCfg()
'�ָ���ҽ����ص����뵥��������
    Dim objControl As Object
    
    Set mobjLastControl = Nothing
    
    '������뵥������Ϣ
    txtFind.Text = ""
    txtCurStudyProject.Text = ""
    
'    cmdInput.Visible = False
    
    labPrice.Caption = "����---"
    
    chkPriority.value = 0
    chkPriority.FontBold = False
    
    dtpExeTime.value = zlDatabase.Currentdate
    
    '�����Ѿ����ص�¼�����
    For Each objControl In labInputPro
        objControl.Visible = False
    Next
    
    For Each objControl In labMustPro
        objControl.Visible = False
    Next
    
    For Each objControl In rtbInputPro
        objControl.Visible = False
    Next
    
    For Each objControl In shpBackLine
        objControl.Visible = False
    Next
    
    For Each objControl In shpInputPro
        objControl.Visible = False
    Next
    
    'ɾ����λѡ��
    vfgList.Rows = 1
    
    'ɾ�������Ŀ
    vfgRequestProject.Rows = 0

    
    'ɾ������ѡ��
    cbxExeRoom.Clear
    
    picInput.Visible = False
    picInput.Tag = 0
End Sub


Private Sub SetInputControl(ByVal strProName As String, ByVal strDefaultContext As String, ByVal lngProOrder As Long, _
    Optional ByVal blnIsMustInput As Boolean, Optional ByVal strElementName As String = "", Optional blnIsBigDistance As Boolean = False, _
    Optional ByVal blnIsReadOnly As Boolean = False)
    
'�������뵥�����¼����Ŀ
'strProName��¼��������
'strDefaultContext��Ĭ��ֵ
'lngProOrder��¼��˳��
'blnIsMustInput���Ƿ��¼��
'strElementName��Ҫ������
'blnIsBigDistance��¼����Ŀ֮���Ƿ�ʹ�ýϴ�ļ������
'blnIsReadOnly����Ŀ�Ƿ�������б༭

    Dim objLab As Label
    Dim objMust As Label
    Dim objRchEdit As TextBox
    Dim objShp As Shape
    
    
    If InStr(strProName, "-") > 0 Then
        '����ָ���
        If Not HasControl("shpInputPro", lngProOrder) Then
            Load shpInputPro(lngProOrder)
        End If
        
        Set objShp = shpInputPro(lngProOrder)
        objShp.Visible = True
        
        If mobjLastControl Is Nothing Then
            objShp.Left = 120
            objShp.Top = 120
            
        Else
            objShp.Left = 120
            objShp.Top = mobjLastControl.Top + mobjLastControl.Height + 360
            
        End If
        
        Set mobjLastControl = objShp
    Else
        '����¼���ǩ
        If Not HasControl("labInputPro", lngProOrder) Then
            Load labInputPro(lngProOrder)
        End If
        
        Set objLab = labInputPro(lngProOrder)
        
        objLab.Caption = strProName
        objLab.Visible = True
        
        If mobjLastControl Is Nothing Then
            objLab.Left = 120
            objLab.Top = IIF(blnIsBigDistance, 360, 180)
            
        Else
            objLab.Left = 120
            objLab.Top = mobjLastControl.Top + mobjLastControl.Height + 120 + IIF(blnIsBigDistance, 360, 180)
            
        End If
        
        '���Ϊ������Ŀ�����������ı�Ǳ�ǩ

        If Not HasControl("labMustPro", lngProOrder) Then
            Load labMustPro(lngProOrder)
        End If
        
        Set objMust = labMustPro(lngProOrder)
        objMust.Left = objLab.Left + objLab.Width + 60   ' Fix(objLab.Width / 2) + objLab.Left - 60
        objMust.Top = objLab.Top ' objLab.Top + objLab.Height - 10
        objMust.Visible = blnIsMustInput
        
        
        
        '����¼���
        If Not HasControl("rtbInputPro", lngProOrder) Then
            Load rtbInputPro(lngProOrder)
        End If
        
        Set objRchEdit = rtbInputPro(lngProOrder)
        
        objRchEdit.Left = objLab.Left + objLab.Width + objMust.Width + 120
        objRchEdit.Top = objLab.Top - 30 ' objLab.Height + 60
        objRchEdit.Width = picInput.ScaleWidth - objRchEdit.Left - 60 - cmdInput.Width  '��ȥcmdInput�Ŀ�ȣ��ǿ��ǵ����ܴ���ʹ�ø���ģ����������
        objRchEdit.Text = strDefaultContext
        objRchEdit.Tag = strProName & IIF(strElementName <> "", ">[" & strElementName & "]", "")
        objRchEdit.Visible = True
        objRchEdit.TabStop = True
        objRchEdit.TabIndex = rtbInputPro.Count - 1
        
        objRchEdit.TabStop = Not (blnIsReadOnly Or strElementName = M_STR_FIXEDELEMENT_DIAGNOSE)
        
        If blnIsReadOnly Or strElementName = M_STR_FIXEDELEMENT_DIAGNOSE Then
            'ֻ��״̬������
            objRchEdit.Locked = True
            objRchEdit.BackColor = &HC0FFFF
        Else
            objRchEdit.Locked = False
            objRchEdit.BackColor = rtbInputPro(0).BackColor
        End If
        
        Call SetInputControlHight(objRchEdit)
        
        '���뱳����
        If Not HasControl("shpBackLine", lngProOrder) Then
            Load shpBackLine(lngProOrder)
        End If
        
        Set objShp = shpBackLine(lngProOrder)
        
        objShp.Left = objRchEdit.Left
        objShp.Top = objRchEdit.Top + objRchEdit.Height
        objShp.Width = objRchEdit.Width
        objShp.Visible = Not (blnIsReadOnly Or strElementName = M_STR_FIXEDELEMENT_DIAGNOSE)
        
        Set mobjLastControl = objRchEdit
    End If
End Sub


Private Function GetInputProName(ByVal strTag As String) As String
'��ȡ¼����Ŀ����Ŀ����
    Dim lngProIndex As Long
    
    GetInputProName = strTag
    
    lngProIndex = InStrRev(strTag, ">[")
    If lngProIndex <= 0 Then Exit Function
    
    GetInputProName = Mid(strTag, 1, lngProIndex - 1)
End Function

Private Function GetInputProElement(ByVal strTag As String) As String
'��ȡ¼����Ŀ��Ҫ������,strTagʾ��: �������>[[�ٴ����]]
    Dim lngProIndex As Long
    Dim strResult As String
    
    GetInputProElement = ""
    
    lngProIndex = InStr(strTag, ">[")
    If lngProIndex <= 0 Then Exit Function
    
    'ȥ���ַ�">["
    strResult = Mid(strTag, lngProIndex + 2, 255)
    
    'ȥ���ַ�"]"
    strResult = Mid(strResult, 1, Len(strResult) - 1)
    
    GetInputProElement = strResult
End Function


Private Function HasControl(ByVal strControlName As String, ByVal lngIndex As Long) As Boolean
On Error GoTo errHandle
    Dim objControl As Object
    Select Case strControlName
        Case "labInputPro"
            For Each objControl In labInputPro
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
            
        Case "labMustPro"
            For Each objControl In labMustPro
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
            
        Case "shpInputPro"
            For Each objControl In shpInputPro
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
            
        Case "rtbInputPro"
            For Each objControl In rtbInputPro
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
        Case "shpBackLine"
            For Each objControl In shpBackLine
                If Not objControl Is Nothing Then
                    If objControl.Index = lngIndex Then
                        HasControl = True
                        Exit Function
                    End If
                End If
            Next
            
    End Select
    
    HasControl = False
Exit Function
errHandle:
    HasControl = False
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo errHandle
    Dim strRequestName As String
    Dim lngHintResult As Long
    
    If Not mblnPageUpdateState Then Exit Sub
    
    strRequestName = Replace(tbcRequest.Tag, Val(tbcRequest.Tag) & "-", "")
    '���Ѿ��ı�����뵥���б�����ʾ
    lngHintResult = MsgBox("��" & strRequestName & "�������Ѿ��ı䣬�Ƿ񱣴棿", vbYesNoCancel, Me.Caption)
    
    If lngHintResult = vbYes Then
        '�������뵥
        If Not SaveRequest(Val(tbcRequest.Tag)) Then
            '�������ʧ�ܣ����˳�ҳ���л�
            Cancel = 1
            Exit Sub
        End If
        
        Call SetRequestPageState(False)
        
    ElseIf lngHintResult = vbNo Then
        Call SetRequestPageState(False)
        
    Else
        Cancel = 1
        Exit Sub
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
On Error GoTo errHandle
    picMenu.Left = 0
    picMenu.Top = 0
    picMenu.Width = ScaleWidth
    
    tbcPage.Left = 3520 '2580
    tbcPage.Top = 60
    tbcPage.Width = ScaleWidth - 3540 ' 2600
    
    tbcRequest.Left = 0
    tbcRequest.Top = picMenu.Height - 100
    tbcRequest.Width = ScaleWidth
    
    picBack.Left = 0
    picBack.Top = tbcRequest.Top + tbcRequest.Height
    picBack.Width = ScaleWidth
    picBack.Height = ScaleHeight - picMenu.Height - stbThis.Height - tbcRequest.Height + 100

Exit Sub
errHandle:
End Sub


Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    
    '���ò˵����͹��������
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True                                '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False                            '�����õĲ˵���������
        .UseFadedIcons = False                                  'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True                                 '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True                                '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True                                      '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24                               '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16                              '����Сͼ��ĳߴ�

        
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '���ÿؼ���ʾ���
        .EnableCustomization False                             '�Ƿ������Զ�������
        .ActiveMenuBar.Visible = False
        
        
        Set .Icons = zlCommFun.GetPubIcons                     '���ù�����ͼ��ؼ�
    End With

    '---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)

    cbrToolBar.ShowTextBelowIcons = True

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TApplicationMenu.amAppSave, "����", "���������� (Ctrl+S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TApplicationMenu.amAppDel, "ɾ��", "ɾ���������"): cbrControl.IconId = 4114: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButtonPopup, TApplicationMenu.amAppType, "�������먋", "ѡ�������뵥������"): cbrControl.IconId = 807
    
    '�������뵥���Ͳ˵�
    Call LoadRequestPageKind(cbrControl)
    If mlngUpdateAppNo > 0 Or mCurPatientInf.lngID <= 0 Then cbrControl.Enabled = False

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TApplicationMenu.amAppEdit, "�˳�", "�˳��������"): cbrControl.IconId = 2613
    cbrControl.BeginGroup = True
        
        
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
End Sub


Private Sub Form_Terminate()
    Set mobjAppDatas = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandle
'    Dim i As Long
    
    Call FreeRequestInputControl
    
'    '�����������ͷ����ݣ���Ϊ���ڹرպ󣬻���Ҫ���ʴ����Է������뵥����
'    For i = mobjAppDatas.Count - 1 To 0 Step -1
'        Call mobjAppDatas.Remove(mobjAppDatas.Keys(i))
'        Set mobjAppDatas.Item(i) = Nothing
'    Next i
    
    Call SaveWinState(Me, App.ProductName)
    
    Set mfrmPacsApplyWord = Nothing
Exit Sub
errHandle:
End Sub

Private Sub SetNullMethod(ByVal lngProjectRowIndex As Long)
'���ü����Ŀ�ķ���Ϊ�շ�������Ϊ�շ���ʱ��������������Ҫȷ��ѡ��λ�ͼ�鷽��
    Dim lngRow As Long
    
    lngRow = lngProjectRowIndex
    If lngRow < 0 Then
        lngRow = vfgRequestProject.RowSel
    End If
    
    If lngRow >= 0 Then vfgRequestProject.Cell(flexcpText, lngRow, TProjectCol.pcMethod) = " "
End Sub

Private Sub LoadRequestPagePart(ByVal lngProjectId As Long, ByVal strSex As String, ByVal strExtData As String)
'���ܣ���ʼ����鲿λ����ʽ������
'������mstrExtData=������鲿λ����Ϣ,Ϊ��ʱ��ʾ�������������Ŀ
On Error GoTo errH
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long, i As Integer
    Dim str���� As String, str���� As String
    Dim arrData As Variant, strNoneRegion As String
    Dim blnNone As Boolean
    Dim Y As Long, str���� As String
    Dim lngRowHideCount As Long
    Dim lngProjectRowIndex As Long
    
    vfgList.Rows = 1
    
    With vfgList
        .WordWrap = True
        .FixedRows = 1
        .FixedCols = 0
        .Rows = .FixedRows + 1
        .Cols = 4
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        
        If str���� = "����" Then
            .TextMatrix(0, 0) = "�걾����"
            .TextMatrix(0, 1) = "�걾����"
            .TextMatrix(0, 2) = "�������"
        Else
            .TextMatrix(0, 0) = "��鲿λ"
            .TextMatrix(0, 1) = "��鲿λ"
            .TextMatrix(0, 2) = "��鷽��"
        End If
        
        .TextMatrix(0, 3) = "��ע"
        .RowHeight(0) = 300
        .ColComboList(2) = "..."
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1: .ColWidth(i) = 1600
        Next
        
    End With
    
    '��ȡ�����Ŀ������Ϣ
    strSQL = "Select ����,��������,ִ�б�� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngProjectId)
    
    If rsTmp.RecordCount <= 0 Then
        lngProjectRowIndex = Val(vfgRequestProject.RowSel)
        
        Call SetNullMethod(lngProjectRowIndex)
        
        Exit Sub
    End If
        
    str���� = rsTmp!��������
    str���� = rsTmp!����
        
    '���˼�鲿λ��Ϣ
    Set rsTmp = mrsRequestPart
    rsTmp.Filter = "������ĿId=" & lngProjectId & " and ����='" & str���� & "'"
    rsTmp.Sort = "����"
    
    blnNone = rsTmp.EOF
    
    If rsTmp.RecordCount <= 0 Then
        lngProjectRowIndex = Val(vfgRequestProject.RowSel)
        
        Call SetNullMethod(lngProjectRowIndex)
        
        Exit Sub
    End If
    
    If rsTmp.RecordCount > 0 Then
        lngRowHideCount = 50
        vfgList.Rows = vfgList.Rows + lngRowHideCount
        
        '���ù̶������������У����ڽ�ѡ�еĲ�λ�ƶ�����Ҫλ��
        For i = 1 To lngRowHideCount
            vfgList.RowHidden(i) = True
        Next i
    End If


    With vfgList
        '��ʾ��׼�Ĳ�λ��Ĭ�Ϸ���
        If blnNone Then
            .Editable = flexEDNone
            .TabStop = False
        Else
            .Editable = flexEDKbdMouse
        End If
        
        If str���� = "����" Then
            .TextMatrix(0, 0) = "�걾����"
            .TextMatrix(0, 1) = "�걾����"
            .TextMatrix(0, 2) = "�������"
        Else
            .TextMatrix(0, 0) = "��鲿λ"
            .TextMatrix(0, 1) = "��鲿λ"
            .TextMatrix(0, 2) = "��鷽��"
        End If
        
        .TextMatrix(0, 3) = "��ע"
        .RowHeight(0) = 300
        .ColComboList(2) = "..."
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        
        Do While Not rsTmp.EOF
            If .TextMatrix(.Rows - 1, 1) <> rsTmp!��λ Then
            
                If .TextMatrix(.Rows - 1, 1) <> "" Then
                    .Rows = .Rows + 1
                End If
            
                .TextMatrix(.Rows - 1, 0) = zlCommFun.GetNeedName("" & rsTmp!����)
                .TextMatrix(.Rows - 1, 1) = rsTmp!��λ
                Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                .Cell(flexcpData, .Rows - 1, 2) = CStr(NVL(rsTmp!��鷽��))  '������ѡ����ʹ��
                .TextMatrix(.Rows - 1, 3) = NVL(rsTmp!��ע)
            End If
            
            If NVL(rsTmp!Ĭ��, 0) = 1 Then '��"������1,������2,..."�ķ�ʽ��ʾ��λ��鷽��
                .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & NVL(rsTmp!����)
                If Left(.TextMatrix(.Rows - 1, 2), 1) = "," Then
                    .TextMatrix(.Rows - 1, 2) = Mid(.TextMatrix(.Rows - 1, 2), 2)
                End If
            End If
            
            rsTmp.MoveNext
        Loop
        
        '�޸�ʱ�������е�����
        '  ���Ϊ�գ�Ҳ��������ǰ�ĵ���λ�����Ŀ����ʱҪ�������ķ�ʽ����ѡ��λ
        '  ���߶�����ǰ�ĵ���λ��Ŀ��ǿ�д�����ǰ�Ĳ�λ(û�з���)���ֻ�������ͬ����λ
        If strExtData <> "" Then
            arrData = Split(Split(strExtData, vbTab)(0), "|")
            
            For i = 0 To UBound(arrData)
                lngIdx = .FindRow(CStr(Split(arrData(i), ";")(0)), 1, 1, , True)
                str���� = ""
                
                If lngIdx <> -1 Then
                
                    '��鷽����û�в����ڵ�
                    For Y = 0 To UBound(Split(Split(arrData(i), ";")(1), ","))
                        If InStr(.Cell(flexcpData, lngIdx, 2), CStr(Split(Split(arrData(i), ";")(1), ",")(Y))) = 0 Then
                            strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0) & "(" & Split(Split(arrData(i), ";")(1), ",")(Y) & ")"
                        Else
                            str���� = str���� & "," & Split(Split(arrData(i), ";")(1), ",")(Y)
                        End If
                    Next
                    
                    '�ò�λ�ķ���:������ǰ������ֻ�в�λû�з���
                    If UBound(Split(arrData(i), ";")) >= 1 Then
                        .TextMatrix(lngIdx, 2) = Mid(str����, 2)
                    Else
                        .TextMatrix(lngIdx, 2) = ""
                    End If
                    
                    .Cell(flexcpData, lngIdx, 1) = 1 '�����ò�λ��ѡ��
                    .Cell(flexcpFontBold, lngIdx, 1, lngIdx, 3) = True
                    .Cell(flexcpBackColor, lngIdx, 1, lngIdx, 3) = &HC0E0FF
                    
                    Set .Cell(flexcpPicture, lngIdx, 1) = img16.ListImages("c1").Picture
                    
                    '����ѡ�����ƶ�������
                    Call MoveToFirst(lngIdx)
                Else
                    '�ò�λ�����Ѳ�����
                    strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0)
                End If
            Next
        End If
        
        .Row = 1: .Col = 1
        .ShowCell .Row, .Col
        
        'ȷ�����ߴ�
        .AutoSize 0, .Cols - 1
        If .ColWidth(0) < 500 Then .ColWidth(0) = 500
        If .ColWidth(0) > 850 Then .ColWidth(0) = 850
        If .ColWidth(1) < 800 Then .ColWidth(1) = 800
        If .ColWidth(1) > 1600 Then .ColWidth(1) = 1600
        If .ColWidth(2) < 2500 Then .ColWidth(2) = 2500
        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
        If .ColWidth(3) < 800 Then .ColWidth(3) = 800
        If .ColWidth(3) > 2000 Then .ColWidth(3) = 2000
        
        lngIdx = 0
        For i = 0 To .Cols - 1
            lngIdx = lngIdx + .ColWidth(i) + 15
        Next
        
    End With
    
    'ɾ�����ص�������
    For i = vfgList.Rows - 1 To 1 Step -1
        If vfgList.RowHidden(i) Then Call vfgList.RemoveItem(i)
    Next i
        
    '�Ѳ����ڵĲ�λ��ʾ
    If strNoneRegion <> "" Then
        If str���� = "����" Then
            MsgBox "���²���걾����Ŀ�������Ѳ����ڣ�" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        Else
            MsgBox "���¼�鲿λ�򷽷�����Ŀ�������Ѳ����ڻ����ã�" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
            
            '������ڲ����õķ������������ֶ��༭����
'            vfgList.Editable = flexEDKbdMouse
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub MoveToFirst(ByVal lngCurRow As Long)
'�ƶ�����һ��
    Dim strRowData As Variant
    Dim strRowText As Variant
    Dim varRowPic  As Variant
    Dim blnFontBold As Boolean
    Dim blnRowHidden As Boolean
    Dim lngBakColor As OLE_COLOR
    
    Dim i As Long
    Dim lngUpRow As Long
    
    If vfgList.Rows = 2 Then Exit Sub

    '��ѯ�״�û�б���ѡ����
    For i = 1 To vfgList.Rows - 1
        If vfgList.Cell(flexcpData, i, 1) <> 1 Then
            lngUpRow = i
            Exit For
        End If
    Next i
    
    blnRowHidden = vfgList.RowHidden(lngUpRow)
    
    For i = 0 To vfgList.Cols - 1
        
        strRowText = vfgList.TextMatrix(lngUpRow, i)
        strRowData = vfgList.Cell(flexcpData, lngUpRow, i)
        blnFontBold = vfgList.Cell(flexcpFontBold, lngUpRow, i)
        lngBakColor = vfgList.Cell(flexcpBackColor, lngUpRow, i)
        
        Set varRowPic = vfgList.Cell(flexcpPicture, lngUpRow, i)
        
        vfgList.TextMatrix(lngUpRow, i) = vfgList.TextMatrix(lngCurRow, i)
        vfgList.Cell(flexcpData, lngUpRow, i) = vfgList.Cell(flexcpData, lngCurRow, i)
        vfgList.Cell(flexcpPicture, lngUpRow, i) = vfgList.Cell(flexcpPicture, lngCurRow, i)
        vfgList.Cell(flexcpFontBold, lngUpRow, i) = vfgList.Cell(flexcpFontBold, lngCurRow, i)
        vfgList.Cell(flexcpBackColor, lngUpRow, i) = vfgList.Cell(flexcpBackColor, lngCurRow, i)
        
        vfgList.TextMatrix(lngCurRow, i) = strRowText
        vfgList.Cell(flexcpData, lngCurRow, i) = strRowData
        vfgList.Cell(flexcpPicture, lngCurRow, i) = varRowPic
        vfgList.Cell(flexcpFontBold, lngCurRow, i) = blnFontBold
        vfgList.Cell(flexcpBackColor, lngCurRow, i) = lngBakColor
    Next i
    
    vfgList.RowHidden(lngUpRow) = vfgList.RowHidden(lngCurRow)
    vfgList.RowHidden(lngCurRow) = blnRowHidden
End Sub


Private Sub picBack_Resize()
On Error Resume Next
    Dim lngFreeHeight As Long
    
    picPart.Height = 3900
    
    Call AdjustRequestFace
    
    uspRequestPage.Left = 0
    uspRequestPage.Top = 0
    uspRequestPage.Width = picBack.ScaleWidth
    uspRequestPage.Height = picBack.ScaleHeight
    
    '���µ�����Ŀѡ��߶ȣ�ʹ���ܹ��Զ���Ӧҳ���С
    If picRequestInf.Top + picRequestInf.Height < picBack.ScaleHeight - 240 Then
        lngFreeHeight = picBack.Height - picRequestInf.Top - picRequestInf.Height - 240
        picPart.Height = 3900 + lngFreeHeight
    Else
        picPart.Height = 3900
    End If

    '��������ҳ�����������Ĵ�С��λ��
    Call AdjustRequestFace

End Sub


Private Sub CloseDiagnoseCodeInput()
'    Call ClipCursor(0)
    
    picBack.Enabled = True
    
    cmdInput.SetFocus
End Sub


Private Sub picInput_Click()
On Error GoTo errHandle
    Dim curPos As PointAPI
    Dim objShp As Shape
    
 
    '��ȡ��ǰָ���Ӧ�ڿؼ��е�λ��
    Call GetCursorPos(curPos)
    
    Call ScreenToClient(picInput.hwnd, curPos)
    
    curPos.X = picInput.ScaleX(curPos.X, vbPixels, vbTwips)
    curPos.Y = picInput.ScaleY(curPos.Y, vbPixels, vbTwips)
     
    For Each objShp In shpBackLine
        If curPos.X > objShp.Left And curPos.X < objShp.Left + objShp.Width And curPos.Y < objShp.Top And objShp.Visible Then
            rtbInputPro(objShp.Index).SetFocus
            Exit Sub
        End If
    Next
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub picPart_Resize()
On Error Resume Next
    '������Ŀѡ���Ĵ�С��ʹ���ܹ��Զ���Ӧ
    vfgRequestProject.Height = picPart.Height - vfgRequestProject.Top - 120
    vfgList.Height = picPart.Height - vfgList.Top - 120
err.Clear
End Sub


Private Sub AutoSizeInputFace()
'�Զ���Ӧ¼�봰�ڽ���
    Dim objLastControl As Object
    Dim objCurControl As Object
    
    Dim objLab As Object
    Dim objMust As Object
    Dim objText As Object
    Dim objShp As Object
    
    Dim lngControlCount As Long
    Dim blnInputObj As Boolean
    Dim blnIsBigDistance As Boolean
    Dim i As Long
    Dim blnExeCase As Boolean
    
    blnIsBigDistance = IIF(rtbInputPro.Count <= 4, True, False)
    Set objLastControl = Nothing
    
    lngControlCount = (shpBackLine.Count - 1) + (shpInputPro.Count - 1)
    
    '��ǩ����¼��־���ı����»���
    For i = 1 To lngControlCount

        blnExeCase = Not HasControl("shpInputPro", i)
        
        If HasControl("shpInputPro", i) Then
            blnExeCase = Not shpInputPro(i).Visible
        End If
        
        If blnExeCase Then
            If HasControl("labInputPro", i) Then
            
                Set objCurControl = labInputPro(i)
    
                If objCurControl.Visible Then
                
                    If Not (objLastControl Is Nothing) Then
                        objCurControl.Top = objLastControl.Top + objLastControl.Height + 120 + IIF(blnIsBigDistance, 360, 180)
                    End If
                    
                    Set objLab = objCurControl
                    
                    If HasControl("labMustPro", i) Then
                        Set objMust = labMustPro(i)
                        objMust.Top = objLab.Top
                    End If
                    
                    If HasControl("rtbInputPro", i) Then
                        Set objText = rtbInputPro(i)
                        objText.Top = objLab.Top - 30
                    End If
                                
                    Set objShp = shpBackLine(i)
                    objShp.Top = objText.Top + objText.Height
                    
                    Set objLastControl = objShp
                End If
            End If
            
        Else
            If HasControl("shpInputPro", i) Then
                Set objCurControl = shpInputPro(i)
                
                If objCurControl.Visible Then
                    If Not (objLastControl Is Nothing) Then
                        objCurControl.Top = objLastControl.Top + objLastControl.Height + 360
                    End If
                    
                    Set objLastControl = objCurControl
                End If
            End If
        End If
        
    Next i
    
    If cmdInput.Visible Then
        cmdInput.Top = rtbInputPro(Val(cmdInput.Tag)).Top - 20
        cmdInput.Height = rtbInputPro(Val(cmdInput.Tag)).Height + 40
    End If
    
    picInput.Height = objLastControl.Top + objLastControl.Height + 120
    
End Sub

Private Function SetInputControlHight(objText As TextBox) As Boolean
'���������ؼ��ĸ߶�
    Dim lngTxtHeight As Long
    Dim lngBase As Long
    Dim lngOldHeight As Long
    Dim strTestText As String
    
    lngBase = Me.TextHeight("��")
    SetInputControlHight = False
    
    With objText
        strTestText = .Text & "������������������������������"
        
        lngTxtHeight = Me.TextHeight(strTestText) * IIF(Me.TextWidth(strTestText) > .Width, Fix(Me.TextWidth(strTestText) / .Width) + 1, 1)
        
        If lngTxtHeight = .Height Then Exit Function
        
        If lngTxtHeight < 375 Then
            If .Height = 375 Then Exit Function
            
            .Height = 375
        ElseIf lngTxtHeight > 1635 Then
            If .Height = 1635 Then Exit Function
            
            .Height = 1635
        Else
            If lngTxtHeight + Fix(lngTxtHeight / lngBase) * 60 > 1635 Then
                .Height = 1635
            Else
                If lngTxtHeight + Fix(lngTxtHeight / lngBase) * 60 = .Height Then Exit Function
                .Height = lngTxtHeight + Fix(lngTxtHeight / lngBase) * 60
            End If
        End If
    End With
    
    SetInputControlHight = True
End Function


Private Sub AdjustRtbInputControlHeight(ByVal Index As Integer)
'����¼����ı����ݣ��������븽��¼��ؼ��߶�
    Dim blnChangeHeight As Boolean
    Dim objCurText As TextBox
    
    Set objCurText = rtbInputPro(Index)
    blnChangeHeight = SetInputControlHight(objCurText)
        
    If blnChangeHeight Then
        Call AutoSizeInputFace
        Call picBack_Resize
        
        Call uspRequestPage.CalcScroll
    End If
End Sub


Private Sub rtbInputPro_Change(Index As Integer)
On Error GoTo errHandle
    
    If rtbInputPro(Index).Visible Then
        Call SetRequestPageState(True)
        
        Call AdjustRtbInputControlHeight(Index)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Resume
End Sub


Private Function GetTabAdviceId(ByVal strSelTabTag As String) As Long
'��ȡtabҲ��Ӧ��ҽ��ID
    Dim lngTemp As Long
    
    lngTemp = Val(strSelTabTag)
    
    GetTabAdviceId = Val(Replace(strSelTabTag, lngTemp & IIF(InStr(strSelTabTag, "-") > 0, "-", ""), ""))
End Function

Private Function IsEqualTabPage(ByVal tabItem As XtremeSuiteControls.ITabControlItem) As Boolean
    IsEqualTabPage = IIF(tabItem.Index & "-" & tabItem.Caption = tbcRequest.Tag, True, False)
End Function

Private Sub rtbInputPro_GotFocus(Index As Integer)
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strPro As String
    Dim strElement As String
    
    strElement = GetInputProElement(rtbInputPro(Index).Tag)
    
    cmdInput.Left = rtbInputPro(Index).Left + rtbInputPro(Index).Width
    cmdInput.Top = rtbInputPro(Index).Top - 20 '+ rtbInputPro(Index).Height - cmdInput.Height
    cmdInput.Height = rtbInputPro(Index).Height + 40
    
    cmdInput.Tag = Index '�洢¼���Ķ�������
        
    If Not rtbInputPro(Index).Visible Or (rtbInputPro(Index).Locked And strElement <> M_STR_FIXEDELEMENT_DIAGNOSE) Then
'        cmdInput.Visible = False
        mblnShowWord = True
        Exit Sub
    End If
    
    strPro = GetInputProName(rtbInputPro(Index).Tag)
    
    strSQL = "select id from ��������ģ�� where �����ļ�id=[1] and ���ݸ���=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ģ��", GetRequestId(tbcRequest.Selected.Tag), strPro)
    
    If rsData.RecordCount > 0 Or strElement = M_STR_FIXEDELEMENT_DIAGNOSE Then
        '������ڸ�������ģ�壬����ʾģ�嵯����ť
'        rtbInputPro(Index).Width = shpBackLine(Index).Width - cmdInput.Width
        
        mblnShowWord = False
'        cmdInput.Visible = True
    Else
'        cmdInput.Visible = False
        mblnShowWord = True
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub rtbInputPro_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHandle
'    If KeyAscii <> 13 Then Exit Sub
'    If Not cmdInput.Visible Then Exit Sub
'
'    Call cmdInput_Click
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errHandle
    Dim lngRequestId As Long
    Dim i As Long
    Dim lngStart As Long
    Dim objCurItem As XtremeSuiteControls.ITabControlItem
    
    If Item.Tag = "" Then Exit Sub  '���û��tagֵ�����������ִ��ɾ������
    lngRequestId = Val(Item.Tag)
    
    lngStart = -1
    For i = 0 To tbcRequest.ItemCount - 1
        If GetRequestId(tbcRequest.Item(i).Tag) <> lngRequestId Then
            tbcRequest.Item(i).Visible = False
        Else
            If lngStart = -1 Then lngStart = i
            If tbcRequest.Item(i).Caption = "����Ŀ" Then lngStart = i
            
            tbcRequest.Item(i).Visible = True
        End If
    Next i
    
    If tbcRequest.ItemCount <= 0 Or lngStart < 0 Then Exit Sub
    
    If mblnIsRestoreTab Then    '�ָ�ҳ�����
        Exit Sub
    End If

    tbcRequest.Item(lngStart).Selected = True
    
    If lngStart <> tbcRequest.Selected.Index Then    '�ָ�ҳ�����
        Exit Sub
    End If
    
    tbcPage.Tag = Item.Index
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tbcRequest_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errHandle
    Dim lngHintResult As Long
    Dim strRequestName As String
    Dim objCurAppData As clsApplicationData
    
    mblnIsRestoreTab = False
    
    If mblnIsRestoreTab Then Exit Sub
    
    '����л����ҳ����ͬ����item�޶�Ӧ���ݣ����˳�����
    If IsEqualTabPage(Item) Or Item.Tag = "" Then Exit Sub
    
    mblnIsLoadRequestPage = True
    
    If mblnPageUpdateState Then
        strRequestName = Replace(tbcRequest.Tag, Val(tbcRequest.Tag) & "-", "")
        lngHintResult = MsgBox("��" & strRequestName & "�������Ѿ��ı䣬�Ƿ񱣴棿", vbYesNoCancel, Me.Caption)
        
        If lngHintResult = vbYes Then
            '�������뵥
            If Not SaveRequest(Val(tbcRequest.Tag)) Then
                mblnIsRestoreTab = True
                
                '�������ʧ�ܣ���ָ�ҳ���л�
                tbcPage.Item(Val(tbcPage.Tag)).Selected = True
                tbcRequest.Item(Val(tbcRequest.Tag)).Selected = True
                Exit Sub
            End If
            
            Call SetRequestPageState(False)
            
        ElseIf lngHintResult = vbNo Then
            Call SetRequestPageState(False)
            
            '�����ϱ༭���еĻ�������
            If Not mclsDiagEdit Is Nothing Then
                Call mclsDiagEdit.DeleteApplyDiag(Val(tbcRequest.Item(Val(tbcRequest.Tag)).Tag))
            End If
        Else
            tbcRequest.Item(Val(tbcRequest.Tag)).Selected = True
            Exit Sub
            
        End If
        
    End If
    
    tbcRequest.Tag = Item.Index & "-" & Item.Caption
    
    Call LoadRequestPage(GetRequestId(Item.Tag), Item.Tag)

    
    mblnIsLoadRequestPage = False
Exit Sub
errHandle:
    mblnIsLoadRequestPage = False
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadRequestPage(ByVal lngApplicationPageId As Long, ByVal strApplicationPageTag As String)
'��������ҳ��
On Error GoTo errHandle
    Dim objCurAppData As clsApplicationData
    Dim strDataKey As String
    Dim lngFreeHeight As Long
    
    Set objCurAppData = New clsApplicationData '���û�����ö�Ӧ���ԣ�������ʹ��Ĭ��ֵ
    
    strDataKey = strApplicationPageTag
    If mobjAppDatas.Exists("_" & strDataKey) Then
        Set objCurAppData = Nothing
        Set objCurAppData = mobjAppDatas.Item("_" & strDataKey)
    End If
    
    '�Ȼָ�����ҳ������
    Call RestoreRequestPageCfg
    
    rtbInputPro(0).BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)

    
    '�����������Ӧ�����м����Ŀ����λ����ʱ���ݼ�
    Call LoadRequestPartDataSet(lngApplicationPageId, mCurPatientInf.strSex)
    
    '��������¼������
    Call LoadRequestAffixInputCfg(lngApplicationPageId, objCurAppData.strRequestAffix)
    
    '����������Ŀ����
    Call LoadRequestPageProject(lngApplicationPageId, objCurAppData.lngProjectId, _
                                objCurAppData.lngExeRoomId, objCurAppData.strExeRoomName, _
                                objCurAppData.lngExeType, objCurAppData.strPartMethod)
    
    '���Ĭ��ѡ����һ����Ŀ�����ȡĬ����Ŀ��Ӧ�ļ�鲿λ���������
    If vfgRequestProject.RowSel < 0 Then
        Call LoadRequestPagePart(0, "", "")
    Else
        Call LoadRequestPagePart(vfgRequestProject.Cell(flexcpData, vfgRequestProject.RowSel, TProjectCol.pcId), _
                                mCurPatientInf.strSex, _
                                vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, TProjectCol.pcMethod))
    End If
    
    '����ҽ����ִ��ʱ��
    If objCurAppData.strStartExeTime <> "" Then
        dtpExeTime.value = objCurAppData.strStartExeTime
    End If
    
    chkPriority.value = IIF(objCurAppData.blnIsPriority, 1, 0)
'    chkPriority.FontBold = IIF(objCurAppData.blnIsPriority, True, False)
    

    '��������ҳ���沼��
    picPart.Height = 3900
    
    Call AdjustRequestFace
    
    '����ҳ����غ����������¼���picPart�ĸ߶ȣ�ʹ���Զ���Ӧ
    If picRequestInf.Top + picRequestInf.Height < picBack.Height - 240 Then
        lngFreeHeight = picBack.Height - picRequestInf.Top - picRequestInf.Height - 240
        picPart.Height = 3900 + lngFreeHeight
    Else
        picPart.Height = 3900
    End If

    '��������ҳ����������Ĵ�С��λ��
    Call AdjustRequestFace
    
    uspRequestPage.UCScrollState = False
    Call uspRequestPage.CalcScroll
        
    Call ShowArrangeState
    Call ShowStudyProjectToTxt
    
    labRequestDoctValue.Caption = IIF(objCurAppData.strRequestDoctor <> "", objCurAppData.strRequestDoctor, UserInfo.����)
     
    '���ҽ���������޸ģ���ؼ��ɱ༭��������Ϊenable
    picAuditing.Visible = Not objCurAppData.blnAllowUpdate
    
    txtFind.Enabled = objCurAppData.blnAllowUpdate
    txtFind.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    
    dtpExeTime.Enabled = objCurAppData.blnAllowUpdate
    cbxExeRoom.Enabled = objCurAppData.blnAllowUpdate
    
    
    chkPriority.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    chkPriority.Enabled = objCurAppData.blnAllowUpdate

    vfgRequestProject.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    vfgRequestProject.BackColorBkg = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    vfgRequestProject.Editable = IIF(objCurAppData.blnAllowUpdate, flexEDKbdMouse, flexEDNone)
    
    vfgList.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    vfgList.BackColorBkg = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    vfgList.Editable = IIF(objCurAppData.blnAllowUpdate, flexEDKbdMouse, flexEDNone)
    
    picBaseInf.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    picInput.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    picPart.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    picRequestInf.BackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    
    uspRequestPage.UCBackColor = IIF(objCurAppData.blnAllowUpdate, vbWhite, &HE0E0E0)
    picInput.Enabled = objCurAppData.blnAllowUpdate
'    picPart.Enabled = objCurAppData.blnAllowUpdate
    
    tmrFocus.Enabled = True
    
    mblnIsLoadRequestPage = False
Exit Sub
errHandle:
    mblnIsLoadRequestPage = False
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadRequestPartDataSet(ByVal lngRequestId As Long, ByVal strSex As String)
'���������뵥�����м�鲿λ��Ϣ
    Dim strSQL As String

    strSQL = "select distinct a.Id,a.����,a.Ĭ��, c.������ĿID, b.�����Ա�, b.����, a.��λ, a.����, b.���� as ��鷽��, b.����,b.��ע " & _
            " from ������Ŀ��λ a, ���Ƽ�鲿λ b, ��������Ӧ�� c " & _
            " where a.����=b.���� and a.��λ=b.���� and a.��ĿId = c.������ĿId " & _
                    " And (b.�����Ա� = [2] or Nvl(b.�����Ա�,0)=0) and c.�����ļ�Id=[1] order by ����,����"

    Set mrsRequestPart = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���뵥��鲿λ", lngRequestId, _
                                                    IIF(strSex = "��", 1, IIF(strSex = "Ů", 2, 0)))
End Sub



Private Sub AdjustRequestFace()
    picBaseInf.Left = IIF(uspRequestPage.Width > picBaseInf.Width, Fix((uspRequestPage.Width - picBaseInf.Width) / 2), 0)
    picBaseInf.Top = 0
    
    picInput.Left = picBaseInf.Left
    picInput.Top = picBaseInf.Top + picBaseInf.Height
    picInput.Width = picBaseInf.Width
    
    picPart.Left = picBaseInf.Left
    picPart.Top = picBaseInf.Top + picBaseInf.Height + IIF(Val(picInput.Tag) = 1, picInput.Height, 0)
    picPart.Width = picBaseInf.Width
    
    picRequestInf.Left = picBaseInf.Left
    picRequestInf.Top = picPart.Top + picPart.Height
    picRequestInf.Width = picBaseInf.Width
End Sub

Private Sub FreeRequestInputControl()
    Dim objFree As Object
    
    'ɾ�����븽���е�¼������......
    For Each objFree In labInputPro
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then
                Unload objFree
            End If
        End If
    Next
    
    For Each objFree In labMustPro
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then
                Unload objFree
            End If
        End If
    Next
    
    For Each objFree In rtbInputPro
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    
    For Each objFree In shpInputPro
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    
    For Each objFree In shpBackLine
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    Set mobjLastControl = Nothing
End Sub

Private Sub Timer1_Timer()
On Error GoTo errHandle
    If mlngUpdateAppNo = 0 Then
        labRequestTimeValue.Caption = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm")
    End If
Exit Sub
errHandle:
End Sub

Private Sub tmrFocus_Timer()
On Error GoTo errHandle
    tmrFocus.Enabled = False
    If rtbInputPro.Count > 1 Then rtbInputPro(1).SetFocus
Exit Sub
errHandle:
End Sub

Private Sub txtFind_Change()
On Error GoTo errHandle
    Call FindRequestProject(txtFind.Text)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub txtFind_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii = 13 And Trim(txtFind.Text) <> "" Then Call FindRequestProject(txtFind.Text, True)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub FindRequestProject(ByVal strFind As String, Optional ByVal blnIsEnter As Boolean = False)
'���Ҽ��������Ŀ
    Dim i As Long
    Dim lngFindIndex As Long
    Dim strPY As String

    lngFindIndex = -1


    '�ӱ�ѡ������ݺ������¿�ʼ����
    For i = IIF(vfgRequestProject.Tag > -1, Val(vfgRequestProject.Tag), 0) To vfgRequestProject.Rows - 1
        strPY = GetPYCode(vfgRequestProject.Cell(flexcpText, i, 1))

        If UCase(strPY) Like "*" & UCase(strFind) & "*" Then
            If lngFindIndex <= -1 And Val(vfgRequestProject.Tag) <> i Then lngFindIndex = i
        End If
    Next i

    '���û���ҵ���������ݿ�ʼλ�����²���
    If lngFindIndex <= -1 Then
        For i = 0 To Val(vfgRequestProject.Tag)
            strPY = GetPYCode(vfgRequestProject.Cell(flexcpText, i, 1))
            If UCase(strPY) Like "*" & UCase(strFind) & "*" Then
                If lngFindIndex <= -1 Then lngFindIndex = i
            End If
        Next i
    End If

    If lngFindIndex >= 0 Then
        Call vfgRequestProject.Select(lngFindIndex, 0)
        Call vfgRequestProject.ShowCell(lngFindIndex, 0)
        
        Call SelectRequestProject(lngFindIndex)
        Call ShowStudyProjectToTxt
    End If
End Sub


Private Function GetPYCode(ByVal strChinese As String) As String
    Dim i As Long
    
    GetPYCode = ""
    
    For i = 1 To Len(strChinese)
        GetPYCode = GetPYCode & GetWordChar1(Mid(strChinese, i, 1))
    Next i

End Function


Private Function GetWordChar1(ByVal strWord As String) As String
'��ú��ֵ�ƴ������
On Error Resume Next
    If Asc(strWord) < 0 Then
        If Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "0":            Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "A":            Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "B":            Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "C":            Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "D":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "E":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "F":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "G":    Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "H":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "J":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "K":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "L":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "M":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("Ŷ") Then
            GetWordChar1 = "N":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("Ŷ") And Asc(Left(strWord, 1)) < Asc("ž") Then
            GetWordChar1 = "O":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("ž") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "P":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("Ȼ") Then
            GetWordChar1 = "Q":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("Ȼ") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "R":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "S":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "T":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "W":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") And Asc(Left(strWord, 1)) < Asc("ѹ") Then
            GetWordChar1 = "X":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("ѹ") And Asc(Left(strWord, 1)) < Asc("��") Then
            GetWordChar1 = "Y":        Exit Function
        End If
        
        If Asc(Left(strWord, 1)) >= Asc("��") Then
            GetWordChar1 = "Z":        Exit Function
        End If
    Else
        If UCase(strWord) <= "Z" And UCase(strWord) >= "A" Then
            GetWordChar1 = UCase(Left(strWord, 1))
        Else
            GetWordChar1 = strWord
        End If
    End If
End Function

Private Sub vfgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strMethod As String, i As Long, j As Long
    Dim arrMethod As Variant, arrSub As Variant
    Dim k As Long
    Dim blnDo As Boolean
    
    If vfgList.Editable = flexEDNone Then Exit Sub
    
    strMethod = vfgList.Cell(flexcpData, Row, Col)
    If strMethod = "" Then
        MsgBox "�ü�鲿λû�����ÿɹ�ѡ��ļ�鷽����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsMethod
        .Rows = 0
        
        arrMethod = Split(Replace(strMethod, vbTab, ";" & vbTab), ";")
        
        For i = 0 To UBound(arrMethod)
            arrSub = Split(arrMethod(i), ",")
            
            For j = 0 To UBound(arrSub)
                .Rows = .Rows + 1
                If j = 0 Then
                    If InStr(1, arrMethod(i), vbTab) > 0 Then
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 2 '�����ǹ�ѡ��
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 3) '��һλ����Ӱ����־
                        If InStr("," & vfgList.TextMatrix(vfgList.Row, 2) & ",", "," & Mid(arrSub(j), 3) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c1").Picture
                            
                            .Cell(flexcpFontBold, .Rows - 1, 0) = True
                            .Cell(flexcpData, .Rows - 1, 0) = 1
                            .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c0").Picture
                            
                            .Cell(flexcpFontBold, .Rows - 1, 0) = False
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                            .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbWhite
                        End If
                    Else
                        '�ų���
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 1 '�������ų���
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 2) '��һλ����Ӱ����־
                        If InStr("," & vfgList.TextMatrix(vfgList.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o1").Picture
                           
                            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 1) = True
                            .Cell(flexcpData, .Rows - 1, 0) = 1 '1Ϊѡ��
                            .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o0").Picture
                            
                            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 1) = False
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                            .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbWhite
                        End If
                    End If
                Else
                    '��ѡ����
                    .RowData(.Rows - 1) = 3 '�����ǹ�ѡ����
                    
                    .Cell(flexcpText, .Rows - 1, 1) = Mid(arrSub(j), 2)

                    If InStr("," & vfgList.TextMatrix(vfgList.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                        blnDo = True
                        '����û��ѡ��ʱ,�����ѡ��
                        For k = .Rows - 2 To 0 Step -1
                            If .RowData(k) <> 3 Then
                                If .Cell(flexcpData, k, 0) = 0 Then blnDo = False
                                Exit For
                            End If
                        Next
                    Else
                        blnDo = False
                    End If
                    
                    If blnDo Then
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c1").Picture
                        
                        .Cell(flexcpFontBold, .Rows - 1, 1) = True
                        .Cell(flexcpData, .Rows - 1, 0) = 1
                        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
                    Else
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                        
                        .Cell(flexcpFontBold, .Rows - 1, 1) = False
                        .Cell(flexcpData, .Rows - 1, 0) = 0
                        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
                    End If
                End If
            Next
        Next
        
        .Row = 0: .Col = 1
        
        .Height = .Rows * (.RowHeightMin + 15) + 30
        If .Height > Me.ScaleHeight - 100 Then .Height = Me.ScaleHeight - 100
        If .Height < 3 * (.RowHeightMin + 15) + 30 Then .Height = 3 * (.RowHeightMin + 15) + 30
        
        .Width = (vfgList.Width - 30) - (vfgList.CellLeft + 15)
        
        .Left = picPart.Left + vfgList.Left + vfgList.CellLeft + 15
        
        .Top = picBack.Top + uspRequestPage.Top + picPart.Top + vfgList.Top + vfgList.CellTop + vfgList.CellHeight '+ .Height
        If .Top + .Height > Me.ScaleHeight Then
            .Top = .Top - .Height - vfgList.CellHeight
        End If
        
        .ZOrder
        If .Tag = "AutoPopup" Then
            .Visible = .Rows > 1
        Else
            .Visible = True
        End If
        If .Visible Then .SetFocus
    End With
End Sub


Private Sub vfgList_DblClick()
On Error GoTo errHandle
    If vfgList.Editable <> flexEDNone And vfgList.MouseCol = 1 And vfgList.MouseRow >= vfgList.FixedRows Then
        Call vfgList_KeyPress(vbKeySpace)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
'���ܣ��Ǳ༭״̬ʱ���Զ��ƶ���Ԫ��
    Dim strPartGroup As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
 
            If vfgList.Col <= 1 Then
                vfgList.Col = vfgList.Col + 1
            ElseIf vfgList.Col = 2 And vfgList.Row <= vfgList.Rows - 2 Then
                vfgList.Row = vfgList.Row + 1
                vfgList.Col = 1
            ElseIf vfgList.Col = 2 And vfgList.Row = vfgList.Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        
    ElseIf KeyAscii = Asc("*") Then
 
            If vfgList.Col = 2 Then
                Call vfgList_CellButtonClick(vfgList.Row, vfgList.Col)
            End If

    ElseIf KeyAscii = vbKeySpace Then
        If vfgList.Editable <> flexEDNone Then
            If vfgList.Col = 1 Then
                If vfgList.Cell(flexcpData, vfgList.Row, vfgList.Col) = 1 Then
                
                    
                    Set vfgList.Cell(flexcpPicture, vfgList.Row, vfgList.Col) = img16.ListImages("c0").Picture
                    
                    vfgList.Cell(flexcpData, vfgList.Row, vfgList.Col) = 0
                    vfgList.Cell(flexcpFontBold, vfgList.Row, 1, vfgList.Row, 3) = False
                    vfgList.Cell(flexcpBackColor, vfgList.Row, 1, vfgList.Row, vfgList.Cols - 1) = vbWhite
                    
                Else
                    
                    Set vfgList.Cell(flexcpPicture, vfgList.Row, vfgList.Col) = img16.ListImages("c1").Picture
                    
                    vfgList.Cell(flexcpData, vfgList.Row, vfgList.Col) = 1
                    vfgList.Cell(flexcpFontBold, vfgList.Row, 1, vfgList.Row, 3) = True
                    vfgList.Cell(flexcpBackColor, vfgList.Row, 1, vfgList.Row, vfgList.Cols - 1) = &HC0E0FF
                    
                    '�Զ���������ѡ����
                    vfgList.Col = 2
                    vsMethod.Tag = "AutoPopup"
                    Call vfgList_CellButtonClick(vfgList.Row, vfgList.Col)
                    vsMethod.Tag = ""
                End If
                            
                '����ǰ��λ���ñ��浽��Ӧ��Ŀ��
                strPartGroup = GetPartMethod
                vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, 3) = strPartGroup
                
                If strPartGroup <> "" Then Call SelectRequestProject(vfgRequestProject.RowSel)
                
                Call ShowStudyProjectToTxt
                
                Call SetRequestPageState(True)
            ElseIf vfgList.Col = 2 Then
                Call vfgList_CellButtonClick(vfgList.Row, vfgList.Col)
            End If
        End If
    End If
End Sub


Private Function GetPartMethod() As String
'�ռ���λ�����������
    Dim i As Long
    
    With vfgList
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, 1) = 1 Then
                
                GetPartMethod = GetPartMethod & "|" & .TextMatrix(i, 1) & ";" & .TextMatrix(i, 2)
            End If
        Next
        
        If GetPartMethod <> "" Then
            GetPartMethod = Mid(GetPartMethod, 2)
        End If
    End With
End Function



Private Sub vfgList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errHandle
    If vfgList.Col = 1 And vfgList.MouseCol = 1 Then
        If X <= vfgList.CellLeft + 250 Then
            Call vfgList_KeyPress(vbKeySpace)
        End If
    End If
 
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SelectRequestProject(ByVal lngRowSel As Long, Optional ByVal blnIsLoad As Boolean = False)
    Dim oldSelectIndex As Long
    
    If lngRowSel <= -1 Then Exit Sub
    
    oldSelectIndex = Val(vfgRequestProject.Tag)
    
    '�ָ�֮ǰ��ѡ�е���ͼ��
    If oldSelectIndex >= 0 Then
        If oldSelectIndex <> lngRowSel Then
            Set vfgRequestProject.Cell(flexcpPicture, oldSelectIndex, TProjectCol.pcId) = img16.ListImages("o0").Picture
            vfgRequestProject.Cell(flexcpFontBold, oldSelectIndex, 0, oldSelectIndex, 2) = False
            vfgRequestProject.Cell(flexcpBackColor, oldSelectIndex, 0, oldSelectIndex, vfgRequestProject.Cols - 1) = vbWhite
        
            vfgRequestProject.Cell(flexcpText, oldSelectIndex, TProjectCol.pcMethod) = ""
        End If
    End If
    
    Set vfgRequestProject.Cell(flexcpPicture, lngRowSel, TProjectCol.pcId) = img16.ListImages("o1").Picture
    vfgRequestProject.Cell(flexcpFontBold, lngRowSel, 0, lngRowSel, 2) = True
    vfgRequestProject.Cell(flexcpBackColor, lngRowSel, 0, lngRowSel, vfgRequestProject.Cols - 1) = &HC0E0FF
    
    '����ִ�п���
    If cbxExeRoom.ListIndex >= 0 Then
        vfgRequestProject.Cell(flexcpText, lngRowSel, TProjectCol.pcExeRoom) = cbxExeRoom.ItemData(cbxExeRoom.ListIndex) & "-" & Replace(cbxExeRoom.Text, Val(cbxExeRoom.Text) & "-", "")
    End If
    
    vfgRequestProject.Tag = lngRowSel
    
    If Not blnIsLoad Then Call SetRequestPageState(True)
    
    Call ShowArrangeState
End Sub

Public Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Private Sub ShowArrangeState()
'��ʾִ����Ŀ�ĵ�ǰ����״̬
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strHint As String
    Dim lngProjectIndex As Long
    Dim lngRoomId As Long
    Dim lngProjectId As Long
    Dim strDateRang As String
    
    
    lngRoomId = 0
    strHint = ""
    
    strDateRang = "To_Date([1],'YYYY-MM-DD HH24:MI:SS') and To_Date([2],'YYYY-MM-DD HH24:MI:SS')"
    
    If cbxExeRoom.ListIndex >= 0 Then
        lngRoomId = cbxExeRoom.ItemData(cbxExeRoom.ListIndex)
        
        '��ѯ���ҵȴ����
        strSQL = "SELECT count(*) as ���� FROM ����ҽ����¼ a, ����ҽ������ b where a.id=b.ҽ��id and a.���id is null and a.�������='D' " & _
                    " and ��ʼִ��ʱ�� between " & strDateRang & " and nvl(b.ִ�й���,0)<3 and nvl(b.ִ��״̬,0)<> 2 and b.ִ�в���id+0=[3]"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ҵȴ����", _
                                            Format(dtpExeTime.value, "yyyy-mm-dd 00:00:00"), _
                                            Format(dtpExeTime.value, "yyyy-mm-dd 23:59:59"), _
                                            lngRoomId)
                                            
        strHint = "��ѡ���ҹ��С�" & Val(NVL(rsTemp!����)) & "���˵ȴ����    "
    End If
    
    lngProjectIndex = Val(vfgRequestProject.Tag)
    
    lngProjectId = 0
    
    If lngProjectIndex >= 0 Then
        lngProjectId = Val(vfgRequestProject.Cell(flexcpData, lngProjectIndex, TProjectCol.pcId))
        
        '��ѯ��Ŀ�ȴ����
        strSQL = "SELECT count(*) as ���� FROM ����ҽ����¼ a, ����ҽ������ b where a.id=b.ҽ��id and a.���id is null and a.�������='D' " & _
                    " and ��ʼִ��ʱ�� between " & strDateRang & " and nvl(b.ִ�й���,0)<3 and nvl(b.ִ��״̬,0)<> 2 and a.������Ŀid+0=[3]"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��Ŀ�ȴ����", _
                                            Format(dtpExeTime.value, "yyyy-mm-dd 00:00:00"), _
                                            Format(dtpExeTime.value, "yyyy-mm-dd 23:59:59"), _
                                            lngProjectId)
                                            
        strHint = strHint & "��ѡ��Ŀ���С�" & Val(NVL(rsTemp!����)) & "���˵ȴ����    "
    End If
    
'    '��ѯ��ǰ�����µ���Ŀ�ȴ����
'    If lngRoomId <> 0 And lngProjectId <> 0 Then
'        strSQL = "SELECT count(*) as ���� FROM ����ҽ����¼ a, ����ҽ������ b where a.id=b.ҽ��id and a.���id is null and a.�������='D' " & _
'                    " and ��ʼִ��ʱ�� between " & strDateRang & " and nvl(b.ִ�й���,0)<3 and nvl(b.ִ��״̬,0)<> 2 and a.������Ŀid=[3] and b.ִ�в���Id=[4]"
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��Ŀ�ȴ����", _
'                                            Format(dtpExeTime.value, "yyyy-mm-dd 00:00:00"), _
'                                            Format(dtpExeTime.value, "yyyy-mm-dd 23:59:59"), _
'                                            lngProjectId, lngRoomId)
'
'        strHint = strHint & "��ǰѡ�����Ŀ�С�" & Val(Nvl(rsTemp!����)) & "�����赽��ѡ���Ҽ��"
'    End If
    
    stbThis.Panels(2).Text = strHint
    
End Sub


Private Sub ShowStudyProjectToTxt()
'��ѡ��ļ����Ŀ�ͷ�����ʾ��txtֻ���ı���
    Dim lngProjectIndex As Long
    Dim dblPrice As Double
    Dim lngExeType As Long
    Dim strExeRoomName As String
    Dim lngExeRoomId As Long
    
    lngProjectIndex = Val(vfgRequestProject.Tag)
    If lngProjectIndex < 0 Then
        txtCurStudyProject.Text = ""
        Exit Sub
    End If
    
    strExeRoomName = vfgRequestProject.Cell(flexcpText, lngProjectIndex, TProjectCol.pcExeRoom)
    strExeRoomName = Replace(strExeRoomName, Val(strExeRoomName) & "-", "")
    
    txtCurStudyProject.Text = Replace("��" & strExeRoomName & "��" & _
                                        vfgRequestProject.Cell(flexcpText, lngProjectIndex, TProjectCol.pcName) & _
                                        "��" & vfgRequestProject.Cell(flexcpText, lngProjectIndex, TProjectCol.pcMethod), "|", "��")
                            
    lngExeType = 0    'ִ������
    If Val(vfgRequestProject.Cell(flexcpData, Val(vfgRequestProject.Tag), TProjectCol.pcBedCol)) = 1 Then
        lngExeType = 1
    ElseIf Val(vfgRequestProject.Cell(flexcpData, Val(vfgRequestProject.Tag), TProjectCol.pcOperCol)) = 1 Then
        lngExeType = 2
    End If
    
    If cbxExeRoom.ListIndex < 0 Then
        lngExeRoomId = 0
    Else
        lngExeRoomId = Val(cbxExeRoom.ItemData(cbxExeRoom.ListIndex))
    End If
                        
    '���������
    dblPrice = GetPrice(mCurPatientInf.lngID, mCurPatientInf.lngPageId, vfgRequestProject.Cell(flexcpData, lngProjectIndex, TProjectCol.pcId), _
                        Trim(vfgRequestProject.Cell(flexcpText, lngProjectIndex, TProjectCol.pcMethod)), _
                        lngExeType, _
                        mCurPatientInf.lngFrom, _
                        lngExeRoomId)
                        
    labPrice = "����" & Format(dblPrice, "0.00")
End Sub


Private Sub vfgList_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vfgList.EditSelStart = 0
    vfgList.EditSelLength = zlCommFun.ActualLen(vfgList.EditText)
End Sub

Private Sub vfgList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo errHandle
    If Col <> 2 Then Cancel = True
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_DblClick()
On Error GoTo errHandle
    If vfgRequestProject.RowSel <= -1 Then Exit Sub
    If vfgRequestProject.Editable = flexEDNone Then Exit Sub
    
    Call SelectRequestProject(vfgRequestProject.RowSel)
    Call ShowStudyProjectToTxt
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_GotFocus()
On Error GoTo errHandle
    If vfgRequestProject.RowSel < 0 And vfgRequestProject.Rows > 0 Then
        vfgRequestProject.Row = 0
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii <> vbKeySpace Then Exit Sub
    If vfgRequestProject.Editable = flexEDNone Then Exit Sub
    
    Call SelectRequestProject(vfgRequestProject.RowSel)

    If vfgRequestProject.Col >= TProjectCol.pcNormalCol Then
        Call SelectProjectExeType(vfgRequestProject.RowSel, vfgRequestProject.Col)
    End If

    
    Call ShowStudyProjectToTxt
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub SelectProjectExeType(ByVal lngRow As Long, ByVal lngCol As Long)
    '���û�����ݣ�����Ҫ����ѡ��״̬
    If vfgRequestProject.Cell(flexcpText, lngRow, lngCol) = "" Then Exit Sub
    
    Set vfgRequestProject.Cell(flexcpPicture, lngRow, lngCol) = img16.ListImages("o1").Picture
    vfgRequestProject.Cell(flexcpData, lngRow, lngCol) = 1
    
    Select Case lngCol
        Case TProjectCol.pcNormalCol
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcBedCol) = img16.ListImages("o0").Picture
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcOperCol) = img16.ListImages("o0").Picture
             
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcBedCol) = 0
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcOperCol) = 0
            
        Case TProjectCol.pcBedCol
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcNormalCol) = img16.ListImages("o0").Picture
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcOperCol) = img16.ListImages("o0").Picture
            
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcNormalCol) = 0
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcOperCol) = 0
            
        Case TProjectCol.pcOperCol
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcNormalCol) = img16.ListImages("o0").Picture
            Set vfgRequestProject.Cell(flexcpPicture, lngRow, TProjectCol.pcBedCol) = img16.ListImages("o0").Picture
            
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcNormalCol) = 0
            vfgRequestProject.Cell(flexcpData, lngRow, TProjectCol.pcBedCol) = 0
    End Select
End Sub


Private Sub vfgRequestProject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errHandle
    'ֻ��ѡ���һ��ʱ�����ܽ�����Ŀѡ��
    If vfgRequestProject.RowSel < 0 Then Exit Sub
    If vfgRequestProject.Editable = flexEDNone Then Exit Sub
    
    If vfgRequestProject.ColSel = 0 Then
        Call SelectRequestProject(vfgRequestProject.RowSel)
        
        Call ShowStudyProjectToTxt
        
    ElseIf vfgRequestProject.ColSel >= TProjectCol.pcNormalCol _
        And vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, vfgRequestProject.Col) <> "" Then
        
        Call SelectRequestProject(vfgRequestProject.RowSel)
        
        Call SelectProjectExeType(vfgRequestProject.RowSel, vfgRequestProject.Col)
        
        Call ShowStudyProjectToTxt
        
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_SelChange()
On Error GoTo errHandle
    Dim blnIsLoading As Boolean
    
    If vfgRequestProject.RowSel < 0 Then
        Call LoadRequestPagePart(0, "", "")
        Exit Sub
    End If
    
    blnIsLoading = False
    
    '�ж�����ԭ���Ƿ���������
    If mblnIsLoadRequestPage Then blnIsLoading = True
    
    mblnIsLoadRequestPage = True

    '������Ŀ��Ӧ�Ĳ�λ
    Call LoadRequestPagePart(vfgRequestProject.Cell(flexcpData, vfgRequestProject.RowSel, TProjectCol.pcId), _
                            mCurPatientInf.strSex, _
                            vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, TProjectCol.pcMethod))
    
    '������Ŀ��ִ�п���
    Call LoadRequestProjectExeRoom(vfgRequestProject.Cell(flexcpData, vfgRequestProject.RowSel, TProjectCol.pcId), _
                                    vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, TProjectCol.pcRoomType), _
                                    Val(vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, TProjectCol.pcExeRoom)))
                                    
                                    
    '����б�༭����Ҫ����Ŀ�Ŀɱ༭������ͬ
    vfgList.Editable = vfgRequestProject.Editable
                                    
    If Not blnIsLoading Then mblnIsLoadRequestPage = False
                                            
Exit Sub
errHandle:
    If Not blnIsLoading Then mblnIsLoadRequestPage = False
    
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgRequestProject_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    Cancel = True
err.Clear
End Sub

Private Sub vsMethod_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 And NewRow <> -1 Then
        If vsMethod.TextMatrix(NewRow, 0) = "" Then
            Cancel = True
            vsMethod.Col = 1
        End If
    End If
End Sub

Private Sub vsMethod_DblClick()
    Call vsMethod_KeyPress(13)
End Sub

Private Sub vsMethod_KeyPress(KeyAscii As Integer)
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    
    With vsMethod
        If KeyAscii = 13 Then
            Call ConfirmMethod
            vsMethod.Visible = False
            vfgList.SetFocus
        ElseIf KeyAscii = vbKeySpace Then
            '��鷽����ѡ����ȡ��
            If .Cell(flexcpData, .Row, 0) = 1 Then
                '��ѡ��ĿǰҲ����ȡ��ѡ��
                .Cell(flexcpData, .Row, 0) = 0
                .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = False
                
                Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o0", "c0")).Picture
                'ͬʱȡ���õ�ѡ�������
                If .RowData(.Row) = 1 Then
                    For i = .Row + 1 To .Rows - 1
                        If .RowData(i) = 3 Then
                            If .Cell(flexcpData, i, 0) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                                
                                Set .Cell(flexcpPicture, i, 1) = img16.ListImages("c0").Picture
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            Else
                blnDo = True
                If .RowData(.Row) = 3 Then
                    '����û��ѡ��ʱ,�����ѡ��
                    For i = .Row - 1 To 0 Step -1
                        If .RowData(i) <> 3 Then
                            If .Cell(flexcpData, i, 0) = 0 Then blnDo = False
                            Exit For
                        End If
                    Next
                End If
                
                If blnDo Then
                    .Cell(flexcpData, .Row, 0) = 1
                    .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
                    
                    Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o1", "c1")).Picture
                    
                    If .RowData(.Row) = 1 Then '��ѡ��ѡ��ʱ��ȡ��������ѡ��
                        For i = 0 To .Rows - 1
                            If i <> .Row And .RowData(i) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                                
                                Set .Cell(flexcpPicture, i, 0, i, 1) = img16.ListImages("o0").Picture
                                For j = i + 1 To .Rows - 1 'ͬʱȡ���õ�ѡ�������
                                    If .RowData(j) = 3 Then
                                        If .Cell(flexcpData, j, 0) = 1 Then
                                            .Cell(flexcpData, j, 0) = 0
                                            .Cell(flexcpFontBold, j, 0, j, .Cols - 1) = False
                                            
                                            Set .Cell(flexcpPicture, j, 1) = img16.ListImages("c0").Picture
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
            
            Call ConfirmMethod
        End If
    End With
End Sub


Private Sub ConfirmMethod()
'���ܣ���鷽����ȷ��
    Dim strMethod As String, i As Long
    Dim strPartGroup As String
        
    With vsMethod
        For i = 0 To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                strMethod = strMethod & "," & .TextMatrix(i, 1)
            End If
        Next
        
        vfgList.TextMatrix(vfgList.Row, 2) = Mid(strMethod, 2)
        
        '�������ú��Զ�ѡ�иò�λ
        If vfgList.TextMatrix(vfgList.Row, 2) <> "" Then
            vfgList.Cell(flexcpData, vfgList.Row, 1) = 1
            vfgList.Cell(flexcpFontBold, vfgList.Row, 1, vfgList.Row, 3) = True
            
            Set vfgList.Cell(flexcpPicture, vfgList.Row, 1) = img16.ListImages("c1").Picture
        End If
    End With
    
    strPartGroup = GetPartMethod
    vfgRequestProject.Cell(flexcpText, vfgRequestProject.RowSel, 3) = strPartGroup
    
    If strPartGroup <> "" Then
        Call SelectRequestProject(vfgRequestProject.RowSel)
    End If
    
    Call ShowStudyProjectToTxt
End Sub

Private Sub vsMethod_LostFocus()
    vsMethod.Visible = False
End Sub

Private Sub vsMethod_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = vsMethod.MouseCol
    If vsMethod.Col = lngCol And vsMethod.Text <> "" Then
        If X <= vsMethod.CellLeft + 250 Then
            Call vsMethod_KeyPress(vbKeySpace)
        End If
    End If
End Sub
