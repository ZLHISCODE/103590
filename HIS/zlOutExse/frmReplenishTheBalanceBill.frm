VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "ZLIDKIND.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReplenishTheBalanceBill 
   Caption         =   "ҽ���������"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReplenishTheBalanceBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDiagnose 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   1140
      ScaleHeight     =   660
      ScaleWidth      =   6780
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1740
      Width           =   6780
      Begin VSFlex8Ctl.VSFlexGrid vsDiagnose 
         Height          =   600
         Left            =   30
         TabIndex        =   5
         Top             =   75
         Width           =   6555
         _cx             =   11562
         _cy             =   1058
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   350
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
   End
   Begin MSCommLib.MSComm msCommSpeak 
      Left            =   14355
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   1650
      Left            =   -180
      ScaleHeight     =   1650
      ScaleWidth      =   14625
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5430
      Width           =   14625
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "ȫѡ(&A)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   15
         TabIndex        =   28
         ToolTipText     =   "�ȼ���Ctrl+A"
         Top             =   1230
         Width           =   1440
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ��(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1530
         TabIndex        =   27
         ToolTipText     =   "�ȼ���Ctrl+R"
         Top             =   1230
         Width           =   1440
      End
      Begin VB.TextBox txtYB 
         Height          =   300
         Left            =   795
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtժҪ 
         Height          =   360
         Left            =   990
         MaxLength       =   100
         TabIndex        =   8
         Top             =   90
         Width           =   6960
      End
      Begin VB.TextBox txt�˿�ϼ� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   7755
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0.00"
         ToolTipText     =   "�����շ�ʱδ�ɿ�ݵ�ʵ�ս��ϼ�"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Frame fraDownSplit 
         Height          =   135
         Left            =   -525
         TabIndex        =   18
         Top             =   945
         Width           =   15075
      End
      Begin VB.CommandButton cmdԤ���� 
         Caption         =   "Ԥ����(&V)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9840
         TabIndex        =   10
         ToolTipText     =   "�ȼ���F5"
         Top             =   1230
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ��(&C)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   13005
         TabIndex        =   12
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   1230
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   11430
         TabIndex        =   11
         Top             =   1230
         Width           =   1440
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   375
         Left            =   -15
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   510
         Width           =   11265
         _cx             =   19870
         _cy             =   661
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   12632256
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReplenishTheBalanceBill.frx":6852
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblժҪ 
         AutoSize        =   -1  'True
         Caption         =   "ժҪ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   7
         Top             =   150
         Width           =   480
      End
      Begin VB.Label lblʵ�� 
         AutoSize        =   -1  'True
         Caption         =   "ʵ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   13305
         TabIndex        =   25
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblӦ�� 
         AutoSize        =   -1  'True
         Caption         =   "Ӧ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11310
         TabIndex        =   24
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbl�˿�ϼ� 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ�˿�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   6450
         TabIndex        =   23
         Top             =   1290
         Width           =   1200
      End
   End
   Begin VB.PictureBox picFeeList 
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   2100
      ScaleHeight     =   1920
      ScaleWidth      =   5055
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3210
      Width           =   5055
      Begin VSFlex8Ctl.VSFlexGrid vsFeeList 
         Height          =   1515
         Left            =   -15
         TabIndex        =   6
         Top             =   -15
         Width           =   5325
         _cx             =   9393
         _cy             =   2672
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReplenishTheBalanceBill.frx":691D
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   4
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
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   -2940
      ScaleHeight     =   1035
      ScaleWidth      =   14085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   270
      Width           =   14085
      Begin VB.TextBox txtMCInvoice 
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9570
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   100
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9550
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   100
         Width           =   1545
      End
      Begin VB.ComboBox cboNO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   12040
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   100
         Width           =   1350
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13500
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F8"
         Top             =   90
         Width           =   400
      End
      Begin VB.ComboBox cboPayMode 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11820
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   630
         Width           =   2010
      End
      Begin VB.Frame fraInfo 
         Height          =   135
         Left            =   -150
         TabIndex        =   14
         Top             =   405
         Width           =   13980
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   375
         Left            =   735
         TabIndex        =   1
         Top             =   630
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmReplenishTheBalanceBill.frx":6A0E
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   2
         InputAppearance =   2
         ShowSortName    =   -1  'True
         DefaultCardType =   "0"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "���ݺ�"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   11245
         TabIndex        =   34
         Top             =   160
         Width           =   720
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ��"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   9085
         TabIndex        =   33
         Top             =   160
         Width           =   480
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   9345
         TabIndex        =   21
         Top             =   60
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblPatient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   180
         TabIndex        =   0
         Top             =   690
         Width           =   480
      End
      Begin VB.Label lblPatiInfor 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4290
         TabIndex        =   2
         Top             =   705
         Width           =   480
      End
      Begin VB.Label lblPayMode 
         AutoSize        =   -1  'True
         Caption         =   "ԭ֧����ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10350
         TabIndex        =   3
         Top             =   705
         Width           =   1200
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   90
         TabIndex        =   20
         Top             =   120
         Width           =   420
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   8070
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceBill.frx":6AC5
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13123
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   88
            Object.Tag             =   "���ڼ��ʻ��շѸ����ʻ���ʾ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   370
            MinWidth        =   88
            Object.Tag             =   "�����շ�Ԥ����ʾ"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   370
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "ҽ������"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceBill.frx":7359
            Key             =   "Calc"
            Object.ToolTipText     =   "������:ALT+?"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmReplenishTheBalanceBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum EM_Balance_Type
    EM_Balance_Register = 0 '�ҺŽ���
    EM_Balance_Charge = 1 '�շѽ���
    EM_Balance_Err_Cancel = 2 'EM_�쳣����
    EM_Balance_Err_ReCharge = 3 'EM_�쳣�����շ�
End Enum
'-----------------------------------------------------------
'�ӿ���ر���
Private mlngModule As Long, mstrPrivs As String
Private mEditType As EM_Balance_Type
Private mstrNo As String '��ǰ�����Ľ��㵥��
Private mstr����ID As String '��ǰ�����Ľ���ID
Private mstr������� As String '��ǰ�����Ľ������
Private mblnFirst As Boolean
Private mblnUnLoad As Boolean
Private mblnElsePersonErrBill As Boolean '�Ƿ������˵��쳣����
'-----------------------------------------------------------
'������ر���
Private mobjPayCards As Cards
Private mblnNotClearLedDisplay As Boolean   '�������ʾ
Private msngMinWidth As Single, msngMinHeight As Single
Private mstrTittle As String
Private mrsList As ADODB.Recordset
Private mblnNotClick As Boolean
Private mstrPreBalance As String '�ϴ�ѡ���֧����ʽ
Private mstrPreDiagnose As String '�ϴ�ѡ������
Private mintInsure As Integer
Private mstrYBPati  As String 'ҽ������
Private mobjPatiInfor As PatiInfor
Private mrs���㷽ʽ As ADODB.Recordset
Private mstrӦ������㷽ʽ As String
Private mcllDiagnose As Collection  '��ǰ������

Private mstr�����ʻ� As String '�Ƿ񽫸����ʻ����õ��շѿ���
Private Enum Pan
    C2��ʾ��Ϣ = 2
    C3�����ʻ� = 3
    C4Ԥ����Ϣ = 4
    C5ҽ������ = 5
End Enum
Private mintSucces As Integer '���óɹ�����
Private mstrPrePati  As String, mlngPrePati   As Long '�ϴβ�����Ϣ
Private mobjInvoice As clsInvoice
Private mobjFactProperty As clsFactProperty
Private mlng����ID As Long
Private mFrmBalanceWin As frmReplenishTheBalanceWin

Private Type Ty_Module_Para
     int����ʣ��Ʊ������ As Integer
     blnģ�����Ҳ��� As Boolean
     intģ������ As Integer
     blnҩ����λ As Boolean
     int�嵥��ӡ��ʽ As Integer
     int��������Ч���� As Integer
     str�����������շѷ�ʽ As String
End Type
Private mtyMoudlePara As Ty_Module_Para
Private Enum mEmPancelIDX
    EM_Pan_Pati = 1
    EM_Pan_Diagnose = 2
    EM_Pan_FeeList = 3
    EM_Pan_Down = 4
End Enum

Private Enum mEM_Diagnose_SelStatu
    EM_dgGrayToSeled = -1 '��ѡ��Ļ�ɫ��Ϊѡ��
    EM_dgClearAllSeled = 0 '�������ѡ�е����
    EM_dgSelAll = 1 'ѡ�����е����
    EM_dgSeledToGray = 5 'ȫ����ѡ�е�����Ϊ��ɫ
End Enum
'-----------------------------------------------------------
'ҽ���������
Private Type TY_Insure
    dbl����͸֧ As Double
    dbl�ʻ���� As Double
End Type
Private mTy_Insure As TY_Insure
 '��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    ҽ���ӿڴ�ӡƱ�� As Boolean
    ����Ԥ���� As Boolean
    �ֱҴ��� As Boolean
    ʵʱ��� As Boolean
    ���Ը� As Boolean
    ȫ�Ը� As Boolean
    ҽ������Ʊ��  As Boolean        'Ԥ����ʱ��Ч
    �Һ�ʹ�ø����ʻ� As Boolean
    ���ղ����� As Boolean   'support�ҺŲ���ȡ������
End Type
Private MCPAR As TYPE_MedicarePAR
Private mcolBalance As Collection 'ҽ��������Ϣ
Private mblnEdit As Boolean  '�Ƿ�༭��
Private mblnPrintBill As Boolean 'Ʊ���Ƿ��ӡ
Private mlng������ϸĿID As Long '�����Ѷ�Ӧ�շ�ϸĿID
Private mcur������ As Currency
'-------------------------------------------------------------------------------------
'API����:
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mrsBalanceNO As ADODB.Recordset '�������(No,�������,����ID)
Private mcllForceDelToCash As Collection 'ǿ��������Ϣ��Array(����Ա,���������,���㷽ʽ)
Private mstr�ų����㷽ʽ As String '����ʹ�õĽ��㷽ʽ,����ö��ŷָ�

Public Function zlEditCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As EM_Balance_Type, Optional ByRef str����ID As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-���õĸ�����
    '     lngModule-ģ���
    '     strPrivs-Ȩ�޴�
    '     EditCard-��ǰ�༭����
    '     str����ID-����ID(�쳣���ռ��쳣����ʱ����)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-16 11:32:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mintSucces = 0: mlngModule = lngModule: mstrPrivs = strPrivs
    mEditType = EditType: mblnFirst = True: mblnUnLoad = False
    mlngModule = 1124
    If CheckDepend = False Then Exit Function
    mstr����ID = str����ID
    Set mobjInvoice = New zlPublicExpense.clsInvoice
    If mobjInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then Exit Function
    If CheckDepend = False Then Unload Me: Exit Function
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    
    zlEditCard = mintSucces > 0
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2014-09-16 16:28:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varTemp As Variant
    With mtyMoudlePara
        .blnҩ����λ = zlDatabase.GetPara("ҩƷ��λ��ʾ", glngSys, mlngModule) = "1"
        .int�嵥��ӡ��ʽ = Val(zlDatabase.GetPara("�շ��嵥��ӡ��ʽ", glngSys, mlngModule))
        strTemp = zlDatabase.GetPara("����ģ�����ҷ�ʽ", glngSys, mlngModule)
        varTemp = Split(strTemp & "|", "|")
        .blnģ�����Ҳ��� = Val(varTemp(0)) = "1"
        .intģ������ = Val(varTemp(1))
        strTemp = Trim(zlDatabase.GetPara("Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", glngSys, mlngModule, "0|10"))
        varTemp = Split(strTemp & "|", "|")
        If Val(varTemp(0)) = 0 Then
            .int����ʣ��Ʊ������ = -1
        Else
            .int����ʣ��Ʊ������ = Val(varTemp(1))
        End If
        '84929
        .int��������Ч���� = Val(zlDatabase.GetPara("��������Ч����", glngSys, mlngModule, "3"))
        .str�����������շѷ�ʽ = zlDatabase.GetPara("����������շѽ��㷽ʽ", glngSys, mlngModule)
    End With
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����������Ϣ����ر���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-10 11:25:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strIDKindStr As String
    
    mstrTittle = "���ò����¼"
    Select Case mEditType
    Case EM_Balance_Charge
        mstrTittle = mstrTittle & "(�շѲ������)"
    Case EM_Balance_Err_Cancel
        mstrTittle = mstrTittle & "(�쳣��������)"
        cmdOK.Caption = "����(&O)"
    Case EM_Balance_Err_ReCharge
        mstrTittle = mstrTittle & "(�쳣��������)"
        cmdOK.Caption = "����(&O)"
    Case EM_Balance_Register
        mstrTittle = mstrTittle & "(�ҺŲ������)"
    Case Else
        mstrTittle = mstrTittle & "(�շѲ������)"
    End Select
    Me.Caption = mstrTittle
    
    '�Ƿ�ɽ��н����˷�
    cmdDelete.Visible = (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) _
        And zlStr.IsHavePrivs(mstrPrivs, "�����˷�")
    
    If gblnLED Then
        zl9LedVoice.Reset msCommSpeak
        zl9LedVoice.Init UserInfo.��� & " �շ�ԱΪ������", mlngModule, gcnOracle
    End If
    
    Call InitModulePara
    
    '��ȡ�����ѵ��շ�ϸĿID,84965
    Dim rsRecord As ADODB.Recordset
    Set rsRecord = zlGetSpecialItemFee("������")
    If Not rsRecord Is Nothing Then
        If Not rsRecord.EOF Then mlng������ϸĿID = Val(Nvl(rsRecord!�շ�ϸĿID))
    End If
    
    Set mobjFactProperty = New clsFactProperty
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_�շ��վ�, 0, 0, 0, mobjFactProperty)
    
    strIDKindStr = "��|��������￨;ҽ|ҽ����;��|���֤��;IC|IC����|1;��|�����;��|�շѵ��ݺ�"
    msngMinWidth = (800 * Screen.TwipsPerPixelX) * 0.5
    msngMinHeight = (600 * Screen.TwipsPerPixelY) * 0.5
    mstrTittle = "���ò����¼"
    lbl����.Caption = ""
    
    Call SetFeeListHead(True)   '��ʼ��������ͷ
    With vsDiagnose
        .Clear 1
        .Rows = 1: .COLS = 1
    End With
    Dim blnVisible As Boolean
    blnVisible = mEditType = EM_Balance_Register Or mEditType = EM_Balance_Charge
    cboPayMode.Visible = blnVisible: lblPayMode.Visible = blnVisible
    txtInvoice.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnStrictCtrl '89302
     
    Call InitPancel
    Call ClearData
    '��ʼ������ϱ�ؼ�
    Call PatiIdentify.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, _
        gobjSquare.objSquareCard, strIDKindStr, gstrSysName)
End Sub

Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2014-09-22 17:24:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PatiIdentify.Text = ""
    lblPatiInfor.Caption = ""
    Set mobjPatiInfor = Nothing
    cboPayMode.Clear
    
    vsBalance.Clear 1
    vsBalance.Rows = 1
    vsBalance.COLS = 1
    Set mcolBalance = New Collection
    
    vsFeeList.Clear 1
    vsFeeList.Rows = 2
    vsDiagnose.Clear 1
    vsDiagnose.Rows = 1
    vsDiagnose.COLS = 1
    lblʵ��.Caption = "ʵ��:0.00"
    lblӦ��.Caption = "Ӧ��:0.00"
    staThis.Panels(Pan.C3�����ʻ�).Text = ""
    staThis.Panels(Pan.C3�����ʻ�).Visible = False
    txtժҪ = "": txt�˿�ϼ�.Text = Format(0, "0.00")
    Call ClearDisplaySHow
    
    mblnEdit = False
    Call SetButtons '���ð�ť
    
    mcur������ = 0
    mstr�ų����㷽ʽ = ""
End Sub

Private Sub cboPayMode_Click()
    If mblnNotClick Then Exit Sub
    If mstrPreBalance = Trim(cboPayMode.Text) Then Exit Sub
    mstrPreBalance = Trim(cboPayMode.Text)
    
    If mrsList Is Nothing Then Exit Sub
    Call LoadFeeData(mrsList)
    Call SetButtons
End Sub

Private Sub cboPayMode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then
        Unload Me: Exit Sub
    End If
    If PatiIdentify.Locked Then
       SetPatientEnableModi True
       Call ClearData
       If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
       Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Call FromNosSel("", False, False, True)
    Call SetButtons
    vsDiagnose.Cell(flexcpChecked, 0, 0, vsDiagnose.Rows - 1, vsDiagnose.COLS - 1) = 2
End Sub

Private Sub cmdDelete_Click()
    '�����˷Ѵ���
    Call frmReplenishTheBalanceDel.zlShowMe(Me, mlngModule, mstrPrivs, EM_RBDTY_�˷�, "")
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim strNos As String, str����IDs As String, str����IDs As String
    Dim dtDate As Date, strNo As String, strReclaimInvoice As String
    Dim curȫ�Ը� As Currency, cur���Ը� As Currency, cur����ͳ�� As Currency
    
    mblnNotClearLedDisplay = True
    strNos = GetSelFeeNos '��ȡ���ν��㵥�ݺ�
    If mEditType = EM_Balance_Err_Cancel Then
        '�쳣����
        If CancelBalance = False Then Call SetButtons: Exit Sub
        Unload Me: mintSucces = mintSucces + 1
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    
    If mEditType = EM_Balance_Err_ReCharge Then
        '�������
        If zlIsCheckExistErrBill(Val(mstr�������), True) = False Then
            MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(Val(mstr�������)) Then
            MsgBox "��ǰ�����������������㴰���н��д����㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        '���������㷽ʽ��Ч�Լ��
        If ThreeBalanceCheck(mobjPayCards, IsRegister(), strNos, mcllForceDelToCash, mstr�ų����㷽ʽ) = False Then Exit Sub
        
        If CheckFactValied(True, mblnPrintBill) = False Then
            Call SetButtons: mblnNotClearLedDisplay = False
            Exit Sub
        End If
        '�쳣����
        dtDate = zlDatabase.Currentdate
        '��ʾ�ͷ�����Ҫ���յķ�Ʊ����ѡ��ȡ�����������
        If ShowReclaimInvoice(strNos, strReclaimInvoice) = False Then Exit Sub
        Call GetAsyncKeyState(VK_RETURN)
        
        If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
        Set mFrmBalanceWin = New frmReplenishTheBalanceWin
        If mFrmBalanceWin.zlChargeWin(Me, EM_Balance_Err_ReCharge, mlngModule, mstrPrivs, mobjPatiInfor, mobjPayCards, mstrNo, dtDate, mstr����ID, _
            mstr�������, MCPAR.�ֱҴ���, strNos, strReclaimInvoice, mcllForceDelToCash, mstr�ų����㷽ʽ, mblnElsePersonErrBill, _
            IsRegister()) = False Then
            If Not gfrmMain Is Nothing Then
                Call zlExeBalanceWinRefrshData(mstrNo, False, dtDate)
            End If
            Call SetButtons
            mblnNotClearLedDisplay = False
            Exit Sub
        End If
        Call SetButtons
        If Not gfrmMain Is Nothing Then
            Call zlExeBalanceWinRefrshData(mstrNo, True, dtDate)
        End If
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    
    If isValied(strNos, str����IDs, str����IDs) = False Then
        If vsFeeList.Enabled And vsFeeList.Visible Then vsFeeList.SetFocus
        Call SetButtons
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    
    '����ҽ��ͳ����
    If SaveItemYbMoney(mobjPatiInfor.����ID, strNos, IIf(mEditType = EM_Balance_Register, 4, 1), _
        curȫ�Ը�, cur���Ը�, cur����ͳ��) = False Then
        If vsFeeList.Enabled And vsFeeList.Visible Then vsFeeList.SetFocus
        Call SetButtons
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    '��֧��Ԥ����ʱ�������֧�����
    If Not MCPAR.����Ԥ���� And mEditType = EM_Balance_Charge Then
        If UpdateBalance(CCur(Val(lblʵ��.Caption)), cur����ͳ��, curȫ�Ը�, cur���Ը�) = False Then
            If vsFeeList.Enabled And vsFeeList.Visible Then vsFeeList.SetFocus
            Call SetButtons
            mblnNotClearLedDisplay = False
            Exit Sub
        End If
    End If
    If CheckFactValied(False, mblnPrintBill) = False Then
        mblnNotClearLedDisplay = False
        Call SetButtons: Exit Sub
    End If
    
    dtDate = zlDatabase.Currentdate
    If SaveData(strNos, str����IDs, str����IDs, dtDate, mblnPrintBill, strNo, curȫ�Ը�, cur���Ը�, cur����ͳ��) = False Then
        If Not gfrmMain Is Nothing Then
            Call zlExeBalanceWinRefrshData(strNo, False, dtDate)
        End If
        If vsFeeList.Enabled And vsFeeList.Visible Then vsFeeList.SetFocus
        Call SetButtons
        mblnNotClearLedDisplay = False
        Exit Sub
    End If
    If Not gfrmMain Is Nothing Then
      Call zlExeBalanceWinRefrshData(strNo, True, dtDate)
    End If
    mblnNotClearLedDisplay = False
End Sub

Private Sub PrintBill(ByVal strNo As String, ByVal dtDate As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ʊ�ݴ�ӡ
    '���:blnPrintBill-��Ʊ�Ƿ������ӡ
    '����:���˺�
    '����:2014-09-24 17:33:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotValiedNos As String
    Dim strReclaimInvoice As String '���յķ�Ʊ��
    Dim blnPrintBillEmpty As Boolean
    Dim blnVirtualPrint As Boolean
    Dim intPrint As Integer
    strNo = IIf(InStr(1, strNo, "'") = 0, "'" & strNo & "'", strNo)
    blnVirtualPrint = MCPAR.ҽ���ӿڴ�ӡƱ��
    If mblnPrintBill And Not (blnVirtualPrint And mstrYBPati <> "") Then
RePrint:
        strReclaimInvoice = ""
        Call frmReplenishTheBalancePrint.ReportPrint(1, strNo, mintInsure, mobjFactProperty, strReclaimInvoice, mlng����ID, txtInvoice.Text, dtDate, _
                blnVirtualPrint, , blnPrintBillEmpty)
        If Not (blnVirtualPrint And mstrYBPati <> "") Then
            If mobjFactProperty.�ϸ���� And blnPrintBillEmpty = False Then
                If zlIsNotSucceedPrintBill(1, strNo, strNotValiedNos) = True Then
                       If MsgBox("����[" & strNotValiedNos & "]Ʊ�ݴ�ӡδ�ɹ�,�Ƿ����½���Ʊ�ݴ�ӡ!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                End If
            End If
        End If
    End If
    
    '��ӡ�����嵥:�̶����ֱ��ӡ
    If zlStr.IsHavePrivs(mstrPrivs, "��������嵥") Then
        intPrint = Val(zlDatabase.GetPara("�����嵥��ӡ��ʽ", glngSys, mlngModule, "0"))
        If intPrint = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO=" & strNo, "ҩƷ��λ=" & IIf(mtyMoudlePara.blnҩ����λ, 1, 0), 2)
        ElseIf intPrint = 2 Then
            If MsgBox("Ҫ��ӡ������շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO=" & strNo, "ҩƷ��λ=" & IIf(mtyMoudlePara.blnҩ����λ, 1, 0), 2)
            End If
        End If
    End If
End Sub

Private Function SaveData(ByVal strNos As String, ByVal str����IDs As String, ByVal str����IDs As String, _
     ByVal dtDate As Date, ByVal blnPrintBill As Boolean, ByRef strNo As String, _
     Optional ByRef curȫ�Ը� As Currency, Optional ByRef cur���Ը� As Currency, _
     Optional ByRef cur����ͳ�� As Currency) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���浥��
    '���:str����IDs-���ر��ζ��ν���ķ��ý���IDs,����ö��ŷ���
    '     str����IDs-���ر��ζ��ν���ķ��ò��ֵĳ���IDs,����ö��ŷ���
    '     blnPrintBill-�Ƿ��ӡƱ��
    '     curȫ�Ը� -ȫ�Էѽ��
    '     cur���Ը�-���Ը����
    '     cur����ͳ��-ͳ����
    '����:strNO-��������㵥��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-17 11:42:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strAdvance As String, strDate As String, strFactNO As String
    Dim str������� As String, str����ID As String, str������� As String, str���ս��� As String
    Dim strSQL As String, strReclaimInvoice As String
    Dim cllIDs As Collection, cllPro As Collection
    Dim blnTrans  As Boolean, i As Long
    Dim varData As Variant
    Dim cur���� As Currency
    
    On Error GoTo errHandle
    
    If ShowReclaimInvoice(strNos, strReclaimInvoice) = False Then Exit Function
    str����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    strFactNO = Trim(txtInvoice.Text)
    strNo = zlDatabase.GetNextNo(13)    '�շѵ�
    str������� = "-" & str����ID
    str���ս��� = GetMedicareBalanceStr(cur����)
    str������� = str���ս���
    strTemp = str����IDs & IIf(str����IDs <> "", "," & str����IDs, "")
    
    Set cllPro = New Collection
    Set cllIDs = New Collection
    If zlCommFun.ActualLen(strTemp) <= 4000 Then
        cllIDs.Add strTemp
    Else
        varData = Split(strTemp, ",")
        strTemp = ""
        For i = 1 To UBound(varData)
            If zlCommFun.ActualLen(strTemp & "," & varData(i)) >= 4000 Then
                strTemp = Mid(strTemp & "," & varData(i), 2)
                cllIDs.Add strTemp
                strTemp = ""
            End If
            strTemp = strTemp & "," & varData(i)
        Next
        If strTemp <> "" Then
            strTemp = Mid(strTemp, 2)
            cllIDs.Add strTemp
        End If
    End If
    strDate = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    For i = 1 To cllIDs.Count
        'Zl_���ò����¼_������
        strSQL = "Zl_���ò����¼_������("
        '  No_In          In ���ò����¼.No%Type,
        strSQL = strSQL & "'" & strNo & "',"
        '  ʵ��Ʊ��_In    In ���ò����¼.ʵ��Ʊ��%Type,
        strSQL = strSQL & IIf(blnPrintBill, "'" & strFactNO & "'", "null") & ","
        '  ����id_In      In ���ò����¼.����id%Type,
        strSQL = strSQL & "" & str����ID & ","
        '  �������_In    In ����Ԥ����¼.�������%Type,
        strSQL = strSQL & "" & str������� & ","
        '  �շѽ���ids_In Varchar2,
        strSQL = strSQL & "'" & cllIDs(i) & "',"
        '  ҽ������_In    Varchar2,:��������,��ʽΪ:���㷽ʽ,������|.."
        strSQL = strSQL & "" & IIf(str���ս��� = "", "NULL", "'" & str���ս��� & "'") & ","
        '  ����Ա���_In  In ���ò����¼.����Ա���%Type,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '  ����Ա����_In  In ���ò����¼.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  �Ǽ�ʱ��_In    In ���ò����¼.�Ǽ�ʱ��%Type := Null,
        strSQL = strSQL & "to_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ��ע_In    In ���ò����¼.��ע%Type := Null,
        strSQL = strSQL & "'" & txtժҪ.Text & "',"
        '  ���ӱ�־_In    In ���ò����¼.��ע%Type := Null,
        strSQL = strSQL & "" & IIf(mEditType = EM_Balance_Register, 1, 0) & ","
        '  ����״̬_In    In ���ò����¼.����״̬%Type := 0
        strSQL = strSQL & "1)"
        zlAddArray cllPro, strSQL
        str���ս��� = ""
    Next
    '80944,Ƚ����,2014-12-18,��Ʊ�ݻ��ղ����ŵ�������ɺ�,ԭ��������������쳣,���Ȳ�����Ʊ��,������ɹ����ٽ��л���
'    If strReclaimInvoice <> "�޿���Ʊ��" Then
'        varData = Split(strNos, ",")
'        For i = 0 To UBound(varData)
'            'Zl_�����շѼ�¼_Reprint
'            strSQL = "zl_�����շѼ�¼_RePrint("
'            '  No_In         ������ü�¼.No%Type,
'            strSQL = strSQL & "'" & varData(i) & "',"
'            '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
'            strSQL = strSQL & "Null,"
'            '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
'            strSQL = strSQL & "0,"
'            '  ʹ����_In     Ʊ��ʹ����ϸ.ʹ����%Type,
'            strSQL = strSQL & "'" & UserInfo.���� & "',"
'            '  ʹ��ʱ��_In   Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
'            strSQL = strSQL & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),"
'            '  �˷�_In       Number := 0,
'            strSQL = strSQL & "0,"
'            '  Ʊ������_In   Number := 0,
'            strSQL = strSQL & "0,"
'            '  �ջ�Ʊ�ݺ�_In Varchar2 := Null,
'            strSQL = strSQL & "'" & strReclaimInvoice & "',"
'            '  Ʊ��_In Number:=1
'            strSQL = strSQL & "" & IIf(mEditType = EM_Balance_Register, 4, 1) & ")"
'            zlAddArray cllPro, strSQL
'        Next
'    End If
    If MCPAR.ҽ���ӿڴ�ӡƱ�� And MCPAR.ҽ������Ʊ�� = False Then
        '38821
        'Ʊ����������(��Ϊ����HIS�Ĵ�ӡ��ҽ���ӿڴ�ӡ����������Ʊ������)
        strSQL = "Zl_�������Ʊ��_Insert('" & strNo & "','" & strFactNO & "'," & ZVal(mlng����ID) & ",'" & UserInfo.���� & "'," & _
                  "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),0,1)"
        zlAddArray cllPro, strSQL
    End If
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If mEditType = EM_Balance_Register Then
        'strAdvance:����ģʽ|�Һŷ���ȡ��ʽ|�Һŵ���|�������־(1-������;0-��ͨ�ҺŽ���)
        strAdvance = "0|0|" & strNos & "|1"
        If Not gclsInsure.RegistSwap(Val(str����ID), cur����, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Function
        End If
        gcnOracle.CommitTrans
        Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistSwap, True, mintInsure)
    Else
        '���ý���ӿ�
        If zlInsureClinicSwap(strFactNO, str����ID, str�������, str�������, curȫ�Ը�, cur���Ը�, _
            cur����ͳ��) = False Then Exit Function
    End If
    
    blnTrans = False
    '��ʾ�˷ѽ��㴰��
    Call GetAsyncKeyState(VK_RETURN)
    If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
    Set mFrmBalanceWin = New frmReplenishTheBalanceWin
    If Not mFrmBalanceWin.zlChargeWin(Me, mEditType, mlngModule, mstrPrivs, mobjPatiInfor, mobjPayCards, strNo, dtDate, str����ID, _
        str�������, MCPAR.�ֱҴ���, strNos, strReclaimInvoice, mcllForceDelToCash, mstr�ų����㷽ʽ, , mEditType = EM_Balance_Register) Then Exit Function
    If Not gfrmMain Is Nothing Then SaveData = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlExeBalanceWinRefrshData(ByVal strNo As String, ByVal blnSaveOK As Boolean, ByVal dtDate As Date)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�н���������ˢ�²���
    '���:blnSaveOK-�Ƿ񱣴�ɹ�
    '����:���˺�
    '����:2014-09-26 10:42:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln���� As Boolean, i As Long, p As Long
    Dim blnGetFact As Boolean
    
    On Error GoTo errHandle
    
    If mEditType = EM_Balance_Err_Cancel Then
        If blnSaveOK Then mintSucces = mintInsure + 1
        Unload Me: Exit Sub
    End If
   
    If mEditType = EM_Balance_Err_ReCharge Then
        If blnSaveOK = False Then Exit Sub
        mintSucces = mintInsure + 1
        '��ӡ����
        Call PrintBill(strNo, dtDate)
        Unload Me: Exit Sub
    End If
    If blnSaveOK Then
        '���뵥����ʷ��¼(�������͵���)
        cboNO.AddItem strNo, 0
        For i = cboNO.ListCount - 1 To 10 Step -1
            cboNO.RemoveItem i 'ֻ��ʾ10��
        Next
        mintSucces = mintInsure + 1
        '��ӡ����
        Call PrintBill(strNo, dtDate)
        Call ReInitPatiInvoice
    End If
    SetPatientEnableModi True
    Call ClearData: Call SetButtons
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlInsureClinicSwap(ByVal strFactNO As String, _
    ByVal str����ID As String, ByVal str������� As String, ByVal strԤ������Ϣ As String, _
    Optional ByRef curȫ�Ը� As Currency, Optional ByRef cur���Ը� As Currency, Optional ByRef cur����ͳ�� As Currency) As Boolean
    '---------------------------------------------------------- -----------------------------------------------------------------------------------
    '����:ҽ������
    '���:strFactNo-��ǰ��Ʊ��
    '     str����ID-��ǰ�Ľ���ID
    '     str�������-��ǰ�Ľ������
    '     strԤ������Ϣ-���㷽ʽ|������||....
    '     curȫ�Ը� -ȫ�Էѽ��
    '     cur���Ը�-���Ը����
    '     cur����ͳ��-ͳ����
    '����:ҽ�����óɹ� ����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTransMedicare As Boolean
    Dim strAdvance As String, strSQL As String
    
    On Error GoTo errHandle
    If mintInsure = 0 Then Exit Function
    
    If MCPAR.ҽ���ӿڴ�ӡƱ�� And MCPAR.ҽ������Ʊ�� = False Then
        '���ϸ����Ʊ��ʱ���浱ǰƱ��
        If Not mobjFactProperty.�ϸ���� = False Then
            zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", strFactNO, glngSys, mlngModule
        End If
    End If
    strAdvance = str�������
    If Not gclsInsure.ClinicSwap(Val(str����ID), _
        GetMedicareBalanceSum(mstr�����ʻ�), GetMedicareBalanceSum("ҽ������"), _
        curȫ�Ը�, cur���Ը�, mintInsure, strAdvance) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    
    blnTransMedicare = True
    If strAdvance = str������� Then strAdvance = ""
     
    If strAdvance = "" Then
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mintInsure)
       zlInsureClinicSwap = True: Exit Function
    End If
    
    strԤ������Ϣ = Replace(Replace(strԤ������Ϣ, "|", "||"), ",", "|") 'ת��Ϊ�ָ�����ͬ���ַ���
    If Not zlInsureCheck(strԤ������Ϣ, strAdvance) Then
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mintInsure)
       zlInsureClinicSwap = True: Exit Function
    End If
    
    '��Ҫ����
    'Zl_���ò������_Modify
    strSQL = "Zl_���ò������_Modify("
    '  ��������_In   Number,
    '  --   0-��ͨ���㷽ʽ:
    '  --     ���㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    strSQL = strSQL & "" & 2 & ","
    '  ����id_In     In ���ò����¼.����id%Type,
    strSQL = strSQL & "" & str����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & strAdvance & "')"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ��ɽ���_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mintInsure)
    zlInsureClinicSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, False, mintInsure)
End Function

Private Function isValied(ByVal strNos As String, ByRef str����IDs As String, ByRef str����IDs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������е���Ч��
    '����:strNOs-���ν���ĵ��ݺ�
    '     str����IDs-���ر��ζ��ν���ķ��ý���IDs,����ö��ŷ���
    '     str����IDs-���ر��ζ��ν���ķ��ò��ֵĳ���IDs,����ö��ŷ���
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-17 10:46:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strTittle As String
    Dim int��¼���� As Integer
    
    On Error GoTo errHandle
    If Not CheckTextLength("ժҪ", txtժҪ) Then Exit Function
    strTittle = IIf(mEditType = EM_Balance_Register, "�Һ�", "�շ�")
    int��¼���� = IIf(mEditType = EM_Balance_Register, 4, 1)
    If strNos = "" Then
        ShowMsgbox "��ǰ����û����Ҫ��������" & strTittle & "���ã���ѡ����Ҫ��������" & strTittle & "���ã�"
        Exit Function
    End If
    
    If mEditType = EM_Balance_Register Then
        If MCPAR.�Һ�ʹ�ø����ʻ� Then
            If mstr�����ʻ� = "" Then
                ShowMsgbox "�Һų���δ���ø����ʻ����㣬�����ʻ�����֧����"
                Exit Function
            End If
        End If
    End If
    '���ѡ�񵥾����Ƿ�����Ѷ��ν����˵�
    strSQL = _
    " Select 1" & _
    " From ���ò����¼ A," & _
    "      (Select /*+Cardinality(b,10)*/ Distinct ����id" & _
    "       From ������ü�¼ A, Table(f_Str2list([1])) B" & _
    "       Where a.No = b.Column_Value And Mod(��¼����, 10)=[2]) B" & _
    " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And Nvl(����״̬,0) <> 2 And a.���ӱ�־=[3] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos, int��¼����, _
        IIf(mEditType = EM_Balance_Register, 1, 0))
    If Not rsTemp.EOF Then
        ShowMsgbox "��ѡ�񵥾��д����Ѳ�������˵����ݻ򲹳�����쳣���ݣ��������ٽ��в�����㣡"
        Exit Function
    End If
    
    strSQL = _
    " Select /*+Cardinality(b,10)*/ a.��¼����, a.����ID," & _
    "       Max(Decode(a.��¼״̬,2,a.����ID,0)) As ����ID " & _
    " From ������ü�¼ A, Table(f_Str2list([1])) B" & _
    " Where a.No = b.Column_Value And Mod(a.��¼����,10)=[2]" & _
    " Group By a.��¼����, a.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos, int��¼����)
    If rsTemp.EOF Then
        ShowMsgbox "��ǰ����û����Ҫ�������ķ��ã���ѡ����Ҫ�������ķ��ã�"
        Exit Function
    End If
    With rsTemp
        str����IDs = "": str����IDs = ""
        Do While Not .EOF
            If Val(Nvl(rsTemp!����ID)) = Val(Nvl(rsTemp!����ID)) Then
                str����IDs = str����IDs & "," & Val(Nvl(rsTemp!����ID))
            Else
                str����IDs = str����IDs & "," & Val(Nvl(rsTemp!����ID))
            End If
            .MoveNext
        Loop
        If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
        If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    End With
    If str����IDs = "" Then
        ShowMsgbox strTittle & "��Ϊ:" & strNos & "��δ�ҵ�ԭʼ��" & strTittle & "��¼�����������ҽ��������㣡"
        Exit Function
    End If

    strSQL = _
    " Select 1" & vbNewLine & _
    " From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
    " Where a.���㷽ʽ = b.����(+) And a.����id In (Select Column_Value From Table(f_Num2list([1])))" & vbNewLine & _
    "       And Decode(Mod(a.��¼����,10),1,0,Decode(b.����,3,1,4,1,0)) = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs)
    If rsTemp.EOF = False Then
        ShowMsgbox strTittle & "��Ϊ:" & strNos & "��" & strTittle & "�����д���ҽ����������ݣ����������ҽ��������㣡"
        Exit Function
    End If
    
    strSQL = _
    " Select 1" & vbNewLine & _
    " From ����Ԥ����¼ A" & vbNewLine & _
    " Where a.����id In (Select Column_Value From Table(f_Num2list([1])))" & vbNewLine & _
    "       And Not Exists(Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� = 9)" & vbNewLine & _
    " Having Count(Distinct Decode(Mod(a.��¼����,10),1,'��Ԥ���',a.���㷽ʽ)) > 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs)
    If rsTemp.EOF = False Then
        ShowMsgbox strTittle & "��Ϊ:" & strNos & "��" & strTittle & "�����д����������ϵĽ��㷽ʽ�����������ҽ��������㣡"
        Exit Function
    End If
    
    '���������㷽ʽ��Ч�Լ��
    If ThreeBalanceCheck(mobjPayCards, mEditType = EM_Balance_Register, strNos, mcllForceDelToCash, mstr�ų����㷽ʽ) = False Then Exit Function

    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ThreeBalanceCheck(objCards As Cards, ByVal blnIsRegister As Boolean, ByVal strNos As String, _
    ByRef cllForceDelToCash As Collection, ByRef str�ų����㷽ʽ As String) As Boolean
    '���������㷽ʽ��Ч�Լ��
    '��Σ�
    '   objCards ������������Ч��֧����ʽ
    '   blnIsRegister �Ƿ�Һŵ�
    '   strNos ����ѡ�񲹳����ĵ��ݺ�
    '���Σ�
    '   cllForceDelToCash ǿ��������Ϣ��Array(����Ա,���������,���㷽ʽ)
    '   str�ų����㷽ʽ �ų����㷽ʽ,����ö��ŷָ�
    '���أ����ͨ��������True�����򣬷���False
    '105432
    Dim objCard As Card
    Dim cllFeeBalance As New Collection, i As Integer
    Dim blnFind As Boolean, blnQuestion As Boolean
    Dim str����Ա As String, strKey As String
    Dim dblMoney  As Double
    Dim j As Integer, lngCount As Long
    Dim varData As Variant
    Dim rsBalance As ADODB.Recordset
    
    On Error GoTo errHandler
    Set cllForceDelToCash = New Collection
    str�ų����㷽ʽ = ""
    Set rsBalance = zlFromIDGetChargeBalance(2, strNos, , , , IIf(blnIsRegister, 4, 1))
    If rsBalance Is Nothing Then ThreeBalanceCheck = True: Exit Function
    
    rsBalance.Filter = "����=3"
    'ȥ��
    With rsBalance
        Do While Not .EOF
            strKey = "_" & Val(Nvl(!�����ID))
            If CollectionExitsValue(cllFeeBalance, strKey) Then
                dblMoney = cllFeeBalance(strKey)(4) + Val(Nvl(!��Ԥ��))
                cllFeeBalance.Remove strKey
            Else
                dblMoney = Val(Nvl(!��Ԥ��))
            End If
            If RoundEx(dblMoney, 6) > 0 Then 'ȫ������ľͲ��ټ���
                'Array(���㷽ʽ,�����ID,�Ƿ�����,���������,��Ԥ��,�Ƿ�ȫ��,�Ƿ�ת�ʼ�����)
                cllFeeBalance.Add Array(Nvl(!���㷽ʽ), Val(Nvl(!�����ID)), Val(Nvl(!�Ƿ�����)), _
                    Nvl(!���������), dblMoney, Val(Nvl(!�Ƿ�ȫ��)), Nvl(!�Ƿ�ת�ʼ�����)), strKey
            End If
            .MoveNext
        Loop
    End With
    If cllFeeBalance.Count = 0 Then ThreeBalanceCheck = True: Exit Function
    
    For i = 1 To cllFeeBalance.Count
        blnQuestion = False
        'ҽ�ƿ����
        If objCards Is Nothing Then
            If MsgBox("��" & cllFeeBalance(i)(3) & "��δ���ã���ҽ�ƿ�֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnQuestion = True
        Else
            blnFind = False
            For Each objCard In objCards
                If objCard.�ӿ���� = cllFeeBalance(i)(1) Then blnFind = True: Exit For
            Next
            If blnFind = False Then
                If MsgBox("��" & cllFeeBalance(i)(3) & "��δ���ã���ҽ�ƿ�֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnQuestion = True
            End If
        End If
        
        If blnQuestion Then
            If cllFeeBalance(i)(2) = 0 Then 'ǿ������
                If str����Ա = "" Then '���ֿ����ʱֻ��֤һ��
                    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
                        str����Ա = UserInfo.����
                    Else
                        str����Ա = zlDatabase.UserIdentifyByUser(Me, "ҽ�ƿ���" & cllFeeBalance(i)(3) & "��ǿ�����֣�Ȩ����֤��", _
                            glngSys, mlngModule, "�����˿�ǿ������", , True)
                        If str����Ա = "" Then Exit Function
                    End If
                End If
                'Array(����Ա,���������,���㷽ʽ)
                cllForceDelToCash.Add Array(str����Ա, cllFeeBalance(i)(3), cllFeeBalance(i)(0))
            End If
        ElseIf cllFeeBalance(i)(5) = 1 Then '����ȫ��
            If cllFeeBalance(i)(2) = 1 Then '�������֣�����ȫ��
                If cllFeeBalance(i)(6) = 0 Then '��֧��ת�ʼ�����
                    If MsgBox("��" & cllFeeBalance(i)(3) & "������ȫ�ˣ���˲����˻�ԭ����" & _
                        "���������������ô��ҽ�ƿ�֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    str�ų����㷽ʽ = str�ų����㷽ʽ & "," & cllFeeBalance(i)(0)
                End If
            ElseIf cllFeeBalance(i)(6) = 0 Then '���������֣�����ȫ�ˣ��Ҳ�֧��ת�ʼ�����
                If MsgBox("��" & cllFeeBalance(i)(3) & "������ȫ���Ҳ������֣�ͬʱҲ��֧��ת�ʼ����ۣ�����޷��˻�ԭ����" & _
                    "���������������ô��ҽ�ƿ�֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                If str����Ա = "" Then '���ֿ����ʱֻ��֤һ��
                    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
                        str����Ա = UserInfo.����
                    Else
                        str����Ա = zlDatabase.UserIdentifyByUser(Me, "��" & cllFeeBalance(i)(3) & "��ǿ�����֣�Ȩ����֤��", _
                            glngSys, mlngModule, "�����˿�ǿ������", , True)
                        If str����Ա = "" Then Exit Function
                    End If
                End If
                'Array(����Ա,���������,���㷽ʽ)
                cllForceDelToCash.Add Array(str����Ա, cllFeeBalance(i)(3), cllFeeBalance(i)(0))
                str�ų����㷽ʽ = str�ų����㷽ʽ & "," & cllFeeBalance(i)(0)
            End If
        End If
    Next
    If str�ų����㷽ʽ <> "" Then str�ų����㷽ʽ = Mid(str�ų����㷽ʽ, 2)
    

    If str�ų����㷽ʽ = "" Then ThreeBalanceCheck = True: Exit Function
    '�ж��Ƿ�����Ч�Ľ��㷽ʽ
    varData = Split(str�ų����㷽ʽ, ",")
    lngCount = mobjPayCards.Count
    For i = 1 To mobjPayCards.Count
        If mobjPayCards(i).�ӿ���� <= 0 Or mobjPayCards(i).�ӿ���� > 0 And mobjPayCards(i).���ѿ� Then
            Exit For
        End If
        
        blnFind = False
        For j = 0 To UBound(varData)
            If mobjPayCards(i).���㷽ʽ = varData(j) Then
                lngCount = lngCount - 1: blnFind = True
            End If
        Next
        If blnFind = False Then Exit For
    Next
    If lngCount <= 0 Then
        MsgBox "�ų�ǿ�����ֵĽ��㷽ʽ����û�п��õĽ��㷽ʽ�����ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    ThreeBalanceCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdSelAll_Click()
    Call FromNosSel("", True, False, True)
    Call SetButtons
    vsDiagnose.Cell(flexcpChecked, 0, 0, vsDiagnose.Rows - 1, vsDiagnose.COLS - 1) = 1
End Sub

Private Sub cmdԤ����_Click()
    Dim strNos As String, strNone As String
    Dim strAdvance As String
    
    If mintInsure = 0 Then
        MsgBox "δ����ҽ�������֤,������Ԥ����!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    strNos = GetSelFeeNos   '��ǰѡ�еĵ���
    If strNos = "" Then
        MsgBox "δѡ����ҪԤ��ķ��õ���,������Ԥ����", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If MCPAR.ʵʱ��� Then
        '1.���뵥�ݣ�2.�޸ĵ��ݣ�3.������ҩ�䷽��4.�޸���ҩ�����������еĸ���ͬʱ�仯��5.��������Զ���������Լ�������ܼ����ۿ�
        '6.�޸ĵ��ۣ�7.����ִ�п��ң�ҩƷ�۸����㣬8.�����ѱ�ʵ�ս������,9.�����������֤ҽ�����,�����ȵ�
        If gclsInsure.CheckItem(mintInsure, 0, 9, MakeDetailRecord(strNos), strAdvance) = False Then Exit Sub
    End If
    
    cmdԤ����.Enabled = False
    'Ԥ����
    If Not ����Ԥ����(strNos, strNone) Then
        If strNone <> "" Then
            MsgBox "��ǰ���ս���ʹ�õĽ��㷽ʽ" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                "������δ���ã����ȵ����㷽ʽ������������Щ���㷽ʽ��", vbInformation, gstrSysName
        End If
        cmdԤ����.TabStop = True: cmdOK.Enabled = False: cmdԤ����.Enabled = True
        If cmdԤ����.Enabled And cmdԤ����.Visible Then cmdԤ����.SetFocus
        mblnEdit = True
        Exit Sub
    End If
    mblnEdit = False
    Call SetButtons
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
End Sub

Private Function ����Ԥ����(ByVal strNos As String, ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ����
    '���:strNos-����Ԥ����ĵ��ݺ�
    '����:strNone-���ز����ڵĽ��㷽ʽ
    '����:Ԥ����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-16 17:30:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���ʺϼ� As Double, dblMoney As Double
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim strDate As String, str���㷽ʽ As String
    Dim dbl�ϼ� As Double
    
    strNone = ""
    
    Screen.MousePointer = 11
    On Error GoTo errH
    '��ʼ�����������
    Call InitBalanceGrid
    
    '��ȡ����ʱ��
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If zlInsureClinicPreSwap(strNos, strDate, strNone) = False Then Exit Function
    'Ҫ�������Ա������ط�ʶ��
    If cmdԤ����.Visible Then
        cmdԤ����.TabStop = False
        cmdOK.Enabled = True
    End If
    
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            dblMoney = dblMoney + Val(.TextMatrix(0, i + 1))
        Next
        txt�˿�ϼ�.Text = Format(dblMoney, "0.00")
    End With
    
    Call zl9InsureLedSpeak
    strNone = Mid(strNone, 2)
    If strNone = "" Then ����Ԥ���� = True
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSelFeeNos() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡѡ��ķ��õ��ݺ�
    '����:����ö��ŷ���
    '����:���˺�
    '����:2014-09-16 17:24:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, i As Long
    
    With vsFeeList
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then
                If Trim(.TextMatrix(i, .ColIndex("NO"))) <> "" _
                    And Abs(Val(.Cell(flexcpChecked, i, .ColIndex("ѡ��")))) = 1 Then
                    strNos = strNos & "," & .TextMatrix(i, .ColIndex("NO"))
                End If
            End If
        Next
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetSelFeeNos = strNos
End Function

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.staThis.Visible Then Bottom = Me.staThis.Height
    staThis.Top = Me.ScaleHeight - Me.staThis.Height
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnNotClearLedDisplay = False
    
    If mblnUnLoad Then Unload Me: Exit Sub
    mblnFirst = False
    Call reSizeWinControl
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
    If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then cmdOK.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '������Shift=-1����ʾ�ǳ���ǿ���ڵ���
    Select Case KeyCode
        Case vbKeyF1  '����
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is PatiIdentify Then
                If mobjPatiInfor Is Nothing Then
                    If MCPatientProcess(mobjPatiInfor) = False Then Exit Sub
                End If
            End If
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF5
            If cmdԤ����.Visible And cmdԤ����.Enabled Then cmdԤ����.SetFocus: cmdԤ����_Click
        Case vbKeyF6 '��λ�����������
            If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Case vbKeyF8 '�˷Ѵ���
            If cmdDelete.Visible Then Call cmdDelete_Click
        Case vbKeyF9 '��λ�����ݺ������
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Case vbKeyEscape
            cmdCancel.SetFocus: Call cmdCancel_Click
        Case 191 '"?"������
            If Shift = vbAltMask Then
                Call staThis_PanelClick(staThis.Panels("Calc"))
            End If
    End Select
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdSelAll.Visible Then Call cmdSelAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdClear.Visible Then Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Set mcolBalance = New Collection
    Call InitFace
    
    If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then
        mblnUnLoad = Not LoadErrBillData(mobjPatiInfor)
    End If
    Call SetControlEnabled
    
    RestoreWinState Me, App.ProductName, mstrTittle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTittle
    If gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset msCommSpeak
    End If
    PatiIdentify.AllowAutoCommCard = False
    PatiIdentify.AllowAutoICCard = False
    PatiIdentify.AllowAutoIDCard = False
    If Not mcllForceDelToCash Is Nothing Then Set mcllForceDelToCash = Nothing
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim bln�շѵ� As Boolean, strNo As String
    Dim lng����ID As Long, str�������� As String
    
    strNo = Trim(PatiIdentify.Text)
    If Left(PatiIdentify.Text, 1) = "." Then bln�շѵ� = True: strNo = Mid(strNo, 2)
    Set mobjPatiInfor = New zlIDKind.PatiInfor
    
    If strNo = "" Then
        If CheckPatiInfor(objCardData) = False Then blnCancel = True: Exit Sub
        Exit Sub
    End If
    
    If objCard.���� Like "*��*��*" And Not blnCard And InStr("-*+/.", Left(Trim(PatiIdentify.Text), 1)) = 0 Then
        Dim strPati As String, vRect As RECT, rsTmp As ADODB.Recordset
        If Not gblnSeekName Then
            blnCancel = True: Exit Sub
        Else
             '�����:50485
            strPati = _
                " Select /*+Rule */distinct 1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ,decode(b.����,Null,Null,'��') As �Ƿ���ҽ�ƿ�" & _
                " From ������Ϣ A, ����ҽ�ƿ���Ϣ B " & _
                " Where Rownum <101 And a.����ID=b.����ID(+) And b.״̬(+)=0 And B.�����ID(+)=[3]  And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & _
                IIf(gintNameDays = 0, "", " And Nvl(A.����ʱ��,A.�Ǽ�ʱ��)>Trunc(Sysdate-[2])")
                
            strPati = strPati & " Order by ����ID,����"
                
            vRect = zlControl.GetControlRect(PatiIdentify.hWnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, PatiIdentify.Height, blnCancel, False, True, strNo & "%", gintNameDays, Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, 0)), "bytSize=1")
            If Not rsTmp Is Nothing Then
                If Nvl(rsTmp!ID) = 0 Then '�����²���
                    blnCancel = True: Exit Sub
                Else '�Բ���ID��ȡ
                    lng����ID = Nvl(rsTmp!ID)
                End If
            Else 'ȡ��ѡ��
                blnCancel = True: Exit Sub
            End If
        End If
    Else
        If Not bln�շѵ� Then
            If objCard.�ӿ���� > 0 Then Exit Sub
            If objCard.���� <> "�շѵ��ݺ�" Then Exit Sub
        End If
        strNo = zlCommFun.GetFullNO(strNo)
        If GetBillNoFromPati(strNo, lng����ID) = False Then
            MsgBox "δ�ҵ���Ӧ��" & IIf(mEditType = EM_Balance_Register, "�Һ�", "�շ�") & "����:" & strNo & ",��������ĵ����Ƿ���ȷ!", vbInformation + vbOKOnly, gstrSysName
            blnCancel = True: Exit Sub
        End If
        If lng����ID = 0 Then
            MsgBox "��Ӧ��" & IIf(mEditType = EM_Balance_Register, "�Һ�", "�շ�") & "����:" & strNo & "���ǽ������˵�" & IIf(mEditType = EM_Balance_Register, "�Һ�", "�շ�") & "��,���ܽ���ҽ��������!", vbInformation + vbOKOnly, gstrSysName
            blnCancel = True: Exit Sub
        End If
    End If
    Set mobjPatiInfor = Nothing
    If zlGetPati(lng����ID, objCardData, str��������) = False Then blnCancel = True: Exit Sub
    strShowText = objCardData.����
    Set mobjPatiInfor = objCardData
    If CheckPatiInfor(objCardData) = False Then blnCancel = True: Exit Sub
    blnFindPatied = True
    If vsFeeList.Enabled And vsFeeList.Visible Then
        vsFeeList.SetFocus
    ElseIf cmdԤ����.Enabled And cmdԤ����.Visible Then
        cmdԤ����.SetFocus
    ElseIf cmdOK.Enabled And cmdOK.Visible Then
        cmdԤ����.SetFocus
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    '�ҵ�����ʱ,�Բ���ID��Ϊ�ж�����
    If objHisPati Is Nothing Then blnCancel = True: Exit Sub
    
    If objHisPati.����ID = 0 Then blnCancel = True: Exit Sub
    Set mobjPatiInfor = objHisPati
    PatiIdentify.Text = mobjPatiInfor.����
    
    If CheckPatiInfor(objHisPati) = False Then blnCancel = True: Exit Sub
    Set objCardData = mobjPatiInfor
    Call SetButtons
    If vsFeeList.Enabled And vsFeeList.Visible Then
        vsFeeList.SetFocus
    ElseIf cmdԤ����.Enabled And cmdԤ����.Visible Then
        cmdԤ����.SetFocus
    ElseIf cmdOK.Enabled And cmdOK.Visible Then
        cmdԤ����.SetFocus
    End If
End Sub

Private Function CheckPatiInfor(ByRef objPatiInfor As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤������Ϣ
    '���:objPatiInfor-��ǰ������Ϣ
    '����:
    '����:��֤�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-16 15:10:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    '�Ƚ���ҽ�������֤
    mblnNotClearLedDisplay = False
    If MCPatientProcess(objPatiInfor) = False Then GoTo GoClear
    '���ط�����Ϣ
    If ReadBills(objPatiInfor) = False Then GoTo GoClear
    Call SetButtons '���ð�ť
    CheckPatiInfor = True
    Exit Function
    
GoClear:
    SetPatientEnableModi True
    ClearData
End Function

Private Function LoadErrBillData(ByRef objPati As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص�����
    '���:strNO-���ݺ�
    '����:
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-10 12:48:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strWhere1 As String, strWithTable As String
    Dim strSQL As String, j As Long
    Dim strTable As String, strFields As String, str�������� As String
    Dim rsTemp As ADODB.Recordset
    Dim dblMoney As Double
    Dim dbl������ As Double
    Dim str���㷽ʽ As String
    Dim strTemp As String
    
    On Error GoTo errHandle
    If Not (mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge) Then Exit Function
    
    strSQL = " Select NO,�������,��ע From ���ò����¼  A   Where a.����ID =[1] and rownum <2"
    If mstr����ID = "" Then mstr����ID = "0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����ID)
    If rsTemp.EOF Then
        MsgBox "δ�ҵ���Ҫ" & IIf(mEditType = EM_Balance_Err_ReCharge, "���½���", "����") & "���쳣���ϼ�¼��", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    mstrNo = Nvl(rsTemp!NO): mstr������� = "" & Val(Nvl(rsTemp!�������))
    cboNO.Text = mstrNo: txtժҪ.Text = Nvl(rsTemp!��ע)
    cboNO.Locked = True: txtժҪ.Enabled = False
    
    strWhere = "": strFields = ",'' as ���": strTable = ""
    strTable = "," & _
    "         ( Select distinct 1 as ��¼����, A1.NO, f_List2str(Cast(COLLECT(distinct Q.������� ) as t_Strlist))  as ���" & _
    "           From  (Select distinct NO,ҽ����� From ������ü�¼ A,�շѵ��� N1��where mod(a.��¼����,10)=1 And a.��¼״̬ in (1,3) ANd A.����ID=N1.�շѽ���ID) A1, " & _
    "               ����ҽ����¼ H,�������ҽ�� J,������ϼ�¼ Q  " & _
    "           Where   A1.ҽ�����=H.ID and Nvl(H.���ID,H.ID)=J.ҽ��ID and J.���ID=Q.ID " & _
    "           Group by  A1.NO ) C " & vbNewLine
    
    strFields = ",Max(C.���) as ���"
    strWithTable = "" & _
    "   With �շѵ��� as ( " & _
    "       Select  Distinct A.�շѽ���ID  From ���ò����¼  A   Where a.������� =[1] )"
    strSQL = "" & strWithTable & vbCrLf & _
    "    Select A.��¼����,A.NO,A.��¼״̬,Nvl(A.�۸񸸺�,A.���) as ���,A.��������,A.��������ID,A.ִ�в���ID,A.�շ����,A.�ѱ�,A.�շ�ϸĿID," & _
    "          A.�������� ,max(A.������) as ������,A.���㵥λ,max(A.ҽ�����) as ҽ�����," & _
    "          Avg(Nvl(A.����,1)) as ����,Avg(A.����) as ����," & _
    "          Sum(A.��׼����) as ����, Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
    "          Max(Decode(a.��¼״̬, 2, '', a.����Ա����)) As ����Ա����, Max(Decode(a.��¼״̬, 2, To_Date('1900-01-01', 'YYYY-MM-DD'),a.�Ǽ�ʱ��)) As �Ǽ�ʱ��," & _
    "          Max(A.ժҪ)  as ժҪ,A.����ID,max(A.����ID) as ����ID" & strFields & _
    "   From ������ü�¼ A,�շѵ��� b " & strTable & _
    "   Where  A.����ID=B.�շѽ���ID " & _
    "          And a.��¼���� = c.��¼����(+) And a.No = c.No(+)" & _
    "   Group by  a.No, a.��¼����, a.����id, a.��¼״̬, Nvl(a.�۸񸸺�, a.���), a.��������, a.��������id, a.ִ�в���id, a.�շ����, a.�ѱ�, a.�շ�ϸĿid,  a.��������, a.���㵥λ, a.����id"
              
    strSQL = _
    " Select Decode(A.��¼����,1,'�շ�',4,'�Һ�','�շ�') as ����, A.NO,A.���,A.��������,A.�ѱ�,a.������,A.�շ�ϸĿID,C.���� as �շ����, " & _
    "       -1 as ѡ��,C.���� as ���,B.����, " & _
    "       Nvl(M1.����,B.����) as ����,E1.���� as ��Ʒ�� ,B.���,Max(Nvl(A.��������,B.��������)) ��������," & _
            IIf(mtyMoudlePara.blnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as  ��λ," & _
    "       Max(A.ҽ�����) as ҽ�����,sum(A.����) as ����," & _
    "       sum(A.����" & IIf(mtyMoudlePara.blnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
    "       Max(A.����" & IIf(mtyMoudlePara.blnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
    "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
    "       D.���� as ִ�п���,E.���� as ��������,Max(a.����Ա����) As ����Ա����, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, " & _
    "       Max(A.ժҪ) as ժҪ,'' as ���㷽ʽ,max(A.���) as ���,A.��¼����,max(A.����ID) as ����ID" & _
    " From (" & strSQL & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E,ҩƷ��� X," & _
    "       �շ���Ŀ���� M1,�շ���Ŀ���� E1" & _
    " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.�շ�ϸĿID=X.ҩƷID(+)" & _
    "       And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID(+) " & _
    "       And A.�շ�ϸĿID=M1.�շ�ϸĿID(+) And M1.����(+)=1 And M1.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
    " Group by A.��¼����,A.NO,A.���,A.��������,A.�ѱ�,a.������,A.�շ�ϸĿID,C.����,C.����,B.����,Nvl(M1.����,B.����)," & _
    "       E1.����,B.���,A.���㵥λ,D.����,E.����,X.ҩƷID,X." & gstrҩ����λ & _
    " Having Sum(A.����)<>0 " & _
    " Order by �Ǽ�ʱ��,NO,���"
    
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�������)
    If mrsList.RecordCount = 0 Then
        MsgBox "δ�ҵ���Ҫ" & IIf(mEditType = EM_Balance_Err_ReCharge, "���½���", "����") & "���쳣���ϼ�¼��", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If zlGetPati(Val(Nvl(mrsList!����ID)), mobjPatiInfor, str��������) = False Then Exit Function
    
    mintInsure = GetBalanceInsure(mstr����ID, str��������)
    txtYB.Text = mintInsure
    mrsList.Filter = ""
    Call LoadFeeData(mrsList)
    
    lbl����.Caption = str��������
    PatiIdentify.Text = mobjPatiInfor.����
    PatiIdentify.PasswordChar = ""
    PatiIdentify.IMEMode = 0
    lblPatiInfor.Caption = "�Ա�:" & mobjPatiInfor.�Ա�
    lblPatiInfor.Caption = lblPatiInfor.Caption & Space(4) & "����:" & mobjPatiInfor.����
    lblPatiInfor.Caption = lblPatiInfor.Caption & Space(4) & "���ʽ:" & mobjPatiInfor.ҽ�Ƹ��ʽ
    initInsurePara mobjPatiInfor.����ID
    
    strSQL = _
    " Select Decode(Mod(a.��¼����,10),1,'��Ԥ���',a.���㷽ʽ) As ���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��" & vbNewLine & _
    " From ����Ԥ����¼ A" & vbNewLine & _
    " Where a.������� = [1]" & vbNewLine & _
    " Group By Decode(Mod(a.��¼����,10),1, '��Ԥ���',a.���㷽ʽ)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�������)

     '����Ԥ���������ý��㼯
    With vsBalance
        .Clear 1
        .Rows = 1
        .COLS = 1
        .TextMatrix(0, 0) = "ҽ������"
        
        Do While Not rsTemp.EOF
            '������ʽ;���;�Ƿ������޸�
            str���㷽ʽ = Nvl(rsTemp!���㷽ʽ, "δ����")
            dbl������ = Val(Nvl(rsTemp!��Ԥ��))
            dblMoney = dblMoney + dbl������
            .COLS = .COLS + 2
            .TextMatrix(0, .COLS - 2) = str���㷽ʽ
            .TextMatrix(0, .COLS - 1) = Format(dbl������, "0.00")
            .Cell(flexcpData, 0, .COLS - 1) = dbl������
            .ColData(.COLS - 1) = 0 '�Ƿ������޸�
            .ColData(.COLS - 2) = 0
            
            '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;�ĺ���
            strTemp = str���㷽ʽ
            strTemp = strTemp & ";" & dbl������
            strTemp = strTemp & ";" & 0
            strTemp = strTemp & ";" & dbl������
            mcolBalance.Add strTemp
            rsTemp.MoveNext
        Loop
        .TabStop = False
    End With
    txt�˿�ϼ�.Text = Format(dblMoney, "0.00")
    Call ReInitPatiInvoice(True, mintInsure, mobjPatiInfor.����ID)
    LoadErrBillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReadBills(ByRef objPati As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص�����
    '���:strNO-���ݺ�
    '     blnFilter-�Ƿ�����ɸѡ
    '����:
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-10 12:48:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strWhere1 As String, strWithTable As String
    Dim strSQL As String, j As Long
    Dim strTable As String, strFields As String
    Dim strTemp As String
    Dim strBalance As String, blnFind As Boolean
    Dim rsTemp As ADODB.Recordset, varData As Variant, i As Long
    Dim strAllNOs As String
    
    On Error GoTo errHandle
    If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then
       '�쳣���ݵĴ���
       ReadBills = LoadErrBillData(objPati)
       Exit Function
    End If
    If objPati.����ID = 0 Then Exit Function
    
    If mEditType = EM_Balance_Register Then
        strWhere = " And a.����ID=[1] And a.��¼���� =4 And a.��¼״̬ In(1, 3)"
    Else
        strWhere = " And a.����ID=[1] And a.��¼���� =1 And a.��¼״̬ In(1, 3)"
    End If
    
    ' --�����:79396,�Һŵ�����������һ���֣��ű�����ѣ�����������
    ' --�����:112811,"�������˹ҺŴ�Ϊ���۵�"ʱ���ܶԸùҺŵ��ݽ���ҽ���������
    If mEditType = EM_Balance_Register Then
        strWhere = strWhere & _
            " And Not Exists (Select 1 From ������ü�¼ Where No = a.NO And ��¼���� = 4 And ��¼״̬ = 2)" & _
            " And Nvl(a.ժҪ,'-') Not Like '����:%'"
    End If
    ' --ֻ�ܶ����õġ�������㷽ʽ�����в�����
    If mtyMoudlePara.str�����������շѷ�ʽ <> "" Then
        strWhere = strWhere & " And Instr('|'||[3]||'|', '|'||Decode(Mod(b.��¼����,10),1,'��Ԥ���',b.���㷽ʽ)||'|') > 0"
    End If
    ' --�������Ѿ����ν����,���������ν��������˵�
    ' --���ܰ���ҽ�������˵�
    ' --���������ѿ������
    ' --�ų����Ѻ�ֻ����һ�ֽ��㷽ʽ
    ' --��ʣ����Ŀ�ĵ��ݲ����в�����
    strWhere = strWhere & _
    " And Not Exists(Select 1 From ���ò����¼ Where �շѽ���id = a.����id And Nvl(����״̬, 0) <> 2)" & vbNewLine & _
    " And Not Exists(Select 1 From ���ս����¼ Where ���� = 1 And ��¼id = a.����id)" & vbNewLine & _
    " And (Mod(b.��¼����, 10) = 1 Or b.���㿨��� Is Null)" & vbNewLine & _
    " And Exists(Select 1" & vbNewLine & _
    "            From ����Ԥ����¼ F" & vbNewLine & _
    "            Where f.����id = b.����id" & vbNewLine & _
    "                  And Not Exists(Select 1 From ���㷽ʽ Where ���� = f.���㷽ʽ And ���� = 9)" & vbNewLine & _
    "            Having Count(Distinct Decode(Mod(��¼����,10),1,'��Ԥ���',���㷽ʽ)) = 1)" & vbNewLine & _
    " And Exists(Select 1" & vbNewLine & _
    "            From ������ü�¼" & vbNewLine & _
    "            Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.��� And �۸񸸺� Is Null" & vbNewLine & _
    "            Having Sum(Nvl(����,1)*����) <> 0)"

    strWithTable = _
    "With �շѵ��� As(" & _
    "    Select a.��¼����, a.No, Max(Decode(Mod(b.��¼����,10),1,'��Ԥ���',b.���㷽ʽ)) As ���㷽ʽ" & vbNewLine & _
    "    From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
    "    Where a.����id = b.����id" & vbNewLine & _
    "          And Decode(a.��¼����,1,a.�Ǽ�ʱ��,a.����ʱ��) Between Trunc(Sysdate)-[2] And Trunc(Sysdate)+1-1/24/60/60" & vbNewLine & _
    "          And a.����id = [1]" & vbNewLine & _
               strWhere & vbNewLine & _
    "    Group By a.��¼����, a.No)"
    
    If mEditType = EM_Balance_Charge Then
       strWithTable = strWithTable & "," & vbNewLine & _
        "ҽ����� As (" & _
        "    Select Distinct 1 As ��¼����, A2.No, f_List2str(Cast(Collect(Distinct c.�������) As t_Strlist)) As ���" & _
        "    From(Select Distinct 1 As ��¼����, B1.No, B1.ҽ�����" & _
        "         From ������ü�¼ B1, �շѵ��� A1" & _
        "         Where B1.No = A1.No And A1.��¼���� = 1 And B1.��¼���� = 1 And B1.��¼״̬ In (1, 3)" & _
        "        ) A2, ����ҽ����¼ A, �������ҽ�� B, ������ϼ�¼ C" & _
        "    Where A2.��¼���� = 1 And A2.ҽ����� = a.Id And Nvl(a.���Id,a.Id) = b.ҽ��id And b.���id = c.Id" & _
        "    Group By A2.No)"
        
        strTable = ",ҽ����� C"
        strFields = ",Max(C.���) as ��� "
        strWhere1 = " And A.NO=C.NO(+)"
    Else
        strTable = ""
        strFields = ",'' as ���"
        strWhere1 = ""
    End If
    
    strSQL = strWithTable & vbNewLine & _
    " Select A.��¼����,A.NO,A.��¼״̬,Nvl(A.�۸񸸺�,A.���) as ���,A.��������,A.��������ID,A.ִ�в���ID," & _
    "       A.�շ����,A.�ѱ�,A.�շ�ϸĿID, A.�������� ,max(A.������) as ������,A.���㵥λ," & _
    "       Max(A.ҽ�����) as ҽ�����,Avg(Nvl(A.����,1)) as ����,Avg(A.����) as ����," & _
    "       Sum(A.��׼����) as ����, Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
    "       Max(Decode(a.��¼״̬, 2, '', a.����Ա����)) As ����Ա����, " & _
    "       Max(Decode(a.��¼״̬, 2, To_Date('1900-01-01', 'YYYY-MM-DD'), a.�Ǽ�ʱ��)) As �Ǽ�ʱ��," & _
    "       Max(A.ժҪ)  as ժҪ,A.����ID,max(A.����ID) as ����ID,max(B.���㷽ʽ) as ���㷽ʽ" & strFields & _
    " From ������ü�¼ A,�շѵ��� B " & strTable & _
    " Where A.��¼���� IN(1,4) And A.��¼����=b.��¼���� And a.No=b.No " & strWhere1 & _
    " Group By a.��¼����, a.��¼״̬, A.NO,Nvl(A.�۸񸸺�,A.���),A.��������,A.��������ID,A.ִ�в���ID," & _
    "       A.�շ����,A.�ѱ�,A.�շ�ϸĿID,A.��������,A.���㵥λ,A.����ID"
    
    strSQL = _
    " Select Decode(A.��¼����,1,'�շ�',4,'�Һ�','�շ�') as ����, A.NO,A.���,A.��������,A.�ѱ�,a.������,A.�շ�ϸĿID,C.���� as �շ����, " & _
    "       " & IIf(mEditType = EM_Balance_Charge, -1, 2) & " as ѡ��,C.���� as ���,B.����, " & _
    "       Nvl(M1.����,B.����) as ����,E1.���� as ��Ʒ�� ,B.���,Max(Nvl(A.��������,B.��������)) ��������," & _
            IIf(mtyMoudlePara.blnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as  ��λ," & _
    "       Max(A.ҽ�����) as ҽ�����,Avg(Decode(a.��¼״̬, 1, a.����, 1)) As ����, " & _
    "       Sum(Decode(a.��¼״̬, 1, 1, a.����) * a.����" & IIf(mtyMoudlePara.blnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ") As ����," & _
    "       Max(A.����" & IIf(mtyMoudlePara.blnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
    "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
    "       D.���� as ִ�п���,E.���� as ��������,Max(a.����Ա����) As ����Ա����, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, " & _
    "       Max(A.ժҪ) as ժҪ,max(A.���㷽ʽ) as ���㷽ʽ,max(A.���) as ���,A.��¼����,max(A.����ID) as ����ID" & _
    " From (" & strSQL & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E,ҩƷ��� X," & _
    "       �շ���Ŀ���� M1,�շ���Ŀ���� E1" & _
    " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.�շ�ϸĿID=X.ҩƷID(+)" & _
    "       And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID(+) " & _
    "       And A.�շ�ϸĿID=M1.�շ�ϸĿID(+) And M1.����(+)=1 And M1.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
    " Group by A.��¼����,A.NO,A.���,A.��������,A.�ѱ�,a.������,A.�շ�ϸĿID,C.����,C.����,B.����,Nvl(M1.����,B.����)," & _
    "       E1.����,B.���,A.���㵥λ,D.����,E.����,X.ҩƷID,X." & gstrҩ����λ & _
    " Having Sum(A.����)<>0 " & _
    " Order by �Ǽ�ʱ��,NO,���"
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objPati.����ID, mtyMoudlePara.int��������Ч���� - 1, _
        mtyMoudlePara.str�����������շѷ�ʽ)
        
    If mrsList.RecordCount = 0 Then
        MsgBox "��ǰ����δ�ҵ���Ҫ������ķ�������!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set mcllDiagnose = New Collection
    strAllNOs = ""
    With mrsList
        Do While Not .EOF
            If InStr("," & strBalance & ",", "," & Nvl(!���㷽ʽ) & ",") = 0 Then
                strBalance = strBalance & "," & Nvl(!���㷽ʽ)
            End If
            If InStr("|" & strAllNOs & "|", "|" & Nvl(!��¼����) & "," & Nvl(!NO) & "|") = 0 Then
                strAllNOs = strAllNOs & "|" & Nvl(!��¼����) & "," & Nvl(!NO)
            End If
            
            If mEditType = EM_Balance_Charge Then
                strTemp = ""
                For i = 1 To mcllDiagnose.Count
                    If mcllDiagnose(i)(0) = Nvl(!���) Then
                        strTemp = mcllDiagnose(i)(1)
                        mcllDiagnose.Remove i: Exit For
                    End If
                Next
                If InStr("," & strTemp & ",", "," & Nvl(!NO) & ",") = 0 Then strTemp = strTemp & "," & Nvl(!NO)
                If Left(strTemp, 1) = "," Then strTemp = Mid(strTemp, 2)
                mcllDiagnose.Add Array(Nvl(!���), strTemp)
            End If
            .MoveNext
        Loop
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    If strAllNOs <> "" Then strAllNOs = Mid(strAllNOs, 2)
    
    '����ԭ�տʽ
    mblnNotClick = True
    cboPayMode.Clear
    strSQL = "Select ����,���� From ���㷽ʽ Where Instr(','||[1]||',',','||����||',')>0 Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strBalance)
    With rsTemp
        Do While Not .EOF
            cboPayMode.AddItem Nvl(!����)
            rsTemp.MoveNext
        Loop
        varData = Split(strBalance, ",")
        For i = 0 To UBound(varData)
            blnFind = False
            For j = 0 To cboPayMode.ListCount - 1
                If varData(i) = cboPayMode.List(j) Then blnFind = True: Exit For
            Next
            If blnFind = False Then
                cboPayMode.AddItem varData(i)
            End If
        Next
    End With
    If cboPayMode.ListCount > 0 Then cboPayMode.ListIndex = 0
    mstrPreBalance = cboPayMode.Text
    mblnNotClick = False
    
    '�������
    With vsDiagnose
        .Clear
        .Rows = 1: .COLS = mcllDiagnose.Count
        If mEditType = EM_Balance_Charge Then
            For i = 1 To mcllDiagnose.Count
               .TextMatrix(0, i - 1) = IIf(mcllDiagnose(i)(0) = "", "������շ�", mcllDiagnose(i)(0))
               .Cell(flexcpData, 0, i - 1) = mcllDiagnose(i)(1)
               .Cell(flexcpChecked, 0, i - 1) = 2
            Next
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .COLS - 1
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
    
    Call LoadFeeData(mrsList)
    Call LoadBalanceNO(strAllNOs)
    ReadBills = True
    Exit Function
errHandle:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadBalanceNO(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ������
    '���:strNOs-��¼����,NO|....
    '����:���˺�
    '����:2014-09-26 16:42:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, strTemp As String
    
    On Error GoTo errHandle
    Set mrsBalanceNO = New ADODB.Recordset
    mrsBalanceNO.Fields.Append "��¼����", adBigInt, , adFldIsNullable
    mrsBalanceNO.Fields.Append "NO", adVarChar, 50, adFldIsNullable
    mrsBalanceNO.Fields.Append "�������", adBigInt, , adFldIsNullable
    mrsBalanceNO.Fields.Append "����ID", adBigInt, , adFldIsNullable
    mrsBalanceNO.CursorLocation = adUseClient
    mrsBalanceNO.LockType = adLockOptimistic
    mrsBalanceNO.CursorType = adOpenStatic
    mrsBalanceNO.Open
    
    If strNos = "" Then Exit Function
    If zlCommFun.ActualLen(strNos) < 4000 Then
        If ReadBalanceData(mrsBalanceNO, strNos) = False Then Exit Function
        LoadBalanceNO = True
        Exit Function
    End If
    
    varData = Split(strNos, "|")
    strTemp = ""
    For i = 0 To UBound(varData)
        If zlCommFun.ActualLen(strTemp & "|" & varData(i)) >= 4000 Then
            If ReadBalanceData(mrsBalanceNO, Mid(strTemp, 2)) = False Then Exit Function
            strTemp = ""
        End If
        strTemp = strTemp & "|" & varData(i)
    Next
    If strTemp <> "" Then
        strTemp = Mid(strTemp, 2)
        If ReadBalanceData(mrsBalanceNO, strTemp) = False Then Exit Function
    End If
    LoadBalanceNO = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadBalanceData(ByRef rsBalanceNO As ADODB.Recordset, ByVal strNos As String) As Boolean
    '��������(No,�������,����ID)��¼���м�������
    '��Σ�
    '   strNos - ��¼����,NO|....
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strFilter As String
    
    On Error GoTo errHandle
    strSQL = _
    " Select Distinct Mod(A.��¼����,10) As ��¼����, A.NO, A.����ID, Nvl(B.�������,0) As �������" & _
    " From ������ü�¼ A,����Ԥ����¼ B  " & _
    " Where a.����ID=b.����ID And ( A.��¼����,A.NO) IN (Select C1,C2 From Table(f_Str2list2([1], '|', ',')))"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    With rsTemp
        Do While Not .EOF
            strFilter = "NO='" & Nvl(!NO) & "'"
            strFilter = strFilter & " And ��¼����= " & Val(Nvl(!��¼����))
            strFilter = strFilter & " And ����ID= " & Val(Nvl(!����ID))
            strFilter = strFilter & " And �������= " & Val(Nvl(!�������))
            rsBalanceNO.Filter = strFilter
            If rsBalanceNO.EOF Then
                rsBalanceNO.Filter = 0
                rsBalanceNO.AddNew
                rsBalanceNO!��¼���� = Val(Nvl(!��¼����))
                rsBalanceNO!NO = CStr(Nvl(!NO))
                rsBalanceNO!����ID = Val(Nvl(!����ID))
                rsBalanceNO!������� = Val(Nvl(!�������))
                rsBalanceNO.Update
            End If
            .MoveNext
        Loop
    End With
    ReadBalanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadFeeData(ByVal rsList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط������ݵ������б���
    '���:rsList-�����б�
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-11 11:16:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, strDiagnose As String
    Dim strNo As String
    Dim strFilter As String
    Dim i As Long, j As Long
    Dim str��� As String, strTemp As String
    
    On Error GoTo errHandle
    If rsList Is Nothing Then Exit Function
    If rsList.State <> 1 Then Exit Function
    
    strBalance = Trim(cboPayMode.Text)
    If mEditType = EM_Balance_Charge Then '�շѽ���
        For i = 0 To vsDiagnose.Rows - 1
            For j = 0 To vsDiagnose.COLS - 1
                If Abs(Val(vsDiagnose.Cell(flexcpChecked, i, j))) = 1 Then
                    strDiagnose = "'" & vsDiagnose.TextMatrix(i, j) & "'"
                    If strDiagnose = "'������շ�'" Then strDiagnose = "null"
                    strFilter = strFilter & "or (���=" & strDiagnose & "" & " And ���㷽ʽ='" & strBalance & "') "
                End If
            Next
        Next
        
        If strBalance = "" And strFilter = "" Then
            rsList.Filter = ""
        ElseIf strFilter = "" Then
            rsList.Filter = "���㷽ʽ='" & strBalance & "'"
        Else
            rsList.Filter = Mid(strFilter, 3)
        End If
    Else '�ҺŽ���
        If strBalance = "" Then
            rsList.Filter = ""
        Else
            rsList.Filter = "���㷽ʽ='" & strBalance & "'"
        End If
    End If
    Set vsFeeList.DataSource = rsList
    Call SetFeeListHead
    
    If Not (mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge) Then
         strNo = ""
        If vsFeeList.Rows < 2 Then
            strNo = ""
        ElseIf vsFeeList.ColIndex("���") >= 0 Then
            strNo = vsFeeList.TextMatrix(1, vsFeeList.ColIndex("���"))
        End If
        vsFeeList.Editable = IIf(strNo <> "", flexEDKbdMouse, flexEDNone)
    End If
    
    If mrsList.RecordCount <> 0 Then
        mrsList.MoveFirst: strTemp = ""
        Do While Not mrsList.EOF
            str��� = IIf(Nvl(mrsList!���) = "", "������շ�", Nvl(mrsList!���))
            If InStr(strTemp & ",", "," & str��� & ",") = 0 Then
                For i = 0 To vsDiagnose.COLS - 1
                    If vsDiagnose.TextMatrix(0, i) = str��� And Abs(vsDiagnose.Cell(flexcpChecked, 0, i)) <> 1 Then
                        vsDiagnose.Cell(flexcpChecked, 0, i) = 1
                    End If
                Next
                strTemp = strTemp & "," & str���
            End If
            mrsList.MoveNext
        Loop
    End If
    
    Call CalcTotalMoney
    Call CalcRegisterYBMoney
    mblnEdit = True
    LoadFeeData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetFeeListHead(Optional blnInitHead As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���÷�����Ϣ��ͷ
    '���:blnInitHead-�Ƿ��ʼ����
    '����:���˺�
    '����:2014-09-10 17:01:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strHead As String, i As Long, varData As Variant
    
    On Error GoTo errHandle
    With vsFeeList
        .Redraw = flexRDNone
        
        If blnInitHead Then
            strHead = "����|NO|���|��������|�ѱ�|������|�շ�ϸĿID|�շ����|ѡ��|���|����|����|��Ʒ��|���|" & _
                      "��������|��λ|ҽ�����|����|����|����|Ӧ�ս��|ʵ�ս��|ִ�п���|��������|����Ա����|" & _
                      "�Ǽ�ʱ��|ժҪ|���㷽ʽ|���|��¼����|����ID"
            
            .Clear
            .Rows = 2
            varData = Split(strHead, "|")
            .COLS = UBound(varData) + 1
            For i = 0 To UBound(varData)
                .TextMatrix(0, i) = varData(i)
            Next
        ElseIf .Rows <= 1 Then
            .Clear 1
            .Rows = 2
        End If
        
        For i = 0 To .COLS - 1
            .ColKey(i) = UCase(Trim(.TextMatrix(0, i)))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            If .ColKey(i) Like "*ID" Or InStr(",���,��������,ҽ�����,�շ����,��¼����,", "," & .ColKey(i) & ",") > 0 Then
                .ColHidden(i) = True
            ElseIf .ColKey(i) Like "*��*" Or .ColKey(i) Like "*��" Or .ColKey(i) Like "*��" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf InStr(",ѡ��,�Ǽ�ʱ��,", "," & .ColKey(i) & ",") > 0 Then
                .ColAlignment(i) = flexAlignCenterCenter
            End If
        Next
        
        Select Case gTy_System_Para.bytҩƷ������ʾ
        Case 0
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("��Ʒ��")) = True
        Case 1
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("��Ʒ��")) = False
        Case 2
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("��Ʒ��")) = False
        End Select
        
        .HighLight = flexHighlightWithFocus
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .COLS - 1
        zl_vsGrid_Para_Restore mlngModule, vsFeeList, mstrTittle, "������Ϣ�б�", True, False
        
        .RowHeight(0) = 350
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        If .TextMatrix(1, .ColIndex("NO")) <> "" Then Call SplitGroupToFeeList
        
        For i = 0 To .COLS - 1
            If i >= .ColIndex("ѡ��") Then Exit For
            .ColHidden(i) = True
        Next
        .ColHidden(.ColIndex("��������")) = True
        If .ColIndex("���㷽ʽ") >= 0 Then .ColHidden(.ColIndex("���㷽ʽ")) = True
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    vsFeeList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SplitGroupToFeeList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϸ���ݷ�����ʾ
    '����:���˺�
    '����:2014-09-10 17:12:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String
    Dim bytCheck As Byte
    
    On Error GoTo errHandle
    bytCheck = 1
    
    With vsFeeList
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("NO"), .ColIndex("ʵ�ս��"), gstrDec, &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("NO"), .ColIndex("Ӧ�ս��"), gstrDec, &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("���")
        .OutlineCol = .ColIndex("���")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                .RowHeight(i) = 350
                .TextMatrix(i, .ColIndex("NO")) = Trim(.Cell(flexcpTextDisplay, i + 1, .ColIndex("NO")))
                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("NO")) & "(" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("����")) & ")"
                 strTemp = strTemp & Space(2) & "�ѱ�:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("�ѱ�"))
                 strTemp = strTemp & Space(2) & "��������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("��������"))
                 strTemp = strTemp & Space(2) & "������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("������"))
                 
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("���"), i, .ColIndex("���")) = 1
                 For j = 0 To .COLS - 1
                    If j < .ColIndex("Ӧ�ս��") Then
                        If j >= .ColIndex("���") Then
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = True
                        ElseIf j = .ColIndex("ѡ��") Then
                            .Cell(flexcpChecked, i, j) = bytCheck
                            .Cell(flexcpAlignment, i, j, i, j) = 4
                            If mEditType = EM_Balance_Register Then bytCheck = 2
                        End If
                    ElseIf .ColIndex("ʵ�ս��") = j Then
                        .Cell(flexcpData, i, j) = Val(.TextMatrix(i, j))
                        .TextMatrix(i, j) = Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    ElseIf .ColIndex("Ӧ�ս��") = j Then
                        .Cell(flexcpData, i, j) = Val(.TextMatrix(i, j))
                        .TextMatrix(i, j) = " " & Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    End If
                 Next
            Else
                .TextMatrix(i, .ColIndex("ѡ��")) = ""
                .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))), gstrFeePrecisionFmt)
                .TextMatrix(i, .ColIndex("����")) = FormatEx(Val(.TextMatrix(i, .ColIndex("����"))), 5)
                .Cell(flexcpData, i, .ColIndex("Ӧ�ս��")) = Val(.TextMatrix(i, .ColIndex("Ӧ�ս��")))
                .Cell(flexcpData, i, .ColIndex("ʵ�ս��")) = Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                
                .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("Ӧ�ս��"))), gstrDec)
                .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ�ս��"))), gstrDec)
                
            End If
        Next
        
        Call .AutoSize(.ColIndex("���"))
        Call .AutoSize(.ColIndex("����"))
        
        For j = 0 To .COLS - 1
            If j < .ColIndex("Ӧ�ս��") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetBillNoFromPati(ByVal strNo As String, ByRef lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݺţ���ȡ��Ӧ�Ĳ���ID
    '���:strNo-���ݺ�
    '����:lng����ID-���ز���ID
    '����:�ҵ���Ӧ�ĵ��ݣ�����true,���򷵻�False
    '����:���˺�
    '����:2014-09-10 12:38:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    lng����ID = 0
    strSQL = "Select ����ID From ������ü�¼ Where ��¼����=1 and NO=[1] and ��¼״̬ in (1,3) and rownum< 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    If rsTemp.EOF Then Exit Function
    lng����ID = Val(Nvl(rsTemp!����ID))
    GetBillNoFromPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPati(ByVal lng����ID As String, ByRef objPati As PatiInfor, ByRef str�������� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID,���»�ȡ����
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-04-06 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere As String
    
    Set objPati = New PatiInfor
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select a.����id, a. �����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ,p.���� as ҽ�Ƹ��ʽ����, a. ����, a.�Ա�, a. ����, a.��������, a.�����ص�, a.���֤��, a.����֤��, a.���, " & _
    "        a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.�໤��, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, " & _
    "        a.��ͬ��λid, a.������λ, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.������, a.������, a.��������, a.����ʱ��, a.����״̬, a.��������, a.��Ժ, a.Ic����, " & _
    "        a.������, a.ҽ����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, '' as ����, 0As ��״̬,'' as ����, '' as ��ʧ��ʽ, " & _
    "       sysdate as ��ʧʱ��, 0  as ��ʧ��Ч����,sysdate as ��ǰʱ��,C.���� as ��������" & _
    "   From ������Ϣ A,ҽ�Ƹ��ʽ P,������� C " & _
    "   Where A.���� = C.���(+) And a.ҽ�Ƹ��ʽ=P.����(+) And ����ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID)
    If rsTemp.EOF Then Exit Function
    objPati.����ID = rsTemp!����ID
    objPati.����� = IIf(Val(Nvl(rsTemp!�����)) = 0, "", Nvl(rsTemp!�����))
    objPati.���� = Nvl(rsTemp!����)
    objPati.�Ա� = Nvl(rsTemp!�Ա�)
    objPati.���� = Nvl(rsTemp!����)
    objPati.�������� = Format(rsTemp!��������, "yyyy-mm-dd")
    objPati.������ַ = Nvl(rsTemp!�����ص�)
    objPati.���֤�� = Nvl(rsTemp!���֤��)
    objPati.����֤�� = Nvl(rsTemp!����֤��)
    objPati.ְҵ = Nvl(rsTemp!ְҵ)
    objPati.���� = Nvl(rsTemp!����)
    objPati.���� = Nvl(rsTemp!����)
    objPati.ѧ�� = Nvl(rsTemp!ѧ��)
    objPati.����״�� = Nvl(rsTemp!����״��)
    objPati.���� = Nvl(rsTemp!����״��)
    objPati.��ͥ��ַ = Nvl(rsTemp!��ͥ��ַ)
    objPati.��ͥ�绰 = Nvl(rsTemp!��ͥ�绰)
    objPati.��ͥ�ʱ� = Nvl(rsTemp!��ͥ��ַ�ʱ�)
    objPati.�໤�� = Nvl(rsTemp!�໤��)
    objPati.��ϵ�� = Nvl(rsTemp!��ϵ������)
    objPati.��ϵ�˹�ϵ = Nvl(rsTemp!��ϵ�˹�ϵ)
    objPati.��ϵ�˵�ַ = Nvl(rsTemp!��ϵ�˵�ַ)
    objPati.��ϵ�˵绰 = Nvl(rsTemp!��ϵ�˵绰)
    objPati.������λ = Nvl(rsTemp!������λ)
    objPati.������λ�绰 = Nvl(rsTemp!��λ�绰)
    objPati.������λ�ʱ� = Nvl(rsTemp!��λ�ʱ�)
    objPati.������λ������ = Nvl(rsTemp!��λ������)
    objPati.������λ�������ʻ� = Nvl(rsTemp!��λ�ʺ�)
    objPati.���ڵ�ַ = Nvl(rsTemp!���ڵ�ַ)
    objPati.���ڵ�ַ�ʱ� = Nvl(rsTemp!���ڵ�ַ�ʱ�)
    objPati.���� = Nvl(rsTemp!����)
    objPati.���� = Nvl(rsTemp!����)
    objPati.ҽ�Ƹ��ʽ���� = Nvl(rsTemp!ҽ�Ƹ��ʽ����)
    objPati.ҽ�Ƹ��ʽ = Nvl(rsTemp!ҽ�Ƹ��ʽ)
    str�������� = Nvl(rsTemp!��������)
    zlGetPati = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Function PatiErrBillPay(ByVal lng����ID As Long) As Boolean
    '����:���ݲ���,�Խ����쳣���ݽ����ؽ�
    '���:lng����ID-ָ���Ĳ���ID
    '����:�����쳣����,���������½���,����true,���򷵻�False
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim str����Ա���� As String, blnDoElsePersonErr As Boolean
    Dim lng����ID As Long
    Dim blnRegister As Boolean
    Dim editTypeTemp As EM_Balance_Type
    
    mblnElsePersonErrBill = False
    blnRegister = mEditType = EM_Balance_Register
   
    On Error GoTo errHandle
    strSQL = "Select ����ID, ����Ա����" & vbNewLine & _
            " From ���ò����¼" & vbNewLine & _
            " Where Nvl(����״̬,0) = 1 And ��¼���� = 1 And ��¼״̬ = 1" & vbNewLine & _
            "       And Nvl(���ӱ�־,0) = [2] And ����id =[1] And Rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, IIf(blnRegister, 1, 0))
    If rsTemp.EOF Then Exit Function
    
    lng����ID = Val(Nvl(rsTemp!����ID))
    str����Ա���� = Nvl(rsTemp!����Ա����)
    
    If str����Ա���� <> UserInfo.���� Then
        '�ж��Ƿ��ܹ������˵��շ��쳣���ݽ�������
        strSQL = "Select �������" & vbNewLine & _
                " From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
                " Where Nvl(a.���㷽ʽ, '-') = b.���� And b.���� Not In ('3', '4') And a.����id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If rsTemp.EOF Then
            '107905�����С��ؽ������쳣���ݡ�Ȩ��ʱ�����Զ�ֻ������ҽ����������˵��쳣���㵥�ݽ����ؽ�
            blnDoElsePersonErr = zlStr.IsHavePrivs(mstrPrivs, "�ؽ������쳣����")
        Else
            '����������ҽ�����㷽ʽ����������Ա�Ͳ��ܴ�����
            blnDoElsePersonErr = False
        End If
        
        If blnDoElsePersonErr = False Then
            If MsgBox("ע��:" & vbCrLf & _
                "       �ò��˴����쳣��" & IIf(blnRegister, "�Һ�", "����") & _
                "������㵥�ݣ�����Ա[" & str����Ա���� & "]��ȡ��һ���֣�" & _
                "ע�⵽����Ա[" & str����Ա���� & "]�����쳣���ݽ����ؽᣡ" & vbCrLf & vbCrLf & _
                "       �Ƿ�����Ըò��˽���" & IIf(blnRegister, "�Һ�", "����") & "������㣿", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                PatiErrBillPay = True
            End If
            Exit Function
        End If
    End If
    
    If MsgBox("ע��:" & vbCrLf & _
            "       �ò��˴����쳣��" & IIf(blnRegister, "�Һ�", "����") & "������㵥��" & _
            IIf(str����Ա���� <> UserInfo.����, "���õ����ǲ���Ա[" & str����Ա���� & "]��ȡ��", "") & _
            "���Ƿ����¶Ըõ��ݽ������½��㣿" & vbCrLf & vbCrLf & _
            "���ǡ��������¶��쳣���ݽ������½���" & vbCrLf & _
            "���񡻴������쳣���ݽ��д�����������" & IIf(blnRegister, "�Һ�", "����") & "����������", _
            vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Function
    End If
    
    '���¶��쳣���ݽ����ؽ�
    mblnElsePersonErrBill = blnDoElsePersonErr
    mEditType = EM_Balance_Err_ReCharge
    mstr����ID = lng����ID
    If LoadErrBillData(mobjPatiInfor) = False Then
        PatiErrBillPay = True
        Exit Function
    End If
    Call cmdOK_Click
    
    PatiErrBillPay = True
    mstr����ID = ""
    mEditType = IIf(blnRegister, EM_Balance_Register, EM_Balance_Charge)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub PatiIdentify_GotFocus()
    zlControl.TxtSelAll PatiIdentify.objTxtInput
    If gblnLED Then zl9LedVoice.Speak "#51" '�����������
End Sub

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(PatiIdentify.Text) = "" Then
        KeyAscii = 0
        Call CheckPatiInfor(mobjPatiInfor)
    End If
End Sub

Private Sub picDiagnose_Resize()
    Err = 0: On Error Resume Next
    With picDiagnose
        vsDiagnose.Left = .ScaleLeft
        vsDiagnose.Top = .ScaleTop
        vsDiagnose.Height = .ScaleHeight
        vsDiagnose.Width = .ScaleWidth
    End With
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        lblʵ��.Left = .ScaleWidth - lblʵ��.Width - 300
        lblӦ��.Left = lblʵ��.Left - lblӦ��.Width - 400
        txtժҪ.Width = lblӦ��.Left - txtժҪ.Left - 400
        vsBalance.Width = .ScaleWidth - vsBalance.Left - 50
        fraDownSplit.Width = .ScaleWidth + 100
        fraDownSplit.Left = .ScaleLeft
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
        cmdԤ����.Left = cmdOK.Left - cmdԤ����.Width - 50
        
        txt�˿�ϼ�.Left = IIf(cmdԤ����.Visible = False, cmdOK.Left, cmdԤ����.Left) - txt�˿�ϼ�.Width - 200
        lbl�˿�ϼ�.Left = txt�˿�ϼ�.Left - lbl�˿�ϼ�.Width - 20
    End With
End Sub

Private Sub picFeeList_Resize()
    Err = 0: On Error Resume Next
    With picFeeList
        vsFeeList.Left = .ScaleLeft
        vsFeeList.Top = .ScaleTop
        vsFeeList.Height = .ScaleHeight
        vsFeeList.Width = .ScaleWidth
    End With
End Sub
 
Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    With picTop
        fraInfo.Left = .ScaleLeft
        fraInfo.Width = .ScaleWidth
        If cmdDelete.Visible Then
            cmdDelete.Left = .ScaleWidth - .ScaleLeft - cmdDelete.Width - 100
            cboNO.Left = .ScaleWidth - .ScaleLeft - cboNO.Width - 550
        Else
            cboNO.Left = .ScaleWidth - .ScaleLeft - cboNO.Width - 50
        End If
        lblNO.Left = cboNO.Left - lblNO.Width - 20
        txtInvoice.Left = lblNO.Left - txtInvoice.Width * 1.3
        txtMCInvoice.Left = txtInvoice.Left
        lblFact.Left = txtInvoice.Left - lblFact.Width - 20
        If txtInvoice.Visible Then
            lblFormat.Left = lblFact.Left - lblFormat.Width - 50
            lblFormat.Top = lblFact.Top
        Else
            lblFormat.Left = lblPayMode.Left - lblPayMode.Width - 50
            lblFormat.Top = lblPayMode.Top
        End If
        cboPayMode.Left = .ScaleWidth - cboPayMode.Width - 50
        lblPayMode.Left = cboPayMode.Left - lblPayMode.Width - 20
        lbl����.Left = .ScaleLeft + 100
    End With
End Sub

Private Function MCPatientProcess(ByRef objPatiInfor As zlIDKind.PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ�������֤
    '����:���˺�
    '����:2014-09-16 09:59:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnTran As Boolean
    Dim lng����ID As Long, lng����IDOut As Long
    Dim lngTemp As Long, str�������� As String
    Dim strAdvance As String
    
    On Error GoTo errH
'    PatiIdentify.AllowAutoCommCard = False
'    PatiIdentify.AllowAutoICCard = False
'    PatiIdentify.AllowAutoIDCard = False
    If Not objPatiInfor Is Nothing Then
        lng����ID = objPatiInfor.����ID
    Else
        lng����ID = 0
    End If
    
    If gblnLED Then zl9LedVoice.Speak "#50"
    lng����IDOut = lng����ID '����Identify�ӿ����޸ĸñ����󷵻���ֵ
    
    '���أ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID,24��������(1=��������),25������������
    '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    strAdvance = "2"
    mstrYBPati = gclsInsure.Identify(IIf(mEditType = EM_Balance_Register, 3, 0), lng����IDOut, mintInsure, strAdvance)
    If mstrYBPati = "" Then GoTo CheckValied:
    
    '��ȡ������Ϣ
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
            lngTemp = Val(CLng(Split(mstrYBPati, ";")(8)))
            If lng����ID <> lngTemp And lng����ID <> 0 Then
                MsgBox "ҽ����֤�Ĳ�����֮ǰ��ȡ�Ĳ��˲���ͬһ������!", vbInformation, gstrSysName
                staThis.Panels(Pan.C2��ʾ��Ϣ) = "ҽ����֤�Ĳ�����֮ǰ��ȡ�Ĳ��˲���ͬһ������!��"
                Call YBIdentifyCancel
                GoTo CheckValied:
                Exit Function
            End If
        End If
        lng����ID = lng����IDOut
    End If
            
    '����:29283
    '  -- ����:���ó���-1-�Һ�;2-�շ�
    '  --        ����id_In-����ID(δ������,������)
    '  --        ����_In: ˢ������;δˢ��ʱ,Ϊ��
    '  --         ˢ����ʽ_In:  1-����ˢ��;2-ҽ��ˢ��
    If zlPatiCardCheck(IIf(mEditType = EM_Balance_Register, 1, 2), lng����ID, CStr(Split(mstrYBPati, ";")(0)), 2) = False Then
        Call YBIdentifyCancel
        GoTo CheckValied: Exit Function
    End If
    
    Call initInsurePara(lng����ID)    '��ʼ��ҽ������
    
    If zlGetPati(lng����ID, objPatiInfor, str��������) = False Then
        Call YBIdentifyCancel
        GoTo CheckValied: Exit Function
    End If
    objPatiInfor.���� = mintInsure
    txtYB.Text = mintInsure
    PatiIdentify.ForeColor = vbRed
    If objPatiInfor.�������� <> "" Then
        Call SetPatiColor(PatiIdentify.objTxtInput, objPatiInfor.��������, vbRed)
    End If
    
    PatiIdentify.Text = Split(mstrYBPati, ";")(3)
    PatiIdentify.PasswordChar = ""
    PatiIdentify.IMEMode = 0
    lblPatiInfor.Caption = "�Ա�:" & objPatiInfor.�Ա�
    lblPatiInfor.Caption = lblPatiInfor.Caption & Space(4) & "����:" & objPatiInfor.����
    lblPatiInfor.Caption = lblPatiInfor.Caption & Space(4) & "���ʽ:" & objPatiInfor.ҽ�Ƹ��ʽ
    lbl����.Caption = str��������
    
    '�����ʻ�
    Dim cur͸֧�� As Currency
    cur͸֧�� = RoundEx(mTy_Insure.dbl����͸֧, 2)
    
    mTy_Insure.dbl�ʻ���� = gclsInsure.SelfBalance(lng����ID, CStr(Split(mstrYBPati, ";")(1)), 10, cur͸֧��, mintInsure)
    staThis.Panels(Pan.C3�����ʻ�).Text = "�����ʻ����:" & Format(mTy_Insure.dbl�ʻ����, "0.00")
    staThis.Panels(Pan.C3�����ʻ�).Visible = True
    mTy_Insure.dbl����͸֧ = cur͸֧��
    
    Call SetButtons '���ð�ť
    
   If MCPAR.����Ԥ���� And mstr�����ʻ� <> "" Then  'ֻ��ʹ�ø����ʻ�����
        vsBalance.COLS = 3
        vsBalance.TextMatrix(0, 0) = "ҽ������"
        vsBalance.TextMatrix(0, 1) = mstr�����ʻ�
        vsBalance.TextMatrix(0, 2) = "0.00"
        vsBalance.ColData(1) = 0
        vsBalance.ColData(2) = 0
    End If
    
    staThis.Panels(Pan.C2��ʾ��Ϣ) = ""
    SetPatientEnableModi (False)
    Call ShowWelcomeByLed
    Call ReInitPatiInvoice(True, mintInsure, lng����ID)
    
    '���ݲ���,���쳣���ݽ����ؽ�
    If PatiErrBillPay(lng����ID) Then
        Call YBIdentifyCancel
        GoTo CheckValied: Exit Function
    End If
    
    MCPatientProcess = True
    
    Exit Function
CheckValied:    '���ʧ��
    mintInsure = 0: mTy_Insure.dbl�ʻ���� = 0: mTy_Insure.dbl����͸֧ = 0
    Set objPatiInfor = Nothing
    staThis.Panels(Pan.C3�����ʻ�).Text = ""
    staThis.Panels(Pan.C3�����ʻ�).Visible = False
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
    Call PatiIdentify_GotFocus
'    PatiIdentify.AllowAutoCommCard = True
'    PatiIdentify.AllowAutoICCard = True
'    PatiIdentify.AllowAutoIDCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function YBIdentifyCancel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ��ҽ�����������֤
    '����:���ؼ�ʱ���˳�������������
    '����:���˺�
    '����:2014-09-16 16:07:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    YBIdentifyCancel = True
    If mstrYBPati = "" Or PatiIdentify.Text = "" Then Exit Function
    If UBound(Split(mstrYBPati, ";")) < 8 Then Exit Function
    If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
        lng����ID = Val(CLng(Split(mstrYBPati, ";")(8)))
    End If
    If lng����ID = 0 Then Exit Function
    YBIdentifyCancel = gclsInsure.IdentifyCancel(IIf(mEditType = EM_Balance_Register, 3, 0), lng����ID, mintInsure)
End Function

Public Function zlPatiCardCheck(ByVal byt���ó��� As Byte, lng����ID As Long, str���� As String, bytˢ����ʽ As Byte) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���鲡��ˢ����ʽ
    '��Σ�byt���ó���: 1-�Һ�;2-�շ�
    '         lng����ID:����ID(δ������,������)
    '         str����;δˢ��ʱ,Ϊ��
    '         bytˢ����ʽ: 1-����ˢ��;2-ҽ��ˢ��
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-04-27 16:09:08
    '˵����һ�����ŵ����ݲ��ˣ�ʹ�õ�ҽ����ͬʱҲ�Ǿ��￨��ҽԺҪ�������ҽ����ʽ����
    '          �����֤�Һš��շѣ����������Էѷ�ʽֱ��ˢ�����У����Ҫ���ڹҺš��շ�ʱ�����ݲ���ˢ�������������ҽ�������֤��ʽˢ�Ŀ���
    '          ����ֱ��ˢ�Ŀ�������ʾ�������������
    '����:29283
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = " Select Zl_Paticardcheck([1],[2],[3],[4]) as ��ʾ��Ϣ From Dual "
    ' Zl_Paticardcheck
    '  ���ó���_IN NUMBER ,
    '  ����id_In Number,
    '  ����_In   Varchar2,
    '  ˢ����ʽ_In Number:=1
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鲡��ˢ����ʽ�Ƿ�Ϸ�", byt���ó���, lng����ID, str����, bytˢ����ʽ)
    strSQL = Nvl(rsTemp!��ʾ��Ϣ)
    If strSQL <> "" Then
        MsgBox strSQL, vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    zlPatiCardCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initInsurePara(ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '����:���˺�
    '����:2011-08-27 12:25:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, mintInsure)
    MCPAR.����Ԥ���� = gclsInsure.GetCapability(support����Ԥ��, lng����ID, mintInsure)
    MCPAR.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, lng����ID, mintInsure)
    MCPAR.���Ը� = gclsInsure.GetCapability(support�շ��ʻ������Ը�, lng����ID, mintInsure)
    MCPAR.ȫ�Ը� = gclsInsure.GetCapability(support�շ��ʻ�ȫ�Է�, lng����ID, mintInsure)
    MCPAR.ʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, lng����ID, mintInsure)
    MCPAR.ҽ������Ʊ�� = False
    MCPAR.�Һ�ʹ�ø����ʻ� = gclsInsure.GetCapability(support�Һ�ʹ�ø����ʻ�, lng����ID, mintInsure)
    MCPAR.���ղ����� = gclsInsure.GetCapability(support�ҺŲ���ȡ������, lng����ID, mintInsure)
End Sub

Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ���ڹ�������
    '����:���������,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 16:49:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '�Ƿ�������ѵĴ���
    If IsCheck���� = False Then Exit Function
    
    '���㷽ʽ���
    Set mrs���㷽ʽ = Get���㷽ʽ("�շ�")
    If mrs���㷽ʽ.RecordCount = 0 Then
        MsgBox "�շѳ���û�п��õĽ��㷽ʽ�����ȵ����㷽ʽ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    If mstr�����ʻ� = "" Then
        mrs���㷽ʽ.Filter = "����=3"
        If Not mrs���㷽ʽ.EOF Then mstr�����ʻ� = mrs���㷽ʽ!����
    End If
    If mstrӦ������㷽ʽ = "" Then
        mrs���㷽ʽ.Filter = "Ӧ����=1"
        If Not mrs���㷽ʽ.EOF Then mstrӦ������㷽ʽ = Nvl(mrs���㷽ʽ!����)
    End If
    mrs���㷽ʽ.Filter = 0
    
    Set mobjPayCards = GetPayCardsObject
    If mobjPayCards Is Nothing Then Exit Function
    If mobjPayCards.Count = 0 Then Exit Function
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPayCardsObject() As Cards
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������֧�ֵĽ���������
    '����:����Cards����
    '����:���˺�
    '����:2015-03-18 09:56:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objCards As Cards, objPayCards As Cards
    Dim rsTemp As ADODB.Recordset
    Dim lngKey As Long, i As Long, blnFind As Boolean
    
    On Error GoTo errHandle
    
    Set objCards = New Cards: Set objPayCards = New Cards
    Set rsTemp = Get���㷽ʽ("������")
    '83533:���ϴ�,2015/3/25,û����Ч�Ĳ�����
    If rsTemp.RecordCount = 0 Then
        MsgBox "������û�п��õĽ��㷽ʽ�����ȵ������㷽ʽ���������ò������Ӧ�ó��ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
        '���:bytType-0-����ҽ�ƿ�;
        '             1-���õ�ҽ�ƿ�,
        '             2-���д��������˻���������
        '             3-���õ������˻���ҽ�ƿ�
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.Count
                If objCards(i).���㷽ʽ = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                '83266:���ϴ�,2015/3/18,ҽ�ƿ������ж��Ƿ�����
                If InStr(",1,2,", "," & Val(Nvl(rsTemp!����)) & ",") > 0 _
                    And Val(Nvl(rsTemp!Ӧ����)) <> 1 Then
                    '������ҽ���Ľ��㷽ʽ����֧Ʊ��
                     Set objCard = New Card
                     objCard.���� = Mid(Nvl(!����), 1, 1)
                     objCard.�ӿڱ��� = Nvl(!����)
                     objCard.�ӿڳ����� = ""
                     objCard.�ӿ���� = -1 * lngKey
                     objCard.���㷽ʽ = Nvl(!����)
                     objCard.���� = Nvl(!����)
                     objCard.���� = True
                     objCard.ȱʡ��־ = Val(Nvl(rsTemp!ȱʡ)) = 1
                     objCard.֧������ = True
                     objCard.�������� = Val(!����)
                    objPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
                End If
            End If
            .MoveNext
        Loop
    End With
    '��������
    For Each objCard In objCards
        If objCard.���ѿ� = False Then
            rsTemp.Filter = "����='" & objCard.���㷽ʽ & "'"
            If Not rsTemp.EOF Then
                objPayCards.Add objCard, "K" & lngKey
                lngKey = lngKey + 1
            End If
        End If
    Next
    If objPayCards.Count = 0 Then
        MsgBox "���㿨��������,ԭ���������:" & vbCrLf & _
        "δ�������ý��㿨,�뵽��ҽ�ƿ���𡻺͡��豸���á�������", vbInformation, gstrSysName
    End If
    Set GetPayCardsObject = objPayCards
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsCheck����() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ���������
    '����:��������,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-05 15:17:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gstr�������� <> "" Then IsCheck���� = True: Exit Function
    If Not (mEditType = EM_Balance_Register Or mEditType = EM_Balance_Charge) Then IsCheck���� = True: Exit Function
    MsgBox "ϵͳ����δ������Ч������,����[���㷽ʽ����]�����á�", vbInformation, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    'ѡ���˷�ԭ��
    If KeyCode <> vbKeyReturn Then Exit Sub

    If Trim(txtժҪ.Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txtժҪ.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txtժҪ, Trim(txtժҪ.Text), "�����˷�ԭ��", "�����˷�ԭ��ѡ��", True, True) = False Then
        If zlCommFun.IsCharChinese(Trim(txtժҪ.Text)) = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtժҪ_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtժҪ
End Sub

Private Sub txtժҪ_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtժҪ, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtժҪ_LostFocus()
    zlCommFun.OpenIme False
    If zlCommFun.ActualLen(txtժҪ.Text) > 100 Then
        MsgBox "ժҪ�����������100���ַ���50�����֣�", vbInformation, gstrSysName
        If txtժҪ.Visible And txtժҪ.Enabled Then txtժҪ.SetFocus
    End If
End Sub

Private Sub txtժҪ_Change()
    txtժҪ.Tag = ""
End Sub

Private Sub vsBalance_DblClick()
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    With vsBalance
      '�������޸ĵ�ҽ����Ŀ
      If Val(.ColData(.Col)) = 0 Then Exit Sub
      .EditCell
      .EditSelStart = 0
      .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsBalance_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Cancel = True: Exit Sub
    With vsBalance
        If Val(.ColData(Col)) = 0 Then Cancel = True: Exit Sub
        '���õ�Ԫ��ı༭����
        .EditMaxLength = 16
    End With
End Sub

Private Sub vsBalance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'ҽ���ӿڼ��
     Dim curOrig As Currency, curTotal As Currency
     Dim i As Integer, strKey As String, str���㷽ʽ As String, varData As Variant
    '������֤
    With vsBalance
        If Val(.ColData(Col)) = 0 Then Exit Sub
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        .EditText = Format(Val(strKey), "0.00")
        If strKey = "" Then Exit Sub
        
        If Not IsNumeric(strKey) Then
            MsgBox .TextMatrix(.Row, 0) & "�����˷Ƿ��ַ�,ֻ�����������ͣ�", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        If Val(strKey) = 0 Then Exit Sub
        
        str���㷽ʽ = Trim(.TextMatrix(0, .Col - 1))
        If str���㷽ʽ = "" Then Exit Sub
        '��������������ص�ԭʼ���(�����ʻ�����͸֧ʱ���ж�)
        curOrig = GetMedicareBalanceSum(str���㷽ʽ, True)      '�ý��㷽ʽ����ԭʼ���ؽ��
        If (str���㷽ʽ <> mstr�����ʻ� Or mTy_Insure.dbl����͸֧ = 0) _
            And Val(strKey) > curOrig And Val(strKey) <> 0 And curOrig <> 0 Then
            Cancel = True
            MsgBox "�����""" & str���㷽ʽ & """������ܳ��� " & Format(curOrig, "0.00") & " ��", vbInformation, gstrSysName
            Exit Sub
        End If
            
        '�����ʻ����
        If str���㷽ʽ = mstr�����ʻ� Then
            '������������͸֧���
            If mTy_Insure.dbl�ʻ���� - Val(strKey) < -1 * mTy_Insure.dbl����͸֧ Then
                Cancel = True
                MsgBox "�ʻ����:" & Format(mTy_Insure.dbl�ʻ����, "0.00") & _
                    IIf(mTy_Insure.dbl����͸֧ = 0, "", "(" & "����͸֧:" & Format(mTy_Insure.dbl����͸֧, "0.00") & ")") & _
                    "����Ҫ����Ľ�", vbInformation, gstrSysName
                Exit Sub
             End If
        End If
            
        '������������ʣ��ɽ�����
        curTotal = RoundEx(Val(lblʵ��.Tag), "0.00")
        For i = 1 To mcolBalance.Count
           '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;�ĺ���
            varData = Split(mcolBalance(i) & ";;;;", ";")
            If varData(0) <> str���㷽ʽ Then
                curTotal = curTotal - CCur(varData(3))
            End If
        Next
        If Val(strKey) > curTotal Then
            Cancel = True
            MsgBox "��������󣬳����������������:" & Format(curTotal, "0.00") & "��", vbInformation, gstrSysName
            Exit Sub
        End If
        .EditText = FormatEx(Val(strKey), 6)
        Call SetBalanceVal(str���㷽ʽ, CCur(Val(strKey)))
    End With
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    With vsBalance
        '�������޸ĵ�ҽ����Ŀ
        If Val(.ColData(Col)) = 0 Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsBalance_GotFocus()
    vsBalance_EnterCell
End Sub

Private Sub vsBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    Call VsFlxGridCheckKeyPress(vsBalance, vsBalance.Row, vsBalance.Col, KeyAscii, m���ʽ)
End Sub

Private Sub vsBalance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Val(vsBalance.ColData(Col)) = 0 Then Exit Sub
    Call VsFlxGridCheckKeyPress(vsBalance, Row, Col, KeyAscii, m���ʽ)
End Sub

Private Sub vsBalance_EnterCell()
    With vsBalance
        If .Col <= 0 Then Exit Sub
    End With
    
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    With vsBalance
        If .ColData(.Col) = 0 Then
             .FocusRect = flexFocusLight
        Else
             .FocusRect = flexFocusHeavy
        End If
    End With
End Sub

Private Sub SetPatientEnableModi(blnModi As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò��˱༭��Ϣ
    '����:���˺�
    '����:2014-09-16 11:42:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PatiIdentify.Locked = Not blnModi
    If blnModi Then
        PatiIdentify.BackColor = &HFFFFFF
    Else
        PatiIdentify.BackColor = &HE0E0E0
    End If
End Sub

Private Sub ShowWelcomeByLed()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ӭ��Ϣ�Ͳ�����Ϣ
    '����:���˺�
    '����:2014-06-06 17:56:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String, lngPatient As Long
    If gblnLED = False Then Exit Sub
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    If gblnLedWelcome Then
        zl9LedVoice.Reset msCommSpeak
        zl9LedVoice.Speak "#1"
        zl9LedVoice.Init UserInfo.��� & " �շ�ԱΪ������", mlngModule, gcnOracle
    End If
    strInfo = Trim(PatiIdentify.Text)
    If Not mobjPatiInfor Is Nothing Then strInfo = strInfo & " " & mobjPatiInfor.�Ա� & " " & mobjPatiInfor.����: lngPatient = mobjPatiInfor.����ID
    zl9LedVoice.DisplayPatient strInfo, lngPatient
End Sub

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True, _
    Optional ByVal intInsure_IN As Integer = 0, Optional ByVal lng����ID_In As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '���:blnFact-�Ƿ�����ȡ��Ʊ��
    '����:���˺�
    '����:2011-04-29 14:17:33
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng����ID As Long
    Dim intInsure As Integer, lngCur����ID As Long, lng��ҳID As Long
    
    If Not mobjPatiInfor Is Nothing Then lngCur����ID = mobjPatiInfor.����ID
    
    lng����ID = IIf(lng����ID_In <> 0, lng����ID_In, lngCur����ID)
    intInsure = IIf(intInsure_IN <> 0, intInsure_IN, mintInsure)
    
    If lng����ID = 0 Then
        '�ϴβ���ID
        If PatiIdentify.Text = mstrPrePati And mlngPrePati <> 0 Then
            lng����ID = mlngPrePati
        End If
    End If
    If lng����ID = 0 Then lng����ID = lngCur����ID
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_�շ��վ�, lng����ID, lng��ҳID, mintInsure, mobjFactProperty)
    Call ZlShowBillFormat(mlngModule, lblFormat, mobjFactProperty.��ӡ��ʽ)
    If blnFact Then Call RefreshFact
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mobjInvoice.zlGetInvoiceGroupID(mlngModule, UserInfo.����, EM_�շ��վ�, mobjFactProperty.ʹ�����, lng����ID, mobjFactProperty.��������ID, lng����ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng����ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng����ID
        Case 0 '����ʧ��
        Case -1
            If Trim(mobjFactProperty.ʹ�����) = "" Then
                MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "��û�����ú͹��õġ�" & mobjFactProperty.ʹ����� & "���շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFactProperty.ʹ�����) = "" Then
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "���صĹ���Ʊ�ݵġ�" & mobjFactProperty.ʹ����� & "���շ�Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub RefreshFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ���շ�Ʊ�ݺ�
    '����:���˺�
    '����:2014-06-06 14:21:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
    If mobjFactProperty Is Nothing Then Exit Sub
    If mobjFactProperty.��ӡ��ʽ = 0 And Not MCPAR.ҽ���ӿڴ�ӡƱ�� Then Exit Sub
    
    If mobjFactProperty.�ϸ���� Then
        'lblFact.tag��Ҫ�Ǽ�鷢Ʊ���Ƿ��ֹ������.�ֹ������,��Ʊ��Ϊ��,�������Զ������ķ�Ʊ��
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            
            If zlGetInvoiceGroupUseID(mlng����ID) = False Then
                txtInvoice.Text = "": txtInvoice.Tag = "": Exit Sub
            End If
            
            '�ϸ�ȡ��һ������
            If mobjInvoice.zlGetNextBill(mlngModule, mlng����ID, strFactNO) = False Then strFactNO = ""
            txtInvoice.Text = strFactNO
            
            'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
            '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
            '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
            txtInvoice.Tag = txtInvoice.Text
            lblFact.Tag = txtInvoice.Tag
            If mobjFactProperty.����ʹ����� Then Call zlCheckFactIsEnough
        End If
    Else
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            '��ɢ��ȡ��һ������
            txtInvoice.Text = zlStr.Increase(UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, mlngModule)))
        End If
        txtInvoice.Tag = txtInvoice.Text
        lblFact.Tag = txtInvoice.Tag
    End If
    txtInvoice.SelStart = Len(txtInvoice.Text)
End Sub

Private Sub zlCheckFactIsEnough(Optional ByVal intInvoicePages As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰƱ���Ƿ�����
    ' ���:intInvoicePages-��Ҫ�ķ�Ʊ����,���Ϊ0,��ϵͳ��������
    '����:���˺�
    '����:2011-05-10 17:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngʣ������ As Long, lngNums As Long
    
    If Not (mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register) Then Exit Sub
    
    '���˺� ����:26948 ����:2009-12-28 17:43:00
    '��Ҫ���ʣ�������Ƿ����:
 
    If intInvoicePages <> 0 Then
        If mobjInvoice.zlCheckInvoiceOverplusEnough(1, intInvoicePages, lngʣ������, mlng����ID, mobjFactProperty.ʹ�����) = False Then
            MsgBox "ע��:" & vbCrLf & _
                   "    ��ǰʣ��Ʊ�ݲ���(" & lngʣ������ & ") ,��ǰ��Ҫ" & intInvoicePages & "��Ʊ��,��ע�������Ʊ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    Else
        If mobjInvoice.zlCheckInvoiceOverplusEnough(1, mtyMoudlePara.int����ʣ��Ʊ������, lngʣ������, mlng����ID, mobjFactProperty.ʹ�����) = False Then
            MsgBox "ע��:" & vbCrLf & _
                   "    ��ǰʣ��Ʊ��(" & lngʣ������ & ") С���˱���������(" & mtyMoudlePara.int����ʣ��Ʊ������ & "),��ע�������Ʊ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    End If
End Sub

Public Function MakeDetailRecord(ByVal strNos As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ��ν�������,������ϸ���ݼ�
    '���:strNOs-���ݺ�,����ö��ŷָ�
    '����:
    '����:�������ݼ�,��ʽ:����ID����ҳID���շ�����շ�ϸĿID�����������ۣ�ʵ�ս������ˣ���������
    '����:���˺�
    '����:2014-09-16 17:20:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim i As Integer, j As Integer, p As Integer, strSQL As String
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl���� As Double, curʵ�� As Currency
    Dim rsTmp As New ADODB.Recordset, rsPrice As ADODB.Recordset
    
    On Error GoTo errHandle
    
    rsTmp.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�շ����", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "������", adVarChar, 100, adFldIsNullable
    '79420,���ϴ�,2014/11/10:������¼���ֶδ�С
    rsTmp.Fields.Append "��������", adVarChar, 100, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    strSQL = "" & _
    "   Select NO,��¼״̬,����ID,nvl(�۸񸸺�,���) as ���,����ID,NULL as ��ҳID,�շ����,�շ�ϸĿID, " & _
    "           Avg(����*Nvl(����,0)) ����,Sum(��׼����) as ����,Sum(ʵ�ս��) ʵ�ս��,max(a.������) as ������,max(C.����) as ��������" & _
    "   From ������ü�¼ A,���ű� C" & _
    "   Where A.NO in (Select Column_Value From Table(f_str2List([1])) ) and A.��������ID=C.ID(+) " & _
    "           And mod(A.��¼����,10)=1 And A.��¼״̬ IN (1,2,3)  " & _
    "   Group By NO,��¼״̬,nvl(�۸񸸺�,���),�շ�ϸĿID,����ID,�շ����,����ID "
    
    Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����շѵ���", strNos)
    With rsPrice
        For i = 1 To .RecordCount
            rsTmp.Filter = "�շ�ϸĿID=" & !�շ�ϸĿID
            If rsTmp.RecordCount = 0 Then
                rsTmp.AddNew
                
                rsTmp!����ID = Nvl(!����ID, mobjPatiInfor.����ID)
                rsTmp!��ҳID = Nvl(!��ҳID, 0)
                rsTmp!�շ���� = !�շ����
                rsTmp!�շ�ϸĿID = !�շ�ϸĿID
                rsTmp!���� = !����
                rsTmp!���� = !����
                rsTmp!ʵ�ս�� = !ʵ�ս��
                rsTmp!������ = !������
                rsTmp!�������� = !��������
            Else
                rsTmp!���� = rsTmp!���� + !����
                rsTmp!���� = (rsTmp!���� + !����) / 2
                rsTmp!ʵ�ս�� = rsTmp!ʵ�ս�� + !ʵ�ս��
            End If
            rsTmp.Update
            .MoveNext
        Next
    End With
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitBalanceGrid(Optional blnOnlyClearBalace As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ս�����
    '���:blnOnlyBalace-�������������Ϣ
    '����:���˺�
    '����:2011-11-02 13:53:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    vsBalance.Clear 1
    vsBalance.Rows = 1
    vsBalance.COLS = 1
    
    vsBalance.ColAlignment(0) = 1
'    vsBalance.ColAlignment(1) = 7
    vsBalance.Row = 0
    vsBalance.Col = 0
    
    vsBalance.TabStop = False
    With vsBalance
        .Cell(flexcpFontBold, 0, 0, 0, .COLS - 1) = False
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .COLS - 1) = Me.ForeColor
    End With
    If mEditType = EM_Balance_Charge And mEditType = EM_Balance_Register Then vsBalance.Editable = flexEDKbdMouse
    For i = 0 To vsBalance.COLS - 1
        vsBalance.ColData(i) = 0
    Next
    If blnOnlyClearBalace Then Exit Sub
    '������㼯����
    Set mcolBalance = New Collection
End Sub

Private Function zlInsureClinicPreSwap(ByVal strNos As String, ByVal strDate As String, _
    ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ԥ����ӿ�
    '���:strNos-��ǰѡ�еĵ���
    '     strDate-����ʱ��
    '����:strNone-��֧�ֵĽ��㷽ʽ
    '����:�ӿڵ��óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-16 17:34:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strBalance As String, strAdvance As String
    Dim varBalance As Variant, varTemp As Variant
    Dim j As Long, strTemp As String
    
    Dim arrPage As Variant, arrBalance() As String, strInvoice As String
    Dim str���㷽ʽ As String, dbl������ As Double, dbl�ɷ���� As Double
    Dim rsTemp  As ADODB.Recordset, i As Long, k As Long, p As Long
    
    On Error GoTo errHandle
    strInvoice = Trim(txtInvoice.Text)
    
    Set rsTemp = MakePreRecord(strNos, strDate, strInvoice)
    
    strAdvance = "2"
    If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, mintInsure, strAdvance) Then
        staThis.Panels(Pan.C2��ʾ��Ϣ).Text = "����Ԥ����ʧ�ܡ�"
        If mstr�����ʻ� <> "" And Not MCPAR.����Ԥ���� Then  'ֻ��ʹ�ø����ʻ�����
            vsBalance.COLS = 3
            vsBalance.TextMatrix(0, 1) = mstr�����ʻ�
            vsBalance.TextMatrix(0, 2) = "0"
            vsBalance.ColData(1) = 0
            vsBalance.ColData(2) = 0
        End If
        Screen.MousePointer = 0
        Exit Function
    End If
    
    If strAdvance <> "" And strAdvance <> "2" Then 'ҽ��Ʊ�ݺ�
        txtMCInvoice.Text = strAdvance
        txtMCInvoice.SelStart = Len(txtMCInvoice.Text)
        txtMCInvoice.Visible = True
    End If
    
    MCPAR.ҽ������Ʊ�� = False
    If InStr(1, strAdvance, ";") > 0 Then
          '38821:strAdvance:��Ʊ��;�Ƿ���Ʊ�ݺ�
          MCPAR.ҽ������Ʊ�� = Val(Split(strAdvance & ";", ";")(1)) = 1
    End If
        

     '����Ԥ���������ý��㼯
    Set mcolBalance = New Collection
    
    With vsBalance
        .Clear 1
        .Rows = 1
        .COLS = 1
        .TextMatrix(0, 0) = "ҽ������"
        
        varBalance = Split(strBalance, "|")
        For i = 0 To UBound(varBalance)
            '������ʽ;���;�Ƿ������޸�
            varTemp = Split(varBalance(i) & ";;;;", ";")
            str���㷽ʽ = varTemp(0)
            dbl������ = Val(varTemp(1))
            
            mrs���㷽ʽ.Filter = "����='" & str���㷽ʽ & "' And  ����>=3 and ����<= 4"
            If mrs���㷽ʽ.EOF Then
                '��¼ҽ���е�����û�еĽ��㷽ʽ
                If InStr(strNone & ",", "," & str���㷽ʽ & ",") = 0 Then
                    strNone = strNone & "," & str���㷽ʽ
                End If
            End If
            If Not mrs���㷽ʽ.EOF And dbl������ <> 0 Then
                .COLS = .COLS + 2
                .TextMatrix(0, .COLS - 2) = str���㷽ʽ
                .TextMatrix(0, .COLS - 1) = FormatEx(dbl������, 6)
                .Cell(flexcpData, 0, .COLS - 1) = dbl������
                .ColData(.COLS - 1) = Val(varTemp(2)) '�Ƿ������޸�
                .ColData(.COLS - 2) = 0
                
                '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;�ĺ���
                strTemp = str���㷽ʽ
                strTemp = strTemp & ";" & dbl������
                strTemp = strTemp & ";" & Val(varTemp(2))
                strTemp = strTemp & ";" & GetYBActualMoeny(str���㷽ʽ, dbl������)
                mcolBalance.Add strTemp
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        For i = 0 To .COLS - 1
            If .ColData(i) <> 0 Then
                .Row = 0:  .Col = i: .TabStop = True
            End If
            If i > 0 And i Mod 2 = 0 Then vsBalance.ColWidth(i) = 1000
        Next
    End With
    
    zlInsureClinicPreSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function MakePreRecord(ByVal strNos As String, ByVal str����ʱ�� As String, ByVal strInvoice As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݶ������ݴ���һ����¼��Ϣ(���ۼ۵�λ)
    '���:strNos-��ǰ�ĵ�����Ϣ
    '     str����ʱ��=����ʱ��(yyyy-mm-dd HH:MM:SS)
    '     strInvoice=Ʊ�ݺ�
    '����:
    '����:ҽ��������ݵ����ݼ�(�������(1--n),����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��)
    '����:���˺�
    '����:2011-08-15 16:40:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, intStartPage As Integer, intPages As Integer
    Dim p As Integer, strSQL As String
    Dim dbl���� As Double, curʵ�� As Currency, curͳ�� As Currency
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    Dim strAllNOs As String
    
    Err = 0: On Error GoTo Errhand:
    rsTmp.Fields.Append "�������", adBigInt, 50, adFldIsNullable
    rsTmp.Fields.Append "�ѱ�", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsTmp.Fields.Append "���", adBigInt, , adFldIsNullable '����:42961
    rsTmp.Fields.Append "ʵ��Ʊ��", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "����ʱ��", adDBTimeStamp, , adFldIsNullable
    rsTmp.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�շ����", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "���㵥λ", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "ͳ����", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "����֧������ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�Ƿ�ҽ��", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "���ձ���", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "ժҪ", adVarChar, 2000, adFldIsNullable
    rsTmp.Fields.Append "�Ƿ���", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "��������ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "ִ�в���ID", adBigInt, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    strSQL = _
    "Select '" & strInvoice & "' as ʵ��Ʊ��,NO,��¼״̬,Nvl( �۸񸸺�, ���) as ���,To_Date('" & str����ʱ�� & "','YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
    "       max(A.����ID) as ����ID ,max(A.�ѱ�) As �ѱ�,�շ����,�վݷ�Ŀ,���㵥λ,������," & _
    "       �շ�ϸĿID,���մ���ID As ����֧������ID,Nvl(������Ŀ��,0) As �Ƿ�ҽ��,���ձ���," & _
    "       Avg(Nvl(����,0)*����) As ����,Avg(��׼����) As ����," & _
    "       Sum(ʵ�ս��) As ʵ�ս��,Sum(ͳ����) As ͳ����,ժҪ," & _
    "       max(�Ӱ��־) as �Ƿ���,��������ID,ִ�в���ID,����ID " & _
    "       From ������ü�¼ a" & _
    "   Where ��¼����=1 And A.NO in (Select Column_Value From Table(f_str2List([1])) ) " & _
    " Group By NO,��¼״̬,Nvl(�۸񸸺�,���),�շ����,�վݷ�Ŀ,���㵥λ,������," & _
    "       �շ�ϸĿID,���մ���ID,Nvl(������Ŀ��,0),���ձ���,ժҪ,��������ID,ִ�в���ID,����ID" & _
    " Order by  NO,���,��¼״̬ "
    Set rsNo = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���۵�����-ҽ��", strNos)
    If rsNo.RecordCount <> 0 Then rsNo.MoveFirst
    p = 0
    Do While Not rsNo.EOF
        rsTmp.AddNew
        If InStr(strAllNOs & ",", "," & Nvl(rsNo!NO) & ",") = 0 Then p = p + 1
        
        rsTmp!������� = p
        rsTmp!�ѱ� = Nvl(rsNo!�ѱ�)
        rsTmp!NO = Nvl(rsNo!NO)   '����ȡ���۵�ʱ����ֵ
        rsTmp!��� = Val(Nvl(rsNo!���))   '����ȡ���۵�ʱ����ֵ
        rsTmp!ʵ��Ʊ�� = strInvoice
        rsTmp!����ʱ�� = CDate(str����ʱ��)
        rsTmp!����ID = Nvl(rsNo!����ID)
        rsTmp!�շ���� = Nvl(rsNo!�շ����)
        rsTmp!�վݷ�Ŀ = Nvl(rsNo!�վݷ�Ŀ)
        rsTmp!������ = Nvl(rsNo!������)
        rsTmp!�շ�ϸĿID = Val(Nvl(rsNo!�շ�ϸĿID))
        rsTmp!���㵥λ = Nvl(rsNo!���㵥λ)
        rsTmp!���� = Val(Nvl(rsNo!����))
        rsTmp!���� = Val(Nvl(rsNo!����))
        rsTmp!ʵ�ս�� = Val(Nvl(rsNo!ʵ�ս��))
        rsTmp!ͳ���� = Val(Nvl(rsNo!ͳ����))
        rsTmp!����֧������ID = IIf(Val(Nvl(rsNo!����֧������ID)) = 0, Null, Val(Nvl(rsNo!����֧������ID)))
        rsTmp!�Ƿ�ҽ�� = Val(Nvl(rsNo!�Ƿ�ҽ��))
        rsTmp!���ձ��� = Nvl(rsNo!���ձ���)
        rsTmp!ժҪ = Nvl(rsNo!ժҪ)
        rsTmp!�Ƿ��� = Val(Nvl(rsNo!�Ƿ���))
        rsTmp!��������ID = Val(Nvl(rsNo!��������ID))
        rsTmp!ִ�в���ID = Val(Nvl(rsNo!ִ�в���ID))
        rsTmp.Update
        If InStr(1, strAllNOs & ",", "," & Nvl(rsNo!NO) & ",") = 0 Then
            strAllNOs = strAllNOs & "," & Nvl(rsNo!NO)
        End If
        rsNo.MoveNext
    Loop
                 
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakePreRecord = rsTmp
    Exit Function
Errhand::
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetYBActualMoeny(ByVal str���㷽ʽ As String, ByVal dbl������ As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�������ʵ��ʹ�ý��
    '���:str���㷽ʽ-ҽ���Ľ��㷽ʽ
    '     dbl������-ҽ���Ľ�����
    '����:ʵ�ʽ��,���򷵻ش����dbl������
    '����:���˺�
    '����:2014-09-16 18:05:13
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim dbl���� As Double, dbl���ʺϼ� As Double
    
    On Error GoTo errHandle
    
    If dbl������ = 0 Then Exit Function
    If str���㷽ʽ <> mstr�����ʻ� Then GetYBActualMoeny = dbl������: Exit Function
    
    '����ҽ���޷��������
     If (mTy_Insure.dbl�ʻ���� > -1 * mTy_Insure.dbl����͸֧ Or mintInsure = 61) _
        And CCur(lblʵ��.Tag) > 0 Then
        dbl���� = dbl������
        If mintInsure <> 61 Then
            '��������ʻ�֧�����
            If RoundEx(mTy_Insure.dbl�ʻ���� - dbl���ʺϼ� - dbl����, 6) = -1 * mTy_Insure.dbl����͸֧ Then
                dbl���� = dbl���� '������͸֧��Χ���㹻(����͸֧0Ϊ����)
            Else
                If mTy_Insure.dbl����͸֧ = 0 And RoundEx(mTy_Insure.dbl�ʻ���� - dbl���ʺϼ�, 6) > 0 Then
                    dbl���� = mTy_Insure.dbl�ʻ���� - dbl���ʺϼ� '������͸֧�������
                Else
                    '��������͸֧��Χ������͸֧ʱ�����
                    If mTy_Insure.dbl����͸֧ <> 0 Then
                        dbl���� = mTy_Insure.dbl�ʻ���� - dbl���ʺϼ� + mTy_Insure.dbl����͸֧ '������͸֧��Χ��֧��
                    Else
                        dbl���� = 0
                    End If
                End If
            End If
        End If
        dbl���ʺϼ� = dbl���ʺϼ� + dbl����
        dbl���� = Format(dbl����, "0.00")
        GetYBActualMoeny = dbl����
    Else
        GetYBActualMoeny = dbl������
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    GetYBActualMoeny = dbl������
End Function

Private Sub zl9InsureLedSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ��Ԥ��Led����
    '����:���˺�
    '����:2014-09-18 13:43:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���ʺϼ� As Double
    If Not gblnLED Then Exit Sub
    dbl���ʺϼ� = GetMedicareBalanceSum(mstr�����ʻ�)
    zl9LedVoice.DisplayBank "ҽ������:", "�ʻ����" & Format(mTy_Insure.dbl�ʻ����, "0.00"), "�ʻ�֧��" & Format(dbl���ʺϼ�, "0.00"), "ͳ��֧��" & Format(GetMedicareBalanceSum - dbl���ʺϼ�, "0.00")
    zl9LedVoice.Speak "#21 " & Format(-1 * GetMedicareBalanceSum, "0.00")
End Sub

Public Function GetMedicareBalanceSum(Optional strItem As String, Optional blnOrig As Boolean) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ս���Ľ��
    '   strItem=�Ƿ�ָ�����㷽ʽ,����Ϊ���н��㷽ʽ
    '   blnOrig=�Ƿ�ȡԭʼ(���)������,����ȡ����(�޸ĺ�)��Ч���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-18 13:44:21
    '˵�����ú�����mcolBalanceΪ׼����,����ҽ�������շ�Ҳ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, curMoney As Currency
    Dim i As Integer
    For i = 1 To mcolBalance.Count
        '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;��Ч���
        varData = Split(mcolBalance(i), ";")
        If strItem = "" Or (strItem <> "" And varData(0) = strItem) Then
            If blnOrig Then
                curMoney = curMoney + CCur(varData(1))
            Else
                curMoney = curMoney + CCur(varData(3))
            End If
        End If
    Next
    GetMedicareBalanceSum = Format(curMoney, "0.00")
End Function

Private Function GetMedicareBalanceStr(ByRef cur���� As Currency) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ս�������
    '����:cur����-���عҺ�ʱ�ĸ����ʻ�֧��
    '����:���ر��ս��㷽ʽ��,"���㷽ʽ,���|...."
    '����:���˺�
    '����:2014-09-17 16:01:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strTemp As String
    Dim varData As Variant
    strTemp = ""
    cur���� = 0
    If mEditType = EM_Balance_Register Then
        With vsBalance
            For i = 1 To .COLS - 1 Step 2
                If .TextMatrix(0, i) = mstr�����ʻ� Then
                    cur���� = Val(.TextMatrix(0, i + 1))
                    strTemp = strTemp & "|" & .TextMatrix(0, i) & "," & Format(Val(.TextMatrix(0, i + 1)), "0.00")
                End If
            Next
        End With
    Else
        For i = 1 To mcolBalance.Count
            '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;��Ч���
            varData = Split(mcolBalance(i), ";")
            strTemp = strTemp & "|" & varData(0) & "," & Format(varData(3), "0.00")
        Next
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    GetMedicareBalanceStr = strTemp
End Function

Private Function IsRegister(Optional ByRef strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϵ�ǰ�����ǹҺŷ���
    '����:strNO-�Һŵ���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-28 12:37:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnRegister As Boolean
    
    On Error GoTo errHandle
    strNo = ""
    If mrsList Is Nothing Then Exit Function
    If mrsList.State <> 1 Then Exit Function
    If mrsList.RecordCount = 0 Then Exit Function
    mrsList.Filter = "��¼����=1"
    blnRegister = mrsList.RecordCount = 0
    mrsList.Filter = 0
    If blnRegister Then
        If Not mrsList.EOF Then strNo = Nvl(mrsList!NO)
    End If
    IsRegister = blnRegister
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CancelBalance() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�쳣����
    '����:���˺�
    '����:2014-06-19 14:42:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str����ID As String, dtDelDate As Date
    Dim blnTrans As Boolean, strSQL As String
    Dim cllPro As Collection, strRegNO As String '�Һŵ���
    Dim blnReg As Boolean
    
    '�������
    If zlIsCheckExistErrBill(Val(mstr�������), True) = False Then
        MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(Val(mstr�������)) Then
        MsgBox "��ǰ�����������������㴰���н��д����㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    blnReg = IsRegister(strRegNO)
    dtDelDate = zlDatabase.Currentdate
    str����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    
    
    Set cllPro = New Collection
    'Zl_���ò����¼_����������
    strSQL = "Zl_���ò����¼_����������("
    '  No_In         In ���ò����¼.No%Type,
    strSQL = strSQL & "'" & mstrNo & "',"
    '  ����id_In     In ���ò����¼.����id%Type,
    strSQL = strSQL & "" & str����ID & ","
    '  �������_In   In ���ò����¼.�������%Type,
    strSQL = strSQL & "" & "-" & str����ID & ","
    '  ����Ա���_In In ���ò����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In In ���ò����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �Ǽ�ʱ��_In   In ���ò����¼.�Ǽ�ʱ��%Type := Null,
    strSQL = strSQL & "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    zlAddArray cllPro, strSQL
    Err = 0: On Error GoTo Errhand:
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If blnReg Then
        '���ùҺ�����
        If ExcuteInsureRegistDel(mintInsure, mobjPatiInfor.����ID, mstr����ID, str����ID, strRegNO) = False Then Exit Function
    Else
        If ExcuteInsureDel(mintInsure, mobjPatiInfor.����ID, mstr����ID, str����ID) = False Then Exit Function
    End If
    blnTrans = False: CancelBalance = True
    Exit Function
Errhand:
   If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ExcuteInsureRegistDel(ByVal intInsure As Integer, ByVal lng����ID As Long, _
    ByVal strԭ����ID As String, ByVal str����ID As Long, ByVal strRegNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ���˺Žӿ�
    '����:���óɹ�,����true,���򷵻�False
    '����:Ƚ����
    '����:2014-10-27
    '˵��:��Ҫ�������������;
    '     ���ʧ��,�����񽫻���(��Ҫ�Ǳ��ⵯ�������������)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, blnTransMedicare As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    strAdvance = "0|" & strRegNO & "|1"
    If Not gclsInsure.RegistDelSwap(Val(strԭ����ID), intInsure, strAdvance) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    blnTransMedicare = True
    'Zl_���ò������_Modify
    strSQL = "Zl_���ò������_Modify("
    '  ��������_In   Number,
    strSQL = strSQL & "" & "2" & ","
    '  ����id_In     In ���ò����¼.����id%Type,
    strSQL = strSQL & "" & str����ID & ","
    '  ���㷽ʽ_In   Varchar2,:���㷽ʽ|������||.."
    strSQL = strSQL & "NULL,"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ��ɽ���_In Number:=0
    strSQL = strSQL & "2)" '1-��ɲ������;0-δ��ɲ������;2-������쳣����
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistDelSwap, True, intInsure)
    
    ExcuteInsureRegistDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistDelSwap, False, intInsure)
    Call ErrCenter
End Function

Private Function ExcuteInsureDel(ByVal intInsure As Integer, ByVal lng����ID As Long, _
    ByVal strԭ����ID As String, ByVal str����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ���˷ѽӿ�
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-18 18:20:38
    '˵��:��Ҫ�������������,�����˷Ѻ�,�ù��̲��ύ,��Ҫ�������ύ;
    '     ���ʧ��,�����񽫻���(��Ҫ�Ǳ��ⵯ�������������)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, blnTransMedicare As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strAdvance = str����ID & "|1"
    
    blnTransMedicare = False
    If Not gclsInsure.ClinicDelSwap(Val(strԭ����ID), , intInsure, strAdvance) Then
         gcnOracle.RollbackTrans: Exit Function
    End If
    blnTransMedicare = True
    If Val(strAdvance) = str����ID Or strAdvance = "" Then
        'Zl_���ò������_Modify
        strSQL = "Zl_���ò������_Modify("
        '  ��������_In   Number,
        strSQL = strSQL & "" & "2" & ","
        '  ����id_In     In ���ò����¼.����id%Type,
        strSQL = strSQL & "" & str����ID & ","
        '  ���㷽ʽ_In   Varchar2,:���㷽ʽ|������||.."
        strSQL = strSQL & "NULL,"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��ɽ���_In Number:=0
        strSQL = strSQL & "2)" '1-��ɲ������;0-δ��ɲ������;2-������쳣����
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
        gcnOracle.CommitTrans: ExcuteInsureDel = True
        Exit Function
    End If
    '���ݷ��صĽ�����Ϣ������Ԥ����¼��strAdvance���ظ�ʽ:���㷽ʽ1|���||���㷽ʽ2|���...
    If InStr(strAdvance, "|") > 0 Then
        '���±�־:
        'Zl_���ò������_Modify
        strSQL = "Zl_���ò������_Modify("
        '  ��������_In   Number,
        strSQL = strSQL & "" & "2" & ","
        '  ����id_In     In ���ò����¼.����id%Type,
        strSQL = strSQL & "" & str����ID & ","
        '  ���㷽ʽ_In   Varchar2,:���㷽ʽ|������||.."
        strSQL = strSQL & "'" & strAdvance & "',"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��ɽ���_In Number:=0
        strSQL = strSQL & "2)" '1-��ɲ������;0-δ��ɲ������;2-������쳣����
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
    gcnOracle.CommitTrans
    ExcuteInsureDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intInsure)
    Call ErrCenter
End Function

Public Function GetBalanceInsure(ByVal str����ID As String, _
    ByRef str�������� As String, Optional ByRef lng����ID As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������
    '����:str��������-��������
    '     lng����ID-����ID
    '����:��������
    '����:���˺�
    '����:2014-09-22 13:57:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errH
    strSQL = "" & _
    "    Select /*+ rule */  B.��¼ID,B.����,B.����ID,C.����" & _
    "    From �����ս����¼ B,������� C" & _
    "    Where B.��¼ID=[1] and B.����=C.���(+) And B.����=1  "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����ID)
    If Not rsTmp.EOF Then
        lng����ID = Nvl(rsTmp!����ID, 0)
        str�������� = Nvl(rsTmp!����)
        GetBalanceInsure = Nvl(rsTmp!����, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsDiagnose_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�޸ĺ�,��Ҫ���¹���
    If Not mEditType = EM_Balance_Charge Then Exit Sub
    'ѡ��
    Call FromDiagnoseSelFee
    
    mblnEdit = True
    Call SetButtons
End Sub

Private Sub FromDiagnoseSelFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ѡ��ָ���ķ���
    '����:���˺�
    '����:2014-09-26 16:23:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, varData As Variant, strTemp As String
    Dim i As Long, j As Long, q As Long, blnHaversBalanceNo As Boolean
    Dim lng������� As Long
    
    blnHaversBalanceNo = False
    If Not mrsBalanceNO Is Nothing Then
        blnHaversBalanceNo = mrsBalanceNO.State = 1
    End If
    '��ȫ��ѡ�е�(��ɫ����),��Ϊѡ��
    Call SetDiagnoseSelStatu(EM_dgGrayToSeled)
    
    '��ȡ����ѡ�е��������Ӧ��Nos
    strNos = GetDiagnoseNos

    '�������ѡ��
    Call FromNosSel("", False, False, True)
    Call FromNosSel(strNos, True, True)
End Sub

Private Function GetDiagnoseNos() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡѡ�����ϵĵ��ݺ�
    '����:�ɹ�,�����������Ӧ��NOs
    '����:���˺�
    '����:2014-09-28 10:39:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNos As String, varData As Variant, strTemp As String
    Dim i As Long, j As Long, q As Long, blnHaversBalanceNo As Boolean
    
    blnHaversBalanceNo = False
    If Not mrsBalanceNO Is Nothing Then
        blnHaversBalanceNo = mrsBalanceNO.State = 1
    End If
    On Error GoTo errHandle
    
    '��ȡѡ�е���ϵĵ��ݺ�(����ö��ŷ���)
    strNos = ""
    With vsDiagnose
        For i = 0 To .Rows - 1
            For j = 0 To .COLS - 1
                If Abs(Val(.Cell(flexcpChecked, i, j))) = 1 And vsDiagnose.TextMatrix(i, j) <> "" Then
                    strTemp = vsDiagnose.Cell(flexcpData, i, j)
                    If strTemp <> "" Then
                         varData = Split(strTemp, ",")
                         For q = 0 To UBound(varData)
                             Call GetRelatedNos(varData(q), strNos)
                         Next
                    End If
                End If
            Next
        Next
    End With
    GetDiagnoseNos = strNos
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetRelatedNos(ByVal strNo As String, ByRef strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ�����ݣ���ȡ�������ݺ�
    '���:strNO-��ǰ���ݺ�
    '����:strNos-���ع����ĵ��ݺ�(��������)
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2014-09-28 10:20:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnHaversBalanceNo As Boolean, varData As Variant
    Dim str������� As String, str������ż� As String
    On Error GoTo errHandle
    
    blnHaversBalanceNo = False
    If Not mrsBalanceNO Is Nothing Then
        blnHaversBalanceNo = mrsBalanceNO.State = 1
    End If
    If Not blnHaversBalanceNo Then GoTo CurNOs:
    
    mrsBalanceNO.Filter = "NO='" & strNo & "'"
    str������ż� = ""
    Do While Not mrsBalanceNO.EOF
        str������� = ""
        If Not mrsBalanceNO.EOF Then str������� = Val(Nvl(mrsBalanceNO!�������))
        If str������� <> 0 And InStr(str������ż� & ",", "," & str������� & ",") = 0 Then
            str������ż� = str������ż� & "," & str�������
        End If
        mrsBalanceNO.MoveNext
    Loop
    mrsBalanceNO.Filter = 0
    If str������ż� = "" Then GoTo CurNOs:
    
    str������ż� = Mid(str������ż�, 2)
    varData = Split(str������ż�, ",")
    
    For i = 0 To UBound(varData)
        mrsBalanceNO.Filter = "�������=" & IIf(varData(i) = "", "0", varData(i))
        Do While Not mrsBalanceNO.EOF
             If InStr(1, "," & strNos & ",", "," & mrsBalanceNO!NO & ",") = 0 Then
                strNos = strNos & "," & mrsBalanceNO!NO
             End If
            mrsBalanceNO.MoveNext
        Loop
    Next
    
    If InStr(1, "," & strNos & ",", "," & strNo & ",") = 0 Then
       strNos = strNos & "," & strNo
    End If
    If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    
    GetRelatedNos = True
    Exit Function
CurNOs:
    strNos = strNos & "," & strNo
    If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    GetRelatedNos = True: Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetDiagnoseSelStatu(ByVal intStatu As mEM_Diagnose_SelStatu)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϵ�ѡ��״̬
    '���:intStatu=-1����ѡ��Ļ�ɫ��Ϊѡ��
    '     intStatu=0  �������ѡ�е����
    '     intStatu=1  ѡ�����е����
    '     intStatu=5  ȫ����ѡ�е�����Ϊ��ɫ
    '����:���˺�
    '����:2014-09-28 10:00:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, intCurStatu As Integer
    With vsDiagnose
        For i = 0 To .Rows - 1
            For j = 0 To .COLS - 1
                intCurStatu = Abs(Val(.Cell(flexcpChecked, i, j)))
                If intCurStatu = 0 Then intCurStatu = 2
                '����ɫ�ĵ���Ϊ
                Select Case intStatu
                Case EM_dgGrayToSeled '��ѡ��Ļ�ɫ��Ϊѡ��
                     If Abs(intCurStatu) = 5 Then intCurStatu = -1
                Case EM_dgClearAllSeled   '�������ѡ�е����
                    intCurStatu = 2
                Case EM_dgClearAllSeled  'ѡ�����е����
                    intCurStatu = -1
                Case Else '5-ȫ����ѡ�е�����Ϊ��ɫ
                    If Abs(intCurStatu) = 1 Then intCurStatu = 5
                End Select
                .Cell(flexcpChecked, i, j) = intCurStatu
            Next
        Next
    End With
End Sub

Private Sub FromNosSel(ByVal strNos As String, ByVal blnSel As Boolean, _
    ByVal blnBeforClearSel As Boolean, Optional blnAllNo As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݺ�ѡ��
    '���:blnSel-ѡ��
    '     blnBeforClearSel-���������ѡ��
    '     blnAllNo-�����ֵ��ݽ��д���
    '����:���˺�
    '����:2014-09-26 17:18:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, i As Long
    
    'ѡ�����е���
    With vsFeeList
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then
                strNo = Trim(.TextMatrix(i, .ColIndex("NO")))
                If strNo <> "" Then
                    If blnAllNo Then
                        .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = IIf(blnSel, -1, 2)
                    Else
                        If InStr(1, "," & strNos & ",", "," & strNo & ",") > 0 Then
                            .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = IIf(blnSel, -1, 2)
                        ElseIf blnBeforClearSel Then
                            .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = 2
                        End If
                    End If
                End If
            End If
        Next
    End With
    Call CalcTotalMoney
    Call CalcRegisterYBMoney
    mblnEdit = True
End Sub

Private Sub vsDiagnose_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub vsFeeList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, strNo As String, strNos As String
    
    Dim blnSel As Boolean
    With vsFeeList
        If .Col <> .ColIndex("ѡ��") Then Exit Sub
        strNo = Trim(.TextMatrix(Row, .ColIndex("NO")))
        If strNo = "" Then Exit Sub
        
        If mEditType = EM_Balance_Register Then
            'ֻ��ѡ��һ���Һŵ�
            If Abs(Val(.Cell(flexcpChecked, Row, Col))) <> 1 Then
                mblnEdit = True
                Call SetButtons
                Call CalcTotalMoney
                Call CalcRegisterYBMoney
                Exit Sub
            End If
            
            For i = 1 To .Rows - 1
                If .IsSubtotal(i) And i <> Row Then
                    .Cell(flexcpChecked, i, Col) = 2
                End If
            Next
        End If
        blnSel = Abs(Val(.Cell(flexcpChecked, Row, Col))) = 1
        '����ѡ��
        Call GetRelatedNos(strNo, strNos)
        
        Call FromNosSel(strNos, blnSel, False)
        '��ѡ�е���Ϊ��ɫ״̬
        Call SetDiagnoseSelStatu(EM_dgSeledToGray)
        
        mblnEdit = True
        Call SetButtons
        Call CalcTotalMoney
        Call CalcRegisterYBMoney
    End With
End Sub

Private Sub vsFeeList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, mstrTittle, "������Ϣ�б�", True, False
End Sub

Private Sub vsFeeList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsFeeList, mstrTittle, "������Ϣ�б�", True, False
End Sub

Private Sub vsFeeList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    With vsFeeList
        If .Col <> .ColIndex("ѡ��") Then Cancel = True: Exit Sub
        If .IsSubtotal(Row) = False Then Cancel = True: Exit Sub
        If .ColIndex("���") < 0 Then Cancel = True: Exit Sub
        If Trim(.TextMatrix(Row, .ColIndex("���"))) = "" Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsFeeList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    With vsFeeList
        If Col <= .ColIndex("ѡ��") Then
             Position = Col
        End If
    End With
End Sub

Private Sub vsFeeList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsFeeList
        If Col <= .ColIndex("ѡ��") Then Cancel = True
    End With
End Sub

Private Sub vsFeeList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cmdԤ����.Enabled And cmdԤ����.Visible Then
        cmdԤ����.SetFocus
    ElseIf cmdOK.Visible And cmdOK.Enabled Then
        cmdOK.SetFocus
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Dim lngR As Long
    If Panel.Key = "Calc" Then
        lngR = FindWindow("SciCalc", "������")
        If lngR <> 0 Then
            BringWindowToTop lngR
        Else
            On Error Resume Next
            Shell "calc.exe", vbNormalFocus
        End If
    End If
End Sub

Private Sub SetButtons()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù��ܰ�ť
    '����:���˺�
    '����:2014-09-23 11:51:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cmdCancel.Enabled = True: cmdCancel.Visible = True
    
    If mEditType = EM_Balance_Register Or mEditType = EM_Balance_Err_ReCharge Or mEditType = EM_Balance_Err_Cancel Then
        cmdԤ����.Visible = False: cmdԤ����.Enabled = False
        cmdOK.Enabled = True: cmdOK.Visible = True
        cmdSelAll.Visible = False: cmdClear.Visible = False
        Call picDown_Resize
        Exit Sub
    End If
    If mobjPatiInfor Is Nothing Then
        cmdOK.Enabled = False: cmdԤ����.Visible = False: cmdԤ����.Enabled = False
        cmdSelAll.Visible = False: cmdClear.Visible = False
        Exit Sub
    End If
    cmdSelAll.Visible = True: cmdClear.Visible = True
    
    '֧��Ԥ����ʱ�Ͳ��̶���ʾ�����ʻ�,������ʾ
    If MCPAR.����Ԥ���� Then
        '��ʾԤ���㰴ť
        cmdԤ����.Enabled = mblnEdit  '�Ƿ�༭����δ����Ԥ�����
        cmdԤ����.Visible = True
        cmdOK.Enabled = Not mblnEdit: cmdOK.Visible = True
        Call picDown_Resize
        Exit Sub
    End If
    If mstr�����ʻ� <> "" Then 'ֻ��ʹ�ø����ʻ�����
        cmdԤ����.Visible = False: cmdԤ����.Enabled = False
        cmdOK.Enabled = True: cmdOK.Visible = True
        Call picDown_Resize
    End If
End Sub

Private Sub CalcTotalMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������صĺϼƽ��
    '����:���˺�
    '����:2014-09-23 12:09:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnSel As Boolean
    Dim dblMoney(0 To 1) As Double
    Dim blnTotalSelect As Boolean
    
    lblӦ��.Tag = "": lblʵ��.Tag = ""
    With vsFeeList
        dblMoney(0) = 0: dblMoney(1) = 0
        For i = 1 To .Rows - 1
            If mEditType = EM_Balance_Err_Cancel Or mEditType = EM_Balance_Err_ReCharge Then
                blnSel = True
            Else
                blnSel = Abs(Val(.Cell(flexcpChecked, i, .ColIndex("ѡ��")))) = 1
            End If
            
            If .IsSubtotal(i) Then
                If blnSel Then
                    dblMoney(0) = dblMoney(0) + Val(.Cell(flexcpData, i, .ColIndex("Ӧ�ս��")))
                    dblMoney(1) = dblMoney(1) + Val(.Cell(flexcpData, i, .ColIndex("ʵ�ս��")))
                End If
                blnTotalSelect = blnSel '��¼������ѡ��״̬
            Else
                '�����б�ѡ�����������Ҳ�ͱ�ѡ��
                If blnTotalSelect Then
                    'ͳ�Ʋ����ѣ�84965
                    If mEditType = EM_Balance_Register _
                        And Val(.TextMatrix(i, .ColIndex("�շ�ϸĿID"))) = mlng������ϸĿID Then
                        mcur������ = mcur������ + Val(.Cell(flexcpData, i, .ColIndex("ʵ�ս��")))
                    End If
                End If
            End If
        Next
    End With
    lblӦ��.Caption = "Ӧ��:" & Format(dblMoney(0), "0.00")
    lblӦ��.Tag = dblMoney(0)
    lblʵ��.Caption = "ʵ��:" & Format(dblMoney(1), "0.00")
    lblʵ��.Tag = dblMoney(1)
    
    '���Ԥ������Ϣ
    vsBalance.Clear 1: vsBalance.COLS = 1: txt�˿�ϼ�.Text = "0.00"
End Sub

Private Sub CalcRegisterYBMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���㲢��ʾ�Һŵ�ǰҽ�����˸����ʻ�����֧�ֵĽ��
    '����:���˺�
    '����:2014-09-23 12:07:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur�ϼ� As Currency
    Dim strInfo As String, i As Long, j As Long, lng����ID As Long
    Dim dbl�����ʻ� As Double
    Dim blnFind As Boolean
    
    '80238,Ƚ����,2014-11-27
    If mEditType <> EM_Balance_Register Then Exit Sub
    If mstrYBPati <> "" Then lng����ID = Val(Split(mstrYBPati, ";")(8))
    
    cur�ϼ� = Val(lblʵ��.Tag)
    If MCPAR.���ղ����� Then '���ղ����ѣ�84965
        cur�ϼ� = cur�ϼ� - mcur������
    End If
    '���㲢��ʾ�����ʻ�֧�����
    'Ҫ��ҽ��֧�ָ����ʻ�֧����ZLHIS����ʹ�ø����ʻ�
    dbl�����ʻ� = 0
    If mintInsure <> 0 And mstr�����ʻ� <> "" Then
        If gclsInsure.GetCapability(support�Һ�ʹ�ø����ʻ�, lng����ID, mintInsure) Then
            If mTy_Insure.dbl�ʻ���� - cur�ϼ� >= -1 * mTy_Insure.dbl����͸֧ Then
               dbl�����ʻ� = Format(cur�ϼ�, "0.00")  '������͸֧��Χ���㹻(����͸֧0Ϊ����)
            Else
                If mTy_Insure.dbl����͸֧ = 0 And mTy_Insure.dbl�ʻ���� > 0 Then
                    dbl�����ʻ� = mTy_Insure.dbl�ʻ����  '������͸֧�������
                Else
                    dbl�����ʻ� = 0 '��������͸֧��Χ������͸֧ʱ�����
                End If
            End If
        End If
    End If
    blnFind = False
    With vsBalance
        .Clear 1
        .Rows = 1
        .COLS = 1
        If blnFind = False Then
            j = -1
            For i = 1 To .COLS - 1 Step 2
                If .TextMatrix(0, i) = "" Then j = i: Exit For
            Next
            If j < 0 Then .COLS = .COLS + 2: j = .COLS - 2
            .TextMatrix(0, i) = mstr�����ʻ�
            .TextMatrix(0, i + 1) = Format(dbl�����ʻ�, "0.00")
        End If
        txt�˿�ϼ� = Format(dbl�����ʻ�, "0.00")
    End With
End Sub

Private Function SaveItemYbMoney(ByVal lng����ID As Long, ByVal strNos As String, _
    ByVal int��¼���� As Integer, Optional ByRef curȫ�Ը� As Currency, _
    Optional ByRef cur���Ը� As Currency, Optional ByRef cur����ͳ�� As Currency) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸ĹҺŻ��շ���Ŀ��ͳ����
    '���:strNo-����,����ö��ŷ���
    '     int��¼����-1-�շ�;4-�Һ�
    '����:curȫ�Ը�-ȫ�Էѽ��
    '     cur���Ը�-���Ը����
    '     cur����ͳ��-ͳ����
    '����:�޸ĳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-23 14:41:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim curʵ�� As Currency, cllPro As Collection
    Dim varData As Variant, strInfo As String
    Dim rsItem As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select ID,NO,��¼״̬,�շ�ϸĿid, ������ĿID, ʵ�ս�� As ʵ�� " & _
    "   From ������ü�¼  " & _
    "   Where NO in (Select Column_Value From Table(f_str2List([1]))) " & _
    "       And ��¼����=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos, int��¼����)
    
    If rsTemp.EOF Then Exit Function
    Set rsItem = New ADODB.Recordset
    
    rsItem.Fields.Append "NO", adVarChar, 100, adFldIsNullable
    rsItem.Fields.Append "�շ�ϸĿid", adBigInt, , adFldIsNullable
    rsItem.Fields.Append "������ĿID", adBigInt, , adFldIsNullable
    rsItem.Fields.Append "ʵ�ս��", adDouble, , adFldIsNullable
    rsItem.Fields.Append "������Ŀ��", adBigInt, , adFldIsNullable
    rsItem.Fields.Append "���մ���id", adBigInt, , adFldIsNullable
    rsItem.Fields.Append "���ձ���", adVarChar, 100, adFldIsNullable
    rsItem.Fields.Append "��������", adVarChar, 100, adFldIsNullable
    rsItem.Fields.Append "ժҪ", adVarChar, 2000, adFldIsNullable
    rsItem.Fields.Append "ͳ����", adDouble, , adFldIsNullable
    rsItem.CursorLocation = adUseClient
    rsItem.LockType = adLockOptimistic
    rsItem.CursorType = adOpenStatic
    rsItem.Open
        
    Do While Not rsTemp.EOF
        rsItem.Filter = "NO='" & Nvl(rsTemp!NO) & "' And �շ�ϸĿid=" & Val(Nvl(rsTemp!�շ�ϸĿID)) & " And ������ĿID=" & Val(Nvl(rsTemp!������ĿID))
        If rsItem.EOF Then
            rsItem.AddNew
            rsItem!NO = CStr(Nvl(rsTemp!NO))
            rsItem!�շ�ϸĿID = Val(Nvl(rsTemp!�շ�ϸĿID))
            rsItem!������ĿID = Val(Nvl(rsTemp!������ĿID))
        End If
        rsItem!ʵ�ս�� = Val(Nvl(rsItem!ʵ�ս��)) + Val(Nvl(rsTemp!ʵ��))
        rsItem.Update
        rsTemp.MoveNext
    Loop
    rsItem.Filter = 0
    Set cllPro = New Collection
    If rsItem.RecordCount <> 0 Then rsItem.MoveFirst
    Do While Not rsItem.EOF
        curʵ�� = Val(Nvl(rsItem!ʵ�ս��))
        strInfo = gclsInsure.GetItemInsure(lng����ID, Val(Nvl(rsItem!�շ�ϸĿID)), curʵ��, True, mintInsure)
        If strInfo <> "" Then
            '������Ŀ��(0/1);���մ���ID;����ͳ����;������Ŀ����;ժҪ;��������
            varData = Split(strInfo & ";;;;;", ";")
            rsItem!������Ŀ�� = Val(varData(0))
            rsItem!���մ���ID = Val(varData(1))
            rsItem!ͳ���� = Val(varData(2))
            rsItem!���ձ��� = Trim(varData(3))
            rsItem!ժҪ = Trim(varData(4))
            rsItem!�������� = Trim(varData(5))
            If Val(varData(2)) = 0 Or Val(varData(0)) = 0 Then
                '��ԭʼ���Ϊ׼,���ֱܷҴ���
                curȫ�Ը� = curȫ�Ը� + curʵ��
            Else
                cur����ͳ�� = cur����ͳ�� + Val(varData(2))
                '��ԭʼ���Ϊ׼,���ֱܷҴ���
                cur���Ը� = cur���Ը� + (curʵ�� - cur����ͳ��)
            End If
            rsItem.Update
        Else
            curȫ�Ը� = curȫ�Ը� + curʵ��
        End If
        rsItem.MoveNext
    Loop
    
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        rsItem.Filter = "NO='" & Nvl(rsTemp!NO) & "' And �շ�ϸĿid=" & Val(Nvl(rsTemp!�շ�ϸĿID)) & " And ������ĿID=" & Val(Nvl(rsTemp!������ĿID))
        If Not rsItem.EOF And InStr(1, ",1,3,", "," & Val(Nvl(rsTemp!��¼״̬))) > 0 Then
            'Zl_�����շѼ�¼_Update
            strSQL = "Zl_�����շѼ�¼_Update("
            '  Id_In         In ������ü�¼.Id%Type,
            strSQL = strSQL & "" & rsTemp!ID & ","
            '  ���մ���id_In In ������ü�¼.���մ���id%Type,
            strSQL = strSQL & "" & ZVal(Val(Nvl(rsItem!���մ���ID))) & ","
            '  ������Ŀ��_In In ������ü�¼.������Ŀ��%Type,
            strSQL = strSQL & "" & ZVal(Val(Nvl(rsItem!������Ŀ��))) & ","
            '  ���ձ���_In   In ������ü�¼.���ձ���%Type,
            strSQL = strSQL & "'" & Nvl(rsItem!���ձ���) & "',"
            '  ��������_In   In ������ü�¼.��������%Type,
            strSQL = strSQL & "'" & Nvl(rsItem!��������) & "',"
            '  ͳ����_In   In ������ü�¼.ͳ����%Type,
            strSQL = strSQL & "" & Val(Nvl(rsItem!ͳ����)) & ","
            '  ժҪ_In       In ������ü�¼.ժҪ%Type
            strSQL = strSQL & "'" & Nvl(rsItem!ժҪ) & "')"
            zlAddArray cllPro, strSQL
        End If
        rsTemp.MoveNext
    Loop
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveItemYbMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub SetBalanceVal(strItem As String, curVal As Currency)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ�����ս��㷽ʽ����Чֵ
    '����:���˺�
    '����:2014-09-24 14:39:57
    '˵�����ú�����mcolBalanceΪ׼����,����ҽ�������շ�Ҳ��
    '˵������������ҽ���շ��޸ı��ս���������۵�ҽ���շ����ø����ʻ��Ƚ�����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, strTemp As String
    Dim cllNewBalance As Collection, i As Long
    
    Set cllNewBalance = New Collection
    If mcolBalance.Count <> 0 Then
        For i = 1 To mcolBalance.Count
            '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;��Ч���
            varTemp = Split(mcolBalance(i), ";")
            If varTemp(0) = strItem And varTemp(3) <> curVal Then
                strTemp = varTemp(0) & ";" & varTemp(1) & ";" & varTemp(2) & ";" & Format(curVal, "0.00")
            Else
                strTemp = varTemp(0) & ";" & varTemp(1) & ";" & varTemp(2) & ";" & varTemp(3)
            End If
            cllNewBalance.Add strTemp
        Next
    Else
        '������ʱǿ������:��֧��Ԥ�����ҽ�������շ�ʱ��
        strTemp = strItem & ";" & Format(curVal, "0.00") & ";0;" & Format(curVal, "0.00")
        cllNewBalance.Add strTemp
    End If
    Set mcolBalance = cllNewBalance
End Sub

Private Function CheckFactValied(Optional blnReCharge As Boolean = False, _
    Optional ByRef blnPrintBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ʊ�ݺ���ĺϷ��Լ��
    '���:blnReCharge-�Ƿ������շѵļ��
    '����:mblnPrintBill-�Ƿ��ӡƱ��
    '����:���ݺϷ�,����tru,���򷵻�false
    '����:���˺�
    '����:2014-09-24 17:30:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    'Ʊ�ݺ�����,�����Ѵ�ӡ���
    blnPrintBill = True
    '����Ƿ��ӡƱ��
    If mobjFactProperty.��ӡ��ʽ = 0 Then
        blnPrintBill = False
    Else
        If mobjFactProperty.��ӡ��ʽ = 2 Then
            If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                blnPrintBill = False
            End If
        End If
    End If
    
    '����ӡֱ���˳�
    If Not blnPrintBill Then CheckFactValied = True: Exit Function

    If Not mobjFactProperty.�ϸ���� Then
        If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
            MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus: Exit Function
        End If
        CheckFactValied = True: Exit Function
    End If
    
    If Trim(txtInvoice.Text) = "" Then
        MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
        If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus: Exit Function
    End If

InvoiceHandle:
    If zlGetInvoiceGroupUseID(mlng����ID, 1, txtInvoice.Text) = False Then Exit Function

    '�����������,Ʊ���Ƿ�����
    If CheckBillRepeat(mlng����ID, 1, txtInvoice.Text) Then
        If txtInvoice.Locked = False And txtInvoice.Tag <> Trim(txtInvoice.Text) Then
            MsgBox "Ʊ�ݺ�""" & txtInvoice.Text & """�Ѿ���ʹ�ã����������롣", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus: Exit Function
        End If
        
        Call RefreshFact
        If txtInvoice.Text = "" Then
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus: Exit Function
        End If
        MsgBox "��ǰƱ�ݺ��Ѿ���ʹ�ã������»�ȡƱ�ݺ�:" & txtInvoice.Text, vbInformation, gstrSysName
        GoTo InvoiceHandle:
    End If
    CheckFactValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case EM_Pan_Pati
        Item.Handle = picTop.hWnd
    Case EM_Pan_Diagnose
        Item.Handle = picDiagnose.hWnd
    Case EM_Pan_FeeList
        Item.Handle = picFeeList.hWnd
    Case EM_Pan_Down
        Item.Handle = picDown.hWnd
    End Select
End Sub

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2009-09-14 18:06:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim sngHight As Single
    Dim panLeft As Pane
    
    '������Ϣ�����ݲ���
    
    sngHight = picTop.Height \ Screen.TwipsPerPixelY
    Set panThis = dkpMan.CreatePane(EM_Pan_Pati, 200, sngHight, DockLeftOf, Nothing)
    panThis.MaxTrackSize.Height = sngHight
    panThis.MinTrackSize.Height = sngHight
    panThis.Title = "": panThis.Tag = EM_Pan_Diagnose
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picTop.hWnd
    
    
    If mEditType = EM_Balance_Charge Then
        '����б�
        sngHight = picDiagnose.Height \ Screen.TwipsPerPixelY
        Set panThis = dkpMan.CreatePane(EM_Pan_Diagnose, 200, sngHight, DockBottomOf, panThis)
        panThis.MaxTrackSize.Height = sngHight
        panThis.MinTrackSize.Height = sngHight
        panThis.Title = "���ѡ��": panThis.Tag = EM_Pan_Diagnose
        panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        panThis.Handle = picDiagnose.hWnd
    Else
        picDiagnose.Visible = False
    End If
    
    
    '������Ϣ�б�
    Set panThis = dkpMan.CreatePane(EM_Pan_FeeList, 250, 580, DockBottomOf, panThis)
    panThis.Title = "��ǰ�ѽ������Ϣ"
    panThis.Tag = EM_Pan_FeeList
    panThis.Handle = picFeeList.hWnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    
    '������
    sngHight = picDown.Height \ Screen.TwipsPerPixelY
    Set panThis = dkpMan.CreatePane(EM_Pan_Down, 200, 580, DockBottomOf, panLeft)
    panThis.Title = "": panThis.Tag = EM_Pan_Down
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picDown.hWnd
    panThis.MaxTrackSize.Height = sngHight
    panThis.MinTrackSize.Height = sngHight
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    zlRestoreDockPanceToReg Me, dkpMan, "����"
End Sub

Private Sub SetControlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�������
    '����:���˺�
    '����:2014-09-26 14:53:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    blnEdit = mEditType = EM_Balance_Charge Or mEditType = EM_Balance_Register
    txtժҪ.Enabled = blnEdit
    txtժҪ.BackColor = IIf(blnEdit, &H80000005, Me.BackColor)
    cboNO.Enabled = blnEdit
    cboNO.BackColor = IIf(blnEdit, &H80000005, Me.BackColor)
    
    PatiIdentify.Enabled = blnEdit
    PatiIdentify.AllowAutoCommCard = blnEdit
    PatiIdentify.AllowAutoICCard = blnEdit
    PatiIdentify.AllowAutoIDCard = blnEdit
    
    blnEdit = Not mEditType = EM_Balance_Err_Cancel
    txtInvoice.Enabled = blnEdit
    txtInvoice.BackColor = IIf(blnEdit, &H80000005, Me.BackColor)
End Sub

Private Sub reSizeWinControl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���µ�������ؼ�λ��
    '����:���˺�
    '����:2014-09-26 14:54:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dkpMan.RecalcLayout
    Call picTop_Resize
    Call picDiagnose_Resize
    Call picDown_Resize
    Call picFeeList_Resize
End Sub

Private Function ShowReclaimInvoice(ByVal strNos As String, ByRef strReclaimInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�ͷ�����Ҫ���յķ�Ʊ
    '���:strNos-��ǰ�ĵ��ݺ�,����ö��ŷ���(����Ǳ������,��Ϊ������㵥��)
    '����:strReclaimInvoice-���ػ��յķ�Ʊ��(����ö��ŷָ�),��ʽ:AAAA,BBB,....)
    '����:��ʾ���ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-10 17:53:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmReInvoiceTemp As frmReInvoice
    Dim strSQL As String, rsTemp As ADODB.Recordset, blnFee As Boolean '��ǰ�����Ƿ�Ϊ�շѽ���
    
    On Error GoTo errHandle
    'ȷ����ǰ�����Ƿ�Ϊ�շѽ���
    If mEditType = EM_Balance_Err_ReCharge Then
        strSQL = "Select 1 From ���ò����¼ Where Nvl(���ӱ�־, 0) = 0 And ������� = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȷ����ǰ�쳣�����Ƿ�Ϊ�շ��쳣����", Val(mstr�������))
        If Not rsTemp.EOF Then blnFee = True
    End If
    blnFee = blnFee Or mEditType = EM_Balance_Charge
    
    Set frmReInvoiceTemp = New frmReInvoice
    If frmReInvoiceTemp.ShowMe(Me, strNos, 0, 0, strReclaimInvoice, True, IIf(blnFee, 1, 4)) = False Then Exit Function
    If Not frmReInvoiceTemp Is Nothing Then Unload frmReInvoiceTemp
    Set frmReInvoiceTemp = Nothing
    ShowReclaimInvoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ClearDisplaySHow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˫����ʾ
    '����:���˺�
    '����:2014-10-13 15:07:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
    If Not gblnLED Then Exit Sub
    If mblnNotClearLedDisplay Then Exit Sub
    zl9LedVoice.DisplayPatient ""
End Sub

Private Function UpdateBalance(ByVal curʵ�պϼ� As Currency, ByVal cur����ͳ�� As Currency, _
        ByVal curȫ�Ը� As Currency, ByVal cur���Ը� As Currency) As Boolean
    '���µ�ǰ���ݸ����ʻ�֧�����:��֧��Ԥ����ʱ
    'ҽ��������������Ӧ�����Ŵ���,�ϼ�Ϊ�������˵������ʻ�
    Dim cur���� As Currency, cur���ø��� As Currency
    Dim i As Integer, j As Integer, blnFind As Boolean
    
    On Error GoTo Errhand
    If mstrYBPati <> "" And mstr�����ʻ� <> "" And mTy_Insure.dbl�ʻ���� > -1 * mTy_Insure.dbl����͸֧ Then
        If curʵ�պϼ� >= 0 Then
            cur���� = cur����ͳ�� + IIf(MCPAR.���Ը�, cur���Ը�, 0) + IIf(MCPAR.ȫ�Ը�, curȫ�Ը�, 0)
            cur���ø��� = mTy_Insure.dbl�ʻ����
            '��������ʻ�֧�����
            If cur���ø��� - cur���� >= -1 * mTy_Insure.dbl����͸֧ Then
                Call SetBalanceVal(mstr�����ʻ�, Format(cur����, "0.00"))   '������͸֧��Χ���㹻(����͸֧0Ϊ����)
            Else
                If mTy_Insure.dbl����͸֧ = 0 And cur���ø��� > 0 Then
                    Call SetBalanceVal(mstr�����ʻ�, Format(cur���ø���, "0.00"))  '������͸֧�������
                Else
                    '��������͸֧��Χ������͸֧ʱ�����
                    If mTy_Insure.dbl����͸֧ <> 0 Then
                        Call SetBalanceVal(mstr�����ʻ�, cur���ø��� + mTy_Insure.dbl����͸֧) '������͸֧��Χ��֧��
                    Else
                        Call SetBalanceVal(mstr�����ʻ�, 0)
                    End If
                End If
            End If
        Else
            Call SetBalanceVal(mstr�����ʻ�, 0)
        End If
        'ˢ����ʾ�����ʻ�֧�����
        '-------------------------------------------------------------------------
        With vsBalance
            For i = 1 To .COLS - 1 Step 2
                If .TextMatrix(0, i) = mstr�����ʻ� And mstr�����ʻ� <> "" Then
                    .TextMatrix(0, i + 1) = Format(GetMedicareBalanceSum(mstr�����ʻ�), "0.00")
                    blnFind = True: Exit For
                End If
            Next
            If blnFind = False Then
                j = -1
                For i = 1 To .COLS - 1 Step 2
                    If .TextMatrix(0, i) = "" Then j = i: Exit For
                Next
                If j < 0 Then .COLS = .COLS + 2: j = .COLS - 2
                .TextMatrix(0, i) = mstr�����ʻ�
                .TextMatrix(0, i + 1) = Format(GetMedicareBalanceSum(mstr�����ʻ�), "0.00")
            End If
        End With
    End If
    UpdateBalance = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
