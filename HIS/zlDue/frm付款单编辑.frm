VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm����༭ 
   Caption         =   "����"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   Icon            =   "frm����༭.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   10125
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2610
      TabIndex        =   24
      Top             =   660
      Width           =   2010
   End
   Begin MSComctlLib.ImageList ilt24 
      Left            =   870
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����༭.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   7080
      ScaleHeight     =   2340
      ScaleWidth      =   2430
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3945
      Width           =   2430
      Begin VSFlex8Ctl.VSFlexGrid vsFp 
         Height          =   1800
         Left            =   60
         TabIndex        =   18
         Top             =   390
         Width           =   2385
         _cx             =   4207
         _cy             =   3175
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm����༭.frx":0E64
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
         ExplorerBar     =   5
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
      Begin XtremeSuiteControls.ShortcutCaption stcFpTittle 
         Height          =   375
         Left            =   -30
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   2415
         _Version        =   589884
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "��ʱ��Ϣ-��Ʊ����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   5025
      ScaleHeight     =   2340
      ScaleWidth      =   1860
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1860
      Begin VSFlex8Ctl.VSFlexGrid vsTemp 
         Height          =   1920
         Left            =   150
         TabIndex        =   15
         Top             =   375
         Width           =   1845
         _cx             =   3254
         _cy             =   3387
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm����༭.frx":0EF9
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
         ExplorerBar     =   5
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
      Begin XtremeSuiteControls.ShortcutCaption stcTempTittle 
         Height          =   375
         Left            =   -30
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   1860
         _Version        =   589884
         _ExtentX        =   3281
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "��ʱ����-�ֶλ���"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   45
      ScaleHeight     =   2235
      ScaleWidth      =   4905
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3960
      Width           =   4905
      Begin VSFlex8Ctl.VSFlexGrid vsԤ�� 
         Height          =   1770
         Left            =   0
         TabIndex        =   6
         Top             =   390
         Width           =   4845
         _cx             =   8546
         _cy             =   3122
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm����༭.frx":0F48
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
         ExplorerBar     =   7
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
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ԥ��:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   4
         Left            =   3615
         TabIndex        =   12
         Top             =   105
         Width           =   630
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ���ۼ�:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   3
         Left            =   1305
         TabIndex        =   11
         Top             =   90
         Width           =   810
      End
      Begin XtremeSuiteControls.ShortcutCaption stcԤ�� 
         Height          =   375
         Left            =   -30
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   4890
         _Version        =   589884
         _ExtentX        =   8625
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "Ԥ���嵥"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picPayList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   0
      ScaleHeight     =   2460
      ScaleWidth      =   9600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1380
      Width           =   9600
      Begin VSFlex8Ctl.VSFlexGrid vsPayList 
         Height          =   1425
         Left            =   45
         TabIndex        =   3
         Top             =   810
         Width           =   8295
         _cx             =   14631
         _cy             =   2514
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm����༭.frx":0FE8
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
         ExplorerBar     =   7
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
      Begin VB.PictureBox picCon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   -30
         ScaleHeight     =   780
         ScaleWidth      =   9570
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   9570
         Begin VB.CommandButton cmdSelDept 
            Caption         =   "��"
            Height          =   255
            Left            =   4410
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   90
            Width           =   255
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            Left            =   675
            TabIndex        =   1
            Top             =   60
            Width           =   4005
         End
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����Ӧ��:"
            ForeColor       =   &H00000040&
            Height          =   180
            Index           =   5
            Left            =   8640
            TabIndex        =   23
            Top             =   510
            Width           =   810
         End
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ۼ�Ӧ��:"
            ForeColor       =   &H00000040&
            Height          =   180
            Index           =   1
            Left            =   3450
            TabIndex        =   22
            Top             =   510
            Width           =   810
         End
         Begin VB.Label lbl��� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������:"
            ForeColor       =   &H00000040&
            Height          =   180
            Index           =   2
            Left            =   6150
            TabIndex        =   21
            Top             =   510
            Width           =   810
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   4875
            TabIndex        =   13
            Top             =   120
            Width           =   90
         End
         Begin VB.Label lbl��Ӧ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ӧ��"
            Height          =   180
            Left            =   75
            TabIndex        =   0
            Top             =   135
            Width           =   540
         End
         Begin XtremeSuiteControls.ShortcutCaption stcPayTittle 
            Height          =   330
            Left            =   90
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   450
            Width           =   9525
            _Version        =   589884
            _ExtentX        =   16801
            _ExtentY        =   582
            _StockProps     =   6
            Caption         =   "δ�����嵥"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin XtremeSuiteControls.ShortcutCaption stcTop 
            Height          =   450
            Left            =   75
            TabIndex        =   20
            Top             =   0
            Width           =   9480
            _Version        =   589884
            _ExtentX        =   16722
            _ExtentY        =   794
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   6375
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm����༭.frx":130C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12779
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   435
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frm����༭.frx":1BA0
      Left            =   -15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm����༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmEdit As frmPayNoEdit
Attribute mfrmEdit.VB_VarHelpID = -1

Private mintStep As Integer
Private mstrFindKey As String       '��������
Private mstrNo As String                   '���ݺ�
Private mlng��λID As Long
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mblnSave As Boolean
Private mfrmMain  As Object
Private mEditType As gEditType
Private mint��¼״̬ As RecBillStatus       '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mErrBillStatusInfor As ErrBillStatusInfor       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mblnEdit As Boolean             '�༭״̬
Private mblnSuccess As Boolean          '�Ƿ��е��ݱ���ɹ�
Private mstrPrivs  As String
Private mlng������� As Long    '�������

Private mdbl�ۼ�Ӧ�� As Double
Private mdbl����Ӧ�� As Double
Private mdbl����Ԥ�� As Double
Private mdbl�ۼ�Ԥ�� As Double
Private mlngЧ�� As Long
Private mstrFind As String
Private mcllFilter As Collection
Private mcbrToolBar As CommandBar
Private mcbrMenuBar As CommandBarPopup
Private mcbrControl As CommandBarControl
Private mobjFindKey As CommandBarControl
Private mintͳ�Ʒ�ʽ As Integer     '0-�����ͬ�е��ŷ���ͳ��,1-����Ʊ��ͳ��
Private mint��� As Integer
Private mbln�����־ As Boolean
Private mint��ʾ��λ As Integer     '0����С��λ��  1�����λ

Private Const mConMenu_Hide = 99
Private Const mConMenu_Hide_TempSave = 9981
Private Const mConMenu_Hide_TempClearAll = 9982
Private Const mConMenu_Popu = 88
Private Const mConMenu_Popu_FP = 8801
Private Const mConMenu_Popu_SH = 8802
Private Const mConMenu_Report = 102

Private Const mlngModule = 1323
Private Enum mPanIndex
    pane_Ӧ���б� = 0
    pane_Ԥ���б� = 1
    pane_��� = 2
    pane_��ʱ���� = 3
    pane_��Ʊ�ϼ� = 4
End Enum

Private Function InitComandBars() As Boolean
    '----------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/1/9
    '----------------------------------------------------------------------------------------
    Dim cbrCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Visible = False
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mConMenu_Hide, "����(&D)", -1, False)
    mcbrMenuBar.Visible = False
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Hide_TempSave, "������ʱ��Ϣ(&S)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Save, IIf(mEditType = gԤ��, "Ԥ��(&O", "����(&O)"))
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Audit, "���(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeOff, "����(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����(&H)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_FilterView, "����(&F)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With
 
    mintͳ�Ʒ�ʽ = Val(zlDatabase.GetPara("ͳ�Ʒ�ʽ", glngSys, mlngModule, "1"))
        
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mConMenu_Popu, "ͳ�Ʒ�ʽ(&T)", -1, False)
    mcbrMenuBar.Visible = False
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Popu_FP, "����Ʊ�Ž��з���ͳ��(&1)")
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Popu_SH, "��������Ž��з���ͳ��(&2)")
    End With
        
    '�����
    With Me.cbsThis.KeyBindings
        .Add FALT, Asc("O"), conMenu_Edit_Save
        .Add FALT, Asc("V"), conMenu_Manage_Audit
        .Add FALT, Asc("S"), conMenu_Edit_ChargeOff
        .Add FALT, Asc("R"), conMenu_View_FilterView
        .Add FCONTROL, Asc("S"), mConMenu_Hide_TempSave
        .Add FCONTROL, Asc("C"), mConMenu_Hide_TempClearAll
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��")
        
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Hide_TempSave, "������ʱ��Ϣ"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3200
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Hide_TempClearAll, "�����ʱ��Ϣ"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3702
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Forward, "��һ��"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Backward, "��һ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Save, IIf(mEditType = gԤ��, "Ԥ��", "����")): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Audit, "���"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeOff, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_FilterView, "����"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 254
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Report, "ҩƷ�����ѯ"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
     
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        mcbrControl.Flags = xtpFlagRightAlign
        mstrFindKey = Trim(zlDatabase.GetPara("��λ����", glngSys, mlngModule, "�������"))
        
        Set mobjFindKey = mcbrToolBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
       ' mobjFindKey.BeginGroup = True
        mobjFindKey.IconId = conMenu_View_Find
        mobjFindKey.Flags = xtpFlagRightAlign
        mobjFindKey.Style = xtpButtonIconAndCaption
        If mstrFindKey = "" Then mstrFindKey = "�������"
        
        With mobjFindKey.CommandBar.Controls
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&1.�������")
            mcbrControl.Parameter = "�������"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&2.��Ʊ��")
            mcbrControl.Parameter = "��Ʊ��"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&3.��ⵥ��")
            mcbrControl.Parameter = "��ⵥ��"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&4.Ʒ��")
            mcbrControl.Parameter = "Ʒ��"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&5.����")
            mcbrControl.Parameter = "����"
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&6.��Ʊ���")
            mcbrControl.Parameter = "��Ʊ���"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&7.�����")
            mcbrControl.Parameter = "�����"
        End With
    
        Set cbrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
        cbrCustom.Handle = txtFind.hwnd
        cbrCustom.Flags = xtpFlagRightAlign
        
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    InitComandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub SaveTempData(Optional blnClsAll As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '����:�ɴ���ʱ���ݵ�����
    '���:blnClsAll-�Ƿ������ʷ�洢����Ϣ
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-20 11:36:23
    '-----------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double, i As Long, lngMax��� As Long, blnHaveData As Boolean
    With vsTemp
        If blnClsAll Then
            .Clear 1
            .Rows = 2
        Else
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("���"))) > lngMax��� Then lngMax��� = Val(.TextMatrix(i, .ColIndex("���")))
            Next
            lngMax��� = lngMax��� + 1
        End If
    End With
    
    blnHaveData = False
    With vsPayList
        dbl��� = 0
        If blnClsAll Then
            .Cell(flexcpText, 1, .ColIndex("�������"), .Rows - 1, .ColIndex("�������")) = ""
            Exit Sub
        Else
            For i = 1 To .Rows - 1
                .Redraw = flexRDNone
                If Trim(.TextMatrix(i, .ColIndex("ѡ��"))) = "��" And Val(.TextMatrix(i, .ColIndex("�������"))) = 0 Then
                    dbl��� = dbl��� + Val(.Cell(flexcpData, i, .ColIndex("��Ʊ���")))
                    .TextMatrix(i, .ColIndex("�������")) = lngMax���
                    blnHaveData = True
                End If
                .Redraw = flexRDBuffered
            Next
        End If
    End With
    If blnHaveData Then
        '����ʱ������
        With vsTemp
            If Val(.TextMatrix(.Rows - 1, .ColIndex("���"))) <> 0 Then
                .Rows = .Rows + 1
            End If
            .TextMatrix(.Rows - 1, .ColIndex("���")) = lngMax���
            .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(dbl���, gVbFmtString.FM_���)
            .Cell(flexcpData, .Rows - 1, .ColIndex("���")) = dbl���
        End With
    End If
End Sub

Private Sub InitPancel()
    '����:��ʼ������ؼ�:2008-07-14 15:04:29
    Dim panThis As Pane
        
    Set mfrmEdit = New frmPayNoEdit
    Load mfrmEdit
    '����27930 by lesfeng 2010-03-23
    mfrmEdit.zlInitPara Me, mlngModule, mstrPrivs, mint���
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_���, 250, 580, DockTopOf, Nothing)
    panThis.Title = "����֪ͨ��"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    panThis.Close
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_Ӧ���б�, 250, 580, DockTopOf, Nothing)
    panThis.Title = "�����嵥"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_Ԥ���б�, 250, 580, DockBottomOf, panThis)
    panThis.Title = "Ԥ�����嵥"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_��ʱ����, 124, 580, DockRightOf, panThis)
    panThis.Title = "Ӧ����ʱ����"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    panThis.MinTrackSize.Width = 124
    panThis.MaxTrackSize.Width = 348
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_��Ʊ�ϼ�, 124, 580, DockRightOf, panThis)
    panThis.Title = "Ӧ����Ʊ�ϼ���Ϣ"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    panThis.MinTrackSize.Width = 124
    panThis.MaxTrackSize.Width = 348
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMan.Options.AlphaDockingContext = True
    Me.dkpMan.Options.HideClient = True
End Sub

Private Sub InitFilter()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-18 16:10:40
    '-----------------------------------------------------------------------------------------------------------
    Set mcllFilter = New Collection
    mcllFilter.Add mlng��λID, "��Ӧ��ID"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "�������"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "��Ʊ����"
    mcllFilter.Add "", "��Ʊ���б�"
    mcllFilter.Add "", "��������б�"
    mcllFilter.Add "", "ϵͳ��ʶ"
    mcllFilter.Add "", "Ʒ��"
    mcllFilter.Add "", "���"
    mcllFilter.Add "", "����"
    mcllFilter.Add Array("", ""), "����"
    mcllFilter.Add Array("", ""), "��ⵥ��"
    mcllFilter.Add Array("", ""), "��Ʊ��"
    mcllFilter.Add Array("", ""), "�������"
    mcllFilter.Add "", "������"
    mcllFilter.Add "", "�����"
    mcllFilter.Add "", "����"
    mcllFilter.Add "", "�ⷿ"
    mcllFilter.Add "0", "������ҩƷ�������С�ڷ�Ʊ����"
    mstrFind = ""
End Sub
'�������������
Private Function GetDepend() As Boolean
    Dim strSQL As String
    Dim rsTemp As New Recordset
    
    GetDepend = False
    
    '��ȡ���㷽ʽ
    Err = 0: On Error GoTo ErrHand:
    
    strSQL = "Select 1 From ���㷽ʽӦ�� Where Ӧ�ó���='������' and rownum<=1 Order by ȱʡ��־ desc"
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "���㷽ʽӦ����Ϣ��ȫ,���ڽ��㷽ʽ�����н������ã�"
        Exit Function
    End If
    GetDepend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub initGrid()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ���Ĭ������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-20 11:11:42
    '-----------------------------------------------------------------------------------------------------------
    With vsPayList
        .Cell(flexcpPicture, 0, .ColIndex("�������"), 0, .ColIndex("�������")) = ilt24.ListImages(1).Picture
        .Cell(flexcpPictureAlignment, 0, .ColIndex("�������"), 0, .ColIndex("�������")) = 4
        .Cell(flexcpAlignment, 0, .ColIndex("�������"), 0, .ColIndex("�������")) = 4
        If (mEditType = g���� Or mEditType = g�޸�) And InStr(1, mstrPrivs, ";�޸ķ�Ʊ��Ϣ;") > 0 Then
            '�д�Ȩ��ʱ���������޸ķ�Ʊ��Ϣ
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub

Private Sub initCard()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ƭ��Ϣ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-19 13:39:41
    '-----------------------------------------------------------------------------------------------------------
    Dim intErrInfor As Integer
    Call initGrid
    '��ʼ���
    If mfrmEdit.zlLoadData(mEditType, mlng��λID, mstrNo, mint��¼״̬, intErrInfor) = False Then
        If intErrInfor = 1 Then
            mErrBillStatusInfor = �Ѿ�ɾ��
        ElseIf intErrInfor = 2 Then
            mErrBillStatusInfor = �Ѿ����
        End If
        Exit Sub
    End If
    If mEditType = g���� Then Exit Sub
    Call LoadPayMoney
    Call ���ܷ�Ʊ��Ϣ
    SetEditPro
End Sub

Public Sub ShowCard(FrmMain As Form, _
    ByVal int�༭״̬ As gEditType, ByVal strPrivs As String, _
    Optional strNO As String = "", _
    Optional lng��λID As Long = 0, _
    Optional int��¼״̬ As RecBillStatus = 1, _
    Optional blnSuccess As Boolean = False, _
    Optional int��� As Integer = 0)
    
    mstrNo = strNO
    mblnSave = False
    mblnSuccess = False
    mEditType = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mstrPrivs = strPrivs

    mlng��λID = lng��λID
    mint��� = int���
    
    mblnChange = False
    mErrBillStatusInfor = �������
    Set mfrmMain = FrmMain
    '����27930 by lesfeng 2010-03-23
    If int��� = 1 Then
        Me.Caption = "��Ǹ���"
    End If
    '��ʼ����������:2008-08-18 17:48:29
    Call InitFilter
    
    '�������������ϵ
    If Not GetDepend Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
     
    If mEditType = g���� Then
        mblnEdit = True
    ElseIf mEditType = g�޸� Then
        mblnEdit = True
    ElseIf mEditType = gԤ�� Then
        mblnEdit = True
    ElseIf mEditType = g��� Then
        mblnEdit = False
    ElseIf mEditType = gȡ�� Then
        mblnEdit = False
    ElseIf mEditType = g�鿴 Then
        mblnEdit = False
    End If
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
End Sub

Private Sub LoadPayMoney()
    '--------------------------------------------------------------
    '���ܣ���乩ѡ���Ӧ��������
    '������
    '���أ�
    '˵����
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long
    Dim strSQL As String, strWhere As String, strMultiPay As String, strHead As String, strAmount As String
    Dim blnMaterialSys As Boolean
    
    '��־,��Ʊ��,��ⵥ��,Ʒ��,���,��λ,����,��Ʊ���
    Call zlCommFun.ShowFlash("�������������¼,���Ժ� ...", Me)
    
    vsPayList.Redraw = False
    DoEvents
    Screen.MousePointer = vbHourglass
    
    '��鰲װϵͳ
    strSQL = "select ��� from zlSystems where ���=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鰲װϵͳ", 400)
    If Not rsTemp.EOF Then
        blnMaterialSys = True
    End If
    rsTemp.Close
    
    '���ݲ��������趨��¼��ȡ����
    If mEditType = g���� Then
        '����ʱ��ȡ�������Ϊ�յ�Ӧ���ѡ��
        If Format(CDate(mcllFilter("�������")(0)), "yyyy-mm-dd") <> "1901-01-01" Then
            strWhere = " and a.������� Is Null and a.������� between [2] and [3] "
        Else
            strWhere = " and a.������� Is Null "
        End If
    ElseIf mEditType = g�޸� Then
        '�༭ʱ��ȡ�������Ϊ�ջ�ǰ�༭�ĸ����������Ӧ��Ӧ����
        If Format(CDate(mcllFilter("�������")(0)), "yyyy-mm-dd") <> "1901-01-01" Then
            strWhere = " And (a.������� Is Null Or a.�������=[22]) And a.������� between [2] and [3] "
        Else
            strWhere = " And (a.������� Is Null Or a.�������=[22]) "
        End If
    Else
        '�鿴�����ʱ����ȡ��ǰ�༭�ĸ������Ӧ��Ӧ����
        strWhere = " And a.�������=[22]"
    End If
    
    If mEditType = g���� Then
        strMultiPay = "(Select a.ID, 0 δ�����, " & _
                      "   Sum(decode(b.�����, null, a.�ƻ����, 0)) ���θ���, " & _
                      "   Sum(decode(b.�����, null, 0, a.�ƻ����)) �Ѹ����  " & _
                      " From Ӧ����¼ A, �����¼ B " & _
                      " Where a.������� = b.������� And instr('12345',a.ϵͳ��ʶ)>0 And a.��¼���� = 2 " & _
                      "   and a.��¼״̬ = 1 and a.������� is not null " & _
                      " Group by a.ID ) A1, "
    Else
        strMultiPay = "select min(a.id) id,a.no,a.���,a.��Ʊ��, " & _
                      "    sum(case when a.��¼����=2 then 0 else a.��Ʊ��� end) ��Ʊ���, a.��¼����,a.�������,a.�ƻ����," & _
                      "    sum(a.����) ���� " & _
                      "from (select distinct a.* from Ӧ����¼ a, Ӧ����¼ b " & _
                      "      where b.�������=[22] and a.no=b.no and a.���=b.��� and a.��Ʊ��=b.��Ʊ�� and a.ϵͳ��ʶ=b.ϵͳ��ʶ and a.��Ŀid=b.��Ŀid " & _
                      "     ) a " & _
                      "group by a.no,a.���,a.��Ʊ��,a.ϵͳ��ʶ,a.��Ŀid,a.��¼����,a.�������,a.�ƻ���� "
        
        strMultiPay = "(Select ID, ��Ʊ��� - nvl(�Ѹ����,0) δ�����, �Ѹ����, ��Ʊ���, ����, " & _
                      "   Case When Nvl(�Ѹ����, 0) = 0 and ���θ��� = 0 then ��Ʊ��� else ���θ��� end ���θ��� " & _
                      " From (" & _
                      "   Select a.ID, Max(a.��Ʊ���) ��Ʊ���, sum(����) ����, " & _
                      "     Sum(Case When a.�ƻ���� Is Not Null And a.��¼���� = 2 And Nvl(a.�ƻ����, 0) <> a.��Ʊ��� And b.Id Is Null Then a.�ƻ���� Else 0 End) �Ѹ����," & _
                      "     Sum(Case When a.�ƻ���� Is Not Null And Nvl(a.�ƻ����, 0) <> a.��Ʊ��� And b.Id Is Not Null Then a.�ƻ���� " & _
                      "              When a.�ƻ���� Is Null And a.������� Is Null    and b.ID is not null   Then a.��Ʊ��� Else 0 End) ���θ��� " & _
                      "   From (" & strMultiPay & ") A, �����¼ B " & _
                      "   Where a.������� = b.�������(+) And b.�������(+) = [22] " & _
                      "     And a.Id in (Select ID From Ӧ����¼ Where ������� = [22])  And a.��¼���� <> -1 " & _
                      "   Group By a.ID ) ) A1,"
                      
    End If
    strSQL = "With T1 as (Select " & IIf(mEditType = gԤ��, "Max(Decode(a.Ԥ��, 1, '��', ''))", "Decode(a.�������, Null, '', '��')") & " As ��־, " & _
             "        min(a.ID) ID,max(a.��¼״̬) ��¼״̬, " & _
             "        '' �ƻ�����,a.�������,a.��Ʊ��,to_char(max(A.��Ʊ����),'yyyy-mm-dd') as ��Ʊ����,a.��ⵥ�ݺ�," & _
             "        max(a.�����) as �����, to_char(max(a.�������),'yyyy-mm-dd') as �������, " & _
             "        max(a.Ʒ��) as Ʒ��,max(a.���) ���,max(a.������λ) ������λ, " & _
             "        max(b.ҩ�ۼ���) as ҩ�ۼ���," & _
             "        sum(nvl(a.����,0)) as ����, " & _
             "        sum(nvl(a.��Ʊ���,0)) as ��Ʊ���, max(a.ϵͳ��ʶ) ϵͳ��ʶ, max(a.�ⷿid) �ⷿid, a.��Ŀid, a.NO " & _
             "  From (" & _
             "        Select Distinct c.* " & vbCr & _
             "        From Ӧ����¼ A, Ӧ����¼ C " & vbCr & _
             "        Where a.ϵͳ��ʶ=c.ϵͳ��ʶ And a.No = c.No And a.��¼���� = c.��¼���� And a.��Ŀid=c.��Ŀid " & _
             "            And a.���=c.��� And a.�ƻ���� = c.�ƻ���� " & vbCr & _
             "            And a.�ƻ����� Is Null and a.��λID = [1] " & vbCr & _
             IIf(mEditType = g���� Or mEditType = g�޸�, " And not A.��¼���� in (-1, 2) ", " And a.��¼���� <> -1 ") & _
             IIf(mbln�����־, " And (A.�����־=1 and nvl(a.ϵͳ��ʶ,0)=1 or nvl(a.ϵͳ��ʶ,0)<>1 ) ", "") & _
                    strWhere & Replace(CStr(mcllFilter("����")), "[alias]", "a.") & _
             ") A, ҩƷ��� B" & _
             "  Where decode(A.ϵͳ��ʶ,1,a.��Ŀid,0)=b.ҩƷID(+) " & _
             "  group by a.��¼����,a.NO,a.��Ŀid,a.���," & IIf(mEditType = gԤ��, "", "a.�������,") & "a.��Ʊ��,a.��Ʊ����,a.��ⵥ�ݺ�,a.������� " & _
             "  having sum(nvl(a.��Ʊ���,0))<>0 " & _
             ") "
    
    If mEditType = g�޸� Then
        strHead = "Select Decode(nvl(a1.���θ���,0), 0, a.��־, '��') ��־,"
    Else
        strHead = "Select a.��־, "
    End If
    
    strSQL = strSQL & _
             strHead & _
             "  a.Id, a.��¼״̬, a.�ƻ�����, a.�������, a.��Ʊ��, a.��Ʊ����, a.��ⵥ�ݺ�, a.�����, a.�������, " & _
             "  a.Ʒ��, a.���, a.ҩ�ۼ���, a.�ⷿid, a.��Ŀid, a.������λ, " & _
             IIf(mEditType = g����, "a.����,", "decode(nvl(a.����,0), 0, a1.����, a.����) ����, ") & _
             IIf(mEditType = g����, "a.��Ʊ���, ", "decode(a.��Ʊ���, null, a1.��Ʊ���, a.��Ʊ���) ��Ʊ���, ") & _
             "  Decode(a.ϵͳ��ʶ, 1, a.���� / b.ҩ���װ, 5, a.���� / e1.����ϵ��, " & IIf(blnMaterialSys, "2, a.���� / e2.����ϵ��,", "") & " null, a.����, null) ҩ������, " & _
             "  Decode(a.ϵͳ��ʶ, 1, b.ȫԺ���, 0) ȫԺ���, " & _
             "  Decode(a.ϵͳ��ʶ, 1, c.��ǰ�ⷿ���, 0) ��ǰ�ⷿ���, " & _
             "  Decode(a.ϵͳ��ʶ, 1, b.ҩ�ⵥλ, 5, E1.��װ��λ, " & IIf(blnMaterialSys, "2, E2.��װ��λ,", "") & " null, a.������λ, '') ҩ�ⵥλ, " & _
             "  d.���� ��ǰ�ⷿ, a.ϵͳ��ʶ, a1.�Ѹ����, a1.���θ���, a1.δ����� " & _
             "From T1 A, " & strMultiPay & _
             "  (Select a.ҩƷid, " & IIf(mint��ʾ��λ = 1, "Round(a.ȫԺ��� / b.ҩ���װ, 5)", "a.ȫԺ���") & " ȫԺ���, b.ҩ�ⵥλ, b.ҩ���װ " & _
             "   From (Select a.ҩƷid, Sum(a.ʵ������) ȫԺ��� From ҩƷ��� a Where Exists(Select 1 From T1 Where ��Ŀid = a.ҩƷid) " & _
             "         Group By a.ҩƷid) A, ҩƷ��� B Where a.ҩƷid = b.ҩƷid) B," & _
             "  (Select a.�ⷿid, a.ҩƷid, " & IIf(mint��ʾ��λ = 1, "Round(a.��ǰ�ⷿ��� / b.ҩ���װ, 5)", "a.��ǰ�ⷿ���") & " ��ǰ�ⷿ��� " & _
             "   From (Select a.�ⷿid, a.ҩƷid, Sum(a.ʵ������) ��ǰ�ⷿ��� From ҩƷ��� A Where Exists(Select 1 From T1 Where ��Ŀid = a.ҩƷid) " & _
             "         Group By a.�ⷿid, a.ҩƷid) A, ҩƷ��� B Where a.ҩƷid = b.ҩƷid) C, " & _
             "  ���ű� D, �������� E1 " & IIf(blnMaterialSys, ", ����Ŀ¼ E2 ", "")
    
    '������ҩƷ�������С�ڷ�Ʊ����
    If Val(mcllFilter("������ҩƷ�������С�ڷ�Ʊ����")) = 1 Then
        strSQL = strSQL & _
                 ",(Select a.��Ŀid, a.����, b.ȫԺ��� " & _
                 "  From (Select ��Ŀid, Sum(����) ���� " & _
                 "        From Ӧ����¼ Where Nvl(ϵͳ��ʶ, 0) = 1 And ��λid = [1] And �����־ = 1 And ������� Is Null " & _
                 "        Group By ��λid, ��Ŀid) A, (Select ҩƷid, Sum(ʵ������) ȫԺ��� From ҩƷ��� Group By ҩƷid) B " & _
                 "  Where a.��Ŀid = b.ҩƷid(+) And a.���� > nvl(b.ȫԺ���,0)) F "
    End If

    strSQL = strSQL & _
             "Where a.ID = a1.ID(+) " & IIf(True, "", " and a.������� = a1.�������(+) ") & _
             "  and a.��Ŀid=b.ҩƷid(+) and a.�ⷿid=c.�ⷿid(+) and a.��Ŀid=c.ҩƷid(+) and a.�ⷿid=d.id(+) And a.��Ŀid = E1.����id(+) " & _
             IIf(blnMaterialSys, " And a.��Ŀid = E2.Id(+) ", "") & _
             IIf(Val(mcllFilter("������ҩƷ�������С�ڷ�Ʊ����")) = 1, " And a.��Ŀid = f.��Ŀid ", "") & _
             IIf(mEditType = g����, "  and nvl(a.��Ʊ���,0) - nvl(a1.�Ѹ����,0) - nvl(a1.���θ���,0) <> 0 ", "") & _
             "Order by a.��Ʊ�� "
             
    strSQL = "Select * From (" & strSQL & ")"
             
'    If Val(mcllFilter("������ҩƷ�������С�ڷ�Ʊ����")) = 1 Then
'        strSQL = "select * from (" & strSQL & ") where ȫԺ��� < ҩ������ "
'    End If
    
    '��Ӧ��ID: [1]
    '�������: [2] [3]
    '��Ʊ����: [4] [5]
    '��Ʊ���б�: [6]
    '��������б�: [7]
    'ϵͳ��ʶ: [8]
    'Ʒ��: [9]
    '���: [10]
    '����: [11]
    '����: [12] , [13]
    '��ⵥ�ݺ�: [14] [15]
    '��Ʊ��: [16] [17]
    '�������: [18] [19]
    '������: [20]
    '�����: [21]
    '�������:[22]
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID, _
     CDate(mcllFilter("�������")(0)), CDate(mcllFilter("�������")(1)), _
     CDate(mcllFilter("��Ʊ����")(0)), CDate(mcllFilter("��Ʊ����")(1)), _
     CStr(mcllFilter("��Ʊ���б�")), CStr(mcllFilter("��������б�")), _
     CStr(mcllFilter("ϵͳ��ʶ")), CStr(mcllFilter("Ʒ��")), _
     CStr(mcllFilter("���")), CStr(mcllFilter("����")), _
     CStr(mcllFilter("����")(0)), CStr(mcllFilter("����")(1)), _
     CStr(mcllFilter("��ⵥ��")(0)), CStr(mcllFilter("��ⵥ��")(1)), _
     CStr(mcllFilter("��Ʊ��")(0)), CStr(mcllFilter("��Ʊ��")(1)), _
     CStr(mcllFilter("�������")(0)), CStr(mcllFilter("�������")(1)), _
     CStr(mcllFilter("������")), CStr(mcllFilter("�����")), _
     mlng�������, Val(mcllFilter("�ⷿ")))
    
    '��ʼ�����������
    With vsPayList
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        .Clear 1
        mdbl����Ӧ�� = 0
        mdbl�ۼ�Ӧ�� = 0
        i = 1
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("ѡ��")) = NVL(rsTemp!��־)
            'ID,��¼״̬,�ƻ�����
            .Cell(flexcpData, i, .ColIndex("ѡ��")) = NVL(rsTemp!ID) & "," & NVL(rsTemp!��¼״̬) & "," & NVL(rsTemp!�ƻ�����)
            '.TextMatrix(i, .ColIndex("�����־")) = Nvl(rsTemp!�����־)
            '.Cell(flexcpData, i, .ColIndex("�����־")) = Nvl(rsTemp!ϵͳ��ʶ)
            .TextMatrix(i, .ColIndex("�������")) = NVL(rsTemp!�������)
            .TextMatrix(i, .ColIndex("��ⵥ��")) = NVL(rsTemp!��ⵥ�ݺ�)
            .TextMatrix(i, .ColIndex("��Ʊ��")) = NVL(rsTemp!��Ʊ��)
            .Cell(flexcpData, i, .ColIndex("��Ʊ��")) = NVL(rsTemp!��Ʊ��)
            .TextMatrix(i, .ColIndex("��Ʊ����")) = NVL(rsTemp!��Ʊ����)
            .Cell(flexcpData, i, .ColIndex("��Ʊ����")) = NVL(rsTemp!��Ʊ����)
            .TextMatrix(i, .ColIndex("ҩ�ۼ���")) = NVL(rsTemp!ҩ�ۼ���)
            .TextMatrix(i, .ColIndex("ϵͳ��ʶ")) = NVL(rsTemp!ϵͳ��ʶ)
            .TextMatrix(i, .ColIndex("ҩƷID")) = NVL(rsTemp!��ĿID)
            .TextMatrix(i, .ColIndex("�ⷿID")) = NVL(rsTemp!�ⷿID)
            .TextMatrix(i, .ColIndex("��ǰ�ⷿ")) = NVL(rsTemp!��ǰ�ⷿ)
            .TextMatrix(i, .ColIndex("��ǰ�ⷿ���")) = Format(Val(NVL(rsTemp!��ǰ�ⷿ���)), gVbFmtString.FM_����)
            .TextMatrix(i, .ColIndex("ȫԺ���")) = Format(Val(NVL(rsTemp!ȫԺ���)), gVbFmtString.FM_����)
'            .TextMatrix(i, .ColIndex("ҩ�ⵥλ")) = Nvl(rsTemp!ҩ�ⵥλ)
'            .TextMatrix(i, .ColIndex("ҩ������")) = Format(Val(Nvl(rsTemp!ҩ������)), gVbFmtString.FM_����)
            .TextMatrix(i, .ColIndex("�����")) = NVL(rsTemp!�����)
            .TextMatrix(i, .ColIndex("�������")) = NVL(rsTemp!�������)
            .TextMatrix(i, .ColIndex("Ʒ��")) = NVL(rsTemp!Ʒ��)
            .TextMatrix(i, .ColIndex("���")) = NVL(rsTemp!���)
            If mint��ʾ��λ = 1 Then
                .TextMatrix(i, .ColIndex("��λ")) = NVL(rsTemp!ҩ�ⵥλ)
                .TextMatrix(i, .ColIndex("����")) = Format(Val(NVL(rsTemp!ҩ������)), gVbFmtString.FM_����)
            Else
                .TextMatrix(i, .ColIndex("��λ")) = NVL(rsTemp!������λ)
                .TextMatrix(i, .ColIndex("����")) = Format(Val(NVL(rsTemp!����)), gVbFmtString.FM_����)
            End If
            .TextMatrix(i, .ColIndex("��Ʊ���")) = Format(Val(NVL(rsTemp!��Ʊ���)), gVbFmtString.FM_���)
            .Cell(flexcpData, i, .ColIndex("��Ʊ���")) = Val(NVL(rsTemp!��Ʊ���))
            
            .TextMatrix(i, .ColIndex("�Ѹ����")) = Format(Val(NVL(rsTemp!�Ѹ����)), gVbFmtString.FM_���)
            .Cell(flexcpData, i, .ColIndex("�Ѹ����")) = Val(NVL(rsTemp!�Ѹ����))
            '.TextMatrix(i, .ColIndex("δ�����")) = Format(.Cell(flexcpData, i, .ColIndex("��Ʊ���")) - .Cell(flexcpData, i, .ColIndex("�Ѹ����")), gVbFmtString.FM_���)
            '.Cell(flexcpData, i, .ColIndex("δ�����")) = .Cell(flexcpData, i, .ColIndex("��Ʊ���")) - .Cell(flexcpData, i, .ColIndex("�Ѹ����"))
            If mEditType = g���� Then
                .TextMatrix(i, .ColIndex("δ�����")) = Format(NVL(rsTemp!��Ʊ���) - Val(NVL(rsTemp!�Ѹ����)), gVbFmtString.FM_���)
                .Cell(flexcpData, i, .ColIndex("δ�����")) = Val(NVL(rsTemp!��Ʊ���) - Val(NVL(rsTemp!�Ѹ����)))
                If IsNull(rsTemp!���θ���) Then
                    .TextMatrix(i, .ColIndex("���θ���")) = Format(.Cell(flexcpData, i, .ColIndex("δ�����")), gVbFmtString.FM_���)
                Else
                    .TextMatrix(i, .ColIndex("���θ���")) = Format(.Cell(flexcpData, i, .ColIndex("δ�����")) - NVL(rsTemp!���θ���), gVbFmtString.FM_���)
                End If
                '���Ʊ��θ�����
                .Cell(flexcpData, i, .ColIndex("���θ���")) = .TextMatrix(i, .ColIndex("���θ���"))
            ElseIf mEditType = g�޸� Then
                .TextMatrix(i, .ColIndex("δ�����")) = Format(NVL(rsTemp!��Ʊ���, 0) - Val(NVL(rsTemp!�Ѹ����)), gVbFmtString.FM_���)
                .Cell(flexcpData, i, .ColIndex("δ�����")) = Val(NVL(rsTemp!��Ʊ���, 0) - Val(NVL(rsTemp!�Ѹ����)))
                If IsNull(rsTemp!���θ���) Then
                    .TextMatrix(i, .ColIndex("���θ���")) = Format(.Cell(flexcpData, i, .ColIndex("δ�����")), gVbFmtString.FM_���)
                Else
                    .TextMatrix(i, .ColIndex("���θ���")) = Format(NVL(rsTemp!���θ���), gVbFmtString.FM_���)
                End If
                '���Ʊ��θ�����
                If Val(.TextMatrix(i, .ColIndex("���θ���"))) < .Cell(flexcpData, i, .ColIndex("δ�����")) Then
                    .Cell(flexcpData, i, .ColIndex("���θ���")) = .Cell(flexcpData, i, .ColIndex("δ�����"))
                Else
                    .Cell(flexcpData, i, .ColIndex("���θ���")) = .TextMatrix(i, .ColIndex("���θ���"))
                End If
            Else
                .TextMatrix(i, .ColIndex("δ�����")) = Format(Val(NVL(rsTemp!δ�����)), gVbFmtString.FM_���)
                .Cell(flexcpData, i, .ColIndex("δ�����")) = Val(NVL(rsTemp!δ�����))
                If IsNull(rsTemp!���θ���) And .Cell(flexcpData, i, .ColIndex("��Ʊ���")) = .Cell(flexcpData, i, .ColIndex("δ�����")) Then
                    .TextMatrix(i, .ColIndex("���θ���")) = Format(.Cell(flexcpData, i, .ColIndex("δ�����")), gVbFmtString.FM_���)
                Else
                    .TextMatrix(i, .ColIndex("���θ���")) = Format(NVL(rsTemp!���θ���), gVbFmtString.FM_���)
                End If
                '���Ʊ��θ�����
                .Cell(flexcpData, i, .ColIndex("���θ���")) = .Cell(flexcpData, i, .ColIndex("δ�����"))
            End If
            
            mdbl�ۼ�Ӧ�� = mdbl�ۼ�Ӧ�� + .Cell(flexcpData, i, .ColIndex("δ�����"))
            If Trim(.TextMatrix(i, .ColIndex("ѡ��"))) <> "" Then
                'mdbl����Ӧ�� = mdbl����Ӧ�� + .Cell(flexcpData, i, .ColIndex("���θ���"))
                mdbl����Ӧ�� = mdbl����Ӧ�� + Val(.TextMatrix(i, .ColIndex("���θ���")))
            End If
            
            i = i + 1
            rsTemp.MoveNext
        Loop
        .Row = 1: .Col = 1
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        Call zl_vsGrid_Para_Restore(mlngModule, vsPayList, Me.Caption, "һ�㸶���б�")
        
'        .ColHidden(.ColIndex("��ǰ�ⷿ")) = Not mbln�����־
'        .ColHidden(.ColIndex("��ǰ�ⷿ���")) = Not mbln�����־
'        .ColHidden(.ColIndex("ȫԺ���")) = Not mbln�����־
'        .ColHidden(.ColIndex("ҩ�ⵥλ")) = True
'        .ColHidden(.ColIndex("ҩ������")) = True
        
        .ColHidden(.ColIndex("�Ѹ����")) = Not mbln�����־
        .ColHidden(.ColIndex("δ�����")) = Not mbln�����־
        .ColHidden(.ColIndex("���θ���")) = Not mbln�����־
        
        .Redraw = flexRDBuffered
    End With
    Call zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    
    Call SetMoneyLbl
    Call GetԤ������             '��ȡԤ����
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlCommFun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub
 
Private Sub GetԤ������()
    '--------------------------------------------------------------
    '���ܣ���ȡ�����Ԥ�����¼��ѡ��
    '������
    '���أ�
    '˵����
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long
    
    Dim strSQL As String
    Dim strWhere As String
    Dim lngLoop As Long
    
    '��־,���㷽ʽ,������,�������
    Call zlCommFun.ShowFlash("��������Ԥ�����¼,���Ժ� ...", Me)
    Screen.MousePointer = vbHourglass
    
    If mEditType = g���� Then
        strWhere = " And ������� Is Null"
    ElseIf mEditType = g�޸� Then
        strWhere = " and (������� Is Null Or �������=[2])"
    Else
        strWhere = " And �������=[2]"
    End If
    On Error GoTo errHandle
    strSQL = "" & _
        "   Select Decode(�������,Null,'','��') As ��־,ID,���㷽ʽ,���,������� " & _
        "   From �����¼ " & _
        "   Where ������� Is not  Null And (��¼״̬=1 and Ԥ����=1)  And ��λID=[1]" & strWhere & _
        "   Order By ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID, mlng�������)
    With vsԤ��
        .Redraw = flexRDNone
        .Clear 1
        .Tag = 0
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        mdbl�ۼ�Ԥ�� = 0: mdbl����Ԥ�� = 0
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("��־")) = NVL(rsTemp!��־)
            .Cell(flexcpData, i, .ColIndex("��־")) = NVL(rsTemp!ID)
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = NVL(rsTemp!���㷽ʽ)
            .TextMatrix(i, .ColIndex("������")) = Format(Val(NVL(rsTemp!���)), gVbFmtString.FM_���)
            .Cell(flexcpData, i, .ColIndex("������")) = Val(NVL(rsTemp!���))
            .TextMatrix(i, .ColIndex("�������")) = NVL(rsTemp!�������)
            
            mdbl�ۼ�Ԥ�� = mdbl�ۼ�Ԥ�� + Val(NVL(rsTemp!���))
            If Trim(.TextMatrix(i, .ColIndex("��־"))) = "��" Then
                mdbl����Ԥ�� = mdbl����Ԥ�� + Val(NVL(rsTemp!���))
            End If
            If Val(NVL(rsTemp!���)) < 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H0&
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    Call SetMoneyLbl
    Call zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlCommFun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub FullԤ��()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��䱾��Ԥ��
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long
    
    With mfrmEdit.vs��Ԥ��
        .Redraw = flexRDNone
        .Clear 1
        .Rows = 2
         lngRow = 1
        For i = 1 To vsԤ��.Rows - 1
            If Trim(vsԤ��.TextMatrix(i, 0)) = "��" Then
                .TextMatrix(lngRow, .ColIndex("ID")) = vsԤ��.Cell(flexcpData, i, vsԤ��.ColIndex("��־"))
                .TextMatrix(lngRow, .ColIndex("���ʽ")) = vsԤ��.TextMatrix(i, vsԤ��.ColIndex("���㷽ʽ"))
                .TextMatrix(lngRow, .ColIndex("�������")) = vsԤ��.TextMatrix(i, vsԤ��.ColIndex("�������"))
                .TextMatrix(lngRow, .ColIndex("������")) = vsԤ��.TextMatrix(i, vsԤ��.ColIndex("������"))
                If Val(.TextMatrix(lngRow, .ColIndex("������"))) < 0 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                Else
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H0&
                End If
                lngRow = lngRow + 1
                .Rows = .Rows + 1
            End If
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lng��� As Long
    '------------------------------------
    Select Case Control.ID
    Case mConMenu_Report    '���ñ���(ҩƷ�����ѯ)
        PrintReport 0
    Case conMenu_File_Print     '��ӡ֪ͨ��
        '��ӡ
        If mEditType = g���� Then Exit Sub
        printbill
    Case conMenu_Edit_Save:       '����
        Call zlSaveData
    Case conMenu_Edit_SelAll:       'ȫѡ
        Call zlAllSelData
    Case conMenu_Edit_ClsAll:       'ȫ��
        Call zlAllClsData
    Case conMenu_View_Backward:       '��һ��
        Call zlBackForward
    Case conMenu_View_Forward:       '��һ��
        Call zlBackward
    Case conMenu_Manage_Audit:       '���
        Call zlSaveCheck
    Case conMenu_Edit_ChargeOff  '����
        Call zlSaveStrike
    Case mConMenu_Hide_TempSave '������ʱ��Ϣ
        Call SaveTempData
    Case mConMenu_Hide_TempClearAll '�����ʱ��Ϣ
        Call SaveTempData(True)
    Case conMenu_View_Location
         If Trim(txtFind.Text) = "" Then Exit Sub
         FindRow Trim(txtFind.Text), IIf(vsPayList.Row + 1 >= vsPayList.Rows - 1, 1, vsPayList.Row + 1)
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsThis.RecalcLayout
    Case conMenu_View_FilterView '��������
        Call ShowFilterCon
    Case mConMenu_Popu_FP   '����Ʊͳ��
        mintͳ�Ʒ�ʽ = 0
        Call Setͳ�Ʒ�ʽ
        Call ���ܷ�Ʊ��Ϣ
    Case mConMenu_Popu_SH
        mintͳ�Ʒ�ʽ = 1
        Call Setͳ�Ʒ�ʽ
        Call ���ܷ�Ʊ��Ϣ
    Case conMenu_View_Refresh   '����ˢ������
        If mEditType = g���� Then
            Me.Tag = -1
            Call FillDeptDue
        Else
            Call initCard
        End If
        
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit
        Unload Me: Exit Sub
    Case Else
        If Control.ID > 401 And Control.ID < 499 Then
            '���ޱ�����
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlAllClsData()
    Dim lngLoop As Long
    Dim objTemp As Object
    
    If mEditType <> g���� And mEditType <> g�޸� And mEditType <> gԤ�� Then Exit Sub
    If Not (Me.ActiveControl Is vsPayList) And Not (Me.ActiveControl Is vsԤ��) Then vsPayList.SetFocus
    
    Set objTemp = Me.ActiveControl
    For lngLoop = 1 To objTemp.Rows - 1
        objTemp.TextMatrix(lngLoop, objTemp.ColIndex(IIf(objTemp.Name = "vsPayList", "ѡ��", "��־"))) = ""
    Next
    If objTemp Is vsPayList Then
        mdbl����Ӧ�� = 0
    Else
        mdbl����Ԥ�� = 0
    End If
    vsFp.Rows = 2
    vsFp.Clear 1
    Call SetMoneyLbl
End Sub

Private Sub zlAllSelData()
    '-----------------------------------------------------------------------------------------------------------
    '����:ȫѡ����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 19:30:01
    '-----------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim objTemp As Object
    Dim blnHaveData As Boolean
    
    If mEditType <> g���� And mEditType <> g�޸� And mEditType <> gԤ�� Then Exit Sub
    If Not (Me.ActiveControl Is vsPayList) And Not (Me.ActiveControl Is vsԤ��) Then vsPayList.SetFocus
    
    Set objTemp = Me.ActiveControl
    For lngLoop = 1 To objTemp.Rows - 1
        blnHaveData = False
        If Me.ActiveControl Is vsPayList Then
            blnHaveData = Trim(objTemp.Cell(flexcpData, lngLoop, objTemp.ColIndex(IIf(objTemp.Name = "vsPayList", "ѡ��", "��־")))) <> ""
        Else
            blnHaveData = Trim(objTemp.Cell(flexcpData, lngLoop, objTemp.ColIndex("���㷽ʽ"))) <> ""
        End If
        If blnHaveData Then
            If objTemp.Name = "vsPayList" Then
                If Trim(objTemp.TextMatrix(lngLoop, objTemp.ColIndex("��Ʊ��"))) = "" Then
                    objTemp.TextMatrix(lngLoop, objTemp.ColIndex("ѡ��")) = ""
                Else
'                    If mbln�����־ Then
'                        '���ҩƷ��
'                        If Trim(objTemp.TextMatrix(lngLoop, objTemp.ColIndex("�����־"))) = "����" And Val(objTemp.Cell(flexcpData, lngLoop, objTemp.ColIndex("�����־"))) = 1 _
'                            Or Trim(objTemp.TextMatrix(lngLoop, objTemp.ColIndex("�����־"))) <> "����" And Val(objTemp.Cell(flexcpData, lngLoop, objTemp.ColIndex("�����־"))) <> 1 Then
'                            objTemp.TextMatrix(lngLoop, objTemp.ColIndex("ѡ��")) = "��"
'                        Else
'                            objTemp.TextMatrix(lngLoop, objTemp.ColIndex("ѡ��")) = ""
'                        End If
'                    Else
                        objTemp.TextMatrix(lngLoop, objTemp.ColIndex("ѡ��")) = "��"
'                    End If
                End If
            Else
                objTemp.TextMatrix(lngLoop, objTemp.ColIndex("��־")) = "��"
            End If
        End If
    Next
    mblnChange = True
    If objTemp Is vsԤ�� Then
        mdbl����Ԥ�� = mdbl�ۼ�Ԥ��
    Else
        mdbl����Ӧ�� = mdbl�ۼ�Ӧ��
    End If
    Call ���ܷ�Ʊ��Ϣ
    Call SetMoneyLbl
End Sub

Private Sub zlBackward()
    '-----------------------------------------------------------------------------------------------------------
    '����:��һ��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-19 19:31:50
    '-----------------------------------------------------------------------------------------------------------
    ChangeMode 1
    zlControl.IsCtrlSetFocus vsPayList
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub zlBackForward()
    Dim dblCount As Double
    Dim lngRow As Long
    Dim i As Long, j As Long
    
    If mEditType = g���� Or mEditType = g�޸� Or mEditType = gԤ�� Then
        If mdbl����Ԥ�� < 0 Then
            MsgBox "���γ�Ԥ�����ܶ��С����", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        
        '����:��ϪҽԺҪ���������븺������Ϊ���ܴ����˻���������繩Ӧ�̲��湩����Ҫ���˻�����Ҫ�˿:2008-08-19 15:45:04
        '        '�������㷽ʽ��Ԥ�����ܶ���ۼ��Ƿ�Ϊ����
        '        Dim str���㷽ʽ As String
        '        Dim dbl��� As Double
        '        str���㷽ʽ = ","
        '        With vsԤ��
        '            For i = 1 To .Rows - 1
        '                dbl��� = 0
        '                If InStr(1, str���㷽ʽ, "," & .TextMatrix(i, .ColIndex("���㷽ʽ")) & ",") = 0 And Trim(.TextMatrix(i, .ColIndex("��־"))) = "��" Then
        '                    For j = 1 To .Rows - 1
        '                        If .TextMatrix(i, .ColIndex("���㷽ʽ")) = .TextMatrix(j, "���㷽ʽ") And Trim(.TextMatrix(j, .ColIndex("��־"))) = "��" Then
        '                            dbl��� = dbl��� + Val(.Cell(flexcpData, j, .ColIndex("������")))
        '                        End If
        '                    Next
        '                    If dbl��� < 0 Then
        '                        MsgBox "���㷽ʽΪ:" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & "���ܶ�ܸ���!", vbInformation + vbDefaultButton1, gstrSysName
        '                        Exit Sub
        '                    End If
        '                    str���㷽ʽ = str���㷽ʽ & .TextMatrix(i, .ColIndex("���㷽ʽ")) & ","
        '                End If
        '            Next
        '        End With
    End If
    mfrmEdit.zldbl����Ӧ�� = mdbl����Ӧ��
    mfrmEdit.zldbl����Ԥ�� = mdbl����Ԥ��
    '���б��γ�Ԥ��������
    If mEditType <> gԤ�� Then
        Call FullԤ��
    End If

    ChangeMode 2
    zlControl.IsCtrlSetFocus mfrmEdit.vsPayEdit
End Sub

Private Sub zlSaveData()
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-19 19:24:16
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim blnSuccess As Boolean
    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard
    If blnSuccess = False Then Exit Sub
    
    mblnChange = False
    If blnSuccess = True Then
        If IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
            '��ӡ
            '����27930 by lesfeng 2010-03-23
            If mint��� = 0 Then
                If InStr(mstrPrivs, ";����֪ͨ��;") <> 0 Then
                    printbill
                End If
            Else
                If InStr(mstrPrivs, ";��Ǹ��;") <> 0 Then
                    printbill
                End If
            End If
        End If
        mblnSuccess = True
        If mEditType = g�޸� Then    '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    Call mfrmEdit.ClearData
    vsFp.Clear 1
    vsFp.Rows = 2
    vsTemp.Clear 1
    vsTemp.Rows = 2
    txtDept.Text = "": txtDept.Tag = "-1": Me.Tag = "-1": mlng��λID = 0:
    ChangeMode 1
    FillDeptDue
    mblnSave = False
    mblnEdit = True
    mblnChange = False
End Sub

Private Sub zlSaveCheck()
    '-----------------------------------------------------------------------------------------------------------
    '����:��˴���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 19:33:35
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim blnSuccess As Boolean
    
    If mEditType = g��� Then        '���
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                '��ӡ
                '����27930 by lesfeng 2010-03-23
                If mint��� = 0 Then
                    If InStr(mstrPrivs, ";����֪ͨ��;") <> 0 Then
                        printbill
                    End If
                Else
                    If InStr(mstrPrivs, ";��Ǹ��;") <> 0 Then
                        printbill
                    End If
                End If
            End If
            mblnChange = False
            mblnSuccess = True
            Unload Me
        End If
        Exit Sub
    End If
End Sub

Private Sub zlSaveStrike()
    '-----------------------------------------------------------------------------------------------------------
    '����:���ó���
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-19 19:37:57
    '-----------------------------------------------------------------------------------------------------------

    Dim strReg As String
    Dim blnSuccess As Boolean
    
   If ValidData = False Then Exit Sub
    
   If mEditType = gȡ�� Then
        If SaveStrike() = True Then
            mblnChange = False
            mblnSuccess = True
            If IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                '��ӡ
                '����27930 by lesfeng 2010-03-23
                If mint��� = 0 Then
                    If InStr(mstrPrivs, ";����֪ͨ��;") <> 0 Then
                        printbill
                    End If
                Else
                    If InStr(mstrPrivs, ";��Ǹ��;") <> 0 Then
                        printbill
                    End If
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case mConMenu_Report
        Control.Enabled = InStr(mstrPrivs, ";ҩƷ�����ѯ;") > 0
    Case conMenu_File_Preview, conMenu_File_Excel
    Case conMenu_Edit_Save:       '����
        Control.Enabled = mintStep >= 2 And mblnChange
        Control.Visible = (mEditType = g�޸� Or mEditType = g���� Or mEditType = gԤ��)
    Case conMenu_Edit_SelAll:       'ȫѡ
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g�޸� Or mEditType = g���� Or mEditType = gԤ��)
    Case conMenu_Edit_ClsAll:       'ȫ��
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g�޸� Or mEditType = g���� Or mEditType = gԤ��)
    Case conMenu_View_FilterView    '����
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g�޸� Or mEditType = g����)
    Case conMenu_View_Backward:     '��һ��
        Control.Enabled = mintStep < 2
    Case conMenu_View_Forward:      '��һ��
        Control.Enabled = mintStep >= 2
    Case conMenu_Manage_Audit:      '���
        Control.Visible = mEditType = g���
    Case conMenu_Edit_ChargeOff  '����
        Control.Visible = mEditType = gȡ��
    Case conMenu_File_Print     '��ӡ֪ͨ��
        Control.Visible = Not (mEditType = g����) And InStr(mstrPrivs, ";����֪ͨ��;") > 0
    Case mConMenu_Hide_TempSave
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g�޸� Or mEditType = g����)
   Case mConMenu_Hide_TempClearAll   '������ʱ��Ϣ
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g�޸� Or mEditType = g����)
    Case conMenu_View_Location
        Control.Enabled = mintStep < 2
    Case conMenu_View_LocationItem
        Control.Enabled = mintStep < 2
    Case mConMenu_Popu_FP
        Control.Checked = mintͳ�Ʒ�ʽ <= 0
    Case mConMenu_Popu_SH
        Control.Checked = mintͳ�Ʒ�ʽ >= 1
    Case conMenu_View_Refresh
        Control.Enabled = mEditType = g�޸� Or mEditType = g����
        Control.Visible = Control.Enabled
'    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
'    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
'    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
'    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
'    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdSelDept_Click()
    Dim strTemp As String
    
    strTemp = frm��Ӧ��ѡ��.SelDept(mstrPrivs)
    If strTemp <> "" Then
        txtDept.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
        mlng��λID = Val(Left(strTemp, InStr(strTemp, ",") - 1))
        FillDeptDue
    End If
    Unload frm��Ӧ��ѡ��
    If vsPayList.Enabled Then vsPayList.SetFocus
End Sub

Private Function ShowFilterCon() As Boolean
    '------------------------------------------------------------------------------------------------------
    '����:��ʾ��������
    '���:
    '����:
    '����:����������,����true,���򷵻�false
    '�޸���:���˺�
    '�޸�ʱ��:2007/1/25
    '------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, lng��Ӧ��ID As Long
    If frm��������.ShowFind(Me, lng��Ӧ��ID, mstrPrivs, cllFilter) = False Then: Exit Function
    Set mcllFilter = cllFilter
 
    If Format(CDate(mcllFilter("�������")(0)), "yyyy-mm-dd") <> "1901-01-01" Then
        lblDate.Caption = "�������:" & Format(CDate(mcllFilter("�������")(0)), "yyyy-mm-dd") & " �� " & Format(CDate(mcllFilter("�������")(1)), "yyyy-mm-dd")
    Else
        lblDate.Caption = ""
    End If
    
    If Format(CDate(mcllFilter("��Ʊ����")(0)), "yyyy-mm-dd") <> "1901-01-01" Then
        lblDate.Caption = lblDate.Caption & IIf(lblDate.Caption = "", "", Space(5))
        lblDate.Caption = lblDate.Caption & "��Ʊ����:" & Format(CDate(mcllFilter("��Ʊ����")(0)), "yyyy-mm-dd") & " �� " & Format(CDate(mcllFilter("��Ʊ����")(1)), "yyyy-mm-dd")
    End If
    Me.Tag = ""
    mlng��λID = Val(mcllFilter("��Ӧ��ID"))
    Call FillDeptDue
    ShowFilterCon = True
End Function
 
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPanIndex.pane_���
        Item.Handle = mfrmEdit.hwnd
    Case mPanIndex.pane_Ӧ���б�
        Item.Handle = picPayList.hwnd
    Case mPanIndex.pane_Ԥ���б�
        Item.Handle = picԤ��.hwnd
    Case mPanIndex.pane_��ʱ����
        Item.Handle = picTemp.hwnd
    Case mPanIndex.pane_��Ʊ�ϼ�
        Item.Handle = picFp.hwnd
    End Select
End Sub

 
Private Sub Form_Activate()
    If mErrBillStatusInfor = �Ѿ�ɾ�� Then
        ShowMsgbox "�õ����Ѿ�������ɾ��,���ܼ���!"
        Unload Me
        Exit Sub
    End If
    If mErrBillStatusInfor = �Ѿ���� Then
        ShowMsgbox "�õ����Ѿ����������,�����ٽ������!"
        Unload Me
        Exit Sub
    End If
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    ChangeMode 1
    
    SetEditPro
    mblnChange = False
    If mEditType = g���� Then
        If ShowFilterCon = False Then Unload Me: Exit Sub
        If txtDept.Enabled And txtDept.Visible Then txtDept.SetFocus
    ElseIf mEditType = g�޸� Then
        If txtDept.Enabled And txtDept.Visible Then txtDept.SetFocus
    ElseIf mEditType = g�鿴 Then
    End If
    mblnChange = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    
    mbln�����־ = Val(zlDatabase.GetPara("�⹺�����Ҫ������Ǹ������ܽ��и������", glngSys, 0)) = 1
    mint��ʾ��λ = Val(zlDatabase.GetPara("��ʾ��λѡ��", glngSys, mlngModule))
     
    mblnFirst = True
    Call InitFilter
    Call InitPancel
    Call InitComandBars
    Call Setͳ�Ʒ�ʽ
    Call zl_vsGrid_Para_Restore(mlngModule, vsPayList, Me.Caption, "һ�㸶���б�")
    Call zl_vsGrid_Para_Restore(mlngModule, vsԤ��, Me.Caption, "һ��Ԥ���б�")
'    If mEditType <> g���� Then
'        Call initCard
'    End If
    Call initCard
    Call vsPayList_LostFocus
    Call vsԤ��_LostFocus
    Call vsTemp_LostFocus
    Call vsFp_LostFocus
    mintStep = 0

End Sub

Private Sub Setͳ�Ʒ�ʽ()
    stcFpTittle.Caption = IIf(mintͳ�Ʒ�ʽ = 0, "��ʱ��Ϣ-��Ʊ����", "��ʱ��Ϣ-���������")
    vsFp.TextMatrix(0, vsFp.ColIndex("��Ʊ��")) = IIf(mintͳ�Ʒ�ʽ = 0, "��Ʊ��", "�������")
End Sub

Private Sub ChangeMode(intMode As Integer)
    Dim panThis As Pane
    
    If intMode = mintStep Then Exit Sub
    mintStep = intMode
    If mintStep = 1 Then
        dkpMan.CloseAll
        dkpMan.ShowPane (mPanIndex.pane_Ӧ���б�)
        '����27930 by lesfeng 2010-03-23
        If mint��� = 0 Then
            dkpMan.ShowPane (mPanIndex.pane_Ԥ���б�)
        End If
        If mEditType = g���� Or mEditType = g�޸� Then
            dkpMan.ShowPane (mPanIndex.pane_��ʱ����)
        End If
        dkpMan.ShowPane (mPanIndex.pane_��Ʊ�ϼ�)
    ElseIf mintStep = 2 Then
        dkpMan.CloseAll
        dkpMan.ShowPane (mPanIndex.pane_���)
        Set panThis = dkpMan.FindPane(mPanIndex.pane_���)
        
        panThis.MaxTrackSize.Height = Me.ScaleHeight
        panThis.MaxTrackSize.Width = Me.ScaleWidth
        dkpMan_AttachPane panThis
        
    ElseIf mintStep = 3 Then
    End If
    dkpMan.RecalcLayout
    
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Width < 10245 Then Me.Width = 10245
    If Me.Height < 7140 Then Me.Height = 7140
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnYes As Boolean
    Call zl_vsGrid_Para_Save(mlngModule, vsPayList, Me.Caption, "һ�㸶���б�")
    Call zl_vsGrid_Para_Save(mlngModule, vsԤ��, Me.Caption, "һ��Ԥ���б�")
    
    If mblnChange Then
        ShowMsgbox "���Ѿ������˵�����Ϣ,�������˳��Ļ�," & vbCrLf & "�����ĵ����ݽ����ܱ���,���Ҫ�˳���?", True, blnYes
        If blnYes = False Then Cancel = 1: Exit Sub
    End If
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        'ʹ�ø��Ի�����
        Call zlDatabase.SetPara("��λ����", mstrFindKey, glngSys, mlngModule)
        Call zlDatabase.SetPara("ͳ�Ʒ�ʽ", mintͳ�Ʒ�ʽ, glngSys, mlngModule)
    End If
    If Not mfrmEdit Is Nothing Then Unload mfrmEdit
    Set mfrmEdit = Nothing
    SaveWinState Me, App.ProductName
End Sub

Private Sub mfrmEdit_InitCard(ByVal lng������� As Long, ByVal lng��λID As Long, ByVal str��λ���� As String)
    '��λ����:
    txtDept.Text = str��λ����
    txtDept.Tag = lng��λID
    mlng��λID = lng��λID: mlng������� = lng�������
End Sub

Private Sub mfrmEdit_zlChangeData(ByVal blnChange As Boolean)
    '���ݷ����ı�ʱ,�������¼�
    mblnChange = blnChange
End Sub

Private Sub picCon_Resize()
    Err = 0: On Error Resume Next
    With picCon
          stcTop.Top = .ScaleTop
          stcTop.Width = .ScaleWidth
          stcTop.Left = .ScaleLeft
          stcPayTittle.Move .ScaleLeft, stcTop.Height + stcTop.Top, .ScaleWidth
    End With
End Sub

Private Sub picPayList_Resize()
    Err = 0: On Error Resume Next
    With picPayList
        picCon.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        vsPayList.Move .ScaleLeft, picCon.Height, .ScaleWidth, .ScaleHeight - picCon.Height
    End With
End Sub

Private Sub picTemp_Resize()
    Err = 0: On Error Resume Next
    With picTemp
        stcTempTittle.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        vsTemp.Move .ScaleLeft, stcTempTittle.Top + stcTempTittle.Height, .ScaleWidth
        vsTemp.Height = .ScaleHeight - vsTemp.Top
    End With
End Sub

Private Sub picFp_Resize()
    Err = 0: On Error Resume Next
    With picFp
        stcFpTittle.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        vsFp.Move .ScaleLeft, stcFpTittle.Top + stcFpTittle.Height, .ScaleWidth
        vsFp.Height = .ScaleHeight - vsFp.Top
    End With
End Sub

Private Sub picԤ��_Resize()
    Err = 0: On Error Resume Next
    With picԤ��
        stcԤ��.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        vsԤ��.Move .ScaleLeft, stcԤ��.Top + stcԤ��.Height, .ScaleWidth, .ScaleHeight - (stcԤ��.Height + stcԤ��.Top)
    End With
End Sub

Private Sub ��ʱ�������ݴ���()
    '-----------------------------------------------------------------------------------------------------------
    '����:��������ܷ�������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-20 14:14:58
    '-----------------------------------------------------------------------------------------------------------
    Dim lngCur��� As Long, blnHaveTempSum As Boolean   '���ڻ������
    Dim i As Long
    Dim lngRow As Long
    Dim dbl��� As Double
    
    Err = 0: On Error GoTo ErrHand:
    With vsPayList
        lngCur��� = Val(.TextMatrix(.Row, .ColIndex("�������")))
        If lngCur��� <= 0 Then Exit Sub
        .TextMatrix(.Row, .ColIndex("�������")) = ""
        dbl��� = Val(.Cell(flexcpData, .Row, .ColIndex("��Ʊ���")))
        blnHaveTempSum = .FindRow(lngCur���, , .ColIndex("�������"), , True) > 0
        
    End With
    
    With vsTemp
        lngRow = .FindRow(lngCur���, 1, .ColIndex("���"), , True)
        If lngRow > 0 And lngRow <= .Rows - 1 Then
            If blnHaveTempSum Then
                .Cell(flexcpData, lngRow, .ColIndex("���")) = Val(.Cell(flexcpData, lngRow, .ColIndex("���"))) - dbl���
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("���"))), gVbFmtString.FM_���)
            Else
                If lngRow = .Rows - 1 And lngRow = 1 Then
                    .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                    .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                Else
                    .RemoveItem lngRow
                End If
            End If
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ���ܷ�Ʊ��Ϣ(Optional ByVal strInvoiceNO As String, Optional ByVal strParamInvoiceDate As String)
    '-----------------------------------------------------------------------------------------------------------
    '����:���ܷ�Ʊ��Ϣ�򷢻�����Ϣ
    '���:
    '  strInvoiceNO: ѡ����Ӧ������ϸ��Ʊ��
    '  strParamInvoiceDate: ѡ����Ӧ������ϸ��Ʊ����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-20 13:20:44
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, strNO As String, intCol As Integer, cllPro As Collection, strKey As String
    Dim bln��Ʊ�� As Boolean
    Dim dbl��� As Double
    Dim strInvoiceDate As String
    Dim varTmp As Variant
    Dim intCountCol As Integer
    Dim intOldRow As Integer
    Dim blnFind As Boolean
    
    intCountCol = vsPayList.ColIndex("���θ���")
    
    bln��Ʊ�� = IIf(mintͳ�Ʒ�ʽ = 0, True, False)
    Set cllPro = New Collection
    
    strNO = ""
    With vsPayList
        mdbl����Ӧ�� = 0

        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, .ColIndex("ѡ��")) <> "" And .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
                strInvoiceDate = IIf(Trim(.TextMatrix(i, .ColIndex("��Ʊ����"))) = "", "", .TextMatrix(i, .ColIndex("��Ʊ����")))
                intCol = IIf(bln��Ʊ��, .ColIndex("��Ʊ��"), .ColIndex("�������"))
                strKey = UCase(Trim(.TextMatrix(i, intCol))) & "_" & strInvoiceDate
                
                'mdbl����Ӧ�� = mdbl����Ӧ�� + Val(.Cell(flexcpData, i, intCountCol))
                mdbl����Ӧ�� = mdbl����Ӧ�� + Val(.TextMatrix(i, intCountCol))
                
                On Error Resume Next
                varTmp = cllPro(strKey)
                If Err.Number = 0 Then
                    '����
                    dbl��� = varTmp(1) + Val(.TextMatrix(i, intCountCol))
                    cllPro.Remove strKey
                Else
                    '������
                    dbl��� = Val(.TextMatrix(i, intCountCol))
                End If
                Err.Clear: On Error GoTo 0
                cllPro.Add Array(Mid(strKey, 1, InStr(strKey, "_") - 1), dbl���, strInvoiceDate), strKey
                
            End If
        Next
    
    End With
    
    '�������
    With vsFp
        intOldRow = .Row
        .Redraw = flexRDNone
        .Clear 1
        .Rows = 2
        For i = 1 To cllPro.Count
            .TextMatrix(i, .ColIndex("���")) = i
            .TextMatrix(i, .ColIndex("��Ʊ��")) = cllPro(i)(0)
            .TextMatrix(i, .ColIndex("��Ʊ����")) = Format(cllPro(i)(2), "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("���")) = Format(Val(cllPro(i)(1)), gVbFmtString.FM_���)
            If cllPro.Count > i Then
                .Rows = .Rows + 1
            End If
            '��դ��λ
            If blnFind = False Then
                If UCase(strInvoiceNO) = UCase(.TextMatrix(i, .ColIndex("��Ʊ��"))) And strParamInvoiceDate = .TextMatrix(i, .ColIndex("��Ʊ����")) Then
                    .Row = i
                    blnFind = True
                End If
            End If
        Next
        .TopRow = IIf(blnFind, .Row, 1)
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub vsFp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsFp, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsFp_AfterSort(ByVal Col As Long, Order As Integer)
    Dim intInvCol As Integer
    With vsFp
        
        If .Rows <= 1 Then Exit Sub
        
        intInvCol = .ColIndex("��Ʊ��")
        
        Select Case Col
            Case intInvCol
                .ColSort(Col) = Order
                .ColSort(.ColIndex("��Ʊ����")) = 0
                .Select 1, 0, .Rows - 1, .Cols - 1
                .Sort = flexSortUseColSort
                zl_VsGridAfterSort vsFp, Col, Order
            Case .ColIndex("��Ʊ����")
                .ColSort(Col) = Order
                .ColSort(intInvCol) = 0
                .Select 1, 0, .Rows - 1, .Cols - 1
                .Sort = flexSortUseColSort
                zl_VsGridAfterSort vsFp, Col, Order
        End Select
        
    End With
End Sub

Private Sub vsFp_GotFocus()
    zl_VsGridGotFocus vsFp
End Sub

Private Sub vsFp_LostFocus()
    zl_VsGridLOSTFOCUS vsFp
End Sub

Private Sub vsFp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
   
   If Button <> vbRightButton Then Exit Sub
 
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.ID, mcbrControl.Caption)
        cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub vsPayList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsPayList
        Select Case Col
        Case .ColIndex("��Ʊ��"), .ColIndex("��Ʊ����")
        Case .ColIndex("���θ���")
            If Val(.TextMatrix(Row, Col)) > .Cell(flexcpData, Row, .ColIndex("���θ���")) Then
                MsgBox "�����θ�����ڡ���δ�����=" & Format(.Cell(flexcpData, Row, .ColIndex("���θ���")), gVbFmtString.FM_���) & "����", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = Format(.Cell(flexcpData, Row, .ColIndex("���θ���")), gVbFmtString.FM_���)
            End If
        Case Else
            Exit Sub
        End Select
        
        If .Cell(flexcpData, Row, Col) <> .TextMatrix(Row, Col) And .ColIndex("���θ���") <> Col Then
            .Cell(flexcpForeColor, Row, Col) = vbRed
        ElseIf .ColIndex("���θ���") = Col Then
            .Cell(flexcpForeColor, Row, Col) = .ForeColor
        Else
            .Cell(flexcpForeColor, Row, Col) = .ForeColor
            Exit Sub
        End If
        
        If .TextMatrix(Row, .ColIndex("ѡ��")) = "" Then
            Call SetӦ��ѡ���־
        Else
            If Col = .ColIndex("���θ���") Then Call SetӦ��ѡ���־
        End If
    End With
End Sub

Private Sub vsPayList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Set mcbrControl = Me.cbsThis.FindControl(, mConMenu_Report)
    If Not mcbrControl Is Nothing Then
        mcbrControl.Enabled = Val(vsPayList.TextMatrix(NewRow, vsPayList.ColIndex("ϵͳ��ʶ"))) = 1 And InStr(mstrPrivs, ";ҩƷ�����ѯ;") > 0
    End If
    Call zl_VsGridRowChange(vsPayList, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsPayList_AfterSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridAfterSort(vsPayList, Col, Order)
End Sub

Private Sub vsPayList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mEditType <> g���� And mEditType <> g�޸� Then Cancel = True: Exit Sub
    
    With vsPayList
        Select Case Col
        Case .ColIndex("��Ʊ��"), .ColIndex("��Ʊ����")
            
'            If Trim(.TextMatrix(.Row, .ColIndex("��Ʊ��"))) = "" And .Row > 1 Then
'                If Trim(.TextMatrix(.Row - 1, .ColIndex("��Ʊ��"))) <> "" Then
'                    .TextMatrix(.Row, .ColIndex("��Ʊ��")) = .TextMatrix(.Row - 1, .ColIndex("��Ʊ��"))
'                    .TextMatrix(.Row, .ColIndex("��Ʊ����")) = .TextMatrix(.Row - 1, .ColIndex("��Ʊ����"))
'                    If .TextMatrix(.Row, .ColIndex("��־")) = "" Then
'                        Call vsPayList_DblClick
'                    End If
'                End If
'            End If
        Case .ColIndex("���θ���")
            Cancel = .TextMatrix(Row, .ColIndex("ѡ��")) = ""         'δѡ��������¼��
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsPayList_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vsPayList_EnterCell()
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
    
    With vsPayList
        .EditMaxLength = 0
        Select Case .Col
        Case .ColIndex("��Ʊ��")
            .EditMaxLength = 200
        Case .ColIndex("��Ʊ����")
            .EditMaxLength = 16
        End Select
    End With
End Sub

Private Sub vsPayList_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPayList
        Select Case Col
        Case .ColIndex("��Ʊ��")
        Case .ColIndex("��Ʊ����")
        Case Else
        End Select
        Call zlVsMoveGridCell(vsPayList, 0, .Cols - 1, False)
    End With
End Sub

Private Sub vsPayList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsPayList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    With vsPayList
        Select Case Col
        Case .ColIndex("��Ʊ��")
            Call VsFlxGridCheckKeyPress(vsPayList, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("��Ʊ����")
            '��Ҫ���ܴ����˿����
            Call VsFlxGridCheckKeyPress(vsPayList, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("���θ���")
            Call VsFlxGridCheckKeyPress(vsPayList, Row, Col, KeyAscii, m���ʽ)
        Case Else
        End Select
    End With
End Sub

Private Sub vsPayList_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If mEditType <> g�޸� And mEditType <> g���� Then Exit Sub
    If Set��Ʊ�ż���Ʊ���� Then
        '��Ҫ�Զ�ѡ�д�Ӧ����¼
        Call SetӦ��ѡ���־
    End If
End Sub

Private Function Set��Ʊ�ż���Ʊ����() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���÷�Ʊ�ż���Ʊ������Ϣ
    '���:
    '����:
    '����:�Զ������˵ģ�����ture,���򷵻�False
    '����:���˺�
    '����:2008-08-21 11:19:39
    '˵��:��Ҫ�ǵ�ǰ�е���Ϣ�������е���Ϣ������ȡ
    '-----------------------------------------------------------------------------------------------------------
    If mEditType <> g�޸� And mEditType <> g���� Then Exit Function
    If InStr(1, mstrPrivs, ";�޸ķ�Ʊ��Ϣ;") = 0 Then Exit Function
    
    With vsPayList
        If Trim(.TextMatrix(.Row, .ColIndex("��Ʊ��"))) = "" And .Row > 1 Then
            If Trim(.TextMatrix(.Row - 1, .ColIndex("��Ʊ��"))) <> "" Then '
                .TextMatrix(.Row, .ColIndex("��Ʊ��")) = .TextMatrix(.Row - 1, .ColIndex("��Ʊ��"))
                .TextMatrix(.Row, .ColIndex("��Ʊ����")) = .TextMatrix(.Row - 1, .ColIndex("��Ʊ����"))
                If .Cell(flexcpData, .Row, .ColIndex("��Ʊ��")) <> .TextMatrix(.Row, .ColIndex("��Ʊ��")) Then
                    .Cell(flexcpForeColor, .Row, .ColIndex("��Ʊ��")) = vbRed
                Else
                    .Cell(flexcpForeColor, .Row, .ColIndex("��Ʊ��")) = .ForeColor
                End If
                If .Cell(flexcpData, .Row, .ColIndex("��Ʊ����")) <> .TextMatrix(.Row, .ColIndex("��Ʊ����")) Then
                    .Cell(flexcpForeColor, .Row, .ColIndex("��Ʊ����")) = vbRed
                Else
                    .Cell(flexcpForeColor, .Row, .ColIndex("��Ʊ����")) = .ForeColor
                End If
                Set��Ʊ�ż���Ʊ���� = True
            End If
        End If
        Select Case .Col
        Case .ColIndex("��Ʊ��")
            .EditText = .TextMatrix(.Row, .ColIndex("��Ʊ��"))
        Case .ColIndex("��Ʊ����")
            .EditText = .TextMatrix(.Row, .ColIndex("��Ʊ����"))
        End Select
    End With
End Function

Private Sub vsPayList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
    
    With vsPayList
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("��Ʊ����")
            If strKey = "" Then Exit Sub
            strKey = CheckIsDate(strKey, "��Ʊ����")
            If strKey = "" Then Cancel = True: Exit Sub
            .EditText = strKey
        End Select
    End With
End Sub

Private Sub vsTemp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsTemp, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsTemp_GotFocus()
    zl_VsGridGotFocus vsTemp
End Sub

Private Sub vsTemp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, lng��� As Long
    With vsTemp
        If KeyCode = vbKeyDelete Then
            
            If (mEditType = g���� Or mEditType = g�޸�) And Val(.TextMatrix(.Row, .ColIndex("���"))) > 0 Then
                lng��� = Val(.TextMatrix(.Row, .ColIndex("���")))
                If MsgBox("�����Ҫɾ���������Ϊ" & lng��� & "  ����ʱ����������?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    '���������������
                    With vsPayList
                        For i = 1 To .Rows - 1
                            If Val(.TextMatrix(.Row, .ColIndex("�������"))) = lng��� Then
                                .TextMatrix(.Row, .ColIndex("�������")) = ""
                            End If
                        Next
                    End With
                    '�Ƴ���ǰ������
                    If .Rows - 1 = .Row And .Row = 1 Then
                        .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                    Else
                        .RemoveItem .Row
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsTemp_LostFocus()
    zl_VsGridLOSTFOCUS vsTemp

End Sub

Private Sub vsԤ��_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsԤ��, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsԤ��_AfterSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridAfterSort(vsԤ��, Col, Order)
End Sub

Private Sub vsԤ��_DblClick()
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
    
    With vsԤ��
        If .Col <> .ColIndex("��־") Then Exit Sub
        If Trim(.TextMatrix(.Row, .ColIndex("���㷽ʽ"))) <> "" Then
            .TextMatrix(.Row, .ColIndex("��־")) = IIf(Trim(.TextMatrix(.Row, .ColIndex("��־"))) = "", "��", "")
            If Trim(.TextMatrix(.Row, 0)) = "" Then
                mdbl����Ԥ�� = mdbl����Ԥ�� - Val(.TextMatrix(.Row, .ColIndex("������")))
            Else
                mdbl����Ԥ�� = mdbl����Ԥ�� + Val(.TextMatrix(.Row, .ColIndex("������")))
            End If
        End If
    End With
    Call SetMoneyLbl
End Sub

Private Sub SetMoneyLbl()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ñ�ǩ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    mfrmEdit.zldbl����Ӧ�� = mdbl����Ӧ��
    mfrmEdit.zldbl����Ԥ�� = mdbl����Ԥ��
    lbl���(1).Caption = "�ۼ�Ӧ��:" & Format(mdbl�ۼ�Ӧ��, "###0.00;-###0.00;0.00;0.00") & ""
    '����27930 by lesfeng 2010-03-23
    If mint��� = 0 Then
        lbl���(2).Caption = "������:" & Format(mdbl����Ӧ��, "###0.00;-###0.00;0.00;0.00") & ""
        lbl���(3).Caption = "Ԥ���ۼ�:" & Format(mdbl�ۼ�Ԥ��, "###0.00;-###0.00;0.00;0.00") & ""
        lbl���(4).Caption = "��Ԥ��:" & Format(mdbl����Ԥ��, "###0.00;-###0.00;0.00;0.00") & ""
        lbl���(5).Caption = "����Ӧ��:" & Format(mdbl����Ӧ�� - mdbl����Ԥ��, "###0.00;-###0.00;0.00;0.00") & ""
    Else
        lbl���(2).Caption = "��Ǹ�����:" & Format(mdbl����Ӧ��, "###0.00;-###0.00;0.00;0.00") & ""
        lbl���(3).Caption = "Ԥ���ۼ�:" & Format(mdbl�ۼ�Ԥ��, "###0.00;-###0.00;0.00;0.00") & ""
        lbl���(4).Caption = "��Ԥ��:" & Format(mdbl����Ԥ��, "###0.00;-###0.00;0.00;0.00") & ""
        lbl���(5).Caption = "���α��Ӧ��:" & Format(mdbl����Ӧ�� - mdbl����Ԥ��, "###0.00;-###0.00;0.00;0.00") & ""
    End If
End Sub

Private Sub vsԤ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        vsԤ��_DblClick
    ElseIf KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub vsԤ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, blnHaveData As Boolean
    
    If Button <> 2 Then Exit Sub
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
    
    blnHaveData = False
    With vsԤ��
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, vsԤ��.ColIndex("��־")) <> "" Then blnHaveData = True: Exit For
        Next
    End With
    If blnHaveData = False Then Exit Sub
    If vsԤ��.Enabled Then vsԤ��.SetFocus
End Sub

Private Sub vsPayList_Click()
    With vsPayList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
           ' SetColumnSort vsPayList, mintPreCol, mintsort
            Exit Sub
         End If
    End With
End Sub

Private Sub vsPayList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
   
    With vsPayList
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeySpace
                Call vsPayList_DblClick
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sngWidth = sngWidth + .ColWidth(i)
                    If sngWidth > .Width Then
                        .LeftCol = i + 1
                        .Col = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
 
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    With vsPayList
        Select Case .Col
        Case .Cols - 1
            .Col = 0: .LeftCol = 0
            Exit Sub
        End Select
        
        Call zlVsMoveGridCell(vsPayList, 0, vsPayList.Cols - 1, False)
    End With
End Sub

Private Sub vsPayList_DblClick()
    Dim intCol As Integer
    With vsPayList
        
        If mEditType <> g���� And mEditType <> g�޸� And mEditType <> gԤ�� Then
            If Val(.TextMatrix(.Row, .ColIndex("ϵͳ��ʶ"))) = 1 And InStr(mstrPrivs, ";ҩƷ�����ѯ;") > 0 Then
                '���ñ���
                PrintReport 1
            End If
            Exit Sub
        End If
        
        If Trim(.TextMatrix(.Row, .ColIndex("��Ʊ��"))) = "" Then Exit Sub
        
        If .Col <> .ColIndex("ѡ��") And .Col <> .ColIndex("���θ���") Then
            If Val(.TextMatrix(.Row, .ColIndex("ϵͳ��ʶ"))) = 1 And InStr(mstrPrivs, ";ҩƷ�����ѯ;") > 0 Then
                '���ñ���
                PrintReport 1
            End If
            Exit Sub
        End If
        
        Call SetӦ��ѡ���־
        If .TextMatrix(.Row, .ColIndex("ѡ��")) = "��" Then
            Call Set��Ʊ�ż���Ʊ����
        End If
        mblnChange = True
    End With
End Sub

Private Sub SetӦ��ѡ���־()
    '-----------------------------------------------------------------------------------------------------------
    '����:����Ӧ��ѡ���־
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-21 11:25:11
    '-----------------------------------------------------------------------------------------------------------
    Dim intCol As Integer
    Dim strInvoice As String
    Dim strInvoiceDate As String
    
    If mEditType <> g���� And mEditType <> g�޸� And mEditType <> gԤ�� Then Exit Sub

    With vsPayList
        If Trim(.Cell(flexcpData, .Row, .ColIndex("ѡ��"))) = "" Then Exit Sub
'        If mbln�����־ Then
'            If .TextMatrix(.Row, .ColIndex("�����־")) = "����" And Val(.Cell(flexcpData, .Row, .ColIndex("�����־"))) = 1 _
'                Or .TextMatrix(.Row, .ColIndex("�����־")) <> "����" And Val(.Cell(flexcpData, .Row, .ColIndex("�����־"))) <> 1 Then
'                .TextMatrix(.Row, .ColIndex("ѡ��")) = IIf(.TextMatrix(.Row, .ColIndex("ѡ��")) = "", "��", "")
'            Else
'                .TextMatrix(.Row, .ColIndex("ѡ��")) = ""
'                Exit Sub
'            End If
'        Else
            '.TextMatrix(.Row, .ColIndex("ѡ��")) = IIf(.TextMatrix(.Row, .ColIndex("ѡ��")) = "", "��", "")
'        End If
        If vsPayList.Col <> vsPayList.ColIndex("���θ���") Then
            .TextMatrix(.Row, .ColIndex("ѡ��")) = IIf(.TextMatrix(.Row, .ColIndex("ѡ��")) = "", "��", "")
            strInvoice = IIf(mintͳ�Ʒ�ʽ = 0, .TextMatrix(.Row, .ColIndex("��Ʊ��")), .TextMatrix(.Row, .ColIndex("�������")))
            strInvoiceDate = .TextMatrix(.Row, .ColIndex("��Ʊ����"))
        End If

'        If Trim(.TextMatrix(.Row, .ColIndex("ѡ��"))) = "" Then
'            mdbl����Ӧ�� = mdbl����Ӧ�� - Val(.TextMatrix(.Row, .ColIndex("���θ���")))
'        Else
'            mdbl����Ӧ�� = mdbl����Ӧ�� + Val(.TextMatrix(.Row, .ColIndex("���θ���")))
'        End If
        mblnChange = True
    End With
    Call ��ʱ�������ݴ���
    Call ���ܷ�Ʊ��Ϣ(strInvoice, strInvoiceDate)
    Call SetMoneyLbl
End Sub

Private Sub vsPayList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, blnHaveData As Boolean
    
    If Button <> 2 Then Exit Sub
    If mEditType <> g���� And mEditType <> g�޸� Then Exit Sub
    blnHaveData = False
    With vsPayList
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, .ColIndex("ѡ��")) <> "" Then blnHaveData = True: Exit Sub
        Next
    End With
    
    If blnHaveData = False Then Exit Sub
    If vsPayList.Enabled Then vsPayList.SetFocus
 End Sub

Private Sub txtDept_Change()
    mlng��λID = 0
End Sub

Private Sub txtDept_GotFocus()
    SetTxtGotFocus txtDept, True
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtDept.Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Val(txtDept.Tag) <> 0 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If SelMltProvide = False Then
        Exit Sub
    End If
End Sub

Private Sub FillDeptDue()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ز�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select ID,����,����,��ַ,�绰,��������,˰��ǼǺ�,������ From ��Ӧ�� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID)
    Err.Clear: On Error GoTo 0
    If Not rsTemp.EOF Then
        txtDept.Text = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
        mlng��λID = Val(NVL(rsTemp!ID))
        mlngЧ�� = NVL(rsTemp!������, 0)
    End If
    Call mfrmEdit.zlLoadPrivder(mlng��λID)
    zlControl.IsCtrlSetFocus vsPayList
    If mlng��λID <> Val(Me.Tag) Then
        Me.Tag = mlng��λID
        LoadPayMoney
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub txtDept_LostFocus()
    ImeLanguage False
End Sub

Private Function ValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:��֤�Ϸ�,����True,����=false
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim lngRow As Long
    Dim strTemp As String
    Dim dblCount As Double
    If mlng��λID = 0 Then
         ShowMsgbox "��Ӧ��ѡ������,������ѡ��!"
         Call zlBackForward
         If txtDept.Enabled Then txtDept.SetFocus
         Exit Function
    End If
    If mfrmEdit.zlValidData = False Then
        Exit Function
    End If
    If mEditType = g���� Or mEditType = g�޸� Then
        If InStr(1, mstrPrivs, "�޸ķ�Ʊ��Ϣ") > 0 Then
            With vsPayList
                For lngRow = 1 To .Rows - 1
                    If Trim(.Cell(flexcpData, lngRow, .ColIndex("ѡ��"))) <> "" _
                        And Trim(.TextMatrix(lngRow, .ColIndex("ѡ��"))) <> "" Then
                        If Trim(.TextMatrix(.Row, .ColIndex("��Ʊ��"))) = "" Then
                            ShowMsgbox "��Ʊ��δ���룬����!"
                            If mintStep >= 2 Then
                                Call zlBackForward
                            End If
                            .Col = .ColIndex("��Ʊ��"): .Row = lngRow: .TopRow = lngRow
                            zlControl.IsCtrlSetFocus vsPayList
                            Exit Function
                        End If
                        strTemp = .TextMatrix(.Row, .ColIndex("��Ʊ����"))
                        If strTemp = "" Then
                            ShowMsgbox "��Ʊ����δ���룬����!"
                            If mintStep >= 2 Then
                                Call zlBackForward
                            End If
                            .Col = .ColIndex("��Ʊ����"): .Row = lngRow: .TopRow = lngRow
                            zlControl.IsCtrlSetFocus vsPayList
                            Exit Function
                        End If
                        If IsDate(strTemp) = False Or IsNumeric(strTemp) Then
                            ShowMsgbox "����ķ�Ʊ���ڲ����������ͣ�����!"
                            If mintStep >= 2 Then
                                Call zlBackForward
                            End If
                            .Col = .ColIndex("��Ʊ����"): .Row = lngRow: .TopRow = lngRow
                            zlControl.IsCtrlSetFocus vsPayList
                            Exit Function
                        End If
                    End If
                Next
            End With
        End If
    End If
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-19 15:20:25
    '-----------------------------------------------------------------------------------------------------------
    Dim cllPro As New Collection
    Dim strNO_IN As String
    Dim lng�������_IN As Long
    Dim str��Ʊ�� As String, str��Ʊ���� As String
    Dim lngRow As Long
    Dim varData As Variant

    SaveCard = False
    Set cllPro = New Collection
    
    Err = 0: On Error GoTo errHandle:
    
    If mEditType = gԤ�� Then
        'Ԥ��
        If vsPayList.Rows <= 1 Then Exit Function
        'ID, ��¼״̬, �ƻ�����
        varData = Split(vsPayList.Cell(flexcpData, 1, vsPayList.ColIndex("ѡ��")), ",")
        '����Ԥ���־
        gstrSQL = "Zl_�������_CheckClear(" & _
                  varData(0) & "," & _
                  mlng������� & _
                  ")"
        AddArray cllPro, gstrSQL
        With vsPayList
            For lngRow = 1 To .Rows - 1
                If Trim(.TextMatrix(lngRow, .ColIndex("ѡ��"))) <> "" Then
                    varData = Split(.Cell(flexcpData, lngRow, .ColIndex("ѡ��")), ",")
                    gstrSQL = "Zl_�������_Check(" & varData(0)     '��Ԥ���־
                    gstrSQL = gstrSQL & "," & mlng�������          '�������
                    gstrSQL = gstrSQL & ",'" & UserInfo.���� & "'"  'Ԥ����
                    gstrSQL = gstrSQL & ",to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd hh24:mi:ss' ) ) "       'Ԥ������
                    
                    AddArray cllPro, gstrSQL
                End If
            Next
        End With
    Else
        '����ı���
        If mfrmEdit.zlSaveCard(cllPro, lng�������_IN, strNO_IN) = False Then Exit Function
        
        '��Ӧ�ɹ��嵥
        With vsPayList
            For lngRow = 1 To .Rows - 1
                If Trim(.TextMatrix(lngRow, .ColIndex("ѡ��"))) <> "" Then
                    If InStr(1, mstrPrivs, "�޸ķ�Ʊ��Ϣ") > 0 Then
                        str��Ʊ�� = Trim(.TextMatrix(lngRow, .ColIndex("��Ʊ��")))
                        If Trim(.Cell(flexcpData, lngRow, .ColIndex("��Ʊ��"))) = str��Ʊ�� Then
                            '���û�����ı䣬�Ͳ����ķ�Ʊ����
                            str��Ʊ�� = "NULL"
                        Else
                            str��Ʊ�� = "'" & str��Ʊ�� & "'"
                        End If
                        str��Ʊ���� = Trim(.TextMatrix(lngRow, .ColIndex("��Ʊ����")))
                        If Trim(.Cell(flexcpData, lngRow, .ColIndex("��Ʊ����"))) = str��Ʊ���� Or str��Ʊ���� = "" Then
                            '���û�����ı䣬�Ͳ����ķ�Ʊ����
                            str��Ʊ���� = "NULL"
                        Else
                            str��Ʊ���� = "To_date('" & str��Ʊ���� & "','yyyy-mm-dd')"
                        End If
                    Else
                        str��Ʊ�� = "NULL": str��Ʊ���� = "NULL"
                    End If
                    
                    ' .Cell(flexcpData, .Row, .ColIndex("��־")) : 'ID,��¼״̬ ,�ƻ�����
                    varData = Split(.Cell(flexcpData, lngRow, .ColIndex("ѡ��")), ",")
                    '���̲���
                    'Zl_�������_Update(
                    gstrSQL = "zl_�������_UPDATE("
                    'Id_In       In Varchar2 := Null,
                    gstrSQL = gstrSQL & Val(varData(0)) & ","
                    '�ƻ����_In In Varchar2 := Null, --��0,1,2,3��ʽ����
                    gstrSQL = gstrSQL & "NULL,"
                    '�������_In In �����¼.�������%Type := Null,
                    gstrSQL = gstrSQL & "" & lng�������_IN & ","
                    'Ԥ����_In   In �����¼.Ԥ����%Type := 0,
                    gstrSQL = gstrSQL & "" & 0 & ","
                    '��Ʊ���_In     In Ӧ����¼.��Ʊ���%Type := 0
                    gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("��Ʊ���"))) & ","
                    '���θ���_In
                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(lngRow, .ColIndex("���θ���"))) & ","
                    '--Ӧ����¼:��Ʊ�źͷ�Ʊ����ΪNULL������£���������Ӧ����¼�еķ�Ʊ�ţ�ͬʱ��ֻ�ܸ�������ͨ����Ŵ���Ʊ��
                    '��Ʊ��_In   In Ӧ����¼.��Ʊ��%Type := Null,
                    gstrSQL = gstrSQL & "" & str��Ʊ�� & ","
                    '��Ʊ����_In In Ӧ����¼.��Ʊ����%Type := Null
                    gstrSQL = gstrSQL & "" & str��Ʊ���� & ")"
                    AddArray cllPro, gstrSQL
                End If
            Next
        End With
        
        '����Ԥ����
        With vsԤ��
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, .ColIndex("��־")) <> "" Then
                    'Zl_�������_Update
                    gstrSQL = "zl_�������_UPDATE("
                    '  Id_In       In Varchar2 := Null,
                    gstrSQL = gstrSQL & Val(.Cell(flexcpData, lngRow, .ColIndex("��־"))) & ","
                    '  �ƻ����_In In Varchar2 := Null, --��0,1,2,3��ʽ����
                    gstrSQL = gstrSQL & "NULL,"
                    '  �������_In In �����¼.�������%Type := Null,
                    gstrSQL = gstrSQL & "" & lng�������_IN & ","
                    '  Ԥ����_In   In �����¼.Ԥ����%Type := 0,
                    gstrSQL = gstrSQL & "" & 1 & ","
                    '  ���_In     In Ӧ����¼.��Ʊ���%Type := 0
                    gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("������"))) & ")"
                    '--Ӧ����¼:��Ʊ�źͷ�Ʊ����ΪNULL������£���������Ӧ����¼�еķ�Ʊ�ţ�ͬʱ��ֻ�ܸ�������ͨ����Ŵ���Ʊ��
                    '��Ʊ��_In   In Ӧ����¼.��Ʊ��%Type := Null,
                    '��Ʊ����_In In Ӧ����¼.��Ʊ����%Type := Null
                    AddArray cllPro, gstrSQL
                End If
            Next
        End With
    End If
    
    Err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllPro, Me.Caption, True
    If mEditType <> gԤ�� Then
        If Check������Ӧ����ϸ(lng�������_IN) = False Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    '�ύ����
    gcnOracle.CommitTrans
    SaveCard = True
    Me.stbThis.Panels(2).Text = "���ŵ��ݺ�Ϊ:" & strNO_IN
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog

    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetEditPro()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ñ༭����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    cmdSelDept.Enabled = mEditType = g����
    txtDept.Enabled = mEditType = g����
End Sub

Private Function SaveCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��˵���
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO_IN As String, cllPro As New Collection
    
    SaveCheck = False
    
    strNO_IN = mfrmEdit.txtNo
    If mfrmEdit.zlCheck(cllPro) = False Then
        ChangeMode 2
        zlControl.IsCtrlSetFocus mfrmEdit.vsPayEdit
        Exit Function
    End If
    '   zl_�������_VERIFY(NO_IN);
    gstrSQL = "zl_�������_VERIFY('" & _
        strNO_IN & "')"
    AddArray cllPro, gstrSQL
    
    Err = 0: On Error GoTo errHandle:
    ExecuteProcedureArrAy cllPro, Me.Caption
    SaveCheck = True
    Exit Function
    
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
 '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO_IN As String
    
    SaveStrike = False
    
    strNO_IN = mfrmEdit.txtNo
    On Error GoTo errHandle:
    '   zl_�������_VERIFY(NO_IN);
    gstrSQL = "zl_�������_strike('" & _
        strNO_IN & "')"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveStrike = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
'��ӡ����
Private Sub printbill()
    '����27930 by lesfeng 2010-03-23
    If mint��� = 0 Then
        ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_1", Me, "���ݱ��=" & mfrmEdit.txtNo.Tag, "��¼״̬=" & mint��¼״̬
    Else
        ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_3", Me, "���ݱ��=" & mfrmEdit.txtNo.Tag, "��¼״̬=" & mint��¼״̬
    End If
End Sub

Private Function SelMltProvide() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ��Ӧ������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String, vRect As RECT, lngH As Long, blnCancel As Boolean
    Dim strȨ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    
    If Trim(txtDept.Text) = "" Then Exit Function
    
    strKey = GetMatchingSting(UCase(txtDept.Text), False)
    SelMltProvide = False
    
    strȨ�� = " and " & Get����Ȩ��(mstrPrivs)
    
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    
    strSQL = "" & _
        "  Select   ID,����,����,����,���֤��," & _
        "           to_char(���֤Ч��,'yyyy-mm-dd') as ���֤Ч��,ִ�պ�," & _
        "           to_char(ִ��Ч��,'yyyy-mm-dd') as ִ��Ч��,˰��ǼǺ�,��ϵ�� " & _
        "  From  ��Ӧ�� " & _
        "   Where (����ʱ�� is null or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01') " & zl_��ȡվ������() & "   " & _
        "           and ĩ��=1 And ( ���� Like upper([1]) or ���� like [1] or ����  like  upper([1]) ) " & strȨ��
    
    
    vRect = zlControl.GetControlRect(txtDept.hwnd)
    lngH = txtDept.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Ӧ��ѡ��", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If txtDept.Enabled Then txtDept.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "û���ҵ����������Ĺ�Ӧ��,����!"
        If txtDept.Enabled Then txtDept.SetFocus
        Exit Function
    End If
    If rsTemp.State = 0 Then Exit Function
    If txtDept.Enabled Then txtDept.SetFocus
    
    txtDept.Text = NVL(rsTemp!����)
    mlng��λID = NVL(rsTemp!ID, 0)
    txtDept.Tag = mlng��λID
    '�������
    FillDeptDue
    zlCommFun.PressKey vbKeyTab
    SelMltProvide = True
End Function
 
Private Sub vsPayList_GotFocus()
    Call zl_VsGridGotFocus(vsPayList)
End Sub

Private Sub vsPayList_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsPayList)
End Sub

Private Sub vsԤ��_GotFocus()
    Call zl_VsGridGotFocus(vsԤ��)
End Sub

Private Sub vsԤ��_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsԤ��)
End Sub

Private Sub FindRow(ByVal strFind As String, Optional lngRow As Long = 1)
    '����:����ָ�е������Ƿ�������ص�����
    '����:intMachType:0-��ƥ��,1-��ȫƥ��
    Dim i As Long, lngCol As Long
    Dim blnAll As Boolean
    With vsPayList
        lngCol = .ColIndex(mstrFindKey)
        If lngCol < 0 Then Exit Sub
        Select Case lngCol
        Case .ColIndex("�������")
            blnAll = True
        Case .ColIndex("����")
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_����)
        Case .ColIndex("��Ʊ���")
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_���)
        Case Else
            blnAll = False
        End Select
       i = .FindRow(strFind, lngRow, lngCol, False, blnAll)
       If i > 0 Then
            .Row = i: .TopRow = i
       Else
            ShowMsgbox "�Ѿ��鵽ĩβ,û�з�����������������,����!"
       End If
    End With
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    If mstrFindKey = "����" Then
        zlCommFun.OpenIme True
    Else
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNO As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtFind) = "" Then Exit Sub
    FindRow Trim(txtFind.Text)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtFind, KeyAscii, m�ı�ʽ
End Sub

Private Sub PrintReport(ByVal bytStyle As Byte) '(ByVal lngDeptID As Long, ByVal lngPurveryID As Long, ByVal lngDrugID As Long)
    Dim lngDeptID As Long, lngDrugID As Long
    Dim strDept As String, strDrug As String, strSupplier As String
    
    strSupplier = Mid(txtDept.Text, InStr(txtDept.Text, "-") + 1)
    With vsPayList
        lngDeptID = Val(.TextMatrix(.Row, .ColIndex("�ⷿID")))
        lngDrugID = Val(.TextMatrix(.Row, .ColIndex("ҩƷID")))
        strDept = .TextMatrix(.Row, .ColIndex("��ǰ�ⷿ"))
        strDrug = .TextMatrix(.Row, .ColIndex("Ʒ��"))
    End With
    
    If bytStyle = 1 Then
        ReportOpen gcnOracle, glngSys, "ZL1_REPORT_1323", Me, _
            "ҩƷ����=" & strDrug & "|" & lngDrugID, _
            "�ⷿ=" & strDept & "|" & IIf(lngDeptID = 0, " is not null ", "=" & lngDeptID), _
            "��ҩ��λ=" & strSupplier & "|" & mlng��λID
    Else
        ReportOpen gcnOracle, glngSys, "ZL1_REPORT_1323", Me, _
            "��ҩ��λ=" & strSupplier & "|" & mlng��λID
    End If
End Sub
