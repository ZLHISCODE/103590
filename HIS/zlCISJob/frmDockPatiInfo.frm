VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#4.0#0"; "ZlPatiAddress.ocx"
Begin VB.Form frmDockPatiInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "������Ϣ"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   129
      Top             =   11160
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   635
      SimpleText      =   $"frmDockPatiInfo.frx":0000
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDockPatiInfo.frx":0047
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15558
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
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8235
      Index           =   2
      Left            =   600
      ScaleHeight     =   8235
      ScaleWidth      =   9915
      TabIndex        =   64
      Top             =   4920
      Width           =   9915
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   24
         Left            =   4065
         MaxLength       =   100
         TabIndex        =   84
         Top             =   1170
         Width           =   3270
      End
      Begin VB.CommandButton cmdE 
         Height          =   220
         Index           =   23
         Left            =   2370
         Picture         =   "frmDockPatiInfo.frx":08DB
         Style           =   1  'Graphical
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   1035
         Width           =   240
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   23
         Left            =   810
         MaxLength       =   20
         TabIndex        =   83
         Text            =   "2013-06-20 18:00"
         Top             =   1065
         Width           =   1500
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   675
         TabIndex        =   135
         Top             =   825
         Width           =   750
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   -30
            Width           =   700
         End
      End
      Begin VB.TextBox txtSL 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Left            =   105
         MaxLength       =   3
         TabIndex        =   81
         Top             =   825
         Width           =   300
      End
      Begin VB.CheckBox chkNoAller 
         Appearance      =   0  'Flat
         Caption         =   "�޹�����¼"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1515
         TabIndex        =   134
         Top             =   1290
         Width           =   1215
      End
      Begin VB.CommandButton cmdDiagMove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Height          =   375
         Index           =   2
         Left            =   9360
         Picture         =   "frmDockPatiInfo.frx":09D1
         Style           =   1  'Graphical
         TabIndex        =   133
         TabStop         =   0   'False
         ToolTipText     =   "�������"
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton cmdDiagMove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Height          =   375
         Index           =   3
         Left            =   9360
         Picture         =   "frmDockPatiInfo.frx":2F79
         Style           =   1  'Graphical
         TabIndex        =   132
         TabStop         =   0   'False
         ToolTipText     =   "�������"
         Top             =   5280
         Width           =   375
      End
      Begin VB.CommandButton cmdDiagMove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Height          =   375
         Index           =   0
         Left            =   9480
         Picture         =   "frmDockPatiInfo.frx":56A4
         Style           =   1  'Graphical
         TabIndex        =   131
         TabStop         =   0   'False
         ToolTipText     =   "�������"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdDiagMove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Height          =   375
         Index           =   1
         Left            =   9480
         Picture         =   "frmDockPatiInfo.frx":7C4C
         Style           =   1  'Graphical
         TabIndex        =   130
         TabStop         =   0   'False
         ToolTipText     =   "�������"
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton cmdSaveZY 
         Height          =   300
         Left            =   8595
         Picture         =   "frmDockPatiInfo.frx":A377
         Style           =   1  'Graphical
         TabIndex        =   118
         TabStop         =   0   'False
         ToolTipText     =   "����ǰժҪ����Ϊ����ժҪ��"
         Top             =   1110
         Width           =   300
      End
      Begin VB.CommandButton cmdShowZY 
         Caption         =   "��"
         Height          =   300
         Left            =   8280
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   1110
         Width           =   300
      End
      Begin VB.PictureBox picPanel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2280
         Index           =   6
         Left            =   1260
         ScaleHeight     =   2280
         ScaleWidth      =   8145
         TabIndex        =   105
         Top             =   5775
         Width           =   8145
         Begin VB.CommandButton cmdE 
            Caption         =   "��"
            Height          =   220
            Index           =   25
            Left            =   2745
            TabIndex        =   112
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   1695
            Width           =   240
         End
         Begin VB.TextBox txtE 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   26
            Left            =   4725
            MaxLength       =   100
            TabIndex        =   94
            Top             =   1710
            Width           =   3270
         End
         Begin VB.TextBox txtE 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   25
            Left            =   810
            MaxLength       =   100
            TabIndex        =   93
            Top             =   1695
            Width           =   1965
         End
         Begin VB.Frame fraC 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   22
            Left            =   4605
            TabIndex        =   108
            Top             =   30
            Width           =   3390
            Begin VB.ComboBox cboE 
               BackColor       =   &H00FDFDFD&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   22
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   -30
               Width           =   3345
            End
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            Caption         =   "��Ⱦ���ϴ�"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1995
            TabIndex        =   90
            Top             =   30
            Width           =   1260
         End
         Begin VB.PictureBox picPanel 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   5
            Left            =   0
            ScaleHeight     =   360
            ScaleWidth      =   1845
            TabIndex        =   107
            Top             =   45
            Width           =   1845
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               Caption         =   "����"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   945
               TabIndex        =   106
               Top             =   0
               Width           =   720
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               Caption         =   "����"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   0
               TabIndex        =   89
               Top             =   0
               Width           =   690
            End
         End
         Begin zl9CISJob.UCPatiVitalSigns UCPatiVitalSigns 
            Height          =   660
            Left            =   150
            TabIndex        =   95
            Top             =   450
            Width           =   5610
            _ExtentX        =   9895
            _ExtentY        =   1164
            TextBackColor   =   -2147483643
            LblBackColor    =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Style           =   1
            XDis            =   200
            YDis            =   200
            LabToTxt        =   0
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   28
            Left            =   0
            TabIndex        =   127
            Top             =   2025
            Width           =   720
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            Caption         =   "ȥ��"
            Height          =   180
            Index           =   22
            Left            =   3780
            TabIndex        =   111
            Top             =   30
            Width           =   360
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            Caption         =   "����ҽѧ��ʾ"
            Height          =   180
            Index           =   26
            Left            =   3060
            TabIndex        =   110
            Top             =   1755
            Width           =   1080
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            Caption         =   "ҽѧ��ʾ"
            Height          =   180
            Index           =   25
            Left            =   0
            TabIndex        =   109
            Top             =   1695
            Width           =   720
         End
      End
      Begin VB.PictureBox picPanel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   3825
         ScaleHeight     =   345
         ScaleWidth      =   3510
         TabIndex        =   96
         Top             =   1260
         Width           =   3510
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            Caption         =   "ҩƷĿ¼"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   0
            TabIndex        =   97
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            Caption         =   "����Դ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   1425
            TabIndex        =   98
            Top             =   0
            Width           =   885
         End
      End
      Begin VB.PictureBox picPanel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   3825
         ScaleHeight     =   345
         ScaleWidth      =   3330
         TabIndex        =   102
         Top             =   2925
         Width           =   3330
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            Caption         =   "��ϱ���"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   1395
            TabIndex        =   104
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   0
            TabIndex        =   103
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.TextBox txtE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDFDFD&
         Height          =   630
         Index           =   27
         Left            =   1515
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Top             =   450
         Width           =   7335
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAller 
         Height          =   1260
         Left            =   1230
         TabIndex        =   85
         Top             =   1665
         Width           =   7980
         _cx             =   14076
         _cy             =   2222
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDockPatiInfo.frx":A901
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
      Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
         Height          =   1260
         Left            =   1305
         TabIndex        =   87
         Top             =   3270
         Width           =   7980
         _cx             =   14076
         _cy             =   2222
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   24
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockPatiInfo.frx":A9B4
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
      Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
         Height          =   960
         Left            =   1215
         TabIndex        =   88
         Top             =   4665
         Width           =   7980
         _cx             =   14076
         _cy             =   1693
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   24
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockPatiInfo.frx":AC70
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
      Begin VB.Line linD 
         Index           =   3
         X1              =   1290
         X2              =   1680
         Y1              =   240
         Y2              =   270
      End
      Begin VB.Line linD 
         Index           =   2
         X1              =   180
         X2              =   390
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Line linD 
         Index           =   1
         X1              =   255
         X2              =   570
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line linD 
         Index           =   0
         X1              =   150
         X2              =   930
         Y1              =   255
         Y2              =   270
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "������ַ"
         Height          =   180
         Index           =   24
         Left            =   2760
         TabIndex        =   138
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   23
         Left            =   0
         TabIndex        =   137
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "������Ϣ"
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
         Index           =   52
         Left            =   4425
         TabIndex        =   120
         Top             =   255
         Width           =   780
      End
      Begin VB.Label lblLinkAdd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������������ժҪ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   1515
         MouseIcon       =   "frmDockPatiInfo.frx":AE62
         MousePointer    =   99  'Custom
         TabIndex        =   116
         ToolTipText     =   "������������������ժҪ�С�"
         Top             =   2985
         Width           =   1800
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "��ϼ�¼"
         Height          =   180
         Index           =   54
         Left            =   360
         TabIndex        =   101
         Top             =   2985
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "������¼"
         Height          =   180
         Index           =   53
         Left            =   495
         TabIndex        =   100
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "����ժҪ"
         Height          =   180
         Index           =   27
         Left            =   495
         TabIndex        =   99
         Top             =   450
         Width           =   720
      End
   End
   Begin VB.VScrollBar vsc 
      Height          =   7935
      Left            =   10905
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   10845
      TabIndex        =   114
      Top             =   5970
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   172883969
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4530
      Index           =   0
      Left            =   900
      ScaleHeight     =   4530
      ScaleWidth      =   9705
      TabIndex        =   0
      Top             =   330
      Width           =   9705
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   55
         Left            =   1605
         MaxLength       =   20
         TabIndex        =   26
         Top             =   3960
         Width           =   2475
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   270
         Index           =   2
         Left            =   1245
         TabIndex        =   16
         Top             =   2375
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Style           =   1
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   270
         Index           =   1
         Left            =   1245
         TabIndex        =   11
         Top             =   1640
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Style           =   1
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   270
         Index           =   0
         Left            =   1245
         TabIndex        =   8
         Top             =   1245
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Items           =   3
         Style           =   1
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   7
         Left            =   8130
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1260
         Width           =   1170
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "��"
         Height          =   220
         Index           =   6
         Left            =   6840
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   1245
         Width           =   240
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "��"
         Height          =   220
         Index           =   4
         Left            =   6840
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   870
         Width           =   240
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "��"
         Height          =   220
         Index           =   1
         Left            =   6840
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   465
         Width           =   240
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   1245
         TabIndex        =   41
         Top             =   450
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   -30
            TabIndex        =   1
            Text            =   "cboEdit"
            Top             =   -30
            Width           =   2505
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   1
         Left            =   4665
         MaxLength       =   30
         TabIndex        =   2
         Top             =   480
         Width           =   2430
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   8130
         TabIndex        =   40
         Top             =   450
         Width           =   1230
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   -30
            Width           =   1200
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   3
         Left            =   1245
         MaxLength       =   20
         TabIndex        =   4
         Top             =   870
         Width           =   2475
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   4
         Left            =   4665
         MaxLength       =   100
         TabIndex        =   5
         Top             =   870
         Width           =   2430
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   8130
         TabIndex        =   37
         Top             =   840
         Width           =   1230
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   -30
            Width           =   1200
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   6
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1260
         Width           =   5850
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   20
         Left            =   1230
         TabIndex        =   35
         Top             =   3570
         Width           =   2535
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   20
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   -30
            Width           =   2520
         End
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "��"
         Height          =   220
         Index           =   8
         Left            =   6840
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   1635
         Width           =   240
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "��"
         Height          =   220
         Index           =   10
         Left            =   6840
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   1995
         Width           =   240
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "��"
         Height          =   220
         Index           =   12
         Left            =   6840
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   2370
         Width           =   240
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   9
         Left            =   8130
         MaxLength       =   6
         TabIndex        =   12
         Top             =   1650
         Width           =   1170
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   11
         Left            =   8130
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2010
         Width           =   1170
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   13
         Left            =   8130
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2385
         Width           =   1170
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   14
         Left            =   1245
         TabIndex        =   31
         Top             =   2775
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   14
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   -30
            Width           =   2505
         End
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   15
         Left            =   4665
         TabIndex        =   30
         Top             =   2775
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   15
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   -30
            Width           =   2490
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   16
         Left            =   8130
         MaxLength       =   6
         TabIndex        =   20
         Top             =   2805
         Width           =   1170
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   17
         Left            =   1245
         TabIndex        =   29
         Top             =   3180
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   17
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   -30
            Width           =   2505
         End
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   18
         Left            =   4665
         TabIndex        =   28
         Top             =   3195
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   18
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   -30
            Width           =   2490
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   19
         Left            =   8130
         MaxLength       =   64
         TabIndex        =   23
         Top             =   3210
         Width           =   1170
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   21
         Left            =   4665
         TabIndex        =   27
         Top             =   3585
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   21
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   -30
            Width           =   2490
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   8
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1650
         Width           =   5850
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   10
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   13
         Top             =   2010
         Width           =   5850
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   12
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   15
         Top             =   2385
         Width           =   5850
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "�໤�����֤��"
         Height          =   180
         Index           =   55
         Left            =   360
         TabIndex        =   139
         Top             =   3960
         Width           =   1260
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "������Ϣ"
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
         Index           =   50
         Left            =   3795
         TabIndex        =   119
         Top             =   255
         Width           =   780
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "չ����ݲ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   480
         MouseIcon       =   "frmDockPatiInfo.frx":C7E4
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   4275
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "��λ�ʱ�"
         Height          =   180
         Index           =   7
         Left            =   7350
         TabIndex        =   92
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblN 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "���֤��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   62
         Top             =   510
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "��ͥ��ַ"
         Height          =   180
         Index           =   12
         Left            =   480
         TabIndex        =   61
         Top             =   2415
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "�໤��"
         Height          =   180
         Index           =   19
         Left            =   7530
         TabIndex        =   60
         Top             =   3210
         Width           =   540
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "����֤��"
         Height          =   180
         Index           =   3
         Left            =   480
         TabIndex        =   59
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "�����ص�"
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   58
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "���ڵ�ַ"
         Height          =   180
         Index           =   8
         Left            =   480
         TabIndex        =   57
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "Rh"
         Height          =   180
         Index           =   21
         Left            =   4410
         TabIndex        =   56
         Top             =   3615
         Width           =   180
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "�Ļ��̶�"
         Height          =   180
         Index           =   2
         Left            =   7350
         TabIndex        =   55
         Top             =   510
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "��λ�绰"
         Height          =   180
         Index           =   11
         Left            =   7350
         TabIndex        =   54
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "��ͥ�绰"
         Height          =   180
         Index           =   13
         Left            =   7350
         TabIndex        =   53
         Top             =   2415
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "Ѫ��"
         Height          =   180
         Index           =   20
         Left            =   810
         TabIndex        =   52
         Top             =   3615
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "����״��"
         Height          =   180
         Index           =   5
         Left            =   7350
         TabIndex        =   51
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   4
         Left            =   4230
         TabIndex        =   50
         Top             =   870
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   4230
         TabIndex        =   49
         Top             =   510
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "�����ʱ�"
         Height          =   180
         Index           =   9
         Left            =   7350
         TabIndex        =   48
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "����״��"
         Height          =   180
         Index           =   14
         Left            =   465
         TabIndex        =   47
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "��λ����"
         Height          =   180
         Index           =   10
         Left            =   480
         TabIndex        =   46
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "��ͥ�ʱ�"
         Height          =   180
         Index           =   16
         Left            =   7350
         TabIndex        =   45
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   15
         Left            =   4230
         TabIndex        =   44
         Top             =   2820
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   18
         Left            =   4230
         TabIndex        =   43
         Top             =   3210
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "ְҵ"
         Height          =   180
         Index           =   17
         Left            =   825
         TabIndex        =   42
         Top             =   3210
         Width           =   360
      End
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   375
      Left            =   90
      TabIndex        =   113
      Top             =   450
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3765
      Index           =   1
      Left            =   810
      ScaleHeight     =   3765
      ScaleWidth      =   9750
      TabIndex        =   63
      Top             =   5280
      Width           =   9750
      Begin VB.PictureBox picOutDoc 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   180
         ScaleHeight     =   3015
         ScaleWidth      =   9195
         TabIndex        =   65
         Top             =   540
         Width           =   9200
         Begin VB.PictureBox picSentence 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   585
            ScaleHeight     =   240
            ScaleWidth      =   1155
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   375
            Visible         =   0   'False
            Width           =   1185
            Begin VB.TextBox txtSentence 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Left            =   0
               TabIndex        =   68
               Top             =   30
               Width           =   930
            End
            Begin VB.Image imgSentence 
               Height          =   210
               Left            =   960
               Picture         =   "frmDockPatiInfo.frx":E166
               ToolTipText     =   "�밴 * �ż�ѡ��"
               Top             =   15
               Width           =   180
            End
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "ȡ��ǩ��(&Q)"
            Height          =   350
            Left            =   6885
            TabIndex        =   77
            Top             =   2400
            Width           =   1200
         End
         Begin VB.PictureBox picPrompt 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4545
            Picture         =   "frmDockPatiInfo.frx":E690
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   66
            Top             =   2078
            Width           =   260
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "ȫ�ı༭(&U)"
            Height          =   350
            Left            =   5595
            TabIndex        =   78
            Top             =   2385
            Width           =   1200
         End
         Begin VB.CommandButton cmdImportEPRDemo 
            Caption         =   "���뷶��(&I)"
            Height          =   350
            Left            =   4290
            TabIndex        =   76
            Top             =   2400
            Width           =   1200
         End
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   3
            Left            =   5580
            TabIndex        =   72
            Top             =   900
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":EA91
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   4
            Left            =   345
            TabIndex        =   74
            Top             =   1785
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":EB2E
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   0
            Left            =   240
            TabIndex        =   69
            Top             =   225
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":EBCB
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   2
            Left            =   435
            TabIndex        =   71
            Top             =   1050
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":EC68
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   1
            Left            =   5250
            TabIndex        =   70
            Top             =   105
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":ED05
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
         Begin VB.Label lblEPRname 
            AutoSize        =   -1  'True
            Caption         =   "(���ﲡ��)"
            Height          =   180
            Left            =   4935
            TabIndex        =   128
            Top             =   1785
            Width           =   900
         End
         Begin VB.Line linDoc 
            BorderColor     =   &H00808080&
            X1              =   795
            X2              =   2040
            Y1              =   2715
            Y2              =   2715
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   4
            Left            =   210
            TabIndex        =   126
            Top             =   1440
            Width           =   360
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "�ֲ�ʷ"
            Height          =   180
            Index           =   1
            Left            =   4605
            TabIndex        =   125
            Top             =   300
            Width           =   540
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   255
            TabIndex        =   124
            Top             =   15
            Width           =   360
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "��ȥʷ"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   123
            Top             =   990
            Width           =   540
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "����ʷ"
            Height          =   180
            Index           =   3
            Left            =   4665
            TabIndex        =   122
            Top             =   990
            Width           =   540
         End
         Begin VB.Label lblDoctor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽ��:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   6705
            TabIndex        =   79
            Top             =   1785
            Width           =   450
         End
         Begin VB.Label lblDoctor 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   7215
            TabIndex        =   75
            Top             =   1785
            Width           =   540
         End
         Begin VB.Label lblTip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���벡��ʱ�� ~ ������ȡ��ѡ��ʾ�ʾ��."
            ForeColor       =   &H00404040&
            Height          =   180
            Left            =   4920
            TabIndex        =   73
            Top             =   2115
            Width           =   3420
         End
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "�������"
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
         Index           =   51
         Left            =   4980
         TabIndex        =   121
         Top             =   285
         Width           =   780
      End
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   11055
      Picture         =   "frmDockPatiInfo.frx":EDA2
      Top             =   1140
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonNew 
      Height          =   240
      Left            =   11055
      Picture         =   "frmDockPatiInfo.frx":155F4
      Top             =   780
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmDockPatiInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event EditFullDoc(ByVal lngEPRFileID As Long, ByVal lngFileID As Long, ByVal strDoctor As String, ByVal strIn As String)    'ȫ�ı༭����
Public Event EPRRefresh() '����ˢ��
Public Event UpdatePatiInfo(ByVal strBirthday As String, ByVal strAge As String, ByVal strSex As String, ByVal strTag As String)  '���֤�ű���������� ���գ����䣬�Ա�strTag ��չ����ժҪ
Public Event UpdatePatiState(ByVal strInfo As String, ByRef strTag As String) '���²�������������
    '������strInfo ��<split>�ָ����ID<split>�Һ�ID<split>���.......... strTag ��չ����
Public Event UpdateDiagInfo(ByVal str����ID As String, ByVal str���ID As String, ByVal strTag As String)  'strTag ��չ����

Private Const DColor = &HEEEEEE, EColor = &HFDFDFD, HColor = &HFFDFDF
Private Const GRD_UNEDITCELL_COLOR = &H8000000B  'δ�༭�ĵ�Ԫ����ɫ������ɫ
Public Event SetEdit()
Private Enum E_ITEM_INDEX
    I���֤�� = 0
    I���� = 1
    I�Ļ��̶� = 2
    I����֤�� = 3
    I���� = 4
    I����״�� = 5
    I�����ص� = 6
    I��λ�ʱ� = 7
    I���ڵ�ַ = 8
    I�����ʱ� = 9
    I��λ���� = 10
    I��λ�绰 = 11
    I��ͥ��ַ = 12
    I��ͥ�绰 = 13
    I����״�� = 14
    I���� = 15
    I��ͥ�ʱ� = 16
    Iְҵ = 17
    I���� = 18
    I�໤�� = 19
    IѪ�� = 20
    IRH = 21
    Iȥ�� = 22
    I����ʱ�� = 23
    I������ַ = 24
    Iҽѧ��ʾ = 25
    I����ҽѧ��ʾ = 26
    I����ժҪ = 27
    
    I������¼ = 53
    I��ϼ�¼ = 54
    I�໤�����֤�� = 55
    
    I���� = 0
    I�ֲ�ʷ = 1
    I��ȥʷ = 2
    I����ʷ = 3
    I���� = 4
    
    I���� = 1

End Enum

Private Enum m_Ctl_ID
    picPanel_������Ϣ = 0
    picPanel_������� = 1
    picPanel_������Ϣ = 2
    picPanel_����Դ = 8
    
    picPanel_��� = 3
    picPanel_������ = 5
    picPanel_���� = 6
    picPanel_�������뷽ʽ = 8
    fraLine_������Ϣ = 0
    fraLine_������� = 1
    fraLine_������Ϣ = 2
     
    optҩƷĿ¼ = 5
    opt����Դ = 4
    opt���� = 1
    opt��� = 0
    opt���� = 3
    opt���� = 2
    
    lbl������� = 50
    lbl���ⲡ�� = 51
    lbl������� = 52
    lbl�������� = 28
End Enum

Private Enum AllerColsIndex
    AI_����ҩ��
    AI_������Ӧ
    AI_����ʱ��
    AI_����Դ����
    AI_ҩ��ID
    AI_������Դ
End Enum

Private Enum Change_State
    CS_ɾ���� = -1
    CS_δ�ı� = 0
    CS_������ = 1
    CS_�滻�� = 2
    CS_������ = 3
End Enum
 
Private Enum PaddType
    PT_�����ص� = 0
    PT_���ڵ�ַ = 1
    PT_��ͥ��ַ = 2
End Enum
 
Private Enum DiagColsIndex
    DI_������� = 0
    DI_���� = 1
    DI_��ϱ��� = 2
    DI_������� = 3
    DI_��ҽ֤�� = 4
    DI_����ʱ�� = 5
    DI_��ע = 6
    DI_ICD���� = 7
    DI_�Ƿ����� = 8
    DI_���� = 9
    DI_Del = 10
    DI_���ID = 11
    DI_����ID = 12
    DI_֤��ID = 13
    DI_ҽ��IDs = 14 '�뵱ǰ��Ϲ�����ҽ��ID��ɵ��ַ�����ҽ��ID���Զ��ŷָ�
    DI_��Ϸ��� = 15
    DI_�̶����� = 16
    DI_����ID = 17
    DI_�����Դ = 18
    DI_�������� = 19
    DI_������� = 20
    DI_֤����� = 21
    DI_��¼���� = 22
    DI_��¼��Ա = 23
End Enum

Private mlng����ID As Long
Private mstr�Һŵ� As String
Private mstr����� As String
Private mlng�Һ�ID As Long
Private mlng����ID As Long '�Һ��е�ִ�п���
Private mstr����״�� As String
Private mlng�����ļ�id As Long
Private mlng����ID As Long
Private mstr������ As String
Private mblnǩ�� As Boolean
Private mlngִ��״̬ As Long
Private mbln�� As Boolean
Private mstr�������� As String
Private mstr���� As String
Private mstr�Ա� As String
Private mstr���� As String
Private mint���� As Integer '��ǰ��������
Private mlng��ͬ��λID As Long
Private mblnEdit��ͬ��λ As Boolean '�Ƿ����޸ĺ�ͬ��λ��Ȩ�ޣ�true �����޸ģ�false�����޸�
Private mclsMipModule As zl9ComLib.clsMipModule
Private mobjKernel As zlPublicAdvice.clsPublicAdvice         '�ٴ����Ĳ���
Private mobjPatient As Object
Private mobjCtl As Object '��ǰ��ؼ�
Private mblnUseEPR As Boolean '�Ƿ��ÿ��õĿ�ݲ���

Private mblnMoved As Boolean '����ҽ��վ����
Private mblnDocInput As Boolean '����ҽ��վ���룬�Ƿ���ʾ��ݲ���
Private mbln��ҽ As Boolean '���� mlng����ID ���ж�
Private mbln¼��ҽ��� As Boolean '������������ҽ������¼����ҽ���
Private mblnEdit As Boolean '�����Ƿ�����޸� false ���Ա༭��true ���ܱ༭
Private mblnChange As Boolean
Private mblnPatiChange As Boolean
Private mblnReturn As Boolean
Private mblnID���� As Boolean '���֤����������ʾ
Private mint������� As Integer '1-������������,2-�����ݿ���ȡ����,3-��ҽ�����˴����ݿ�����
Private mint���� As Integer '���ó��ϣ�0-����ҽ������վ��1-����ҽ���༭����
Private mblnOK As Boolean '�Ƿ����̨�ύ�����ݣ�����������������ݵ��ύ

Private mint���� As Integer
Private mintAllerInput As Integer                     'AllerInput:����������Դ��0-��ҩƷĿ¼���룬1-������Դ����
Private mintDiagInput As Integer                      '0-������ϱ�׼����,1-���ݼ�����������
Private mblnSizeTmp As Boolean
Private mbytSize As Byte '9��С���壬12��������
Private mlngTopVsc As Long

Private mclsZip As zlRichEPR.cZip
Private mclsUnZip As zlRichEPR.cUnzip
Private mrsMainInfo As ADODB.Recordset
Private mrsSecdInfo As ADODB.Recordset
Private mrsPreEditCtl As ADODB.Recordset '��һ���ؼ��༭��Ϣ����ʽ���ؼ�����|�Ƿ�Ϊ����ؼ�|�ؼ���|�ؼ��±�
Private mblnNoSave As Boolean '���������ݣ�trueʱ�����棬false����
Private mblnCboNoClick As Boolean '���������б�Click�¼�
 
Private mstrCtlName As String '��ǰ���༭�Ŀؼ�����
Private mintCtlIndex As Integer '��ǰ���༭�Ŀؼ�������
Private mstrTagAller As String '��¼��Դ���
Private mstrTagDiagXY As String
Private mstrTagDiagZY As String
Private mblnFreeInput As Boolean      '����Ƿ��������ɵ���
Private mdatCurDate As Date
Private mblnChk  As Boolean  '�Ƿ�ִ��chk����¼�
Private mblnStructAdress As Boolean, mblnShowTown As Boolean
Private mblnUpdate As Boolean '�Ƿ���½ṹ����ַ
Private mbln��������� As Boolean

Public Function zlRefresh(ByVal lng����ID As Long, ByVal lng�Һ�id As Long, ByVal blnEdit As Boolean, ByVal blnMoved As Boolean, Optional ByRef objMip As Object, Optional ByVal int���� As Integer) As Boolean
'���ܣ��������У�����ˢ��
    Dim blnTmp As Boolean
    
        mblnUpdate = True
    Call SavePreItem
    mint���� = int����
    mlng����ID = lng����ID
    mlng�Һ�ID = lng�Һ�id
    mblnMoved = blnMoved
    mblnEdit = blnEdit
    If mlng�Һ�ID = 0 Then
        mblnEdit = True
    End If
    If Not objMip Is Nothing Then Set mclsMipModule = objMip
    mblnNoSave = True
    Call ClearPatiInfo
    If lng�Һ�id <> 0 Then
        Call LoadPatiInfo
        Call LoadAllerData
        Call LoadDiagData
        If mint���� = 1 Then
            mblnUseEPR = False
        Else
            mblnUseEPR = CanUseFastEPR
        End If
        If mblnUseEPR Then
            If mblnDocInput Then
                Call LoadDocData
                lblLink.Caption = "�����ݲ���"
            Else
                lblLink.Caption = "չ����ݲ���"
            End If
            lblLink.Visible = True
        Else
            lblLink.Visible = False
            PicPanel(picPanel_�������).Visible = False
            mblnDocInput = False
        End If
        If Not mblnEdit Then mdatCurDate = zlDatabase.Currentdate
    End If
    Call SetFaceEditable(mblnEdit)
    mblnNoSave = False
    Set mobjCtl = Nothing
End Function

Private Sub SetAllerEdit(ByVal blnReadOnly As Boolean)
    vsAller.Editable = IIf(blnReadOnly, flexEDNone, flexEDKbdMouse)
    vsAller.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
    vsAller.BackColorBkg = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
    vsAller.TextMatrix(1, AI_����ҩ��) = IIf(blnReadOnly, "��", "")
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����ҽ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    LockWindowUpdate Me.hwnd
    mbytSize = IIf(bytSize = 0, 9, 12)
    Call zlControl.SetPubFontSize(Me, bytSize)
    Call Grid.SetFontSize(vsAller, mbytSize)
    Call Grid.SetFontSize(vsDiagZY, mbytSize)
    Call Grid.SetFontSize(vsDiagXY, mbytSize)
    Set UCPatiVitalSigns.Font = txtE(I����).Font
    If mblnStructAdress Then
        PatiAddress(PT_�����ص�).Font.Size = mbytSize
        PatiAddress(PT_���ڵ�ַ).Font.Size = mbytSize
        PatiAddress(PT_��ͥ��ַ).Font.Size = mbytSize
    End If
    Call SetRTFEditFontSize
    Call SetCtlPos(0)
    Call SetCtlPos(1)
    Call SetCtlPos(2)
    Call Form_Resize
    LockWindowUpdate 0
    Me.Refresh
End Sub

Private Sub SetFaceEditable(ByVal blnReadOnly As Boolean)
'���ܣ����ݵ�ǰ�Ƿ�ֻ�������ý���Ŀɱ༭����
    Dim strObjName As String
    Dim i As Long
    Dim objControl As Object
    Dim intTmp As Integer
    
    On Error GoTo errH
    
    For Each objControl In Me.Controls
        strObjName = TypeName(objControl)
        If InStr("TextBox;ComboBox;CheckBox;VSFlexGrid;OptionButton;PatiAddress", strObjName) > 0 Then '
            If strObjName = "TextBox" Then
                objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                objControl.Locked = blnReadOnly
            ElseIf strObjName = "PatiAddress" Then
                objControl.ControlLock = blnReadOnly
                objControl.Enabled = Not blnReadOnly
            ElseIf strObjName = "ComboBox" Then
                objControl.Enabled = True
                objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                objControl.Locked = blnReadOnly
            ElseIf strObjName = "CheckBox" Or strObjName = "OptionButton" Then
                'û��Locked����,��Enabledʵ��
                objControl.Enabled = Not blnReadOnly
            ElseIf strObjName = "VSFlexGrid" Then
                'ͬʱע��Ҫ�ڼ�������¼��н���һЩ����
                objControl.Editable = IIf(blnReadOnly, flexEDNone, flexEDKbdMouse)
                objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                objControl.BackColorBkg = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
            End If
        End If
    Next
    PicPanel(picPanel_������Ϣ).Enabled = True
    UCPatiVitalSigns.Enabled = Not blnReadOnly
    picOutDoc.Enabled = Not blnReadOnly
    
    For i = 0 To rtfEdit.UBound
        rtfEdit(i).BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
    Next
     
    lblLinkAdd.ForeColor = IIf(blnReadOnly, &HC0C0C0, &HC00000)
    
    If Not blnReadOnly Then
        If chkNoAller.Value = 1 Then
            Call SetAllerEdit(True)
        End If
        Call SetDocEditable
    End If
    
    intTmp = IIf(optInfo(opt����).Value, 1, 0)
    Call SetInputRoot(0, gint�����Դ, intTmp, optInfo(opt���), optInfo(opt����))
    If gint�����Դ = 1 Then
        If intTmp = 1 Then
            optInfo(opt����).Value = True
            optInfo(opt���).Value = False
        Else
            optInfo(opt����).Value = False
            optInfo(opt���).Value = True
        End If
    End If
    
    intTmp = IIf(optInfo(opt����Դ).Value, 2, 1)
    If Not gobjPass Is Nothing Then
        Call SetInputRoot(2, gint����������Դ, intTmp, optInfo(optҩƷĿ¼), optInfo(opt����Դ))
    Else
        Call SetInputRoot(1, 1, 1, optInfo(optҩƷĿ¼), optInfo(opt����Դ))
    End If
    If gint����������Դ = 0 Then
        If intTmp = 2 Then
            optInfo(optҩƷĿ¼).Value = False
            optInfo(opt����Դ).Value = True
        Else
            optInfo(optҩƷĿ¼).Value = True
            optInfo(opt����Դ).Value = False
        End If
    End If
    
    If blnReadOnly Then
        For i = 0 To 5
            optInfo(i).Enabled = Not blnReadOnly
        Next
    End If
    If Not blnReadOnly Then
        If Not mblnEdit��ͬ��λ And mlng��ͬ��λID <> 0 Then
            txtE(I��λ����).BackColor = vbButtonFace
            txtE(I��λ����).Locked = True
        End If
    End If
    If mint���� = 0 Then
        UCPatiVitalSigns.Visible = False
        lblN(lbl��������).Visible = False
    End If
    Exit Sub
errH:
    If 2 = 1 Then
        Resume
    End If
    err.Clear
End Sub

Private Sub AllerEnterNextCell()
    Dim i As Long, j As Long
    With vsAller
        If .Col = AI_������Ӧ Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = AI_����ҩ��
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            .Col = .Col + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Private Sub SetAllerInput(ByVal lngRow As Long, Optional rsInput As ADODB.Recordset, Optional ByVal strTYTInput As String)
'���ܣ��������ҩ�������
'������strTYTInput=̫Ԫͨ������ҩ�ӿڷ��ص��ַ���
    Dim strSQL As String, curDate As Date
    Dim arrTmp As Variant
    Dim strAllerOld As String, strAllerNew As String
    
    With vsAller
        
        strAllerOld = .Cell(flexcpData, lngRow, AI_����ҩ��) & ";" & .TextMatrix(lngRow, AI_����Դ����)
        If Not gobjPass Is Nothing Then
            If optInfo(opt����Դ).Value Then
                arrTmp = Split(strTYTInput, ";")
                
                If UBound(arrTmp) < 1 Then Exit Sub
                If strAllerOld <> strTYTInput Or Val(.RowData(lngRow) & "") <> 0 Then
                    .TextMatrix(lngRow, AI_����ҩ��) = arrTmp(1)
                    .TextMatrix(lngRow, AI_����Դ����) = arrTmp(0)
                    .RowData(lngRow) = 0
                End If
            Else
                If Not rsInput Is Nothing Then
                    .RowData(lngRow) = CLng(rsInput!ID)
                    .TextMatrix(lngRow, AI_����ҩ��) = NVL(rsInput!����)
                Else
                    .RowData(lngRow) = 0
                    .TextMatrix(lngRow, AI_����ҩ��) = .EditText
                End If
                
                strAllerNew = .TextMatrix(lngRow, AI_����ҩ��) & ";" & .TextMatrix(lngRow, AI_����Դ����)
                
                If strAllerOld <> strAllerNew Or Val(.RowData(lngRow) & "") <> 0 Then
                    .TextMatrix(lngRow, AI_����Դ����) = ""
                End If
            End If
        Else
            If optInfo(optҩƷĿ¼).Value Then
                If Not rsInput Is Nothing Then
                    .RowData(lngRow) = CLng(rsInput!ID)
                    .TextMatrix(lngRow, AI_����ҩ��) = NVL(rsInput!����)
                Else
                    .RowData(lngRow) = 0
                    .TextMatrix(lngRow, AI_����ҩ��) = .EditText
                End If
                
                strAllerNew = .TextMatrix(lngRow, AI_����ҩ��) & ";" & .TextMatrix(lngRow, AI_����Դ����)
                
                If strAllerOld <> strAllerNew Or Val(.RowData(lngRow) & "") <> 0 Then
                    .TextMatrix(lngRow, AI_����Դ����) = ""
                End If
            Else
                If Not rsInput Is Nothing Then
                    .TextMatrix(lngRow, AI_����ҩ��) = rsInput!���� & ""
                    .TextMatrix(lngRow, AI_����Դ����) = rsInput!���� & ""
                    .RowData(lngRow) = 0
                Else
                    .RowData(lngRow) = 0
                    .TextMatrix(lngRow, AI_����ҩ��) = .EditText
                End If
            End If
        End If
        .Cell(flexcpData, lngRow, AI_����ҩ��) = .TextMatrix(lngRow, AI_����ҩ��)
        .TextMatrix(lngRow, AI_ҩ��ID) = Val(.RowData(lngRow) & "")

        If .TextMatrix(lngRow, AI_����ʱ��) = "" Then
            .TextMatrix(lngRow, AI_����ʱ��) = Format(mdatCurDate, "YYYY-MM-DD")
        End If
        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
    End With
End Sub

Private Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True, Optional ByVal strMintime As String, Optional strMaxtTime As String) As String
'���ܣ�������������ڼ�,�������������ڴ�(yyyy-MM-dd[ HH:mm])
'������blnTime=�Ƿ���ʱ�䲿��
'������strMintime=����ʱ�������
'          strOutTime=����ʱ�������
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = mdatCurDate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '���봮�а������ڷָ���
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                'ֻ���������ڲ���
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                'ֻ������ʱ�䲿��
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '����Ƿ�����,����ԭ����
            strTmp = strText
        End If
    Else
        '���������ڷָ���
        If Len(strTmp) <= 2 Then
            '��������dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '��������MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '��������yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '��������MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '��������yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '��������yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) Then
        If strMintime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(strMintime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMintime
            End If
        End If
        If strMaxtTime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") > Format(strMaxtTime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMaxtTime
            End If
        End If
        If Not blnTime Then
            strTmp = Format(strTmp, "yyyy-MM-dd")
        End If
    End If
    GetFullDate = strTmp
End Function

Private Sub InitEditData()
'���ܣ���ʼ���༭�����ͱ�Ҫ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    Call zlControl.CboSetHeight(cboE(I����), cboE(I����).Height * 16)
    Call zlControl.CboSetHeight(cboE(I����), cboE(I����).Height * 16)
    Call zlControl.CboSetHeight(cboE(Iְҵ), cboE(Iְҵ).Height * 16)
    
    vsDiagXY.MergeCol(0) = True
    vsDiagZY.MergeCol(0) = True
 
    strSQL = _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, 'ְҵ' ���� From ְҵ Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '����' ���� From ���� Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '����' ���� From ���� Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, 'Ѫ��' ���� From Ѫ�� Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, 'ѧ��' ���� From ѧ�� Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '����ȥ��' ���� From ����ȥ�� Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '����״��' ���� From ����״�� Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '���֤δ¼ԭ��' ���� From ���֤δ¼ԭ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call SetCboFromRec(rsTmp, Array("ְҵ", "����", "����", "Ѫ��", "ѧ��", "����ȥ��", "����״��", "���֤δ¼ԭ��"), Array(Iְҵ, I����, I����, IѪ��, I�Ļ��̶�, Iȥ��, I����״��, I���֤��))
    Call SetCboFromList(Array("0-δ����", "1-����1̥", "2-����2̥������", "4-����"), Array(I����״��))
    Call SetCboFromList(Array("0-δ��", "1-��", "2-��", "3-����"), Array(IRH))
    Call SetCboFromList(Array(" ", "Сʱǰ", "��ǰ", "��ǰ", "��ǰ", "��ǰ"), Array(I����), 2)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetCboFromRec(ByVal rsTmp As ADODB.Recordset, ByVal arrTab As Variant, ByVal arrCboIdx As Variant)
    Dim i As Long, j As Long
    Dim objCboTmp As ComboBox
     
    For i = 0 To UBound(arrTab)
        rsTmp.Filter = "����='" & arrTab(i) & "'"
        If Not rsTmp.EOF Then
            rsTmp.Sort = "����,ID"
            Set objCboTmp = cboE(arrCboIdx(i))
                objCboTmp.Clear
                
            For j = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!����) Then
                    objCboTmp.AddItem rsTmp!����
                Else
                    objCboTmp.AddItem rsTmp!���� & "-" & rsTmp!����
                End If
                objCboTmp.ItemData(objCboTmp.NewIndex) = NVL(rsTmp!ID, 0)
                If Val(rsTmp!ȱʡ & "") = 1 Then
                    Call zlControl.CboSetIndex(objCboTmp.hwnd, objCboTmp.NewIndex)
                    objCboTmp.Tag = objCboTmp.NewIndex
                End If
                rsTmp.MoveNext
            Next
        End If
    Next
End Sub

Private Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'���ܣ���ָ������װ��ָ��ComboBox
'������arrList=List String����
'      arrCboIdx=ComboBox��������,���ComboBoxʱ,װ��������ͬ
'      intDefaut=ȱʡ����
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arrCboIdx)
        cboE(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            cboE(arrCboIdx(i)).AddItem arrList(j)
        Next
        cboE(arrCboIdx(i)).ListIndex = intDefault 'ȱʡΪδѡ��
    Next
End Sub


Private Sub ClearPatiInfo()
    Dim i As Long
  
    mrsMainInfo.Filter = 0
    mrsMainInfo.MoveFirst
    mblnNoSave = True
    For i = 1 To mrsMainInfo.RecordCount
        If mrsMainInfo!�ؼ��� = "txtE" Then
            txtE(mrsMainInfo!Index).Text = ""
        ElseIf mrsMainInfo!�ؼ��� = "PatiAddress" Then
            PatiAddress(mrsMainInfo!Index).Tag = ""
            PatiAddress(mrsMainInfo!Index).Value = ""
        ElseIf mrsMainInfo!�ؼ��� = "cboE" Then
            cboE(mrsMainInfo!Index).ListIndex = -1
            If cboE(mrsMainInfo!Index).Style = 0 Then
                cboE(mrsMainInfo!Index).Text = ""
            End If
        End If
        mrsMainInfo!��Ϣԭֵ = Null
        mrsMainInfo.Update
        mrsMainInfo.MoveNext
    Next
    txtSL.Text = ""
    vsAller.Rows = vsAller.FixedRows
    vsAller.AddItem ""
    vsDiagXY.Rows = vsAller.FixedRows
    vsDiagXY.AddItem ""
    vsDiagZY.Rows = vsAller.FixedRows
    vsDiagZY.AddItem ""
    '1��������¼��
    '2��������¼¼�뷽ʽ
    '3����ϼ�¼  ���������ҽ�����ؼ�������ʾ����
    '4�����¼�뷽ʽ
    UCPatiVitalSigns.ClearData
    For i = I���� To I����
        rtfEdit(i).Text = ""
    Next
    mblnNoSave = False
End Sub

Private Sub InitBaseInfo()
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant

    '��ҳ��׼�ı䣬��ʼ����¼��
    '1������¼�ṹ����
    Set mrsMainInfo = New ADODB.Recordset
    With mrsMainInfo
        .Fields.Append "���", adInteger, , adFldKeyColumn              '��������ʶ��Ϣ
        .Fields.Append "��Ϣ��", adVarChar, 100, adFldKeyColumn   '��Ϣ����
        '�ü�¼������¼һ����Ϣ��Ӧһ���ؼ������������Ϣ��Ӧһ���ؼ��������������д
        .Fields.Append "�ؼ���", adVarChar, 100, adFldIsNullable      'չʾ��Ϣ�Ŀؼ�����
        .Fields.Append "Index", adInteger, , adFldIsNullable                'Ϊ��ʱ��ʾ���ǿؼ�����
        .Fields.Append "ExpState", adInteger                                        '��Ϣ��չ״̬��0-����չ��1-��ʼ��չ��2-������չ
        .Fields.Append "ҳ��", adInteger                                                '��Ϣ���ڵ�ҳ��
        .Fields.Append "��Ϣԭֵ", adVarChar, 2000, adFldIsNullable  '��Ϣ����ҳ����ʱ��ֵ
        .Fields.Append "��Ϣ��ֵ", adVarChar, 2000, adFldIsNullable  '��Ϣ����ҳ���ʱ��ֵ
        .Fields.Append "ErrInfo", adVarChar, 4000, adFldIsNullable  '�ؼ�¼����Ϣ���Ϸ���ʾ��Ϣ��
        .Fields.Append "Edit", adInteger                                                 '0-�ɱ༭,1-���ɱ༭��ֻ����չʾ,2-���ɱ༭������
        .Fields.Append "�Ƿ�ı�", adInteger                                          '��Ϣ�Ƿ��иı�0-δ�ı䣬1-�ı���
        .Fields.Append "��Դ", adInteger  '��Ϣ���������ű�0��������Ϣ��1��������Ϣ�ӱ�,-1����ı�
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    '2���μ���Ϣ��¼���ṹ����
    Set mrsSecdInfo = New ADODB.Recordset
    With mrsSecdInfo
        .Fields.Append "Sort", adInteger                                              '����¼��������
        .Fields.Append "���", adInteger                                              '��ʶ��Ϣ����������¼��
        .Fields.Append "�ؼ���", adVarChar, 100                                       'չʾ��Ϣ�Ŀؼ�����
        .Fields.Append "IndexEx", adInteger, , adFldIsNullable               '�кŻ�ؼ�����Index
        .Fields.Append "ҳ��", adInteger                                                     '��Ϣ���ڵ�ҳ��
        .Fields.Append "ԭID", adBigInt, , adFldIsNullable
        .Fields.Append "��Ϣԭֵ", adVarChar, 2000, adFldIsNullable      '��Ϣ����ҳ����ʱ��ֵ
        .Fields.Append "����Ϣԭֵ", adVarChar, 2000, adFldIsNullable    '��Ϣ����Ҫ���֣���ʶһ����Ϣ�Ƿ񱻳��׸ı䣬��Ϣ����ҳ����ʱ��ֵ
        .Fields.Append "��ID", adBigInt, , adFldIsNullable
        .Fields.Append "��Ϣ��ֵ", adVarChar, 2000, adFldIsNullable      '��Ϣ����ҳ���ʱ��ֵ
        .Fields.Append "����Ϣ��ֵ", adVarChar, 2000, adFldIsNullable    '��Ϣ����ҳ���ʱ��ֵ
        .Fields.Append "��ҽ���", adVarChar, 2000, adFldIsNullable    '��Ϣ����ҳ���ʱ��ֵ ͬһ����ϣ���ͬ����ܳ���3����¼
        .Fields.Append "Edit", adInteger                                                      '0-�ɱ༭,1-���ɱ༭��ֻ����չʾ,2-���ɱ༭������
        .Fields.Append "�ı�״̬", adInteger                                              '��Ϣ�ı�̶�0-δ�ı䣬1-�μ���Ϣ�ı䣬2-����Ϣ�ı�,3-����,-1��ɾ��
        .Fields.Append "ID", adBigInt, , adFldIsNullable                             '��Ϣ�������ݿ��е�ID,һ������ؼ�ʹ��
        .Fields.Append "Tag", adVarChar, 2000                                           '�洢��������
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    With mrsMainInfo
        arrMainFileds = Array("��Ϣ��", "�ؼ���", "Index", "��Դ")
        '������Ϣҳ
        .AddNew arrMainFileds, Array("����", "cboE", I����, 0)
        .AddNew arrMainFileds, Array("����", "cboE", I����, 0)
        .AddNew arrMainFileds, Array("����״��", "cboE", I����״��, 0)
        .AddNew arrMainFileds, Array("ְҵ", "cboE", Iְҵ, 0)
        .AddNew arrMainFileds, Array("���֤��", "��", -1, 0)
        .AddNew arrMainFileds, Array("���֤��״̬", "cboE", I���֤��, 0)
        .AddNew arrMainFileds, Array("�⼮���֤��", "cboE", I���֤��, 0)
        .AddNew arrMainFileds, Array("����", "txtE", I����, 0)
        .AddNew arrMainFileds, Array("����֤��", "txtE", I����֤��, 0)
        .AddNew arrMainFileds, Array("����", "txtE", I����, 0)
        .AddNew arrMainFileds, Array("�����ص�", "txtE", I�����ص�, 0)
        .AddNew arrMainFileds, Array("��λ����", "txtE", I��λ����, 0) '������ʾ�����ݿ��ֶβ�һ���������ֶ�Ϊ  ������λ
        .AddNew arrMainFileds, Array("��λ�绰", "txtE", I��λ�绰, 0)
        .AddNew arrMainFileds, Array("��λ�ʱ�", "txtE", I��λ�ʱ�, 0)
        .AddNew arrMainFileds, Array("��ͥ��ַ", "txtE", I��ͥ��ַ, 0)
        .AddNew arrMainFileds, Array("��ͥ�绰", "txtE", I��ͥ�绰, 0)
        .AddNew arrMainFileds, Array("��ͥ��ַ�ʱ�", "txtE", I��ͥ�ʱ�, 0)
        .AddNew arrMainFileds, Array("���ڵ�ַ", "txtE", I���ڵ�ַ, 0)
        .AddNew arrMainFileds, Array("���ڵ�ַ�ʱ�", "txtE", I�����ʱ�, 0)
        .AddNew arrMainFileds, Array("�໤��", "txtE", I�໤��, 0)
        .AddNew arrMainFileds, Array("�Ļ��̶�", "cboE", I�Ļ��̶�, 1)
        .AddNew arrMainFileds, Array("����״��", "cboE", I����״��, 1)
        .AddNew arrMainFileds, Array("Ѫ��", "cboE", IѪ��, 1)
        .AddNew arrMainFileds, Array("RH", "cboE", IRH, 1)
        .AddNew arrMainFileds, Array("ժҪ", "txtE", I����ժҪ, 0)
        .AddNew arrMainFileds, Array("��Ⱦ���ϴ�", "chkInfo", Null, 0)
        .AddNew arrMainFileds, Array("ȥ��", "cboE", Iȥ��, 1)
        .AddNew arrMainFileds, Array("������ַ", "txtE", I������ַ, 0)
        .AddNew arrMainFileds, Array("����ʱ��", "txtE", I����ʱ��, 0)
        .AddNew arrMainFileds, Array("ҽѧ��ʾ", "txtE", Iҽѧ��ʾ, 1)
        .AddNew arrMainFileds, Array("����ҽѧ��ʾ", "txtE", I����ҽѧ��ʾ, 1)
        .AddNew arrMainFileds, Array("�޹�����¼", "chkNoAller", Null, 1)
        .AddNew arrMainFileds, Array("�໤�����֤��", "txtE", I�໤�����֤��, 1)
         
        If mblnStructAdress Then
            .AddNew arrMainFileds, Array("�����ص�ṹ��", "PatiAddress", PT_�����ص�, 0)
            .AddNew arrMainFileds, Array("���ڵ�ַ�ṹ��", "PatiAddress", PT_���ڵ�ַ, 0)
            .AddNew arrMainFileds, Array("��ͥ��ַ�ṹ��", "PatiAddress", PT_��ͥ��ַ, 0)
        End If
        
        .AddNew arrMainFileds, Array("��������", "UCPatiVitalSigns", Null, -1)
        .AddNew arrMainFileds, Array("����", "rtfEdit", I����, -1)
        .AddNew arrMainFileds, Array("�ֲ�ʷ", "rtfEdit", I�ֲ�ʷ, -1)
        .AddNew arrMainFileds, Array("��ȥʷ", "rtfEdit", I��ȥʷ, -1)
        .AddNew arrMainFileds, Array("����ʷ", "rtfEdit", I����ʷ, -1)
        .AddNew arrMainFileds, Array("����", "rtfEdit", I����, -1)
        
    End With
    
    Set mrsPreEditCtl = New ADODB.Recordset
    With mrsPreEditCtl
        .Fields.Append "�ؼ�����", adVarChar, 100              '��TextBox,Combox
        .Fields.Append "�ؼ���", adVarChar, 100
        .Fields.Append "Index", adInteger, , adFldIsNullable   '�кŻ�ؼ�����Index��������ǿؼ�����Ϊ��1
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
        .AddNew Array("�ؼ�����", "�ؼ���", "Index"), Array("", "", -1)
        .MoveFirst
    End With
End Sub

Private Sub SetCurCtlInfo(ByVal strType As String, ByVal strName As String, Optional ByVal Index As Integer = -1)
'���ܣ���¼��ǰ�ؼ�
    '�ڼ�¼��ǰ�ؼ�֮ǰ������һ���ؼ���ֵ
    Call SavePreItem(1)
    Call mrsPreEditCtl.Update(Array("�ؼ�����", "�ؼ���", "Index"), Array(strType, strName, Index))
    If strName = "txtE" Then
        Set mobjCtl = txtE(Index)
'    ElseIf strName = "PatiAddress" Then
'        Set mobjCtl = PatiAddress(Index)
    ElseIf strName = "cboE" Then
        Set mobjCtl = cboE(Index)
    Else
        Set mobjCtl = Nothing
    End If
End Sub

Private Sub SavePreItem(Optional intTpye As Integer = 0)
'���ܣ���������
    Dim strType As String
    Dim strName As String
    Dim Index As Integer
    Dim blnDo As Boolean
    Dim objCtl As Object
    Dim strValue As String
    Dim str���֤�� As String
    Dim strMsg As String
    
    If mblnNoSave Then Exit Sub
    If mblnEdit Then Exit Sub
    If intTpye = 0 Then mblnUpdate = True
    If Not mrsPreEditCtl Is Nothing Then
        If mrsPreEditCtl.RecordCount > 0 And Not mrsPreEditCtl.EOF Then
            If mrsPreEditCtl!�ؼ��� & "" <> "" Then
                blnDo = True
            End If
        End If
    End If
    
    If blnDo Then
        strType = mrsPreEditCtl!�ؼ����� & ""
        strName = mrsPreEditCtl!�ؼ��� & ""
        Index = Val(mrsPreEditCtl!Index)
        Select Case strType
        Case "TextBox"
            If strName = "txtE" And Not txtE(Index).Locked Then
                If Index <> I�໤�����֤�� Then
                    Call UpDateInfo(txtE(Index).Text, "txtE", Index)
                End If
            End If
        Case "CheckBox"
            If "chkNoAller" = strName Then
                 Call UpDateInfo(chkNoAller.Value, "chkNoAller")
            Else
                Call UpDate�Һ���Ϣ("��Ⱦ���ϴ�", chkInfo.Value)
            End If
        Case "PatiAddress"
            Call UpDate�ṹ����ַ(Index)
        Case "OptionButton"
            If Index = opt���� Or Index = opt���� Then
                Call UpDate�Һ���Ϣ("����", IIf(optInfo(opt����).Value, 1, 0))
            End If
        Case "VSFlexGrid"
            If strName = "vsAller" Then
                Call UpDateAller
            ElseIf strName = "vsDiagXY" Then
                Call UpDateDiag(vsDiagXY)
            ElseIf strName = "vsDiagZY" Then
                Call UpDateDiag(vsDiagZY)
            End If
        End Select
        If strName = "UCPatiVitalSigns" Then
            Call UCPatiVitalSigns_Validate(False)
        ElseIf strName = "rtfEdit" Then
            Call rtfEdit_Validate(Index, False)
        End If
    End If
End Sub

Private Sub LoadPatiInfo()
'���ܣ����ز��˻�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim rsOther As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngRow As Long, bln��ҽ As Boolean
    Dim lngidx As Long
    Dim strValue As String
    
    On Error GoTo errH
    
    Call ClearPatiInfo
    
    strSQL = "Select b.NO,B.Id As �Һ�id,A.�����,A.����id,A.�����ص�,A.��������,A.���֤��,Null as ���֤��״̬,Null as �⼮���֤��,A.����֤��, A.ְҵ, A.����, A.����, A.����, A.����,A.����״��," & vbNewLine & _
        "       A.��ͥ��ַ, A.��ͥ�绰, A.��ͥ��ַ�ʱ�, A.�໤��, A.���ڵ�ַ, A.���ڵ�ַ�ʱ�, A.��ͬ��λid, A.������λ as ��λ����, A.��λ�绰, A.��λ�ʱ�, Nvl(A.����, 0) as ����," & vbNewLine & _
        "       Nvl(B.����, A.����) ����, Nvl(B.�Ա�, A.�Ա�) �Ա�, Nvl(B.����, A.����) ����, B.����ʱ��, B.������ַ, B.��Ⱦ���ϴ�,B.����, B.����," & vbNewLine & _
        "       Nvl(Nvl(B.�������id, Decode(B.ת��״̬, 1, B.ת�����id, Null)), B.ִ�в���id) As ����id, B.ժҪ, B.����,b.ִ��״̬,a.��ͬ��λID" & vbNewLine & _
        "From ������Ϣ A, ���˹Һż�¼ B" & vbNewLine & _
        "Where A.����id = B.����id And b.id=[1] And B.��¼����=1 And B.��¼״̬=1"
     If mblnMoved Then
        strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�Һ�ID)
    If rsTmp.EOF Then Exit Sub
    
    mstr�Һŵ� = rsTmp!NO & ""
    mstr����� = rsTmp!����� & ""
    mstr�������� = Format(rsTmp!�������� & "", "yyyy-MM-dd")
    mstr���� = rsTmp!���� & ""
    mstr���� = rsTmp!���� & ""
    mstr�Ա� = rsTmp!�Ա� & ""
    mint���� = Val(rsTmp!���� & "")
    mlng��ͬ��λID = Val(rsTmp!��ͬ��λid & "")
    mlngִ��״̬ = Decode(NVL(rsTmp!ִ��״̬, 0), 0, 0, 2, 1, 1, 2)
    
    strSQL = "Select ��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And (����ID=[2] Or ����ID is Null and instr(',ȥ��,�޹�����¼,',','||��Ϣ��||',')=0) Order by Nvl(����ID,999999999)"
    Set rsOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    
    Call LoadCache(rsTmp, rsOther)
    
    mrsMainInfo.Filter = 0
    mrsMainInfo.MoveFirst
    mblnNoSave = True

    For i = 1 To mrsMainInfo.RecordCount
        lngidx = Val(mrsMainInfo!Index & "")
        strValue = mrsMainInfo!��Ϣԭֵ & ""
        
        If mrsMainInfo!�ؼ��� = "txtE" Then
            txtE(lngidx).Text = strValue
            txtE(lngidx).Tag = ""
            txtE(lngidx).BackColor = vbWindowBackground
        ElseIf mrsMainInfo!�ؼ��� = "PatiAddress" Then
            Call zlReadAddrInfo(PatiAddress(Val(mrsMainInfo!Index)), Val(NVL(mlng����ID)), 0, Decode(Val(mrsMainInfo!Index), PT_�����ص�, 1, PT_���ڵ�ַ, 4, PT_��ͥ��ַ, 3), strValue)
            If mrsMainInfo!��Ϣԭֵ <> strValue Then
                mrsMainInfo!��Ϣԭֵ = PatiAddress(Val(mrsMainInfo!Index)).Value
                mrsMainInfo.Update
            End If
        ElseIf mrsMainInfo!�ؼ��� = "cboE" Then
            If lngidx <> I���֤�� Then
                If Mid(strValue, 1, 1) = Chr(30) Then
                    strValue = zlStr.TrimEx(strValue, Chr(30))
                End If
                Call GetCboIndex(cboE(lngidx), strValue)
            End If
        End If
        mrsMainInfo.MoveNext
    Next
    
    mblnReturn = True
    '*���֤�ŵ�������
    lngidx = I���֤��
    mrsMainInfo.Filter = "�ؼ���='��' and Index=-1 and ��Ϣ��='���֤��'"
    If Not mrsMainInfo.EOF Then
        strValue = mrsMainInfo!��Ϣԭֵ & ""
        If Mid(strValue, 1, 1) = Chr(30) Then
            strValue = zlStr.TrimEx(strValue, Chr(30))
        End If
        Call GetCboIndex(cboE(lngidx), strValue)
        If cboE(lngidx).ListIndex = -1 And strValue <> "" Then
            cboE(lngidx).Tag = strValue
            If mblnID���� Then
                strValue = Mid(strValue, 1, 12) & String(Len(Mid(strValue, 13, 2)), "*") & Mid(strValue, 15)
            End If
            cboE(lngidx).Text = strValue
        End If
    End If
    If cboE(lngidx).Text = "" Then
        mrsMainInfo.Filter = "�ؼ���='cboE' and Index=" & lngidx & " and ��Ϣ��='���֤��״̬'"
        If Not mrsMainInfo.EOF Then
            strValue = mrsMainInfo!��Ϣԭֵ & ""
            If Mid(strValue, 1, 1) = Chr(30) Then
                strValue = zlStr.TrimEx(strValue, Chr(30))
            End If
            Call GetCboIndex(cboE(lngidx), strValue)
        End If
        If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) <> "�й�" Then
            If cboE(lngidx).Text = "" Then
                mrsMainInfo.Filter = "�ؼ���='cboE' and Index=" & lngidx & " and ��Ϣ��='�⼮���֤��'"
                If Not mrsMainInfo.EOF Then
                    strValue = mrsMainInfo!��Ϣԭֵ & ""
                    cboE(lngidx).Tag = strValue
                    cboE(lngidx).Text = strValue
                End If
            End If
        End If
    End If
    mblnReturn = False
    '1��������¼��
    '2��������¼¼�뷽ʽ
    '3����ϼ�¼  ���������ҽ�����ؼ�������ʾ����
    '4�����¼�뷽ʽ
    
    optInfo(opt����).Value = Val(rsTmp!���� & "") = 1
    optInfo(opt����).Value = Val(rsTmp!���� & "") = 0
    
    chkInfo.Value = Val(rsTmp!��Ⱦ���ϴ� & "")
    
    If mint���� = 1 Then
        Call UCPatiVitalSigns.LoadPatiVitalSigns(mlng����ID, mlng�Һ�ID)
        strSQL = UCPatiVitalSigns.GetSaveSQL(mlng����ID, mlng�Һ�ID)
        UCPatiVitalSigns.Visible = True
        lblN(lbl��������).Visible = True
    Else
        strSQL = ""
        UCPatiVitalSigns.Visible = False
        lblN(lbl��������).Visible = False
    End If
    
    mrsMainInfo.Filter = "�ؼ���='UCPatiVitalSigns'"
    mrsMainInfo!��Ϣԭֵ = strSQL
    mrsMainInfo.Update
    
    mbln�� = Val(rsTmp!���� & "") = 1
    mlng����ID = Val(rsTmp!����ID & "")
    
    mbln��ҽ = Sys.DeptHaveProperty(mlng����ID, "��ҽ��")
    mbln��ҽ = mbln��ҽ Or mbln¼��ҽ���
    mblnNoSave = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCache(ByVal rsMain As ADODB.Recordset, ByVal rsSecond As ADODB.Recordset)
'���ܣ���ʼ��������Ϣ
    Dim i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    mrsMainInfo.Filter = "��Դ=0"
    For i = 1 To mrsMainInfo.RecordCount
        If mrsMainInfo!�ؼ��� <> "PatiAddress" Then
            mrsMainInfo!��Ϣԭֵ = rsMain(CStr(mrsMainInfo!��Ϣ��)).Value
        Else
            mrsMainInfo!��Ϣԭֵ = rsMain(Replace(CStr(mrsMainInfo!��Ϣ��), "�ṹ��", "")).Value
        End If
        mrsMainInfo.Update
        mrsMainInfo.MoveNext
    Next
    
    mrsMainInfo.Filter = "��Դ=1"
    For i = 1 To mrsMainInfo.RecordCount
        rsSecond.Filter = "��Ϣ��='" & mrsMainInfo!��Ϣ�� & "'"
        If Not rsSecond.EOF Then
            mrsMainInfo!��Ϣԭֵ = NVL(rsSecond!��Ϣֵ)
            mrsMainInfo.Update
        End If
        mrsMainInfo.MoveNext
    Next
    
    mrsMainInfo.Filter = "�ؼ���='txtE' and Index=" & I����ʱ��
    If Not mrsMainInfo.EOF Then
        strTmp = rsMain(CStr(mrsMainInfo!��Ϣ��)).Value & ""
        If strTmp <> "" Then
            mrsMainInfo!��Ϣԭֵ = Format(rsMain(CStr(mrsMainInfo!��Ϣ��)).Value, "yyyy-MM-dd HH:MM")
            mrsMainInfo.Update
        End If
    End If
    
    mrsMainInfo.Filter = "�ؼ���='cboE' and Index=" & I���֤�� & " and ��Ϣ��='���֤��״̬' or ��Ϣ��='�⼮���֤��'"
    If Not mrsMainInfo.EOF Then
        rsSecond.Filter = "��Ϣ��='���֤��״̬'"
        If Not rsSecond.EOF Then
            mrsMainInfo!��Ϣԭֵ = NVL(rsSecond!��Ϣֵ)
            mrsMainInfo.Update
        End If
    End If
     
    mrsMainInfo.Filter = "�ؼ���='cboE' and Index=" & I���֤�� & " and  ��Ϣ��='�⼮���֤��'"
    If Not mrsMainInfo.EOF Then
        rsSecond.Filter = "��Ϣ��='�⼮���֤��'"
        If Not rsSecond.EOF Then
            mrsMainInfo!��Ϣԭֵ = NVL(rsSecond!��Ϣֵ)
            mrsMainInfo.Update
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboE_Change(Index As Integer)
    If Index = I���֤�� Then
        Call cboSpecificInfoChange(Index)
    End If
End Sub

Private Sub chkInfo_Click()
    Call UpDate�Һ���Ϣ("��Ⱦ���ϴ�", chkInfo.Value)
End Sub

Private Function SaveRTFData(ByVal lng����ID As Long, Optional blnSign As Boolean) As Boolean
'���ܣ����没�˲�����ʽRTF����
'������
    Dim strZipFile As String, strTempFile As String, i As Long
    Dim bFinded As Boolean, lngStartPos As Long, lngEndPos As Long, arrTmp As Variant
    Dim strContent As String, lngRecID As Long
    
    If mlng����ID = 0 Then
        lngRecID = lng����ID
    Else
        lngRecID = mlng����ID
    End If
    
    If blnSign = False Then
        '�滻�������
        edtEditor.Freeze
        edtEditor.ForceEdit = True
        
        For i = 0 To lblDoc.UBound
            bFinded = FindOutLinePosition(edtEditor, CStr(lblDoc(i).Tag), lngStartPos, lngEndPos)
            If bFinded Then
                strContent = rtfEdit(i).Text    'ȥ��β���Ļس�����
                Do While Len(strContent) > 2
                    If Mid(strContent, Len(strContent) - 1) = vbLf Or Mid(strContent, Len(strContent) - 1) = vbCr Then
                        strContent = Mid(strContent, Len(strContent) - 1)
                    Else
                        Exit Do
                    End If
                Loop
                edtEditor.Range(lngStartPos, lngEndPos).Text = strContent
            End If
        Next
        
        edtEditor.UnFreeze
        edtEditor.ForceEdit = False
        'Ҫ�����ݸ���
        If mlng����ID = 0 Then Call ElementsUpdate(lngRecID)
    End If
    
    On Error GoTo errH
    strTempFile = App.Path & "\TMP.rtf"
    If Dir(strTempFile) <> "" Then Kill strTempFile
    edtEditor.SaveDoc strTempFile
    'ѹ���ļ�
    strZipFile = zlFileZip(strTempFile)
    '�����ʽ
    Call Sys.SaveLob(glngSys, 5, lngRecID, strZipFile)
    
    'ɾ����ʱ�ļ�
    If strTempFile <> "" Then Kill strTempFile
    If strZipFile <> "" Then Kill strZipFile

    SaveRTFData = True
    Exit Function
errH:
    SaveRTFData = False
End Function

Private Function ElementsUpdate(ByVal lng����ID As Long) As Boolean
'���ܣ�����Editor�ؼ��е��滻Ҫ�����ݣ��Ա㱣��ΪRTF�ļ�
    Dim ThisElements As New zlRichEPR.cEPRElements
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, lngKey As Long
    Dim bFinded As Boolean, bNeeded As Boolean, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long

    strSQL = "Select ������,ID From ���Ӳ������� Where �ļ�ID= [1] And �������� = 4 And ��ֹ��=0 and �������� =0 And �滻�� =1 order by ������ "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    For i = 1 To rsTmp.RecordCount
        lngKey = ThisElements.Add(NVL(rsTmp("������"), 0))
        ThisElements("K" & lngKey).GetElementFromDB cprET_�������༭, rsTmp("ID"), True
        rsTmp.MoveNext
    Next

     For i = 1 To ThisElements.Count
        If ThisElements(i).�滻�� = 1 Then
            ThisElements(i).�����ı� = GetReplaceEleValue(ThisElements(i).Ҫ������, mlng����ID, mlng�Һ�ID, 1, 0)
            bFinded = FindNextKey(edtEditor, 0, "E", ThisElements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            ThisElements(i).Refresh edtEditor
        End If
        If ThisElements(i).�滻�� = 1 And ThisElements(i).�Զ�ת�ı� Then
            EleToString edtEditor, ThisElements(i)     '�Զ�ת��Ϊ���ı�����ʱ��ɾ����Ҫ�أ�
        End If
    Next
    Set ThisElements = Nothing
End Function

Private Sub EleToString(ByRef edtThis As Object, Ele As cEPRElement)
    Dim sKeyType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bNeeded As Boolean, bBeteenKeys As Boolean
    Dim bForce As Boolean, strOldTag As String
    
    bBeteenKeys = FindNextKey(edtThis, 0, "E", Ele.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bBeteenKeys Then
        Dim lngLen As Long, str���� As String
        str���� = Ele.�����ı�
        lngLen = Len(str����)
        With edtThis
            .Freeze
            strOldTag = .Tag
            .Tag = "EleToString"
            bForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = str����
            .Range(lKSS, lKSS + lngLen).Font.Protected = False
            .Range(lKSS, lKSS + lngLen).Font.Hidden = False
            .Range(lKSS, lKSS + lngLen).Font.BackColor = tomAutoColor
            .Range(lKSS, lKSS + lngLen).Font.Underline = cprNone
            .ForceEdit = bForce
            .UnFreeze
            .Tag = strOldTag
        End With
    End If
End Sub

Private Function GetReplaceEleValue(ByVal ElementName As String, _
    ByVal sPatientID As String, _
    ByVal sPageID As String, _
    ByVal iPatientType As PatiFromEnum, _
    ByVal lngҽ��ID As Long) As String

    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5]) From Dual"
    err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�滻��", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lngҽ��ID)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value))
    End If
    Exit Function

DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

Private Function FindOutLinePosition(ByRef edtThis As Object, ByVal strOName As String, ByRef lngS As Long, lngE As Long) As Boolean
'���ܣ�����ָ����������ƣ�������������ı�����ֹλ��
    Dim blnFindedNext As Boolean, lngCur As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, bFinded As Boolean, bNeeded As Boolean
    Dim strTmp As String
    
    bFinded = True
    While bFinded
        bFinded = FindNextKey(edtThis, lngCur, "O", 0, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            strTmp = edtThis.Range(lKEE, lKEE + Len(strOName))
            If strOName = strTmp Then
                lngS = lKEE + Len(strOName)
                blnFindedNext = FindNextAnyKey(edtThis, lngS, strTmp, lKSS, lKSE, lKES, lKEE, 0, bNeeded)
                If blnFindedNext Then
                    lngE = lKSS
                Else
                    lngE = Len(edtThis.Text)
                End If
                Do While lngE > lngS + 1    'ȥ��β���Ļس�����
                    If edtThis.Range(lngE - 1, lngE) = vbLf Or edtThis.Range(lngE - 1, lngE) = vbCr Then
                        lngE = lngE - 1
                    Else
                        Exit Do
                    End If
                Loop
                FindOutLinePosition = True
                Exit Function
            Else
                lngCur = lKEE
            End If
        End If
    Wend
End Function

'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Private Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If strFile = "" Then Exit Function
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    Set mclsZip = New zlRichEPR.cZip
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
    Set mclsZip = Nothing
End Function

Private Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim objFSO As New Scripting.FileSystemObject    'FSO����
    
    Dim strZipPath As String
    If strZipFile = "" Then Exit Function
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If objFSO.FileExists(strZipPath & "TMP.RTF") Then objFSO.DeleteFile strZipPath & "TMP.RTF"
    
    Set mclsUnZip = New zlRichEPR.cUnzip
    With mclsUnZip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
    Set mclsUnZip = Nothing
End Function

Private Function FindNextKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindNextKey = True
            End If
        End If
    End With
End Function

Private Function FindNextAnyKey(ByRef edtThis As Object, _
    ByRef lngCurPosition As Long, _
    ByRef strKeyType As String, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef lngKey As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = "S("
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = .TOM.TextDocument.Range(i - 2, i - 1)
                lngKSS = i - 2 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 14
                lngKES = j - 2
                lngKEE = j + 14
                lngKey = Val(.TOM.TextDocument.Range(i + 1, i + 9))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 10, i + 11))
                FindNextAnyKey = True
            End If
        End If
    End With
End Function
   
Private Sub GetSQLOutDoc(ByRef arrSQL As Variant, ByVal lng����ID As Long)
'���ܣ���֯��ݲ��������ݱ���SQL
'������lng����ID-����ʱ������ȡ�Ĳ���ID
    Dim i As Long, k As Long
    Dim strTmp(5) As String
    
    If mlng����ID = 0 Then
        For i = 0 To rtfEdit.UBound
            If Trim(rtfEdit(i).Text) <> "" Then Exit For
        Next
        If i > rtfEdit.UBound Then Exit Sub     '����ʱ�����û�������ݣ��򲻱���
    End If
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
     
    If mlng����ID = 0 Then
        If rtfEdit(I����).Locked Then Exit Sub
        arrSQL(UBound(arrSQL)) = "Zl_�����ﲡ��_Update(1," & mlng����ID & "," & _
            mlng�Һ�ID & "," & mlng����ID & "," & mlng�����ļ�id & "," & lng����ID & ",'" & UserInfo.���� & "','" & _
            Replace(Trim(rtfEdit(I����).Text), "'", "��") & "','" & Replace(Trim(rtfEdit(I����ʷ).Text), "'", "��") & "','" & Replace(Trim(rtfEdit(I�ֲ�ʷ).Text), "'", "��") & "','" & _
            Replace(Trim(rtfEdit(I����).Text), "'", "��") & "','" & Replace(Trim(rtfEdit(I��ȥʷ).Text), "'", "��") & "')"
    Else
        k = 0
        For i = 0 To rtfEdit.UBound
            If rtfEdit(i).Locked = False Then
                strTmp(i) = rtfEdit(i).Tag & "|" & Replace(Trim(rtfEdit(i).Text), "'", "��")
                k = k + 1
            End If
        Next
        If k = 0 Then Exit Sub
        arrSQL(UBound(arrSQL)) = "Zl_�����ﲡ��_Update(2," & mlng����ID & "," & _
            mlng�Һ�ID & "," & mlng����ID & ",0," & mlng����ID & ",'" & UserInfo.���� & "','" & _
            strTmp(0) & "','" & strTmp(3) & "','" & strTmp(1) & "','" & strTmp(4) & "','" & strTmp(2) & "')"
    End If
End Sub

Private Function GetEPRDoc() As zlRichEPR.cEPRDocument
'���ܣ���ȡ�����ļ���RTF���ݵ�editor�ؼ��У��������ĵ�����
    Dim objDoc As New zlRichEPR.cEPRDocument
   
    objDoc.InitEPRDoc cprEM_�޸�, cprET_�������༭, mlng����ID, cprPF_����, mlng����ID, mlng�Һ�ID, , mlng����ID
    
    If objDoc.ReadFileStructure(edtEditor) = True Then
        Set GetEPRDoc = objDoc
    End If
End Function

Private Sub chkInfo_GotFocus()
    Call SetCurCtlInfo(TypeName(chkInfo), "chkInfo")
End Sub

Private Sub chkInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chkNoAller_Click()
    Dim strValue As String
    If mblnChk = False Then
        strValue = chkNoAller.Value
        If vsAller.TextMatrix(vsAller.FixedRows, AI_����ҩ��) <> "" And vsAller.TextMatrix(vsAller.FixedRows, AI_����ҩ��) <> "��" Then
            If strValue = "1" Then
                MsgBox "�Ѿ��й���ҩ����ܱ��Ϊ�ޡ�", vbInformation, gstrSysName
                mblnChk = True
                chkNoAller.Value = 0
                Exit Sub
            End If
        End If
        Call SetAllerEdit(strValue = "1")
        Call UpDateInfo(strValue, "chkNoAller")
    End If
    mblnChk = False
End Sub

Private Sub cmdDiagMove_Click(Index As Integer)
'���ܣ��ƶ������
    Dim strTmp As String
    Dim i As Long, lngRow As Long
    Dim vsDiag As VSFlexGrid                '��ǰ��ϱ��
    Dim intStep As Integer                  '�ƶ�λ�ã�1-�����ƶ���-1�����ƶ�

    If Index = 0 Or Index = 1 Then
        Set vsDiag = vsDiagXY               '��ҽ
    Else
        Set vsDiag = vsDiagZY               '��ҽ
    End If
    
    If vsDiag.Editable = flexEDNone Then
        Exit Sub
    End If
    
'    intStep=�ƶ�λ�ã�1-�����ƶ���-1�����ƶ�
    If Index = 0 Or Index = 2 Then
        intStep = -1
    ElseIf Index = 1 Or Index = 3 Then
        intStep = 1
    End If
    
    With vsDiag
        If .Row < 0 Then
            Exit Sub
        Else
            If Not DiagRowCanMove(vsDiag, intStep, .Row) Then Exit Sub
            lngRow = IIf(intStep = 1, .Row, .Row + intStep)
        End If
        
        For i = .FixedCols To .Cols - 1
            '������������
            strTmp = .TextMatrix(.Row + intStep, i)
            .TextMatrix(.Row + intStep, i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = strTmp
            '������������
            strTmp = .Cell(flexcpData, .Row + intStep, i)
            .Cell(flexcpData, .Row + intStep, i) = .Cell(flexcpData, .Row, i)
            .Cell(flexcpData, .Row, i) = strTmp
        Next
        
        '������������
        strTmp = .RowData(.Row + intStep)
        .RowData(.Row + intStep) = .RowData(.Row)
        .RowData(.Row) = Val(strTmp)
        Call SetDiagReletedInfo(vsDiag)
        .Row = .Row + intStep
    End With
End Sub

Private Function DiagRowCanMove(ByRef vsDiag As VSFlexGrid, ByVal intStep As Integer, ByVal lngRow As Long) As Boolean
'���ܣ���������ƶ��ؼ�״̬
'������intStep=�ƶ�λ�ã�1-�����ƶ���-1�����ƶ�
'   lngRow=���ж����У�һ��Ϊ��ǰ��
    Dim lngBgn As Long, lngEnd As Long
    Dim i As Long
    
    lngBgn = 0
    lngEnd = 0
    '���ݵ�ǰ�е�λ�������ƶ���Ͽؼ��Ŀ�����
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_�������) <> "" Then
                If lngBgn < .FixedRows Then lngBgn = i
                lngEnd = i
            End If
        Next
    End With
    
    If lngBgn = lngEnd Then 'ֻ��һ����ϣ��򲻿��ƶ�
        DiagRowCanMove = False
    ElseIf lngRow = lngBgn Then '��ǰ���Ǳ������һ�У���ֻ������
        DiagRowCanMove = intStep = 1
    ElseIf lngRow = lngEnd Then '��ǰ���Ǳ��������һ�У���ֻ����
        DiagRowCanMove = intStep = -1
    Else  '��ǰ���Ǳ������м�ĳһ�У�����������ƶ�
        DiagRowCanMove = True
    End If
End Function

Private Sub cmdSaveZY_Click()
    Dim strSQL As String, i As Integer
    Dim rsTmp As Recordset
    Dim Index As Long
    Dim objTxt As Object
    Set objTxt = txtE(I����ժҪ)
    Index = I����ժҪ
    If objTxt.Locked Then Exit Sub
    If Trim(objTxt.Text) = "" Then
        MsgBox "������ժҪ���ݡ�", vbInformation, gstrSysName
        If objTxt.Enabled Then objTxt.SetFocus
        Exit Sub
    End If
    On Error GoTo errH
    strSQL = "Select 1 From ���þ���ժҪ Where ����=[1] And ��ԱID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(objTxt.Text), UserInfo.ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "�������Ѿ��ڳ���ժҪ�С�", vbInformation, gstrSysName
        If objTxt.Enabled Then objTxt.SetFocus
        Exit Sub
    End If
    
    strSQL = zlCommFun.zlGetSymbol(objTxt.Text, CByte(mint����))
    strSQL = "Zl_���þ���ժҪ_Update(0,Null,'" & Replace(objTxt.Text, "'", "''") & "','" & strSQL & "'," & UserInfo.ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    AddComboItem objTxt.hwnd, CB_ADDSTRING, 0, objTxt.Text
    MsgBox "������Ϊ����ժҪ��", vbInformation, gstrSysName
    If objTxt.Enabled Then objTxt.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdShowZY_Click()
    If txtE(I����ժҪ).Locked Then Exit Sub
    If AbstractSelect("") Then Exit Sub
End Sub

Private Sub cmdSign_Click()
    Dim i As Long, str�������� As String, strSource As String, strSQL As String
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim patiSign As cEPRSign, objEPRDoc As cEPRDocument
    
    If mblnǩ�� = False Then
        For i = 0 To rtfEdit.UBound
            If Trim(rtfEdit(i).Text) <> "" Then Exit For
        Next
        If i > rtfEdit.UBound Then
            MsgBox "�������벡����Ϣ���ٽ���ǩ����", vbInformation, gstrSysName
            Exit Sub    '����ʱ�����û�������ݣ��򲻱���
        End If
                                         
        If edtEditor.Text = "" Then
            If ReadRTFData(mlng����ID) = False Then Exit Sub
        End If
        
        strSource = edtEditor.Text
        If cmdSign.Visible And cmdSign.Enabled Then cmdSign.SetFocus
        Set patiSign = frmOutDocterSign.ShowMe(Me, strSource, mlng����ID, mlng�Һ�ID)
        If patiSign Is Nothing Then Exit Sub
        With patiSign
            .Key = "1"
            str�������� = .ǩ����ʽ & ";" & .ǩ������ & ";" & .֤��ID & ";" & IIf(.��ʾ��ǩ, 1, 0) & ";" & _
                    Format(.ǩ��ʱ��, "yyyy-mm-dd hh:mm:ss") & ";" & .��ʾʱ�� & ";" & .ǩ��Ҫ��
                    
            strSQL = "Zl_�����ﲡ��_ǩ��(1," & mlng����ID & ",'" & str�������� & "','" & UserInfo.���� & "','" & _
                    .ǰ������ & "','" & .ʱ��� & "','" & .ǩ������ & "','" & .ǩ����Ϣ & "')"
        End With
        
        Set objEPRDoc = GetEPRDoc()
        If objEPRDoc Is Nothing Then Exit Sub
        Call patiSign.InsertIntoEditor(edtEditor, Len(edtEditor.Text), , objEPRDoc)
        Set objEPRDoc = Nothing
        Set patiSign = Nothing
    Else
        If MsgBox("��ȷ��Ҫȡ��ǩ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        Set patiSign = GetSign(mlng����ID)
        If patiSign Is Nothing Then Exit Sub
        
        Set objEPRDoc = GetEPRDoc()
        If objEPRDoc Is Nothing Then Exit Sub
        Call patiSign.DeleteFromEditor(edtEditor, objEPRDoc)
        Set objEPRDoc = Nothing
        Set patiSign = Nothing
        
        strSQL = "Zl_�����ﲡ��_ǩ��(0," & mlng����ID & ")"
    End If
    
   
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If SaveRTFData(mlng����ID, True) = False Then GoTo errH
    gcnOracle.CommitTrans: blnTrans = False
    
        
    Call LoadDocData
    
    RaiseEvent EPRRefresh
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Set objEPRDoc = Nothing: Set patiSign = Nothing
    Call SaveErrLog
End Sub

Private Sub cmdImportEPRDemo_Click()
    Dim objImportEPRDemo As New frmImportEPRDemo
    Dim rsDemo As New Recordset
    
    If mlng����ID <> 0 Then
        MsgBox "�ò����Ѿ������˲����ļ��������ٵ��뷶�ġ�", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If objImportEPRDemo.ShowMe(Me, mlng�����ļ�id, mlng����ID, mlng�Һ�ID, rsDemo) > 0 Then
    Call SetDocData(rsDemo, 1)
    End If
End Sub

Private Sub SetDocData(ByVal rsTmp As Recordset, ByVal intType As Integer)
'���ܣ����ÿ����������
'������intType=0������ȡ��intType=1���ĵ��룬����ղ���ID
    Dim i As Long, j As Long, arrTmp As Variant
    Dim strContent As String
    
    With rsTmp
        If .RecordCount > 0 Then
            arrTmp = Split("-10,2,3,5,6", ",") '��������,�ֲ�ʷ,����ʷ,����ʷ,�����
            For i = 0 To UBound(arrTmp)
                .Filter = "Ԥ�����id=" & arrTmp(i)
                rtfEdit(i).Text = ""
                If intType = 1 Then
                    '���뷶�ĺ����ж����á�
                    rtfEdit(i).Locked = False
                    rtfEdit(i).BackColor = HColor
                End If
                For j = 1 To .RecordCount
                    If j = 1 Then
                        strContent = "" & !�����ı�
                        If InStr(strContent, lblDoc(i).Tag) = 1 Then strContent = Mid(strContent, Len(lblDoc(i).Tag) + 1)
                        rtfEdit(i).Text = strContent
                        If intType = 0 Then rtfEdit(i).Tag = !ID
                    Else
                        rtfEdit(i).Text = rtfEdit(i).Text & vbCrLf & !�����ı�
                        If intType = 0 Then rtfEdit(i).Tag = rtfEdit(i).Tag & "," & !ID
                    End If
                    .MoveNext
                Next
            Next
        End If
    End With
End Sub

Private Function ReadRTFData(ByVal lng����ID As Long) As Boolean
'���ܣ���ȡ�����ļ���RTF���ݵ�editor�ؼ���
    Dim strZipFile As String, strTempFile As String
    Dim lngRecID As Long
    
    If mlng����ID = 0 Then
        lngRecID = lng����ID
    Else
        lngRecID = mlng����ID
    End If
    
    On Error GoTo errH
        
    '�жϱ����ǲ��������Ĳ���
    If mlng����ID = 0 Then
        strZipFile = Sys.ReadLob(glngSys, 1, mlng�����ļ�id)
    Else
        strZipFile = Sys.ReadLob(glngSys, 5, lngRecID)
    End If
    
    strTempFile = zlFileUnzip(strZipFile)
    edtEditor.OpenDoc strTempFile
     'ɾ����ʱ�ļ�
    If strTempFile <> "" Then Kill strTempFile
    If strZipFile <> "" Then Kill strZipFile
   
    ReadRTFData = True
    Exit Function
errH:
    ReadRTFData = False
End Function

Private Function GetSign(ByVal lng����ID As Long) As cEPRSign
'���ܣ���ȡ��ǰ�û���ǩ������
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim OneSign As New cEPRSign, intSign As Integer, strUserName As String
    
    strUserName = UserInfo.����
    intSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0)
    If intSign = 1 Then
        strSQL = "Select ǩ�� From ��Ա�� Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
        If rsTemp.RecordCount > 0 Then
            If Not IsNull(rsTemp!ǩ��) Then strUserName = rsTemp!ǩ��
        End If
    End If
    strSQL = "Select Id,������ From ���Ӳ������� Where �ļ�id= [1] And ��������=8 And Instr(';'||�����ı�||';',[2])>0 Order By ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, ";" & strUserName & ";")
    If rsTemp.RecordCount > 0 Then
        OneSign.Key = NVL(rsTemp!������, 0)
        If OneSign.GetSignFromDB(rsTemp!ID) = True Then Set GetSign = OneSign
    End If
End Function

Private Sub cmdUpdate_Click()
    RaiseEvent EditFullDoc(mlng�����ļ�id, mlng����ID, mstr������, lblDoctor(1).Tag)
End Sub

Private Sub lblLink_Click()
'���ܣ����ÿ�ݲ����Ŀɼ���
    mblnDocInput = Not mblnDocInput
    Call zlDatabase.SetPara("��ʾ�����������", IIf(mblnDocInput, 1, 0), glngSys, p����ҽ��վ, InStr(";" & gstrPrivs & ";", ";��������;") > 0)
    
    If mblnDocInput Then Call LoadDocData
    PicPanel(picPanel_�������).Visible = mblnDocInput
    If mblnDocInput Then
        lblLink.Caption = "�����ݲ���"
    Else
        lblLink.Caption = "չ����ݲ���"
    End If
    If Not mblnEdit Then Call SetDocEditable
 
    Call Form_Resize
End Sub

Private Sub lblLinkAdd_Click()
    If lblLinkAdd.ForeColor = &HC0C0C0 Then Exit Sub
    Call MakeLog
End Sub

Private Sub optInfo_Click(Index As Integer)
    Dim blnTmp As Boolean
    Dim strValue As String
    If mblnNoSave Then Exit Sub
    If Index = opt���� Or Index = opt���� Then
        Call UpDate�Һ���Ϣ("����", IIf(optInfo(opt����).Value, 1, 0))
    End If
End Sub

Private Sub optInfo_GotFocus(Index As Integer)
    Call SetCurCtlInfo(TypeName(optInfo(Index)), "optInfo", Index)
End Sub

Private Sub PatiAddress_GotFocus(Index As Integer)
    Call SetCurCtlInfo(TypeName(PatiAddress(Index)), "PatiAddress", Index)
End Sub

Private Sub PatiAddress_SetEdit(Index As Integer, blnEdit As Boolean)
    mblnUpdate = blnEdit
    RaiseEvent SetEdit
End Sub
Private Sub picPanel_Click(Index As Integer)
    Call SavePreItem
End Sub

Private Sub picPanel_Paint(Index As Integer)
    Call DrawLine(picPanel_����)
    Call DrawLine(picPanel_������Ϣ)
End Sub

Private Sub SetCtlPos(Index As Integer)
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim lngH As Long
    Dim lngW1 As Long, lngW2 As Long
    Dim lngTmp As Long
    Dim lngDXCboLin As Long
    
    On Error Resume Next
 
    If Index = picPanel_������Ϣ Then
        lblN(I���֤��).Top = lblN(lbl�������).Top + lblN(lbl�������).Height + 200
        lngH = 150
        Call zlControl.SetPubCtrlPos(True, 1, lblN(I���֤��), lngH, lblN(I����֤��), lngH, lblN(I�����ص�), lngH, lblN(I���ڵ�ַ), lngH, _
            lblN(I��λ����), lngH, lblN(I��ͥ��ַ), lngH, lblN(I����״��), lngH, lblN(Iְҵ), lngH, lblN(IѪ��), lngH, lblN(I�໤�����֤��), lngH, lblLink)
        
        lngW1 = 40: lngW2 = 350
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I���֤��), lngW1, fraC(I���֤��), lngW2, lblN(I����), lngW1, txtE(I����), lngW2, _
            lblN(I�Ļ��̶�), lngW1, fraC(I�Ļ��̶�))
            
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I�໤�����֤��), lngW1, txtE(I�໤�����֤��), lngW2 - 100)
        '--------------------------------------------------------------------------
        Call zlControl.SetPubCtrlPos(True, -1, fraC(I���֤��), lngH, txtE(I����֤��), lngH, txtE(I�����ص�), lngH, txtE(I���ڵ�ַ), lngH, _
            txtE(I��λ����), lngH, txtE(I��ͥ��ַ), lngH, fraC(I����״��), lngH, fraC(Iְҵ), lngH, txtE(I�໤�����֤��), lngH, fraC(IѪ��))
            
        Call zlControl.SetPubCtrlPos(True, -1, fraC(I���֤��), lngH, txtE(I����֤��), lngH, txtE(I�����ص�), lngH, txtE(I���ڵ�ַ), lngH, _
            txtE(I��λ����), lngH, txtE(I��ͥ��ַ), lngH, fraC(I����״��), lngH, fraC(Iְҵ), lngH, txtE(I�໤�����֤��), lngH, fraC(IѪ��))
        
        Call zlControl.SetPubCtrlPos(True, 1, lblN(I����), lngH, lblN(I����), lngH, lblN(I����), lngH, lblN(I����), lngH, lblN(IRH))
        
        Call zlControl.SetPubCtrlPos(True, -1, txtE(I����), lngH, txtE(I����), lngH, fraC(I����), lngH, fraC(I����), lngH, fraC(IRH))
        
        Call zlControl.SetPubCtrlPos(True, 1, lblN(I�Ļ��̶�), lngH, lblN(I����״��), lngH, lblN(I��λ�ʱ�), lngH, lblN(I�����ʱ�), lngH, _
            lblN(I��λ�绰), lngH, lblN(I��ͥ�绰), lngH, lblN(I��ͥ�ʱ�), lngH, lblN(I�໤��))
            
        Call zlControl.SetPubCtrlPos(True, -1, fraC(I�Ļ��̶�), lngH, fraC(I����״��), lngH, txtE(I��λ�ʱ�), lngH, txtE(I�����ʱ�), lngH, _
             txtE(I��λ�绰), lngH, txtE(I��ͥ�绰), lngH, txtE(I��ͥ�ʱ�), lngH, txtE(I�໤��))
             
        '----------------------------------------------------------------------------
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I����֤��), lngW1, txtE(I����֤��))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I����), lngW1, txtE(I����))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I����״��), lngW1, fraC(I����״��))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I�໤�����֤��), lngW1, txtE(I�໤�����֤��))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I�����ص�), lngW1, txtE(I�����ص�))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I��λ�ʱ�), lngW1, txtE(I��λ�ʱ�))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I���ڵ�ַ), lngW1, txtE(I���ڵ�ַ))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I�����ʱ�), lngW1, txtE(I�����ʱ�))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I��λ����), lngW1, txtE(I��λ����))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I��λ�绰), lngW1, txtE(I��λ�绰))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I��ͥ��ַ), lngW1, txtE(I��ͥ��ַ))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I��ͥ�绰), lngW1, txtE(I��ͥ�绰))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I����״��), lngW1, fraC(I����״��))
        lblN(I����).Top = lblN(I����״��).Top
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I����), lngW1, fraC(I����))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I��ͥ�ʱ�), lngW1, txtE(I��ͥ�ʱ�))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(Iְҵ), lngW1, fraC(Iְҵ))
        lblN(I����).Top = lblN(Iְҵ).Top
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I����), lngW1, fraC(I����))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I�໤��), lngW1, txtE(I�໤��))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(IѪ��), lngW1, fraC(IѪ��))
        lblN(IRH).Top = lblN(IѪ��).Top
        Call zlControl.SetPubCtrlPos(False, 0, lblN(IRH), lngW1, fraC(IRH))
        
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I����), lngW1, cmdE(I����))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I����), lngW1, cmdE(I����))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I�����ص�), lngW1, cmdE(I�����ص�))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I��λ����), lngW1, cmdE(I��λ����))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I��ͥ��ַ), lngW1, cmdE(I��ͥ��ַ))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I���ڵ�ַ), lngW1, cmdE(I���ڵ�ַ))
        
        lngW1 = 60
        
        cboE(I���֤��).Left = -30
        cboE(I���֤��).Top = -30
        fraC(I���֤��).Width = cboE(I���֤��).Width
        fraC(I���֤��).Height = cboE(I���֤��).Height - lngW1
        
        cboE(I����״��).Left = -30
        cboE(I����״��).Top = -30
        fraC(I����״��).Width = cboE(I����״��).Width
        fraC(I����״��).Height = cboE(I����״��).Height - lngW1
        
        cboE(Iְҵ).Left = -30
        cboE(Iְҵ).Top = -30
        fraC(Iְҵ).Width = cboE(Iְҵ).Width
        fraC(Iְҵ).Height = cboE(Iְҵ).Height - lngW1
        
        cboE(IѪ��).Left = -30
        cboE(IѪ��).Top = -30
        fraC(IѪ��).Width = cboE(IѪ��).Width
        fraC(IѪ��).Height = cboE(IѪ��).Height - lngW1
        
        cboE(IRH).Left = -30
        cboE(IRH).Top = -30
        fraC(IRH).Width = cboE(IRH).Width
        fraC(IRH).Height = cboE(IRH).Height - lngW1
        
        cboE(I����).Left = -30
        cboE(I����).Top = -30
        fraC(I����).Width = cboE(I����).Width
        fraC(I����).Height = cboE(I����).Height - lngW1
        
        cboE(I����).Left = -30
        cboE(I����).Top = -30
        fraC(I����).Width = cboE(I����).Width
        fraC(I����).Height = cboE(I����).Height - lngW1
        
        cboE(I�Ļ��̶�).Left = -30
        cboE(I�Ļ��̶�).Top = -30
        fraC(I�Ļ��̶�).Width = cboE(I�Ļ��̶�).Width
        fraC(I�Ļ��̶�).Height = cboE(I�Ļ��̶�).Height - lngW1
        
        cboE(I����״��).Left = -30
        cboE(I����״��).Top = -30
        fraC(I����״��).Width = cboE(I����״��).Width
        fraC(I����״��).Height = cboE(I����״��).Height - lngW1
 
        
        txtE(I����).Width = txtE(I����).Width
        
        txtE(I�����ص�).Width = txtE(I����).Width + txtE(I����).Left - txtE(I�����ص�).Left
        txtE(I���ڵ�ַ).Width = txtE(I�����ص�).Width
        txtE(I��λ����).Width = txtE(I�����ص�).Width
        txtE(I��ͥ��ַ).Width = txtE(I�����ص�).Width
        
        txtE(I�����ʱ�).Width = txtE(I��λ�ʱ�).Width
        txtE(I��λ�绰).Width = txtE(I��λ�ʱ�).Width
        txtE(I��ͥ�绰).Width = txtE(I��λ�ʱ�).Width
        txtE(I��ͥ�ʱ�).Width = txtE(I��λ�ʱ�).Width
        txtE(I�໤��).Width = txtE(I��λ�ʱ�).Width
        
        cmdE(I����).Left = txtE(I����).Left + txtE(I����).Width - cmdE(I����).Width
        cmdE(I����).Left = cmdE(I����).Left
        cmdE(I�����ص�).Left = cmdE(I����).Left
        cmdE(I���ڵ�ַ).Left = cmdE(I����).Left
        cmdE(I��λ����).Left = cmdE(I����).Left
        cmdE(I��ͥ��ַ).Left = cmdE(I����).Left
        
        lblLink.Left = lblN(I���֤��).Left
        
        If mblnStructAdress Then
            PatiAddress(PT_�����ص�).Top = txtE(I�����ص�).Top: PatiAddress(PT_�����ص�).Left = txtE(I�����ص�).Left: PatiAddress(PT_�����ص�).Height = txtE(I�����ص�).Height: PatiAddress(PT_�����ص�).Width = txtE(I�����ص�).Width
            PatiAddress(PT_���ڵ�ַ).Top = txtE(I���ڵ�ַ).Top: PatiAddress(PT_���ڵ�ַ).Left = txtE(I���ڵ�ַ).Left: PatiAddress(PT_���ڵ�ַ).Height = txtE(I���ڵ�ַ).Height: PatiAddress(PT_���ڵ�ַ).Width = txtE(I���ڵ�ַ).Width
            PatiAddress(PT_��ͥ��ַ).Top = txtE(I��ͥ��ַ).Top: PatiAddress(PT_��ͥ��ַ).Left = txtE(I��ͥ��ַ).Left: PatiAddress(PT_��ͥ��ַ).Height = txtE(I��ͥ��ַ).Height: PatiAddress(PT_��ͥ��ַ).Width = txtE(I��ͥ��ַ).Width
        End If
        lblN(I�໤�����֤��).Left = lblN(I��ͥ��ַ).Left
        txtE(I�໤�����֤��).Left = lblN(I�໤�����֤��).Left + lblN(I�໤�����֤��).Width + 30
        Call DrawLine(Index)
    ElseIf picPanel_������Ϣ = Index Then
        lngW1 = 40: lngH = 150
        lngW2 = 350
        
        lblN(I����ʱ��).Top = lblN(lbl�������).Top + lblN(lbl�������).Height + 200
        
        lblN(I����ʱ��).Left = lblN(I����ժҪ).Left
     
        lblN(I������¼).Top = lblN(I����ʱ��).Top + lblN(I����ʱ��).Height + 200
        
        lblN(I������¼).Left = lblN(I����ժҪ).Left
        
        chkNoAller.Left = lblN(I������¼).Left + lblN(I������¼).Width + 200
        chkNoAller.Top = lblN(I������¼).Top - 20
            
        PicPanel(picPanel_����Դ).Top = lblN(I������¼).Top - 20
        PicPanel(picPanel_����Դ).Height = lblN(I������¼).Height + 40
        optInfo(opt����Դ).Left = optInfo(optҩƷĿ¼).Width
        
        vsAller.Top = lblN(I������¼).Top + lblN(I������¼).Height + 20
        vsAller.Left = lblN(I������¼).Left
        vsAller.Width = txtE(I�໤��).Width + txtE(I�໤��).Left - lblN(I������¼).Left
        
        lblN(I����ժҪ).Top = vsAller.Top + vsAller.Height + 200
        txtE(I����ժҪ).Top = lblN(I����ժҪ).Top + lblN(I����ժҪ).Height + 20
        txtE(I����ժҪ).Left = lblN(I����ժҪ).Left
        txtE(I����ժҪ).Height = IIf(mbytSize = 0, 600, 700)
        txtE(I����ժҪ).Width = txtE(I�໤��).Width + txtE(I�໤��).Left - txtE(I����ժҪ).Left
      
        cmdSaveZY.Top = txtE(I����ժҪ).Top + txtE(I����ժҪ).Height + 20
        cmdSaveZY.Left = txtE(I����ժҪ).Width + txtE(I����ժҪ).Left - cmdSaveZY.Width
        
        cmdShowZY.Left = cmdSaveZY.Left - cmdShowZY.Width - 10
        cmdShowZY.Top = cmdSaveZY.Top
        
        '----���
        lblN(I��ϼ�¼).Left = lblN(I����ժҪ).Left
        lblN(I��ϼ�¼).Top = cmdShowZY.Top + cmdShowZY.Height + 160
        
        lblLinkAdd.Top = lblN(I��ϼ�¼).Top
        lblLinkAdd.Left = chkNoAller.Left + PicPanel(picPanel_���).Width
        
        PicPanel(picPanel_���).Left = chkNoAller.Left
        PicPanel(picPanel_���).Top = lblN(I��ϼ�¼).Top - 20
        PicPanel(picPanel_���).Height = lblN(I��ϼ�¼).Height + 40
        optInfo(opt���).Left = optInfo(opt����).Width
        
        vsDiagXY.Top = lblN(I��ϼ�¼).Top + lblN(I��ϼ�¼).Height + 20
        vsDiagXY.Left = vsAller.Left
        vsDiagXY.Width = txtE(I����ժҪ).Width
        
        cmdDiagMove(0).Move vsDiagXY.Left + vsDiagXY.Width + 70, vsDiagXY.Top + 150, 375, 375
        cmdDiagMove(1).Move vsDiagXY.Left + vsDiagXY.Width + 70, vsDiagXY.Top + cmdDiagMove(0).Height + 250, 375, 375
        
        If mbln��ҽ Then
            vsDiagZY.Visible = True
            vsDiagZY.Top = vsDiagXY.Top + vsDiagXY.Height + lngH
            vsDiagZY.Left = vsDiagXY.Left
            vsDiagZY.Width = txtE(I����ժҪ).Width
            lngTmp = vsDiagZY.Height
            
            cmdDiagMove(2).Visible = True
            cmdDiagMove(3).Visible = True
            cmdDiagMove(2).Move vsDiagZY.Left + vsDiagZY.Width + 70, vsDiagZY.Top + 80, 375, 375
            cmdDiagMove(3).Move vsDiagZY.Left + vsDiagZY.Width + 70, vsDiagZY.Top + cmdDiagMove(2).Height + 180, 375, 375
            
        Else
            vsDiagZY.Visible = False
            lngTmp = 0
            cmdDiagMove(2).Visible = False
            cmdDiagMove(3).Visible = False
        End If
        
        PicPanel(picPanel_����).Left = vsDiagXY.Left
        PicPanel(picPanel_����).Top = vsDiagXY.Height + vsDiagXY.Top + lngTmp + 2 * lngH
        PicPanel(picPanel_����).Width = txtE(I����ժҪ).Width + 1200
            optInfo(opt����).Left = optInfo(opt����).Width

        lblN(Iҽѧ��ʾ).Top = lblN(Iȥ��).Top + lblN(Iȥ��).Height + 180
        
        lblN(lbl��������).Top = lblN(Iҽѧ��ʾ).Top + lblN(Iҽѧ��ʾ).Height + 180
        
        UCPatiVitalSigns.Top = lblN(lbl��������).Top + lblN(lbl��������).Height + 100
        
        fraC(Iȥ��).Width = cboE(Iȥ��).Width
        fraC(Iȥ��).Height = cboE(Iȥ��).Height - 60
        
        Call zlControl.SetPubCtrlPos(False, 0, chkInfo, lngW2, lblN(Iȥ��), lngW1, fraC(Iȥ��))
        
        fraC(I����).Width = cboE(I����).Width
        fraC(I����).Height = cboE(I����).Height - 60
         
        txtE(I������ַ).Width = fraC(Iȥ��).Width
        
        txtE(I����ҽѧ��ʾ).Width = fraC(Iȥ��).Width
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I����ʱ��), lngW1, txtSL, 0, fraC(I����), 30, txtE(I����ʱ��), lngW1, cmdE(I����ʱ��), lngW2, lblN(I������ַ), lngW1, txtE(I������ַ))
        
        txtE(I������ַ).Left = vsDiagXY.Left + vsDiagXY.Width - txtE(I������ַ).Width
        lblN(I������ַ).Left = txtE(I������ַ).Left - lblN(I������ַ).Width - 10
        
        
        txtE(I����ʱ��).Width = txtE(Iҽѧ��ʾ).Width
        
        cmdE(I����ʱ��).Left = txtE(I����ʱ��).Left + txtE(I����ʱ��).Width
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(Iҽѧ��ʾ), lngW1, txtE(Iҽѧ��ʾ), 900, lblN(I����ҽѧ��ʾ), lngW1, txtE(I����ҽѧ��ʾ))
        
        Call zlControl.SetPubCtrlPos(False, 0, txtE(Iҽѧ��ʾ), lngW1, cmdE(Iҽѧ��ʾ))
        
        cmdE(Iҽѧ��ʾ).Left = txtE(Iҽѧ��ʾ).Left + txtE(Iҽѧ��ʾ).Width
        
        fraC(Iȥ��).Left = txtE(I����ժҪ).Width - fraC(Iȥ��).Width
        lblN(Iȥ��).Left = fraC(Iȥ��).Left - lngW1 - lblN(Iȥ��).Width
        
        lblN(I����ҽѧ��ʾ).Left = lblN(Iȥ��).Left + lblN(Iȥ��).Width - lblN(I����ҽѧ��ʾ).Width
        txtE(I����ҽѧ��ʾ).Left = lblN(I����ҽѧ��ʾ).Left + lblN(I����ҽѧ��ʾ).Width + lngW1
        
        chkInfo.Left = cmdE(Iҽѧ��ʾ).Left + cmdE(Iҽѧ��ʾ).Width - chkInfo.Width + 130
        
        Call DrawLine(picPanel_����)
        
        '�»���λ���趨
        lngDXCboLin = 60
        
        x1 = txtSL.Left
        y1 = txtSL.Top + txtSL.Height
        x2 = txtSL.Left + txtSL.Width + txtSL.Width
        y2 = y1
        linD(0).x1 = x1
        linD(0).x2 = x2
        linD(0).y1 = y1
        linD(0).y2 = y2
        
        x1 = fraC(I����).Left
        y1 = fraC(I����).Top + fraC(I����).Height
        x2 = fraC(I����).Left + fraC(I����).Width - lngDXCboLin
        y2 = y1
        linD(1).x1 = x1
        linD(1).x2 = x2
        linD(1).y1 = y1
        linD(1).y2 = y2
        
        x1 = txtE(I����ʱ��).Left
        y1 = txtE(I����ʱ��).Top + txtE(I����ʱ��).Height
        x2 = txtE(I����ʱ��).Left + txtE(I����ʱ��).Width + cmdE(I����ʱ��).Width
        y2 = y1

        linD(2).x1 = x1
        linD(2).x2 = x2
        linD(2).y1 = y1
        linD(2).y2 = y2
        

        x1 = txtE(I������ַ).Left
        y1 = txtE(I������ַ).Top + txtE(I������ַ).Height
        x2 = txtE(I������ַ).Left + txtE(I������ַ).Width
        y2 = y1

        linD(3).x1 = x1
        linD(3).x2 = x2
        linD(3).y1 = y1
        linD(3).y2 = y2
        
    ElseIf picPanel_������� = Index Then
        '''''
        cmdSign.Left = txtE(I�໤��).Width + txtE(I�໤��).Left - cmdSign.Width
       
        cmdUpdate.Left = cmdSign.Left - cmdUpdate.Width - 30
        
        cmdImportEPRDemo.Left = cmdUpdate.Left - cmdImportEPRDemo.Width - 30
        
        lngW1 = txtE(I����ժҪ).Width \ 2
        
        For i = 0 To I����
            rtfEdit(i).Width = lngW1 - 200
        Next
        
        rtfEdit(I�ֲ�ʷ).Left = txtE(I�໤��).Width + txtE(I�໤��).Left - rtfEdit(I�ֲ�ʷ).Width
        
        lblDoc(I����).Top = 0
        lblDoc(I�ֲ�ʷ).Top = 0
        
        lblDoc(I����).Left = lblN(I���֤��).Left
        
        rtfEdit(I����).Left = lblDoc(I����).Left
        rtfEdit(I����).Top = lblDoc(I����).Top + lblDoc(I����).Height + 20
        
        
        lblDoc(I��ȥʷ).Top = rtfEdit(I����).Top + rtfEdit(I����).Height + 100
        lblDoc(I��ȥʷ).Left = lblN(I���֤��).Left
        
        rtfEdit(I��ȥʷ).Left = lblDoc(I��ȥʷ).Left
        rtfEdit(I��ȥʷ).Top = lblDoc(I��ȥʷ).Top + lblDoc(I��ȥʷ).Height + 20
        
        lblDoc(I����).Top = rtfEdit(I��ȥʷ).Top + rtfEdit(I��ȥʷ).Height + 100
        lblDoc(I����).Left = lblN(I���֤��).Left
        
        rtfEdit(I����).Left = lblDoc(I����).Left
        rtfEdit(I����).Top = lblDoc(I����).Top + lblDoc(I����).Height + 20
        
        lblDoc(I�ֲ�ʷ).Left = rtfEdit(I�ֲ�ʷ).Left
        
        rtfEdit(I�ֲ�ʷ).Left = lblDoc(I�ֲ�ʷ).Left
        rtfEdit(I�ֲ�ʷ).Top = lblDoc(I�ֲ�ʷ).Top + lblDoc(I�ֲ�ʷ).Height + 20
        
        lblDoc(I����ʷ).Top = lblDoc(I��ȥʷ).Top
        lblDoc(I����ʷ).Left = lblDoc(I�ֲ�ʷ).Left
        
        rtfEdit(I����ʷ).Left = lblDoc(I����ʷ).Left
        rtfEdit(I����ʷ).Top = lblDoc(I����ʷ).Top + lblDoc(I����ʷ).Height + 20
        
        picPrompt.Top = rtfEdit(I����).Top + rtfEdit(I����).Height / 2 - lblTip.Height
        picPrompt.Left = lblDoc(I����ʷ).Left
        
        lblTip.Top = picPrompt.Top
        lblTip.Left = picPrompt.Left + picPrompt.Width + 20
        
        lngW1 = rtfEdit(I����).Top + rtfEdit(I����).Height + 120
        
        linDoc.y1 = lngW1 - 60
        linDoc.y2 = linDoc.y1
        linDoc.x1 = lblDoc(I����).Left
        linDoc.x2 = rtfEdit(I����ʷ).Left + rtfEdit(I����ʷ).Width
        
        cmdSign.Top = lngW1
        cmdUpdate.Top = lngW1
        cmdImportEPRDemo.Top = lngW1
        
        lblEPRname.Left = lblDoc(I����).Left
        lblEPRname.Top = lngW1 + 20
        
        lblDoctor(0).Left = lblEPRname.Left + lblEPRname.Width + 300
        lblDoctor(0).Top = lblEPRname.Top
        
        lblDoctor(1).Left = lblDoctor(0).Left + lblDoctor(0).Width + 20
        lblDoctor(1).Top = lblEPRname.Top
        
    End If
End Sub

Private Sub DrawLine(ByVal intIdx As Integer)
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim objPic As Object
    Dim lngDXCboLin As Long
    
    If Not (intIdx = picPanel_���� Or intIdx = picPanel_������Ϣ) Then
        Exit Sub
    End If
    
    On Error Resume Next
    lngDXCboLin = 60
    Set objPic = PicPanel(intIdx)
    
    objPic.Cls
    If intIdx = picPanel_���� Then
        
        x1 = txtE(Iҽѧ��ʾ).Left
        y1 = txtE(Iҽѧ��ʾ).Top + txtE(Iҽѧ��ʾ).Height
        x2 = txtE(Iҽѧ��ʾ).Left + txtE(Iҽѧ��ʾ).Width + cmdE(Iҽѧ��ʾ).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I����ҽѧ��ʾ).Left
        y1 = txtE(I����ҽѧ��ʾ).Top + txtE(I����ҽѧ��ʾ).Height
        x2 = txtE(I����ҽѧ��ʾ).Left + txtE(I����ҽѧ��ʾ).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
    
        x1 = fraC(Iȥ��).Left
        y1 = fraC(Iȥ��).Top + fraC(Iȥ��).Height
        x2 = fraC(Iȥ��).Left + fraC(Iȥ��).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        '--------�����������ı߿� 4 ����-------------
        If mint���� = 1 Then
            x1 = 0
            y1 = UCPatiVitalSigns.Top - 80
            x2 = txtE(I����ҽѧ��ʾ).Left + txtE(I����ҽѧ��ʾ).Width
            y2 = y1
            objPic.Line (x1, y1)-(x2, y2)
        
            x1 = 0
            y1 = UCPatiVitalSigns.Top + UCPatiVitalSigns.Height + 20
            x2 = txtE(I����ҽѧ��ʾ).Left + txtE(I����ҽѧ��ʾ).Width
            y2 = y1
            objPic.Line (x1, y1)-(x2, y2)
            
            x1 = 0
            y1 = UCPatiVitalSigns.Top - 80
            x2 = 0
            y2 = UCPatiVitalSigns.Top + UCPatiVitalSigns.Height + 20
            objPic.Line (x1, y1)-(x2, y2)
            
            x1 = txtE(I����ҽѧ��ʾ).Left + txtE(I����ҽѧ��ʾ).Width
            y1 = UCPatiVitalSigns.Top - 80
            x2 = x1
            y2 = UCPatiVitalSigns.Top + UCPatiVitalSigns.Height + 20
            objPic.Line (x1, y1)-(x2, y2)
        End If
        '-------------------------
        
    ElseIf intIdx = picPanel_������Ϣ Then
        x1 = txtE(I����).Left
        y1 = txtE(I����).Top + txtE(I����).Height
        x2 = txtE(I����).Left + txtE(I����).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I����֤��).Left
        y1 = txtE(I����֤��).Top + txtE(I����֤��).Height
        x2 = txtE(I����֤��).Left + txtE(I����֤��).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I�໤�����֤��).Left
        y1 = txtE(I�໤�����֤��).Top + txtE(I�໤�����֤��).Height
        x2 = txtE(I�໤�����֤��).Left + txtE(I�໤�����֤��).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I����).Left
        y1 = txtE(I����).Top + txtE(I����).Height
        x2 = txtE(I����).Left + txtE(I����).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I��λ�ʱ�).Left
        y1 = txtE(I��λ�ʱ�).Top + txtE(I��λ�ʱ�).Height
        x2 = txtE(I��λ�ʱ�).Left + txtE(I��λ�ʱ�).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I�����ʱ�).Left
        y1 = txtE(I�����ʱ�).Top + txtE(I�����ʱ�).Height
        x2 = txtE(I�����ʱ�).Left + txtE(I�����ʱ�).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I��ͥ�ʱ�).Left
        y1 = txtE(I��ͥ�ʱ�).Top + txtE(I��ͥ�ʱ�).Height
        x2 = txtE(I��ͥ�ʱ�).Left + txtE(I��ͥ�ʱ�).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I��ͥ�绰).Left
        y1 = txtE(I��ͥ�绰).Top + txtE(I��ͥ�绰).Height
        x2 = txtE(I��ͥ�绰).Left + txtE(I��ͥ�绰).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I��λ�绰).Left
        y1 = txtE(I��λ�绰).Top + txtE(I��λ�绰).Height
        x2 = txtE(I��λ�绰).Left + txtE(I��λ�绰).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I�໤��).Left
        y1 = txtE(I�໤��).Top + txtE(I�໤��).Height
        x2 = txtE(I�໤��).Left + txtE(I�໤��).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I��λ����).Left
        y1 = txtE(I��λ����).Top + txtE(I��λ����).Height
        x2 = txtE(I��λ����).Left + txtE(I��λ����).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        If Not mblnStructAdress Then
           x1 = txtE(I�����ص�).Left
           y1 = txtE(I�����ص�).Top + txtE(I�����ص�).Height
           x2 = txtE(I�����ص�).Left + txtE(I�����ص�).Width
           y2 = y1
           objPic.Line (x1, y1)-(x2, y2)
           
           x1 = txtE(I��ͥ��ַ).Left
           y1 = txtE(I��ͥ��ַ).Top + txtE(I��ͥ��ַ).Height
           x2 = txtE(I��ͥ��ַ).Left + txtE(I��ͥ��ַ).Width
           y2 = y1
           objPic.Line (x1, y1)-(x2, y2)
           
           x1 = txtE(I���ڵ�ַ).Left
           y1 = txtE(I���ڵ�ַ).Top + txtE(I���ڵ�ַ).Height
           x2 = txtE(I���ڵ�ַ).Left + txtE(I���ڵ�ַ).Width
           y2 = y1
           objPic.Line (x1, y1)-(x2, y2)
        End If
        
        x1 = fraC(I���֤��).Left
        y1 = fraC(I���֤��).Top + fraC(I���֤��).Height
        x2 = fraC(I���֤��).Left + fraC(I���֤��).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I�Ļ��̶�).Left
        y1 = fraC(I�Ļ��̶�).Top + fraC(I�Ļ��̶�).Height
        x2 = fraC(I�Ļ��̶�).Left + fraC(I�Ļ��̶�).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I�Ļ��̶�).Left
        y1 = fraC(I�Ļ��̶�).Top + fraC(I�Ļ��̶�).Height
        x2 = fraC(I�Ļ��̶�).Left + fraC(I�Ļ��̶�).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(IRH).Left
        y1 = fraC(IRH).Top + fraC(IRH).Height
        x2 = fraC(IRH).Left + fraC(IRH).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(IѪ��).Left
        y1 = fraC(IѪ��).Top + fraC(IѪ��).Height
        x2 = fraC(IѪ��).Left + fraC(IѪ��).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(Iְҵ).Left
        y1 = fraC(Iְҵ).Top + fraC(Iְҵ).Height
        x2 = fraC(Iְҵ).Left + fraC(Iְҵ).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I����״��).Left
        y1 = fraC(I����״��).Top + fraC(I����״��).Height
        x2 = fraC(I����״��).Left + fraC(I����״��).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I����).Left
        y1 = fraC(I����).Top + fraC(I����).Height
        x2 = fraC(I����).Left + fraC(I����).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        
        x1 = fraC(I����).Left
        y1 = fraC(I����).Top + fraC(I����).Height
        x2 = fraC(I����).Left + fraC(I����).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        
        x1 = fraC(I����״��).Left
        y1 = fraC(I����״��).Top + fraC(I����״��).Height
        x2 = fraC(I����״��).Left + fraC(I����״��).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
    End If
End Sub

Private Sub SetRTFEditFontSize()
'���ܣ����ò��˲�ʷ��Ϣ�����������
    Dim i As Long
    lblEPRname.Visible = True
    mblnSizeTmp = True
    For i = 0 To rtfEdit.UBound
        Call zlControl.RTFSetFontSize(rtfEdit(i), IIf(mbytSize = 0, 9, 12))
    Next
    mblnSizeTmp = False
End Sub

Private Sub lblDoctor_Click(Index As Integer)
    lblEPRname.Visible = Not lblEPRname.Visible
End Sub

Private Sub rtfEdit_Change(Index As Integer)
    If mblnSizeTmp = True Then Exit Sub
    mblnChange = True
End Sub

Private Sub rtfEdit_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True)
    Call SetCurCtlInfo(TypeName(rtfEdit(Index)), "rtfEdit", Index)
End Sub

Private Sub rtfEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not mblnPatiChange Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '�������λس������ת
            With rtfEdit(Index)
                If Trim(.Text) = "" Then
                    KeyAscii = 0
                    Call zlCommFun.PressKey(vbKeyTab)
                ElseIf .SelStart - 1 > 0 Then
                    If Mid(.Text, .SelStart - 1, 2) = vbCrLf Then
                        KeyAscii = 0
                        Call zlCommFun.PressKey(vbKeyBack)
                        Call zlCommFun.PressKey(vbKeyTab)
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub rtfEdit_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme
End Sub

Private Sub rtfEdit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTmp As String
    mrsMainInfo.Filter = "�ؼ���='rtfEdit' and Index=" & Index
    strTmp = mrsMainInfo!ErrInfo & ""
    Call zlCommFun.ShowTipInfo(rtfEdit(Index).hwnd, strTmp, True, True)
End Sub

Private Sub rtfEdit_SelChange(Index As Integer)
    With rtfEdit(Index)
        If .SelLength = 0 And .SelStart > 0 And PicPanel(picPanel_�������).Tag = "" Then
            If Mid(.Text, .SelStart, 1) = "`" Or Mid(.Text, .SelStart, 1) = "��" Then
                PicPanel(picPanel_�������).Tag = "UnChange"
                .SelStart = .SelStart - 1
                .SelLength = 1
                .SelText = ""
                Call ShowWordInput(rtfEdit(Index))
                PicPanel(picPanel_�������).Tag = ""
            End If
        End If
    End With
End Sub

Private Sub rtfEdit_Validate(Index As Integer, Cancel As Boolean)
    Call UpDate����(Index)
End Sub

Private Sub txtE_Change(Index As Integer)
    Dim txtTmp As Object
    Dim lngPos As Long, lngLen As Long
    If Index = I�໤�����֤�� Then
        If mblnReturn Then Exit Sub
        Set txtTmp = txtE(Index)
        mblnReturn = True
        '�����������
        If Not zlStr.CheckCharScope(txtTmp.Text, "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ*") Then
            txtTmp.Text = ""
        Else
            If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
                If zlCommFun.ActualLen(txtTmp.Text) > 18 Then
                    txtTmp.Text = Mid(txtTmp.Text, 1, 18)
                End If
            End If
        End If
        If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
            lngPos = InStr(txtTmp.Text, "*")
            lngLen = Len(Mid(txtTmp.Text, 13, 2))
            Select Case lngPos
                Case 0
                    txtTmp.Tag = txtTmp.Text
                Case Else
                    txtTmp.Tag = Mid(txtTmp.Text, 1, lngPos - 1)
                    txtTmp.Text = txtTmp.Tag
                    txtTmp.SelStart = Len(txtTmp.Text)
            End Select
        Else
            txtTmp.Tag = txtTmp.Text
        End If
        mblnReturn = False
    End If
End Sub

Private Sub txtE_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(txtE(Index).hwnd, txtE(Index).Tag, True, True)
End Sub

Private Sub txtE_Validate(Index As Integer, Cancel As Boolean)
    Dim strValue As String, str���֤�� As String
    If Not txtE(Index).Locked Then
        If Index = I����ʱ�� Then
            txtE(Index).Text = GetFullDate(txtE(Index).Text)
        End If
        If Index = I����ʱ�� Then
            txtSL.Text = ""
            cboE(I����).ListIndex = -1
        End If
        If Index = I�໤�����֤�� Then
            '���������֤���Ǵ��ڲ�������  cboE(Index).Tag
            strValue = txtE(Index).Tag
            mrsMainInfo.Filter = "��Ϣ��='���֤��'"
            str���֤�� = mrsMainInfo!��Ϣԭֵ & ""
            If strValue <> str���֤�� Then
                If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
                    If Not Check���֤��(strValue, txtE(Index)) Then
                        Cancel = True
                        txtE(Index).SetFocus
                        Exit Sub
                    End If
                End If
            End If
            Call Update�໤�����֤
        Else
            Call UpDateInfo(txtE(Index).Text, "txtE", Index)
        End If
    End If
End Sub

Private Sub txtSentence_GotFocus()
    Call zlCommFun.OpenIme(True)
    Call zlControl.TxtSelAll(txtSentence)
End Sub

Private Sub txtSentence_KeyPress(KeyAscii As Integer)
    Dim strSentence As String, blnCancel As Boolean, strType As String
       
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        Select Case Val(picSentence.Tag)
        Case I����
            strType = "��������"
        Case I����ʷ
            strType = "����ʷ"
        Case I�ֲ�ʷ
            strType = "�ֲ�ʷ"
        Case I����
            strType = "���һ����"
        Case I��ȥʷ
            strType = "����ʷ"
        End Select
                
        strSentence = frmSentenceSel.ShowMe(Me, mlng�����ļ�id, mstr�Ա�, mstr����״��, strType, txtSentence.Text, picSentence.hwnd, blnCancel)
        If strSentence <> "" Then
            rtfEdit(Val(picSentence.Tag)).SelText = strSentence
            Call HideWordInput
        Else
            If Not blnCancel Then
                MsgBox "û���ҵ�ƥ��Ĵʾ䡣", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txtSentence)
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call imgSentence_Click
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
        Call HideWordInput
    End If
End Sub


Private Sub imgSentence_Click()
    Dim strSentence As String, strType As String
    
    Select Case Val(picSentence.Tag)
    Case I����
        strType = "��������"
    Case I����ʷ
        strType = "����ʷ"
    Case I�ֲ�ʷ
        strType = "�ֲ�ʷ"
    Case I����
        strType = "���һ����"
    Case I��ȥʷ
        strType = "����ʷ"
    End Select
    
    strSentence = frmSentenceSel.ShowMe(Me, mlng�����ļ�id, mstr�Ա�, mstr����״��, strType)
    If strSentence <> "" Then
        rtfEdit(Val(picSentence.Tag)).SelText = strSentence
        Call HideWordInput
    End If
End Sub

Private Sub txtSentence_LostFocus()
    If Not frmSentenceSel.mblnShow Then
        Call HideWordInput   '���شʾ�����
    End If
End Sub

Private Sub ShowWordInput(ByRef txtThis As RichTextBox)
'���ܣ���ʾ�ʾ�����
    Dim vPos As POINTAPI
    
    If txtThis.Visible And txtThis.Enabled And Not txtThis.Locked Then
        picSentence.Tag = txtThis.Index '�����Ա����ط��غ�λ
        
        If txtThis.Text = "" Then PicPanel(picPanel_�������).Tag = "UnChange": txtThis.Text = " " '����Ҫ��һ�����ַ����ܷ���������
        vPos = zlControl.GetCaretPos(txtThis.hwnd)
        If txtThis.Text = " " Then PicPanel(picPanel_�������).Tag = "UnChange": txtThis.Text = ""
        
        If vPos.X <> -1 And vPos.Y <> -1 Then
            If txtThis.Left + vPos.X + Screen.TwipsPerPixelX * 2 < txtThis.Left + txtThis.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX Then
                picSentence.Left = txtThis.Left + vPos.X + Screen.TwipsPerPixelX * 2
            Else
                picSentence.Left = txtThis.Left + txtThis.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX
            End If
            picSentence.Top = txtThis.Top + vPos.Y + Screen.TwipsPerPixelY
            txtSentence.Text = ""
            picSentence.Visible = True
            txtSentence.SetFocus
        End If
    End If
End Sub

Private Sub HideWordInput()
'���ܣ����شʾ�����
    Dim idx As Long
    
    If picSentence.Visible Then
        picSentence.Visible = False
        txtSentence.Text = ""
        
        idx = Val(picSentence.Tag)
        picSentence.Tag = ""
        
        If rtfEdit(idx).Visible And rtfEdit(idx).Enabled And Not rtfEdit(idx).Locked Then
            rtfEdit(idx).SetFocus
        End If
    End If
End Sub

Private Sub LoadDocData()
    Dim rsTmp As ADODB.Recordset, strSQL As String, strContent As String
    Dim i As Long, j As Long, arrTmp  As Variant, blnLoading As Boolean
    
    mblnNoSave = True
    For i = 0 To rtfEdit.UBound
        rtfEdit(i).Text = ""
        rtfEdit(i).Tag = ""
    Next
    lblEPRname.Caption = ""    '�����ڼ�����Ա��ԭ�����磺ѡ����ٴʾ�ʱ��û���г�Ԥ�ƵĴʾ䣬�ɸ��ݲ����ļ����Ʋ��Ƿ���������ٴʾ��Ӧ
    
    'ֻ��ʾ�򵥲���ģʽ�²������ļ�
    strSQL = "Select id,�ļ�id,ǩ������,��������,������ From ���Ӳ�����¼ A Where ����id = [1] And ��ҳid = [2] And �������� = 1" & vbNewLine & _
        " And Exists(Select 1 From �����ļ��б� B Where A.�ļ�ID = B.ID And B.���� = '3')"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    If rsTmp.RecordCount > 0 Then
        If SetCompendsTag(Val("" & rsTmp!�ļ�ID)) Then
            mlng�����ļ�id = Val("" & rsTmp!�ļ�ID)
            mblnǩ�� = IIf(Val("" & rsTmp!ǩ������) > 0, True, False)
            mlng����ID = rsTmp!ID
            mstr������ = "" & rsTmp!������
            lblEPRname.Caption = "" & rsTmp!��������
            '��ȡ����µĶ����ı�,��������Ϊ-1��ʾ��ٱ����ı�
            strSQL = "Select A.Ԥ�����id, B.�����ı�, B.ID" & vbNewLine & _
                    "From ���Ӳ������� A, ���Ӳ������� B" & vbNewLine & _
                    "Where A.�ļ�id = [1] And A.�������� = 1 And A.Ԥ�����id+0 In(-10,5,2,6,3)" & vbNewLine & _
                    "      And B.��id = A.ID And B.�������� = 2 Order By A.Ԥ�����id, B.�������"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
            With rsTmp
                If .RecordCount > 0 Then
                    arrTmp = Split("-10,2,3,5,6", ",") '��������,�ֲ�ʷ,����ʷ,����ʷ,�����
                    For i = 0 To UBound(arrTmp)
                        .Filter = "Ԥ�����id=" & arrTmp(i)
                        For j = 1 To .RecordCount
                            If j = 1 Then
                                strContent = "" & !�����ı�
                                If InStr(strContent, lblDoc(i).Tag) = 1 Then strContent = Mid(strContent, Len(lblDoc(i).Tag) + 1)
                                rtfEdit(i).Text = strContent
                                rtfEdit(i).Tag = !ID
                            Else
                                rtfEdit(i).Text = rtfEdit(i).Text & vbCrLf & !�����ı�
                                rtfEdit(i).Tag = rtfEdit(i).Tag & "," & !ID
                            End If
                            .MoveNext
                        Next
                    Next
                End If
            End With
        End If
    Else
        If mbln�� Then
            strSQL = " And (R.�¼� = '����'  OR R.�¼� IS NUll)"
        Else
            If optInfo(opt����).Value Then
                strSQL = " And (R.�¼� = '����' Or R.�¼� = '����'  OR R.�¼� IS NUll)"
            Else
                strSQL = " And (R.�¼� = '����' Or R.�¼� = '����'  OR R.�¼� IS NUll )"
            End If
        End If
        'ϵͳ��������(��)�ﲡ���ҶԵ�ǰ�������ã�����5���̶�Ԥ�����,����ʾ����¼�����.
        strSQL = "Select F.ID, F.���� as ��������" & vbNewLine & _
                "From (Select F.ID, F.ͨ��, A.����id, F.����,Decode(R.�¼�,Null,2,1) �¼�" & vbNewLine & _
                "       From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� R" & vbNewLine & _
                "       Where F.ID = A.�ļ�id(+) And F.ID = R.�ļ�id(+) And F.���� = 1 And F.����= '3'" & strSQL & ") F" & vbNewLine & _
                "Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [2]" & vbNewLine & _
                "Order By F.�¼�,F.ͨ�� Desc,F.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�Һ�ID, mlng����ID)
        If rsTmp.RecordCount > 0 Then
            mlng�����ļ�id = rsTmp!ID
            lblEPRname.Caption = "" & rsTmp!��������
            If SetCompendsTag(mlng�����ļ�id) = False Then
                mlng�����ļ�id = 0: lblEPRname.Caption = ""
            End If
        Else
            mlng�����ļ�id = 0: lblEPRname.Caption = ""
        End If
        
        mlng����ID = 0
        mblnǩ�� = False
    End If
    
    mrsMainInfo.Filter = "�ؼ���='rtfEdit'"
    
    For i = 1 To mrsMainInfo.RecordCount
        mrsMainInfo!��Ϣԭֵ = rtfEdit(mrsMainInfo!Index).Text
        mrsMainInfo.Update
        mrsMainInfo.MoveNext
    Next
    mblnNoSave = False
    cmdImportEPRDemo.Visible = mlng�����ļ�id <> 0
    Call SetDocEditable
    Call SetRTFEditFontSize
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetCompendsTag(ByVal lng�����ļ�id As Long) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    
    strSQL = "Select Decode(A.Ԥ�����id, -10, 0, 5, 3, 2, 1, 6, 4, 3, 2) As ���, B.�����ı�" & vbNewLine & _
            "From �����ļ��ṹ A, �����ļ��ṹ B" & vbNewLine & _
            "Where A.�ļ�id = [1] And A.Ԥ�����id+0 In (-10,5,2,6,3) And A.Id = B.��id And B.�������� = 2" & vbNewLine & _
            "Order By ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�����ļ�id)
    If rsTmp.RecordCount > 0 And rsTmp.RecordCount <= 5 Then
        If rsTmp!��� & "" = "0" Then  '�����������
            For i = 0 To rsTmp.RecordCount - 1
                lblDoc(Val(rsTmp!��� & "")).Tag = rsTmp!�����ı�       '���ڱ���Rtf�ļ��滻����ʱ��λ
                rsTmp.MoveNext
            Next
            For i = 1 To lblDoc.Count - 1
                If lblDoc(i).Tag = "" Then
                    lblDoc(i).Visible = False
                    rtfEdit(i).Visible = False
                Else
                    lblDoc(i).Visible = True
                    rtfEdit(i).Visible = True
                End If
            Next
            SetCompendsTag = True
        End If
    End If
End Function

Private Function CanUseFastEPR() As Boolean
'���ܣ��Ƿ��п��õĿ�ݲ����ļ�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTmp As Boolean
    
    On Error GoTo errH
    
    blnTmp = InStr(GetInsidePrivs(p���ﲡ������), "������д") > 0
    
    If Not blnTmp Then Exit Function
    
    If mbln�� Then
        strSQL = " And (R.�¼� = '����'  OR R.�¼� IS NUll)"
    Else
        If optInfo(opt����).Value Then
            strSQL = " And (R.�¼� = '����' Or R.�¼� = '����'  OR R.�¼� IS NUll)"
        Else
            strSQL = " And (R.�¼� = '����' Or R.�¼� = '����'  OR R.�¼� IS NUll )"
        End If
    End If
    'ϵͳ��������(��)�ﲡ���ҶԵ�ǰ�������ã�����5���̶�Ԥ�����,����ʾ����¼�����.
    strSQL = "Select F.ID, F.���� as ��������" & vbNewLine & _
            "From (Select F.ID, F.ͨ��, A.����id, F.����,Decode(R.�¼�,Null,2,1) �¼�" & vbNewLine & _
            "       From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� R" & vbNewLine & _
            "       Where F.ID = A.�ļ�id(+) And F.ID = R.�ļ�id(+) And F.���� = 1 And F.����= '3'" & strSQL & ") F" & vbNewLine & _
            "Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [2]" & vbNewLine & _
            "Order By F.�¼�,F.ͨ�� Desc,F.id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�Һ�ID, mlng����ID)
    If rsTmp.RecordCount > 0 Then
        mlng�����ļ�id = rsTmp!ID
        lblEPRname.Caption = "" & rsTmp!��������
        If SetCompendsTag(mlng�����ļ�id) Then
            CanUseFastEPR = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function NeedName(strList As String) As String
'˵��:1-strList��()��[]�ָ����������ʱ��������[����]��(����)��ͷ,�������Ϊ���ֻ���ĸ
'     2-�ָ��������ȼ����س���(Chr(13)��> - > [] > ()

    '�����ж��Իس����ָ�
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
        Exit Function
    End If
    '��[]�ָ�
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
        If zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
            Exit Function
        End If
    End If
    '��()�ָ�
    If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
        If zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
            Exit Function
        End If
    End If
    '��-�ָ�
    NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    
End Function

Private Sub GetCboIndex(objCbo As Object, ByVal strFind As String)
'���ܣ����ַ�����ComboBox�в�������
'������blnKeep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Dim i As Integer

    '�Ⱦ�ȷ����
    For i = 0 To objCbo.ListCount - 1
        If objCbo.List(i) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        ElseIf NeedName(objCbo.List(i)) = strFind And strFind <> "" Then
            objCbo.ListIndex = i: Exit Sub
        End If
    Next
    
    '���ģ������
    If strFind <> "" Then
        For i = 0 To objCbo.ListCount - 1
            If InStr(objCbo.List(i), strFind) > 0 And strFind <> "" Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
End Sub

Private Sub SetCtlBackColor()
    Dim objCtl As Object
    For Each objCtl In Me.Controls
        If UCase(TypeName(objCtl)) <> "LINE" And InStr(",LINE,VSCROLLBAR,IMAGE,STATUSBAR,", "," & UCase(TypeName(objCtl)) & ",") = 0 Then
        objCtl.BackColor = vbWhite ' &HFFFFFF
        End If
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = vsc.Value
    lngMin = vsc.Min
    lngMax = vsc.Max
    
    If KeyCode = vbKeyPageDown Then '��
        If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsc.Value = lngCur + (lngMax - lngMin) / 10
        Else
            vsc.Value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '��
        If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsc.Value = lngCur - (lngMax - lngMin) / 10
        Else
            vsc.Value = lngMin
        End If
    End If
    
End Sub

Private Sub Form_Activate()
'������
    glngPreHWnd = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf FlexScroll
    If Not mobjCtl Is Nothing Then
        mobjCtl.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()
'������
    SetWindowLong Me.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If Me.ActiveControl.Name = "dtpDate" Then
            dtpDate.Visible = False
        End If
    ElseIf 39 = KeyAscii Then '������
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strHead As String
    Dim blnTmp As Boolean
    Dim blnHave As Boolean
    Dim strTmp As String
    Dim intTmp As Integer
    
    Call RestoreWinState(Me, App.ProductName)
    
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    
    mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '���˵�ַ�ṹ��¼��
    mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '�����ַ�ṹ��¼��
    mblnUpdate = True

    Call InitBaseInfo
    
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ"))
    mblnID���� = Val(zlDatabase.GetPara(247, glngSys, , 0)) = 1
    mblnFreeInput = Val(zlDatabase.GetPara("��������������ɵ���", glngSys, 0)) = 1
    mbln¼��ҽ��� = Val(zlDatabase.GetPara("������ҽ������¼����ҽ���", glngSys, p����ҽ���´�, "0")) = 1

    '������뷽ʽ
    strTmp = gstr�������
    If Val(Mid(strTmp, 1, 1)) = 0 Then
        strTmp = "1"
    Else
        strTmp = Mid(strTmp, 1, 1)
    End If
    mint������� = Val(strTmp)
    
    mbytSize = zlDatabase.GetPara("����", glngSys, p����ҽ��վ, "0")
    mblnEdit��ͬ��λ = InStr(GetInsidePrivs(p����ҽ��վ), "��Լ���˵Ǽ�") > 0
    '�ȶ��������˵���������Ҫ�ж�
    blnTmp = InStr(GetInsidePrivs(p���ﲡ������), "������д") > 0
    If blnTmp Then
        lblDoctor(1).Tag = IIf(InStr(GetInsidePrivs(p���ﲡ������), "���˲���") > 0, 1, 0)
        mblnDocInput = Val(zlDatabase.GetPara("��ʾ�����������", glngSys, p����ҽ��վ, 0)) = 1
        blnHave = IIf(InStr(GetInsidePrivs(1070), "ǩ��Ȩ") > 0, True, False)
        lblDoctor(1).Caption = UserInfo.����
        lblLink.Visible = True
    Else
        mblnDocInput = False
        blnHave = False
        lblLink.Visible = False
    End If
    
    '��ʼ����ַ�ؼ�
    PatiAddress(PT_�����ص�).Visible = mblnStructAdress
    PatiAddress(PT_���ڵ�ַ).Visible = mblnStructAdress
    PatiAddress(PT_��ͥ��ַ).Visible = mblnStructAdress
    txtE(I�����ص�).Visible = Not mblnStructAdress: cmdE(I�����ص�).Visible = Not mblnStructAdress
    txtE(I���ڵ�ַ).Visible = Not mblnStructAdress: cmdE(I���ڵ�ַ).Visible = Not mblnStructAdress
    txtE(I��ͥ��ַ).Visible = Not mblnStructAdress: cmdE(I��ͥ��ַ).Visible = Not mblnStructAdress
    If mblnStructAdress Then
        PatiAddress(PT_�����ص�).ShowTown = False
        PatiAddress(PT_���ڵ�ַ).ShowTown = mblnShowTown
        PatiAddress(PT_��ͥ��ַ).ShowTown = mblnShowTown
    End If
 
    cmdSign.Visible = blnHave
    lblDoctor(0).Visible = blnHave
    lblDoctor(1).Visible = blnHave
    PicPanel(picPanel_�������).Visible = mblnDocInput

    intTmp = Val(zlDatabase.GetPara("�����������", glngSys, p����ҽ��վ, 0, Array(optInfo(opt����), optInfo(opt���))))
    Call SetInputRoot(0, gint�����Դ, intTmp, optInfo(opt���), optInfo(opt����))
    
    intTmp = Val(zlDatabase.GetPara("����������Դ", glngSys, p����ҽ��վ, 0, Array(optInfo(optҩƷĿ¼), optInfo(opt����Դ))))
    If Not gobjPass Is Nothing Then
        Call SetInputRoot(2, gint����������Դ, intTmp, optInfo(optҩƷĿ¼), optInfo(opt����Դ))
    Else
        Call SetInputRoot(intTmp, 1, intTmp, optInfo(optҩƷĿ¼), optInfo(opt����Դ))
    End If
    
    strHead = "������,3100,1;������Ӧ,3800,1;����ʱ��,1100,4;����Դ����;ҩ��ID;������Դ"
    Call InitVSFlexGrid(vsAller, strHead)
    
    strHead = ",460,4;����;��ϱ���,840,4;�������,4000,1;��ҽ֤��;����ʱ��,1600,1;��ע;ICD����;����,460,4;,270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
    Call InitVSFlexGrid(vsDiagXY, strHead, "0,��ҽ,18,1", 1, 1)
    
    strHead = ",460,4;����;��ϱ���,840,4;�������,2700,1;��ҽ֤��,1300,1;����ʱ��,1600,1;��ע;ICD����;����,460,4;,270,4;,270,4;���ID;����ID;֤��ID;ҽ��IDs;��Ϸ���;�̶�����;����ID;�����Դ;��������;�������;֤�����;��¼����;��¼��Ա"
    Call InitVSFlexGrid(vsDiagZY, strHead, "0,��ҽ,18,11", 1, 0)
    
    Call InitEditData
    
    vsc.Max = 600
    vsc.Min = 0
    vsc.LargeChange = 100
    Call SetCtlBackColor
    
    Call SetCtlPos(0)
    Call SetCtlPos(1)
    Call SetCtlPos(2)
    If mint���� = 1 Then
        stbThis.Visible = True
    Else
        stbThis.Visible = False
    End If
    Set mobjCtl = Nothing
    mblnCboNoClick = False
    mblnOK = False
End Sub

Private Sub LoadDiagData()
'���ܣ����������Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    Set rsTmp = GetPatiDiagData(False)
    
    'ɾ��֮ǰ�Ļ���
    mrsSecdInfo.Filter = "�ؼ���='vsDiagXY' or �ؼ���='vsDiagZY'"
    If Not mrsSecdInfo.EOF Then
        For i = 1 To mrsSecdInfo.RecordCount
            mrsSecdInfo.Delete
            mrsSecdInfo.Update
            mrsSecdInfo.MoveNext
        Next
    End If
    Call LoadVsDiagData(vsDiagXY, rsTmp, "1")
    Call LoadVsDiagData(vsDiagZY, rsTmp, "11")
End Sub

Private Function FilterDiagByType(ByRef rsInput As ADODB.Recordset, ByVal intDiagType As Integer) As Boolean
'���ܣ���ϵĹ���
'���أ�true-��ҳ���
    rsInput.Filter = "��¼��Դ=3 And �������=" & intDiagType
    FilterDiagByType = Not rsInput.EOF
    If rsInput.EOF Then rsInput.Filter = "��¼��Դ=2 And �������=" & intDiagType
    If rsInput.EOF Then rsInput.Filter = "��¼��Դ=1 And �������=" & intDiagType
    If rsInput.EOF Then rsInput.Filter = "��¼��Դ=4 And �������=" & intDiagType
End Function

Private Sub LoadVsDiagData(ByRef vsDiagInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal strDiagType As String)
'���ܣ�����ϼ��ص�����в��һ���
'������vsDiagInput=��Ҫ������ϵı��
'      rsInput=��ȡ����ϼ�¼��
'      strDiagType=��������ַ������������Զ��ŷָ�
'      blnOnlyCache=�Ƿ�ֻ�������ݣ�True-���澭�����󻺴棬False-�����ʼ���ػ���
'˵����LoadMedPageData���Ӻ���

    Dim strTmp As String
    Dim i As Long, j As Long, k As Long, lngRow As Long
    Dim bln�ֻ��̶� As Boolean
    Dim bln��ҽ As Boolean
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String
    Dim arrWhole As Variant, arrMain As Variant
    Dim blnFreeDiag As Boolean
    Dim lngTmp As Long
    
    On Error GoTo errH
    With vsDiagInput
        bln��ҽ = vsDiagInput.Name = "vsDiagXY"
        .Rows = .FixedRows
        
        .Rows = .FixedRows + 1
        .TextMatrix(.Rows - 1, 0) = IIf(bln��ҽ, "��ҽ", "��ҽ")
        .TextMatrix(.Rows - 1, DI_��Ϸ���) = IIf(bln��ҽ, 1, 11)
        If Not FilterDiagByType(rsInput, Val(strDiagType)) Then
            .Tag = "1"
        Else
            .Tag = ""
        End If
        If bln��ҽ Then
            mstrTagDiagXY = .Tag
        Else
            mstrTagDiagZY = .Tag
        End If
        .Tag = ""
        
        Do While Not rsInput.EOF
            'ȷ����ǰ��ʾ��
            lngRow = .FindRow(strDiagType, , DI_��Ϸ���, , True)
            For j = lngRow To .Rows - 1
                If Val(.TextMatrix(j, DI_��Ϸ���)) = Val(strDiagType) Then
                    lngRow = j
                    If .TextMatrix(j, DI_�������) = "" Then Exit For
                Else
                    Exit For
                End If
            Next
            '������
            If .TextMatrix(lngRow, DI_�������) <> "" Then
                lngRow = lngRow + 1: .AddItem "", lngRow
                .TextMatrix(lngRow, DI_��Ϸ���) = strDiagType
            End If
            
            strTmp = rsInput!������� & ""
            If Not (IsNull(rsInput!���id) And IsNull(rsInput!����id)) Then
                '��ȡ��ϱ��룬�������Ϊ(����)��������(����)����(֤��) ���͵Ŀ��Ի�ȡ�������
                If strTmp Like "(?*)?*" Then
                    lngPos = InStr(1, strTmp, ")")
                    .TextMatrix(lngRow, DI_��ϱ���) = Mid(strTmp, 2, lngPos - 2)
                    strTmp = Mid(strTmp, lngPos + 1)
                End If
            End If
            If .TextMatrix(lngRow, DI_��ϱ���) = "" And Not (IsNull(rsInput!���id) And IsNull(rsInput!����id)) Then
                '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                .TextMatrix(lngRow, DI_��ϱ���) = IIf(Not IsNull(rsInput!����id), rsInput!�������� & "", rsInput!��ϱ��� & "")
            End If
            '��ȡ��ҽ֤����������������ܻ�����ǰ��׺��ǰ��׺�������ţ����Է����ȡ�ַ���
            If strTmp Like "?*(?*)" And Not bln��ҽ Then
                strTmp = StrReverse(strTmp)
                lngPos = InStr(1, strTmp, "(")
                .TextMatrix(lngRow, DI_��ҽ֤��) = StrReverse(Mid(strTmp, 2, lngPos - 2))
                strTmp = StrReverse(Mid(strTmp, lngPos + 1))
            End If
            'ȡ�������
            .TextMatrix(lngRow, DI_�������) = strTmp
            '��������ı�������
            If Not (IsNull(rsInput!���id) And IsNull(rsInput!����id)) Then
                .Cell(flexcpData, lngRow, DI_�������) = IIf(Not IsNull(rsInput!����id), rsInput!�������� & "", rsInput!������� & "")
            Else
                .Cell(flexcpData, lngRow, DI_�������) = .TextMatrix(lngRow, DI_�������)
            End If
             .Cell(flexcpData, lngRow, DI_��ϱ���) = .TextMatrix(lngRow, DI_��ϱ���)
             .Cell(flexcpData, lngRow, DI_��ҽ֤��) = .TextMatrix(lngRow, DI_��ҽ֤��)
            '���������ݼ���
            .TextMatrix(lngRow, DI_����ʱ��) = Format(rsInput!����ʱ�� & "", "YYYY-MM-DD HH:mm")
            .TextMatrix(lngRow, DI_��ע) = rsInput!��ע & ""
            .TextMatrix(lngRow, DI_ICD����) = rsInput!���� & ""
            .TextMatrix(lngRow, DI_�Ƿ�����) = IIf(Val(rsInput!�Ƿ����� & "") = 1, "��", "")
            .TextMatrix(lngRow, DI_���ID) = rsInput!���id & ""
            .TextMatrix(lngRow, DI_����ID) = rsInput!����id & ""
            .TextMatrix(lngRow, DI_֤��ID) = rsInput!֤��id & ""
            .TextMatrix(lngRow, DI_ҽ��IDs) = rsInput!ҽ��ID & ""
            .TextMatrix(lngRow, DI_�̶�����) = IIf(IsNull(rsInput!����), "", "1")
            .TextMatrix(lngRow, DI_����ID) = IIf(IsNull(rsInput!����), "0", "1")
            .TextMatrix(lngRow, DI_�����Դ) = Val(rsInput!��¼��Դ & "") '�����¼��Դ���Ա㱣��ʱ������Ϊ��ҳ�򲡰���Դ
            .TextMatrix(lngRow, DI_��������) = rsInput!�������� & ""
            .TextMatrix(lngRow, DI_�������) = rsInput!������� & ""
            .TextMatrix(lngRow, DI_֤�����) = rsInput!֤����� & ""
            .TextMatrix(lngRow, DI_��¼����) = Format(rsInput!��¼���� & "", "YYYY-MM-DD HH:mm")
            .TextMatrix(lngRow, DI_��¼��Ա) = rsInput!��¼�� & ""
            .RowData(lngRow) = Val(rsInput!ID & "")
            rsInput.MoveNext
        Loop
   
        '������������Ϣ
        .Cell(flexcpForeColor, .FixedRows, DI_�Ƿ�����, .Rows - 1, DI_�Ƿ�����) = vbRed
        .Cell(flexcpBackColor, .FixedRows, DI_��ϱ���, .Rows - 1, DI_��ϱ���) = GRD_UNEDITCELL_COLOR      '����ɫ
        .Cell(flexcpText, .FixedRows, DI_�������, .Rows - 1, DI_�������) = IIf(bln��ҽ, "��ҽ", "��ҽ")
        
        '���ݻ���
        lngTmp = 1
        strTmp = ""
        arrMain = Array(DI_��ϱ���, DI_��Ϸ���, DI_���ID, DI_����ID)
        arrWhole = Array(DI_��Ϸ���, DI_��������, DI_��ϱ���, DI_ICD����, DI_�������, DI_֤�����, DI_��ҽ֤��, DI_�Ƿ�����, DI_���ID, DI_����ID, DI_�������, DI_��ע, DI_����ʱ��)
        For i = .FixedRows To .Rows - 1
            blnFreeDiag = Val(.TextMatrix(i, DI_���ID)) = 0 And Val(.TextMatrix(i, DI_����ID)) = 0 '����¼�����
            If .TextMatrix(i, DI_�������) <> "" Then
                If strTmp <> .TextMatrix(i, DI_��Ϸ���) Then
                    j = 1: strTmp = .TextMatrix(i, DI_��Ϸ���)
                Else
                    j = j + 1
                End If
                strInfo = j: strMainInfo = ""
                For k = LBound(arrWhole) To UBound(arrWhole)
                    strInfo = strInfo & "|" & .TextMatrix(i, arrWhole(k))
                Next
                For k = LBound(arrMain) To UBound(arrMain)
                    strMainInfo = strMainInfo & "|" & .TextMatrix(i, arrMain(k))
                Next

                If blnFreeDiag Then strMainInfo = strMainInfo & "|" & .TextMatrix(i, DI_�������) '����¼����ϼ����������
                mrsSecdInfo.AddNew Array("���", "ԭID", "�ؼ���", "��Ϣԭֵ", "����Ϣԭֵ", "Tag", "��Ϣ��ֵ", "����Ϣ��ֵ"), Array(lngTmp, Val(.RowData(i)), vsDiagInput.Name, strInfo, strMainInfo, IIf(.TextMatrix(i, DI_�����Դ) = "", 3, .TextMatrix(i, DI_�����Դ)), Null, Null)
                lngTmp = lngTmp + 1
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

Private Sub LoadAllerData()
'���ܣ����ع�����Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lngRow As Long, j As Long
    Dim strInfo As String 'ȫ��Ϣ
    Dim strMainInfo As String '������Ϣ
    Dim str����ʱ�� As String
    Dim lngTmp As Long
    Dim strAll As String
    Dim blnTmp As Boolean
    blnTmp = mblnNoSave
    mblnNoSave = True
    On Error GoTo errH
    chkNoAller.Value = 0
    mrsMainInfo.Filter = "��Ϣ��='�޹�����¼'"
    If Val(mrsMainInfo!��Ϣԭֵ & "") = 1 Then
        chkNoAller.Value = 1
        mblnNoSave = blnTmp
        Exit Sub
    End If
    mblnNoSave = blnTmp
 
    strSQL = "Select a.ID,a.��¼��Դ,a.����ʱ��,a.ҩ��id,a.ҩ����,a.������Ӧ,a.����Դ����" & vbNewLine & _
        "From ���˹�����¼ A" & vbNewLine & _
        "Where a.��� = 1 And a.����id =[1] And a.��ҳid =[2] And Not Exists" & vbNewLine & _
        " (Select b.ҩ��id From ���˹�����¼ b" & vbNewLine & _
        "       Where (Nvl(b.ҩ��id, 0) = Nvl(A.ҩ��id, 0) Or Nvl(ҩ����, 'Null') = Nvl(A.ҩ����, 'Null')) And Nvl(���, 0) = 0 And" & vbNewLine & _
        "             b.��¼ʱ�� > A.��¼ʱ�� And b.����id =[1] And b.��ҳid =[2])" & vbNewLine & _
        "Order By Nvl(a.����ʱ��,a.��¼ʱ��) Desc,a.��¼ʱ�� desc, a.ҩ����"

    If mblnMoved Then
        strSQL = Replace(strSQL, "���˹�����¼", "H���˹�����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ҳ��ȡ������Ϣ", mlng����ID, mlng�Һ�ID)
    
    mstrTagAller = ""
    rsTmp.Filter = "��¼��Դ=3"
    If rsTmp.EOF Then
        rsTmp.Filter = "��¼��Դ<>3"
        mstrTagAller = "1" '����ҳ��Դ
    End If
    
    'ɾ��֮ǰ�Ļ���
    mrsSecdInfo.Filter = "�ؼ���='vsAller'"
    If Not mrsSecdInfo.EOF Then
        For i = 1 To mrsSecdInfo.RecordCount
            mrsSecdInfo.Delete
            mrsSecdInfo.Update
            mrsSecdInfo.MoveNext
        Next
    End If
    
    '������ʷ����103674
    If rsTmp.RecordCount > 0 Then
        chkNoAller.Value = 0
        Call SetAllerEdit(False)
        Call UpDateInfo(0, "chkNoAller")
    End If
    
    lngTmp = 1
    With vsAller
        .Rows = .FixedRows
        For i = 1 To rsTmp.RecordCount
            '������Դ�Ŀ������ظ� Ψһ����ҩ��ID��ҩ����������Դ���룬����ʱ��
            str����ʱ�� = Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd")
            strMainInfo = Val(rsTmp!ҩ��ID & "") & "|" & rsTmp!ҩ���� & "|" & rsTmp!����Դ���� & "|" & str����ʱ��
            
            If InStr("," & strAll & ",", "," & strMainInfo & ",") = 0 Then
                strAll = strAll & "," & strMainInfo
                .Rows = .Rows + 1: lngRow = .Rows - 1
                .TextMatrix(lngRow, AI_����ʱ��) = str����ʱ��
                .TextMatrix(lngRow, AI_����ҩ��) = NVL(rsTmp!ҩ����)
                .TextMatrix(lngRow, AI_������Ӧ) = NVL(rsTmp!������Ӧ)
                .TextMatrix(lngRow, AI_����Դ����) = NVL(rsTmp!����Դ����)
                .TextMatrix(lngRow, AI_ҩ��ID) = Val(rsTmp!ҩ��ID & "")
                .TextMatrix(lngRow, AI_������Դ) = rsTmp!��¼��Դ & ""
                '���ݱ��ݴ洢
                .Cell(flexcpData, lngRow, AI_����ҩ��) = .TextMatrix(lngRow, AI_����ҩ��)
                .RowData(lngRow) = Val(rsTmp!ID & "")
                
                strInfo = strMainInfo & "|" & rsTmp!������Ӧ
                mrsSecdInfo.AddNew Array("���", "ԭID", "�ؼ���", "��Ϣԭֵ", "����Ϣԭֵ"), Array(lngTmp, Val(rsTmp!ID & ""), "vsAller", strInfo, strMainInfo)
                lngTmp = lngTmp + 1
            End If
            rsTmp.MoveNext
        Next
        .Rows = .Rows + 1 '����һ�п���
        .Row = 1: .Col = AI_����ҩ��
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Function GetAllerSaveSQL(ByRef arrSQL As Variant) As Boolean
'���ܣ���ȡ����ҩ�����SQL
    Dim i As Long
    Dim lng״̬ As Long
    Dim strTmp As String
    Dim strInfo As String
    Dim strMainInfo As String
    Dim lngRow As Long
    Dim lngTmp As Long
    Dim strAll As String
    Dim strDels As String
    
    On Error GoTo errH
    
    arrSQL = Array()
    
    With vsAller
        .Tag = ""
        lngTmp = 1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, AI_����ҩ��) <> "" Then
                strMainInfo = Val(.TextMatrix(i, AI_ҩ��ID)) & "|" & .TextMatrix(i, AI_����ҩ��) & "|" & .TextMatrix(i, AI_����Դ����) & "|" & .TextMatrix(i, AI_����ʱ��)
                strInfo = strMainInfo & "|" & .TextMatrix(i, AI_������Ӧ)
 
                If InStr("," & strAll & ",", "," & strMainInfo & ",") > 0 Then
                    '��ͬ��ÿ��¼
                    .Tag = i
                    .Cell(flexcpBackColor, i, .FixedCols, i, AI_������Ӧ) = &HC0C0FF
                    Call .ShowCell(i, AI_����ҩ��)
                    Exit Function
                Else
                    strAll = strAll & "," & strMainInfo '�ռ�������������ж��Ƿ����ظ���
                End If
                
                mrsSecdInfo.Filter = "�ؼ���='vsAller' and ���=" & lngTmp
                If mrsSecdInfo.EOF Then
                    mrsSecdInfo.AddNew
                    mrsSecdInfo!��� = lngTmp
                    mrsSecdInfo!�ؼ��� = "vsAller"
                End If
                mrsSecdInfo!��ID = Val(.RowData(i))
                mrsSecdInfo!��Ϣ��ֵ = strInfo
                mrsSecdInfo!����Ϣ��ֵ = strMainInfo
                mrsSecdInfo!IndexEx = i
                mrsSecdInfo.Update
                lngTmp = lngTmp + 1
 
                mrsSecdInfo.Filter = 0
            End If
        Next
    
        mrsSecdInfo.Filter = "�ؼ���='vsAller'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng״̬ = CS_δ�ı�
            If mrsSecdInfo!��Ϣԭֵ & "" <> mrsSecdInfo!��Ϣ��ֵ & "" Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣԭֵ) Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣ��ֵ) Then
                lng״̬ = CS_ɾ����
            End If
            If lng״̬ = CS_������ And mrsSecdInfo!����Ϣԭֵ & "" <> mrsSecdInfo!����Ϣ��ֵ & "" Then
                lng״̬ = CS_�滻��
            End If
            mrsSecdInfo.Update "�ı�״̬", lng״̬
            mrsSecdInfo.MoveNext
        Next
   
        'ɾ�����Լ�����Ϣ�ı�����Ҫ����ɾ������
        mrsSecdInfo.Filter = "(�ı�״̬=" & CS_ɾ���� & " And �ؼ���='vsAller') OR (�ı�״̬=" & CS_�滻�� & " And �ؼ���='vsAller')"
        Do While Not mrsSecdInfo.EOF
            strDels = strDels & "," & mrsSecdInfo!ԭID
            mrsSecdInfo.MoveNext
        Loop
        
        '����Ϣ�ı��Լ���������Ҫ���ò������        '�μ���Ϣ�ı䣬���ø��¹���
        mrsSecdInfo.Filter = "�ؼ���='vsAller' And �ı�״̬>" & CS_δ�ı�
        
        If mstrTagAller = "1" Then
            '����޸��˹�����¼��������Դ�Ĺ�����¼����һ�ݡ�
            If (strDels <> "" Or Not mrsSecdInfo.EOF) Then
                For lngRow = .FixedRows To .Rows - 1
                    If .TextMatrix(lngRow, AI_����ҩ��) <> "" And .TextMatrix(lngRow, AI_����ҩ��) <> "��" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "zl_���˹�����¼_Insert(" & mlng����ID & "," & mlng�Һ�ID & "," & _
                            "3," & ZVal(.TextMatrix(lngRow, AI_ҩ��ID)) & ",'" & .TextMatrix(lngRow, AI_����ҩ��) & "',1," & _
                            ToDateOracle(.TextMatrix(lngRow, AI_����ʱ��), "ymd") & ",SysDate,'" & _
                            .TextMatrix(lngRow, AI_������Ӧ) & "','" & .TextMatrix(lngRow, AI_����Դ����) & "')"
                    End If
                Next
            End If
        Else
            If strDels <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_���˹�����¼_Delete(" & mlng����ID & "," & mlng�Һ�ID & ",3,'" & Mid(strDels, 2) & "')"
            End If
            Do While Not mrsSecdInfo.EOF
                lngRow = mrsSecdInfo!IndexEx
                If .TextMatrix(lngRow, AI_����ҩ��) <> "��" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mrsSecdInfo!�ı�״̬ <> CS_������ Then
                        arrSQL(UBound(arrSQL)) = "zl_���˹�����¼_Insert(" & mlng����ID & "," & mlng�Һ�ID & "," & _
                                "3," & ZVal(.TextMatrix(lngRow, AI_ҩ��ID)) & ",'" & .TextMatrix(lngRow, AI_����ҩ��) & "',1," & _
                                ToDateOracle(.TextMatrix(lngRow, AI_����ʱ��), "ymd") & ",SysDate,'" & _
                                .TextMatrix(lngRow, AI_������Ӧ) & "','" & .TextMatrix(lngRow, AI_����Դ����) & "')"
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_���˹�����¼_Update(" & mrsSecdInfo!ԭID & "," & mlng����ID & "," & mlng�Һ�ID & "," & _
                                "3," & ZVal(.TextMatrix(lngRow, AI_ҩ��ID)) & ",'" & .TextMatrix(lngRow, AI_����ҩ��) & "',1," & _
                                ToDateOracle(.TextMatrix(lngRow, AI_����ʱ��), "ymd") & ",'" & _
                                .TextMatrix(lngRow, AI_������Ӧ) & "','" & .TextMatrix(lngRow, AI_����Դ����) & "')"
                    End If
                End If
                mrsSecdInfo.MoveNext
            Loop
        End If
    End With
    If UBound(arrSQL) <> -1 Then GetAllerSaveSQL = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function GetDiagSaveSQL(ByRef vsDiagInput As VSFlexGrid, ByRef arrSQL As Variant) As Boolean
'���ܣ���ȡ��ϱ����SQL
    Dim i As Long, j As Long, k As Long
    Dim lng״̬ As Long
    Dim strTmp As String
    Dim strInfo As String
    Dim strMainInfo As String
    Dim lngRow As Long
    Dim strVsName As String
    Dim arrWhole As Variant
    Dim arrOther As Variant
    Dim arrMain As Variant
    Dim blnFreeDiag As Boolean
    Dim datCur As Date
    Dim lngID As Long
    Dim lngTmp As Long
    Dim strAllDiag As String
    Dim strTag As String
    
    On Error GoTo errH
    arrSQL = Array()
    With vsDiagInput
        .Tag = ""
        strVsName = .Name
        strTmp = ""
        lngTmp = 1
        arrMain = Array(DI_��ϱ���, DI_��Ϸ���, DI_���ID, DI_����ID)
        arrWhole = Array(DI_��Ϸ���, DI_��������, DI_��ϱ���, DI_ICD����, DI_�������, DI_֤�����, DI_��ҽ֤��, DI_�Ƿ�����, DI_���ID, DI_����ID, DI_�������, DI_��ע, DI_����ʱ��)
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_�������) <> "" Then
                blnFreeDiag = Val(.TextMatrix(i, DI_���ID)) = 0 And Val(.TextMatrix(i, DI_����ID)) = 0 '����¼�����
                If strTmp <> .TextMatrix(i, DI_��Ϸ���) Then
                    j = 1: strTmp = .TextMatrix(i, DI_��Ϸ���)
                Else
                    j = j + 1
                End If
                strInfo = j: strMainInfo = ""
                For k = LBound(arrWhole) To UBound(arrWhole)
                    strInfo = strInfo & "|" & .TextMatrix(i, arrWhole(k))
                Next
                For k = LBound(arrMain) To UBound(arrMain)
                    strMainInfo = strMainInfo & "|" & .TextMatrix(i, arrMain(k))
                Next
                If blnFreeDiag Then strMainInfo = strMainInfo & "|" & .TextMatrix(i, DI_�������) '����¼����ϼ����������
                
                If strVsName = "vsDiagZY" Then
                    If InStr("," & strAllDiag & ",", "," & strMainInfo & ",") = 0 Then
                        strAllDiag = strAllDiag & "," & strMainInfo '�ռ�������������ж��Ƿ����ظ���
                    End If
                Else
                    If InStr("," & strAllDiag & ",", "," & strMainInfo & ",") > 0 Then
                        '����ͬ���
                        .Tag = i
                        .Cell(flexcpBackColor, i, .FixedCols, i, DI_�Ƿ�����) = &HC0C0FF
                        Call .ShowCell(i, DI_�������)
                        Exit Function
                    Else
                        strAllDiag = strAllDiag & "," & strMainInfo '�ռ�������������ж��Ƿ����ظ���
                    End If
                End If
                mrsSecdInfo.Filter = "�ؼ���='" & strVsName & "' and ���=" & lngTmp
 
                If mrsSecdInfo.EOF Then
                    mrsSecdInfo.AddNew
                    mrsSecdInfo!��� = lngTmp
                    mrsSecdInfo!�ؼ��� = strVsName
                End If
                mrsSecdInfo!��Ϣ��ֵ = strInfo
                mrsSecdInfo!����Ϣ��ֵ = strMainInfo
                mrsSecdInfo!��ҽ��� = .TextMatrix(i, DI_��ҽ֤��)
                mrsSecdInfo!IndexEx = i
                mrsSecdInfo.Update
                lngTmp = lngTmp + 1
                mrsSecdInfo.Filter = 0
            End If
        Next
        
        If strVsName = "vsDiagZY" Then
            If strAllDiag <> "" Then
                arrOther = Split(strAllDiag, ",")
                For i = 0 To UBound(arrOther)
                    strTmp = arrOther(i)
                    If strTmp <> "" Then
                        mrsSecdInfo.Filter = "�ؼ���='vsDiagZY' and ����Ϣ��ֵ='" & strTmp & "'"
                        If mrsSecdInfo.RecordCount > 2 Then
                            strInfo = "|ͬһ��ҽ��ϳ�����2����"
                            '����3����
                            .Tag = mrsSecdInfo!IndexEx & strInfo
                            lngTmp = Val(mrsSecdInfo!IndexEx)
                            .Cell(flexcpBackColor, lngTmp, .FixedCols, lngTmp, DI_�Ƿ�����) = &HC0C0FF
                            Call .ShowCell(lngTmp, DI_�������)
                            Exit Function
                        ElseIf mrsSecdInfo.RecordCount > 1 Then
                            strAllDiag = "��"
                            For j = 1 To mrsSecdInfo.RecordCount
                                If InStr("," & strAllDiag & ",", "," & mrsSecdInfo!��ҽ��� & ",") > 0 Then
                                    strInfo = "|ͬһ��ҽ��϶�Ӧ��������ͬ��֤��"
                                    '����ͬ���
                                    .Tag = mrsSecdInfo!IndexEx & strInfo
                                    lngTmp = Val(mrsSecdInfo!IndexEx)
                                    .Cell(flexcpBackColor, lngTmp, .FixedCols, lngTmp, DI_�Ƿ�����) = &HC0C0FF
                                    Call .ShowCell(lngTmp, DI_�������)
                                    Exit Function
                                Else
                                    strAllDiag = strAllDiag & "," & mrsSecdInfo!��ҽ��� '�ռ�������������ж��Ƿ����ظ���
                                End If
                                mrsSecdInfo.MoveNext
                            Next
                        End If
                    End If
                Next
            End If
        End If
        mrsSecdInfo.Filter = "�ؼ���='" & strVsName & "'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng״̬ = CS_δ�ı�
            If mrsSecdInfo!��Ϣԭֵ & "" <> mrsSecdInfo!��Ϣ��ֵ & "" Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣԭֵ) Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣ��ֵ) Then
                lng״̬ = CS_ɾ����
            End If
            If lng״̬ = CS_������ And mrsSecdInfo!����Ϣԭֵ & "" <> mrsSecdInfo!����Ϣ��ֵ & "" Then
                lng״̬ = CS_�滻��
            End If
            mrsSecdInfo.Update "�ı�״̬", lng״̬
            mrsSecdInfo.MoveNext
        Next
        
        'ɾ�����Լ�����Ϣ�ı�����Ҫ����ɾ������
        mrsSecdInfo.Filter = "(�ı�״̬=" & CS_ɾ���� & " And �ؼ���='" & strVsName & "') OR (�ı�״̬=" & CS_�滻�� & " And �ؼ���='" & strVsName & "')": strTmp = ""
        Do While Not mrsSecdInfo.EOF
            strTmp = strTmp & "," & mrsSecdInfo!ԭID
            mrsSecdInfo.MoveNext
        Loop
        '����Ϣ�ı��Լ���������Ҫ���ò������
        '�μ���Ϣ�ı䣬���ø��¹���
        mrsSecdInfo.Filter = "�ı�״̬>" & CS_δ�ı� & " And �ؼ���='" & strVsName & "'"
        
        datCur = mdatCurDate
         
        If strVsName = "vsDiagXY" Then
            strTag = mstrTagDiagXY
        Else
            strTag = mstrTagDiagZY
        End If
        
        If strTag = "1" Then
            If strTmp <> "" Or Not mrsSecdInfo.EOF Then
                j = 1
                For lngRow = .FixedRows To .Rows - 1
                    If .TextMatrix(lngRow, DI_�������) <> "" Then
                        If Trim(.TextMatrix(lngRow, DI_��ϱ���)) = "" Then
                            strTmp = .TextMatrix(lngRow, DI_�������) & IIf(.TextMatrix(lngRow, DI_��ҽ֤��) <> "", "(" & .TextMatrix(lngRow, DI_��ҽ֤��) & ")", "")
                        Else
                            strTmp = "(" & .TextMatrix(lngRow, DI_��ϱ���) & ")" & .TextMatrix(lngRow, DI_�������) & IIf(.TextMatrix(lngRow, DI_��ҽ֤��) <> "", "(" & .TextMatrix(lngRow, DI_��ҽ֤��) & ")", "")
                        End If
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3,NULL," & .TextMatrix(lngRow, DI_��Ϸ���) & "," & _
                                ZVal(.TextMatrix(lngRow, DI_����ID)) & "," & ZVal(.TextMatrix(lngRow, DI_���ID)) & "," & ZVal(.TextMatrix(lngRow, DI_֤��ID)) & ",'" & _
                                strTmp & "',null,null," & IIf(.TextMatrix(lngRow, DI_�Ƿ�����) = "", 0, 1) & "," & ToDateOracle(datCur, "ymdhms") & ",'" & .TextMatrix(lngRow, DI_ҽ��IDs) & "' ," & j & ",'" & _
                                .TextMatrix(lngRow, DI_��ע) & "',Null," & ToDateOracle(.TextMatrix(lngRow, DI_����ʱ��), "ymdhm") & ")"
                        j = j + 1
                    End If
                Next
            End If
        Else
            If strTmp <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                '����ϵͳ�洢�������ϵͳ���
                arrSQL(UBound(arrSQL)) = "Zl_������ϼ�¼_Delete(" & mlng����ID & "," & mlng�Һ�ID & ",3,NULL,NUll,'" & Mid(strTmp, 2) & "')"
            End If
            
            Do While Not mrsSecdInfo.EOF
                lngRow = mrsSecdInfo!IndexEx: j = Val(Mid(mrsSecdInfo!��Ϣ��ֵ, 1, InStr(mrsSecdInfo!��Ϣ��ֵ, "|") - 1))
                If Trim(.TextMatrix(lngRow, DI_��ϱ���)) = "" Then
                    strTmp = .TextMatrix(lngRow, DI_�������) & IIf(.TextMatrix(lngRow, DI_��ҽ֤��) <> "", "(" & .TextMatrix(lngRow, DI_��ҽ֤��) & ")", "")
                Else
                    strTmp = "(" & .TextMatrix(lngRow, DI_��ϱ���) & ")" & .TextMatrix(lngRow, DI_�������) & IIf(.TextMatrix(lngRow, DI_��ҽ֤��) <> "", "(" & .TextMatrix(lngRow, DI_��ҽ֤��) & ")", "")
                End If
    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If mrsSecdInfo!�ı�״̬ <> CS_������ Then
                    lngID = zlDatabase.GetNextId("������ϼ�¼")
                    arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3,NULL," & .TextMatrix(lngRow, DI_��Ϸ���) & "," & _
                                        ZVal(.TextMatrix(lngRow, DI_����ID)) & "," & ZVal(.TextMatrix(lngRow, DI_���ID)) & "," & ZVal(.TextMatrix(lngRow, DI_֤��ID)) & ",'" & _
                                        strTmp & "',null,null," & IIf(.TextMatrix(lngRow, DI_�Ƿ�����) = "", 0, 1) & "," & ToDateOracle(datCur, "ymdhms") & ",'" & .TextMatrix(lngRow, DI_ҽ��IDs) & "' ," & j & ",'" & .TextMatrix(lngRow, DI_��ע) & "'," & _
                                        "null," & ToDateOracle(.TextMatrix(lngRow, DI_����ʱ��), "ymdhm") & ",Null," & lngID & ")"
                Else
                    arrSQL(UBound(arrSQL)) = "Zl_������ϼ�¼_Update(" & mrsSecdInfo!ԭID & "," & mlng����ID & "," & mlng�Һ�ID & ",3," & .TextMatrix(lngRow, DI_��Ϸ���) & "," _
                                        & ZVal(.TextMatrix(lngRow, DI_����ID)) & "," & ZVal(.TextMatrix(lngRow, DI_���ID)) & "," & ZVal(.TextMatrix(lngRow, DI_֤��ID)) & ",'" & _
                                        strTmp & "',null,null," & IIf(.TextMatrix(lngRow, DI_�Ƿ�����) = "", 0, 1) & "," & j & ",'" & .TextMatrix(lngRow, DI_��ע) & "',null," & ToDateOracle(.TextMatrix(lngRow, DI_����ʱ��), "ymdhm") & ")"
                End If
                mrsSecdInfo.MoveNext
            Loop
        End If
    End With
    If UBound(arrSQL) <> -1 Then GetDiagSaveSQL = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function ToDateOracle(ByVal strDate As String, ByVal strType As String) As String
'���ܣ���ȡORACLE Date���ʹ�
'������strDate=ʱ���ַ���
'      strType=��ʽ�ַ������ͣ�ymd-�����գ�yyyy-mm-dd)��ymdhm-��yyyy-mm-dd hh:mm),ymdhms-��yyyy-mm-dd hh:mm:ss)
    If Not IsDate(strDate) Then ToDateOracle = "Null": Exit Function
    Select Case UCase(strType)
        Case "YMD"
           ToDateOracle = "To_Date('" & Format(strDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "YMDHM"
           ToDateOracle = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Case "YMDHMS"
           ToDateOracle = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
           ToDateOracle = "Null"
    End Select

End Function

Private Sub UpDateAller()
'���ܣ����������¼
    Dim i As Long
    Dim blnTrans As Boolean
    Dim arrSQL As Variant
    Dim blnUpdate As Boolean
    On Error GoTo errH
    
    If mblnNoSave Then Exit Sub
    
    If Not GetAllerSaveSQL(arrSQL) Then
        '������ֻ���
        mrsSecdInfo.Filter = "�ؼ���='vsAller'"
        For i = 1 To mrsSecdInfo.RecordCount
            Call mrsSecdInfo.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ"), Array(Null, Null))
            mrsSecdInfo.MoveNext
        Next
        Exit Sub
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        blnUpdate = True
    Next
    gcnOracle.CommitTrans: blnTrans = False
    If blnUpdate Then
        mblnOK = True
        Call LoadAllerData
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub UpDateDiag(ByRef vsDiagInput As VSFlexGrid)
'���ܣ�������ϼ�¼
    Dim i As Long
    Dim blnTrans As Boolean
    Dim arrSQL As Variant
    Dim str����ID As String
    Dim str���ID As String
    Dim lngRowXY As Long
    Dim lngColXY As Long
    Dim lngRowZY As Long
    Dim lngColZY As Long
    
    On Error GoTo errH
    
    If mblnNoSave Then Exit Sub
    If mbln��������� Then Exit Sub
    
    If Not GetDiagSaveSQL(vsDiagInput, arrSQL) Then
        '������ֻ���
        mrsSecdInfo.Filter = "�ؼ���='" & vsDiagInput.Name & "'"
        For i = 1 To mrsSecdInfo.RecordCount
            Call mrsSecdInfo.Update(Array("��Ϣ��ֵ", "����Ϣ��ֵ", "��ҽ���"), Array(Null, Null, ""))
            mrsSecdInfo.MoveNext
        Next
        Exit Sub
    End If
    Call MsgDis
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    mblnOK = True
    
    lngRowXY = vsDiagXY.Row
    lngColXY = vsDiagXY.Col
    lngRowZY = vsDiagZY.Row
    lngColZY = vsDiagZY.Col
    
    Call LoadDiagData
    
    With vsDiagXY
        If lngRowXY < .Rows And lngRowXY >= .FixedRows Then .Row = lngRowXY
        If lngColXY < .Cols And lngColXY >= .FixedCols Then .Col = lngColXY
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, DI_���ID)) <> 0 Then
                If InStr("," & str���ID & ",", "," & Val(.TextMatrix(i, DI_���ID)) & ",") = 0 Then
                    str���ID = str���ID & "," & Val(.TextMatrix(i, DI_���ID))
                End If
            End If
            If Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                If InStr("," & str����ID & ",", "," & Val(.TextMatrix(i, DI_����ID)) & ",") = 0 Then
                    str����ID = str����ID & "," & Val(.TextMatrix(i, DI_����ID))
                End If
            End If
        Next
    End With
    
    If mbln��ҽ Then
        With vsDiagZY
            If lngRowZY < .Rows And lngRowZY >= .FixedRows Then .Row = lngRowZY
            If lngColZY < .Cols And lngColZY >= .FixedCols Then .Col = lngColZY
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, DI_���ID)) <> 0 Then
                    If InStr("," & str���ID & ",", "," & Val(.TextMatrix(i, DI_���ID)) & ",") = 0 Then
                        str���ID = str���ID & "," & Val(.TextMatrix(i, DI_���ID))
                    End If
                End If
                If Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                    If InStr("," & str����ID & ",", "," & Val(.TextMatrix(i, DI_����ID)) & ",") = 0 Then
                        str����ID = str����ID & "," & Val(.TextMatrix(i, DI_����ID))
                    End If
                End If
            Next
        End With
    End If
    str����ID = Mid(str����ID, 2): str���ID = Mid(str���ID, 2)
    If str����ID <> "" Or str���ID <> "" Then RaiseEvent UpdateDiagInfo(str����ID, str���ID, "")

    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Function GetPatiDiagData(ByVal blnLast As Boolean) As ADODB.Recordset
'���ܣ���ȡ���˵���ϼ�¼���Լ�¼����ʽ���ط�����ص������
'������blnLast �Ƿ������һ�ξ���
    Dim strSQL As String, strSQLTmp As String, strDiagType As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If blnLast Then
    '���һ�εľ���ID
        strSQLTmp = "(Select Max(ID) As ��ҳid" & vbNewLine & _
            "From ���˹Һż�¼" & vbNewLine & _
            "Where ����id = [1] And ��¼���� = 1 And ��¼״̬ = 1 And" & vbNewLine & _
            "      �Ǽ�ʱ�� =" & vbNewLine & _
            "      (Select Max(A.�Ǽ�ʱ��)" & vbNewLine & _
            "       From ���˹Һż�¼ A" & vbNewLine & _
            "       Where A.����id = [1] And A.��¼���� = 1 And A.��¼״̬ = 1 And A.�Ǽ�ʱ�� < (Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID = [2])))"
    End If
    
    '���ö�ȡ��ϵ�����Լ������Դ
    If mbln��ҽ Then
        strDiagType = " And A.��¼��Դ IN(1,3) And A.������� IN(1,11) "
    Else
        strDiagType = " And A.��¼��Դ IN(1,3) And A.�������=1 "
    End If

    '��װSQL,���Ӳ������Ĳ��ò�ѯҽ����¼
    strSQL = "Select A.��ע, A.Id, A.����id, A.��ҳid, A.ҽ��id, A.��¼��Դ, A.��ϴ���, A.�������, A.�������, A.��Ժ����, A.����id, A.���id, A.֤��id,B.���� ��������,C.���� �������,D.���� ֤������," & vbNewLine & _
            "       A.�������, A.��Ժ���, A.�Ƿ�δ��, A.�Ƿ�����, A.����ʱ��, B.���� As ��������,B.��� As ������� , C.���� As ��ϱ���, D.���� As ֤�����,B.����," & vbNewLine & _
            "(Select F_List2str(Cast(Collect(C.ҽ��id|| '') As T_Strlist)) ҽ��id From �������ҽ�� C Where C.���id = A.Id) As ҽ��id," & _
            "B.�Ա�����, B.��Ч����, B.����, B.����, E.Id As ����, E.�Ƿ���,A.��¼����,A.��¼�� " & vbNewLine & _
            "From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C, ��������Ŀ¼ D,����������� E" & vbNewLine & _
            "Where A.����id = B.Id(+) And A.���id = C.Id(+) And A.֤��id = D.Id(+)  And  B.����id = E.Id(+)" & strDiagType & "And A.ȡ��ʱ�� Is Null And A.������� Is Not Null And ����id = [1] And ��ҳid =[2]" & strSQLTmp & vbNewLine & _
            "Order By A.�������, A.��¼��Դ Desc, A.��ϴ���, A.�������, A.Id"
    If mblnMoved Then
        strSQL = Replace(strSQL, "������ϼ�¼", "H������ϼ�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ҳ���", mlng����ID, mlng�Һ�ID)
    Set GetPatiDiagData = rsTmp
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub InitVSFlexGrid(ByRef vsInput As VSFlexGrid, ByVal strHead As String, Optional ByVal strRowContent As String, Optional ByVal intFixedCols As Integer, Optional ByVal intFixedRows As Integer = 1)
'���ܣ���ʼ��������ݣ����ڴ�����Ի����ûָ�֮ǰ
'������vsInput=Ҫ���ø�ʽ�ı��
'          strHead=�����и�ʽ����ʽΪ���б���1,�п�1,����1,��������1,��ʽ��1;�б���2,�п�2,����2,��������2,��ʽ��2.....
'          strRowContent=����Ԥ����������,��ʽΪ����1,����1,��2,����2:��1;��1,����1,��2,����2:��2;....(��Ҫ��С�������У�:��1
    Dim i As Integer, lngRow As Long, j As Long
    Dim arrHead As Variant, arrCol As Variant, arrRow As Variant
    Dim arrTmp As Variant
    On Error GoTo errH
    '������
    With vsInput
        If strHead <> "" Then
            arrHead = Split(strHead, ";")
            .Clear: .Cols = 0: .Rows = 0
            .Rows = intFixedRows + 1: .Cols = UBound(arrHead) + 1
            .FixedRows = intFixedRows: .FixedCols = intFixedCols
            For i = LBound(arrHead) To UBound(arrHead)
                arrCol = Split(arrHead(i), ",")
                .FixedAlignment(i) = 4
                If intFixedRows <> 0 Then .TextMatrix(0, i) = arrCol(0)
                If UBound(arrCol) > 0 Then
                    .ColWidth(i) = Val(arrCol(1))
                Else
                    .ColHidden(i) = True
                End If
                If UBound(arrCol) > 1 Then .ColAlignment(i) = Val(arrCol(2))
                If UBound(arrCol) > 2 Then .ColDataType(i) = Val(arrCol(3))
                If UBound(arrCol) > 3 Then .ColFormat(i) = arrCol(4)
                If UBound(arrCol) > 4 Then .ColHidden(i) = Val(arrCol(5))
            Next
        End If
        '���ý�����
        If strRowContent <> "" Then
            .Rows = .FixedRows
            lngRow = .FixedRows - 1: arrRow = Split(strRowContent, ";")
            For i = LBound(arrRow) To UBound(arrRow)
                arrTmp = Split(arrRow(i), ";")
                'ȷ���к�
                lngRow = lngRow + 1
                If UBound(arrTmp) > 0 Then lngRow = Val(arrTmp(1))
                .Rows = lngRow + 1 '��������
                '��������
                arrCol = Split(arrTmp(0), ",")
                For j = LBound(arrCol) To UBound(arrCol) Step 2
                    .TextMatrix(lngRow, Val(arrCol(j))) = arrCol(j + 1)
                Next
            Next
        End If
    End With
    vsAller.ExtendLastCol = False
    Exit Sub
errH:
    Debug.Print err.Source & "-InitVSFlexGrid:" & err.Description
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub MovePanel()

    PicPanel(picPanel_������Ϣ).Top = mlngTopVsc
    
    If mblnDocInput Then
        PicPanel(picPanel_�������).Top = PicPanel(picPanel_������Ϣ).Height + PicPanel(picPanel_������Ϣ).Top
        PicPanel(picPanel_������Ϣ).Top = PicPanel(picPanel_�������).Height + PicPanel(picPanel_�������).Top
    Else
        PicPanel(picPanel_������Ϣ).Top = PicPanel(picPanel_������Ϣ).Height + PicPanel(picPanel_������Ϣ).Top
    End If
    
    dtpDate.Top = PicPanel(picPanel_������Ϣ).Top - dtpDate.Height + txtE(I����ʱ��).Top - 20
    
    dtpDate.Left = PicPanel(picPanel_������Ϣ).Left + txtE(I����ʱ��).Left
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intTmp As Integer
    
    intTmp = IIf(optInfo(opt����).Value, 1, 0)
    Call zlDatabase.SetPara("�����������", intTmp, glngSys, p����ҽ��վ, InStr(gstrPrivs, "��������") > 0)
    
    intTmp = IIf(optInfo(opt����Դ).Value, 2, 1)
    Call zlDatabase.SetPara("����������Դ", intTmp, glngSys, p����ҽ��վ, gint����������Դ = 0 And gbytPass = 3 And InStr(gstrPrivs, "��������") > 0)
    Call SavePreItem
    Call ClearPatiInfo
    Set mobjKernel = Nothing
    Set mobjPatient = Nothing
    Set mobjCtl = Nothing
    Set mclsZip = Nothing
    Set mclsUnZip = Nothing
    Set mrsMainInfo = Nothing
    Set mrsSecdInfo = Nothing
    Set mrsPreEditCtl = Nothing
    mblnCboNoClick = False
    mlng�Һ�ID = 0
End Sub

Private Sub txtSL_GotFocus()
    Call zlControl.TxtSelAll(txtSL)
End Sub

Private Sub txtSL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtSL_Validate(Cancel As Boolean)
    Dim datCur As Date, datRes As Date
    
    If Trim(txtSL.Text) <> "" Then
        If IsNumeric(txtSL.Text) Then
            If Val(txtSL.Text) <= 0 Then
                MsgBox "����ʱ������ֵ����Ϊ������", vbInformation, gstrSysName
                txtSL.SetFocus: Exit Sub
            End If
        Else
            MsgBox "����ʱ������ֵ����Ϊ���֡�", vbInformation, gstrSysName
            txtSL.Text = "": txtSL.SetFocus: Exit Sub
        End If
    Else
         Exit Sub
    End If
    If cboE(I����).ListIndex <= 0 Then Exit Sub
    datCur = Format(mdatCurDate, "yyyy-MM-dd HH:mm")
    Select Case cboE(I����).ListIndex
    Case 1 'Сʱ
        datRes = DateAdd("n", -1 * Val(txtSL.Text) * 60, CDate(datCur))
    Case 2 '��
        datRes = DateAdd("h", -1 * Val(txtSL.Text) * 24, CDate(datCur))
    Case 3 '��
        datRes = DateAdd("d", -1 * 7 * Val(txtSL.Text), CDate(datCur))
    Case 4 '��
        datRes = DateAdd("M", -1 * Int(Val(txtSL.Text)), CDate(datCur))
        datRes = DateAdd("d", -1 * (Val(txtSL.Text) - Int(Val(txtSL.Text))) * 30, datRes)
    Case 5 '��
        If Val(txtSL.Text) < 100 Then
            datRes = DateAdd("yyyy", -1 * Int(Val(txtSL.Text)), CDate(datCur))
            datRes = DateAdd("d", -1 * (Val(txtSL.Text) - Int(Val(txtSL.Text))) * 365, datRes)
        Else
            MsgBox "����ʱ�����㲻�ܳ���100�ꡣ", vbInformation, gstrSysName
            txtSL.SetFocus: Exit Sub
        End If
    End Select
    txtE(I����ʱ��).Text = Format(CDate(datRes), "YYYY-MM-DD HH:mm")
    If Not txtE(I����ʱ��).Locked Then
        Call UpDateInfo(txtE(I����ʱ��).Text, "txtE", I����ʱ��)
    End If
End Sub

Private Sub UCPatiVitalSigns_Change(ByVal int��� As Integer)
    mblnChange = True
End Sub

Private Sub UCPatiVitalSigns_GotFocus()
    Call SetCurCtlInfo(TypeName(UCPatiVitalSigns), "UCPatiVitalSigns")
End Sub

Private Sub UCPatiVitalSigns_Validate(Cancel As Boolean)
    Dim strSQL As String
    If mblnNoSave Then Exit Sub
    If mblnChange Then
        strSQL = UCPatiVitalSigns.GetSaveSQL(mlng����ID, mlng�Һ�ID)
        mrsMainInfo.Filter = "�ؼ���='UCPatiVitalSigns'"
        If mrsMainInfo!��Ϣԭֵ & "" <> strSQL Then
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            mrsMainInfo!��Ϣԭֵ = strSQL
            mrsMainInfo.Update
            strSQL = ""
            With UCPatiVitalSigns
                strSQL = mlng����ID & "<split>" & mlng�Һ�ID & "<split>" & .value��� & "<split>" & .value���� & "<split>" & .value���� & "<split>" & _
                    .value���� & "<split>" & .value���� & "<split>" & .value����ѹ & "<split>" & .value����ѹ & "<split>" & .valueѪѹ��λ
            End With
            RaiseEvent UpdatePatiState(strSQL, "")
            mblnOK = True
        End If
    End If
    mblnChange = False
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'���ܣ���ʾ��ʾ��Ϣ����λ��������Ŀ��
    Dim lngColor As Long
    
    On Error GoTo errH
 
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then
        If TypeName(objTmp) = "TextBox" Then zlControl.TxtSelAll objTmp
        objTmp.SetFocus
    End If
    Me.Refresh
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsAller_GotFocus()
    Call SetCurCtlInfo(TypeName(vsAller), "vsAller")
End Sub

Private Sub vsAller_Validate(Cancel As Boolean)
''''''''''''''
    Dim vsTmp As Object
    Dim i As Long
    Dim j As Long
    
    Set vsTmp = vsAller
    With vsAller
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, AI_����ҩ��)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, AI_����ҩ��)) > 60 Then
                    .Row = i: .Col = AI_����ҩ��
                    Call ShowMessage(vsTmp, "����ҩ����̫����ֻ����60���ַ���30�����֡�")
                    Cancel = True
                    Exit Sub
                End If
                If zlCommFun.ActualLen(.TextMatrix(i, AI_������Ӧ)) > 100 Then
                    .Row = i: .Col = AI_������Ӧ
                    Call ShowMessage(vsTmp, "������Ӧ����̫����ֻ����100���ַ���50�����֡�")
                    Cancel = True
                    Exit Sub
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, AI_����ҩ��)) <> "" And Format(.TextMatrix(i, AI_����ʱ��), "yyyy-mm-dd") = Format(.TextMatrix(j, AI_����ʱ��), "yyyy-mm-dd") Then
                        If .TextMatrix(j, AI_����ҩ��) = .TextMatrix(i, AI_����ҩ��) Then
                            .Row = i: .Col = AI_����ҩ��
                            Call ShowMessage(vsTmp, "����" & Format(.TextMatrix(j, AI_����ʱ��), "yyyy��mm��dd��") & "�ڴ�����ͬ�Ĺ���ҩ���¼��")
                            Cancel = True
                            Exit Sub
                        ElseIf Val(.TextMatrix(i, AI_ҩ��ID)) <> 0 And .TextMatrix(i, AI_ҩ��ID) = .TextMatrix(j, AI_ҩ��ID) Then
                            .Row = i: .Col = AI_����ҩ��
                             Call ShowMessage(vsTmp, "����" & Format(.TextMatrix(j, AI_����ʱ��), "yyyy��mm��dd��") & "�ڴ�����ͬ�Ĺ���ҩ���¼��")
                             Cancel = True
                            Exit Sub
                        ElseIf .TextMatrix(i, AI_����Դ����) <> "" And .TextMatrix(i, AI_����Դ����) = .TextMatrix(j, AI_����Դ����) Then
                            .Row = i: .Col = AI_����ҩ��
                            Call ShowMessage(vsTmp, "����" & Format(.TextMatrix(j, AI_����ʱ��), "yyyy��mm��dd��") & "�ڴ�����ͬ�Ĺ���ҩ���¼��")
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                Next
            End If
        Next
    End With
    Call UpDateAller
End Sub

Private Sub vsc_Change()
    Call vsc_Scroll
End Sub

Private Sub vsc_Scroll()
    mlngTopVsc = -1 * vsc.Value * Screen.TwipsPerPixelY
    Call MovePanel
End Sub

Private Sub cboE_Click(Index As Integer)
    Dim strValue As String
    Dim datCur As Date, datRes As Date
    
    If mblnCboNoClick Then Exit Sub
    strValue = NeedName(cboE(Index).Text)
    Call UpDateInfo(strValue, "cboE", Index)
    
    If Index = I���� Then
        If cboE(I����).ListIndex <= 0 Then Exit Sub
        If Trim(txtSL.Text) = "" Then Exit Sub
        datCur = Format(mdatCurDate, "yyyy-MM-dd HH:mm")
        Select Case cboE(I����).ListIndex
            Case 1 'Сʱ
                datRes = DateAdd("n", -1 * Val(txtSL.Text) * 60, CDate(datCur))
            Case 2 '��
                datRes = DateAdd("h", -1 * Val(txtSL.Text) * 24, CDate(datCur))
            Case 3 '��
                datRes = DateAdd("d", -1 * 7 * Val(txtSL.Text), CDate(datCur))
            Case 4 '��
                datRes = DateAdd("M", -1 * Int(Val(txtSL.Text)), CDate(datCur))
                datRes = DateAdd("d", -1 * (Val(txtSL.Text) - Int(Val(txtSL.Text))) * 30, datRes)
            Case 5 '��
                If Val(txtSL.Text) < 100 Then
                    datRes = DateAdd("yyyy", -1 * Int(Val(txtSL.Text)), CDate(datCur))
                    datRes = DateAdd("d", -1 * (Val(txtSL.Text) - Int(Val(txtSL.Text))) * 365, datRes)
                Else
                    MsgBox "����ʱ�����㲻�ܳ���100�ꡣ", vbInformation, gstrSysName
                    txtSL.SetFocus: Exit Sub
                End If
        End Select
        txtE(I����ʱ��).Text = Format(CDate(datRes), "YYYY-MM-DD HH:mm")
        If Not txtE(I����ʱ��).Locked Then
            Call UpDateInfo(txtE(I����ʱ��).Text, "txtE", I����ʱ��)
        End If
    End If
End Sub

Private Sub cboE_GotFocus(Index As Integer)
    If Index = I���֤�� Then
        Call zlControl.TxtSelAll(cboE(Index))
    End If
    Call SetCurCtlInfo(TypeName(cboE(Index)), "cboE", Index)
End Sub

Private Sub cboE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cboE(Index).Style = 0 Then
            cboE(Index).ListIndex = -1
            cboE(Index).Text = ""
        Else
            If cboE(Index).ListIndex <> -1 Then
               cboE(Index).ListIndex = -1
            End If
        End If
    End If
End Sub

Private Sub cboE_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strMask As String
    Dim lngidx As Long
    Dim blnCancel As Boolean
    
    If Index = I���֤�� Then
        Call cboSpecificInfoKeyPress(Index, KeyAscii)
    End If
    
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = I���֤�� Then
            Call cboE_Validate(I���֤��, blnCancel)
        End If
        If Not blnCancel Then
            
            If Index = IRH Then
                cboE(I���֤��).SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        If Index = I���֤�� Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            Else
                lblN(I���֤��).Tag = "1"
            End If
        Else
            lngidx = Cbo.MatchIndex(cboE(Index).hwnd, KeyAscii)
            If lngidx = -1 And cboE(Index).ListCount > 0 Then lngidx = 0
            cboE(Index).ListIndex = lngidx
        End If
    ElseIf KeyAscii = 8 Then
        If Index = I���֤�� Then
            lblN(I���֤��).Tag = "1"
        End If
    End If
End Sub

Private Sub cboE_Validate(Index As Integer, Cancel As Boolean)
    Dim strValue As String
    Dim str���֤�� As String
    
    If Index = I���֤�� Then
        '���������֤���Ǵ��ڲ�������  cboE(Index).Tag
        If cboE(Index).ListIndex = -1 Then
            strValue = cboE(Index).Tag
            mrsMainInfo.Filter = "��Ϣ��='���֤��'"
            str���֤�� = mrsMainInfo!��Ϣԭֵ & ""
            If strValue <> str���֤�� Then
                If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
                    If Not Check���֤��(strValue, cboE(Index)) Then
                        Cancel = True
                        cboE(Index).SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        Call UpDate���֤��
    End If
End Sub

Private Function Check���֤��(ByVal strNO As String, objTmp As Object) As Boolean
'���ܣ����֤�ż��
    Dim strTmp As String
    Dim lngColor As Long
    Dim str�������� As String
    Dim lng�Ա� As Long
    Dim strBirthday As String, strAge As String, strSex As String, strErrIfno As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If mobjPatient Is Nothing Then
        On Error Resume Next
        Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        err.Clear: On Error GoTo 0
    End If
    If mobjPatient Is Nothing Then
        MsgBox "����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ�", vbInformation, Me.Caption
        Exit Function
    End If

    On Error GoTo errH
    Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.�û���)
    
    strTmp = strNO
    lngColor = objTmp.BackColor
    
    If mobjPatient.CheckPatiIdcard(strTmp, strBirthday, strAge, strSex, strErrIfno) Then 'ʡ��֤�Ϸ������Ƿ�ƥ��
        '�ж��Ƿ��Ѿ����ˣ�Ҫ��ֹ
        If gblnPatiByID Then
            strSQL = "select 1 from ������Ϣ a where a.���֤��=[1] and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
            If Not rsTmp.EOF Then
                cboE(I���֤��).BackColor = &HC0C0FF
                MsgBox "�����֤���Ѿ�������ͬһ���ֻ֤�ܶ�Ӧһ����������!", vbInformation, Me.Caption
                cboE(I���֤��).BackColor = lngColor
                Exit Function
            End If
        End If

        If objTmp.Index = I���֤�� Then
            strTmp = ""
            If Format(strBirthday, "yyyy-MM-dd") <> Format(mstr��������, "yyyy-MM-dd") Then
                strTmp = "��������"
            End If
            If mstr�Ա� <> strSex Then
                strTmp = strTmp & IIf(strTmp <> "", "��", "") & "�Ա�"
            End If
            If strAge <> mstr���� Then
                strTmp = strTmp & IIf(strTmp <> "", "��", "") & "����"
            End If
            
            If strTmp <> "" Then
                If InStr(GetInsidePrivs(p������Ϣ��������), "������Ϣ����") = 0 Then
                    cboE(I���֤��).BackColor = &HC0C0FF
                    If MsgBox("���֤�����ȡ��" & strTmp & "�벡�˵�ǰ��" & strTmp & "��������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        cboE(I���֤��).BackColor = lngColor
                        Exit Function
                    Else
                        cboE(I���֤��).BackColor = lngColor
                    End If
                Else
                    strErrIfno = "���֤�����ȡ��" & strTmp & "�벡�˵�ǰ��" & strTmp & "��������Ƿ�������������Զ����½����ϵ�" & strTmp & "��"
                    
                    cboE(I���֤��).BackColor = &HC0C0FF
                    If MsgBox(strErrIfno, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        cboE(I���֤��).BackColor = lngColor
                        Exit Function
                    Else
                        cboE(I���֤��).BackColor = lngColor
                    End If
                    
                    If mobjPatient.SavePatiBaseInfo(mlng����ID, mlng�Һ�ID, mstr����, strSex, strAge, strBirthday, "������ҳ", 1, strErrIfno) Then
                        
                    Else
                        cboE(I���֤��).BackColor = &HC0C0FF
                        If MsgBox(strErrIfno & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            cboE(I���֤��).BackColor = lngColor
                            Exit Function
                        Else
                            cboE(I���֤��).BackColor = lngColor
                        End If
                    End If
                    mstr�Ա� = strSex
                    mstr���� = strAge
                    mstr�������� = Format(strBirthday, "yyyy-MM-dd")
                    RaiseEvent UpdatePatiInfo(strBirthday, strAge, strSex, "")
                End If
            End If
        End If
    Else '���֤���Ϸ����˳�
        objTmp.BackColor = &HC0C0FF
        MsgBox strErrIfno, vbInformation, gstrSysName
        objTmp.BackColor = lngColor
        Exit Function
    End If
         
    Check���֤�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendMsgDiag(ByVal datCur As Date) As Boolean
'���ܣ����������Ϣ
    Dim i As Long
    Dim arrTmp As Variant
    Dim strFilter As String
    On Error GoTo errH
    If mclsMipModule Is Nothing Then SendMsgDiag = True: Exit Function
  
    mrsSecdInfo.Filter = "�ı�״̬<>" & CS_δ�ı� & " And �ı�״̬<>" & CS_������
    Do While Not mrsSecdInfo.EOF
        If mrsSecdInfo!�ؼ��� = "vsDiagXY" Or mrsSecdInfo!�ؼ��� = "vsDiagZY" Then
            arrTmp = Split(mrsSecdInfo!��Ϣԭֵ & "", "|")
            If mrsSecdInfo!�ı�״̬ <> CS_������ Then 'ɾ�������滻���ȴ���ɾ�������Ϣ
'                Call ZLHIS_CIS_011(mclsMipModule, mlng����ID, mstr����, 1, mlng�Һ�ID, gclsPros.��Ժ����ID, mrsSecdInfo!ID, arrTmp(DMP_��ϱ���), arrTmp(DMP_��������))
            End If
            arrTmp = Split(mrsSecdInfo!��Ϣ��ֵ & "", "|")
            If mrsSecdInfo!�ı�״̬ <> CS_ɾ���� Then  '���������滻�д����´������Ϣ
'                Call ZLHIS_CIS_010(mclsMipModule, mlng����ID, mstr����, 1, mlng�Һ�ID, gclsPros.��Ժ����ID, Val(mrsSecdInfo!Tag & ""), arrTmp(DMP_�������), arrTmp(DMP_�Ƿ�����), arrTmp(DMP_��ϴ���), arrTmp(DMP_��ϱ���), arrTmp(DMP_��������), arrTmp(DMP_��������), arrTmp(DMP_�������), arrTmp(DMP_֤�����), arrTmp(DMP_֤������), datCur, UserInfo.����)
            End If
        End If
        mrsSecdInfo.MoveNext
    Loop
    
    SendMsgDiag = True
    '��ԭѧ��ϲ�������Ϣ
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function UpDateInfo(ByVal strValue As String, ByVal strCtlName As String, Optional ByVal intIdx As Integer = -1) As Boolean
    Dim strSQL As String, strSqlTwo As String
    Dim strFilter As String
    Dim str��Ϣ�� As String
    Dim strInfo As String
    Dim objTmp As Object
    Dim strMask As String
    Dim strTmp As String
    Dim i As Long
    Dim lngTmp As Long
    Dim strPar As String
    Dim datCur As Date
    Dim blnTrans As Boolean, blnEMPI As Boolean, strMsg As String
    
    If mblnNoSave Then Exit Function
    If mblnEdit Then Exit Function
    If strCtlName = "cboE" And intIdx = I���֤�� Then
        '���֤�����⴦��
        Call UpDate���֤��
    ElseIf strCtlName = "PatiAddress" Then
        '�ṹ����ַ���⴦��
        Call UpDate�ṹ����ַ(intIdx)
    Else
        strFilter = "�ؼ���='" & strCtlName & "' and Index= " & intIdx
        If strCtlName = "chkNoAller" Then
            '�޹�����¼����
            strFilter = "�ؼ���='" & strCtlName & "'"
        End If
        mrsMainInfo.Filter = strFilter
        If Not mrsMainInfo.EOF Then
            If strCtlName = "txtE" Then
                Set objTmp = txtE(intIdx)
                If (intIdx = I�����ص� Or intIdx = I���ڵ�ַ Or intIdx = I��ͥ��ַ) And mblnStructAdress Then
                    strInfo = PatiAddress(Decode(intIdx, I�����ص�, PT_�����ص�, I���ڵ�ַ, PT_���ڵ�ַ, I��ͥ��ַ, PT_��ͥ��ַ)).Value
                Else
                    strInfo = Trim(objTmp.Text)
                End If
                strInfo = Replace(strInfo, "'", "��")
                If InStr(",ժҪ,����ʱ��,������ַ,", "," & mrsMainInfo!��Ϣ�� & ",") = 0 Then
                    If zlCommFun.ActualLen(strInfo) > objTmp.MaxLength Then
                        objTmp.BackColor = &HC0C0FF
                        objTmp.Tag = mrsMainInfo!��Ϣ�� & "-����̫��(����¼��" & objTmp.MaxLength & "���ַ���" & objTmp.MaxLength \ 2 & "������)��"
                        Exit Function
                    Else
                        objTmp.BackColor = vbWindowBackground
                        If intIdx <> I�໤�����֤�� Then
                            objTmp.Tag = ""
                        End If
                    End If
                End If
                
                If strInfo <> "" Then
                    Select Case intIdx
                        Case I��ͥ�绰, I��λ�绰
                            strMask = "1234567890-()"
                            lngTmp = Len(strInfo)
                            strTmp = strInfo
                            objTmp.BackColor = vbWindowBackground
                            objTmp.Tag = ""
                            For i = 1 To lngTmp
                                If InStr(strMask, Mid(strTmp, i, 1)) = 0 Then
                                    objTmp.BackColor = &HC0C0FF
                                    objTmp.Tag = mrsMainInfo!��Ϣ�� & "-�����а����Ƿ��ַ�(����¼�������ַ�����" & strMask & "��)��"
                                    Exit Function
                                End If
                            Next
                        Case I��ͥ�ʱ�, I��λ�ʱ�, I�����ʱ�
                            strMask = "1234567890"
                            If (Not IsNumeric(strInfo)) Or InStr(strInfo, ".") > 0 Then
                                objTmp.BackColor = &HC0C0FF
                                objTmp.Tag = mrsMainInfo!��Ϣ�� & "-�����а����Ƿ��ַ�(����¼��0��9������)��"
                                Exit Function
                            Else
                                objTmp.BackColor = vbWindowBackground
                                objTmp.Tag = ""
                            End If
                    End Select
                End If
                objTmp.Text = strInfo
                strValue = strInfo
            End If
            
            If mrsMainInfo!��Ϣԭֵ & "" <> strValue Then
                strTmp = strValue
                If InStr(",ժҪ,����ʱ��,������ַ,", "," & mrsMainInfo!��Ϣ�� & ",") > 0 Then
                    Call UpDate�Һ���Ϣ(mrsMainInfo!��Ϣ�� & "", strValue)
                ElseIf mrsMainInfo!��Ϣ�� = "�໤��" Then
                    If mrsMainInfo!��Ϣԭֵ & "" <> "" And strValue = "" Then
                        MsgBox "�໤����Ϣֻ���޸ģ����������", vbInformation, gstrSysName
                        txtE(I�໤��).Text = mrsMainInfo!��Ϣԭֵ & ""
                        Exit Function
                    Else
                        strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                        strSQL = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'" & mrsMainInfo!��Ϣ�� & "'," & strValue & ")"
                        blnEMPI = True
                    End If
                ElseIf InStr(",RH,Ѫ��,����ҽѧ��ʾ,ҽѧ��ʾ,", "," & mrsMainInfo!��Ϣ�� & ",") > 0 Then
                    strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                    strSQL = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'" & mrsMainInfo!��Ϣ�� & "'," & strValue & "," & mlng�Һ�ID & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                ElseIf mrsMainInfo!��Ϣ�� = "��λ����" Then
                    If strValue = "" Then
                        mlng��ͬ��λID = 0
                    Else
                        strValue = "'" & strValue & "'"
                    End If
                    strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                    strSQL = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'������λ'," & strValue & ")"
                    strSqlTwo = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'��ͬ��λid'," & IIf(mlng��ͬ��λID = 0, "Null", mlng��ͬ��λID) & ")"
                    blnEMPI = True
                ElseIf mrsMainInfo!��Ϣ�� = "����״��" Then
                    i = 0
                    Set objTmp = cboE(I����״��)
                    If objTmp.Text <> "" And objTmp.ListIndex <> -1 Then
                        If InStr(objTmp.Text, "δ��") = 0 And InStr(objTmp.Text, "����") = 0 Then
                            If IsDate(mstr��������) Then
                                datCur = mdatCurDate
                                If DateDiff("yyyy", CDate(mstr��������), datCur) < 15 Then
                                     If MsgBox("�ò�������С��15�꣬����״��Ӧ��дΪδ����������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        strTmp = mrsMainInfo!��Ϣԭֵ & ""
                                        i = 1
                                        '�ָ�ԭֵ
                                        mblnCboNoClick = True
                                        objTmp.ListIndex = -1
                                        Call GetCboIndex(objTmp, strTmp)
                                        mblnCboNoClick = False
                                     End If
                                End If
                            End If
                        End If
                    End If
                    If i = 0 Then
                        strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                        strSQL = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'" & mrsMainInfo!��Ϣ�� & "'," & strValue & ")"
                        blnEMPI = True
                    End If
                ElseIf InStr(",�໤�����֤��,", "," & mrsMainInfo!��Ϣ�� & ",") > 0 Then
                    Call Update�໤�����֤
                Else
                    strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                    If mrsMainInfo!��Դ = 0 Then
                        strSQL = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'" & mrsMainInfo!��Ϣ�� & "'," & strValue & ")"
                        blnEMPI = True
                    Else
                        If mrsMainInfo!��Ϣ�� = "�޹�����¼" Then
                            strSQL = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'" & mrsMainInfo!��Ϣ�� & "'," & strValue & "," & mlng�Һ�ID & ")"
                        Else
                            strSQL = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'" & mrsMainInfo!��Ϣ�� & "'," & strValue & "," & mlng�Һ�ID & ")"
                            If mrsMainInfo!��Ϣ�� = "�Ļ��̶�" Then
                                blnEMPI = True
                            End If
                        End If
                    End If
                    On Error GoTo errH
                    gcnOracle.BeginTrans: blnTrans = True
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    If blnEMPI Then
                        If EMPIModifyPatiInfo(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, strMsg) = 0 Then
                            gcnOracle.RollbackTrans
                            blnTrans = False
                            MsgBox strMsg, vbInformation, gstrSysName
                            Exit Function
                        End If
                        blnEMPI = False
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                    If InStr(",��ͥ��ַ,��ͥ�绰,", "," & mrsMainInfo!��Ϣ�� & ",") > 0 Then
                        If HaveRIS Then
                            If gobjRis.HISModPati(1, mlng����ID, mlng�Һ�ID) <> 1 Then
                                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                            End If
                        ElseIf gbln����Ӱ����Ϣϵͳ�ӿ� = True Then
                            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                        End If
                    End If
                End If
                
                If blnEMPI Then
                    gcnOracle.BeginTrans: blnTrans = True
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    If strSqlTwo <> "" Then
                        Call zlDatabase.ExecuteProcedure(strSqlTwo, Me.Caption)
                    End If
                    If EMPIModifyPatiInfo(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, strMsg) = 0 Then
                        gcnOracle.RollbackTrans
                        blnTrans = False
                        MsgBox strMsg, vbInformation, gstrSysName
                        Exit Function
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                End If
                
                mrsMainInfo!��Ϣԭֵ = strTmp
                mrsMainInfo.Update
                mblnOK = True
            End If
        End If
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Update�໤�����֤()
'���ܣ��໤�����֤��
    Dim strSQL As String, i As Long, blnTrans As Boolean
    Dim strValue As String
    Dim intIndex As Integer
    Dim strPar As String, strMsg As String
    On Error GoTo errH
    
    If mblnNoSave Then Exit Sub
    mblnReturn = True
    intIndex = cboE(I���֤��).ListIndex
    strValue = Trim(txtE(I�໤�����֤��).Text)
    strPar = IIf(strValue = "", "null", "'" & strValue & "'")
    strSQL = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'�໤�����֤��'," & strPar & ",'')"
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    If EMPIModifyPatiInfo(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, strMsg) = 0 Then
        gcnOracle.RollbackTrans
        blnTrans = False
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    End If
    gcnOracle.CommitTrans: blnTrans = False
    mblnOK = True
    mblnReturn = False
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function EMPIModifyPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngClinicID As Long, ByRef strMsg As String) As Integer
    If Not gobjPlugIn Is Nothing Then
        err.Clear: On Error Resume Next
        If gobjPlugIn.EMPI_ModifyPatiInfo(lngSys, lngModule, lngPatiID, 0, lngClinicID, strMsg) = 0 Then
            If err.Number = 0 Then
                strMsg = "��ǰ������EMPIϵͳ�ӿڣ���EMPIϵͳ�ӿ�(EMPI_ModifyPatiInfo)δ���óɹ���" & strMsg
                EMPIModifyPatiInfo = 0
                Exit Function
            End If
        End If
        If err.Number <> 0 And err.Number <> 438 Then
            strMsg = "zlPlugIn ��Ҳ���ִ�� EMPI_ModifyPatiInfo ����ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description
            EMPIModifyPatiInfo = 0
            Exit Function
        End If
        err.Clear: On Error GoTo 0
    End If
    EMPIModifyPatiInfo = 1
End Function

Private Function UpDate�Һ���Ϣ(ByVal strInfoName As String, ByVal strValue As String)
'���ܣ����� ���ժҪ����Ⱦ���ϴ�������ʱ�䣬������ַ  �����ֶε�ֵ
'������strInfoName �ֶ�����,strValue �ֶε�ֵ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String
    Dim strArr(4) As String
    Dim strInfo As String
    Dim objTmp As Object
    Dim datCur As Date
    
    If mblnNoSave Then Exit Function
    On Error GoTo errH
    '���
    Set objTmp = txtE(I����ժҪ)
    strInfo = Trim(objTmp.Text)
    strInfo = Replace(strInfo, "'", "��")
    If zlCommFun.ActualLen(strInfo) > objTmp.MaxLength Then
        objTmp.BackColor = &HC0C0FF
        objTmp.Tag = "����ժҪ-����̫��(����¼��" & objTmp.MaxLength & "���ַ���" & objTmp.MaxLength \ 2 & "������)��"
        Exit Function
    Else
        objTmp.BackColor = vbWindowBackground
        objTmp.Tag = ""
    End If
    objTmp.Text = strInfo
    
    Set objTmp = txtE(I����ʱ��)
    strInfo = Trim(objTmp.Text)
    If Not IsDate(strInfo) And strInfo <> "" Then
        objTmp.BackColor = &HC0C0FF
        objTmp.Tag = "����ʱ��-���ڸ�ʽ���ԡ�"
        Exit Function
    ElseIf IsDate(strInfo) Then
        datCur = mdatCurDate
        If CDate(strInfo) > datCur Then
            objTmp.BackColor = &HC0C0FF
            objTmp.Tag = "����ʱ��Ӧ��С�ڵ�ǰʱ�䡣"
            Exit Function
        End If
    Else
        objTmp.BackColor = vbWindowBackground
        objTmp.Tag = ""
    End If
    objTmp.Text = strInfo
    
    strInfo = Trim(txtE(I������ַ).Text)
    strInfo = Trim(objTmp.Text)
    strInfo = Replace(strInfo, "'", "��")
    If zlCommFun.ActualLen(strInfo) > objTmp.MaxLength Then
        objTmp.BackColor = &HC0C0FF
        objTmp.Tag = "������ַ-����̫��(����¼��" & objTmp.MaxLength & "���ַ���" & objTmp.MaxLength \ 2 & "������)��"
        Exit Function
    Else
        objTmp.BackColor = vbWindowBackground
        objTmp.Tag = ""
    End If
    objTmp.Text = strInfo
    
    strSQL = "select ����,�Ա�,����,����,����,����,����,ְҵ,��������,�����ص�,���֤��,����֤��,����״��,ҽ�Ƹ��ʽ," & _
        "��ͥ��ַ,��ͥ�绰,��ͥ��ַ�ʱ�,���ڵ�ַ,���ڵ�ַ�ʱ�,��ͬ��λid,������λ,��λ�绰,��λ�ʱ�,��ϵ������,��ϵ�˹�ϵ," & _
        "��ϵ�˵绰,��ϵ�˵�ַ,Email,Qq,�໤�� from ������Ϣ where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If rsTmp.EOF Then Exit Function
    
    strSQL = "Zl_������Ϣ_��ҳ����(" & mlng����ID & ",'" & mstr����� & "',"
    For i = 0 To rsTmp.Fields.Count - 1
        If rsTmp.Fields(i).Name = "��������" Then
            strTmp = IIf(IsNull(rsTmp!��������), "NULL,", "To_Date('" & Format(rsTmp!��������, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),")
        Else
            strTmp = IIf(IsNull(rsTmp.Fields(i).Value), "NULL,", "'" & rsTmp.Fields(i).Value & "',")
        End If
        strSQL = strSQL & strTmp
    Next
    strSQL = strSQL & "'" & mstr�Һŵ� & "'"
    
    strArr(0) = IIf(optInfo(opt����).Value, 1, 0)
    
    strTmp = Trim(txtE(I����ժҪ).Text)
    strArr(1) = IIf(strTmp = "", "NULL", "'" & strTmp & "'")
    
    strArr(2) = chkInfo.Value
    
    strTmp = Trim(txtE(I����ʱ��).Text)
    If strTmp <> "" Then
        strArr(3) = "To_Date('" & Format(strTmp, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    Else
        strArr(3) = "NULL"
    End If
    
    strTmp = Trim(txtE(I������ַ).Text)
    strArr(4) = IIf(strTmp = "", "NULL", "'" & strTmp & "'")
    
    Select Case strInfoName
    Case "����"
        strArr(0) = strValue
    Case "ժҪ"
        If strValue <> "" Then
            strArr(1) = "'" & strValue & "'"
        Else
            strArr(1) = "NULL"
        End If
    Case "��Ⱦ���ϴ�"
        strArr(2) = strValue
    Case "����ʱ��"
        If strValue <> "" Then
            strArr(3) = "To_Date('" & Format(strValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        Else
            strArr(3) = "NULL"
        End If
    Case "������ַ"
        strArr(4) = IIf(strValue = "", "NULL", "'" & strValue & "'")
    End Select
    strSQL = strSQL & "," & strArr(0) & "," & strArr(1) & "," & strArr(2) & "," & strArr(3) & "," & strArr(4) & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If "ժҪ" = strInfoName Then
        strSQL = IIf(strValue = "", "NULL", strValue)
        RaiseEvent UpdatePatiInfo("", "", "", strSQL)
    End If
    mblnOK = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub UpDate����(ByVal intIdx As Integer)
'���ܣ����没��
    Dim arrSQL As Variant, i As Long, blnTrans As Boolean
    Dim lng����ID As Long
    Dim strValue As String
    Dim strTmp As String
 
    On Error GoTo errH
    If mblnNoSave Then Exit Sub
    If Not mblnChange Then Exit Sub
    If rtfEdit(intIdx).BackColor = DColor Then
        Exit Sub
    End If
    strValue = Trim(rtfEdit(intIdx).Text)
 
    mrsMainInfo.Filter = "�ؼ���='rtfEdit' and Index=" & intIdx
    
    strValue = Replace(strValue, "'", "��")
    If zlCommFun.ActualLen(strValue) > 4000 Then
        rtfEdit(intIdx).BackColor = &HC0C0FF
        strTmp = mrsMainInfo!��Ϣ�� & "-����̫��(����¼��4000���ַ���2000������)��"
        mrsMainInfo.Update "ErrInfo", strTmp
        Exit Sub
    Else
        rtfEdit(intIdx).BackColor = vbWindowBackground
        mrsMainInfo.Update "ErrInfo", ""
    End If
    
    If mrsMainInfo!��Ϣԭֵ & "" = strValue Then
        Exit Sub
    End If

    arrSQL = Array()
    If mlng����ID = 0 Then lng����ID = zlDatabase.GetNextId("���Ӳ�����¼")
    
    Call GetSQLOutDoc(arrSQL, lng����ID)
    If UBound(arrSQL) = -1 Then Exit Sub
    
    '�ύ����
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
  
    If ReadRTFData(lng����ID) = False Then GoTo errH
    If SaveRTFData(lng����ID) = False Then GoTo errH
      
    gcnOracle.CommitTrans: blnTrans = False
    
    mrsMainInfo!��Ϣԭֵ = strValue
    mrsMainInfo.Update
 
    If mlng����ID = 0 Then Call LoadDocData
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub UpDate�ṹ����ַ(ByVal intIdx As Integer)
'���ܣ����½ṹ����ַ
    Dim strSQL As String
    Dim strValue As String
    Dim blnTrans As Boolean
    Dim strSQLTmp As String
On Error GoTo errH
    If mblnNoSave Then Exit Sub
    If Not mblnUpdate Then Exit Sub
    strValue = Trim(PatiAddress(intIdx).Value)
    mrsMainInfo.Filter = "�ؼ���='PatiAddress' and Index=" & intIdx

    If mrsMainInfo!��Ϣԭֵ & "" = strValue Then
        Exit Sub
    End If

    If mblnStructAdress Then
        If PatiAddress(intIdx).Value <> "" Then
           strSQL = "zl_���˵�ַ��Ϣ_update(1," & mlng����ID & ",NULL," & Decode(intIdx, PT_�����ص�, 1, PT_���ڵ�ַ, 4, PT_��ͥ��ַ, 3) & ",'" & PatiAddress(intIdx).valueʡ & "','" & _
               PatiAddress(intIdx).value�� & "','" & PatiAddress(intIdx).value���� & "','" & PatiAddress(intIdx).value���� & "','" & _
               PatiAddress(intIdx).value��ϸ��ַ & "','" & PatiAddress(intIdx).Code & "')"
        Else
           strSQL = "zl_���˵�ַ��Ϣ_update(2," & mlng����ID & ",NULL," & Decode(intIdx, PT_�����ص�, 1, PT_���ڵ�ַ, 4, PT_��ͥ��ַ, 3) & ")"
        End If
    End If
    
    strSQLTmp = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'" & Decode(intIdx, PT_�����ص�, "�����ص�", PT_���ڵ�ַ, "���ڵ�ַ", PT_��ͥ��ַ, "��ͥ��ַ") & "','" & PatiAddress(intIdx).Value & "')"
    
On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call zlDatabase.ExecuteProcedure(strSQLTmp, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    mrsMainInfo!��Ϣԭֵ = strValue
    mrsMainInfo.Update
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub UpDate���֤��()
'���ܣ����֤��
'���ܣ�intType 0-���֤�ţ�1�����֤��״̬
    Dim arrSQL As Variant, i As Long, blnTrans As Boolean
    Dim strValue As String
    Dim intIndex As Integer
    Dim str���֤�� As String
    Dim str���֤��״̬ As String
    Dim blnDo As Boolean
    Dim strPar As String, strMsg As String
    On Error GoTo errH
    
    If mblnNoSave Then Exit Sub
    mblnReturn = True
    intIndex = cboE(I���֤��).ListIndex
    
    arrSQL = Array()
    If intIndex = -1 Then
        strValue = cboE(I���֤��).Tag
    Else
        strValue = cboE(I���֤��).Text
    End If
    
    strPar = IIf(strValue = "", "null", "'" & strValue & "'")
    If intIndex = -1 Then
        If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'���֤��'," & strPar & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'�⼮���֤��',Null,Null)"
        Else
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'�⼮���֤��'," & strPar & ",'')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'���֤��',Null)"
        End If
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'���֤��״̬',Null,Null)"
    Else
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_������Ϣ_������Ϣ(" & mlng����ID & ",'���֤��',Null)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'�⼮���֤��',Null,Null)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'���֤��״̬'," & strPar & ",Null)"
    End If
    
    '�ύ����
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If EMPIModifyPatiInfo(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, strMsg) = 0 Then
        gcnOracle.RollbackTrans
        blnTrans = False
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    End If
    gcnOracle.CommitTrans: blnTrans = False
    mblnOK = True
    If intIndex = -1 Then
        If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
            mrsMainInfo.Filter = "��Ϣ��='���֤��'"
            mrsMainInfo.Update "��Ϣԭֵ", strValue
        Else
            mrsMainInfo.Filter = "��Ϣ��='�⼮���֤��'"
            mrsMainInfo.Update "��Ϣԭֵ", strValue
        End If
        mrsMainInfo.Filter = "��Ϣ��='���֤��״̬'"
        mrsMainInfo.Update "��Ϣԭֵ", ""
        
        cboE(I���֤��).Tag = strValue
        If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
            If mblnID���� Then
                strValue = Mid(strValue, 1, 12) & String(Len(Mid(strValue, 13, 2)), "*") & Mid(strValue, 15)
            End If
        End If
        cboE(I���֤��).Text = strValue
    Else
        mrsMainInfo.Filter = "��Ϣ��='���֤��'"
        mrsMainInfo.Update "��Ϣԭֵ", ""
        mrsMainInfo.Filter = "��Ϣ��='�⼮���֤��'"
        mrsMainInfo.Update "��Ϣԭֵ", ""
        mrsMainInfo.Filter = "��Ϣ��='���֤��״̬'"
        mrsMainInfo.Update "��Ϣԭֵ", strValue
    End If
    cboE(I���֤��).ToolTipText = ""
    cboE(I���֤��).BackColor = vbWindowBackground
    mblnReturn = False
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtE_GotFocus(Index As Integer)
    If Index <> I����ժҪ Then
        Call zlControl.TxtSelAll(txtE(Index))
    ElseIf txtE(Index).SelLength = 0 Then
        Call zlControl.TxtSelAll(txtE(Index))
    End If
    Call SetCurCtlInfo(TypeName(txtE(Index)), "txtE", Index)
End Sub

Private Sub txtE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Index = Iҽѧ��ʾ Then
            txtE(Iҽѧ��ʾ) = ""
        End If
    End If
End Sub

Private Sub txtE_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    Dim txtTmp As Object
    Dim strValue As String, str���֤�� As String
    If Index = I�໤�����֤�� Then
        If KeyAscii = vbKeyReturn Then
            zlCommFun.PressKey vbKeyTab
        Else
            Set txtTmp = txtE(Index)
            If Not (KeyAscii >= 0 And KeyAscii < 32) Then
                If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
                    If zlCommFun.ActualLen(txtTmp.Text) >= 18 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 Then
                            KeyAscii = 0
                        ElseIf zlCommFun.IsCharChinese(txtTmp.Text) Then
                            txtTmp.Text = "": txtTmp.Tag = ""
                        End If
                        If KeyAscii <> 0 Then
                            Select Case zlCommFun.ActualLen(txtTmp.Text)
                                Case 12
                                    txtTmp.Tag = txtTmp.Text & Chr(KeyAscii)
                                Case 13
                                    txtTmp.Tag = txtTmp.Tag & Chr(KeyAscii)
                            End Select
                        End If
                    End If
                Else
                    If Not (KeyAscii >= 0 And KeyAscii < 32) Then
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 Then
                            KeyAscii = 0
                        ElseIf zlCommFun.IsCharChinese(txtTmp.Text) Then
                            txtTmp.Text = "": txtTmp.Tag = ""
                        End If
                        If KeyAscii <> 0 Then
                            txtTmp.Tag = txtTmp.Text & Chr(KeyAscii)
                        End If
                    End If
                End If
            End If
        End If
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = I���� Or Index = I����) And txtE(Index).Text <> "" Then
            '������������
            strSQL = "Select Rownum as ID,����,����,���� From ���� " & _
                " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2]) And Nvl(����, 0) < 3" & _
                " Order by ����"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(Index = I����, "����", "����"), False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, False, False, _
                UCase(txtE(Index).Text) & "%", gstrLike & UCase(txtE(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtE(Index).Text = rsTmp!����
            End If
            txtE(Index).SetFocus
        ElseIf (Index = I�����ص� Or Index = I��ͥ��ַ Or Index = I���ڵ�ַ) And txtE(Index).Text <> "" Then
            '�����������
            strSQL = "Select Rownum as ID,����,����,���� From ���� " & _
                " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, False, False, _
                UCase(txtE(Index).Text) & "%", gstrLike & UCase(txtE(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtE(Index).Text = rsTmp!����
            End If
            txtE(Index).SetFocus
            
        ElseIf Index = I��ͥ�ʱ� Or Index = I�����ʱ� Or Index = I��λ�ʱ� Then
            If ((Not IsNumeric(txtE(Index).Text)) Or Len(txtE(Index).Text) > 6 Or InStr(txtE(Index).Text, ".") > 0) And txtE(Index).Text <> "" Then
                If txtE(Index).Text <> "" Then
                    If zlCommFun.IsCharChinese(txtE(Index).Text) Then
                        strSQL = strSQL & " And A.���� Like [1] "
                    Else
                        strSQL = strSQL & " And A.���� Like [1] "
                    End If
                End If
                strSQL = "Select Rownum as ID,����,����,�ʱ�  From ���� A " & _
                "Where �ʱ� is not null " & strSQL & " Order by ����"
                vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, _
                                        False, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, False, False, _
                                        UCase(txtE(Index).Text) & "%")
                '������������,��һ��Ҫƥ��
                If Not rsTmp Is Nothing Then
                    txtE(Index).Text = rsTmp!�ʱ� & ""
                End If
                txtE(Index).SetFocus
            End If
        ElseIf Index = I��λ���� And txtE(Index).Text <> "" Then
            '���빤����λ
            strSQL = "Select ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " And (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������λ", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, False, False, _
                UCase(txtE(Index).Text) & "%", gstrLike & UCase(txtE(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtE(Index).Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                If mblnEdit��ͬ��λ Then
                    mlng��ͬ��λID = Val(rsTmp!ID)
                Else
                    mlng��ͬ��λID = 0
                End If
                If txtE(I��λ�绰).Text = "" Then
                    txtE(I��λ�绰).Text = NVL(rsTmp!�绰)
                End If
            Else
                txtE(Index).Tag = ""
                mlng��ͬ��λID = 0
            End If
            txtE(Index).SetFocus
        ElseIf Index = I����ժҪ Then
            Call TxtKeyPressժҪ(KeyAscii)
        ElseIf Index = I�໤�����֤�� Then
            Call txtE_Validate(I�໤�����֤��, blnCancel)
            If blnCancel Then
                mblnNoSave = True
                Exit Sub
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        '�ǿ��ư���
        If Index = Iҽѧ��ʾ Then
            KeyAscii = 0
        ElseIf Index = I�໤�����֤�� Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            Else
                lblN(I�໤�����֤��).Tag = "1"
            End If
        End If
        If KeyAscii = 39 Then KeyAscii = 0 '�����ű���
        'ѡ���ݼ�
        If KeyAscii = Asc("*") Then
            'ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
            On Error Resume Next
            strSQL = ""
            strSQL = cmdE(Index).Name
            err.Clear: On Error GoTo 0
            If strSQL <> "" Then
                KeyAscii = 0
                Call cmdE_Click(Index)
                Exit Sub
            End If
        End If
        
        '�������볤��
        If txtE(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtE(Index).Text) > txtE(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '������������
        Select Case Index
            Case I��ͥ�绰, I��λ�绰
                strMask = "1234567890-()"
            Case I����ʱ��
                strMask = "1234567890-: "
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    Else
        If Index = Iҽѧ��ʾ Then
            KeyAscii = 0
        ElseIf Index = I�໤�����֤�� Then
            lblN(I�໤�����֤��).Tag = "1"
        End If
    End If
End Sub

Private Sub TxtKeyPressժҪ(KeyAscii As Integer)
    Dim objTxt As Object
    Set objTxt = txtE(I����ժҪ)
    If objTxt.Text <> "" Then
        If AbstractSelect(objTxt.Text) Then Exit Sub
    End If
End Sub

Private Sub cmdE_Click(Index As Integer)
'˵����ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, blnLevel As Boolean
    Dim strResult As String
    Dim blnNoSave As Boolean
    
    'ʹ��Lock�ķ�ʽ,������Enabled�ķ�ʽ
    If Not cmdE(Index).Enabled Or txtE(Index).Locked Then
        If txtE(Index).Enabled Then txtE(Index).SetFocus
        Exit Sub
    End If
    
    Select Case Index
        Case I�����ص�, I��ͥ��ַ, I���ڵ�ַ
            'ѡ���������
            strSQL = "Select Rownum as ID,����,����,���� From ���� Order by ����"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""����""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                Call txtE(Index).SetFocus
                blnNoSave = True
            Else
                txtE(Index).Text = rsTmp!����
                txtE(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case I��λ����
            'ѡ��λ��Ϣ
            strSQL = "Select ID,�ϼ�ID,ĩ��,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ��" & _
                " From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "��Լ��λ", , , , , True, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""��Լ��λ""���ݣ����ȵ���Լ��λ���������á�", vbInformation, gstrSysName
                End If
                txtE(Index).Tag = ""
                If txtE(Index).Enabled Then txtE(Index).SetFocus
                blnNoSave = True
            Else
                txtE(Index).Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                If mblnEdit��ͬ��λ Then
                    mlng��ͬ��λID = Val(rsTmp!ID)
                Else
                    mlng��ͬ��λID = 0
                End If
                If txtE(I��λ�绰).Text = "" Then
                    txtE(I��λ�绰).Text = NVL(rsTmp!�绰)
                End If
                If txtE(Index).Enabled Then txtE(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case I����, I����
            'ѡ����������
            strSQL = "Select 1  From ���� Where Nvl(����,0)<>0 And RowNum<2"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTmp.RecordCount > 0 Then blnLevel = True
            
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            If blnLevel Then
                strSQL = _
                        "Select Id, �ϼ�id, Id ����, ����, ����, ĩ��" & vbNewLine & _
                        "From (Select Rpad(����, 15, '0') As Id, Rpad(Substr(����, 1, Decode(Nvl(����, 0), 0, 0, 1, 2, 4)), 15, '0') As �ϼ�id, ����, ����," & vbNewLine & _
                        "              Decode(Nvl(����, 0), 2, 1, 3, 1, 0) As ĩ��" & vbNewLine & _
                        "       From ����" & vbNewLine & _
                        "       Where Nvl(����, 0) < 3" & vbNewLine & _
                        "       Order By ����)" & vbNewLine & _
                        "Start With �ϼ�id Is Null" & vbNewLine & _
                        "Connect By Prior Id = �ϼ�id"
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "����", , , , , , , vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel)
            Else
                strSQL = "Select Rownum as ID,����,����,���� From ���� Order by ����"
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""����""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtE(Index).SetFocus
                blnNoSave = True
            Else
                txtE(Index).Text = rsTmp!����
                txtE(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case Iҽѧ��ʾ
            'ѡ��ҽѧ��ʾ
            On Error GoTo errH
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            strSQL = "Select Rownum ID,����,����,���� From ҽѧ��ʾ Order by ����"
            Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "", True, "", "", True, True, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, True, True)

            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""ҽѧ��ʾ""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtE(Index).SetFocus
                blnNoSave = True
            Else
                While Not rsTmp.EOF
                    strResult = strResult & "," & rsTmp!����
                    rsTmp.MoveNext
                Wend
                txtE(Index).Text = Mid(strResult, 2)
                txtE(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case I����ʱ��
            dtpDate.ZOrder
            If IsDate(txtE(I����ʱ��).Text) Then
                 dtpDate.Value = CDate(txtE(I����ʱ��).Text)
             Else
                 dtpDate.Value = mdatCurDate
             End If
             dtpDate.Visible = True
             dtpDate.SetFocus
        Case Else
            blnNoSave = True
    End Select
    
    If Not blnNoSave Then
        Call UpDateInfo(txtE(Index).Text, "txtE", Index)
    End If
    Set mobjCtl = txtE(Index)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    'ȡֵ
    If IsDate(txtE(I����ʱ��).Text) Then
        strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtE(I����ʱ��).Text, "yyyy-MM-dd HH:mm"), 12, 5)
    Else
        strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(mdatCurDate, "yyyy-MM-dd HH:mm"), 12, 5)
    End If
    
    txtE(I����ʱ��).Text = strDate
    dtpDate.Tag = ""
    txtE(I����ʱ��).SetFocus
    dtpDate.Visible = False
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        txtE(I����ʱ��).SetFocus
        dtpDate.Tag = ""
        dtpDate.Visible = False
    End If
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'''
    Dim strDate As String
    
    With vsAller
        Select Case Col
            Case AI_����ʱ��
                strDate = GetFullDate(.TextMatrix(Row, Col), False)
                If IsDate(strDate) Then
                    .TextMatrix(Row, Col) = strDate
                End If
        End Select
        Call vsAller_AfterRowColChange(-1, -1, .Row, .Col)
    End With
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'''
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    With vsAller
        If NewCol = AI_����ҩ�� Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = IIf(Trim(.TextMatrix(NewRow, AI_����ҩ��)) = "", flexFocusLight, flexFocusSolid)
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int�Ա� As Integer
    Dim vPoint As POINTAPI
    
    With vsAller
        If Not gobjPass Is Nothing Then
            If optInfo(opt����Դ).Value Then
                strSQL = gobjPass.zlPassInputAllergy()
                If InStr(strSQL, ";") > 0 Then
                    Call SetAllerInput(Row, , strSQL)
                    Call AllerEnterNextCell
                End If
            Else
                If mstr�Ա� Like "*��*" Then
                    int�Ա� = 1
                ElseIf mstr�Ա� Like "*Ů*" Then
                    int�Ա� = 2
                End If
                
                strSQL = _
                    " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                    " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                    " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                    " Select ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
                    " NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ��" & _
                    " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                    " Union All" & _
                    " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����,A.����," & _
                    " A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                    " From ������ĿĿ¼ A,ҩƷ���� B" & _
                    " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
                    IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[1])", "") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����ҩ��", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int�Ա�)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û��ҩƷ���ݿ���ѡ��", vbInformation, gstrSysName
                    End If
                Else
                    Call SetAllerInput(Row, rsTmp)
                    Call AllerEnterNextCell
                End If
            End If
        Else
            If mstr�Ա� Like "*��*" Then
                int�Ա� = 1
            ElseIf mstr�Ա� Like "*Ů*" Then
                int�Ա� = 2
            End If
            If optInfo(optҩƷĿ¼).Value Then
                strSQL = _
                    " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                    " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                    " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                    " Select ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
                    " NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ��" & _
                    " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                    " Union All" & _
                    " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����,A.����," & _
                    " A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                    " From ������ĿĿ¼ A,ҩƷ���� B" & _
                    " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
                    IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[1])", "") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����ҩ��", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int�Ա�)
            Else
                strSQL = "Select Rownum As ID, ����, ����, ���� From ����Դ Order By ����"
                vPoint = zlControl.GetCoordPos(vsAller.hwnd, vsAller.Left, vsAller.CellTop)
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "����Դ", , , , , True, True, vPoint.X, vPoint.Y, vsAller.Height, blnCancel)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    If optInfo(optҩƷĿ¼).Value Then
                        MsgBox "û��ҩƷ���ݿ���ѡ��", vbInformation, gstrSysName
                    Else
                        MsgBox "û�й���Դ���ݿ���ѡ��", vbInformation, gstrSysName
                    End If
                End If
            Else
                Call SetAllerInput(Row, rsTmp)
                Call AllerEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If vsAller.Editable = flexEDNone Then Exit Sub
    
    With vsAller
        If KeyCode = vbKeyF4 Then
            If .Col = 1 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, AI_����ҩ��) <> "" Then
                If MsgBox("ȷʵҪ������й���ҩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    Call UpDateAller
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsAller_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    If vsAller.Editable = flexEDNone Then Exit Sub
    If 39 = KeyAscii Then KeyAscii = 0 '������
    With vsAller
        If KeyAscii = vbKeySpace Then   'Space
            If .Col = AI_����ҩ�� And Not gobjPass Is Nothing And optInfo(opt����Դ).Value Then KeyAscii = 0: Exit Sub
        End If
        If KeyAscii = 13 Then
             KeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = AI_����ҩ�� Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsAller_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsAller_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If 39 = KeyAscii Then KeyAscii = 0
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    With vsAller
        If Col = AI_����ʱ�� Then
            If KeyAscii = 13 Then
                .Col = .Col + 1
                .ShowCell Row, Col
                .Col = .Col - 1
            End If
        ElseIf Col = AI_����ҩ�� Then
            If KeyAscii <> 13 Then
                If Not gobjPass Is Nothing And optInfo(opt����Դ).Value Then KeyAscii = 0
            End If
        End If
    End With
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsAller
        If Col = AI_����ҩ�� Or Col = AI_����ʱ�� Then
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
        End If
    End With
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = AI_������Ӧ And Trim(vsAller.TextMatrix(Row, AI_����ҩ��)) = "" Then Cancel = True
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim int�Ա�  As Integer
    Dim curDate As Date
    Dim strDate As String
    
    With vsAller
        If Col = AI_����ҩ�� Then
            If .EditText = "" Then
                If .Cell(flexcpData, Row, Col) <> "" Then
                    If MsgBox("ȷʵҪ������й���ҩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        .RemoveItem .Row
                        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
                    Else
                        .EditText = .Cell(flexcpData, Row, Col)
                    End If
                End If
                If mblnReturn Then Call AllerEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call AllerEnterNextCell
            Else
                If LenB(StrConv(.EditText, vbFromUnicode)) > 60 Then
                    MsgBox "ҩ�����Ʋ��ܳ���30�����ֵĳ��ȡ�", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
                strInput = UCase(.EditText)
                If mstr�Ա� Like "*��*" Then
                    int�Ա� = 1
                ElseIf mstr�Ա� Like "*Ů*" Then
                    int�Ա� = 2
                End If
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                If optInfo(optҩƷĿ¼) Then
                    strSQL = _
                        " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ," & _
                        " B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                        " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
                        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.ID=C.������ĿID" & _
                        " And (A.���� Like [1] Or A.���� Like [2] Or C.���� Like [2] Or C.���� Like [2])" & _
                        IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[3])", "") & _
                        Decode(mint����, 0, " And C.����=[4]", 1, " And C.����=[4]", "") & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " Order by A.����"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҩ��", False, "", "", False, _
                        False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gstrLike & strInput & "%", int�Ա�, mint���� + 1)
                Else
                    If zlCommFun.IsCharChinese(strInput) Then
                        strSQL = "Select Rownum As ID, ����, ����, ���� From ����Դ Where ���� Like [1] Order By ����"
                    Else
                        If mint���� = 1 Then
                            strSQL = "Select Rownum As ID, ����, ����, ���� From ����Դ Where zlWbCode(����) Like [1] Order By ����"
                        Else
                            strSQL = "Select Rownum As ID, ����, ����, ���� From ����Դ Where ���� Like [1] Order By ����"
                        End If
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����Դ", False, "", "", False, _
                        False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        gstrLike & UCase(strInput) & "%")
                End If
                If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    Cancel = True
                Else
                    Call SetAllerInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call AllerEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = AI_����ʱ�� Then
            If .EditText <> "" Then
                strDate = GetFullDate(.EditText, False)
                If IsDate(strDate) Then
                    curDate = mdatCurDate
                    If CDate(strDate) > curDate Then
                        MsgBox "����������ڲ��ܴ��ڵ�ǰʱ�䡣��ǰʱ�䣺" & Format(curDate, "yyyy-mm-dd") & "��"
                        Cancel = True
                        .EditText = .TextMatrix(Row, Col)
                    End If
                    .EditText = Format(strDate, "yyyy-MM-dd")
                Else
                    MsgBox "��������ȷ�Ĺ���ʱ�䣬���磺""2012-12-21""��""121221""��"
                    Cancel = True
                End If
            End If
        Else
            If LenB(StrConv(.EditText, vbFromUnicode)) > 100 Then
                MsgBox "������Ӧ���ܳ���50�����ֵĳ��ȡ�", vbInformation, Me.Caption
                Cancel = True
                Exit Sub
            End If
        End If
    End With
End Sub

Private Function DiagCellEditable(ByRef vsDiagTmp As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    Dim bln��ҽ As Boolean
    Dim blnJudge As Boolean
    Dim dtTmp As DiagType
    Dim lng��ԺRow As Long
    
    If lngRow < 0 Then
        Exit Function
    End If
    With vsDiagTmp
        bln��ҽ = .Name = "vsDiagXY"
        '�����в��ɱ༭
        If .ColHidden(lngCol) Then Exit Function
        '���������������������������������(�����߼�����������
        If .TextMatrix(lngRow, DI_�������) = "" Then
            If Not (lngCol = DI_Del Or lngCol = DI_�������) Then Exit Function
        Else
            If lngCol <> DI_���� And lngCol <> DI_Del And lngCol <> DI_���� Then
                '����ҽ�����ɱ༭
                If .TextMatrix(lngRow, DI_ҽ��IDs) <> "" Then Exit Function
            End If
        End If
        DiagCellEditable = True
    End With
End Function

Private Sub DiagKeyDown(ByRef vsDiag As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsDiagXY_KeyDown�¼���vsDiagZY_KeyDown�¼�
    Dim i As Long, j As Long
    Dim dtCurRow As DiagType, lngRow As Long
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        If intKeyCode = vbKeyF4 Then
            If .Col = DI_������� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf intKeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, DI_�������) <> "" Or .Rows = .FixedRows + 1 Then
                If .TextMatrix(.Row, DI_�������) = "" Then Exit Sub
                If Not DiagCellEditable(vsDiag, .Row, DI_�������) Then Exit Sub
                If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    'ɾ����/��Ҫ��Ϻ������ҽӿ�
                    If Not gobjPlugIn Is Nothing Then
                        On Error Resume Next
                        Call gobjPlugIn.DiagnosisDeleted(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, Val(.TextMatrix(.Row, DI_���ID)), .TextMatrix(.Row, DI_�������))
                        Call zlPlugInErrH(err, "DiagnosisDeleted")
                        err.Clear: On Error GoTo 0
                    End If
                    dtCurRow = Val(.TextMatrix(.Row, DI_��Ϸ���))
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedCols, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, DI_��Ϸ���) = dtCurRow
                    '�����ͬ�������������
                    If .TextMatrix(.Row, DI_�������) = "" Or .Rows <> .FixedRows + 1 Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            For i = .Row + 1 To .Rows - 1
                                '�����ƶ���ǰ��һ��,����ǰ��Ϊ��һ�����࣬��ǰ��һ���Ƿ�����һ���������ʼ�У����ǣ���ɾ����
                                If .TextMatrix(i, DI_��Ϸ���) <> "" Then
                                    If .TextMatrix(i - 1, DI_��Ϸ���) = "" Then .RemoveItem i - 1
                                    Exit For
                                End If
                                '�����ƶ���ǰ��һ��
                                For j = .FixedCols To .Cols - 1
                                    .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                    .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                Next
                                .RowData(i - 1) = .RowData(i)
                                '���һ��ɾ��
                                If i = .Rows - 1 Then
                                    .RemoveItem i: Exit For
                                End If
                            Next
                        End If
                    End If
                    Call UpDateDiag(vsDiag)
                End If
            ElseIf .TextMatrix(.Row, DI_�������) = "" Or .Rows <> .FixedRows + 1 Then
                .RemoveItem .Row
            End If
            '������������Ϣ
            '�����д�˷���ʱ�䣬������ķ���ʱ����������д��
'            If gclsPros.FuncType <> f���ѡ�� Then Call SetDiagReletedInfo(vsDiag)
        ElseIf intKeyCode = vbKeyInsert Then '������
            lngRow = .Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, DI_��Ϸ���) = .TextMatrix(lngRow - 1, DI_��Ϸ���)
            .TextMatrix(lngRow, DI_�������) = .TextMatrix(lngRow - 1, DI_�������)
            .Cell(flexcpForeColor, .FixedRows, DI_�Ƿ�����, .Rows - 1, DI_�Ƿ�����) = vbRed
            .Cell(flexcpBackColor, .FixedRows, DI_��ϱ���, .Rows - 1, DI_��ϱ���) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Row = lngRow: .Col = DI_��ϱ���
            .ShowCell .Row, .Col
        ElseIf intKeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call DiagKeyPress(vsDiag, intKeyCode)
        End If
    End With
End Sub

Private Sub DiagKeyPress(ByRef vsDiag As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsDiagXY_KeyPress�¼���vsDiagZY_KeyPress�¼�
    If vsDiag.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = 39 Then intKeyAscii = 0 '�����ű���
    With vsDiag
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Call EnterNextCellDiag(vsDiag)
        Else
            If .Col <> DI_�Ƿ����� Then
                If Not DiagCellEditable(vsDiag, .Row, .Col) Then Exit Sub
            End If
            Select Case .Col
                Case DI_�Ƿ�����
                    If intKeyAscii <> vbKeySpace Then Exit Sub
                    intKeyAscii = 0
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", IIf(.Col = DI_�Ƿ�����, "��", "��"), "")
                Case DI_��ϱ���, DI_�������, DI_��ҽ֤��, DI_ICD���� '��ҽ��ҽ֤������,��ҽ��ICD��������
                    If intKeyAscii = Asc("*") Then
                        intKeyAscii = 0
                        Call DiagCellButtonClick(vsDiag, .Row, .Col)
                    Else
                        .ComboList = "" 'ʹ��ť״̬��������״̬
                    End If
            End Select
        End If
    End With
End Sub

Private Sub EnterNextCellDiag(ByRef vsDiagTmp As VSFlexGrid)
    Dim i As Long, j As Long
    
    With vsDiagTmp
        '����һ��Ԫ��ʼѭ������
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, DI_�������) To DI_Del
                If Not .ColHidden(j) Then
                    If DiagCellEditable(vsDiagTmp, i, j) And .ColWidth(j) <> 0 Then Exit For
                End If
            Next
            If j <= DI_Del Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > DI_Del And .TextMatrix(.Rows - 1, DI_�������) <> "" Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, DI_��Ϸ���) = .TextMatrix(.Rows - 2, DI_��Ϸ���)
            .TextMatrix(.Rows - 1, DI_�������) = .TextMatrix(.Rows - 2, DI_�������)
            .ShowCell i, DI_�������
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub DiagCellButtonClick(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
'vsDiagZY_CellButtonClick�¼���vsDiagXY_CellButtonClick�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngCurRow As Long
    Dim bln��ҽ As Boolean
    
    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        Select Case lngCol
            Case DI_�������, DI_��ϱ���
                If optInfo(opt���).Value Then
                    '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, IIf(bln��ҽ, "1", "2"), mlng����ID, , True, False)
                Else
                    'B-��ҽ�������룬7-�����ж���Y-�����ж����ⲿԭ��6-������ϣ�M-������̬ѧ���룻������ϣ�D-ICD-10��������
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, IIf(bln��ҽ, "D", "B"), mlng����ID, mstr�Ա�, True, True, , glngSys)
                End If
                If Not rsTmp Is Nothing Then
                    Call SetDiagInput(vsDiag, lngRow, rsTmp)
                    Call EnterNextCellDiag(vsDiag)
                    zlControl.ControlSetFocus vsDiag, True
                End If
            Case DI_��ҽ֤��
                If optInfo(opt���).Value Then
                    '���������:�Ȳ��Ƿ��ж�Ӧ
                    If Set��ҽ֤��(lngRow, Val(.TextMatrix(lngRow, DI_���ID))) Then
                        zlControl.ControlSetFocus vsDiag, True
                        Exit Sub
                    End If
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng����ID, mstr�Ա�, True, , , glngSys)
                Else
                    'Z-��ҽ��������
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng����ID, mstr�Ա�, True, , , glngSys)
                End If
                If Not rsTmp Is Nothing Then
                    Call Set��ҽ֤��(lngRow, 0, rsTmp)
                    Call EnterNextCellDiag(vsDiag)
                    zlControl.ControlSetFocus vsDiag, True
                End If
            Case DI_����
                If Not .Cell(flexcpPicture, lngRow, DI_����) Is Nothing Or Not .CellButtonPicture Is Nothing Then
                    Call DiagKeyDown(vsDiag, vbKeyInsert, 0)
                End If
            Case DI_Del
                If Not .Cell(flexcpPicture, lngRow, DI_Del) Is Nothing Or Not .CellButtonPicture Is Nothing Then
                    Call DiagKeyDown(vsDiag, vbKeyDelete, 0)
                End If
        End Select
    End With
End Sub

Private Sub SetDiagInput(ByRef vsDiagTmp As VSFlexGrid, ByVal lngRow As Long, rsInput As ADODB.Recordset, Optional bln���� As Boolean)
'���ܣ����������Ŀ������
'      bln����=�Ƿ��Ǹ�������
    Dim str�Ա� As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long, j As Long
    Dim strTmp As String, bln�ֻ��̶� As Boolean
    Dim bln��ҽ As Boolean, blnRCodeIn As Boolean
    Dim lngTmpRow As Long, lng��ԺRow As Long
    Dim lngԭ���ID As Long, int��ϴ��� As Integer
    
    With vsDiagTmp
        bln��ҽ = .Name = "vsDiagXY"
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                '���ǵ����ĸ�������
                If Not bln���� Then
                    If i > 1 Then
                        '���һ����������ҽ��������ϣ���ҽ�������ж���ѡ�����ʱ�Ĵ���
                        lngԭ���ID = 0
                        If lngRow = .Rows - 1 Then
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, DI_��Ϸ���) = .TextMatrix(lngRow, DI_��Ϸ���)
                            .TextMatrix(.Rows - 1, DI_�������) = .TextMatrix(lngRow, DI_�������)
                        End If
                        'ȷ����ǰ��ʾ��
                        If Val(.TextMatrix(lngRow + 1, DI_��Ϸ���)) = Val(.TextMatrix(lngRow, DI_��Ϸ���)) Then
                            For j = lngRow + 1 To .Rows - 1
                                If Val(.TextMatrix(j, DI_��Ϸ���)) = Val(.TextMatrix(lngRow, DI_��Ϸ���)) Then
                                    lngRow = j
                                    If .TextMatrix(j, DI_�������) = "" Then Exit For
                                Else
                                    Exit For
                                End If
                            Next
                            If .TextMatrix(lngRow, DI_�������) <> "" Then
                                lngRow = lngRow + 1: .AddItem "", lngRow
                                .TextMatrix(lngRow, DI_��Ϸ���) = .TextMatrix(lngRow - 1, DI_��Ϸ���)
                                .TextMatrix(lngRow, DI_�������) = .TextMatrix(lngRow - 1, DI_�������)
                            End If
                        Else
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, DI_��Ϸ���) = .TextMatrix(lngRow - 1, DI_��Ϸ���)
                            .TextMatrix(lngRow, DI_�������) = .TextMatrix(lngRow - 1, DI_�������)
                        End If
                    Else
                        lngԭ���ID = Val(.TextMatrix(lngRow, DI_���ID))
                    End If
                    
                    .TextMatrix(lngRow, DI_��ϱ���) = rsInput!���� & ""
                    .TextMatrix(lngRow, DI_�������) = rsInput!����
                    .Cell(flexcpData, lngRow, DI_�������) = rsInput!���� & ""  '����ԭ��
                    .Cell(flexcpData, lngRow, DI_��ϱ���) = rsInput!���� & ""
                    .TextMatrix(lngRow, DI_���ID) = rsInput!���id & ""
                    .TextMatrix(lngRow, DI_����ID) = rsInput!����id & ""
                    .TextMatrix(lngRow, DI_��������) = rsInput!�������� & ""
                    .TextMatrix(lngRow, DI_�������) = rsInput!������� & ""
                End If
                If Not bln���� Then .TextMatrix(lngRow, DI_�̶�����) = IIf(Not IsNull(rsInput!����), "1", "")
                .TextMatrix(lngRow, DI_ICD����) = IIf(bln����, rsInput!���� & "", rsInput!���� & "")
                .TextMatrix(lngRow, DI_����ID) = IIf(bln����, rsInput!��ĿID & "", rsInput!����ID & "")
                
                If Not bln��ҽ Then
                    '��ҽ���ݼ�����ϲο�ȡ֤��
                    Call Set��ҽ֤��(lngRow, Val(.TextMatrix(lngRow, DI_���ID)))
                End If
                
                If CreatePlugInOK(p����ҽ��վ, -1) Then
                    int��ϴ��� = 0
                    int��ϴ��� = IIf(lngRow = .FixedRows, -1, -2)
                    On Error Resume Next
                    Select Case int��ϴ���
                        Case -1
                            Call gobjPlugIn.DiagnosisEnter(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, Val(rsInput!��ĿID), .TextMatrix(lngRow, DI_�������), lngԭ���ID)
                            Call zlPlugInErrH(err, "DiagnosisEnter")
                        Case -2
                            Call gobjPlugIn.DiagnosisOtherEnter(glngSys, p����ҽ��վ, mlng����ID, mlng�Һ�ID, Val(rsInput!��ĿID), .TextMatrix(lngRow, DI_�������), lngԭ���ID)
                            Call zlPlugInErrH(err, "DiagnosisOtherEnter")
                    End Select
                    err.Clear: On Error GoTo errH
                End If
                
                rsInput.MoveNext
            Next
        Else
            If Not bln���� Then
                .TextMatrix(lngRow, DI_�������) = .EditText
              
                .Cell(flexcpData, lngRow, DI_�������) = .TextMatrix(lngRow, DI_�������)
                .TextMatrix(lngRow, DI_��ϱ���) = ""
                 .Cell(flexcpData, lngRow, DI_�������) = ""
                .TextMatrix(lngRow, DI_���ID) = ""
                .TextMatrix(lngRow, DI_����ID) = ""
                .TextMatrix(lngRow, DI_֤��ID) = ""
           
            Else
                .TextMatrix(lngRow, DI_�̶�����) = ""
                .TextMatrix(lngRow, DI_ICD����) = ""
                .TextMatrix(lngRow, DI_����ID) = ""
            End If
        End If
        .Cell(flexcpForeColor, .FixedRows, DI_�Ƿ�����, .Rows - 1, DI_�Ƿ�����) = vbRed
        .Cell(flexcpBackColor, .FixedRows, DI_��ϱ���, .Rows - 1, DI_��ϱ���) = GRD_UNEDITCELL_COLOR      '����ɫ
        
        '������������Ϣ
        Call SetDiagReletedInfo(vsDiagTmp, lngRow)
        If optInfo(opt����).Value = False Then
            If PatiReSeeDoctor Then
                If MsgBox("���˾�����ҡ�ҽ����������ϴ���ͬ��Ҫ���Ϊ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    optInfo(opt����).Value = True
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'cmdMakeLog�¼�
Private Sub MakeLog()
'���ܣ��������������ժҪ��
    Dim strLog As String, i As Long
    Dim strTmp As String
 
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_�������) <> "" Then
                strLog = strLog & "��" & .TextMatrix(i, DI_�������) & IIf(.TextMatrix(i, DI_�Ƿ�����) <> "", "(��)", "")
            End If
        Next
    End With

    With vsDiagZY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_�������) <> "" Then
                strLog = strLog & "��" & .TextMatrix(i, DI_�������) & IIf(.TextMatrix(i, DI_�Ƿ�����) <> "", "(��)", "")
            End If
        Next
    End With
    If strLog <> "" Then
        With txtE(I����ժҪ)
            If .SelStart = 0 And .SelLength = Len(.Text) Then
                .SelStart = Len(.Text)
            End If
            i = .SelStart
            .SelText = Mid(strLog, 2)
            .SelStart = i
            .SelLength = Len(Mid(strLog, 2))
            .SetFocus
            strTmp = .Text
        End With
        
        Call UpDateInfo(strTmp, "txtE", I����ժҪ)
    End If
End Sub


Private Function PatiReSeeDoctor() As Boolean
'���ܣ��жϲ��˱����Ƿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String
    Dim strSQL As String
    Dim vsTmp As VSFlexGrid
    
    On Error GoTo errH
    
    'ҽ�����������ϴ���ͬ��û��ת������
    strSQL1 = "Select ����ID,ִ���� as ҽ��,ִ�в���ID as ����ID From ���˹Һż�¼ Where ID=[2] And ת�����ID Is Null And �������ID Is Null"
    
    strSQL2 = "Select Max(ID) as ID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
            " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
    strSQL2 = "Select ����ID,ִ���� as ҽ��,ִ�в���ID as ����ID From ���˹Һż�¼ Where ID=(" & strSQL2 & ") And ת�����ID Is Null And �������ID Is Null"
    
    strSQL = "Select 1 From (" & strSQL1 & ") A,(" & strSQL2 & ") B Where A.����ID=B.����ID And A.ҽ��=B.ҽ�� And A.����ID=B.����ID"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng����ID, mlng�Һ�ID)
    If rsTmp.EOF Then Exit Function
    
    '��Ҫ������ϴ���ͬ
    With vsDiagXY
        If .TextMatrix(.FixedRows, DI_�������) <> "" Then
            strSQL = "Select Max(ID) as ��ҳID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
                    " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
            strSQL = "Select 1 From ������ϼ�¼" & _
                " Where ����ID=[1] And ��ҳID=(" & strSQL & ")" & _
                " And �������=1 And ��¼��Դ IN(1,3) And ��ϴ���=1" & _
                " And (����ID=[3] And ����ID<>0 Or ���ID=[4] And ���ID<>0 Or �������=[5])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng����ID, mlng�Һ�ID, _
                Val(.TextMatrix(.FixedRows, DI_����ID)), Val(.TextMatrix(.FixedRows, DI_���ID)), .TextMatrix(.FixedRows, DI_�������))
            If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
        End If
    End With
    
    If mbln��ҽ Then
        With vsDiagZY
            If .TextMatrix(.FixedRows, DI_�������) <> "" Then
                strSQL = "Select Max(ID) as ��ҳID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
                       " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
                strSQL = "Select 1 From ������ϼ�¼" & _
                    " Where ����ID=[1] And ��ҳID=(" & strSQL & ")" & _
                    " And �������=11 And ��¼��Դ IN(1,3) And ��ϴ���=1" & _
                    " And (����ID=[3] And ����ID<>0 Or ���ID=[4] And ���ID<>0 Or �������=[5])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng����ID, mlng�Һ�ID, _
                    Val(.TextMatrix(.FixedRows, DI_����ID)), Val(.TextMatrix(.FixedRows, DI_���ID)), .TextMatrix(.FixedRows, DI_�������))
                If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
            End If
        End With
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDiagReletedInfo(ByRef vsDiagTmp As VSFlexGrid, Optional ByVal lngRow As Long = -1)
'����������������������������
    '�����д�˷���ʱ�䣬������ķ���ʱ����������д��
End Sub

Private Function Set��ҽ֤��(ByVal lngRow As Long, ByVal lng���ID As Long, Optional ByVal rsInput As Recordset, Optional ByVal blnFreeInput As Boolean) As Boolean
'���ܣ���ҽ���ݼ�����ϲο�ȡ֤��
'������rsInput-�����Ϊ�գ������ָ������ҩ֤���¼��
'���أ��Ƿ��ж�Ӧ��ϵ
    Dim strSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String
    
    On Error GoTo errH
    
    With vsDiagZY
        If blnFreeInput Then
            .TextMatrix(lngRow, DI_֤��ID) = ""
            .TextMatrix(lngRow, DI_֤�����) = ""
            .TextMatrix(lngRow, DI_��ҽ֤��) = .EditText
        Else
            'ȥ�����е�֤��
            If .TextMatrix(lngRow, DI_�������) Like "?*(?*)" Then
                strTmp = Mid(.TextMatrix(lngRow, DI_�������), 1, InStrRev(.TextMatrix(lngRow, DI_�������), "(") - 1)
            Else
                strTmp = .TextMatrix(lngRow, DI_�������)
            End If
            
            If rsInput Is Nothing Then
                If lng���ID = 0 Then Exit Function
                strSQL = "Select Distinct A.֤����� As ID, A.֤��id As ��Ŀid, B.����, B.����, A.֤������ ����," & IIf(mint���� = 0, "B.����", "B.����� As ����") & ", B.˵��" & vbNewLine & _
                            "From ������ϲο� A, ��������Ŀ¼ B" & vbNewLine & _
                            "Where A.֤��id = B.Id(+) And A.���id = [1] And A.֤������ Is Not Null" & vbNewLine & _
                            "Order By A.֤�����"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsInput = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng���ID)
                If rsInput Is Nothing Then
                    If Not blnCancel Then Exit Function
                    If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, DI_��ҽ֤��)
                    Set��ҽ֤�� = True: Exit Function
                End If
            End If
            
            .TextMatrix(lngRow, DI_֤��ID) = NVL(rsInput!��ĿID)
            .TextMatrix(lngRow, DI_֤�����) = NVL(rsInput!����)
            If Not IsNull(rsInput!����) Then
                .TextMatrix(lngRow, DI_�������) = strTmp
                .Cell(flexcpData, lngRow, DI_�������) = .TextMatrix(lngRow, DI_�������)
                .TextMatrix(lngRow, DI_��ҽ֤��) = NVL(rsInput!����)
                .Cell(flexcpData, lngRow, DI_��ҽ֤��) = .TextMatrix(lngRow, DI_��ҽ֤��)
                If .EditText <> "" Then .EditText = .TextMatrix(lngRow, DI_��ҽ֤��)
            End If
        End If
        Set��ҽ֤�� = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DiagAfterEdit(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
    'vsDiagXY_AfterEdit�¼�,vsDiagZY_AfterEdit�¼�
    Dim bln��ҽ As Boolean
    
    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        If lngCol = DI_������� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If .EditText = "" And .Cell(flexcpData, lngRow, lngCol) <> "" Then
                '�ڵ���vsDiagXY_KeyDown(vbKeyDelete, 0)���ǿ���ɾ����ǰ�У������ָ�ԭʼ����
                .TextMatrix(lngRow, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                Call DiagKeyDown(vsDiag, vbKeyDelete, 0)
            End If
        End If
        Call DiagAfterRowColChange(vsDiag, -1, -1, .Row, .Col)
        zlControl.ControlSetFocus vsDiag, True
    End With
End Sub

Private Sub DiagAfterRowColChange(ByRef vsDiag As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsDiagZY_AfterRowColChange�¼���vsDiagXY_AfterRowColChange�¼�
    Dim i As Long
    Dim bln��ҽ As Boolean
    Dim vPoint As POINTAPI
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        '���ͼƬ
        For i = .FixedRows To .Rows - 1
            Set .Cell(flexcpPicture, i, DI_����) = Nothing
            Set .Cell(flexcpPicture, i, DI_Del) = Nothing
        Next
        
        If Not DiagCellEditable(vsDiag, lngNewRow, lngNewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .ComboList = ""
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing
             
            Select Case lngNewCol
                Case DI_�������
                    .ComboList = "..."
                Case DI_����, DI_Del
                    .ComboList = "..."
                    .FocusRect = flexFocusNone
                    Set .CellButtonPicture = IIf(lngNewCol = DI_����, imgButtonNew.Picture, imgButtonDel.Picture)
                Case DI_��ҽ֤��
                    If .TextMatrix(lngNewRow, DI_�������) = "" Then
                        .ComboList = ""
                        .FocusRect = flexFocusLight
                    Else
                        .ComboList = "..."
                    End If
                Case Else
                    .ComboList = ""
            End Select
        End If
        If lngNewRow >= .FixedRows Then
            '��ʾͼƬ
            If lngNewCol <> DI_���� And .TextMatrix(lngNewRow, DI_�������) <> "" Then
                If .Rows - 1 <> lngNewRow Then
                    '��һ�����Ϊ������������
                    If Not (.TextMatrix(lngNewRow, DI_��Ϸ���) = .TextMatrix(lngNewRow + 1, DI_��Ϸ���) And .TextMatrix(lngNewRow + 1, DI_�������) = "") Then
                         Set .Cell(flexcpPicture, lngNewRow, DI_����) = imgButtonNew.Picture
                    End If
                Else
                    Set .Cell(flexcpPicture, lngNewRow, DI_����) = imgButtonNew.Picture
                End If
            End If
            '��ʾͼƬ
            If lngNewCol <> DI_Del Then Set .Cell(flexcpPicture, lngNewRow, DI_Del) = imgButtonDel.Picture
        End If
        zlControl.ControlSetFocus vsDiag, True
    End With
End Sub

Private Sub DiagAfterUserResize(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
'vsDiagZY_BeforeUserResize�¼���vsDiagXY_BeforeUserResize�¼�
    If lngCol = DI_������� Then
        If vsDiagZY.ColWidth(DI_��ҽ֤��) < vsDiagXY.ColWidth(lngCol) Then
             vsDiagZY.ColHidden(DI_��ҽ֤��) = False
             vsDiagZY.ColWidth(lngCol) = vsDiagXY.ColWidth(lngCol) - vsDiagZY.ColWidth(DI_��ҽ֤��)
        Else
             vsDiagZY.ColHidden(DI_��ҽ֤��) = True
             vsDiagZY.ColWidth(lngCol) = vsDiagXY.ColWidth(lngCol)
        End If
    Else
         vsDiagZY.ColWidth(lngCol) = vsDiagXY.ColWidth(lngCol)
    End If
End Sub

Private Sub DiagBeforeUserResize(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByRef blnCancel As Boolean)
'vsDiagZY_BeforeUserResize�¼���vsDiagXY_BeforeUserResize�¼�
    If lngCol = DI_���� Or lngCol = DI_Del Or lngCol < DI_������� Then blnCancel = True
End Sub

Private Sub DiagClick(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_Click�¼���vsDiagZY_Click�¼�
    Dim bln��ҽ As Boolean
    
    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        If (.MouseCol = DI_���� Or .MouseCol = DI_Del) And .MouseRow >= .FixedRows Then
            If .MouseCol = DI_���� Then
                If .TextMatrix(.MouseRow, DI_�������) = "" Or .TextMatrix(.MouseRow, 0) = IIf(bln��ҽ, "��Ժ���", "��Ҫ���") Then Exit Sub
            End If
            .Select .MouseRow, .MouseCol
            Call DiagCellButtonClick(vsDiag, .MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub DiagDblClick(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_DblClick�¼���vsDiagZY_DblClick�¼�
    Call DiagKeyPress(vsDiag, vbKeySpace)
End Sub

Private Sub DiagGotFocus(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_GotFocus�¼���vsDiagZY_GotFocus�¼�
    Call SetCurCtlInfo(TypeName(vsDiag), vsDiag.Name)

    If vsDiag.Row >= vsDiag.FixedRows And vsDiag.Col >= vsDiag.FixedCols Then
        Call DiagAfterRowColChange(vsDiag, -1, -1, vsDiag.Row, vsDiag.Col)
    End If

    zlControl.ControlSetFocus vsDiag, True
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagXY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagXY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterUserResize(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_Click()
    Call DiagClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_DblClick()
    Call DiagDblClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_GotFocus()
    Call DiagGotFocus(vsDiagXY)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagXY, KeyCode, Shift)
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagXY, KeyAscii)
End Sub

Private Sub vsAller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim strInfo As String
 
    With vsAller
        If .Tag <> "" Then
            lngRow = Val(.Tag)
            If .MouseRow = lngRow Then
                strInfo = "����������ͬ�Ĺ�����¼��"
            End If
        End If
    End With
    Call zlCommFun.ShowTipInfo(vsAller.hwnd, strInfo, True, True)
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagZY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagXY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim strInfo As String
 
    With vsDiagXY
        If .Tag <> "" Then
            lngRow = Val(.Tag)
            If .MouseRow = lngRow Then
                strInfo = "����������ͬ��ϡ�"
            End If
        End If
    End With
    Call zlCommFun.ShowTipInfo(vsDiagXY.hwnd, strInfo, True, True)
End Sub

Private Sub vsDiagZY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Dim lngRow As Long
    Dim strTmp As String
    
    strInfo = vsDiagZY.Tag
    If InStr(strInfo, "|") <> 0 Then
        lngRow = Val(Split(strInfo, "|")(0))
        strInfo = Split(strInfo, "|")(1)
        With vsDiagZY
            If .MouseRow = lngRow Then
                strTmp = strInfo
            End If
        End With
    End If
    strInfo = strTmp
    Call zlCommFun.ShowTipInfo(vsDiagZY.hwnd, strInfo, True, True)
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagXY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub DiagSetupEditWindow(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsDiagXY_SetupEditWindow�¼���vsDiagZY_SetupEditWindow�¼�
    With vsDiag
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub DiagStartEdit(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByRef blnCancel As Boolean)
'vsDiagXY_StartEdit�¼���vsDiagZY_StartEdit�¼�
    
    If Not DiagCellEditable(vsDiag, lngRow, lngCol) Then
        blnCancel = True
    ElseIf lngCol = DI_�Ƿ����� Then
        blnCancel = True '��ֱ�ӱ༭
    End If
    
End Sub

Private Sub vsDiagXY_Validate(Cancel As Boolean)
    Call UpDateDiag(vsDiagXY)
End Sub

Private Sub vsDiagZY_Validate(Cancel As Boolean)
    Call UpDateDiag(vsDiagZY)
End Sub

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub DiagValidateEdit(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByRef blnCancel As Boolean)
'vsDiagXY_ValidateEdit�¼���vsDiagZY_ValidateEdit�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnInputCancel As Boolean
    Dim int������� As Integer
    Dim strInput As String, vPoint As POINTAPI
    Dim strDiagType As String
    Dim bln��ҽ As Boolean
    Dim str�Ա� As String
    
    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        Select Case lngCol
            Case DI_�������, DI_��ϱ���
                If bln��ҽ Then
                    strDiagType = "'D'"
                Else
                    strDiagType = IIf(optInfo(opt���).Value, "", "B")
                End If
                
                If .EditText = "" And .Cell(flexcpData, lngRow, lngCol) <> "" Then
                    .EditText = ""
                ElseIf .EditText = .Cell(flexcpData, lngRow, lngCol) Then
                    If mblnReturn Then Call EnterNextCellDiag(vsDiag)
                ElseIf .TextMatrix(lngRow, DI_��ϱ���) <> "" And .Cell(flexcpData, lngRow, lngCol) <> "" And .EditText Like "*" & .Cell(flexcpData, lngRow, lngCol) & "*" Then
                    '�жϼ���ǰ׺��������Ƿ������������ϱ���
                    strInput = UCase(.EditText)
                    strSQL = GetMedInputSQL(IIf(bln��ҽ, 0, 1), strInput, str�Ա�)
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, strDiagType, str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng����ID)
                    If rsTmp.RecordCount = 1 Then
                        Call SetDiagInput(vsDiag, lngRow, rsTmp)
                        .EditText = .Text
                    Else
                        '�����ڱ�׼������ǰ�����븽����Ϣ
                        '������.Cell(flexcpData, lngRow, lngCol)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
                        .TextMatrix(lngRow, DI_�������) = .EditText
                    End If
                ElseIf .TextMatrix(lngRow, DI_��ϱ���) <> "" And .Cell(flexcpData, lngRow, lngCol) <> "" And mblnFreeInput Then
                    strInput = UCase(.EditText)
                    strSQL = GetMedInputSQL(IIf(bln��ҽ, 0, 1), strInput, str�Ա�)
                    On Error GoTo errH
                    vPoint = zlControl.GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInfo(opt����).Value, "�������", "��������"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gstrLike & strInput & "%", strDiagType, str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                    If blnInputCancel Then
                        blnCancel = True
                    Else
                        If rsTmp Is Nothing Then
                            .TextMatrix(lngRow, DI_�������) = .EditText
                        Else
                             Call SetDiagInput(vsDiag, lngRow, rsTmp): .EditText = .Text
                        End If
                    End If
                Else
                    int������� = mint�������
                    strInput = UCase(.EditText)
                    strSQL = GetMedInputSQL(IIf(bln��ҽ, 0, 1), strInput, str�Ա�)
                    '�����ж��룺Y-�����ж����ⲿԭ�򣻲����������M-������̬ѧ���룻������ϣ�D-ICD-10��������
                    vPoint = zlControl.GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInfo(opt����).Value, "�������", "��������"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gstrLike & strInput & "%", strDiagType, str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                    If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        blnCancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And (int������� = 2 Or int������� = 3 And mint���� <> 0) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            blnCancel = True
                        ElseIf Not (rsTmp Is Nothing) Then
                            Call SetDiagInput(vsDiag, lngRow, rsTmp)
                            .EditText = .Text
                        Else
                            'û��ƥ��ɹ��ٴε�������¼��
                            If int������� = 1 Or (int������� = 3 And (rsTmp Is Nothing) And mint���� = 0) Then
                                Call SetDiagInput(vsDiag, lngRow, Nothing)
                                .EditText = .Text
                            Else
                                blnCancel = True
                            End If
                        End If
                    End If
                End If
      
                mblnReturn = False
            Case DI_��ҽ֤��
                If .EditText = "" And .Cell(flexcpData, lngRow, lngCol) <> "" Then
                    .EditText = ""
                    '��ҽ֤���������������
                    .Cell(flexcpData, lngRow, lngCol) = ""
                ElseIf .EditText = .Cell(flexcpData, lngRow, lngCol) Then
                    If mblnReturn Then Call EnterNextCellDiag(vsDiag)
                
                ElseIf .TextMatrix(lngRow, DI_��ϱ���) <> "" And .Cell(flexcpData, lngRow, lngCol) <> "" And mblnFreeInput Then
                    strInput = UCase(.EditText)
                    strDiagType = "Z"
                    strSQL = GetMedInputSQL(1, strInput, str�Ա�, strDiagType)
                    vPoint = zlControl.GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gstrLike & strInput & "%", strDiagType, str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                    If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        blnCancel = True
                    Else
                        If rsTmp Is Nothing Then
                            .TextMatrix(lngRow, DI_��ҽ֤��) = .EditText
                        Else
                            Call Set��ҽ֤��(lngRow, 0, rsTmp, rsTmp Is Nothing)
                        End If
                    End If
                Else
                    int������� = mint�������
                    strInput = UCase(.EditText)
                    strDiagType = "Z"
                    strSQL = GetMedInputSQL(1, strInput, str�Ա�, strDiagType)
                    If optInfo(opt����).Value Then
                        '���������:�Ȳ��Ƿ��ж�Ӧ
                        If Set��ҽ֤��(lngRow, Val(.TextMatrix(lngRow, DI_���ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = zlControl.GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gstrLike & strInput & "%", strDiagType, str�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                    If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        blnCancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And (int������� = 2 Or int������� = 3 And mint���� <> 0) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            blnCancel = True
                        Else
                            Call Set��ҽ֤��(lngRow, 0, rsTmp, rsTmp Is Nothing)
                        End If
                    End If
                End If
                mblnReturn = False
            Case DI_����ʱ��
                If .EditText <> "" Then
                    strInput = GetFullDate(.EditText)
                    If IsDate(strInput) Then
                        .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    Else
                        MsgBox "��������ȷ�ķ���ʱ�䣬���磺""2012-12-21 00:00""��"
                        blnCancel = True
                    End If
                End If
                If lngRow = .FixedRows Then
'                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_����ʱ��), IsDate(.EditText), True)
'                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��������), IsDate(.EditText), True)
                End If
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetMedInputSQL(ByVal intType As Integer, ByVal strInput As String, ByRef str�Ա� As String, Optional ByVal strOtherInfo As String) As String
'���ܣ���ò�ѯ��ҳ�����ѯ��SQL
'������intType:��ȡ��SQL����,0-��ҽ��ϣ�1-��ҽ��ϣ�2-��������
'    strInput-��ѯ������str�Ա�--���˵��Ա�
'    strOtherInfo:��ҽ���-������������
'���أ�strsql--��ѯ��ϵ�SQL

    Dim strSQL As String

    If mstr�Ա� Like "*��*" Then
        str�Ա� = "��"
    ElseIf mstr�Ա� Like "*Ů*" Then
        str�Ա� = "Ů"
    End If

    Select Case intType
        Case 0, 1 '��ҽ���,��ҽ���
            If intType = 0 And optInfo(opt���).Value Or intType = 1 And optInfo(opt���).Value And strOtherInfo <> "Z" Then
            '���������:һ����Ͽ������ڶ������
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                End If
                strSQL = "Select A.Id, A.Id ��ĿID, A.����, Null ���, Null ����, Null ����id, Null ��������, A.����, A.˵��, A.����, B.����, 0 ��Ч����, 0 ����," & vbNewLine & _
                                "              0 �Ƿ���, Max(D.����id) ����id, A.Id ���id" & vbNewLine & _
                                "       From �������Ŀ¼ A, ������ϱ��� B, ������϶��� D" & vbNewLine & _
                                " Where A.ID=B.���ID And A.ID=D.���ID(+) And A.���=" & IIf(intType = 0, 1, 2) & vbNewLine & _
                                " And B.����=[5] And (" & strSQL & ")" & vbNewLine & _
                                " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                                "Group By A.Id, A.����, A.����, A.˵��, A.����,B.����"
                '��ȡ��϶�Ӧ�������븽��
                strSQL = "Select distinct A.ID,A.��ĿID, A.����, B.���, B.����, Null ����id, Null ��������, A.����, A.˵��, Null ����,A.����, A.��Ч����, A.����, A.�Ƿ���," & vbNewLine & _
                                "       B.���� ��������, B.Id ����id, B.��� �������, A.���id," & vbNewLine & _
                                "      Decode(a.����, [6], 1, Decode(A.����,[6],1,decode(A.����,[6],1,NULL))) As ����1ID,Decode(d.���id, Null, Decode(c.���id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                                "      Decode(Substr(A.����, 1, Length([6])), [6], 1, Decode(Substr(A.����, 1, Length([6])),[6],1,decode(Substr(a.����, 1, Length([6])),[6],1,NULL))) As ����3ID" & _
                                " From (" & strSQL & ") A, ��������Ŀ¼ B, ������Ͽ��� C, ������Ͽ��� D" & vbNewLine & _
                                " Where A.����id = B.Id(+)" & vbNewLine & _
                                " And c.���id(+) = a.Id And d.���id(+) = a.Id And c.����id(+)=[8]  And d.��Աid(+) = [7]" & _
                                " Order By ����1ID, ����2ID, ����3ID, A.����"
            Else
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "A.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "A.���� Like [1] Or A.���� Like [2] Or " & IIf(mint���� = 0, "A.����", "A.�����") & " Like [2]"
                End If
           
                strSQL = _
                    "Select A.Id,A.Id ��ĿID, A.����, A.���, A.����,Null ����ID, Null ��������, A.����, A.˵��, Null ����, A.����id, " & IIf(mint���� = 0, "A.����", "A.�����") & " as ����,  A.��Ч����, A.����, C.�Ƿ���,A.���� ��������, A.Id ����id,A.��� �������," & vbNewLine & _
                    "       Max(B.���id) ���id" & vbNewLine & _
                    "From ��������Ŀ¼ A, ������϶��� B, ����������� C " & vbNewLine & _
                    "Where A.Id = B.����id(+) And A.����id = C.Id(+)  And" & vbNewLine & _
                    " Instr([3],A.���)>0 And (" & strSQL & ")" & _
                    IIf(str�Ա� <> "", " And (A.�Ա�����=[4] Or A.�Ա����� is NULL)", "") & _
                    " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    "Group By A.Id, A.����, A.���, A.����, A.����, A.˵��, A.����id," & IIf(mint���� = 0, "A.����", "A.�����") & ", A.��Ч����, A.����, A.���,C.�Ƿ���"
                 
                strSQL = "Select distinct A.Id,A.��ĿID, A.����, A.���, A.����,A.����ID, A.��������, A.����, A.˵��, A.����, A.����id, A.����,  A.��Ч����, A.����, A.�Ƿ���,A.��������, A.����id,A.�������,A.���id, " & _
                        " Decode(a.����, [6], 1, Decode(A.����,[6],1,decode(a.����,[6],1,NULL))) As ����1ID," & vbNewLine & _
                        "    Decode(d.����id, Null, Decode(c.����id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                        "   Decode(Substr(a.����, 1, Length([6])), [6], 1, Decode(Substr(A.����, 1, Length([6])),[6],1,decode(Substr(a.����, 1, Length([6])),[6],1,NULL))) As ����3ID" & vbNewLine & _
                        " From (" & strSQL & ") A, ����������� C, ����������� D " & _
                        " Where  c.����id(+) = a.Id And d.����id(+) = a.Id And c.����id(+)=[8]  And d.��Աid(+) = [7] " & _
                        " Order By ����1ID, ����2ID, ����3ID, A.����"
            End If
    End Select
    GetMedInputSQL = strSQL
End Function

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagZY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagZY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterUserResize(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    mbln��������� = True
    Call DiagCellButtonClick(vsDiagZY, Row, Col)
    mbln��������� = False
    Call UpDateDiag(vsDiagZY)
End Sub

Private Sub vsDiagZY_Click()
    Call DiagClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_DblClick()
    Call DiagDblClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_GotFocus()
    Call DiagGotFocus(vsDiagZY)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagZY, KeyCode, Shift)
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagZY, KeyAscii)
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagZY, Row, Col, KeyAscii)
End Sub

Private Sub DiagKeyPressEdit(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByRef intKeyAscii As Integer)
    If intKeyAscii = 39 Then intKeyAscii = 0 '�����ű���
    If intKeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagZY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mbln��������� = True
    Call DiagValidateEdit(vsDiagZY, Row, Col, Cancel)
    mbln��������� = False
    Call UpDateDiag(vsDiagZY)
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long
    Dim lngWidth As Long
    Dim lngTop As Long, lngHeight As Long
    Dim lngH As Long
    Dim lngFrmWidth As Long
    Dim blnС���� As Boolean
 
    On Error Resume Next
    If mint���� = 1 Then
        stbThis.Visible = True
    Else
        stbThis.Visible = False
    End If
    vsc.Top = 0
    vsc.Height = Me.ScaleHeight - IIf(mint���� = 1, stbThis.Height, 0)
    vsc.Left = Me.ScaleWidth - vsc.Width
    vsc.Visible = True
    
    Me.BackColor = &H808080  ' &HE0E0E0
    
    PicPanel(picPanel_������Ϣ).Top = mlngTopVsc
    
    blnС���� = mbytSize = 9
 
    lngFrmWidth = vsc.Left - 200
    
    If blnС���� Then
        PicPanel(picPanel_������Ϣ).Width = 9665
        If lngFrmWidth < 9665 Then
            PicPanel(picPanel_������Ϣ).Left = 0
        Else
            PicPanel(picPanel_������Ϣ).Left = 100 + (lngFrmWidth - 9665) / 2
        End If
        PicPanel(picPanel_������Ϣ).Height = 4600
        PicPanel(picPanel_�������).Height = 4500
        If mbln��ҽ Then
            PicPanel(picPanel_������Ϣ).Height = 9300
        Else
            PicPanel(picPanel_������Ϣ).Height = 8340
        End If
        PicPanel(picPanel_����).Height = 5000
    Else
        PicPanel(picPanel_������Ϣ).Width = 12440
        If lngFrmWidth < 12440 Then
            PicPanel(picPanel_������Ϣ).Left = 0
        Else
            PicPanel(picPanel_������Ϣ).Left = 100 + (lngFrmWidth - 12440) / 2
        End If
        PicPanel(picPanel_������Ϣ).Height = 5300
        PicPanel(picPanel_�������).Height = 5400
        If mbln��ҽ Then
            PicPanel(picPanel_������Ϣ).Height = 10500
        Else
            PicPanel(picPanel_������Ϣ).Height = 9540
        End If
        PicPanel(picPanel_����).Height = 5000
    End If
    
    lngWidth = PicPanel(picPanel_������Ϣ).Width
    
    lblN(lbl�������).Left = (lngWidth - lblN(lbl�������).Width) / 2 - 50
    lblN(lbl�������).Caption = "������Ϣ"
    lblN(lbl�������).Top = 500
    
    lblN(lbl�������).Left = lblN(lbl�������).Left
    lblN(lbl�������).Top = 100
    
    lblN(lbl���ⲡ��).Left = lblN(lbl�������).Left
    lblN(lbl���ⲡ��).Top = 100
    
    PicPanel(picPanel_�������).Top = PicPanel(picPanel_������Ϣ).Height + PicPanel(picPanel_������Ϣ).Top
    PicPanel(picPanel_�������).Width = lngWidth
    PicPanel(picPanel_�������).Left = PicPanel(picPanel_������Ϣ).Left

    picOutDoc.Width = lngWidth
    picOutDoc.Top = lblN(lbl�������).Top + lblN(lbl�������).Height + 200
    picOutDoc.Left = 0
    picOutDoc.Height = PicPanel(picPanel_�������).Height - picOutDoc.Top
    
    PicPanel(picPanel_������Ϣ).Top = IIf(mblnDocInput, PicPanel(picPanel_�������).Height, 0) + PicPanel(picPanel_�������).Top
    PicPanel(picPanel_������Ϣ).Left = PicPanel(picPanel_������Ϣ).Left
    PicPanel(picPanel_������Ϣ).Width = lngWidth
    
    lngH = PicPanel(picPanel_������Ϣ).Height + PicPanel(picPanel_������Ϣ).Height + IIf(mblnDocInput, PicPanel(picPanel_�������).Height, 0)
    lngH = lngH - Me.ScaleHeight
    
    vsc.Max = lngH \ (Screen.TwipsPerPixelY)
    
End Sub

Private Function AbstractSelect(ByVal strFind As String) As Boolean
'����ժҪѡ����
    Dim blnCancle As Boolean
    Dim strRetrun As String
    Dim lngLeft As Long, lngTop As Long
    Dim strName As String
    
    Dim objTxt As Object
    Set objTxt = txtE(I����ժҪ)
    
    lngLeft = objTxt.Left + objTxt.Container.Left + 5800
    lngTop = objTxt.Top + objTxt.Container.Top - 210
    
    strRetrun = mobjKernel.ShowCommItem(Me, strFind, blnCancle, lngLeft, lngTop, 4)
    If Not blnCancle Then
        If strRetrun = "" Then
            If strFind = "" Then
                MsgBox "û���ҵ����õľ���ժҪ��", vbInformation, Me.Caption
                Exit Function
            End If
            objTxt.Text = strFind
        Else
            objTxt.Text = strRetrun
        End If
        Call UpDateInfo(objTxt.Text, "txtE", I����ժҪ)
    End If
    AbstractSelect = blnCancle
End Function

Private Sub SetDocEditable()
'���ܣ���ݲ����Ŀɱ༭��
    Dim blnDoc As Boolean
    Dim k As Long, i As Long
    
    If mblnDocInput Then
        blnDoc = mlng����ID <> 0 And (mlng����ID = 0 And mlng�����ļ�id <> 0 Or mlng����ID <> 0 And mblnǩ�� = False) And (mlngִ��״̬ = 1 Or mlngִ��״̬ = 5)
        If blnDoc And mlng����ID <> 0 And lblDoctor(1).Tag = "0" Then    'û���޸����˲�����Ȩ��
            blnDoc = mstr������ = UserInfo.����
        End If
        k = 0
        For i = 0 To rtfEdit.Count - 1
            rtfEdit(i).Locked = Not blnDoc Or InStr(rtfEdit(i).Tag, ",") > 0   '���ڶ�������ʱ(������ȫ�ı༭����)�����������޸�
            If rtfEdit(i).Locked = False Then
                rtfEdit(i).BackColor = vbWindowBackground
                k = k + 1
            Else
                rtfEdit(i).BackColor = DColor
            End If
        Next
        If mlng����ID = 0 Or mlng����ID <> 0 And mblnǩ�� = False Then
            cmdSign.Caption = "ǩ��(&S)"
        Else
            cmdSign.Caption = "ȡ��ǩ��(&S)"
        End If
        cmdSign.Enabled = mlng����ID <> 0 And (mlng����ID = 0 And mlng�����ļ�id <> 0 Or mlng����ID <> 0) And (mlngִ��״̬ = 1 Or mlngִ��״̬ = 5)
        
        If cmdSign.Enabled And mlng����ID <> 0 And lblDoctor(1).Tag = "0" Then   'û���޸����˲�����Ȩ��
            cmdSign.Enabled = mstr������ = UserInfo.����
        End If
        cmdUpdate.Enabled = cmdSign.Enabled
        cmdImportEPRDemo.Enabled = cmdSign.Enabled
    End If
End Sub

Private Sub cboSpecificInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'cboSpecificInfo_KeyPress�¼�
    Dim lngidx As Long
    Dim cboTmp As ComboBox
    If intKeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        Set cboTmp = cboE(intIndex)
        If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
            If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
                If zlCommFun.ActualLen(cboTmp.Text) >= 18 Then
                    intKeyAscii = 0
                Else
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                        intKeyAscii = 0
                    ElseIf zlCommFun.IsCharChinese(cboTmp.Text) Then
                        cboTmp.Text = "": cboTmp.Tag = ""
                    End If
                    If intKeyAscii <> 0 Then
                        Select Case zlCommFun.ActualLen(cboTmp.Text)
                            Case 12
                                cboTmp.Tag = cboTmp.Text & Chr(intKeyAscii)
                            Case 13
                                cboTmp.Tag = cboTmp.Tag & Chr(intKeyAscii)
                        End Select
                    End If
                End If
            Else
                If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                        intKeyAscii = 0
                    ElseIf zlCommFun.IsCharChinese(cboTmp.Text) Then
                        cboTmp.Text = "": cboTmp.Tag = ""
                    End If
                    If intKeyAscii <> 0 Then
                        cboTmp.Tag = cboTmp.Text & Chr(intKeyAscii)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cboSpecificInfoChange(ByRef intIndex As Integer)
'cboSpecificInfo_Change�¼�
    Dim cboTmp As ComboBox
    Dim lngPos As Long, lngLen As Long

    If mblnReturn Then Exit Sub
    Select Case intIndex
        Case I���֤��
            Set cboTmp = cboE(intIndex)
            mblnReturn = True
            If Cbo.FindIndex(cboTmp, cboTmp.Text, True) = -1 Then
                '�����������
                If Not zlStr.CheckCharScope(cboTmp.Text, "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ*") Then
                    cboTmp.Text = ""
                Else
                    If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
                        If zlCommFun.ActualLen(cboTmp.Text) > 18 Then
                            cboTmp.Text = Mid(cboTmp.Text, 1, 18)
                        End If
                    End If
                End If
            End If
            If Trim(zlCommFun.GetNeedName(cboE(I����).Text)) = "�й�" Then
                lngPos = InStr(cboTmp.Text, "*")
                lngLen = Len(Mid(cboTmp.Text, 13, 2))
                Select Case lngPos
                    Case 0
                        cboTmp.Tag = cboTmp.Text
                    Case Else
                        cboTmp.Tag = Mid(cboTmp.Text, 1, lngPos - 1)
                        cboTmp.Text = cboTmp.Tag
                        cboTmp.SelStart = Len(cboTmp.Text)
                End Select
            Else
                cboTmp.Tag = cboTmp.Text
            End If
            mblnReturn = False
    End Select
End Sub

Private Function SetInputRoot(ByVal intType As Integer, ByVal intSysPara As Integer, ByRef intModPara As Integer, ParamArray arrControls() As Variant) As Boolean
'˵�����ú�������ϵͳ������ģ�������ͬ����һ�鵥ѡ��ť��ϵͳ����ֵһ��ΪA(0��1),A+1,A+2....,ģ�����ΪB,B+1,....ϵͳ����ΪAʱ��ģ�����������,������������
'           ģ�����=B(ϵͳ����=A)������ҵ��Ч����ϵͳ����=A+1��ͬ
'           ģ�����=B+1(ϵͳ����=A)������ҵ��Ч����ϵͳ����=A+2��ͬ
'���ܣ�������Դ����������ģ�������ҽ�����Դ����ҽ�����Դ������������Դ
'������intType=0-��ҽ�����Դ���ã�1-��ҽ�����Դ��2-���������Դ
'      intSysPara=ϵͳ����������ֵΪA(0��1),A+1,A+2��..��ֵΪAʱģ�����������
'      intModPara=ģ�����
'���أ��Ƿ�ɹ�
'      intModPara=ʵ�ʲ���ֵ����ϵͳ����Ϊ��0��1��2��ģ��Ϊ0��1 ��ϵͳΪ0ʱģ�������ã���ʱģ�����ʵ��ֵ=ģ�����ֵ����ϵͳ����<>0����1��ģ�����ʵ��ֵ=ϵͳ����-1

    Dim blnVisual As Boolean, blnEnable As Boolean
    Dim i As Long
 
    On Error GoTo errH
    '����������Դ����������̫Ԫͨʱ�ؼ����ɼ�,��������ɼ�
    blnVisual = intType = 2 And gbytPass = 3 Or intType <> 2
    blnEnable = intSysPara = IIf(intType <> 2, 1, 0)
    If Not blnVisual Then intModPara = 0
    If Not blnEnable Then intModPara = intSysPara - IIf(intType <> 2, 2, 1)
    '���ÿؼ���ֵ�Լ�������
    For i = LBound(arrControls) To UBound(arrControls)
        arrControls(i).Visible = blnVisual
        If blnVisual Then
            arrControls(i).Enabled = blnEnable And arrControls(i).Enabled
            'ʵ��ģ�����ֵ��ؼ������±���ʼֵһ����˳��һ��
            If i = intModPara Then
                arrControls(i).Value = 1
            Else
                arrControls(i).Value = 0
            End If
        End If
    Next
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function UpdateLastItem() As String
'���ܣ������ϴα༭�Ŀؼ���Ϣ
    Call SavePreItem
End Function

Public Function IsDataSaved() As Boolean
    IsDataSaved = mblnOK
End Function


Private Sub MsgDis()
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strSQL As String
    Dim str����ID As String
    Dim str���ID As String
    On Error GoTo ErrHand
    '�жϵ�ǰ�����Ƿ���д��Ⱦ�����濨
    strSQL = "Select �ļ�ID From ���Ӳ�����¼ Where ����ID=[1] And ��ҳID=[2] And ��������=5  and ������=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "MsgDis", mlng����ID, mlng�Һ�ID, UserInfo.����)
    If rsTmp.RecordCount > 0 Then
        '�ж��û��Ƿ��޸Ļ�ɾ�����
        With vsDiagXY
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, DI_���ID)) <> 0 Then
                    If InStr("," & str���ID & ",", "," & Val(.TextMatrix(i, DI_���ID)) & ",") = 0 Then
                        str���ID = str���ID & "," & Val(.TextMatrix(i, DI_���ID))
                    End If
                End If
                If Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                    If InStr("," & str����ID & ",", "," & Val(.TextMatrix(i, DI_����ID)) & ",") = 0 Then
                        str����ID = str����ID & "," & Val(.TextMatrix(i, DI_����ID))
                    End If
                End If
            Next
        End With
        
        If mbln��ҽ Then
            With vsDiagZY
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, DI_���ID)) <> 0 Then
                        If InStr("," & str���ID & ",", "," & Val(.TextMatrix(i, DI_���ID)) & ",") = 0 Then
                            str���ID = str���ID & "," & Val(.TextMatrix(i, DI_���ID))
                        End If
                    End If
                    If Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                        If InStr("," & str����ID & ",", "," & Val(.TextMatrix(i, DI_����ID)) & ",") = 0 Then
                            str����ID = str����ID & "," & Val(.TextMatrix(i, DI_����ID))
                        End If
                    End If
                Next
            End With
        End If
        str����ID = Mid(str����ID, 2): str���ID = Mid(str���ID, 2)
        strSQL = ""
        If str����ID <> "" Then
            strSQL = " Union Select ����id,���id From ��������ǰ�� Where ����ID IN (Select Column_Value From Table(f_Num2list([3])))"
        End If
        If str���ID <> "" Then
            strSQL = strSQL & " Union Select ����id,���id From ��������ǰ�� Where ���ID IN (Select Column_Value From Table(f_Num2list([4])))"
        End If
        strSQL = "Select a.����id, a.���id From ������ϼ�¼ A, ��������ǰ�� B Where a.����id = [1] And a.��ҳid = [2] And a.������� = 1 And (a.����id = b.����id Or a.���id = b.���id) " & IIf(strSQL = "", "", "Minus (" & Mid(strSQL, 8) & ") ")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "MsgDis", mlng����ID, mlng�Һ�ID, str����ID, str���ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "��ǰ���˴�Ⱦ��������ݷ����˸ı�,���޸Ĵ�Ⱦ�����濨��", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

