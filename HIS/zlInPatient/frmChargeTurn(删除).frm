VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "ZLIDKIND.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeTurn 
   AutoRedraw      =   -1  'True
   Caption         =   "��(��)�����תסԺ"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11715
   Icon            =   "frmChargeTurn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11715
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBill 
      Height          =   2100
      Left            =   90
      ScaleHeight     =   2040
      ScaleWidth      =   10485
      TabIndex        =   21
      Top             =   1365
      Width           =   10545
      Begin VSFlex8Ctl.VSFlexGrid mshList 
         Height          =   1470
         Left            =   75
         TabIndex        =   22
         Top             =   90
         Width           =   5490
         _cx             =   9684
         _cy             =   2593
         Appearance      =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   Begin VB.PictureBox picBalance 
      Height          =   1950
      Left            =   6285
      ScaleHeight     =   1890
      ScaleWidth      =   2985
      TabIndex        =   19
      Top             =   4035
      Width           =   3045
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   1335
         Left            =   0
         TabIndex        =   20
         Top             =   135
         Width           =   2565
         _cx             =   4524
         _cy             =   2355
         Appearance      =   3
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "ת���ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   105
         TabIndex        =   23
         Top             =   1605
         Width           =   1155
      End
   End
   Begin VB.PictureBox picList 
      Height          =   1935
      Left            =   105
      ScaleHeight     =   1875
      ScaleWidth      =   5415
      TabIndex        =   17
      Top             =   3945
      Width           =   5475
      Begin VSFlex8Ctl.VSFlexGrid mshDetail 
         Height          =   1185
         Left            =   30
         TabIndex        =   18
         Top             =   165
         Width           =   5130
         _cx             =   9049
         _cy             =   2090
         Appearance      =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   11715
      TabIndex        =   12
      Top             =   0
      Width           =   11715
      Begin VB.Frame fraFixed 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   60
         TabIndex        =   24
         Top             =   480
         Width           =   9435
         Begin VB.CheckBox chkShow 
            Caption         =   "����ʾ��ת������"
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
            Left            =   0
            TabIndex        =   3
            Top             =   75
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "ˢ��(&R)"
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
            Left            =   7155
            TabIndex        =   6
            Top             =   0
            Width           =   1300
         End
         Begin VB.ComboBox cbo�������� 
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
            Left            =   4040
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   15
            Width           =   2040
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            Caption         =   "��������"
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
            Left            =   2780
            TabIndex        =   4
            Top             =   75
            Width           =   960
         End
      End
      Begin VB.Frame fraPati 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   60
         TabIndex        =   14
         Top             =   80
         Width           =   2820
         Begin VB.TextBox txtPatient 
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1140
            MaxLength       =   64
            TabIndex        =   25
            ToolTipText     =   "�ȼ���F11"
            Top             =   0
            Width           =   1650
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   345
            Left            =   510
            TabIndex        =   26
            Top             =   0
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   609
            Appearance      =   2
            IDKindStr       =   $"frmChargeTurn.frx":058A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   12
            FontName        =   "����"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            NotContainFastKey=   "F1;CTRL+F1;F2;F3;CTRL+F4;F5;F6;F7;CTRL+F7;F8;F9;F10;F11;F12;CTRL+F12;CTRL+S;CTRL+A;CTRL+R;CTRL+D;CTRL+Q;ESC;ALT+?"
            MustSelectItems =   "����,���￨"
            BackColor       =   -2147483633
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   0
            TabIndex        =   27
            Top             =   45
            Width           =   480
         End
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   4155
         TabIndex        =   1
         Top             =   90
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   182386691
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   7110
         TabIndex        =   2
         Top             =   90
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   182386691
         CurrentDate     =   36588
      End
      Begin zlIDKind.IDKindNew IDKindTime 
         Height          =   240
         Left            =   2880
         TabIndex        =   28
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   423
         ShowSortName    =   0   'False
         IDKindStr       =   "����ʱ��|����ʱ��|0|0|0|0|0|0|0|0|0;�Ǽ�ʱ��|�Ǽ�ʱ��|0|0|0|0|0|0|0|0|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "����"
         IDKind          =   -1
         DefaultCardType =   "0"
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   6870
         TabIndex        =   13
         Top             =   135
         Width           =   120
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11715
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7665
      Width           =   11715
      Begin VB.CommandButton cmdParaSet 
         Caption         =   "��������(&R)"
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
         Left            =   4590
         TabIndex        =   16
         Top             =   0
         Width           =   1500
      End
      Begin VB.CommandButton cmdOk 
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
         Left            =   8220
         TabIndex        =   15
         Top             =   -15
         Width           =   1300
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
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
         Left            =   150
         TabIndex        =   9
         Top             =   0
         Width           =   1300
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ȫ��(&C)"
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
         Index           =   1
         Left            =   3210
         TabIndex        =   7
         Top             =   0
         Width           =   1300
      End
      Begin VB.CommandButton cmdAll 
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
         Index           =   0
         Left            =   1845
         TabIndex        =   0
         Top             =   0
         Width           =   1300
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
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
         Left            =   9570
         TabIndex        =   8
         Top             =   0
         Width           =   1300
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   8100
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeTurn.frx":0620
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15584
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
            AutoSize        =   2
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
Attribute VB_Name = "frmChargeTurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrNOs As String 'Ҫ���з���ת��ĵ�����Ϣ,��ʽ������,Ʊ��,����ID,����(��ҽ��Ϊ��),��������,������㵥�ݺ�:H0000001,F000023,81235,901,�շѵ�(���ʵ�),S0000001;...
Private mlngPatient As Long
Private mobjParent As Object
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private mblnSelPati As Boolean '�Ƿ�ѡ����
Private mintPatientRange As Integer
Private mrsInfo As ADODB.Recordset
Private mstrPrivs As String, mlngModule As Long
Private mbln����תסԺ����� As Boolean
Private mbln�������� As Boolean
Private Enum mObjPancel
    Pan_Search = 1
    Pan_Bill = 2
    Pan_List = 3
    Pan_Balance = 4
    Pan_Bottom = 5
End Enum
Private mstr�����ʻ� As String

'�������ѿ��Ĵ������
Private Type Ty_SquareCard
    blnExistsObjects As Boolean     '��װ�����ѿ���
    rsSquare As ADODB.Recordset
    dblˢ���ܶ� As Double
    bln������ As Boolean '��ǰ��ȡ�ĵ����ǿ�����
    strˢ������ As String   'ˢ�����㷽ʽ;���;�Ƿ������޸�|..."
End Type
Private mtySquareCard As Ty_SquareCard
Private mintIDKind As Integer
Private mobjSquare As Object
Private mblnNotClick As Boolean
Private mstrTitle As String
Private mrsFeeList As ADODB.Recordset
Private mobjThirdSwap As clsThirdSwap
Private mblnRefreshData As Boolean

Private mobjExpenceSvr As Object 'zlPublicExpense.clsExpenceSvr

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2011-03-25 17:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim panThis As Pane
    Dim panTop As Pane, panRight As Pane
    
    Set panTop = dkpMan.CreatePane(mObjPancel.Pan_Search, 200, 580, DockTopOf, Nothing)
    panTop.Title = "��������"
    panTop.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panTop.Tag = mObjPancel.Pan_Search
    panTop.Handle = picTop.hWnd
    If mbln����תסԺ����� Then
        panTop.MaxTrackSize.Height = 495 / Screen.TwipsPerPixelY
        panTop.MinTrackSize.Height = 495 / Screen.TwipsPerPixelY
    Else
        panTop.MaxTrackSize.Height = 850 / Screen.TwipsPerPixelY
        panTop.MinTrackSize.Height = 850 / Screen.TwipsPerPixelY
    End If
    
    Set panThis = dkpMan.CreatePane(mObjPancel.Pan_Bill, 250, 580, DockBottomOf, panTop)
    panThis.Title = "����תסԺ�б�"
    panThis.Tag = mObjPancel.Pan_Bill
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picBill.hWnd
    

    Set panRight = dkpMan.CreatePane(mObjPancel.Pan_Balance, 1500 / Screen.TwipsPerPixelX, 580, DockRightOf, panThis)
    panRight.Title = "����תסԺ������Ϣ"
    panRight.Tag = mObjPancel.Pan_Balance
    panRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panRight.Handle = picBalance.hWnd
    
    Set panThis = dkpMan.CreatePane(mObjPancel.Pan_List, 250, 580, DockBottomOf, panThis)
    panThis.Title = "������ϸ�б�"
    panThis.Tag = mObjPancel.Pan_List
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picList.hWnd
 
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
End Sub

Private Sub cbo��������_Click()
    If mblnNotClick Then Exit Sub
    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
End Sub

Private Sub chkShow_Click()
    If mblnNotClick Then Exit Sub
    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pan_Search
        Item.Handle = picTop.hWnd
    Case Pan_Bill
        Item.Handle = picBill.hWnd
    Case Pan_List
        Item.Handle = picList.hWnd
    Case Pan_Balance
        Item.Handle = picBalance.hWnd
    End Select
End Sub

Public Sub ShowMe(objParent As Object, ByVal lngPatient As Long, ByRef strNos As String, _
    Optional blnSelPati As Boolean = False, Optional intPatientRange As Integer = 0, _
    Optional strPrivs As String, Optional lngModule As Long, Optional ByRef blnRefreshData As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������תסԺ����
    '���:lngPatient-����ID
    '      blnSelPati-�Ƿ���Ҫѡ����
    '      intPatientRange:(0-���в���,1-�κη���δ���岡��;2-���δ����Ĳ���;3-סԺδ����Ĳ���;4-����δ����Ĳ���)
    '����:
    '   strNOS:Ҫ���з���ת��ĵ�����Ϣ,��ʽ��
    '       ����,Ʊ��,����ID,����(��ҽ��Ϊ��),��������,������㵥�ݺ�:H0000001,F000023,81235,901,�շѵ�(���ʵ�),S0000001;...
    '   blnRefreshData-�������תסԺ���Ƿ�ˢ������
    '����:
    '����:���˺�
    '����:2010-11-09 17:09:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnSelPati = blnSelPati: mintPatientRange = intPatientRange
    mlngPatient = lngPatient: mstrPrivs = strPrivs: mlngModule = lngModule
    mstrNOs = strNos: mblnRefreshData = False: txtPatient.Tag = lngPatient
    Set mobjParent = objParent
    
    On Error Resume Next
    Call Me.Show(vbModal, objParent)
    strNos = mstrNOs
    blnRefreshData = mblnRefreshData
End Sub

Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-11-09 17:30:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mshList.Redraw = flexRDNone
    mshList.Clear 1: mshList.Rows = 2
    sta.Panels(2).Text = ""
    Call setHeader: Call SetBillColor
    mshList.Redraw = flexRDBuffered
    Set mrsFeeList = Nothing
    cbo��������.Clear
End Sub

Private Sub SetBillSelected(ByVal strNos As String)
'˵��:���ת�뼸���ʧ��,�ٽ���ѡ����,��ǰѡ������ѱ�ת��ĵ���������"����ת��",���Բ�Ӧ��ѡ��
    Dim i As Long
    With mshList
        For i = 1 To .Rows - 1
            If InStr(";" & strNos, ";" & .TextMatrix(i, .ColIndex("���ݺ�"))) > 0 And .TextMatrix(i, .ColIndex("���")) = "��ת��" Then
                .TextMatrix(i, .ColIndex("ѡ��")) = "��"
            Else
                .TextMatrix(i, .ColIndex("ѡ��")) = ""
            End If
        Next
    End With
End Sub

Public Function CheckExistTurn(ByVal lngPatient As Long, ByRef dat��Ժʱ�� As Date) As Boolean
'����:�����Ժʱ��֮���Ƿ����ת������
'����:ת�����ݵĵǼ�ʱ��
    Dim rsTmp As ADODB.Recordset, strSql As String
        
    On Error GoTo errH
    strSql = "" & _
    " Select Max(����ʱ��) ����ʱ�� " & _
    " From סԺ���ü�¼" & vbNewLine & _
    " Where ��¼���� = 2 And ��¼״̬ In(1,3) And ����id = [1] And ��ҳid Is Null And ��ʶ�� Is Null And �����־=2" & vbNewLine & _
    "       And ժҪ='�������ת��'"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����Ƿ������ת����", lngPatient, dat��Ժʱ��)
    
    If Not IsNull(rsTmp!����ʱ��) Then
        dat��Ժʱ�� = rsTmp!����ʱ��
        CheckExistTurn = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsYBSingle(ByVal strNO As String, Optional blnYBAllDel_Out As Boolean, Optional ByRef blnThirdAllDel_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҽ���Ƿ�ȫ�˻��Ƿֵ��ݾ�
    '���:strNo-ָ������
    '����:blnThirdAllDel-�������Ƿ����ȫ��
    '     blnYBAllDel_Out-ҽ���Ƿ����ȫ��
    '����:�ֵ����ˣ�����true,���򷵻�False
    '����:���˺�
    '����:2018-09-13 14:16:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    blnThirdAllDel_Out = False: blnYBAllDel_Out = False
    
    strSql = "Select 1 From ҽ��������ϸ Where NO = [1] And Rownum < 2 And �����ID is NULL "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    
    blnYBAllDel_Out = rsTmp.EOF
    If rsTmp.EOF Then IsYBSingle = False: Exit Function
    
    blnThirdAllDel_Out = CheckAllTurn(strNO)
    IsYBSingle = Not blnThirdAllDel_Out
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPatiObjectFromNo(ByVal strNO As String, ByVal int���� As Integer, _
    ByRef objPati_out As clsPatiInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շѵ�����ȡָ���Ĳ�����Ϣ
    '���:strNo-�շѵ��ݺ�
    '     int����-��������:=1-�շѵ�;2-���ʵ�
    '����:objPati_out-���ز�����Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-17 14:54:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql  As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    Set objPati_out = New clsPatiInfo
    strSql = _
        " Select b.����id, b.����, b.�Ա�, b.�Ա�, b.����" & _
        " From ������ü�¼ A, ������Ϣ B" & _
        " Where a.����id = b.����id And a.No = [1] And ��¼���� = [2] And ��¼״̬ In (0, 1, 3) And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, int����)
    If rsTemp.EOF Then Exit Function
    With objPati_out
        .����ID = Val(NVL(rsTemp!����ID))
        .���� = Val(NVL(rsTemp!����))
        .�Ա� = Val(NVL(rsTemp!�Ա�))
        .���� = Val(NVL(rsTemp!����))
        .�Ա� = Val(NVL(rsTemp!�Ա�))
    End With
    GetPatiObjectFromNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExcuteTranSaveOver(ByVal objPati As clsPatiInfo, ByRef objBalanceInfor As clsBalanceInfo, ByRef cllBillPro As Collection, Optional blnNotModify As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ת����ɱ���
    '���:objBalanceInfor-������Ϣ
    '     objPati-������Ϣ
    '     blnNotModify-�Ƿ񲻽�����������
    '����:
    '����:ת�ʳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-17 16:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, cllPro As Collection
    Dim blnTrans As Boolean, i As Long
    
    On Error GoTo errHandle
    
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    
    Set cllPro = New Collection
    
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    If Not blnNotModify Then
    
        '    '���ִ��
        '    Zl_�������תסԺ_Modify
        strSql = "Zl_�������תסԺ_Modify("
        '    ��������_In   Number,  '0-������У�Ա�־:ֻ���¹�������ID��У�Ա�־;1-��ͨ�˷ѷ�ʽ:2.�������˷ѽ���:;3-ҽ������;4-���ѿ�����:
        strSql = strSql & "1,"
        '    ����id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & objBalanceInfor.����ID & ","
        '    ����id_In     ���˽��ʼ�¼.����id%Type,
        strSql = strSql & "" & objPati.����ID & ","
        '    ���㷽ʽ_In   Varchar2,
        strSql = strSql & "NULL,"
        '    ����Ա���_In ����Ԥ����¼.����Ա���%Type := Null,
        strSql = strSql & "'" & UserInfo.��� & "' ,"
        '    ����Ա����_In ����Ԥ����¼.����Ա����%Type := Null,
        strSql = strSql & "'" & UserInfo.���� & "' ,"
        '    ����˷�_In   Number := 0,0-δ����˷�;1-�쳣����˷�;2-����˷�
        strSql = strSql & "2)"
        '    ��������id_In ����Ԥ����¼.Id%Type := Null,
        '    �˿�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null,
        '    У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := Null,
        '    �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        '    �����id_In   ����Ԥ����¼.�����id%Type := Null,
        '    ����_In       ����Ԥ����¼.����%Type := Null,
        '    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        '    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        '    ���ԭ����_In Number:=0
        zlAddArray cllPro, strSql
    End If
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    blnTrans = False
    ExcuteTranSaveOver = True
    Set cllBillPro = New Collection
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteTurn(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal strNos As String, ByVal strסԺ�� As String, ByVal lng��ҳID As Long, _
    ByVal dat��Ժʱ�� As Date, ByVal lng��Ժ����ID As Long, ByVal lng��Ժ����ID As Long, _
    Optional ByRef strOutDelDate As String, Optional ByRef blnReflashData_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ĵ��ݺ�����,ִ���������תסԺ����,��ҽ���˷ѽ������
    '���:
    '   strNos:Ҫ���з���ת��ĵ�����Ϣ,��ʽ��
    '       ����,Ʊ��,����ID,����(��ҽ��Ϊ��),��������,������㵥�ݺ�:H0000001,F000023,81235,901,�շѵ�(���ʵ�),S0000001;...
    '   lngסԺ��-סԺ��,lng��ҳID-��ҳID,��������������ҽ����Ժ����Ǽ�ʱ�Ŵ���
    '����:strDelDate-����ת������(Ŀǰ��Ҫ�����»�ȡԤ��������)
    '   blnReflashData_Out-�Ƿ�����ˢ������
    '����:
    '����:���˺�
    '����:2011-02-16 10:26:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNO As Variant, arrInfo As Variant
    Dim i As Long, j As Long, lngcnt As Long, blnҽ�������� As Boolean
    Dim strSql As String, strInvoice As String, strInDate As String, strDelDate As String
    Dim cllPro As Collection, str��ת����ID As String
    Dim intInsure As Integer, blnTurnAll As Boolean
    Dim objBalanceInfor As clsBalanceInfo, objPati As clsPatiInfo
    Dim objSquareDelItems As clsBalanceItems
    Dim strSfNos As String, blnBillPrintInited As Boolean
    Dim strCheckNos As String '��ʽ����¼����,���ݺ�|��¼����,���ݺ�|... ���У���¼���ʣ�1-�����շѣ�2-�������
    Dim lngCount As Long, blnDataSaved As Boolean
    Dim lngStep As Long, bln���ڽ��ʵ� As Boolean
    Dim strNewNo As String, strNewNos As String, varNos As Variant, p As Integer
    
    '�������ĵ��ݴ���˼·���Ƚ����õ���תΪסԺ���ü�¼���ٵ������������˷�
    Dim strReplenishNo As String, strReplenishNos As String 'Array(�����㵥�ݺ�,ת����SQL,�µ��ݺ�)
    Dim cllReplenishPro As Collection
    
    On Error GoTo errHandle
    Set mobjThirdSwap = New clsThirdSwap
    If mobjThirdSwap.zlInitCompents(Me, lngModule, mobjICCard) = False Then Exit Function
     
    mstrPrivs = strPrivs: mlngModule = lngModule
    If strNos = "" Then Exit Function
    strInDate = "To_Date('" & Format(dat��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    strOutDelDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strDelDate = "To_Date('" & strOutDelDate & "','YYYY-MM-DD HH24:MI:SS')"
    
    arrNO = Split(strNos, ";")
    
    '����,Ʊ��,����ID,����(��ҽ��Ϊ��),��������,������㵥�ݺ�
    arrInfo = Split(arrNO(0), ",")
    If GetPatiObjectFromNo(arrInfo(0), IIf(arrInfo(4) = "���ʵ�", 2, 1), objPati) = False Then
        MsgBox "δ�ҵ�ָ���Ĳ��ˣ��������������תסԺ��", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mbln�������� Then Call zlBillPrint_Initialize(Val("1137-���˽��ʹ���"))
    
    For i = 0 To UBound(arrNO)
        arrInfo = Split(arrNO(i), ",")
        strCheckNos = strCheckNos & "|" & IIf(arrInfo(4) = "���ʵ�", 2, 1) & "," & arrInfo(0)
    Next
    If strCheckNos <> "" Then strCheckNos = Mid(strCheckNos, 2)
    If mobjExpenceSvr.zlChargeTurnCheck(strCheckNos, objPati.����ID, lng��ҳID, Me.Caption) = False Then Exit Function
    
    Set objBalanceInfor = New clsBalanceInfo
    With objBalanceInfor
        .����ʱ�� = CDate(strOutDelDate)
        .�������� = 3  '��������:1-�������;2-סԺ����;3-�������תסԺ
    End With
    
    Set cllPro = New Collection
    Set cllReplenishPro = New Collection
    
    blnReflashData_Out = False
    lngCount = UBound(arrNO) + 1
    
    zlControl.StaShowPercent 0, sta.Panels(2), Me
    lngStep = 0
    i = LBound(arrNO)
    Do While i <= UBound(arrNO)
        lngStep = lngStep + 1
        
        lngcnt = 1
        strInvoice = Trim(Split(arrNO(i), ",")(1))
        If strInvoice <> "" Then
            For j = i + 1 To UBound(arrNO)
                If strInvoice = Split(arrNO(j), ",")(1) Then
                    lngcnt = lngcnt + 1
                Else
                    Exit For
                End If
            Next
        End If
        
        
        'ҽ��Ҫ������һ�ſ�ʼ��,�����������ǰ����ݺŵ������еģ����Դ˴����򼴿�
        For j = i To i + lngcnt - 1
            '����,Ʊ��,����ID,����(��ҽ��Ϊ��),��������,������㵥�ݺ�
            arrInfo = Split(arrNO(j), ",")
            blnҽ�������� = False: blnTurnAll = False
            
            strReplenishNo = arrInfo(5)
            If strReplenishNo = "" Then
                If Val(arrInfo(3)) <> 0 Then    '���ʵ�������Ϊ0
                    blnҽ�������� = IsYBSingle(arrInfo(0))
                Else
                    blnTurnAll = CheckAllTurn(arrInfo(0))
                    If InStr("," & str��ת����ID & ",", "," & arrInfo(2) & ",") > 0 Then blnTurnAll = True
                End If
            End If
            
            With objBalanceInfor
                .����ID = Val(arrInfo(2))
                .���ʵ��ݺ� = arrInfo(0)
                .objInsure.���� = Val(arrInfo(3))
            End With
            
            '�ȴ���ļ��ʵ�����ǰ���ݲ��Ǽ��ʵ���˵�����ʵ��Ѵ�����
            If arrInfo(4) <> "���ʵ�" And mbln�������� And Not blnBillPrintInited Then
                Call zlBillPrint_Initialize(Val("1121-�����շѹ���"))
                blnBillPrintInited = True
            End If
            
            If blnҽ�������� Or (objBalanceInfor.objInsure.���� = 0 And Not blnTurnAll) Or strReplenishNo <> "" Then
                
                If InStr("," & str��ת����ID & ",", "," & arrInfo(2) & ",") = 0 Then ' ����һ�ν��ʷֵ��ݵģ��Ѿ�ת��������Ҫ�ж�
                    strNewNo = zlDatabase.NextNo(14)
                    
                    'Zl_�������תסԺ_Insert
                    strSql = "Zl_�������תסԺ_insert("
                    '  No_In         סԺ���ü�¼.NO%Type,
                    strSql = strSql & "'" & arrInfo(0) & "',"
                    '  Newno_In        סԺ���ü�¼.No%Type,
                    strSql = strSql & "'" & strNewNo & "',"
                    '  סԺ��_In     סԺ���ü�¼.��ʶ��%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                    strSql = strSql & "" & ZVal(strסԺ��) & ","
                    '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                    strSql = strSql & "" & ZVal(lng��ҳID) & ","
                    '  ��Ժʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                    strSql = strSql & "" & strInDate & ","
                    '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                    strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                    '  ��Ժ����id_In סԺ���ü�¼.���˲���id%Type,
                    strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                    '  ת��ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ���ת��ʱ,ÿ�ŵ��ݵ�ת��ʱ����ͬ,����ϵͳ��ǰʱ��
                    strSql = strSql & "" & strDelDate & ","
                    '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                    strSql = strSql & "'" & UserInfo.���� & "',"
                    '  ��������_In Number := 1, --1-�����շѵ�;2-������ʵ�
                    strSql = strSql & "" & IIf(arrInfo(4) = "���ʵ�", 2, 1) & ")"
                    
                    If strReplenishNo <> "" And mbln�������� Then
                        If InStr(strReplenishNos & ";", ";" & strReplenishNo & "," & arrInfo(3) & ";") = 0 Then
                            strReplenishNos = strReplenishNos & ";" & strReplenishNo & "," & arrInfo(3)
                        End If
                        'Array(�����㵥�ݺ�,ת����SQL,�µ��ݺ�)
                        cllReplenishPro.Add Array(strReplenishNo, strSql, strNewNo)
                    Else
                        zlAddArray cllPro, strSql
                        If arrInfo(4) = "���ʵ�" And mbln�������� Then
                            'Zl_����תסԺ_����ת��
                            strSql = "Zl_����תסԺ_����ת��("
                            '  No_In         סԺ���ü�¼.No%Type,
                            strSql = strSql & "'" & arrInfo(0) & "',"
                            '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                            strSql = strSql & "'" & UserInfo.��� & "',"
                            '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                            strSql = strSql & "'" & UserInfo.���� & "',"
                            '  ����ʱ��_In   סԺ���ü�¼.����ʱ��%Type
                            strSql = strSql & "" & strDelDate & ")"
                            zlAddArray cllPro, strSql
                            
                            If DelBalaceMz(objPati, cllPro, lng��ҳID, lng��Ժ����ID, objBalanceInfor) = False Then
                                blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                                Exit Function
                            End If
                            bln���ڽ��ʵ� = True
                        ElseIf mbln�������� And arrInfo(4) <> "���ʵ�" Then
                            strSfNos = "'" & arrInfo(0) & "'"
                            If zlBillPrint_EraseBill(strSfNos, 0) = False Then Exit Function
                    
                            objBalanceInfor.����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
                            'Zl_����תסԺ_�շ�ת��
                            strSql = "Zl_����תסԺ_�շ�ת��("
                            '  No_In         סԺ���ü�¼.No%Type,
                            strSql = strSql & "'" & arrInfo(0) & "',"
                            '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                            strSql = strSql & "'" & UserInfo.��� & "',"
                            '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                            strSql = strSql & "'" & UserInfo.���� & "',"
                            '  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                            strSql = strSql & "" & strDelDate & ","
                            '  �����˷�_In   Number := 0,--0-����תסԺ��������;1-�����˷�ģʽ;=1ʱ:��Ժ����id_In����ҳID_IN���Բ�����
                            strSql = strSql & "" & 0 & ","
                            '  ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
                            strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                            '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null,
                            strSql = strSql & "" & ZVal(lng��ҳID) & ","
                            '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
                            strSql = strSql & "" & "NULL" & ","
                            '  ����id_In     ����Ԥ����¼.����id%Type := Null,
                            strSql = strSql & "" & objBalanceInfor.����ID & ")"
                            '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
                            '  ����_In     ����Ԥ����¼.��Ԥ��%Type := Null,
                            '  �ɿ���id_In   ����Ԥ����¼.�ɿ���id%Type := Null
                            zlAddArray cllPro, strSql
                            
                            intInsure = objBalanceInfor.objInsure.����
                             'ִ��ҽ��:
                            If ExcuteInsureDel(objBalanceInfor, intInsure, objBalanceInfor.���ʵ��ݺ�, cllPro) = False Then
                                blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                                Exit Function
                            End If
                            'ִ��һ��ͨ
                            If Not ExecuteThirdReturnMoneySwap(objPati, objBalanceInfor, cllPro, objSquareDelItems) Then
                                blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                                Exit Function
                            End If
                            '���
                            If ExcuteTranSaveOver(objPati, objBalanceInfor, cllPro) = False Then
                                blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                                Exit Function
                            End If
                        Else
                            'ֱ���������תסԺ
                            If Not ExcuteTranSaveOver(objPati, objBalanceInfor, cllPro, True) Then Exit Function
                        End If
                        
                        Call mobjExpenceSvr.zlAdjustFeeData(strNewNo)
                    End If
                End If
            Else
                If InStr("," & str��ת����ID & ",", "," & arrInfo(2) & ",") = 0 Then
                    If arrInfo(4) = "���ʵ�" Then
                        varNos = Array(arrInfo(0))
                    Else '�շѵ���һ��ת����������е���
                        strSfNos = zlGetBalanceNos(1, arrInfo(2))
                        varNos = Split(strSfNos, ",")
                    End If
                    
                    For p = 0 To UBound(varNos)
                        strNewNo = zlDatabase.NextNo(14)
                        strNewNos = strNewNos & "," & strNewNo
                        
                        'Zl_�������תסԺ_Insert
                        strSql = "Zl_�������תסԺ_insert("
                        '  No_In         סԺ���ü�¼.NO%Type,
                        strSql = strSql & "'" & varNos(0) & "',"
                        '  Newno_In        סԺ���ü�¼.No%Type,
                        strSql = strSql & "'" & strNewNo & "',"
                        '  סԺ��_In     סԺ���ü�¼.��ʶ��%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                        strSql = strSql & "" & ZVal(strסԺ��) & ","
                        '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                        strSql = strSql & "" & ZVal(lng��ҳID) & ","
                        '  ��Ժʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                        strSql = strSql & "" & strInDate & ","
                        '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                        strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                        '  ��Ժ����id_In סԺ���ü�¼.���˲���id%Type,
                        strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                        '  ת��ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ���ת��ʱ,ÿ�ŵ��ݵ�ת��ʱ����ͬ,����ϵͳ��ǰʱ��
                        strSql = strSql & "" & strDelDate & ","
                        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSql = strSql & "'" & UserInfo.���� & "',"
                        '  ��������_In Number := 1, --1-�����շѵ�;2-������ʵ�
                        strSql = strSql & "" & IIf(arrInfo(4) = "���ʵ�", 2, 1) & ")"
                        zlAddArray cllPro, strSql
                    Next
                    If strNewNos <> "" Then strNewNos = Mid(strNewNos, 2)
                    
                    If arrInfo(4) = "���ʵ�" And mbln�������� Then
                        'Zl_����תסԺ_����ת��
                        strSql = "Zl_����תסԺ_����ת��("
                        '  No_In         סԺ���ü�¼.No%Type,
                        strSql = strSql & "'" & arrInfo(0) & "',"
                        '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                        strSql = strSql & "'" & UserInfo.��� & "',"
                        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSql = strSql & "'" & UserInfo.���� & "',"
                        '  ����ʱ��_In   סԺ���ü�¼.����ʱ��%Type
                        strSql = strSql & "" & strDelDate & ")"
                        zlAddArray cllPro, strSql
                        
                        If DelBalaceMz(objPati, cllPro, lng��ҳID, lng��Ժ����ID, objBalanceInfor) = False Then
                            blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                            Exit Function
                        End If
                        bln���ڽ��ʵ� = True
                    ElseIf mbln�������� And arrInfo(4) <> "���ʵ�" Then
                        strSfNos = "'" & Replace(strSfNos, ",", "','") & "'"
                        If zlBillPrint_EraseBill(strSfNos, 0) = False Then Exit Function
                        
                        objBalanceInfor.����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
                        'Zl_����תסԺ_�շ�ת��
                        strSql = "Zl_����תסԺ_�շ�ת��("
                        '  No_In         סԺ���ü�¼.No%Type,
                        strSql = strSql & "'" & arrInfo(0) & "',"
                        '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                        strSql = strSql & "'" & UserInfo.��� & "',"
                        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSql = strSql & "'" & UserInfo.���� & "',"
                        '  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                        strSql = strSql & "" & strDelDate & ","
                        '  �����˷�_In   Number := 0,--0-����תסԺ��������;1-�����˷�ģʽ;=1ʱ:��Ժ����id_In����ҳID_IN���Բ�����
                        strSql = strSql & "" & 0 & ","
                        '  ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
                        strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                        '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null,
                        strSql = strSql & "" & ZVal(lng��ҳID) & ","
                        '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
                        strSql = strSql & "" & "NULL" & ","
                        '  ����id_In     ����Ԥ����¼.����id%Type := Null,
                        strSql = strSql & "" & objBalanceInfor.����ID & ","
                        '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
                        strSql = strSql & "" & objBalanceInfor.����ID & ")"
                        '  ����_In     ����Ԥ����¼.��Ԥ��%Type := Null,
                        '  �ɿ���id_In   ����Ԥ����¼.�ɿ���id%Type := Null
                        zlAddArray cllPro, strSql
                        
                        intInsure = objBalanceInfor.objInsure.����
                         'ִ��ҽ��:
                        If ExcuteInsureDel(objBalanceInfor, intInsure, "", cllPro) = False Then
                            blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                            Exit Function
                        End If
                        'ִ��һ��ͨ
                        If Not ExecuteThirdReturnMoneySwap(objPati, objBalanceInfor, cllPro, objSquareDelItems) Then
                            blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                            Exit Function
                        End If
                        '���
                        If ExcuteTranSaveOver(objPati, objBalanceInfor, cllPro) = False Then Exit Function
                    Else
                        'ֱ���������תסԺ
                        If Not ExcuteTranSaveOver(objPati, objBalanceInfor, cllPro, True) Then Exit Function
                    End If
                    
                    Call mobjExpenceSvr.zlAdjustFeeData(strNewNos)
                End If
                str��ת����ID = str��ת����ID & "," & arrInfo(2)
            End If
        Next
        
        zlControl.StaShowPercent lngStep / lngCount, sta.Panels(2), Me
        i = i + lngcnt
    Loop
    
    sta.Panels(2).Text = ""
    
    '�Բ�����㵥�ݽ����˷Ѵ���
    If strReplenishNos <> "" Then
        strReplenishNos = Mid(strReplenishNos, 2)
        If ExecuteReplenishDel(strReplenishNos, cllReplenishPro, lng��ҳID, lng��Ժ����ID, strOutDelDate) = False Then
            Exit Function
        End If
    End If
    
    '��ӡԤ�����
    Call PrintPrePayPrint(strOutDelDate)
    
    '��ʾ���ʴ���
    If bln���ڽ��ʵ� And mbln�������� Then
       Call ShowBalanceWindows(strOutDelDate)
    End If
    
    ExecuteTurn = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteReplenishDel(ByVal strNos As String, ByVal cllPro As Collection, _
    ByVal lng��ҳID As Long, ByVal lng��Ժ����ID As Long, ByVal strDelDate As String) As Boolean
    '����:�Բ������ĵ��ݽ���ת���ü��˷Ѵ���
    '���:
    '   strNos �����㵥��,��ʽ�����ݺ�,����;...
    '   cllPro ������˷ѹ��̵ļ��ϣ�Array(�����㵥�ݺ�,ת����SQL,�µ��ݺ�)
    '   strDelDate �˷�ʱ��
    Dim strSql As String, strNoTemp As String
    Dim varNos As Variant, i As Long, p As Long, blnTrans As Boolean
    Dim strNO As String, intInsure As Integer
    Dim lng�������ID  As Long, lng���ó���ID As Long, lng������� As Long
    Dim lngԭ����ID As Long, strAdvance As String
    Dim strNewNos As String, strNewNo As String
    
    Err = 0: On Error GoTo errH
    If strNos = "" Then ExecuteReplenishDel = True: Exit Function
    
    Call zlBillPrint_Initialize(Val("1124-���ղ������"))
    varNos = Split(strNos, ";")
    For i = 0 To UBound(varNos)
        '���ݺ�,����;...
        strNO = Split(varNos(i), ",")(0): intInsure = Split(varNos(i), ",")(1)
        
        If zlBillPrint_EraseBill(strNO, 0) = False Then Exit Function
        
        lng���ó���ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        lng�������ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        lng������� = -1 * lng���ó���ID
        
        gcnOracle.BeginTrans: blnTrans = True
        For p = 1 To cllPro.Count
            'Array(�����㵥�ݺ�,ת����SQL,�µ��ݺ�)
            strNoTemp = cllPro(p)(0): strSql = cllPro(p)(1): strNewNo = cllPro(p)(2)
            If strNoTemp = strNO Then
                strNewNos = strNewNos & "," & strNewNo
                zlDatabase.ExecuteProcedure strSql, Me.Caption
            End If
        Next
        If strNewNos <> "" Then strNewNos = Mid(strNewNos, 2)
        
        'Zl_����תסԺ_������ת��(
        strSql = "Zl_����תסԺ_������ת��("
        '  No_In         ���ò����¼.No%Type,
        strSql = strSql & "'" & strNO & "',"
        '  ���ó���id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & lng���ó���ID & ","
        '  �������id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & lng�������ID & ","
        '  �������_In     ����Ԥ����¼.�������%Type,
        strSql = strSql & "" & lng������� & ","
        '  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
        strSql = strSql & "To_Date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
        strSql = strSql & "'" & UserInfo.��� & "',"
        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
        strSql = strSql & "'" & UserInfo.���� & "',"
        '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
        strSql = strSql & "" & lng��ҳID & ","
        '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & lng��Ժ����ID & ")"
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        
        'Public Function ClinicDelSwap(lngStlID As Long, Optional ByVal bln�˷� As Boolean = True, _
            Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:�������˷ѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ��
            '���:lngStlID-��Ҫ�˵ķѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
            '     bln�˷� -�������˷ѽ��׻��Ǹķѽ����ڵ��ñ��ӿ�
            '     strAdvance:��ʽ:����ID|��������־|��,ÿλ|�ָ�
            '           ��һλ:�������ID,ҽ�����Ը��ݳ���ID������ȡ��
            '           �ڶ�λ:��������־,1-����������;0�ǲ���������
            '           ����λ:NO:��ǰ�����NO
            '           ����λ��: ���Ժ���չ
            '     ע�⣺
            '           strAdvance��10.34.0��ǰ(�������ʽ���)
            '               �൥��һ�ν���ʱ,�������ԭ����IDs:����ID1,����ID2,...
            '               �����������ʽΪ:�˷ѵ���������|��ǰ�˵ڼ��ŵ���
            '����:strAdvance:1.ԭ���˻�ʱ�����ؿ�
            '                2.�˷ѽ��㷽ʽ���շѽ��㷽ʽ��һ��ʱ�����ظ�ʽΪ�����㷽ʽ|���||���㷽ʽ|���||�������У����Ϊ����
            '���أ����׳ɹ�����true�����򣬷���false
        strAdvance = lng�������ID & "|1"
        lngԭ����ID = zlGetFromNOToLastBalanceID(strNO, , , , True)
        If Not gclsInsure.ClinicDelSwap(lngԭ����ID, True, intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            MsgBox "ҽ������ʧ�ܣ��޷����������������תסԺ������", vbInformation, gstrSysName
            Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
        
        Call mobjExpenceSvr.zlAdjustFeeData(strNewNos)
    Next
    ExecuteReplenishDel = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Function

Private Function zlGetFromNOToLastBalanceID(ByVal strNos As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal bln��ʷ��ͬ���� As Boolean = False, _
    Optional lng������� As Long, Optional bln������ As Boolean = False) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ���շѵ��ݵ�NO���������һ����Ч�Ľ��ʵ�ID
    '���:blnNoMoved�Ƿ��ں󱸱��У���ѯ����֮ǰ���ж���Ҫ���������
    '     bln��ʷ��ͬ����-�Ƿ�������ʷ��һ���ѯ
    '     bln������-�Ƿ񲹳����
    '����:lng�������-�������һ����Ч�Ľ������
    '����:����ID
    '����:���˺�
    '����:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String, strSQL1 As String
    
    On Error GoTo errHandle:
    '87975
    strSql = _
            " Select /*+cardinality(m,10)*/ Max(a.����id) As ����id" & vbNewLine & _
            " From ������ü�¼ A, Table(f_Str2list([1])) M" & vbNewLine & _
            " Where a.No = m.Column_Value" & vbNewLine & _
            "       And a.�Ǽ�ʱ�� + 0 =" & vbNewLine & _
            "           (Select /*+cardinality(j,10)*/ Max(m.�Ǽ�ʱ��)" & vbNewLine & _
            "            From ������ü�¼ M, Table(f_Str2list([1])) J" & vbNewLine & _
            "            Where m.No = j.Column_Value And Mod(m.��¼����, 10) = 1 And m.��¼״̬ In (1, 3) And Nvl(m.����״̬, 0) <> 1)" & vbNewLine & _
            "            And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And Nvl(a.����״̬, 0) <> 1"

    If bln������ Then
        strSql = Replace(strSql, "������ü�¼", "���ò����¼")
        strSql = Replace(strSql, "Max(a.����id)", "Max(a.����id)")
    End If

    strSql = "" & _
            "   Select A.����ID,B.������� " & _
            "   From (" & strSql & ") A,����Ԥ����¼ B " & _
            "   Where A.����ID=B.����ID(+) And Rownum<2"

    If Not blnNOMoved And bln��ʷ��ͬ���� Then
        strSQL1 = Replace(strSql, "������ü�¼", "H������ü�¼")
        strSQL1 = Replace(strSql, "���ò����¼", "H���ò����¼")
        strSQL1 = Replace(strSql, "����Ԥ����¼", "H����Ԥ����¼")
        strSql = strSql & " Union ALL " & strSQL1
    ElseIf blnNOMoved Then
        strSql = Replace(strSql, "������ü�¼", "H������ü�¼")
        strSQL1 = Replace(strSql, "���ò����¼", "H���ò����¼")
        strSql = Replace(strSql, "����Ԥ����¼", "H����Ԥ����¼")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���ݵ��ݻ�ȡ���һ���������ʵĽ���ID", strNos)

    If rsTemp.EOF Then Exit Function

    lng������� = Val(NVL(rsTemp!�������))
    zlGetFromNOToLastBalanceID = Val(NVL(rsTemp!����ID))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExcuteInsureDel(ByVal objBalanceInfor As clsBalanceInfo, _
    ByVal intInsure As Integer, ByVal strNO As String, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ҽ���˷��ò���
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-17 16:31:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cllPro As Collection
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strAdvance As String
    
    On Error GoTo errHandle
        
    If intInsure = 0 Then ExcuteInsureDel = True: Exit Function
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    Set cllPro = New Collection
    
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    blnTrans = True: blnTransMedicare = False
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    strAdvance = objBalanceInfor.����ID & "|0" & IIf(strNO <> "", "|" & strNO, "")
    If Not gclsInsure.ClinicDelSwap(objBalanceInfor.����ID, , intInsure, strAdvance) Then
        gcnOracle.RollbackTrans
        MsgBox "ҽ������ʧ�ܣ��޷������������ת��Ժ������", vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTransMedicare = True: blnTrans = False
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
    objBalanceInfor.�Ƿ񱣴���ʵ� = True
    Set cllBillPro = New Collection
    ExcuteInsureDel = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTransMedicare And mbln�������� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intInsure)
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetYBBalance(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    Optional ByVal blnDelCheck As Boolean = True, Optional ByVal blnDel As Boolean = True, _
    Optional ByVal intInsure As Integer, Optional ByVal bln����������� As Boolean, _
    Optional ByVal str�����ʻ� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��ԭ���㷽ʽ�ͽ�����
    '����:���ؽ�����Ϣ,��ʽ:���㷽ʽ|������||...
    '����:���˺�
    '����:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    Dim strSql As String, rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    strSql = _
        " Select ���㷽ʽ, Sum(��Ԥ��) As ��Ԥ��" & _
        " From ����Ԥ����¼ A, ���㷽ʽ B" & _
        " Where a.���㷽ʽ = b.���� And a.����id = [1] And b.���� In (3, 4) And a.�����id Is Null" & _
        " Group By ���㷽ʽ"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    Do While Not rsData.EOF
        If blnDelCheck Then
            If bln����������� Then
                '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
                If gclsInsure.GetCapability(support�����������, lng����ID, intInsure, NVL(rsData!���㷽ʽ)) Then
                    str���㷽ʽ = str���㷽ʽ & "||" & NVL(rsData!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(NVL(rsData!��Ԥ��))
                End If
            Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                If NVL(rsData!���㷽ʽ) <> str�����ʻ� Then
                    str���㷽ʽ = str���㷽ʽ & "||" & NVL(rsData!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(NVL(rsData!��Ԥ��))
                End If
            End If
        Else
            str���㷽ʽ = str���㷽ʽ & "||" & NVL(rsData!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(NVL(rsData!��Ԥ��))
        End If
            
        rsData.MoveNext
    Loop
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 3)
    GetYBBalance = str���㷽ʽ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlInsureCheck(ByVal strԤ���� As String, ByVal strAdvance As String) As Boolean
    '��鵱ǰ��ҽ���Ƿ���Ҫ�϶�
    '���:
    '   strԤ����-���ս���
    '   strAdvance-ҽ�����صĽ���
    '˵����
    '   ��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
    Dim blnFind  As Boolean, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo ErrHandler
    If strAdvance = "" Or strԤ���� = strAdvance Then Exit Function
    
    zlInsureCheck = True
    
    varData = Split(strԤ����, "||")
    varData1 = Split(strAdvance, "||")
    If UBound(varData) <> UBound(varData1) Then Exit Function
    
    For i = 0 To UBound(varData)
        blnFind = False
        varTemp = Split(varData(i), "|")
        For j = 0 To UBound(varData1)
            varTemp1 = Split(varData1(j), "|")
            If varTemp(0) = varTemp1(0) Then
                blnFind = True
                If Val(varTemp(1)) <> Val(varTemp1(1)) Then Exit Function
            End If
        Next
        If Not blnFind Then Exit Function
    Next
    zlInsureCheck = False
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteInsureDel_JZ(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal intInsure As Integer, ByVal str�����ʻ����� As String, _
    ByRef cllBillPro As Collection, ByRef objBalanceInfor As clsBalanceInfo) As Boolean
    '����:ִ�н���ҽ���˷��ò���
    '���:
    '   lng����ID - ԭ����ID
    Dim strSql As String, blnTransMedicare As Boolean
    Dim strAdvance As String, strSavedAdvance As String
    Dim bln����������� As Boolean
    Dim blnTrans As Boolean, cllPro As Collection
    Dim i As Integer
    
    On Error GoTo errHandle
    If intInsure = 0 Then ExecuteInsureDel_JZ = True: Exit Function
    
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    blnTrans = True
    bln����������� = gclsInsure.GetCapability(support�����������, lng����ID, intInsure)
    strSavedAdvance = GetYBBalance(lng����ID, lng����ID, True, True, intInsure, bln�����������, str�����ʻ�����)
    
    'Zl_���˽�������_Modify(
    strSql = "Zl_���˽�������_Modify("
    '  ��������_In      Number,
    strSql = strSql & "" & 3 & ","
    '  ����id_In        ������ü�¼.����id%Type,
    strSql = strSql & "" & lng����ID & ","
    '  ����id_In        ����Ԥ����¼.����id%Type,
    strSql = strSql & "" & objBalanceInfor.����ID & ","
    '  ���㷽ʽ_In      Varchar2,
    strSql = strSql & "'" & strSavedAdvance & "')"
    zlAddArray cllPro, strSql
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
          
    If bln����������� Then
        strAdvance = objBalanceInfor.����ID & "|0"
        If Not gclsInsure.ClinicDelSwap(lng����ID, True, intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            MsgBox "ҽ������ʧ�ܣ��޷����������������תסԺ������", vbInformation, gstrSysName
            Exit Function
        End If
        blnTransMedicare = True
    
        '���������Ƿ���ҪУ��
        If zlInsureCheck(strSavedAdvance, strAdvance) Then
            'Zl_���˽�������_Modify(
            strSql = "Zl_���˽�������_Modify("
            '  ��������_In      Number,
            strSql = strSql & "" & 3 & ","
            '  ����id_In        ������ü�¼.����id%Type,
            strSql = strSql & "" & lng����ID & ","
            '  ����id_In        ����Ԥ����¼.����id%Type,
            strSql = strSql & "" & objBalanceInfor.����ID & ","
            '  ���㷽ʽ_In      Varchar2,
            strSql = strSql & "'" & strAdvance & "')"
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        End If
    End If
    
    gcnOracle.CommitTrans: blnTrans = False
    objBalanceInfor.�Ƿ񱣴���ʵ� = True
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
    
    Set cllBillPro = New Collection
    ExecuteInsureDel_JZ = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intInsure)
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThirdReturnMoneySwap_JZ(ByVal objPati As clsPatiInfo, ByRef objBalanceInfor As clsBalanceInfo, _
    ByRef cllBillPro As Collection) As Boolean
    '����:ִ�������������˿�
    '���:objPati-��ǰ����Ĳ�����Ϣ
    '     objBalanceInfor-��ǰ�Ľ�����Ϣ
    '����:
    '����:ִ�гɹ�����true,���򷵻�False
    Dim strSql As String, rsTemp As ADODB.Recordset, rsBalance As ADODB.Recordset
    Dim i As Integer, lng�����ID As Long, lngԭ����ID As Long, lng��������ID As Long
    Dim objThirdDelItems As clsBalanceItems, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim objItems As clsBalanceItems, blnChangeMoney As Boolean
    Dim blnFinded As Boolean, blnSaveed As Boolean
    Dim cllPro As Collection, blnTrans As Boolean
    
    On Error GoTo errHandle
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    '������ִ��
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    Set cllPro = New Collection
    
    strSql = _
        " Select �����id, ���㷽ʽ, ��Ԥ�� As �����ܶ�, ��Ԥ��, ������ˮ��, ����˵��," & _
        "        ����, ��������id, �������, ժҪ, �տ�ʱ��" & _
        " From ����Ԥ����¼ A" & _
        " Where ��¼���� = 12 And a.����id = [1] And a.�����ID Is Not Null And a.У�Ա�־ = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, objBalanceInfor.����ID)
    '������������ֱ���˳�
    If rsTemp.RecordCount = 0 Then
        gcnOracle.RollbackTrans
        ExecuteThirdReturnMoneySwap_JZ = True: Exit Function
    End If
    
    strSql = _
        " Select Distinct a.����id, Nvl(a.�����id,0) as �����id,a.������ˮ��,Nvl(a.��������id,0) as ��������id " & _
        " From ����Ԥ����¼ A, " & _
        "  (Select a.ID" & _
        "   From ���˽��ʼ�¼ A, ���˽��ʼ�¼ B" & _
        "   Where a.No = b.No And a.��¼״̬ In (1, 3) And b.Id = [1]) B" & _
        " Where a.����id = b.id And Mod(a.��¼����,10)<>1"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, objBalanceInfor.����ID)
    
    Set objThirdDelItems = New clsBalanceItems
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lng�����ID = Val(NVL(rsTemp!�����ID))
            lng��������ID = Val(NVL(rsTemp!��������ID))
            
            lngԭ����ID = 0
            rsBalance.Filter = "�����ID=" & lng�����ID & " and ��������ID=" & lng��������ID
            If Not rsBalance.EOF Then lngԭ����ID = Val(NVL(rsBalance!����ID))
            If lngԭ����ID = 0 Then
                rsBalance.Filter = "�����ID=" & lng�����ID & " and ������ˮ��='" & NVL(!������ˮ��) & "'"
                If Not rsBalance.EOF Then lngԭ����ID = Val(NVL(rsBalance!����ID))
                If lngԭ����ID = 0 Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    MsgBox NVL(rsTemp!���㷽ʽ) & "δ�ҵ�ԭʼ�����¼ ������!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            objBalanceInfor.����ID = lngԭ����ID
            
            Set objItem = New clsBalanceItem
            With objItem
                Set .objCard = mobjThirdSwap.zlGetCardFromCardType(lng�����ID, False, NVL(rsTemp!���㷽ʽ))
                .����ID = objBalanceInfor.����ID
                .����IDs = lngԭ����ID
                .����ID = lngԭ����ID
                .��������ID = lng��������ID
                .������ˮ�� = NVL(rsTemp!������ˮ��)
                .����˵�� = NVL(rsTemp!����˵��)
                .���㷽ʽ = NVL(rsTemp!���㷽ʽ)
                .������� = NVL(rsTemp!�������)
                .����ժҪ = NVL(rsTemp!ժҪ)
                .������ = Val(NVL(rsTemp!��Ԥ��))
                .�������� = 3  '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                .�������� = .objCard.��������
                .����ʱ�� = Format(rsTemp!�տ�ʱ��, "yyyy-mm-dd HH:MM:SS")
                .���� = NVL(rsTemp!����)
                .�����ID = lng�����ID
                .ʣ���� = Val(NVL(rsTemp!��Ԥ��))
                .δ�˽�� = Val(NVL(rsTemp!��Ԥ��))
                .ԭʼ��� = Val(NVL(rsTemp!��Ԥ��))
            End With
            
            blnFinded = False
            For i = 1 To objThirdDelItems.Count
                Set objItemTemp = objThirdDelItems(i)
                If objItemTemp.�����ID = objItem.�����ID And objItemTemp.��������ID = objItem.��������ID Then
                    Set objItems = objItemTemp.objTag
                    If objItems Is Nothing Then Set objItems = New clsBalanceItems
                    objItems.AddItem objItem
                    objItems.������ = objItems.������ + objItem.������
                    Set objThirdDelItems(i).objTag = objItems
                    objThirdDelItems.������ = objThirdDelItems.������ + objItem.������
                    blnFinded = True: Exit For
                End If
            Next
            If Not blnFinded Then
                Set objItems = objItem.objTag
                If objItems Is Nothing Then Set objItems = New clsBalanceItems
                Set objItemTemp = objItem.zlCopyNewItemFromBalanceItem(objItem)
                Call objItems.AddItem(objItemTemp)
                objItems.������ = objItems.������ + objItem.������
                Set objItem.objTag = objItems
                objThirdDelItems.AddItem objItem
                objThirdDelItems.������ = objThirdDelItems.������ + objItem.������
            End If
            
            .MoveNext
        Loop
    End With
    
    Set rsBalance = Nothing: Set rsTemp = Nothing
   'ִ�������˿�
    For Each objItem In objThirdDelItems
        blnSaveed = False
        'byt��������-0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
        If mobjThirdSwap.zlThird_ReturnMoney_IsValied(objPati, objItem.objCard, 2, objBalanceInfor, objItem.objTag, objItems, False) = False Then
            If blnTrans Then gcnOracle.RollbackTrans
            If objBalanceInfor.�Ƿ񱣴���ʵ� Then
                 Call MsgBox(objItem.objCard.���� & "�˿�ʧ�ܣ����ڲ��˽��ʴ����н����쳣���ˣ�", vbInformation + vbOKOnly, gstrSysName)
            End If
            Exit Function
        End If
        If mobjThirdSwap.zlThird_ReturnMoney(objPati, objItem.objCard, objBalanceInfor, objItems, cllPro, False, objItems, blnSaveed, False, blnChangeMoney, False, blnTrans) = False Then
            If blnSaveed Or objBalanceInfor.�Ƿ񱣴���ʵ� Then
                objBalanceInfor.�Ƿ񱣴���ʵ� = True
                Call MsgBox(objItem.objCard.���� & "�˿�ʧ�ܣ����ڲ��˽��ʴ����н����쳣���ˣ�", vbInformation + vbOKOnly, gstrSysName)
            Else
                Call MsgBox(objItem.objCard.���� & "�˿�ʧ�ܣ������������תסԺʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName)
            End If
            Exit Function
        End If
        If blnSaveed And Not objBalanceInfor.�Ƿ񱣴���ʵ� Then objBalanceInfor.�Ƿ񱣴���ʵ� = True
    Next
    
    If blnTrans Then gcnOracle.CommitTrans
    objBalanceInfor.�Ƿ񱣴���ʵ� = True
    Set cllBillPro = New Collection
    ExecuteThirdReturnMoneySwap_JZ = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DelBalaceMz(ByVal objPati As clsPatiInfo, cllBillPro As Collection, _
    ByVal lng��ҳID As Long, ByVal lng��Ժ����ID As Long, ByRef objBalanceInfor As clsBalanceInfo) As Boolean
    '����:���˵������ͽ�������
    Dim strSql As String, rsData As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim intInsure As Integer, rsOneCard As ADODB.Recordset
    Dim lng����ID As Long, strNO As String, lng����ID As Long
    Dim blnDataSaved As Boolean, strBalanceIDs As String, strBalanceNos As String
    
    On Error GoTo ErrHandler
    strSql = _
        " Select /*+cardinality(j,10)*/ Distinct b.Id As ����ID, b.No, c.����, b.����ID" & _
        " From ������ü�¼ A, ���˽��ʼ�¼ B, ���ս����¼ C" & _
        " Where a.����id = b.Id And a.��¼���� In (2, 12) And a.No = [1] And b.��¼״̬ = 1" & _
        "       And b.ID=c.��¼id(+) And c.����(+) = 1 And c.�����id(+) Is Null" & _
        " Order By No"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, objBalanceInfor.���ʵ��ݺ�)
    If rsData.EOF Then
        'δ���ˣ�����ת�����
        blnTrans = True
        zlExecuteProcedureArrAy cllBillPro, Me.Caption
        blnTrans = False
        
        objBalanceInfor.�Ƿ񱣴���ʵ� = True
        Set cllBillPro = New Collection
        DelBalaceMz = True
        Exit Function
    End If
    
    Do While Not rsData.EOF
        strBalanceIDs = strBalanceIDs & "," & NVL(rsData!����ID)
        strBalanceNos = strBalanceNos & "," & NVL(rsData!NO)
        rsData.MoveNext
    Loop
    If strBalanceIDs <> "" Then
        '����Ƿ����һ��ͨ����
        Set rsOneCard = zlGetOneCard(Mid(strBalanceIDs, 2))
        If rsOneCard.RecordCount > 0 Then
            MsgBox "�ڽ��ʵ���" & Mid(strBalanceNos, 2) & "�д���һ��ͨ���㣬�ݲ�֧������תסԺ����,����!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If rsData.RecordCount > 0 Then rsData.MoveFirst
    Do While Not rsData.EOF
        With objBalanceInfor
            .�������� = 1  '��������:1-�������;2-סԺ����;3-�������תסԺ
            .����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        End With
        
        lng����ID = Val(NVL(rsData!����ID))
        strNO = NVL(rsData!NO)
        lng����ID = Val(NVL(rsData!����ID))
        intInsure = Val(NVL(rsData!����))
        
        If zlBillPrint_EraseBill("", lng����ID) = False Then Exit Function
        
        'Zl_���˽��ʼ�¼_Cancel
        strSql = "Zl_���˽��ʼ�¼_Cancel("
        '  No_In         ���˽��ʼ�¼.No%Type,
        strSql = strSql & "'" & strNO & "',"
        '  ����id_In     ���˽��ʼ�¼.Id%Type,
        strSql = strSql & "'" & objBalanceInfor.����ID & "',"
        '  ����Ա���_In ���˽��ʼ�¼.����Ա���%Type,
        strSql = strSql & "'" & UserInfo.��� & "',"
        '  ����Ա����_In ���˽��ʼ�¼.����Ա����%Type,
        strSql = strSql & "'" & UserInfo.���� & "',"
        '  ����ʱ��_In   ���˽��ʼ�¼.�շ�ʱ��%Type := Null
        strSql = strSql & "" & "To_Date('" & objBalanceInfor.����ʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ")"
        zlAddArray cllBillPro, strSql
        
        'Zl_����תסԺ_��������
        strSql = "Zl_����תסԺ_��������("
        '  No_In       ���˽��ʼ�¼.No%Type,
        strSql = strSql & "'" & strNO & "',"
        '  ����id_In   ���˽��ʼ�¼.Id%Type,
        strSql = strSql & "'" & objBalanceInfor.����ID & "',"
        '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
        strSql = strSql & "" & ZVal(lng��ҳID) & ","
        '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
        '  �������_In Number:=0 --0-��ʼ��������;1-��ɽ�������
        strSql = strSql & "" & 0 & ")"
        zlAddArray cllBillPro, strSql
        
        'ҽ���˿�
        If ExecuteInsureDel_JZ(lng����ID, lng����ID, intInsure, mstr�����ʻ�, cllBillPro, objBalanceInfor) = False Then Exit Function
        
        'һ��ͨ�˿�
        If ExecuteThirdReturnMoneySwap_JZ(objPati, objBalanceInfor, cllBillPro) = False Then Exit Function
        
        '��ɽ�������
        'Zl_����תסԺ_��������
        strSql = "Zl_����תסԺ_��������("
        '  No_In       ���˽��ʼ�¼.No%Type,
        strSql = strSql & "'" & NVL(rsData!NO) & "',"
        '  ����id_In   ���˽��ʼ�¼.Id%Type,
        strSql = strSql & "'" & objBalanceInfor.����ID & "',"
        '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
        strSql = strSql & "" & ZVal(lng��ҳID) & ","
        '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
        '  �������_In Number:=0 --0-��ʼ��������;1-��ɽ�������
        strSql = strSql & "" & 1 & ")"
        zlAddArray cllBillPro, strSql
        
        '���һ�ν������Ͼ��ύ
        blnTrans = True
        zlExecuteProcedureArrAy cllBillPro, Me.Caption
        blnTrans = False
        
        objBalanceInfor.�Ƿ񱣴���ʵ� = True
        Set cllBillPro = New Collection
        
        rsData.MoveNext
    Loop
    DelBalaceMz = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowBalanceWindows(ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���ʴ���
    ' ���:strDelDate-��������(��ҪӦ�����ٴν���ʱ��Ԥ����)
    '����:���˺�
    '����:2011-03-29 17:38:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objInExse As Object
    Dim lng����ID As Long
    
   '4.�������ʲ���
    If objInExse Is Nothing Then
        Err = 0: On Error Resume Next
        Set objInExse = CreateObject("zl9InExse.clsFeeQuery")
        If Err <> 0 Then
            MsgBox "ע��:" & "�ڴ���סԺ���ò���ʱ����,���ܸò���δ����ע��,����ʧ��,��ע�����½���!", vbInformation + vbOKOnly, gstrSysName
            ShowBalanceWindows = True
            Exit Function
        End If
    End If
    
    On Error GoTo errHandle
    If mlngPatient <> 0 Then
        lng����ID = mlngPatient
    ElseIf Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng����ID = Val(NVL(mrsInfo!����ID))
    End If
    
    'zlPatiBalance(ByVal frmMain As Object, _
    '    ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, strDBUser As String, _
    '    ByVal lng����ID As Long, ByVal lng��ҳID As   long ) as boolean
    If objInExse.zlPatiBalance(Me, gcnOracle, glngSys, gstrDBUser, lng����ID, 0, strDelDate) = False Then
        '���ý���
    End If
    ShowBalanceWindows = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowBills(ByVal lngPatient As Long, ByVal DatBegin As Date, ByVal DatEnd As Date, _
    Optional ByVal blnFilter As Boolean)
'����:��ȡ����ʾ����ָ�������ڵ�������õ���
    Dim DatTmp As Date, strSql As String
    Dim strWhere As String
    Dim strFilter As String
    Dim strIDs As String, lngPre��������ID As Long
    Dim strVerifyWhere As String
    Dim strErrWhere As String, strBalanceErrWhere As String
    
    On Error GoTo errH
    If mrsFeeList Is Nothing Or blnFilter = False Then
        zlCommFun.ShowFlash "���ڶ�ȡ�շѵ���,���Ժ� ..."
        If DatBegin > DatEnd Then
            DatTmp = DatEnd
            DatEnd = DatBegin
            DatBegin = DatTmp
        End If
        
        '�ų��շ��쳣�ĵ���
        strErrWhere = _
            " And Not Exists (Select 1" & _
            "     From ������ü�¼ J1, ������ü�¼ J2, ������ü�¼ J3" & _
            "     Where a.No = J1.No And a.��� = J1.��� And J1.��¼���� = 1 And J1.��¼״̬ In (1,3)" & _
            "           And J1.����id = J2.����id And J1.��� =  J2.���" & _
            "           And J2.No = J3.No And J2.��� =  J3.��� And Mod(J3.��¼����,10) = 1 And Nvl(J3.����״̬,0)=1)" & vbCrLf
        strErrWhere = strErrWhere & _
            " And Not Exists(Select 1 From ���ò����¼ where �շѽ���ID=a.����ID And ��¼����=1 And Nvl(����״̬,0)=1) " & vbCrLf
        
        '�ų������쳣�ĵ���
        strBalanceErrWhere = _
            " And Not Exists(Select 1" & _
            "     From ������ü�¼ J1, ���˽��ʼ�¼ J2" & _
            "     Where J1.No = a.No And J1.��¼���� In (2,12) And J1.����id = J2.Id And Nvl(J2.����״̬,0)=1)"
        
        If mbln����תסԺ����� Then
           strWhere = " And A.����id = [1] "
        Else
            If DatEnd - DatBegin < 4 Then   '36170
                If IDKindTime.IDKind = 1 Then
                    strWhere = " And A.����id+0 = [1] And A.����ʱ�� Between [2] And [3]  "
                Else
                    strWhere = " And A.����id+0 = [1] And A.�Ǽ�ʱ�� Between [2] And [3]  "
                End If
            Else
                If IDKindTime.IDKind = 1 Then
                    strWhere = " And A.����id = [1] And A.����ʱ��+0 Between [2] And [3]  "
                Else
                    strWhere = " And A.����id = [1] And A.�Ǽ�ʱ��+0 Between [2] And [3]  "
                End If
            End If
        End If
        
        If mbln����תסԺ����� Then
            strVerifyWhere = _
            " And Exists (Select 1 From ������ü�¼ M,������˼�¼ J " & _
            "             Where M.ID=J.����ID And M.����ID = [1] and M.NO=A.NO And Mod(M.��¼����,10)=Mod(A.��¼����,10)  " & _
            "                   And J.������� is Not NULL and  nvl(J.��¼״̬,0)=0 and J.����=1) " & vbNewLine
        Else
            strVerifyWhere = _
            " And Not Exists (Select 1 From ������ü�¼ M,������˼�¼ J " & _
            "                 Where M.ID=J.����ID And M.����ID = [1] and M.NO=A.NO And Mod(M.��¼����,10)=Mod(A.��¼����,10) " & _
            "                       And J.������� is Not NULL and  nvl(J.��¼״̬,0) > 0 and J.����=1)"
        End If
        
        strSql = strSql & _
            " Select x.ѡ��, x.���, x.����, Max(Decode(Nvl(z.����, 0),0,'','��')) As ҽ��,Max(z.����) As һ��ͨҽ��," & _
            "       x.No As ���ݺ�, x.Ʊ�ݺ�," & vbNewLine & _
            "       x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Max(y.����id) As ����id," & vbNewLine & _
            "       Max(Decode(z.�����ID,NULL,Nvl(z.����,0),0)) As ����" & vbNewLine & _
            " From ( Select  '��' As ѡ��, '��ת��' As ���, '�շѵ�' As ����, a.No," & vbNewLine & _
            "               a.ʵ��Ʊ�� As Ʊ�ݺ�, a.������, a.��������ID, LTrim(To_Char(Sum(a.Ӧ�ս��), '9999999990.0000')) As Ӧ�ս��," & vbNewLine & _
            "               LTrim(To_Char(Sum(a.ʵ�ս��), '9999999990.0000')) As ʵ�ս��, To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & vbNewLine & _
            "        From ������ü�¼ A" & vbNewLine & _
            "        Where Mod(a.��¼����, 10) = 1 And nvl(a.����״̬,0)<>1 And a.��¼״̬ <> 0 " & strWhere & " " & strVerifyWhere & vbCrLf & strErrWhere & _
            "              And Exists (Select 1 From ������ü�¼ K" & _
            "                          Where k.No = a.No And k.����id = [1] And Mod(k.��¼����, 10) = Mod(a.��¼����, 10)" & _
            "                                And Nvl(k.���ӱ�־, 0) <> 9" & _
            "                          Group By k.��� Having Sum(k.ʵ�ս��) <> 0)" & vbNewLine & _
            "        Group By a.No, a.ʵ��Ʊ��, a.������, a.��������ID, a.����ʱ�� " & _
            "      ) X, ������ü�¼ Y," & vbNewLine & _
            "      ( Select Distinct a.��¼id, a.����,a.�����ID,b.����" & vbNewLine & _
            "        From ���ս����¼ A,ҽ�ƿ���� B" & vbNewLine & _
            "        Where a.���� = 1 And a.����id = [1] And a.�����ID=b.ID(+)) Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.��¼����, 10) = 1 And y.��¼״̬ In (1, 3) And y.����ID = [1]" & _
            "        And y.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = x.No And Mod(��¼����, 10) = 1 And ����ID = [1] And ��¼״̬ In (1, 3)) And y.����id = z.��¼id(+)" & _
            " Group By x.ѡ��, x.���, x.����, x.No, x.Ʊ�ݺ�, x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ�� "
 
        strSql = strSql & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select x.ѡ��, x.���, x.����, Max(Decode(Nvl(z.����, 0),0,'','��')) As ҽ��,Max(z.����) As һ��ͨҽ��," & _
            "       x.No As ���ݺ�, x.Ʊ�ݺ�," & vbNewLine & _
            "       x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Max(y.����id) As ����id," & vbNewLine & _
            "       Max(Decode(z.�����ID,NULL,Nvl(z.����,0),0)) As ����" & vbNewLine & _
            " From ( " & _
            "       Select " & vbNewLine & _
            "           '' As ѡ��, '����ת��' As ���, '�շѵ�' As ����, a.No," & vbNewLine & _
            "           a.ʵ��Ʊ�� As Ʊ�ݺ�, a.������, a.��������ID, LTrim(To_Char(Sum(a.Ӧ�ս��), '9999999990.0000')) As Ӧ�ս��," & vbNewLine & _
            "           LTrim(To_Char(Sum(a.ʵ�ս��), '9999999990.0000')) As ʵ�ս��, To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & vbNewLine & _
            "       From ������ü�¼ A" & vbNewLine & _
            "       Where Mod(a.��¼����, 10) = 1 And nvl(a.����״̬,0)<>1 And a.��¼״̬ = 3 " & strWhere & " And Nvl(a.���ӱ�־, 0) <> 9 " & vbCrLf & strErrWhere & _
            "           And Not Exists (Select 1 From ������ü�¼ K  Where k.No = a.No And k.����id = [1] And Mod(k.��¼����, 10) = Mod(a.��¼����, 10) And Nvl(k.���ӱ�־, 0) <> 9 Group By k.���  Having Sum(k.ʵ�ս��) <> 0)" & vbNewLine & _
            "       Group By a.No, a.ʵ��Ʊ��, a.������, a.��������ID, a.����ʱ�� " & _
            "       ) X, ������ü�¼ Y," & vbNewLine & _
            "       (Select Distinct a.��¼id, a.����,a.�����ID, b.����" & vbNewLine & _
            "        From ���ս����¼ A,ҽ�ƿ���� B" & vbNewLine & _
            "        Where a.���� = 1 And a.����id = [1] And a.�����ID=b.ID(+)) Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.��¼����, 10) = 1 And y.��¼״̬ In (1, 3) And y.����ID = [1]" & _
            "       And y.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = x.No And Mod(��¼����, 10) = 1 And ����ID = [1] And ��¼״̬ In (1, 3)) And y.����id = z.��¼id(+)" & _
            " Group By x.ѡ��, x.���, x.����, x.No, x.Ʊ�ݺ�, x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��"

            
        strSql = strSql & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select x.ѡ��, x.���, x.����, Max(Decode(Nvl(z.����, 0),0,'','��')) As ҽ��,Max(z.����) As һ��ͨҽ��," & _
            "       x.No As ���ݺ�, x.Ʊ�ݺ�," & vbNewLine & _
            "       x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Max(y.����id) As ����id," & vbNewLine & _
            "       Max(Decode(z.�����ID,NULL,Nvl(z.����,0),0)) As ����" & vbNewLine & _
            "From (Select " & vbNewLine & _
            "        '' As ѡ��, '����ת��' As ���, '�շѵ�' As ����, a.No," & vbNewLine & _
            "        a.ʵ��Ʊ�� As Ʊ�ݺ�, a.������, a.��������ID, LTrim(To_Char(Sum(a.Ӧ�ս��), '9999999990.0000')) As Ӧ�ս��," & vbNewLine & _
            "        LTrim(To_Char(Sum(a.ʵ�ս��), '9999999990.0000')) As ʵ�ս��, To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & vbNewLine & _
            "       From ������ü�¼ A" & vbNewLine & _
            "       Where Mod(a.��¼����, 10) = 1 And nvl(a.����״̬,0)<>1 And a.��¼״̬ <> 0 " & strWhere & " " & vbCrLf & strErrWhere & _
            "           And Exists (Select 1 From ������ü�¼ M,������˼�¼ J Where M.ID=J.����ID And M.����ID = [1] and M.NO=A.NO And Mod(M.��¼����,10)=Mod(A.��¼����,10) And J.������� is Not NULL and  nvl(J.��¼״̬,0) = 1 and J.����=1)" & _
            "           And Exists��(Select 1�� From ������ü�¼ K��Where k.No = a.No And k.����id = [1] And Mod(k.��¼����, 10) = Mod(a.��¼����, 10) And Nvl(k.���ӱ�־, 0) <> 9��Group By k.��š�Having Sum(k.ʵ�ս��) <> 0)" & vbNewLine & _
            "       Group By a.No, a.ʵ��Ʊ��, a.������, a.��������ID, a.����ʱ��) X, ������ü�¼ Y," & vbNewLine & _
            "     (  Select Distinct a.��¼id, a.����,a.�����ID, b.����" & vbNewLine & _
            "        From ���ս����¼ A,ҽ�ƿ���� B" & vbNewLine & _
            "        Where a.���� = 1 And a.����id = [1] And a.�����ID=b.ID(+)) Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.��¼����, 10) = 1 And y.��¼״̬ In (1, 3) And y.����ID = [1]" & _
            " And y.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = x.No And Mod(��¼����, 10) = 1 And ����ID = [1] And ��¼״̬ In (1, 3)) And y.����id = z.��¼id(+)" & _
            " Group By x.ѡ��, x.���, x.����, x.No, x.Ʊ�ݺ�, x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��"
     
        strSql = strSql & " UNION ALL " & _
                " Select    '��' as ѡ��,'��ת��' as ���,'���ʵ�' as ����,'' as ҽ��,'' As һ��ͨҽ��," & _
                "       A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������, a.��������ID," & vbNewLine & _
                "       LTrim(To_Char(Sum(A.Ӧ�ս��), '999999999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
                "       LTrim(To_Char(Sum(A.ʵ�ս��), '999999999" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
                "       To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, 0 as ����ID,0 as ����" & vbNewLine & _
                " From ������ü�¼ A" & vbNewLine & _
                " Where A.��¼���� =2 And A.��¼״̬ <> 0 " & strWhere & strBalanceErrWhere & vbNewLine & _
                "           And Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��¼����=A.��¼���� And Nvl(k.���ӱ�־, 0) <> 9 Group By K.��� Having Sum(K.����) <> 0) " & vbNewLine & _
                            IIf(mbln����תסԺ�����, "           And Exists(Select 1 From ������ü�¼ M,������˼�¼ J where M.ID=J.����ID and M.NO=A.NO And M.��¼����=A.��¼���� And J.������� is Not NULL and  nvl(J.��¼״̬,0)=0 and J.����=1) " & vbNewLine, " And Not Exists(Select 1 From ������ü�¼ M,������˼�¼ J where M.ID=J.����ID and M.NO=A.NO And M.��¼����=A.��¼���� And J.������� is Not NULL and  nvl(J.��¼״̬,0) > 0 and J.����=1) ") & _
                "Group By A.NO, A.ʵ��Ʊ��, A.������, a.��������ID, A.����ʱ�� "
             
        strSql = strSql & " UNION ALL " & _
            " Select C.ѡ��,C.���,C.����,C.ҽ��,c.һ��ͨҽ��,C.���ݺ�, C.Ʊ�ݺ�, C.������, c.��������ID," & vbNewLine & _
            "       LTrim(To_Char(Sum(D.Ӧ�ս��), '999999999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(D.ʵ�ս��), '999999999" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
            "       C.����ʱ��, C.����ID, C.����" & vbNewLine & _
            " From " & _
            " (Select    '' as ѡ��,'����ת��' as ���,'���ʵ�' as ����,'' as ҽ��,'' As һ��ͨҽ��," & _
            "       A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������, a.��������ID," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.Ӧ�ս��), '999999999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.ʵ�ս��), '999999999" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
            "       To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��,0 as ����ID,0 as ����" & vbNewLine & _
            " From ������ü�¼  A" & vbNewLine & _
            " Where A.��¼���� = 2 And A.��¼״̬ In (2,3)" & strWhere & strBalanceErrWhere & vbNewLine & _
            "       And Not Exists (Select 1 From ������ü�¼ Where NO=A.NO And ��¼״̬=1 And ��¼����=2) " & vbNewLine & _
            "       And Not Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��¼����=A.��¼���� And Nvl(k.���ӱ�־, 0) <> 9 Group By K.��� Having Sum(K.ʵ�ս��) <> 0) " & vbNewLine & _
            " Group By A.NO, A.ʵ��Ʊ��, A.������, a.��������ID, A.����ʱ�� Having Sum(A.ʵ�ս��)=0) C,������ü�¼ D Where C.���ݺ�=D.NO And D.��¼����=2 And D.��¼״̬=3" & vbNewLine & _
            " Group By C.ѡ��,C.���,C.����,C.ҽ��,C.���ݺ�, C.Ʊ�ݺ�, C.������, c.��������ID,C.����ʱ��, C.����ID, C.���� "
            
        strSql = strSql & " UNION ALL " & _
            " Select    '' as ѡ��,'����ת��' as ���,'���ʵ�' as ����,'' as ҽ��,'' As һ��ͨҽ��, " & _
            "       A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������, a.��������ID," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.Ӧ�ս��), '999999999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.ʵ�ս��), '999999999" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
            "       To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, 0 as ����ID,0 as ����" & vbNewLine & _
            " From ������ü�¼ A" & vbNewLine & _
            " Where A.��¼���� = 2 And A.��¼״̬ <> 0 " & strWhere & strBalanceErrWhere & vbNewLine & _
            "       And Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��¼����=A.��¼���� And Nvl(k.���ӱ�־, 0) <> 9 Group By K.��� Having Sum(K.����) <> 0) " & vbNewLine & _
            " And  Exists (Select 1 From ������ü�¼ M,������˼�¼ J where M.ID=J.����ID and M.NO=A.NO And M.��¼����=A.��¼���� And J.������� is Not NULL and  nvl(J.��¼״̬,0) = 1 and J.����=1) " & _
            "Group By A.NO, A.ʵ��Ʊ��, A.������, a.��������ID, A.����ʱ�� "
        
        strSql = _
            " Select ѡ��, ���, ����, ҽ��, һ��ͨҽ��, ���ݺ�, Ʊ�ݺ�, ������, b.���� As ��������, Ӧ�ս��, ʵ�ս��, ����ʱ��," & vbNewLine & _
            "        ����id, ����, ��������id As ��������ID, b.���� As �������ұ���" & _
            " From (" & strSql & ") A,���ű� B" & vbNewLine & _
            " Where a.��������ID = b.ID" & vbNewLine & _
            " Order By ����,���, Ʊ�ݺ� Desc, ���ݺ� Desc"
        'ע��:����ҽ��Ҫ������һ�ſ�ʼ��,��������ܹؼ�
        Set mrsFeeList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatient, DatBegin, DatEnd)
    
        '���ؿ�ѡ����
        mblnNotClick = True
        If cbo��������.ListIndex <> -1 Then lngPre��������ID = Val(cbo��������.ItemData(cbo��������.ListIndex))
        cbo��������.Clear
        cbo��������.AddItem "���п���"
        Do While Not mrsFeeList.EOF
            If InStr("," & strIDs & ",", "," & NVL(mrsFeeList!��������ID) & ",") = 0 Then
                strIDs = strIDs & "," & NVL(mrsFeeList!��������ID)
                
                cbo��������.AddItem IIf(zlIsShowDeptCode, NVL(mrsFeeList!�������ұ���) & "-", "") & NVL(mrsFeeList!��������)
                cbo��������.ItemData(cbo��������.NewIndex) = NVL(mrsFeeList!��������ID)
                If Val(NVL(mrsFeeList!��������ID)) = lngPre��������ID Then cbo��������.ListIndex = cbo��������.NewIndex
            End If
            mrsFeeList.MoveNext
        Loop
        cbo.SetListWidthAuto cbo��������
        If cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
        mblnNotClick = False
        
        zlCommFun.StopFlash
    End If
    
    Screen.MousePointer = vbHourglass
    strFilter = ""
    If chkShow.Value = vbChecked Then strFilter = strFilter & " And  ���='��ת��'"
    If Val(cbo��������.ItemData(cbo��������.ListIndex)) <> 0 Then
        strFilter = strFilter & " And ��������ID=" & cbo��������.ItemData(cbo��������.ListIndex)
    End If
    mrsFeeList.Filter = Mid(strFilter, 5)
    
    mshList.Redraw = flexRDNone: mshList.Clear
    mshList.Rows = 2
    Set mshList.DataSource = mrsFeeList
    If mrsFeeList.EOF Then
        sta.Panels(2).Text = "û���ҵ�ָ��ʱ�䷶Χ���շѻ���ʵ���!"
        mshList.Rows = 2
    Else
        sta.Panels(2).Text = "�� " & mrsFeeList.RecordCount & " ���շѵ���"
    End If
    Call setHeader
    Call SetInsure
    Call SetBillColor
    mshList.Redraw = flexRDBuffered
    Call mshList_AfterRowColChange(0, 0, 1, 0)
    If mshList.Rows >= 2 Then mshList.Select 1, 0
    Call SetSumMoney
    Screen.MousePointer = vbDefault
    Exit Sub
errH:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetInsure()
    Dim intInsure As Integer, lngRow As Long
    Dim str���� As String
    
    With mshList
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("���")) = "��ת��" And .TextMatrix(lngRow, .ColIndex("ѡ��")) = "��" Then
                intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
                str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
                If intInsure > 0 And str���� = "�շѵ�" Then
                    If Not gclsInsure.GetCapability(support�����������, mlngPatient, intInsure) Then
                        .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
                    End If
                End If
            End If
        Next lngRow
    End With
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Function ExecuteThirdReturnMoneySwap(ByVal objPati As clsPatiInfo, ByRef objBalanceInfor As clsBalanceInfo, ByRef cllBillPro As Collection, Optional objSequareDelItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���������˿�
    '���:objPati-��ǰ����Ĳ�����Ϣ
    '     cllBillPro-��ǰִ�еĹ��̼�
    '     objBalanceInfor-��ǰ�Ľ�����Ϣ
    '����:objSequareDelItems_Out-���ѿ��˿���Ϣ��
    '����:ִ�гɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-08-17 15:01:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, rsBalance As ADODB.Recordset
    Dim i As Integer, lng�����ID As Long, lngԭ����ID As Long, bln���ѿ� As Boolean, lng��������ID As Long, lng���㿨��� As Long
    Dim objThirdDelItems As clsBalanceItems, objSequareDelItems As clsBalanceItems, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim objItems As clsBalanceItems, blnChangeMoney As Boolean
    Dim blnFinded As Boolean, blnSaveed As Boolean
    Dim cllPro As Collection, blnTrans As Boolean
    Dim rsTotal As ADODB.Recordset
    Dim str������Ϣ As String
    
    On Error GoTo errHandle
    '������ִ�к����������ݣ�����Ҫ��ִ��
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    
    strSql = _
    " Select '' as NO, �����ID,���㿨���,���㷽ʽ,��Ԥ�� as �����ܶ�,��Ԥ��,������ˮ��,����˵��,����,��������ID,�������,ժҪ,�տ�ʱ��" & vbNewLine & _
    " From ����Ԥ����¼ A" & vbNewLine & _
    " Where ��¼���� = 3 And ��¼״̬ = 2 and ���ӱ�־=-1  And ����id = [1] and У�Ա�־=1 " & vbNewLine & _
    "       and Not Exists(Select 1 From ҽ��������ϸ where ����ID=[1] And A.�����ID=�����ID  And a.��������ID=��������ID )" & vbNewLine & _
    " Union all " & vbNewLine & _
    " Select distinct b.NO,A.�����ID,A.���㿨���,b.���㷽ʽ,A.��Ԥ�� as �����ܶ�,b.��� as ��Ԥ��,nvl(b.������ˮ��,A.������ˮ��) as ������ˮ��,nvl(b.����˵��,A.����˵��) as ����˵��," & vbNewLine & _
    "        A.����,A.��������ID,A.�������,A.ժҪ,A.�տ�ʱ��" & vbNewLine & _
    " From ����Ԥ����¼ A ,ҽ��������ϸ B" & vbNewLine & _
    " Where A.��¼���� = 3 And A.��¼״̬ = 2 and A.���ӱ�־=-1  And A.����id = [1] and A.У�Ա�־=1 " & vbNewLine & _
    "       and A.����ID=B.����ID And A.�����ID=B.�����ID and A.��������ID=B.��������ID and a.���㷽ʽ=b.���㷽ʽ(+) " & vbNewLine & _
    " Order by �����ID,��������ID,NO,���㷽ʽ"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, objBalanceInfor.����ID)
    If rsTemp.RecordCount = 0 Then '�����������ѿ�����ֱ���˳�
        ExecuteThirdReturnMoneySwap = True: Exit Function
    End If
    
    Set rsTotal = New ADODB.Recordset
    With rsTotal
        .Fields.Append "�����ID", adInteger, , adFldIsNullable
        .Fields.Append "��������ID", adInteger, , adFldIsNullable
        .Fields.Append "���ݺ�", adVarChar, 20, adFldIsNullable
        .Fields.Append "���������", adVarChar, 100, adFldIsNullable
        .Fields.Append "�����ܶ�", adDouble, , adFldIsNullable
        .Fields.Append "��ϸ�ܶ�", adDouble, , adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    Set cllPro = New Collection
    
    strSql = " " & _
    "   Select distinct a.����id, nvl(a.�����id,0) as �����id,a.������ˮ��,nvl(a.���㿨���,0) as ���㿨���,nvl(a.��������id,0) as ��������id " & _
    "   From ����Ԥ����¼ A, " & _
    "        (Select Distinct ����id " & _
    "          From ������ü�¼ " & _
    "          Where NO In (Select Distinct NO From ������ü�¼ Where ����id = [1]) And Mod(��¼����, 10) = 1 And ��¼״̬ In (3, 1)) B " & _
    "   Where a.����id = b.����id and mod(a.��¼����,10)<>1"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, objBalanceInfor.����ID)
    
    Set objSequareDelItems = New clsBalanceItems
    Set objThirdDelItems = New clsBalanceItems
    
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lng�����ID = Val(NVL(rsTemp!�����ID))
            lng���㿨��� = Val(NVL(rsTemp!���㿨���))
            bln���ѿ� = lng���㿨��� <> 0
            lng��������ID = Val(NVL(rsTemp!��������ID))
            
            rsBalance.Filter = "�����ID=" & lng�����ID & " and ��������ID=" & lng��������ID & " and ���㿨���=" & lng���㿨���
            lngԭ����ID = 0
            If Not rsBalance.EOF Then lngԭ����ID = Val(NVL(rsBalance!����ID))
            If lngԭ����ID = 0 And Not bln���ѿ� Then
                rsBalance.Filter = "�����ID=" & lng�����ID & " and ������ˮ��='" & NVL(!������ˮ��) & "'"
                If Not rsBalance.EOF Then lngԭ����ID = Val(NVL(rsBalance!����ID))
                If lngԭ����ID = 0 Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    MsgBox NVL(rsTemp!���㷽ʽ) & "δ�ҵ�ԭʼ�����¼ ������!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            
            Set objItem = New clsBalanceItem
            With objItem
                Set .objCard = mobjThirdSwap.zlGetCardFromCardType(lng�����ID, bln���ѿ�, NVL(rsTemp!���㷽ʽ))
                .����ID = objBalanceInfor.����ID
                .����IDs = lngԭ����ID
                .����ID = lngԭ����ID
                .��������ID = lng��������ID
                .������ˮ�� = NVL(rsTemp!������ˮ��)
                .����˵�� = NVL(rsTemp!����˵��)
                .���㷽ʽ = NVL(rsTemp!���㷽ʽ)
                .������� = NVL(rsTemp!�������)
                .����ժҪ = NVL(rsTemp!ժҪ)
                .������ = Val(NVL(rsTemp!��Ԥ��))
                .�������� = IIf(bln���ѿ�, 5, 3)  '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                .�������� = .objCard.��������
                .����ʱ�� = Format(rsTemp!�տ�ʱ��, "yyyy-mm-dd HH:MM:SS")
                .���� = NVL(rsTemp!����)
                .�����ID = IIf(bln���ѿ�, lng���㿨���, lng�����ID)
                .ʣ���� = Val(NVL(rsTemp!��Ԥ��))
                .δ�˽�� = Val(NVL(rsTemp!��Ԥ��))
                .ԭʼ��� = Val(NVL(rsTemp!��Ԥ��))
                .���ѿ� = bln���ѿ�
                .���ݺ� = NVL(rsTemp!NO)
            End With
            If objItem.���ݺ� <> "" And Not objItem.���ѿ� And objItem.�����ID <> 0 Then
                rsTotal.Filter = "�����ID=" & objItem.�����ID & " and ��������ID=" & objItem.��������ID
                If rsTotal.EOF Then
                    rsTotal.AddNew
                    rsTotal!�����ID = objItem.�����ID
                    rsTotal!��������ID = objItem.��������ID
                    'rsTotal!���ݺ� = objItem.���ݺ�
                    rsTotal!��������� = IIf(objItem.objCard.���� = "", objItem.���㷽ʽ, objItem.objCard.����)
                End If
                If InStr(str������Ϣ & ",", "," & objItem.���㷽ʽ & ",") = 0 Then
                    str������Ϣ = str������Ϣ & "," & objItem.���㷽ʽ
                    rsTotal!�����ܶ� = Val(NVL(rsTotal!�����ܶ�)) + Val(NVL(rsTemp!�����ܶ�))
                End If
                rsTotal!��ϸ�ܶ� = RoundEx(Val(NVL(rsTotal!��ϸ�ܶ�)) + objItem.������, 6)
                rsTotal.Update
            End If
            
            If objItem.���ѿ� Then
                objSequareDelItems.AddItem objItem
                objSequareDelItems.������ = objSequareDelItems.������ + objItem.������
            Else
                blnFinded = False
                For i = 1 To objThirdDelItems.Count
                    Set objItemTemp = objThirdDelItems(i)
                    If objItemTemp.�����ID = objItem.�����ID And objItemTemp.��������ID = objItem.��������ID Then
                        Set objItems = objItemTemp.objTag
                        If objItems Is Nothing Then Set objItems = New clsBalanceItems
                        objItems.AddItem objItem
                        objItems.������ = objItems.������ + objItem.������
                        Set objThirdDelItems(i).objTag = objItems
                        objThirdDelItems.������ = objThirdDelItems.������ + objItem.������
                        blnFinded = True
                        Exit For
                    End If
                Next
                If Not blnFinded Then
                    Set objItems = objItem.objTag
                    If objItems Is Nothing Then Set objItems = New clsBalanceItems
                    Set objItemTemp = objItem.zlCopyNewItemFromBalanceItem(objItem)
                    Call objItems.AddItem(objItemTemp)
                    objItems.������ = objItems.������ + objItem.������
                    Set objItem.objTag = objItems
                    objThirdDelItems.AddItem objItem
                    objThirdDelItems.������ = objThirdDelItems.������ + objItem.������
                End If
            End If
            .MoveNext
        Loop
    End With
    
    Set rsBalance = Nothing: Set rsTemp = Nothing
    '���ҽ��������ϸ�뵱ǰ���˿��ܶ��Ƿ�һ�£���һ�£���ֹת��
    rsTotal.Filter = 0
    With rsTotal
         If .RecordCount <> 0 Then .MoveFirst
         Do While Not .EOF
            If RoundEx(Val(NVL(!�����ܶ�)), 6) <> RoundEx(Val(NVL(!��ϸ�ܶ�)), 6) Then
                If blnTrans Then gcnOracle.RollbackTrans
                MsgBox "���ݺ�Ϊ" & !���ݺ� & "���˿��ܶ���ҽ��������ϸ�е��˿��һ�£���ֹ�������תסԺ!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
             .MoveNext
         Loop
    End With
   
   'ִ�������˿�
    For Each objItem In objThirdDelItems
    
        blnSaveed = False
        'byt��������-0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
        If mobjThirdSwap.zlThird_ReturnMoney_IsValied(objPati, objItem.objCard, 2, objBalanceInfor, objItem.objTag, objItems, False) = False Then
            If blnTrans Then gcnOracle.RollbackTrans
            If objBalanceInfor.�Ƿ񱣴���ʵ� Then
                 Call MsgBox(objItem.objCard.���� & "�˿�ʧ�ܣ����������շѴ����н����쳣����", vbInformation + vbOKOnly, gstrSysName)
            End If
            Exit Function
        End If
        If mobjThirdSwap.zlThird_ReturnMoney(objPati, objItem.objCard, objBalanceInfor, objItems, cllPro, False, objItems, blnSaveed, False, blnChangeMoney, False, blnTrans) = False Then
            If blnSaveed Or objBalanceInfor.�Ƿ񱣴���ʵ� Then
                objBalanceInfor.�Ƿ񱣴���ʵ� = True
                Call MsgBox(objItem.objCard.���� & "�˿�ʧ�ܣ����������շѴ����н����쳣���ˣ�", vbInformation + vbOKOnly, gstrSysName)
            Else
                Call MsgBox(objItem.objCard.���� & "�˿�ʧ��,�����������תסԺʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName)
            End If
            Exit Function
        End If
        If blnSaveed And Not objBalanceInfor.�Ƿ񱣴���ʵ� Then objBalanceInfor.�Ƿ񱣴���ʵ� = True
    Next
    If objThirdDelItems.Count = 0 Then  '���ѿ��������ʱһ������
        If blnTrans Then gcnOracle.RollbackTrans
         ExecuteThirdReturnMoneySwap = True: Exit Function
    End If
    
    If blnTrans Then gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    ExecuteThirdReturnMoneySwap = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub setHeader()
    Dim strHead As String
    Dim i As Long
    With mshList
        strHead = "ѡ��,4,500|���,4,850|����,4,800|ҽ��,4,500|һ��ͨҽ��,1,550|���ݺ�,4,850|Ʊ�ݺ�,4,1100|������,1,800|��������,1,1200|" & _
            "Ӧ�ս��,7,850|ʵ�ս��,7,850|����ʱ��,4,1850|����ID,4,0|����,4,0|��������ID,4,0|�������ұ���,4,0"
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
             If .ColKey(i) Like "*ID" Or .ColKey(i) = "����" Or .ColKey(i) = "�������ұ���" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
             End If
             .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .RowHeight(0) = 320
        
        '�ϲ������С�ҽ����
        .MergeCellsFixed = flexMergeRestrictRows
        .MergeRow(0) = True
        .TextMatrix(0, .ColIndex("һ��ͨҽ��")) = .TextMatrix(0, .ColIndex("ҽ��"))
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore 1131, mshList, Me.Caption, "����תסԺ�б�", True
        
        .ColHidden(.ColIndex("һ��ͨҽ��")) = True
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("һ��ͨҽ��"))) <> "" Then
                .ColHidden(.ColIndex("һ��ͨҽ��")) = False: Exit For
            End If
        Next
        
        .Row = 1
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub
Private Sub SetBillColor()
    Dim i As Long
    
    With mshList
        For i = 1 To .Rows - 1
            .Row = i
            If .TextMatrix(i, .ColIndex("���")) = "����ת��" Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H8000000C
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = .ForeColor
            End If
        Next
    End With
End Sub

Private Sub cmdParaSet_Click()
    frmChargeTurnParSet.ShowSet Me, 1131, mstrPrivs
    mbln�������� = Val(zlDatabase.GetPara("����ת�������˷�", glngSys, 1131)) = 1
End Sub

Private Sub LockScreen(blnLock As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ļ
    '����:���˺�
    '����:2018-09-12 10:54:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean
    
    blnEnabled = Not blnLock
    cmdOk.Enabled = blnEnabled
    cmdCancel.Enabled = blnEnabled
    cmdHelp.Enabled = blnEnabled
    cmdAll(0).Enabled = blnEnabled
    cmdAll(1).Enabled = blnEnabled
    picTop.Enabled = blnEnabled
    mshList.Enabled = blnEnabled
End Sub

Private Sub cmdOk_Click()
    Dim i As Long, strNO As String, strNos As String
    Dim blnThirdAllDel As Boolean, bnYBAllDel As Boolean
    Dim lng����ID As Long, str���ݺ� As String, intInsure As Long
    Dim strReplenishNo As String, strNotSelectNos As String
    Dim varData As Variant, strTemp As String, blnErrBill As Boolean
    
    mstrNOs = ""
    If mlngPatient = 0 Then
        MsgBox "δ���ֲ�����Ϣ�����飡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    zlCommFun.ShowFlash "����׼��ת�����ݣ����Ժ�..."
    
    'ֱ�ӱ���
    With mshList
        strNO = ""
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = "��ת��" And .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
            
                lng����ID = Val(.TextMatrix(i, .ColIndex("����ID")))
                str���ݺ� = .TextMatrix(i, .ColIndex("���ݺ�"))
                intInsure = Val(.TextMatrix(i, .ColIndex("����")))
                strReplenishNo = "": strNotSelectNos = ""
                blnErrBill = False
                
                If InStr(1, "," & strNO, "," & str���ݺ� & ",") = 0 Then
                    strNO = strNO & "," & str���ݺ�
                End If
                
                If .TextMatrix(i, .ColIndex("����")) = "�շѵ�" Then
                    If CheckBillExistReplenishData(1, , str���ݺ�, strReplenishNo, blnErrBill) Then
                        If blnErrBill Then
                            zlCommFun.StopFlash
                            MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼�ѽ���ҽ��������㣬���������쳣����״̬�����ȵ������ղ�����㡿���д���", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        If CheckReplenishAllNosIsSelected(strReplenishNo, .TextMatrix(i, .ColIndex("����")), strNotSelectNos) = False Then
                            zlCommFun.StopFlash
                            MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼�ѽ���ҽ��������㣬���µ���Ҳ����һ��ת����" & vbCrLf & strNotSelectNos, vbInformation, gstrSysName
                            Exit Sub
                        End If
                        '��ȡҽ������
                        intInsure = GetReplenishInsure(strReplenishNo)
                        If intInsure = 0 Then
                            zlCommFun.StopFlash
                            MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼�ѽ���ҽ��������㣬��δ��ȡ��ҽ������,����ת����", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        '���ҽ���Ƿ��ܹ�ԭ������
                        strTemp = CheckInsureCancel(mlngPatient, intInsure, strReplenishNo, True)
                        If strTemp <> "" Then
                            zlCommFun.StopFlash
                            MsgBox strTemp, vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                If strReplenishNo = "" Then
                    If intInsure <> 0 Then
                        '���ҽ�������Ƿ�ȫת��
                        If IsYBSingle(str���ݺ�, bnYBAllDel, blnThirdAllDel) = False Then
                            If CheckBalanceAllNosIsSelected(lng����ID, .TextMatrix(i, .ColIndex("����")), strNos) = False Then
                                zlCommFun.StopFlash
                                MsgBox "ҽ�����ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼����δת��ȫ����ؽ��㵥��,���ܼ���!", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            
                            '���ҽ���ֵ��ݣ�������Ϊȫ�ˣ�Ŀǰֻ�ܽ�ֹת��
                            If InStr(strNos, ",") > 0 And bnYBAllDel = False And blnThirdAllDel Then
                                MsgBox "�ݲ�֧���ڱ�������תסԺ�����д���ҽ���ֵ��ݽ��㣬��һ��ͨ����ȫ�˵���������ܳɹ�תסԺ�ĵ������£�" & vbCrLf & strNos, vbInformation + vbOKOnly, gstrSysName
                                zlCommFun.StopFlash
                                Exit Sub
                            End If
                        End If
                    Else
                        If CheckAllTurn(str���ݺ�) = True Then
                            If CheckBalanceAllNosIsSelected(lng����ID, .TextMatrix(i, .ColIndex("����"))) = False Then
                                zlCommFun.StopFlash
                                MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼����δת��ȫ����ؽ��㵥��,���ܼ���!", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                mstrNOs = mstrNOs & ";" & str���ݺ� & "," & .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) & "," & _
                    lng����ID & "," & intInsure & "," & .TextMatrix(i, .ColIndex("����")) & "," & strReplenishNo
            End If
        Next
    End With
    If strNO <> "" Then strNO = Mid(strNO, 2)
    If mstrNOs <> "" Then mstrNOs = Mid(mstrNOs, 2)
    
    If mstrNOs = "" Then
        zlCommFun.StopFlash
        MsgBox "�㻹δѡ��Ҫת��סԺ���õĵ��ݣ��������̣�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    varData = Split(strNO, ","): strNO = ""
    For i = 0 To UBound(varData)
        If i > 60 Then strNO = strNO & ",...": Exit For
        strNO = strNO & IIf(strNO = "", "", ",")
        strNO = strNO & IIf(i > 0 And i Mod 6 = 0, vbCrLf, "")
        strNO = strNO & varData(i)
    Next
    If MsgBox("���Ƿ���Ҫ�������������ת��סԺ������" & vbCrLf & _
        strNO, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        zlCommFun.StopFlash
        mstrNOs = ""
        Exit Sub
    End If
    
    '����Ҫѡ����
    If mblnSelPati = False Then Unload Me: Exit Sub
    
    Err = 0: On Error GoTo ErrHand:
    If Val(NVL(mrsInfo!��ҳID)) = 0 Then
        zlCommFun.StopFlash
        MsgBox "�ò��˻�δ��Ժ,�����������תסԺ����,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    LockScreen True
    If ExecuteTurn(Me, mlngModule, mstrPrivs, mstrNOs, NVL(mrsInfo!סԺ��), Val(NVL(mrsInfo!��ҳID)), _
        CDate(Format(mrsInfo!��Ժ����, "yyyy-mm-dd HH:MM:SS")), Val(NVL(mrsInfo!��Ժ����ID)), Val(NVL(mrsInfo!��Ժ����ID))) = False Then
        LockScreen False
        Set mrsFeeList = Nothing
        Call cmdRefresh_Click
        zlCommFun.StopFlash
        Exit Sub
    Else
        If Val(txtPatient.Tag) <> 0 And Val(txtPatient.Tag) = Val(NVL(mrsInfo!����ID)) Then mblnRefreshData = True
    End If
    zlCommFun.StopFlash
    LockScreen False
    
    If mlngModule = 1137 Then
       txtPatient.Text = ""
       Set mrsInfo = Nothing
       mshDetail.Clear 1
       mshDetail.Rows = 2
       mshList.Clear 1
       mshList.Rows = 2
       vsBalance.Clear 1
       vsBalance.Rows = 2
       zlControl.ControlSetFocus txtPatient
       mlngPatient = 0
       Exit Sub
    End If
    Unload Me
    Exit Sub
ErrHand:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockScreen False
End Sub

Private Function GetReplenishAllNos(ByVal strNO As String) As String
    '��ȡ�����������з��õ���
    '���أ�
    '   �����������з��õ���:A001,A002,...
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strNos As String
    
    On Error GoTo ErrHandler
    strSql = _
        " Select Distinct a.No" & vbNewLine & _
        " From ������ü�¼ A, ������ü�¼ B, ���ò����¼ C" & vbNewLine & _
        " Where a.No = b.No And a.��� = b.��� And a.��¼���� In (1, 11)" & vbNewLine & _
        "       And b.����id = c.�շѽ���id" & vbNewLine & _
        "       And c.��¼���� = 1 And c.���ӱ�־ = 0 And c.No = [1]" & vbNewLine & _
        " Group By a.No, a.���" & vbNewLine & _
        " Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    Do While Not rsTmp.EOF
        strNos = strNos & "," & NVL(rsTmp!NO)
        rsTmp.MoveNext
    Loop
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    GetReplenishAllNos = strNos
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckReplenishAllNosIsSelected(ByVal strNO As String, ByVal str���� As String, _
    Optional ByRef strNotSelectNos As String) As Boolean
    '��鲹����������ʣ��δ�˷��ñ����Ƿ�ѡ����ת��
    '��Σ�
    '   str���� �շѵ�/���ʵ�
    '���Σ�
    '   strNotSelectNos û�б�ѡ�����Ҫһ��ת���ĵ���
    Dim i As Integer, k As Long, blnFind As Boolean
    Dim strNos As String, varNos As Variant
    
    On Error GoTo ErrHandler
    strNotSelectNos = ""
    strNos = GetReplenishAllNos(strNO)
    
    varNos = Split(strNos, ",")
    With mshList
        For i = 0 To UBound(varNos)
            blnFind = False
            For k = 1 To .Rows - 1
                If .TextMatrix(k, .ColIndex("����")) = str���� And .TextMatrix(k, .ColIndex("���ݺ�")) = varNos(i) Then
                    If .TextMatrix(k, .ColIndex("���")) = "��ת��" And .TextMatrix(k, .ColIndex("ѡ��")) = "��" Then
                        blnFind = True: Exit For
                    End If
                End If
            Next
            
            If blnFind = False Then
                strNotSelectNos = strNotSelectNos & "," & varNos(i)
            End If
        Next
    End With
    
    If strNotSelectNos <> "" Then
        strNotSelectNos = Mid(strNotSelectNos, 2)
        Exit Function
    End If
    CheckReplenishAllNosIsSelected = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetReplenishInsure(ByVal strNO As String) As Long
    '��ȡ��������ҽ������
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSql = _
        " Select Max(b.����) As ����" & vbNewLine & _
        " From ����Ԥ����¼ A, ���ս����¼ B, ���ò����¼ C" & vbNewLine & _
        " Where a.����id = b.��¼id And a.��¼���� = 6" & vbNewLine & _
        "       And a.����id = c.����id And c.��¼���� = 1" & vbNewLine & _
        "       And c.��¼״̬ In(1,3) And c.���ӱ�־ = 0 And c.No = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    If Not rsTmp.EOF Then GetReplenishInsure = NVL(rsTmp!����)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBalanceAllNosIsSelected(ByVal lng����ID As Long, ByVal str���� As String, _
    Optional ByRef strNos_Out As String) As Boolean
    '���һ�ν��������ʣ��δ�˷��ñ����Ƿ�ѡ����ת��
    '��Σ�
    '   str���� �շѵ�/���ʵ�
    '����:
    '   strNos_Out-��ǰһ�ν��ʵ�ʣ����õ��ݣ�����ö��ŷ���
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim i As Integer, blnFind As Boolean, blnNotIsSelected As Boolean
    
    On Error GoTo ErrHandler
    strNos_Out = ""
    strSql = _
        " Select Distinct a.No" & vbNewLine & _
        " From ������ü�¼ A, ������ü�¼ B" & vbNewLine & _
        " Where a.No = b.No And Mod(a.��¼����,10) = Mod(b.��¼����,10)" & vbNewLine & _
        "       And a.���=b.��� And b.����id = [1]" & vbNewLine & _
        " Group By a.No,a.���" & vbNewLine & _
        " Having Nvl(Sum(Nvl(a.����,1)*a.����),0) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    Do While Not rsTmp.EOF
        With mshList
            If blnNotIsSelected = False Then
                blnFind = False
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("����")) = str���� And .TextMatrix(i, .ColIndex("���ݺ�")) = NVL(rsTmp!NO) Then
                        If .TextMatrix(i, .ColIndex("���")) = "��ת��" And .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
                            blnFind = True: Exit For
                        End If
                    End If
                Next
                If blnFind = False Then blnNotIsSelected = True
            End If
            strNos_Out = strNos_Out & "," & NVL(rsTmp!NO)
        End With
        rsTmp.MoveNext
    Loop
    strNos_Out = Mid(strNos_Out, 2)
    CheckBalanceAllNosIsSelected = Not blnNotIsSelected
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Activate()
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Call picTop_Resize
End Sub

Private Function Get�����ʻ�����() As String
    '����:��ȡ��������ʻ�����
    Dim rs���㷽ʽ As ADODB.Recordset
    
    On Error GoTo errHandle
    Set rs���㷽ʽ = Get���㷽ʽ("�շ�", "3")
    If rs���㷽ʽ.EOF Then Exit Function
    
    Get�����ʻ����� = NVL(rs���㷽ʽ!����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Initialize()
    Call zlBillPrint_Initialize
End Sub

Private Sub Form_Load()
    Dim strTmp As String, Datsys As Date
    
    If CreateExpenceSvr(mobjExpenceSvr, mlngModule) = False Then Exit Sub
    If Not gobjSquare Is Nothing Then
        Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
         '��ʼ����صı������ݼ�
        Set mtySquareCard.rsSquare = New ADODB.Recordset
        mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
        If Not gobjSquare.objSquareCard Is Nothing Then IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        Set mobjSquare = gobjSquare.objSquareCard
    End If
    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTmp)
    mintIDKind = Val(strTmp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    mstrTitle = Me.Caption
    
    Call RestoreWinState(Me, App.ProductName)
    IDKindTime.NotAutoAppendKind = True
    IDKindTime.IDKindStr = "����ʱ��|����ʱ��|0|0|0|0|0|0|0|0|0;�Ǽ�ʱ��|�Ǽ�ʱ��|0|0|0|0|0|0|0|0|0"
    IDKindTime.IDKind = Val(zlDatabase.GetPara("�ϴ�ѡ��ʱ��ͳ������", glngSys, 1143, 0)) + 1
    mbln����תסԺ����� = IIf(Val(zlDatabase.GetPara("����תסԺ�����", glngSys, 1143, 0)) = 1, True, False)
    mbln�������� = Val(zlDatabase.GetPara("����ת�������˷�", glngSys, 1131)) = 1
    mstr�����ʻ� = Get�����ʻ�����()
    
    mblnNotClick = True
    chkShow.Value = IIf(Val(zlDatabase.GetPara("����ʾ��ת������", glngSys, 1131, 1, Array(chkShow))) = 1, 1, 0)
    mblnNotClick = False
    picBalance.BorderStyle = 0: picList.BorderStyle = 0:    picBill.BorderStyle = 0
    Call InitPancel
    Datsys = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ʼʱ��")
    If IsDate(strTmp) Then
        dtpBegin.Value = CDate(strTmp)
    Else
        dtpBegin.Value = Format(DateAdd("d", -3, Datsys), "yyyy-mm-dd 00:00:00")
    End If
    dtpBegin.MaxDate = Format(Datsys, "yyyy-mm-dd 23:59:59")
    If mstrNOs <> "" Then
        strTmp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����ʱ��")
    Else
        strTmp = ""
    End If
    If IsDate(strTmp) Then
        dtpEnd.Value = CDate(strTmp)
    Else
        dtpEnd.Value = Format(Datsys, "yyyy-mm-dd 23:59:59")
    End If
    Call SetVisibleCtl
    Call setHeader: Call SetDetail: Call SetBalanceHead
    Call zlCreateObject
    
    If mblnSelPati = False Then
        Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
        Call SetBillSelected(mstrNOs)
    Else
        If mlngPatient <> 0 Then
            If GetPatient(IDKind.GetCurCard, "-" & mlngPatient, 0) Then
                Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
            End If
        Else
            Call ClearData
        End If
    End If
    If mblnSelPati = False Then
        fraPati.Visible = False: cmdOk.Visible = True
    Else
        fraPati.Visible = True: cmdOk.Visible = True
    End If
    Call picTop_Resize
End Sub

Private Sub SetVisibleCtl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���visible����
    '����:���˺�
    '����:2011-03-29 21:49:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dtpBegin.Visible = Not mbln����תסԺ�����
    dtpEnd.Visible = Not mbln����תסԺ�����
    lbl��.Visible = Not mbln����תסԺ�����
    IDKindTime.Visible = Not mbln����תסԺ�����
End Sub

Private Sub cmdCancel_Click()
    mstrNOs = ""
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdRefresh_Click()
    If mlngPatient = 0 Then
        MsgBox "����ѡ���ˣ����飡", vbInformation, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value, False)
    If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Terminate()
    Call zlBillPrint_Terminate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ʼʱ��", Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss")
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����ʱ��", Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")
    Call SaveWinState(Me, App.ProductName)
    Set mtySquareCard.rsSquare = Nothing
    Call zlDatabase.SetPara("����ʾ��ת������", chkShow.Value, glngSys, 1131)
    zlDatabase.SetPara "�ϴ�ѡ��ʱ��ͳ������", IDKindTime.IDKind - 1, glngSys, 1143, InStr(1, mstrPrivs, ";��������;") > 0
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "����תסԺ��ϸ�б�", True
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "����תסԺ�����б�", True
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "����תסԺ�б�", True
    Call zlCloseObject
    Set mrsFeeList = Nothing
End Sub
 
Private Sub IDKind_Click(objCard As zlOneCardComLib.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.���� Like "*IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, True, Trim(txtPatient.Text))
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub

   '    zlReadCard(frmMain As Object, _
   '    ByVal lngModule As Long, _
   '    ByVal lngCardTypeID As Long, _
   '    ByVal blnOlnyCardNO As Boolean, _
   '    ByVal strExpand As String, _
   '    ByRef strOutCardNO As String, _
   '    ByRef strOutPatiInforXML As String) As Boolean
   '    '---------------------------------------------------------------------------------------------------------------------------------------------
   '    '����:�����ӿ�
   '    '���:frmMain-���õĸ�����
   '    '       lngModule-���õ�ģ���
   '    '       strExpand-��չ����,������
   '    '       blnOlnyCardNO-������ȡ����
   '    '����:strOutCardNO-���صĿ���
   '    '       strOutPatiInforXML-(������Ϣ����.XML��)
   '    '����:��������    True:���óɹ�,False:����ʧ��\
   If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
   txtPatient.Text = strOutCardNO
   If txtPatient.Text <> "" Then Call FindPati(objCard, True, Trim(txtPatient.Text))
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlOneCardComLib.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mshDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "����תסԺ��ϸ�б�", True
End Sub

Private Sub mshDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "����תסԺ��ϸ�б�", True
End Sub

Private Sub mshList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "����תסԺ�б�", True
End Sub

Private Sub mshList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNO As String, str���� As String
    
    If NewRow = OldRow Then Exit Sub
    With mshList
        strNO = Trim(.TextMatrix(NewRow, .ColIndex("���ݺ�")))
        str���� = Trim(.TextMatrix(NewRow, .ColIndex("����")))
        If NewRow = 0 Or strNO = "" Then
            mshDetail.Clear 1: mshDetail.Rows = 2
            Call SetDetail
        Else
            Call ShowDetail(str����, strNO)
        End If
    End With
End Sub

Private Sub mshList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "����תסԺ�б�", True
End Sub

Private Sub mshList_DblClick()
    With mshList
        If .MouseRow = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("���ݺ�")) = "" Then Exit Sub
        Call SetRowSelected(.Row, Trim(.TextMatrix(.Row, .ColIndex("ѡ��"))) = "")
    End With
    Call SetSumMoney
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
     If KeyAscii <> 32 Then Exit Sub
    With mshList
        If .TextMatrix(.Row, .ColIndex("���ݺ�")) = "" Then Exit Sub
       Call SetRowSelected(.Row, Trim(.TextMatrix(.Row, .ColIndex("ѡ��"))) = "")
    End With
    Call SetSumMoney
End Sub

Private Sub cmdAll_Click(Index As Integer)
    Dim i As Long
    
    With mshList
        .Redraw = False
        For i = 1 To .Rows - 1
            If Index = 1 Then
                .TextMatrix(i, .ColIndex("ѡ��")) = ""
            Else
                If Not SetRowSelected(i, Index = 0) Then
                    .Row = i: .Col = 0: .ColSel = .Cols - 1
                    Call mshList_AfterRowColChange(0, 0, .Row, .Col)
                    Exit For
                End If
            End If
        Next
        .Redraw = True
    End With
    Call SetSumMoney(Index = 1)
End Sub

Private Function CheckInsureCancel(ByVal lng����ID As Long, ByVal lngInsure As Long, _
    ByVal strNO As String, Optional ByVal bln������ As Long) As String
    '���ҽ���Ƿ��ܹ�ԭ������
    '���أ�����ԭ�����ϣ��򷵻ؿգ����򣬷�����ʾ��Ϣ
    Dim strTmp As String, i As Integer
    Dim arrBalanceType As Variant, strBalanceType As String
    
    On Error GoTo ErrHandler
    If Not gclsInsure.GetCapability(support�����������, lng����ID, lngInsure) Then
        CheckInsureCancel = IIf(bln������, "ҽ���������", "") & "����[" & strNO & "]�Ĳ������಻֧������������ϣ�������ת����"
        Exit Function
    Else
        '���жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
        strTmp = GetBalanceType(strNO, bln������)
        arrBalanceType = Split(strTmp, ",")
        For i = 0 To UBound(arrBalanceType)
            strBalanceType = arrBalanceType(i)
            If Not gclsInsure.GetCapability(support�����������, lng����ID, lngInsure, strBalanceType) Then
                CheckInsureCancel = IIf(bln������, "ҽ���������", "") & "����[" & strNO & "]�Ĳ������಻֧��" & strBalanceType & "�������ϣ�������ת����"
                Exit Function
            End If
        Next
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetRowSelected(ByVal lngRow As Long, blnSelect As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ�е�ѡ��״̬
    '       ����Ƕ��ŵ����е�һ��,����ͬʱ���ö����е���������
    '����:���˺�
    '����:2011-02-21 16:10:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, strTmp As String
    Dim str���� As String
    Dim blnAll As Boolean
    
    With mshList
        If .TextMatrix(lngRow, .ColIndex("���")) = "��ת��" And .TextMatrix(lngRow, .ColIndex("ѡ��")) <> IIf(blnSelect, "��", "") Then
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
            str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
            strNO = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
            
            If intInsure > 0 And blnSelect And str���� = "�շѵ�" Then
                strTmp = CheckInsureCancel(mlngPatient, intInsure, strNO)
                If strTmp <> "" Then
                    sta.Panels(2).Text = strTmp
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
                    Exit Function
                End If
            End If
            
            .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
            If str���� = "�շѵ�" Then
                If intInsure > 0 Then      'ȫ��ѡ���ȡ��
                    blnAll = gclsInsure.GetCapability(support�൥���շѱ���ȫ��, mlngPatient, intInsure)
                    If Not blnAll Then blnAll = Not IsYBSingle(strNO)
                    If blnAll Then If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
                    
                Else '�ֽ�����Ҫ����൥���շ����
                    If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
                End If
            End If
        End If
        If .TextMatrix(lngRow, .ColIndex("���")) = "����ת��" Then .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
    End With
    SetRowSelected = True
End Function

Private Function CheckAllTurn(ByVal strNO As String) As Boolean
    Dim strSql As String, rsData As ADODB.Recordset, lngCardTypeID As Long
    Dim strCardTypeIDs As String, strTemp As String
    Dim strWhere As String
       
    On Error GoTo errHandle
           
    strWhere = "And  Not Exists(select 1 From ҽ��������ϸ Where NO=[1] And A.�����ID=�����ID and A.��������ID=��������ID) "
    
    strSql = "" & _
    "   Select A.���㷽ʽ,nvl(A.�����ID,0) as �����ID,nvl(A.���㿨���,0) as ���㿨���,nvl(A.��������ID,0) as ��������ID," & _
    "       max(Nvl(D.�Ƿ�ȫ��,nvl(E.�Ƿ�ȫ��,0))) as �Ƿ�ȫ��,nvl(max(decode(nvl(C.����,0),3,1,4,1,0)),0) as �Ƿ�ҽ��" & vbNewLine & _
    "   From ����Ԥ����¼ A, " & _
    "       (   Select Distinct ����id  " & _
    "           From ������ü�¼ " & _
    "           Where Mod(��¼����,10) = 1 And ��¼״̬ <> 0  " & _
    "                 And NO In (   Select Distinct NO  From ������ü�¼ Where ����id In  (Select ����id" & vbNewLine & _
    "                               From ����Ԥ����¼" & vbNewLine & _
    "                               Where ������� In (Select b.�������" & vbNewLine & _
    "                                          From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
    "                                          Where a.No = [1] And a.��¼���� = 1 And a.��¼״̬ <> 0 And a.����id = b.����id))) " & vbNewLine & _
    "                 " & _
    "         ) B,���㷽ʽ C,ҽ�ƿ���� D,���ѿ����Ŀ¼ E" & vbNewLine & _
    "   Where a.����id = b.����id And a.��¼���� = 3 And A.���㷽ʽ=C.����(+) And A.�����ID=D.ID(+) and A.���㿨���=E.���(+) " & vbNewLine & _
    "       " & strWhere & vbNewLine & _
    "   Group By A.���㷽ʽ,nvl(A.�����ID,0),nvl(A.���㿨���,0),nvl(A.��������ID,0) " & vbNewLine & _
    "   Having Sum(��Ԥ��) <> 0" & _
    "   Order by �����ID,��������ID"

    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    If rsData.EOF Then CheckAllTurn = False: Exit Function
    rsData.Filter = "�Ƿ�ȫ��=1"
    If Not rsData.EOF Then CheckAllTurn = True: Exit Function   '����ȫ�˵�������������������
    
    rsData.Filter = "�Ƿ�ҽ�� =1 And �����ID<>0"
    If Not rsData.EOF Then CheckAllTurn = True: Exit Function   'һ��ͨ����ҽ������ʱ����ȫ��(��SQL���ų��˷ֵ��ݽ����)������������
    
    rsData.Filter = "�����ID<>0"
    rsData.Sort = "�����ID,��������ID"
    
    With rsData
        strCardTypeIDs = ""
        Do While Not .EOF
            lngCardTypeID = Val(NVL(rsData!�����ID))
            strTemp = lngCardTypeID & ":" & Val(NVL(rsData!��������ID))
            If InStr(strCardTypeIDs & ",", "," & strTemp & ",") > 0 Then    '�϶���һ��ͨ���ڶ��ֽ��㷽ʽ������Ҳ����ȫ��
                CheckAllTurn = True: Exit Function   'һ��ͨ����ҽ������ʱ����ȫ��(��SQL���ų��˷ֵ��ݽ����)������������
            End If
            strCardTypeIDs = strCardTypeIDs & "," & strTemp
            .MoveNext
        Loop
    End With
    CheckAllTurn = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intInsure As Integer) As Boolean
'����:���ŵ�������ѡ���ȡ��
'     ���ҽ�����ŵ���Ҫ�������˷�,ѡ������һ��ʱ,ȫѡ����,ȡ��ʱȫȡ��
    Dim i As Long, j As Long, k As Long, strNO As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant, blnAllTurn As Boolean
    Dim str���� As String, strReplenishNo As String
    Dim strNos As String, varNos As Variant
    
    With mshList
        str���� = .TextMatrix(lngRow, .ColIndex("����"))
        If str���� = "���ʵ�" Then SetMultiOther = True: Exit Function
        If intInsure = 0 Then
            '����Ƿ�Ϊ�����㵥��
            If CheckBillExistReplenishData(1, , .TextMatrix(lngRow, .ColIndex("���ݺ�")), strReplenishNo) Then
                strNos = GetReplenishAllNos(strReplenishNo)
                varNos = Split(strNos, ",")
                For i = 0 To UBound(varNos)
                    For k = 1 To .Rows - 1
                        If .TextMatrix(k, .ColIndex("����")) = str���� And .TextMatrix(k, .ColIndex("���ݺ�")) = varNos(i) Then
                            .TextMatrix(k, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
                            Exit For
                        End If
                    Next
                Next
                SetMultiOther = True
                Exit Function
            End If
        
            blnAllTurn = CheckAllTurn(.TextMatrix(lngRow, .ColIndex("���ݺ�")))
            
            If gblnMultiBalance Or blnAllTurn Then     '   �൥��,���ֽ��㷽ʽ
                '33635:ԭ���Ƕ൥���Ҷ��ֽ��㷽ʽ,���ܲ�����
                strNO = ""
                For k = 1 To .Rows - 1
                      If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                            And .TextMatrix(k, .ColIndex("����")) = str���� _
                            And Trim(.TextMatrix(lngRow, .ColIndex("����ID"))) <> "" Then
                            strNO = strNO & "," & .TextMatrix(k, .ColIndex("���ݺ�"))
                      End If
                Next
                If strNO <> "" Then strNO = Mid(strNO, 2)
                If InStr(1, strNO, ",") > 0 Then    '֤��Ϊ�൥��
                    For k = 1 To .Rows - 1
                          If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                            And .TextMatrix(k, .ColIndex("����")) = str���� _
                            And Trim(.TextMatrix(lngRow, .ColIndex("����ID"))) <> "" Then
                                .TextMatrix(k, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
                          End If
                    Next
                End If
            End If
            SetMultiOther = True
            Exit Function
        End If
        
        If IsYBSingle(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) Then SetMultiOther = True: Exit Function
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = "��ת��" _
                And .TextMatrix(i, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                And i <> lngRow Then
                If .TextMatrix(i, .ColIndex("ѡ��")) <> .TextMatrix(lngRow, .ColIndex("ѡ��")) Then
                   If intInsure <> 0 And blnSelect Then
                        strNO = .TextMatrix(i, .ColIndex("���ݺ�"))
                        '�жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                         strTmp = GetBalanceType(strNO)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                 strBalanceType = arrBalanceType(j)
                                 If Not gclsInsure.GetCapability(support�����������, mlngPatient, intInsure, strBalanceType) Then
                                     sta.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
                                     For k = 1 To .Rows - 1
                                        If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(i, .ColIndex("����ID")) _
                                            And .TextMatrix(k, .ColIndex("����")) = str���� Then
                                            .TextMatrix(k, .ColIndex("ѡ��")) = ""
                                        End If
                                     Next
                                     Exit Function
                                 End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function

Private Function zlGetOneCard(ByVal strIDs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ�һ��ͨ���㵥��
    '���:strIDs-����ID(����Ϊ����,�ö��ŷ���)
    '����:
    '����:һ��ͨ��������,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    
    On Error GoTo errHandle
    
     strSql = "" & _
    "   Select /*+cardinality(j,10)*/  A.����ID,A.��λ�ʺ�, A.�������, B.ҽԺ����, A.��Ԥ�� as ���" & vbNewLine & _
    "   From ����Ԥ����¼ A, һ��ͨĿ¼ B,Table(f_Num2list([1])) J " & vbNewLine & _
    "   Where A.����id = J.Column_Value  And A.���㷽ʽ = B.���㷽ʽ" & _
    "   Order by ����ID"
    Set zlGetOneCard = zlDatabase.OpenSQLRecord(strSql, "��ȡһ��ͨ��������", strIDs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetBalanceType(ByVal strNO As String, _
    Optional ByVal bln������ As Boolean) As String
    '����:��ȡһ�ŵ����е�ҽ�����㷽ʽ��
    Dim rsTmp As ADODB.Recordset, strSql As String
        
    On Error GoTo errH
    If bln������ Then
        strSql = _
            " Select Distinct a.���㷽ʽ" & vbNewLine & _
            " From ����Ԥ����¼ A, ���㷽ʽ B, ���ò����¼ C" & vbNewLine & _
            " Where a.���㷽ʽ = b.���� And a.��¼���� = 6 And b.���� In(3,4)" & vbNewLine & _
            "       And a.����id = c.����id And c.��¼���� = 1" & vbNewLine & _
            "       And c.���ӱ�־ = 0 And Nvl(c.����״̬, 0) <> 2 And c.No = [1]"
    Else
        strSql = _
            " Select Distinct a.���㷽ʽ" & vbNewLine & _
            " From ����Ԥ����¼ A, ���㷽ʽ B, ������ü�¼ C" & vbNewLine & _
            " Where a.���㷽ʽ = b.���� And b.���� In(3,4)" & vbNewLine & _
            "       And a.����id = c.����ID And c.��¼���� = 1 And c.No = [1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    Do While Not rsTmp.EOF
        GetBalanceType = GetBalanceType & "," & rsTmp!���㷽ʽ
        rsTmp.MoveNext
    Loop
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowDetail(ByVal str���� As String, ByVal strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ����
    '���:str����:�շѵ�(���ʵ�)
    '        strNO-���ݺ�
    '����:���˺�
    '����:2011-02-22 11:14:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    Err = 0: On Error GoTo errH
    If mshList.Row < 0 Then Exit Sub
    If mshList.TextMatrix(mshList.Row, mshList.ColIndex("���")) = "��ת��" Then
        strSql = _
        " Select C.���� As ���, max(Decode(Decode(J1.�������||':'||J1.ҽ������,'7:***',1,0), 1, '***', Nvl(E.����, B.����))) As ����, " & _
        "       B.���, A.���㵥λ As ��λ, Sum(Nvl(A.����, 1) * A.����) As ����," & _
        "       LTrim(To_Char(A.��׼����, '999990.00000')) As ����, LTrim(To_Char(Sum(A.Ӧ�ս��), '99999" & gstrDec & "')) As Ӧ�ս��," & _
        "       LTrim(To_Char(Sum(A.ʵ�ս��), '99999" & gstrDec & "')) As ʵ�ս��, D.���� As ִ�п���, 3 As ��¼״̬" & _
        " From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E,����ҽ����¼ J1" & _
        " Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And Mod(A.��¼����,10) = [2]" & _
        "      And A.ҽ�����=J1.id(+)" & _
        "      And A.��¼״̬ In (2,3) And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3 And Nvl(a.���ӱ�־, 0) <> 9 " & _
        " Group By A.��׼����,A.���, C.����, B.���, A.���㵥λ, D.����" & _
        " Having Sum(A.����) <> 0 " & _
        " Union All" & _
        " Select C.���� As ���,max(Decode(Decode(J1.�������||':'||J1.ҽ������,'7:***',1,0), 1, '***', Nvl(E.����, B.����))) As ����," & _
        "       B.���, A.���㵥λ As ��λ, Sum(Nvl(A.����, 1) * A.����) As ����," & _
        "       LTrim(To_Char(A.��׼����, '999990.00000')) As ����, LTrim(To_Char(Sum(A.Ӧ�ս��), '99999" & gstrDec & "')) As Ӧ�ս��," & _
        "       LTrim(To_Char(Sum(A.ʵ�ս��), '99999" & gstrDec & "')) As ʵ�ս��, D.���� As ִ�п���, 1 As ��¼״̬" & _
        " From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E,����ҽ����¼ J1" & _
        " Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And Mod(A.��¼����,10) = [2] " & _
        "      And A.ҽ�����=J1.id(+) " & _
        "      And A.��¼״̬=1 And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3 And Nvl(A.���ӱ�־,0) <> 9 " & _
        " Group By A.��׼����,A.���, C.����, B.���, A.���㵥λ, D.����" & _
        " Having Sum(A.����) <> 0 "
    
    ElseIf mshList.TextMatrix(mshList.Row, mshList.ColIndex("���")) = "����ת��" Then
        strSql = _
        " Select C.���� As ���, max(Decode(Decode(J1.�������||':'||J1.ҽ������,'7:***',1,0), 1, '***', Nvl(E.����, B.����))) As ����," & _
        "       B.���, A.���㵥λ As ��λ, Sum(Nvl(A.����, 1) * A.����) As ����," & _
        "       LTrim(To_Char(A.��׼����, '999990.00000')) As ����, LTrim(To_Char(Sum(A.Ӧ�ս��), '99999" & gstrDec & "')) As Ӧ�ս��," & _
        "       LTrim(To_Char(Sum(A.ʵ�ս��), '99999" & gstrDec & "')) As ʵ�ս��, D.���� As ִ�п���, 2 As ��¼״̬" & _
        " From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E,����ҽ����¼ J1" & _
        " Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And Mod(A.��¼����,10) = [2] " & _
        "      And A.ҽ�����=J1.id(+) " & _
        "      And A.��¼״̬ In (1,3) And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3 And Nvl(A.���ӱ�־,0) <> 9 " & _
        " Group By A.��׼����,A.���, C.����,B.���, A.���㵥λ, D.���� Having Sum(A.����) <> 0 "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, IIf(str���� = "���ʵ�", 2, 1))
    
    mshDetail.Redraw = flexRDNone
    mshDetail.Clear
    Set mshDetail.DataSource = rsTmp
    If rsTmp.EOF Then mshDetail.Rows = 2
    Call SetDetail
    mshDetail.Redraw = flexRDBuffered
    Exit Sub
errH:
    mshDetail.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    strHead = "���,1,650|����,1,1500|���,1,1450|��λ,4,500|����,7,500|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ִ�п���,4,1000|��¼״̬,4,0"
    With mshDetail
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        .ColHidden(9) = True
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 9)) = 1 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbBlack
            'If Val(.TextMatrix(i, 9)) = 2 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbRed
            If Val(.TextMatrix(i, 9)) = 3 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbBlue
        Next i
        .AutoSize 0, .Cols - 1
        zl_vsGrid_Para_Restore 1131, mshDetail, Me.Caption, "����תסԺ��ϸ�б�", True
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub
Private Sub SetBalanceHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����б�
    '����:���˺�
    '����:2011-03-28 11:27:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strHead As String
    Dim i As Long
    strHead = "���,4,650|��־,1,600|���㵥��,1,1500|������,7,1000|���㷢Ʊ,1, 2600"
    With vsBalance
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        zl_vsGrid_Para_Restore 1131, vsBalance, Me.Caption, "����תסԺ�����б�", True
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub
Private Sub picBill_Resize()
    Err = 0: On Error Resume Next
    With picBill
        mshList.Left = .ScaleLeft
        mshList.Top = .ScaleTop
        mshList.width = .ScaleWidth
        mshList.Height = .ScaleHeight
    End With
End Sub
Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Top = .ScaleTop
        vsBalance.width = .ScaleWidth
        lblSum.Top = .ScaleHeight - lblSum.Height
        vsBalance.Height = lblSum.Top - mshDetail.Top
    End With
End Sub

Private Sub picBottom_Resize()
    Err = 0: On Error Resume Next
    With picBottom
            cmdCancel.Left = .ScaleLeft + .ScaleWidth - cmdCancel.width - 400
            cmdOk.Left = cmdCancel.Left - cmdOk.width - 20
            cmdOk.Top = cmdCancel.Top
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        mshDetail.Left = .ScaleLeft
        mshDetail.Top = .ScaleTop
        mshDetail.width = .ScaleWidth
        mshDetail.Height = .ScaleHeight
    End With
End Sub

Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    If mblnSelPati Then
        fraPati.Left = picTop.ScaleLeft + 150
        IDKindTime.Left = fraPati.Left + fraPati.width + 20
    Else
        IDKindTime.Left = picTop.ScaleLeft + 150
    End If
    dtpBegin.Left = IDKindTime.Left + IDKindTime.width + 30
    lbl��.Left = dtpBegin.Left + dtpBegin.width + 50
    dtpEnd.Left = lbl��.Left + lbl��.width + 50
    
    fraFixed.Left = fraPati.Left + IIf(mbln����תסԺ����� And mblnSelPati, fraPati.width + 150, 150)
    fraFixed.Top = IIf(mbln����תסԺ�����, 80, 450)
End Sub
Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
End Sub
Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If Not mobjIDCard Is Nothing And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not txtPatient.Locked Then Call IDKind.SetAutoReadCard(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    
    If txtPatient.Locked Then Exit Sub
    '����ѡ����
    If Not (Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13) Then
       If IDKind.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        End If
    End If
    
    Me.Refresh
    'ˢ����ϻ���������س�
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-10-18 16:35:27
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    If GetPatient(objCard, strInput, blnCard) Then
        '69526:������,2014-02-13,��Ժ�����޷���������תסԺ����
        If Val(zlDatabase.GetPara("��Ժ������������תסԺ", glngSys, 1137, "0")) = 0 Then
            If HaveOut(mlngPatient) Then
                MsgBox "����" & mrsInfo!���� & "�Ѿ���Ժ��δ����סԺ������������������תסԺ������", vbInformation, gstrSysName
                txtPatient.Text = "": mlngPatient = 0
                Call ClearData
                Set mrsInfo = Nothing
                If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
                Exit Sub
            End If
        End If
        
        strSql = "Select 1 From ���˹Һż�¼ Where Nvl(���ӱ�־,0) = 3 And Id = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(NVL(mrsInfo!�Һ�ID)))
        If Not rsTemp.EOF Then
            '�����ﲡ��
            Me.Hide
            On Error Resume Next
            Call frmChargeTurnNew.ShowMe(Me, Val(mrsInfo!�Һ�ID), True)
            Err = 0: On Error GoTo 0
            
            txtPatient.Text = "": mlngPatient = 0
            Call ClearData
            Set mrsInfo = Nothing
            If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
            
            Me.Show vbModal, mobjParent
            Exit Sub
        End If
        
        '��ʱ������ʽ�����¼�Form_Load
        Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
        If mshList.TextMatrix(1, mshList.ColIndex("���ݺ�")) <> "" Then
            If mshList.TextMatrix(1, mshList.ColIndex("ѡ��")) <> "" Then
                If cmdOk.Visible And cmdOk.Enabled Then Call cmdOk.SetFocus
            Else
                If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
            End If
        Else
            If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
        End If
    Else
        txtPatient.Text = "": mlngPatient = 0
        Call ClearData
        If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
    End If
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
End Sub
Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    Call IDKind.SetAutoReadCard(False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If txtPatient.Text <> mrsInfo!���� Then txtPatient.Text = mrsInfo!����
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card '54894
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
                            
 
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��,lng��ҳID=��ȡָ��סԺ�����Ĳ�����Ϣ
    '����:
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmrsInfo�а���������Ϣ,ʧ��ʱmrsInfo=Close,strInput�����������ж��Ƿ�����ʾ��,�����ٴ���ʾû���ҵ�����
    '����:���˺�
    '����:2010-11-09 17:17:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, strRange As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    strSql = _
    " Select b.�Һ�ID,A.����ID,Nvl(B.��ҳID,0) as ��ҳID,A.סԺ��,A.��ǰ����,B.��Ժ����ID,B.��Ժ����ID,B.��Ժ����," & _
    "        Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(B.����,A.����) as ����,A.IC����,A.���￨��,A.����֤��," & _
    "       Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�,C.���� as ��ǰ����,A.��ǰ����ID,D.���� as ��Ժ����,B.��Ժ����ID,A.���� as ����,E.����,E.ҽ����,E.����," & _
    "       A.�Ǽ�ʱ��,Nvl(B.״̬,0) as ״̬,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Nvl(B.��˱�־,0) as ��˱�־,B.��Ժ����,B.��Ժ����,B.��������,B.��������" & _
    " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,ҽ�����˵��� E,ҽ�����˹����� F" & _
    " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID(+) And " & IIf(lng��ҳID = 0, "A.��ҳID=B.��ҳID(+)", "B.��ҳID=[3]") & _
    "           And A.����ID=F.����ID(+) And F.��־(+)=1 And F.ҽ����=E.ҽ����(+) And F.����=E.����(+) And F.���� = E.����(+)" & _
    "           And A.��ǰ����ID=C.ID(+) And B.��Ժ����ID=D.ID(+) "
        
    If blnCard = True And objCard.���� Like "����*" Then  'ˢ��
        lng�����ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSql = strSql & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSql = strSql & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSql = strSql & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSql = strSql & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    Else '��������
        Select Case objCard.����
            Case "����", "��������￨"
                If Not mrsInfo Is Nothing Then
                    If mrsInfo.State = 1 Then
                        If mrsInfo!���� = Trim(txtPatient.Text) Then
                            mlngPatient = Val(NVL(mrsInfo!����ID))
                            GetPatient = True
                            Exit Function
                        End If
                    End If
                End If
                If mintPatientRange > 0 Then
                    Select Case mintPatientRange
                        Case 1  '�κη���δ���岡��
                            strRange = ""
                        Case 2  '���δ����Ĳ���
                            strRange = " And C.��Դ;�� = 4"
                        Case 3  'סԺδ����Ĳ���
                            strRange = " And C.��Դ;�� = 2"
                        Case 4  '����δ����Ĳ���
                            strRange = " And C.��Դ;�� = 1"
                    End Select
                    strPati = " And Exists(Select 1 From ����δ����� C Where C.����id=A.����ID And Nvl(C.��ҳID,0)=A.��ҳID" & strRange & ")"
                End If
                 'ͨ����������
                strPati = "Select A.����ID as ID,A.����ID,A.סԺ��, A.�����, Nvl(b.�Ա�, a.�Ա�) as �Ա�, Nvl(b.����, a.����) as ����, A.סԺ����, A.��ͥ��ַ, A.������λ," & vbNewLine & _
                        "To_Char(A.��������,'YYYY-MM-DD') as ��������,  To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����, To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����" & vbNewLine & _
                        "From ������Ϣ A, ������ҳ B,��Ժ���� C" & vbNewLine & _
                        "Where A.����id = B.����id(+) And A.��ҳID = B.��ҳid(+) And A.ͣ��ʱ�� Is Null And A.����ID=C.����ID And A.���� = [1] " & vbNewLine & strPati & vbNewLine & _
                        "Order By Decode(סԺ��, Null, 1, 0), ��Ժ���� Desc"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If Not mrsInfo Is Nothing Then
                    strInput = Val(mrsInfo!����ID)
                    strSql = strSql & " And A.����ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strSql = strSql & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSql = strSql & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Mid(strInput, 2), strInput, lng��ҳID)
    If mrsInfo.EOF Then Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
    
    txtPatient.Text = NVL(mrsInfo!����): mlngPatient = Val(NVL(mrsInfo!����ID))
    If IsDate(Format(mrsInfo!��Ժ����, "yyyy-mm-dd HH:MM:SS")) Then
        '�������Ϊ��Ժ����,����ת��סԺ�����е��������
        dtpEnd.MaxDate = CDate(Format(mrsInfo!��Ժ����, "yyyy-mm-dd 23:59:59"))
        dtpEnd.Value = dtpEnd.MaxDate
        dtpEnd.MaxDate = dtpEnd.MaxDate + 1
        dtpBegin.MaxDate = dtpEnd.Value
        '   ����: 36609����Ժʱ��Ҫ��һ��,��Ϊ���ܴ��ڲ�����û���������ʱ,����Ժ,��ȥ�������,�Ӷ�����������ת���˵����.
    End If
    
    GetPatient = True
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function
  
Private Function PrintPrePayPrint(ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡԤ����
    '���:strDelDate-����ת������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-16 10:30:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, bytPrepayPrint As Byte
    Dim strNos As String
    
    On Error GoTo errHandle
    If zlStr.IsHavePrivs(mstrPrivs, "Ԥ�����վݴ�ӡ") = False Then
       PrintPrePayPrint = True: Exit Function '����ӡ
    End If
    bytPrepayPrint = Val(zlDatabase.GetPara("����תסԺԤ����ӡ", glngSys, 1131))
    If bytPrepayPrint = 0 Then PrintPrePayPrint = True: Exit Function '����ӡ
    
    strSql = "Select Distinct NO From ����Ԥ����¼ Where ��¼���� = 1 And �տ�ʱ�� = [1] And ժҪ = '����תסԺԤ��'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡתԤ����", CDate(strDelDate))
    If rsTemp.EOF Then
        'û��תΪԤ�����ݣ���Ҳ����ӡ
        PrintPrePayPrint = True: Exit Function
    End If
    If bytPrepayPrint = 2 Then   '��ʾ��ӡ
        If MsgBox("�����������תסԺ����ʱ�������ֽ�Ƚ��㷽ʽתΪ��Ԥ������Ƿ�Ҫ��ӡԤ����Ʊ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            PrintPrePayPrint = True: Exit Function
        End If
    End If
    
    If Val(zlDatabase.GetPara(283, glngSys)) = 1 Then  '112862
        '1-����������Ԥ������һ���Դ�ӡ
        strNos = ""
        Do While Not rsTemp.EOF
            strNos = strNos & "," & NVL(rsTemp!NO)
            rsTemp.MoveNext
        Loop
        If strNos <> "" Then
            strNos = Mid(strNos, 2)
            If zlPrintInvoice(strNos, strDelDate) = False Then Exit Function
        End If
    Else
        '0-�����ɵ�Ԥ�����ݷֱ��ӡ
        Do While Not rsTemp.EOF
            If zlPrintInvoice(NVL(rsTemp!NO), strDelDate) = False Then Exit Function
            rsTemp.MoveNext
        Loop
    End If
    PrintPrePayPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetSumMoney(Optional blnCls As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ú���ʾ�ϼ�
    '����:���˺�
    '����:2011-03-04 14:17:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblSumMoney As Double
    Dim strJzNOs As String, strSfNos As String
    With mshList
        If blnCls = False Then
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ѡ��"))) <> "" Then
                    dblSumMoney = dblSumMoney + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                End If
                If .TextMatrix(i, .ColIndex("���")) = "��ת��" And .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
                    If .TextMatrix(i, .ColIndex("����")) = "���ʵ�" Then
                        strJzNOs = strJzNOs & "," & .TextMatrix(i, .ColIndex("���ݺ�"))
                    Else
                        strSfNos = strSfNos & "," & .TextMatrix(i, .ColIndex("���ݺ�"))
                    End If
                End If
            Next
            If strJzNOs <> "" Then strJzNOs = Mid(strJzNOs, 2)
            If strSfNos <> "" Then strSfNos = Mid(strSfNos, 2)
        Else
            dblSumMoney = 0
        End If
    End With
    lblSum.Caption = "����ת���ϼ�:" & Format(dblSumMoney, "###0.00;-###0.00;0.00;0.00")
    
    '����ѡ�������ͨ��
    Call LoadBalance(strJzNOs, strSfNos)
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = sta.Height + picBottom.Height + 100
End Sub

Private Sub LoadBalance(ByVal strJzNOs As String, ByVal strSfNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ�����Ϣ
    '����:���˺�
    '����:2011-03-28 11:33:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, i As Long
    Dim strSFTable As String, strJzTable As String
    Dim varPara() As Variant
    
    On Error GoTo errHandle
    If strJzNOs = "" And strSfNos = "" Then
        vsBalance.Clear 1: vsBalance.Rows = 2: Exit Sub
    End If
    
    ReDim Preserve varPara(0) As Variant
    If strJzNOs <> "" Then
        If zlGetVarBoundSQL(1, strJzNOs, strJzTable, varPara, UBound(varPara) + 1) = False Then Exit Sub
        strJzTable = _
            " Select A.��־, A.NO, A.������, f_List2str(Cast(COLLECT(distinct C.����) as t_Strlist)) As ��Ʊ�� " & _
            " From (Select /*+cardinality(j,10)*/ '����' As ��־, B.NO, To_Char(Sum(a.���ʽ��),'9999990.00') As ������ " & _
            "       From ������ü�¼ A, ���˽��ʼ�¼ B, (" & strJzTable & ") J " & _
            "       Where A.NO = J.Column_Value  And A.����id = B.ID  And B.��¼״̬=1 And A.��¼���� In (2, 12) " & _
            "       Group By B.NO) A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C,Ʊ��ʹ����ϸ D " & _
            " Where A.NO = B.NO(+) and B.��������(+)=3 And B.ID = C.��ӡid(+) And C.����(+)=1 " & _
            " Group By A.��־, A.NO, A.������"
    End If
    
    If strSfNos <> "" Then
        If zlGetVarBoundSQL(1, strSfNos, strSFTable, varPara, UBound(varPara) + 1) = False Then Exit Sub
        strSFTable = _
            IIf(strJzNOs = "", "", " Union All") & _
            " Select A.��־, A.NO, A.������, f_List2str(Cast(COLLECT(distinct C.����) as t_Strlist))  As ��Ʊ�� " & _
            " From (Select /*+cardinality(j,10)*/ '�շ�' As ��־, A.NO, To_Char(Sum(a.���ʽ��),'9999990.00') As ������ " & _
            "       From ������ü�¼ A, (" & strSFTable & ") J " & _
            "       Where A.NO = J.Column_Value And Mod(A.��¼����,10) = 1 " & _
            "       Group By A.NO) A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C " & _
            " Where A.NO = B.NO(+) and B.��������(+)=1 And B.ID = C.��ӡid(+) And C.����(+)=1 " & _
            " Group By A.��־, A.NO, A.������"
    End If
    strSql = _
        " Select Rownum As ���, ��־, NO As ���㵥��, ������, ��Ʊ�� " & _
        " From (" & strJzTable & strSFTable & ")"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSql, Me.Caption, varPara)
    
    Set vsBalance.DataSource = rsTemp
    If rsTemp.RecordCount = 0 Then
        vsBalance.Rows = 2
    End If
    Call SetBalanceHead
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "����תסԺ�����б�", True
End Sub

Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "����תסԺ�����б�", True
End Sub

Private Function zlPrintInvoice(ByVal strNos As String, ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ʊ����
    '��Σ�
    '   strNos ���δ�ӡԤ�����ݺţ���ʽ��A001,A002,A003,...
    '����:���˺�
    '����:2011-04-02 09:48:13
    '����:36984
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngShareUseID As Long, lng����ID As Long, strInvoice As String
    Dim blnInput As Boolean, blnValid As Boolean
    Dim strSql As String
    Dim intInvoiceFormat As Integer
    
    '����ϸ����Ʊ��ʹ��
    On Error GoTo errHandle
    If gblnPrepayStrict Then
        lngShareUseID = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, 1131, 0)
        '1.�ϸ����Ʊ��ʱ������ʵ�ʵ�Ʊ������,���¼������ID��Ʊ�ݺ�
        lng����ID = GetInvoiceGroupID(2, 1, lng����ID, lngShareUseID, strInvoice, "2")
        If lng����ID <= 0 Then
            Select Case lng����ID
                Case -1
                    MsgBox "Ԥ������[" & strNos & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "��û���㹻�����ú͹��õ�Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "����[" & strNos & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "��û���㹻�ĵĹ���Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox "����[" & strNos & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & strInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                        "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -4
                    MsgBox "����[" & strNos & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & strInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                        "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ�,�ش�õ��ݣ�", vbInformation, gstrSysName
                Case Else
                    MsgBox "Ʊ��������Ϣ����ʧ�ܣ�������������ش򵥾�[" & strNos & "]", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        Do
            '����Ʊ�����ö�ȡ
            blnInput = False
            strInvoice = GetNextBill(lng����ID)
            If strInvoice = "" Then
                '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                strInvoice = UCase(InputBox("�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                "", Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            Else
                strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                strInvoice, Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            End If
            
            '�û�ȡ������,����ӡ
            If strInvoice = "" Then Exit Function
            '���������Ч��
            If blnInput Then
                If GetInvoiceGroupID(2, 1, lng����ID, lngShareUseID, strInvoice, "2") = -3 Then
                    MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                Else
                    blnValid = True
                End If
            Else
                blnValid = True
            End If
        Loop While Not blnValid
    Else
        '�п����ǵ�һ��ʹ��
         Do
             blnInput = False
             '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
             strInvoice = UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, 1131, ""))
             If strInvoice = "" Then
                 strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                 vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                 "", Me.Left + 1500, Me.Top + 1500))
                 blnInput = True
             Else
                 strInvoice = zlCommFun.IncStr(strInvoice)
                 strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                 strInvoice, Me.Left + 1500, Me.Top + 1500))
                 blnInput = True
             End If
                 
             '�û�ȡ������,�����ӡ
             If strInvoice = "" Then
                 If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                 blnValid = True
             Else
                 '���������Ч��
                 If blnInput Then
                     If zlCommFun.ActualLen(strInvoice) <> gbytPrepayLen Then
                         MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytPrepayLen & " λ��", vbInformation, gstrSysName
                     Else
                         blnValid = True
                     End If
                 Else
                     blnValid = True
                 End If
             End If
         Loop While Not blnValid
    End If
    
    'ִ�����ݴ���
    'Zl_����Ԥ����¼_Reprint
    strSql = "Zl_����Ԥ����¼_Reprint("
    '  ���ݺ�_In Varchar2,
    strSql = strSql & "'" & strNos & "',"
    '  Ʊ�ݺ�_In Ʊ��ʹ����ϸ.����%Type,
    strSql = strSql & "'" & strInvoice & "',"
    '  ����id_In Ʊ��ʹ����ϸ.����id%Type,
    strSql = strSql & "" & IIf(lng����ID = 0, "NULL", lng����ID) & ","
    '  ʹ����_In Ʊ��ʹ����ϸ.ʹ����%Type
    strSql = strSql & "'" & UserInfo.���� & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    '���Ʊ��
    intInvoiceFormat = Val(zlDatabase.GetPara(284, glngSys, , "0"))
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, _
        "NO=" & strNos, "�տ�ʱ��=" & Format(strDelDate, "yyyy-mm-dd HH:MM:SS"), _
        "����ID=" & mlngPatient, IIf(intInvoiceFormat = 0, "", "ReportFormat=" & intInvoiceFormat), 2)
    
    '���±���Ʊ��
    If Not gblnPrepayStrict Then
        zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", strInvoice, glngSys, 1131
    End If
    zlPrintInvoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Sub zlCreateObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������¼�����
    '����: �����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-28 16:16:00
    '˵��:
    '����:54894
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '������������
    Err = 0: On Error Resume Next
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
         Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
End Sub

Private Sub zlCloseObject()
    '�ر���ض���
    Err = 0: On Error Resume Next
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
    End If
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
End Sub

Public Function CreateExpenceSvr(ByRef objExpenceSvr As Object, ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݲ�����������
    '���:
    '����:
    '����:
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set objExpenceSvr = CreateObject("zlPublicExpense.clsExpenceSvr")
    If Err <> 0 Then
        MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense.clsExpenceSvr)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
        Exit Function
    End If
    If objExpenceSvr Is Nothing Then Exit Function
    
    If objExpenceSvr.zlInitCommon(glngSys, lngModule, gcnOracle, gstrDBUser) = False Then
        MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense.clsExpenceSvr)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    CreateExpenceSvr = True
End Function
