VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#3.0#0"; "zlIDKind.ocx"
Begin VB.Form frmChargeTurn 
   AutoRedraw      =   -1  'True
   Caption         =   "��(��)�����תסԺ"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   312
   ClientWidth     =   11712
   Icon            =   "frmChargeTurn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   11712
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBill 
      Height          =   2100
      Left            =   90
      ScaleHeight     =   2052
      ScaleWidth      =   10500
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
      ScaleHeight     =   1908
      ScaleWidth      =   3000
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
      ScaleHeight     =   1884
      ScaleWidth      =   5424
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
      ScaleHeight     =   888
      ScaleWidth      =   11712
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
            _ExtentX        =   1101
            _ExtentY        =   614
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
            AllowAutoICCard =   -1  'True
            AllowAutoIDCard =   -1  'True
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
         _ExtentX        =   4720
         _ExtentY        =   614
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
         Format          =   271122435
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   7110
         TabIndex        =   2
         Top             =   90
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   614
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
         Format          =   271122435
         CurrentDate     =   36588
      End
      Begin zlIDKind.IDKindNew IDKindTime 
         Height          =   240
         Left            =   2880
         TabIndex        =   28
         Top             =   120
         Width           =   855
         _ExtentX        =   2350
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
      ScaleHeight     =   432
      ScaleWidth      =   11712
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
      _ExtentX        =   20659
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "frmChargeTurn.frx":0620
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15621
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
      _ExtentX        =   360
      _ExtentY        =   339
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
Private mcllNOs As Collection 'Ҫ���з���ת��ĵ�����Ϣ,��Ա��
                              ' |-cllNO(Collection),��Ա�����ݺ�,Ʊ�ݺ�,����ID,����,��������,�����㵥��,��������ID,������
Private mlng����ID As Long
Private mfrmMain As Object
Private mblnOk As Boolean
Private mbln����ִ�� As Boolean '�Ƿ����ִ��

Private mintPatientRange As Integer
Private mobjPati As clsPatientInfo
Private mstrPrivs As String, mlngModule As Long
Private mbln����תסԺ����� As Boolean
Private mbln�������� As Boolean
Private mblnMultiBalance As Boolean
Private mblnPrepayStrict As Boolean, mbytPrepayLen As Byte

Private Enum mObjPancel
    Pan_Search = 1
    Pan_Bill = 2
    Pan_List = 3
    Pan_Balance = 4
    Pan_Bottom = 5
End Enum
Private mstr�����ʻ� As String

Private mintIDKind As Integer
Private mblnNotClick As Boolean
Private mstrTitle As String
Private mrsFeeList As ADODB.Recordset
Private mobjThirdSwap As clsThreeSwap
Private mblnRefreshData As Boolean

Private mobjExpenceSvr As zlPublicExpense.clsExpenceSvr
Private mobjOneCardComLib As zlOneCardComLib.clsOneCardComLib
Private mblnNewClinicPati As Boolean '�Ƿ�Ϊ�����ﲡ��

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
    Call ShowBills(mlng����ID, dtpBegin.value, dtpEnd.value)
End Sub

Private Sub chkShow_Click()
    If mblnNotClick Then Exit Sub
    Call ShowBills(mlng����ID, dtpBegin.value, dtpEnd.value)
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

Public Function ShowMe(frmMain As Object, ByVal lng����ID As Long, Optional ByVal bln����ִ�� As Boolean, _
    Optional ByVal strPrivs As String, Optional ByVal lngModule As Long, Optional ByRef blnRefreshData As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������תסԺ����
    '���:
    '   bln����ִ��:�Ƿ����ִ�У�����Ƕ���ִ������ύ���ݵ����ݿ⣬������ ExecuteTurn �ӿڵ���ִ��
    '����:
    '   blnRefreshData-�������תסԺ���Ƿ�ˢ������
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbln����ִ�� = bln����ִ��
    mlng����ID = lng����ID: mstrPrivs = strPrivs: mlngModule = lngModule
    mblnRefreshData = False: txtPatient.Tag = lng����ID
    Set mfrmMain = frmMain
    
    mblnOk = False
    On Error Resume Next
    Me.Show vbModal, frmMain
    ShowMe = mblnOk
    blnRefreshData = mblnRefreshData
End Function

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

Private Function IsYBSingle(ByVal strNo As String, Optional blnYBAllDel_Out As Boolean, Optional ByRef blnThirdAllDel_Out As Boolean) As Boolean
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ҽ���Ƿ�ȫ�˻��Ƿֵ��ݾ�", strNo)
    
    blnYBAllDel_Out = rsTmp.EOF
    If rsTmp.EOF Then IsYBSingle = False: Exit Function
    
    blnThirdAllDel_Out = CheckAllTurn(strNo)
    IsYBSingle = Not blnThirdAllDel_Out
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExcuteTranSaveOver(ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, _
    ByRef cllBillPro As Collection, Optional blnNotModify As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ת����ɱ���
    '���:objBalanceInfor-������Ϣ
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
        '    Zl_�������תסԺ_Modify_s
        strSql = "Zl_�������תסԺ_Modify_s("
        '    ��������_In   Number,  '0-������У�Ա�־:ֻ���¹�������ID��У�Ա�־;1-��ͨ�˷ѷ�ʽ:2.�������˷ѽ���:;3-ҽ������;4-���ѿ�����:
        strSql = strSql & "1,"
        '    ����id_In     ����Ԥ����¼.����id%Type,
        strSql = strSql & "" & objBalanceInfor.����ID & ","
        '    ����id_In     ���˽��ʼ�¼.����id%Type,
        strSql = strSql & "" & objPati.����ID & ","
        '  ����_In         ����Ԥ����¼.����%Type,
        strSql = strSql & "'" & objPati.���� & "' ,"
        '  �Ա�_In         ����Ԥ����¼.�Ա�%Type,
        strSql = strSql & "'" & objPati.�Ա� & "' ,"
        '  ����_In         ����Ԥ����¼.����%Type,
        strSql = strSql & "'" & objPati.���� & "' ,"
        '  �����_In       ����Ԥ����¼.�����%Type,
        strSql = strSql & "'" & objPati.����� & "' ,"
        '  סԺ��_In       ����Ԥ����¼.סԺ��%Type,
        strSql = strSql & "'" & objPati.סԺ�� & "' ,"
        '  ���ʽ����_In ����Ԥ����¼.���ʽ����%Type,
        strSql = strSql & "'" & objPati.ҽ�Ƹ��ʽ & "' ,"
        '    ���㷽ʽ_In   Varchar2,
        strSql = strSql & "NULL,"
        '    ����Ա���_In ����Ԥ����¼.����Ա���%Type := Null,
        strSql = strSql & "'" & UserInfo.��� & "' ,"
        '    ����Ա����_In ����Ԥ����¼.����Ա����%Type := Null,
        strSql = strSql & "'" & UserInfo.���� & "' ,"
        '    ����˷�_In   Number := 0,0-δ����˷�;1-����˷�
        strSql = strSql & "1)"
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
    zlExecuteProcedureArrAy cllPro, "����������תסԺ"
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

Public Function ExecuteTurn(ByVal frmMain As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal strסԺ�� As String, ByVal dat��Ժʱ�� As Date, ByVal lng��Ժ����ID As Long, ByVal lng��Ժ����ID As Long, _
    ByRef strErrmsg_Out As String, Optional ByRef blnReflashData_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ĵ��ݺ�����,ִ���������תסԺ����,��ҽ���˷ѽ������
    '���:
    '   lngסԺ��-סԺ��,lng��ҳID-��ҳID,��������������ҽ����Ժ����Ǽ�ʱ�Ŵ���
    '����:
    '   strErrMsg_Out=ʧ��ʱ���ش���ԭ��
    '   blnReflashData_Out=�Ƿ�������ת��
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngcnt As Long, blnҽ�������� As Boolean
    Dim strSql As String, strInvoice As String
    Dim cllPro As Collection, str��ת����ID As String
    Dim intInsure As Integer, blnTurnAll As Boolean
    Dim objBalanceInfor As clsBalanceInfo
    Dim strSfNos As String, blnBillPrintInited As Boolean
    Dim lngStep As Long, bln���ڽ��ʵ� As Boolean
    Dim strNewNo As String, strNewNos As String, varNos As Variant, p As Integer
    Dim strDelDate As String, cllNO As Collection
    '�������ĵ��ݴ���˼·���Ƚ����õ���תΪסԺ���ü�¼���ٵ������������˷�
    Dim strReplenishNo As String, strReplenishNos As String 'Array(�����㵥�ݺ�,ת����SQL,�µ��ݺ�)
    Dim cllReplenishPro As Collection, lngҽ��С��ID As Long
    
    On Error GoTo errHandle
    blnReflashData_Out = False
    If mbln����ִ�� = False And mblnNewClinicPati Then
        ExecuteTurn = frmChargeTurnNew.ExecuteTurn(frmMain, lng����ID, lng��ҳID, _
            strסԺ��, dat��Ժʱ��, lng��Ժ����ID, lng��Ժ����ID, strErrmsg_Out, blnReflashData_Out)
        Exit Function
    End If
    
    If mlng����ID <> lng����ID Then
        strErrmsg_Out = "����ѡ��ת��������������뵱ǰ���˲�ͬ��������ִ���������תסԺ��": Exit Function
    End If
    
    If mcllNOs Is Nothing Then Exit Function
    If mcllNOs.Count = 0 Then ExecuteTurn = True: Exit Function
     
    strDelDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If mbln�������� Then Call zlBillPrint_Initialize(Val("1137-���˽��ʹ���"))
    
    Set mobjThirdSwap = New clsThreeSwap
    If mobjThirdSwap.Init(mobjOneCardComLib, frmMain, mlngModule, _
        mobjPati.����ID, mobjPati.����, mobjPati.�Ա�, mobjPati.����) = False Then Exit Function
    
    Set objBalanceInfor = New clsBalanceInfo
    With objBalanceInfor
        .����ʱ�� = CDate(strDelDate)
        .�������� = 3  '��������:1-�������;2-סԺ����;3-�������תסԺ
    End With
    
    Set cllPro = New Collection
    Set cllReplenishPro = New Collection
    
    zlCommFun.ShowFlash "���ڽ����������תסԺ�������Ժ�...", frmMain
    
    '���ݺ�,Ʊ�ݺ�,����ID,����,��������,�����㵥��,��������ID,������
    lngStep = 0
    i = 1
    Do While i <= mcllNOs.Count
        lngStep = lngStep + 1
        Set cllNO = mcllNOs(i)
        
        'ͬһ����Ʊ�ŵ�һ��ת
        lngcnt = 1
        strInvoice = cllNO("Ʊ�ݺ�")
        If strInvoice <> "" Then
            For j = i + 1 To mcllNOs.Count
                Set cllNO = mcllNOs(j)
                If strInvoice = cllNO("Ʊ�ݺ�") Then
                    lngcnt = lngcnt + 1
                Else
                    Exit For
                End If
            Next
        End If
        
        
        'ҽ��Ҫ������һ�ſ�ʼ��,�����������ǰ����ݺŵ������еģ����Դ˴����򼴿�
        For j = i To i + lngcnt - 1
            Set cllNO = mcllNOs(j)
            blnҽ�������� = False: blnTurnAll = False
            
            strReplenishNo = cllNO("�����㵥��")
            If strReplenishNo = "" Then
                If Val(cllNO("����")) <> 0 Then '���ʵ�������Ϊ0
                    blnҽ�������� = IsYBSingle(cllNO("���ݺ�"))
                Else
                    blnTurnAll = CheckAllTurn(cllNO("���ݺ�"))
                    If InStr("," & str��ת����ID & ",", "," & cllNO("����ID") & ",") > 0 Then blnTurnAll = True
                End If
            End If
            
            With objBalanceInfor
                .����ID = Val(cllNO("����ID"))
                .���ʵ��ݺ� = cllNO("���ݺ�")
            End With
            intInsure = Val(cllNO("����"))
            
            '�ȴ���ļ��ʵ�����ǰ���ݲ��Ǽ��ʵ���˵�����ʵ��Ѵ�����
            If cllNO("��������") <> "���ʵ�" And mbln�������� And Not blnBillPrintInited Then
                Call zlBillPrint_Initialize(Val("1121-�����շѹ���"))
                blnBillPrintInited = True
            End If
    
            lngҽ��С��ID = ZlGetMedicalGroupID(lng����ID, lng��ҳID, cllNO("��������ID"), cllNO("������"), dat��Ժʱ��)
            
            If blnҽ�������� Or (intInsure = 0 And Not blnTurnAll) Or strReplenishNo <> "" Then
                
                If InStr("," & str��ת����ID & ",", "," & cllNO("����ID") & ",") = 0 Then ' ����һ�ν��ʷֵ��ݵģ��Ѿ�ת��������Ҫ�ж�
                    strNewNo = zlDatabase.NextNo(14)
                    
                    'Zl_�������תסԺ_Insert_S
                    strSql = "Zl_�������תסԺ_insert_S("
                    '  No_In         סԺ���ü�¼.NO%Type,
                    strSql = strSql & "'" & cllNO("���ݺ�") & "',"
                    '  Newno_In        סԺ���ü�¼.No%Type,
                    strSql = strSql & "'" & strNewNo & "',"
                    '  סԺ��_In     סԺ���ü�¼.��ʶ��%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                    strSql = strSql & "" & ZVal(strסԺ��) & ","
                    '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                    strSql = strSql & "" & ZVal(lng��ҳID) & ","
                    '  ��Ժʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                    strSql = strSql & "To_Date('" & Format(dat��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                    '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                    strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                    '  ��Ժ����id_In סԺ���ü�¼.���˲���id%Type,
                    strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                    '  ת��ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ���ת��ʱ,ÿ�ŵ��ݵ�ת��ʱ����ͬ,����ϵͳ��ǰʱ��
                    strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'),"
                    '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                    strSql = strSql & "'" & UserInfo.���� & "',"
                    '  ҽ��С��id_In   סԺ���ü�¼.ҽ��С��id%Type,
                    strSql = strSql & "" & ZVal(lngҽ��С��ID) & ","
                    '  ����_In         סԺ���ü�¼.����%Type,
                    strSql = strSql & "'" & mobjPati.���� & "',"
                    '  ��������_In Number := 1, --1-�����շѵ�;2-������ʵ�
                    strSql = strSql & "" & IIf(cllNO("��������") = "���ʵ�", 2, 1) & ")"
                    
                    If strReplenishNo <> "" And mbln�������� Then
                        If InStr(strReplenishNos & ";", ";" & strReplenishNo & "," & cllNO("����") & ";") = 0 Then
                            strReplenishNos = strReplenishNos & ";" & strReplenishNo & "," & cllNO("����")
                        End If
                        'Array(�����㵥�ݺ�,ת����SQL,�µ��ݺ�)
                        cllReplenishPro.Add Array(strReplenishNo, strSql, strNewNo)
                    Else
                        zlAddArray cllPro, strSql
                        If cllNO("��������") = "���ʵ�" And mbln�������� Then
                            'Zl_����תסԺ_����ת��
                            strSql = "Zl_����תסԺ_����ת��("
                            '  No_In         סԺ���ü�¼.No%Type,
                            strSql = strSql & "'" & cllNO("���ݺ�") & "',"
                            '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                            strSql = strSql & "'" & UserInfo.��� & "',"
                            '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                            strSql = strSql & "'" & UserInfo.���� & "',"
                            '  ����ʱ��_In   סԺ���ü�¼.����ʱ��%Type
                            strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'))"
                            zlAddArray cllPro, strSql
                            
                            If DelBalaceMz(mobjPati, cllPro, lng��ҳID, lng��Ժ����ID, objBalanceInfor) = False Then
                                blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                                Exit Function
                            End If
                            bln���ڽ��ʵ� = True
                        ElseIf mbln�������� And cllNO("��������") <> "���ʵ�" Then
                            strSfNos = "'" & cllNO("���ݺ�") & "'"
                            If zlBillPrint_EraseBill(strSfNos, 0) = False Then Exit Function
                            
                            With objBalanceInfor
                                .�������� = 3 '��������:1-�������;2-סԺ����;3-�������תסԺ
                                .����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
                            End With
                            
                            'Zl_����תסԺ_�շ�ת��_s
                            strSql = "Zl_����תסԺ_�շ�ת��_s("
                            '  ����id_In       ���˽��ʼ�¼.����id%Type,
                            strSql = strSql & "" & mobjPati.����ID & ","
                            '  ����_In         ����Ԥ����¼.����%Type,
                            strSql = strSql & "'" & mobjPati.���� & "' ,"
                            '  �Ա�_In         ����Ԥ����¼.�Ա�%Type,
                            strSql = strSql & "'" & mobjPati.�Ա� & "' ,"
                            '  ����_In         ����Ԥ����¼.����%Type,
                            strSql = strSql & "'" & mobjPati.���� & "' ,"
                            '  �����_In       ����Ԥ����¼.�����%Type,
                            strSql = strSql & "'" & mobjPati.����� & "' ,"
                            '  סԺ��_In       ����Ԥ����¼.סԺ��%Type,
                            strSql = strSql & "'" & mobjPati.סԺ�� & "' ,"
                            '  ���ʽ����_In ����Ԥ����¼.���ʽ����%Type,
                            strSql = strSql & "'" & mobjPati.ҽ�Ƹ��ʽ & "' ,"
                            '  No_In         סԺ���ü�¼.No%Type,
                            strSql = strSql & "'" & cllNO("���ݺ�") & "',"
                            '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                            strSql = strSql & "'" & UserInfo.��� & "',"
                            '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                            strSql = strSql & "'" & UserInfo.���� & "',"
                            '  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                            strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'),"
                            '  ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
                            strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                            '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null,
                            strSql = strSql & "" & ZVal(lng��ҳID) & ","
                            '  ����id_In     ����Ԥ����¼.����id%Type := Null,
                            strSql = strSql & "" & objBalanceInfor.����ID & ")"
                            '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null
                            zlAddArray cllPro, strSql
                            
                             'ִ��ҽ��:
                            If ExcuteInsureDel(objBalanceInfor, intInsure, objBalanceInfor.���ʵ��ݺ�, cllPro) = False Then
                                blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                                Exit Function
                            End If
                            'ִ��һ��ͨ
                            If Not ExecuteThirdReturnMoneySwap(mobjPati, objBalanceInfor, cllPro) Then
                                blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                                Exit Function
                            End If
                            '���
                            If ExcuteTranSaveOver(mobjPati, objBalanceInfor, cllPro) = False Then
                                blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                                Exit Function
                            End If
                        Else
                            'ֱ���������תסԺ
                            If Not ExcuteTranSaveOver(mobjPati, objBalanceInfor, cllPro, True) Then Exit Function
                        End If
                        
                        Call mobjExpenceSvr.zlAdjustFeeData(strNewNo)
                    End If
                End If
            Else
                If InStr("," & str��ת����ID & ",", "," & cllNO("����ID") & ",") = 0 Then
                    If cllNO("��������") = "���ʵ�" Then
                        varNos = Array(cllNO("���ݺ�"))
                    Else '�շѵ���һ��ת����������е���
                        strSfNos = GetBalanceNos(1, cllNO("����ID"))
                        varNos = Split(strSfNos, ",")
                    End If
                    
                    strNewNos = ""
                    For p = 0 To UBound(varNos)
                        strNewNo = zlDatabase.NextNo(14)
                        strNewNos = strNewNos & "," & strNewNo
                        
                        'Zl_�������תסԺ_Insert_S
                        strSql = "Zl_�������תסԺ_insert_S("
                        '  No_In         סԺ���ü�¼.NO%Type,
                        strSql = strSql & "'" & varNos(p) & "',"
                        '  Newno_In        סԺ���ü�¼.No%Type,
                        strSql = strSql & "'" & strNewNo & "',"
                        '  סԺ��_In     סԺ���ü�¼.��ʶ��%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                        strSql = strSql & "" & ZVal(strסԺ��) & ","
                        '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                        strSql = strSql & "" & ZVal(lng��ҳID) & ","
                        '  ��Ժʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                        strSql = strSql & "To_Date('" & Format(dat��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                        strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                        '  ��Ժ����id_In סԺ���ü�¼.���˲���id%Type,
                        strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                        '  ת��ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ���ת��ʱ,ÿ�ŵ��ݵ�ת��ʱ����ͬ,����ϵͳ��ǰʱ��
                        strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'),"
                        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSql = strSql & "'" & UserInfo.���� & "',"
                        '  ҽ��С��id_In   סԺ���ü�¼.ҽ��С��id%Type,
                        strSql = strSql & "" & ZVal(lngҽ��С��ID) & ","
                        '  ����_In         סԺ���ü�¼.����%Type,
                        strSql = strSql & "'" & mobjPati.���� & "',"
                        '  ��������_In Number := 1, --1-�����շѵ�;2-������ʵ�
                        strSql = strSql & "" & IIf(cllNO("��������") = "���ʵ�", 2, 1) & ")"
                        zlAddArray cllPro, strSql
                    Next
                    If strNewNos <> "" Then strNewNos = Mid(strNewNos, 2)
                    
                    If cllNO("��������") = "���ʵ�" And mbln�������� Then
                        'Zl_����תסԺ_����ת��
                        strSql = "Zl_����תסԺ_����ת��("
                        '  No_In         סԺ���ü�¼.No%Type,
                        strSql = strSql & "'" & cllNO("���ݺ�") & "',"
                        '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                        strSql = strSql & "'" & UserInfo.��� & "',"
                        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSql = strSql & "'" & UserInfo.���� & "',"
                        '  ����ʱ��_In   סԺ���ü�¼.����ʱ��%Type
                        strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'))"
                        zlAddArray cllPro, strSql
                        
                        If DelBalaceMz(mobjPati, cllPro, lng��ҳID, lng��Ժ����ID, objBalanceInfor) = False Then
                            blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                            Exit Function
                        End If
                        bln���ڽ��ʵ� = True
                    ElseIf mbln�������� And cllNO("��������") <> "���ʵ�" Then
                        strSfNos = "'" & Replace(strSfNos, ",", "','") & "'"
                        If zlBillPrint_EraseBill(strSfNos, 0) = False Then Exit Function
                        
                        With objBalanceInfor
                            .�������� = 3 '��������:1-�������;2-סԺ����;3-�������תסԺ
                            .����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
                        End With
                        
                        'Zl_����תסԺ_�շ�ת��_s
                        strSql = "Zl_����תסԺ_�շ�ת��_s("
                        '  ����id_In       ���˽��ʼ�¼.����id%Type,
                        strSql = strSql & "" & mobjPati.����ID & ","
                        '  ����_In         ����Ԥ����¼.����%Type,
                        strSql = strSql & "'" & mobjPati.���� & "' ,"
                        '  �Ա�_In         ����Ԥ����¼.�Ա�%Type,
                        strSql = strSql & "'" & mobjPati.�Ա� & "' ,"
                        '  ����_In         ����Ԥ����¼.����%Type,
                        strSql = strSql & "'" & mobjPati.���� & "' ,"
                        '  �����_In       ����Ԥ����¼.�����%Type,
                        strSql = strSql & "'" & mobjPati.����� & "' ,"
                        '  סԺ��_In       ����Ԥ����¼.סԺ��%Type,
                        strSql = strSql & "'" & mobjPati.סԺ�� & "' ,"
                        '  ���ʽ����_In ����Ԥ����¼.���ʽ����%Type,
                        strSql = strSql & "'" & mobjPati.ҽ�Ƹ��ʽ & "' ,"
                        '  No_In         סԺ���ü�¼.No%Type,
                        strSql = strSql & "'" & cllNO("���ݺ�") & "',"
                        '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                        strSql = strSql & "'" & UserInfo.��� & "',"
                        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSql = strSql & "'" & UserInfo.���� & "',"
                        '  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                        strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'),"
                        '  ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
                        strSql = strSql & "" & ZVal(lng��Ժ����ID) & ","
                        '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null,
                        strSql = strSql & "" & ZVal(lng��ҳID) & ","
                        '  ����id_In     ����Ԥ����¼.����id%Type := Null,
                        strSql = strSql & "" & objBalanceInfor.����ID & ","
                        '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
                        strSql = strSql & "" & objBalanceInfor.����ID & ")"
                        zlAddArray cllPro, strSql
                        
                         'ִ��ҽ��:
                        If ExcuteInsureDel(objBalanceInfor, intInsure, "", cllPro) = False Then
                            blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                            Exit Function
                        End If
                        'ִ��һ��ͨ
                        If Not ExecuteThirdReturnMoneySwap(mobjPati, objBalanceInfor, cllPro) Then
                            blnReflashData_Out = objBalanceInfor.�Ƿ񱣴���ʵ�
                            Exit Function
                        End If
                        '���
                        If ExcuteTranSaveOver(mobjPati, objBalanceInfor, cllPro) = False Then Exit Function
                    Else
                        'ֱ���������תסԺ
                        If Not ExcuteTranSaveOver(mobjPati, objBalanceInfor, cllPro, True) Then Exit Function
                    End If
                    
                    Call mobjExpenceSvr.zlAdjustFeeData(strNewNos)
                End If
                str��ת����ID = str��ת����ID & "," & cllNO("����ID")
            End If
        Next
        i = i + lngcnt
    Loop
    
    sta.Panels(2).Text = ""
    
    '�Բ�����㵥�ݽ����˷Ѵ���
    If strReplenishNos <> "" Then
        strReplenishNos = Mid(strReplenishNos, 2)
        If ExecuteReplenishDel(strReplenishNos, cllReplenishPro, lng��ҳID, lng��Ժ����ID, strDelDate) = False Then
            Exit Function
        End If
    End If
    
    '��ӡԤ�����
    Call PrintPrePayPrint(strDelDate)
    
    '��ʾ���ʴ���
    If bln���ڽ��ʵ� And mbln�������� Then
       Call ShowBalanceWindows(frmMain, strDelDate)
    End If
    
    ExecuteTurn = True
    Exit Function
errHandle:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteReplenishDel(ByVal strNOs As String, ByVal cllPro As Collection, _
    ByVal lng��ҳID As Long, ByVal lng��Ժ����ID As Long, ByVal strDelDate As String) As Boolean
    '����:�Բ������ĵ��ݽ���ת���ü��˷Ѵ���
    '���:
    '   strNos �����㵥��,��ʽ�����ݺ�,����;...
    '   cllPro ������˷ѹ��̵ļ��ϣ�Array(�����㵥�ݺ�,ת����SQL,�µ��ݺ�)
    '   strDelDate �˷�ʱ��
    Dim strSql As String, strNoTemp As String
    Dim varNos As Variant, i As Long, p As Long, blnTrans As Boolean
    Dim strNo As String, intInsure As Integer
    Dim lng�������ID  As Long, lng���ó���ID As Long, lng������� As Long
    Dim lngԭ����ID As Long, strAdvance As String
    Dim strNewNos As String, strNewNo As String
    
    Err = 0: On Error GoTo errH
    If strNOs = "" Then ExecuteReplenishDel = True: Exit Function
    
    Call zlBillPrint_Initialize(Val("1124-���ղ������"))
    varNos = Split(strNOs, ";")
    For i = 0 To UBound(varNos)
        '���ݺ�,����;...
        strNo = Split(varNos(i), ",")(0): intInsure = Split(varNos(i), ",")(1)
        
        If zlBillPrint_EraseBill(strNo, 0) = False Then Exit Function
        
        lng���ó���ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        lng�������ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        lng������� = -1 * lng���ó���ID
        
        gcnOracle.BeginTrans: blnTrans = True
        For p = 1 To cllPro.Count
            'Array(�����㵥�ݺ�,ת����SQL,�µ��ݺ�)
            strNoTemp = cllPro(p)(0): strSql = cllPro(p)(1): strNewNo = cllPro(p)(2)
            If strNoTemp = strNo Then
                strNewNos = strNewNos & "," & strNewNo
                zlDatabase.ExecuteProcedure strSql, "ִ�в���������"
            End If
        Next
        If strNewNos <> "" Then strNewNos = Mid(strNewNos, 2)
        
        'Zl_����תסԺ_������ת��_s(
        strSql = "Zl_����תסԺ_������ת��_s("
        '  No_In         ���ò����¼.No%Type,
        strSql = strSql & "'" & strNo & "',"
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
        zlDatabase.ExecuteProcedure strSql, "ִ�в�����ת��"
        
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
        lngԭ����ID = GetFromNOToLastBalanceID(strNo, , , , True)
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

Private Function GetFromNOToLastBalanceID(ByVal strNOs As String, _
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

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���ݵ��ݻ�ȡ���һ���������ʵĽ���ID", strNOs)

    If rsTemp.EOF Then Exit Function

    lng������� = Val(Nvl(rsTemp!�������))
    GetFromNOToLastBalanceID = Val(Nvl(rsTemp!����ID))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExcuteInsureDel(ByVal objBalanceInfor As clsBalanceInfo, _
    ByVal intInsure As Integer, ByVal strNo As String, ByRef cllBillPro As Collection) As Boolean
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
    zlExecuteProcedureArrAy cllPro, "ִ��ҽ������", True
    
    strAdvance = objBalanceInfor.����ID & "|0" & IIf(strNo <> "", "|" & strNo, "")
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

Private Function GetYBBalance(ByVal lng����ID As Long, ByVal lng����ID As Long, _
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
                If gclsInsure.GetCapability(support�����������, lng����ID, intInsure, Nvl(rsData!���㷽ʽ)) Then
                    str���㷽ʽ = str���㷽ʽ & "||" & Nvl(rsData!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(rsData!��Ԥ��))
                End If
            Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                If Nvl(rsData!���㷽ʽ) <> str�����ʻ� Then
                    str���㷽ʽ = str���㷽ʽ & "||" & Nvl(rsData!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(rsData!��Ԥ��))
                End If
            End If
        Else
            str���㷽ʽ = str���㷽ʽ & "||" & Nvl(rsData!���㷽ʽ) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(rsData!��Ԥ��))
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
    zlExecuteProcedureArrAy cllPro, "ִ��ҽ������", True
          
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
            zlDatabase.ExecuteProcedure strSql, "У��ҽ������"
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

Private Function ExecuteThirdReturnMoneySwap_JZ(objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, _
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
    zlExecuteProcedureArrAy cllPro, "ִ�з���ת��", True
    
    Set cllPro = New Collection
    
    strSql = _
        " Select �����id, ���㷽ʽ, ��Ԥ�� As �����ܶ�, ��Ԥ��, ������ˮ��, ����˵��," & _
        "        ����, ��������id, �������, ժҪ, �տ�ʱ��" & _
        " From ����Ԥ����¼ A" & _
        " Where ��¼���� = 12 And a.����id = [1] And a.�����ID Is Not Null And a.У�Ա�־ = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯ����������", objBalanceInfor.����ID)
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
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, "��ѯ��������", objBalanceInfor.����ID)
    
    Set objThirdDelItems = New clsBalanceItems
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lng�����ID = Val(Nvl(rsTemp!�����ID))
            lng��������ID = Val(Nvl(rsTemp!��������ID))
            
            lngԭ����ID = 0
            rsBalance.Filter = "�����ID=" & lng�����ID & " and ��������ID=" & lng��������ID
            If Not rsBalance.EOF Then lngԭ����ID = Val(Nvl(rsBalance!����ID))
            If lngԭ����ID = 0 Then
                rsBalance.Filter = "�����ID=" & lng�����ID & " and ������ˮ��='" & Nvl(!������ˮ��) & "'"
                If Not rsBalance.EOF Then lngԭ����ID = Val(Nvl(rsBalance!����ID))
                If lngԭ����ID = 0 Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    MsgBox Nvl(rsTemp!���㷽ʽ) & "δ�ҵ�ԭʼ�����¼ ������!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            objBalanceInfor.����ID = lngԭ����ID
            
            Set objItem = New clsBalanceItem
            With objItem
                Set .objCard = mobjThirdSwap.zlGetCardFromCardType(lng�����ID, False, Nvl(rsTemp!���㷽ʽ))
                .����ID = objBalanceInfor.����ID
                .����IDs = lngԭ����ID
                .����ID = lngԭ����ID
                .��������ID = lng��������ID
                .������ˮ�� = Nvl(rsTemp!������ˮ��)
                .����˵�� = Nvl(rsTemp!����˵��)
                .���㷽ʽ = Nvl(rsTemp!���㷽ʽ)
                .������� = Nvl(rsTemp!�������)
                .����ժҪ = Nvl(rsTemp!ժҪ)
                .������ = Val(Nvl(rsTemp!��Ԥ��))
                .�������� = 3  '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                .�������� = .objCard.��������
                .����ʱ�� = Format(rsTemp!�տ�ʱ��, "yyyy-mm-dd HH:MM:SS")
                .���� = Nvl(rsTemp!����)
                .�����ID = lng�����ID
                .ʣ���� = Val(Nvl(rsTemp!��Ԥ��))
                .δ�˽�� = Val(Nvl(rsTemp!��Ԥ��))
                .ԭʼ��� = Val(Nvl(rsTemp!��Ԥ��))
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
                Set objItemTemp = objItem.Clone()
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
        If mobjThirdSwap.zlThird_ReturnMoney_IsValied(objItem.objCard, 2, objBalanceInfor, objItem.objTag, objItems, False) = False Then
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

Private Function DelBalaceMz(objPati As clsPatientInfo, cllBillPro As Collection, _
    ByVal lng��ҳID As Long, ByVal lng��Ժ����ID As Long, ByRef objBalanceInfor As clsBalanceInfo) As Boolean
    '����:���˵������ͽ�������
    Dim strSql As String, rsData As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim intInsure As Integer
    Dim lng����ID As Long, strNo As String, lng����ID As Long
    Dim strBalanceIDs As String, strBalanceNos As String
    
    On Error GoTo ErrHandler
    strSql = _
        " Select /*+cardinality(j,10)*/ Distinct b.Id As ����ID, b.No, c.����, b.����ID" & _
        " From ������ü�¼ A, ���˽��ʼ�¼ B, ���ս����¼ C" & _
        " Where a.����id = b.Id And a.��¼���� In (2, 12) And a.No = [1] And b.��¼״̬ = 1" & _
        "       And b.ID=c.��¼id(+) And c.����(+) = 1 And c.�����id(+) Is Null" & _
        " Order By No"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ��������", objBalanceInfor.���ʵ��ݺ�)
    If rsData.EOF Then
        'δ���ˣ�����ת�����
        blnTrans = True
        zlExecuteProcedureArrAy cllBillPro, "ִ�н��ʷ���ת��"
        blnTrans = False
        
        objBalanceInfor.�Ƿ񱣴���ʵ� = True
        Set cllBillPro = New Collection
        DelBalaceMz = True
        Exit Function
    End If
    
    Do While Not rsData.EOF
        strBalanceIDs = strBalanceIDs & "," & Nvl(rsData!����ID)
        strBalanceNos = strBalanceNos & "," & Nvl(rsData!NO)
        rsData.MoveNext
    Loop
    
    If rsData.RecordCount > 0 Then rsData.MoveFirst
    Do While Not rsData.EOF
        With objBalanceInfor
            .�������� = 1  '��������:1-�������;2-סԺ����;3-�������תסԺ
            .����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        End With
        
        lng����ID = Val(Nvl(rsData!����ID))
        strNo = Nvl(rsData!NO)
        lng����ID = Val(Nvl(rsData!����ID))
        intInsure = Val(Nvl(rsData!����))
        
        If zlBillPrint_EraseBill("", lng����ID) = False Then Exit Function
        
        'Zl_���˽��ʼ�¼_Cancel
        strSql = "Zl_���˽��ʼ�¼_Cancel("
        '  No_In         ���˽��ʼ�¼.No%Type,
        strSql = strSql & "'" & strNo & "',"
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
        strSql = strSql & "'" & strNo & "',"
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
        strSql = strSql & "'" & Nvl(rsData!NO) & "',"
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
        zlExecuteProcedureArrAy cllBillPro, "ִ�н�������"
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

Private Function ShowBalanceWindows(frmMain As Object, ByVal strDelDate As String) As Boolean
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
    If mlng����ID <> 0 Then
        lng����ID = mlng����ID
    ElseIf Not mobjPati Is Nothing Then
        lng����ID = mobjPati.����ID
    End If
    
    'zlPatiBalance(ByVal frmMain As Object, _
    '    ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, strDBUser As String, _
    '    ByVal lng����ID As Long, ByVal lng��ҳID As   long ) as boolean
    If objInExse.zlPatiBalance(frmMain, gcnOracle, glngSys, gstrDBUser, lng����ID, 0, strDelDate) = False Then
        '���ý���
    End If
    ShowBalanceWindows = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowBills(ByVal lngPatient As Long, ByVal DatBegin As Date, ByVal datEnd As Date, _
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
        If DatBegin > datEnd Then
            DatTmp = datEnd
            datEnd = DatBegin
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
            "     Where J1.No = a.No And J1.��� = a.��� And J1.��¼���� In (2,12) And J1.����id = J2.Id And Nvl(J2.����״̬,0)=1)"
        
        If mbln����תסԺ����� Then
           strWhere = " And A.����id = [1] "
        Else
            If datEnd - DatBegin < 4 Then   '36170
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
            "             Where M.ID=J.����ID And M.����ID = [1] and M.NO=A.NO And M.��� = A.��� And Mod(M.��¼����,10)=Mod(A.��¼����,10)  " & _
            "                   And J.������� is Not NULL and  nvl(J.��¼״̬,0)=0 and J.����=1) " & vbNewLine
        Else
            strVerifyWhere = _
            " And Not Exists (Select 1 From ������ü�¼ M,������˼�¼ J " & _
            "                 Where M.ID=J.����ID And M.����ID = [1] and M.NO=A.NO And M.��� = A.��� And Mod(M.��¼����,10)=Mod(A.��¼����,10) " & _
            "                       And J.������� is Not NULL and  nvl(J.��¼״̬,0) > 0 and J.����=1)"
        End If
        
        strSql = strSql & _
            " Select x.ѡ��, x.���, x.����, Max(Decode(Nvl(z.����, 0),0,'','��')) As ҽ��,Max(z.�����ID) As һ��ͨҽ��," & _
            "       x.No As ���ݺ�, x.Ʊ�ݺ�," & vbNewLine & _
            "       x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Max(y.����id) As ����id," & vbNewLine & _
            "       Max(Decode(z.�����ID,NULL,Nvl(z.����,0),0)) As ����" & vbNewLine & _
            " From ( Select  '��' As ѡ��, '��ת��' As ���, '�շѵ�' As ����, a.No," & vbNewLine & _
            "               a.ʵ��Ʊ�� As Ʊ�ݺ�, a.������, a.��������ID, Sum(a.Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
            "               Sum(a.ʵ�ս��) As ʵ�ս��, To_Char(Max(a.����ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & vbNewLine & _
            "        From ������ü�¼ A" & vbNewLine & _
            "        Where Mod(a.��¼����, 10) = 1 And nvl(a.����״̬,0)<>1 And a.��¼״̬ <> 0 " & strWhere & " " & strVerifyWhere & vbCrLf & strErrWhere & _
            "              And Exists (Select 1 From ������ü�¼ K" & _
            "                          Where k.No = a.No And k.��� = a.��� And k.����id = [1] And Mod(k.��¼����, 10) = Mod(a.��¼����, 10)" & _
            "                                And Nvl(k.���ӱ�־, 0) <> 9" & _
            "                          Group By k.��� Having Sum(k.ʵ�ս��) <> 0)" & vbNewLine & _
            "        Group By a.No, a.ʵ��Ʊ��, a.������, a.��������ID" & _
            "      ) X, ������ü�¼ Y," & vbNewLine & _
            "      ( Select Distinct a.��¼id, a.����,a.�����ID" & vbNewLine & _
            "        From ���ս����¼ A" & vbNewLine & _
            "        Where a.���� = 1 And a.����id = [1]) Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.��¼����, 10) = 1 And y.��¼״̬ In (1, 3) And y.����ID = [1]" & _
            "        And y.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = x.No And Mod(��¼����, 10) = 1 And ����ID = [1] And ��¼״̬ In (1, 3)) And y.����id = z.��¼id(+)" & _
            " Group By x.ѡ��, x.���, x.����, x.No, x.Ʊ�ݺ�, x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ�� "
 
        strSql = strSql & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select x.ѡ��, x.���, x.����, Max(Decode(Nvl(z.����, 0),0,'','��')) As ҽ��,Max(z.�����ID) As һ��ͨҽ��," & _
            "       x.No As ���ݺ�, x.Ʊ�ݺ�," & vbNewLine & _
            "       x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Max(y.����id) As ����id," & vbNewLine & _
            "       Max(Decode(z.�����ID,NULL,Nvl(z.����,0),0)) As ����" & vbNewLine & _
            " From ( " & _
            "       Select " & vbNewLine & _
            "           '' As ѡ��, '����ת��' As ���, '�շѵ�' As ����, a.No," & vbNewLine & _
            "           a.ʵ��Ʊ�� As Ʊ�ݺ�, a.������, a.��������ID, Sum(a.Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
            "           Sum(a.ʵ�ս��) As ʵ�ս��, To_Char(Max(a.����ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & vbNewLine & _
            "       From ������ü�¼ A" & vbNewLine & _
            "       Where Mod(a.��¼����, 10) = 1 And nvl(a.����״̬,0)<>1 And a.��¼״̬ = 3 " & strWhere & " And Nvl(a.���ӱ�־, 0) <> 9 " & vbCrLf & strErrWhere & _
            "           And Not Exists (Select 1 From ������ü�¼ K  Where k.No = a.No And k.��� = a.��� And k.����id = [1] And Mod(k.��¼����, 10) = Mod(a.��¼����, 10) And Nvl(k.���ӱ�־, 0) <> 9 Group By k.���  Having Sum(k.ʵ�ս��) <> 0)" & vbNewLine & _
            "       Group By a.No, a.ʵ��Ʊ��, a.������, a.��������ID " & _
            "       ) X, ������ü�¼ Y," & vbNewLine & _
            "       (Select Distinct a.��¼id, a.����,a.�����ID" & vbNewLine & _
            "        From ���ս����¼ A" & vbNewLine & _
            "        Where a.���� = 1 And a.����id = [1]) Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.��¼����, 10) = 1 And y.��¼״̬ In (1, 3) And y.����ID = [1]" & _
            "       And y.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = x.No And Mod(��¼����, 10) = 1 And ����ID = [1] And ��¼״̬ In (1, 3)) And y.����id = z.��¼id(+)" & _
            " Group By x.ѡ��, x.���, x.����, x.No, x.Ʊ�ݺ�, x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��"

            
        strSql = strSql & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select x.ѡ��, x.���, x.����, Max(Decode(Nvl(z.����, 0),0,'','��')) As ҽ��,Max(z.�����ID) As һ��ͨҽ��," & _
            "       x.No As ���ݺ�, x.Ʊ�ݺ�," & vbNewLine & _
            "       x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Max(y.����id) As ����id," & vbNewLine & _
            "       Max(Decode(z.�����ID,NULL,Nvl(z.����,0),0)) As ����" & vbNewLine & _
            "From (Select " & vbNewLine & _
            "        '' As ѡ��, '����ת��' As ���, '�շѵ�' As ����, a.No," & vbNewLine & _
            "        a.ʵ��Ʊ�� As Ʊ�ݺ�, a.������, a.��������ID, Sum(a.Ӧ�ս��) As Ӧ�ս��," & vbNewLine & _
            "        Sum(a.ʵ�ս��) As ʵ�ս��, To_Char(Max(a.����ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & vbNewLine & _
            "       From ������ü�¼ A" & vbNewLine & _
            "       Where Mod(a.��¼����, 10) = 1 And nvl(a.����״̬,0)<>1 And a.��¼״̬ <> 0 " & strWhere & " " & vbCrLf & strErrWhere & _
            "           And Exists (Select 1 From ������ü�¼ M,������˼�¼ J Where M.ID=J.����ID And M.����ID = [1] and M.NO=A.NO And M.��� = A.��� And Mod(M.��¼����,10)=Mod(A.��¼����,10) And J.������� is Not NULL and  nvl(J.��¼״̬,0) = 1 and J.����=1)" & _
            "           And Exists��(Select 1�� From ������ü�¼ K��Where k.No = a.No And K.��� = a.��� And k.����id = [1] And Mod(k.��¼����, 10) = Mod(a.��¼����, 10) And Nvl(k.���ӱ�־, 0) <> 9��Group By k.��š�Having Sum(k.ʵ�ս��) <> 0)" & vbNewLine & _
            "       Group By a.No, a.ʵ��Ʊ��, a.������, a.��������ID) X, ������ü�¼ Y," & vbNewLine & _
            "     (  Select Distinct a.��¼id, a.����,a.�����ID" & vbNewLine & _
            "        From ���ս����¼ A" & vbNewLine & _
            "        Where a.���� = 1 And a.����id = [1]) Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.��¼����, 10) = 1 And y.��¼״̬ In (1, 3) And y.����ID = [1]" & _
            " And y.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = x.No And Mod(��¼����, 10) = 1 And ����ID = [1] And ��¼״̬ In (1, 3)) And y.����id = z.��¼id(+)" & _
            " Group By x.ѡ��, x.���, x.����, x.No, x.Ʊ�ݺ�, x.������, x.��������ID, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��"
     
        strSql = strSql & " UNION ALL " & _
                " Select    '��' as ѡ��,'��ת��' as ���,'���ʵ�' as ����,'' as ҽ��,0 As һ��ͨҽ��," & _
                "       A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������, a.��������ID," & vbNewLine & _
                "       Sum(A.Ӧ�ս��) As Ӧ�ս��, Sum(A.ʵ�ս��) As ʵ�ս��," & vbNewLine & _
                "       To_Char(Max(A.����ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, 0 as ����ID,0 as ����" & vbNewLine & _
                " From ������ü�¼ A" & vbNewLine & _
                " Where A.��¼���� =2 And A.��¼״̬ <> 0 " & strWhere & strBalanceErrWhere & vbNewLine & _
                "       And Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��� = A.��� And K.��¼����=A.��¼���� And Nvl(k.���ӱ�־, 0) <> 9 Group By K.��� Having Sum(K.����) <> 0) " & vbNewLine & _
                        strVerifyWhere & _
                "Group By A.NO, A.ʵ��Ʊ��, A.������, a.��������ID "
             
        strSql = strSql & " UNION ALL " & _
            " Select C.ѡ��,C.���,C.����,C.ҽ��,c.һ��ͨҽ��,C.���ݺ�, C.Ʊ�ݺ�, C.������, c.��������ID," & vbNewLine & _
            "       Sum(D.Ӧ�ս��) As Ӧ�ս��, Sum(D.ʵ�ս��) As ʵ�ս��, C.����ʱ��, C.����ID, C.����" & vbNewLine & _
            " From " & _
            " (Select    '' as ѡ��,'����ת��' as ���,'���ʵ�' as ����,'' as ҽ��,0 As һ��ͨҽ��," & _
            "       A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������, a.��������ID," & vbNewLine & _
            "       Sum(A.Ӧ�ս��) As Ӧ�ս��, Sum(A.ʵ�ս��) As ʵ�ս��," & vbNewLine & _
            "       To_Char(Max(A.����ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��,0 as ����ID,0 as ����" & vbNewLine & _
            "   From ������ü�¼  A" & vbNewLine & _
            "   Where A.��¼���� = 2 And A.��¼״̬ In (2,3)" & strWhere & strBalanceErrWhere & vbNewLine & _
            "       And Not Exists (Select 1 From ������ü�¼ Where NO=A.NO And ��� = A.��� And ��¼״̬=1 And ��¼����=2) " & vbNewLine & _
            "       And Not Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��� = A.��� And K.��¼����=A.��¼���� And Nvl(k.���ӱ�־, 0) <> 9 Group By K.��� Having Sum(K.ʵ�ս��) <> 0) " & vbNewLine & _
            "   Group By A.NO, A.ʵ��Ʊ��, A.������, a.��������ID" & _
            "   Having Sum(A.ʵ�ս��)=0) C,������ü�¼ D Where C.���ݺ�=D.NO And D.��¼����=2 And D.��¼״̬=3" & vbNewLine & _
            " Group By C.ѡ��,C.���,C.����,C.ҽ��,C.���ݺ�, C.Ʊ�ݺ�, C.������, c.��������ID,C.����ʱ��, C.����ID, C.���� "
            
        strSql = strSql & " UNION ALL " & _
            " Select    '' as ѡ��,'����ת��' as ���,'���ʵ�' as ����,'' as ҽ��,0 As һ��ͨҽ��, " & _
            "       A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������, a.��������ID," & vbNewLine & _
            "       Sum(A.Ӧ�ս��) As Ӧ�ս��, Sum(A.ʵ�ս��) As ʵ�ս��," & vbNewLine & _
            "       To_Char(Max(A.����ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, 0 as ����ID,0 as ����" & vbNewLine & _
            " From ������ü�¼ A" & vbNewLine & _
            " Where A.��¼���� = 2 And A.��¼״̬ <> 0 " & strWhere & strBalanceErrWhere & vbNewLine & _
            "       And Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��� = A.��� And K.��¼����=A.��¼���� And Nvl(k.���ӱ�־, 0) <> 9 Group By K.��� Having Sum(K.����) <> 0) " & vbNewLine & _
            " And  Exists (Select 1 From ������ü�¼ M,������˼�¼ J where M.ID=J.����ID and M.NO=A.NO And M.��� = A.��� And M.��¼����=A.��¼���� And J.������� is Not NULL and  nvl(J.��¼״̬,0) = 1 and J.����=1) " & _
            "Group By A.NO, A.ʵ��Ʊ��, A.������, a.��������ID "
        
        strSql = _
            " Select ѡ��, ���, ����, ҽ��, һ��ͨҽ��, ���ݺ�, Ʊ�ݺ�, ������, b.���� As ��������,a.��������ID As ��������ID, " & _
            "        To_Char(Ӧ�ս��, '" & gSysPara.Money_Decimal.strFormt_ORA & "') As Ӧ�ս��," & _
            "        To_Char(ʵ�ս��, '" & gSysPara.Money_Decimal.strFormt_ORA & "') As ʵ�ս��, " & _
            "       ����ʱ��, ����id, ����, ��������id As ��������ID, b.���� As �������ұ���" & _
            " From (" & strSql & ") A,���ű� B" & _
            " Where a.��������ID = b.ID" & _
            " Order By ����,���, Ʊ�ݺ� Desc, ���ݺ� Desc"
        'ע��:����ҽ��Ҫ������һ�ſ�ʼ��,��������ܹؼ�
        Set mrsFeeList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatient, DatBegin, datEnd)
    
        '���ؿ�ѡ����
        mblnNotClick = True
        If cbo��������.ListIndex <> -1 Then lngPre��������ID = Val(cbo��������.ItemData(cbo��������.ListIndex))
        cbo��������.Clear
        cbo��������.AddItem "���п���"
        Do While Not mrsFeeList.EOF
            If InStr("," & strIDs & ",", "," & Nvl(mrsFeeList!��������ID) & ",") = 0 Then
                strIDs = strIDs & "," & Nvl(mrsFeeList!��������ID)
                
                cbo��������.AddItem IIf(zlIsShowDeptCode, Nvl(mrsFeeList!�������ұ���) & "-", "") & Nvl(mrsFeeList!��������)
                cbo��������.ItemData(cbo��������.NewIndex) = Nvl(mrsFeeList!��������ID)
                If Val(Nvl(mrsFeeList!��������ID)) = lngPre��������ID Then cbo��������.ListIndex = cbo��������.NewIndex
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
    If chkShow.value = vbChecked Then strFilter = strFilter & " And  ���='��ת��'"
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
                    If Not gclsInsure.GetCapability(support�����������, mlng����ID, intInsure) Then
                        .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
                    End If
                End If
            End If
        Next lngRow
    End With
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.value
End Sub

Private Function ExecuteThirdReturnMoneySwap(objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���������˿�
    '���:objPati-��ǰ����Ĳ�����Ϣ
    '     cllBillPro-��ǰִ�еĹ��̼�
    '     objBalanceInfor-��ǰ�Ľ�����Ϣ
    '����:
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
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, "ִ�������˿�", True
    Set cllPro = New Collection
    
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
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯ�����¼", objBalanceInfor.����ID)
    If rsTemp.RecordCount = 0 Then '�����������ѿ�����ֱ���˳�
        gcnOracle.RollbackTrans
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
    
    strSql = " " & _
    "   Select distinct a.����id, nvl(a.�����id,0) as �����id,a.������ˮ��,nvl(a.���㿨���,0) as ���㿨���,nvl(a.��������id,0) as ��������id " & _
    "   From ����Ԥ����¼ A, " & _
    "        (Select Distinct ����id " & _
    "          From ������ü�¼ " & _
    "          Where NO In (Select Distinct NO From ������ü�¼ Where ����id = [1]) And Mod(��¼����, 10) = 1 And ��¼״̬ In (3, 1)) B " & _
    "   Where a.����id = b.����id and mod(a.��¼����,10)<>1"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, "��ѯ�����˿���Ϣ", objBalanceInfor.����ID)
    
    Set objSequareDelItems = New clsBalanceItems
    Set objThirdDelItems = New clsBalanceItems
    
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lng�����ID = Val(Nvl(rsTemp!�����ID))
            lng���㿨��� = Val(Nvl(rsTemp!���㿨���))
            bln���ѿ� = lng���㿨��� <> 0
            lng��������ID = Val(Nvl(rsTemp!��������ID))
            
            rsBalance.Filter = "�����ID=" & lng�����ID & " and ��������ID=" & lng��������ID & " and ���㿨���=" & lng���㿨���
            lngԭ����ID = 0
            If Not rsBalance.EOF Then lngԭ����ID = Val(Nvl(rsBalance!����ID))
            If lngԭ����ID = 0 And Not bln���ѿ� Then
                rsBalance.Filter = "�����ID=" & lng�����ID & " and ������ˮ��='" & Nvl(!������ˮ��) & "'"
                If Not rsBalance.EOF Then lngԭ����ID = Val(Nvl(rsBalance!����ID))
                If lngԭ����ID = 0 Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    MsgBox Nvl(rsTemp!���㷽ʽ) & "δ�ҵ�ԭʼ�����¼ ������!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            
            Set objItem = New clsBalanceItem
            With objItem
                Set .objCard = mobjThirdSwap.zlGetCardFromCardType(lng�����ID, bln���ѿ�, Nvl(rsTemp!���㷽ʽ))
                .����ID = objBalanceInfor.����ID
                .����IDs = lngԭ����ID
                .����ID = lngԭ����ID
                .��������ID = lng��������ID
                .������ˮ�� = Nvl(rsTemp!������ˮ��)
                .����˵�� = Nvl(rsTemp!����˵��)
                .���㷽ʽ = Nvl(rsTemp!���㷽ʽ)
                .������� = Nvl(rsTemp!�������)
                .����ժҪ = Nvl(rsTemp!ժҪ)
                .������ = Val(Nvl(rsTemp!��Ԥ��))
                .�������� = IIf(bln���ѿ�, 5, 3)  '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                .�������� = .objCard.��������
                .����ʱ�� = Format(rsTemp!�տ�ʱ��, "yyyy-mm-dd HH:MM:SS")
                .���� = Nvl(rsTemp!����)
                .�����ID = IIf(bln���ѿ�, lng���㿨���, lng�����ID)
                .ʣ���� = Val(Nvl(rsTemp!��Ԥ��))
                .δ�˽�� = Val(Nvl(rsTemp!��Ԥ��))
                .ԭʼ��� = Val(Nvl(rsTemp!��Ԥ��))
                .���ѿ� = bln���ѿ�
                .���ݺ� = Nvl(rsTemp!NO)
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
                    rsTotal!�����ܶ� = Val(Nvl(rsTotal!�����ܶ�)) + Val(Nvl(rsTemp!�����ܶ�))
                End If
                rsTotal!��ϸ�ܶ� = RoundEx(Val(Nvl(rsTotal!��ϸ�ܶ�)) + objItem.������, 6)
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
                    Set objItemTemp = objItem.Clone
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
            If RoundEx(Val(Nvl(!�����ܶ�)), 6) <> RoundEx(Val(Nvl(!��ϸ�ܶ�)), 6) Then
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
        If mobjThirdSwap.zlThird_ReturnMoney_IsValied(objItem.objCard, 2, objBalanceInfor, objItem.objTag, objItems, False) = False Then
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
    Dim strHead As String, lngCardTypeID As Long, objCard As Card
    Dim i As Long
    With mshList
        strHead = "ѡ��,4,500|���,4,850|����,4,800|ҽ��,4,500|һ��ͨҽ��,1,550|���ݺ�,4,850|Ʊ�ݺ�,4,1100|������,1,800|��������,1,1200|" & _
            "��������ID,1,0|Ӧ�ս��,7,850|ʵ�ս��,7,850|����ʱ��,4,1850|����ID,4,0|����,4,0|��������ID,4,0|�������ұ���,4,0"
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
            'תҽ�ƿ����IDΪ������ʾ
            lngCardTypeID = Val(.TextMatrix(i, .ColIndex("һ��ͨҽ��")))
            .TextMatrix(i, .ColIndex("һ��ͨҽ��")) = ""
            If lngCardTypeID > 0 Then
                .Cell(flexcpData, i, .ColIndex("һ��ͨҽ��")) = lngCardTypeID
                If GetPayCard(lngCardTypeID, objCard) Then
                    .TextMatrix(i, .ColIndex("һ��ͨҽ��")) = objCard.����
                End If
                .ColHidden(.ColIndex("һ��ͨҽ��")) = False
            End If
        Next
        
        .Row = 1
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Function GetPayCard(ByVal lngCardTypeID As Long, ByRef objCard As Card, _
    Optional ByVal bln������ As Boolean, Optional ByVal bln���ѿ� As Boolean) As Boolean
    '���ݿ����ID��ȡ�������Ϣ
    On Error GoTo ErrHandler
    Set objCard = Nothing
    
    'zlGetCard:(ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean, ByRef objCard As Card)
    If mobjOneCardComLib.zlGetCard(lngCardTypeID, bln���ѿ�, objCard) = False Then Exit Function
    If Not objCard Is Nothing Then
        If bln������ And Not objCard.���� Then Set objCard = Nothing
    End If
    GetPayCard = Not objCard Is Nothing
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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

Private Sub cmdOK_Click()
    Dim i As Long, strNOs As String
    Dim blnThirdAllDel As Boolean, bnYBAllDel As Boolean
    Dim lng����ID As Long, str���ݺ� As String, intInsure As Long
    Dim strReplenishNo As String, strNotSelectNos As String
    Dim strTemp As String, blnErrBill As Boolean, strErrMsg As String
    Dim cllNO As Collection, cllPati As Collection
    
    Set mcllNOs = New Collection
    If mlng����ID = 0 Then
        MsgBox "δ���ֲ�����Ϣ�����飡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    zlCommFun.ShowFlash "����׼��ת�����ݣ����Ժ�...", Me
    
    'ֱ�ӱ���
    With mshList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = "��ת��" And .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
            
                lng����ID = Val(.TextMatrix(i, .ColIndex("����ID")))
                str���ݺ� = .TextMatrix(i, .ColIndex("���ݺ�"))
                intInsure = Val(.TextMatrix(i, .ColIndex("����")))
                strReplenishNo = "": strNotSelectNos = ""
                blnErrBill = False
                
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
                        strTemp = CheckInsureCancel(mlng����ID, intInsure, strReplenishNo, True)
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
                            If CheckBalanceAllNosIsSelected(lng����ID, .TextMatrix(i, .ColIndex("����")), strNOs) = False Then
                                zlCommFun.StopFlash
                                MsgBox "ҽ�����ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼����δת��ȫ����ؽ��㵥��,���ܼ���!", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            
                            '���ҽ���ֵ��ݣ�������Ϊȫ�ˣ�Ŀǰֻ�ܽ�ֹת��
                            If InStr(strNOs, ",") > 0 And bnYBAllDel = False And blnThirdAllDel Then
                                MsgBox "�ݲ�֧���ڱ�������תסԺ�����д���ҽ���ֵ��ݽ��㣬��һ��ͨ����ȫ�˵���������ܳɹ�תסԺ�ĵ������£�" & vbCrLf & strNOs, vbInformation + vbOKOnly, gstrSysName
                                zlCommFun.StopFlash
                                Exit Sub
                            End If
                        End If
                    Else
                        If CheckAllTurn(str���ݺ�) Then
                            If CheckBalanceAllNosIsSelected(lng����ID, .TextMatrix(i, .ColIndex("����"))) = False Then
                                zlCommFun.StopFlash
                                MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼����δת��ȫ����ؽ��㵥��,���ܼ���!", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                '���ݺ�,Ʊ�ݺ�,����ID,����,��������,�����㵥��,��������ID,������
                Set cllNO = New Collection
                cllNO.Add str���ݺ�, "���ݺ�"
                cllNO.Add .TextMatrix(i, .ColIndex("Ʊ�ݺ�")), "Ʊ�ݺ�"
                cllNO.Add lng����ID, "����ID"
                cllNO.Add intInsure, "����"
                cllNO.Add .TextMatrix(i, .ColIndex("����")), "��������"
                cllNO.Add strReplenishNo, "�����㵥��"
                cllNO.Add .TextMatrix(i, .ColIndex("��������ID")), "��������ID"
                cllNO.Add .TextMatrix(i, .ColIndex("������")), "������"
                mcllNOs.Add cllNO
            End If
        Next
    End With
    
    If mcllNOs.Count = 0 Then
        zlCommFun.StopFlash
        MsgBox "�㻹δѡ��Ҫת��סԺ���õĵ��ݣ��������̣�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ʽ����¼����,���ݺ�|��¼����,���ݺ�|... ���У���¼���ʣ�1-�����շѣ�2-�������
    strNOs = ""
    For i = 1 To mcllNOs.Count
        strNOs = strNOs & IIf(strNOs = "", "", "|")
        strNOs = strNOs & IIf(mcllNOs(i)("��������") = "���ʵ�", 2, 1) & "," & mcllNOs(i)("���ݺ�")
    Next
    
    Set cllPati = New Collection
    cllPati.Add mobjPati.����ID, "����ID"
    cllPati.Add mobjPati.��ҳID, "��ҳID"
    cllPati.Add mobjPati.����, "����"
    cllPati.Add mobjPati.��˱�־, "��˱�־"
    cllPati.Add mobjPati.סԺ״̬, "סԺ״̬"
    If mobjExpenceSvr.zlChargeTurnCheck(strNOs, cllPati, "�������תתסԺ���") = False Then Exit Sub
     
    strNOs = ""
    For i = 1 To mcllNOs.Count
        If i > 60 Then strNOs = strNOs & ",...": Exit For
        strNOs = strNOs & IIf(strNOs = "", "", ",")
        strNOs = strNOs & IIf(i > 0 And i Mod 6 = 0, vbCrLf, "")
        strNOs = strNOs & mcllNOs(i)("���ݺ�")
    Next
    If MsgBox("��ȷ��Ҫ�������������ת��סԺ������" & vbCrLf & _
        strNOs, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        zlCommFun.StopFlash
        Set mcllNOs = Nothing
        Exit Sub
    End If
    
    '����Ҫѡ����
    If mbln����ִ�� = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    
    Err = 0: On Error GoTo errHand:
    If mobjPati.��ҳID = 0 Then
        zlCommFun.StopFlash
        MsgBox "�ò��˻�δ��Ժ,�����������תסԺ����,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    LockScreen True
    If ExecuteTurn(Me, mobjPati.����ID, mobjPati.��ҳID, mobjPati.סԺ��, _
        CDate(mobjPati.��Ժ����), mobjPati.��ǰ����id, mobjPati.��ǰ����id, strErrMsg) = False Then
        LockScreen False
        Set mrsFeeList = Nothing
        Call cmdRefresh_Click
        zlCommFun.StopFlash
        If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Sub
    Else
        If Val(txtPatient.Tag) <> 0 And Val(txtPatient.Tag) = mobjPati.����ID Then mblnRefreshData = True
    End If
    zlCommFun.StopFlash
    LockScreen False
    
    If mlngModule = 1137 Then
       txtPatient.Text = ""
       Set mobjPati = Nothing
       mshDetail.Clear 1
       mshDetail.Rows = 2
       mshList.Clear 1
       mshList.Rows = 2
       vsBalance.Clear 1
       vsBalance.Rows = 2
       zlControl.ControlSetFocus txtPatient
       mlng����ID = 0
       Exit Sub
    End If
    Unload Me
    Exit Sub
errHand:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockScreen False
End Sub

Private Function GetReplenishAllNos(ByVal strNo As String) As String
    '��ȡ�����������з��õ���
    '���أ�
    '   �����������з��õ���:A001,A002,...
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strNOs As String
    
    On Error GoTo ErrHandler
    strSql = _
        " Select Distinct a.No" & vbNewLine & _
        " From ������ü�¼ A, ������ü�¼ B, ���ò����¼ C" & vbNewLine & _
        " Where a.No = b.No And a.��� = b.��� And a.��¼���� In (1, 11)" & vbNewLine & _
        "       And b.����id = c.�շѽ���id" & vbNewLine & _
        "       And c.��¼���� = 1 And c.���ӱ�־ = 0 And c.No = [1]" & vbNewLine & _
        " Group By a.No, a.���" & vbNewLine & _
        " Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    Do While Not rsTmp.EOF
        strNOs = strNOs & "," & Nvl(rsTmp!NO)
        rsTmp.MoveNext
    Loop
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    
    GetReplenishAllNos = strNOs
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckReplenishAllNosIsSelected(ByVal strNo As String, ByVal str���� As String, _
    Optional ByRef strNotSelectNos As String) As Boolean
    '��鲹����������ʣ��δ�˷��ñ����Ƿ�ѡ����ת��
    '��Σ�
    '   str���� �շѵ�/���ʵ�
    '���Σ�
    '   strNotSelectNos û�б�ѡ�����Ҫһ��ת���ĵ���
    Dim i As Integer, k As Long, blnFind As Boolean
    Dim strNOs As String, varNos As Variant
    
    On Error GoTo ErrHandler
    strNotSelectNos = ""
    strNOs = GetReplenishAllNos(strNo)
    
    varNos = Split(strNOs, ",")
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

Private Function GetReplenishInsure(ByVal strNo As String) As Long
    '��ȡ��������ҽ������
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSql = _
        " Select Max(b.����) As ����" & vbNewLine & _
        " From ����Ԥ����¼ A, ���ս����¼ B, ���ò����¼ C" & vbNewLine & _
        " Where a.����id = b.��¼id And a.��¼���� = 6" & vbNewLine & _
        "       And a.����id = c.����id And c.��¼���� = 1" & vbNewLine & _
        "       And c.��¼״̬ In(1,3) And c.���ӱ�־ = 0 And c.No = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    If Not rsTmp.EOF Then GetReplenishInsure = Nvl(rsTmp!����)
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
                    If .TextMatrix(i, .ColIndex("����")) = str���� And .TextMatrix(i, .ColIndex("���ݺ�")) = Nvl(rsTmp!NO) Then
                        If .TextMatrix(i, .ColIndex("���")) = "��ת��" And .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
                            blnFind = True: Exit For
                        End If
                    End If
                Next
                If blnFind = False Then blnNotIsSelected = True
            End If
            strNos_Out = strNos_Out & "," & Nvl(rsTmp!NO)
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
    
    Get�����ʻ����� = Nvl(rs���㷽ʽ!����)
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
    Dim strValue As String
    
    If zlGetOneCardComLibObject(Me, mlngModule, mobjOneCardComLib) = False Then Unload Me: Exit Sub
    If zlGetExpenceSvrObject(mobjExpenceSvr) = False Then Unload Me: Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjOneCardComLib, "", txtPatient)

    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTmp)
    mintIDKind = Val(strTmp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    mstrTitle = Me.Caption
    
    Call RestoreWinState(Me, App.ProductName)
    IDKindTime.NotAutoAppendKind = True
    IDKindTime.IDKindStr = "����ʱ��|����ʱ��|0|0|0|0|0|0|0|0|0;�Ǽ�ʱ��|�Ǽ�ʱ��|0|0|0|0|0|0|0|0|0"
    IDKindTime.IDKind = Val(zlDatabase.GetPara("�ϴ�ѡ��ʱ��ͳ������", glngSys, 1143, 0)) + 1
    
    mintPatientRange = Val(zlDatabase.GetPara("��ʾ���岡��", glngSys, 1137, 0))
    mbln����תסԺ����� = IIf(Val(zlDatabase.GetPara("����תסԺ�����", glngSys, 1143, 0)) = 1, True, False)
    mbln�������� = Val(zlDatabase.GetPara("����ת�������˷�", glngSys, 1131)) = 1
    mstr�����ʻ� = Get�����ʻ�����()
    '���ŵ���ʹ�ö��ֽ��㷽ʽģʽ
    mblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
    
    'Ʊ�ݺ��볤�ȡ����￨�ų���
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    mbytPrepayLen = Val(Split(strValue, "|")(1))
    If mbytPrepayLen = 0 Then mbytPrepayLen = 7
    'Ʊ���ϸ����
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    mblnPrepayStrict = Mid(strValue, 2, 1) = "1"
    
    mblnNotClick = True
    chkShow.value = IIf(Val(zlDatabase.GetPara("����ʾ��ת������", glngSys, 1131, 1, Array(chkShow))) = 1, 1, 0)
    mblnNotClick = False
    picBalance.BorderStyle = 0: picList.BorderStyle = 0:    picBill.BorderStyle = 0
    
    Call InitPancel
    
    Datsys = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ʼʱ��")
    If IsDate(strTmp) Then
        dtpBegin.value = CDate(strTmp)
    Else
        dtpBegin.value = Format(DateAdd("d", -3, Datsys), "yyyy-mm-dd 00:00:00")
    End If
    dtpBegin.MaxDate = Format(Datsys, "yyyy-mm-dd 23:59:59")
    dtpEnd.value = Format(Datsys, "yyyy-mm-dd 23:59:59")
    
    Call SetVisibleCtl
    Call setHeader: Call SetDetail: Call SetBalanceHead
    
    mblnNewClinicPati = False
    If mlng����ID = 0 Then
        Call ClearData
    Else
        If GetPatient(IDKind.GetCurCard, "-" & mlng����ID, False, True) Then
            If IsNewClinicPati(mobjPati.�Һ�ID) Then '�����ﲡ��
                Me.Hide
                If frmChargeTurnNew.ShowMe(Me, mobjPati.�Һ�ID, mbln����ִ��) Then
                    mblnOk = True: mblnNewClinicPati = True
                End If
                Unload Me: Exit Sub
            End If
        
            Call ShowBills(mlng����ID, dtpBegin.value, dtpEnd.value)
        End If
    End If
    
    If mbln����ִ�� = False Then
        fraPati.Visible = False: cmdOk.Visible = True
    Else
        fraPati.Visible = True: cmdOk.Visible = True
    End If
    Call picTop_Resize
End Sub

Private Function IsNewClinicPati(ByVal lng�Һ�ID As Long) As Boolean
    '�ж��Ƿ�Ϊ�����ﲡ��
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If lng�Һ�ID = 0 Then Exit Function
    strSql = "Select 1 From ���˹Һż�¼ Where Nvl(���ӱ�־,0) = 3 And Id = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�ж��Ƿ�Ϊ�����ﲡ��", lng�Һ�ID)
    IsNewClinicPati = Not rsTemp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
    Set mcllNOs = Nothing
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdRefresh_Click()
    If mlng����ID = 0 Then
        MsgBox "����ѡ���ˣ����飡", vbInformation, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    Call ShowBills(mlng����ID, dtpBegin.value, dtpEnd.value, False)
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
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ʼʱ��", Format(dtpBegin.value, "yyyy-MM-dd HH:mm:ss")
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����ʱ��", Format(dtpEnd.value, "yyyy-MM-dd HH:mm:ss")
    Call SaveWinState(Me, App.ProductName)
    
    Call zlDatabase.SetPara("����ʾ��ת������", chkShow.value, glngSys, 1131)
    zlDatabase.SetPara "�ϴ�ѡ��ʱ��ͳ������", IDKindTime.IDKind - 1, glngSys, 1143, InStr(1, mstrPrivs, ";��������;") > 0
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "����תסԺ��ϸ�б�", True
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "����תסԺ�����б�", True
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "����תסԺ�б�", True
    Set mrsFeeList = Nothing
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlOneCardComLib.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    zlControl.ControlSetFocus txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlOneCardComLib.Card, objPatiInfor As zlOneCardComLib.clsPatientInfo, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text = "" Then Exit Sub
    Call FindPati(objCard, False, Trim(txtPatient.Text))
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
    Dim strNo As String, str���� As String
    
    If NewRow = OldRow Then Exit Sub
    With mshList
        strNo = Trim(.TextMatrix(NewRow, .ColIndex("���ݺ�")))
        str���� = Trim(.TextMatrix(NewRow, .ColIndex("����")))
        If NewRow = 0 Or strNo = "" Then
            mshDetail.Clear 1: mshDetail.Rows = 2
            Call SetDetail
        Else
            Call ShowDetail(str����, strNo)
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
    ByVal strNo As String, Optional ByVal bln������ As Long) As String
    '���ҽ���Ƿ��ܹ�ԭ������
    '���أ�����ԭ�����ϣ��򷵻ؿգ����򣬷�����ʾ��Ϣ
    Dim strTmp As String, i As Integer
    Dim arrBalanceType As Variant, strBalanceType As String
    
    On Error GoTo ErrHandler
    If Not gclsInsure.GetCapability(support�����������, lng����ID, lngInsure) Then
        CheckInsureCancel = IIf(bln������, "ҽ���������", "") & "����[" & strNo & "]�Ĳ������಻֧������������ϣ�������ת����"
        Exit Function
    Else
        '���жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
        strTmp = GetBalanceType(strNo, bln������)
        arrBalanceType = Split(strTmp, ",")
        For i = 0 To UBound(arrBalanceType)
            strBalanceType = arrBalanceType(i)
            If Not gclsInsure.GetCapability(support�����������, lng����ID, lngInsure, strBalanceType) Then
                CheckInsureCancel = IIf(bln������, "ҽ���������", "") & "����[" & strNo & "]�Ĳ������಻֧��" & strBalanceType & "�������ϣ�������ת����"
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
    Dim intInsure As Integer, strNo As String, strTmp As String
    Dim str���� As String
    Dim blnAll As Boolean
    
    With mshList
        If .TextMatrix(lngRow, .ColIndex("���")) = "��ת��" And .TextMatrix(lngRow, .ColIndex("ѡ��")) <> IIf(blnSelect, "��", "") Then
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
            str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
            strNo = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
            
            If intInsure > 0 And blnSelect And str���� = "�շѵ�" Then
                strTmp = CheckInsureCancel(mlng����ID, intInsure, strNo)
                If strTmp <> "" Then
                    sta.Panels(2).Text = strTmp
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
                    Exit Function
                End If
            End If
            
            .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
            If str���� = "�շѵ�" Then
                If intInsure > 0 Then      'ȫ��ѡ���ȡ��
                    blnAll = gclsInsure.GetCapability(support�൥���շѱ���ȫ��, mlng����ID, intInsure)
                    If Not blnAll Then blnAll = Not IsYBSingle(strNo)
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

Private Function CheckAllTurn(ByVal strNo As String) As Boolean
    Dim strSql As String, rsData As ADODB.Recordset, lngCardTypeID As Long
    Dim strCardTypeIDs As String, strTemp As String
    Dim strWhere As String, objCard As Card
       
    On Error GoTo errHandle
           
    strWhere = "And  Not Exists(select 1 From ҽ��������ϸ Where NO=[1] And A.�����ID=�����ID and A.��������ID=��������ID) "
    
    strSql = "" & _
    "   Select A.���㷽ʽ,nvl(A.�����ID,0) as �����ID,nvl(A.���㿨���,0) as ���㿨���,nvl(A.��������ID,0) as ��������ID," & _
    "       max(nvl(E.�Ƿ�ȫ��,0)) as �Ƿ�ȫ��,nvl(max(decode(nvl(C.����,0),3,1,4,1,0)),0) as �Ƿ�ҽ��" & vbNewLine & _
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
    "         ) B,���㷽ʽ C,���ѿ����Ŀ¼ E" & vbNewLine & _
    "   Where a.����id = b.����id And a.��¼���� = 3 And A.���㷽ʽ=C.����(+) and A.���㿨���=E.���(+) " & vbNewLine & _
    "       " & strWhere & vbNewLine & _
    "   Group By A.���㷽ʽ,nvl(A.�����ID,0),nvl(A.���㿨���,0),nvl(A.��������ID,0) " & vbNewLine & _
    "   Having Sum(��Ԥ��) <> 0" & _
    "   Order by �����ID,��������ID"

    Set rsData = zlDatabase.OpenSQLRecord(strSql, "����Ƿ�ȫ��", strNo)
    If rsData.EOF Then CheckAllTurn = False: Exit Function
    
    rsData.Filter = "���㿨���<>0 And �Ƿ�ȫ��=1"
    If Not rsData.EOF Then CheckAllTurn = True: Exit Function   '����ȫ�˵�������������������
    
    rsData.Filter = "�����ID<>0 "
    Do While Not rsData.EOF
        If GetPayCard(rsData!�����ID, objCard) Then
            If objCard.�Ƿ�ȫ�� Then CheckAllTurn = True: Exit Function   '����ȫ�˵�������������������
        End If
        rsData.MoveNext
    Loop
    If Not rsData.EOF Then CheckAllTurn = True: Exit Function   '����ȫ�˵�������������������
    
    rsData.Filter = "�Ƿ�ҽ�� =1 And �����ID<>0"
    If Not rsData.EOF Then CheckAllTurn = True: Exit Function   'һ��ͨ����ҽ������ʱ����ȫ��(��SQL���ų��˷ֵ��ݽ����)������������
    
    rsData.Filter = "�����ID<>0"
    rsData.Sort = "�����ID,��������ID"
    
    With rsData
        strCardTypeIDs = ""
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(rsData!�����ID))
            strTemp = lngCardTypeID & ":" & Val(Nvl(rsData!��������ID))
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
    Dim i As Long, j As Long, k As Long, strNo As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant, blnAllTurn As Boolean
    Dim str���� As String, strReplenishNo As String
    Dim strNOs As String, varNos As Variant
    
    With mshList
        str���� = .TextMatrix(lngRow, .ColIndex("����"))
        If str���� = "���ʵ�" Then SetMultiOther = True: Exit Function
        If intInsure = 0 Then
            '����Ƿ�Ϊ�����㵥��
            If CheckBillExistReplenishData(1, , .TextMatrix(lngRow, .ColIndex("���ݺ�")), strReplenishNo) Then
                strNOs = GetReplenishAllNos(strReplenishNo)
                varNos = Split(strNOs, ",")
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
            
            If mblnMultiBalance Or blnAllTurn Then     '   �൥��,���ֽ��㷽ʽ
                '33635:ԭ���Ƕ൥���Ҷ��ֽ��㷽ʽ,���ܲ�����
                strNo = ""
                For k = 1 To .Rows - 1
                      If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                            And .TextMatrix(k, .ColIndex("����")) = str���� _
                            And Trim(.TextMatrix(lngRow, .ColIndex("����ID"))) <> "" Then
                            strNo = strNo & "," & .TextMatrix(k, .ColIndex("���ݺ�"))
                      End If
                Next
                If strNo <> "" Then strNo = Mid(strNo, 2)
                If InStr(1, strNo, ",") > 0 Then    '֤��Ϊ�൥��
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
                        strNo = .TextMatrix(i, .ColIndex("���ݺ�"))
                        '�жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                         strTmp = GetBalanceType(strNo)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                 strBalanceType = arrBalanceType(j)
                                 If Not gclsInsure.GetCapability(support�����������, mlng����ID, intInsure, strBalanceType) Then
                                     sta.Panels(2).Text = "����[" & strNo & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
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

Private Function GetBalanceType(ByVal strNo As String, _
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
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

Private Sub ShowDetail(ByVal str���� As String, ByVal strNo As String)
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
            " Select C.���� As ���, max(Decode(a.�Ƿ���, 1, '***', Nvl(E.����, B.����))) As ����, " & _
            "       B.���, A.���㵥λ As ��λ, Sum(Nvl(A.����, 1) * A.����) As ����," & _
            "       To_Char(A.��׼����, '" & gSysPara.Price_Decimal.strFormt_ORA & "') As ����, " & _
            "       To_Char(Sum(A.Ӧ�ս��), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As Ӧ�ս��," & _
            "       To_Char(Sum(A.ʵ�ս��), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As ʵ�ս��, D.���� As ִ�п���, 3 As ��¼״̬" & _
            " From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E" & _
            " Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And Mod(A.��¼����,10) = [2]" & _
            "      And A.��¼״̬ In (2,3) And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3 And Nvl(a.���ӱ�־, 0) <> 9 " & _
            " Group By A.��׼����,A.���, C.����, B.���, A.���㵥λ, D.����" & _
            " Having Sum(A.����) <> 0 "
        
        strSql = strSql & " Union All" & _
            " Select C.���� As ���,max(Decode(a.�Ƿ���, 1, '***', Nvl(E.����, B.����))) As ����," & _
            "       B.���, A.���㵥λ As ��λ, Sum(Nvl(A.����, 1) * A.����) As ����," & _
            "       To_Char(A.��׼����, '" & gSysPara.Price_Decimal.strFormt_ORA & "') As ����, " & _
            "       To_Char(Sum(A.Ӧ�ս��), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As Ӧ�ս��," & _
            "       To_Char(Sum(A.ʵ�ս��), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As ʵ�ս��, D.���� As ִ�п���, 1 As ��¼״̬" & _
            " From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E" & _
            " Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And Mod(A.��¼����,10) = [2] " & _
            "      And A.��¼״̬=1 And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3 And Nvl(A.���ӱ�־,0) <> 9 " & _
            " Group By A.��׼����,A.���, C.����, B.���, A.���㵥λ, D.����" & _
            " Having Sum(A.����) <> 0 "
    
    ElseIf mshList.TextMatrix(mshList.Row, mshList.ColIndex("���")) = "����ת��" Then
        strSql = _
        " Select C.���� As ���, max(Decode(a.�Ƿ���, 1, '***', Nvl(E.����, B.����))) As ����," & _
        "       B.���, A.���㵥λ As ��λ, Sum(Nvl(A.����, 1) * A.����) As ����," & _
        "       To_Char(A.��׼����, '" & gSysPara.Price_Decimal.strFormt_ORA & "') As ����, " & _
        "       To_Char(Sum(A.Ӧ�ս��), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As Ӧ�ս��," & _
        "       To_Char(Sum(A.ʵ�ս��), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As ʵ�ս��, D.���� As ִ�п���, 2 As ��¼״̬" & _
        " From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E" & _
        " Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And Mod(A.��¼����,10) = [2] " & _
        "      And A.��¼״̬ In (1,3) And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3 And Nvl(A.���ӱ�־,0) <> 9 " & _
        " Group By A.��׼����,A.���, C.����,B.���, A.���㵥λ, D.���� Having Sum(A.����) <> 0 "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo, IIf(str���� = "���ʵ�", 2, 1))
    
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
        mshList.Width = .ScaleWidth
        mshList.Height = .ScaleHeight
    End With
End Sub
Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Top = .ScaleTop
        vsBalance.Width = .ScaleWidth
        lblSum.Top = .ScaleHeight - lblSum.Height
        vsBalance.Height = lblSum.Top - mshDetail.Top
    End With
End Sub

Private Sub picBottom_Resize()
    Err = 0: On Error Resume Next
    With picBottom
            cmdCancel.Left = .ScaleLeft + .ScaleWidth - cmdCancel.Width - 400
            cmdOk.Left = cmdCancel.Left - cmdOk.Width - 20
            cmdOk.Top = cmdCancel.Top
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        mshDetail.Left = .ScaleLeft
        mshDetail.Top = .ScaleTop
        mshDetail.Width = .ScaleWidth
        mshDetail.Height = .ScaleHeight
    End With
End Sub

Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    If mbln����ִ�� Then
        fraPati.Left = picTop.ScaleLeft + 150
        IDKindTime.Left = fraPati.Left + fraPati.Width + 20
    Else
        IDKindTime.Left = picTop.ScaleLeft + 150
    End If
    dtpBegin.Left = IDKindTime.Left + IDKindTime.Width + 30
    lbl��.Left = dtpBegin.Left + dtpBegin.Width + 50
    dtpEnd.Left = lbl��.Left + lbl��.Width + 50
    
    fraFixed.Left = fraPati.Left + IIf(mbln����תסԺ����� And mbln����ִ��, fraPati.Width + 150, 150)
    fraFixed.Top = IIf(mbln����תסԺ�����, 80, 450)
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
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
    Dim blnOutMsg As Boolean
    
    If GetPatient(objCard, strInput, blnCard, , blnOutMsg) Then
        '69526:������,2014-02-13,��Ժ�����޷���������תסԺ����
        If Val(zlDatabase.GetPara("��Ժ������������תסԺ", glngSys, 1137, "0")) = 0 Then
            If Not mobjPati.��Ժ Then
                MsgBox "����" & mobjPati.���� & "�Ѿ���Ժ��δ����סԺ������������������תסԺ������", vbInformation, gstrSysName
                txtPatient.Text = "": mlng����ID = 0
                Call ClearData
                Set mobjPati = Nothing
                If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
                Exit Sub
            End If
        End If
        
        If IsNewClinicPati(mobjPati.�Һ�ID) Then '�����ﲡ��
            Me.Hide
            On Error Resume Next
            Call frmChargeTurnNew.ShowMe(Me, mobjPati.�Һ�ID, True)
            Err = 0: On Error GoTo 0
            
            txtPatient.Text = "": mlng����ID = 0
            Call ClearData
            Set mobjPati = Nothing
            If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
            
            Me.Show vbModal, mfrmMain
            Exit Sub
        End If
        
        '��ʱ������ʽ�����¼�Form_Load
        Call ShowBills(mlng����ID, dtpBegin.value, dtpEnd.value)
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
        If Not blnOutMsg Then MsgBox "û���ҵ��ò���,�������������Ƿ���ȷ��", vbInformation, gstrSysName
        txtPatient.Text = "": mlng����ID = 0
        Call ClearData
        If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus
    End If
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
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
    If mobjPati Is Nothing Then Exit Sub
    If txtPatient.Text <> mobjPati.���� Then txtPatient.Text = mobjPati.����
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    ByVal blnCard As Boolean, Optional ByVal blnFindByPatiID As Boolean, Optional ByRef blnOutMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '���:
    '   objCard=������
    '   strInput=�����ı�
    '   blnCard=�Ƿ�ˢ��
    '   blnFindByPatiID=ֱ�Ӱ�����ID����
    '   blnOutMsg-�Ѿ���ʾ,�������ⲿ����ʾ
    '����:
    '   blnCancel=���ڱ�ʾ����ȡ��
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePassWord As Boolean
    
    On Error GoTo errH
    Call GetPatientInfo(mobjPati, objCard, strInput, blnCard, IDKind.Cards.��ȱʡ������, IDKind.DefaultCardType, _
        Me, txtPatient, blnFindByPatiID, blnHavePassWord, blnOutMsg, mintPatientRange)
    If mobjPati Is Nothing Then Set mobjPati = Nothing: Exit Function
    
    txtPatient.Text = mobjPati.����: mlng����ID = mobjPati.����ID
    If mobjPati.��Ժ���� <> "" Then
        '�������Ϊ��Ժ����,����ת��סԺ�����е��������
        dtpEnd.MaxDate = CDate(Format(mobjPati.��Ժ����, "yyyy-mm-dd 23:59:59"))
        dtpEnd.value = dtpEnd.MaxDate
        dtpEnd.MaxDate = dtpEnd.MaxDate + 1
        dtpBegin.MaxDate = dtpEnd.value
        '����: 36609 ����Ժʱ��Ҫ��һ��,��Ϊ���ܴ��ڲ�����û���������ʱ,����Ժ,��ȥ�������,�Ӷ�����������ת���˵����.
    End If
    
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mobjPati = Nothing
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
    Dim strNOs As String
    Dim blnNotFirst As Boolean
    
    On Error GoTo errHandle
    If zlstr.IsHavePrivs(mstrPrivs, "Ԥ�����վݴ�ӡ") = False Then
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
        strNOs = ""
        Do While Not rsTemp.EOF
            strNOs = strNOs & "," & Nvl(rsTemp!NO)
            rsTemp.MoveNext
        Loop
        If strNOs <> "" Then
            strNOs = Mid(strNOs, 2)
            If PrintInvoice(strNOs, strDelDate) = False Then Exit Function
        End If
    Else
        '0-�����ɵ�Ԥ�����ݷֱ��ӡ
        Do While Not rsTemp.EOF
            If PrintInvoice(Nvl(rsTemp!NO), strDelDate, Not blnNotFirst) = False Then Exit Function
            blnNotFirst = True
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
    Dim rsTemp As ADODB.Recordset, strSql As String
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
            "       Group By B.NO) A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C " & _
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

Private Function PrintInvoice(ByVal strNOs As String, ByVal strDelDate As String, Optional ByVal blnFirstBill As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ʊ����
    '��Σ�
    '   strNos ���δ�ӡԤ�����ݺţ���ʽ��A001,A002,A003,...
    '   blnFirstBill �Ƿ��һ��Ʊ�ݣ�����Ĳ����ظ���ʾ
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
    If mblnPrepayStrict Then
        lngShareUseID = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, 1131, 0)
        '1.�ϸ����Ʊ��ʱ������ʵ�ʵ�Ʊ������,���¼������ID��Ʊ�ݺ�
        lng����ID = GetInvoiceGroupID(2, 1, lng����ID, lngShareUseID, strInvoice, "2")
        If lng����ID <= 0 Then
            Select Case lng����ID
                Case -1
                    MsgBox "Ԥ������[" & strNOs & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "��û���㹻�����ú͹��õ�Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "����[" & strNOs & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "��û���㹻�ĵĹ���Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox "����[" & strNOs & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & strInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                        "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -4
                    MsgBox "����[" & strNOs & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & strInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                        "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ�,�ش�õ��ݣ�", vbInformation, gstrSysName
                Case Else
                    MsgBox "Ʊ��������Ϣ����ʧ�ܣ�������������ش򵥾�[" & strNOs & "]", vbInformation, gstrSysName
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
                If blnFirstBill Or strInvoice = "" Then
                    strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, strInvoice, Me.Left + 1500, Me.Top + 1500))
                End If
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
                If blnFirstBill Or strInvoice = "" Then
                    strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, strInvoice, Me.Left + 1500, Me.Top + 1500))
                End If
                blnInput = True
            End If
                 
             '�û�ȡ������,�����ӡ
             If strInvoice = "" Then
                 If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                 blnValid = True
             Else
                 '���������Ч��
                 If blnInput Then
                     If zlCommFun.ActualLen(strInvoice) <> mbytPrepayLen Then
                         MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & mbytPrepayLen & " λ��", vbInformation, gstrSysName
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
    strSql = strSql & "'" & strNOs & "',"
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
        "NO=" & strNOs, "�տ�ʱ��=" & Format(strDelDate, "yyyy-mm-dd HH:MM:SS"), _
        "����ID=" & mlng����ID, IIf(intInvoiceFormat = 0, "", "ReportFormat=" & intInvoiceFormat), 2)
    
    '���±���Ʊ��
    If Not mblnPrepayStrict Then
        zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", strInvoice, glngSys, 1131
    End If
    PrintInvoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetBalanceNos(ByVal bytTYPE As Byte, _
    ByVal strFindValue As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional bln������ As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ���շѵ��ݵ�NO�����ID�������ţ�����ͬһ�ν����NOs
    '���:bytType-0-����NO������;1-���ݽ���ID������,2-���ݽ������������
    '    strFindValue-���ҵ�ֵ
    '    blnNOMoved-�Ƿ��ں󱸱��У���ѯ����֮ǰ���ж���Ҫ���������
    '    bln������-�Ƿ�ҽ��������
    '����:��ʽ��"AAA,BBB,CCC,..."
    '����:���˺�
    '����:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, strNOs As String

    On Error GoTo errHandle:
    Select Case bytTYPE
    Case 0 '0-����NO������
        If bln������ Then
            strSql = "" & _
            "   Select distinct A.NO " & _
            "   From ������ü�¼ A,(Select distinct �շѽ���ID as ����ID From ���ò����¼ Where NO=[1] and ��¼����=1 ) B" & _
            "   Where A.����ID=B.����ID" & _
            "   Order by NO"
        Else
            strSql = _
                "Select Distinct a.No" & vbNewLine & _
                "From ������ü�¼ A, ������ü�¼ B" & vbNewLine & _
                "Where a.����id = b.����id And Mod(a.��¼����, 10) = 1" & vbNewLine & _
                "      And Mod(b.��¼����, 10) = 1 And b.No = [1]" & vbNewLine & _
                "Order By NO"
        End If
    Case 1  '1-���ݽ���ID������
        If bln������ Then
            strSql = "" & _
            "    Select Distinct A.No " & _
            "    From ������ü�¼ A," & _
            "        (Select distinct C1.�շѽ���ID as ����ID " & _
            "         From ���ò����¼ A1,���ò����¼ B1,���ò����¼ C1  " & _
            "         Where A1.����ID=[2] and A1.��¼����=1  " & _
            "               And A1.NO=B1.NO and A1.��¼����=B1.��¼���� " & _
            "               And B1.�������=C1.������� and C1.��¼״̬ in (1,3) ) B " & _
            "    Where A.����ID=B.����ID    " & _
            "    Order By NO"
        Else
            strSql = _
                "Select Distinct a.No" & vbNewLine & _
                "From ������ü�¼ A, ������ü�¼ B, ������ü�¼ C" & vbNewLine & _
                "Where a.No = b.No And Mod(a.��¼����, 10) = 1" & vbNewLine & _
                "      And b.����id = c.����id And c.����id = [2]" & vbNewLine & _
                "Order By NO"
        End If
    Case 2  '2-���ݽ������������
        If bln������ Then
            strSql = "" & _
            "    Select Distinct A.No " & _
            "    From ������ü�¼ A," & _
            "        (Select distinct C1.�շѽ���ID as ����ID " & _
            "         From ���ò����¼ A1,���ò����¼ B1,���ò����¼ C1  " & _
            "         Where A1.�������=[2] and A1.��¼����=1  " & _
            "               And A1.NO=B1.NO and A1.��¼����=B1.��¼���� " & _
            "               And B1.�������=C1.������� and C1.��¼״̬ in (1,3) ) B " & _
            "    Where A.����ID=B.����ID    " & _
            "    Order By NO"
        Else
            strSql = _
                "Select Distinct a.No" & vbNewLine & _
                "From ������ü�¼ A, ������ü�¼ B, ������ü�¼ C" & vbNewLine & _
                "Where a.No = b.No And Mod(a.��¼����, 10) = 1 And b.����id = c.����id" & vbNewLine & _
                "      And c.����id In (Select ����id From ����Ԥ����¼ Where ������� = [2])" & vbNewLine & _
                "Order By NO"
        End If
    End Select
    If blnNOMoved Then
        strSql = Replace(strSql, "������ü�¼", "H������ü�¼")
        strSql = Replace(strSql, "����Ԥ����¼", "H����Ԥ����¼")
        strSql = Replace(strSql, "���ò����¼", "H���ò����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���ݵ��ݻ�ȡһ�ν��ʵĵ���", strFindValue, Val(strFindValue))
    
    With rsTemp
        Do While Not .EOF
            strNOs = strNOs & "," & !NO
            .MoveNext
        Loop
    End With
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    GetBalanceNos = strNOs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBillExistReplenishData(intTYPE As Integer, _
    Optional lngBalance As Long, Optional strNOs As String, _
    Optional ByRef strReplenishNo As String, Optional ByRef blnErrBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ���ڶ��ν���
    '���:intType:0-�շ����ݣ�ʹ��lngBalanceΪ�������
    '     intType:1-�շ����ݣ�ʹ��strNosΪ���ݺ�
    '���Σ�
    '   strReplenishNo ������㵥�ݺ�
    '   blnErrBill �Ƿ��쳣���㵥��
    '����:True-���ڶ��ν������� False-�����ڶ��ν�������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strReplenishNo = ""
    If intTYPE = 0 Then
        strSql = _
            " Select Max(a.NO) As No,Max(a.����״̬) As ����״̬" & vbNewLine & _
            " From ���ò����¼ A, (Select Distinct ����id From ����Ԥ����¼ Where ������� = [1]) B" & vbNewLine & _
            " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2"
        strSql = strSql & _
            " Union All" & _
            " Select Max(a.NO) As No,Max(a.����״̬) As ����״̬ From ���ò����¼ A Where a.������� = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����ν���", lngBalance)
    Else
        strSql = _
            " Select Max(a.NO) As No,Max(a.����״̬) As ����״̬" & vbNewLine & _
            " From ���ò����¼ A," & vbNewLine & _
            "      (Select /*+cardinality(j,10)*/Distinct a.����id" & vbNewLine & _
            "       From ������ü�¼ A,Table(f_Str2list([1])) J" & vbNewLine & _
            "       Where Mod(a.��¼����, 10) = 1 And a.NO=j.Column_Value) B" & vbNewLine & _
            " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����ν���", strNOs)
    End If
    
    strReplenishNo = Nvl(rsTmp!NO)
    blnErrBill = Val(Nvl(rsTmp!����״̬)) = 1
    CheckBillExistReplenishData = strReplenishNo <> ""
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatientInfo(ByRef objPati As clsPatientInfo, _
    ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, _
    ByVal blnDefaultCardFind As Boolean, ByVal lngDefaultCardTypeID As Long, _
    frmMain As Object, objText As Object, _
    Optional ByVal blnFindByPatiID As Boolean, _
    Optional ByRef blnHavePassWord As Boolean, _
    Optional ByRef blnCancel As Boolean, _
    Optional ByVal intPatientRange As Integer = -1) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ������Ϣ
    '��Σ�
    '   objCard=����ʶ��Ŀ�����
    '   strInput=�û�����ֵ
    '   blnCard=�Ƿ�ˢ��
    '   blnDefaultCardFind=�Ƿ�ȱʡ������
    '   lngDefaultCardTypeID=ȱʡ�����ID
    '   frmMain=����ؼ����ڴ���
    '   objText=����ؼ�
    '   blnFindByPatiID=ֱ�Ӱ�����ID����
    '   lng��ҳID=ָ��סԺ����
    '   intPatientRange-����������ʱ,�Ƿ�ֻ��ʾδ����õĲ���,0-���ѽ���,1-δ����,2-���δ����,3-סԺδ����
    '���Σ�
    '   blnHavePassWord=�Ƿ���Ҫ������֤
    '���أ����ز�����Ϣ
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, lng��ҳID As Long
    Dim varCardType As Variant, blnFind As Boolean
    Dim strCardPass As String, lng�����ID As Long
    
    On Error GoTo ErrHandler
    blnHavePassWord = False
    
    blnFind = False
    If blnCard And objCard.���� Like "����*" And InStr("-+*", Left(strInput, 1)) = 0 Then '103563,ȱʡ�����
        If blnDefaultCardFind And lngDefaultCardTypeID > 0 Then
            varCardType = lngDefaultCardTypeID
        Else
            varCardType = -1
        End If
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Or blnFindByPatiID Then  '����ID
        lng����ID = Mid(strInput, 2)
        blnFind = True
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strInput = Mid(strInput, 2)
        varCardType = "�����"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strInput = zlFormatID(Mid(strInput, 2))
        If mobjOneCardComLib.zlGetPatiIDFromInpatientNum(strInput, lng����ID, , , lng��ҳID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        blnFind = True
    ElseIf Left(strInput, 1) = "^" And IsNumeric(Mid(strInput, 2)) Then '���ۺ�
        strInput = Mid(strInput, 2)
        varCardType = "���ۺ�"
    Else
        If Not objPati Is Nothing Then
            If objPati.���� = strInput Then GetPatientInfo = True: Exit Function
        End If
        
        Select Case objCard.����
        Case "����", "��������￨"
            If GetPatiIdFromPatiName(objText, strInput, lng����ID, frmMain, blnCancel, intPatientRange) = False Then GoTo NotFoundPati:
            strInput = lng����ID
            blnFind = True
        Case "ҽ����"
            strInput = UCase(strInput)
            varCardType = objCard.����
        Case "�����"
            If Not IsNumeric(strInput) Then GoTo NotFoundPati:
            varCardType = objCard.����
        Case "סԺ��"
            If Not IsNumeric(strInput) Then GoTo NotFoundPati:
            strInput = zlFormatID(strInput)
            If mobjOneCardComLib.zlGetPatiIDFromInpatientNum(strInput, lng����ID, , , lng��ҳID) = False Then GoTo NotFoundPati:
            If lng����ID <= 0 Then GoTo NotFoundPati:
            blnFind = True
        Case "���ۺ�"
            If Not IsNumeric(strInput) Then GoTo NotFoundPati:
            varCardType = objCard.����
        Case Else
            If objCard.�ӿ���� > 0 Then
                varCardType = objCard.�ӿ����
            Else
                varCardType = objCard.����
            End If
            blnHavePassWord = True
        End Select
    End If
    
    If blnFind = False Then
        If mobjOneCardComLib.zlGetPatiID(varCardType, strInput, , lng����ID, strCardPass, , lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID > 0 Then blnFind = True
    End If
    If blnFind = False Then GoTo NotFoundPati:
    
    Set objPati = GetPatiInfo(lng����ID, lng��ҳID)
    If objPati Is Nothing Then GoTo NotFoundPati:
    
    objPati.���� = strCardPass
    GetPatientInfo = True
    Exit Function
NotFoundPati:
    Set objPati = Nothing
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set objPati = Nothing
End Function

Private Function GetPatiIdFromPatiName(ByVal objControl As Object, ByVal strName As String, ByRef lngPatiID As Long, _
    Optional frmMain As Object, Optional ByRef blnCancel As Boolean, Optional ByVal intPatientRange As Integer = -1) As Boolean
    '����:���ݲ�����������ȡ������Ϣ
    '���:
    '   objControl-���õĿؼ�
    '   strName-����Ĳ�����Ϣ
    '   frmMain-���õ�������
    '   intPatientRange-����������ʱ,�Ƿ�ֻ��ʾδ����õĲ���,0-���ѽ���,1-δ����,2-���δ����,3-סԺδ����
    '���Σ�
    '   lngPatiId=ѡ��Ĳ���ID
    '   blnCancel=�Ƿ��û�ȡ��ѡ��
    '����:�ɹ�����true,���򷵻�False
    '˵��:������ʱ����
    Dim rsPati As ADODB.Recordset
    Dim i As Long
    Dim strSql As String, strWhere As String
    Dim cllFilter As Collection, rsPatiPageInfo As ADODB.Recordset
    Dim str����IDs As String, strSubTable As String, varPara() As Variant
    Dim rsFee As ADODB.Recordset
    Dim str��ҳIDs As String
    Dim vRect As RECT, rsOutSel As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mobjOneCardComLib.zlGetPatiRecordFromPatiName(strName, rsPati) = False Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    '   rsPati-������Ϣ��,�ֶΣ�����ID,ID,����ID,��ҳID,����,�Ա�,����,��������,����,�����,סԺ��,����,
    '                           ����,��������,���֤��,��ͥ��ַ,������λ,��Ժ��־,�ֻ���,�Ƿ���ҽ�ƿ�,����ʱ��,��Ժ����,��Ժ����
    Do While Not rsPati.EOF
        str����IDs = str����IDs & "," & Nvl(rsPati!����ID)
        If Val(Nvl(rsPati!��ҳID)) <> 0 Then
            str��ҳIDs = str��ҳIDs & "," & Nvl(rsPati!����ID) & ":" & Nvl(rsPati!��ҳID)
        End If
        rsPati.MoveNext
    Loop
    rsPati.MoveFirst
    
    If str��ҳIDs <> "" Then
        Set cllFilter = New Collection
        cllFilter.Add Array("��ҳIDS", Mid(str��ҳIDs, 2))
        If GetPatiPageInfByRange(cllFilter, rsPatiPageInfo) = False Then Exit Function
    End If
    
    If intPatientRange >= 0 Then
        '��ȡ����δ����õĲ���
        str����IDs = Mid(str����IDs, 2)
        If zlGetVarBoundSQL(0, str����IDs, strSubTable, varPara, 0) = False Then Exit Function
        
        Select Case intPatientRange
        Case 1  '�κη���δ���岡��
            strWhere = ""
        Case 2  '���δ����Ĳ���
            strWhere = " And a.��Դ;�� = 4"
        Case 3  'סԺδ����Ĳ���
            strWhere = " And a.��Դ;�� = 2"
        Case 4  '����δ����Ĳ���
            strWhere = " And a.��Դ;�� = 1"
        End Select
        strSql = "Select a.����ID" & _
                " From ����δ����� A,(" & strSubTable & ") B" & _
                " Where a.����ID=b.Column_Value" & strWhere & _
                " Group By a.����ID"
        Set rsFee = zlDatabase.OpenSQLRecordByArray(strSql, "��ѯ����δ����õĲ���", varPara)
        
        For i = rsPati.RecordCount To 1 Step -1
            rsFee.Filter = "����ID=" & Nvl(rsPati!����ID)
            If rsFee.EOF Then
                rsPati.Delete adAffectCurrent
            ElseIf Not rsPatiPageInfo Is Nothing Then
                rsPatiPageInfo.Filter = "����ID=" & Nvl(rsPati!����ID)
                If Not rsPatiPageInfo.EOF Then
                    rsPati!��Ժ���� = Format(Nvl(rsPatiPageInfo!��Ժʱ��), "yyyy-MM-dd")
                    rsPati!��Ժ���� = Format(Nvl(rsPatiPageInfo!��Ժʱ��), "yyyy-MM-dd")
                End If
            End If
            rsPati.MoveNext
        Next
    End If
    
    rsPati.Sort = "��Ժ��־ Desc,��Ժ���� Desc"
    If rsPati.RecordCount = 0 Then Exit Function
    If rsPati.RecordCount = 1 Then
        lngPatiID = Val(rsPati!����ID)
        GetPatiIdFromPatiName = True: Exit Function
    End If
    
    vRect = zlControl.GetControlRect(objControl.hWnd)
    Set rsOutSel = zlDatabase.ShowRecSelect(frmMain, rsPati, 0, "����ѡ����", _
        False, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, objControl.Height, _
        blnCancel, False, True, 1, "����ID,ID,��ҳID,��������,����,����,�ֻ���,�Ƿ���ҽ�ƿ�,����ʱ��")
    If rsOutSel Is Nothing Then Exit Function
    If rsOutSel.EOF Then Exit Function
    
    lngPatiID = Val(rsOutSel!����ID)
    GetPatiIdFromPatiName = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiInfoByPage(objPati As clsPatientInfo, _
    Optional ByVal lng��ҳID As Long, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ӳ�����ҳ�л�ȡ������Ϣ
    '���:
    '   objPati-���в�����Ϣ
    '   lng��ҳID-��ҳID��Ϊ0ʱ��ȡ���һ��סԺ��
    '����:
    '   objPati-���ز�����Ϣ����
    '����:�ɹ�����True�����򷵻�False
    '˵��:������� objPati ��ΪNothing���������Ϣ�ϲ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    If objPati Is Nothing Then Exit Function
    If objPati.����ID = 0 Then Exit Function
    If zlGetServiceObject(objService) = False Then Exit Function
    
    If lng��ҳID = 0 Then lng��ҳID = objPati.��ҳID '������Ϣ�е���ҳIDΪ���һ����ҳID
    If objService.ZlCissvr_GetPatiPageInfo(1, objPati.����ID & ":" & lng��ҳID, rsTemp, , , lngModule) = False Then Exit Function
    If rsTemp Is Nothing Then GetPatiInfoByPage = True: Exit Function
    If rsTemp.EOF Then GetPatiInfoByPage = True: Exit Function
    
    If objPati Is Nothing Then Set objPati = New clsPatientInfo
    With objPati
        .��ҳID = Nvl(rsTemp!��ҳID)
        .���� = Nvl(rsTemp!����)
        .�Ա� = Nvl(rsTemp!�Ա�)
        .���� = Nvl(rsTemp!����)
        .�ѱ� = Nvl(rsTemp!�ѱ�)
        .ҽ�Ƹ��ʽ = Nvl(rsTemp!ҽ�Ƹ��ʽ����)
        .ҽ�Ƹ��ʽ���� = Nvl(rsTemp!ҽ�Ƹ��ʽ����)
        .���� = Val(Nvl(rsTemp!����))
        .�������� = GetInsureName(Val(Nvl(rsTemp!����)))
        .�������� = Nvl(rsTemp!��������)
        .��ǰ����id = Val(Nvl(rsTemp!��ǰ����id))
        .��ǰ�������� = Nvl(rsTemp!��ǰ��������)
        .��ǰ����id = Val(Nvl(rsTemp!��ǰ����id))
        .��ǰ�������� = Nvl(rsTemp!��ǰ��������)
        .���� = Nvl(rsTemp!��ǰ����)
        .סԺ�� = Nvl(rsTemp!סԺ��)
        .�������� = Val(Nvl(rsTemp!��������))
        .��Ժ���� = Nvl(rsTemp!��Ժʱ��)
        .��Ժ���� = Nvl(rsTemp!��Ժʱ��)
        .סԺҽʦ = Nvl(rsTemp!סԺҽʦ)
        .���˱�ע = Nvl(rsTemp!���˱�ע)
        .סԺ״̬ = Val(Nvl(rsTemp!סԺ״̬))
        .��˱�־ = Val(Nvl(rsTemp!��˱�־))
        .��Ŀ���� = Nvl(rsTemp!��Ŀ����)
        .ҽ���� = Nvl(rsTemp!ҽ����)
        .�Һ�ID = Val(Nvl(rsTemp!�Һ�ID))
    End With
    GetPatiInfoByPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPatiInfo(ByVal lng����ID As Long, _
    Optional ByVal lng��ҳID As Long, Optional ByVal lngModule As Long) As clsPatientInfo
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ���ȴӲ�����Ϣ�л�ȡ���ٴӲ�����ҳ�л�ȡ���кϲ�
    '���:
    '   objPati-���в�����Ϣ
    '   lng��ҳID-��ҳID��Ϊ0ʱ��ȡ���һ��סԺ��
    '����:
    '   objPati-���ز�����Ϣ����
    '����:�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPati As clsPatientInfo
    
    On Error GoTo errHandle
    '��ȡ������Ϣ
    If mobjOneCardComLib.zlGetPatiInforFromPatiID(lng����ID, objPati) = False Then Exit Function
    If objPati Is Nothing Then Exit Function
    
    '2.��ȡ������ҳ
    If lng��ҳID = 0 Then lng��ҳID = objPati.��ҳID
    If GetPatiInfoByPage(objPati, lng��ҳID, lngModule) = False Then Exit Function
    
    Set GetPatiInfo = objPati
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPatiPageInfByRange(ByVal cllFilter As Collection, _
    ByRef rsPatiPageInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ѯ������ҳ��Ϣ
    '���:
    '   cllFilter ��ѯ������:��Ա(Array(Key,Value),Array(Key,Value),,...)
    '       Key:����IDS,����IDS,����IDS,��ҳIDS,��Ժ��ʼʱ��,��Ժ����ʱ��,��Ժ��ʼʱ��,��Ժ����ʱ��,
    '           �ѱ�,סԺ״̬,��������,����,վ����,��ѯת�Ʋ���,���һ��סԺ,����,����վ����
    '       סԺ״̬:0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ
    '       �������ʣ�����ö��ŷ�0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�NULL-��ʾ������
    '       ����:���Դ�%�ֺű������ƥ��
    '       �ѳ�Ժ������סԺ״̬Ϊ1��2ʱ��Ч
    '       վ����:���Ҷ�Ӧ��վ����
    '       ����:>0:ָ������ҽ������,0:ҽ������ͨ����,-1:��ͨ����,-2:ҽ������
    '����:
    '   rsPatiPageInfo ���˲�����ҳ��Ϣ������ID,��ҳID,����,�Ա�,����,סԺ��,����,����,�ѱ�,��������,ҽ����,
    '                                   ��Ժʱ��,��Ժʱ��,סԺ״̬,��������,��ǰ����ID,��ǰ��������,��ǰ����ID,��ǰ��������,
    '                                   ҽ�Ƹ��ʽ����,ҽ�Ƹ��ʽ����,סԺҽʦ,���˱�ע,��Ŀ����,����ȼ�,
    '                                   ����ת��,��˱�־,�����,Ԥ��Ժʱ��,�ϴδ߿���
    '       סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)
    '       ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    '       ����ת��:0-δת����1-��ת��
    '       ��˱�־:0���-δ���,1-����˻�ʼ���;2-������
    '����:��ȡ�ɹ�����True�����򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If zlGetServiceObject(objService) = False Then Exit Function
    GetPatiPageInfByRange = objService.ZlCissvr_GetPatiPageInfByRange(cllFilter, rsPatiPageInfo, lngModule)
End Function

