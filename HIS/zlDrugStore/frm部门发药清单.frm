VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm���ŷ�ҩ�嵥 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   8880
      ScaleHeight     =   7185
      ScaleWidth      =   3705
      TabIndex        =   17
      Top             =   120
      Width           =   3735
      Begin VB.PictureBox picHscSend 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3735
         TabIndex        =   25
         Tag             =   "0"
         Top             =   0
         Width           =   3735
         Begin VB.CheckBox chk������� 
            BackColor       =   &H00FFEDDD&
            Caption         =   "����"
            Height          =   180
            Left            =   2520
            TabIndex        =   26
            Top             =   30
            Width           =   735
         End
         Begin VB.Label lblDiag 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "�ٴ����"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   30
            Width           =   2280
         End
      End
      Begin VB.PictureBox Pic��ҩ���� 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3735
         TabIndex        =   23
         Tag             =   "0"
         Top             =   1800
         Width           =   3735
         Begin VB.Label lbl��ҩ���� 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "����ҩ�������Ϣ"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   600
            TabIndex        =   24
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.TextBox txt��ҩ���� 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1215
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   3735
      End
      Begin VB.PictureBox picDoctor 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3735
         TabIndex        =   20
         Tag             =   "0"
         Top             =   3480
         Width           =   3735
         Begin VB.Label lblDoctor 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "����ҽ��ǩ��ͼƬ"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   780
            TabIndex        =   21
            Top             =   0
            Width           =   1440
         End
      End
      Begin VB.PictureBox picǩ��ͼƬ 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000B&
         Height          =   1050
         Left            =   960
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   90
         TabIndex        =   19
         Top             =   3960
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.PictureBox picSign 
         AutoRedraw      =   -1  'True
         Height          =   210
         Left            =   2640
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   18
         Top             =   4200
         Visible         =   0   'False
         Width           =   330
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf��� 
         Height          =   1335
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   3720
         _cx             =   6562
         _cy             =   2355
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm���ŷ�ҩ�嵥.frx":0000
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
   Begin VB.Frame fraH 
      BackColor       =   &H80000007&
      Height          =   5895
      Left            =   6600
      MousePointer    =   9  'Size W E
      TabIndex        =   16
      Top             =   1680
      Width           =   15
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   5535
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   960
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1800
      _cx             =   3175
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
      BackColorSel    =   13487565
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm���ŷ�ҩ�嵥.frx":003D
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
      ExplorerBar     =   3
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   960
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1800
      _cx             =   3175
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
      BackColorSel    =   13487565
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm���ŷ�ҩ�嵥.frx":00B2
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
   Begin MSComctlLib.ImageList imgGroup 
      Left            =   5040
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":0127
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":0281
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":03DB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAssist 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9375
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   9375
      Begin VB.ComboBox cbo�˲��� 
         Height          =   300
         Left            =   3480
         TabIndex        =   15
         Text            =   "cbo�˲���"
         Top             =   60
         Width           =   1900
      End
      Begin VB.ComboBox cbo��ҩ�� 
         Height          =   300
         Left            =   600
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "cbo��ҩ��"
         Top             =   60
         Width           =   1900
      End
      Begin VB.ComboBox cbo��ҩ����ʽ 
         Height          =   300
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   60
         Width           =   1905
      End
      Begin VB.Label lbl�˲��� 
         Caption         =   "�˲���"
         Height          =   180
         Left            =   2880
         TabIndex        =   14
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbl��ҩ�� 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ��"
         Height          =   180
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lbl��ҩ����ʽ 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ����ʽ"
         Height          =   180
         Left            =   6000
         TabIndex        =   9
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4500
      TabIndex        =   0
      Top             =   300
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frm���ŷ�ҩ�嵥.frx":0535
         ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
         Top             =   0
         Width           =   195
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   1508
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   14737632
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm���ŷ�ҩ�嵥.frx":0A83
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   960
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1800
      _cx             =   3175
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
      BackColorSel    =   13487565
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm���ŷ�ҩ�嵥.frx":0AD1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   960
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1800
      _cx             =   3175
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
      BackColorSel    =   13487565
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm���ŷ�ҩ�嵥.frx":0B46
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   960
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1800
      _cx             =   3175
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
      BackColorSel    =   13487565
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm���ŷ�ҩ�嵥.frx":0BBB
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
   Begin MSComctlLib.ImageList imgCheck 
      Left            =   5040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":0C30
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":11CA
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":1764
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfChargeOff 
      Height          =   960
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   1800
      _cx             =   3175
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
      BackColorSel    =   13487565
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm���ŷ�ҩ�嵥.frx":18BE
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   5040
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":1933
            Key             =   "��ӡ11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":1CCD
            Key             =   "��ǰ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":852F
            Key             =   "ָʾ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":ED91
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":F32B
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":F6C5
            Key             =   "��־"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":FA5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":FDF9
            Key             =   "ͼ��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":10193
            Key             =   "ѡ��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":10BA5
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":17407
            Key             =   "δ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":1DC69
            Key             =   "�ڼ�"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":244CB
            Key             =   "�Ѽ�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":2AD2D
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":3158F
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":31929
            Key             =   "����_ѡ��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":31CC3
            Key             =   "�ײ�"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":38525
            Key             =   "����"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":3ED87
            Key             =   "��Ƭ"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":455E9
            Key             =   "����"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":4BE4B
            Key             =   "ָ��"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":526AD
            Key             =   "���"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":58F0F
            Key             =   "������ʽ"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":5F771
            Key             =   "�����ļ�"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":65FD3
            Key             =   "����"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":6C835
            Key             =   "�շ�"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":6D247
            Key             =   "���"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":73AA9
            Key             =   "����"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":7A30B
            Key             =   "ȷ��"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":80B6D
            Key             =   "��ʼ"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":873CF
            Key             =   "����"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":8DC31
            Key             =   "����"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":8DFCB
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":8E365
            Key             =   "�����ܼ�"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":8E6FF
            Key             =   "ȫ���ܼ�"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":8EA99
            Key             =   "�ܼ�"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":8EE33
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":8F845
            Key             =   "�Ѿ���ӡ"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":90257
            Key             =   "ҩƷ"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ŷ�ҩ�嵥.frx":96AB9
            Key             =   "��Σ"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   3840
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frm���ŷ�ҩ�嵥"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mlngMode As Long

Private mblnOutPut As Boolean

Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

'�������˵�
Private Const conMenu_Tool_ShowShortage = 101       '��ʾȱҩ
'Private Const conMenu_Tool_ShowRefuse = 102         '��ʾ�ܷ�
Private Const conMenu_Tool_ShowReturnSend = 103     '��ʾ��ҩ����
Private Const conMenu_Tool_SumByBatch = 104         '�����λ���
Private Const conMenu_Tool_ShowAllProcess = 105     '��ʾ���й��̵���
Private Const conMenu_Tool_ShowPlug = 106           '���ò����������ҩ
Private Const conMenu_Tool_ShowInfo = 107           '��ʾ��չ��Ϣ

'�����˵�
'��ҩʱ
Private Const conMenu_StatusPopup = 3                 '����״̬
Private Const conMenu_Status_Verify = 301             '��ҩ
Private Const conMenu_Status_Reject = 303             '�ܷ�ȷ��
Private Const conMenu_Status_Return = 304             '��ҩ
Private Const conMenu_Status_RefuseRestore = 308      '�ܷ��ָ�
Private Const conMenu_Status_Shortage = 309           'ȱҩ
Private Const conMenu_Status_NoProcess = 310          '������
Private Const conMenu_Status_AllSend = 311            'ȫ����ҩ
Private Const conMenu_Status_AllReject = 312          'ȫ���ܷ�
Private Const conMenu_Status_AllNoProcess = 313       'ȫ��������
'��ҩʱ
Private Const conMenu_Status_AllReturn = 321          'ȫ����ҩ
Private Const conMenu_Status_AllCancel = 322          'ȫ��ȡ����ҩ
'ҩƷ����
Private Const conMenu_MediPopup = 4                   'ҩƷ������ʾ
Private Const conMenu_Medi_CodeAddName = 401          '��ʾ���������
Private Const conMenu_Medi_Code = 402                 '��ʾ����
Private Const conMenu_Medi_Name = 403                 '��ʾ����

'��������
Private mdblSumListHeight As Double                 '��¼���ܷ�ҩ�б�ԭʼ�ĸ߶�
Private mdblSendListHeight As Double                 '��¼��ҩ�б�ԭʼ�ĸ߶�
Private mblnResize As Boolean

Private mstrFindCondition As String                 '��������

Private mblnShowReject As Boolean

'���ݼ�
Private mrsSendList As ADODB.Recordset              '��ҩ���ݼ�
Private mrsChargeOff As ADODB.Recordset             '�������ݼ�
Private mrsReturnList As ADODB.Recordset            '��ҩ���ݼ�

'�б���ʾ����
Private Type Type_ShowListCondition
    intListType As Integer                          '0-δ��;1-����;2-ȱҩ;3-�ܷ�;4-�ѷ�
    bln�����λ��� As Boolean
    bln�����һ��� As Boolean
    intShowPass As Integer                           '�Ƿ���ʾ������ҩ��PASS��
    blnҽ����ѯ As Boolean
    bln��ʾ��ҩ�������� As Boolean
    bln��ʾ��չ��Ϣ As Boolean
    bln��ʾȱҩ As Boolean
    bln��ʾ���̵��� As Boolean
    bln��ʾ������� As Boolean
    lngҩ��id As Long
    bln�޸��������� As Boolean
    bln������ҩ���� As Boolean
    blnҩƷ���� As Boolean
    bln����δ��˴�����ҩ As Boolean
    intҩƷ���Ʊ�����ʾ As Integer
    int��ҩ����ʽ As Integer
    bln�������� As Boolean
    str��Σ���� As String
    str��Σ���� As String
    int��ҩ��������Ĭ��Ϊ��ҩ״̬ As Integer
    bln��ʾԭ���� As Boolean
End Type
Private mcondition As Type_ShowListCondition

'���±�־
Private mblnRefresh As Boolean                      'ˢ�±�־
Private mblnSendChange As Boolean                   '����ҩ�嵥�е�״̬�����仯
Private mblnDrop As Boolean                     '��KeyDown���ж������б��Ƿ񵯳�
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

'�б�����
Private Enum mListType
    ��ҩ = 0
    ���� = 1
    ȱҩ = 2
    �ܷ� = 3
    ��ҩ = 4
End Enum

'ִ��״̬
Private Enum mState
    ȱҩ = 0
    ��ҩ = 1
    �ܷ� = 2
    ������ = 3
    �ܷ�_�ָ� = 4
    �ܷ�_������ = 5
    ��ҩ = 6
    ��ҩ_ԭʼ��¼ = 7
    ��ҩ_��ҩ��¼ = 8
    ��ҩ_��ҩ��¼ = 9
    ת����¼ = 10
End Enum

'ͨ���˵��ı�״̬�Ĵ���
Private Enum mChangeState
    ��ҩ = 0
    �ܷ� = 1
    ȱҩ = 2
    ������ = 3
    ȫ����ҩ = 4
    ȫ���ܷ� = 5
    ȫ�������� = 6
End Enum

'���ݻ�������
Private Enum mSubTotalType
    SubSum = 0                  '�ϼ�
    SubByDept = 1               '����ҩ����С��
    SubByPeople = 2             '������С��
    SubByNo = 3                 '������С��
    SubByDrug = 4               '��ҩƷС��
    SubByHosNumber = 5          '��סԺ��С��
    SubByBedNumber = 6          '������С��
    SubByPeopleDept = 7         '�����˿���
End Enum

Private mstrUnallowSetColHide(0 To 4) As String         '�������������ص���
Private mstrUnallowShow(0 To 4) As String                   '��������ʾ����

'δ��ҩ�б�
Private Const mconIntCol��ҩ_���� As Integer = 61
Private mIntCol��ҩ_��ǰ�� As Integer
Private mIntCol��ҩ_����� As Integer
Private mIntCol��ҩ_����� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_����ҽ�� As Integer
Private mIntCol��ҩ_״̬ As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_��ҩ���� As Integer
Private mIntCol��ҩ_NO As Integer
Private mIntCol��ҩ_����Ա As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_�������� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_�Ա� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_סԺ�� As Integer
Private mIntCol��ҩ_Ʒ�� As Integer
Private mIntCol��ҩ_Ƥ�Խ�� As Integer
Private mIntCol��ҩ_������ As Integer
Private mIntCol��ҩ_Ӣ���� As Integer
Private mIntCol��ҩ_�䷽���� As Integer
Private mIntCol��ҩ_��� As Integer
Private mIntCol��ҩ_������ As Integer
Private mIntCol��ҩ_ԭ���� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_Ч�� As Integer
Private mIntCol��ҩ_�� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_��� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_Ƶ�� As Integer
Private mIntCol��ҩ_�÷� As Integer
Private mIntCol��ҩ_��ҩ���� As Integer
Private mIntCol��ҩ_��ҩĿ�� As Integer
Private mIntCol��ҩ_����ʱ�� As Integer
Private mIntCol��ҩ_˵�� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_ҽ��id As Integer
Private mIntCol��ҩ_��ҩ�� As Integer
Private mIntCol��ҩ_�ⷿ��λ As Integer
Private mIntCol��ҩ_���ID As Integer
Private mIntCol��ҩ_ҩƷID As Integer
Private mIntCol��ҩ_������λ As Integer
Private mIntCol��ҩ_��ҩ���� As Integer
Private mIntCol��ҩ_��ҩ����id As Integer
Private mIntCol��ҩ_ҩƷ��������� As Integer
Private mIntCol��ҩ_ҩƷ���� As Integer
Private mIntCol��ҩ_ҩƷ���� As Integer
Private mIntCol��ҩ_�շ�ID As Integer
Private mIntCol��ҩ_ִ��״̬ As Integer
Private mIntCol��ҩ_��ҩ�� As Integer
Private mIntCol��ҩ_���շ� As Integer
Private mIntCol��ҩ_����ID As Integer
Private mIntCol��ҩ_��ҳID As Integer
Private mIntCol��ҩ_��ҩ���� As Integer
Private mIntCol��ҩ_��ΣҩƷ As Integer
Private mIntCol��ҩ_��� As Integer
Private mIntCol��ҩ_����ҩƷ˵�� As Integer
Private mIntCol��ҩ_��������id As Integer
Private mIntCol��ҩ_��ע As Integer

'�����б�
Private Const mconIntCol����_���� As Integer = 14
Private mIntCol����_��ǰ�� As Integer
Private mIntCol����_Ʒ�� As Integer
Private mIntCol����_��� As Integer
Private mIntCol����_������ As Integer
Private mIntCol����_ԭ���� As Integer
Private mIntCol����_���� As Integer
Private mIntCol����_Ч�� As Integer
Private mIntCol����_���� As Integer
Private mIntCol����_��λ As Integer
Private mIntCol����_���� As Integer
Private mIntCol����_��� As Integer
Private mIntCol����_ҩƷ��������� As Integer
Private mIntCol����_ҩƷ���� As Integer
Private mIntCol����_ҩƷ���� As Integer

Private Const mconIntCol���һ���_���� As Integer = 25
Private mIntCol���һ���_��ǰ�� As Integer
Private mIntCol���һ���_���� As Integer
Private mIntCol���һ���_Ʒ�� As Integer
Private mIntCol���һ���_��� As Integer
Private mIntCol���һ���_������ As Integer
Private mIntCol���һ���_ԭ���� As Integer
Private mIntCol���һ���_���� As Integer
Private mIntCol���һ���_Ч�� As Integer
Private mIntCol���һ���_Ӧ������ As Integer
Private mIntCol���һ���_�������� As Integer
Private mIntCol���һ���_�������� As Integer
Private mIntCol���һ���_ʵ������ As Integer
Private mIntCol���һ���_��λ As Integer
Private mIntCol���һ���_���� As Integer
Private mIntCol���һ���_Ӧ����� As Integer
Private mIntCol���һ���_ʵ����� As Integer
Private mIntCol���һ���_���� As Integer
Private mIntCol���һ���_����ID As Integer
Private mIntCol���һ���_ҩƷID As Integer
Private mIntCol���һ���_��ҩ���� As Integer
Private mIntCol���һ���_��ҩ����id As Integer
Private mIntCol���һ���_ҩƷ��������� As Integer
Private mIntCol���һ���_ҩƷ���� As Integer
Private mIntCol���һ���_ҩƷ���� As Integer
Private mIntCol���һ���_��װ As Integer

'�����б�
Private Const mconIntCol����_���� As Integer = 14
Private mIntCol����_��ǰ�� As Integer
Private mIntCol����_������� As Integer
Private mIntCol����_���� As Integer
Private mIntCol����_NO As Integer
Private mIntCol����_ҩƷID As Integer
Private mIntCol����_����ʱ�� As Integer
Private mIntCol����_�շ���� As Integer
Private mIntCol����_������ As Integer
Private mIntCol����_���� As Integer
Private mIntCol����_Ч�� As Integer
Private mIntCol����_׼������ As Integer
Private mIntCol����_�������� As Integer
Private mIntCol����_��װ As Integer
Private mIntCol����_��λ As Integer

'ȱҩ�б�
Private Const mconIntColȱҩ_���� As Integer = 21
Private mIntColȱҩ_��ǰ�� As Integer
Private mIntColȱҩ_���� As Integer
Private mIntColȱҩ_NO As Integer
Private mIntColȱҩ_���� As Integer
Private mIntColȱҩ_��ҩ���� As Integer
Private mIntColȱҩ_���� As Integer
Private mIntColȱҩ_���� As Integer
Private mIntColȱҩ_�Ա� As Integer
Private mIntColȱҩ_Ʒ�� As Integer
Private mIntColȱҩ_��� As Integer
Private mIntColȱҩ_������ As Integer
Private mIntColȱҩ_ԭ���� As Integer
Private mIntColȱҩ_���� As Integer
Private mIntColȱҩ_Ч�� As Integer
Private mIntColȱҩ_���� As Integer
Private mIntColȱҩ_���� As Integer
Private mIntColȱҩ_��� As Integer
Private mIntColȱҩ_ҩƷ��������� As Integer
Private mIntColȱҩ_ҩƷ���� As Integer
Private mIntColȱҩ_ҩƷ���� As Integer
Private mIntColȱҩ_��ע As Integer

'�ܷ��б�
Private Const mconIntCol�ܷ�_���� As Integer = 24
Private mIntCol�ܷ�_��ǰ�� As Integer
Private mIntCol�ܷ�_���� As Integer
Private mIntCol�ܷ�_״̬ As Integer
Private mIntCol�ܷ�_NO As Integer
Private mIntCol�ܷ�_���� As Integer
Private mIntCol�ܷ�_��ҩ���� As Integer
Private mIntCol�ܷ�_���� As Integer
Private mIntCol�ܷ�_���� As Integer
Private mIntCol�ܷ�_�Ա� As Integer
Private mIntCol�ܷ�_Ʒ�� As Integer
Private mIntCol�ܷ�_��� As Integer
Private mIntCol�ܷ�_������ As Integer
Private mIntCol�ܷ�_ԭ���� As Integer
Private mIntCol�ܷ�_���� As Integer
Private mIntCol�ܷ�_Ч�� As Integer
Private mIntCol�ܷ�_���� As Integer
Private mIntCol�ܷ�_���� As Integer
Private mIntCol�ܷ�_��� As Integer
Private mIntCol�ܷ�_ҩƷ��������� As Integer
Private mIntCol�ܷ�_ҩƷ���� As Integer
Private mIntCol�ܷ�_ҩƷ���� As Integer
Private mIntCol�ܷ�_ִ��״̬ As Integer
Private mIntCol�ܷ�_�շ�ID As Integer
Private mIntCol�ܷ�_��ע As Integer

'��ҩ�б�
Private Const mconIntCol��ҩ_���� As Integer = 50
Private mIntCol��ҩ_��ǰ�� As Integer
Private mIntCol��ҩ_����� As Integer
Private mIntCol��ҩ_����� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_״̬ As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_��ҩ���� As Integer
Private mIntCol��ҩ_NO As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_�Ա� As Integer
Private mIntCol��ҩ_סԺ�� As Integer
Private mIntCol��ҩ_Ʒ�� As Integer
Private mIntCol��ҩ_������ As Integer
Private mIntCol��ҩ_Ӣ���� As Integer
Private mIntCol��ҩ_��� As Integer
Private mIntCol��ҩ_������ As Integer
Private mIntCol��ҩ_ԭ���� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_Ч�� As Integer
Private mIntCol��ҩ_�� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_������ As Integer
Private mIntCol��ҩ_׼���� As Integer
Private mIntCol��ҩ_��ҩ�� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_��� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_Ƶ�� As Integer
Private mIntCol��ҩ_�÷� As Integer
Private mIntCol��ҩ_����Ա As Integer
Private mIntCol��ҩ_��ҩʱ�� As Integer
Private mIntCol��ҩ_���� As Integer
Private mIntCol��ҩ_ҽ��id As Integer
Private mIntCol��ҩ_��ҩ�� As Integer
Private mIntCol��ҩ_�ⷿ��λ As Integer
Private mIntCol��ҩ_���ID As Integer
Private mIntCol��ҩ_ҩƷID As Integer
Private mIntCol��ҩ_������λ As Integer
Private mIntCol��ҩ_ҩƷ��������� As Integer
Private mIntCol��ҩ_ҩƷ���� As Integer
Private mIntCol��ҩ_ҩƷ���� As Integer
Private mIntCol��ҩ_�շ�ID As Integer
Private mIntCol��ҩ_ִ��״̬ As Integer
Private mIntCol��ҩ_��ҩ�� As Integer
Private mIntCol��ҩ_��ҩ����id As Integer
Private mIntCol��ҩ_����ʱ�� As Integer
Private mIntCol��ҩ_��ע As Integer
Private mIntCol��ҩ_����ID As Integer
Private mIntCol��ҩ_��ҳID As Integer

Public Sub SetSendBillStateByCustom(ByVal str�շ�ids As String)
    '�Զ�����˹��ܣ����ݷ��ص����ݸ��½��淢ҩ״̬��ȡ����ҩ��
    Dim intState As Integer
    Dim strState As String
    Dim lngColor As Long
    Dim i As Long
    Dim lng���ID As Long
    Dim strNo As String
    
    If mcondition.intListType <> mListType.��ҩ Then Exit Sub
    If str�շ�ids = "" Then Exit Sub
    
    With vsfList(mListType.��ҩ)
        intState = mState.������
        strState = "������"
        lngColor = mListColor.State_UnProcess
        
        .Redraw = flexRDNone
        For i = 1 To .rows - 1
            If .IsSubtotal(i) = False And Val(.TextMatrix(i, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ _
                And InStr("," & str�շ�ids & ",", "," & Val(.TextMatrix(i, mIntCol��ҩ_�շ�ID)) & ",") > 0 Then
                .TextMatrix(i, mIntCol��ҩ_ִ��״̬) = intState
                .TextMatrix(i, mIntCol��ҩ_״̬) = strState
                
                .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                
                mrsSendList.Filter = "�շ�ID=" & Val(.TextMatrix(i, mIntCol��ҩ_�շ�ID))
                
                mrsSendList!ִ��״̬ = intState
                mrsSendList!״̬ = strState
                
                mrsSendList.Update
                
                mblnSendChange = True
            End If
        Next
        
        .Redraw = flexRDDirect
    End With
End Sub
Private Sub GetDiagnosis(ByVal lngRow As Long)
    '���������еĲ���ID�õ����˵���ϼ�¼
    Dim strTmp As String
    Dim i As Integer
    
    
    With vsf���
        
        If vsfList(mListType.��ҩ).IsSubtotal(lngRow) = True Then 'Or vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_��ҩ����) = "" Then
            .rows = 1
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = ""
            lblDiag.Caption = "�ٴ����"
            .Tag = ""
            Exit Sub
        End If
        
        If Val(.Tag) = Val(vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_����ID)) Then Exit Sub
        
        strTmp = RecipeSendWork_GetDiagnosis(IIf(chk�������.Value = 1, 3, 2), Val(vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_����ID)), Val(vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_��ҳID)))
    
        .Tag = vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_����ID)
        
        lblDiag.Caption = "[" & vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_����) & "]���ٴ����"
        
        .Redraw = flexRDNone
        .rows = 1
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = ""
            
        If strTmp <> "" Then
            strTmp = strTmp & "|"
            For i = 0 To UBound(Split(strTmp, "|"))
                If Split(strTmp, "|")(i) <> "" Then
                    If i > 0 Then .rows = .rows + 1
                    .TextMatrix(i, 0) = Split(Split(strTmp, "|")(i), ",")(0)
                    .TextMatrix(i, 1) = Split(Split(strTmp, "|")(i), ",")(1)
                End If
            Next
        End If
        .Redraw = flexRDDirect
        
        If .TextMatrix(0, 0) = "" Then
            lblDiag.Caption = lblDiag.Caption & "(0)"
        Else
            lblDiag.Caption = lblDiag.Caption & "(" & .rows & ")"
        End If
    End With
End Sub

Public Function GetRecordInfo() As String
    '���ص�ǰ��¼����Ϣ
    '���أ�����|NO|����ID
    
    If mcondition.intListType = mListType.��ҩ Or mcondition.intListType = mListType.��ҩ Then
        With vsfList(mcondition.intListType)
            If .Row = 0 Then Exit Function
            If .IsSubtotal(.Row) = True Then Exit Function
            If .TextMatrix(.Row, .ColIndex("�շ�ID")) = "" Then Exit Function
            
            GetRecordInfo = .TextMatrix(.Row, .ColIndex("����")) & "|" & .TextMatrix(.Row, .ColIndex("NO")) & "|" & .TextMatrix(.Row, .ColIndex("ҩƷID"))
        End With
    End If
End Function
Public Sub ClearList(ByVal intType As Integer)
    '����б�
    
    Select Case intType
        Case mListType.��ҩ
            Set mrsSendList = Nothing
            Set mrsChargeOff = Nothing
            
            vsfList(mListType.��ҩ).rows = 1
            vsfList(mListType.��ҩ).rows = 2
            vsfList(mListType.����).rows = 1
            vsfList(mListType.����).rows = 2
            vsfList(mListType.�ܷ�).rows = 1
            vsfList(mListType.�ܷ�).rows = 2
            vsfList(mListType.ȱҩ).rows = 1
            vsfList(mListType.ȱҩ).rows = 2
            
            vsfChargeOff.rows = 1
            Me.picǩ��ͼƬ.Visible = False
            Me.txt��ҩ����.Text = ""
            vsf���.rows = 0
            vsf���.rows = 4
        Case mListType.��ҩ
            Set mrsReturnList = Nothing
            
            vsfList(mListType.��ҩ).rows = 1
            vsfList(mListType.��ҩ).rows = 2
    End Select
End Sub
Public Function GetPrintObject(ByVal blnOutPut As Boolean) As Object
    mblnOutPut = blnOutPut
    If vsfList(mcondition.intListType).rows = 1 Then
        Set GetPrintObject = Nothing
    Else
        Set GetPrintObject = vsfList(mcondition.intListType)
    End If
End Function

Private Sub Modify��ҩ����(ByVal blnShow As Boolean)
    '�л���ʾ��ҩ����״̬ʱ�������ʾ����¼�¼Ϊ��ҩ״̬���������ʾ�����Ϊ������״̬
    
    If mrsSendList Is Nothing Then Exit Sub

    mrsSendList.Filter = "��¼״̬>1"
    If blnShow = True Then
        Do While Not mrsSendList.EOF
            If mrsSendList!ִ��״̬ = mState.������ And mcondition.int��ҩ��������Ĭ��Ϊ��ҩ״̬ = 1 Then
                If mrsSendList!��ΣҩƷ = 0 Or (mrsSendList!��ΣҩƷ > 0 And InStr(1, mcondition.str��Σ����, mrsSendList!��ΣҩƷ) = 0) And Not (InStr("," & mcondition.str��Σ���� & ",", "," & mrsSendList!��ΣҩƷ & ",") > 0 And mrsSendList!��ΣҩƷ > 0) Then
                    mrsSendList!״̬ = "��ҩ"
                    mrsSendList!ִ��״̬ = mState.��ҩ
                    mrsSendList.Update
                End If
            End If
            mrsSendList.MoveNext
        Loop
    Else
        Do While Not mrsSendList.EOF
            If mrsSendList!ִ��״̬ = mState.��ҩ Then
                mrsSendList!״̬ = "������"
                mrsSendList!ִ��״̬ = mState.������
                mrsSendList.Update
            End If
            mrsSendList.MoveNext
        Loop
    End If
End Sub

Public Sub SetAllReturn()
    '��ҩ״̬ʱ������ȫ��
    Dim n As Long
    
    If mcondition.intListType <> mListType.��ҩ Then Exit Sub
    
    With vsfList(mListType.��ҩ)
        For n = 1 To .rows - 1
            If .IsSubtotal(n) = False Then
                If Val(.TextMatrix(n, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ_ԭʼ��¼ And Val(.TextMatrix(n, mIntCol��ҩ_׼����)) > 0 Then
                    .TextMatrix(n, mIntCol��ҩ_��ҩ��) = Val(.TextMatrix(n, mIntCol��ҩ_׼����))
                    .TextMatrix(n, mIntCol��ҩ_״̬) = "��ҩ"
                    .TextMatrix(n, mIntCol��ҩ_ִ��״̬) = mState.��ҩ
                    
                    mrsReturnList.Filter = "�շ�ID=" & Val(.TextMatrix(n, mIntCol��ҩ_�շ�ID))
                    
                    mrsReturnList!״̬ = .TextMatrix(n, mIntCol��ҩ_״̬)
                    mrsReturnList!ִ��״̬ = Val(.TextMatrix(n, mIntCol��ҩ_ִ��״̬))
                    mrsReturnList!��ҩ�� = Val(.TextMatrix(n, mIntCol��ҩ_��ҩ��))
                    mrsReturnList.Update
                End If
            End If
        Next
    End With
End Sub


Public Sub SetAllNotReturn()
    '��ҩ״̬ʱ������ȫ����
    Dim n As Long
    
    If mcondition.intListType <> mListType.��ҩ Then Exit Sub
    
    With vsfList(mListType.��ҩ)
        For n = 1 To .rows - 1
            If .IsSubtotal(n) = False Then
                If Val(.TextMatrix(n, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ Then
                    .TextMatrix(n, mIntCol��ҩ_��ҩ��) = ""
                    .TextMatrix(n, mIntCol��ҩ_״̬) = "������"
                    .TextMatrix(n, mIntCol��ҩ_ִ��״̬) = mState.��ҩ_ԭʼ��¼
                    
                    mrsReturnList.Filter = "�շ�ID=" & Val(.TextMatrix(n, mIntCol��ҩ_�շ�ID))
                    
                    mrsReturnList!״̬ = .TextMatrix(n, mIntCol��ҩ_״̬)
                    mrsReturnList!ִ��״̬ = Val(.TextMatrix(n, mIntCol��ҩ_ִ��״̬))
                    mrsReturnList!��ҩ�� = Val(.TextMatrix(n, mIntCol��ҩ_��ҩ��))
                    mrsReturnList.Update
                End If
            End If
        Next
    End With
End Sub

Public Sub SetFontSize(ByVal intFont As Integer)
    Dim objVSF As VSFlexGrid
    
    For Each objVSF In vsfList
        objVSF.Font.Size = intFont
        Me.Font.Size = objVSF.Font.Size
        objVSF.Cell(flexcpFontSize, 0, 0, objVSF.rows - 1, objVSF.Cols - 1) = objVSF.Font.Size
        
        objVSF.RowHeightMin = TextHeight("��") + 100
        objVSF.RowHeightMax = TextHeight("��") + 100
        objVSF.Refresh
    Next
End Sub
Public Sub AfterSendRefresh()
    '��ҩ����·�ҩ���ݼ�
    
    'ɾ���ѷ�ҩ�ļ�¼
    If Not mrsSendList Is Nothing Then
        With mrsSendList
            .Filter = "ִ��״̬=1"
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
            .UpdateBatch
            
            .Filter = ""
        End With
    End If
    
    If Not mrsChargeOff Is Nothing Then
        With mrsChargeOff
            .Filter = "ִ�б�־=1"
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
            .UpdateBatch
            
            .Filter = ""
        End With
    End If
    
    '������ϸ����
    RefreshList mListType.��ҩ, mrsSendList, mrsChargeOff
End Sub

Public Sub AfterReturnRefresh()
    '��ҩ����·�ҩ���ݼ�
    
    'ɾ���ѷ�ҩ�ļ�¼
    With mrsReturnList
        .Filter = "ִ��״̬=" & mState.��ҩ
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
        .UpdateBatch
        
        .Filter = ""
    End With
    
    '������ϸ����
    RefreshList mListType.��ҩ, mrsReturnList
End Sub
Public Sub AfterRejectRefresh()
    '�ܷ�����·�ҩ���ݼ�
    
    '�޸ľܷ�ҩ�ļ�¼
    With mrsSendList
        .Filter = "ִ��״̬=" & mState.�ܷ�
        Do While Not .EOF
            !ִ��״̬ = mState.�ܷ�_������
            !״̬ = "������"
            
            .MoveNext
        Loop
        .UpdateBatch
        
        .Filter = ""
    End With
    
    '������ϸ����
    RefreshList mListType.��ҩ, mrsSendList
End Sub

Public Sub AfterRejectRestoreRefresh()
    '�ܷ��ָ�����·�ҩ���ݼ�
    
    '�޸ľܷ��ָ��ļ�¼
    With mrsSendList
        .Filter = "ִ��״̬=" & mState.�ܷ�_�ָ�
        Do While Not .EOF
            !ִ��״̬ = mState.��ҩ
            !״̬ = "��ҩ"
            
            .MoveNext
        Loop
        .UpdateBatch
        
        .Filter = ""
    End With
    
    '������ϸ����
    RefreshList mListType.��ҩ, mrsSendList
End Sub
Public Sub FindRecord(ByVal intType As Integer, Optional ByVal strFind As String = "")
    Dim lng�շ�ID As Long
    Dim int�շ�ID�� As Integer
    Dim strFilter As String
    
    '�����������Ϊ�գ������ϴβ�������ҲΪ�գ����˳�
    If strFind = "" And mstrFindCondition = "" Then Exit Sub
     
    If strFind <> "" And strFind <> mstrFindCondition Then
        '����������Ϊ�գ����Ҳ������ϴβ��������������¹������ݼ�
        mstrFindCondition = strFind
        If intType = mListType.��ҩ Then
            If mrsReturnList Is Nothing Then Exit Sub
            
            mrsReturnList.Filter = mstrFindCondition
            If mrsReturnList.RecordCount = 0 Then Exit Sub
            
            lng�շ�ID = mrsReturnList!�շ�ID
            int�շ�ID�� = mIntCol��ҩ_�շ�ID
        Else
            If mrsSendList Is Nothing Then Exit Sub
            
            mrsSendList.Filter = mstrFindCondition
            If mrsSendList.RecordCount = 0 Then Exit Sub
            
            lng�շ�ID = mrsSendList!�շ�ID
            int�շ�ID�� = mIntCol��ҩ_�շ�ID
        End If
    Else
        '��������Ϊ�գ����ߵ����ϴβ��������������ѹ��˷�Χ�в���������¼
        If intType = mListType.��ҩ Then
            If mrsReturnList Is Nothing Then Exit Sub
            
            If mrsReturnList.RecordCount = 0 Then Exit Sub
                
            int�շ�ID�� = mIntCol��ҩ_�շ�ID
            
            mrsReturnList.MoveNext
            
            If Not mrsReturnList.EOF Then
                lng�շ�ID = mrsReturnList!�շ�ID
            Else
                mrsReturnList.MoveFirst
                lng�շ�ID = mrsReturnList!�շ�ID
            End If
                    
        Else
            If mrsSendList Is Nothing Then Exit Sub
            
            If mrsSendList.RecordCount = 0 Then Exit Sub
            
            int�շ�ID�� = mIntCol��ҩ_�շ�ID
            
            mrsSendList.MoveNext
            
            If Not mrsSendList.EOF Then
                lng�շ�ID = mrsSendList!�շ�ID
            Else
                mrsSendList.MoveFirst
                lng�շ�ID = mrsSendList!�շ�ID
            End If
                
        End If
    End If
    
    '���ݲ��ҵ����շ�ID���ڱ���ж�λ
    With vsfList(intType)
        .Row = .FindRow(lng�շ�ID, 1, int�շ�ID��)
    End With
End Sub

Public Function GetReturnDate() As String
    'ȡ����ҩ��������
    '���أ�����
    
    With vsfList(mListType.��ҩ)
        If .Row = 0 Then Exit Function
        If .IsSubtotal(.Row) = True Then Exit Function
        If .TextMatrix(.Row, mIntCol��ҩ_��ҩʱ��) = "" Then Exit Function

        GetReturnDate = .TextMatrix(.Row, mIntCol��ҩ_��ҩʱ��)
    End With
End Function

Public Function GetSendedInfo() As String
    '�����ѷ�ҩ������Ϣ����ҩ����|��ҩ����ID|���ܷ�ҩ��
    
    With vsfList(mListType.��ҩ)
        If .Row = 0 Then Exit Function
        If .IsSubtotal(.Row) = True Then Exit Function
        If .TextMatrix(.Row, mIntCol��ҩ_��ҩ��) = "" Then Exit Function

        GetSendedInfo = .TextMatrix(.Row, mIntCol��ҩ_����) & "|" & .TextMatrix(.Row, mIntCol��ҩ_��ҩ����id) & "|" & .TextMatrix(.Row, mIntCol��ҩ_��ҩ��)
    End With
End Function
Public Function GetSendRecord() As ADODB.Recordset
    '�����������淵�ط�ҩ��¼��
    
    If mrsSendList Is Nothing Then
        Set GetSendRecord = Nothing
        Exit Function
    Else
        mrsSendList.Filter = ""
        Set GetSendRecord = mrsSendList
    End If
End Function

Public Function GetReturnRecord() As ADODB.Recordset
    '�����������淵����ҩ��¼��
    
    If mrsReturnList Is Nothing Then
        Set GetReturnRecord = Nothing
    Else
        mrsReturnList.Filter = ""
        Set GetReturnRecord = mrsReturnList
    End If
End Function

Public Function GetChargeOffRecord() As ADODB.Recordset
    '�����������淵�����ʼ�¼��
    Dim i As Integer
    Dim lng��ҩ����ID As Long
    Dim lngҩƷid As Long
    
    With vsfList(mListType.����)
        If mcondition.bln�����һ��� = True Then
            For i = 1 To .rows - 1
                lng��ҩ����ID = Val(.TextMatrix(i, mIntCol���һ���_��ҩ����id))
                lngҩƷid = Val(.TextMatrix(i, mIntCol���һ���_ҩƷID))
                 
                If Not mrsChargeOff Is Nothing Then
                    mrsChargeOff.Filter = "��˱�־>0 And ��ҩ����ID=" & lng��ҩ����ID & " And ҩƷID=" & lngҩƷid
                    Do While Not mrsChargeOff.EOF
                        mrsChargeOff!ִ�б�־ = 1
                        mrsChargeOff.Update
                        mrsChargeOff.MoveNext
                    Loop
                End If
            Next
        End If
    End With
    
    Set GetChargeOffRecord = mrsChargeOff
End Function

Public Function GetStayRecord() As ADODB.Recordset
    '�����������淵�������¼��
    Dim i As Integer
    Dim rsStay As ADODB.Recordset                  '�������ݼ�
    
    Set rsStay = New ADODB.Recordset
    With rsStay
        If .State = 1 Then .Close
        .Fields.Append "��ҩ����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    If mcondition.bln�����һ��� = True Then
        With vsfList(mListType.����)
            For i = 1 To .rows - 1
                If Val(.TextMatrix(i, mIntCol���һ���_��ҩ����id)) > 0 And Val(.TextMatrix(i, mIntCol���һ���_��������)) > 0 Then
                    rsStay.AddNew
                    rsStay!��ҩ����ID = Val(.TextMatrix(i, mIntCol���һ���_��ҩ����id))
                    rsStay!ҩƷID = Val(.TextMatrix(i, mIntCol���һ���_ҩƷID))
                    rsStay!���� = Val(.TextMatrix(i, mIntCol���һ���_����))
                    rsStay!�������� = Val(.TextMatrix(i, mIntCol���һ���_��������)) * Val(.TextMatrix(i, mIntCol���һ���_��װ))
                    rsStay!���� = Val(.TextMatrix(i, mIntCol���һ���_����)) / Val(.TextMatrix(i, mIntCol���һ���_��װ))
                    rsStay.Update
                End If
            Next
        End With
    End If
    
    Set GetStayRecord = rsStay
End Function

Public Function Get��ǰ��ҩ����ʽ() As Integer
    Get��ǰ��ҩ����ʽ = IIf(cbo��ҩ����ʽ.ListIndex = -1, 1, cbo��ҩ����ʽ.ListIndex + 1)
End Function

Public Function Get��ǰ��ҩ��() As String
    '�����������ڷ�����ҩ��
    
    If InStr(cbo��ҩ��.Text, "-") > 0 Then
        Get��ǰ��ҩ�� = Mid(cbo��ҩ��.Text, InStr(cbo��ҩ��.Text, "-") + 1)
    Else
        Get��ǰ��ҩ�� = cbo��ҩ��.Text
    End If
End Function

Public Function Get��ǰ�˲���() As String
    '�����������ڷ�����ҩ��
    
    If InStr(cbo�˲���.Text, "-") > 0 Then
        Get��ǰ�˲��� = Mid(cbo�˲���.Text, InStr(cbo�˲���.Text, "-") + 1)
    Else
        Get��ǰ�˲��� = cbo�˲���.Text
    End If
End Function


Private Sub InitColSelList(ByVal intListType As Integer)
    Dim i As Integer
    
    With vsfColSel
        .Tag = intListType
        
        .rows = .FixedRows
        For i = 1 To vsfList(intListType).Cols - 1
            '���ڲ�������ʾ�б���в��ܼ�����ѡ���б�
            If IsInString(mstrUnallowShow(intListType), vsfList(intListType).ColKey(i), ";") = False Then
                If (mcondition.bln��ʾԭ���� And vsfList(intListType).ColKey(i) = "ԭ����") Or vsfList(intListType).ColKey(i) <> "ԭ����" Then
                    .rows = .rows + 1
                    .TextMatrix(.rows - 1, 1) = vsfList(intListType).ColKey(i)
                    .RowData(.rows - 1) = i
                End If
                
                '�п�Ϊ�ջ������ص�������Ϊ����ѡ
                If Not (vsfList(intListType).ColWidth(i) = 0 Or vsfList(intListType).ColHidden(i)) Then
                    .TextMatrix(.rows - 1, 0) = 0
                End If
                
                'ָ����������Ϊ������������
                If IsInString(mstrUnallowSetColHide(intListType), vsfList(intListType).ColKey(i), ";") = True Then
                    .Cell(flexcpForeColor, .rows - 1, 1) = .BackColorFixed
                End If
            End If
        Next
    End With
End Sub

Private Sub InitList(ByVal intType As Integer)
    '���ݲ�����ʼ���б�
    
    Select Case intType
        Case mListType.��ҩ
            Call InitList_Send
        Case mListType.����
            Call InitList_Sum
            Call InitList_ChargeOff
        Case mListType.ȱҩ
            Call InitList_Shortage
        Case mListType.�ܷ�
            Call InitList_Reject
        Case mListType.��ҩ
            Call InitList_Return
        Case Else
            Call InitList_Send
            Call InitList_Sum
            Call InitList_ChargeOff
            Call InitList_Shortage
            Call InitList_Reject
            Call InitList_Return
    End Select
End Sub

Private Sub SaveListColState(Optional intType As Integer = -1)
    Dim str������ As String
    Dim i As Integer
    Dim strType As String
    Dim n As Integer
    Dim intStart As Integer
    Dim intEnd As Integer
    
    If Val(zlDataBase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    
    If intType = -1 Then
        intStart = 0
        intEnd = vsfList.count - 1
    Else
        intStart = intType
        intEnd = intType
    End If
    
    For n = intStart To intEnd
        Select Case n
            Case mListType.��ҩ
                strType = "��ҩ"
            Case mListType.����
                If mcondition.bln�����һ��� = True Then
                    strType = "���һ���"
                Else
                    strType = "����"
                End If
            Case mListType.ȱҩ
                strType = "ȱҩ"
            Case mListType.�ܷ�
                strType = "�ܷ�"
            Case mListType.��ҩ
                strType = "��ҩ"
        End Select
        
        str������ = ""
        With vsfList(n)
            For i = 0 To .Cols - 1
                str������ = IIf(str������ = "", "", str������ & "|") & .ColKey(i) & "," & IIf(.ColHidden(i) = True, 0, .ColWidth(i))
            Next
        End With
        
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList(n)), strType, str������)
    Next
End Sub

Private Function LoadListColState(ByVal intType As Integer) As String
    Dim str������ As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zlDataBase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    
    Select Case intType
        Case mListType.��ҩ
            strType = "��ҩ"
        Case mListType.����
            If mcondition.bln�����һ��� = True Then
                strType = "���һ���"
            Else
                strType = "����"
            End If
        Case mListType.ȱҩ
            strType = "ȱҩ"
        Case mListType.�ܷ�
            strType = "�ܷ�"
        Case mListType.��ҩ
            strType = "��ҩ"
    End Select
    
    LoadListColState = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList(intType)), strType, "")
End Function
Private Sub InitList_Send()
    Dim i As Integer
    Dim n As Integer
    Dim str������ As String
    Dim arr������
    
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntCol��ҩ_��ǰ�� = 0
    mIntCol��ҩ_����� = 1
    mIntCol��ҩ_����� = 2
    mIntCol��ҩ_��ҩ���� = 3
    mIntCol��ҩ_���� = 4
    mIntCol��ҩ_����ҽ�� = 5
    mIntCol��ҩ_״̬ = 6
    mIntCol��ҩ_���� = 7
    mIntCol��ҩ_��ҩ���� = 8
    mIntCol��ҩ_NO = 9
    mIntCol��ҩ_����Ա = 10
    mIntCol��ҩ_���� = 11
    mIntCol��ҩ_�������� = 12
    mIntCol��ҩ_���� = 13
    mIntCol��ҩ_�Ա� = 14
    mIntCol��ҩ_���� = 15
    mIntCol��ҩ_סԺ�� = 16
    mIntCol��ҩ_Ʒ�� = 17
    mIntCol��ҩ_Ƥ�Խ�� = 18
    mIntCol��ҩ_������ = 19
    mIntCol��ҩ_Ӣ���� = 20
    mIntCol��ҩ_�䷽���� = 21
    mIntCol��ҩ_��� = 22
    mIntCol��ҩ_������ = 23
    mIntCol��ҩ_ԭ���� = 24
    mIntCol��ҩ_���� = 25
    mIntCol��ҩ_Ч�� = 26
    mIntCol��ҩ_�� = 27
    mIntCol��ҩ_���� = 28
    mIntCol��ҩ_���� = 29
    mIntCol��ҩ_��� = 30
    mIntCol��ҩ_���� = 31
    mIntCol��ҩ_Ƶ�� = 32
    mIntCol��ҩ_�÷� = 33
    mIntCol��ҩ_��ҩ���� = 34
    mIntCol��ҩ_��ҩĿ�� = 35
    mIntCol��ҩ_����ҩƷ˵�� = 36
    mIntCol��ҩ_����ʱ�� = 37
    mIntCol��ҩ_˵�� = 38
    mIntCol��ҩ_���� = 39
    mIntCol��ҩ_ҽ��id = 40
    mIntCol��ҩ_��ҩ�� = 41
    mIntCol��ҩ_�ⷿ��λ = 42
    mIntCol��ҩ_���ID = 43
    mIntCol��ҩ_ҩƷID = 44
    mIntCol��ҩ_������λ = 45
    mIntCol��ҩ_��ҩ����id = 46
    mIntCol��ҩ_ҩƷ��������� = 47
    mIntCol��ҩ_ҩƷ���� = 48
    mIntCol��ҩ_ҩƷ���� = 49
    mIntCol��ҩ_�շ�ID = 50
    mIntCol��ҩ_ִ��״̬ = 51
    mIntCol��ҩ_��ҩ�� = 52
    mIntCol��ҩ_���շ� = 53
    mIntCol��ҩ_����ID = 54
    mIntCol��ҩ_��ҳID = 55
    mIntCol��ҩ_��ҩ���� = 56
    mIntCol��ҩ_��ΣҩƷ = 57
    mIntCol��ҩ_��� = 58
    mIntCol��ҩ_��������id = 59
    mIntCol��ҩ_��ע = 60
    
    '�ָ��û��Զ���˳��
    str������ = LoadListColState(mListType.��ҩ)
    If str������ <> "" Then
        arr������ = Split(str������, "|")
        If UBound(arr������) + 1 <> mconIntCol��ҩ_���� Then
            str������ = ""
        Else
            For n = 0 To UBound(arr������)
                SetColumnValue mListType.��ҩ, Split(arr������(n), ",")(0), n
            Next
        End If
    End If
    
    '��ʼ��δ��ҩ�嵥
    With vsfList(mListType.��ҩ)
        .Redraw = flexRDNone
        
        .rows = 2
        
'        .RowHeightMax = 255
        .Cols = mconIntCol��ҩ_����
        
'        .Cell(flexcpPicture, 1, mIntCol��ҩ_��ǰ��, 1, mIntCol��ҩ_��ǰ��) = Me.imgList.ListImages(2).Picture
'        .Cell(flexcpPictureAlignment, 1, mIntCol��ҩ_��ǰ��, .Rows - 1, mIntCol��ҩ_��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ǰ��, "", 250, flexAlignCenterCenter, "��ǰ��"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�����, "��", IIf(Not (gobjPass Is Nothing) And IsInString(gstrprivs, "������ҩ���", ";"), 300, 0), flexAlignCenterCenter, "�����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�����, "��", 300, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "���˿���", 1000, flexAlignLeftCenter, "���˿���"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����ҽ��, "����ҽ��", IIf(mcondition.blnҽ����ѯ = True, 1100, 0), flexAlignLeftCenter, "����ҽ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_״̬, "״̬", 1000, flexAlignLeftCenter, "״̬"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1000, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ����, "��ҩ����", 1000, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_NO, "NO", 900, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����Ա, "����Ա", 800, flexAlignLeftCenter, "����Ա"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1000, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��������, "��������", 1000, flexAlignLeftCenter, "��������"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 700, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�Ա�, "�Ա�", 700, flexAlignLeftCenter, "�Ա�"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 700, flexAlignLeftCenter, "����"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_סԺ��, "סԺ��", 1200, flexAlignLeftCenter, "סԺ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_Ʒ��, "ҩƷ����", 2500, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_Ƥ�Խ��, "", 800, flexAlignLeftCenter, "Ƥ�Խ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_������, "������", 2000, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_Ӣ����, "Ӣ����", 2000, flexAlignLeftCenter, "Ӣ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�䷽����, "�䷽����", 2000, flexAlignLeftCenter, "�䷽����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_���, "���", 1500, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_������, "������", 1500, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ԭ����, "ԭ����", 1500, flexAlignLeftCenter, "ԭ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1500, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_Ч��, "Ч��", 1500, flexAlignLeftCenter, "Ч��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��, "��", 300, flexAlignRightCenter, "��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1200, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1200, flexAlignRightCenter, "����"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_���, "���", 1200, flexAlignRightCenter, "���"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1200, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_Ƶ��, "Ƶ��", 500, flexAlignLeftCenter, "Ƶ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�÷�, "�÷�", 800, flexAlignLeftCenter, "�÷�"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ����, "��ҩ����", 900, flexAlignRightCenter, "��ҩ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩĿ��, "��ҩĿ��", 0, flexAlignLeftCenter, "��ҩĿ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����ҩƷ˵��, "����ҩƷ˵��", 1500, flexAlignLeftCenter, "����ҩƷ˵��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����ʱ��, "����ʱ��", 1800, flexAlignLeftCenter, "����ʱ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_˵��, "˵��", 1200, flexAlignLeftCenter, "˵��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҽ��id, "ҽ��id", 0, flexAlignCenterCenter, "ҽ��id"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ��, "��ҩ��", 1000, flexAlignLeftCenter, "��ҩ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�ⷿ��λ, "�ⷿ��λ", IIf(mcondition.blnҩƷ���� = True, 1200, 0), flexAlignLeftCenter, "�ⷿ��λ"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_���ID, "���ID", 0, flexAlignCenterCenter, "���ID"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҩƷID, "ҩƷID", 0, flexAlignCenterCenter, "ҩƷID"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_������λ, "������λ", 0, flexAlignLeftCenter, "������λ"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ����, "��ҩ����", 1000, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ����id, "��ҩ����id", 0, flexAlignCenterCenter, "��ҩ����id"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҩƷ���������, "ҩƷ���������", 0, flexAlignCenterCenter, "ҩƷ���������"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҩƷ����, "ҩƷ����", 0, flexAlignCenterCenter, "ҩƷ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҩƷ����, "ҩ��", 0, flexAlignCenterCenter, "ҩ��"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�շ�ID, "�շ�ID", 0, flexAlignCenterCenter, "�շ�ID"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ִ��״̬, "״̬��־", 0, flexAlignCenterCenter, "״̬��־"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ��, "��ҩ��", 1000, flexAlignLeftCenter, "��ҩ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_���շ�, "���շ�", 0, flexAlignLeftCenter, "���շ�"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����ID, "����ID", 0, flexAlignCenterCenter, "����ID"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҳID, "��ҳID", 0, flexAlignCenterCenter, "��ҳID"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ����, "��ҩ����", 0, flexAlignCenterCenter, "��ҩ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ΣҩƷ, "��ΣҩƷ", 0, flexAlignCenterCenter, "��ΣҩƷ"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_���, "���", 0, flexAlignCenterCenter, "���"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��������id, "��������id", 0, flexAlignCenterCenter, "��������id"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ע, "��ע", 1000, flexAlignLeftCenter, "��ע"
        
        mstrUnallowSetColHide(mListType.��ҩ) = "״̬;ҩƷ����;����"
        mstrUnallowShow(mListType.��ҩ) = "��ǰ��;�����;����;����;ҽ��id;��ҩĿ��;���ID;ҩƷID;Ƥ�Խ��;������λ;��ҩ����id;ҩƷ���������;ҩƷ����;ҩ��;�շ�ID;״̬��־;���շ�;����ID;��ҳID;��ҩ����;��ΣҩƷ;���;��������id"
        If mcondition.blnҩƷ���� = False Then mstrUnallowShow(mListType.��ҩ) = mstrUnallowShow(mListType.��ҩ) & ";�ⷿ��λ"
        If mcondition.blnҽ����ѯ = False Then mstrUnallowShow(mListType.��ҩ) = mstrUnallowShow(mListType.��ҩ) & ";����ҽ��"
        
        '�ָ��Զ����п���������������ʾ���У�
        If str������ <> "" Then
            arr������ = Split(str������, "|")
            For n = 0 To UBound(arr������)
                If IsInString(mstrUnallowShow(mListType.��ҩ), Split(arr������(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr������(n), ",")(0) = .ColKey(i) Then
                            .ColWidth(i) = Val(Split(arr������(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln��ʾԭ���� = False Then VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ԭ����, "ԭ����", 0, flexAlignLeftCenter, "ԭ����"
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub InitList_Reject()
    Dim i As Integer
    Dim n As Integer
    Dim str������ As String
    Dim arr������
    
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntCol�ܷ�_��ǰ�� = 0
    mIntCol�ܷ�_���� = 1
    mIntCol�ܷ�_״̬ = 2
    mIntCol�ܷ�_NO = 3
    mIntCol�ܷ�_���� = 4
    mIntCol�ܷ�_��ҩ���� = 5
    mIntCol�ܷ�_���� = 6
    mIntCol�ܷ�_���� = 7
    mIntCol�ܷ�_�Ա� = 8
    mIntCol�ܷ�_Ʒ�� = 9
    mIntCol�ܷ�_��� = 10
    mIntCol�ܷ�_������ = 11
    mIntCol�ܷ�_ԭ���� = 12
    mIntCol�ܷ�_���� = 13
    mIntCol�ܷ�_Ч�� = 14
    mIntCol�ܷ�_���� = 15
    mIntCol�ܷ�_���� = 16
    mIntCol�ܷ�_��� = 17
    mIntCol�ܷ�_ҩƷ��������� = 18
    mIntCol�ܷ�_ҩƷ���� = 19
    mIntCol�ܷ�_ҩƷ���� = 20
    mIntCol�ܷ�_ִ��״̬ = 21
    mIntCol�ܷ�_�շ�ID = 22
    mIntCol�ܷ�_��ע = 23
    
    '�ָ��û��Զ���˳��
    str������ = LoadListColState(mListType.�ܷ�)
    If str������ <> "" Then
        arr������ = Split(str������, "|")
        If UBound(arr������) + 1 <> mconIntCol�ܷ�_���� Then
            str������ = ""
        Else
            For n = 0 To UBound(arr������)
               SetColumnValue mListType.�ܷ�, Split(arr������(n), ",")(0), n
            Next
        End If
    End If
    
    '��ʼ���ܷ�ҩ�嵥
    With vsfList(mListType.�ܷ�)
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntCol�ܷ�_����
        
        .Cell(flexcpPicture, 1, mIntCol�ܷ�_��ǰ��, 1, mIntCol�ܷ�_��ǰ��) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol�ܷ�_��ǰ��, .rows - 1, mIntCol�ܷ�_��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_��ǰ��, "", 250, flexAlignCenterCenter, "��ǰ��"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_����, "���˿���", 1200, flexAlignLeftCenter, "���˿���"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_״̬, "״̬", 1000, flexAlignLeftCenter, "״̬"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_NO, "NO", 800, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_����, "����", 1000, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_��ҩ����, "��ҩ����", 1000, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_����, "����", 800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_����, "����", 1000, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_�Ա�, "�Ա�", 1000, flexAlignLeftCenter, "�Ա�"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_Ʒ��, "ҩƷ����", 2500, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_���, "���", 1500, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_������, "������", 1500, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_ԭ����, "ԭ����", 1500, flexAlignLeftCenter, "ԭ����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_����, "����", 1500, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_Ч��, "Ч��", 1500, flexAlignLeftCenter, "Ч��"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_����, "����", 1200, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_����, "����", 1200, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_���, "���", 1200, flexAlignRightCenter, "���"
        
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_ҩƷ���������, "ҩƷ���������", 0, flexAlignCenterCenter, "ҩƷ���������"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_ҩƷ����, "ҩƷ����", 0, flexAlignCenterCenter, "ҩƷ����"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_ҩƷ����, "ҩƷ����", 0, flexAlignCenterCenter, "ҩ��"
        
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_ִ��״̬, "ִ��״̬", 0, flexAlignCenterCenter, "ִ��״̬"
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_�շ�ID, "�շ�ID", 0, flexAlignCenterCenter, "�շ�ID"
        
        VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_��ע, "��ע", 1200, flexAlignLeftCenter, "��ע"
        
        mstrUnallowSetColHide(mListType.�ܷ�) = "״̬;ҩƷ����;����"
        mstrUnallowShow(mListType.�ܷ�) = "��ǰ��;ҩƷ���������;ҩƷ����;ҩ��;ִ��״̬;�շ�ID"
        
        '�ָ��Զ����п���������������ʾ���У�
        If str������ <> "" Then
            arr������ = Split(str������, "|")
            For n = 0 To UBound(arr������)
                If IsInString(mstrUnallowShow(mListType.�ܷ�), Split(arr������(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr������(n), ",")(0) = .ColKey(i) Then
                            .ColWidth(i) = Val(Split(arr������(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln��ʾԭ���� = False Then VsfGridColFormat vsfList(mListType.�ܷ�), mIntCol�ܷ�_ԭ����, "ԭ����", 0, flexAlignLeftCenter, "ԭ����"
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub InitList_Shortage()
    Dim i As Integer
    Dim n As Integer
    Dim str������ As String
    Dim arr������
    
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntColȱҩ_��ǰ�� = 0
    mIntColȱҩ_���� = 1
    mIntColȱҩ_NO = 2
    mIntColȱҩ_���� = 3
    mIntColȱҩ_��ҩ���� = 4
    mIntColȱҩ_���� = 5
    mIntColȱҩ_���� = 6
    mIntColȱҩ_�Ա� = 7
    mIntColȱҩ_Ʒ�� = 8
    mIntColȱҩ_��� = 9
    mIntColȱҩ_������ = 10
    mIntColȱҩ_ԭ���� = 11
    mIntColȱҩ_���� = 12
    mIntColȱҩ_Ч�� = 13
    mIntColȱҩ_���� = 14
    mIntColȱҩ_���� = 15
    mIntColȱҩ_��� = 16
    mIntColȱҩ_ҩƷ��������� = 17
    mIntColȱҩ_ҩƷ���� = 18
    mIntColȱҩ_ҩƷ���� = 19
    mIntColȱҩ_��ע = 20
    
    '�ָ��û��Զ���˳��
    str������ = LoadListColState(mListType.ȱҩ)
    If str������ <> "" Then
        arr������ = Split(str������, "|")
        If UBound(arr������) + 1 <> mconIntColȱҩ_���� Then
            str������ = ""
        Else
            For n = 0 To UBound(arr������)
                SetColumnValue mListType.ȱҩ, Split(arr������(n), ",")(0), n
            Next
        End If
    End If
    
    '��ʼ��ȱҩ�嵥
    With vsfList(mListType.ȱҩ)
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntColȱҩ_����
        
        .Cell(flexcpPicture, 1, mIntColȱҩ_��ǰ��, 1, mIntColȱҩ_��ǰ��) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntColȱҩ_��ǰ��, .rows - 1, mIntColȱҩ_��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_��ǰ��, "", 250, flexAlignCenterCenter, "��ǰ��"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_����, "���˿���", 1200, flexAlignLeftCenter, "���˿���"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_NO, "NO", 800, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_����, "����", 1000, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_��ҩ����, "��ҩ����", 1000, flexAlignCenterCenter, "��ҩ����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_����, "����", 800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_����, "����", 1000, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_�Ա�, "�Ա�", 1000, flexAlignLeftCenter, "�Ա�"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_Ʒ��, "ҩƷ����", 2500, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_���, "���", 1500, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_������, "������", 1500, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_ԭ����, "ԭ����", 1500, flexAlignLeftCenter, "ԭ����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_����, "����", 1500, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_Ч��, "Ч��", 1500, flexAlignLeftCenter, "Ч��"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_����, "����", 1200, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_����, "����", 1200, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_���, "���", 1200, flexAlignRightCenter, "���"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_ҩƷ���������, "ҩƷ���������", 0, flexAlignRightCenter, "ҩƷ���������"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_ҩƷ����, "ҩƷ����", 0, flexAlignRightCenter, "ҩƷ����"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_ҩƷ����, "ҩ��", 0, flexAlignRightCenter, "ҩ��"
        VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_��ע, "��ע", 1200, flexAlignLeftCenter, "��ע"
        
        mstrUnallowSetColHide(mListType.ȱҩ) = "ҩƷ����;����"
        mstrUnallowShow(mListType.ȱҩ) = "��ǰ��;ҩƷ���������;ҩƷ����;ҩ��"
        
        '�ָ��Զ����п���������������ʾ���У�
        If str������ <> "" Then
            arr������ = Split(str������, "|")
            For n = 0 To UBound(arr������)
                If IsInString(mstrUnallowShow(mListType.ȱҩ), Split(arr������(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr������(n), ",")(0) = .ColKey(i) Then
                            .ColWidth(i) = Val(Split(arr������(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln��ʾԭ���� = False Then VsfGridColFormat vsfList(mListType.ȱҩ), mIntColȱҩ_ԭ����, "ԭ����", 0, flexAlignLeftCenter, "ԭ����"
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub InitList_ChargeOff()
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntCol����_��ǰ�� = 0
    mIntCol����_������� = 1
    mIntCol����_���� = 2
    mIntCol����_NO = 3
    mIntCol����_ҩƷID = 4
    mIntCol����_����ʱ�� = 5
    mIntCol����_�շ���� = 6
    mIntCol����_������ = 7
    mIntCol����_���� = 8
    mIntCol����_Ч�� = 9
    mIntCol����_׼������ = 10
    mIntCol����_�������� = 11
    mIntCol����_��װ = 12
    mIntCol����_��λ = 13
    
    '�ָ��û��Զ���˳��
    
    '��ʼ��ȱҩ�嵥
    With vsfChargeOff
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntCol����_����
        
        .Cell(flexcpPicture, 1, mIntCol����_��ǰ��, 1, mIntCol����_��ǰ��) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol����_��ǰ��, .rows - 1, mIntCol����_��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfChargeOff, mIntCol����_��ǰ��, "", 250, flexAlignCenterCenter
        
        VsfGridColFormat vsfChargeOff, mIntCol����_�������, "�������", 1500, flexAlignLeftCenter, "�������"
        VsfGridColFormat vsfChargeOff, mIntCol����_����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfChargeOff, mIntCol����_NO, "NO", 1200, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfChargeOff, mIntCol����_ҩƷID, "ҩƷID", 0, flexAlignLeftCenter, "ҩƷID"
        VsfGridColFormat vsfChargeOff, mIntCol����_����ʱ��, "����ʱ��", 2000, flexAlignRightCenter, "����ʱ��"
        VsfGridColFormat vsfChargeOff, mIntCol����_�շ����, "�շ����", 0, flexAlignLeftCenter, "�շ����"
        
        VsfGridColFormat vsfChargeOff, mIntCol����_������, "������", 2000, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfChargeOff, mIntCol����_����, "����", 1000, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfChargeOff, mIntCol����_Ч��, "Ч��", 1500, flexAlignLeftCenter, "Ч��"
        VsfGridColFormat vsfChargeOff, mIntCol����_׼������, "׼������", 1000, flexAlignRightCenter, "׼������"
        VsfGridColFormat vsfChargeOff, mIntCol����_��������, "��������", 1000, flexAlignRightCenter, "��������"
        
        VsfGridColFormat vsfChargeOff, mIntCol����_��װ, "��װ", 0, flexAlignLeftCenter, "��װ"
        VsfGridColFormat vsfChargeOff, mIntCol����_��λ, "��λ", 1000, flexAlignLeftCenter, "��λ"
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Function GetChargeOffCount(ByVal lng����ID As Long, ByVal lngҩƷid As Long, ByVal lng���� As Long) As Double
    Dim dblSum As Double
    
    With mrsChargeOff
        If mrsChargeOff Is Nothing Then Exit Function

        
        If mcondition.bln�����λ��� Then
            .Filter = "��ҩ����id=" & lng����ID & " And ҩƷID=" & lngҩƷid & " And ����=" & lng���� & " And ��˱�־ = 1"
        Else
            .Filter = "��ҩ����id=" & lng����ID & " And ҩƷID=" & lngҩƷid & " And ��˱�־ = 1"
        End If
        If .EOF Then Exit Function
        
        Do While Not .EOF
            Debug.Print mrsChargeOff!����
            dblSum = dblSum + !�������� / !��װ
            .MoveNext
        Loop
        
    End With
    
    GetChargeOffCount = dblSum
End Function
Private Sub InitList_Return()
    Dim i As Integer
    Dim n As Integer
    Dim str������ As String
    Dim arr������
    
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntCol��ҩ_��ǰ�� = 0
    mIntCol��ҩ_����� = 1
    mIntCol��ҩ_����� = 2
    mIntCol��ҩ_���� = 3
    mIntCol��ҩ_״̬ = 4
    mIntCol��ҩ_���� = 5
    mIntCol��ҩ_��ҩ���� = 6
    mIntCol��ҩ_NO = 7
    mIntCol��ҩ_���� = 8
    mIntCol��ҩ_���� = 9
    mIntCol��ҩ_�Ա� = 10
    mIntCol��ҩ_סԺ�� = 11
    mIntCol��ҩ_Ʒ�� = 12
    mIntCol��ҩ_������ = 13
    mIntCol��ҩ_Ӣ���� = 14
    mIntCol��ҩ_��� = 15
    mIntCol��ҩ_������ = 16
    mIntCol��ҩ_ԭ���� = 17
    mIntCol��ҩ_���� = 18
    mIntCol��ҩ_Ч�� = 19
    mIntCol��ҩ_�� = 20
    mIntCol��ҩ_���� = 21
    mIntCol��ҩ_������ = 22
    mIntCol��ҩ_׼���� = 23
    mIntCol��ҩ_��ҩ�� = 24
    mIntCol��ҩ_���� = 25
    mIntCol��ҩ_��� = 26
    mIntCol��ҩ_���� = 27
    mIntCol��ҩ_Ƶ�� = 28
    mIntCol��ҩ_�÷� = 29
    mIntCol��ҩ_����Ա = 30
    mIntCol��ҩ_��ҩʱ�� = 31
    mIntCol��ҩ_��ҩ�� = 32
    mIntCol��ҩ_��ҩ�� = 33
    mIntCol��ҩ_����ʱ�� = 34
    mIntCol��ҩ_���� = 35
    mIntCol��ҩ_ҽ��id = 36
    
    mIntCol��ҩ_�ⷿ��λ = 37
    mIntCol��ҩ_���ID = 38
    mIntCol��ҩ_ҩƷID = 39
    mIntCol��ҩ_������λ = 40
    mIntCol��ҩ_ҩƷ��������� = 41
    mIntCol��ҩ_ҩƷ���� = 42
    mIntCol��ҩ_ҩƷ���� = 43
    mIntCol��ҩ_�շ�ID = 44
    mIntCol��ҩ_ִ��״̬ = 45
    mIntCol��ҩ_��ҩ����id = 46
    mIntCol��ҩ_��ע = 47
    mIntCol��ҩ_����ID = 48
    mIntCol��ҩ_��ҳID = 49
    
    '�ָ��û��Զ���˳��
    str������ = LoadListColState(mListType.��ҩ)
    If str������ <> "" Then
        arr������ = Split(str������, "|")
        If UBound(arr������) + 1 <> mconIntCol��ҩ_���� Then
            str������ = ""
        Else
            For n = 0 To UBound(arr������)
                SetColumnValue mListType.��ҩ, Split(arr������(n), ",")(0), n
            Next
        End If
    End If

    '��ʼ����ҩ�嵥
    With vsfList(mListType.��ҩ)
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntCol��ҩ_����
        
        .Cell(flexcpPicture, 1, mIntCol��ҩ_��ǰ��, 1, mIntCol��ҩ_��ǰ��) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol��ҩ_��ǰ��, .rows - 1, mIntCol��ҩ_��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ǰ��, "", 250, flexAlignCenterCenter, "��ǰ��"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�����, "��", IIf(mcondition.intShowPass <> 0 And IsInString(gstrprivs, "������ҩ���", ";"), 300, 0), flexAlignCenterCenter, "�����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�����, "��", 300, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "���˿���", 1200, flexAlignLeftCenter, "���˿���"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_״̬, "״̬", 1000, flexAlignLeftCenter, "״̬"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1000, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ����, "��ҩ����", 1000, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_NO, "NO", 900, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 600, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 700, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�Ա�, "�Ա�", 700, flexAlignLeftCenter, "�Ա�"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_סԺ��, "סԺ��", 1200, flexAlignLeftCenter, "סԺ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_Ʒ��, "ҩƷ����", 2500, flexAlignLeftCenter, "ҩƷ����"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_������, "������", 2000, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_Ӣ����, "Ӣ����", 2000, flexAlignLeftCenter, "Ӣ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_���, "���", 1500, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_������, "������", 1500, flexAlignCenterCenter, "������"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ԭ����, "ԭ����", 1500, flexAlignCenterCenter, "ԭ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1500, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_Ч��, "Ч��", 1500, flexAlignLeftCenter, "Ч��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��, "��", 300, flexAlignRightCenter, "��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1000, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_������, "������", 1000, flexAlignRightCenter, "������"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_׼����, "׼����", 1000, flexAlignRightCenter, "׼����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ��, "��ҩ��", 1000, flexAlignRightCenter, "��ҩ��"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1000, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_���, "���", 1000, flexAlignRightCenter, "���"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 1000, flexAlignRightCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_Ƶ��, "Ƶ��", 500, flexAlignLeftCenter, "Ƶ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�÷�, "�÷�", 800, flexAlignLeftCenter, "�÷�"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����Ա, "����Ա", 800, flexAlignLeftCenter, "����Ա"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩʱ��, "��ҩʱ��", 1500, flexAlignLeftCenter, "��ҩʱ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҽ��id, "ҽ��id", 0, flexAlignCenterCenter, "ҽ��id"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ��, "��/��ҩ��", 1000, flexAlignLeftCenter, "��/��ҩ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ��, "��ҩ��", 1200, flexAlignLeftCenter, "��ҩ��"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����ʱ��, "����ʱ��", 1200, flexAlignLeftCenter, "����ʱ��"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҽ��id, "ҽ��id", 0, flexAlignCenterCenter, "ҽ��id"
        
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�ⷿ��λ, "�ⷿ��λ", IIf(mcondition.blnҩƷ���� = True, 1200, 0), flexAlignLeftCenter, "�ⷿ��λ"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_���ID, "���ID", 0, flexAlignCenterCenter, "���ID"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҩƷID, "ҩƷID", 0, flexAlignCenterCenter, "ҩƷID"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_������λ, "������λ", 0, flexAlignCenterCenter, "������λ"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҩƷ���������, "ҩƷ���������", 0, flexAlignCenterCenter, "ҩƷ���������"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҩƷ����, "ҩƷ����", 0, flexAlignCenterCenter, "ҩƷ����"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ҩƷ����, "ҩ��", 0, flexAlignCenterCenter, "ҩ��"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_�շ�ID, "�շ�ID", 0, flexAlignCenterCenter, "�շ�ID"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ִ��״̬, "ִ��״̬", 0, flexAlignCenterCenter, "ִ��״̬"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҩ����id, "��ҩ����id", 0, flexAlignLeftCenter, "��ҩ����id"
        
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ע, "��ע", 1200, flexAlignLeftCenter, "��ע"
            
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_����ID, "����ID", 0, flexAlignLeftCenter, "����ID"
        VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_��ҳID, "��ҳID", 0, flexAlignLeftCenter, "��ҳID"
            
        mstrUnallowSetColHide(mListType.��ҩ) = "ҩƷ����;����;��ҩ��"
        mstrUnallowShow(mListType.��ҩ) = "��ǰ��;�����;����;����;ҽ��id;���ID;ҩƷID;������λ;ҩƷ���������;ҩƷ����;ҩ��;ִ��״̬;�շ�ID;��ҩ����id;����ID;��ҳID"
        If mcondition.blnҩƷ���� = False Then mstrUnallowShow(mListType.��ҩ) = mstrUnallowShow(mListType.��ҩ) & ";�ⷿ��λ"
        
        '�ָ��Զ����п���������������ʾ���У�
        If str������ <> "" Then
            arr������ = Split(str������, "|")
            For n = 0 To UBound(arr������)
                If IsInString(mstrUnallowShow(mListType.��ҩ), Split(arr������(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr������(n), ",")(0) = .ColKey(i) Then
                            .ColWidth(i) = Val(Split(arr������(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln��ʾԭ���� = False Then VsfGridColFormat vsfList(mListType.��ҩ), mIntCol��ҩ_ԭ����, "ԭ����", 0, flexAlignCenterCenter, "ԭ����"
        
        .Redraw = flexRDDirect
    End With
    
End Sub
Private Sub InitList_Sum()
    Dim int��ǰ�� As Integer
    Dim i As Integer
    Dim n As Integer
    Dim str������ As String
    Dim arr������
    
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntCol����_��ǰ�� = 0
    mIntCol����_Ʒ�� = 1
    mIntCol����_��� = 2
    mIntCol����_������ = 3
    mIntCol����_ԭ���� = 4
    mIntCol����_���� = 5
    mIntCol����_Ч�� = 6
    mIntCol����_���� = 7
    mIntCol����_��λ = 8
    mIntCol����_���� = 9
    mIntCol����_��� = 10
    mIntCol����_ҩƷ��������� = 11
    mIntCol����_ҩƷ���� = 12
    mIntCol����_ҩƷ���� = 13
    
    mIntCol���һ���_��ǰ�� = 0
    mIntCol���һ���_��ҩ���� = 1
    mIntCol���һ���_���� = 2
    mIntCol���һ���_Ʒ�� = 3
    mIntCol���һ���_��� = 4
    mIntCol���һ���_������ = 5
    mIntCol���һ���_ԭ���� = 6
    mIntCol���һ���_���� = 7
    mIntCol���һ���_Ч�� = 8
    mIntCol���һ���_Ӧ������ = 9
    mIntCol���һ���_�������� = 10
    mIntCol���һ���_�������� = 11
    mIntCol���һ���_ʵ������ = 12
    mIntCol���һ���_��λ = 13
    mIntCol���һ���_���� = 14
    mIntCol���һ���_Ӧ����� = 15
    mIntCol���һ���_ʵ����� = 16
    mIntCol���һ���_���� = 17
    mIntCol���һ���_����ID = 18
    mIntCol���һ���_ҩƷID = 19
    mIntCol���һ���_��ҩ����id = 20
    mIntCol���һ���_ҩƷ��������� = 21
    mIntCol���һ���_ҩƷ���� = 22
    mIntCol���һ���_ҩƷ���� = 23
    mIntCol���һ���_��װ = 24
    
    '�ָ��û��Զ���˳��
    str������ = LoadListColState(mListType.����)
    If str������ <> "" Then
        arr������ = Split(str������, "|")
        If UBound(arr������) + 1 <> IIf(mcondition.bln�����һ��� = True, mconIntCol���һ���_����, mconIntCol����_����) Then
            str������ = ""
        Else
            For n = 0 To UBound(arr������)
                SetColumnValue mListType.����, Split(arr������(n), ",")(0), n
            Next
        End If
    End If
    
    '''��ʼ�����ܷ�ҩ�嵥
    With vsfList(mListType.����)
        .Redraw = flexRDNone
        .rows = 2
        .Cols = IIf(mcondition.bln�����һ���, mconIntCol���һ���_����, mconIntCol����_����)
        
        If mcondition.bln�����һ��� Then
            int��ǰ�� = mIntCol���һ���_��ǰ��
        Else
            int��ǰ�� = mIntCol����_��ǰ��
        End If
        
        .Cell(flexcpPicture, 1, int��ǰ��, 1, int��ǰ��) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, int��ǰ��, .rows - 1, int��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.����), int��ǰ��, "", 250, flexAlignCenterCenter, "��ǰ��"
        
        If mcondition.bln�����һ��� = False Then
            VsfGridColFormat vsfList(mListType.����), mIntCol����_Ʒ��, "ҩƷ����", 2500, flexAlignLeftCenter, "ҩƷ����"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_���, "���", 1500, flexAlignLeftCenter, "���"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_������, "������", 1500, flexAlignLeftCenter, "������"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_ԭ����, "ԭ����", 1500, flexAlignLeftCenter, "ԭ����"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_����, "����", 1200, flexAlignLeftCenter, "����"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_Ч��, "Ч��", 1200, flexAlignLeftCenter, "Ч��"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_����, "����", 1200, flexAlignRightCenter, "����"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_��λ, "��λ", 500, flexAlignCenterCenter, "��λ"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_����, "����", 1200, flexAlignRightCenter, "����"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_���, "���", 1200, flexAlignRightCenter, "���"
            
            VsfGridColFormat vsfList(mListType.����), mIntCol����_ҩƷ���������, "ҩƷ���������", 0, flexAlignLeftCenter, "ҩƷ���������"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_ҩƷ����, "ҩƷ����", 0, flexAlignLeftCenter, "ҩƷ����"
            VsfGridColFormat vsfList(mListType.����), mIntCol����_ҩƷ����, "ҩ��", 0, flexAlignLeftCenter, "ҩ��"
        Else
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_����, "��������", 1200, flexAlignLeftCenter, "��������"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_Ʒ��, "ҩƷ����", 2500, flexAlignLeftCenter, "ҩƷ����"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_���, "���", 1500, flexAlignLeftCenter, "���"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_������, "������", 1500, flexAlignLeftCenter, "������"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_ԭ����, "ԭ����", 1500, flexAlignLeftCenter, "ԭ����"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_����, "����", 1200, flexAlignLeftCenter, "����"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_Ч��, "Ч��", 1200, flexAlignLeftCenter, "Ч��"
            
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_Ӧ������, "Ӧ������", 1200, flexAlignRightCenter, "Ӧ������"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_��������, "��������", 1200, flexAlignRightCenter, "��������"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_��������, "��������", IIf(mcondition.bln������ҩ���� = True, 1200, 0), flexAlignRightCenter, "��������"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_ʵ������, "ʵ������", 1200, flexAlignRightCenter, "ʵ������"

            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_��λ, "��λ", 500, flexAlignCenterCenter, "��λ"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_����, "����", 1200, flexAlignRightCenter, "����"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_Ӧ�����, "Ӧ�����", 1200, flexAlignRightCenter, "Ӧ�����"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_ʵ�����, "ʵ�����", 1200, flexAlignRightCenter, "ʵ�����"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_����, "����", 0, flexAlignRightCenter, "����"

            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_����ID, "����ID", 0, flexAlignCenterCenter, "����ID"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_ҩƷID, "ҩƷID", 0, flexAlignLeftCenter, "ҩƷID"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_��ҩ����, "��ҩ����", 1200, flexAlignLeftCenter, "��ҩ����"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_��ҩ����id, "��ҩ����id", 0, flexAlignLeftCenter, "��ҩ����id"
            
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_ҩƷ���������, "ҩƷ���������", 0, flexAlignLeftCenter, "ҩƷ���������"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_ҩƷ����, "ҩƷ����", 0, flexAlignLeftCenter, "ҩƷ����"
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_ҩƷ����, "ҩ��", 0, flexAlignLeftCenter, "ҩ��"
            
            VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_��װ, "��װ", 0, flexAlignLeftCenter, "��װ"
        End If
        
        mstrUnallowSetColHide(mListType.����) = "ҩƷ����;����;Ӧ������;ʵ������;��λ"
        If mcondition.bln������ҩ���� = True Then mstrUnallowSetColHide(mListType.����) = mstrUnallowSetColHide(mListType.����) & ";��������"
        
        mstrUnallowShow(mListType.����) = "��ǰ��;����;����ID;ҩƷID;��ҩ����id;ҩƷ���������;ҩƷ����;ҩ��;��װ"
        If mcondition.bln������ҩ���� = False Then mstrUnallowShow(mListType.����) = mstrUnallowShow(mListType.����) & ";��������"
        
        '�ָ��Զ����п���������������ʾ���У�
        If str������ <> "" Then
            arr������ = Split(str������, "|")
            For n = 0 To UBound(arr������)
                If IsInString(mstrUnallowShow(mListType.����), Split(arr������(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr������(n), ",")(0) = .ColKey(i) Then
                            If IsInString(mstrUnallowSetColHide(mListType.����), Split(arr������(n), ",")(0), ";") = True Then
                                '����ǲ��������ص��У����п���Ϊ0
                                If Val(Split(arr������(n), ",")(1)) <> 0 Then
                                    .ColWidth(i) = Val(Split(arr������(n), ",")(1))
                                End If
                            Else
                                .ColWidth(i) = Val(Split(arr������(n), ",")(1))
                            End If
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln�����һ��� = False Then
            If mcondition.bln��ʾԭ���� = False Then VsfGridColFormat vsfList(mListType.����), mIntCol����_ԭ����, "ԭ����", 0, flexAlignLeftCenter, "ԭ����"
        Else
            If mcondition.bln��ʾԭ���� = False Then VsfGridColFormat vsfList(mListType.����), mIntCol���һ���_ԭ����, "ԭ����", 0, flexAlignLeftCenter, "ԭ����"
        End If
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub SetGroup(ByVal Bill As VSFlexGrid, ByVal bln�Ƿ���� As Boolean)
    Dim n As Integer
    Dim str������� As String
    Dim str������� As String
    Dim str������� As String
    Dim int����_���ID As Integer
    Dim int����_NO As Integer
    Dim int����_����� As Integer
    Dim bln�Ƿ���ڷ��� As Boolean
    Dim bln�����з���� As Boolean
    
    '������С������ʱû�б�Ҫ���飨1�й̶��У�2�л����У�
    If Bill.rows < 4 Then Exit Sub
    
    str������� = "-1"
        
    '�����ID����
    With Bill
        Select Case Bill.index
        Case mListType.��ҩ
            int����_���ID = mIntCol��ҩ_���ID
            int����_NO = mIntCol��ҩ_NO
            int����_����� = mIntCol��ҩ_�����
        Case mListType.��ҩ
            int����_���ID = mIntCol��ҩ_���ID
            int����_NO = mIntCol��ҩ_NO
            int����_����� = mIntCol��ҩ_�����
        End Select
        
        .Redraw = flexRDNone
        
        .Cell(flexcpPicture, 1, int����_�����, .rows - 1, int����_�����) = Nothing
                
        If Not bln�Ƿ���� Then
            .ColWidth(int����_�����) = 0
            .Redraw = flexRDDirect
            Exit Sub
        Else
            .ColWidth(int����_�����) = 250
        End If
        
        For n = 1 To .rows - 1
            .Row = n
            .Col = int����_�����
            If .IsSubtotal(n) = False And .TextMatrix(n, int����_���ID) <> "" Then
                str������� = IIf(.TextMatrix(n, int����_���ID) = 0, "0", .TextMatrix(n, int����_NO) & .TextMatrix(n, int����_���ID))
                If n + 1 <= .rows - 1 Then
                    If .IsSubtotal(n + 1) = False And .TextMatrix(n + 1, int����_���ID) <> "" Then  '�������Ϊ��¼��ʱ
                        str������� = IIf(.TextMatrix(n + 1, int����_���ID) = 0, "-1", .TextMatrix(n + 1, int����_NO) & .TextMatrix(n + 1, int����_���ID))
                    ElseIf n + 2 <= .rows - 1 Then  '�������Ϊ��������ʱ
                        If .IsSubtotal(n + 2) = False And .TextMatrix(n + 2, int����_���ID) <> "" Then    '���������Ϊ��¼��ʱ
                            str������� = IIf(.TextMatrix(n + 2, int����_���ID) = 0, "-1", .TextMatrix(n + 2, int����_NO) & .TextMatrix(n + 2, int����_���ID))
                        Else
                            str������� = "-1"
                        End If
                    Else
                        str������� = "-1"
                    End If
                Else
                    str������� = "-1"
                End If
                
                If str������� = str������� Then
                    If str������� = str������� Then
                        .Cell(flexcpPicture, n, int����_�����) = imgGroup.ListImages(2).Picture
                    Else
                        .Cell(flexcpPicture, n, int����_�����) = imgGroup.ListImages(3).Picture
                    End If
                ElseIf str������� = str������� Then
                        .Cell(flexcpPicture, n, int����_�����) = imgGroup.ListImages(1).Picture
                    bln�Ƿ���ڷ��� = True
                End If
            
                str������� = IIf(str������� = "0", "-1", str�������)
            Else
                '��������ǻ����У���Ҫ�������е����ID�жϷ������
                If n + 1 <= .rows - 1 Then
                    If .IsSubtotal(n + 1) = False And .TextMatrix(n + 1, int����_���ID) <> "" Then
                        If str������� <> "-1" And str������� = IIf(.TextMatrix(n + 1, int����_���ID) = 0, "-1", .TextMatrix(n + 1, int����_NO) & .TextMatrix(n + 1, int����_���ID)) Then
                            .Cell(flexcpPicture, n, int����_�����) = imgGroup.ListImages(2).Picture
                        End If
                    End If
                End If
            End If
        Next
        
        .Cell(flexcpPictureAlignment, 1, int����_�����, .rows - 1, int����_�����) = flexAlignRightCenter
        
        If Not bln�Ƿ���ڷ��� Then .ColWidth(int����_�����) = 0
        
        .Redraw = flexRDDirect

    End With
    
End Sub

Private Sub Load��ҩ����ʽ()
    Dim rsData As ADODB.Recordset
    
    Set rsData = DeptSendWork_Get��ҩ����ʽ("ZL1_BILL_1342")
    
    With cbo��ҩ����ʽ
        .Clear
        
        Do While Not rsData.EOF
            .AddItem rsData!��ʽ
            rsData.MoveNext
        Loop
        
        If mcondition.int��ҩ����ʽ <= .ListCount - 1 And mcondition.int��ҩ����ʽ >= 0 Then
            .ListIndex = mcondition.int��ҩ����ʽ
        Else
            .ListIndex = 0
        End If
        
        If rsData.RecordCount = 1 Then
            .Enabled = False
        End If
    End With
End Sub

Private Sub Load��ҩ��(ByVal lngҩ��id As Long)
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    Set rsData = DeptSendWork_Get��ҩ��(lngҩ��id)
    
    With rsData
        Me.cbo��ҩ��.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            cbo��ҩ��.AddItem !����

            If gstrUserName = !���� Then
                intIndex = .AbsolutePosition - 1
            End If

            .MoveNext
        Loop

        cbo��ҩ��.Enabled = Not cbo��ҩ��.ListCount = 0

        If intIndex <> -1 Then cbo��ҩ��.ListIndex = intIndex
    End With
End Sub

Public Sub Load�˲���(ByVal lngҩ��id As Long)
    Dim rsData As ADODB.Recordset
    
    Set rsData = DeptSendWork_Get�˲���(lngҩ��id)
    
    With cbo�˲���
        .Clear
        
        Do While Not rsData.EOF
            .AddItem rsData!����
            rsData.MoveNext
        Loop
    End With
    
    cbo�˲���.Text = gstrUserAbbr & "-" & gstrUserName
End Sub
Public Sub RefreshList(ByVal intType As Integer, ByVal rsData As ADODB.Recordset, Optional ByVal rsChargeOff As ADODB.Recordset)
    'ˢ���б�
    
    Select Case intType
        Case mListType.��ҩ
            '������ҩ���ݼ��ĸ���
            Set mrsSendList = rsData
            
            '�����������ݼ��ĸ���
            Set mrsChargeOff = rsChargeOff
    
            '��ҩ״̬ʱ��ͬʱˢ�´���ҩ�����ܷ�ҩ��ȱҩ���ܷ����б�
            mblnRefresh = True
            
            '���ݽ���ѡ�����ı���ҩ�������ݵķ�ҩ״̬
            Modify��ҩ���� mcondition.bln��ʾ��ҩ��������
            
            Call InitList_Send
            Call RefreshList_Send
            
            Call InitList_ChargeOff
            
            Call InitList_Sum
            Call RefreshList_Sum
            
            Call InitList_Shortage
            Call RefreshList_Shortage
            
            Call InitList_Reject
            Call RefreshList_Reject
            
            mblnRefresh = False
        Case mListType.��ҩ
            '������ҩ���ݼ��ĸ���
            Set mrsReturnList = rsData
            
            Call InitList_Return
            Call RefreshList_Return
    End Select
    
    Call InitColSelList(mcondition.intListType)
    Form_Resize
End Sub


Private Sub RefreshList_Send()
    'ˢ�´���ҩ�б�
    Dim lngRow As Long
    Dim str���� As String
    Dim lngStateColor As Double
    Dim strFilter As String
    Dim i As Long
    Dim dateCurrent As Date
    
    If mrsSendList Is Nothing Then Exit Sub
    
    dateCurrent = Sys.Currentdate
    
    '�Ƿ���ʾ��ҩ����ҩƷ
    If mcondition.bln��ʾ��ҩ�������� = True Then
        '��ʾ�����ķ�ҩҩƷ
        strFilter = strFilter & "ִ��״̬=" & mState.��ҩ
        
        '�Ƿ���ʾȱҩҩƷ
        If mcondition.bln��ʾȱҩ = True Then
           strFilter = strFilter & " Or ִ��״̬=" & mState.ȱҩ
        End If
    
        '��ʾ�������ܷ���ҩƷ���ϴβ�����ҩƷ��
        strFilter = strFilter & " Or ִ��״̬=" & mState.������ & " Or ִ��״̬=" & mState.�ܷ�
    Else
        '��ʾ�����ķ�ҩҩƷ
        strFilter = "(��¼״̬=1 And ִ��״̬=" & mState.��ҩ & ")"
        
        '�Ƿ���ʾȱҩҩƷ
        If mcondition.bln��ʾȱҩ = True Then
           strFilter = strFilter & " Or (��¼״̬=1 And ִ��״̬=" & mState.ȱҩ & ")"
        End If
        
        '��ʾ�������ܷ���ҩƷ���ϴβ�����ҩƷ��
        strFilter = strFilter & " Or (��¼״̬=1 And ִ��״̬=" & mState.������ & ")" & " Or (��¼״̬=1 And ִ��״̬=" & mState.�ܷ� & ")"
    End If

    With vsfList(mListType.��ҩ)
        mrsSendList.Filter = strFilter
'        mrsSendList.Filter = "(��¼״̬=1 And ִ��״̬=1) Or (��¼״̬=1 and ִ��״̬=0) Or (��¼״̬=1 and ִ��״̬=3) Or (��¼״̬=1 and ִ��״̬=2)"
'        mrsSendList.Sort = "��ҩ����,��ҩ��,����,NO,���"
        mrsSendList.Sort = "��ҩ����,NO,���ID"
        
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
        
        .Subtotal flexSTClear
        
        If mrsSendList.EOF Then
            .rows = 2
            .Cell(flexcpText, 1, mIntCol��ҩ_����, 1, .Cols - 1) = "û���ҵ����������ļ�¼......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
        Else
            Do While Not mrsSendList.EOF
                lngRow = lngRow + 1
                .rows = lngRow + 1
                
                .TextMatrix(lngRow, mIntCol��ҩ_�����) = ""
                .TextMatrix(lngRow, mIntCol��ҩ_����) = mrsSendList!����
                .TextMatrix(lngRow, mIntCol��ҩ_����ҽ��) = IIf(IsNull(mrsSendList!����ҽ��), "", mrsSendList!����ҽ��)
                
                .TextMatrix(lngRow, mIntCol��ҩ_����) = mrsSendList!����
                .TextMatrix(lngRow, mIntCol��ҩ_��ҩ����) = IIf(InStr(1, mrsSendList!����, "3") > 1, "��Ժ��ҩ", IIf(InStr(1, mrsSendList!����, "4") > 1, "��ȡҩ", "Ժ����ҩ"))
                .TextMatrix(lngRow, mIntCol��ҩ_NO) = mrsSendList!NO
                .TextMatrix(lngRow, mIntCol��ҩ_����Ա) = mrsSendList!����Ա
                .TextMatrix(lngRow, mIntCol��ҩ_����) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
                .TextMatrix(lngRow, mIntCol��ҩ_����) = mrsSendList!����
                .TextMatrix(lngRow, mIntCol��ҩ_�Ա�) = mrsSendList!�Ա�
                .TextMatrix(lngRow, mIntCol��ҩ_��������) = zlStr.NVL(mrsSendList!��������)
                .Cell(flexcpForeColor, lngRow, mIntCol��ҩ_����, lngRow, mIntCol��ҩ_����) = zlStr.NVL(mrsSendList!��ɫ, 0)
                
                .TextMatrix(lngRow, mIntCol��ҩ_����) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
                
                .TextMatrix(lngRow, mIntCol��ҩ_סԺ��) = IIf(IsNull(mrsSendList!סԺ��), "", mrsSendList!סԺ��)
                
                If mrsSendList!������ <> 0 Then
                    .Cell(flexcpPicture, lngRow, mIntCol��ҩ_Ʒ��) = Me.ImgList.ListImages(39).Picture
                    .Cell(flexcpPictureAlignment, lngRow, mIntCol��ҩ_Ʒ��) = flexPicAlignLeftCenter
                End If
                
                If mrsSendList!��ΣҩƷ > 0 Then
                    .Cell(flexcpPicture, lngRow, mIntCol��ҩ_Ʒ��) = Me.ImgList.ListImages("��Σ").Picture
                    .Cell(flexcpPictureAlignment, lngRow, mIntCol��ҩ_Ʒ��) = flexPicAlignLeftCenter
                End If
                
                If mcondition.intҩƷ���Ʊ�����ʾ = 0 Then
                    .TextMatrix(lngRow, mIntCol��ҩ_Ʒ��) = mrsSendList!Ʒ��
                ElseIf mcondition.intҩƷ���Ʊ�����ʾ = 1 Then
                    .TextMatrix(lngRow, mIntCol��ҩ_Ʒ��) = mrsSendList!ҩƷ����
                Else
                    .TextMatrix(lngRow, mIntCol��ҩ_Ʒ��) = mrsSendList!ҩƷ����
                End If
                
                .TextMatrix(lngRow, mIntCol��ҩ_������) = IIf(IsNull(mrsSendList!������), "", mrsSendList!������)
                .TextMatrix(lngRow, mIntCol��ҩ_Ӣ����) = IIf(IsNull(mrsSendList!Ӣ����), "", mrsSendList!Ӣ����)
                .TextMatrix(lngRow, mIntCol��ҩ_�䷽����) = IIf(IsNull(mrsSendList!�䷽����), "", mrsSendList!�䷽����)
                .TextMatrix(lngRow, mIntCol��ҩ_���) = IIf(IsNull(mrsSendList!���), "", mrsSendList!���)
                .TextMatrix(lngRow, mIntCol��ҩ_������) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
                .TextMatrix(lngRow, mIntCol��ҩ_ԭ����) = IIf(IsNull(mrsSendList!ԭ����), "", mrsSendList!ԭ����)
                .TextMatrix(lngRow, mIntCol��ҩ_����) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
                .TextMatrix(lngRow, mIntCol��ҩ_Ч��) = IIf(IsNull(mrsSendList!Ч��), "", mrsSendList!Ч��)
                .TextMatrix(lngRow, mIntCol��ҩ_��) = mrsSendList!��
                .TextMatrix(lngRow, mIntCol��ҩ_����) = mrsSendList!����
                .TextMatrix(lngRow, mIntCol��ҩ_����) = Format(mrsSendList!����, "#0." & String(mintPriceDigit, "0"))
                
                .TextMatrix(lngRow, mIntCol��ҩ_���) = zlStr.FormatEx(mrsSendList!���, mintMoneyDigit, , True)
                .TextMatrix(lngRow, mIntCol��ҩ_����) = mrsSendList!����
                .TextMatrix(lngRow, mIntCol��ҩ_Ƶ��) = IIf(IsNull(mrsSendList!Ƶ��), "", mrsSendList!Ƶ��)
                .TextMatrix(lngRow, mIntCol��ҩ_�÷�) = IIf(IsNull(mrsSendList!�÷�), "", mrsSendList!�÷�)
                .TextMatrix(lngRow, mIntCol��ҩ_��ҩ����) = IIf(IsNull(mrsSendList!��ҩ����), "", mrsSendList!��ҩ����)
                .TextMatrix(lngRow, mIntCol��ҩ_��ҩĿ��) = mrsSendList!��ҩĿ��
                .TextMatrix(lngRow, mIntCol��ҩ_����ʱ��) = mrsSendList!����ʱ��
                .TextMatrix(lngRow, mIntCol��ҩ_˵��) = IIf(IsNull(mrsSendList!˵��), "", mrsSendList!˵��)
                .TextMatrix(lngRow, mIntCol��ҩ_����) = mrsSendList!����
                .TextMatrix(lngRow, mIntCol��ҩ_ҽ��id) = mrsSendList!ҽ��id
                .TextMatrix(lngRow, mIntCol��ҩ_��ҩ��) = ""
                .TextMatrix(lngRow, mIntCol��ҩ_�ⷿ��λ) = IIf(IsNull(mrsSendList!�ⷿ��λ), "", mrsSendList!�ⷿ��λ)
                
                .TextMatrix(lngRow, mIntCol��ҩ_���ID) = IIf(IsNull(mrsSendList!���ID), 0, mrsSendList!���ID)
                .TextMatrix(lngRow, mIntCol��ҩ_ҩƷID) = mrsSendList!ҩƷID
                .TextMatrix(lngRow, mIntCol��ҩ_������λ) = mrsSendList!������λ
                .TextMatrix(lngRow, mIntCol��ҩ_��ҩ����) = mrsSendList!��ҩ����
                .TextMatrix(lngRow, mIntCol��ҩ_��ҩ����id) = mrsSendList!��ҩ����ID
                
                .TextMatrix(lngRow, mIntCol��ҩ_ҩƷ���������) = mrsSendList!ҩƷ���������
                .TextMatrix(lngRow, mIntCol��ҩ_ҩƷ����) = mrsSendList!ҩƷ����
                .TextMatrix(lngRow, mIntCol��ҩ_ҩƷ����) = mrsSendList!ҩƷ����
                .TextMatrix(lngRow, mIntCol��ҩ_����ҩƷ˵��) = zlStr.NVL(mrsSendList!����ҩƷ˵��)
                
                .TextMatrix(lngRow, mIntCol��ҩ_�շ�ID) = mrsSendList!�շ�ID
                
                .TextMatrix(lngRow, mIntCol��ҩ_��ҩ��) = IIf(IsNull(mrsSendList!��ҩ��), "", mrsSendList!��ҩ��)
                .TextMatrix(lngRow, mIntCol��ҩ_���շ�) = mrsSendList!���շ�
                
                If mrsSendList!�Ƿ�Ƥ�� = 1 Then
                    .TextMatrix(lngRow, mIntCol��ҩ_Ƥ�Խ��) = GetƤ�Խ��(mrsSendList!����ID, mrsSendList!ҩ��ID, dateCurrent, mrsSendList!����ʱ��, mrsSendList!��ҳID)
                End If
                
                .TextMatrix(lngRow, mIntCol��ҩ_����ID) = mrsSendList!����ID
                .TextMatrix(lngRow, mIntCol��ҩ_��ҳID) = mrsSendList!��ҳID
                .TextMatrix(lngRow, mIntCol��ҩ_��ҩ����) = zlStr.NVL(mrsSendList!��ҩ����)
                .TextMatrix(lngRow, mIntCol��ҩ_��ΣҩƷ) = zlStr.NVL(mrsSendList!��ΣҩƷ, 0)
                .TextMatrix(lngRow, mIntCol��ҩ_���) = NVL(mrsSendList!�����, 0)
                .TextMatrix(lngRow, mIntCol��ҩ_��������id) = NVL(mrsSendList!��������id, 0)
                                
                .TextMatrix(lngRow, mIntCol��ҩ_״̬) = mrsSendList!״̬
                .TextMatrix(lngRow, mIntCol��ҩ_ִ��״̬) = mrsSendList!ִ��״̬
                            
                .TextMatrix(lngRow, mIntCol��ҩ_��ע) = NVL(mrsSendList!ҽ������, "")
                            
                '����״̬����ɫ
                If mrsSendList!ִ��״̬ = mState.ȱҩ Then
                    lngStateColor = mListColor.State_Shortage
                ElseIf mrsSendList!ִ��״̬ = mState.��ҩ Then
                    lngStateColor = mListColor.State_Send
                ElseIf mrsSendList!ִ��״̬ = mState.�ܷ� Then
                    lngStateColor = mListColor.State_Reject
                ElseIf mrsSendList!ִ��״̬ = mState.������ Then
                    lngStateColor = mListColor.State_UnProcess
                End If
                
                .Cell(flexcpBackColor, lngRow, 1, lngRow, .Cols - 1) = lngStateColor
                
'                ���ú�����ҩ��־ (PASS)
                If Not gobjPass Is Nothing Then
                    .Cell(flexcpPicture, lngRow, mIntCol��ҩ_�����, lngRow, mIntCol��ҩ_�����) = gobjPass.zlPassSetWarnLight_YF(Val(mrsSendList!�����))
                    .Cell(flexcpPictureAlignment, lngRow, mIntCol��ҩ_�����, lngRow, mIntCol��ҩ_�����) = flexPicAlignCenterCenter
                End If
                
                '�������ڲ���PASS
'                .Cell(flexcpPicture, lngRow, mIntCol��ҩ_�����, lngRow, mIntCol��ҩ_�����) = frmPublic.imgPass.ListImages(Val(mrsSendList!�����) + 2).Picture
'                .Cell(flexcpPictureAlignment, lngRow, mIntCol��ҩ_�����, lngRow, mIntCol��ҩ_�����) = flexPicAlignCenterCenter
                
                '����ҩƷ������ʾ
                If IsInString("����ҩ;����ҩ;����I��;����II��", zlStr.NVL(mrsSendList!�������), ";") = True And zlStr.NVL(mrsSendList!�������) <> "" Then
                    .Cell(flexcpFontBold, lngRow, mIntCol��ҩ_Ʒ��, lngRow, mIntCol��ҩ_Ʒ��) = True
                End If
                
                '���������
                If mcondition.blnҩƷ���� = True And mrsSendList!������� = 0 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = mListColor.LowerLimit
                End If
                
                '��ҽ���Ƿ��ڲ��ŷ�ҩ�н��й���ҩ����
                If mrsSendList!ҩʦ��˱�־ = 0 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\1345", "δ���ҽ����ɫ", 33023)
                End If
                
                mrsSendList.MoveNext
            Loop
            
            '����������Ϊ�Ӵ���ʾ
            .Cell(flexcpFontBold, 1, mIntCol��ҩ_����, .rows - 1, mIntCol��ҩ_����) = True
            
            'Ƥ�Խ����ʾ
            For i = 1 To .rows - 1
                If .IsSubtotal(i) = False Then
                    If .TextMatrix(i, mIntCol��ҩ_Ƥ�Խ��) = "(+)" Then
                        .Cell(flexcpForeColor, i, mIntCol��ҩ_Ƥ�Խ��, i, mIntCol��ҩ_Ƥ�Խ��) = vbRed
                    ElseIf .TextMatrix(i, mIntCol��ҩ_Ƥ�Խ��) = "(-)" Then
                        .Cell(flexcpForeColor, i, mIntCol��ҩ_Ƥ�Խ��, i, mIntCol��ҩ_Ƥ�Խ��) = vbBlue
                    Else
                        .Cell(flexcpForeColor, i, mIntCol��ҩ_Ƥ�Խ��, i, mIntCol��ҩ_Ƥ�Խ��) = &H80000008
                    End If
                End If
            Next
                    
            SetSubTotal vsfList(mListType.��ҩ), "��ҩ����"
        End If
        
        .Redraw = flexRDDirect
    End With
    
    SetGroup vsfList(mListType.��ҩ), True
End Sub

Private Sub RefreshList_Return()
    'ˢ����ҩ�б�
    
    Dim intRow As Integer
    Dim lngStateColor As Double
    Dim strFilter As String
    
    If mrsReturnList Is Nothing Then Exit Sub
    
    '��ʾ��������ҩҩƷ
    strFilter = "ִ��״̬=" & mState.��ҩ_ԭʼ��¼ & " And ׼����>0 "
    
    '�Ƿ���ʾ���й��̵���
    If mcondition.bln��ʾ���̵��� = True Then
        strFilter = "ִ��״̬=" & mState.��ҩ_ԭʼ��¼ & " Or ִ��״̬=" & mState.��ҩ_��ҩ��¼ & " Or ִ��״̬=" & mState.��ҩ_��ҩ��¼
    End If
    
    strFilter = strFilter & " Or ִ��״̬=" & mState.��ҩ
    
    With vsfList(mListType.��ҩ)
        mrsReturnList.Filter = strFilter
'        mrsReturnList.Sort = "����,��ҩ��,����,NO,���"
        mrsReturnList.Sort = "����,NO,���ID"
        
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
        
        If mrsReturnList.EOF Then
            .rows = 2
            .Cell(flexcpText, 1, mIntCol��ҩ_����, 1, .Cols - 1) = "û���ҵ����������ļ�¼......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
        Else
            Do While Not mrsReturnList.EOF
                intRow = intRow + 1
                .rows = intRow + 1
                
                .TextMatrix(intRow, mIntCol��ҩ_����) = mrsReturnList!����
                .TextMatrix(intRow, mIntCol��ҩ_����) = mrsReturnList!����
                .TextMatrix(intRow, mIntCol��ҩ_��ҩ����) = IIf(Right(mrsReturnList!����, 1) = 3, "��Ժ��ҩ", IIf(Right(mrsReturnList!����, 1) = 4, "��ȡҩ", "Ժ����ҩ"))
                .TextMatrix(intRow, mIntCol��ҩ_NO) = mrsReturnList!NO
                
                .TextMatrix(intRow, mIntCol��ҩ_����) = IIf(IsNull(mrsReturnList!����), "", mrsReturnList!����)
                .TextMatrix(intRow, mIntCol��ҩ_����) = mrsReturnList!����
                .TextMatrix(intRow, mIntCol��ҩ_�Ա�) = mrsReturnList!�Ա�
                .TextMatrix(intRow, mIntCol��ҩ_סԺ��) = IIf(IsNull(mrsReturnList!סԺ��), "", mrsReturnList!סԺ��)
                
                If mcondition.intҩƷ���Ʊ�����ʾ = 0 Then
                    .TextMatrix(intRow, mIntCol��ҩ_Ʒ��) = mrsReturnList!Ʒ��
                ElseIf mcondition.intҩƷ���Ʊ�����ʾ = 1 Then
                    .TextMatrix(intRow, mIntCol��ҩ_Ʒ��) = mrsReturnList!ҩƷ����
                Else
                    .TextMatrix(intRow, mIntCol��ҩ_Ʒ��) = mrsReturnList!ҩƷ����
                End If
                
                If mrsReturnList!��ΣҩƷ > 0 Then
                    .Cell(flexcpPicture, intRow, mIntCol��ҩ_Ʒ��) = Me.ImgList.ListImages("��Σ").Picture
                    .Cell(flexcpPictureAlignment, intRow, mIntCol��ҩ_Ʒ��) = flexPicAlignLeftCenter
                End If
                
                .TextMatrix(intRow, mIntCol��ҩ_������) = IIf(IsNull(mrsReturnList!������), "", mrsReturnList!������)
                
                .TextMatrix(intRow, mIntCol��ҩ_Ӣ����) = IIf(IsNull(mrsReturnList!Ӣ����), "", mrsReturnList!Ӣ����)
                .TextMatrix(intRow, mIntCol��ҩ_���) = IIf(IsNull(mrsReturnList!���), "", mrsReturnList!���)
                .TextMatrix(intRow, mIntCol��ҩ_������) = IIf(IsNull(mrsReturnList!����), "", mrsReturnList!����)
                .TextMatrix(intRow, mIntCol��ҩ_ԭ����) = IIf(IsNull(mrsReturnList!ԭ����), "", mrsReturnList!ԭ����)
                .TextMatrix(intRow, mIntCol��ҩ_����) = IIf(IsNull(mrsReturnList!����), "", mrsReturnList!����)
                .TextMatrix(intRow, mIntCol��ҩ_Ч��) = IIf(IsNull(mrsReturnList!Ч��), "", mrsReturnList!Ч��)
                .TextMatrix(intRow, mIntCol��ҩ_��) = mrsReturnList!��
                
                .TextMatrix(intRow, mIntCol��ҩ_����) = mrsReturnList!����
                .TextMatrix(intRow, mIntCol��ҩ_������) = mrsReturnList!������
                .TextMatrix(intRow, mIntCol��ҩ_׼����) = mrsReturnList!׼����
                
                If IIf(IsNull(mrsReturnList!��ҩ��), 0, mrsReturnList!��ҩ��) > 0 And mrsReturnList!ִ��״̬ = mState.��ҩ Then
                    .TextMatrix(intRow, mIntCol��ҩ_��ҩ��) = mrsReturnList!��ҩ��
                End If
                
                .TextMatrix(intRow, mIntCol��ҩ_����) = Format(mrsReturnList!����, "#0." & String(mintPriceDigit, "0"))
        
                .TextMatrix(intRow, mIntCol��ҩ_���) = Format(mrsReturnList!���, "#0." & String(mintPriceDigit, "0"))
                .TextMatrix(intRow, mIntCol��ҩ_����) = mrsReturnList!����
                .TextMatrix(intRow, mIntCol��ҩ_Ƶ��) = IIf(IsNull(mrsReturnList!Ƶ��), "", mrsReturnList!Ƶ��)
                .TextMatrix(intRow, mIntCol��ҩ_�÷�) = IIf(IsNull(mrsReturnList!�÷�), "", mrsReturnList!�÷�)
                
                .TextMatrix(intRow, mIntCol��ҩ_����Ա) = mrsReturnList!����Ա
                
                .TextMatrix(intRow, mIntCol��ҩ_��ҩʱ��) = mrsReturnList!��ҩʱ��
                .TextMatrix(intRow, mIntCol��ҩ_����) = mrsReturnList!����
                .TextMatrix(intRow, mIntCol��ҩ_ҽ��id) = mrsReturnList!ҽ��id
                .TextMatrix(intRow, mIntCol��ҩ_��ҩ��) = IIf(IsNull(mrsReturnList!��ҩ��), "", mrsReturnList!��ҩ��)
                .TextMatrix(intRow, mIntCol��ҩ_��ҩ����id) = mrsReturnList!��ҩ����ID
                .TextMatrix(intRow, mIntCol��ҩ_�ⷿ��λ) = IIf(IsNull(mrsReturnList!�ⷿ��λ), "", mrsReturnList!�ⷿ��λ)
               
                .TextMatrix(intRow, mIntCol��ҩ_���ID) = IIf(IsNull(mrsReturnList!���ID), 0, mrsReturnList!���ID)
                .TextMatrix(intRow, mIntCol��ҩ_ҩƷID) = mrsReturnList!ҩƷID
                .TextMatrix(intRow, mIntCol��ҩ_������λ) = mrsReturnList!������λ
                
                .TextMatrix(intRow, mIntCol��ҩ_ҩƷ���������) = mrsReturnList!ҩƷ���������
                .TextMatrix(intRow, mIntCol��ҩ_ҩƷ����) = mrsReturnList!ҩƷ����
                .TextMatrix(intRow, mIntCol��ҩ_ҩƷ����) = mrsReturnList!ҩƷ����
                
                .TextMatrix(intRow, mIntCol��ҩ_�շ�ID) = mrsReturnList!�շ�ID
                .TextMatrix(intRow, mIntCol��ҩ_״̬) = mrsReturnList!״̬
                .TextMatrix(intRow, mIntCol��ҩ_ִ��״̬) = mrsReturnList!ִ��״̬
                .TextMatrix(intRow, mIntCol��ҩ_��ҩ��) = mrsReturnList!��ҩ��
                .TextMatrix(intRow, mIntCol��ҩ_����ʱ��) = zlStr.NVL(mrsReturnList!����ʱ��)
                .TextMatrix(intRow, mIntCol��ҩ_��ע) = IIf(IsNull(mrsReturnList!ҽ������), "", mrsReturnList!ҽ������)
                            
                .TextMatrix(intRow, mIntCol��ҩ_����ID) = mrsReturnList!����ID
                .TextMatrix(intRow, mIntCol��ҩ_��ҳID) = mrsReturnList!��ҳID
                            
                '����״̬����ɫ
                If mrsReturnList!ִ��״̬ = mState.��ҩ_ԭʼ��¼ Then
                    lngStateColor = mListColor.Return_Original
                ElseIf mrsReturnList!ִ��״̬ = mState.��ҩ_��ҩ��¼ Then
                    lngStateColor = mListColor.Return_Sended
                ElseIf mrsReturnList!ִ��״̬ = mState.��ҩ_��ҩ��¼ Then
                    lngStateColor = mListColor.Return_Returned
                End If
                
                '���ü�¼��ǰ��ɫ
                .Cell(flexcpForeColor, intRow, 1, intRow, .Cols - 1) = lngStateColor
                
                '���ú�����ҩ��־��PASS��
                If mcondition.intShowPass = 1 Then
                    If mrsReturnList!����� > -1 And mrsReturnList!����� < 5 Then
                        .Cell(flexcpPicture, intRow, mIntCol��ҩ_�����, intRow, mIntCol��ҩ_�����) = frmPublic.imgPass.ListImages(Val(mrsReturnList!�����) + 1).Picture
                        .Cell(flexcpPictureAlignment, intRow, mIntCol��ҩ_�����, intRow, mIntCol��ҩ_�����) = flexPicAlignCenterCenter
                    End If
                ElseIf mcondition.intShowPass = 3 Then
                    If mrsReturnList!����� = 1 Then
                        .Cell(flexcpPicture, intRow, mIntCol��ҩ_�����, intRow, mIntCol��ҩ_�����) = frmPublic.imgPass.ListImages(3).Picture
                    ElseIf mrsReturnList!����� = 2 Then
                        .Cell(flexcpPicture, intRow, mIntCol��ҩ_�����, intRow, mIntCol��ҩ_�����) = frmPublic.imgPass.ListImages(2).Picture
                    ElseIf mrsReturnList!����� = 3 Then
                        .Cell(flexcpPicture, intRow, mIntCol��ҩ_�����, intRow, mIntCol��ҩ_�����) = frmPublic.imgPass.ListImages(1).Picture
                    End If
                    .Cell(flexcpPictureAlignment, intRow, mIntCol��ҩ_�����, intRow, mIntCol��ҩ_�����) = flexPicAlignCenterCenter
                End If
                
                '�������ڲ���PASS
'                .Cell(flexcpPicture, intRow, mIntCol��ҩ_�����, intRow, mIntCol��ҩ_�����) = frmPublic.imgPass.ListImages(Val(mrsReturnList!�����) + 2).Picture
'                .Cell(flexcpPictureAlignment, intRow, mIntCol��ҩ_�����, intRow, mIntCol��ҩ_�����) = flexPicAlignCenterCenter
                
                '����ҩƷ������ʾ
                If IsInString("����ҩ;����ҩ;����I��;����II��", zlStr.NVL(mrsReturnList!�������), ";") = True And zlStr.NVL(mrsReturnList!�������) <> "" Then
                    .Cell(flexcpFontBold, intRow, mIntCol��ҩ_Ʒ��, intRow, mIntCol��ҩ_Ʒ��) = True
                End If
                
                mrsReturnList.MoveNext
            Loop
            
            '����������Ϊ�Ӵ���ʾ
            .Cell(flexcpFontBold, 1, mIntCol��ҩ_��ҩ��, .rows - 1, mIntCol��ҩ_��ҩ��) = True
        End If
        
        .Redraw = flexRDDirect
    End With
    
    SetGroup vsfList(mListType.��ҩ), mcondition.bln��ʾ���̵��� = False
End Sub
Private Sub RefreshList_Reject()
    'ˢ�¾ܷ��б�
    Dim intRow As Integer
    
    If mrsSendList Is Nothing Then Exit Sub
    
    mrsSendList.Filter = "ִ��״̬=" & mState.�ܷ� & " Or ִ��״̬=" & mState.�ܷ�_������
    mrsSendList.Sort = "��ҩ����,����,NO,Ʒ��"
    
    With vsfList(mListType.�ܷ�)
        .Redraw = flexRDNone
        .rows = 1

        Do While Not mrsSendList.EOF
            intRow = intRow + 1
            .rows = intRow + 1
            
            .TextMatrix(intRow, mIntCol�ܷ�_����) = mrsSendList!����
            .TextMatrix(intRow, mIntCol�ܷ�_NO) = mrsSendList!NO
            .TextMatrix(intRow, mIntCol�ܷ�_����) = mrsSendList!����
            .TextMatrix(intRow, mIntCol�ܷ�_��ҩ����) = IIf(mrsSendList!���� = 3, "��Ժ��ҩ", IIf(mrsSendList!���� = 4, "��ȡҩ", "Ժ����ҩ"))
            .TextMatrix(intRow, mIntCol�ܷ�_����) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
            .TextMatrix(intRow, mIntCol�ܷ�_����) = mrsSendList!����
            .TextMatrix(intRow, mIntCol�ܷ�_�Ա�) = mrsSendList!�Ա�
            
            If mcondition.intҩƷ���Ʊ�����ʾ = 0 Then
                .TextMatrix(intRow, mIntCol�ܷ�_Ʒ��) = mrsSendList!Ʒ��
            ElseIf mcondition.intҩƷ���Ʊ�����ʾ = 1 Then
                .TextMatrix(intRow, mIntCol�ܷ�_Ʒ��) = mrsSendList!ҩƷ����
            Else
                .TextMatrix(intRow, mIntCol�ܷ�_Ʒ��) = mrsSendList!ҩƷ����
            End If
            
            If mrsSendList!��ΣҩƷ > 0 Then
                .Cell(flexcpPicture, intRow, mIntCol�ܷ�_Ʒ��) = Me.ImgList.ListImages("��Σ").Picture
                .Cell(flexcpPictureAlignment, intRow, mIntCol�ܷ�_Ʒ��) = flexPicAlignLeftCenter
            End If
            
            .TextMatrix(intRow, mIntCol�ܷ�_���) = IIf(IsNull(mrsSendList!���), "", mrsSendList!���)
            .TextMatrix(intRow, mIntCol�ܷ�_������) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
            .TextMatrix(intRow, mIntCol�ܷ�_ԭ����) = IIf(IsNull(mrsSendList!ԭ����), "", mrsSendList!ԭ����)
            .TextMatrix(intRow, mIntCol�ܷ�_����) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
            .TextMatrix(intRow, mIntCol�ܷ�_Ч��) = IIf(IsNull(mrsSendList!Ч��), "", mrsSendList!Ч��)
            .TextMatrix(intRow, mIntCol�ܷ�_����) = mrsSendList!����
            
            .TextMatrix(intRow, mIntCol�ܷ�_����) = Format(mrsSendList!����, "#0." & String(mintPriceDigit, "0"))
            .TextMatrix(intRow, mIntCol�ܷ�_���) = Format(mrsSendList!���, "#0." & String(mintPriceDigit, "0"))
            
            .TextMatrix(intRow, mIntCol�ܷ�_ҩƷ���������) = mrsSendList!ҩƷ���������
            .TextMatrix(intRow, mIntCol�ܷ�_ҩƷ����) = mrsSendList!ҩƷ����
            .TextMatrix(intRow, mIntCol�ܷ�_ҩƷ����) = mrsSendList!ҩƷ����
            
            .TextMatrix(intRow, mIntCol�ܷ�_ִ��״̬) = mrsSendList!ִ��״̬
            
            If mrsSendList!ִ��״̬ = mState.�ܷ� Then
                .TextMatrix(intRow, mIntCol�ܷ�_״̬) = ""
            Else
                .TextMatrix(intRow, mIntCol�ܷ�_״̬) = mrsSendList!״̬
            End If
            
            .TextMatrix(intRow, mIntCol�ܷ�_�շ�ID) = mrsSendList!�շ�ID
            .TextMatrix(intRow, mIntCol�ܷ�_��ע) = NVL(mrsSendList!ҽ������, "")
            
            '����ҩƷ������ʾ
            If IsInString("����ҩ;����ҩ;����I��;����II��", zlStr.NVL(mrsSendList!�������), ";") = True And zlStr.NVL(mrsSendList!�������) <> "" Then
                .Cell(flexcpFontBold, intRow, mIntCol�ܷ�_Ʒ��, intRow, mIntCol�ܷ�_Ʒ��) = True
            End If
            
            mrsSendList.MoveNext
        Loop
        
        .Redraw = flexRDDirect
   End With
End Sub

Private Sub RefreshList_Shortage()
    'ˢ��ȱҩ�б�
    
    Dim intRow As Integer
    
    If mrsSendList Is Nothing Then Exit Sub
    
    mrsSendList.Filter = "ִ��״̬=" & mState.ȱҩ
    mrsSendList.Sort = "��ҩ����,����,NO,Ʒ��"
    
    With vsfList(mListType.ȱҩ)
        .Redraw = flexRDNone
        .rows = 1

        Do While Not mrsSendList.EOF
            intRow = intRow + 1
            .rows = intRow + 1
            
            .TextMatrix(intRow, mIntColȱҩ_����) = mrsSendList!����
            .TextMatrix(intRow, mIntColȱҩ_NO) = mrsSendList!NO
            .TextMatrix(intRow, mIntColȱҩ_����) = mrsSendList!����
            .TextMatrix(intRow, mIntColȱҩ_��ҩ����) = IIf(mrsSendList!���� = 3, "��Ժ��ҩ", IIf(mrsSendList!���� = 4, "��ȡҩ", "Ժ����ҩ"))
            .TextMatrix(intRow, mIntColȱҩ_����) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
            .TextMatrix(intRow, mIntColȱҩ_����) = mrsSendList!����
            .TextMatrix(intRow, mIntColȱҩ_�Ա�) = mrsSendList!�Ա�
            
            If mcondition.intҩƷ���Ʊ�����ʾ = 0 Then
                .TextMatrix(intRow, mIntColȱҩ_Ʒ��) = mrsSendList!Ʒ��
            ElseIf mcondition.intҩƷ���Ʊ�����ʾ = 1 Then
                .TextMatrix(intRow, mIntColȱҩ_Ʒ��) = mrsSendList!ҩƷ����
            Else
                .TextMatrix(intRow, mIntColȱҩ_Ʒ��) = mrsSendList!ҩƷ����
            End If
            
            If mrsSendList!������ <> 0 Then
                .Cell(flexcpPicture, intRow, mIntColȱҩ_Ʒ��) = Me.ImgList.ListImages(39).Picture
                .Cell(flexcpPictureAlignment, intRow, mIntColȱҩ_Ʒ��) = flexPicAlignLeftCenter
            End If
            
            If mrsSendList!��ΣҩƷ > 0 Then
                .Cell(flexcpPicture, intRow, mIntColȱҩ_Ʒ��) = Me.ImgList.ListImages("��Σ").Picture
                .Cell(flexcpPictureAlignment, intRow, mIntColȱҩ_Ʒ��) = flexPicAlignLeftCenter
            End If
                        
            .TextMatrix(intRow, mIntColȱҩ_���) = IIf(IsNull(mrsSendList!���), "", mrsSendList!���)
            .TextMatrix(intRow, mIntColȱҩ_������) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
            .TextMatrix(intRow, mIntColȱҩ_ԭ����) = IIf(IsNull(mrsSendList!ԭ����), "", mrsSendList!ԭ����)
            .TextMatrix(intRow, mIntColȱҩ_����) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
            .TextMatrix(intRow, mIntColȱҩ_Ч��) = IIf(IsNull(mrsSendList!Ч��), "", mrsSendList!Ч��)
            .TextMatrix(intRow, mIntColȱҩ_����) = mrsSendList!����
            
            .TextMatrix(intRow, mIntColȱҩ_����) = Format(mrsSendList!����, "#0." & String(mintPriceDigit, "0"))
            .TextMatrix(intRow, mIntColȱҩ_���) = Format(mrsSendList!���, "#0." & String(mintPriceDigit, "0"))
            
            .TextMatrix(intRow, mIntColȱҩ_ҩƷ���������) = mrsSendList!ҩƷ���������
            .TextMatrix(intRow, mIntColȱҩ_ҩƷ����) = mrsSendList!ҩƷ����
            .TextMatrix(intRow, mIntColȱҩ_ҩƷ����) = mrsSendList!ҩƷ����
            .TextMatrix(intRow, mIntColȱҩ_��ע) = NVL(mrsSendList!ҽ������, "")
            
            '����ҩƷ������ʾ
            If IsInString("����ҩ;����ҩ;����I��;����II��", zlStr.NVL(mrsSendList!�������), ";") = True And zlStr.NVL(mrsSendList!�������) <> "" Then
                .Cell(flexcpFontBold, intRow, mIntColȱҩ_Ʒ��, intRow, mIntColȱҩ_Ʒ��) = True
            End If
            
            mrsSendList.MoveNext
        Loop
        
        .Redraw = flexRDDirect
   End With
End Sub

Private Function RefreshList_ChargeOff(ByVal lng����ID As Long, ByVal lngҩƷid As Long) As Boolean
    'ˢ�������б�
    Dim intRow As Integer
    Dim dblSumNum As Double
    
    If mrsChargeOff Is Nothing Then Exit Function
    
    mrsChargeOff.Filter = "��ҩ����id=" & lng����ID & " And ҩƷID=" & lngҩƷid & " And ��˱�־ = 1"
    mrsChargeOff.Sort = "NO,�շ���� Desc"
    
    If mrsChargeOff.EOF Then Exit Function
    
    With vsfChargeOff
        .Redraw = flexRDNone
        .rows = 1

        Do While Not mrsChargeOff.EOF
            intRow = intRow + 1
            .rows = intRow + 1
            
            .TextMatrix(intRow, mIntCol����_�������) = mrsChargeOff!��ҩ����
            .TextMatrix(intRow, mIntCol����_����) = mrsChargeOff!����
            .TextMatrix(intRow, mIntCol����_NO) = mrsChargeOff!NO
            .TextMatrix(intRow, mIntCol����_ҩƷID) = mrsChargeOff!ҩƷID
            .TextMatrix(intRow, mIntCol����_����ʱ��) = Format(mrsChargeOff!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(intRow, mIntCol����_������) = IIf(IsNull(mrsChargeOff!����), "", mrsChargeOff!����)
            .TextMatrix(intRow, mIntCol����_����) = IIf(IsNull(mrsChargeOff!����), "", mrsChargeOff!����)
            .TextMatrix(intRow, mIntCol����_Ч��) = Format(mrsChargeOff!Ч��, "yyyy-mm-dd")
            .TextMatrix(intRow, mIntCol����_׼������) = zlStr.FormatEx(mrsChargeOff!׼������ / mrsChargeOff!��װ, 5)
            .TextMatrix(intRow, mIntCol����_��������) = zlStr.FormatEx(mrsChargeOff!�������� / mrsChargeOff!��װ, 5)
            .TextMatrix(intRow, mIntCol����_��װ) = IIf(IsNull(mrsChargeOff!��װ), "", mrsChargeOff!��װ)
            .TextMatrix(intRow, mIntCol����_��λ) = IIf(IsNull(mrsChargeOff!��λ), "", mrsChargeOff!��λ)
            .TextMatrix(intRow, mIntCol����_�շ����) = IIf(IsNull(mrsChargeOff!�շ����), "", mrsChargeOff!�շ����)
            
            dblSumNum = dblSumNum + mrsChargeOff!�������� / mrsChargeOff!��װ
            
           mrsChargeOff.MoveNext
        Loop
        
        intRow = intRow + 1
        .rows = intRow + 1
            
        .TextMatrix(intRow, mIntCol����_NO) = "�ϼ�"
        .TextMatrix(intRow, mIntCol����_��������) = zlStr.FormatEx(dblSumNum, 5)
        
        .Redraw = flexRDDirect
   End With
   
   RefreshList_ChargeOff = True
End Function


Private Sub RefreshList_Sum()
    'ˢ�»����б�
    
    Dim str���һ��� As String
    Dim strҩƷ���� As String
    Dim dblSumNumber As Double
    Dim dblSumMoney As Double
    Dim intRow As Integer
    Dim strSum As String
    Dim intSumType As Integer
    Dim strFilter As String
    Dim n As Integer
    
    If mrsSendList Is Nothing Then Exit Sub
    
    strFilter = "ִ��״̬=" & mState.��ҩ
    
    '�Ƿ���ʾ��ҩ����ҩƷ
    If mcondition.bln��ʾ��ҩ�������� = False Then
        strFilter = strFilter & " And ��¼״̬=1 "
    End If

    mrsSendList.Filter = strFilter
    
    With vsfList(mListType.����)
        .Redraw = flexRDNone
        .rows = 1
        .MergeCells = flexMergeNever
        .Subtotal flexSTClear
        
        If mrsSendList.EOF Then
            .rows = 2
            .Cell(flexcpText, 1, IIf(mcondition.bln�����һ��� = True, mIntCol���һ���_����, mIntCol����_Ʒ��), 1, .Cols - 1) = "û���ҵ����������ļ�¼......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
        Else
            '�����ҡ�ҩƷ����
            If mcondition.bln�����һ��� = True Then
                mrsSendList.Sort = "��ҩ����,Ʒ��,����"
'                 If mcondition.bln�����λ��� = True Then
'                    .ColWidth(mIntCol���һ���_����) = 1200
                Do While Not mrsSendList.EOF
                    Debug.Print mrsSendList!��ҩ���� & mrsSendList!Ʒ�� & IIf(mcondition.bln�����λ���, IIf(IsNull(mrsSendList!����), 0, mrsSendList!����), "")
                    If str���һ��� <> mrsSendList!��ҩ���� & mrsSendList!Ʒ�� & IIf(mcondition.bln�����λ���, IIf(IsNull(mrsSendList!����), 0, mrsSendList!����), "") Then
                        intRow = intRow + 1
                        .rows = intRow + 1
                        
                        str���һ��� = mrsSendList!��ҩ���� & mrsSendList!Ʒ�� & IIf(mcondition.bln�����λ���, IIf(IsNull(mrsSendList!����), 0, mrsSendList!����), "")
                        dblSumNumber = mrsSendList!ʵ������
                        dblSumMoney = Val(mrsSendList!���)
                        
                        .TextMatrix(intRow, mIntCol���һ���_��ǰ��) = ""
                        
                        .TextMatrix(intRow, mIntCol���һ���_����) = mrsSendList!����
                        If mcondition.intҩƷ���Ʊ�����ʾ = 0 Then
                            .TextMatrix(intRow, mIntCol���һ���_Ʒ��) = mrsSendList!Ʒ��
                        ElseIf mcondition.intҩƷ���Ʊ�����ʾ = 1 Then
                            .TextMatrix(intRow, mIntCol���һ���_Ʒ��) = mrsSendList!ҩƷ����
                        Else
                            .TextMatrix(intRow, mIntCol���һ���_Ʒ��) = mrsSendList!ҩƷ����
                        End If
                        
                        If mrsSendList!������ <> 0 Then
                            .Cell(flexcpPicture, intRow, mIntCol���һ���_Ʒ��) = Me.ImgList.ListImages(39).Picture
                            .Cell(flexcpPictureAlignment, intRow, mIntCol���һ���_Ʒ��) = flexPicAlignLeftCenter
                        End If
                        
                        If mrsSendList!��ΣҩƷ > 0 Then
                            .Cell(flexcpPicture, intRow, mIntCol���һ���_Ʒ��) = Me.ImgList.ListImages("��Σ").Picture
                            .Cell(flexcpPictureAlignment, intRow, mIntCol���һ���_Ʒ��) = flexPicAlignLeftCenter
                        End If
                
                        .TextMatrix(intRow, mIntCol���һ���_���) = IIf(IsNull(mrsSendList!���), "", mrsSendList!���)
                        .TextMatrix(intRow, mIntCol���һ���_������) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
                        .TextMatrix(intRow, mIntCol���һ���_ԭ����) = IIf(IsNull(mrsSendList!ԭ����), "", mrsSendList!ԭ����)
                        .TextMatrix(intRow, mIntCol���һ���_����) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
                        .TextMatrix(intRow, mIntCol���һ���_Ч��) = IIf(IsNull(mrsSendList!Ч��), "", mrsSendList!Ч��)
                        
                        .TextMatrix(intRow, mIntCol���һ���_Ӧ������) = mrsSendList!ʵ������
                        .TextMatrix(intRow, mIntCol���һ���_��������) = zlStr.FormatEx(mrsSendList!��������, 5)
                        If mcondition.bln������ҩ���� = True Then
                            .TextMatrix(intRow, mIntCol���һ���_��������) = FormatEx(GetChargeOffCount(mrsSendList!��ҩ����ID, mrsSendList!ҩƷID, mrsSendList!����), 5)
                        Else
                            .TextMatrix(intRow, mIntCol���һ���_��������) = "0"
                        End If
                        
                        .TextMatrix(intRow, mIntCol���һ���_ʵ������) = mrsSendList!ʵ������
                        .TextMatrix(intRow, mIntCol���һ���_��λ) = mrsSendList!��λ
                        
                        .TextMatrix(intRow, mIntCol���һ���_����) = Format(mrsSendList!����, "#0." & String(mintPriceDigit, "0"))
                        .TextMatrix(intRow, mIntCol���һ���_Ӧ�����) = Format(mrsSendList!���, "#0." & String(mintPriceDigit, "0"))
                        
                        .TextMatrix(intRow, mIntCol���һ���_����) = IIf(IsNull(mrsSendList!����), 0, mrsSendList!����)
                        .TextMatrix(intRow, mIntCol���һ���_����ID) = mrsSendList!����ID
                        .TextMatrix(intRow, mIntCol���һ���_ҩƷID) = mrsSendList!ҩƷID
                        
                        .TextMatrix(intRow, mIntCol���һ���_��ҩ����) = mrsSendList!��ҩ����
                        .TextMatrix(intRow, mIntCol���һ���_��ҩ����id) = mrsSendList!��ҩ����ID
                        
                        .TextMatrix(intRow, mIntCol���һ���_ҩƷ���������) = mrsSendList!Ʒ��
                        .TextMatrix(intRow, mIntCol���һ���_ҩƷ����) = mrsSendList!ҩƷ����
                        .TextMatrix(intRow, mIntCol���һ���_ҩƷ����) = mrsSendList!ҩƷ����
                        
                        .TextMatrix(intRow, mIntCol���һ���_��װ) = mrsSendList!��װ
                        
                        '���������һ�У���һ�в��ǹ̶���ʱ������ʽ����һ�е����������
                        If intRow - 1 > 0 Then
                            .TextMatrix(intRow - 1, mIntCol���һ���_Ӧ������) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol���һ���_Ӧ������)), 5)
                            .TextMatrix(intRow - 1, mIntCol���һ���_��������) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol���һ���_��������)), 5)
                            .TextMatrix(intRow - 1, mIntCol���һ���_��������) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol���һ���_��������)), 5)
                            .TextMatrix(intRow - 1, mIntCol���һ���_ʵ������) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol���һ���_ʵ������)), 5)
                            .TextMatrix(intRow - 1, mIntCol���һ���_Ӧ�����) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol���һ���_Ӧ�����)), mintMoneyDigit, , True)
                            .TextMatrix(intRow - 1, mIntCol���һ���_ʵ�����) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol���һ���_ʵ�����)), mintMoneyDigit, , True)
                        End If
                        
                        '����ҩƷ������ʾ
                        If IsInString("����ҩ;����ҩ;����I��;����II��", zlStr.NVL(mrsSendList!�������), ";") = True And zlStr.NVL(mrsSendList!�������) <> "" Then
                            .Cell(flexcpFontBold, intRow, mIntCol���һ���_Ʒ��, intRow, mIntCol���һ���_Ʒ��) = True
                        End If
                    Else
                        dblSumNumber = dblSumNumber + mrsSendList!ʵ������
                        dblSumMoney = dblSumMoney + Val(mrsSendList!���)
                        
                        .TextMatrix(intRow, mIntCol���һ���_Ӧ������) = dblSumNumber
                        .TextMatrix(intRow, mIntCol���һ���_ʵ������) = dblSumNumber
                        .TextMatrix(intRow, mIntCol���һ���_Ӧ�����) = zlStr.FormatEx(dblSumMoney, mintMoneyDigit, , True)
                        
                        
                    End If
                    
                    mrsSendList.MoveNext
                Loop
                
                '����������Ϊ�Ӵ���ʾ
                .Cell(flexcpFontBold, 1, mIntCol���һ���_ʵ������, .rows - 1, mIntCol���һ���_ʵ������) = True
                
                'ͳ��ʵ�ʷ�ҩ����
                For n = 1 To .rows - 1
                    If .TextMatrix(n, 0) <> "С��" Then
                        'Ӧ������С��������������ʵ��Ϊ��������ʾ���ҽ�ʵ����ҩ����������Ϊ0
                        If Val(.TextMatrix(n, mIntCol���һ���_Ӧ������)) - Val(.TextMatrix(n, mIntCol���һ���_��������)) < 0 Then
                            .TextMatrix(n, mIntCol���һ���_ʵ������) = zlStr.FormatEx(Val(.TextMatrix(n, mIntCol���һ���_Ӧ������)) - Val(.TextMatrix(n, mIntCol���һ���_��������)), 5)
                            .TextMatrix(n, mIntCol���һ���_��������) = 0
                        Else
                            If Val(.TextMatrix(n, mIntCol���һ���_��������)) > 0 Then
                                '�������������Ϊ0����ҩƷ����ƻ�ȡֵ��������ʵ��Ӧ���������㣨ʵ��Ӧ����Ӧ������������������
                                If Val(.TextMatrix(n, mIntCol���һ���_��������)) > Val(.TextMatrix(n, mIntCol���һ���_Ӧ������)) - Val(.TextMatrix(n, mIntCol���һ���_��������)) Then
                                    '��������������ʵ��Ӧ��������������������ʵ��Ӧ������
                                    .TextMatrix(n, mIntCol���һ���_��������) = zlStr.FormatEx(Val(.TextMatrix(n, mIntCol���һ���_Ӧ������)) - Val(.TextMatrix(n, mIntCol���һ���_��������)), 5)
                                End If
                                
                                'ʵ��������Ӧ��������������������������
                                .TextMatrix(n, mIntCol���һ���_ʵ������) = zlStr.FormatEx(Val(.TextMatrix(n, mIntCol���һ���_Ӧ������) - Val(.TextMatrix(n, mIntCol���һ���_��������)) - Val(.TextMatrix(n, mIntCol���һ���_��������))), 5)
                            ElseIf Val(.TextMatrix(n, mIntCol���һ���_��������)) = 0 Then
                                .TextMatrix(n, mIntCol���һ���_ʵ������) = FormatEx(Val(.TextMatrix(n, mIntCol���һ���_Ӧ������) - Val(.TextMatrix(n, mIntCol���һ���_��������))), 5)
                            End If
                        End If
                        
                        .TextMatrix(n, mIntCol���һ���_ʵ�����) = zlStr.FormatEx(Val(.TextMatrix(n, mIntCol���һ���_Ӧ�����)) / Val(.TextMatrix(n, mIntCol���һ���_Ӧ������)) * Val(.TextMatrix(n, mIntCol���һ���_ʵ������)), mintMoneyDigit, , True)
                        
                        .Row = n
                        .Col = mIntCol���һ���_ʵ������
                        .CellFontBold = True
                        If Val(.TextMatrix(n, mIntCol���һ���_ʵ������)) < 0 Then
                            .CellForeColor = vbRed
                        ElseIf Val(.TextMatrix(n, mIntCol���һ���_ʵ������)) > 0 Then
                            .CellForeColor = vbBlue
                        End If
                    End If
                Next
               
                '�����λ��ܣ���ʾ�����У�
                If mcondition.bln�����λ��� = True Then
                    .ColWidth(mIntCol���һ���_����) = 1200
                Else
                    .ColWidth(mIntCol���һ���_����) = 0
                End If
                
                '����С�ƣ��ϼ�
                SetSubTotal vsfList(mListType.����), "��ҩ����"
            Else
            '��ҩƷ����
                mrsSendList.Sort = "Ʒ��"
                If mrsSendList.EOF Then mrsSendList.MoveFirst
                
                Do While Not mrsSendList.EOF
                    If strҩƷ���� <> mrsSendList!Ʒ�� & IIf(mcondition.bln�����λ���, IIf(IsNull(mrsSendList!����), 0, mrsSendList!����), "") Then
                        intRow = intRow + 1
                        .rows = intRow + 1
                        
                        strҩƷ���� = mrsSendList!Ʒ�� & IIf(mcondition.bln�����λ���, IIf(IsNull(mrsSendList!����), 0, mrsSendList!����), "")
                        dblSumNumber = mrsSendList!ʵ������
                        dblSumMoney = Val(mrsSendList!���)
                        
                        .TextMatrix(intRow, mIntCol����_��ǰ��) = ""
                        
                        If mcondition.intҩƷ���Ʊ�����ʾ = 0 Then
                            .TextMatrix(intRow, mIntCol����_Ʒ��) = mrsSendList!Ʒ��
                        ElseIf mcondition.intҩƷ���Ʊ�����ʾ = 1 Then
                            .TextMatrix(intRow, mIntCol����_Ʒ��) = mrsSendList!ҩƷ����
                        Else
                            .TextMatrix(intRow, mIntCol����_Ʒ��) = mrsSendList!ҩƷ����
                        End If
                        
                        If mrsSendList!������ <> 0 Then
                            .Cell(flexcpPicture, intRow, mIntCol����_Ʒ��) = Me.ImgList.ListImages(39).Picture
                            .Cell(flexcpPictureAlignment, intRow, mIntCol����_Ʒ��) = flexPicAlignLeftCenter
                        End If
                        
                        If mrsSendList!��ΣҩƷ > 0 Then
                            .Cell(flexcpPicture, intRow, mIntCol����_Ʒ��) = Me.ImgList.ListImages("��Σ").Picture
                            .Cell(flexcpPictureAlignment, intRow, mIntCol����_Ʒ��) = flexPicAlignLeftCenter
                        End If
                        
                        .TextMatrix(intRow, mIntCol����_���) = IIf(IsNull(mrsSendList!���), "", mrsSendList!���)
                        .TextMatrix(intRow, mIntCol����_������) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
                        .TextMatrix(intRow, mIntCol����_ԭ����) = IIf(IsNull(mrsSendList!ԭ����), "", mrsSendList!ԭ����)
                        .TextMatrix(intRow, mIntCol����_����) = IIf(IsNull(mrsSendList!����), "", mrsSendList!����)
                        .TextMatrix(intRow, mIntCol����_Ч��) = IIf(IsNull(mrsSendList!Ч��), "", mrsSendList!Ч��)
                        .TextMatrix(intRow, mIntCol����_����) = mrsSendList!ʵ������
                        .TextMatrix(intRow, mIntCol����_��λ) = mrsSendList!��λ
                        .TextMatrix(intRow, mIntCol����_����) = Format(mrsSendList!����, "#0." & String(mintPriceDigit, "0"))
                        .TextMatrix(intRow, mIntCol����_���) = Format(mrsSendList!���, "#0." & String(mintPriceDigit, "0"))
                        
                        .TextMatrix(intRow, mIntCol����_ҩƷ���������) = mrsSendList!Ʒ��
                        .TextMatrix(intRow, mIntCol����_ҩƷ����) = mrsSendList!ҩƷ����
                        .TextMatrix(intRow, mIntCol����_ҩƷ����) = mrsSendList!ҩƷ����
                        
                        '���������һ�У���һ�в��ǹ̶���ʱ������ʽ����һ�е����������
                        If intRow - 1 > 0 Then
                            .TextMatrix(intRow - 1, mIntCol����_����) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol����_����)), 5)
                            .TextMatrix(intRow - 1, mIntCol����_���) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol����_���)), mintMoneyDigit, , True)
                        End If
                        
                        '����ҩƷ������ʾ
                        If IsInString("����ҩ;����ҩ;����I��;����II��", zlStr.NVL(mrsSendList!�������), ";") = True And zlStr.NVL(mrsSendList!�������) <> "" Then
                            .Cell(flexcpFontBold, intRow, mIntCol����_Ʒ��, intRow, mIntCol����_Ʒ��) = True
                        End If
                    Else
                        dblSumNumber = dblSumNumber + mrsSendList!ʵ������
                        dblSumMoney = dblSumMoney + Val(mrsSendList!���)
                        
                        .TextMatrix(intRow, mIntCol����_����) = dblSumNumber
                        .TextMatrix(intRow, mIntCol����_���) = zlStr.FormatEx(dblSumMoney, mintMoneyDigit, , True)
                    End If
                    
                    mrsSendList.MoveNext
                Loop
                
                '����������Ϊ�Ӵ���ʾ
                .Cell(flexcpFontBold, 1, mIntCol����_����, .rows - 1, mIntCol����_����) = True
                
                '�����λ��ܣ���ʾ�����У�
                If mcondition.bln�����λ��� = True Then
                    .ColWidth(mIntCol����_����) = 1200
                Else
                    .ColWidth(mIntCol����_����) = 0
                End If
                
                '����С�ƣ��ϼ�
                SetSubTotal vsfList(mListType.����), ""
            End If
            .Row = 1
            vsfChargeOff_EnterCell
        End If
        
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub ResizeChargeOffList()
    On Error Resume Next
    
    vsfList(mListType.����).Height = mdblSumListHeight

    If vsfChargeOff.Visible = True Then
        vsfList(mListType.����).Height = mdblSumListHeight / 4 * 3
                        
        picHsc.Top = vsfList(mListType.����).Top + vsfList(mListType.����).Height
        picHsc.Left = vsfList(mListType.����).Left
        picHsc.Width = vsfList(mListType.����).Width
        
        vsfChargeOff.Top = picHsc.Top + picHsc.Height
        vsfChargeOff.Left = vsfList(mListType.����).Left
        vsfChargeOff.Height = mdblSumListHeight / 4 - picHsc.Height
        vsfChargeOff.Width = vsfList(mListType.����).Width
    End If
End Sub

Public Sub SetParams()
    Dim intType As Integer
    
    With mcondition
        .lngҩ��id = Val(zlDataBase.GetPara("��ҩҩ��", glngSys, 1342))
        .bln�����һ��� = (Val(zlDataBase.GetPara("�����һ�����ʾ�����嵥", glngSys, 1342)) = 1)
        .bln������ҩ���� = (Val(zlDataBase.GetPara("��ҩʱ������ҩ���ʼ�¼", glngSys, 1342, 0)) = 1) And IsInString(gstrprivs, "��ҩ����", ";")
        .bln����δ��˴�����ҩ = (gtype_UserSysParms.P6_δ��˼��ʴ�����ҩ = 1)
        .intҩƷ���Ʊ�����ʾ = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ���ŷ�ҩ����", "ҩƷ������ʾ��ʽ", 0)))
        .int��ҩ����ʽ = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ���ŷ�ҩ����", "��ҩ����ʽ", 0)))
        .bln�������� = CheckIsCenter(.lngҩ��id)
        .str��Σ���� = zlDataBase.GetPara("��ΣҩƷ����", glngSys, 1342, "")
        .str��Σ���� = zlDataBase.GetPara("��Σ����", glngSys, 1342, "")
        .int��ҩ��������Ĭ��Ϊ��ҩ״̬ = Val(zlDataBase.GetPara("��ҩ��������Ĭ��Ϊ��ҩ״̬", glngSys, 1342, ""))
        
        .bln��ʾԭ���� = Is��ҩ�ⷿ(.lngҩ��id)
        
        If .intҩƷ���Ʊ�����ʾ > 2 Or .intҩƷ���Ʊ�����ʾ < 0 Then .intҩƷ���Ʊ�����ʾ = 0
        If .blnҩƷ���� <> (Val(zlDataBase.GetPara("�ⷿ��λ�����������ʾ", glngSys, 1342, 0)) = 1) Then
            .blnҩƷ���� = (Val(zlDataBase.GetPara("�ⷿ��λ�����������ʾ", glngSys, 1342, 0)) = 1)
            
            intType = mcondition.intListType
            
            If intType = mListType.��ҩ Or intType = mListType.��ҩ Then
                If .blnҩƷ���� = True Then
                    mstrUnallowShow(intType) = Replace(mstrUnallowShow(intType), ";�ⷿ��λ", "")
                    vsfList(intType).ColWidth(mIntCol��ҩ_�ⷿ��λ) = 1200
                Else
                    mstrUnallowShow(intType) = mstrUnallowShow(intType) & ";�ⷿ��λ"
                    vsfList(intType).ColWidth(mIntCol��ҩ_�ⷿ��λ) = 0
                End If
                
                InitColSelList intType
            End If
        End If
        Call GetDrugDigit(.lngҩ��id, "ҩƷ������ҩ", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End With
    
    
End Sub

Private Sub SetPassMenuButton(ByVal intListType As Integer, ByVal lngRow As Long)
    '����cmdAlley��ť״̬
    Dim cbrControl As CommandBarControl
    Dim rsData As ADODB.Recordset
    
    If mcondition.intShowPass <> 1 Or Not IsInString(gstrprivs, "������ҩ���", ";") Then Exit Sub
    
    '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�����Ͳ���ʾcmdAlley��ť
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] "
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, vsfList(intListType).TextMatrix(lngRow, vsfList(intListType).ColIndex("NO")), Val(vsfList(intListType).TextMatrix(lngRow, vsfList(intListType).ColIndex("����"))))
    
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, conMenu_Tool_ShowPlug, , True)
    
    If rsData.RecordCount = 0 Then
        If Not cbrControl Is Nothing Then cbrControl.Enabled = False
    Else
        If Not cbrControl Is Nothing Then cbrControl.Enabled = True
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetSendBillState(ByVal intChangeType As Integer, Optional ByVal intRow As Integer = 0)
    Dim intState As Integer
    Dim strState As String
    Dim lngColor As Long
    Dim i As Long
    Dim lng���ID As Long
    Dim strNo As String
    
    If (intChangeType = mChangeState.��ҩ Or intChangeType = mChangeState.�ܷ� Or intChangeType = mChangeState.ȱҩ _
        Or intChangeType = mChangeState.������) And intRow = 0 Then Exit Sub
    
    With vsfList(mListType.��ҩ)
        '������¼����ѡ�Ķ�����¼�ı�״̬
        If intChangeType = mChangeState.��ҩ Or intChangeType = mChangeState.�ܷ� Or intChangeType = mChangeState.������ Then
            Select Case intChangeType
                Case mChangeState.��ҩ
                    intState = mState.��ҩ
                    strState = "��ҩ"
                    lngColor = mListColor.State_Send
                Case mChangeState.�ܷ�
                    intState = mState.�ܷ�
                    strState = "�ܷ�"
                    lngColor = mListColor.State_Reject
                Case mChangeState.������
                    intState = mState.������
                    strState = "������"
                    lngColor = mListColor.State_UnProcess
            End Select
            
            For i = 1 To .rows - 1
                If .IsSubtotal(i) = False And .IsSelected(i) = True And Val(.TextMatrix(i, mIntCol��ҩ_ִ��״̬)) <> intState And Val(.TextMatrix(i, mIntCol��ҩ_ִ��״̬)) <> mState.ȱҩ Then
                    .TextMatrix(i, mIntCol��ҩ_ִ��״̬) = intState
                    .TextMatrix(i, mIntCol��ҩ_״̬) = strState
                    
                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                    
                    mrsSendList.Filter = "�շ�ID=" & Val(.TextMatrix(i, mIntCol��ҩ_�շ�ID))
                    
                    mrsSendList!ִ��״̬ = intState
                    mrsSendList!״̬ = strState
                    
                    mrsSendList.Update
                    
                    mblnSendChange = True
                End If
            Next
            
            'ͬ��ҽ����ҩƷ״̬��Ҫͬ���ı䣻����Ǹ�ΣҩƷ������Ҫ�󵥶�����ʱ��ͬ���ı�
            If mcondition.bln�������� = True And InStr(1, mcondition.str��Σ����, .TextMatrix(intRow, mIntCol��ҩ_��ΣҩƷ)) = 0 Then
                strNo = .TextMatrix(intRow, mIntCol��ҩ_NO)
                lng���ID = Val(.TextMatrix(intRow, mIntCol��ҩ_���ID))
                If lng���ID > 0 Then
                    For i = 1 To .rows - 1
                        If .IsSubtotal(i) = False And .TextMatrix(i, mIntCol��ҩ_NO) = strNo And Val(.TextMatrix(i, mIntCol��ҩ_���ID)) = lng���ID _
                            And Val(.TextMatrix(i, mIntCol��ҩ_ִ��״̬)) <> intState And Val(.TextMatrix(i, mIntCol��ҩ_ִ��״̬)) <> mState.ȱҩ _
                            And InStr(1, mcondition.str��Σ����, .TextMatrix(i, mIntCol��ҩ_��ΣҩƷ)) = 0 Then
                            .TextMatrix(i, mIntCol��ҩ_ִ��״̬) = intState
                            .TextMatrix(i, mIntCol��ҩ_״̬) = strState
                            
                            .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                            
                            mrsSendList.Filter = "�շ�ID=" & Val(.TextMatrix(i, mIntCol��ҩ_�շ�ID))
                            
                            mrsSendList!ִ��״̬ = intState
                            mrsSendList!״̬ = strState
                            
                            mrsSendList.Update
                            
                            mblnSendChange = True
                        End If
                    Next
                End If
            End If
            
            SetMainComandBars mListType.��ҩ, intRow
        Else
        '���д���ҩ����ļ�¼�ı�״̬
            Select Case intChangeType
                Case mChangeState.ȫ����ҩ
                    intState = mState.��ҩ
                    strState = "��ҩ"
                    lngColor = mListColor.State_Send
                Case mChangeState.ȫ���ܷ�
                    intState = mState.�ܷ�
                    strState = "�ܷ�"
                    lngColor = mListColor.State_Reject
                Case mChangeState.ȫ��������
                    intState = mState.������
                    strState = "������"
                    lngColor = mListColor.State_UnProcess
            End Select
            
            For i = 1 To .rows - 1
                If .IsSubtotal(i) = False And Val(.TextMatrix(i, mIntCol��ҩ_ִ��״̬)) <> mState.ȱҩ Then
                    .TextMatrix(i, mIntCol��ҩ_ִ��״̬) = intState
                    .TextMatrix(i, mIntCol��ҩ_״̬) = strState
                
                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                    
                    mrsSendList.Filter = "�շ�ID=" & Val(.TextMatrix(i, mIntCol��ҩ_�շ�ID))
                    
                    mrsSendList!ִ��״̬ = intState
                    mrsSendList!״̬ = strState
                    
                    mrsSendList.Update
                End If
            Next
            
            mblnSendChange = True
        End If
    End With
End Sub

Private Sub SetMainComandBars(ByVal intListType As Integer, ByVal lngRow As Long)
    '���ݵ�ǰ��¼�嵥���ͼ���ǰ��¼������������Ĳ˵�״̬
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim blnExists As Boolean
    
    If lngRow = 0 Then Exit Sub
    
    Select Case intListType
        Case mListType.��ҩ
            '���ܷ���״̬���л�
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol��ҩ_�շ�ID)) = 0 Then Exit Sub
            
            Set cbrMenu = frm���ŷ�ҩ����New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            Set cbrControl = frm���ŷ�ҩ����New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol��ҩ_ִ��״̬)) = mState.�ܷ� Then
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                If Not cbrControl Is Nothing Then cbrControl.Enabled = True
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            End If
            
            Set cbrMenu = frm���ŷ�ҩ����New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            Set cbrControl = frm���ŷ�ҩ����New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ Then
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                If Not cbrControl Is Nothing Then cbrControl.Enabled = True
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            End If
        Case mListType.�ܷ�
            '���ܷ��������ָ���״̬���л�
            Set cbrMenu = frm���ŷ�ҩ����New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            Set cbrControl = frm���ŷ�ҩ����New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol�ܷ�_�շ�ID)) = 0 Then
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol�ܷ�_ִ��״̬)) = mState.�ܷ� Then
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = True
                Else
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = False
                End If
            End If
            
            Set cbrMenu = frm���ŷ�ҩ����New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            Set cbrControl = frm���ŷ�ҩ����New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol�ܷ�_�շ�ID)) = 0 Then
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol�ܷ�_ִ��״̬)) = mState.�ܷ�_�ָ� Then
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = True
                Else
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = False
                End If
            End If
        Case mListType.��ҩ
            If vsfList(intListType).TextMatrix(lngRow, mIntCol��ҩ_�շ�ID) = "" Or (Not IsNumeric(vsfList(intListType).TextMatrix(lngRow, mIntCol��ҩ_�շ�ID))) Then
                Exit Sub
            End If
            
            Set cbrMenu = frm���ŷ�ҩ����New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = frm���ŷ�ҩ����New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            
            With vsfList(intListType)
                blnExists = RecipeSendWork_JudgeSign(.TextMatrix(lngRow, mIntCol��ҩ_����), .TextMatrix(lngRow, mIntCol��ҩ_NO), IIf(Val(.TextMatrix(lngRow, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ_ԭʼ��¼, 2, 3), .TextMatrix(lngRow, mIntCol��ҩ_�շ�ID))
                
                If blnExists Then
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = True
                Else
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = False
                End If
            End With
    End Select
End Sub

Private Sub SetSubTotal(ByVal vsfObj As VSFlexGrid, ByVal strSub As String)
    Dim lngRow As Long
    Dim intCol As Integer
    Dim intSumType As Integer
    Dim strSum As String
    
    '����С�ƣ��ϼ�
    With vsfObj
        .OutlineCol = 0
        .OutlineBar = 0
        .SubtotalPosition = flexSTBelow
        
        If .index = mListType.��ҩ Then
            .Subtotal flexSTSum, -1, mIntCol��ҩ_���, "###.00", , mListColor.SumTotal, True
            If strSub = "��ҩ����" Then
                .Subtotal flexSTSum, mIntCol��ҩ_��ҩ����, mIntCol��ҩ_���, "###.00", , mListColor.SumTotal, False, "", mIntCol��ҩ_��ҩ����, False
                .Subtotal flexSTSum, mIntCol��ҩ_NO, mIntCol��ҩ_���, "###.00", , mListColor.SumTotal, False, "", mIntCol��ҩ_NO, False
            ElseIf strSub = "NO" Then
                .Subtotal flexSTSum, mIntCol��ҩ_NO, mIntCol��ҩ_���, "###.00", , mListColor.SumTotal, False, "", mIntCol��ҩ_NO, False
            ElseIf strSub = "ҩƷ����" Then
                .Subtotal flexSTSum, mIntCol��ҩ_Ʒ��, mIntCol��ҩ_���, "###.00", , mListColor.SumTotal, False, "", mIntCol��ҩ_Ʒ��, False
            ElseIf strSub = "סԺ��" Then
                .Subtotal flexSTSum, mIntCol��ҩ_סԺ��, mIntCol��ҩ_���, "###.00", , mListColor.SumTotal, False, "", mIntCol��ҩ_סԺ��, False
            ElseIf strSub = "����" Then
                .Subtotal flexSTSum, mIntCol��ҩ_����, mIntCol��ҩ_���, "###.00", , mListColor.SumTotal, False, "", mIntCol��ҩ_����, False
            ElseIf strSub = "����" Then
                .Subtotal flexSTSum, mIntCol��ҩ_����, mIntCol��ҩ_���, "###.00", , mListColor.SumTotal, False, "", mIntCol��ҩ_����, False
            End If
        ElseIf .index = mListType.���� Then
            
            If mcondition.bln�����һ��� = True Then
                .Subtotal flexSTSum, -1, mIntCol���һ���_Ӧ�����, "###.00", , mListColor.SumTotal, True
                If strSub = "��ҩ����" Then
                    .Subtotal flexSTSum, mIntCol���һ���_��ҩ����, mIntCol���һ���_Ӧ�����, "###.00", , mListColor.SumTotal, False, "", mIntCol���һ���_��ҩ����, False
                ElseIf strSub = "ҩƷ����" Then
                    .Subtotal flexSTSum, mIntCol���һ���_Ʒ��, mIntCol���һ���_Ӧ�����, "###.00", , mListColor.SumTotal, False, "", mIntCol���һ���_Ʒ��, False
                End If
            Else
                .Subtotal flexSTSum, -1, mIntCol����_���, "###.00", , mListColor.SumTotal, True
'                If strSub = "ҩƷ����" Then
'                    .Subtotal flexSTSum, mIntCol����_Ʒ��, mIntCol���һ���_Ӧ�����, "###.00", , mListColor.SumTotal, False, "", mIntCol����_Ʒ��, False
'                End If
            End If
        End If
        
        For lngRow = 1 To .rows - 1
            If .IsSubtotal(lngRow) = True Then
                '�ҵ�����һ�кϼ�
                If lngRow = .rows - 1 Then
                    '��������һ�У����Ǻϼ�
                    strSum = ""
                    intSumType = mSubTotalType.SubSum
                Else
                    '������ǣ����������ɫ�жϺϼƵ�����
                    For intCol = 1 To .Cols - 1
                        .Row = lngRow
                        .Col = intCol
                        If .CellForeColor = mListColor.SumTotal Then
                            If intCol = .ColIndex("��ҩ����") Then
                                intSumType = mSubTotalType.SubByDept
                            ElseIf intCol = .ColIndex("����") Then
                                intSumType = mSubTotalType.SubByPeople
                            ElseIf intCol = .ColIndex("NO") Then
                                intSumType = mSubTotalType.SubByNo
                            ElseIf intCol = .ColIndex("ҩƷ����") Then
                                intSumType = mSubTotalType.SubByDrug
                            ElseIf intCol = .ColIndex("סԺ��") Then
                                intSumType = mSubTotalType.SubByHosNumber
                            ElseIf intCol = .ColIndex("����") Then
                                intSumType = mSubTotalType.SubByBedNumber
                            End If

                            strSum = Trim(Replace(.TextMatrix(lngRow, intCol), "Total", ""))

                            Exit For
                        End If
                    Next
                End If

                SetGridSubTotal vsfObj, lngRow, strSum, mrsSendList, intSumType
            End If
        Next
    End With
End Sub

Private Sub SetTempOperate(ByVal intLastType As Integer, ByVal intThisType As Integer)
    '�������ȡ��ʱ�Ĳ���
    Dim strValue As String
    
    '�����ϸ�ҳ��Ĳ���
    Select Case intLastType
        Case mListType.��ҩ
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾȱҩҩƷ", IIf(mcondition.bln��ʾȱҩ, 1, 0)
            
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ��ҩ��������", IIf(mcondition.bln��ʾ��ҩ��������, 1, 0)
            
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ�����Ϣ", IIf(mcondition.bln��ʾ��չ��Ϣ, 1, 0)
        Case mListType.����
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\����", "�����λ���", IIf(mcondition.bln�����λ���, 1, 0)
        Case mListType.��ҩ
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ���̵���", IIf(mcondition.bln��ʾ���̵���, 1, 0)
    End Select
    
    '�����ϸ�ҳ���������
    SaveListColState intLastType
    
    'ȡ��ǰҳ��Ĳ���
    Select Case intThisType
        Case mListType.��ҩ
            strValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾȱҩҩƷ", "1")
            mcondition.bln��ʾȱҩ = (Val(strValue) = 1)
            
            strValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ��ҩ��������", "1")
            mcondition.bln��ʾ��ҩ�������� = (Val(strValue) = 1)
        Case mListType.����
            strValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\����", "�����λ���", "0")
            mcondition.bln�����λ��� = (Val(strValue) = 1)
        Case mListType.��ҩ
            strValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ���̵���", "0")
            mcondition.bln��ʾ���̵��� = (Val(strValue) = 1)
    End Select
End Sub

Private Sub SetGridSubTotal(ByVal vsfObj As VSFlexGrid, ByVal intRow As Integer, _
            ByVal strSub As String, ByVal rsData As ADODB.Recordset, ByVal intSubType As Integer)
    '���ڲ������С�ơ��ϼƣ�ͳ�Ƶ������ɺϼ����;�����
    '�ϼƣ�ͳ�ƿ���������������������������������ҩƷ���������
    '����ҩ����С�ƣ�ͳ�Ʋ�����������������������ҩƷ���������
    '������С�ƣ�ͳ�Ƶ�������������ҩƷ���������
    '������С�ƣ�ͳ�ƴ���ҩƷ���������
    '��Ʒ��С�ƣ�ͳ�Ʋ����������������������
    '������������סԺ�ţ�����С�ƣ���������������ҩƷ���������
    Dim rsSub As ADODB.Recordset
    
    Dim str��ǰҩƷ As String
    Dim str��ǰNO As String
    Dim lng��ǰ���� As Long
    Dim str��ǰ��ҩ���� As String
    
    Dim dbl����ҩƷ���� As Double
    Dim dblNO���� As Double
    Dim dbl�������� As Double
    Dim dbl�������� As Double
    Dim Dbl��� As Double
    Dim dblʵ����� As Double
    Dim dbl������ As Double
    Dim strTemp As String
    Dim str����id As String
    Dim str���� As String
    
    
    Dim strSumText As String
    
    Dim intSumCol As Integer
    
    Dim strFilter As String
    
    '�Ƿ���ʾ��ҩ����ҩƷ
    If mcondition.bln��ʾ��ҩ�������� = False Then
        strFilter = " And ��¼״̬=1 "
    End If
    
    Set rsSub = rsData
    
    vsfObj.MergeCells = flexMergeRestrictRows
    vsfObj.MergeRow(intRow) = True
    
    Select Case intSubType
        Case mSubTotalType.SubByNo
            '�����ݺ�С��
            rsSub.Filter = "NO='" & strSub & "'" & strFilter
            rsSub.Sort = "Ʒ��"
            
            str��ǰҩƷ = ""
            dbl����ҩƷ���� = 0
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        If str��ǰҩƷ <> rsSub!Ʒ�� Then
                            str��ǰҩƷ = rsSub!Ʒ��
                            dbl����ҩƷ���� = dbl����ҩƷ���� + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                Dbl��� = 0
                rsSub.MoveFirst
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        Dbl��� = Dbl��� + rsSub!���
                    End If
                    rsSub.MoveNext
                Loop
            End If
        Case mSubTotalType.SubByPeople, mSubTotalType.SubByHosNumber, mSubTotalType.SubByBedNumber
            '������������סԺ�ţ����Ż���ʱ
            If intSubType = mSubTotalType.SubByPeople Then
                rsSub.Filter = "����='" & strSub & "'" & strFilter
            ElseIf intSubType = mSubTotalType.SubByHosNumber Then
                rsSub.Filter = "סԺ��='" & strSub & "'" & strFilter
            ElseIf intSubType = mSubTotalType.SubByBedNumber Then
                rsSub.Filter = "����='" & strSub & "'" & strFilter
            End If
            
            rsSub.Sort = "NO"
            
            str��ǰҩƷ = ""
            str��ǰNO = ""
            dblNO���� = 0
            dbl����ҩƷ���� = 0
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        If InStr(1, "," & str��ǰNO, "," & rsSub!NO & ",") < 1 Then
                            str��ǰNO = str��ǰNO & rsSub!NO & ","
                            dblNO���� = dblNO���� + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                rsSub.Sort = "Ʒ��"
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        If str��ǰҩƷ <> rsSub!Ʒ�� Then
                            str��ǰҩƷ = rsSub!Ʒ��
                            dbl����ҩƷ���� = dbl����ҩƷ���� + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                Dbl��� = 0
                rsSub.MoveFirst
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        Dbl��� = Dbl��� + rsSub!���
                    End If
                    rsSub.MoveNext
                Loop
            End If
        Case mSubTotalType.SubByDept
            '����ҩ����С��
            rsSub.Filter = "��ҩ����='" & strSub & "'" & strFilter
            rsSub.Sort = "����,NO"
            
            lng��ǰ���� = 0
            str��ǰҩƷ = ""
            str��ǰNO = ""
            dbl�������� = 0
            dblNO���� = 0
            dbl����ҩƷ���� = 0
            str����id = ""
            str��ǰNO = ""
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        If InStr(1, "," & str����id, "," & rsSub!����ID & ",") < 1 Or InStr(1, "," & str����, "," & rsSub!���� & ",") < 1 Then
                            str����id = str����id & rsSub!����ID & ","
                            str���� = str���� & rsSub!���� & ","
                            dbl�������� = dbl�������� + 1
                        End If
                   
                        If InStr(1, "," & str��ǰNO, "," & rsSub!NO & ",") < 1 Then
                            str��ǰNO = str��ǰNO & rsSub!NO & ","
                            dblNO���� = dblNO���� + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                rsSub.Sort = "Ʒ��,����"
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        If str��ǰҩƷ <> rsSub!Ʒ�� Then
                            str��ǰҩƷ = rsSub!Ʒ��
                            dbl����ҩƷ���� = dbl����ҩƷ���� + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                dbl������ = 0
                Dbl��� = 0
                rsSub.MoveFirst
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        Dbl��� = Dbl��� + rsSub!���
                        
                        If vsfObj.index = mListType.���� Then
                            If mcondition.bln�����λ��� Then
                                If InStr(1, strTemp & ",", rsSub!��ҩ����ID & "," & rsSub!ҩƷID & "," & rsSub!���� & ",") < 1 Then
                                    dbl������ = dbl������ + ((rsSub!�������� + FormatEx(GetChargeOffCount(rsSub!��ҩ����ID, rsSub!ҩƷID, rsSub!����), 5)) * rsSub!����)
                                    strTemp = strTemp & rsSub!��ҩ����ID & "," & rsSub!ҩƷID & "," & rsSub!����
                                End If
                            Else
                                If InStr(1, strTemp & ",", rsSub!��ҩ����ID & "," & rsSub!ҩƷID & ",") < 1 Then
                                
                                    dbl������ = dbl������ + ((rsSub!�������� + FormatEx(GetChargeOffCount(rsSub!��ҩ����ID, rsSub!ҩƷID, rsSub!����), 5)) * rsSub!����)
                                    strTemp = strTemp & rsSub!��ҩ����ID & "," & rsSub!ҩƷID
        
                                End If
                            End If
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
            End If
        Case mSubTotalType.SubByDrug
            '��ҩƷС��
            rsSub.Filter = "Ʒ��='" & strSub & "'" & strFilter
            rsSub.Sort = "����,����,NO"
            
            lng��ǰ���� = 0
            str��ǰNO = ""
            dbl�������� = 0
            dblNO���� = 0
            dbl����ҩƷ���� = 0
            str����id = ""
            str��ǰNO = ""
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        If InStr(1, "," & str����id, "," & rsSub!����ID & ",") < 1 Or InStr(1, "," & str����, "," & rsSub!���� & ",") < 1 Then
                            str����id = str����id & rsSub!����ID & ","
                            str���� = str���� & rsSub!���� & ","
                            dbl�������� = dbl�������� + 1
                        End If
                   
                        If InStr(1, "," & str��ǰNO, "," & rsSub!NO & ",") < 1 Then
                            str��ǰNO = str��ǰNO & rsSub!NO & ","
                            dblNO���� = dblNO���� + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                Dbl��� = 0
                dbl������ = 0
                rsSub.MoveFirst
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        Dbl��� = Dbl��� + rsSub!���
                        
                        If vsfObj.index = mListType.���� Then
                            If mcondition.bln�����λ��� Then
                                If InStr(1, strTemp & ",", rsSub!��ҩ����ID & "," & rsSub!ҩƷID & "," & rsSub!���� & ",") < 1 Then
                                    dbl������ = dbl������ + ((rsSub!�������� + FormatEx(GetChargeOffCount(rsSub!��ҩ����ID, rsSub!ҩƷID, rsSub!����), 5)) * rsSub!����)
                                    strTemp = strTemp & rsSub!��ҩ����ID & "," & rsSub!ҩƷID & "," & rsSub!����
                                End If
                            Else
                                If InStr(1, strTemp & ",", rsSub!��ҩ����ID & "," & rsSub!ҩƷID & ",") < 1 Then
                                
                                    dbl������ = dbl������ + ((rsSub!�������� + FormatEx(GetChargeOffCount(rsSub!��ҩ����ID, rsSub!ҩƷID, rsSub!����), 5)) * rsSub!����)
                                    strTemp = strTemp & rsSub!��ҩ����ID & "," & rsSub!ҩƷID
        
                                End If
                            End If
                        End If
                    End If
                    rsSub.MoveNext
                Loop
            End If
        Case mSubTotalType.SubSum
            '�ϼ�
            rsSub.Filter = IIf(mcondition.bln��ʾ��ҩ�������� = False, "��¼״̬=1", "")
            rsSub.Sort = "��ҩ����,ҩƷid,����,����,NO"
            
            str��ǰ��ҩ���� = ""
            lng��ǰ���� = 0
            str��ǰҩƷ = ""
            str��ǰNO = ""
            dbl�������� = 0
            dbl�������� = 0
            dblNO���� = 0
            dbl����ҩƷ���� = 0
            Dbl��� = 0
            dbl������ = 0
            str����id = ""
            str��ǰNO = ""
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        If str��ǰ��ҩ���� <> rsSub!��ҩ���� Then
                            str��ǰ��ҩ���� = rsSub!��ҩ����
                            dbl�������� = dbl�������� + 1
                        End If
                        
                        If InStr(1, "," & str����id, "," & rsSub!����ID & ",") < 1 Or InStr(1, "," & str����, "," & rsSub!���� & ",") < 1 Then
                            str����id = str����id & rsSub!����ID & ","
                            str���� = str���� & rsSub!���� & ","
                            dbl�������� = dbl�������� + 1
                        End If
                        
                        If InStr(1, "," & str��ǰNO, "," & rsSub!NO & ",") < 1 Then
                            str��ǰNO = str��ǰNO & rsSub!NO & ","
                            dblNO���� = dblNO���� + 1
                        End If
                        
                        Dbl��� = Dbl��� + rsSub!���
                        
                        If vsfObj.index = mListType.���� Then
                            If mcondition.bln�����λ��� Then
                                If InStr(1, strTemp & ",", rsSub!��ҩ����ID & "," & rsSub!ҩƷID & "," & rsSub!���� & ",") < 1 Then
                                    dbl������ = dbl������ + ((rsSub!�������� + FormatEx(GetChargeOffCount(rsSub!��ҩ����ID, rsSub!ҩƷID, rsSub!����), 5)) * rsSub!����)
                                    strTemp = strTemp & rsSub!��ҩ����ID & "," & rsSub!ҩƷID & "," & rsSub!����
                                End If
                            Else
                                If InStr(1, strTemp & ",", rsSub!��ҩ����ID & "," & rsSub!ҩƷID & ",") < 1 Then
                                
                                    dbl������ = dbl������ + ((rsSub!�������� + FormatEx(GetChargeOffCount(rsSub!��ҩ����ID, rsSub!ҩƷID, rsSub!����), 5)) * rsSub!����)
                                    strTemp = strTemp & rsSub!��ҩ����ID & "," & rsSub!ҩƷID
        
                                End If
                            End If
                        End If
                    End If
                    
                    rsSub.MoveNext
                Loop
                
                rsSub.Sort = "Ʒ��"
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.���� And rsSub!ִ��״̬ = mState.��ҩ) Or _
                        (vsfObj.index <> mListType.���� And (mcondition.bln��ʾȱҩ = True And rsSub!ִ��״̬ <= 3) Or (mcondition.bln��ʾȱҩ = False And ((rsSub!ִ��״̬ = mState.��ҩ Or rsSub!ִ��״̬ = mState.������ Or rsSub!ִ��״̬ = mState.�ܷ�)))) Then
                        If str��ǰҩƷ <> rsSub!Ʒ�� Then
                            str��ǰҩƷ = rsSub!Ʒ��
                            dbl����ҩƷ���� = dbl����ҩƷ���� + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
            End If
    End Select
    
    If intSubType = mSubTotalType.SubSum Then
        If vsfObj.index = mListType.��ҩ Then
            strSumText = "�ϼƣ� " & dbl�������� & "������  " & dbl�������� & "������  " & dblNO���� & "�ŵ���  " & dbl����ҩƷ���� & "��ҩƷ���� " & "����" & Format(Dbl���, "#####0.00;-#####0.00; ;") & "Ԫ"
        ElseIf vsfObj.index = mListType.���� Then
            If mcondition.bln�����һ��� = True Then
                strSumText = "�ϼƣ� " & dbl�������� & "������  " & dbl����ҩƷ���� & "��ҩƷ���� " & "Ӧ������" & Format(Dbl���, "#####0.00;-#####0.00; ;") & "Ԫ ʵ������" & Format(Dbl��� - dbl������, "#####0.00;-#####0.00; ;") & "Ԫ"
            Else
                strSumText = "�ϼƣ� " & dbl����ҩƷ���� & "��ҩƷ���� " & "Ӧ������" & Format(Dbl���, "#####0.00;-#####0.00; ;") & "Ԫ ʵ������" & Format(Dbl��� - dbl������, "#####0.00;-#####0.00; ;") & "Ԫ"
            End If
        End If
        
        vsfObj.Cell(flexcpText, intRow, 1, intRow, vsfObj.Cols - 1) = strSumText
        vsfObj.Cell(flexcpAlignment, intRow, 1, intRow, 1) = flexAlignLeftCenter
        vsfObj.Cell(flexcpForeColor, intRow, 1, intRow, vsfObj.Cols - 1) = mListColor.SumTotal
    Else
        If intSubType = mSubTotalType.SubByDept Then
            strSumText = "[��ҩ���ţ�" & strSub & "]С�ƣ� "
        ElseIf intSubType = mSubTotalType.SubByNo Then
            strSumText = "[NO��" & strSub & "]С�ƣ� "
        ElseIf intSubType = mSubTotalType.SubByDrug Then
            strSumText = "[ҩƷ��" & strSub & "]С�ƣ� "
        ElseIf intSubType = mSubTotalType.SubByPeople Then
            strSumText = "[������" & strSub & "]С�ƣ� "
        ElseIf intSubType = mSubTotalType.SubByHosNumber Then
            strSumText = "[סԺ�ţ�" & strSub & "]С�ƣ� "
        ElseIf intSubType = mSubTotalType.SubByBedNumber Then
            strSumText = "[��λ�ţ�" & strSub & "]С�ƣ� "
        End If
        
        If vsfObj.index = mListType.��ҩ Then
            If dbl�������� > 0 Then
                strSumText = strSumText & dbl�������� & "������  "
            End If
            If dblNO���� > 0 Then
                strSumText = strSumText & dblNO���� & "�ŵ���  "
            End If
        End If
        
        If intSubType = mSubTotalType.SubByDrug Then
            If vsfObj.index = mListType.���� Then
                strSumText = strSumText & "Ӧ������" & Format(Dbl���, "#####0.00;-#####0.00; ;") & "Ԫ ʵ������" & Format(Dbl��� - dbl������, "#####0.00;-#####0.00; ;") & "Ԫ"
            Else
                strSumText = strSumText & "����" & Format(Dbl���, "#####0.00;-#####0.00; ;") & "Ԫ"
            End If
        ElseIf dbl����ҩƷ���� > 0 Then
            If vsfObj.index = mListType.���� Then
                strSumText = strSumText & dbl����ҩƷ���� & "��ҩƷ���� " & "Ӧ������" & Format(Dbl���, "#####0.00;-#####0.00; ;") & "Ԫ ʵ������" & Format(Dbl��� - dbl������, "#####0.00;-#####0.00; ;") & "Ԫ"
            Else
                strSumText = strSumText & dbl����ҩƷ���� & "��ҩƷ���� " & "����" & Format(Dbl���, "#####0.00;-#####0.00; ;") & "Ԫ"
            End If
        End If
        
        If vsfObj.index = mListType.��ҩ Then
            intSumCol = mIntCol��ҩ_����� + 1
        ElseIf vsfObj.index = mListType.���� Then
            If mcondition.bln�����һ��� = True Then
                intSumCol = mIntCol���һ���_��ҩ����
            Else
                intSumCol = mIntCol����_Ʒ��
            End If
        End If
        
        vsfObj.Cell(flexcpText, intRow, intSumCol, intRow, vsfObj.Cols - 1) = strSumText
        vsfObj.Cell(flexcpAlignment, intRow, intSumCol, intRow, intSumCol) = flexAlignLeftCenter
        vsfObj.Cell(flexcpForeColor, intRow, intSumCol, intRow, vsfObj.Cols - 1) = mListColor.SumTotal
    End If
End Sub

Public Sub ShowList(ByVal intType As Integer, ByVal lng�ⷿID As Long)
    Dim objVSF As VSFlexGrid
    
    With mcondition
        If lng�ⷿID > 0 And .lngҩ��id <> lng�ⷿID Then
            .lngҩ��id = lng�ⷿID
            .bln�������� = CheckIsCenter(.lngҩ��id)
            Call GetDrugDigit(.lngҩ��id, "ҩƷ���ŷ�ҩ", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        End If
         
        If .intListType <> intType Then
            Call SetTempOperate(.intListType, intType)
            .intListType = intType
        End If
        Call Load��ҩ��(.lngҩ��id)
        Call Load�˲���(.lngҩ��id)
    End With
    
    For Each objVSF In vsfList
        If objVSF.index = mcondition.intListType Then
            objVSF.Visible = True
        Else
            objVSF.Visible = False
        End If
    Next
    
    mblnShowReject = (intType = mListType.�ܷ�)
    
    If mcondition.intListType = mListType.��ҩ Or mcondition.intListType = mListType.���� Then
        picAssist.Visible = True
    Else
        picAssist.Visible = False
    End If
            
    Call SetComandBars(mcondition.intListType)
    
    Call InitColSelList(mcondition.intListType)
    
    vsfChargeOff.Visible = False
    
    If mcondition.intListType = mListType.��ҩ Then
        picInfo.Visible = mcondition.bln��ʾ��չ��Ϣ
        fraH.Visible = True
        
    Else
        picInfo.Visible = False
        fraH.Visible = False
    End If

    Call Form_Resize
    '���ݷ�ҩ���ݼ��ı仯�ж��Ƿ�����ˢ���б�
    If (mcondition.intListType = mListType.���� Or mcondition.intListType = mListType.�ܷ�) And mblnSendChange = True Then
        Call RefreshList_Sum
        Call RefreshList_Reject
        mblnSendChange = False
    End If
End Sub

Private Sub cbo��ҩ��_Click()
'    Exit Sub
End Sub

Private Sub cbo��ҩ��_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo��ҩ��.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo��ҩ��_KeyPress(KeyAscii As Integer)
Dim i As Long, intIdx As Integer
    Dim strText As String, strResult As String, strFilter As String

    If KeyAscii = 13 Then
        strText = UCase(cbo��ҩ��.Text)
        If cbo��ҩ��.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo��ҩ��.List(cbo��ҩ��.ListIndex) Then Call zlControl.CboSetIndex(cbo��ҩ��.hWnd, -1)
        End If
        If strText = "" Then
            cbo��ҩ��.ListIndex = -1
        ElseIf cbo��ҩ��.ListIndex = -1 Then
            intIdx = -1

            For i = 1 To cbo��ҩ��.ListCount - 1
                If Mid(cbo��ҩ��.List(i), 1, InStr(1, cbo��ҩ��.List(i), "-") - 1) = strText _
                    Or Mid(cbo��ҩ��.List(i), InStr(1, cbo��ҩ��.List(i), "-")) = strText Then
                    intIdx = i
                    Exit For
                End If
            Next

            If intIdx = -1 Then
                For i = 1 To cbo��ҩ��.ListCount - 1
                    If UCase(cbo��ҩ��.List(i)) Like strText & "*" Then
                        intIdx = i
                    End If
                Next
            End If

            cbo��ҩ��.ListIndex = intIdx
            SendMessage cbo��ҩ��.hWnd, CB_SHOWDROPDOWN, True, 0
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo��ҩ��_Click
            Exit Sub
        End If
        If cbo��ҩ��.ListIndex = -1 Then
            cbo��ҩ��.ListIndex = 0
        Else
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo��ҩ��_Click
            ElseIf intIdx <> cbo��ҩ��.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo��ҩ��.SetFocus
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo��ҩ��_Click
            End If
        End If
    End If
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim Int���� As Integer
    Dim strNo As String
    Dim objPopup As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim strsql As String
    Dim rsTmp As Recordset
    Dim lngPatiID As Long
    Dim str�Һŵ� As String
    Dim lng��ҳID As Long
    Dim lngҽ��id As Long
    
    Select Case Control.Id
        Case conMenu_Tool_ShowShortage
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mcondition.bln��ʾȱҩ = Control.Checked
            
            Call RefreshList_Send
        Case conMenu_Tool_ShowReturnSend
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mblnSendChange = True
            
            mcondition.bln��ʾ��ҩ�������� = Control.Checked
            
            Modify��ҩ���� mcondition.bln��ʾ��ҩ��������
            
            DoEvents
            Call RefreshList_Send
        Case conMenu_Tool_ShowInfo
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mcondition.bln��ʾ��չ��Ϣ = Control.Checked
            picInfo.Visible = Control.Checked
            fraH.Visible = Control.Checked
            
            If picInfo.Visible Then
                vsfList(mListType.��ҩ).Width = vsfList(mListType.��ҩ).Width - picInfo.Width
            Else
                vsfList(mListType.��ҩ).Width = vsfList(mListType.��ҩ).Width + picInfo.Width
            End If
            
            Call cbsMain_Resize
        Case conMenu_Tool_SumByBatch
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mcondition.bln�����λ��� = Control.Checked
            
            Call RefreshList_Sum
        Case conMenu_Tool_ShowAllProcess
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mcondition.bln��ʾ���̵��� = Control.Checked
            
            Call RefreshList_Return
        
        Case conMenu_Tool_ShowPlug
            '���ܣ��Բ��˹���ʷ/����״̬���й���
            'Pass
            If vsfList(mcondition.intListType).Row = 0 Then Exit Sub
            If vsfList(mcondition.intListType).IsSubtotal(vsfList(mcondition.intListType).Row) = True Then Exit Sub
            If mcondition.intListType = mListType.��ҩ Then
                Int���� = vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_����)
                strNo = vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_NO)
            ElseIf mcondition.intListType = mListType.��ҩ Then
                Int���� = vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_����)
                strNo = vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_NO)
            End If
            
'            Call AdviceCheckWarn(Int����, strNo, 21)
          
            '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�������˳�
            strsql = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.����id,0 ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,������ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] "
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, strNo, Int����)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!����ID
            str�Һŵ� = NVL(rsTmp!�Һŵ�)
            lng��ҳID = rsTmp!��ҳID
            
            Call gobjPass.zlPassCmdAlleyManage_YF(mlngMode, lngPatiID, lng��ҳID, str�Һŵ�)
        
        '�����˵���PASS����
        Case mconMenu_PASS * 10# To mconMenu_PASS * 10# + 99
            If vsfList(mcondition.intListType).Row = 0 Then Exit Sub
            If vsfList(mcondition.intListType).IsSubtotal(vsfList(mcondition.intListType).Row) = True Then Exit Sub
            If mcondition.intListType = mListType.��ҩ Then
                Int���� = vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_����)
                strNo = vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_NO)
                lngҽ��id = Val(vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_ҽ��id))
            ElseIf mcondition.intListType = mListType.��ҩ Then
                Int���� = vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_����)
                strNo = vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_NO)
                lngҽ��id = Val(vsfList(mListType.��ҩ).TextMatrix(vsfList(mListType.��ҩ).Row, mIntCol��ҩ_ҽ��id))
            End If
            
            strsql = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.����id,0 ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,������ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] "
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, strNo, Int����)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!����ID
            str�Һŵ� = NVL(rsTmp!�Һŵ�)
            lng��ҳID = rsTmp!��ҳID
            
            Call gobjPass.zlPassCommandBarExe_YF(mlngMode, Control.Id - (mconMenu_PASS * 10#), lngPatiID, lng��ҳID, str�Һŵ�, lngҽ��id)
        
        '�����˵�����ҩ����״̬����
        Case conMenu_Status_Verify
            SetSendBillState mChangeState.��ҩ, vsfList(mListType.��ҩ).Row
        Case conMenu_Status_Reject
            SetSendBillState mChangeState.�ܷ�, vsfList(mListType.��ҩ).Row
        Case conMenu_Status_Shortage
        Case conMenu_Status_NoProcess
            SetSendBillState mChangeState.������, vsfList(mListType.��ҩ).Row
        Case conMenu_Status_AllSend
            SetSendBillState mChangeState.ȫ����ҩ
        Case conMenu_Status_AllReject
            SetSendBillState mChangeState.ȫ���ܷ�
        Case conMenu_Status_AllNoProcess
            SetSendBillState mChangeState.ȫ��������
        
        '�����˵�����ҩ����״̬����
        Case conMenu_Status_AllReturn
            'ȫ����ҩ
            SetAllReturn
        Case conMenu_Status_AllCancel
            'ȫ��ȡ����ҩ
            SetAllNotReturn
    End Select
End Sub


Private Function AdviceCheckWarn(ByVal Int���� As Integer, ByVal strNo As String, ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'���ܣ�����Passϵͳ��ع���
'������lngCmd=
'        0-�������PASS�˵�״̬
'        21-����״̬/����ʷ����(ֻ��)
'      lngRow=��ǰҩƷҽ�����кţ�lngCmd=0ʱ��Ҫ
'���أ����PASS�˵�ʱ������>=0��ʾ���Ե����˵�,��������-1
'˵������ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New ADODB.Recordset
    Dim strҩƷ As String, str�÷� As String, lngҩƷid As Long, str������λ As String
    Dim strsql As String, i As Long, k As Long
    Dim lngPatiID As Long
    Dim lngPassPati As Long
    Dim lng��ҳID As Long
    Dim str�Һŵ� As String

    AdviceCheckWarn = -1

    On Error GoTo errH
    Screen.MousePointer = 11

    If strNo = "" Then Exit Function

    '����PASS����״̬
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�������˳�
    strsql = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] "
    Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, strNo, Int����)

    If rsTmp.RecordCount = 0 Then
        rsTmp.Close
        Exit Function
    End If

    lngPatiID = rsTmp!����ID
    str�Һŵ� = zlStr.NVL(rsTmp!�Һŵ�)
    lng��ҳID = rsTmp!��ҳID

    '���벡�˾�����Ϣ(PASS��Ҫ�Ļ�������,ͬһ���˿ɲ��ظ�����)
    '-------------------------------------------------------------
    If lngPatiID <> lngPassPati Then
        If str�Һŵ� <> "" Then               '���ﲡ��
            strsql = "Select ����ID,Count(Distinct Trunc(�Ǽ�ʱ��)) as ������� From ���˹Һż�¼ Where ��¼����=1 And ��¼״̬=1 And ����ID=[1] Group by ����ID"
            strsql = "Select D.�������,A.����,A.�Ա�,A.��������," & _
                " C.���� as ������,C.���� as ������,E.��� as ҽ����,E.���� as ҽ����" & _
                " From ������Ϣ A,���˹Һż�¼ B,���ű� C,(" & strsql & ") D,��Ա�� E" & _
                " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And A.����ID=D.����ID" & _
                " And B.ִ����=E.����(+) And A.����ID=[1] And B.NO=[2]"
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, lngPatiID, str�Һŵ�)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

            Call PassSetPatientInfo(lngPatiID, rsTmp!�������, rsTmp!����, zlStr.NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), zlStr.NVL(rsTmp!ҽ����) & "/" & zlStr.NVL(rsTmp!ҽ����), ""), "")
        Else                                    'סԺ����
            strsql = _
                " Select A.����,A.�Ա�,A.��������,B.��Ժ����,B.��Ժ����," & _
                " C.���� as ������,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
                " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
                " Where A.����ID=B.����ID And A.��ҳid=B.��ҳid And B.��Ժ����ID=C.ID" & _
                " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[2]"
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, lngPatiID, lng��ҳID)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

            Call PassSetPatientInfo(lngPatiID, lng��ҳID, rsTmp!����, zlStr.NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), zlStr.NVL(rsTmp!ҽ����) & "/" & zlStr.NVL(rsTmp!ҽ����), ""), _
                IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy-MM-dd")))
        End If
        lngPassPati = lngPatiID
    End If

    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        If mcondition.intListType = mListType.��ҩ Then
           'ȡҩƷ����
            strҩƷ = vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_ҩƷ����)
            lngҩƷid = vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_ҩƷID)
            str������λ = vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_������λ)
            'ȡҩƷ��ҩ;��
            str�÷� = vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_�÷�)
        Else
            'ȡҩƷ����
            strҩƷ = vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_ҩƷ����)
            lngҩƷid = vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_ҩƷID)
            str������λ = vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_������λ)
            'ȡҩƷ��ҩ;��
            str�÷� = vsfList(mListType.��ҩ).TextMatrix(lngRow, mIntCol��ҩ_�÷�)
        End If
        
        '�����ѯҩƷ��Ϣ
        Call PassSetQueryDrug(lngҩƷid, strҩƷ, str������λ, str�÷�)

        '���ò˵�����״̬
        Call SetPassMenuState

        AdviceCheckWarn = 1 '��ʾ���Ե����˵�

        Screen.MousePointer = 0: Exit Function
    End If

    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    Screen.MousePointer = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetPassMenuState()
    '���ܣ�����Pass�˵�����״̬
    'Pass
    Dim objPopup As CommandBarControl

    ''''һ���˵�
    'ҩ���ٴ���Ϣ�ο�
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CPRRes") = 1

    'ҩƷ˵����
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 1, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("Directions") = 1

    '�й�ҩ��
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 2, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("Chp") = 1

    '������ҩ����
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 3, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CPERes") = 1

    '����ֵ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 4, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CheckRes") = 1

    'ר����Ϣ
'    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 5, , True)
'    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("") = 1

    'ҽҩ��Ϣ����
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 6, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MEDInfo") = 1

    'ҩƷ�����Ϣ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 7, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MATCH-DRUG") = 1

    '��ҩ;�������Ϣ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 8, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MATCH-ROUTE") = 1

    'ҽԺҩƷ��Ϣ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 9, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("HisDrugInfo") = 1
    
    
    ''''ר����Ϣ�����˵�
    'ҩ��-ҩ���໥����
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DDIM") = 1
    
    'ҩ��-ʳ���໥ʹ��
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 1, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DFIM") = 1
    
    '����ע�����������
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 2, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MatchRes") = 1
    
    '����ע�����������
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 3, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("TriessRes") = 1
    
    '����֢
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 4, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DDCM") = 1
    
    '������
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 5, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("SIDE") = 1
    
    '��������ҩ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 6, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("GERI") = 1
    
    '��ͯ��ҩ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 7, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("PEDI") = 1
    
    '��������ҩ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 8, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("PREG") = 1
    
    '��������ҩ
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 9, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("LACT") = 1
End Sub
Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim objVSF As VSFlexGrid
    
    On Error Resume Next
    
    If cbsMain.count > 1 Then
        Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
        For Each objVSF In vsfList
            If objVSF.Visible Then
                objVSF.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop - IIf(picAssist.Visible, picAssist.Height + 50, 0)
                
                fraColSel.Left = objVSF.Left + objVSF.ColWidth(0) - fraColSel.Width - 50
                fraColSel.Top = objVSF.Top + (objVSF.RowHeight(0) - fraColSel.Height) / 2 + 30
                fraColSel.ZOrder
                
                If objVSF.index = mListType.��ҩ Then
                    mdblSendListHeight = objVSF.Height
                    If Me.picInfo.Visible Then
                        objVSF.Width = objVSF.Width - picInfo.Width
                    End If
                    
                    With Me.picInfo
                        .Left = objVSF.Left + objVSF.Width + 50
                        .Height = objVSF.Height
                        .Top = objVSF.Top
                    End With
                    
                    fraH.Left = picInfo.Left - 20
                    fraH.Height = objVSF.Height
                    fraH.Top = objVSF.Top
                    picInfo.Top = objVSF.Top
                End If
                
                If objVSF.index = mListType.���� Then
                    mdblSumListHeight = objVSF.Height
                    Call ResizeChargeOffList
                End If
                Exit For
            End If
        Next
    End If
End Sub

Private Sub InitComandBars()
    Dim cbrControl As CommandBarControl
    Dim objCmdBar As CommandBar
    Dim lngCount As Integer
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
'        .SetIconSize False, 24, 24
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.ActiveMenuBar.Visible = False
    Me.cbsMain.AddImageList Me.imgCheck
End Sub

Private Sub SetComandBars(ByVal intListType As Integer)
    Dim cbrControl As CommandBarControl
    Dim cbrControlSub As CommandBarControl
    Dim objCmdBar As CommandBar
    Dim lngCount As Integer
    Dim objMenu As CommandBarPopup
        
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    If intListType = mListType.�ܷ� Or intListType = mListType.ȱҩ Then Exit Sub
    
    Set objCmdBar = cbsMain.Add("����", xtpBarTop)
    objCmdBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objCmdBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objCmdBar.ContextMenuPresent = False
    
    'ҩƷ������ʾ
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_MediPopup, "ҩ����ʾ", 1, False)
    objMenu.Id = conMenu_MediPopup
    With objMenu.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Medi_CodeAddName, "ҩ��(���������)")
        cbrControl.Checked = (cbrControl.Id = 401 + mcondition.intҩƷ���Ʊ�����ʾ)
        Set cbrControl = .Add(xtpControlButton, conMenu_Medi_Code, "ҩ��(������)")
        cbrControl.Checked = (cbrControl.Id = 401 + mcondition.intҩƷ���Ʊ�����ʾ)
        Set cbrControl = .Add(xtpControlButton, conMenu_Medi_Name, "ҩ��(������)")
        cbrControl.Checked = (cbrControl.Id = 401 + mcondition.intҩƷ���Ʊ�����ʾ)
    End With
    
    Select Case intListType
        Case mListType.��ҩ
            '���ù������˵�
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowShortage, "��ʾȱҩҩƷ")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "��ʾ���Ƿ���ʾȱҩҩƷ"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln��ʾȱҩ, 2, 1)
            cbrControl.Checked = mcondition.bln��ʾȱҩ
            
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowReturnSend, "��ʾ��ҩ��������")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "��ʾ���Ƿ���ʾ��ҩ��������"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln��ʾ��ҩ��������, 2, 1)
            cbrControl.Checked = mcondition.bln��ʾ��ҩ��������
            
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowInfo, "��ʾ��չ��Ϣ")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "��ʾ���Ƿ���ʾ�ٴ���ϣ�����ҩ�������Ϣ������ҽ��ǩ��ͼƬ"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln��ʾ��չ��Ϣ, 2, 1)
            cbrControl.Checked = mcondition.bln��ʾ��չ��Ϣ
            
'            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowPlug, "����ʷ/����״̬")
'            cbrControl.BeginGroup = True
'            cbrControl.ToolTipText = "��ʾ����ʾ����ʷ/����״̬"
'            cbrControl.Style = xtpButtonIconAndCaption
'            cbrControl.IconId = 3
'            cbrControl.Visible = (mcondition.intShowPass = 1 And IsInString(gstrprivs, "������ҩ���", ";"))
            
            If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, objCmdBar.Controls, conMenu_Tool_ShowPlug, 3)
            
            '���õ����˵�
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_StatusPopup, "�༭(&E)", 1, False)
            objMenu.Id = conMenu_StatusPopup
            With objMenu.CommandBar.Controls
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_Verify, "��ҩ(&C)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_Reject, "�ܷ�(&H)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_Shortage, "ȱҩ(&L)")
                cbrControl.Enabled = False
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_NoProcess, "������(&H)")
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllSend, "ȫ����ҩ(&S)")
                cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllReject, "ȫ���ܷ�(&J)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllNoProcess, "ȫ��������(&B)")
            End With
        Case mListType.����
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_SumByBatch, "��ҩƷ���λ���")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "��ʾ���Ƿ�ҩƷ���λ���"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln�����λ���, 2, 1)
            cbrControl.Checked = mcondition.bln�����λ���
        Case mListType.��ҩ
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowAllProcess, "��ʾ���й��̵���")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "��ʾ���Ƿ���ʾ���й��̵���"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln��ʾ���̵���, 2, 1)
            cbrControl.Checked = mcondition.bln��ʾ���̵���
            
'            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowPlug, "����ʷ/����״̬")
'            cbrControl.BeginGroup = True
'            cbrControl.ToolTipText = "��ʾ����ʾ����ʷ/����״̬"
'            cbrControl.Style = xtpButtonIconAndCaption
'            cbrControl.IconId = 3
'            cbrControl.Visible = (mcondition.intShowPass = 1 And IsInString(gstrprivs, "������ҩ���", ";"))
            If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, objCmdBar.Controls, conMenu_Tool_ShowPlug, 3)
            
            '���õ����˵�
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_StatusPopup, "�༭(&E)", 1, False)
            objMenu.Id = conMenu_StatusPopup
            With objMenu.CommandBar.Controls
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllReturn, "ȫ����ҩ(&R)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllCancel, "ȫ��ȡ��(&C)")
            End With
    End Select
    
    
    Select Case intListType
        Case mListType.��ҩ, mListType.��ҩ
            '���õ����˵���PASS
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_PASS, "PASS��&P)", 1, False)
            objMenu.Id = mconMenu_PASS
'            If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, cbsMain, mconMenu_PASS, 1)
    End Select
End Sub


Private Sub chk�������_Click()
    If mcondition.intListType <> mListType.��ҩ Then Exit Sub
    If vsfList(mcondition.intListType).Row = 0 Then Exit Sub
    If vsfList(mListType.��ҩ).IsSubtotal(vsfList(mListType.��ҩ).Row) = True Then Exit Sub
    
    vsf���.Tag = ""
    Call GetDiagnosis(vsfList(mcondition.intListType).Row)
End Sub

Private Sub Form_Load()
    With mcondition
        If Val(zlDataBase.GetPara("ʹ�ø��Ի����")) = 1 Then
            .bln��ʾȱҩ = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾȱҩҩƷ", "1")) = 1)
            .bln��ʾ��ҩ�������� = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ��ҩ��������", "1")) = 1)
            .bln��ʾ��չ��Ϣ = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ�����Ϣ", "1")) = 1)
            .bln�����λ��� = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\����", "�����λ���", "0")) = 1)
            .bln��ʾ���̵��� = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ���̵���", "0")) = 1)
            .bln��ʾ������� = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ�������", "0")) = 1)
        Else
            .bln��ʾȱҩ = True
            .bln�����λ��� = False
            .bln��ʾ��չ��Ϣ = True
            .bln��ʾ���̵��� = False
            .bln��ʾ��ҩ�������� = True
            .bln��ʾ������� = False
        End If
        
        .intShowPass = gintPass
        .blnҽ����ѯ = IsInString(gstrprivs, "ҽ����ѯ", ";")
        
'        .blnҽ����ѯ = False
        
        '���ڲ���PASS
'        .blnShowPass = True

        .bln�޸��������� = IsInString(gstrprivs, "�޸���������", ";")
    End With
    
    Call SetParams
    
    Call Load��ҩ����ʽ
    
    Call InitComandBars
    Call InitList(-1)
    
    vsfChargeOff.Visible = False
    Me.txt��ҩ����.Text = ""
    
    Me.chk�������.Value = IIf(mcondition.bln��ʾ������� = True, 1, 0)
End Sub


Private Sub Form_Resize()
    Dim objVSF As VSFlexGrid
    Dim i As Integer
    
    On Error Resume Next
    
    With picAssist
        If .Visible Then
            .Left = 0
            .Top = Me.Height - .Height
            .Width = Me.Width - 50
        End If
    End With
    
    If cbsMain.count = 1 Then
        For Each objVSF In vsfList
            If objVSF.Visible = True Then
                objVSF.Move 0, 0, Me.Width, IIf(picAssist.Visible, Me.Height - picAssist.Height - 50, Me.Height)
                
                fraColSel.Left = objVSF.Left + objVSF.ColWidth(0) - fraColSel.Width - 50
                fraColSel.Top = objVSF.Top + (objVSF.RowHeight(0) - fraColSel.Height) / 2 + 30
                fraColSel.ZOrder
                
                If objVSF.index = mListType.��ҩ Then
                    mdblSendListHeight = objVSF.Height
                    objVSF.Width = objVSF.Width - picInfo.Width
                    
                    With Me.picInfo
                        .Left = objVSF.Left + objVSF.Width + 50
                        .Height = objVSF.Height
                        .Top = objVSF.Top
                    End With
                    fraH.Left = picInfo.Left - 20
                    fraH.Height = objVSF.Height
                    fraH.Top = objVSF.Top
                    picInfo.Top = objVSF.Top
                End If
                
                If objVSF.index = mListType.���� Then
                    mdblSumListHeight = objVSF.Height
                    Call ResizeChargeOffList
                End If
                Exit For
            End If
        Next
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetTempOperate mcondition.intListType, mcondition.intListType
'    SaveListColState
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ���ŷ�ҩ����", "ҩƷ������ʾ��ʽ", mcondition.intҩƷ���Ʊ�����ʾ)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ���ŷ�ҩ����", "��ҩ����ʽ", cbo��ҩ����ʽ.ListIndex)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "���ŷ�ҩ����\��ҩ", "��ʾ�������", chk�������.Value)
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList(Val(.Tag)).SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList(Val(.Tag)).ColHidden(.RowData(i)) Or vsfList(Val(.Tag)).ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = fraColSel.Top + fraColSel.Height
                
                If .Top + .Height > Me.ScaleHeight - vsfList(Val(.Tag)).Top Then
                    .Height = Me.ScaleHeight - .Top - vsfList(Val(.Tag)).Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                
                .Left = fraColSel.Left
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub picInfo_Resize()
    picHscSend.Width = picInfo.Width
    Pic��ҩ����.Width = picInfo.Width
    picDoctor.Width = picInfo.Width
    Pic��ҩ����.Top = (1 / 3) * Me.picInfo.Height
    picDoctor.Top = (2 / 3) * Me.picInfo.Height
    
    vsf���.Height = Pic��ҩ����.Top - picHscSend.Top - picHscSend.Height
    vsf���.Top = Me.picHscSend.Height
    vsf���.Width = picInfo.Width
    
    txt��ҩ����.Top = Pic��ҩ����.Top + Pic��ҩ����.Height
    txt��ҩ����.Height = Me.picDoctor.Top - Pic��ҩ����.Top - Pic��ҩ����.Height
    txt��ҩ����.Width = picInfo.Width
    
    picǩ��ͼƬ.Top = picDoctor.Height + picDoctor.Top + 20
    picǩ��ͼƬ.Left = (picInfo.Width / 2) - (picǩ��ͼƬ.Width / 2)
End Sub

Private Sub fraH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfList(mListType.��ҩ).Width + x <= 1200 Then Exit Sub
        If picInfo.Width - x < 1200 Then Exit Sub

        fraH.Left = fraH.Left + x
        picInfo.Left = picInfo.Left + x
        picInfo.Width = picInfo.Width - x
        vsfList(mListType.��ҩ).Width = vsfList(mListType.��ҩ).Width + x
        
        Me.Refresh
    End If
End Sub

Private Sub picHscSend_Resize()
    On Error Resume Next
    
    With lblDiag
        .Left = (picHscSend.Width - .Width) / 2
    End With
    
    With chk�������
        .Left = picHscSend.Width - .Width - 50
    End With
End Sub


Private Sub Pic��ҩ����_Resize()
    With lbl��ҩ����
        .Left = (Pic��ҩ����.Width - .Width) / 2
    End With
End Sub

Private Sub picDoctor_Resize()
    With lblDoctor
        .Left = (picDoctor.Width - .Width) / 2
    End With
End Sub

Private Sub picAssist_Resize()
    On Error Resume Next
    
    With Lbl��ҩ��
    End With
    
    With cbo��ҩ��
    End With
    
    
    
    With cbo�˲���
        .Left = (picAssist.Width - .Width + 400) / 2
    End With
    
    With lbl�˲���
        .Left = cbo�˲���.Left - 50 - .Width
    End With
    
    With cbo��ҩ����ʽ
        .Left = picAssist.Width - .Width - 50
    End With
    
    With lbl��ҩ����ʽ
        .Left = cbo��ҩ����ʽ.Left - 50 - .Width
    End With
End Sub

Private Sub vsfChargeOff_EnterCell()
    With vsfChargeOff
        If .Row = 0 Then Exit Sub
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = frmPublic.ImgList.ListImages(2).Picture

    End With
End Sub


Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsfColSel.RowData(Row)
        If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
            vsfList(Val(vsfColSel.Tag)).ColWidth(lngCol) = vsfList(Val(vsfColSel.Tag)).ColData(lngCol)
            vsfList(Val(vsfColSel.Tag)).ColHidden(lngCol) = False
        Else
'            vsfList(Val(vsfColSel.Tag)).ColWidth(lngCol) = 0
            vsfList(Val(vsfColSel.Tag)).ColHidden(lngCol) = True
        End If
    End If
    
    SaveListColState Val(vsfColSel.Tag)
End Sub

Private Sub vsfColSel_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfColSel
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub


Private Sub vsfColSel_LostFocus()
    vsfColSel.Visible = False
End Sub


Private Sub vsfColSel_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsfColSel.Cell(flexcpForeColor, Row, 1) = vsfColSel.BackColorFixed Then Cancel = True
End Sub

Private Sub vsfList_AfterEdit(index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    Dim j As Integer
    Dim intRow As Integer
    Dim dblMoney As Double
    Dim dblCurMoney As Double
    Dim strCont As String
    
    With vsfList(index)
        Select Case index
            Case mListType.��ҩ
                If Row = 0 Then Exit Sub
                If Val(.TextMatrix(Row, mIntCol��ҩ_�շ�ID)) = 0 Then Exit Sub
                If Val(.TextMatrix(Row, mIntCol��ҩ_ִ��״̬)) <> mState.��ҩ And Val(.TextMatrix(Row, mIntCol��ҩ_ִ��״̬)) <> mState.��ҩ_ԭʼ��¼ Then Exit Sub
                If Val(.TextMatrix(Row, mIntCol��ҩ_׼����)) = 0 Then Exit Sub
                If Col = mIntCol��ҩ_��ҩ�� Then
                    If Val(.TextMatrix(Row, mIntCol��ҩ_׼����)) >= 0 Then
                        If Val(.TextMatrix(Row, mIntCol��ҩ_��ҩ��)) > Val(.TextMatrix(Row, mIntCol��ҩ_׼����)) Or Val(.TextMatrix(Row, mIntCol��ҩ_��ҩ��)) < 0 Then
                            .TextMatrix(Row, mIntCol��ҩ_��ҩ��) = Val(.TextMatrix(Row, mIntCol��ҩ_׼����))
                        End If
                    Else
                        If Val(.TextMatrix(Row, mIntCol��ҩ_��ҩ��)) < Val(.TextMatrix(Row, mIntCol��ҩ_׼����)) Or Val(.TextMatrix(Row, mIntCol��ҩ_��ҩ��)) >= 0 Then
                            .TextMatrix(Row, mIntCol��ҩ_��ҩ��) = Val(.TextMatrix(Row, mIntCol��ҩ_׼����))
                        End If
                    End If
                    
                    If Val(.TextMatrix(Row, mIntCol��ҩ_��ҩ��)) = 0 Then
                        .TextMatrix(Row, mIntCol��ҩ_��ҩ��) = ""
                        
                        If Val(.TextMatrix(Row, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ_ԭʼ��¼ Then
                            Exit Sub
                        ElseIf Val(.TextMatrix(Row, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ Then
                            mrsReturnList.Filter = "�շ�ID=" & Val(.TextMatrix(Row, mIntCol��ҩ_�շ�ID))
                            
                            .TextMatrix(Row, mIntCol��ҩ_״̬) = "������"
                            .TextMatrix(Row, mIntCol��ҩ_ִ��״̬) = mState.��ҩ_ԭʼ��¼
                            mrsReturnList!״̬ = .TextMatrix(Row, mIntCol��ҩ_״̬)
                            mrsReturnList!ִ��״̬ = Val(.TextMatrix(Row, mIntCol��ҩ_ִ��״̬))
                            mrsReturnList!��ҩ�� = Val(.TextMatrix(Row, mIntCol��ҩ_��ҩ��))
                            mrsReturnList.Update
                        End If
                    Else
                        mrsReturnList.Filter = "�շ�ID=" & Val(.TextMatrix(Row, mIntCol��ҩ_�շ�ID))
                        
                        .TextMatrix(Row, mIntCol��ҩ_״̬) = "��ҩ"
                        .TextMatrix(Row, mIntCol��ҩ_ִ��״̬) = mState.��ҩ
                        mrsReturnList!״̬ = .TextMatrix(Row, mIntCol��ҩ_״̬)
                        mrsReturnList!ִ��״̬ = Val(.TextMatrix(Row, mIntCol��ҩ_ִ��״̬))
                        mrsReturnList!��ҩ�� = Val(.TextMatrix(Row, mIntCol��ҩ_��ҩ��))
                        mrsReturnList.Update
                    End If
                End If
            Case mListType.����
                If Row = 0 Then Exit Sub
                If mcondition.bln�����һ��� = False Or mcondition.bln�޸��������� = False Then Exit Sub
                If Val(.TextMatrix(Row, mIntCol���һ���_��ҩ����id)) = 0 Then Exit Sub
                
                Dim dbl������ As Double
                Dim dblӦ���� As Double
                Dim dblʵ���� As Double
                
                dblӦ���� = Val(.TextMatrix(Row, mIntCol���һ���_Ӧ������)) - Val(.TextMatrix(Row, mIntCol���һ���_��������))
                
                If Col = mIntCol���һ���_ʵ������ Then
                    dblʵ���� = Val(.TextMatrix(Row, mIntCol���һ���_��������))
                    If dblʵ���� > dblӦ���� Or dblʵ���� < 0 Then
                        .TextMatrix(Row, mIntCol���һ���_ʵ������) = zlStr.FormatEx(dblӦ����, 5)
                        .TextMatrix(Row, mIntCol���һ���_��������) = 0
                    Else
                        .TextMatrix(Row, mIntCol���һ���_ʵ������) = zlStr.FormatEx(dblʵ����, 5)
                        .TextMatrix(Row, mIntCol���һ���_��������) = zlStr.FormatEx(dblӦ���� - Val(.TextMatrix(Row, mIntCol���һ���_ʵ������)), 5)
                    End If
                ElseIf Col = mIntCol���һ���_�������� Then
                    dbl������ = Val(.TextMatrix(Row, mIntCol���һ���_��������))
                    If dbl������ > dblӦ���� Or dbl������ < 0 Then
                        .TextMatrix(Row, mIntCol���һ���_ʵ������) = zlStr.FormatEx(dblӦ����, 5)
                        .TextMatrix(Row, mIntCol���һ���_��������) = 0
                    Else
                        .TextMatrix(Row, mIntCol���һ���_ʵ������) = zlStr.FormatEx(dblӦ���� - Val(.TextMatrix(Row, mIntCol���һ���_��������)), 5)
                        .TextMatrix(Row, mIntCol���һ���_��������) = zlStr.FormatEx(dbl������, 5)
                    End If
                End If
                
                .TextMatrix(Row, mIntCol���һ���_ʵ�����) = zlStr.FormatEx(Val(.TextMatrix(Row, mIntCol���һ���_Ӧ�����)) / Val(.TextMatrix(Row, mIntCol���һ���_Ӧ������)) * Val(.TextMatrix(Row, mIntCol���һ���_ʵ������)), mintMoneyDigit, , True)
                        
                If mcondition.bln�����һ��� = True Then
                    mrsSendList.Filter = "��ҩ����id=" & Val(.TextMatrix(Row, mIntCol���һ���_��ҩ����id)) & " and ҩƷid=" & Val(.TextMatrix(Row, mIntCol���һ���_ҩƷID))
                Else
                    mrsSendList.Filter = "ҩƷid=" & Val(.TextMatrix(Row, mIntCol���һ���_ҩƷID)) & "And ����=" & Val(.TextMatrix(Row, mIntCol���һ���_����))
                End If
                
                mrsSendList!�������� = .TextMatrix(Row, mIntCol���һ���_��������)
                
                mrsSendList.Update
                        
                DoEvents
                
                .Row = Row
                .Col = mIntCol���һ���_ʵ������
                If Val(.TextMatrix(Row, mIntCol���һ���_ʵ������)) < 0 Then
                    .CellForeColor = vbRed
                ElseIf Val(.TextMatrix(Row, mIntCol���һ���_ʵ������)) > 0 Then
                    .CellForeColor = vbBlue
                End If
                
                '��ȡ��һ���ϼ��У���ȡ��һ���ϼ��У�Ȼ������ͳ��ʵ�����
                For i = Row To .rows - 1
                    If Not IsNumeric(.TextMatrix(i, mIntCol���һ���_Ӧ�����)) Then
                        Exit For
                    End If
                Next
                
                
                For j = Row To 1 Step -1
                    If Not IsNumeric(.TextMatrix(j, mIntCol���һ���_ʵ�����)) Then
                        Exit For
                    End If
                Next
                
                For intRow = j + 1 To i - 1
                    dblMoney = dblMoney + CDbl(.TextMatrix(intRow, mIntCol���һ���_ʵ�����))
                Next
                
                '��ȡԭʼ��ʵ�ʽ��ͳ��ֵ
                strCont = .TextMatrix(i, 1)
                dblCurMoney = Mid(Mid(strCont, InStr(1, strCont, "ʵ������") + 6), 1, InStr(1, Mid(strCont, InStr(1, strCont, "ʵ������") + 6), "Ԫ") - 1)
                
                '�޸ĵ�ǰ��ͳ�ƽ��
                .Cell(flexcpText, i, 1, i, .Cols - 1) = Mid(strCont, 1, InStr(1, strCont, "ʵ������") + 5) & Format(dblMoney, "#####0.00;-#####0.00;0.00;") & "Ԫ"
                
                '��ȡ���α༭���
                dblMoney = dblCurMoney - dblMoney
                
                If i <> .rows - 1 Then
                  '��ȡԭʼ�ϼƽ��
                  strCont = .TextMatrix(.rows - 1, 1)
                  dblCurMoney = Mid(Mid(strCont, InStr(1, strCont, "ʵ������") + 6), 1, InStr(1, Mid(strCont, InStr(1, strCont, "ʵ������") + 6), "Ԫ") - 1)
                
                  
                  '�޸ĺϼ�ͳ�ƽ��
                  .Cell(flexcpText, .rows - 1, 1, .rows - 1, .Cols - 1) = Mid(strCont, 1, InStr(1, strCont, "ʵ������") + 5) & Format(dblCurMoney - dblMoney, "#####0.00;-#####0.00;0.00;") & "Ԫ"
                End If
        End Select
    End With
    
End Sub

Private Sub vsfList_AfterMoveColumn(index As Integer, ByVal Col As Long, Position As Long)
    Dim i As Integer
    
    '������ѡ���б�
    Call InitColSelList(index)
    
    '������˳���
    For i = 0 To vsfList(index).Cols - 1
        Call SetColumnValue(index, vsfList(index).TextMatrix(0, i), i)
    Next
    
    '����ҳ���������
    SaveListColState index
End Sub

Private Sub vsfList_AfterSort(index As Integer, ByVal Col As Long, Order As Integer)
    If vsfList(index).rows > 1 Then
        If vsfList(index).TextMatrix(1, vsfList(index).ColIndex("ҩƷ����")) <> "" Then
            If index = mListType.��ҩ Then
                SetSubTotal vsfList(index), vsfList(index).TextMatrix(0, Col)
                SetGroup vsfList(index), Col = mIntCol��ҩ_NO
            ElseIf index = mListType.���� Then
                SetSubTotal vsfList(index), vsfList(index).TextMatrix(0, Col)
            ElseIf index = mListType.��ҩ Then
                SetGroup vsfList(index), Col = mIntCol��ҩ_NO And mcondition.bln��ʾ���̵��� = False
            End If
        End If
        
        If Val(zlDataBase.GetPara("ʹ�ø��Ի����")) = 1 Then
            '���洦���嵥���û��������
            '�������
            '����б�����
            'ֵ=�к�|��/����
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ���ŷ�ҩ", "�����嵥����" & index, Col & "|" & Order)
        End If
    End If
End Sub

Private Sub vsfList_AfterUserResize(index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Val(zlDataBase.GetPara("ʹ�ø��Ի����")) = 1 Then
        '����ҳ���������
        SaveListColState index
    End If
    
    If index = mListType.��ҩ Then
        If Col = mIntCol��ҩ_Ƥ�Խ�� Then
            With vsfList(index)
                If .ColWidth(mIntCol��ҩ_Ƥ�Խ��) > 800 Then
                    .ColWidth(mIntCol��ҩ_Ʒ��) = .ColWidth(mIntCol��ҩ_Ʒ��) + (.ColWidth(mIntCol��ҩ_Ƥ�Խ��) - 800)
                    .ColWidth(mIntCol��ҩ_Ƥ�Խ��) = 800
                Else
                    .ColWidth(mIntCol��ҩ_Ʒ��) = .ColWidth(mIntCol��ҩ_Ʒ��) - (800 - .ColWidth(mIntCol��ҩ_Ƥ�Խ��))
                    .ColWidth(mIntCol��ҩ_Ƥ�Խ��) = 800
                End If
            End With
        End If
    End If
End Sub


Private Sub vsfList_BeforeMoveColumn(index As Integer, ByVal Col As Long, Position As Long)
    '���ò����ƶ�����
    Select Case index
        Case mListType.��ҩ
            If Col = mIntCol��ҩ_����� Then
                Position = mIntCol��ҩ_�����
            End If
            
            If Col = mIntCol��ҩ_����� Then
                Position = mIntCol��ҩ_�����
            End If
            
            If Col = mIntCol��ҩ_Ʒ�� Then
                Position = mIntCol��ҩ_Ʒ��
            End If
            
            If Col = mIntCol��ҩ_Ƥ�Խ�� Then
                Position = mIntCol��ҩ_Ƥ�Խ��
            End If
            
            If (Col <> mIntCol��ҩ_Ʒ�� And Position = mIntCol��ҩ_Ʒ��) Or (Col <> mIntCol��ҩ_Ƥ�Խ�� And Position = mIntCol��ҩ_Ƥ�Խ��) Or (Col <> mIntCol��ҩ_����� And Position = mIntCol��ҩ_�����) Or (Col <> mIntCol��ҩ_����� And Position = mIntCol��ҩ_�����) Then
                Position = Col
            End If
    End Select

End Sub

Private Sub vsfList_BeforeSort(index As Integer, ByVal Col As Long, Order As Integer)
    If index <> mListType.��ҩ And index <> mListType.���� Then Exit Sub
    
    With vsfList(index)
        .Subtotal flexSTClear
    End With
End Sub

Private Sub vsfList_BeforeUserResize(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '���ò��ܵ����п����
    Select Case index
        Case mListType.��ҩ
            If Col = mIntCol��ҩ_��ǰ�� Or Col = mIntCol��ҩ_����� Or Col = mIntCol��ҩ_����� _
                Or Col = mIntCol��ҩ_״̬ Then Cancel = True
        Case mListType.��ҩ
            If Col = mIntCol��ҩ_��ǰ�� Or Col = mIntCol��ҩ_����� Or Col = mIntCol��ҩ_����� Then Cancel = True
        Case Else
            If Col = 0 Then Cancel = True
    End Select
End Sub

Private Sub vsfList_DblClick(index As Integer)
    Dim lngColor As Long
    Dim i As Long
    Dim strNo As String
    Dim lng���ID As Long
    Dim intState As Integer
    Dim strState As String
    
    With vsfList(index)
        Select Case index
            Case mListType.��ҩ
                If .Row = 0 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol��ҩ_�շ�ID)) = 0 Then Exit Sub
                
                If .Col <> mIntCol��ҩ_����� Then
                    If Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬)) = mState.ȱҩ Then Exit Sub
                    
                    If Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ Then
                        .TextMatrix(.Row, mIntCol��ҩ_״̬) = "�ܷ�"
                        .TextMatrix(.Row, mIntCol��ҩ_ִ��״̬) = mState.�ܷ�
                        lngColor = mListColor.State_Reject
                    ElseIf Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬)) = mState.�ܷ� Then
                        .TextMatrix(.Row, mIntCol��ҩ_״̬) = "������"
                        .TextMatrix(.Row, mIntCol��ҩ_ִ��״̬) = mState.������
                        lngColor = mListColor.State_UnProcess
                    ElseIf Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬)) = mState.������ Then
                        If mcondition.bln����δ��˴�����ҩ = False And Val(.TextMatrix(.Row, mIntCol��ҩ_���շ�)) = 0 Then
                            .TextMatrix(.Row, mIntCol��ҩ_״̬) = "�ܷ�"
                            .TextMatrix(.Row, mIntCol��ҩ_ִ��״̬) = mState.�ܷ�
                            lngColor = mListColor.State_Reject
                        Else
                            .TextMatrix(.Row, mIntCol��ҩ_״̬) = "��ҩ"
                            .TextMatrix(.Row, mIntCol��ҩ_ִ��״̬) = mState.��ҩ
                            lngColor = mListColor.State_Send
                        End If
                    End If
                    
                    intState = Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬))
                    strState = .TextMatrix(.Row, mIntCol��ҩ_״̬)
                    
                    .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = lngColor
                    
                    mrsSendList.Filter = "�շ�ID=" & Val(.TextMatrix(.Row, mIntCol��ҩ_�շ�ID))
                    
                    mrsSendList!״̬ = .TextMatrix(.Row, mIntCol��ҩ_״̬)
                    mrsSendList!ִ��״̬ = Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬))
                    mrsSendList.Update
                    
                    mblnSendChange = True
                
                    'ͬ��ҽ����ҩƷ״̬��Ҫͬ���ı䣻����Ǹ�ΣҩƷ������Ҫ�󵥶�����ʱ��ͬ���ı�
                    If mcondition.bln�������� And InStr(1, mcondition.str��Σ����, .TextMatrix(.Row, mIntCol��ҩ_��ΣҩƷ)) = 0 Then
                        strNo = .TextMatrix(.Row, mIntCol��ҩ_NO)
                        lng���ID = Val(.TextMatrix(.Row, mIntCol��ҩ_���ID))
                        If lng���ID > 0 Then
                            For i = 1 To .rows - 1
                                If .IsSubtotal(i) = False And .TextMatrix(i, mIntCol��ҩ_NO) = strNo And Val(.TextMatrix(i, mIntCol��ҩ_���ID)) = lng���ID _
                                    And Val(.TextMatrix(i, mIntCol��ҩ_ִ��״̬)) <> intState And Val(.TextMatrix(i, mIntCol��ҩ_ִ��״̬)) <> mState.ȱҩ _
                                    And InStr(1, mcondition.str��Σ����, .TextMatrix(i, mIntCol��ҩ_��ΣҩƷ)) = 0 Then
                                    .TextMatrix(i, mIntCol��ҩ_ִ��״̬) = intState
                                    .TextMatrix(i, mIntCol��ҩ_״̬) = strState
                                    
                                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                                    
                                    mrsSendList.Filter = "�շ�ID=" & Val(.TextMatrix(i, mIntCol��ҩ_�շ�ID))
                                    
                                    mrsSendList!ִ��״̬ = intState
                                    mrsSendList!״̬ = strState
                                    
                                    mrsSendList.Update
                                End If
                            Next
                        End If
                    End If
                    
                    SetMainComandBars index, .Row
                    DoEvents
                    Call RefreshList_Sum
                Else
                    If mcondition.intShowPass = 3 And Not gobjPass Is Nothing And IsInString(gstrprivs, "������ҩ���", ";") Then
                        Call gobjPass.queryCheckResult(.TextMatrix(.Row, mIntCol��ҩ_סԺ��), "2")
                    End If
                End If
            Case mListType.�ܷ�
                If .Row = 0 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol�ܷ�_�շ�ID)) = 0 Then Exit Sub
                If .Col = mIntCol�ܷ�_״̬ Then
                    If Val(.TextMatrix(.Row, mIntCol�ܷ�_ִ��״̬)) = mState.�ܷ�_������ Then
                        .TextMatrix(.Row, .Col) = "�ָ�"
                        .TextMatrix(.Row, mIntCol�ܷ�_ִ��״̬) = mState.�ܷ�_�ָ�
                        lngColor = mListColor.State_RejectRestore
                    ElseIf Val(.TextMatrix(.Row, mIntCol�ܷ�_ִ��״̬)) = mState.�ܷ�_�ָ� Then
                        .TextMatrix(.Row, .Col) = "������"
                        .TextMatrix(.Row, mIntCol�ܷ�_ִ��״̬) = mState.�ܷ�_������
                        lngColor = mListColor.State_RejectUnProcess
                    End If
                    
                    .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = lngColor
                    
                    mrsSendList.Filter = "�շ�ID=" & Val(.TextMatrix(.Row, mIntCol�ܷ�_�շ�ID))
                    
                    mrsSendList!״̬ = .TextMatrix(.Row, mIntCol�ܷ�_״̬)
                    mrsSendList!ִ��״̬ = Val(.TextMatrix(.Row, mIntCol�ܷ�_ִ��״̬))
                    mrsSendList.Update
                    
                    SetMainComandBars index, .Row
                End If
            Case mListType.��ҩ
                If .Row = 0 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol��ҩ_�շ�ID)) = 0 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol��ҩ_׼����)) = 0 Then Exit Sub
                If .Col = mIntCol��ҩ_״̬ Then
                    If Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ_ԭʼ��¼ Then
                        .TextMatrix(.Row, .Col) = "��ҩ"
                        .TextMatrix(.Row, mIntCol��ҩ_��ҩ��) = .TextMatrix(.Row, mIntCol��ҩ_׼����)
                        .TextMatrix(.Row, mIntCol��ҩ_ִ��״̬) = mState.��ҩ
                    ElseIf Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ Then
                        .TextMatrix(.Row, .Col) = "������"
                        .TextMatrix(.Row, mIntCol��ҩ_��ҩ��) = ""
                        .TextMatrix(.Row, mIntCol��ҩ_ִ��״̬) = mState.��ҩ_ԭʼ��¼
                    End If
                    
                    mrsReturnList.Filter = "�շ�ID=" & Val(.TextMatrix(.Row, mIntCol��ҩ_�շ�ID))
                    
                    mrsReturnList!״̬ = .TextMatrix(.Row, mIntCol��ҩ_״̬)
                    mrsReturnList!ִ��״̬ = Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬))
                    mrsReturnList.Update
                End If
        End Select
    End With
End Sub

Private Sub vsfList_DrawCell(index As Integer, ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
    '      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
    '      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long
    
    Dim lngStateColor As Long
    
    If index <> mListType.��ҩ Then Exit Sub
    
    With vsfList(index)
        If Col = mIntCol��ҩ_Ʒ�� And .IsSubtotal(Row) = False Then
            '������Ԫ���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom - 1
            
            If Row = 0 Then
                SetBkColor hDC, SysColor2RGB(.BackColorFixed)
            Else
                If .IsSelected(Row) = True Then
                    lngStateColor = .BackColorSel
                Else
                    '����״̬����ɫ
                    If .TextMatrix(Row, mIntCol��ҩ_ִ��״̬) <> "" Then
                        If .TextMatrix(Row, mIntCol��ҩ_ִ��״̬) = mState.ȱҩ Then
                            lngStateColor = mListColor.State_Shortage
                        ElseIf .TextMatrix(Row, mIntCol��ҩ_ִ��״̬) = mState.��ҩ Then
                            lngStateColor = mListColor.State_Send
                        ElseIf .TextMatrix(Row, mIntCol��ҩ_ִ��״̬) = mState.�ܷ� Then
                            lngStateColor = mListColor.State_Reject
                        ElseIf .TextMatrix(Row, mIntCol��ҩ_ִ��״̬) = mState.������ Then
                            lngStateColor = mListColor.State_UnProcess
                        End If
                    Else
                        lngStateColor = .BackColorSel
                    End If
                End If
                
                SetBkColor hDC, SysColor2RGB(lngStateColor)
            End If
            
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
    End With
End Sub

Private Sub vsfList_EnterCell(index As Integer)
    Dim lng������id As Long
    Dim strTempFile As String
    
    If mblnOutPut = True Then Exit Sub
    If mblnRefresh = True Then Exit Sub
    
    vsfChargeOff.Visible = False
    
    With vsfList(index)
        
        If .Row = 0 Then Exit Sub
        
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = frmPublic.ImgList.ListImages(2).Picture
        
        .Editable = flexEDNone
        
        Select Case index
            Case mListType.��ҩ
                If mblnShowReject = False Then
                    If vsfList(mListType.��ҩ).IsSubtotal(.Row) = False And (.TextMatrix(.Row, mIntCol��ҩ_��ҩĿ��) <> "" Or .TextMatrix(.Row, mIntCol��ҩ_��ҩ����) <> "") Then
                        Me.txt��ҩ����.Text = "��ҩĿ�ģ�" & .TextMatrix(.Row, mIntCol��ҩ_��ҩĿ��) & vbCrLf & "��ҩ���ɣ�" & .TextMatrix(.Row, mIntCol��ҩ_��ҩ����)
                    Else
                        Me.txt��ҩ����.Text = ""
                    End If
                    
                    '��ʾ�������
                    Call GetDiagnosis(.Row)
                    
                    '��ȡ�����˵�ǩ��ͼƬ
                    picǩ��ͼƬ.Picture = Nothing
                    picǩ��ͼƬ.Visible = False
                    If vsfList(mListType.��ҩ).IsSubtotal(.Row) = False And IsNumeric(.TextMatrix(.Row, mIntCol��ҩ_�շ�ID)) Then
                        lng������id = get������id(.TextMatrix(.Row, mIntCol��ҩ_����ҽ��), Val(.TextMatrix(.Row, mIntCol��ҩ_��������id)))
                        strTempFile = Sys.ReadLob(100, 15, lng������id)
                        If strTempFile <> "" Then
                            picSign.Picture = LoadPicture(strTempFile)
                            If Not picSign.Picture Is Nothing Then
                                picǩ��ͼƬ.Visible = True
                                picǩ��ͼƬ.PaintPicture picSign.Picture, 0, 0, picǩ��ͼƬ.ScaleX(picǩ��ͼƬ.Width, vbTwips, vbPixels), picǩ��ͼƬ.ScaleY(picǩ��ͼƬ.Height, vbTwips, vbPixels)
                            End If
                            Kill strTempFile
                        End If
                    End If
                    
                    If Not gobjPass Is Nothing Then Call gobjPass.zlPassSetDrug_YF(.TextMatrix(.Row, mIntCol��ҩ_ҩƷID), .TextMatrix(.Row, mIntCol��ҩ_ҩƷ����))
                    If Not gobjPass Is Nothing Then Call gobjPass.zlPassClearLight_YF
                End If
            Case mListType.����
                If .Row = 0 Then Exit Sub
                
                If mcondition.bln�����һ��� = False Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol���һ���_��ҩ����id)) = 0 Then Exit Sub
                
                If .Col = mIntCol���һ���_�������� Or .Col = mIntCol���һ���_�������� Then
                    If mcondition.bln�޸��������� = True Then .Editable = flexEDKbdMouse
                End If
                

                vsfList(mListType.����).Height = mdblSumListHeight
                
                If RefreshList_ChargeOff(Val(.TextMatrix(.Row, mIntCol���һ���_��ҩ����id)), Val(.TextMatrix(.Row, mIntCol���һ���_ҩƷID))) = True Then
                    vsfChargeOff.Visible = True
                    picHsc.Visible = True
                    
                    Call ResizeChargeOffList
                End If
            Case mListType.��ҩ
                If .Row = 0 Then Exit Sub
                SetMainComandBars index, .Row
                If Val(.TextMatrix(.Row, mIntCol��ҩ_�շ�ID)) = 0 Then Exit Sub
                
                '����PASS��ť״̬
                SetPassMenuButton index, .Row
                
                If Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬)) <> mState.��ҩ And Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬)) <> mState.��ҩ_ԭʼ��¼ Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol��ҩ_׼����)) = 0 Then Exit Sub
                Select Case .Col
                    Case mIntCol��ҩ_��ҩ��
                        .Editable = flexEDKbdMouse
                        
                        If Val(.TextMatrix(.Row, mIntCol��ҩ_��ҩ��)) = 0 Then
                            .TextMatrix(.Row, mIntCol��ҩ_��ҩ��) = .TextMatrix(.Row, mIntCol��ҩ_׼����)
                            .TextMatrix(.Row, mIntCol��ҩ_״̬) = "��ҩ"
                            .TextMatrix(.Row, mIntCol��ҩ_ִ��״̬) = mState.��ҩ
                            
                            mrsReturnList.Filter = "�շ�ID=" & Val(.TextMatrix(.Row, mIntCol��ҩ_�շ�ID))
                            mrsReturnList!״̬ = .TextMatrix(.Row, mIntCol��ҩ_״̬)
                            mrsReturnList!ִ��״̬ = Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬))
                            mrsReturnList!��ҩ�� = Val(.TextMatrix(.Row, mIntCol��ҩ_��ҩ��))
                            mrsReturnList.Update
                        End If
                End Select
            End Select
            
            SetMainComandBars index, .Row
    End With
End Sub

Private Function get������id(ByVal str���� As String, ByVal lng��������id As Long) As Long
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    
    strsql = "select A.ID from ��Ա�� A,������Ա B where A.id=B.��Աid and A.����=[1] and B.����id=[2]"
    Set rstemp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, str����, lng��������id)
    
    If Not rstemp.EOF Then
        get������id = rstemp!Id
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfList_KeyPressEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    
    With vsfList(index)
        Select Case index
            Case mListType.��ҩ
                If Col = mIntCol��ҩ_��ҩ�� Then
                    strKey = .EditText
                    If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13) + Chr(Asc("-")), Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                        If .EditSelLength = Len(strKey) Then Exit Sub
                        If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                            KeyAscii = 0
                            Exit Sub
                        End If
                        If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= mintNumberDigit And strKey Like "*.*" Then
                            KeyAscii = 0
                            Exit Sub
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            Case mListType.����
                If Col = mIntCol���һ���_�������� Or Col = mIntCol���һ���_ʵ������ Then
                    If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    ElseIf KeyAscii = Asc(".") Then
                        If InStr(.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                            KeyAscii = 0
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfList_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim LngID As Long
    Dim Int���� As Integer
    Dim strNo As String
    Dim str����� As String
    Dim lngҽ��id As Long
    Dim strsql As String
    Dim rsTmp As Recordset
    Dim lngPatiID As Long
    Dim str�Һŵ� As String
    Dim lng��ҳID As Long
    
    '�����ڱ��е����˵�
    If vsfList(index).MouseRow < 1 Then Exit Sub
    If vsfList(index).MouseCol < 1 Then Exit Sub
    If vsfList(index).IsSubtotal(vsfList(index).MouseRow) = True Then Exit Sub
    
    '��ҩ״̬�����˵�
    If index = mListType.��ҩ Then
        If Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("�շ�ID"))) = 0 Then Exit Sub
        If vsfList(index).MouseCol <> vsfList(index).ColIndex("�����") Then
            If Button = 2 Then
                Select Case Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, mIntCol��ҩ_ִ��״̬))
                    Case mState.��ҩ
                        LngID = conMenu_Status_Verify
                    Case mState.�ܷ�
                        LngID = conMenu_Status_Reject
                    Case mState.������
                        LngID = conMenu_Status_NoProcess
                End Select
            
                If Me.cbsMain Is Nothing Then Exit Sub
                Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_StatusPopup)
                If Not objPopup Is Nothing Then
                    For Each cbrControl In objPopup.Controls
                        If cbrControl.Id < 320 Then
                            cbrControl.Visible = True
                        Else
                            cbrControl.Visible = False
                        End If
                    Next
                    
                    Set cbrControl = objPopup.Controls.Find(xtpControlButton, conMenu_Status_Verify, , True)
                    If Not cbrControl Is Nothing Then
                        cbrControl.Checked = (cbrControl.Id = LngID)
                    End If
                    
                    Set cbrControl = objPopup.Controls.Find(xtpControlButton, conMenu_Status_Reject, , True)
                    If Not cbrControl Is Nothing Then
                        cbrControl.Checked = (cbrControl.Id = LngID)
                    End If
        
                    Set cbrControl = objPopup.Controls.Find(xtpControlButton, conMenu_Status_NoProcess, , True)
                    If Not cbrControl Is Nothing Then
                        cbrControl.Checked = (cbrControl.Id = LngID)
                    End If
                    
                    objPopup.CommandBar.ShowPopup
                End If
            End If
        End If
    End If
    
    '��ҩ״̬�����˵�
    If index = mListType.��ҩ Then
        If Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("�շ�ID"))) = 0 Then Exit Sub
        If Button = 2 Then
            If Me.cbsMain Is Nothing Then Exit Sub
            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_StatusPopup)
            If Not objPopup Is Nothing Then
                For Each cbrControl In objPopup.Controls
                    If cbrControl.Id < 320 Then
                        cbrControl.Visible = False
                    Else
                        cbrControl.Visible = True
                    End If
                Next
                
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
    
    If Button = 1 And vsfList(index).MouseCol = vsfList(index).ColIndex("�����") And index = mListType.��ҩ Then
        If IsInString(gstrprivs, "������ҩ���", ";") And (index = mListType.��ҩ Or index = mListType.��ҩ) Then
            If Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("�շ�ID"))) = 0 Then Exit Sub
'            If vsfList(index).Cell(flexcpPicture, vsfList(index).MouseRow, mIntCol��ҩ_�����, vsfList(index).MouseRow, mIntCol��ҩ_�����) Is Nothing Then Exit Sub
            Int���� = Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("����")))
            strNo = vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("NO"))
            str����� = vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("���"))
            lngҽ��id = Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("ҽ��id")))
            
            '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�������˳�
            strsql = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.����id,0 ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,������ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] "
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, strNo, Int����)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!����ID
            str�Һŵ� = NVL(rsTmp!�Һŵ�)
            lng��ҳID = rsTmp!��ҳID
            
            '��ȡpass�˵�������
            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_PASS)
            
            Call gobjPass.zlPASSPopupCommandBars_YF(mlngMode, objPopup.CommandBar, mconMenu_PASS, lngPatiID, lng��ҳID, str�Һŵ�, str�����, lngҽ��id)
            
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub SetColumnValue(ByVal intType As Integer, ByVal str���� As String, ByVal intValue As Integer)
    Select Case intType
        Case mListType.��ҩ
            Select Case str����
                Case "���˿���"
                    mIntCol��ҩ_���� = intValue
                Case "����ҽ��"
                    mIntCol��ҩ_����ҽ�� = intValue
                Case "״̬"
                    mIntCol��ҩ_״̬ = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "��ҩ����"
                    mIntCol��ҩ_��ҩ���� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "NO"
                    mIntCol��ҩ_NO = intValue
                Case "����Ա"
                    mIntCol��ҩ_����Ա = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "�Ա�"
                    mIntCol��ҩ_�Ա� = intValue
                Case "סԺ��"
                    mIntCol��ҩ_סԺ�� = intValue
                Case "ҩƷ����"
                    mIntCol��ҩ_Ʒ�� = intValue
                Case "Ƥ�Խ��"
                    mIntCol��ҩ_Ƥ�Խ�� = intValue
                Case "������"
                    mIntCol��ҩ_������ = intValue
                Case "Ӣ����"
                    mIntCol��ҩ_Ӣ���� = intValue
                Case "�䷽����"
                    mIntCol��ҩ_�䷽���� = intValue
                Case "���"
                    mIntCol��ҩ_��� = intValue
                Case "������"
                    mIntCol��ҩ_������ = intValue
                Case "ԭ����"
                    mIntCol��ҩ_ԭ���� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "Ч��"
                    mIntCol��ҩ_Ч�� = intValue
                Case "��"
                    mIntCol��ҩ_�� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "���"
                    mIntCol��ҩ_��� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "Ƶ��"
                    mIntCol��ҩ_Ƶ�� = intValue
                Case "�÷�"
                    mIntCol��ҩ_�÷� = intValue
                Case "��ҩ����"
                    mIntCol��ҩ_��ҩ���� = intValue
                Case "����ʱ��"
                    mIntCol��ҩ_����ʱ�� = intValue
                Case "˵��"
                    mIntCol��ҩ_˵�� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                 Case "��ҩ��"
                    mIntCol��ҩ_��ҩ�� = intValue
                Case "�ⷿ��λ"
                    mIntCol��ҩ_�ⷿ��λ = intValue
                Case "������λ"
                    mIntCol��ҩ_������λ = intValue
                Case "��ҩ����"
                    mIntCol��ҩ_��ҩ���� = intValue
                Case "��������"
                    mIntCol��ҩ_�������� = intValue
                Case "����ҩƷ˵��"
                    mIntCol��ҩ_����ҩƷ˵�� = intValue
                Case "��ע"
                    mIntCol��ҩ_��ע = intValue
            End Select
        Case mListType.����
            If mcondition.bln�����һ��� Then
                Select Case str����
                    Case "���˿���"
                        mIntCol���һ���_���� = intValue
                    Case "ҩƷ����"
                        mIntCol���һ���_Ʒ�� = intValue
                    Case "���"
                        mIntCol���һ���_��� = intValue
                    Case "������"
                        mIntCol���һ���_������ = intValue
                    Case "ԭ����"
                        mIntCol���һ���_ԭ���� = intValue
                    Case "����"
                        mIntCol���һ���_���� = intValue
                    Case "Ч��"
                        mIntCol���һ���_Ч�� = intValue
                    Case "Ӧ������"
                        mIntCol���һ���_Ӧ������ = intValue
                    Case "��������"
                        mIntCol���һ���_�������� = intValue
                    Case "��������"
                        mIntCol���һ���_�������� = intValue
                    Case "ʵ������"
                        mIntCol���һ���_ʵ������ = intValue
                    Case "��λ"
                        mIntCol���һ���_��λ = intValue
                
                    Case "����"
                        mIntCol���һ���_���� = intValue
                    Case "Ӧ�����"
                        mIntCol���һ���_Ӧ����� = intValue
                    Case "����"
                        mIntCol���һ���_���� = intValue
                    Case "����ID"
                        mIntCol���һ���_����ID = intValue
                    Case "ҩƷID"
                        mIntCol���һ���_ҩƷID = intValue
                        
                    Case "��ҩ����"
                        mIntCol���һ���_��ҩ���� = intValue
                    Case "��ҩ����id"
                        mIntCol���һ���_��ҩ����id = intValue
                End Select
            Else
                Select Case str����
                    Case "ҩƷ����"
                        mIntCol����_Ʒ�� = intValue
                    Case "���"
                        mIntCol����_��� = intValue
                    Case "������"
                        mIntCol����_������ = intValue
                    Case "ԭ����"
                        mIntCol����_ԭ���� = intValue
                    Case "����"
                        mIntCol����_���� = intValue
                    Case "Ч��"
                        mIntCol����_Ч�� = intValue
                    Case "����"
                        mIntCol����_���� = intValue
                    Case "��λ"
                        mIntCol����_��λ = intValue
                    Case "����"
                        mIntCol����_���� = intValue
                    Case "���"
                        mIntCol����_��� = intValue
                End Select
            End If
        Case mListType.ȱҩ
            Select Case str����
                Case "���˿���"
                    mIntColȱҩ_���� = intValue
                Case "NO"
                    mIntColȱҩ_NO = intValue
                Case "����"
                    mIntColȱҩ_���� = intValue
                Case "����"
                    mIntColȱҩ_���� = intValue
                Case "����"
                    mIntColȱҩ_���� = intValue
                Case "�Ա�"
                    mIntColȱҩ_�Ա� = intValue
                Case "��ҩ����"
                    mIntColȱҩ_��ҩ���� = intValue
                Case "ҩƷ����"
                    mIntColȱҩ_Ʒ�� = intValue
                Case "���"
                    mIntColȱҩ_��� = intValue
                Case "������"
                    mIntColȱҩ_������ = intValue
                Case "ԭ����"
                    mIntColȱҩ_ԭ���� = intValue
                Case "����"
                    mIntColȱҩ_���� = intValue
                Case "Ч��"
                    mIntColȱҩ_Ч�� = intValue
                Case "����"
                    mIntColȱҩ_���� = intValue
                Case "����"
                    mIntColȱҩ_���� = intValue
                Case "���"
                    mIntColȱҩ_��� = intValue
                Case "��ע"
                    mIntColȱҩ_��ע = intValue
            End Select
        Case mListType.�ܷ�
            Select Case str����
                Case "���˿���"
                    mIntCol�ܷ�_���� = intValue
                Case "״̬"
                    mIntCol�ܷ�_״̬ = intValue
                Case "NO"
                    mIntCol�ܷ�_NO = intValue
                Case "����"
                    mIntCol�ܷ�_���� = intValue
                Case "��ҩ����"
                    mIntCol�ܷ�_��ҩ���� = intValue
                Case "����"
                    mIntCol�ܷ�_���� = intValue
                Case "����"
                    mIntCol�ܷ�_���� = intValue
                Case "�Ա�"
                    mIntCol�ܷ�_�Ա� = intValue
                Case "ҩƷ����"
                    mIntCol�ܷ�_Ʒ�� = intValue
                Case "���"
                    mIntCol�ܷ�_��� = intValue
                Case "������"
                    mIntCol�ܷ�_������ = intValue
                Case "ԭ����"
                    mIntCol�ܷ�_ԭ���� = intValue
                Case "����"
                    mIntCol�ܷ�_���� = intValue
                Case "Ч��"
                    mIntCol�ܷ�_Ч�� = intValue
                Case "����"
                    mIntCol�ܷ�_���� = intValue
                Case "����"
                    mIntCol�ܷ�_���� = intValue
                Case "���"
                    mIntCol�ܷ�_��� = intValue
                Case "��ע"
                    mIntCol�ܷ�_��ע = intValue
            End Select
        Case mListType.��ҩ
            Select Case str����
                Case "����ʱ��"
                    mIntCol��ҩ_����ʱ�� = intValue
                Case "���˿���"
                    mIntCol��ҩ_���� = intValue
                Case "״̬"
                    mIntCol��ҩ_״̬ = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "NO"
                    mIntCol��ҩ_NO = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "�Ա�"
                    mIntCol��ҩ_�Ա� = intValue
                Case "סԺ��"
                    mIntCol��ҩ_סԺ�� = intValue
                Case "ҩƷ����"
                    mIntCol��ҩ_Ʒ�� = intValue
                Case "��ҩ����"
                    mIntCol��ҩ_��ҩ���� = intValue
                Case "������"
                    mIntCol��ҩ_������ = intValue
                Case "Ӣ����"
                    mIntCol��ҩ_Ӣ���� = intValue
                Case "���"
                    mIntCol��ҩ_��� = intValue
                Case "������"
                    mIntCol��ҩ_������ = intValue
                Case "ԭ����"
                    mIntCol��ҩ_ԭ���� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "Ч��"
                    mIntCol��ҩ_Ч�� = intValue
                Case "��"
                    mIntCol��ҩ_�� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "������"
                    mIntCol��ҩ_������ = intValue
                Case "׼����"
                    mIntCol��ҩ_׼���� = intValue
                Case "��ҩ��"
                    mIntCol��ҩ_��ҩ�� = intValue
        
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "���"
                    mIntCol��ҩ_��� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "Ƶ��"
                    mIntCol��ҩ_Ƶ�� = intValue
                Case "�÷�"
                    mIntCol��ҩ_�÷� = intValue
                Case "����Ա"
                    mIntCol��ҩ_����Ա = intValue
                Case "��ҩʱ��"
                    mIntCol��ҩ_��ҩʱ�� = intValue
                Case "����"
                    mIntCol��ҩ_���� = intValue
                Case "ҽ��id"
                    mIntCol��ҩ_ҽ��id = intValue
                Case "��/��ҩ��"
                    mIntCol��ҩ_��ҩ�� = intValue
                    
                Case "�ⷿ��λ"
                    mIntCol��ҩ_�ⷿ��λ = intValue
                Case "���ID"
                    mIntCol��ҩ_���ID = intValue
                Case "ҩƷID"
                    mIntCol��ҩ_ҩƷID = intValue
                Case "������λ"
                    mIntCol��ҩ_������λ = intValue
                Case "��ҩ��"
                    mIntCol��ҩ_��ҩ�� = intValue
                Case "��ע"
                    mIntCol��ҩ_��ע = intValue
            End Select
    End Select
End Sub

Public Sub VerifySign()
    Dim rsData As Recordset
    
    With vsfList(mListType.��ҩ)
        If Val(.TextMatrix(.Row, mIntCol��ҩ_�շ�ID)) = 0 Then Exit Sub
        
        '����������˵���ǩ��������Ҫ�Է�ҩ/��ҩ�˽��е���ǩ������
        If Val(.TextMatrix(.Row, mIntCol��ҩ_ִ��״̬)) = mState.��ҩ_��ҩ��¼ Then
            '��ҩ��¼��֤
            If VerifySignatureRecored_bak(EsignTache.returnStep, .TextMatrix(.Row, mIntCol��ҩ_����), .TextMatrix(.Row, mIntCol��ҩ_NO), mcondition.lngҩ��id, .TextMatrix(.Row, mIntCol��ҩ_�շ�ID)) = False Then
                Exit Sub
            End If
        Else
            '��ҩ��¼��֤
            If VerifySignatureRecoredGather(EsignTache.send, .TextMatrix(.Row, mIntCol��ҩ_�շ�ID)) = False Then
                Exit Sub
            End If
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub





