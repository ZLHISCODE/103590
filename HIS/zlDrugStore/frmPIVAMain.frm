VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPIVAMain 
   Caption         =   "������Һ�������Ĺ���"
   ClientHeight    =   11700
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   17910
   Icon            =   "frmPIVAMain.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   11700
   ScaleWidth      =   17910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraTip 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9240
      TabIndex        =   107
      Top             =   1200
      Width           =   840
      Begin VB.PictureBox pic�Ա�ҩ 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   0
         Picture         =   "frmPIVAMain.frx":058A
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   108
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lbl�Ա�ҩ 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�ҩ"
         Height          =   180
         Index           =   3
         Left            =   255
         TabIndex        =   109
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   3480
      ScaleHeight     =   780
      ScaleWidth      =   9375
      TabIndex        =   50
      Top             =   7200
      Width           =   9375
      Begin VB.TextBox txtFindItem 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   7320
         MaxLength       =   13
         TabIndex        =   51
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6480
         TabIndex        =   53
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblFindItem 
         AutoSize        =   -1  'True
         Caption         =   "ƿǩ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   52
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.PictureBox picPacker 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   8760
      Picture         =   "frmPIVAMain.frx":6DDC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   45
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picDept 
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   1
      Left            =   8280
      ScaleHeight     =   1695
      ScaleWidth      =   3015
      TabIndex        =   37
      Top             =   8160
      Width           =   3015
      Begin VSFlex8Ctl.VSFlexGrid vsfDept 
         Height          =   1200
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   2760
         _cx             =   4868
         _cy             =   2117
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
         BackColorSel    =   16771280
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":D62E
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
   End
   Begin VB.PictureBox picDept 
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   0
      Left            =   3720
      ScaleHeight     =   1695
      ScaleWidth      =   3015
      TabIndex        =   26
      Top             =   7800
      Width           =   3015
      Begin VSFlex8Ctl.VSFlexGrid vsfDept 
         Height          =   1200
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   2760
         _cx             =   4868
         _cy             =   2117
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
         BackColorSel    =   16771280
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":D6DD
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
   End
   Begin VB.PictureBox picPacker 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   8520
      Picture         =   "frmPIVAMain.frx":D78C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picPacker 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   8280
      Picture         =   "frmPIVAMain.frx":DD16
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picPrint 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   8520
      Picture         =   "frmPIVAMain.frx":E2A0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picPrint 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   8280
      Picture         =   "frmPIVAMain.frx":E82A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picDetailList 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   3600
      ScaleHeight     =   6255
      ScaleWidth      =   13695
      TabIndex        =   13
      Top             =   1800
      Width           =   13695
      Begin VB.TextBox txtDia 
         Enabled         =   0   'False
         Height          =   855
         Left            =   10800
         TabIndex        =   92
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Frame fraMedis 
         Height          =   615
         Left            =   360
         TabIndex        =   79
         Top             =   240
         Width           =   9855
         Begin VB.CheckBox chkCheck 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "ȫѡ"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboType 
            Height          =   300
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   84
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5370
            TabIndex        =   83
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   6420
            TabIndex        =   82
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   7470
            TabIndex        =   81
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   8520
            TabIndex        =   80
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblType 
            BackColor       =   &H80000004&
            Caption         =   "��ҩ����"
            Height          =   180
            Left            =   1200
            TabIndex        =   87
            Top             =   277
            Width           =   780
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":EDB4
            Height          =   240
            Index           =   0
            Left            =   4560
            Picture         =   "frmPIVAMain.frx":15606
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":1BE58
            Height          =   240
            Index           =   1
            Left            =   5610
            Picture         =   "frmPIVAMain.frx":226AA
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":28EFC
            Height          =   240
            Index           =   2
            Left            =   6660
            Picture         =   "frmPIVAMain.frx":2F74E
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":35FA0
            Height          =   240
            Index           =   3
            Left            =   7710
            Picture         =   "frmPIVAMain.frx":3C7F2
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":43044
            Height          =   240
            Index           =   4
            Left            =   8760
            Picture         =   "frmPIVAMain.frx":49896
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.Frame fraDetailCtr 
         BackColor       =   &H00FFEDDD&
         Height          =   840
         Left            =   -120
         TabIndex        =   54
         Top             =   840
         Width           =   15015
         Begin VB.ComboBox cboSort 
            Height          =   300
            Left            =   12720
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   480
            Width           =   2500
         End
         Begin VB.ComboBox cboFrequency 
            Height          =   300
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkSure 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "��ȷ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   73
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkSure 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "δȷ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   72
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "���"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   71
            Top             =   150
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.ComboBox cboMedi 
            Height          =   300
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   480
            Width           =   3700
         End
         Begin VB.ComboBox cboLevel 
            Height          =   300
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   480
            Width           =   615
         End
         Begin VB.ComboBox cboBatch 
            Height          =   300
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkSendType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "�ѷ���"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5400
            TabIndex        =   67
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkSendType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "δ����"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   4800
            TabIndex        =   66
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "��ҩ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   65
            Top             =   150
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkPack 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "����ҩ(���)����"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   64
            Top             =   150
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.OptionButton optShowType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "��Ҫ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   7080
            TabIndex        =   63
            Top             =   150
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton optShowType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "��ϸ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   8040
            TabIndex        =   62
            Top             =   150
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chkAll 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "ȫѡ"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   150
            Width           =   735
         End
         Begin VB.CheckBox chkDept 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "����������"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   150
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "�Ѵ�ӡ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   59
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "δ��ӡ"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   58
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkChange 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "�ѱ�"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   6360
            TabIndex        =   57
            Top             =   150
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkChange 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "δ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   56
            Top             =   150
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.ComboBox cboDosType 
            Height          =   300
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblSort 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "����ʽ"
            Height          =   180
            Left            =   11880
            TabIndex        =   90
            Top             =   540
            Width           =   720
         End
         Begin VB.Label lblFrequency 
            BackColor       =   &H00FFEDDD&
            Caption         =   "Ƶ��"
            Height          =   180
            Left            =   5280
            TabIndex        =   89
            Top             =   540
            Width           =   420
         End
         Begin VB.Label lblMedi 
            BackColor       =   &H00FFEDDD&
            Caption         =   "ҩƷ"
            Height          =   180
            Left            =   7560
            TabIndex        =   78
            Top             =   540
            Width           =   420
         End
         Begin VB.Label lblBatch 
            BackColor       =   &H00FFEDDD&
            Caption         =   "����"
            Height          =   180
            Left            =   2280
            TabIndex        =   77
            Top             =   540
            Width           =   420
         End
         Begin VB.Label lblLevel 
            BackColor       =   &H00FFEDDD&
            Caption         =   "���ȼ�"
            Height          =   180
            Left            =   3840
            TabIndex        =   76
            Top             =   540
            Width           =   540
         End
         Begin VB.Label lblVolu 
            BackColor       =   &H00FFEDDD&
            Caption         =   "������0"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   9960
            TabIndex        =   75
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblDosType 
            BackColor       =   &H00FFEDDD&
            Caption         =   "����"
            Height          =   180
            Left            =   120
            TabIndex        =   74
            Top             =   540
            Width           =   420
         End
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   9000
         TabIndex        =   47
         ToolTipText     =   "�ȼ���F2"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtLog 
         Height          =   855
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   3960
         Width           =   3855
      End
      Begin VB.Frame fraH 
         Height          =   30
         Left            =   240
         MousePointer    =   7  'Size N S
         TabIndex        =   31
         Top             =   2760
         Width           =   9375
      End
      Begin VB.PictureBox picHelp 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   240
         ScaleHeight     =   225
         ScaleWidth      =   9975
         TabIndex        =   22
         Top             =   0
         Width           =   9975
         Begin VB.PictureBox picHelpIcon 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   0
            Picture         =   "frmPIVAMain.frx":500E8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   23
            Top             =   20
            Width           =   240
         End
         Begin VB.Label lblCount 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   5760
            TabIndex        =   28
            Top             =   45
            Width           =   4170
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "��ʾ��"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   260
            TabIndex        =   24
            Top             =   50
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfTrans 
         Height          =   840
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   4560
         _cx             =   8043
         _cy             =   1482
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777215
         GridColor       =   10329501
         GridColorFixed  =   10329501
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   61
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":5693A
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
         ExplorerBar     =   2
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
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmPIVAMain.frx":5710D
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSumDrug 
         Height          =   1200
         Left            =   8280
         TabIndex        =   16
         Top             =   1680
         Width           =   1920
         _cx             =   3387
         _cy             =   2117
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   15724527
         GridColor       =   10329501
         GridColorFixed  =   10329501
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":5765B
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
      Begin VSFlex8Ctl.VSFlexGrid vsfMedis 
         Height          =   840
         Left            =   5400
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   2400
         _cx             =   4233
         _cy             =   1482
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777215
         GridColor       =   10329501
         GridColorFixed  =   10329501
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   35
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":5781B
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
         Editable        =   1
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
      Begin VSFlex8Ctl.VSFlexGrid VSFLook 
         Height          =   1440
         Left            =   240
         TabIndex        =   30
         Top             =   2880
         Width           =   8640
         _cx             =   15240
         _cy             =   2540
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777215
         GridColor       =   10329501
         GridColorFixed  =   10329501
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   20
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":57C6F
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
         ExplorerBar     =   2
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
      Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
         Height          =   855
         Left            =   0
         TabIndex        =   49
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
         FormatString    =   $"frmPIVAMain.frx":57F0D
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
      Begin VB.Label lblLog 
         Caption         =   "�������"
         Height          =   255
         Left            =   8280
         TabIndex        =   94
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label lblDia 
         Caption         =   "�������"
         Height          =   255
         Left            =   10800
         TabIndex        =   93
         Top             =   3480
         Width           =   1335
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   3480
      ScaleHeight     =   1455
      ScaleWidth      =   2295
      TabIndex        =   10
      Top             =   120
      Width           =   2295
      Begin VB.Frame fraLineV1 
         BackColor       =   &H80000012&
         Height          =   2085
         Left            =   120
         TabIndex        =   11
         Top             =   -120
         Width           =   50
      End
      Begin XtremeSuiteControls.TabControl tbcDetail 
         Height          =   975
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin VB.PictureBox picCondition 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   120
      ScaleHeight     =   8055
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox picMsg 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         MouseIcon       =   "frmPIVAMain.frx":57F5B
         ScaleHeight     =   2055
         ScaleWidth      =   2895
         TabIndex        =   39
         Tag             =   "0"
         Top             =   6000
         Width           =   2895
         Begin VB.Frame fraMsg 
            Height          =   50
            Left            =   -20
            TabIndex        =   42
            Top             =   0
            Width           =   3405
         End
         Begin VB.PictureBox picUpOrDown 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   2400
            Picture         =   "frmPIVAMain.frx":58265
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   40
            Top             =   60
            Width           =   270
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfMsg 
            Height          =   1560
            Left            =   0
            TabIndex        =   44
            Top             =   405
            Width           =   2880
            _cx             =   5080
            _cy             =   2752
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
            BackColorSel    =   16771280
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483633
            BackColorAlternate=   16777215
            GridColor       =   10329501
            GridColorFixed  =   10329501
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPIVAMain.frx":585A7
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
         Begin VB.Label lblMsgComment 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "��Ϣ����(0)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   45
            TabIndex        =   41
            Top             =   90
            Width           =   1095
         End
      End
      Begin VB.PictureBox picLook 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1305
         ScaleWidth      =   2025
         TabIndex        =   34
         Top             =   4320
         Width           =   2055
         Begin XtremeSuiteControls.TabControl tbcLook 
            Height          =   1215
            Left            =   0
            TabIndex        =   35
            Top             =   240
            Width           =   2055
            _Version        =   589884
            _ExtentX        =   3625
            _ExtentY        =   2143
            _StockProps     =   64
            Enabled         =   -1  'True
         End
      End
      Begin VB.PictureBox picWork 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   1200
         ScaleHeight     =   1305
         ScaleWidth      =   1905
         TabIndex        =   32
         Top             =   3840
         Width           =   1935
         Begin XtremeSuiteControls.TabControl tabWork 
            Height          =   1095
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1695
            _Version        =   589884
            _ExtentX        =   2990
            _ExtentY        =   1931
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picDeptList 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   0
         ScaleHeight     =   1695
         ScaleWidth      =   2895
         TabIndex        =   8
         Top             =   3840
         Width           =   2895
         Begin VB.CommandButton cmdRefreshTrans 
            Caption         =   "ˢ����ϸ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   48
            Top             =   0
            Width           =   1095
         End
         Begin VB.CheckBox chkAllDept 
            Appearance      =   0  'Flat
            Caption         =   "ȫѡ"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   43
            Top             =   40
            Width           =   735
         End
         Begin VB.Frame fraLineH1 
            Height          =   50
            Left            =   -20
            TabIndex        =   9
            Top             =   0
            Width           =   3405
         End
         Begin XtremeSuiteControls.TabControl tabDeptList 
            Height          =   1455
            Left            =   120
            TabIndex        =   36
            Top             =   0
            Width           =   2535
            _Version        =   589884
            _ExtentX        =   4471
            _ExtentY        =   2566
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picTime 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   3135
         TabIndex        =   1
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtdept 
            Height          =   315
            Left            =   840
            TabIndex        =   105
            Top             =   1920
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton cmdDrug 
            Caption         =   "..."
            Height          =   255
            Left            =   2640
            TabIndex        =   104
            Top             =   2760
            Width           =   255
         End
         Begin VB.TextBox txtTag 
            Height          =   315
            Left            =   840
            TabIndex        =   103
            Top             =   3120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtDrug 
            Height          =   315
            Left            =   840
            TabIndex        =   101
            Top             =   2700
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   840
            TabIndex        =   99
            Top             =   2280
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.PictureBox picShowSendType 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   0
            MouseIcon       =   "frmPIVAMain.frx":586C4
            ScaleHeight     =   270
            ScaleWidth      =   3015
            TabIndex        =   95
            Tag             =   "0"
            Top             =   1560
            Width           =   3015
            Begin VB.PictureBox picUpOrDown1 
               BackColor       =   &H00FFEDDD&
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   2640
               Picture         =   "frmPIVAMain.frx":589CE
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   96
               Top             =   0
               Width           =   270
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "������������"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   120
               TabIndex        =   97
               Top             =   45
               Width           =   1080
            End
         End
         Begin VB.ComboBox cboʱ�䷶Χ 
            Height          =   300
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   420
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker Dtp����ʱ�� 
            Height          =   315
            Left            =   885
            TabIndex        =   3
            Top             =   1140
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   116654083
            CurrentDate     =   39998
         End
         Begin MSComCtl2.DTPicker Dtp��ʼʱ�� 
            Height          =   300
            Left            =   885
            TabIndex        =   4
            Top             =   780
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   116654083
            CurrentDate     =   39998
         End
         Begin VB.Label lbldept 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "�� ��"
            Height          =   180
            Left            =   240
            TabIndex        =   106
            Top             =   1980
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "ƿǩ��"
            Height          =   180
            Left            =   225
            TabIndex        =   102
            Top             =   3180
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblDrug 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "ҩ Ʒ"
            Height          =   180
            Left            =   240
            TabIndex        =   100
            Top             =   2760
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "������"
            Height          =   180
            Left            =   225
            TabIndex        =   98
            Top             =   2340
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "��Һ��ִ��ʱ�䷶Χ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   45
            TabIndex        =   15
            Top             =   120
            Width           =   1755
         End
         Begin VB.Label lblTimeBegin 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "��ʼʱ��"
            Height          =   180
            Left            =   45
            TabIndex        =   7
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblTimeEnd 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   45
            TabIndex        =   6
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label lblʱ�䷶Χ 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "ʱ�䷶Χ"
            Height          =   180
            Left            =   45
            TabIndex        =   5
            Top             =   480
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   11340
      Width           =   17916
      _ExtentX        =   31591
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPIVAMain.frx":58D10
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24712
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
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
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   840
      Left            =   6000
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   2280
      _cx             =   4022
      _cy             =   1482
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
      BackColorSel    =   16771280
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPIVAMain.frx":595A4
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
   Begin MSComctlLib.ImageList ImgList 
      Left            =   10080
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":59632
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":5FE94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":666F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":6CF58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":737BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":7A01C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":8087E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":870E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgPro 
      Left            =   10800
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":8D942
            Key             =   "��ȡҩ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":941A4
            Key             =   "�Ա�ҩ"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   7320
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPIVAMain.frx":9AA06
      Left            =   6480
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPIVAMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngMode As Long
Private mstrPrivs As String
Private mblnLoad As Boolean
Private mblnActive As Boolean
Private mobjCISJOB As Object  '���Ӳ������Ķ���
Private mobjPlugIn As Object    '��ҽӿڶ���
    
Private mrsDeptAdvice As ADODB.Recordset        '������Ӧ��ҽ����
Private mrsTrans As ADODB.Recordset             '��Һ����¼��������Һ�����ݣ�ҩƷ��
Private mrsDeptTrans As ADODB.Recordset         '������Ӧ����Һ����

Private mrsWorkBatch As ADODB.Recordset         '��Һ�������ĵĹ�������

Private mstr�ϴβ���ID As String                '�ϴ�ѡ��Ĳ���
Private mstr�ϴ�IDS As String                   '�ϴι�������
Public mblnParamsRefresh As Boolean
Private mstrFilter As String
Private mstrUnVisble  As String
Private mstrUnallowSetColHide As String
Private mblnFilter As Boolean

Private mstrCenterName As String
Private mlngPassPati As Long

'��Ϣ��ض������
Private WithEvents mobjMipModule As zl9ComLib.clsMipModule
Attribute mobjMipModule.VB_VarHelpID = -1
Private mdateToday As Date                      '��������
Private mrsMsg As Recordset
Private mrsSendMsg As Recordset

Private mstrLastLabel As String                 '�ϴ�ѡ���ƿǩ��
Private mintCountPack As Integer                '��ҩ֮���л����״̬�ĵ�������
Private mintBeginRow As Integer
Private mintEndRow As Integer
Private mlng��ɨ��  As Long
Private mlngδɨ�� As Long
Private mlngNum As Long


Private mfrmPIVCard As frmPIVCard
Private mfrmPrintPlan As frmPrintPlan
Private mfrmPlan As frmPlan

Private mstr���� As String
Private mstr��� As String

Private mint��־ As Integer

Private mstr��ҩid As String
Private mblnLock As Boolean

Private mrsPRI As Recordset
Private mrsVol As Recordset
Private mrstemp As Recordset
Private mrsMedi As Recordset

Private mblnShowOhters As Boolean             '�Ƿ���ʾ�Ա�ҩ�벻ȡҩ

'��Һ���б��У������޸ĵ�����ɫ
Private Const CSTCOLOR_MODIFY = &HE1FFE1        'ǳ��ɫ
'��Һ���б��У��������޸ĵ�����ɫ
Private Const CSTCOLOR_UNMODIFY = &H80000005    '��ɫ
'����״̬����Һ���м�¼ʱ��ť����ɫ
Private Const CSTCOLOR_RECORDS = &HE1FFE1       'ǳ��ɫ
'����״̬����Һ��û�м�¼ʱ��ť����ɫ
Private Const CSTCOLOR_NORECORDS = &HFFFFFF   '��ɫ
'��ǰ״̬��ť����ɫ
Private Const CSTCOLOR_COMMAND = &HFFEDDD       'ǳ��ɫ

'Ȩ��
Private Type Type_Privs
    bln�˲�ȷ�� As Boolean
    blnȡ����� As Boolean
    bln��ҩȷ�� As Boolean
    blnȡ����ҩ As Boolean
    bln��ҩȷ�� As Boolean
    blnȡ����ҩ As Boolean
    bln����ȷ�� As Boolean
    blnȡ������ As Boolean
    bln�������� As Boolean
    bln������� As Boolean
    blnȷ�Ͼܾ� As Boolean
    bln���ʾܾ� As Boolean
    bln�Ű����� As Boolean
End Type
Private mPrives As Type_Privs

'ʹ�õ��Ĳ���������ϵͳ�����������������򱾻�ע���
Private Type Type_Params
    '�������е�ϵͳ����
    lng�������� As Long
    bln����δ��˴�����ҩ As Boolean
    bln����δ�շѴ�����ҩ As Boolean
    bln����ȡ����ҩ As Boolean
    blnҽ������ As Boolean
    bln��˻��۵� As Boolean
    bln�����������۷��� As Boolean
    intҩƷ������ʾ As Integer          '0-��������ƣ�1-�����룬2-������

    '�������е���������
    int��ҩ���ӡ As Integer            '0-��ʾ��ӡ;1-�Զ���ӡ;2-����
    int���ͺ��ӡ As Integer            '0-��ʾ��ӡ;1-�Զ���ӡ;2-����
    bln�������� As Boolean
    bln������� As Boolean
    bln��� As Boolean
    intƤ����Ч���� As Integer          'Ƥ����Ч����
    int��ӡ���� As Integer            '0-��ʾ��ӡ;1-�Զ���ӡ;2-����
    blnLastBatch As Boolean             '�����ϴ�����
    bln������� As Boolean
    blnByMedi As Boolean                '��ҩƷ,��������
    blnFilter As Boolean             '�Ƿ����õĳ���ҩƷ���й���
    
'    intƿǩ�Զ���ӡ As Integer          '0-���Զ���ӡ;1-��ҩ���Զ���ӡ;2-��ҩ���Զ���ӡ
    intƿǩ��ҩ���ӡ As Integer        '0-��ʾ��ӡ;1-�Զ���ӡ;2-����
    intƿǩ��ҩ���ӡ As Integer        '0-��ʾ��ӡ;1-�Զ���ӡ;2-����
    blnƿǩ�ֹ���ӡ As Boolean
    strBatchList  As String             '���������б�
    intCount As Integer                 '��Ƭģʽ�£�������ʾ����
    intNum As Integer
    str����ҩƷ As String
    blnTwoCode As Boolean               'ɨ��һ�ν��з��ͻ�����ҩ����
    intCheck As Integer                 '��˸�ҩ������ҽ��
    blnRePeople As Boolean              '��ӡƿǩʱ�Ƿ���дʵ�ʲ���Ա
    
    'ע������
    intFont As Integer                  '��������С
    intAutoSelect As Integer            '�Զ�ѡ���ϴ�ѡ�����Һ��
    strSort As String                   '��Һ������
    strVsfTrans As String               '��ϸ����п�
    strVsfLook As String                 '�Ѱ�ҩ����п�
    strVsfSum As String                  '���ܱ���п�
    '�����
    IntCheckStock As Integer            '0-�����;1-��������;2-�����ֹ
    
    '�Ƿ���ʾ������ҩ��PASS��
    intShowPass As Integer
    
    int�������� As Integer
    intҩƷ������ʾ��ʽ As Integer      '0-��������ƣ�1-�����ƣ�2-������
    strSourceDep As String              '��ʾ��Դ����
End Type
Private mParams As Type_Params

Private Type Type_Condition
    lngCenterID As Long            '��Һ�������ĵĲ���ID
    strCenterName As String
    intTransTimeSel As Integer
    strTransStartTime As String
    strTransEndTime As String
    strTransStep As String
End Type
Private mcondition As Type_Condition

'��ϸ��ҳ
Private Enum mDetailType
    ��Һ���б� = 0
    ��Һ����Ƭ
    ҩƷ�����б�
End Enum

'ҵ��/�鿴��ҳ
Private Const CNUMWORK = 0
Private Const CNUMLOOK = 1

'ҵ�����/����
Private Const M_STR_CALSS_AUDIT = "00"              '���ҽ��
Private Const M_STR_CALSS_PREPARE = "01"            '��ҩӡǩ
Private Const M_STR_CALSS_DOSAGE = "02"             '��ҩ�˲�
Private Const M_STR_CALSS_SEND = "03"               '���ͺ˲�
Private Const M_STR_CALSS_VERIFY = "04"             '�������
Private Const M_STR_CALSS_PASSEDAUDIT = "10"        '�����ͨ��ҽ��
Private Const M_STR_CALSS_FAILAUDIT = "11"          '���δͨ��ҽ��
Private Const M_STR_CALSS_SENDED = "12"             '�ѷ��Ͳ鿴
Private Const M_STR_CALSS_SIGNED = "13"             '��ǩ�ղ鿴
Private Const M_STR_CALSS_REFUSETOSIGN = "14"       '�ܾ�ǩ�ղ鿴
Private Const M_STR_CALSS_INVALID = "15"            '�����ϲ鿴
Private Const M_STR_CALSS_DEVICERETURN = "16"       'ҽ�����˲鿴

Private Enum mTransStatus
    ���� = 1
    ��ҩ = 2
    У�� = 3
    ��ҩ = 4
    ���� = 5
    ǩ�� = 6
    �ܾ�ǩ�� = 7
    ȷ�Ͼ��� = 8
    �������� = 9
    �������ͨ�� = 10
    �������δͨ�� = 11
End Enum


'ҵ����ܱ��
Private Const MINTSUMCOLS = 13      '������
Private mintcolsum���� As Integer
Private mintcolsum��� As Integer
Private mintcolsumҩƷ���� As Integer
Private mintcolsum��Ʒ�� As Integer
Private mintcolsumӢ���� As Integer
Private mintcolsum��� As Integer
Private mintcolsum���� As Integer
Private mintcolsum���� As Integer
Private mintcolsum���� As Integer
Private mintcolsum��ҩ���� As Integer
Private mintcolsum������� As Integer
Private mintcolsumȱҩ��־ As Integer
Private mintcolsum�Ƿ��� As Integer


Private Const MINTCOLS = 63      '������
'Private mIntColƿǩ�� As Integer
'Private mintcol���� As Integer
'Private mIntCol��ҩ�� As Integer
'Private mIntCol��ҩʱ�� As Integer
'Private mIntCol��ҩ���� As Integer
'Private mIntColҽ������ʱ�� As Integer
'Private mIntColִ��ʱ�� As Integer
'Private mIntColҩƷ���� As Integer
'Private mintcol��� As Integer
'Private mIntCol���� As Integer
'Private mintcol���� As Integer
'Private mIntColNO As Integer
'Private mIntCol���� As Integer
'Private mIntCol������λ As Integer
'Private mIntCol�÷� As Integer
'Private mintcolҩƷid As Integer
'Private mIntCol��ҩid As Integer
Private mIntCol��ǰ�� As Integer
Private mIntCol�� As Integer
Private mintcolѡ�� As Integer
Private mIntCol�� As Integer
Private mIntCol�� As Integer
Private mIntColҽ��     As Integer
Private mIntCol�� As Integer
Private mIntCol��ӡ As Integer
Private mIntCol��� As Integer
Private mintcol���� As Integer
Private mIntCol����ԭ�� As Integer
Private mIntCol���ȼ� As Integer
Private mIntCol���� As Integer
Private mIntCol���� As Integer
Private mIntCol���� As Integer
Private mIntCol�Ա� As Integer
Private mIntCol���� As Integer
Private mIntCol���� As Integer
Private mIntColסԺ�� As Integer
Private mIntCol�� As Integer
Private mIntColҩƷ���� As Integer
Private mIntColƤ As Integer
Private mintcol��� As Integer
Private mIntCol��ҩ���� As Integer
Private mIntCol���� As Integer
Private mintcol���� As Integer
Private mIntColִ��ʱ�� As Integer
Private mIntColִ��Ƶ�� As Integer
Private mIntColƿǩ�� As Integer
Private mIntCol��ҩ���� As Integer
Private mIntColҽ������ʱ�� As Integer
Private mIntCol��ҩ�� As Integer
Private mIntCol��ҩʱ�� As Integer
Private mIntCol��ҩ�� As Integer
Private mIntCol��ҩʱ�� As Integer
Private mIntCol������ As Integer
Private mIntCol����ʱ�� As Integer
Private mIntCol���������� As Integer
Private mIntCol��������ʱ�� As Integer
Private mIntCol��������� As Integer
Private mIntCol�������ʱ�� As Integer
Private mIntCol����ԭ�� As Integer
Private mIntCol����״̬ As Integer

Private mIntCol��־ As Integer
Private mIntCol�Ƿ����� As Integer
Private mIntCol�������� As Integer
Private mIntColNO As Integer
Private mIntCol���� As Integer
Private mIntCol������λ As Integer
Private mIntCol�÷� As Integer
Private mintcolҩƷid As Integer
Private mIntCol�˲��� As Integer
Private mIntCol�˲�ʱ�� As Integer
Private mIntCol��ӡ��־ As Integer
Private mIntCol��ҩid As Integer
Private mIntCol�Ƿ��� As Integer
Private mIntColԭ���� As Integer
Private mIntCol����ҩ�� As Integer
Private mIntCol��ҳid As Integer
Private mIntCol����ID As Integer
Private mIntCol������ As Integer
Private mIntCol���� As Integer
Private mIntCol��ý As Integer
Private mIntCol��Ӧҽ��ID As Integer

'Private Enum mTransStep
'    ���ҽ�� = 0
'    ��ҩӡǩ
'    ��ҩ�˲�
'    ���ͺ˲�
'    �������
'End Enum
'
'Private Enum mTransLook
'    �����ͨ��ҽ�� = 0
'    ���δͨ��ҽ��
'    �ѷ��Ͳ鿴
'    ��ǩ�ղ鿴
'    �ܾ�ǩ�ղ鿴
'    �����ϲ鿴
'End Enum

'�����˵�
Private Const conMenu_OperPopup = 300                   '����

Private Const conMenu_Oper_PrintLabel = 301
Private Const conMenu_Oper_PrintLabel_SelRow = 302    '��ӡ��ǩ����ǰѡ���У�
Private Const conMenu_Oper_PrintLabel_SelBatch = 303    '��ӡ��ǩ����ǰѡ�����Σ�
Private Const conMenu_Oper_PrintLabel_SelDept = 304    '��ӡ��ǩ����ǰѡ�в�����
Private Const conMenu_Oper_PrintLabel_SelPati = 305    '��ӡ��ǩ����ǰѡ�в��ˣ�
Private Const conMenu_Oper_PrintLabel_AllRow = 306    '��ӡ��ǩ������ѡ����У�
Private Const conMenu_Oper_PrintLabel_SelSendNo = 307    '��ӡ��ǩ����ǰѡ�еİ�ҩ���ţ�

Private Const conMenu_Oper_DelBatch = 311
Private Const conMenu_Oper_DelBatch_SelRow = 312       'ɾ�����Σ���ǰѡ���У�
Private Const conMenu_Oper_DelBatch_SelBatch = 313   'ɾ�����Σ���ǰѡ�����Σ�
Private Const conMenu_Oper_DelBatch_SelDept = 314    'ɾ�����Σ���ǰѡ�в�����
Private Const conMenu_Oper_DelBatch_SelPati = 315    'ɾ�����Σ���ǰѡ�в��ˣ�
Private Const conMenu_Oper_DelBatch_AllRow = 316    'ɾ�����Σ�����ѡ����У�

Private Const conMenu_Oper_Select = 321
Private Const conMenu_Oper_Select_SelRow = 322       'ѡ�񣨵�ǰѡ���У�
Private Const conMenu_Oper_Select_SelBatch = 323   'ѡ�񣨵�ǰѡ�����Σ�
Private Const conMenu_Oper_Select_SelDept = 324    'ѡ�񣨵�ǰѡ�в�����
Private Const conMenu_Oper_Select_SelPati = 325    'ѡ�񣨵�ǰѡ�в��ˣ�
Private Const conMenu_Oper_Select_SelAll = 326    'ѡ�������У�
Private Const conMenu_Oper_Select_SelSendNo = 327    'ѡ�񣨵�ǰѡ�еİ�ҩ���ţ�
Private Const conMenu_Oper_Select_SelMed = 328    'ѡ�񣨵�ǰ���еĿ���ҩ�
Private Const conMenu_Oper_Select_CancleSelDept = 329    'ȡ��ѡ�񣨵�ǰѡ�в�����
Private Const conMenu_Oper_Select_CancleSelPati = 330    'ȡ��ѡ�񣨵�ǰѡ�в��ˣ�

Private Const conMenu_Oper_Bag = 331
Private Const conMenu_Oper_Bag_Batch = 332   '����������ǰ���Σ�
Private Const conMenu_Oper_Bag_All = 333   'ȫ������������ǰ���Σ�

Private Const conMenu_Oper_Look = 341        '���Ӳ�������

Private Const mconMenu_SortPopup = 6000                  '����ʽ
Private Const mconMenu_SortPopup_ByCode = 6001           '������
Private Const mconMenu_SortPopup_ByName = 6002           '������
          
'ҽ���ӿ�
Private gclsInsure As New clsInsure

Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
Private Function CheckPriceAdjustByID() As Boolean
    '�����շ�ID�����ҩƷ����
    Dim rstemp As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    Dim strDrugList As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    '���û����ȫ�ֵ����۹����򲻽��к�����飬����true
    If Val(zlDatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjustByID = True: Exit Function
    
    If mrsTrans Is Nothing Then
        MsgBox "��ȡ�����쳣��������ˢ�����ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Filter = "ִ�б�־=1"
    
    If mrsTrans.RecordCount = 0 Then
        MsgBox "��ȡ�����쳣��������ˢ�����ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Sort = "�շ�ID"
    
    Set rstemp = mrsTrans
    If mrsTrans.RecordCount = 0 Then
        MsgBox "��ȡ�����쳣��������ˢ�����ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Do While Not rstemp.EOF
        If i >= 5 Then Exit Do
        gstrSQL = "Select a.ҩƷid, Nvl(a.����, 0) As ����," & vbNewLine & _
            "       '[' || c.���� || ']' || c.���� || Decode(c.����, Null, Null, '(' || c.���� || ')') || c.��� As ͨ����" & vbNewLine & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C" & vbNewLine & _
            " Where a.ҩƷid = b.ҩƷid And b.ҩƷid = c.Id And b.�Ƿ����۹��� = 1 And a.Id = [1] "
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjustByID", Val(rstemp!�շ�ID))
        
        If Not rsData.EOF Then
            If InStr(1, "," & strDrugList & ",", "," & rsData!ҩƷID & ",") = 0 Then
                strDrugList = IIf(strDrugList = "", "", strDrugList & ",") & rsData!ҩƷID
                If CheckPriceAdjust(rsData!ҩƷID, mcondition.lngCenterID, rsData!����) = False Then
                    i = i + 1
                    strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & rsData!ͨ����
                End If
            End If
        End If
        
        rstemp.MoveNext
    Loop
    
    If strMsg = "" Then
        CheckPriceAdjustByID = True
        Exit Function
    Else
        MsgBox "����ҩƷ���������۹�����������ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡" & vbCrLf & strMsg, vbInformation, gstrSysName
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitVSFLook()
    Dim arr������ As Variant
    Dim n As Integer
    Dim i As Integer
    
    mIntCol����״̬ = 0
    mIntColƿǩ�� = 1
    mintcol���� = 2
    mIntCol��� = 3
    mIntCol��ҩ�� = 4
    mIntCol��ҩʱ�� = 5
    mIntCol��ҩ���� = 6
    mIntColҽ������ʱ�� = 7
    mIntColִ��ʱ�� = 8
    mIntColҩƷ���� = 9
    mintcol��� = 10
    mIntCol���� = 11
    mintcol���� = 12
    mIntColNO = 13
    mIntCol���� = 14
    mIntCol������λ = 15
    mIntCol�÷� = 16
    mintcolҩƷid = 17
    mIntCol��ҩid = 18
    
    '�ָ��û��Զ�����˳��
    If mParams.strVsfLook <> "" Then
        arr������ = Split(mParams.strVsfLook, "|")
        
        For n = 0 To UBound(arr������)
            SetVsfLookValue Split(arr������(n), ",")(0), n
        Next
    End If
    
    With VSFLook
        .rows = 1
        .rows = 2
'        .Cols = 17
        
        VsfGridColFormat VSFLook, mIntCol����״̬, "����״̬", 1200, flexAlignLeftCenter, "����״̬"
        VsfGridColFormat VSFLook, mIntColƿǩ��, "ƿǩ��", 2000, flexAlignRightCenter, "ƿǩ��"
        VsfGridColFormat VSFLook, mintcol����, "����", 1000, flexAlignLeftCenter, "����"
        VsfGridColFormat VSFLook, mIntCol���, "���", 1000, flexAlignLeftCenter, "���"
        VsfGridColFormat VSFLook, mIntCol��ҩ��, "��ҩ��", 1200, flexAlignLeftCenter, "��ҩ��"
        VsfGridColFormat VSFLook, mIntCol��ҩʱ��, "��ҩʱ��", 2000, flexAlignLeftCenter, "��ҩʱ��"
        VsfGridColFormat VSFLook, mIntCol��ҩ����, "��ҩ����", 2000, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat VSFLook, mIntColҽ������ʱ��, "ҽ������ʱ��", 2000, flexAlignLeftCenter, "ҽ������ʱ��"
        VsfGridColFormat VSFLook, mIntColִ��ʱ��, "ִ��ʱ��", 1800, flexAlignLeftCenter, "ִ��ʱ��"
        VsfGridColFormat VSFLook, mIntColҩƷ����, "ҩƷ����", 1800, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat VSFLook, mintcol���, "���", 1800, flexAlignLeftCenter, "���"
        VsfGridColFormat VSFLook, mIntCol����, "����", 1800, flexAlignLeftCenter, "����"
        VsfGridColFormat VSFLook, mintcol����, "����", 1800, flexAlignLeftCenter, "����"
        VsfGridColFormat VSFLook, mIntColNO, "NO", 1800, flexAlignLeftCenter, "NO"
        VsfGridColFormat VSFLook, mIntCol����, "����", 1800, flexAlignLeftCenter, "����"
        VsfGridColFormat VSFLook, mIntCol������λ, "������λ", 1800, flexAlignLeftCenter, "������λ"
        VsfGridColFormat VSFLook, mIntCol�÷�, "�÷�", 1800, flexAlignLeftCenter, "�÷�"
        VsfGridColFormat VSFLook, mintcolҩƷid, "ҩƷid", 1800, flexAlignLeftCenter, "ҩƷid"
        VsfGridColFormat VSFLook, mIntCol��ҩid, "��ҩid", 1800, flexAlignLeftCenter, "��ҩid"
    End With
    
    '�ָ��п�
    If mParams.strVsfLook <> "" Then
        arr������ = Split(mParams.strVsfLook, "|")
        For n = 0 To UBound(arr������)
            For i = 0 To VSFLook.Cols - 1
                If Split(arr������(n), ",")(0) = VSFLook.ColKey(i) Then
                    VSFLook.ColWidth(i) = Val(Split(arr������(n), ",")(1))
                End If
            Next
        Next
    End If
End Sub

Private Sub SetVsfLookValue(ByVal str���� As String, ByVal intValue As Integer)
    Select Case str����
        Case "ƿǩ��"
            mIntColƿǩ�� = intValue
        Case "����"
            mintcol���� = intValue
        Case "���"
            mIntCol��� = intValue
        Case "��ҩ��"
            mIntCol��ҩ�� = intValue
        Case "��ҩʱ��"
            mIntCol��ҩʱ�� = intValue
        Case "��ҩ����"
            mIntCol��ҩ���� = intValue
        Case "ҽ������ʱ��"
            mIntColҽ������ʱ�� = intValue
        Case "ִ��ʱ��"
            mIntColִ��ʱ�� = intValue
        Case "ҩƷ����"
            mIntColҩƷ���� = intValue
        Case "����"
            mintcol���� = intValue
        Case "���"
            mintcol��� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "NO"
            mIntColNO = intValue
        Case "����"
            mIntCol���� = intValue
        Case "������λ"
            mIntCol������λ = intValue
        Case "�÷�"
            mIntCol�÷� = intValue
        Case "�÷�"
            mintcolҩƷid = intValue
        Case "��ҩid"
            mIntCol��ҩid = intValue
    End Select
End Sub
Private Sub InitVsfSum()
    Dim arr������ As Variant
    Dim n As Integer
    Dim i As Integer
    
    mintcolsum���� = 0
    mintcolsum��� = 1
    mintcolsumҩƷ���� = 2
    mintcolsum��Ʒ�� = 3
    mintcolsumӢ���� = 4
    mintcolsum��� = 5
    mintcolsum���� = 6
    mintcolsum���� = 7
    mintcolsum���� = 8
    mintcolsum��ҩ���� = 9
    mintcolsum������� = 10
    mintcolsumȱҩ��־ = 11
    mintcolsum�Ƿ��� = 12
    
    '�ָ��û��Զ�����˳��
    If mParams.strVsfSum <> "" Then
        arr������ = Split(mParams.strVsfSum, "|")
        
        For n = 0 To UBound(arr������)
            SetColumnValue Split(arr������(n), ",")(0), n
        Next
    End If
    
    With vsfSumDrug
        .rows = 1
        .rows = 2
        .Cols = MINTSUMCOLS

        VsfGridColFormat vsfSumDrug, mintcolsum����, "����", 450, flexAlignRightCenter, "����"
        VsfGridColFormat vsfSumDrug, mintcolsum���, "���", 2500, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfSumDrug, mintcolsumҩƷ����, "ҩƷ����", 400, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfSumDrug, mintcolsum��Ʒ��, "��Ʒ��", 2000, flexAlignLeftCenter, "��Ʒ��"
        VsfGridColFormat vsfSumDrug, mintcolsumӢ����, "Ӣ����", 2000, flexAlignLeftCenter, "Ӣ����"
        VsfGridColFormat vsfSumDrug, mintcolsum���, "���", 1800, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfSumDrug, mintcolsum����, "����", 1800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfSumDrug, mintcolsum����, "����", 1800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfSumDrug, mintcolsum����, "����", 1800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfSumDrug, mintcolsum��ҩ����, "��ҩ����", 1800, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfSumDrug, mintcolsum�������, "�������", 1800, flexAlignLeftCenter, "�������"
        VsfGridColFormat vsfSumDrug, mintcolsumȱҩ��־, "ȱҩ��־", 1800, flexAlignLeftCenter, "ȱҩ��־"
        VsfGridColFormat vsfSumDrug, mintcolsum�Ƿ���, "�Ƿ���", 1800, flexAlignLeftCenter, "�Ƿ���"
    End With
    
    '�ָ��п�
    If mParams.strVsfSum <> "" Then
        arr������ = Split(mParams.strVsfSum, "|")
        For n = 0 To UBound(arr������)
            For i = 0 To vsfSumDrug.Cols - 1
                If Split(arr������(n), ",")(0) = vsfSumDrug.ColKey(i) Then
                    vsfSumDrug.ColWidth(i) = Val(Split(arr������(n), ",")(1))
                End If
            Next
        Next
    End If
End Sub

Private Sub SetColumnValue(ByVal str���� As String, ByVal intValue As Integer)
    Select Case str����
        Case "����"
            mintcolsum���� = intValue
        Case "���"
            mintcolsum��� = intValue
        Case "ҩƷ����"
            mintcolsumҩƷ���� = intValue
        Case "Ӣ����"
            mintcolsumӢ���� = intValue
        Case "��Ʒ��"
            mintcolsum��Ʒ�� = intValue
        Case "���"
            mintcolsum��� = intValue
        Case "����"
            mintcolsum���� = intValue
        Case "����"
            mintcolsum���� = intValue
        Case "����"
            mintcolsum���� = intValue
        Case "��ҩ����"
            mintcolsum��ҩ���� = intValue
        Case "�������"
            mintcolsum������� = intValue
        Case "ȱҩ��־"
            mintcolsumȱҩ��־ = intValue
        Case "�Ƿ���"
            mintcolsum�Ƿ��� = intValue
    End Select
                   
End Sub

Private Sub SetTransColumnValue(ByVal str���� As String, ByVal intValue As Integer)
    Select Case str����
        Case "��ǰ��"
            mIntCol��ǰ�� = intValue
        Case "��"
            mIntCol�� = intValue
        Case "ѡ��"
            mintcolѡ�� = intValue
        Case "��"
            mIntCol�� = intValue
        Case "��"
            mIntCol�� = intValue
        Case "ҽ��"
            mIntColҽ�� = intValue
        Case "��"
            mIntCol�� = intValue
        Case "��ӡ"
            mIntCol��ӡ = intValue
        Case "���"
            mIntCol��� = intValue
        Case "��ҩ����"
            mintcol���� = intValue
        Case "����ԭ��"
            mIntCol����ԭ�� = intValue
        Case "���ȼ�"
            mIntCol���ȼ� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "�Ա�"
            mIntCol�Ա� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "סԺ��"
            mIntColסԺ�� = intValue
        Case "�����"
            mIntCol�� = intValue
        Case "ҩƷ����"
            mIntColҩƷ���� = intValue
        Case "Ƥ"
            mIntColƤ = intValue
        Case "���"
            mintcol��� = intValue
        Case "��ҩ����"
            mIntCol��ҩ���� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "����"
            mintcol���� = intValue
        Case "ִ��ʱ��"
            mIntColִ��ʱ�� = intValue
        Case "ƿǩ��"
            mIntColƿǩ�� = intValue
        Case "��ҩ����"
            mIntCol��ҩ���� = intValue
        Case "ҽ������ʱ��"
            mIntColҽ������ʱ�� = intValue
        Case "��ҩ��"
            mIntCol��ҩ�� = intValue
        Case "��ҩʱ��"
            mIntCol��ҩʱ�� = intValue
        Case "��ҩ��"
            mIntCol��ҩ�� = intValue
        Case "��ҩʱ��"
            mIntCol��ҩʱ�� = intValue
        Case "������"
            mIntCol������ = intValue
        Case "����ʱ��"
            mIntCol����ʱ�� = intValue
        Case "����������"
            mIntCol���������� = intValue
        Case "��������ʱ��"
            mIntCol��������ʱ�� = intValue
        Case "���������"
            mIntCol��������� = intValue
        Case "�������ʱ��"
            mIntCol�������ʱ�� = intValue
        Case "��־"
            mIntCol��־ = intValue
        Case "��������"
            mIntCol�������� = intValue
        Case "NO"
            mIntColNO = intValue
        Case "������λ"
            mIntCol������λ = intValue
        Case "�÷�"
            mIntCol�÷� = intValue
        Case "ҩƷid"
            mintcolҩƷid = intValue
        Case "�˲���"
            mIntCol�˲��� = intValue
        Case "�˲�ʱ��"
            mIntCol�˲�ʱ�� = intValue
        Case "��ӡ��־"
            mIntCol��ӡ��־ = intValue
        Case "��ҩid"
            mIntCol��ҩid = intValue
        Case "����ҩ��"
            mIntCol����ҩ�� = intValue
        Case "ԭ����"
            mIntColԭ���� = intValue
        Case "��ҳid"
            mIntCol��ҳid = intValue
        Case "����ID"
            mIntCol����ID = intValue
        Case "������"
            mIntCol������ = intValue
        Case "�Ƿ�����"
            mIntCol�Ƿ����� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "�Ƿ���"
            mIntCol�Ƿ��� = intValue
        Case "����ԭ��"
            mIntCol����ԭ�� = intValue
    End Select
End Sub

Private Sub InitVsfTrans()
    Dim arr������ As Variant
    Dim n As Integer
    Dim i As Integer
    Dim strRows As String
    
    '��ʼ����
    mIntCol��ǰ�� = 0
    mIntCol�� = 1
    mintcolѡ�� = 2
    mIntCol�������� = 3
    mIntCol�� = 4
    mIntCol�� = 5
    mIntColҽ�� = 6
    mIntCol�� = 7
    mIntCol��ӡ = 8
    mIntCol��� = 9
    mintcol���� = 10
    mIntCol����ԭ�� = 11
    mIntCol���ȼ� = 12
    mIntCol���� = 13
    mIntCol���� = 14
    mIntCol���� = 15
    mIntCol�Ա� = 16
    mIntCol���� = 17
    mIntCol���� = 18
    mIntColסԺ�� = 19
    mIntCol�� = 20
    mIntColҩƷ���� = 21
    mIntColƤ = 22
    mintcol��� = 23
    mIntCol��ҩ���� = 24
    mIntCol���� = 25
    mintcol���� = 26
    mIntColִ��ʱ�� = 27
    mIntColִ��Ƶ�� = 28
    mIntColƿǩ�� = 29
    mIntCol��ҩ���� = 30
    mIntColҽ������ʱ�� = 31
    mIntCol��ҩ�� = 32
    mIntCol��ҩʱ�� = 33
    mIntCol��ҩ�� = 34
    mIntCol��ҩʱ�� = 35
    mIntCol������ = 36
    mIntCol����ʱ�� = 37
    mIntCol���������� = 38
    mIntCol��������ʱ�� = 39
    mIntCol��������� = 40
    mIntCol�������ʱ�� = 41
    mIntCol����ԭ�� = 42
    mIntCol��־ = 43
    mIntCol�Ƿ����� = 44
    mIntColNO = 45
    mIntCol���� = 46
    mIntCol������λ = 47
    mIntCol�÷� = 48
    mintcolҩƷid = 49
    mIntCol�˲��� = 50
    mIntCol�˲�ʱ�� = 51
    mIntCol��ӡ��־ = 52
    mIntCol��ҩid = 53
    mIntCol�Ƿ��� = 54
    mIntColԭ���� = 55
    mIntCol����ҩ�� = 56
    mIntCol��ҳid = 57
    mIntCol����ID = 58
    mIntCol������ = 59
    mIntCol���� = 60
    mIntCol��ý = 61
    mIntCol��Ӧҽ��ID = 62
    
    '���û���ǰû�б���"��ҩ����"��,����г�ʼ��
    If InStr(mParams.strVsfTrans, "��ҩ����") = 0 Then mParams.strVsfTrans = ""
    
    '�ָ��û��Զ�����˳��
    If mParams.strVsfTrans <> "" Then
        arr������ = Split(mParams.strVsfTrans, "|")
        
        For n = 0 To UBound(arr������)
            SetTransColumnValue Split(arr������(n), ",")(0), n
        Next
    End If
    
    With vsfTrans
        .Cols = MINTCOLS
        
        VsfGridColFormat vsfTrans, mIntCol��ǰ��, " ", 200, flexAlignRightCenter, "��ǰ��"
        VsfGridColFormat vsfTrans, mIntCol��, "��", 400, flexAlignRightCenter, "��"
        VsfGridColFormat vsfTrans, mintcolѡ��, "ѡ��", 400, flexAlignLeftCenter, "ѡ��"
        VsfGridColFormat vsfTrans, mIntCol��, "��", 400, flexAlignLeftCenter, "��"
        VsfGridColFormat vsfTrans, mIntCol��, "��", 400, flexAlignLeftCenter, "��"
        VsfGridColFormat vsfTrans, mIntColҽ��, "ҽ��", 400, flexAlignLeftCenter, "ҽ��"
        VsfGridColFormat vsfTrans, mIntCol��, "��", 400, flexAlignLeftCenter, "��"
        VsfGridColFormat vsfTrans, mIntCol��ӡ, "��ӡ", 400, flexAlignLeftCenter, "��ӡ"
        VsfGridColFormat vsfTrans, mIntCol���, "���", 400, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfTrans, mintcol����, "����", 400, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfTrans, mIntCol����ԭ��, "����ԭ��", 0, flexAlignLeftCenter, "����ԭ��"
        VsfGridColFormat vsfTrans, mIntCol���ȼ�, "���ȼ�", 600, flexAlignLeftCenter, "���ȼ�"
        VsfGridColFormat vsfTrans, mIntCol����, "����", 1800, flexAlignLeftCenter, "����"

        VsfGridColFormat vsfTrans, mIntCol����, "����", 1800, flexAlignRightCenter, "����"
        VsfGridColFormat vsfTrans, mIntCol����, "����", 1800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfTrans, mIntCol�Ա�, "�Ա�", 400, flexAlignLeftCenter, "�Ա�"
        VsfGridColFormat vsfTrans, mIntCol����, "����", 800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfTrans, mIntCol����, "����", 800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfTrans, mIntColסԺ��, "סԺ��", 800, flexAlignLeftCenter, "סԺ��"
        VsfGridColFormat vsfTrans, mIntCol��, "��", 400, flexAlignLeftCenter, "�����"
        VsfGridColFormat vsfTrans, mIntColҩƷ����, "ҩƷ����", 1800, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfTrans, mIntColƤ, "Ƥ", 600, flexAlignLeftCenter, "Ƥ"
        VsfGridColFormat vsfTrans, mintcol���, "���", 1800, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfTrans, mIntCol��ҩ����, "��ҩ����", 1800, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfTrans, mIntCol����, "����", 1800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfTrans, mintcol����, "����", 1800, flexAlignLeftCenter, "����"


        VsfGridColFormat vsfTrans, mIntColִ��ʱ��, "ִ��ʱ��", 2000, flexAlignRightCenter, "ִ��ʱ��"
        VsfGridColFormat vsfTrans, mIntColִ��Ƶ��, "ִ��Ƶ��", 1200, flexAlignRightCenter, "ִ��Ƶ��"
        VsfGridColFormat vsfTrans, mIntColƿǩ��, "ƿǩ��", 2500, flexAlignLeftCenter, "ƿǩ��"
        VsfGridColFormat vsfTrans, mIntCol��ҩ����, "��ҩ����", 2000, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfTrans, mIntColҽ������ʱ��, "ҽ������ʱ��", 2000, flexAlignLeftCenter, "ҽ������ʱ��"
        VsfGridColFormat vsfTrans, mIntCol��ҩ��, "��ҩ��", 2000, flexAlignLeftCenter, "��ҩ��"
        VsfGridColFormat vsfTrans, mIntCol��ҩʱ��, "��ҩʱ��", 1800, flexAlignLeftCenter, "��ҩʱ��"
        VsfGridColFormat vsfTrans, mIntCol��ҩ��, "��ҩ��", 1800, flexAlignLeftCenter, "��ҩ��"
        VsfGridColFormat vsfTrans, mIntCol��ҩʱ��, "��ҩʱ��", 1800, flexAlignLeftCenter, "��ҩʱ��"
        VsfGridColFormat vsfTrans, mIntCol������, "������", 1800, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfTrans, mIntCol����ʱ��, "����ʱ��", 1800, flexAlignLeftCenter, "����ʱ��"
        VsfGridColFormat vsfTrans, mIntCol����������, "����������", 1800, flexAlignLeftCenter, "����������"
        VsfGridColFormat vsfTrans, mIntCol��������ʱ��, "��������ʱ��", 1800, flexAlignLeftCenter, "��������ʱ��"

        VsfGridColFormat vsfTrans, mIntCol���������, "���������", 450, flexAlignRightCenter, "���������"
        VsfGridColFormat vsfTrans, mIntCol�������ʱ��, "�������ʱ��", 2500, flexAlignLeftCenter, "�������ʱ��"
        VsfGridColFormat vsfTrans, mIntCol����ԭ��, "����ԭ��", 2500, flexAlignLeftCenter, "����ԭ��"
        VsfGridColFormat vsfTrans, mIntCol��־, "��־", 400, flexAlignLeftCenter, "��־"
        VsfGridColFormat vsfTrans, mIntCol�Ƿ�����, "�Ƿ�����", 0, flexAlignLeftCenter, "�Ƿ�����"
        VsfGridColFormat vsfTrans, mIntCol��������, "��������", 2000, flexAlignLeftCenter, "��������"
        VsfGridColFormat vsfTrans, mIntColNO, "NO", 1800, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfTrans, mIntCol����, "����", 1800, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfTrans, mIntCol������λ, "������λ", 1800, flexAlignLeftCenter, "������λ"
        VsfGridColFormat vsfTrans, mIntCol�÷�, "�÷�", 1800, flexAlignLeftCenter, "�÷�"
        VsfGridColFormat vsfTrans, mintcolҩƷid, "ҩƷid", 1800, flexAlignLeftCenter, "ҩƷid"
        VsfGridColFormat vsfTrans, mIntCol�˲���, "�˲���", 1800, flexAlignLeftCenter, "�˲���"
        VsfGridColFormat vsfTrans, mIntCol�˲�ʱ��, "�˲�ʱ��", 1800, flexAlignLeftCenter, "�˲�ʱ��"
        VsfGridColFormat vsfTrans, mIntCol��ӡ��־, "��ӡ��־", 0, flexAlignLeftCenter, "��ӡ��־"

        VsfGridColFormat vsfTrans, mIntCol��ҩid, "��ҩid", 0, flexAlignLeftCenter, "��ҩid"
        VsfGridColFormat vsfTrans, mIntCol�Ƿ���, "�Ƿ���", 0, flexAlignLeftCenter, "�Ƿ���"
        VsfGridColFormat vsfTrans, mIntColԭ����, "ԭ����", 0, flexAlignLeftCenter, "ԭ����"
        VsfGridColFormat vsfTrans, mIntCol����ҩ��, "����ҩ��", 0, flexAlignLeftCenter, "����ҩ��"
        VsfGridColFormat vsfTrans, mIntCol��ҳid, "��ҳid", 0, flexAlignLeftCenter, "��ҳid"
        VsfGridColFormat vsfTrans, mIntCol����ID, "����ID", 0, flexAlignLeftCenter, "����ID"
        VsfGridColFormat vsfTrans, mIntCol������, "������", 0, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfTrans, mIntCol����, "����", 0, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfTrans, mIntCol��ý, "��ý", 0, flexAlignLeftCenter, "��ý"
        VsfGridColFormat vsfTrans, mIntCol��Ӧҽ��ID, "��Ӧҽ��ID", 0, flexAlignLeftCenter, "��Ӧҽ��ID"
    End With

    '�ָ���������
    If mParams.strVsfTrans <> "" Then
        arr������ = Split(mParams.strVsfTrans, "|")
        For n = 0 To UBound(arr������)
            For i = 0 To vsfTrans.Cols - 1
                If Split(arr������(n), ",")(0) = vsfTrans.ColKey(i) Then
                    vsfTrans.ColWidth(i) = Val(Split(arr������(n), ",")(1))
                End If
            Next
        Next
    End If
    
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        strRows = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���", mcondition.strTransStep, "")
    End If
    
    If strRows = "" Then
        strRows = mstrUnVisble & "��;��ҩ����;��������;��ҩ��;��ҩʱ��;��ҩ��;��ҩʱ��;������;����ʱ��;����������;��������ʱ��;���������;�������ʱ��;����ԭ��;"
    End If
    
    
    If strRows <> "" Then
        For n = 1 To Me.vsfTrans.Cols - 1
            If InStr(1, ";" & strRows & ";", ";" & vsfTrans.ColKey(n) & ";") > 0 Then
                vsfTrans.ColHidden(n) = True
            Else
                vsfTrans.ColHidden(n) = False
            End If
        Next
    End If
    
    '��ʼ����Ϊ��ҩ����
    Call InitColSelList(mstrUnVisble & "��;��ҩ����;��������;��ҩ��;��ҩʱ��;��ҩ��;��ҩʱ��;������;����ʱ��;����������;��������ʱ��;���������;�������ʱ��;����ԭ��;")
End Sub




Private Sub SetSortFlag(Optional ByVal blnSpecial As Boolean = False)
    '���������־
    Dim intCol, intSortCount As Integer
    
    With vsfTrans
        .Redraw = flexRDNone
        
        'ȡ�������־
        For intCol = 0 To .Cols - 1
            If InStr(1, .TextMatrix(0, intCol), "��") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "��", "")
            If InStr(1, .TextMatrix(0, intCol), "��") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "��", "")
            If InStr(1, .TextMatrix(0, intCol), "��") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "��", "")
            If InStr(1, .TextMatrix(0, intCol), "��") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "��", "")
            If InStr(1, .TextMatrix(0, intCol), "��") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "��", "")
        Next
        
        '���������־
        If blnSpecial = True Then
            '����İ�ҩƷ����(����ҩƷ����ý������ҩƷ����)
            .TextMatrix(0, .ColIndex("ҩƷ����")) = .TextMatrix(0, .ColIndex("ҩƷ����")) & "��"
        ElseIf mParams.strSort <> "" Then
            For intCol = 0 To .Cols - 1
                For intSortCount = 0 To UBound(Split(mParams.strSort, ","))
                    If .ColKey(intCol) = IIf(Split(mParams.strSort, ",")(intSortCount) = "���򴲺�", "����", Split(mParams.strSort, ",")(intSortCount)) Then
                        Select Case intSortCount + 1
                            Case 1
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "��"
                            Case 2
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "��"
                            Case 3
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "��"
                            Case 4
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "��"
                            Case 5
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "��"
                        End Select
                        
                        Exit For
                    End If
                Next
            Next
        End If
        
        .Redraw = flexRDDirect
    End With
End Sub
Private Sub GetCount()
    '����ͳ����Ϣ����ѡ����������ѡ��Һ����
    '��״̬����ʾ
    Dim lngCount As Long
    Dim lngRow As Long
    Dim lng���ID As Long
    Dim lngVolume As Long
    
    stbThis.Panels(2).Text = ""
    
    With vsfDept(Me.tabDeptList.Selected.index)
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("����ID")) <> "" Then
                If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                    lngCount = lngCount + 1
                End If
            End If
        Next
    End With
    
    If lngCount = 0 Then Exit Sub
    stbThis.Panels(2).Text = "��ǰѡ������" & lngCount
    
    lngCount = 0
    lngVolume = 0
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("��ҩID")) <> "" Then
                If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                    If Val(.TextMatrix(lngRow, .ColIndex("��ý"))) = 1 Then
                        lngVolume = lngVolume + Val(.TextMatrix(lngRow, .ColIndex("����")))
                    End If
                    
                    If lng���ID <> Val(.TextMatrix(lngRow, .ColIndex("��ҩid"))) Then
                        lng���ID = Val(.TextMatrix(lngRow, .ColIndex("��ҩid")))
                        lngCount = lngCount + 1
                    End If
                End If
            End If
        Next
    End With
    
    If Not mrsTrans Is Nothing Then
        If mrsTrans.RecordCount > 0 Then
            mrsTrans.Filter = ""
            mrsTrans.Sort = "���"
            mrsTrans.MoveLast
            lblCount.Caption = "��Һ����" & mlng��ɨ�� + mlngδɨ�� & " �ѣ�" & mlng��ɨ�� & "  δ��" & mlngδɨ�� & " ��ǰѡ����Һ����" & lngCount
            mrsTrans.MoveFirst
        End If
    End If
    
    lblVolu.Visible = False
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        lblVolu.Visible = True
        Me.lblVolu.Caption = "������" & lngVolume
    End If
    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "  ��ǰѡ����Һ����" & lngCount & IIf(mcondition.strTransStep = M_STR_CALSS_DOSAGE, " ��ǰ�ı��˴��״̬����Һ����" & mintCountPack, "")
    
End Sub

Private Function Check�Ա�ҩ() As Boolean
    '���ܣ�����Ա�ҩ�Ĵ���ҩ����
    Dim strSQL As String
    Dim rs���� As ADODB.Recordset
    Dim rsʵ�� As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    Dim str����ids As String     '���磺����1,����2...
    Dim lng����id As Long
    Dim str��ҩids As String     '���磺��ҩid1,��ҩid2...
    Dim str��ǰ���� As String
    Dim str��ǰҩƷ As String
    Dim lng��ǰ��ҩid As Long
    
    On Error GoTo errHandle
    
    Check�Ա�ҩ = False
    
    If Not mblnShowOhters Then Exit Function
    
    If mrsTrans Is Nothing Then
        MsgBox "��ȡ�����쳣��������ˢ�����ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Filter = "ִ�б�־=1"
    
    If mrsTrans.RecordCount = 0 Then
        MsgBox "��ȡ�����쳣��������ˢ�����ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Set rsData = mrsTrans
    
    '����Һ�����а����˷���
    With rsData
        
        .Sort = "����id"
        
        Do While Not .EOF
            '�ռ�����ID
            If InStr("," & str����ids & ",", "," & !����ID & ",") = 0 Then
                str����ids = str����ids & IIf(str����ids = "", "", ",") & !����ID
            End If
            
            .MoveNext
        Loop
        
        '���в�ѯ�ò��˵��Ƿ�����Ա�ҩ��¼
        For i = 0 To UBound(Split(str����ids, ","))
            lng����id = Split(str����ids, ",")(i)
            
            .Filter = "ִ�б�־=1 and ����id =" & lng����id
            .Sort = "��ҩid"
            
            '��ͬ������Ҫ��ʼ��
            str��ҩids = ""
            
            '�ռ���ҩid
            Do While Not .EOF
                If InStr("," & str��ҩids & ",", "," & !��ҩid & ",") = 0 Then
                    str��ҩids = str��ҩids & IIf(str��ҩids = "", "", ",") & !��ҩid
                End If

                .MoveNext
            Loop
            
            '��1.�Ƚ���ҩƷ���ܼ�顿
            strSQL = "Select c.ҩƷid, Sum((b.�������� / c.����ϵ��)) As ����, a.����, e.����" & vbNewLine & _
                    "From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ҩƷ��� C, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) D, �շ���ĿĿ¼ E" & vbNewLine & _
                    "Where a.ҽ��id = b.���id And a.Id = d.Column_Value And b.�շ�ϸĿid = c.ҩƷid And c.ҩƷid = e.Id And b.ִ������ = 5 And b.ִ�б�� = 0 And" & vbNewLine & _
                    "      b.�շ�ϸĿid In (Select d.ҩƷid From ��Һ�Ա�ҩ�嵥 D Where d.�Ƿ����� = 1)" & vbNewLine & _
                    "Group By c.ҩƷid, a.����, e.����"
                    
            Set rs���� = zlDatabase.OpenSQLRecord(strSQL, "����Ա�ҩ", str��ҩids)
                        
            '����ӦҩƷ�����Ƿ��㹻
            Do While Not rs����.EOF
                str��ǰ���� = rs����!����
                str��ǰҩƷ = rs����!����
                
                strSQL = "Select Sum(b.ʵ������) As ʵ������, c.����" & vbNewLine & _
                        "From δ��ҩƷ��¼ A, ҩƷ�շ���¼ B, �շ���ĿĿ¼ C" & vbNewLine & _
                        "Where a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And b.ҩƷid = c.Id And b.����� Is Null And b.������� Is Null And" & vbNewLine & _
                        "      Mod(b.��¼״̬, 3) = 1 And a.����id = [1] And a.�ⷿid = [2] And b.ҩƷid = [3] And Exists (Select 1 From ������ü�¼ C Where c.Id = b.����id)" & vbNewLine & _
                        "Group By ����"
                
                Set rsʵ�� = zlDatabase.OpenSQLRecord(strSQL, "�˶��Ա�ҩ����", lng����id, mParams.lng��������, rs����!ҩƷID)
                
                If rsʵ��.EOF Then
                    MsgBox "���� " & str��ǰ���� & " ���Ա�ҩҩƷ��" & str��ǰҩƷ & "���Ĵ���ҩ�������㣬�޷����а�ҩ��", vbExclamation, "�Ա�ҩ�����������"
                    Exit Function
                Else
                    '������������������ʾ����ֹ���
                    If nvl(rsʵ��!ʵ������, 0) < nvl(rs����!����, 0) Then
                        MsgBox "���� " & str��ǰ���� & " ���Ա�ҩҩƷ��" & str��ǰҩƷ & "���Ĵ���ҩ�������㣬�޷����а�ҩ��", vbExclamation, "�Ա�ҩ�����������"
                        Exit Function
                    End If
                End If
                
                rs����.MoveNext
            Loop
            
            '��2.�ٽ�����ҩ������ҩƷ���ܼ�顿
            strSQL = "Select c.ҩƷid, Sum((b.�������� / c.����ϵ��)) As ����, a.����, e.����, a.Id" & vbNewLine & _
                    "From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ҩƷ��� C, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) D, �շ���ĿĿ¼ E" & vbNewLine & _
                    "Where a.ҽ��id = b.���id And a.Id = d.Column_Value And b.�շ�ϸĿid = c.ҩƷid And c.ҩƷid = e.Id And b.ִ������ = 5 And b.ִ�б�� = 0 And" & vbNewLine & _
                    "      b.�շ�ϸĿid In (Select d.ҩƷid From ��Һ�Ա�ҩ�嵥 D Where d.�Ƿ����� = 1)" & vbNewLine & _
                    "Group By c.ҩƷid, a.����, e.����, a.Id"
                    
            Set rs���� = zlDatabase.OpenSQLRecord(strSQL, "����Ա�ҩ", str��ҩids)
            
            '�����ҩ������ӦҩƷ�����Ƿ��㹻
            Do While Not rs����.EOF
                str��ǰ���� = rs����!����
                str��ǰҩƷ = rs����!����
                lng��ǰ��ҩid = rs����!Id
                
                strSQL = "Select Sum(b.ʵ������) As ʵ������, c.����" & vbNewLine & _
                        "From δ��ҩƷ��¼ A, ҩƷ�շ���¼ B, �շ���ĿĿ¼ C" & vbNewLine & _
                        "Where a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And b.ҩƷid = c.Id And b.����� Is Null And b.������� Is Null And b.�ƻ�id = [4] And" & vbNewLine & _
                        "      Mod(b.��¼״̬, 3) = 1 And a.����id = [1] And a.�ⷿid = [2] And b.ҩƷid = [3]" & vbNewLine & _
                        "Group By ����"
                
                Set rsʵ�� = zlDatabase.OpenSQLRecord(strSQL, "�˶��Ա�ҩ����", lng����id, mParams.lng��������, rs����!ҩƷID, lng��ǰ��ҩid)
                
                If Not rsʵ��.EOF Then
                    '������������������ʾ����ֹ���
                    If nvl(rsʵ��!ʵ������, 0) < nvl(rs����!����, 0) Then
                        MsgBox "���� " & str��ǰ���� & " ���Ա�ҩҩƷ��" & str��ǰҩƷ & "���Ĵ���ҩ�������㣬�޷����а�ҩ��", vbExclamation, "�Ա�ҩ�����������"
                        Exit Function
                    End If
                End If
                
                rs����.MoveNext
            Loop
        Next
    End With

    Check�Ա�ҩ = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InputIsScaner(ByRef txtInput As Object, ByVal KeyAscii As Integer) As Boolean
'���ܣ��ж�ָ���ı����е�ǰ�����Ƿ����������豸����
'������KeyAscii=��KeyPress�¼��е��õĲ���
    Static sngInputBegin As Single
    Dim sngNow As Single, blnScaner As Boolean, strText As String
    
    '����ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 10 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    
    '�ж��Ƿ��������豸����
    sngNow = Timer
    If txtInput.Text = "" Or strText = "" Then
        sngInputBegin = sngNow
    Else
        If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnScaner = True
    End If
    
    InputIsScaner = blnScaner
End Function



Private Function CheckBill(ByVal str��ҩID�� As String) As Boolean
    Dim str�շ�ID�� As String
    Dim lngCount As Long
    
    If str��ҩID�� = "" Then Exit Function
    If mrsTrans Is Nothing Then Exit Function
    
    mrsTrans.Filter = ""
    If mrsTrans.RecordCount = 0 Then Exit Function
    
    With mrsTrans
        For lngCount = 0 To UBound(Split(str��ҩID��, ","))
            If Val(Split(str��ҩID��, ",")(lngCount)) > 0 Then
                .Filter = "��ҩID=" & Val(Split(str��ҩID��, ",")(lngCount))
                If .RecordCount > 0 Then
                    Do While Not .EOF
                        If InStr(1, "," & str�շ�ID�� & ",", "," & Val(!�շ�ID) & ",") = 0 Then
                            str�շ�ID�� = IIf(str�շ�ID�� = "", "", str�շ�ID�� & ",") & Val(!�շ�ID)
                        End If
                        .MoveNext
                    Loop
                End If
            End If
        Next
    End With
    
    If str�շ�ID�� = "" Then Exit Function
    
    For lngCount = 0 To UBound(Split(str�շ�ID��, ","))
        If Val(Split(str�շ�ID��, ",")(lngCount)) > 0 Then
            If DeptSendWork_CheckBill(1, Val(Split(str�շ�ID��, ",")(lngCount)), mParams.bln����δ��˴�����ҩ) > 0 Then Exit Function
        End If
    Next
    
    CheckBill = True
End Function
Private Function CheckStock() As Boolean
    '�����
    Dim lng�շ�ID As Long
    Dim str�շ�ID As String
    Dim rsData As ADODB.Recordset
    Dim blnIsShort As Boolean
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    If mParams.IntCheckStock = 0 Then
        CheckStock = True
        Exit Function
    End If
    
    If mrsTrans Is Nothing Then
        MsgBox "��ȡ�����쳣��������ˢ�����ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Filter = "ִ�б�־=1"
    
    If mrsTrans.RecordCount = 0 Then
        MsgBox "��ȡ�����쳣��������ˢ�����ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Sort = "�շ�ID"
    
    Set rstemp = mrsTrans
    If mrsTrans.RecordCount = 0 Then
        MsgBox "��ȡ�����쳣��������ˢ�����ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Do While Not rstemp.EOF
        If lng�շ�ID <> rstemp!�շ�ID Then
            lng�շ�ID = rstemp!�շ�ID
            str�շ�ID = IIf(str�շ�ID = "", "", str�շ�ID & ",") & rstemp!�շ�ID
        End If
        rstemp.MoveNext
        
        If Len(str�շ�ID) >= 3950 Then
            '���¸��¿��
            gstrSQL = "Select /*+ Rule*/ " & _
                " A.ID As �շ�id, A.ʵ������ * Nvl(����, 1) / D.סԺ��װ As ��ҩ����, B.ʵ������ / D.סԺ��װ As ������� " & _
                " From ҩƷ�շ���¼ A, " & _
                " (Select �ⷿid, ҩƷid, Nvl(����, 0) As ����, Nvl(ʵ������, 0) As ʵ������ " & _
                " From ҩƷ��� Where ���� = 1 And �ⷿid = [1]) B, " & _
                " Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)) C, ҩƷ��� D " & _
                " Where A.�ⷿid + 0 = B.�ⷿid(+) And A.ҩƷid + 0 = B.ҩƷid(+) And Nvl(A.����, 0) = B.����(+) And A.ҩƷid + 0 = D.ҩƷid " & _
                " And A.������� Is Null And A.�ⷿid + 0 = [1] And A.ID = C.Column_Value"
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�����", mcondition.lngCenterID, str�շ�ID)
            
            Do While Not rsData.EOF
                If rsData!��ҩ���� > rsData!������� Then
                    blnIsShort = True
                End If
                
                mrsTrans.Filter = "�շ�ID=" & rsData!�շ�ID
                Do While Not mrsTrans.EOF
                    mrsTrans!������� = rsData!�������
                    mrsTrans.Update
                    
                    mrsTrans.MoveNext
                Loop
                rsData.MoveNext
            Loop
            
            str�շ�ID = ""
        End If
    Loop
    
    If blnIsShort = True Then
        If mParams.IntCheckStock = 1 Then
            CheckStock = (MsgBox("����ѡ����ҩ����Һ����Ӧ����ЩҩƷ��治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
            Exit Function
        Else
            MsgBox "����ѡ����ҩ����Һ����Ӧ����ЩҩƷ��治�㣬���ܼ���������ҩƷ�����б��в鿴��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub GetPrivs()
    With mPrives
        .bln�˲�ȷ�� = IsInString(mstrPrivs, "�˲�ȷ��", ";")
        .blnȡ����� = IsInString(mstrPrivs, "ȡ�����", ";")
        .bln��ҩȷ�� = IsInString(mstrPrivs, "��ҩȷ��", ";")
        .blnȡ����ҩ = IsInString(mstrPrivs, "ȡ����ҩ", ";")
        .bln��ҩȷ�� = IsInString(mstrPrivs, "��ҩȷ��", ";")
        .blnȡ����ҩ = IsInString(mstrPrivs, "ȡ����ҩ", ";")
        .bln����ȷ�� = IsInString(mstrPrivs, "����ȷ��", ";")
        .blnȡ������ = IsInString(mstrPrivs, "ȡ������", ";")
        .bln�������� = IsInString(mstrPrivs, "��������", ";")
        .bln������� = IsInString(mstrPrivs, "�������", ";")
        .blnȷ�Ͼܾ� = IsInString(mstrPrivs, "ȷ�Ͼܾ�", ";")
        .bln���ʾܾ� = IsInString(mstrPrivs, "���ʾܾ�", ";")
        .bln�Ű����� = IsInString(mstrPrivs, "�Ű�����", ";")
    End With
End Sub

Private Sub GetParams()
    Dim strAutoPrint As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    With mParams
        'ϵͳ����
        .lng�������� = Val(zlDatabase.GetPara("��������", glngSys, 1345, 0))
        .bln����δ��˴�����ҩ = (gtype_UserSysParms.P6_δ��˼��ʴ�����ҩ = 1)
        .bln����δ�շѴ�����ҩ = (gtype_UserSysParms.P148_δ�շѴ�����ҩ = 1)
        .bln����ȡ����ҩ = (gtype_UserSysParms.P15_�����շ��뷢ҩ���� = 1 Or gtype_UserSysParms.P16_סԺ�����뷢ҩ���� = 1)
        .blnҽ������ = (gtype_UserSysParms.P68_����ҩ�������Ϻ���ҩ = 0)
        .bln��˻��۵� = True
        '��ȡ�������ϵͳ��ϵͳ����
        .bln������� = (gtype_UserSysParms.P240_ҩ��������� = 2 Or gtype_UserSysParms.P240_ҩ��������� = 3)
        .bln��� = (gtype_UserSysParms.P214_�״�ҽ��ִ����Ҫ��� = 1 And Not .bln�������)
        .intƤ����Ч���� = (gtype_UserSysParms.P70_�����Ǽ���Ч���� = 1)
        
        '�������ã�����
        .int��ҩ���ӡ = Val(zlDatabase.GetPara("��ҩ���ӡ", glngSys, 1345, 0))
        .int���ͺ��ӡ = Val(zlDatabase.GetPara("���ͺ��ӡ", glngSys, 1345, 0))
        .bln�������� = (Val(zlDatabase.GetPara("��������", glngSys, 1345, 0)) = 1)
        .bln������� = (Val(zlDatabase.GetPara("�������", glngSys, 1345, 0)) = 1)
        strAutoPrint = zlDatabase.GetPara("ƿǩ�Զ���ӡ", glngSys, 1345, "00|00")
        .blnƿǩ�ֹ���ӡ = (Val(zlDatabase.GetPara("ƿǩ�ֹ���ӡ", glngSys, 1345, 0)) = 1)
        .intCount = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\��Һ��Ƭ", "��Ƭ����", 3))
        .intNum = Val(zlDatabase.GetPara("ƿǩ��ӡ����", glngSys, 1345, 0))
        .int��ӡ���� = Val(zlDatabase.GetPara("��ӡ��ǩ���Ƿ��ӡ���ܱ���", glngSys, 1345, 2))
        .blnLastBatch = (Val(zlDatabase.GetPara("�����ϴ�����", glngSys, 1345, 0)) = 1)
        .blnTwoCode = (Val(zlDatabase.GetPara("ɨ����ƿǩ���Զ�����", glngSys, 1345, 0)) = 1)
        .blnByMedi = (Val(zlDatabase.GetPara("�����Σ�ҩƷ����", glngSys, 1345, 0)) = 1)
        .intCheck = zlDatabase.GetPara("��˸�ҩ������������", glngSys, 1345, 0)
        .blnFilter = (Val(zlDatabase.GetPara("�Ƿ����õĳ���ҩƷ����ҩƷ���˲���", glngSys, 1345, 0)) = 1)
        .blnRePeople = (Val(zlDatabase.GetPara("��ӡƿǩʱ��д�������ڵ�ʵ�ʲ���Ա", glngSys, 1345, 0)) = 1)
        
        .intҩƷ������ʾ��ʽ = Val(zlDatabase.GetPara("ҩƷ������ʾ��ʽ", glngSys, 1345, 0))
        
        .strSourceDep = zlDatabase.GetPara("��ʾ��Դ����", glngSys, 1345, "")
        
        If InStr(1, strAutoPrint, "|") = 0 Or Len(strAutoPrint) <> 5 Then
            strAutoPrint = "00|00"
        End If
        
        If Mid(strAutoPrint, 1, 1) = 1 Then
            If Val(Mid(strAutoPrint, 2, 1)) = 1 Then
                .intƿǩ��ҩ���ӡ = 1
            Else
                .intƿǩ��ҩ���ӡ = 0
            End If
        Else
            .intƿǩ��ҩ���ӡ = 2
        End If
        If Mid(strAutoPrint, 4, 1) = 1 Then
            If Val(Mid(strAutoPrint, 5, 1)) = 1 Then
                .intƿǩ��ҩ���ӡ = 1
            Else
                .intƿǩ��ҩ���ӡ = 0
            End If
        Else
            .intƿǩ��ҩ���ӡ = 2
        End If
            
        '��������
        .IntCheckStock = MediWork_GetCheckStockRule(.lng��������)

        'PASS
        .intShowPass = gintPass
'        .blnShowPass = True
    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub BillPrint_Prepare()
    '��ӡ��ҩ��
    Dim StrDate As String
    
    With vsfTrans
        If .Row > 0 Then StrDate = .TextMatrix(.Row, .ColIndex("��ҩʱ��"))
    End With
      
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_1", Me, _
        "����=" & mcondition.lngCenterID, _
        "��ҩʱ��=" & StrDate, "������Ա=" & gstrUserName, "PrintEmpty=0", 1)
End Sub

Private Sub BillPrint_Send()
    '��ӡ���͵�
    Dim StrDate As String
    
    With vsfTrans
        If .Row > 0 Then StrDate = .TextMatrix(.Row, .ColIndex("����ʱ��"))
    End With
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_2", Me, _
            "����=" & mcondition.lngCenterID, _
            "����ʱ��=" & StrDate, "������Ա=" & gstrUserName, "PrintEmpty=0", 1)
End Sub

Private Sub BillPrint_Return()
    '��ӡ��ҩ�����嵥
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_2", Me, "��װϵ��=C.סԺ��װ", 2)
End Sub

Private Sub ClearDetailList()
    '�����ϸ�б�����
    vsfTrans.rows = 1
    vsfTrans.rows = 2
    
'    vsfDrug.rows = 1
'    vsfDrug.rows = 2
    
    vsfSumDrug.rows = 1
    vsfSumDrug.rows = 2
    
    Me.VSFLook.rows = 1
    Me.VSFLook.rows = 2
    
    Me.vsfMedis.rows = 1
    Me.vsfMedis.rows = 2
End Sub

Private Sub RefreshPrintSign(ByVal str��ҩid As String, ByVal dateNow As Date, Optional ByVal str������Ա As String)
    On Error GoTo errHandle

    '���´�ӡ��־
    gstrSQL = "Zl_��Һ��ҩ��¼_��ӡ("
    '��ҩID
    gstrSQL = gstrSQL & "'" & str��ҩid & "'"
    gstrSQL = gstrSQL & ",To_Date('" & dateNow & "','yyyy-MM-dd hh24:mi:ss')"
    gstrSQL = gstrSQL & IIf(str������Ա <> "", ",'" & str������Ա & " '", ",Null")
    gstrSQL = gstrSQL & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���´�ӡ��־")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ResizeConditionArea()
    Dim intCount As Integer
    
    On Error Resume Next

    'ʱ��������
    With picTime
        .Left = 0
        .Top = 0
        .Width = picCondition.Width
    End With
    
    With cboʱ�䷶Χ
        .Width = picTime.Width - .Left - 100
    End With
    
    lblTimeBegin.Visible = (cboʱ�䷶Χ.ListIndex = 3)
    With Dtp��ʼʱ��
        .Visible = (cboʱ�䷶Χ.ListIndex = 3)
        .Width = cboʱ�䷶Χ.Width
    End With
    
    lblTimeEnd.Visible = (cboʱ�䷶Χ.ListIndex = 3)
    With Dtp����ʱ��
        .Visible = (cboʱ�䷶Χ.ListIndex = 3)
        .Width = cboʱ�䷶Χ.Width
    End With

    With picShowSendType
        If cboʱ�䷶Χ.ListIndex = 3 Then
            .Top = Dtp����ʱ��.Top + Dtp����ʱ��.Height + 20
        Else
            .Top = cboʱ�䷶Χ.Top + cboʱ�䷶Χ.Height + 20
        End If
        .Width = picTime.Width
    End With
    
    picUpOrDown1.Left = picShowSendType.Width - picUpOrDown1.Width - 10
    
    Me.lbldept.Top = picShowSendType.Top + picShowSendType.Height + 20
    Me.txtdept.Top = picShowSendType.Top + picShowSendType.Height + 20
    txtdept.Width = Dtp��ʼʱ��.Width
    
    Me.lblName.Top = txtdept.Top + txtdept.Height + 20
    Me.txtName.Top = txtdept.Top + txtdept.Height + 20
    txtName.Width = Dtp��ʼʱ��.Width
    
    Me.lblDrug.Top = txtName.Top + txtName.Height + 20
    Me.txtDrug.Top = txtName.Top + txtName.Height + 20
    txtDrug.Width = Dtp��ʼʱ��.Width
    cmdDrug.Top = txtDrug.Top
    cmdDrug.Left = txtDrug.Left + txtDrug.Width - cmdDrug.Width
    
    Me.lblTag.Top = txtDrug.Top + txtDrug.Height + 20
    Me.txtTag.Top = txtDrug.Top + txtDrug.Height + 20
    txtTag.Width = Dtp��ʼʱ��.Width
    
    
    With picTime
        If txtTag.Visible = True Then
            .Height = txtTag.Top + txtTag.Height
        Else
            .Height = picShowSendType.Top + picShowSendType.Height
        End If
    End With
    
    '��Ϣ�б�
    With Me.picMsg
        .Left = 0
        .Width = picCondition.Width
        .Height = picUpOrDown.Top + picUpOrDown.Height + 50 + IIf(lblMsgComment.Tag = "1", vsfMsg.Height + 50, 0)
        .Top = picCondition.Height - .Height - 50
    End With
    
    '�����б���
    With picDeptList
        .Left = 0
        .Top = picTime.ScaleTop + picTime.ScaleHeight
        .Width = picCondition.Width
        .Height = picCondition.Height - .Top - IIf(picMsg.Visible, Me.picMsg.Height, 0) - 50
    End With
   
    With fraLineH1
        .Top = 50
        .Width = picTime.Width + 100
    End With
   
 End Sub

Private Sub DeleteBatch(ByVal lngType As Long)
    'ɾ����������
    Dim strInputID As String
    Dim lngRow As Long
    Dim strCom As String
    
    On Error GoTo errHandle
    
    With vsfTrans
        If lngType = conMenu_Oper_DelBatch_SelBatch Then
            strCom = .TextMatrix(.Row, .ColIndex("��ҩ����"))
        ElseIf lngType = conMenu_Oper_DelBatch_SelDept Then
            strCom = .TextMatrix(.Row, .ColIndex("����"))
        ElseIf lngType = conMenu_Oper_DelBatch_SelPati Then
            strCom = .TextMatrix(.Row, .ColIndex("����")) & .TextMatrix(.Row, .ColIndex("����")) & .TextMatrix(.Row, .ColIndex("����"))
        End If
        
        If lngType = conMenu_Oper_DelBatch_SelRow Then
            '��ǰ��
            If .TextMatrix(.Row, .ColIndex("��ҩ����")) <> "" Then
                strInputID = Val(.TextMatrix(.Row, .ColIndex("��ҩID")))
            End If
        Else
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                    If lngType = conMenu_Oper_DelBatch_SelBatch Then
                        If .TextMatrix(lngRow, .ColIndex("��ҩ����")) <> "" And .TextMatrix(lngRow, .ColIndex("��ҩ����")) = strCom Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("��ҩID"))
                        End If
                    ElseIf lngType = conMenu_Oper_DelBatch_SelDept Then
                        If .TextMatrix(lngRow, .ColIndex("��ҩ����")) <> "" And .TextMatrix(lngRow, .ColIndex("����")) = strCom Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("��ҩID"))
                        End If
                    ElseIf lngType = conMenu_Oper_DelBatch_SelPati Then
                        If .TextMatrix(lngRow, .ColIndex("��ҩ����")) <> "" And .TextMatrix(lngRow, .ColIndex("����")) & .TextMatrix(lngRow, .ColIndex("����")) & .TextMatrix(lngRow, .ColIndex("����")) = strCom Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("��ҩID"))
                        End If
                    ElseIf lngType = conMenu_Oper_DelBatch_AllRow Then
                        If .TextMatrix(lngRow, .ColIndex("��ҩ����")) <> "" And Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("��ҩID"))
                        End If
                    End If
                End If
            Next
        End If
    End With
    
    If strInputID = "" Then Exit Sub
    
    gstrSQL = "Zl_��Һ��ҩ��¼_�������("
    '��ҩID
    gstrSQL = gstrSQL & "'" & strInputID & "'"
    gstrSQL = gstrSQL & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
    
    DoEvents
    
    '�������ݼ�����
    strInputID = "," & strInputID & ","
    With mrsTrans
        .Filter = ""
        Do While Not .EOF
            If InStr(strInputID, "," & !��ҩid & ",") > 0 Then
                !��ҩ���� = ""
                .Update
            End If
            .MoveNext
        Loop
    End With
    
    DoEvents
    
    '�����б���ʾ
    With vsfTrans
        .Redraw = flexRDNone
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                If InStr(strInputID, "," & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ",") > 0 Then
                    .TextMatrix(lngRow, .ColIndex("��ҩ����")) = ""
                End If
            End If
        Next
        .Redraw = flexRDDirect
    End With
    
    MsgBox "���������ɣ�", vbInformation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub GetCondition()
    Dim strStartTime As String
    Dim strEndTime As String
    
    'ʱ�䷶Χ
    Select Case cboʱ�䷶Χ.ListIndex
        Case 0
            '����
            mcondition.intTransTimeSel = 0
            mcondition.strTransStartTime = Format(mdateToday, "yyyy-mm-dd") & " 00:00:00"
            mcondition.strTransEndTime = Format(mdateToday, "yyyy-mm-dd") & " 23:59:59"
        Case 1
            '����
            mcondition.intTransTimeSel = 1
            mcondition.strTransStartTime = Format(DateAdd("d", 1, mdateToday), "yyyy-mm-dd") & " 00:00:00"
            mcondition.strTransEndTime = Format(DateAdd("d", 1, mdateToday), "yyyy-mm-dd") & " 23:59:59"
        Case 2
            '���պ�����
            mcondition.intTransTimeSel = 2
            mcondition.strTransStartTime = Format(mdateToday, "yyyy-mm-dd") & " 00:00:00"
            mcondition.strTransEndTime = Format(DateAdd("d", 1, mdateToday), "yyyy-mm-dd") & " 23:59:59"
        Case 3
            'ָ�����ڷ�Χ
            mcondition.intTransTimeSel = 3
            mcondition.strTransStartTime = Format(Dtp��ʼʱ��.Value, "yyyy-mm-dd hh:mm:ss")
            mcondition.strTransEndTime = Format(Dtp����ʱ��.Value, "yyyy-mm-dd hh:mm:ss")
    End Select
End Sub

Private Sub GetTransCount(ByVal dateStart As Date, ByVal dateEnd As Date)
    'ȡ������������Ӧ����Һ��������
    Dim rsTmp As ADODB.Recordset
    Dim lngCount As Long
    Dim intTabIndex As Integer
    Dim strCaption As String
    Dim lng����id As Long
    Dim intType As Integer
    Dim lng���� As Long
    Dim str��¼id As String
    Dim str���� As String
    
    '������Ӧ����Һ����
    Set mrsDeptTrans = New ADODB.Recordset
    With mrsDeptTrans
        If .State = 1 Then .Close
        
        .Fields.Append "ѡ��", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adDouble, 1, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 10, adFldIsNullable
        .Fields.Append "��¼id", adLongVarChar, 20000, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        If Mid(Me.lblName.Caption, 1, Len(Me.lblName.Caption) - 1) = "����" Then
            intType = 1
        ElseIf Mid(Me.lblName.Caption, 1, Len(Me.lblName.Caption) - 1) = "����" Then
            intType = 2
        Else
            intType = 3
        End If
        
        Set rsTmp = PIVA_GetTransCount(mcondition.lngCenterID, dateStart, dateEnd, mParams.bln���, mParams.bln�������, intType, Me.txtName.Text, Val(Me.txtDrug.Tag), Me.txtTag.Text, Val(Me.txtdept.Tag), mParams.intCheck, mParams.strSourceDep)
        
        '�����¼��
        rsTmp.Sort = "����,����id"
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                If rsTmp!���� = "00" And mParams.intCheck = 1 Then
                    
                    If lng����id <> rsTmp!����ID Or str���� <> rsTmp!���� Then
                    
                        If lng����id <> 0 Then
                            !���� = lng����
                            !��¼id = str��¼id
                        End If
                        
                        lng����id = rsTmp!����ID
                        .AddNew
                        !ѡ�� = 0
                        !���� = rsTmp!����
                        !����ID = rsTmp!����ID
                        !���� = rsTmp!����
                        !���� = rsTmp!����
                        !���� = rsTmp!����
                        str���� = rsTmp!����
                        If nvl(rsTmp!ҩʦ��˱�־, 0) = 0 Or nvl(rsTmp!ҩʦ��˱�־, 0) = 3 Then
                            lng���� = 1
                        Else
                            lng���� = 0
                        End If
                        str��¼id = rsTmp!Id
                    Else
                        If nvl(rsTmp!ҩʦ��˱�־, 0) = 0 Or nvl(rsTmp!ҩʦ��˱�־, 0) = 3 Then
                            lng���� = lng���� + 1
                        End If
                        str��¼id = str��¼id & "," & rsTmp!Id
                    End If
                Else
                    If lng����id <> rsTmp!����ID Or str���� <> rsTmp!���� Then
                    
                        If lng����id <> 0 Then
                            !���� = lng����
                            !��¼id = str��¼id
                        End If
                        
                        lng����id = rsTmp!����ID
                        .AddNew
                        !ѡ�� = 0
                        !���� = rsTmp!����
                        !����ID = rsTmp!����ID
                        !���� = rsTmp!����
                        !���� = rsTmp!����
                        !���� = rsTmp!����
                        str���� = rsTmp!����
                        lng���� = 1
                        str��¼id = rsTmp!Id
                    Else
                        lng���� = lng���� + 1
                        str��¼id = str��¼id & "," & rsTmp!Id
                    End If
                End If
                .Update
                
                rsTmp.MoveNext
                
                If rsTmp.EOF Then
                    !���� = lng����
                    !��¼id = str��¼id
                End If
                
            Loop
        End If
    End With
    
    '�����ҵ�񻷽ڵ�ҽ������Һ������������ʾ�ڷ�ҳ��ǩ��
    For intTabIndex = 0 To Me.tbcLook.ItemCount - 1
        lngCount = 0

        If Not mrsDeptTrans Is Nothing Then
            mrsDeptTrans.Filter = "����='" & tbcLook.Item(intTabIndex).Tag & "'"
            Do While Not mrsDeptTrans.EOF
                lngCount = lngCount + mrsDeptTrans!����

                mrsDeptTrans.MoveNext
            Loop
            
            strCaption = tbcLook.Item(intTabIndex).Caption
            strCaption = Mid(strCaption, 1, InStr(1, strCaption, "(")) & lngCount & ")"
            tbcLook.Item(intTabIndex).Caption = strCaption
        End If
    Next
    
    For intTabIndex = 0 To Me.tabWork.ItemCount - 1
        lngCount = 0

        If Not mrsDeptTrans Is Nothing Then
            mrsDeptTrans.Filter = "����='" & tabWork.Item(intTabIndex).Tag & "'"
            Do While Not mrsDeptTrans.EOF
                lngCount = lngCount + mrsDeptTrans!����

                mrsDeptTrans.MoveNext
            Loop
            
            strCaption = Me.tabWork.Item(intTabIndex).Caption
            strCaption = Mid(strCaption, 1, InStr(1, strCaption, "(")) & lngCount & ")"
            tabWork.Item(intTabIndex).Caption = strCaption
        End If
    Next
    Call SetTabColor(tabWork)
    Call SetTabColor(tbcLook)
End Sub
Private Sub GetWorkBatchRec()
    'ȡ��Һ�������ĵĹ�������
    On Error GoTo errHandle
    gstrSQL = "Select ����,��ɫ, ��ҩʱ��, ��ҩʱ��, ���, 1 ���� From ��ҩ�������� Where ����=1 and ��������ID=[1] " & _
        " Union All " & _
        " Select Max(Nvl(����, 0))+1 ����,0 ��ɫ, '' ��ҩʱ��, '' ��ҩʱ��, 0 ���, 0 ���� From ��ҩ�������� where ��������ID=[1] " & _
        " Order By ����"
    Set mrsWorkBatch = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Һ�������Ĺ�������", mParams.lng��������)
    
    mstr���� = ""
    mParams.strBatchList = ""
    mstr��� = ""
    With mrsWorkBatch
        cboBatch.Clear
        cboBatch.AddItem "<ȫ��>"
        Do While Not .EOF
            If !���� = 1 Then
                mParams.strBatchList = IIf(mParams.strBatchList = "", "", mParams.strBatchList & "|") & !���� & "#" & _
                    vbTab & "��ҩʱ��" & !��ҩʱ�� & _
                    vbTab & "��ҩʱ��" & !��ҩʱ��
                    
                mstr��� = mstr��� & "," & !���� & "#," & zlStr.nvl(!���, 0)
                    
                mstr���� = IIf(mstr���� = "", "", mstr���� & "/") & !���� & "#" & _
                    "|" & "��ҩʱ��" & !��ҩʱ�� & _
                    "|" & "��ҩʱ��" & !��ҩʱ�� & "," & IIf(zlStr.nvl(!��ɫ) = "", 0, !��ɫ)
            Else
                mParams.strBatchList = IIf(mParams.strBatchList = "", "", mParams.strBatchList & "|") & zlStr.nvl(!����, 1) & "#" & vbTab & "���������Σ�"
                mstr���� = IIf(mstr���� = "", "", mstr���� & "/") & zlStr.nvl(!����, 1) & "#|���������Σ�" & "," & IIf(zlStr.nvl(!��ɫ) = "", 0, !��ɫ)
                mstr��� = IIf(mstr��� = "", "", mstr��� & ",") & zlStr.nvl(!����, 1) & "#,0"
            End If
            
            '����������Ϣ��������   IIf(mstr��� = "", "", mstr��� & ",") & NVL(!����, 1) & "#,0"
            cboBatch.AddItem !���� & "#"
            
            .MoveNext
        Loop
        cboBatch.Text = "<ȫ��>"
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub IniPriRec()
    Set mrstemp = New ADODB.Recordset
    With mrstemp
        If .State = 1 Then .Close
        .Fields.Append "��ҩid", adDouble, 18, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub IniTransRec()
    '��Һ����¼��
    Set mrsTrans = New ADODB.Recordset
    With mrsTrans
        If .State = 1 Then .Close
        
        '�ü�¼��Ӧ����Һ��ҩ��¼��Ϣ
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩid", adDouble, 18, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 3, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ִ��ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҳid", adDouble, 18, adFldIsNullable
        .Fields.Append "���ȼ�", adDouble, 18, adFldIsNullable
        .Fields.Append "���˿���id", adDouble, 18, adFldIsNullable
        .Fields.Append "���ʱ��", adLongVarChar, 20, adFldIsNullable
        
        '��Һ��ҩ��¼ҵ�������Ϣ
        .Fields.Append "��ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ƿǩ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ӡ��־", adDouble, 1, adFldIsNullable
        .Fields.Append "�Ƿ���", adDouble, 1, adFldIsNullable
        .Fields.Append "�˲���", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�˲�ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��ҩʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�������ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ҩ��", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "ҩʦ���ʱ��", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "�Ƿ��������", adDouble, 1, adFldIsNullable
        .Fields.Append "�Ƿ�����", adDouble, 1, adFldIsNullable
        .Fields.Append "����ҩ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�ֹ���������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ԭ��", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "�Ƿ�ȷ�ϵ���", adDouble, 1, adFldIsNullable
        
        '��Һ��ҩ��¼��Ӧ��ҩƷ��Ϣ
        .Fields.Append "�շ�id", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 2, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable   '����+ͨ����/��Ʒ��
        .Fields.Append "ҩƷ��������", adLongVarChar, 50, adFldIsNullable   '�̶���ʾ����+ͨ����/��Ʒ��,���ڻ����б������
        .Fields.Append "ͨ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ӣ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adDouble, 20, adFldIsNullable
        .Fields.Append "������λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "�÷�", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 3, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ý", adDouble, 1, adFldIsNullable
        .Fields.Append "�Ƿ�Ƥ��", adDouble, 1, adFldIsNullable
        .Fields.Append "��ҩ����1", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���α��", adDouble, 1, adFldIsNullable
        .Fields.Append "��ýid", adDouble, 18, adFldIsNullable
        .Fields.Append "����ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "������", adDouble, 20, adFldIsNullable
        .Fields.Append "�Ƿ���", adDouble, 1, adFldIsNullable
        
        .Fields.Append "��ҩ����", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable
        
        .Fields.Append "�����", adDouble, 1, adFldIsNullable
        .Fields.Append "ҽ������ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƥ�Խ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҽ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "��Ӧҽ��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���ͺ�", adDouble, 18, adFldIsNullable
        .Fields.Append "ִ��Ƶ��", adLongVarChar, 50, adFldIsNullable
        
        .Fields.Append "ִ�б�־", adDouble, 1, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adDouble, 5, adFldIsNullable
        .Fields.Append "��ɫ", adDouble, 18, adFldIsNullable
        .Fields.Append "����ԭ��", adLongVarChar, 200, adFldIsNullable
        
        .Fields.Append "ʵ����ҩ����", adLongVarChar, 50, adFldIsNullable       '������ʾ����ҩƷ����ҩ���ͣ���������ýҩƷ��
        
        .Fields.Append "ִ������", adDouble, 1, adFldIsNullable
        .Fields.Append "ִ�б��", adDouble, 1, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub


Private Sub InitPanes()
    '��ʼ�������ؼ�
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
'    Me.dkpMain.Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
    
    Dim objPaneCon As Pane

    Set objPaneCon = Me.dkpMain.CreatePane(mconPane_PIVA_Condition, 225, 100, DockLeftOf, Nothing)
    objPaneCon.Title = mstrCenterName
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
End Sub

Private Function Find��������(ByVal rsData As ADODB.Recordset, ByVal lng��ҩid As Long) As String
    '���ܣ���ʾ�Ա�ҩ��ȡҩ����������
    
    Find�������� = ""
    
    '��δ���������ʾ���򲻽��в���
    If Not mblnShowOhters Then Exit Function
    
    rsData.Filter = "��ҩid = " & lng��ҩid
    
    Do While Not rsData.EOF
        If nvl(rsData!��������) <> "" Then
            Find�������� = rsData!��������
            Exit Function
        End If
        
        rsData.MoveNext
    Loop
    
    Find�������� = ""
    
End Function

Private Sub LoadTrans(ByVal strIDS As String, ByVal strStep As String, ByVal intPack As Integer, ByVal intSend As Integer)
    Dim rsTrans As ADODB.Recordset
    Dim lng��ҩid As Long
    Dim int��ҩ���� As Integer
    Dim rstemp As Recordset
    Dim dbl���� As Double
    Dim strOldִ��ʱ�� As String
    Dim strOld���� As String
    Dim lngOld����id As Long
    Dim lngOld��ҩid As Long
    Dim lng���� As Long
    Dim lngCount As Long
    Dim lng���ȼ� As Long
    Dim int��� As Integer
    Dim i As Integer
    Dim lng��ýid  As Long
    Dim lngҩƷid As Long
    Dim str��ҩ���� As String
    Dim dbl���� As Double
    Dim arrExecute As Variant
    Dim rsSel As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Call IniTransRec
    
    Call IniPriRec
    mintCountPack = 0
    If Not mParams.blnTwoCode Then Me.cboBatch.ListIndex = 0
    
    mblnFilter = False
    Me.cboLevel.ListIndex = 0
    Me.cboMedi.ListIndex = 0
    Me.cboDosType.ListIndex = 0
    Me.cboFrequency.ListIndex = 0
    mblnFilter = True
    
    If Not (mcondition.strTransStep = M_STR_CALSS_SEND And mParams.blnTwoCode = True) Then Me.cboBatch.ListIndex = 0
    
    arrExecute = GetArrayByStr(strIDS, 3950, ",")
    For i = 0 To UBound(arrExecute)
    
        
        Set rsTrans = Piva_GetTrans(CStr(arrExecute(i)), mParams.lng��������, strStep, intPack, mblnShowOhters)
        
        Set rsSel = rsTrans.Clone
        
        With rsTrans
            Set rstemp = rsTrans
            If .RecordCount > 0 Then
                rsTrans.Sort = "����id,��ҩid,ִ��ʱ��,��ҩ����,��ý,ҽ�����"
                Do While Not .EOF
                    lngCount = lngCount + 1
                                    
                    mrsTrans.AddNew
                    mrsTrans!��ҩid = !��ҩid
                    mrsTrans!����ID = !����ID
                    mrsTrans!��� = !���
                    mrsTrans!���� = IIf(IsNull(!����), "", !����)
                    mrsTrans!�Ա� = IIf(IsNull(!�Ա�), "", !�Ա�)
                    mrsTrans!���� = IIf(IsNull(!����), "", !����)
                    mrsTrans!סԺ�� = IIf(IsNull(!סԺ��), "", !סԺ��)
                    mrsTrans!���� = IIf(IsNull(!����), "", !����)
                    mrsTrans!�������� = IIf(IsNull(!��������), "", !��������)
                    mrsTrans!���� = IIf(IsNull(!����), "", !����)
                    mrsTrans!���� = !���˲���
                    mrsTrans!���� = !���˿���
                    mrsTrans!ִ��ʱ�� = IIf(IsNull(!ִ��ʱ��), "", Format(!ִ��ʱ��, "YYYY-MM-DD HH:MM"))
                    mrsTrans!����ID = IIf(IsNull(!����ID), 0, !����ID)
                    mrsTrans!��ҳid = IIf(IsNull(!��ҳid), 0, !��ҳid)
                    mrsTrans!���˿���id = IIf(IsNull(!���˿���id), 0, !���˿���id)
                    mrsTrans!���ʱ�� = nvl(!���ʱ��)
                    
                    mrsTrans!��ҩ���� = IIf(IsNull(!��ҩ����), "", !��ҩ���� & "#")
                    mrsTrans!����ҩ���� = IIf(IsNull(!��ҩ����), "", !��ҩ���� & "#")
                    mrsTrans!ƿǩ�� = IIf(IsNull(!ƿǩ��), "", !ƿǩ��)
                    mrsTrans!��ӡ��־ = IIf(IIf(IsNull(!��ӡ��־), 0, !��ӡ��־) = 0, 0, 1)
                    mrsTrans!�Ƿ��� = IIf(IsNull(!�Ƿ���), 0, !�Ƿ���)
                    mrsTrans!�˲��� = IIf(IsNull(!������Ա), "", !������Ա)
                    mrsTrans!�˲�ʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!��ҩ�� = IIf(IsNull(!������Ա), "", !������Ա)
                    mrsTrans!��ҩʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!��ҩ���� = IIf(IsNull(!��ҩ����), "", !��ҩ����)
                    mrsTrans!��ҩ�� = IIf(IsNull(!������Ա), "", !������Ա)
                    mrsTrans!��ҩʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!������ = IIf(IsNull(!������Ա), "", !������Ա)
                    mrsTrans!����ʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!���������� = IIf(IsNull(!������Ա), "", !������Ա)
                    mrsTrans!��������ʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!��������� = IIf(IsNull(!������Ա), "", !������Ա)
                    mrsTrans!�������ʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!����ҩ�� = 1
                    mrsTrans!ҩʦ���ʱ�� = IIf(IsNull(!ҩʦ���ʱ��), 0, !ҩʦ���ʱ��)
                    mrsTrans!�Ƿ�������� = IIf(IsNull(!�Ƿ��������), 0, !�Ƿ��������)
                    mrsTrans!�Ƿ����� = IIf(IsNull(!�Ƿ�����), 0, !�Ƿ�����)
                    mrsTrans!�ֹ��������� = IIf(IsNull(!�ֹ���������), 0, !�ֹ���������)
                    mrsTrans!����ԭ�� = nvl(!����ԭ��)
                    mrsTrans!�Ƿ�ȷ�ϵ��� = IIf(IsNull(!�Ƿ�ȷ�ϵ���), 0, !�Ƿ�ȷ�ϵ���)
                    
                    mrsTrans!�շ�ID = !�շ�ID
                    mrsTrans!���� = !����
                    mrsTrans!NO = nvl(!NO)
                    
                    mrsTrans!ҩƷ�������� = IIf(IsNull(!ҩƷ����), !ͨ����, "[" & !ҩƷ���� & "]" & !ͨ����)
                    
                    If mParams.intҩƷ������ʾ��ʽ = 0 Then
                        '���������
                        mrsTrans!ҩƷ���� = IIf(IsNull(!ҩƷ����), !ͨ����, "[" & !ҩƷ���� & "]" & !ͨ����)
                    ElseIf mParams.intҩƷ������ʾ��ʽ = 1 Then
                        '����
                        mrsTrans!ҩƷ���� = !ͨ����
                    ElseIf mParams.intҩƷ������ʾ��ʽ = 2 Then
                        '����
                        mrsTrans!ҩƷ���� = IIf(IsNull(!ҩƷ����), "", "[" & !ҩƷ���� & "]")
                    End If

                    mrsTrans!ͨ���� = !ͨ����
                    mrsTrans!��Ʒ�� = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
                    mrsTrans!Ӣ���� = IIf(IsNull(!Ӣ����), "", !Ӣ����)
                    mrsTrans!��� = IIf(IsNull(!���), "", !���)
                    mrsTrans!���� = IIf(IsNull(!����), "", !����)
                    mrsTrans!���� = IIf(IsNull(!����), "", !����)
                    mrsTrans!���� = FormatEx(nvl(!����, 0), 2)
                    mrsTrans!������λ = !������λ
                    mrsTrans!Ƶ�� = IIf(IsNull(!Ƶ��), "", !Ƶ��)
                    mrsTrans!���� = nvl(!����, 0)
                    mrsTrans!��λ = !��λ
                    mrsTrans!���� = !����
                    mrsTrans!�÷� = IIf(IsNull(!�÷�), "", !�÷�)
                    mrsTrans!ҩƷID = nvl(!ҩƷID, 0)
                    mrsTrans!ҩ��ID = !ҩ��ID
                    mrsTrans!������� = !�������
                    mrsTrans!����ID = !����ID
                    mrsTrans!���ȼ� = nvl(!���ȼ�, 0)
                    mrsTrans!���α�� = nvl(!���α��, 0)
                    mrsTrans!��ҩ���� = !��ҩ����1
                    mrsTrans!ʵ����ҩ���� = !��ҩ����1
                    mrsTrans!��ҩ���� = nvl(!��ҩ����, 0)
                    mrsTrans!������� = nvl(!�������, 0)
                    mrsTrans!ʵ������ = nvl(!ʵ������, 0)
                    
                    mrsTrans!ҽ��id = !ҽ��id
                    mrsTrans!��Ӧҽ��ID = !��Ӧҽ��ID
                    mrsTrans!���ͺ� = !���ͺ�
                    mrsTrans!����ʱ�� = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!Ƥ�Խ�� = nvl(!Ƥ�Խ��)
                    mrsTrans!ҽ������ʱ�� = IIf(IsNull(!ҽ������ʱ��), "", Format(!ҽ������ʱ��, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!����� = nvl(!�����, 0)
                    mrsTrans!�������� = IIf(IsNull(!��������), Find��������(rsSel, !��ҩid), !��������)
                    mrsTrans!���� = !����
                    mrsTrans!��ɫ = !��ɫ
                    mrsTrans!ִ�б�־ = 0
                    mrsTrans!��ý = nvl(!��ý, 0)
                    mrsTrans!�Ƿ�Ƥ�� = nvl(!�Ƿ�Ƥ��, 0)
                    mrsTrans!ִ��Ƶ�� = nvl(!ִ��Ƶ��, 0)
                    mrsTrans!����ҩƷID = nvl(!ҩƷID, 0)
                    mrsTrans!������ = FormatEx(nvl(!����, 0), 2)
                    mrsTrans!����ԭ�� = IIf(IsNull(!����ԭ��), "", !����ԭ��)
                    
                    mrsTrans!ִ������ = nvl(!ִ������, 0)
                    mrsTrans!ִ�б�� = nvl(!ִ�б��, 0)
                    
                    If mParams.blnFilter And mParams.str����ҩƷ <> "" Then
                        mrsTrans!�Ƿ��� = IIf(InStr(1, "," & mParams.str����ҩƷ & ",", !ҩƷID) > 0, 1, 0)
                    End If
                    
                    If !��ҩid <> lng��ҩid Then
                        int��� = int��� + 1
                    End If
                    mrsTrans!��� = int���
                    mrsTrans.Update
                    
                    
                    If !��ҩid = lng��ҩid Then
                        
                        If Val(!��ҩ����) > 0 Then
                            int��ҩ���� = 1
                        ElseIf int��ҩ���� = 0 And Val(!��ҩ����) = 0 Then
                            int��ҩ���� = 0
                        End If
                        
                        If str��ҩ���� = "" Then
                            str��ҩ���� = nvl(!��ҩ����1)
                        End If
                    Else
                        int��ҩ���� = Val(!��ҩ����)
                        If nvl(!��ý, 0) = 0 Then
                            lngҩƷid = nvl(!ҩƷID, 0)
                            dbl���� = FormatEx(nvl(!����, 0), 2)
                            str��ҩ���� = nvl(!��ҩ����1)
                        End If
                    End If
                    
                    If !��ý = 1 Then
                        lng��ýid = nvl(!ҩƷID, 0)
                    End If
                    
                    mrsTrans.Filter = ""
                    lng��ҩid = !��ҩid
                    
                    .MoveNext
                    
                    If .EOF Then
                        mrsTrans.Filter = "��ҩid=" & lng��ҩid
                        mrsTrans.MoveFirst
                        Do While Not mrsTrans.EOF
                            mrsTrans.Update "����ҩ��", int��ҩ����
                            mrsTrans.Update "��ýid", lng��ýid
                            mrsTrans.Update "����ҩƷid", lngҩƷid
                            mrsTrans.Update "��ҩ����", str��ҩ����
                            mrsTrans.MoveNext
                        Loop
                    Else
                        If lng��ҩid <> !��ҩid Then
                            mrsTrans.Filter = "��ҩid=" & lng��ҩid
                            mrsTrans.MoveFirst
                            Do While Not mrsTrans.EOF
                                mrsTrans.Update "����ҩ��", int��ҩ����
                                mrsTrans.Update "��ýid", lng��ýid
                                mrsTrans.Update "����ҩƷid", lngҩƷid
                                mrsTrans.Update "��ҩ����", str��ҩ����
                                mrsTrans.Update "������", dbl����
                                mrsTrans.MoveNext
                            Loop
                            lngҩƷid = 0
                            lng��ýid = 0
                            str��ҩ���� = ""
                        End If
                    End If
                    
                    
                Loop
            End If
        End With
    Next
    
    mrsTrans.Filter = ""
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboDosType_Click()
  Call SetFilter
End Sub

Private Sub cboFrequency_Click()
    Call SetFilter
End Sub

Private Sub cboMedi_Click()
    Call SetFilter
End Sub



Private Sub cboType_Click()
    Call LoadVsfMedi(mstr�ϴ�IDS, False)
    If mcondition.strTransStep = M_STR_CALSS_SEND Then Me.txtFindItem.SetFocus
End Sub

Private Sub chkAllDept_Click()
    Dim lngRow As Long
    
    With vsfDept(tabDeptList.Selected.index)
        If .rows = 1 Then Exit Sub
        If Val(.TextMatrix(1, .ColIndex("����ID"))) = 0 Then Exit Sub
        
        For lngRow = 1 To .rows - 1
            .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(chkAllDept.Value = 0, 0, -1)
        Next
    End With
    
'    DoEvents
'    Call RefreshDetailList(Me.tabDeptList.Selected.index)
End Sub

Private Sub chkChange_Click(index As Integer)
    Call UpdateExeSign(0, 0)
    chkAll.Value = 0
    
    Call SetFilter
End Sub

Private Sub chkCheck_Click()
    Dim lngRow As Long
    
    With Me.vsfMedis
        For lngRow = 1 To .rows - 1
            If mcondition.strTransStep = M_STR_CALSS_AUDIT Then
                .TextMatrix(lngRow, .ColIndex("��־")) = IIf(Me.chkCheck.Value = 1, "1", "0")
                .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = IIf(Me.chkCheck.Value = 1, Me.ImgList.ListImages(3).Picture, Nothing)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
            Else
                If nvl(.TextMatrix(lngRow, .ColIndex("���id")), "0") <> "00" And Val(.TextMatrix(lngRow, .ColIndex("��ҩ��־"))) <> 1 Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(chkCheck.Value = 0, 0, -1)
                End If
            End If
        Next
    End With
End Sub

Private Sub chkResult_Click(index As Integer)
    Call LoadVsfMedi(mstr�ϴ�IDS, True)
End Sub

Private Sub chkSure_Click(index As Integer)
    '�л�ȷ��״̬ʱ�����ѡ��־
    Call UpdateExeSign(0, 0)
    chkAll.Value = 0
    
    Call SetFilter
End Sub

Private Sub chkPrint_Click(index As Integer)
    '�л�ȷ��״̬ʱ�����ѡ��־
    Call UpdateExeSign(0, 0)
    chkAll.Value = 0
    
    Call SetFilter
End Sub

Private Sub SetBeach()
    Dim lngRow As Long
    Dim strInput As String
    Dim lng��ҩid As Long
    Dim arrExecute As Variant
    
    With Me.vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩid"))) <> 0 Then
                mrsTrans.Filter = "��ҩid=" & Val(.TextMatrix(lngRow, .ColIndex("��ҩid")))
                If mrsTrans!�ֹ��������� <> 1 And mrsTrans!�Ƿ�������� = 1 Then
                    .TextMatrix(lngRow, .ColIndex("��ҩ����")) = mrsTrans!����ҩ����
                    If lng��ҩid <> Val(.TextMatrix(lngRow, .ColIndex("��ҩid"))) Then
                        lng��ҩid = Val(.TextMatrix(lngRow, .ColIndex("��ҩid")))
                        
                        If zlStr.nvl(mrsTrans!����ҩ����) = "" Then
                            strInput = IIf(strInput = "", "", strInput & "|") & mrsTrans!��ҩid & ",:" & zlStr.nvl(mrsTrans!���ȼ�)
                        Else
                            strInput = IIf(strInput = "", "", strInput & "|") & mrsTrans!��ҩid & "," & Mid(mrsTrans!����ҩ����, 1, IIf(Len(mrsTrans!����ҩ����) = 0, 0, Len(mrsTrans!����ҩ����) - 1)) & ":" & zlStr.nvl(mrsTrans!���ȼ�)
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    On Error GoTo errHandle
    
    arrExecute = GetArrayByStr(strInput, 3950, "|")
    For lngRow = 0 To UBound(arrExecute)
        If mrsPRI.RecordCount > 0 Or mrsVol.RecordCount > 0 Then
            gstrSQL = "Zl_��Һ��ҩ��¼_����("
            '��ҩID,����
            gstrSQL = gstrSQL & "'" & arrExecute(lngRow) & "'"
            gstrSQL = gstrSQL & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
        End If
    Next

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub SetLock(ByVal intType As Integer, ByVal str��ҩid As String, Optional ByVal blnRow As Boolean)
'��ָ����ҽ�����������
'intType:1-����,0-����
'str��ҩid:��ҩid���ַ���
    Dim arrExecute As Variant
    Dim blnBeginTrans As Boolean
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    If str��ҩid = "" Then
        If mrsTrans Is Nothing Then Exit Sub
        With mrsTrans
            .Filter = "ִ�б�־=1"
            .Sort = "����,��ҩ����,סԺ��"
            If .RecordCount > 0 Then
                .MoveFirst
            Else
                Exit Sub
            End If
            
            Do While Not .EOF
                If InStr(1, "," & str��ҩid & ",", "," & !��ҩid & ",") = 0 Then
                    str��ҩid = IIf(str��ҩid = "", "", str��ҩid & ",") & !��ҩid
                End If
                .MoveNext
            Loop
        End With
    End If
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(str��ҩid, 3950, ",")
    For lngRow = 0 To UBound(arrExecute)
        gstrSQL = "Zl_��Һ��ҩ��¼_����("
        '��ҩID
        gstrSQL = gstrSQL & "'" & arrExecute(lngRow) & "'"
        '�Ƿ�����
        gstrSQL = gstrSQL & "," & intType
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "SetLock-����")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    If Not blnRow Then
        With Me.vsfTrans
            For lngRow = 1 To .rows - 1
                If InStr(1, "," & str��ҩid & ",", "," & .TextMatrix(lngRow, .ColIndex("��ҩid")) & ",") > 0 And Val(.TextMatrix(lngRow, .ColIndex("��ҩid"))) > 0 Then
                    .TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = intType
                    .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = IIf(.TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = "1", Me.ImgList.ListImages(5).Picture, Me.ImgList.ListImages(6).Picture)
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                End If
            Next
        End With
    End If
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub DelTransRec()
    'ִ�й��ܺ���¼�¼����ɾ��������ѡ��ļ�¼
    Dim lngRow As Long
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                    mrsTrans.Filter = "��ҩID=" & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                    Do While Not mrsTrans.EOF
                        mrsTrans.Delete
                        mrsTrans.Update
                        mrsTrans.MoveNext
                    Loop
                End If
            End If
        Next
    End With
End Sub
Private Sub RefreshDeptList(ByVal index As Integer)
    If mblnActive = True And vsfDept(index).Visible = True Then vsfDept(index).SetFocus
    mlng��ɨ�� = 0
    mlngδɨ�� = 0
    If Me.cboType.ListCount <> 0 Then
        Me.cboType.ListIndex = 0
    End If
    
    '�����Һ����ϸ�б�
    Call ClearDetailList
    '�����Һ����Ƭ
    mfrmPIVCard.ClearCard
    '��֯��ѯ����
    Call GetCondition
    'ȡ������������Ӧ����Һ������
    Call GetTransCount(CDate(mcondition.strTransStartTime), CDate(mcondition.strTransEndTime))
    '��ʾ������������Ӧ����Һ��������
    Call ShowDeptTrans(index, IIf(index = CNUMWORK, tabWork.Selected.Tag, tbcLook.Selected.Tag))
    
    chkAllDept.Value = 0
End Sub

Public Sub RefreshDetailList(ByVal index As Long)
    'ˢ�º���ʾ��Һ�����б�
    Dim str����id As String
    Dim i As Integer
    Dim bln�б� As Boolean
    Dim bln��Ƭ As Boolean
    Dim strIDS As String
    
    On Error GoTo errHandle
    
    Call AviShow(Me)
    
    chkAll.Enabled = False
    chkAll.Value = 0
    Me.chkCheck.Value = 0
    Me.lblVolu.Caption = "������0"
    Me.lblMsg.Visible = True
    Me.lblMsg.Caption = ""
    
    Call ClearDetailList
    Call mfrmPIVCard.ClearCard
    
    If vsfDept(index).Visible Then vsfDept(index).SetFocus

    If Not tbcDetail.Item(mDetailType.��Һ����Ƭ).Selected Then
        tbcDetail.Item(mDetailType.��Һ���б�).Selected = True
    End If
    
    mstr�ϴβ���ID = ""
    mstr�ϴ�IDS = ""
    
    With vsfDept(index)
        For i = 1 To .rows - 1
            If Val(.TextMatrix(i, .ColIndex("����id"))) > 0 And Val(.TextMatrix(i, .ColIndex("ѡ��"))) = -1 Then
                strIDS = IIf(strIDS = "", "", strIDS & ",") & .TextMatrix(i, .ColIndex("��¼id"))
                str����id = IIf(str����id = "", "", str����id & ",") & .TextMatrix(i, .ColIndex("����id"))
            End If
        Next
    End With
    
    If strIDS = "" Then Call AviShow(Me, False): Exit Sub
    
    If Not mParams.blnFilter Then
        mblnFilter = False
        Call zlGetMediNum(CDate(mcondition.strTransStartTime), CDate(mcondition.strTransEndTime), str����id, mcondition.strTransStep)
        mblnFilter = True
    End If
    
    mstr�ϴβ���ID = str����id
    mstr�ϴ�IDS = strIDS
    
    Call GetCondition
    
    If mParams.bln��� And (mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) Then
        Me.tbcDetail.Item(mDetailType.��Һ���б�).Caption = "����ҽ���б�"
        Call LoadVsfMedi(strIDS)
        bln��Ƭ = mfrmPIVCard.ShowDetailCard(mrsMedi, mstr����, False, mParams.intCount, mParams.bln��������, mParams.bln�������, mcondition.strTransStep, mParams.bln���)
    Else
        Me.tbcDetail.Item(mDetailType.��Һ���б�).Caption = "��Һ���б�"
        '����ȡ��ѡ������Ӧ����Һ����ϸ
        Call LoadTrans(strIDS, mcondition.strTransStep, Val(vsfTrans.Tag), Val(fraDetailCtr.Tag))
         '��״̬����ʾ��ѡ�Ĳ�������Һ������
'        Call GetCount
'        mrsTrans.Filter = ""
'
'        '��ʾ��Һ����Ƭ
'        bln��Ƭ = mfrmPIVCard.ShowDetailCard(mrsTrans, mstr����, mcondition.strTransStep = M_STR_CALSS_PREPARE, mParams.intCount, mParams.bln��������, mParams.bln�������, mcondition.strTransStep, mParams.bln���)
'        '��ʾ��Һ����ϸ�б�
'        bln�б� = ShowTrans(index)
'        '��ʾ��Һ��ҩƷ�����б�
'        Call ShowSumDrug

        Call SetFilter
       
        If bln�б� And bln��Ƭ Then
            chkAll.Enabled = True
        End If
        
        If mParams.blnFilter Then Call zlGetMediNumNew
    End If
    
    Call AviShow(Me, False)
    Exit Sub
errHandle:
    Call AviShow(Me, False)
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlGetMediNum(ByVal dateBegin As Date, ByVal dateEnd As Date, ByVal str����ids As String, ByVal int����״̬ As Integer)
    Dim rstemp As Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle

    strTmp = " Select distinct a.ƿǩ��, c.No, c.ҩƷid, f.����" & vbNewLine & _
        "              From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C, ҩƷ���� D, ҩƷ��� E, �շ���ĿĿ¼ F, ����ҽ����¼ M, סԺ���ü�¼ N" & vbNewLine & _
        "              Where a.ִ��ʱ�� Between [1] And [2] And a.Id = b.��¼id And b.�շ�id = c.Id And" & vbNewLine & _
        "                    c.ҩƷid = e.ҩƷid And d.��ý <> 1 And a.����id = [3] And e.ҩ��id = d.ҩ��id And m.���id = a.ҽ��id And m.Id = n.ҽ����� And" & vbNewLine & _
        "                    c.����id = n.Id And e.ҩƷid = f.Id And a.���˲���id In (Select Column_Value From Table(Cast(f_Str2list([4]) As Zltools.t_Strlist))) And ����״̬ = [5] "
    strTmp = strTmp & " Union All " & Replace(strTmp, "סԺ���ü�¼", "������ü�¼")
    
    gstrSQL = "Select ҩƷid, ����, ����" & vbNewLine & _
        "From (Select ҩƷid, ����, Count(ƿǩ��) ����" & vbNewLine & _
        "       From (" & strTmp & ")" & vbNewLine & _
        "       Group By ҩƷid, ����" & vbNewLine & _
        "       Order By ���� Desc) where rownum<4 "
        
       
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��Ϣ", dateBegin, dateEnd, mParams.lng��������, str����ids, int����״̬)
        
    cboMedi.Clear
    Me.cboMedi.AddItem "<ȫ��>"
    Do While Not rstemp.EOF
        Me.cboMedi.AddItem rstemp!���� & IIf(mParams.blnFilter, "", "(" & rstemp!���� & ")")
        Me.cboMedi.ItemData(Me.cboMedi.ListCount - 1) = rstemp!ҩƷID
        
'        mParams.str����ҩƷ = IIf(mParams.str����ҩƷ = "", "", mParams.str����ҩƷ & ",") & rsTemp!ҩƷID
        rstemp.MoveNext
    Loop
    Me.cboMedi.Text = "<ȫ��>"
        
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub zlGetMediNumNew()
    Dim rsFilter As Recordset
    Dim lngҩƷid As Long
    Dim strҩƷ���� As String
    Dim lngCount As Long
    Dim lng��ҩid As Long
    
    On Error GoTo errHandle
    
    Set rsFilter = New ADODB.Recordset
    With rsFilter
        If .State = 1 Then .Close
        
        '�ü�¼��Ӧ����Һ��ҩ��¼��Ϣ
        .Fields.Append "ҩƷid", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    cboMedi.Clear
    Me.cboMedi.AddItem "<ȫ��>", 0
    Me.cboMedi.Text = "<ȫ��>"
     
    With mrsTrans
        .Filter = "��ý<>1 And �Ƿ���=1"
        If .RecordCount = 0 Then Exit Sub
        
        .Sort = "ҩƷid,��ҩid"
        
        Do While Not .EOF
            If lng��ҩid <> mrsTrans!��ҩid Then
                If lngҩƷid <> mrsTrans!ҩƷID Then
                    If lngҩƷid <> 0 Then
                        rsFilter.AddNew
                        rsFilter!ҩƷID = lngҩƷid
                        rsFilter!ҩƷ���� = strҩƷ����
                        rsFilter!���� = lngCount
                        
                        rsFilter.Update
                    End If
                
                    lngҩƷid = mrsTrans!ҩƷID
                    strҩƷ���� = mrsTrans!ҩƷ����
                    lngCount = 1
                Else
                    lngCount = lngCount + 1
                End If
            End If
            
            lng��ҩid = mrsTrans!��ҩid
                        
            .MoveNext
            
            If .EOF Then
                rsFilter.AddNew
                rsFilter!ҩƷID = lngҩƷid
                rsFilter!ҩƷ���� = strҩƷ����
                rsFilter!���� = lngCount
                
                rsFilter.Update
            End If
        Loop
    End With
    
    
    
    rsFilter.Filter = ""
    rsFilter.Sort = "���� Desc"
    Do While Not rsFilter.EOF
        Me.cboMedi.AddItem rsFilter!ҩƷ���� & "(" & rsFilter!���� & ")"
        Me.cboMedi.ItemData(Me.cboMedi.NewIndex) = rsFilter!ҩƷID
        
'        mParams.str����ҩƷ = IIf(mParams.str����ҩƷ = "", "", mParams.str����ҩƷ & ",") & rsTemp!ҩƷID
        rsFilter.MoveNext
    Loop
       
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub SelectBatch(ByVal lngType As Long, ByVal index As Long)
    Dim lngRow As Long
    Dim strCompare As String
    Dim str��ҩid As String
    Dim i As Integer
    Dim str���id As String
    Dim intFirst As Integer
    Dim datCur As Date
    Dim lng��ҩid As Long
    
    With vsfTrans
        If tbcDetail.Item(mDetailType.��Һ����Ƭ).Selected Then
            str��ҩid = mfrmPIVCard.ChooseOne

            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("��ҩID")) = str��ҩid Then
                    .Row = i
                    Exit For
                End If
            Next
        End If
        
        str��ҩid = ";"
        
        If .Row = 0 Then Exit Sub
        
        If .TextMatrix(.Row, .ColIndex("��ҩID")) = "" Then Exit Sub
        
        .Redraw = flexRDNone
        
        Select Case lngType
            Case conMenu_Oper_Select_SelRow
                'ѡ��ǰ��
                str��ҩid = str��ҩid & .TextMatrix(.Row, .ColIndex("��ҩID")) & ";"
                
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("��ҩID")) = .TextMatrix(.Row, .ColIndex("��ҩID")) Then
                        If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = 0 Then
                            .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1
                        End If
                        str��ҩid = str��ҩid & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ";"
                    End If
                Next
            Case conMenu_Oper_Select_SelBatch
                'ѡ��ǰ����
                strCompare = .TextMatrix(.Row, .ColIndex("��ҩ����"))
                
                If strCompare <> "" Then
                    For lngRow = 1 To .rows - 1
                        If .TextMatrix(lngRow, .ColIndex("��ҩID")) <> "" Then
                            If .TextMatrix(lngRow, .ColIndex("��ҩ����")) = strCompare Then
                                If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = 0 Then
                                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1
                                End If
                                str��ҩid = str��ҩid & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ";"
                            End If
                        End If
                    Next
                End If
            Case conMenu_Oper_Select_SelDept, conMenu_Oper_Select_CancleSelDept
                'ѡ��ǰ����
                strCompare = .TextMatrix(.Row, .ColIndex("����"))
                
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("��ҩID")) <> "" Then
                        If .TextMatrix(lngRow, .ColIndex("����")) = strCompare Then
                            .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(lngType = conMenu_Oper_Select_SelDept, -1, 0)
                            str��ҩid = str��ҩid & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ";"
                        End If
                    End If
                Next
                
            Case conMenu_Oper_Select_SelPati, conMenu_Oper_Select_CancleSelPati
                'ѡ��ǰ����
                strCompare = .TextMatrix(.Row, .ColIndex("����")) & .TextMatrix(.Row, .ColIndex("����")) & .TextMatrix(.Row, .ColIndex("����")) & .TextMatrix(.Row, .ColIndex("����"))
                
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("��ҩID")) <> "" Then
                        If .TextMatrix(lngRow, .ColIndex("����")) & .TextMatrix(lngRow, .ColIndex("����")) & .TextMatrix(lngRow, .ColIndex("����")) & .TextMatrix(lngRow, .ColIndex("����")) = strCompare Then
                            .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(lngType = conMenu_Oper_Select_SelPati, -1, 0)
                            str��ҩid = str��ҩid & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ";"
                        End If
                    End If
                Next
            Case conMenu_Oper_Select_SelSendNo
                'ѡ��ǰ��ҩ��
                If mcondition.strTransStep = M_STR_CALSS_PREPARE Then Exit Sub
                strCompare = .TextMatrix(.Row, .ColIndex("��ҩ����"))
                
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("��ҩID")) <> "" Then
                        If .TextMatrix(lngRow, .ColIndex("��ҩ����")) = strCompare Then
                            If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = 0 Then
                                .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1
                            End If
                            str��ҩid = str��ҩid & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ";"
                        End If
                    End If
                Next
            Case conMenu_Oper_Select_SelMed
            'ѡ�����еĿ���ҩ��
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("����ҩ��")) = 1 Then
                        If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = "0" Then
                            .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1
                        End If
                        str��ҩid = str��ҩid & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ";"
                    End If
                Next
            Case conMenu_Oper_Select_SelAll
                'ѡ��������
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("��ҩID")) <> "" Then
                        If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = "0" Then
                            .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1
                        End If
                        str��ҩid = str��ҩid & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ";"
                    End If
                Next
            Case conMenu_Oper_Bag_Batch
                strCompare = .TextMatrix(.Row, .ColIndex("��ҩ����"))
                datCur = Sys.Currentdate
                
                If strCompare <> "" Then
                    For lngRow = 1 To .rows - 1
                        intFirst = 0
                        If .TextMatrix(lngRow, .ColIndex("��ҩID")) <> "" Then
                            If .TextMatrix(lngRow, .ColIndex("��ҩ����")) = strCompare Then
                                If InStr("|" & str���id, "|" & .TextMatrix(lngRow, .ColIndex("��ҩID"))) < 1 Then
                                    str���id = IIf(str���id = "", "", str���id & "|") & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ",2"
                                End If
                                
                                '������Һ(���)ͼ��
                                .Col = .ColIndex("���")
                                .Cell(flexcpPicture, lngRow, .ColIndex("���")) = picPacker(2).Picture
                                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("���")) = flexPicAlignCenterCenter
                                .TextMatrix(lngRow, .ColIndex("�Ƿ���")) = 2
                                
                                mrsTrans.Filter = "��ҩID=" & Val(.TextMatrix(.Row, .ColIndex("��ҩID")))
                                Do While Not mrsTrans.EOF
                                    intFirst = intFirst + 1
                                    mrsTrans!�Ƿ��� = Val(.TextMatrix(.Row, .ColIndex("�Ƿ���")))
                                    
                                    If mcondition.strTransStep = M_STR_CALSS_DOSAGE And intFirst = 1 And .TextMatrix(.Row, .ColIndex("�Ƿ���")) > 0 Then
                                        mintCountPack = mintCountPack + IIf(IIf(IsNull(mrsTrans!��ҩʱ��), "", Format(mrsTrans!��ҩʱ��, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!���ʱ��), "", Format(mrsTrans!���ʱ��, "YYYY-MM-DD HH:MM:SS")), 0, 1)
                                    Else
                                        If IIf(IsNull(mrsTrans!��ҩʱ��), "", Format(mrsTrans!��ҩʱ��, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!���ʱ��), "", Format(mrsTrans!���ʱ��, "YYYY-MM-DD HH:MM:SS")) Then
                                            mintCountPack = mintCountPack - 1
                                        End If
                                    End If
                                    
                                    mrsTrans!���ʱ�� = IIf(.TextMatrix(.Row, .ColIndex("�Ƿ���")) = 0, "", datCur)
                                    mrsTrans.Update
                                    mrsTrans.MoveNext
                                Loop
                                
                                mfrmPIVCard.PackCard Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))), 2
                            End If
                        End If
                    Next
                End If
                
                Call GetCount
                
                If str���id <> "" Then
                    gstrSQL = "Zl_��Һ��ҩ��¼_���("
                    '��ҩID,���
                    gstrSQL = gstrSQL & "'" & str���id & "'"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
                End If
            Case conMenu_Oper_Bag_All
                datCur = Sys.Currentdate
                
                '���´��ͼ��
                With Me.vsfTrans
                    .Col = .ColIndex("���")
                    .Cell(flexcpPicture, 1, .ColIndex("���"), .rows - 1, .ColIndex("���")) = picPacker(2).Picture
                    .Cell(flexcpPictureAlignment, 1, .ColIndex("���"), .rows - 1, .ColIndex("���")) = flexPicAlignCenterCenter
                    .Cell(flexcpText, 1, .ColIndex("�Ƿ���"), .rows - 1, .ColIndex("�Ƿ���")) = 2
                
                    For lngRow = 1 To .rows - 1
                        If lng��ҩid <> Val(.TextMatrix(lngRow, .ColIndex("��ҩid"))) And Val(.TextMatrix(lngRow, .ColIndex("��ҩid"))) <> 0 Then
                            lng��ҩid = Val(.TextMatrix(lngRow, .ColIndex("��ҩid")))
                            If InStr("|" & str���id, "|" & .TextMatrix(lngRow, .ColIndex("��ҩID"))) < 1 Then
                                str���id = IIf(str���id = "", "", str���id & "|") & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ",2"
                            End If
                            
                            mrsTrans.Filter = "��ҩID=" & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                            
                            '�ı��ڲ����ݼ���ֵ
                            mrsTrans!�Ƿ��� = Val(.TextMatrix(lngRow, .ColIndex("�Ƿ���")))
                            
                            If mcondition.strTransStep = M_STR_CALSS_DOSAGE And .TextMatrix(lngRow, .ColIndex("�Ƿ���")) > 0 Then
                                mintCountPack = mintCountPack + IIf(IIf(IsNull(mrsTrans!��ҩʱ��), "", Format(mrsTrans!��ҩʱ��, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!���ʱ��), "", Format(mrsTrans!���ʱ��, "YYYY-MM-DD HH:MM:SS")), 0, 1)
                            Else
                                If IIf(IsNull(mrsTrans!��ҩʱ��), "", Format(mrsTrans!��ҩʱ��, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!���ʱ��), "", Format(mrsTrans!���ʱ��, "YYYY-MM-DD HH:MM:SS")) Then
                                    mintCountPack = mintCountPack - 1
                                End If
                            End If
                            
                            mrsTrans!���ʱ�� = IIf(.TextMatrix(lngRow, .ColIndex("�Ƿ���")) = 0, "", datCur)
                            mrsTrans.Update
                            mrsTrans.MoveNext
                        End If
                    Next
                    
                    Call GetCount
                
                    If str���id <> "" Then
                        gstrSQL = "Zl_��Һ��ҩ��¼_���("
                        '��ҩID,���
                        gstrSQL = gstrSQL & "'" & str���id & "'"
                        gstrSQL = gstrSQL & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
                    End If
                End With
        End Select
        
        Call mfrmPIVCard.BatchChoose(str��ҩid)
            
        .Redraw = flexRDDirect
        
        '�������ݼ�
        If lngType = conMenu_Oper_Select_SelRow Then
            Call UpdateExeSign(Val(.TextMatrix(.Row, .ColIndex("��ҩID"))), IIf(Val(.TextMatrix(.Row, .ColIndex("ѡ��"))) = -1, 1, 0))
        ElseIf lngType = conMenu_Oper_Select_SelAll Then
            Call UpdateExeSign(0, 1)
        Else
            Call UpdateExeSign(-1, index)
        End If
    End With
End Sub

Private Sub SetCommand()
    '���ò˵�����ݰ�ť
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    'ȡ����ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Cancel, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Cancel, , True)

    If Not cbrMenu Is Nothing Then
        cbrMenu.Visible = False
        If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then cbrMenu.Visible = (mPrives.blnȡ����� And mParams.bln���)
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then cbrMenu.Visible = mPrives.blnȡ����ҩ
        If mcondition.strTransStep = M_STR_CALSS_SEND Then cbrMenu.Visible = mPrives.blnȡ����ҩ
        If mcondition.strTransStep = M_STR_CALSS_SENDED Then cbrMenu.Visible = mPrives.blnȡ������
        
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            cbrMenu.Caption = "ȡ����ҩ(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
            cbrMenu.Caption = "ȡ����ҩ(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SENDED Then
            cbrMenu.Caption = "ȡ������(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
            cbrMenu.Caption = "ȡ�����(&C)"
        End If
    End If
    
    If Not cbrControl Is Nothing Then
        cbrControl.Visible = False
        If mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then cbrControl.Visible = (mPrives.blnȡ����� And mParams.bln���)
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then cbrControl.Visible = mPrives.blnȡ����ҩ
        If mcondition.strTransStep = M_STR_CALSS_SEND Then cbrControl.Visible = mPrives.blnȡ����ҩ
        If mcondition.strTransStep = M_STR_CALSS_SENDED Then cbrControl.Visible = mPrives.blnȡ������
        
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            cbrControl.Caption = "ȡ����ҩ(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
            cbrControl.Caption = "ȡ����ҩ(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SENDED Then
            cbrControl.Caption = "ȡ������(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
            cbrControl.Caption = "ȡ�����(&C)"
        End If
    End If

    '��ӡƿǩ
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButtonPopup, mconMenu_File_PIVA_BillPrintLable, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButtonPopup, mconMenu_File_PIVA_BillPrintLable, , True)
    
    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mParams.blnƿǩ�ֹ���ӡ = True And (mcondition.strTransStep <> M_STR_CALSS_AUDIT And mcondition.strTransStep <> M_STR_CALSS_PASSEDAUDIT And mcondition.strTransStep <> M_STR_CALSS_FAILAUDIT))
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mParams.blnƿǩ�ֹ���ӡ = True And (mcondition.strTransStep <> M_STR_CALSS_AUDIT And mcondition.strTransStep <> M_STR_CALSS_PASSEDAUDIT And mcondition.strTransStep <> M_STR_CALSS_FAILAUDIT))
    
    
    '��ӡ��ҩ��
'    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_File_PIVA_BillPrintWait, , True)
'    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_File_PIVA_BillPrintWait, , True)

'    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (Val(cmdNext.Tag) = mType.��Һ�� And mcondition.strTransStep = M_STR_CALSS_PREPARE)
'    If Not cbrControl Is Nothing Then cbrControl.Visible = (Val(cmdNext.Tag) = mType.��Һ�� And mcondition.strTransStep = M_STR_CALSS_PREPARE)
    
    '���,�ܾ�ҽ����ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Approve, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Approve, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_AUDIT And mPrives.bln�˲�ȷ�� And mParams.bln���)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_AUDIT And mPrives.bln�˲�ȷ�� And mParams.bln���)
    
    '�Ű�����
    Set cbrMenu = cbsMain.FindControl(, mconMenu_PlanPopup)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = mPrives.bln�Ű�����
    
    '����,������ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Lock, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Lock, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE)
    
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_UnLock, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_UnLock, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE)
    
    '�������ΰ�ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Beach, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Beach, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln��ҩȷ��)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln��ҩȷ��)

    'ȷ�ϵ�����ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, MCONMENU_EDIT_PIVA_SURE, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_EDIT_PIVA_SURE, , True)
    
    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln��ҩȷ��)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln��ҩȷ��)
    
    '��ҩ��ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Prepare, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Prepare, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln��ҩȷ��)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln��ҩȷ��)

    '��ҩ��ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Dosage, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Dosage, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_DOSAGE And mPrives.bln��ҩȷ��)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_DOSAGE And mPrives.bln��ҩȷ��)
    
    '���Ͱ�ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Send, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Send, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_SEND And mPrives.bln����ȷ��)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_SEND And mPrives.bln����ȷ��)
    
    'ȷ�Ͼܾ���ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, MCONMENU_EDIT_PIVA_REFUSE, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_EDIT_PIVA_REFUSE, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_REFUSETOSIGN And mPrives.blnȷ�Ͼܾ�)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_REFUSETOSIGN And mPrives.blnȷ�Ͼܾ�)
        
    '���ʰ�ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_ReVerify, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_ReVerify, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_VERIFY And mPrives.bln�������)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_VERIFY And mPrives.bln�������)
    
    'ɾ����ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Delete, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Delete, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_INVALID)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_INVALID)
    
    '���������ť
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButtonPopup, conMenu_Oper_Bag, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButtonPopup, conMenu_Oper_Bag, , True)
    
    If mParams.bln������� Then
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
    Else
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
    End If
End Sub

Private Sub SetTabColor(ByVal tbcObj As TabControl)
    '���÷�ҳ��ť����ɫ
    Dim intTabIndex As Integer
    Dim strCount As String

    With tbcObj
        For intTabIndex = 0 To .ItemCount - 1
            If .Item(intTabIndex).Selected = True Then
                .Item(intTabIndex).Color = CSTCOLOR_COMMAND
            Else
                strCount = Mid(.Item(intTabIndex).Caption, InStr(1, .Item(intTabIndex).Caption, "(") + 1)
                strCount = Mid(strCount, 1, InStr(1, strCount, ")") - 1)
                If Val(strCount) > 0 Then
                    .Item(intTabIndex).Color = CSTCOLOR_RECORDS
                Else
                    .Item(intTabIndex).Color = CSTCOLOR_NORECORDS
                End If
            End If
        Next
    End With
End Sub

Private Sub SetListBar()
    '������ϸ�б�ҳ��ѡ����ʾ��ͬ��ѡ������
    Select Case tbcDetail.Selected.index
        Case mDetailType.��Һ���б�
            chkDept.Visible = False
            chkPack.Visible = False
            chkAll.Visible = True
            chkType(0).Visible = True
            chkType(1).Visible = True
            
            If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
                chkSendType(0).Visible = True
                chkSendType(1).Visible = True
            Else
                chkSendType(0).Visible = False
                chkSendType(1).Visible = False
            End If
            
            Me.lblBatch.Visible = True
            Me.cboBatch.Visible = True
            Me.lblLevel.Visible = True
            Me.cboLevel.Visible = True
            
            Me.cboFrequency.Visible = True
            Me.lblFrequency.Visible = True
            
            lblMedi.Visible = True
            cboMedi.Visible = True
            
            Me.lblVolu.Visible = True
            
            chkPrint(0).Visible = True
            chkPrint(1).Visible = True
            
            chkChange(0).Visible = True
            chkChange(1).Visible = True
            
            chkSure(0).Visible = mcondition.strTransStep = M_STR_CALSS_PREPARE
            chkSure(1).Visible = mcondition.strTransStep = M_STR_CALSS_PREPARE
            
            optShowType(0).Visible = True
            optShowType(1).Visible = True
            
            lblDosType.Visible = True
            cboDosType.Visible = True
        Case mDetailType.ҩƷ�����б�
            chkDept.Visible = True
            chkPack.Visible = True
            chkAll.Visible = False
            chkType(0).Visible = False
            chkType(1).Visible = False
            chkSendType(0).Visible = False
            chkSendType(1).Visible = False
            optShowType(0).Visible = False
            optShowType(1).Visible = False
            
            Me.lblBatch.Visible = False
            Me.cboBatch.Visible = False
            Me.lblLevel.Visible = False
            Me.cboLevel.Visible = False
            
            chkSure(0).Visible = False
            chkSure(1).Visible = False
            
            chkPrint(0).Visible = False
            chkPrint(1).Visible = False
            
            chkChange(0).Visible = False
            chkChange(1).Visible = False
            
            Me.cboFrequency.Visible = False
            Me.lblFrequency.Visible = False
            
            lblMedi.Visible = False
            cboMedi.Visible = False
            
            lblDosType.Visible = False
            cboDosType.Visible = False
            
            Me.lblVolu.Visible = False
            
            vsfTrans.Visible = False
            vsfSumDrug.Visible = True
    End Select
    
    chkSure(0).Left = IIf(chkType(1).Visible, chkType(1).Left + chkType(1).Width, chkAll.Left + chkAll.Width) + 200
    chkSure(1).Left = chkSure(0).Left + chkSure(0).Width + 50
    
    chkPrint(0).Left = IIf(chkSure(1).Visible, chkSure(1).Left + chkSure(1).Width, chkType(1).Left + chkType(1).Width) + 200
    chkPrint(1).Left = chkPrint(0).Left + chkPrint(0).Width + 50
    
    chkChange(0).Left = IIf(chkPrint(1).Visible, chkPrint(1).Left + chkPrint(1).Width, chkSure(1).Left + chkSure(1).Width) + 200
    chkChange(1).Left = chkChange(0).Left + chkChange(0).Width + 50
    
    chkSendType(0).Left = IIf(chkChange(1).Visible, chkChange(1).Left + chkChange(1).Width, chkPrint(1).Left + chkPrint(1).Width) + 200
    chkSendType(1).Left = chkSendType(0).Left + chkSendType(0).Width + 50
    
    optShowType(0).Left = IIf(chkSendType(1).Visible, chkSendType(1).Left + chkSendType(1).Width, chkChange(1).Left + chkChange(1).Width) + 200
    optShowType(1).Left = optShowType(0).Left + optShowType(0).Width + 50
    
End Sub

Private Sub BillPrint_Sum()
    '��ӡ���ܵ���
    Dim StrDate As String
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_3", Me, _
        "����=" & mcondition.lngCenterID, _
        "��ӡʱ��=" & StrDate, "PrintEmpty=0", 1)
End Sub

Private Sub SetSumDrugColHide()
    '���û���ҩƷ�б��е�����������
    With vsfSumDrug
        .ColHidden(.ColIndex("����")) = (chkDept.Value = 0)
        .ColHidden(.ColIndex("���")) = (chkPack.Value = 0)
        
        If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            .ColHidden(.ColIndex("�������")) = False
        Else
            .ColHidden(.ColIndex("�������")) = True
        End If
    End With
End Sub

Private Sub SetTransColHide()
    '������Һ���ݱ������������

    With vsfTrans
        .ColHidden(.ColIndex("����������")) = (mcondition.strTransStep <> M_STR_CALSS_VERIFY)
        .ColHidden(.ColIndex("��������ʱ��")) = (mcondition.strTransStep <> M_STR_CALSS_VERIFY)
        
        .ColHidden(.ColIndex("��������")) = (mcondition.strTransStep <> M_STR_CALSS_INVALID)
        .ColHidden(.ColIndex("���������")) = (mcondition.strTransStep <> M_STR_CALSS_INVALID)
        .ColHidden(.ColIndex("�������ʱ��")) = (mcondition.strTransStep <> M_STR_CALSS_INVALID)
        
        .ColHidden(.ColIndex("��ҩ����")) = (mcondition.strTransStep <= M_STR_CALSS_PREPARE)
        .ColHidden(.ColIndex("��")) = (mcondition.strTransStep <> M_STR_CALSS_PREPARE)
        
        
        .ColHidden(.ColIndex("��")) = (mcondition.strTransStep > M_STR_CALSS_PREPARE)
        .ColHidden(.ColIndex("ҽ��")) = (mcondition.strTransStep > M_STR_CALSS_PREPARE)
        
        .ColHidden(.ColIndex("����ԭ��")) = (mcondition.strTransStep <> M_STR_CALSS_REFUSETOSIGN)
        .ColHidden(.ColIndex("��")) = (mcondition.strTransStep <> M_STR_CALSS_PREPARE)
        
        .ColHidden(.ColIndex("��")) = (mcondition.strTransStep <> M_STR_CALSS_VERIFY)
        .ColHidden(.ColIndex("ѡ��")) = (mcondition.strTransStep = M_STR_CALSS_VERIFY)
        
        If optShowType(0).Value = True Then
            .ColHidden(.ColIndex("��ҩ��")) = True
            .ColHidden(.ColIndex("��ҩʱ��")) = True
            .ColHidden(.ColIndex("��ҩ��")) = True
            .ColHidden(.ColIndex("��ҩʱ��")) = True
            .ColHidden(.ColIndex("������")) = True
            .ColHidden(.ColIndex("����ʱ��")) = True
            .ColHidden(.ColIndex("ҽ������ʱ��")) = True
        Else
            .ColHidden(.ColIndex("ҽ������ʱ��")) = False
            .ColHidden(.ColIndex("��ҩ��")) = (mcondition.strTransStep <> M_STR_CALSS_DOSAGE)
            .ColHidden(.ColIndex("��ҩʱ��")) = (mcondition.strTransStep <> M_STR_CALSS_DOSAGE)
            .ColHidden(.ColIndex("��ҩ��")) = (mcondition.strTransStep <> M_STR_CALSS_SEND)
            .ColHidden(.ColIndex("��ҩʱ��")) = (mcondition.strTransStep <> M_STR_CALSS_SEND)
            .ColHidden(.ColIndex("������")) = (mcondition.strTransStep <> M_STR_CALSS_SENDED)
            .ColHidden(.ColIndex("����ʱ��")) = (mcondition.strTransStep <> M_STR_CALSS_SENDED)
        End If
    End With
End Sub

Private Sub ShowMedicalRecord()
    '�����ܡ�:���ĵ�ǰ���˵ĵ��Ӳ���

    With vsfMedis
        '���
        If .Row < 1 Then Exit Sub
        
        '���õ��Ӳ������Ľӿ�
        If Not mobjCISJOB Is Nothing Then
            On Error Resume Next
            Call mobjCISJOB.ShowArchive(Me, Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("����id"))), Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("��ҳid"))))
            err.Clear: On Error GoTo 0
        End If
        
    End With
End Sub

Private Sub ShowComment(ByVal intTab As Integer, ByVal strStep As String)
    '��ʾ��ǰ���̵���ʾ��Ϣ
    
    lblHelp.Caption = ""
    
    If intTab = mDetailType.��Һ���б� Then
        If strStep = M_STR_CALSS_PREPARE Then
            If mPrives.bln��ҩȷ�� = False Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��û�а�ҩȷ�ϵ�Ȩ��"
            Else
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ҩȷ��"
                If mParams.int��ҩ���ӡ = 0 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ҩȷ�Ϻ���ʾ��ӡ��ҩ��"
                ElseIf mParams.int��ҩ���ӡ = 1 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ҩȷ�Ϻ��Զ���ӡ��ҩ��"
                Else
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ҩȷ�Ϻ󲻴�ӡ��ҩ��"
                End If
                If mParams.intƿǩ��ҩ���ӡ = 0 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ҩȷ�Ϻ���ʾ��ӡ��ǩ"
                ElseIf mParams.intƿǩ��ҩ���ӡ = 1 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ҩȷ�Ϻ��Զ���ӡ��ǩ"
                End If
            End If
            If mParams.bln�������� = True Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "�����������"
            End If
            If mParams.bln������� = True Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "����������״̬"
            End If
        ElseIf strStep = M_STR_CALSS_DOSAGE Then
            If mPrives.bln��ҩȷ�� = False Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��û����ҩȷ�ϵ�Ȩ��"
            Else
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ҩȷ��"
                If mParams.intƿǩ��ҩ���ӡ = 0 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ҩȷ�Ϻ���ʾ��ӡ��ǩ"
                ElseIf mParams.intƿǩ��ҩ���ӡ = 1 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ҩȷ�Ϻ��Զ���ӡ��ǩ"
                End If
            End If
            If mParams.bln������� = True Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "����������״̬"
            End If
        ElseIf strStep = M_STR_CALSS_SEND Then
            If mPrives.bln����ȷ�� = False Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��û�з���ȷ�ϵ�Ȩ��"
            Else
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "����ȷ��"
                If mParams.int���ͺ��ӡ = 0 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "����ȷ�Ϻ���ʾ��ӡ���͵�"
                ElseIf mParams.int���ͺ��ӡ = 1 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "����ȷ�Ϻ��Զ���ӡ���͵�"
                Else
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "����ȷ�Ϻ󲻴�ӡ���͵�"
                End If
            End If
        ElseIf strStep = M_STR_CALSS_SENDED Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "�ѷ��Ͳ鿴"
        ElseIf strStep = M_STR_CALSS_SIGNED Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ǩ�ղ鿴"
        ElseIf strStep = M_STR_CALSS_REFUSETOSIGN Then
            If mPrives.blnȷ�Ͼܾ� = False Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��û�оܾ�ȷ�ϵ�Ȩ��"
            Else
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "�ܾ�ǩ������鿴"
            End If
        ElseIf strStep = M_STR_CALSS_VERIFY Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��������鿴"
            If mPrives.bln������� = True Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "�����������"
            End If
        ElseIf strStep = M_STR_CALSS_INVALID Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "�����ϲ鿴"
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ʾ��������˵���Һ��"
        ElseIf strStep = M_STR_CALSS_DEVICERETURN Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "ҽ�����˲鿴"
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "��ʾ��ҽ���������ϵ���Һ��"
        ElseIf strStep = M_STR_CALSS_AUDIT Then
            lblHelp.Caption = "��δ��˵�ҽ��������˲���"
        ElseIf strStep = M_STR_CALSS_PASSEDAUDIT Then
            lblHelp.Caption = "����ͨ����˵�ҽ������ȡ������"
        ElseIf strStep = M_STR_CALSS_FAILAUDIT Then
            lblHelp.Caption = "��δͨ����˵�ҽ������ȡ������"
        End If
        
        mfrmPIVCard.LoadHelp lblHelp.Caption
    ElseIf intTab = mDetailType.ҩƷ�����б� Then
        lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "ҩƷ���ܲ鿴;��ɫ��ʾȱҩ"
    End If
    
    
End Sub

Private Sub ShowSumDrug()
    '������ѡ�����Һ��������ҩƷ
    Dim lngRow As Long
    Dim strCurr As String
    Dim dblSum As Double
    Dim dblSendSum As Double
    Dim lng�շ�ID As Long
    
    With vsfSumDrug
        .rows = 1
        .rows = 2
        
        .MergeCells = flexMergeNever
        
        If mrsTrans Is Nothing Then Exit Sub
        
        mrsTrans.Filter = "ִ�б�־=1"
        
        If mrsTrans.RecordCount = 0 Then
            mrsTrans.Filter = ""
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        .rows = 1
        
        If chkDept.Value = 1 And chkPack.Value = 1 Then
            mrsTrans.Sort = "����,�Ƿ���,ҩƷ��������,����,�շ�ID"
            
            Do While Not mrsTrans.EOF
                If strCurr <> mrsTrans!���� & mrsTrans!�Ƿ��� & mrsTrans!ҩƷ�������� & mrsTrans!���� Then
                    lngRow = lngRow + 1
                    .rows = .rows + 1
                    
                    strCurr = mrsTrans!���� & mrsTrans!�Ƿ��� & mrsTrans!ҩƷ�������� & mrsTrans!����
                    dblSum = zlStr.nvl(mrsTrans!����, 0)
                    
                    lng�շ�ID = mrsTrans!�շ�ID
                    dblSendSum = mrsTrans!��ҩ����
                    
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("�Ƿ���")) = mrsTrans!�Ƿ���
                    .TextMatrix(lngRow, .ColIndex("ҩƷ����")) = mrsTrans!ҩƷ��������
                    .TextMatrix(lngRow, .ColIndex("��Ʒ��")) = mrsTrans!��Ʒ��
                    .TextMatrix(lngRow, .ColIndex("Ӣ����")) = mrsTrans!Ӣ����
                    .TextMatrix(lngRow, .ColIndex("���")) = zlStr.nvl(mrsTrans!���)
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(dblSum, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("��ҩ����")) = FormatEx(dblSendSum, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("�������")) = FormatEx(mrsTrans!�������, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("ȱҩ��־")) = IIf(dblSendSum > mrsTrans!�������, 1, 0)
                Else
                    dblSum = dblSum + zlStr.nvl(mrsTrans!����, 0)
                    .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(dblSum, 2) & mrsTrans!��λ
                    
                    If lng�շ�ID <> mrsTrans!�շ�ID Then
                        lng�շ�ID = mrsTrans!�շ�ID
                        dblSendSum = dblSendSum + mrsTrans!��ҩ����
                        .TextMatrix(lngRow, .ColIndex("��ҩ����")) = FormatEx(dblSendSum, 2) & mrsTrans!��λ
                        .TextMatrix(lngRow, .ColIndex("ȱҩ��־")) = IIf(dblSendSum > mrsTrans!�������, 1, 0)
                    End If
                End If
                
                mrsTrans.MoveNext
            Loop
            
            mrsTrans.Filter = ""
        ElseIf chkDept.Value = 1 Then
            mrsTrans.Sort = "����,ҩƷ��������,����,�շ�ID"
            
            Do While Not mrsTrans.EOF
                If strCurr <> mrsTrans!���� & mrsTrans!ҩƷ�������� & mrsTrans!���� Then
                    lngRow = lngRow + 1
                    .rows = .rows + 1
                    
                    strCurr = mrsTrans!���� & mrsTrans!ҩƷ�������� & mrsTrans!����
                    dblSum = zlStr.nvl(mrsTrans!����, 0)
                    
                    lng�շ�ID = mrsTrans!�շ�ID
                    dblSendSum = mrsTrans!��ҩ����
                    
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("ҩƷ����")) = mrsTrans!ҩƷ��������
                    .TextMatrix(lngRow, .ColIndex("��Ʒ��")) = mrsTrans!��Ʒ��
                    .TextMatrix(lngRow, .ColIndex("Ӣ����")) = mrsTrans!Ӣ����
                    .TextMatrix(lngRow, .ColIndex("���")) = zlStr.nvl(mrsTrans!���)
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(dblSum, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("��ҩ����")) = FormatEx(dblSendSum, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("�������")) = FormatEx(mrsTrans!�������, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("ȱҩ��־")) = IIf(dblSendSum > mrsTrans!�������, 1, 0)
                Else
                    dblSum = dblSum + zlStr.nvl(mrsTrans!����, 0)
                    .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(dblSum, 2) & mrsTrans!��λ
                    
                    If lng�շ�ID <> mrsTrans!�շ�ID Then
                        lng�շ�ID = mrsTrans!�շ�ID
                        dblSendSum = dblSendSum + mrsTrans!��ҩ����
                        .TextMatrix(lngRow, .ColIndex("��ҩ����")) = FormatEx(dblSendSum, 2) & mrsTrans!��λ
                        .TextMatrix(lngRow, .ColIndex("ȱҩ��־")) = IIf(dblSendSum > mrsTrans!�������, 1, 0)
                    End If
                End If
                
                mrsTrans.MoveNext
            Loop
        ElseIf chkPack.Value = 1 Then
            mrsTrans.Sort = "�Ƿ���,ҩƷ��������,����,�շ�ID"
            
            Do While Not mrsTrans.EOF
                If strCurr <> mrsTrans!�Ƿ��� & mrsTrans!ҩƷ�������� & mrsTrans!���� Then
                    lngRow = lngRow + 1
                    .rows = .rows + 1
                    
                    strCurr = mrsTrans!�Ƿ��� & mrsTrans!ҩƷ�������� & mrsTrans!����
                    dblSum = zlStr.nvl(mrsTrans!����, 0)
                    
                    lng�շ�ID = mrsTrans!�շ�ID
                    dblSendSum = mrsTrans!��ҩ����
                    
                    .TextMatrix(lngRow, .ColIndex("�Ƿ���")) = mrsTrans!�Ƿ���
                    .TextMatrix(lngRow, .ColIndex("ҩƷ����")) = mrsTrans!ҩƷ��������
                    .TextMatrix(lngRow, .ColIndex("��Ʒ��")) = mrsTrans!��Ʒ��
                    .TextMatrix(lngRow, .ColIndex("Ӣ����")) = mrsTrans!Ӣ����
                    .TextMatrix(lngRow, .ColIndex("���")) = zlStr.nvl(mrsTrans!���)
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(dblSum, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("��ҩ����")) = FormatEx(dblSendSum, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("�������")) = FormatEx(mrsTrans!�������, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("ȱҩ��־")) = IIf(dblSendSum > mrsTrans!�������, 1, 0)
                Else
                    dblSum = dblSum + zlStr.nvl(mrsTrans!����, 0)
                    .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(dblSum, 2) & mrsTrans!��λ
                    
                    If lng�շ�ID <> mrsTrans!�շ�ID Then
                        lng�շ�ID = mrsTrans!�շ�ID
                        dblSendSum = dblSendSum + mrsTrans!��ҩ����
                        .TextMatrix(lngRow, .ColIndex("��ҩ����")) = FormatEx(dblSendSum, 2) & mrsTrans!��λ
                        .TextMatrix(lngRow, .ColIndex("ȱҩ��־")) = IIf(dblSendSum > mrsTrans!�������, 1, 0)
                    End If
                End If
                
                mrsTrans.MoveNext
            Loop
        Else
            mrsTrans.Sort = "ҩƷ��������,����,�շ�ID"
            
            Do While Not mrsTrans.EOF
                If strCurr <> mrsTrans!ҩƷ�������� & mrsTrans!���� Then
                    lngRow = lngRow + 1
                    .rows = .rows + 1
                    
                    strCurr = mrsTrans!ҩƷ�������� & mrsTrans!����
                    dblSum = zlStr.nvl(mrsTrans!����, 0)
                    
                    lng�շ�ID = mrsTrans!�շ�ID
                    dblSendSum = mrsTrans!��ҩ����

                    .TextMatrix(lngRow, .ColIndex("ҩƷ����")) = mrsTrans!ҩƷ��������
                    .TextMatrix(lngRow, .ColIndex("��Ʒ��")) = mrsTrans!��Ʒ��
                    .TextMatrix(lngRow, .ColIndex("Ӣ����")) = mrsTrans!Ӣ����
                    .TextMatrix(lngRow, .ColIndex("���")) = zlStr.nvl(mrsTrans!���)
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                    .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(dblSum, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("��ҩ����")) = FormatEx(dblSendSum, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("�������")) = FormatEx(mrsTrans!�������, 2) & mrsTrans!��λ
                    .TextMatrix(lngRow, .ColIndex("ȱҩ��־")) = IIf(dblSendSum > mrsTrans!�������, 1, 0)
                Else
                    dblSum = dblSum + zlStr.nvl(mrsTrans!����, 0)
                    .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(dblSum, 2) & mrsTrans!��λ
                    
                    If lng�շ�ID <> mrsTrans!�շ�ID Then
                        lng�շ�ID = mrsTrans!�շ�ID
                        dblSendSum = dblSendSum + mrsTrans!��ҩ����
                        .TextMatrix(lngRow, .ColIndex("��ҩ����")) = FormatEx(dblSendSum, 2) & mrsTrans!��λ
                        .TextMatrix(lngRow, .ColIndex("ȱҩ��־")) = IIf(dblSendSum > mrsTrans!�������, 1, 0)
                    End If
                End If
                
                mrsTrans.MoveNext
            Loop
        End If
        
        For lngRow = 1 To .rows - 1
            '��ʶȱҩҩƷ
            If .TextMatrix(lngRow, .ColIndex("ҩƷ����")) <> "" Then
                If .TextMatrix(lngRow, .ColIndex("ȱҩ��־")) = 1 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                End If
            End If
            
            '������Һ(���)ͼ��
            If chkPack.Value = 1 Then
                If Val(.TextMatrix(lngRow, .ColIndex("�Ƿ���"))) > 0 Then
'                    .Row = lngRow
'                    .Col = .ColIndex("���")
'                    .CellPicture = picPacker(1).Picture
'                    .CellPictureAlignment = flexPicAlignCenterCenter
                    .Cell(flexcpPicture, lngRow, .ColIndex("���"), lngRow, .ColIndex("���")) = picPacker(Val(.TextMatrix(lngRow, .ColIndex("�Ƿ���")))).Picture
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("���"), lngRow, .ColIndex("���")) = flexPicAlignCenterCenter
                End If
            End If
        Next
        
        '�ϲ�����
        If chkDept.Value = 1 Then
            .MergeCells = flexMergeRestrictRows
            .MergeCol(.ColIndex("����")) = True
        End If
        
        Call SetSumDrugColHide
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Function ShowTrans(ByVal index As Long) As Boolean
    Dim lngRow As Long
    Dim lng��ҩid As Long
    Dim strMaxNo As String
    Dim j As Integer
    Dim dteCur As Date
    Dim intBlackColor As Integer
    Dim lng����id As Long
    Dim lngҩƷid As Long
    Dim LngID As Long
    Dim dateCurrent As Date
    Dim strSort As String
    Dim lngNum As Long
'    chkAll.Value = 0
    
    vsfTrans.rows = 1
    vsfTrans.rows = 2
    intBlackColor = 2
'    vsfDrug.rows = 1
'    vsfDrug.rows = 2

    mlngδɨ�� = 0
'
    dteCur = Sys.Currentdate
    If mrsTrans Is Nothing Then Exit Function
    If mrsTrans.RecordCount = 0 Then Exit Function
    mrsTrans.MoveFirst
    
    If Val(vsfTrans.Tag) = 1 And mParams.blnByMedi = True Then
        Call MediSort
    Else
        strSort = mParams.strSort
        If InStr(1, strSort, "��ҩ����") > 0 Then
            strSort = Replace(strSort, "��ҩ����", "��ҩ����,���ȼ�")
        End If
        
        If InStr(1, strSort, "����") > 0 Then
            strSort = Replace(strSort, "����", "����,��������")
        End If
        
        mrsTrans.Sort = IIf(strSort <> "", strSort & ",��ҩid", "��ҩid")
    End If
    
    If mrsTrans.RecordCount = 0 Then Exit Function
    
    dateCurrent = Sys.Currentdate
    
    With Me.vsfDept(0)
        For lngNum = 1 To .rows - 1
            If Val(.TextMatrix(lngNum, .ColIndex("����id"))) > 0 And Val(.TextMatrix(lngNum, .ColIndex("ѡ��"))) = -1 Then
                .TextMatrix(lngNum, .ColIndex("����")) = 0
            End If
        Next
    End With
    
    With vsfTrans
        .MergeCells = flexMergeFree
        .Redraw = flexRDNone
        .rows = 1
'        .rows = mrsTrans.RecordCount + 1
        lngRow = 1
        Do While Not mrsTrans.EOF
            .rows = .rows + 1
            If mstrFilter <> "" Then
                If lng��ҩid <> LngID Or LngID = 0 Then
                    LngID = Split(mstrFilter, ",")(0)
                    mrsTrans.Filter = "��ҩid=" & Split(mstrFilter, ",")(0)
                    mstrFilter = Mid(mstrFilter, Len(Split(mstrFilter, ",")(0)) + 2)
                End If
            End If
            
            .MergeCol(.ColIndex("����")) = True
            .MergeRow(lngRow) = False
            
            If lng����id <> mrsTrans!����ID Then
                lng����id = mrsTrans!����ID
                If intBlackColor = 2 Then
                    intBlackColor = 1
                ElseIf intBlackColor = 1 Then
                    intBlackColor = 2
                End If
            End If
            
            If lng��ҩid <> mrsTrans!��ҩid Then
                With Me.vsfDept(0)
                    For lngNum = 1 To .rows - 1
                        If mrsTrans!���� = Mid(.TextMatrix(lngNum, .ColIndex("����")), InStr(1, .TextMatrix(lngNum, .ColIndex("����")), "]") + 1) And Val(.TextMatrix(lngNum, .ColIndex("ѡ��"))) = -1 Then
                            .TextMatrix(lngNum, .ColIndex("����")) = Val(.TextMatrix(lngNum, .ColIndex("����"))) + 1
                        End If
                    Next
                End With
                mlngNum = mlngNum + 1
                mlngδɨ�� = mlngδɨ�� + 1
                If lng��ҩid <> 0 Then
                    .rows = .rows + 1
                    .RowHidden(lngRow) = True
                    For j = 0 To .Cols - 1
                        .TextMatrix(lngRow, j) = "00"
                    Next
                    lngRow = lngRow + 1
                End If
                lng��ҩid = mrsTrans!��ҩid
            Else
                .MergeCol(.ColIndex("ѡ��")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("�Ա�")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("סԺ��")) = True
                .MergeCol(.ColIndex("��ҩ����")) = True
                .MergeCol(.ColIndex("��������")) = True
                .MergeCol(.ColIndex("ִ��ʱ��")) = True
                .MergeCol(.ColIndex("ƿǩ��")) = True
                .MergeCol(.ColIndex("���ȼ�")) = True
                .MergeCol(.ColIndex("�˲���")) = True
                .MergeCol(.ColIndex("�˲�ʱ��")) = True
                .MergeCol(.ColIndex("��ҩ��")) = True
                .MergeCol(.ColIndex("��ҩʱ��")) = True
                .MergeCol(.ColIndex("��ҩ����")) = True
                .MergeCol(.ColIndex("��ҩ��")) = True
                .MergeCol(.ColIndex("��ҩʱ��")) = True
                .MergeCol(.ColIndex("������")) = True
                .MergeCol(.ColIndex("����ʱ��")) = True
                .MergeCol(.ColIndex("����������")) = True
                .MergeCol(.ColIndex("��������ʱ��")) = True
                .MergeCol(.ColIndex("���������")) = True
                .MergeCol(.ColIndex("�������ʱ��")) = True
                .MergeCol(.ColIndex("ҽ������ʱ��")) = True
                .MergeCol(.ColIndex("�Ƿ���")) = True
                .MergeCol(.ColIndex("��ӡ��־")) = True
                .MergeCol(.ColIndex("��ҩID")) = True
                .MergeCol(.ColIndex("����ҩ��")) = True
                .MergeCol(.ColIndex("��ӡ")) = True
                .MergeCol(.ColIndex("ҽ��")) = True
                .MergeCol(.ColIndex("���")) = True
                .MergeCol(.ColIndex("��")) = True
                .MergeCol(.ColIndex("��")) = True
                .MergeCol(.ColIndex("������")) = True
                .MergeCol(.ColIndex("����ԭ��")) = True
                .MergeCol(.ColIndex("��")) = True
                .MergeCol(.ColIndex("��")) = True
                .MergeCol(.ColIndex("ִ��Ƶ��")) = True
                .MergeCol(.ColIndex("����ԭ��")) = True
            End If
            
            
            
                lngҩƷid = mrsTrans!ҩƷID
                .TextMatrix(lngRow, .ColIndex("������")) = intBlackColor
                .TextMatrix(lngRow, 1) = lngRow
                
                .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(mrsTrans!ִ�б�־ = 1, -1, "")
                
                If mParams.intAutoSelect = 1 Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(InStr(1, mstrLastLabel, mrsTrans!ƿǩ��) > 0, -1, "")
                End If
                
                .TextMatrix(lngRow, .ColIndex("��������")) = IIf(IsNull(mrsTrans!��������), "", mrsTrans!��������)
                
                .TextMatrix(lngRow, .ColIndex("��ӡ")) = " "
                .TextMatrix(lngRow, .ColIndex("ҽ��")) = " "
                .TextMatrix(lngRow, .ColIndex("���")) = " "
                .TextMatrix(lngRow, .ColIndex("��")) = " "
                .TextMatrix(lngRow, .ColIndex("��")) = " "
                .TextMatrix(lngRow, .ColIndex("��")) = " "
                .TextMatrix(lngRow, .ColIndex("��")) = " "
                .TextMatrix(lngRow, .ColIndex("��־")) = 0
                .TextMatrix(lngRow, .ColIndex("��ҩ����")) = IIf(zlStr.nvl(mrsTrans!��ҩ����) = "", " ", zlStr.nvl(mrsTrans!��ҩ����))
                .TextMatrix(lngRow, .ColIndex("ԭ����")) = mrsTrans!��ҩ����
                .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                .TextMatrix(lngRow, .ColIndex("����")) = IIf(zlStr.nvl(mrsTrans!����) = "", "<��>", mrsTrans!����)
                .TextMatrix(lngRow, .ColIndex("סԺ��")) = mrsTrans!סԺ��
                .TextMatrix(lngRow, .ColIndex("�Ա�")) = mrsTrans!�Ա�
                .TextMatrix(lngRow, .ColIndex("����")) = mrsTrans!����
                .TextMatrix(lngRow, .ColIndex("ִ��ʱ��")) = mrsTrans!ִ��ʱ��
                .TextMatrix(lngRow, .ColIndex("ƿǩ��")) = mrsTrans!ƿǩ��
                .TextMatrix(lngRow, .ColIndex("����id")) = mrsTrans!����ID
                .TextMatrix(lngRow, .ColIndex("��ҳid")) = mrsTrans!��ҳid
                .TextMatrix(lngRow, .ColIndex("���ȼ�")) = Val(zlStr.nvl(mrsTrans!���ȼ�))
                .TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = zlStr.nvl(mrsTrans!�Ƿ�����, 0)
                                                    
                .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = IIf(.TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = "1", Me.ImgList.ListImages(5).Picture, Me.ImgList.ListImages(6).Picture)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                
                .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = IIf(mrsTrans!�Ƿ�ȷ�ϵ��� = 1, Me.ImgList.ListImages(8).Picture, Nothing)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                                
                .TextMatrix(lngRow, .ColIndex("�˲���")) = mrsTrans!�˲���
                .TextMatrix(lngRow, .ColIndex("�˲�ʱ��")) = mrsTrans!�˲�ʱ��
                .TextMatrix(lngRow, .ColIndex("��ҩ��")) = mrsTrans!��ҩ��
                .TextMatrix(lngRow, .ColIndex("��ҩʱ��")) = mrsTrans!��ҩʱ��
                .TextMatrix(lngRow, .ColIndex("��ҩ����")) = IIf(zlStr.nvl(mrsTrans!��ҩ����) = "", " ", mrsTrans!��ҩ����)
                .TextMatrix(lngRow, .ColIndex("��ҩ��")) = mrsTrans!��ҩ��
                .TextMatrix(lngRow, .ColIndex("��ҩʱ��")) = mrsTrans!��ҩʱ��
                .TextMatrix(lngRow, .ColIndex("������")) = mrsTrans!������
                .TextMatrix(lngRow, .ColIndex("����ʱ��")) = mrsTrans!����ʱ��
                .TextMatrix(lngRow, .ColIndex("����������")) = mrsTrans!����������
                .TextMatrix(lngRow, .ColIndex("��������ʱ��")) = mrsTrans!��������ʱ��
                .TextMatrix(lngRow, .ColIndex("���������")) = mrsTrans!���������
                .TextMatrix(lngRow, .ColIndex("�������ʱ��")) = mrsTrans!�������ʱ��
                .TextMatrix(lngRow, .ColIndex("ҽ������ʱ��")) = mrsTrans!ҽ������ʱ��
                .TextMatrix(lngRow, .ColIndex("����ԭ��")) = mrsTrans!����ԭ��
                .TextMatrix(lngRow, .ColIndex("����ԭ��")) = mrsTrans!����ԭ��
                
                .TextMatrix(lngRow, .ColIndex("�Ƿ���")) = mrsTrans!�Ƿ���
                .TextMatrix(lngRow, .ColIndex("��ӡ��־")) = mrsTrans!��ӡ��־
                .TextMatrix(lngRow, .ColIndex("��ҩID")) = mrsTrans!��ҩid
                
                .TextMatrix(lngRow, .ColIndex("����ҩ��")) = mrsTrans!����ҩ��
                
                '����ҩƷ��Ϣ
                .TextMatrix(lngRow, .ColIndex("ҩƷ����")) = mrsTrans!ҩƷ����
                .TextMatrix(lngRow, .ColIndex("���")) = zlStr.nvl(mrsTrans!���)
                .TextMatrix(lngRow, .ColIndex("��ҩ����")) = IIf(IsNull(mrsTrans!ʵ����ҩ����), "", mrsTrans!ʵ����ҩ����)
                .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(mrsTrans!����, 2) & mrsTrans!������λ
                .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(mrsTrans!����, 2) & mrsTrans!��λ
                .TextMatrix(lngRow, .ColIndex("NO")) = zlStr.nvl(mrsTrans!NO)
                .TextMatrix(lngRow, .ColIndex("����")) = nvl(mrsTrans!����)
                .TextMatrix(lngRow, .ColIndex("ҩƷid")) = mrsTrans!ҩƷID
                .TextMatrix(lngRow, .ColIndex("ִ��Ƶ��")) = mrsTrans!ִ��Ƶ��
                .TextMatrix(lngRow, .ColIndex("��ý")) = mrsTrans!��ý
                .TextMatrix(lngRow, .ColIndex("��Ӧҽ��ID")) = mrsTrans!��Ӧҽ��ID
                
                If mrsTrans!�Ƿ�Ƥ�� = 1 Then
                    .TextMatrix(lngRow, .ColIndex("Ƥ")) = GetƤ�Խ��(Val(mrsTrans!����ID), Val(mrsTrans!ҩ��ID), dateCurrent, CDate(mrsTrans!����ʱ��), mrsTrans!��ҳid)
                End If
                
                .Cell(flexcpPicture, lngRow, .ColIndex("ҽ��"), lngRow, .ColIndex("ҽ��")) = IIf(Format(zlStr.nvl(mrsTrans!ҩʦ���ʱ��), "YYYY-MM-dd") = Format(dteCur, "YYYY-MM-dd"), Me.ImgList.ListImages(2).Picture, Me.ImgList.ListImages(1).Picture)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("ҽ��"), lngRow, .ColIndex("ҽ��")) = flexPicAlignCenterCenter
                
                .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = IIf(mrsTrans!�Ƿ�������� = 1, Me.ImgList.ListImages(7).Picture, Nothing)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                
                If Not gobjPass Is Nothing Then
                    .Cell(flexcpPicture, lngRow, .ColIndex("�����"), lngRow, .ColIndex("�����")) = gobjPass.zlPassSetWarnLight_YF(Val(mrsTrans!�����))
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("�����"), lngRow, .ColIndex("�����")) = flexPicAlignCenterCenter
                End If
                
                '��ʾ[�Ա�ҩ]��־
                If mrsTrans!ִ������ = 5 And mrsTrans!ִ�б�� = 0 Then
                    .Cell(flexcpPicture, lngRow, .ColIndex("ҩƷ����"), lngRow, .ColIndex("ҩƷ����")) = Me.ImgPro.ListImages("�Ա�ҩ").Picture
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("ҩƷ����"), lngRow, .ColIndex("ҩƷ����")) = flexPicAlignLeftCenter
                End If
                
                .Cell(flexcpForeColor, lngRow, .ColIndex("��ҩ����"), lngRow, .ColIndex("��ҩ����")) = IIf(mrsTrans!���α�� = 2, vbRed, IIf(mcondition.strTransStep = M_STR_CALSS_PREPARE And mParams.bln�������� = True, vbBlue, vbBlack))
                lngRow = lngRow + 1
                
                mintCountPack = mintCountPack + IIf(IIf(IsNull(mrsTrans!��ҩʱ��), "", Format(mrsTrans!��ҩʱ��, "YYYY-MM-DD HH:MM:SS")) > IIf(IsNull(mrsTrans!���ʱ��), "", Format(mrsTrans!���ʱ��, "YYYY-MM-DD HH:MM:SS")), 0, 1)
            
            mrsTrans.MoveNext
            
            If mstrFilter <> "" And mrsTrans.EOF Then
                mrsTrans.Filter = ""
                LngID = 0
            End If
        Loop
        
'        'ѡ���мӴ֣���ɫ��ʾ
'        .Cell(flexcpFontBold, 0, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = True
'        .Cell(flexcpForeColor, 0, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = vbBlue
        .Cell(flexcpBackColor, 1, .ColIndex("ѡ��"), .rows - 1, .ColIndex("ѡ��")) = CSTCOLOR_MODIFY
        
        .Cell(flexcpFontBold, 0, .ColIndex("��"), 0, .ColIndex("��")) = True
        .Cell(flexcpForeColor, 0, .ColIndex("��"), 0, .ColIndex("��")) = vbBlue
        
        '�������������ʾ�����ݲ�����ͬ����
        .Cell(flexcpFontBold, 0, .ColIndex("���"), 0, .ColIndex("���")) = IIf((mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE) And mParams.bln������� = True, True, False)
        .Cell(flexcpForeColor, 0, .ColIndex("���"), .rows - 1, .ColIndex("���")) = IIf((mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE) And mParams.bln������� = True, vbBlue, vbBlack)
        .Cell(flexcpBackColor, 1, .ColIndex("���"), .rows - 1, .ColIndex("���")) = IIf((mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE) And mParams.bln������� = True, CSTCOLOR_MODIFY, CSTCOLOR_UNMODIFY)
       
        .Cell(flexcpFontBold, 0, .ColIndex("��ҩ����"), 0, .ColIndex("��ҩ����")) = IIf(mcondition.strTransStep = M_STR_CALSS_PREPARE And mParams.bln�������� = True, True, False)
        .Cell(flexcpBackColor, 1, .ColIndex("��ҩ����"), .rows - 1, .ColIndex("��ҩ����")) = IIf(mcondition.strTransStep = M_STR_CALSS_PREPARE And mParams.bln�������� = True, CSTCOLOR_MODIFY, CSTCOLOR_UNMODIFY)
        
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                '���´�ӡ��־ͼ��
                If Val(.TextMatrix(lngRow, .ColIndex("��ӡ��־"))) = 1 Then
                    .Row = lngRow
                    .Col = .ColIndex("��ӡ")
                    .CellPicture = picPrint(1).Picture
                    .CellPictureAlignment = flexPicAlignCenterCenter
                End If
                
                '������Һ(���)ͼ��
                If Val(.TextMatrix(lngRow, .ColIndex("�Ƿ���"))) > 0 Then
                    .Row = lngRow
                    .Col = .ColIndex("���")
                    .CellPicture = picPacker(Val(.TextMatrix(lngRow, .ColIndex("�Ƿ���")))).Picture
                    .CellPictureAlignment = flexPicAlignCenterCenter
                End If
                
                '���õ������˵ı���ɫ
                If Val(.TextMatrix(lngRow, .ColIndex("������"))) = 1 Then
                    .Cell(flexcpBackColor, lngRow, 1, lngRow, .Cols - 1) = &H80000005
                Else
                    .Cell(flexcpBackColor, lngRow, 1, lngRow, .Cols - 1) = &HC0FFC0
                End If
            End If
        Next
        
'        Call SetTransColHide
        Call GetCount
        
        .Redraw = flexRDDirect
    End With
    
    Call UpdateExeSign(-1)
    ShowTrans = True
End Function

Private Sub PIVAWork_Sure()
'ȷ�ϵ���
Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                    If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                        strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("��ҩID"))
                    End If
                    
                    .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Me.ImgList.ListImages(8).Picture
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                    
                    mrsTrans.Filter = "��ҩID=" & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                    Do While Not mrsTrans.EOF
                        mrsTrans!�Ƿ�ȷ�ϵ��� = 1
                        mrsTrans.Update
                        mrsTrans.MoveNext
                    Loop
                End If
                
            End If
        Next
    End With
    
    If strInputID = "" Then
        MsgBox "��ѡ��Ҫȷ�ϵ�������Һ���ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_��Һ��ҩ��¼_ȷ�ϵ���("
        '��ҩID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ȷ�ϵ���")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub ShowDeptAdvice(ByVal intStep As Integer, ByVal index As Long)
    Dim lng����id As Long
    Dim dblCount As Double
    
    With vsfDept(0)
        .rows = 1
        .rows = 2
            
        .Cell(flexcpText, 1, .ColIndex("ѡ��"), 1, .Cols - 1) = "û��ҽ����Ϣ......"
        .MergeCells = flexMergeRestrictRows
        .MergeRow(1) = True
     
        If mrsDeptAdvice Is Nothing Then Exit Sub
        
        mrsDeptAdvice.Filter = "�˲��־=" & intStep
        mrsDeptAdvice.Sort = "����,����ID,�˲��־"
        
        If mrsDeptAdvice.RecordCount = 0 Then Exit Sub
        
        .Redraw = flexRDNone
        
        .rows = 1
        
        mrsDeptAdvice.MoveFirst
        Do While Not mrsDeptAdvice.EOF
            If lng����id <> mrsDeptAdvice!����ID Then
                lng����id = mrsDeptAdvice!����ID
                
                .rows = .rows + 1
                
                If mstr�ϴβ���ID <> "" Then
                    .TextMatrix(.rows - 1, .ColIndex("ѡ��")) = IIf(InStr(1, mstr�ϴβ���ID, mrsDeptAdvice!����ID) > 0, -1, 0)
                Else
                    .TextMatrix(.rows - 1, .ColIndex("ѡ��")) = IIf(mrsDeptAdvice!ѡ�� = 1, -1, 0)
                End If
                .TextMatrix(.rows - 1, .ColIndex("����")) = mrsDeptAdvice!����
                .TextMatrix(.rows - 1, .ColIndex("����")) = mrsDeptAdvice!����
                .TextMatrix(.rows - 1, .ColIndex("����ID")) = mrsDeptAdvice!����ID
            Else
                .TextMatrix(.rows - 1, .ColIndex("����")) = Val(.TextMatrix(.rows - 1, .ColIndex("����"))) + mrsDeptAdvice!����
            End If
            
            mrsDeptAdvice.MoveNext
        Loop
        
        .Cell(flexcpFontBold, 1, .ColIndex("ѡ��"), .rows - 1, .ColIndex("ѡ��")) = True
        .Cell(flexcpForeColor, 1, .ColIndex("ѡ��"), .rows - 1, .ColIndex("ѡ��")) = vbBlue
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub ShowDeptTrans(ByVal intType As Integer, ByVal strType As String)
    '��ʾ������������Ӧ����Һ��������
    With vsfDept(intType)
        .rows = 1
        .rows = 2
        
        .Cell(flexcpText, 1, .ColIndex("ѡ��"), 1, .Cols - 1) = "û����Һ����Ϣ......"
        .MergeCells = flexMergeRestrictRows
        .MergeRow(1) = True
        
        If mrsDeptTrans Is Nothing Then Exit Sub
        
        mrsDeptTrans.Filter = "����='" & strType & "'"
        
        If mrsDeptTrans.RecordCount = 0 Then Exit Sub
        
        .Redraw = flexRDNone
        
        .rows = 1
        
        If mParams.int�������� = 1 Then
            mrsDeptTrans.Sort = "����"
        Else
            mrsDeptTrans.Sort = "����"
        End If
        
        Do While Not mrsDeptTrans.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("ѡ��")) = IIf(mrsDeptTrans!ѡ�� = 1, -1, 0)
            .TextMatrix(.rows - 1, .ColIndex("����")) = mrsDeptTrans!����
            .TextMatrix(.rows - 1, .ColIndex("����")) = mrsDeptTrans!����
            .TextMatrix(.rows - 1, .ColIndex("����ID")) = mrsDeptTrans!����ID
            .TextMatrix(.rows - 1, .ColIndex("��¼id")) = mrsDeptTrans!��¼id
            
            mrsDeptTrans.MoveNext
        Loop
'        .Cell(flexcpFontBold, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = True
'        .Cell(flexcpForeColor, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = vbBlue
        .Cell(flexcpBackColor, 1, .ColIndex("ѡ��"), .rows - 1, .ColIndex("ѡ��")) = CSTCOLOR_MODIFY
        
        .Redraw = flexRDDirect
    End With
    
    'ͳ��ѡ��Ĳ�������Һ��������
    Call GetCount
End Sub

Private Sub UpdateExeSign(ByVal LngID As Long, Optional ByVal intSign As Integer)
    '���ݱ�����ݸ������ݼ�ִ�б�־
    'lngID��0-���м�¼��intSignֵ����;>0-��Ӧ�����ݼ���¼��intSignֵ����(ҽ��ʱ��ʾ���ID����Һ��ʱ��ʾ��ҩID);"-1"-���ݱ�����ݸ���
    'intSign����lngID=0,>0ʱ����
    Dim lngCount As Long
    
    If mrsTrans Is Nothing Then Exit Sub
    
    If LngID = 0 Then
        If mrsTrans.RecordCount = 0 And mrsTrans.Filter <> "" Then mrsTrans.Filter = ""
        Do While Not mrsTrans.EOF
            mrsTrans!ִ�б�־ = intSign
            mrsTrans.Update
            mrsTrans.MoveNext
        Loop
    ElseIf LngID = -1 Then
        With vsfTrans
            For lngCount = 1 To .rows - 1
                If .TextMatrix(lngCount, .ColIndex("��ҩID")) <> "" Then
                    mrsTrans.Filter = "��ҩID=" & Val(.TextMatrix(lngCount, .ColIndex("��ҩID")))
                    Do While Not mrsTrans.EOF
                        If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
                            mrsTrans!ִ�б�־ = Val(.TextMatrix(lngCount, .ColIndex("��־")))
                        Else
                            mrsTrans!ִ�б�־ = IIf(Val(.TextMatrix(lngCount, .ColIndex("ѡ��"))) = -1, 1, 0)
                        End If
                        mrsTrans.Update
                        mrsTrans.MoveNext
                    Loop
                End If
            Next
        End With
    Else
        mrsTrans.Filter = "��ҩID=" & LngID
        Do While Not mrsTrans.EOF
            mrsTrans!ִ�б�־ = intSign
            mrsTrans.Update
            mrsTrans.MoveNext
        Loop
    End If
    
    DoEvents
    
    Call GetCount
End Sub

Private Sub cboBatch_Click()
    mlng��ɨ�� = 0
    Call SetFilter
    If mcondition.strTransStep = M_STR_CALSS_SEND Then Me.txtFindItem.SetFocus
End Sub

Private Sub cboLevel_Click()
    Call SetFilter
End Sub

Private Sub SetFilter()
    Dim bln��Ƭ As Boolean
    Dim bln�б� As Boolean
    Dim lngRow As Long
    Dim lngCount As Long
    
    Me.chkAll.Value = 0
    
    Call ClearDetailList
    If mblnFilter = False Then Exit Sub
    With vsfDept(Me.tabDeptList.Selected.index)
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("����ID")) <> "" Then
                If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                    lngCount = lngCount + 1
                End If
            End If
        Next
    End With
    
    If lngCount = 0 Then Exit Sub
    If mrsTrans Is Nothing Then Exit Sub
    
    mrsTrans.Filter = ""
    If vsfTrans.TextMatrix(1, vsfTrans.ColIndex("����id")) = "" Then
        If mrsTrans Is Nothing Then
            Exit Sub
        Else
            If mrsTrans.RecordCount = 0 Then Exit Sub
        End If
    End If
    
    If cboBatch.Text = "<ȫ��>" And cboLevel.Text = "<ȫ��>" Then
        mrsTrans.Filter = ""
    ElseIf cboBatch.Text <> "<ȫ��>" And cboLevel.Text = "<ȫ��>" Then
        mrsTrans.Filter = "��ҩ����=" & cboBatch.Text
    ElseIf cboBatch.Text = "<ȫ��>" And cboLevel.Text <> "<ȫ��>" Then
        mrsTrans.Filter = "���ȼ�=" & cboLevel.Text
    Else
        mrsTrans.Filter = "��ҩ����=" & cboBatch.Text & IIf(cboLevel.Text = "<ȫ��>", "", " And ���ȼ�=" & cboLevel.Text)
    End If
    
    If Me.cboFrequency.Text <> "<ȫ��>" Then
        mrsTrans.Filter = IIf(mrsTrans.Filter = 0, "ִ��Ƶ��='" & Me.cboFrequency.Text & "'", mrsTrans.Filter & " And ִ��Ƶ��='" & Me.cboFrequency.Text & "'")
    End If
    
    If Me.chkSure(1).Value = 1 And Me.chkSure(0).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and �Ƿ�ȷ�ϵ���=1", "�Ƿ�ȷ�ϵ���=1")
    ElseIf Me.chkSure(0).Value = 1 And Me.chkSure(1).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and �Ƿ�ȷ�ϵ���=0", "�Ƿ�ȷ�ϵ���=0")
    End If
    
    If Me.chkPrint(1).Value = 1 And Me.chkPrint(0).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and ��ӡ��־=1", "��ӡ��־=1")
    ElseIf Me.chkPrint(0).Value = 1 And Me.chkPrint(1).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and ��ӡ��־=0", "��ӡ��־=0")
    End If
    
    If Me.chkChange(1).Value = 1 And Me.chkChange(0).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and �Ƿ��������=1", "�Ƿ��������=1")
    ElseIf Me.chkChange(0).Value = 1 And Me.chkChange(1).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and �Ƿ��������=0", "�Ƿ��������=0")
    End If
    
    If Me.cboDosType.ListIndex <> 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and ��ҩ����='" & Me.cboDosType.Text & "'", "��ҩ����='" & Me.cboDosType.Text & "'")
    End If

    '�����˵�ҩƷ��������
    If Me.cboMedi.Text <> "<ȫ��>" Or (mParams.blnByMedi = True Or mParams.blnFilter = False) Then
        Call MediSort
        
        If mrsTrans.RecordCount = 0 Then
            Me.vsfTrans.rows = 1
            Me.vsfTrans.rows = 2
            Exit Sub
        End If
        
        Call SetSortFlag(True)
    Else
        Call SetSortFlag
    End If
    
    '��ʾ��Һ����ϸ�б�
    bln�б� = ShowTrans(Me.tabDeptList.Selected.index)
    '��ʾ��Һ��ҩƷ�����б�
    Call ShowSumDrug
    '��״̬����ʾ��ѡ�Ĳ�������Һ������
    Call GetCount
    '��ʾ��Һ����Ƭ
    bln��Ƭ = mfrmPIVCard.ShowDetailCard(mrsTrans, mstr����, mcondition.strTransStep = M_STR_CALSS_PREPARE, mParams.intCount, mParams.bln��������, mParams.bln�������, mcondition.strTransStep, mParams.bln���)
    
    If bln�б� And bln��Ƭ Then
        chkAll.Enabled = True
    End If
End Sub

Private Sub LoadData()
    Dim rstemp As Recordset
    On Error GoTo errHandle
    
    gstrSQL = "select ����id,��������,��ҩ����,Ƶ��,��Ч,���ȼ� from ��ҺҩƷ���ȼ� order by ���ȼ�"
    Set mrsPRI = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ȼ�����")
    
    gstrSQL = "select distinct ���ȼ� from ��ҺҩƷ���ȼ� order by ���ȼ�"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ȼ�����")
    
    cboLevel.Clear
    Me.cboLevel.AddItem "<ȫ��>"
    Do While Not rstemp.EOF
        Me.cboLevel.AddItem rstemp!���ȼ�
        rstemp.MoveNext
    Loop
    cboLevel.Text = "<ȫ��>"

    gstrSQL = "select ����id,��������,����,��ҩ���� from ������������ where ��������ID=[1]"
    Set mrsVol = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������������", mParams.lng��������)
    
    cboMedi.Clear
    Me.cboMedi.AddItem "<ȫ��>"
    If mParams.blnFilter Then
        gstrSQL = "select distinct ҩƷid,���� from ��Һ���ȴ�ӡҩƷ"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��Ϣ")

        Do While Not rstemp.EOF
            Me.cboMedi.AddItem rstemp!����
            Me.cboMedi.ItemData(Me.cboMedi.ListCount - 1) = rstemp!ҩƷID

            mParams.str����ҩƷ = IIf(mParams.str����ҩƷ = "", "", mParams.str����ҩƷ & ",") & rstemp!ҩƷID
            rstemp.MoveNext
        Loop
    End If
    cboMedi.Text = "<ȫ��>"
    
    Set rstemp = DeptSendWork_Get��ҩ����
    cboType.Clear
    cboDosType.Clear
    Me.cboType.AddItem "<ȫ��>"
    Me.cboDosType.AddItem "<ȫ��>"
    Do While Not rstemp.EOF
        Me.cboType.AddItem rstemp!���� & "-" & rstemp!����
        Me.cboDosType.AddItem rstemp!���� & "-" & rstemp!����
        rstemp.MoveNext
    Loop
    cboDosType.Text = "<ȫ��>"
    cboType.Text = "<ȫ��>"
    
    gstrSQL = "select ���� from ����Ƶ����Ŀ"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "ִ��Ƶ��")
    
    cboFrequency.Clear
    Me.cboFrequency.AddItem "<ȫ��>"
    Do While Not rstemp.EOF
        Me.cboFrequency.AddItem rstemp!����
        rstemp.MoveNext
    Loop
    Me.cboFrequency.Text = "<ȫ��>"
    
    Me.cboBatch.Text = "<ȫ��>"
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub MediSort()
    Dim lngҩƷid As Long
    Dim strFilter As String
    Dim str��ҩids As String
    Dim lng��ҩid As Long
    Dim i As Long
    Dim j As Long
    Dim strSort As String
    Dim strIDSOld As String
    Dim str��ҩid As String
    Dim strIDS As String
    Dim strTemp As String
    
    lngҩƷid = Val(Me.cboMedi.ItemData(Me.cboMedi.ListIndex))
    strFilter = IIf(mrsTrans.Filter = 0, "", mrsTrans.Filter)
    mstrFilter = ""
    
    If lngҩƷid <> 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and ҩƷid=" & lngҩƷid, "ҩƷid=" & lngҩƷid)
        mrsTrans.Sort = "ҩƷid,����"
        
        Do While Not mrsTrans.EOF
            strIDSOld = IIf(strIDSOld = "", mrsTrans!��ҩid, strIDSOld & "," & mrsTrans!��ҩid)
            str��ҩid = IIf(str��ҩid = "", "��ҩid=" & mrsTrans!��ҩid, str��ҩid & " or ��ҩid=" & mrsTrans!��ҩid)
            mrsTrans.MoveNext
        Loop
        
        If str��ҩid = "" Then Exit Sub
        mrsTrans.Filter = str��ҩid
        mrsTrans.Sort = "��ý,ҩƷid"
        
        lngҩƷid = 0
        Do While Not mrsTrans.EOF
            If mrsTrans!��ý = 1 Then
                If lngҩƷid <> mrsTrans!ҩƷID Or lngҩƷid = 0 Then
                    lngҩƷid = mrsTrans!ҩƷID
                    strIDS = IIf(strIDS = "", mrsTrans!��ҩid, strIDS & "|" & mrsTrans!��ҩid)
                Else
                    strIDS = IIf(strIDS = "", mrsTrans!��ҩid, strIDS & "," & mrsTrans!��ҩid)
                End If
'                mstrFilter = IIf(mstrFilter = "", mrsTrans!��ҩid, mstrFilter & "," & mrsTrans!��ҩid)
            End If
            mrsTrans.MoveNext
        Loop
        
        str��ҩid = ""
        For i = 0 To UBound(Split(strIDS, "|"))
            For j = 0 To UBound(Split(strIDSOld, ","))
                If InStr(1, "," & Split(strIDS, "|")(i) & ",", "," & Split(strIDSOld, ",")(j) & ",") > 0 Then
                    If InStr(1, "," & mstrFilter & ",", "," & Split(strIDSOld, ",")(j) & ",") < 1 Then
                        mstrFilter = IIf(mstrFilter = "", Split(strIDSOld, ",")(j), mstrFilter & "," & Split(strIDSOld, ",")(j))
                    End If
                End If
            Next
        Next
        
        For j = 0 To UBound(Split(strIDSOld, ","))
            If InStr(1, "," & mstrFilter & ",", "," & Split(strIDSOld, ",")(j) & ",") < 1 Then
                If InStr(1, "," & strTemp & ",", "," & Split(strIDSOld, ",")(j) & ",") < 1 Then
                    strTemp = strTemp & "," & Split(strIDSOld, ",")(j)
                End If
            End If
        Next
        
        mstrFilter = IIf(mstrFilter = "", Mid(strTemp, 2), mstrFilter & strTemp)
        mrsTrans.Filter = ""
    Else
        mrsTrans.Sort = "��ҩ����,����ҩƷid,��ýid,������,ƿǩ��,��ý"
    End If
    
    
    
'    If mParams.blnByMedi = True And Val(Me.vsfTrans.Tag) = 1 Then
'        mrsTrans.Sort = "��ҩ����,����ҩƷid,��ýid,������,ƿǩ��,��ý"
'    Else
'        strSort = mParams.strSort
'        If InStr(1, mParams.strSort, "��ҩ����") > 0 Then
'            strSort = Replace(mParams.strSort, "��ҩ����", "��ҩ����,���ȼ�")
'        End If
'        If InStr(1, mParams.strSort, "����") > 0 Then
'            strSort = Replace(mParams.strSort, "����", "����,����")
'        End If
'        mrsTrans.Sort = IIf(strSort <> "", strSort & ",��ҩid,��ý", "��ҩid,��ý")
'    End If
End Sub













Private Sub cboʱ�䷶Χ_Click()
    Dim dteTime As Date
    
    With cboʱ�䷶Χ
        If .ListIndex <> Val(.Tag) Then
            If (Val(.Tag) = 3 And .ListIndex < 3) Or (Val(.Tag) < 3 And .ListIndex = 3) Then
                Call ResizeConditionArea
            End If
            .Tag = .ListIndex
            
            dteTime = Sys.Currentdate
            
            If .ListIndex = 0 Then
                Dtp��ʼʱ��.Value = CDate(Format(dteTime, "YYYY-MM-DD"))
                Dtp����ʱ��.Value = CDate(Format(dteTime, "YYYY-MM-DD"))
            ElseIf .ListIndex = 1 Then
                Dtp��ʼʱ��.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD"))
                Dtp����ʱ��.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD"))
            ElseIf .ListIndex = 2 Then
                Dtp��ʼʱ��.Value = CDate(Format(dteTime, "YYYY-MM-DD"))
                Dtp����ʱ��.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD"))
            ElseIf .ListIndex = 3 Then
                If mcondition.strTransStartTime <> "" Then
                    Dtp��ʼʱ��.Value = CDate(Format(mcondition.strTransStartTime, "YYYY-MM-DD hh:mm:ss"))
                    Dtp����ʱ��.Value = CDate(Format(mcondition.strTransEndTime, "YYYY-MM-DD hh:mm:ss"))
                Else
                    Dtp��ʼʱ��.Value = CDate(Format(dteTime, "YYYY-MM-DD hh:mm:ss"))
                    Dtp����ʱ��.Value = CDate(Format(dteTime, "YYYY-MM-DD hh:mm:ss"))
                End If
            End If
            mcondition.intTransTimeSel = .ListIndex
        End If
    End With
    
    Call RefreshDeptList(Me.tabDeptList.Selected.index)
End Sub


Private Sub cbsMain_ControlSelected(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    If Me.cbsMain Is Nothing Then Exit Sub
    If Control Is Nothing Then Exit Sub
    
    '�����˵���ѡ��
    If Control.Id = conMenu_Oper_Select Then
        '������ѡ��
        Set cbrControl = Me.cbsMain.FindControl(xtpControlButton, conMenu_Oper_Select_SelBatch, False, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Enabled = True
        End If
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim lngPatiID As Long
    Dim str�Һŵ� As String
    Dim lng��ҳID As Long
    Dim lngCurrAdviceID As Long

    Select Case Control.Id
        ''''�ļ�
        Case mconMenu_File_PrintSet     '��ӡ����
            zlPrintSet
        Case mconMenu_File_Preview      '��ӡԤ��
            zlSubPrint 2
        Case mconMenu_File_Print        '��ӡ
            zlSubPrint 1
        Case mconMenu_File_Excel        '�����Excel
            zlSubPrint 3

        Case mconMenu_File_PIVA_BillPrintLable
            '��ӡ��ǩ
            Call BillPrint_Label(Control.Id)
        Case mconMenu_Edit_PIVA_Approve
            '���
            Call PIVAWork_Approve
        Case mconMenu_Edit_PIVA_Beach
            '��������
            Call SetBeach
        Case MCONMENU_EDIT_PIVA_SURE
            'ȷ�ϵ���
            Call PIVAWork_Sure
        Case mconMenu_Edit_PIVA_Prepare
            'ִ��Ԥ����
            Call setNOtExcetePrice
    
            '��ҩȷ��
            Call PIVAWork_Prepare(1)
        Case mconMenu_Edit_PIVA_Dosage
            '��ҩȷ��
            Call PIVAWork_Dosage
        Case mconMenu_Edit_PIVA_Send
            '����ȷ��
            Call PIVAWork_Send
        Case MCONMENU_EDIT_PIVA_REFUSE
            'ȷ�Ͼܾ�
            Call PIVAWork_Refuse
        Case mconMenu_Edit_PIVA_ReVerify
            'ִ��Ԥ����
            Call setNOtExcetePrice
            
            '�������
            Call PIVAWork_ReturnVerify
        Case mconMenu_Edit_PIVA_Cancel
            'ִ��Ԥ����
            Call setNOtExcetePrice
            
            'ȡ����һ������
            Call PIVAWork_Cancel
        Case mconMenu_Edit_PIVA_Delete
            'ɾ���ѻ�����ҽ������Һ��ҩ��¼
            Call PIVAWork_Delete
            
        Case MCONMENU_PLAN_PIVA_DESK
            Call frmDesk.ShowMe(mParams.lng��������, Me)
        Case MCONMENU_PLAN_PIVA_DESKDRUG
            Call frmDeskMedi.ShowMe(mParams.lng��������, Me)
        Case MCONMENU_PLAN_PIVA_PERWORK
            Call frmPlan.ShowMe(mParams.lng��������, Me)
        Case mconMenu_Edit_PIVA_PASS
            '���ܣ��Բ��˹���ʷ/����״̬���й���
            'Pass
            If Not Me.vsfMedis Is ActiveControl Then Exit Sub
            If vsfMedis.Row = 0 Then Exit Sub
            
            lngPatiID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("����id")))
            lng��ҳID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("��ҳid")))
            
            Call gobjPass.zlPassCmdAlleyManage_YF(mlngMode, lngPatiID, lng��ҳID, "")
        
        '��ӡ����
        Case mconMenu_File_PIVA_BillPrintWait
            '��ӡ��ҩ��
            Call BillPrint_Prepare
        Case mconMenu_File_PIVA_BillPrintTotal
            '��ӡ���͵�
            Call BillPrint_Send
        Case mconMenu_File_PIVA_BillPrintReturn
            '��ӡ��ҩ�����嵥
            Call BillPrint_Return
        Case mconMenu_File_PIVA_BillPrintNext
            Call frmPrint.ShowMe(Me)
        Case mconMenu_File_PIVA_BillPrintSum
            '��ӡ���ܱ���
            Call BillPrint_Sum
        Case mconMenu_File_Parameter
            '��������
            ResetParams
        Case MCONMENU_EDIT_PIVA_SORTSET
            '�����������
            frmPIVASortSet.Show 1, Me
            Call SetSort(True)
        Case mconMenu_View_Refresh
            'ˢ��
            Call RefreshDeptList(Me.tabDeptList.Selected.index)
        
        Case mconMenu_File_Exit
            '�˳�
            Unload Me
        
        ''''�鿴
        Case mconMenu_View_ToolBar_Button               '��׼��ť
            Control.Checked = Not Control.Checked
            Me.cbsMain(2).Visible = Control.Checked
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Text                 '�ı���ǩ
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbsMain(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Size                 '��ͼ��
            Control.Checked = Not Control.Checked
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_StatusBar                    '״̬��
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3                   '�ֺ�����
            mParams.intFont = Val(Control.Parameter)
            Call SetFontSize
        Case mconMenu_View_ShowHistory
            Control.Checked = Not Control.Checked
            mParams.intAutoSelect = IIf(Control.Checked, 1, 0)
        Case mconMenu_Edit_PIVA_Lock
            '��ǰ����ȫ������
            Call SetLock(1, "")
        Case mconMenu_Edit_PIVA_UnLock
            '��ǰ����ȫ������
            Call SetLock(0, "")
            
        ''''����
        Case mconMenu_Help_Help                         '����
            Call ShowHelp(App.ProductName, Me.hWnd, "Frm���ŷ�ҩ����")
        Case mconMenu_Help_Web                          'WEB�ϵ�����
        Case mconMenu_Help_Web_Home                     '������ҳ
            Call zlHomePage(Me.hWnd)
        Case mconMenu_Help_Web_Forum                    '������̳
            Call zlWebForum(Me.hWnd)
        Case mconMenu_Help_Web_Mail                     '���ͷ���
            Call zlMailTo(Me.hWnd)
        Case mconMenu_Help_About                        '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        
        ''�����˵�
        Case conMenu_Oper_PrintLabel_SelRow, conMenu_Oper_PrintLabel_SelBatch, conMenu_Oper_PrintLabel_SelDept, conMenu_Oper_PrintLabel_SelPati, conMenu_Oper_PrintLabel_AllRow, conMenu_Oper_PrintLabel_SelSendNo
            '��ӡ��ǩ
            Call BillPrint_Label(Control.Id)
        Case conMenu_Oper_DelBatch_SelRow, conMenu_Oper_DelBatch_SelBatch, conMenu_Oper_DelBatch_SelDept, conMenu_Oper_DelBatch_SelPati, conMenu_Oper_DelBatch_AllRow
            'ɾ������
            Call DeleteBatch(Control.Id)
        Case conMenu_Oper_Select_SelRow, conMenu_Oper_Select_SelBatch, conMenu_Oper_Select_SelDept, conMenu_Oper_Select_CancleSelDept, conMenu_Oper_Select_SelPati, conMenu_Oper_Select_CancleSelPati, conMenu_Oper_Select_SelSendNo, conMenu_Oper_Select_SelAll, conMenu_Oper_Select_SelMed, conMenu_Oper_Bag_Batch, conMenu_Oper_Bag_All
            '����ѡ��,���
            Call SelectBatch(Control.Id, (Me.tabDeptList.Selected.index))
'        Case conMenu_Oper_DelLabel_SelRow, conMenu_Oper_DelLabel_SelBatch, conMenu_Oper_DelLabel_SelDept, conMenu_Oper_DelLabel_SelPati, conMenu_Oper_DelLabel_AllRow
'            'ɾ����ǩ
'            Call DeleteLabel(Control.Id)
        Case conMenu_Oper_Look
            On Error Resume Next
            '���Ӳ�������
            If Not mobjCISJOB Is Nothing Then
                Call mobjCISJOB.ShowArchive(Me, Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("����id"))), Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("��ҳid"))))
            End If
            
            err.Clear
        Case MCONMENU_EDIT_PIVA_MedicalRecord
            '���Ӳ�������(������)
            Call ShowMedicalRecord
        Case mconMenu_Edit_PlugIn + 1 To mconMenu_Edit_PlugIn + 99 '��ҷ�ҩҵ���ܵ���
            PivaExPlugNormal Control.Parameter
        Case mconMenu_PASS * 10# To mconMenu_PASS * 10# + 99
            lngCurrAdviceID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("ҽ��id")))
            lngPatiID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("����id")))
            lng��ҳID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("��ҳid")))
            
            Call gobjPass.zlPassCommandBarExe_YF(mlngMode, Control.Id - (mconMenu_PASS * 10#), lngPatiID, lng��ҳID, "", lngCurrAdviceID)
        '�����˵���PASS����
'        Case mconMenu_PASS_Item
'            'ҩ���ٴ���Ϣ�ο�
'            Call PassDoCommand(101)
'        Case mconMenu_PASS_Item + 1
'            'ҩƷ˵����
'            Call PassDoCommand(102)
'        Case mconMenu_PASS_Item + 2
'            '�й�ҩ��
'            Call PassDoCommand(107)
'        Case mconMenu_PASS_Item + 3
'            '������ҩ����
'            PassDoCommand (103)
'        Case mconMenu_PASS_Item + 4
'             '����ֵ
'            Call PassDoCommand(104)
'        Case mconMenu_PASS_Item + 6
'            'ҽҩ��Ϣ����
'            Call PassDoCommand(106)
'        Case mconMenu_PASS_Item + 7
'            'ҩƷ�����Ϣ
'             Call PassDoCommand(13)
'        Case mconMenu_PASS_Item + 8
'            '��ҩ;�������Ϣ
'            Call PassDoCommand(14)
'        Case mconMenu_PASS_Item + 9
'            'ҽԺҩƷ��Ϣ
'            Call PassDoCommand(105)
'
'        '���ܣ�ִ��ר��PASS����
'        Case mconMenu_PASS_Spec
'            'ҩ��-ҩ���໥����
'            Call PassDoCommand(201)
'        Case mconMenu_PASS_Spec + 1
'            'ҩ��-ʳ���໥ʹ��
'            Call PassDoCommand(202)
'        Case mconMenu_PASS_Spec + 2
'            '����ע�������
'            Call PassDoCommand(203)
'        Case mconMenu_PASS_Spec + 3
'            '����ע�������
'            Call PassDoCommand(204)
'        Case mconMenu_PASS_Spec + 4
'            '����֢
'            Call PassDoCommand(205)
'        Case mconMenu_PASS_Spec + 5
'            '������
'            Call PassDoCommand(206)
'        Case mconMenu_PASS_Spec + 6
'            '��������ҩ
'            Call PassDoCommand(207)
'        Case mconMenu_PASS_Spec + 7
'            '��ͯ��ҩ
'            Call PassDoCommand(208)
'        Case mconMenu_PASS_Spec + 8
'            '��������ҩ
'            Call PassDoCommand(209)
'        Case mconMenu_PASS_Spec + 9
'            '��������ҩ
'            Call PassDoCommand(210)
'        Case mconMenu_PASS_Spec + 10
'            Call AdviceCheckWarn(9, "0000000", 2, 1, Me.vsfMedis.TextMatrix(Me.vsfMedis.Row, Me.vsfMedis.ColIndex("ҽ��id")))
        Case Else
            '���Ҳ˵�
            If Control.Id > mconMenu_Look And Control.Id < mconMenu_Look + 10 Then
                lblFindItem.Caption = Control.Caption
                
                Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Look)
                If Not objPopup Is Nothing Then
                    For Each cbrControl In objPopup.CommandBar.Controls
                        cbrControl.Checked = False
                        If cbrControl.Caption = lblFindItem.Caption Then
                            cbrControl.Checked = True
                        End If
                    Next
                End If
            End If
            
            If Control.Id > mconMenu_Filter And Control.Id < mconMenu_Filter + 10 Then
                lblName.Caption = Control.Caption & "��"
                
                Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Filter)
                If Not objPopup Is Nothing Then
                    For Each cbrControl In objPopup.CommandBar.Controls
                        cbrControl.Checked = False
                        If cbrControl.Caption = Mid(lblName.Caption, 1, Len(lblName.Caption) - 1) Then
                            cbrControl.Checked = True
                        End If
                    Next
                End If
            End If
             
            If Control.Id > 401 And Control.Id < 499 Then
                'ִ���Զ��屨��
                Call BillPrint_Custom(Control)
            End If
        
            '�������򵯳��˵�
            If Control.Id > mconMenu_SortPopup And Control.Id < mconMenu_SortPopup + 10 Then
                Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_SortPopup)
                If Not objPopup Is Nothing Then
                    For Each cbrControl In objPopup.CommandBar.Controls
                        cbrControl.Checked = False
                    Next
                End If
                
                Control.Checked = True
                If mParams.int�������� <> Control.Id - mconMenu_SortPopup Then
                    mParams.int�������� = Control.Id - mconMenu_SortPopup
                    Call ShowDeptTrans(Me.tabDeptList.Selected.index, IIf(Me.tabDeptList.Selected.index = CNUMWORK, tabWork.Selected.Tag, tbcLook.Selected.Tag))
                End If
            End If
    End Select
End Sub

Private Sub PivaExPlugNormal(ByVal strFunName As String)
    Dim lng��ҩid As Long
    
    If Not mobjPlugIn Is Nothing Then
        If vsfTrans.rows > 1 Then
            lng��ҩid = Val(vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("��ҩID")))
        End If
        
        On Error Resume Next
        Call mobjPlugIn.PivaWorkNormal(glngModul, strFunName, mParams.lng��������, lng��ҩid)
        err.Clear: On Error GoTo 0
    End If
    
End Sub

Private Sub SetSort(Optional ByVal BlnRefresh As Boolean = False)
    '������Һ������
    Dim strSortString As String
    Dim i As Integer
    Const ALL_SORT_ITEM As String = "����,����,����,��ҩ����,����,ƿǩ��,ִ��ʱ��"
    Const DEFAULT_SORT As String = "����,����,����"
    
    strSortString = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "��Һ�������Ĺ���", "��Һ������", "")
    If strSortString = "" Then
        strSortString = DEFAULT_SORT
    ElseIf InStr(1, strSortString, "|") = 0 Then
        strSortString = DEFAULT_SORT
    ElseIf Mid(strSortString, InStr(1, strSortString, "|") + 1) = "" Then
        strSortString = DEFAULT_SORT
    Else
        strSortString = Mid(strSortString, InStr(1, strSortString, "|") + 1)
        For i = 0 To UBound(Split(strSortString, ","))
            If Split(strSortString, ",")(i) <> "" Then
                If InStr(1, "," & ALL_SORT_ITEM & ",", "," & Split(strSortString, ",")(i) & ",") = 0 Then
                    strSortString = DEFAULT_SORT
                    Exit For
                End If
            End If
        Next
    End If
    
    If strSortString <> mParams.strSort Then
        mParams.strSort = strSortString
        If BlnRefresh = True Then Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
    
    Call SetSortFlag
End Sub
Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strName As String
    
    strName = Split(Control.Parameter, ",")(1)

    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strName, Me, _
        "��ʼʱ��=" & mcondition.strTransStartTime, _
        "����ʱ��=" & mcondition.strTransEndTime, _
        "��ҩ��¼=", _
        "ƿǩ��=")

End Sub
Private Sub SetFontSize()
    Dim intFont As Integer
    Dim stdfnt As StdFont
    
    Select Case mParams.intFont
        Case 0
            intFont = 9
        Case 1
            intFont = 11
        Case 2
            intFont = 15
        Case Else
            intFont = 9
    End Select
    
    With vsfDept(0)
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("��") + 120
        .RowHeightMax = TextHeight("��") + 120
        .Refresh
    End With
    
    With vsfTrans
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("��") + 120
        .RowHeightMax = TextHeight("��") + 120
        .Refresh
    End With
    
'    With vsfDrug
'        .Font.Size = intFont
'        Me.Font.Size = .Font.Size
'        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
'
'        .RowHeightMin = TextHeight("��") + 120
'        .RowHeightMax = TextHeight("��") + 120
'        .Refresh
'    End With
    
    With vsfSumDrug
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("��") + 120
        .RowHeightMax = TextHeight("��") + 120
        .Refresh
    End With
    
    If Not tbcDetail.PaintManager.Font Is Nothing Then
        With tbcDetail
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = intFont
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    Me.FontSize = intFont
End Sub
Private Sub zlSubPrint(ByVal bytMode As Byte)
    'bytMode��1-��ӡ��2-Ԥ����3-�����Excel
    Dim ObjThis As Object
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim strTitle As String
    
    'ȡ��ӡ�б����
    Select Case tbcDetail.Selected.index
        Case mDetailType.��Һ���б�
            If vsfTrans.rows = 1 Then Exit Sub
            If vsfTrans.TextMatrix(1, vsfTrans.ColIndex("��ҩID")) = "" Then Exit Sub
            
            Set ObjThis = GetPrintObj(vsfTrans)
            
            Select Case mcondition.strTransStep
                Case M_STR_CALSS_PREPARE
                    strTitle = "����ҩ��Һ���嵥"
                Case M_STR_CALSS_DOSAGE
                    strTitle = "����ҩ��Һ���嵥"
                Case M_STR_CALSS_SEND
                    strTitle = "��������Һ���嵥"
                Case M_STR_CALSS_SENDED
                    strTitle = "�ѷ�����Һ���嵥"
            End Select
        Case mDetailType.ҩƷ�����б�
            If vsfSumDrug.rows = 1 Then Exit Sub
            If vsfSumDrug.TextMatrix(1, vsfSumDrug.ColIndex("ҩƷ����")) = "" Then Exit Sub
            
            Set ObjThis = GetPrintObj(vsfSumDrug)
            
            Select Case mcondition.strTransStep
                Case M_STR_CALSS_PREPARE
                    strTitle = "����ҩҩƷ�嵥"
                Case M_STR_CALSS_DOSAGE
                    strTitle = "����ҩҩƷ�嵥"
                Case M_STR_CALSS_SEND
                    strTitle = "������ҩƷ�嵥"
                Case M_STR_CALSS_SENDED
                    strTitle = "�ѷ���ҩƷ�嵥"
            End Select
    End Select
    
    If ObjThis Is Nothing Then Exit Sub
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "��ӡ��:" & gstrUserName
    ObjAppRow.Add "��ӡʱ��:" & Format(Sys.Currentdate, "yyyy-MM-dd HH:MM:SS")
    objPrint.BelowAppRows.Add ObjAppRow
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "��ʼʱ��:" & Format(Dtp��ʼʱ��.Value, "yyyy-MM-dd ")
    ObjAppRow.Add "����ʱ��:" & Format(Dtp����ʱ��.Value, "yyyy-MM-dd ")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = strTitle
    Set objPrint.Body = ObjThis
    
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Function GetPrintObj(ByVal vsfObj As VSFlexGrid, Optional ByVal strMerge As String = "") As VSFlexGrid
    '���ش�ӡ�ؼ�
'    Dim vsfPrint As VSFlexGrid
    Dim lngCol As Long, lngRow As Long
    Dim lngCols As Long
    Dim lngPrintCol As Long
    
    With vsfPrint
        .Cols = 1
        .rows = 1
        
        For lngCol = 0 To vsfObj.Cols - 1
            If vsfObj.ColHidden(lngCol) = False And vsfObj.ColWidth(lngCol) > 0 And vsfObj.ColKey(lngCol) <> "��ǰ��" Then
                lngCols = lngCols + 1
                
                If lngCols > 1 Then
                    .Cols = .Cols + 1
                End If
                
                .ColKey(lngCols - 1) = vsfObj.ColKey(lngCol)
                .ColWidth(lngCols - 1) = vsfObj.ColWidth(lngCol)
                .FixedAlignment(lngCols - 1) = vsfObj.FixedAlignment(lngCol)
                .ColAlignment(lngCols - 1) = vsfObj.ColAlignment(lngCol)
            End If
        Next
                
        For lngRow = 0 To vsfObj.rows - 1
            For lngCol = 0 To vsfObj.Cols - 1
                For lngPrintCol = 0 To .Cols - 1
                    If .ColKey(lngPrintCol) = vsfObj.ColKey(lngCol) Then
                        .TextMatrix(lngRow, lngPrintCol) = vsfObj.TextMatrix(lngRow, lngCol)
                        If .ColKey(lngPrintCol) = "ѡ��" And lngRow > 0 Then
                            If Val(vsfObj.TextMatrix(lngRow, vsfObj.ColIndex("ѡ��"))) = -1 Then
                                .TextMatrix(lngRow, lngPrintCol) = "��"
                            Else
                                .TextMatrix(lngRow, lngPrintCol) = ""
                            End If
                        End If
                        
                        If .ColKey(lngPrintCol) = "��ӡ" And lngRow > 0 Then
                            If vsfObj.TextMatrix(lngRow, vsfObj.ColIndex("��ӡ��־")) = 1 Then
                                .TextMatrix(lngRow, lngPrintCol) = "��"
                            Else
                                .TextMatrix(lngRow, lngPrintCol) = ""
                            End If
                        End If
                        
                        If .ColKey(lngPrintCol) = "���" And lngRow > 0 Then
                            If vsfObj.TextMatrix(lngRow, vsfObj.ColIndex("�Ƿ���")) = 1 Then
                                .TextMatrix(lngRow, lngPrintCol) = "��"
                            Else
                                .TextMatrix(lngRow, lngPrintCol) = ""
                            End If
                        End If
                        Exit For
                    End If
                Next
            Next
            .rows = .rows + 1
        Next
    End With
    
    Set GetPrintObj = vsfPrint
End Function
Private Sub ResetParams()
    mblnParamsRefresh = False
    
    With frmPIVAParaSet
        .mstrPrivs = mstrPrivs
        .mlng�ⷿid = mParams.lng��������
        .Show 1, Me
    End With
    
    If mblnParamsRefresh = True Then
        Call GetParams
        
        If mcondition.lngCenterID <> mParams.lng�������� Then
            mcondition.lngCenterID = mParams.lng��������
        End If
        
        Call ShowComment(tbcDetail.Selected.index, mcondition.strTransStep)
        Call SetCommand
        
        DoEvents
        
        Call RefreshDeptList(0)
        
        DoEvents
        
        Call RefreshDetailList(0)
    End If
End Sub

Private Sub BillPrint_Label(ByVal lngType As Long)
    '��ӡ��ǩ
    Dim strInputID As String    '��ҩID...
    Dim strPrintID As String    '��ҩID,ƿǩ��|��ҩ��,ƿǩ��...
    Dim lngRow As Long
    Dim strCom As String
    Dim arrParams
    Dim strMsg As String
    Dim i As Integer
    Dim str��ҩid As String
    Dim dateNow As Date
    Dim blnPrint As Boolean
    Dim str����Ա As String
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    If vsfTrans.rows = 1 Then Exit Sub
    If vsfTrans.TextMatrix(1, vsfTrans.ColIndex("��ҩID")) = "" Then Exit Sub
    
    
    With vsfTrans
'        If Me.tbcDetail.Item(mDetailType.��Һ����Ƭ).Selected Then
'            For i = 1 To .rows - 1
'                If .TextMatrix(i, .ColIndex("��ҩID")) = mstr��ҩid Then
'                    .Row = i
'                    Exit For
'                End If
'            Next
'        End If
        If lngType = conMenu_Oper_PrintLabel_SelBatch Then
            strCom = .TextMatrix(.Row, .ColIndex("��ҩ����"))
            strMsg = "��ӡ��ǰ����Ϊ��" & .TextMatrix(.Row, .ColIndex("��ҩ����")) & "��������ƿǩ���Ƿ������"
        ElseIf lngType = conMenu_Oper_PrintLabel_SelDept Then
            strCom = .TextMatrix(.Row, .ColIndex("����"))
            strMsg = "��ӡ��ǰ����Ϊ��" & .TextMatrix(.Row, .ColIndex("����")) & "��������ƿǩ���Ƿ������"
        ElseIf lngType = conMenu_Oper_PrintLabel_SelPati Then
            strCom = .TextMatrix(.Row, .ColIndex("����")) & .TextMatrix(.Row, .ColIndex("����")) & .TextMatrix(.Row, .ColIndex("����"))
            strMsg = "��ӡ��ǰ����Ϊ��" & .TextMatrix(.Row, .ColIndex("����")) & "��������ƿǩ���Ƿ������"
        ElseIf lngType = conMenu_Oper_PrintLabel_SelSendNo Then
            strCom = .TextMatrix(.Row, .ColIndex("��ҩ����"))
            strMsg = "��ӡ��ǰ��ҩ����Ϊ��" & .TextMatrix(.Row, .ColIndex("��ҩ����")) & "��������ƿǩ���Ƿ������"
        ElseIf lngType = conMenu_Oper_PrintLabel_AllRow Then
            strMsg = "��ӡ��ǰ��ѡ�������ƿǩ���Ƿ������"
        End If
        
        If lngType = conMenu_Oper_PrintLabel_SelRow Then
            '��ǰ��
            strInputID = Val(.TextMatrix(.Row, .ColIndex("��ҩID")))
            strPrintID = Val(.TextMatrix(.Row, .ColIndex("��ҩID"))) & "," & .TextMatrix(.Row, .ColIndex("ƿǩ��"))
        Else
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                    If lngType = conMenu_Oper_PrintLabel_SelBatch Then
                        If .TextMatrix(lngRow, .ColIndex("��ҩ����")) = strCom Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & "," & .TextMatrix(lngRow, .ColIndex("ƿǩ��"))
                            End If
                        End If
                    ElseIf lngType = conMenu_Oper_PrintLabel_SelDept Then
                        If .TextMatrix(lngRow, .ColIndex("����")) = strCom Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & "," & .TextMatrix(lngRow, .ColIndex("ƿǩ��"))
                            End If
                        End If
                    ElseIf lngType = conMenu_Oper_PrintLabel_SelPati Then
                        If .TextMatrix(lngRow, .ColIndex("����")) & .TextMatrix(lngRow, .ColIndex("����")) & .TextMatrix(lngRow, .ColIndex("����")) = strCom Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & "," & .TextMatrix(lngRow, .ColIndex("ƿǩ��"))
                            End If
                        End If
                    ElseIf lngType = conMenu_Oper_PrintLabel_SelSendNo Then
                        If .TextMatrix(lngRow, .ColIndex("��ҩ����")) = strCom Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & "," & .TextMatrix(lngRow, .ColIndex("ƿǩ��"))
                            End If
                        End If
                    ElseIf lngType = mconMenu_File_PIVA_BillPrintLable Or lngType = conMenu_Oper_PrintLabel_AllRow Then
                        If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & "," & .TextMatrix(lngRow, .ColIndex("ƿǩ��"))
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End With
    
    If strPrintID = "" Then
        MsgBox "��ѡ��Ҫ��ӡ��ǩ����Һ���ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If strMsg <> "" Then
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    If mParams.blnRePeople And (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_SEND) Then
        str����Ա = Mid(frmPeople.ShowMe(mParams.lng��������), 2)
    End If
    
    dateNow = Sys.Currentdate
    
    arrParams = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrParams)
        Call RefreshPrintSign(CStr(arrParams(i)), dateNow, str����Ա)
    Next
    
    DoEvents
    
    '�������ݼ�����
    With mrsTrans
        arrParams = Split(strPrintID, "|")
        For lngRow = 0 To UBound(arrParams)
            If arrParams(lngRow) <> "" Then
                .Filter = "��ҩID=" & Val(Split(arrParams(lngRow), ",")(0))
                
                Do While Not .EOF
                    !��ӡ��־ = 1
                    !ƿǩ�� = Split(arrParams(lngRow), ",")(1)
                    .Update
                    .MoveNext
                Loop
            End If
        Next
    End With
    
    DoEvents
    
    '�����б���ʾ
    With vsfTrans
        str��ҩid = ";"
        .Redraw = flexRDNone
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("��ӡ��־"))) = 0 And InStr(1, "|" & strPrintID, "|" & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ",") > 0 Then
                    str��ҩid = str��ҩid & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ";"
                    .Row = lngRow
                    .Col = .ColIndex("��ӡ")
                    .CellPicture = picPrint(1).Picture
                    .CellPictureAlignment = flexPicAlignCenterCenter
                End If
            End If
        Next
        .Redraw = flexRDDirect
    End With
    
    mfrmPIVCard.BatchPrint str��ҩid
    
    DoEvents
    
    '���ñ����ӡ��ǩ
    arrParams = Split(strPrintID, "|")
    For lngRow = 0 To UBound(arrParams)
        If arrParams(lngRow) <> "" Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_1", Me, _
                "��ҩID=" & Val(Split(arrParams(lngRow), ",")(0)), _
                "ƿǩ��=" & Split(arrParams(lngRow), ",")(1), _
                "PrintEmpty=0", 2)
        End If
    Next
    
    '��ӡ�����嵥
    If mParams.int��ӡ���� = 0 Then
        blnPrint = (MsgBox("�Ƿ��ӡ���δ�ӡ��ҩƷ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        
        If blnPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_3", Me, _
            "����=" & mcondition.lngCenterID, _
            "��ӡʱ��=" & dateNow, "PrintEmpty=0", 1)
        End If
    ElseIf mParams.int��ӡ���� = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_3", Me, _
        "����=" & mcondition.lngCenterID, _
        "��ӡʱ��=" & dateNow, "PrintEmpty=0", 1)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Cancel()
    'PIVA������ȡ�������ݵ�ǰ���账��
    Dim strID As String
    Dim lngRow As Long
    Dim strMsg As String
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    Dim strErr As String
    Dim intRow As Integer
    Dim lng��ҩid As Long
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing And mcondition.strTransStep <> M_STR_CALSS_FAILAUDIT And mcondition.strTransStep <> M_STR_CALSS_PASSEDAUDIT Then Exit Sub
    
'    If MsgBox("�Ƿ�" & strMsg & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
         With Me.vsfMedis
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) <> 0 Then
                    If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
                        If Val(.TextMatrix(lngRow, .ColIndex("ҽ��id"))) <> 0 Then
                            '���ҽ���Ƿ��Ѿ���ҩ
                            If Not CheckIs��ҩ(Val(.TextMatrix(lngRow, .ColIndex("ҽ��id")))) Then
                                strID = strID & .TextMatrix(lngRow, .ColIndex("ҽ��id")) & "," & .TextMatrix(lngRow, .ColIndex("��־")) & "|"
                            Else
                                strErr = "��ѡ��ҽ�������Ѿ���ҩ��ҽ�����Ѿ���ҩ��ҽ�����ܽ���ȡ����˲������Ƿ����ȡ������ҽ������ˣ�"
                            End If
                        End If
                    Else
                        strID = strID & .TextMatrix(lngRow, .ColIndex("ҽ��id")) & "," & .TextMatrix(lngRow, .ColIndex("��־")) & "|"
                    End If
                End If
            Next
        End With
        
        If strErr <> "" Then
            If MsgBox(strErr, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        End If
    Else
        If mcondition.strTransStep <> M_STR_CALSS_PREPARE Then
            With mrsTrans
                .Filter = "ִ�б�־=1"
                .Sort = "����,��ҩ����,סԺ��"
                
                Do While Not .EOF
                    If InStr(1, "," & strID & ",", "," & !��ҩid & ",") = 0 Then
                        strID = IIf(strID = "", "", strID & ",") & !��ҩid
                    End If
                    
                    If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_SEND Then
                        If IsOutPatient(mstrPrivs, Val(mrsTrans!����), CStr(nvl(mrsTrans!NO)), 2, 2, mrsTrans!����ID, mrsTrans!��ҳid, 3) = False Then Exit Sub
                        If IsReceiptBalance_Charge(1, mstrPrivs, Val(mrsTrans!����), CStr(nvl(mrsTrans!NO)), Val(nvl(mrsTrans!�������, 0)), 2, 2, 3) = False Then Exit Sub
                    ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
                        If IsOutPatient(mstrPrivs, Val(mrsTrans!����), CStr(nvl(mrsTrans!NO)), 2, 2, mrsTrans!����ID, mrsTrans!��ҳid, 4) = False Then Exit Sub
                        If IsReceiptBalance_Charge(1, mstrPrivs, Val(mrsTrans!����), CStr(nvl(mrsTrans!NO)), Val(nvl(mrsTrans!�������, 0)), 2, 2, 4) = False Then Exit Sub
                    End If
                    
                    .MoveNext
                Loop
            End With
        
            If strID = "" Then
                MsgBox "��ѡ��Ҫȡ������Һ���ݣ�", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            With Me.VSFLook
                For intRow = 1 To .rows - 1
                    If lng��ҩid <> Val(.TextMatrix(intRow, .ColIndex("��ҩid"))) And .TextMatrix(intRow, .ColIndex("����״̬")) = "�Ѱ�ҩ" Then
                        lng��ҩid = Val(.TextMatrix(intRow, .ColIndex("��ҩid")))
                        strID = IIf(strID = "", "", strID & ",") & lng��ҩid
                    End If
                Next
            End With
        End If
    End If
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    If mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
        arrExecute = GetArrayByStr(strID, 3950, "|")
    Else
        arrExecute = GetArrayByStr(strID, 3950, ",")
    End If
    For i = 0 To UBound(arrExecute)
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            strMsg = "ȡ����ҩ"
            
            gstrSQL = "Zl_��Һ��ҩ��¼_ȡ����ҩ("
            '��ҩID��
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            gstrSQL = gstrSQL & ")"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
            strMsg = "ȡ����ҩ"
            
            gstrSQL = "Zl_��Һ��ҩ��¼_ȡ����ҩ("
            '��ҩID��
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            gstrSQL = gstrSQL & ")"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SENDED Then
            strMsg = "ȡ������"
            
            gstrSQL = "Zl_��Һ��ҩ��¼_ȡ������("
            '��ҩID��
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            gstrSQL = gstrSQL & ")"
        ElseIf mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
            strMsg = "ȡ�����"
            
            gstrSQL = "Zl_��Һ��ҩ��¼_���("
            'ҽ��ID��
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "|'"
            gstrSQL = gstrSQL & ")"
        End If
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, strMsg)
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
'    If mcondition.strTransStep >= M_STR_CALSS_PREPARE Then
'
'        '�������ݼ�����
'        Call DelTransRec
'
'        mrsTrans.Filter = ""
'        Call ShowTrans
'    End If
'
    'ˢ��
    
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        Me.VSFLook.rows = 1
    End If
    
    Call RefreshDeptList(Me.tabDeptList.Selected.index)
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Prepare(ByVal intType As Integer)
    'PIVA��������ҩȷ��
    Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim strPrintID As String    '��ҩID,ƿǩ��|��ҩ��,ƿǩ��...
    Dim arrParams
    Dim blnBeginTrans As Boolean
    Dim str�շ�ID�� As String
    Dim arrExecute As Variant
    Dim i As Integer
    Dim arrSql As Variant
    Dim strInput As String
    Dim blnlock As Boolean
    Dim strOn As String
    Dim strOff As String
    Dim dateNow As Date
    Dim curPrepareNo As Currency
    Dim rsExcStatus As ADODB.Recordset
    Dim strExcID As String
    Dim intCount As Integer
    Dim strExcLable As String
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    arrSql = Array()

    With mrsTrans
        .Filter = "ִ�б�־=1 and �Ƿ�ȷ�ϵ���=0"
        If Not .EOF Then
            If MsgBox("��ǰ������δȷ�ϵ�������Һ�����Ƿ��ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        .Filter = "ִ�б�־=1 " & IIf(intType = 1, "", " and �Ƿ���=1")
        .Sort = "����,��ҩ����,����"
        
        Do While Not .EOF
           
            If InStr(1, "," & str�շ�ID�� & ",", "," & !�շ�ID & ",") = 0 Then
                str�շ�ID�� = IIf(str�շ�ID�� = "", "", str�շ�ID�� & ",") & !�շ�ID
            End If
            
            .MoveNext
        Loop
    End With
    
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                    
                    If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                        strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                    End If
                    
                    If InStr(1, "|" & strPrintID & "|", "|" & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & "," & .TextMatrix(lngRow, .ColIndex("ƿǩ��")) & "|") = 0 Then
                        strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & "," & .TextMatrix(lngRow, .ColIndex("ƿǩ��"))
                    End If
                
                    If .TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = 1 Then
                        blnlock = True
                    End If
                End If
            End If
        Next
    End With
    
    If strInputID = "" Then
        MsgBox "��ѡ��Ҫ��ҩ����Һ���ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '����쳣״̬
    If strInputID <> "" Then
        strInputID = "," & strInputID
        
        Set rsExcStatus = PIVA_GetExcStatus(strInputID, 1)
        If Not rsExcStatus Is Nothing Then
            Do While Not rsExcStatus.EOF
                '��¼�쳣״̬����Һ��id
                strExcID = IIf(strExcID = "", "", strExcID & ",") & rsExcStatus!Id
                
                '�Ӵ����͵���Һ������ȥ���쳣״̬����Һ��id
                strInputID = Replace(strInputID, "," & rsExcStatus!Id, "")
                
                intCount = intCount + 1
                
                '��¼���5��ƿǩ��������ʾ
                If intCount <= 5 Then
                    strExcLable = IIf(strExcLable = "", "", strExcLable & vbCrLf) & rsExcStatus!ƿǩ��
                End If
                
                rsExcStatus.MoveNext
            Loop
        End If
        
        'ȥ��ǰ���","
        If strInputID <> "" Then
            strInputID = Mid(strInputID, 2)
        End If
        
        '��֯��ʾ����
        If strExcLable <> "" Then
            strExcLable = "ע�⣺������Һ�����ܰ�ҩ�������ѱ������˰�ҩ�����ˣ�" & vbCrLf & strExcLable
            
            If intCount > 5 Then
                strExcLable = strExcLable & vbCrLf & "��������" & intCount - 5 & "����Һ��......"
            End If
        End If
    End If
    
    '�����쳣���ݺʹ������ݵ�����ֱ���ʾ
    If strExcLable <> "" Then
        '���쳣����ʱ
        If strInputID = "" Then
            '��ѡ��Ķ����쳣����ʱ
            MsgBox strExcLable & vbCrLf & "��ѡ�����Һ�����ѱ������˰�ҩ�����ˣ�������ѡ��", vbInformation, gstrSysName
                       
            'ˢ��
            Call RefreshDeptList(Me.tabDeptList.Selected.index)
            Call RefreshDetailList(Me.tabDeptList.Selected.index)
            
            Exit Sub
        Else
            '�ų��쳣�����⻹����������ʱ
            If MsgBox(strExcLable & vbCrLf & "�Ƿ��ʣ�����Һ����ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    Else
        If MsgBox("�Ƿ��ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    
    '�����
    If CheckStock = False Then Exit Sub
    
    '�Ա�ҩ���
    If Check�Ա�ҩ = False Then Exit Sub
    
    '���۹���
    If CheckPriceAdjustByID = False Then Exit Sub
    
    '����Ƿ����
    mrsTrans.Filter = "ִ�б�־=1"
    mrsTrans.Sort = "����, NO, �������"
    Do While Not mrsTrans.EOF
        If IsOutPatient(mstrPrivs, Val(mrsTrans!����), CStr(nvl(mrsTrans!NO)), 2, 2, mrsTrans!����ID, mrsTrans!��ҳid, 2) = False Then Exit Sub
        If IsReceiptBalance_Charge(1, mstrPrivs, Val(mrsTrans!����), CStr(nvl(mrsTrans!NO)), Val(nvl(mrsTrans!�������, 0)), 2, 2, 2) = False Then Exit Sub
        
        mrsTrans.MoveNext
    Loop
    
    'ȡ��ҩ����(���ܷ�ҩ��)
    curPrepareNo = Val(zlDatabase.GetNextNo(20))
    
    StrCurDate = Format(Sys.Currentdate, "YYYY-MM-DD HH:MM:SS")
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_��Һ��ҩ��¼_��ҩ("
        '����ID
        gstrSQL = gstrSQL & mcondition.lngCenterID
        '��ҩID
        gstrSQL = gstrSQL & ",'" & arrExecute(i) & "'"
        '��ҩ����
        gstrSQL = gstrSQL & "," & curPrepareNo
        '��ҩ��
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '��ҩʱ��
        gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
        gstrSQL = gstrSQL & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
    Next
    
    If mParams.bln��˻��۵� = True Then
        arrExecute = GetArrayByStr(str�շ�ID��, 3950, ",")
        For i = 0 To UBound(arrExecute)
            gstrSQL = "Zl_סԺ���ʼ�¼_��ҩ���("
            '�շ�ID��
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            '����Ա���
            gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
            '����Ա����
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            '���ʱ��
            gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        Next
    End If
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "סԺ�������")
    Next
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '�������ݼ�����
'    Call DelTransRec
'
'    DoEvents
'
'    mrsTrans.Filter = ""
'    'ˢ����ϸ
'    Call ShowTrans
'    Call ShowSumDrug
'
    '����
    If blnlock Then
        Call SetLock(0, strInputID)
    End If
    'ˢ��
    Call RefreshDeptList(0)
    Call RefreshDetailList(0)
    
    lblCount.Caption = "��Һ����0 �ѣ�0  δ��0 ��ǰѡ����Һ����0"
    
    '��ӡ��ҩ����
    If mParams.int��ҩ���ӡ = 0 Then
        blnPrint = (MsgBox("�Ƿ��ӡ���ΰ�ҩ��ҩƷ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    ElseIf mParams.int��ҩ���ӡ = 1 Then
        blnPrint = True
    End If
    
    If blnPrint = True Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_1", Me, _
            "����=" & mcondition.lngCenterID, _
            "��ҩʱ��=" & StrCurDate, "������Ա=" & gstrUserName, "PrintEmpty=0", 2)
    End If
    
    '��ӡƿǩ
    blnPrint = False
    If mParams.intƿǩ��ҩ���ӡ = 0 Then
        blnPrint = (MsgBox("�Ƿ��ӡ��Һƿǩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    ElseIf mParams.intƿǩ��ҩ���ӡ = 1 Then
        blnPrint = True
    End If
    
    If blnPrint = True Then
        '���ö���������д�ӡ
        Call mfrmPrintPlan.ShowMe(Me, strInputID, mParams.intNum)
'        dateNow = Sys.Currentdate
'        arrExecute = GetArrayByStr(strInputID, 3950, ",")
'        For i = 0 To UBound(arrExecute)
'            Call RefreshPrintSign(arrExecute(i), dateNow)
'        Next
'
'        arrParams = Split(strPrintID, "|")
'        For lngRow = 0 To UBound(arrParams)
'            If arrParams(lngRow) <> "" Then
'                For i = 1 To mParams.intNum
'                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_1", Me, _
'                        "��ҩID=" & Val(Split(arrParams(lngRow), ",")(0)), _
'                        "PrintEmpty=0", 2)
'                Next
'            End If
'        Next
    End If
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PIVAWork_Approve()
    'PIVA���������
    Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim strPrintID As String    '��ҩID,ƿǩ��|��ҩ��,ƿǩ��...
    Dim str�շ�ID�� As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim strInput As String
    Dim strTemp As String
    Dim arrExecute As Variant
    
    On Error GoTo errHandle
        
    Call InitSendMsgRs
    With Me.vsfMedis
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("ҽ��id"))) > 0 And Val(.TextMatrix(lngRow, .ColIndex("��־"))) > 0 Then
                strInputID = strInputID & .TextMatrix(lngRow, .ColIndex("ҽ��id")) & "," & .TextMatrix(lngRow, .ColIndex("��־")) & "|"
                
                If .TextMatrix(lngRow, .ColIndex("��־")) = 2 Then
                    mrsSendMsg.AddNew
                    mrsSendMsg!ҽ��id = .TextMatrix(lngRow, .ColIndex("ҽ��id"))
                    mrsSendMsg!���ͺ� = "111"
                    mrsSendMsg!����ID = .TextMatrix(lngRow, .ColIndex("����ID"))
                    mrsSendMsg!���� = .TextMatrix(lngRow, .ColIndex("����"))
                    mrsSendMsg!סԺ�� = .TextMatrix(lngRow, .ColIndex("סԺ��"))
                    mrsSendMsg!��ҳid = .TextMatrix(lngRow, .ColIndex("��ҳid"))
                    mrsSendMsg!����ID = .TextMatrix(lngRow, .ColIndex("����id"))
                    mrsSendMsg!����ID = .TextMatrix(lngRow, .ColIndex("����id"))
                    mrsSendMsg!���� = .TextMatrix(lngRow, .ColIndex("����"))
                    mrsSendMsg.Update
                End If
            End If
        Next
    End With
    
    If strInputID = "" Then Exit Sub
    
    arrExecute = GetArrayByStr(strInputID, 3950, "|")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_��Һ��ҩ��¼_���("
        'ҽ��ID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "|'"
        '���ҩʦ
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        gstrSQL = gstrSQL & ")"
    
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�״�ִ��ҽ�����")
    Next
    
    '������Ϣ
    Call SendMsgModule
    
    Call RefreshDeptList(0)
    Call RefreshDetailList(0)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Delete()
    'PIVA������ɾ����ҽ�����˶����ϵ���Һ��ҩ��¼
    Dim strID As String
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .strID("ѡ��"))) = -1 Then
                    If InStr(1, "," & strID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                        strID = IIf(strID = "", "", strID & ",") & .TextMatrix(lngRow, .ColIndex("��ҩID"))
                    End If
                End If
            End If
        Next
    End With
    
    If strID = "" Then
        MsgBox "��ѡ��Ҫɾ������Һ���ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("�Ƿ�Ҫɾ�������ϵ���Һ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "Zl_��Һ��ҩ��¼_ɾ��("
    '��ҩID��
    gstrSQL = gstrSQL & "'" & strID & "'"
    gstrSQL = gstrSQL & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ɾ��")
    
    DoEvents
    
'    '�������ݼ�����
'    Call DelTransRec
'
'    DoEvents
'
'    mrsTrans.Filter = ""
'    'ˢ����ϸ
'    Call ShowTrans
'    Call ShowSumDrug
    
    
    'ˢ��
    Call RefreshDeptList(0)
    Call RefreshDetailList(0)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PIVAWork_Dosage(Optional ByVal lng��ҩid As Long, Optional ByVal str����˵�� As String)
    'PIVA��������ҩȷ��
    Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim strPrintID As String    '��ҩID,ƿǩ��|��ҩ��,ƿǩ��...
    Dim arrParams
    Dim blnPrint As Boolean
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    Dim dateNow As Date
    Dim rsExcStatus As ADODB.Recordset
    Dim strExcID As String
    Dim strExcLable As String
    Dim intCount As Integer
    Dim blnlock As Boolean
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    If lng��ҩid = 0 Then
        With vsfTrans
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                    If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                        If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                        End If
                        
                        If InStr(1, "," & strPrintID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                            strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & "," & .TextMatrix(lngRow, .ColIndex("ƿǩ��"))
                        End If
                        
                        If .TextMatrix(lngRow, .ColIndex("�Ƿ�����")) = 1 Then
                            blnlock = True
                        End If
                    End If
                End If
            Next
        End With
    Else
        mlng��ɨ�� = mlng��ɨ�� + 1
        mlngδɨ�� = mlngδɨ�� - 1
        
        strInputID = lng��ҩid
    End If
    
    If strInputID = "" Then Exit Sub
    
    '����쳣״̬
    If strInputID <> "" Then
        strInputID = "," & strInputID
        
        Set rsExcStatus = PIVA_GetExcStatus(strInputID, mTransStatus.��ҩ)
        If Not rsExcStatus Is Nothing Then
            If rsExcStatus.EOF And lng��ҩid <> 0 Then
                lblMsg.Caption = "����"
                
                With Me.vsfDept(0)
                    For i = 1 To .rows - 1
                        If Mid(.TextMatrix(i, .ColIndex("����")), InStr(1, .TextMatrix(i, .ColIndex("����")), "]") + 1) = vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("����")) Then
                            .TextMatrix(i, .ColIndex("����")) = Val(.TextMatrix(i, .ColIndex("����"))) - 1
                            Exit For
                        End If
                    Next
                End With
            ElseIf Not rsExcStatus.EOF And lng��ҩid <> 0 Then
                If rsExcStatus!����״̬ = 4 Then
                    Me.lblMsg.Caption = "��ƿǩ��ɨ��"
                ElseIf rsExcStatus!����״̬ = 5 Then
                    Me.lblMsg.Caption = "��ƿǩ�ѷ���"
                ElseIf rsExcStatus!����״̬ >= 9 Then
                    Me.lblMsg.Caption = "����ҽ����ֹͣ������"
                ElseIf nvl(rsExcStatus!�Ƿ���, 0) > 0 Then
                    Me.lblMsg.Caption = "��ƿǩ�Ѵ��"
                End If
                Exit Sub
            Else
                Do While Not rsExcStatus.EOF
                    If rsExcStatus!����״̬ <> 2 Then
                    
                        '��¼�쳣״̬����Һ��id
                        strExcID = IIf(strExcID = "", "", strExcID & ",") & rsExcStatus!Id
                        
                        '�Ӵ����͵���Һ������ȥ���쳣״̬����Һ��id
                        strInputID = Replace(strInputID, "," & rsExcStatus!Id, "")
                        
                        intCount = intCount + 1
                        
                        '��¼���5��ƿǩ��������ʾ
                        If intCount <= 5 Then
                            strExcLable = IIf(strExcLable = "", "", strExcLable & vbCrLf) & rsExcStatus!ƿǩ��
                        End If
                    End If
                    rsExcStatus.MoveNext
                Loop
            End If
        End If
        
        'ȥ��ǰ���","
        If strInputID <> "" Then
            strInputID = Mid(strInputID, 2)
        End If
        
        '��֯��ʾ����
        If strExcLable <> "" Then
            strExcLable = "ע�⣺������Һ��������ҩ�������ѱ���������ҩ�����ˣ�" & vbCrLf & strExcLable
            
            If intCount > 5 Then
                strExcLable = strExcLable & vbCrLf & "��������" & intCount - 5 & "����Һ��......"
            End If
        End If
    End If
    
    '�����쳣���ݺʹ������ݵ�����ֱ���ʾ
    If strExcLable = "" Then
        '���쳣����ʱ
        If lng��ҩid = 0 Then If MsgBox("�Ƿ���ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        '���쳣����ʱ
        If strInputID = "" Then
            '��ѡ��Ķ����쳣����ʱ
            MsgBox strExcLable & vbCrLf & "��ѡ�����Һ�����ѱ���������ҩ�����ˣ�������ѡ��", vbInformation, gstrSysName
                       
            'ˢ��
            Call RefreshDeptList(0)
            Call RefreshDetailList(0)
            
            Exit Sub
        Else
            '�ų��쳣�����⻹����������ʱ
            If MsgBox(strExcLable & vbCrLf & "�Ƿ��ʣ�����Һ����ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    StrCurDate = Format(Sys.Currentdate, "YYYY-MM-DD HH:MM:SS")
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_��Һ��ҩ��¼_��ҩ("
        '��ҩID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
        '��ҩ��
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '��ҩʱ��
        gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
        '����˵��
        gstrSQL = gstrSQL & ",'" & str����˵�� & "'"
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��ҩȷ��")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
'
'    '�������ݼ�����
'    Call DelTransRec
'
'    DoEvents
'
'    mrsTrans.Filter = ""
'    'ˢ����ϸ
'    Call ShowTrans
'    Call ShowSumDrug
    
    '����
    If blnlock Then
        Call SetLock(0, strInputID)
    End If
    
    'ˢ��
    If lng��ҩid = 0 Then
        lblCount.Caption = "��Һ����0 �ѣ�0  δ��0 ��ǰѡ����Һ����0"
        Call RefreshDeptList(Me.tabDeptList.Selected.index)
        Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
    
    If lng��ҩid <> 0 Then
        mrsTrans.Filter = ""
        mrsTrans.Filter = "��ҩid=" & lng��ҩid
        Do While Not mrsTrans.EOF
            mrsTrans.Delete (adAffectCurrent)
            mrsTrans.MoveNext
        Loop
        Call SetFilter
        lblCount.Caption = "��Һ����" & mlng��ɨ�� + mlngδɨ�� & " �ѣ�" & mlng��ɨ�� & "  δ��" & mlngδɨ�� & " ��ǰѡ����Һ����0"
        Me.txtFindItem.SetFocus
    End If
    
    '��ӡƿǩ
    If lng��ҩid = 0 Then
        blnPrint = False
        If mParams.intƿǩ��ҩ���ӡ = 0 Then
            blnPrint = (MsgBox("�Ƿ��ӡ��Һƿǩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        ElseIf mParams.intƿǩ��ҩ���ӡ = 1 Then
            blnPrint = True
        End If
        
        If blnPrint = True Then
            '���ö���������д�ӡ
            Call mfrmPrintPlan.ShowMe(Me, strInputID, mParams.intNum)
    '        dateNow = Sys.Currentdate
    '        arrExecute = GetArrayByStr(strInputID, 3950, ",")
    '        For i = 0 To UBound(arrExecute)
    '            Call RefreshPrintSign(arrExecute(i), dateNow)
    '        Next
    '
    '        arrParams = Split(strPrintID, "|")
    '        For lngRow = 0 To UBound(arrParams)
    '            If arrParams(lngRow) <> "" Then
    '                For i = 1 To mParams.intNum
    '                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_1", Me, _
    '                        "��ҩID=" & Val(Split(arrParams(lngRow), ",")(0)), _
    '                        "PrintEmpty=0", 2)
    '                Next
    '            End If
    '        Next
        End If
    End If
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_ReturnVerify()
    'PIVA�������������
    Dim strCurrent As String
    Dim str��ҩid As String
    Dim strNo As String
    Dim str������� As String
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    If MsgBox("�Ƿ�������ˣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strCurrent = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    With mrsTrans
        .Filter = "ִ�б�־>0"
        
        If .EOF Then
            MsgBox "��ѡ��Ҫ������˵���Һ���ݣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '����Ƿ����
        .Filter = "ִ�б�־>0"
        .Sort = "����, NO, �������"
        Do While Not .EOF
            If IsOutPatient(mstrPrivs, Val(!����), CStr(!NO), 2, 2, mrsTrans!����ID, mrsTrans!��ҳid, 1) = False Then Exit Sub
            If IsReceiptBalance_Charge(1, mstrPrivs, Val(!����), CStr(!NO), Val(!�������), 2, 2, 1) = False Then Exit Sub
            
            .MoveNext
        Loop
        
        '��ҩ��¼���ʴ�����"Zl_��Һ��ҩ��¼_�������"��ͳһ������ҩ��������ˣ�
        .Filter = "ִ�б�־>0"
        .Sort = "��ҩID"
        Do While Not .EOF
            If InStr(1, str��ҩid, !��ҩid) = 0 Then
                str��ҩid = IIf(str��ҩid = "", "", str��ҩid & ",") & !��ҩid & "," & !ִ�б�־
            End If
            
            .MoveNext
        Loop
        If str��ҩid <> "" Then
            gstrSQL = "Zl_��Һ��ҩ��¼_�������("
            'str��ҩID
            gstrSQL = gstrSQL & "'" & str��ҩid & "'"
            '�����
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            '���ʱ��
            gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
            gstrSQL = gstrSQL & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, "PIVAWork_ReturnVerify")
        End If
        
        gclsInsure.InitOracle gcnOracle
        
        .Filter = "ִ�б�־=1"
        .Sort = "NO,�������"
        Do While Not .EOF
            If strNo <> !NO Or str������� <> !������� & ":" & !ʵ������ Then
                strNo = !NO
                str������� = !������� & ":" & !ʵ������
                
                'ҽ������
                If Not IsNull(!����) And InStr(1, strMCNO, !NO) = 0 Then
                    MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(!����))
                    MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(!����))
                    strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !���� & _
                            "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
                End If
            End If
            .MoveNext
        Loop
    End With
    
    'ҽ�������������ϴ�������ʱ�ϴ�
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans:
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Next
    End If
    
    'ҽ�������������ϴ�����ɺ��ϴ�
    If strMCNO <> "" Then
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    MsgBox "����""" & CStr(arrMCPar(0)) & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                End If
            End If
        Next
    End If
    
    If MsgBox("����Ҫ��ӡ��ҩ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_2", Me, "��ҩʱ��=" & strCurrent, "��װϵ��=C.סԺ��װ", 2)
    End If
    
'    '�������ݼ�����
'    Call DelTransRec
'
'    DoEvents
'
'    mrsTrans.Filter = ""
'    'ˢ����ϸ
'    Call ShowTrans
'    Call ShowSumDrug
    
    'ˢ��
    Call RefreshDeptList(0)
    Call RefreshDetailList(0)
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Refuse()
'ȷ�Ͼܷ�
Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                    If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                        strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("��ҩID"))
                    End If
                End If
                
            End If
        Next
    End With
    
    If strInputID = "" Then
        MsgBox "��ѡ��Ҫȷ�Ͼܾ�����Һ���ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("�Ƿ�ȷ�Ͼܾ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    StrCurDate = Format(Sys.Currentdate, "YYYY-MM-DD HH:MM:SS")
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_��Һ��ҩ��¼_ȷ�Ͼܾ�("
        '��ҩID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
        '�ܾ���
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ȷ�Ͼܾ�")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    'ˢ��
    Call RefreshDeptList(Me.tabDeptList.Selected.index)
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Send(Optional ByVal lng��ҩid As Long, Optional ByVal str����˵�� As String)
    'PIVA����������ȷ��
    Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    Dim rsExcStatus As ADODB.Recordset
    Dim strExcID As String
    Dim strExcLable As String
    Dim intCount As Integer
    Dim blnAutoPrint As Boolean
    Dim lng����id As Long
    Dim str�������� As String
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    If lng��ҩid = 0 Then
        With vsfTrans
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) > 0 Then
                    If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                        If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) & ",") = 0 Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("��ҩID"))
                        End If
                    End If
                    
                End If
            Next
        End With
    Else
        mlng��ɨ�� = mlng��ɨ�� + 1
        mlngδɨ�� = mlngδɨ�� - 1
        With Me.vsfDept(0)
            For i = 1 To .rows - 1
                If Mid(.TextMatrix(i, .ColIndex("����")), InStr(1, .TextMatrix(i, .ColIndex("����")), "]") + 1) = vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("����")) Then
                    .TextMatrix(i, .ColIndex("����")) = Val(.TextMatrix(i, .ColIndex("����"))) - 1
                    blnAutoPrint = (.TextMatrix(i, .ColIndex("����")) = 0)
                    lng����id = Val(.TextMatrix(i, .ColIndex("����id")))
                    str�������� = vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("����"))
                    Exit For
                End If
            Next
        End With
        strInputID = lng��ҩid
    End If
    
    '����쳣״̬
    If strInputID <> "" Then
        strInputID = "," & strInputID
        
        Set rsExcStatus = PIVA_GetExcStatus(strInputID, mTransStatus.��ҩ)
        If Not rsExcStatus Is Nothing Then
            If rsExcStatus.EOF And lng��ҩid <> 0 Then
                lblMsg.Caption = "����"
            ElseIf Not rsExcStatus.EOF And lng��ҩid <> 0 Then
                If rsExcStatus!����״̬ = 5 Then
                    Me.lblMsg.Caption = "��ƿǩ��ɨ��"
                ElseIf rsExcStatus!����״̬ >= 9 Then
                    Me.lblMsg.Caption = "����ҽ����ֹͣ������"
                End If
                Exit Sub
            Else
                Do While Not rsExcStatus.EOF
                    '��¼�쳣״̬����Һ��id
                    strExcID = IIf(strExcID = "", "", strExcID & ",") & rsExcStatus!Id
                    
                    '�Ӵ����͵���Һ������ȥ���쳣״̬����Һ��id
                    strInputID = Replace(strInputID, "," & rsExcStatus!Id, "")
                    
                    intCount = intCount + 1
                    
                    '��¼���5��ƿǩ��������ʾ
                    If intCount <= 5 Then
                        strExcLable = IIf(strExcLable = "", "", strExcLable & vbCrLf) & rsExcStatus!ƿǩ��
                    End If
                    
                    rsExcStatus.MoveNext
                Loop
            End If
        End If
        
        
        'ȥ��ǰ���","
        If strInputID <> "" Then
            strInputID = Mid(strInputID, 2)
        End If
        
        '��֯��ʾ����
        If strExcLable <> "" Then
            strExcLable = "ע�⣺������Һ�����ܷ��ͣ������ѱ������˷��ͻ����ˣ�" & vbCrLf & strExcLable
            
            If intCount > 5 Then
                strExcLable = strExcLable & vbCrLf & "��������" & intCount - 5 & "����Һ��......"
            End If
        End If
    End If
    
    '�����쳣���ݺʹ������ݵ�����ֱ���ʾ
    If strExcLable = "" Then
        '���쳣����ʱ
        If lng��ҩid = 0 Then If MsgBox("�Ƿ��ͣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        '���쳣����ʱ
        If strInputID = "" Then
            '��ѡ��Ķ����쳣����ʱ
            MsgBox strExcLable & vbCrLf & "��ѡ�����Һ�����ѱ������˷��ͻ����ˣ�������ѡ��", vbInformation, gstrSysName
                       
            'ˢ��
            Call RefreshDeptList(Me.tabDeptList.Selected.index)
            Call RefreshDetailList(Me.tabDeptList.Selected.index)
            
            Exit Sub
        Else
            '�ų��쳣�����⻹����������ʱ
            If MsgBox(strExcLable & vbCrLf & "�Ƿ���ʣ�����Һ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    StrCurDate = Format(Sys.Currentdate, "YYYY-MM-DD HH:MM:SS")
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_��Һ��ҩ��¼_����("
        '��ҩID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
        '������
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '����ʱ��
        gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
        '����˵��
        gstrSQL = gstrSQL & ",'" & str����˵�� & "'"
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ȷ��")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '��ӡ���ͱ���
    If lng��ҩid = 0 Then
        If mParams.int���ͺ��ӡ = 0 Then
            blnPrint = (MsgBox("�Ƿ��ӡ���η��͵�ҩƷ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        ElseIf mParams.int���ͺ��ӡ = 1 Then
            blnPrint = True
        End If
        
        If blnPrint = True Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_2", Me, _
            "����=" & mcondition.lngCenterID, _
            "����ʱ��=" & StrCurDate, "������Ա=" & gstrUserName, "PrintEmpty=0", 2)
        End If
    End If
    
    'ˢ��
    If lng��ҩid = 0 Then
        lblCount.Caption = "��Һ����0 �ѣ�0  δ��0 ��ǰѡ����Һ����0"
        Call RefreshDeptList(Me.tabDeptList.Selected.index)
        Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
    
    If lng��ҩid <> 0 Then
        mrsTrans.Filter = ""
        mrsTrans.Filter = "��ҩid=" & lng��ҩid
        Do While Not mrsTrans.EOF
            mrsTrans.Delete (adAffectCurrent)
            mrsTrans.MoveNext
        Loop
        Call SetFilter
        Me.txtFindItem.SetFocus
    End If
    
    DoEvents
    If blnAutoPrint Then
        lblCount.Caption = "��Һ����" & mlng��ɨ�� + mlngδɨ�� & " �ѣ�" & mlng��ɨ�� & "  δ��" & mlngδɨ�� & " ��ǰѡ����Һ����0"
        If Me.cboBatch.Text = "<ȫ��>" Then Exit Sub
        If MsgBox("�Ƿ��ӡ" & str�������� & "����" & Me.cboBatch.Text & "���η��͵�ҩƷ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_4", Me, _
                "����=" & mcondition.lngCenterID, _
                "���˲���=" & lng����id, _
                "��ҩ����=" & Mid(Me.cboBatch.Text, 1, 1), "PrintEmpty=0", 2)
        End If
    End If
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_View_StatusBar '״̬��
            Control.Checked = Me.stbThis.Visible
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3       '����
            Control.Checked = Val(Control.Parameter) = mParams.intFont
        Case MCONMENU_EDIT_PIVA_MedicalRecord               '���Ӳ�������
            If mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
    End Select
End Sub

Private Sub chkAll_Click()
    mint��־ = 0
    Chk_all
End Sub

Private Sub Chk_all()
    Dim lngRow As Long
    Dim str��ҩid As String
    Dim strFilter As String
    
    mstrLastLabel = ""
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))) <> 0 Then
                 str��ҩid = str��ҩid & .TextMatrix(lngRow, .ColIndex("��ҩID")) & ","
                .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(chkAll.Value = 1, -1, 0)
                
                If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
                    If chkAll.Value = 1 Then
                        .TextMatrix(lngRow, .ColIndex("��־")) = "1"
                        .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Me.ImgList.ListImages(3).Picture
                        .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                    Else
                        .TextMatrix(lngRow, .ColIndex("��־")) = "0"
                        .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Nothing
                    End If
                End If
                
                mstrLastLabel = IIf(mstrLastLabel = "", "", mstrLastLabel & ",") & .TextMatrix(lngRow, .ColIndex("ƿǩ��"))
            End If
        Next
    End With
        
    If chkAll.Value = 1 Then
        Call UpdateExeSign(-1, 1)
        
        DoEvents
    Else
        Call UpdateExeSign(-1, 0)
    End If
    
    If mint��־ <> 1 Then
        'ͬ����Ƭ������
        mfrmPIVCard.chkClick Me.chkAll.Value
        mint��־ = 0
    End If
End Sub

Private Sub chkDept_Click()
    Call ShowSumDrug
End Sub


Private Sub chkPack_Click()
    Call ShowSumDrug
End Sub

Private Sub chkSendType_Click(index As Integer)
    Dim n As Integer
    
    If chkSendType(0).Value = 0 And chkSendType(1).Value = 0 Then
        chkSendType(index).Value = 1
    End If
    
    If chkSendType(0).Value = 1 And chkSendType(1).Value = 1 Then
        n = 0
    ElseIf chkSendType(0).Value = 1 Then
        n = 1
    Else
        n = 2
    End If
    
    If n <> Val(fraDetailCtr.Tag) Then
        fraDetailCtr.Tag = n
        
        Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
End Sub

Private Sub chkType_Click(index As Integer)
    Dim n As Integer
    
    If chkType(0).Value = 0 And chkType(1).Value = 0 Then
        chkType(index).Value = 1
    End If
    
    'ͬ����Ƭ���б������
    mfrmPIVCard.CheckType index, chkType(index).Value
    
    If chkType(0).Value = 1 And chkType(1).Value = 1 Then
        n = 0
    ElseIf chkType(0).Value = 1 Then
        n = 1
    Else
        n = 2
    End If
    
    If n <> Val(vsfTrans.Tag) Then
        vsfTrans.Tag = n
        
        Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
End Sub

Private Sub cmdDrug_Click()
    Dim RecReturn As Recordset
    '��ȡҩƷ������
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "������������", mParams.lng��������, mParams.lng��������)
    End If
    
    Set RecReturn = frmSelector.ShowMe(Me, 0, 1, Me.txtDrug.Text, , , mParams.lng��������, , , 0, True, True, True, , , mstrPrivs)
    
    If Not RecReturn.EOF Then
        Me.txtDrug.Text = "(" & RecReturn!ҩƷ���� & ")" & RecReturn!ͨ����
        Me.txtDrug.Tag = RecReturn!ҩƷID
    End If
     
End Sub

Private Sub cmdRefreshTrans_Click()
    Me.cboType.ListIndex = 0
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
    If Me.cboBatch.ListIndex > 0 Then Call cboBatch_Click
    Me.txtFindItem.SetFocus
End Sub

Private Sub CmdSave_Click()
    '����ԭ��
    Dim strSQL As String
    
     On Error GoTo errHandle
     
     If txtLog.Text = "" Then
        MsgBox "����дҩʦ���ԭ��", vbInformation, gstrSysName
        Exit Sub
     End If
     
     strSQL = "Zl_����ҽ����¼_SaveReason("
     strSQL = strSQL & Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("ҽ��id")))
     strSQL = strSQL & ",'" & txtLog.Text & "'"
     strSQL = strSQL & ")"
     
     Call zlDatabase.ExecuteProcedure(strSQL, "����ԭ��")
     Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("ҩʦ���ԭ��")) = txtLog.Text
     
     MsgBox "ҩʦ���ԭ�򱣴�ɹ���", vbInformation, gstrSysName
     Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
   SetListBar
   mblnActive = True
End Sub

Private Function ShowOhters() As Boolean
    '���ܣ����ݵ�ǰ�û������������ã��ж��Ƿ���Ҫ��ʾ[��ȡҩ]��[�Ա�ҩ]
    Dim strSQL As String
    Dim rstemp As ADODB.Recordset
    Dim bln�Ա�ҩ As Boolean
    Dim bln��ȡҩ As Boolean
    
    On Error GoTo errHandle
    
    ShowOhters = False
    
    strSQL = "Select 1 From ��Һ�Ա�ҩ�嵥 Where �Ƿ����� = 1 And Rownum < 2"
    
    Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Һ�Ա�ҩ�嵥")
    
    bln�Ա�ҩ = (Val(zlDatabase.GetPara("�Ա�ҩ��������������", glngSys, 1345, 0)) = 1)
    bln��ȡҩ = (Val(zlDatabase.GetPara("��ȡҩ��������������", glngSys, 1345, 0)) = 1)
    
    If Not rstemp.EOF Or bln�Ա�ҩ Or bln��ȡҩ Then
        ShowOhters = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub Form_Load()
    Dim cbrControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim i As Integer
    
    mblnLoad = True
    mblnActive = False
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    mintBeginRow = 0
    mintEndRow = 0
    
    lblMsgComment.Tag = 1
    picUpOrDown.Picture = frmPublic.ImgList.ListImages.Item("DownArrow").Picture
    
    mblnShowOhters = ShowOhters()
    fraTip.Visible = mblnShowOhters
    
    '��ȡȨ�޺Ͳ���
    Call GetPrivs
    Call GetParams
    
    '���ݼ��
    If DependOnCheck = False Then
        Exit Sub
    End If
    stbThis.Panels(3).Text = gstrUserName
    
    '�������Ӳ������Ķ���
    If mobjCISJOB Is Nothing Then
        On Error Resume Next
        Set mobjCISJOB = CreateObject("zl9CISJob.clsCISJob")
        
        If Not mobjCISJOB Is Nothing Then
            Call mobjCISJOB.InitCISJob(gcnOracle, Me, glngSys, mstrPrivs, gobjBrower.mobjEmr)
        End If
        err.Clear
        
        On Error GoTo 0
    End If
    
    mdateToday = Sys.Currentdate
    
    '��ʼ���������
    Set mfrmPIVCard = New frmPIVCard
    Set mfrmPlan = New frmPlan
    Set mfrmPrintPlan = New frmPrintPlan
    
    If Not mParams.bln��� Then
        mcondition.strTransStep = "01"
        fraMedis.Visible = False
    Else
        mcondition.strTransStep = "00"
        fraMedis.Visible = True
    End If
    
    If mParams.bln��� Then
        mcondition.strTransStep = "00"
        If mParams.intShowPass = 1 Then
            For i = 0 To Me.ImgResult.count - 1
                Me.chkResult(i).Visible = True
                Me.ImgResult(i).Visible = True
                
            Next
        Else
            For i = 0 To Me.ImgResult.count - 1
                Me.chkResult(i).Visible = False
                Me.ImgResult(i).Visible = False
                
            Next
        End If
    Else
        mcondition.strTransStep = "01"
        For i = 0 To Me.ImgResult.count - 1
            Me.chkResult(i).Visible = False
            Me.ImgResult(i).Visible = False
            
        Next
        fraMedis.Visible = False
    End If
    
    '��ʼ����������
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        mParams.int�������� = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��������", "1")
    Else
        mParams.int�������� = 1
    End If
    
    mcondition.intTransTimeSel = 0
    mcondition.lngCenterID = mParams.lng��������
    
    '��ҽӿ�
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    
    '���ؽ���ؼ�
    InitPanes
    InitTabControl
    InitComandBars
    
    '�ָ�������ɫ
    picDeptList.BackColor = &HFFFFFF
    picTime.BackColor = &HFFFFFF
    lblNote.BackColor = &HFFFFFF
    lblʱ�䷶Χ.BackColor = &HFFFFFF
    lblTimeBegin.BackColor = &HFFFFFF
    lblTimeEnd.BackColor = &HFFFFFF
    lblName.BackColor = &HFFFFFF
    lblDrug.BackColor = &HFFFFFF
    lblTag.BackColor = &HFFFFFF
    lbldept.BackColor = &HFFFFFF
    
    '���岻��ʾ����
    mstrUnVisble = "��ǰ��;��־;�Ƿ�����;NO;����;������λ;�÷�;ҩƷid;�˲���;�˲�ʱ��;��ӡ��־;��ҩid;�Ƿ���;ԭ����;����ҩ��;��ҳid;����ID;����;��ý;������;��Ӧҽ��ID;"
    mstrUnallowSetColHide = "ѡ��;���;����;����;ƿǩ��;ҩƷ����;����;"
    
    '����Զ��屨��
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '�ָ�����
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        '�ָ����Ի�����
        Call LoadCustomSet
        
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
        
        Call RestoreWinState(Me, App.ProductName)
    End If
    
    Call GetWorkBatchRec
    
    Call InitVsfTrans
    Call InitVsfSum
    Call InitVSFLook
    
    Call SetCommand
    
    '������Һ���������
    Call SetSort
    
    Select Case mcondition.strTransStep
        Case M_STR_CALSS_AUDIT, M_STR_CALSS_PASSEDAUDIT, M_STR_CALSS_FAILAUDIT
            lblFindItem.Caption = "����"
        Case Else
            lblFindItem.Caption = "ƿǩ��"
    End Select
    
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Look)
    If Not objPopup Is Nothing Then
        For Each cbrControl In objPopup.CommandBar.Controls
            cbrControl.Checked = False
            If cbrControl.Caption = lblFindItem.Caption Then
                cbrControl.Checked = True
            End If
        Next
    End If
    
    err = 0
    On Error Resume Next
    Set mobjMipModule = New zl9ComLib.clsMipModule
    Call mobjMipModule.InitMessage(glngSys, mlngMode, mstrPrivs)
    Call AddMipModule(mobjMipModule)
    
    mblnLoad = False
    
    Call Loadʱ�䷶Χ
    
    lblSort.Visible = False
    cboSort.Visible = False
    With cboSort
        .Clear
        .AddItem "0-�������������������"
        .AddItem "1-������ҩƷ����������"
        .AddItem "2-����ý����������"
    End With
    
    If mobjMipModule Is Nothing Then
        picMsg.Visible = False
    Else
        picMsg.Visible = True
    End If
    
    '����ͼ��
    For i = 0 To Me.ImgResult.count
        Me.ImgResult(i).Picture = frmPublic.imgPass.ListImages(i + 1).Picture
    Next
    
    '�����б��ʼ������
    LoadData
End Sub

Private Sub LoadCustomSet()
    Dim cbrMenu As CommandBarControl
    
    mParams.intFont = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "����", 0))
    mParams.intAutoSelect = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "�Զ�ѡ��", 0))
    mParams.strVsfTrans = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��ҩ������ϸ", "")
    mParams.strVsfSum = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "���ܱ���п�", "")
    mParams.strVsfLook = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "�Ѱ�ҩ����п�", "")
    
    Call SetFontSize
    
    mcondition.intTransTimeSel = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "ʱ�䷶Χ", 0))
    If mcondition.intTransTimeSel < 0 Or mcondition.intTransTimeSel > 3 Then
        mcondition.intTransTimeSel = 0
    End If
    
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ShowHistory, , True)
    If Not cbrMenu Is Nothing Then
        cbrMenu.Checked = (mParams.intAutoSelect = 1)
    End If
End Sub

Private Sub SaveCustomSet()
    Dim i As Integer
    Dim str������ As String
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "����", mParams.intFont
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "ʱ�䷶Χ", mcondition.intTransTimeSel
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "�Զ�ѡ��", mParams.intAutoSelect
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��������", mParams.int��������
    
    str������ = ""
    With Me.vsfTrans
        For i = 0 To .Cols - 1
            str������ = IIf(str������ = "", "", str������ & "|") & .ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "��ҩ������ϸ", str������)
    
    str������ = ""
    With Me.vsfSumDrug
        For i = 0 To .Cols - 1
            str������ = IIf(str������ = "", "", str������ & "|") & .ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "���ܱ���п�", str������)
    
    str������ = ""
    With Me.VSFLook
        For i = 0 To .Cols - 1
            str������ = IIf(str������ = "", "", str������ & "|") & .ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name, "�Ѱ�ҩ����п�", str������)
End Sub
Private Function DependOnCheck() As Boolean
    '�������ݼ��
    Dim rsTmp As ADODB.Recordset
    
    DependOnCheck = False
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct A.ID, A.����" & _
        " From ���ű� A, ��������˵�� B " & _
        " Where A.ID = B.����id And B.�������� = '��������' And " & _
        " B.����id In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��') " & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������")
    
    '��ǰ����
    If mParams.lng�������� = 0 Then
'        MsgBox "���ڲ������������õ�ǰ���������ģ�", vbInformation, gstrSysName
        frmPIVAParaSet.Show 1, Me
        Call GetParams
        DependOnCheck = True
        Exit Function
    Else
        Do While Not rsTmp.EOF
            If mParams.lng�������� = rsTmp!Id Then
                mstrCenterName = rsTmp!����
                Exit Do
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    '��鲿����Ա
    gstrSQL = "Select Distinct P.ID, P.����" & _
        " From ���ű� P " & _
        " Where (P.վ�� = '" & gstrNodeNo & "' Or P.վ�� is Null) And P.ID In (Select Distinct A.����id " & _
        " From ������Ա A, ��������˵�� B " & _
        " Where A.��Աid = [1] And A.����id = B.����id And B.�������� = '��������' And " & _
        " B.����id In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��')) And " & _
        " (P.����ʱ�� Is Null Or P.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����������Ա", glngUserId)
    
    If rsTmp.RecordCount = 0 Then
        MsgBox "�㲻����Һ����������Ա������ʹ�ñ�ģ�飡", vbInformation, gstrSysName
        Exit Function
    End If
    
    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picCondition.hWnd
    End Select
End Sub

Private Sub InitTabControl()
    '��ʼ����ҳ�ؼ�
    Dim lngColor As Long
    
    With Me.tabDeptList
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(CNUMWORK, "ҵ��", picWork.hWnd, 0).Tag = "ҵ��_"
        .InsertItem(CNUMLOOK, "�鿴", picLook.hWnd, 0).Tag = "�鿴_"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
        
'        Call SetTabColor(tabDeptList)
    End With
    
    
    With Me.tabWork
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
            .Position = xtpTabPositionTop
        End With
        
        Me.fraH.Tag = "1"
        
        If mParams.bln��� Then
            .InsertItem(0, "���ҽ��(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_AUDIT
            
            lblCount.Visible = False
            vsfMedis.Visible = True
            If mParams.intShowPass = 1 Then fraMedis.Visible = True
            vsfTrans.Visible = False
    
            fraDetailCtr.Visible = False
            vsfSumDrug.Visible = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("��")) = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("ѡ��")) = True
            Me.fraH.Tag = "2"
        Else
            fraMedis.Visible = False
        End If
        
        .InsertItem(1, "��ҩӡǩ(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_PREPARE
        .InsertItem(2, "��ҩ�˲�(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_DOSAGE
        .InsertItem(3, "���ͺ˲�(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_SEND
        .InsertItem(4, "�������(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_VERIFY
        
        .Item(1).Selected = True
        .Item(0).Selected = True

        Call SetTabColor(tabWork)
    End With
    
    
    With Me.tbcLook
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
            .Position = xtpTabPositionTop
        End With
        
        Me.fraH.Tag = "1"

        If mParams.bln��� Then
            .InsertItem(0, "�����ͨ��ҽ��(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_PASSEDAUDIT
            .InsertItem(1, "���δͨ��ҽ��(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_FAILAUDIT
        
            lblCount.Visible = False
            vsfMedis.Visible = True
            vsfTrans.Visible = False
            fraDetailCtr.Visible = False
            vsfSumDrug.Visible = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("��")) = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("ѡ��")) = True
            Me.fraH.Tag = "2"
        End If
        
        .InsertItem(2, "�ѷ��Ͳ鿴(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_SENDED
        .InsertItem(3, "��ǩ�ղ鿴(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_SIGNED
        .InsertItem(4, "�ܾ�ǩ�ղ鿴(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_REFUSETOSIGN
        .InsertItem(5, "��������˲鿴(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_INVALID
        .InsertItem(6, "ҽ�����˲鿴(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_DEVICERETURN
        
        .Item(1).Selected = True
        .Item(0).Selected = True
        
        Call SetTabColor(tbcLook)
    End With
    
    With Me.tbcDetail
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
            .Position = xtpTabPositionTop
        End With
        
        .InsertItem(mDetailType.��Һ���б�, "��Һ���б�(&0)", picDetailList.hWnd, 0).Tag = "��Һ���б�_"
        .InsertItem(mDetailType.��Һ����Ƭ, "��Һ����Ƭ(&1)", mfrmPIVCard.hWnd, 0).Tag = "��Һ����Ƭ_"
        .InsertItem(mDetailType.ҩƷ�����б�, "ҩƷ�����б�(&2)", picDetailList.hWnd, 0).Tag = "ҩƷ�����б�_"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
        
        If mParams.bln��� Then .Item(mDetailType.ҩƷ�����б�).Visible = False
        If mParams.bln��� Then .Item(mDetailType.��Һ���б�).Caption = "����ҽ���б�"
    End With
End Sub

Private Sub InitComandBars()
    '��ʼ���˵�������ȫ���˵����������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim cbrControlCustom As CommandBarControlCustom
    
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
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = frmPublic.imgPIVA.Icons
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsMain.ActiveMenuBar.Title = "�˵�"
    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.Id = mconMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ(&P)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Excel, "�����&Excel��")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrint, "���ݴ�ӡ(&B)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintWait, "��ӡҩƷ��ҩ��(&C)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintTotal, "��ӡ�����嵥(&W)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintReturn, "��ӡ��ҩ�����嵥(&W)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintSum, "��ӡ���ܱ���(&S)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintNext, "����ƿǩ(&W)")
             
        Set cbrControlMain = .Add(xtpControlButtonPopup, mconMenu_File_PIVA_BillPrintLable, "��ӡ��ǩ(&R)")
        cbrControlMain.Visible = mParams.blnƿǩ�ֹ���ӡ
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelRow, "��ӡ��ǰ��¼(&R)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelBatch, "��ӡ��ǰ����(&B)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelDept, "��ӡ��ǰ����(&D)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelPati, "��ӡ��ǰ����(&P)", -1, False)
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelSendNo, "��ӡ��ǰ��ҩ����(&S)", -1, False)
        cbrControl.BeginGroup = True
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_AllRow, "��ӡ����ѡ��ļ�¼(&A)", -1, False)
        cbrControl.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Parameter, "��������(&T)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�(&X)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.Id = mconMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Cancel, "ȡ��(&C)")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Approve, "���(&A)")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Lock, "����(&S)")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_UnLock, "����(&S)")
        cbrControlMain.Visible = False
        
        '��������
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Beach, "��������(&B)")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_SURE, "ȷ�ϵ���(&O)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Prepare, "��ҩ(&H)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Dosage, "��ҩ(���)(&R)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Send, "����(&B)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_REFUSE, "ȷ�Ͼܾ�(&R)")
        cbrControlMain.Visible = False
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Delete, "ɾ��(&D)")
'        cbrControlMain.BeginGroup = True
'        cbrControlMain.Visible = False

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_ReVerify, "����ȷ��(&V)")
        cbrControlMain.Visible = False
        
'        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_PLAN, "�Ű�(&P)")
'        cbrControlMain.BeginGroup = True
        
        If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, cbrMenuBar.CommandBar.Controls, mconMenu_Edit_PIVA_PASS, mconMenu_Edit_PIVA_PASS)
        
        '��Ҳ�������չ����
        Call zlPlugIn_SetMenu(glngSys, glngModul, mobjPlugIn, cbrMenuBar.CommandBar.Controls, mconMenu_Edit_PlugIn)
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_PlanPopup, "�Ű�����(&E)", -1, False)
    cbrMenuBar.Id = mconMenu_PlanPopup
    
    With cbrMenuBar.CommandBar.Controls
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_PLAN_PIVA_DESK, "��Һ̨����(&D)")
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_PLAN_PIVA_DESKDRUG, "��Һ̨ҩƷ����(&M)")
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_PLAN_PIVA_PERWORK, "��Ա��������(&P)")
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.Id = mconMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_ToolBar, "������(&T)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        cbrControl.Checked = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_StatusBar, "״̬��(&S)")
        cbrControlMain.Checked = True
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_FontSize, "����(&F)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_1, "С����(&S)", -1, False)
        cbrControl.Parameter = 0
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_2, "������(&M)", -1, False)
        cbrControl.Parameter = 1
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_3, "������(&B)", -1, False)
        cbrControl.Parameter = 2
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_ShowHistory, "�Զ���ѡ�ϴ�ѡ�����Һ��(&A)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_SORTSET, "�����������(&S)")
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.Id = mconMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "��������(&H)")
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_Help_Web, "&WEB�ϵ�����")
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Home, "������ҳ(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Forum, "������̳(&F)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_About, "����(&A)��")
        cbrControlMain.BeginGroup = True
    End With
    
    '�����
    With Me.cbsMain.KeyBindings
'        .Add FCONTROL, Asc("S"), mconMenu_Edit_Save
'        .Add FCONTROL, Asc("Z"), mconMenu_Edit_Untread
'        .Add FCONTROL, Asc("M"), mconMenu_Edit_Modify
'        .Add FSHIFT, VK_DELETE, mconMenu_Edit_Delete
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_View_Refresh
        .Add 0, VK_F1, mconMenu_Help_Help
    End With

    '���ò����ò˵�
    With Me.cbsMain.Options
        .AddHiddenCommand mconMenu_File_PrintSet
        .AddHiddenCommand mconMenu_File_Excel
    End With
    
    '���ò������з�ʽ�˵�
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_SortPopup, "��������(&P)", -1, False)
    cbrMenuBar.Id = mconMenu_SortPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_SortPopup_ByCode, "����������(&0)")
        cbrControlMain.Checked = (mParams.int�������� = 1)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_SortPopup_ByName, "����������(&1)")
        cbrControlMain.Checked = (mParams.int�������� = 2)
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Cancel, "ȡ��")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButtonPopup, mconMenu_File_PIVA_BillPrintLable, "��ӡ��ǩ")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = mParams.blnƿǩ�ֹ���ӡ
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelRow, "��ӡ��ǰ��¼(&R)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelBatch, "��ӡ��ǰ����(&B)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelDept, "��ӡ��ǰ����(&D)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelPati, "��ӡ��ǰ����(&P)", -1, False)
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelSendNo, "��ӡ��ǰ��ҩ����(&S)", -1, False)
        cbrControl.BeginGroup = True
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_AllRow, "��ӡ����ѡ��ļ�¼(&A)", -1, False)
        cbrControl.BeginGroup = True
         
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintWait, "��ӡ��ҩ��")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Approve, "���")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
               
        
        '����,������ť
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Lock, "����")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_UnLock, "����")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Beach, "��������")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_SURE, "ȷ�ϵ���")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Prepare, "��ҩ")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Dosage, "��ҩ(���)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Send, "����")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_REFUSE, "ȷ�Ͼܾ�")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Delete, "ɾ��")
'        cbrControlMain.BeginGroup = True
'        cbrControlMain.Visible = False

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_ReVerify, "����ȷ��")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_PASS, "����ʷ/����״̬")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = (mParams.intShowPass = 1 And IsInString(gstrprivs, "������ҩ���", ";"))
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��")
        cbrControlMain.BeginGroup = True
        
        '���Ӳ�������
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_MedicalRecord, "���Ӳ�������")
        cbrControlMain.BeginGroup = True
        
        '��Ҳ�������չ����
        Call zlPlugIn_SetToolbar(glngSys, glngModul, mobjPlugIn, cbrToolBar.Controls, mconMenu_Edit_PlugIn)

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlCustom = .Add(xtpControlCustom, mconMenu_View_Find, "����")
        cbrControlCustom.Handle = picFind.hWnd
        cbrControlCustom.Flags = xtpFlagRightAlign
    End With
    
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
    
    '���õ����˵�
    '��ӡ��ǩ
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_OperPopup, "����(&O)", -1, False)
    cbrMenuBar.Id = conMenu_OperPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Oper_Select, "ѡ��(&S)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelRow, "ѡ��ǰ��¼(&R)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelBatch, "ѡ��ǰ����(&B)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelDept, "ѡ��ǰ����(&D)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_CancleSelDept, "ȡ��ѡ��ǰ����(&C)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelPati, "ѡ��ǰ����(&P)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_CancleSelPati, "ȡ��ѡ��ǰ����(&R)")
        
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelSendNo, "ѡ��ǰ��ҩ����(&S)")
        cbrControl.BeginGroup = True
        
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelMed, "ѡ�����п���ҩ��(&M)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelAll, "ѡ�����м�¼(&A)")
        cbrControl.BeginGroup = True
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Oper_PrintLabel, "��ӡ��ǩ(&P)")
        objPopup.IconId = mconMenu_File_PIVA_BillPrintLable
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelRow, "��ӡ��ǰ��¼(&R)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelBatch, "��ӡ��ǰ����(&B)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelDept, "��ӡ��ǰ����(&D)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelPati, "��ӡ��ǰ����(&P)")
        
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelSendNo, "��ӡ��ǰ��ҩ����(&S)")
        cbrControl.BeginGroup = True
        
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_AllRow, "��ӡ����ѡ��ļ�¼(&A)")
        cbrControl.BeginGroup = True
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Oper_Bag, "���(&B)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Bag_Batch, "�����ǰ����(&B)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Bag_All, "������м�¼(&B)")
        
        '�鿴����ҽ���������Ϣ
        Set cbrControl = .Add(xtpControlButton, conMenu_Oper_Look, "���Ӳ�������(&I)")
                
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Oper_DelBatch, "ɾ������(&D)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_SelRow, "ɾ����ǰ������(&R)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_SelBatch, "ɾ����ǰ����(&B)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_SelDept, "ɾ����ǰ�в�������(&D)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_SelPati, "ɾ����ǰ�в�������(&P)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_AllRow, "ɾ������ѡ���������(&A)")
    End With
    
    '���õ����˵���PASS
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_PASS, "PASS��&P)", 1, False)
    cbrMenuBar.Id = mconMenu_PASS
    cbrMenuBar.Visible = False
'    With cbrMenuBar.CommandBar.Controls
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 0, "ҩ���ٴ���Ϣ�ο�(&C)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 1, "ҩƷ˵����(&D)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 2, "�й�ҩ��(&N)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 3, "������ҩ����(&S)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 4, "����ֵ(&T)")
'
'        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_PASS_Item + 5, "ר����Ϣ(&P)")
'        cbrControlMain.BeginGroup = True
'
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 0, "ҩ��-ҩ���໥����(&D)", -1, False)
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 1, "ҩ��-ʳ���໥����(&F)", -1, False)
'
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 2, "����ע�������(&M)", -1, False)
'        cbrControl.BeginGroup = True
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 3, "����ע�������(&T)", -1, False)
'
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 4, "����֢(&C)", -1, False)
'        cbrControl.BeginGroup = True
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 5, "������(&S)", -1, False)
'
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 6, "��������ҩ(&G)", -1, False)
'        cbrControl.BeginGroup = True
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 7, "��ͯ��ҩ(&P)", -1, False)
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 8, "��������ҩ(&E)", -1, False)
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 9, "��������ҩ(&L)", -1, False)
'
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 6, "ҽҩ��Ϣ����(&I)")
'        cbrControlMain.BeginGroup = True
'
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 7, "ҩƷ�����Ϣ(&M)")
'        cbrControlMain.BeginGroup = True
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 8, "��ҩ;�������Ϣ(&R)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 9, "ҽԺҩƷ��Ϣ(&F)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 10, "���(&S)")
'    End With
    
    '���õ����˵�������
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_Look, "������Ŀ��&F)", 1, False)
    cbrMenuBar.Id = mconMenu_Look
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 1, "ƿǩ��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 2, "סԺ��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 3, "����")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 4, "����")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 5, "��ҩ����")
    End With
    
     '�����������˲˵�������
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_Filter, "������Ŀ��&F)", 1, False)
    cbrMenuBar.Id = mconMenu_Filter
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Filter + 1, "����")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Filter + 2, "סԺ��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Filter + 3, "����")
    End With
End Sub

Private Sub Loadʱ�䷶Χ()
    Dim dteTime As Date
    
    dteTime = Sys.Currentdate
    Dtp��ʼʱ��.Value = Format(dteTime, "yyyy-MM-dd") & " 00:00:00"
    Dtp����ʱ��.Value = Format(dteTime, "yyyy-MM-dd") & " 23:59:59"
    
    With cboʱ�䷶Χ
        .Clear
        .AddItem "0-����"
        .AddItem "1-����"
        .AddItem "2-���պ�����"
        .AddItem "3-ָ��ʱ�䷶Χ"
    End With
    
    With cboʱ�䷶Χ
        .ListIndex = mcondition.intTransTimeSel
        
        If .ListIndex <> Val(.Tag) Then
            .Tag = .ListIndex
        End If
        
        If .ListIndex = 0 Then
            Dtp��ʼʱ��.Value = CDate(Format(dteTime, "YYYY-MM-DD") & " 00:00:00")
            Dtp����ʱ��.Value = CDate(Format(dteTime, "YYYY-MM-DD") & " 23:59:59")
        ElseIf .ListIndex = 1 Then
            Dtp��ʼʱ��.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD") & " 00:00:00")
            Dtp����ʱ��.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD") & " 23:59:59")
        ElseIf .ListIndex = 2 Then
            Dtp��ʼʱ��.Value = CDate(Format(dteTime, "YYYY-MM-DD") & " 00:00:00")
            Dtp����ʱ��.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD") & " 23:59:59")
        ElseIf .ListIndex = 3 Then
            If mcondition.strTransStartTime = "" Then
                mcondition.strTransStartTime = Format(dteTime, "YYYY-MM-DD") & " 00:00:00"
            End If
            If mcondition.strTransEndTime = "" Then
                mcondition.strTransEndTime = Format(dteTime, "YYYY-MM-DD") & " 23:59:59"
            End If
            
            Dtp��ʼʱ��.Value = CDate(Format(mcondition.strTransStartTime, "YYYY-MM-DD") & " 00:00:00")
            Dtp����ʱ��.Value = CDate(Format(mcondition.strTransEndTime, "YYYY-MM-DD") & " 23:59:59")
        End If
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Height < 9000 Then Me.Height = 9000
    If Me.Width < 12000 Then Me.Width = 12000
    
    ResizeConditionArea
    picDetailList_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsDeptAdvice = Nothing
    Set mrsTrans = Nothing
    Set mrsDeptTrans = Nothing
    
    Set mobjCISJOB = Nothing
        
    '���洰�ڼ�����
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
        
        Call SaveWinState(Me, App.ProductName)
    
        '������Ի�����
        Call SaveCustomSet
    End If
    
    'ж����Ϣ����
    If Not mobjMipModule Is Nothing Then
        Call mobjMipModule.CloseMessage
        Call DelMipModule(mobjMipModule)
        Set mobjMipModule = Nothing
    End If
    mcondition.strTransStep = ""
    
    If Not mfrmPIVCard Is Nothing Then Unload mfrmPIVCard
    If Not mfrmPlan Is Nothing Then Unload mfrmPlan
    If Not mfrmPrintPlan Is Nothing Then Unload mfrmPrintPlan
    
    'ж����ҽӿ�
    Call zlPlugIn_Unload(mobjPlugIn)
    
    Unload Me
End Sub

Private Sub lblTransDrug_Click()
'    If Val(picHscTransDrug.Tag) = "1" Then
'        Call imgDown_Click
'    Else
'        Call imgUp_Click
'    End If
End Sub

Private Sub fraH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfTrans.Height + y <= 1200 Then Exit Sub
        If VSFLook.Height - y < 1200 Then Exit Sub

        fraH.Top = fraH.Top + y
        VSFLook.Top = VSFLook.Top + y
        VSFLook.Height = VSFLook.Height - y
        vsfTrans.Height = vsfTrans.Height + y
        
        txtLog.Top = txtLog.Top + y
        txtLog.Height = txtLog.Height - y
        Me.vsfMedis.Height = vsfMedis.Height + y
        Me.Refresh
    End If
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfTrans.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfTrans.ColHidden(.RowData(i)) Or vsfTrans.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = vsfTrans.Top + vsfTrans.RowHeight(0) + 30
                
                If .Top + .Height > Me.ScaleHeight - vsfTrans.Top Then
                    .Height = Me.ScaleHeight - .Top - vsfTrans.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                
                .Left = vsfTrans.Left
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub


Private Sub lblFindItem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    If Button = 1 Then
        If Me.cbsMain Is Nothing Then Exit Sub
        
        Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Look)
        If Not objPopup Is Nothing Then
            For Each cbrControl In objPopup.CommandBar.Controls
                If cbrControl.Caption = "��ҩ����" Then
                    If mcondition.strTransStep = M_STR_CALSS_AUDIT _
                        Or mcondition.strTransStep = M_STR_CALSS_PREPARE _
                        Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
                        Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
                        cbrControl.Visible = False
                    Else
                        cbrControl.Visible = True
                    End If
                ElseIf cbrControl.Caption = "ƿǩ��" Then
                    If mcondition.strTransStep = M_STR_CALSS_AUDIT _
                        Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
                        Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
                        cbrControl.Visible = False
                    Else
                        cbrControl.Visible = True
                    End If
                Else
                    cbrControl.Visible = True
                End If
            Next
                
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub


Private Sub lblName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    If Button = 1 Then
        If Me.cbsMain Is Nothing Then Exit Sub
        
        Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Filter)
        If Not objPopup Is Nothing Then
'            For Each cbrControl In objPopup.CommandBar.Controls
'                If cbrControl.Caption = "��ҩ����" Then
'                    If mcondition.strTransStep = M_STR_CALSS_AUDIT _
'                        Or mcondition.strTransStep = M_STR_CALSS_PREPARE _
'                        Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
'                        Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
'                        cbrControl.Visible = False
'                    Else
'                        cbrControl.Visible = True
'                    End If
'                ElseIf cbrControl.Caption = "ƿǩ��" Then
'                    If mcondition.strTransStep = M_STR_CALSS_AUDIT _
'                        Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
'                        Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
'                        cbrControl.Visible = False
'                    Else
'                        cbrControl.Visible = True
'                    End If
'                Else
'                    cbrControl.Visible = True
'                End If
'            Next
                
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub optShowType_Click(index As Integer)
    Call SetTransColHide
End Sub

Private Sub picCondition_Resize()
    Call ResizeConditionArea
End Sub

Private Sub picDept_Resize(index As Integer)
    On Error Resume Next

    With vsfDept(index)
        .Move picDept(index).ScaleLeft, picDept(index).ScaleTop, picDept(index).ScaleWidth, picDept(index).ScaleHeight
    End With
End Sub

Private Sub picDeptList_Resize()
    On Error Resume Next
    
    With Me.tabDeptList
        .Move picDeptList.ScaleLeft, picDeptList.ScaleTop + 150, picDeptList.ScaleWidth, picDeptList.ScaleHeight - 150
    End With
    
    With Me.cmdRefreshTrans
        .Move picDeptList.Width - .Width - 50, Me.tabDeptList.Top - 50
    End With
    
    With Me.chkAllDept
'        .Move picDeptList.Width - .Width - 50, Me.tabDeptList.Top + 50
        .Move cmdRefreshTrans.Left - .Width - 50, Me.tabDeptList.Top + 50
    End With
End Sub


Private Sub picDetail_Resize()
    On Error Resume Next
    
    With fraLineV1
'        .Top = 0
        .Left = 0
        .Height = picDetail.Height + 100
    End With
    
    With tbcDetail
        .Top = 0
        .Left = fraLineV1.Left + 50
        .Width = picDetail.Width - fraLineV1.Width
        .Height = picDetail.Height - 50
    End With
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picDetail.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    
    With fraTip
        .ZOrder 0
        .Top = stbThis.Top + 90
        .Left = stbThis.Panels(2).Left + stbThis.Panels(2).Width - .Width - 50
    End With
End Sub

Private Sub picDetailList_Resize()
    On Error Resume Next
    
    With picHelp
        .Top = 0
        .Left = 0
        .Width = picDetailList.Width
    End With
     
    With fraDetailCtr
        .Top = picHelp.Top + picHelp.Height
        .Left = 0
        .Width = picDetailList.Width - 50
    End With
    
    With Me.fraMedis
        .Top = picHelp.Top + picHelp.Height
        .Left = 0
        .Width = picDetailList.Width - 50
    End With
    
    With vsfTrans
        .Top = fraDetailCtr.Top + fraDetailCtr.Height + 50
        .Left = 0
        .Width = picDetailList.Width - 50
'        .Height = picDetailList.Height - .Top
    End With
    
    With Me.vsfMedis
        .Top = IIf(fraMedis.Visible, fraMedis.Top + fraMedis.Height, picHelp.Top + picHelp.Height) + 50
        .Left = fraDetailCtr.Left
        .Width = Me.vsfTrans.Width
        .Height = picDetailList.Height - 100 - IIf(fraMedis.Visible, fraMedis.Height, 0)
    End With
    
    With vsfSumDrug
        .Top = fraDetailCtr.Top + fraDetailCtr.Height + 50
        .Left = 0
        .Width = picDetailList.Width
        .Height = picDetailList.Height - .Top
    End With
    
    If fraH.Tag = "1" Then
        VSFLook.Visible = True
        Me.fraH.Visible = True
        txtLog.Visible = False
        CmdSave.Visible = False
        Me.txtDia.Visible = False
        Me.lblDia.Visible = False
        Me.lblLog.Visible = False
        
        VSFLook.Top = picDetailList.Height - VSFLook.Height + 20
        VSFLook.Left = 0
        VSFLook.Width = picDetailList.Width - 50
        
        fraH.Top = VSFLook.Top - fraH.Height - 50
        fraH.Left = VSFLook.Left
        fraH.Width = VSFLook.Width
        vsfTrans.Height = fraH.Top - vsfTrans.Top - 50
'        VSFLook.Height = picDetailList.Height - fraH.Top - fraH.Height - 50
    ElseIf fraH.Tag = "2" Then
        txtLog.Text = ""
        txtDia.Text = ""
        VSFLook.Visible = False
        Me.fraH.Visible = False
        txtLog.Visible = True
        CmdSave.Visible = True
        Me.txtDia.Visible = True
        Me.lblDia.Visible = True
        Me.lblLog.Visible = True
        
        
        
        Me.txtLog.Top = picDetailList.Height - txtLog.Height - IIf(Me.tabDeptList.Selected.index = 1, 0, CmdSave.Height) - 50
        Me.lblLog.Top = picDetailList.Height - txtLog.Height - IIf(Me.tabDeptList.Selected.index = 1, 0, CmdSave.Height) - 50 - Me.lblLog.Height
        
        Me.txtDia.Top = picDetailList.Height - txtLog.Height - IIf(Me.tabDeptList.Selected.index = 1, 0, CmdSave.Height) - 50 - Me.lblLog.Height - Me.txtDia.Height
        
        Me.lblDia.Top = picDetailList.Height - txtLog.Height - IIf(Me.tabDeptList.Selected.index = 1, 0, CmdSave.Height) - 50 - Me.lblLog.Height - Me.txtDia.Height - lblDia.Height
        
'        txtLog.Height = VSFLook.Height - CmdSave.Height - 50
        txtLog.Left = 0
        txtLog.Width = picDetailList.Width - 50
        txtDia.Left = 0
        txtDia.Width = picDetailList.Width - 50
        lblDia.Left = 0
        lblLog.Left = 0
        
        CmdSave.Left = txtLog.Left + txtLog.Width / 2
        CmdSave.Top = txtLog.Height + txtLog.Top + 50
        
        fraH.Top = lblDia.Top - fraH.Height - 50
        fraH.Left = txtLog.Left
        fraH.Width = txtLog.Width
        vsfMedis.Height = fraH.Top - vsfMedis.Top - 50
        
        If Me.tabDeptList.Selected.index = 1 Then
            txtLog.Enabled = False
        Else
            txtLog.Enabled = True
        End If
    Else
        VSFLook.Visible = False
        txtLog.Visible = False
        Me.txtDia.Visible = False
        Me.lblDia.Visible = False
        Me.lblLog.Visible = False
        Me.fraH.Visible = False
        CmdSave.Visible = False
        vsfTrans.Height = picDetailList.Height - fraDetailCtr.Height - 300
        vsfMedis.Height = picDetailList.Height - fraDetailCtr.Height - 300
    End If
    
End Sub

Private Sub picHelp_Resize()
    Me.lblCount.Left = picHelp.Width - lblCount.Width - 50
End Sub


Private Sub picLook_Resize()
    On Error Resume Next
     
    Me.tbcLook.Move picLook.ScaleLeft, picLook.ScaleTop, picLook.ScaleWidth, picLook.ScaleHeight
    err.Clear
End Sub

Private Sub picMsg_Resize()
    On Error Resume Next

    With vsfMsg
        .Move picMsg.ScaleLeft, picMsg.ScaleTop + lblMsgComment.Top + lblMsgComment.Height + 100, picMsg.ScaleWidth, .Height
    End With
    
    Me.picUpOrDown.Left = picMsg.Width - picUpOrDown.Width - 50
    fraMsg.Width = picMsg.Width
    
    err.Clear
End Sub

Private Sub picUpOrDown_Click()
    If Me.lblMsgComment.Tag = "1" Then
        Me.lblMsgComment.Tag = "0"
        vsfMsg.Visible = False
        picUpOrDown.Picture = frmPublic.ImgList.ListImages.Item("UpArrow").Picture
    Else
        Me.lblMsgComment.Tag = "1"
        vsfMsg.Visible = True
        picUpOrDown.Picture = frmPublic.ImgList.ListImages.Item("DownArrow").Picture
    End If
    
    Call ResizeConditionArea
End Sub

Private Sub picUpOrDown1_Click()
    If Me.txtTag.Visible = False Then
        lblName.Visible = True
        txtName.Visible = True
        
        lblDrug.Visible = True
        txtDrug.Visible = True
        
        lblTag.Visible = True
        txtTag.Visible = True
        
        lbldept.Visible = True
        txtdept.Visible = True
    Else
        lblName.Visible = False
        txtName.Visible = False
        
        lblDrug.Visible = False
        txtDrug.Visible = False
        
        lblTag.Visible = False
        txtTag.Visible = False
        
        lbldept.Visible = False
        txtdept.Visible = False
    End If
    
     Call ResizeConditionArea
End Sub

Private Sub picWork_Resize()
    On Error Resume Next
     
    Me.tabWork.Move picWork.ScaleLeft, picWork.ScaleTop, picWork.ScaleWidth, picWork.ScaleHeight
    err.Clear
End Sub

Private Sub tabDeptList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnLoad = True Then Exit Sub
    
    chkAllDept.Value = 0
    
    If Item.index = 0 Then
        mcondition.strTransStep = tabWork.Selected.Tag
        Call tabWork_SelectedChanged(tabWork.Selected)
    Else
        mcondition.strTransStep = tbcLook.Selected.Tag
        Call tbcLook_SelectedChanged(tbcLook.Selected)
    End If
    
    DoEvents
    Call SetCommand
    DoEvents
    
    Call RefreshDeptList(Item.index)
    
    Select Case mcondition.strTransStep
        Case M_STR_CALSS_AUDIT, M_STR_CALSS_PASSEDAUDIT, M_STR_CALSS_FAILAUDIT
            lblFindItem.Caption = "����"
        Case Else
            lblFindItem.Caption = "ƿǩ��"
    End Select
    
End Sub

'Private Sub picHscTransDrug_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        If vsfTrans.Height + Y <= 1200 Then Exit Sub
'        If vsfDrug.Height - Y < 1200 Then Exit Sub
'
'        picHscTransDrug.Top = picHscTransDrug.Top + Y
'        vsfDrug.Top = vsfDrug.Top + Y
'        vsfDrug.Height = vsfDrug.Height - Y
'
'        vsfTrans.Height = vsfTrans.Height + Y
'        Me.Refresh
'    End If
'End Sub
    

Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.index
        Case mDetailType.��Һ���б�
            If (mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) And mParams.bln��� Then
                Me.vsfMedis.Visible = True
                vsfTrans.Visible = False
                fraDetailCtr.Visible = False
            Else
                Me.vsfMedis.Visible = False
                vsfTrans.Visible = True
                fraDetailCtr.Visible = True
            End If
            
            vsfSumDrug.Visible = False
        Case mDetailType.ҩƷ�����б�
            vsfTrans.Visible = False
            vsfSumDrug.Visible = True
            vsfMedis.Visible = False
            fraDetailCtr.Visible = True
            Call ShowSumDrug
    End Select
    
    Call SetListBar
    Call ShowComment(Item.index, mcondition.strTransStep)
End Sub

Private Sub tabWork_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intCount As Integer
    Dim intSum As Integer
    Dim strUnvisble As String
    Dim strRows As String
    Dim i As Integer
    
    If mblnLoad = True Then Exit Sub
    
    lblCount.Visible = True
    mcondition.strTransStep = Item.Tag
    
    Call SetTabColor(tabWork)
    
    DoEvents
    Call SetCommand
    DoEvents
    
    mstrLastLabel = ""
    
    chkAllDept.Value = 0
    chkAll.Value = 0
    
    If Me.cboType.ListCount <> 0 Then
        Me.cboType.ListIndex = 0
    End If

    If mParams.bln��� Then
        If mcondition.strTransStep >= M_STR_CALSS_PREPARE Then
            lblFindItem.Caption = "ƿǩ��"
            fraMedis.Visible = False
            vsfMedis.Visible = False
            vsfTrans.Visible = True
            fraDetailCtr.Visible = True
            Me.tbcDetail.Item(mDetailType.ҩƷ�����б�).Visible = True
            lblCount.Visible = True
        Else
            lblFindItem.Caption = "����"
            If mParams.intShowPass = 1 Then
                For i = 0 To Me.ImgResult.count - 1
                    Me.chkResult(i).Visible = True
                    Me.ImgResult(i).Visible = True
                Next
            End If
            fraMedis.Visible = True
            vsfMedis.Visible = True
            vsfTrans.Visible = False
            fraDetailCtr.Visible = False
            vsfSumDrug.Visible = False
            Me.tbcDetail.Item(mDetailType.ҩƷ�����б�).Visible = False
            lblCount.Visible = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("��")) = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("ѡ��")) = True
        End If
    Else
        fraMedis.Visible = False
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        Me.fraH.Tag = "1"
    ElseIf mcondition.strTransStep = M_STR_CALSS_AUDIT Then
        Me.fraH.Tag = "2"
    Else
        Me.fraH.Tag = ""
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        Me.lblVolu.Visible = True
    Else
        Me.lblVolu.Visible = False
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
        Me.lblNote.Caption = "��Һ����������ʱ�䷶Χ"
    Else
        Me.lblNote.Caption = "��Һ��ִ��ʱ�䷶Χ"
    End If
    
    If mParams.bln��� And (mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) Then
        Me.tbcDetail.Item(mDetailType.��Һ���б�).Caption = "����ҽ���б�"
    Else
        Me.tbcDetail.Item(mDetailType.��Һ���б�).Caption = "��Һ���б�"
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        strUnvisble = mstrUnVisble & "��;��ҩ��;��������;��ҩʱ��;��ҩ��;��ҩʱ��;������;����ʱ��;����������;��������ʱ��;���������;�������ʱ��;����ԭ��;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_DOSAGE Then
        strUnvisble = mstrUnVisble & "��;��ҩ��;��������;��ҩʱ��;������;����ʱ��;����������;��������ʱ��;���������;�������ʱ��;����ԭ��;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
        strUnvisble = mstrUnVisble & "��;��ҩ��;��������;��ҩʱ��;������;����ʱ��;����������;��������ʱ��;���������;�������ʱ��;����ԭ��;��ҩ����;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_VERIFY Then
        strUnvisble = mstrUnVisble & "ѡ��;��������;��ҩ��;��ҩʱ��;��ҩ��;��ҩʱ��;������;����ʱ��;���������;�������ʱ��;��ҩ����;"
    End If
    
    Me.VSFLook.rows = 1
    Call picDetailList_Resize
    
    Call SetListBar
    
    vsfColSel.Visible = False
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        strRows = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���", mcondition.strTransStep, "")
    End If
    
    If strRows = "" Then
        strRows = strUnvisble
    End If
    
    If strRows <> "" Then
        For i = 1 To Me.vsfTrans.Cols - 1
            If InStr(1, ";" & strRows & ";", ";" & vsfTrans.ColKey(i) & ";") > 0 Then
                vsfTrans.ColHidden(i) = True
            Else
                vsfTrans.ColHidden(i) = False
            End If
        Next
    End If
    
    Call SetLookMenu
'    Call SetTransColHide
    Call InitColSelList(strUnvisble)
    Call SetSumDrugColHide
    Call ShowComment(tbcDetail.Selected.index, mcondition.strTransStep)
    
    Set mrsTrans = Nothing
    Call ShowDeptTrans(Me.tabDeptList.Selected.index, tabWork.Selected.Tag)
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
    
End Sub

Private Sub tbcLook_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intCount As Integer
    Dim intSum As Integer
    Dim strUnvisble As String
    Dim strRows As String
    Dim i As Integer
    
    If mblnLoad = True Then Exit Sub
    
    lblCount.Visible = True
    mcondition.strTransStep = Item.Tag
    
    Call SetTabColor(tbcLook)
    
    DoEvents
    Call SetCommand
    DoEvents
    
    mstrLastLabel = ""
   
    chkAllDept.Value = 0
    chkAll.Value = 0
    
    If Me.cboType.ListCount <> 0 Then
        Me.cboType.ListIndex = 0
    End If
    
    If mParams.bln��� Then
        If mcondition.strTransStep >= M_STR_CALSS_SENDED Then
            lblFindItem.Caption = "ƿǩ��"
            fraMedis.Visible = False
            vsfMedis.Visible = False
            vsfTrans.Visible = True
            fraDetailCtr.Visible = True
            Me.tbcDetail.Item(mDetailType.ҩƷ�����б�).Visible = True
            lblCount.Visible = True
        Else
            lblFindItem.Caption = "����"
            If mParams.intShowPass = 1 Then
                For i = 0 To Me.ImgResult.count - 1
                    Me.chkResult(i).Visible = True
                    Me.ImgResult(i).Visible = True

                Next
            End If
            fraMedis.Visible = True
            vsfMedis.Visible = True
            vsfTrans.Visible = False
            fraDetailCtr.Visible = False
            vsfSumDrug.Visible = False
            Me.tbcDetail.Item(mDetailType.ҩƷ�����б�).Visible = False
            lblCount.Visible = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("ѡ��")) = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("��")) = True
        End If
    Else
        fraMedis.Visible = False
'        For i = 0 To Me.ImgResult.count - 1
'            Me.chkResult(i).Visible = False
'            Me.ImgResult(i).Visible = False
'
'        Next
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
        Me.fraH.Tag = "2"
    Else
        Me.fraH.Tag = ""
    End If
    
    Me.lblVolu.Visible = False
    If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
        Me.lblNote.Caption = "��Һ����������ʱ�䷶Χ"
    Else
        Me.lblNote.Caption = "��Һ��ִ��ʱ�䷶Χ"
    End If
                                                                                          
    If mParams.bln��� And (mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) Then
        Me.tbcDetail.Item(mDetailType.��Һ���б�).Caption = "����ҽ���б�"
    Else
        Me.tbcDetail.Item(mDetailType.��Һ���б�).Caption = "��Һ���б�"
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_INVALID Then
        strUnvisble = mstrUnVisble & "��;��ҩ��;��ҩʱ��;��ҩ��;��ҩʱ��;������;����ʱ��;����������;��������ʱ��;����ԭ��;��ҩ����;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_SENDED Then
        strUnvisble = mstrUnVisble & "��;��ҩ��;��������;��ҩʱ��;��ҩ��;��ҩʱ��;����������;��������ʱ��;���������;�������ʱ��;����ԭ��;��ҩ����;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_DEVICERETURN Then
        strUnvisble = mstrUnVisble & "��;��ҩ����;��������;ҽ������ʱ��;��ҩ��;��ҩʱ��;��ҩ��;��ҩʱ��;������;����ʱ��;����������;��������ʱ��;���������;�������ʱ��;����ԭ��;"
    Else
        strUnvisble = mstrUnVisble & "��;��������;��ҩ��;��ҩʱ��;��ҩ��;��ҩʱ��;������;����ʱ��;����������;��������ʱ��;���������;�������ʱ��;����ԭ��;��ҩ����;"
    End If
    
    Me.VSFLook.rows = 1
    Call picDetailList_Resize
    
    Call SetListBar
    
    
    vsfColSel.Visible = False
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        strRows = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���", mcondition.strTransStep, "")
    End If
    
    If strRows = "" Then
        strRows = strUnvisble
    End If
    
    If strRows <> "" Then
        For i = 1 To Me.vsfTrans.Cols - 1
            If InStr(1, ";" & strRows & ";", ";" & vsfTrans.ColKey(i) & ";") > 0 Then
                vsfTrans.ColHidden(i) = True
            Else
                vsfTrans.ColHidden(i) = False
            End If
        Next
    End If
'    Call SetTransColHide
    
    Call SetLookMenu
    
    Call InitColSelList(strUnvisble)
    Call SetSumDrugColHide
    Call ShowComment(tbcDetail.Selected.index, mcondition.strTransStep)
    
    Set mrsTrans = Nothing
    Call ShowDeptTrans(Me.tabDeptList.Selected.index, tbcLook.Selected.Tag)
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
End Sub


Private Sub SetLookMenu()
    Dim cbrControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Look)
    
    If Not objPopup Is Nothing Then
        For Each cbrControl In objPopup.CommandBar.Controls
            cbrControl.Checked = False
            If cbrControl.Caption = lblFindItem.Caption Then
                cbrControl.Checked = True
            End If
        Next
    End If
End Sub


Private Sub txtdept_GotFocus()
    Call zlControl.TxtSelAll(txtdept)
End Sub

Private Sub txtdept_KeyPress(KeyAscii As Integer)
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    Dim vRect As RECT
    Dim rstemp As ADODB.Recordset
    
    Me.txtdept.Tag = ""
    On Error GoTo errHandle
    If KeyAscii = 13 Then
        If txtdept.Text <> "" Then
            gstrSQL = " Select A.ID,b.���� As վ������, b.��� As վ��,A.����||'-'||A.���� ���� From ���ű� A, Zlnodelist B " & _
                " Where a.վ�� = b.���(+) And A.ID in (Select ����ID From ��������˵�� Where ��������='�ٴ�' And ������� IN(2,3))" & _
                " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                " And (A.���� Like [1] Or A.���� Like [1] Or A.���� Like [1])"
        Else
            Exit Sub
        End If
        
        gstrSQL = gstrSQL & " Order By a.���� || '-' || a.���� "
        
        '�ж�����¼����ʾ����ѡ��
        vRect = zlControl.GetControlRect(txtdept.hWnd)
        sngX = vRect.Left
        sngY = vRect.Top
        sngH = txtdept.Height
        
        Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ѡ���������", False, "", "ѡ���������", False, False, True, sngX, sngY, sngH, True, False, False, UCase(txtdept.Text) & "%")
        
        If Not rstemp Is Nothing Then
            If Not rstemp.EOF Then
                txtdept.Tag = rstemp!Id
                txtdept.Text = rstemp!����
            End If
        End If
        
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDrug_GotFocus()
    Call zlControl.TxtSelAll(txtDrug)
End Sub

Private Sub txtDrug_KeyPress(KeyAscii As Integer)
    Dim rsReturn As Recordset
    
    Me.txtDrug.Tag = ""
    If KeyAscii = 13 Then
    
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(2, "������������", mParams.lng��������, mParams.lng��������)
        End If
    
        Set rsReturn = frmSelector.ShowMe(Me, 1, 1, Me.txtDrug.Text, , , mParams.lng��������, mParams.lng��������, , 0, True, True, True, , , mstrPrivs)
'        Set RecReturn = frmSelector.showMe(Me, 1, IIf(mint�༭״̬ = 8 Or mbln�˻�, 2, 1), strKey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint�༭״̬ = 8 Or mbln�˻�, Val(txtProvider.Tag), 0), True, True, True, , , mstrPrivs)
    
        If Not rsReturn.EOF Then
            Me.txtDrug.Text = "(" & rsReturn!ҩƷ���� & ")" & rsReturn!ͨ����
            Me.txtDrug.Tag = rsReturn!ҩƷID
        End If
    End If
End Sub

Private Sub txtFinditem_GotFocus()
    Call zlControl.TxtSelAll(txtFindItem)
End Sub

Private Sub txtFinditem_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim StrDate As String
    Dim blnScaner As Boolean
    Dim blnDoIt As Boolean
    Dim intCol As Integer
    Dim blnFindItem As Boolean
    Dim strFind As String
    Dim rstemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    'Me.lblMsg.Caption = ""
    If KeyAscii <> 13 Then
        If lblFindItem.Caption = "ƿǩ��" Then
            blnScaner = InputIsScaner(txtFindItem, KeyAscii)
        End If
    Else
        txtFindItem.Text = Trim(txtFindItem.Text)
        If txtFindItem.Text = "" Then Exit Sub
        blnScaner = InputIsScaner(txtFindItem, KeyAscii)
        blnDoIt = True
    End If
    
    If blnDoIt = True Then
        If mcondition.strTransStep = M_STR_CALSS_AUDIT _
            Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
            Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
            With vsfMedis
                If .rows = 1 Then Exit Sub
                If .TextMatrix(1, .ColIndex("ҽ��ID")) = "" Then Exit Sub
            
            
                '�����Һ���б����Ƿ��в�����Ŀ
                blnFindItem = False
                For intCol = 1 To .Cols - 1
                    If .ColKey(intCol) = lblFindItem.Caption Then
                        blnFindItem = True
                        Exit For
                    End If
                Next
                If blnFindItem = False Then Exit Sub
                
                strFind = txtFindItem.Text
            
                lngRow = .FindRow(strFind, 1, .ColIndex(lblFindItem.Caption))
                 
                If lngRow > 0 Then
                    .Row = lngRow
                    .TopRow = lngRow
                Else
                    MsgBox "û�ҵ�" & Me.lblFindItem.Caption & "Ϊ[" & strFind & "]��ҽ���� ��", vbInformation, gstrSysName
                    If tbcDetail.Item(mDetailType.��Һ���б�).Selected Then txtFindItem.SetFocus
                End If
            
                txtFindItem.Text = ""
                If tbcDetail.Item(mDetailType.��Һ���б�).Selected Then txtFindItem.SetFocus
            End With
        Else
            With vsfTrans
                
                
                If .rows = 1 Then Exit Sub
                If .TextMatrix(1, .ColIndex("��ҩID")) = "" Then Exit Sub
                If Me.txtFindItem.Text = "" Then Exit Sub
                
                '�����Һ���б����Ƿ��в�����Ŀ
                blnFindItem = False
                For intCol = 1 To .Cols - 1
                    If .ColKey(intCol) = lblFindItem.Caption Then
                        blnFindItem = True
                        Exit For
                    End If
                Next
                If blnFindItem = False Then Exit Sub
                
                strFind = txtFindItem.Text
            
                lngRow = .FindRow(strFind, 1, .ColIndex(lblFindItem.Caption))
                 
                
                If lngRow > 0 Then
                    .Row = lngRow
                    .TopRow = lngRow
                    
                    
                    
                    If blnScaner = True And Me.lblFindItem.Caption = "ƿǩ��" And Me.txtFindItem.Text <> "" Then
                        If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = 0 Then
                            For i = 1 To .rows - 1
                                If .TextMatrix(i, .ColIndex("��ҩID")) = .TextMatrix(lngRow, .ColIndex("��ҩID")) Then
                                    .TextMatrix(i, .ColIndex("ѡ��")) = -1
                                End If
                            Next
                            
                            Call UpdateExeSign(Val(.TextMatrix(lngRow, .ColIndex("��ҩID"))), IIf(Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1, 1, 0))
                            
                            DoEvents
                            If InStr(1, mstrLastLabel, .TextMatrix(lngRow, .ColIndex("ƿǩ��"))) = 0 Then
                                mstrLastLabel = IIf(mstrLastLabel = "", "", mstrLastLabel & ",") & .TextMatrix(lngRow, .ColIndex("ƿǩ��"))
                            End If
                        End If
                    End If
                    
                    If mParams.blnTwoCode = True And lblFindItem.Caption = "ƿǩ��" And blnScaner And Me.txtFindItem.Text <> "" Then
                        
                        
                        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Then
                            txtFindItem.Text = ""
                            Call PIVAWork_Dosage(Val(vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("��ҩID"))), "ɨ��")
                            Exit Sub
                        ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
                            txtFindItem.Text = ""
                            
                            
                            Call PIVAWork_Send(Val(vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("��ҩID"))), "ɨ��")
                            Exit Sub
                        End If
                    End If
                Else

                    If mParams.blnTwoCode = True And lblFindItem.Caption = "ƿǩ��" And blnScaner And Me.txtFindItem.Text <> "" Then
                        
                        gstrSQL = "select ����״̬ from ��Һ��ҩ��¼ where ƿǩ��=[1] and ִ��ʱ�� between [2] and [3]"
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯƿǩ��", txtFindItem.Text, CDate(mcondition.strTransStartTime), CDate(mcondition.strTransEndTime))
                        
                        Debug.Print "2"
                        
                        DoEvents
                        If rstemp.EOF Then
                            Me.lblMsg.Caption = "��ƿǩ������"
                        Else
                            If mcondition.strTransStep = M_STR_CALSS_DOSAGE Then
                                If rstemp!����״̬ = 4 Then
                                    Me.lblMsg.Caption = "��ƿǩ��ɨ��"
                                ElseIf rstemp!����״̬ = 1 Then
                                    Me.lblMsg.Caption = "��ƿǩ�ڰ�ҩ����"
                                ElseIf rstemp!����״̬ = 2 Then
                                    Me.lblMsg.Caption = "��ƿǩ����ҩ����"
                                ElseIf rstemp!����״̬ = 5 Then
                                    Me.lblMsg.Caption = "��ƿǩ���ѷ��ͻ���"
                                ElseIf rstemp!����״̬ >= 9 Then
                                    Me.lblMsg.Caption = "����ҽ����ֹͣ������"
                                End If
                            ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
                                If rstemp!����״̬ = 5 Then
                                    Me.lblMsg.Caption = "��ƿǩ��ɨ��"
                                ElseIf rstemp!����״̬ = 1 Then
                                    Me.lblMsg.Caption = "��ƿǩ�ڰ�ҩ����"
                                ElseIf rstemp!����״̬ = 2 Then
                                    Me.lblMsg.Caption = "��ƿǩ����ҩ����"
                                ElseIf rstemp!����״̬ >= 9 Then
                                    Me.lblMsg.Caption = "����ҽ����ֹͣ������"
                                End If
                            End If
                        End If
                    Else
                        MsgBox "û�ҵ�" & Me.lblFindItem.Caption & "Ϊ[" & strFind & "]����Һ�� ��", vbInformation, gstrSysName
                    End If
                End If
            
                txtFindItem.Text = ""
                If tbcDetail.Item(mDetailType.��Һ���б�).Selected Then txtFindItem.SetFocus
                If tbcDetail.Item(mDetailType.��Һ����Ƭ).Selected And lngRow > 0 Then
                    mfrmPIVCard.GetForce Val(.TextMatrix(lngRow, .ColIndex("��ҩID")))
                End If
            End With
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtLog_GotFocus()
    If Me.txtLog.ForeColor = &H80000000 Then
        Me.txtLog.ForeColor = &H80000001
        txtLog.Text = ""
    End If
End Sub

Private Sub txtName_GotFocus()
    Call zlControl.TxtSelAll(txtdept)
End Sub

Private Sub txtTag_GotFocus()
    Call zlControl.TxtSelAll(txtTag)
End Sub

Private Sub txtTag_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsfColSel.RowData(Row)
        If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
'            vsfTrans.ColWidth(lngCol) = vsfTrans.ColData(lngCol)
            vsfTrans.ColHidden(lngCol) = False
        Else
'            vsfList(Val(vsfColSel.Tag)).ColWidth(lngCol) = 0
            vsfTrans.ColHidden(lngCol) = True
        End If
    End If
    
    Call SaveListColState
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



Private Sub InitColSelList(ByVal strUnvisble As String)
    Dim i As Integer
    
    With vsfColSel
        .rows = .FixedRows
        For i = 1 To vsfTrans.Cols - 1
            '���ڲ�������ʾ�б���в��ܼ�����ѡ���б�
            If IsInString(strUnvisble, vsfTrans.ColKey(i), ";") = False Then
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 1) = vsfTrans.ColKey(i)
                .RowData(.rows - 1) = i
'
'                '�п�Ϊ�ջ������ص�������Ϊ����ѡ
'                If Not (vsfTrans.ColWidth(i) = 0 Or vsfTrans.ColHidden(i)) Then
'                    .TextMatrix(.rows - 1, 0) = 0
'                End If
                
                'ָ����������Ϊ������������
                If IsInString(mstrUnallowSetColHide, vsfTrans.ColKey(i), ";") = True Then
                    .Cell(flexcpForeColor, .rows - 1, 1) = .BackColorFixed
                End If
            End If
        Next
    End With
End Sub

Private Sub SaveListColState()
    Dim strType As String
    Dim str������ As String
    Dim i As Integer
    
    With vsfTrans
        For i = 0 To .Cols - 1
            If .ColHidden(i) = True Then
                str������ = IIf(str������ = "", "", str������ & ";") & .ColKey(i)
            End If
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���", mcondition.strTransStep, str������)
End Sub



Private Sub vsfDept_AfterEdit(index As Integer, ByVal Row As Long, ByVal Col As Long)
    With vsfDept(0)
        If Row = 0 Then Exit Sub
        If Col <> .ColIndex("ѡ��") Then Exit Sub
        If .MouseRow <> Row Or .MouseCol <> Col Then Exit Sub
    End With

    Call GetCount

'    DoEvents
'    Call RefreshDetailList(index)
End Sub


Private Sub vsfDept_EnterCell(index As Integer)
    With vsfDept(0)
        .Editable = flexEDNone
        If .Row = 0 Then Exit Sub
        If .Col <> .ColIndex("ѡ��") Then Exit Sub
        
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Function AdviceCheckWarn(ByVal Int���� As Integer, ByVal strNo As String, ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional ByVal lngҽ��id As Long) As Long
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
    Dim strSQL As String, i As Long, k As Long
    Dim lngPatiID As Long
    Dim lng��ҳID As Long
    Dim str�Һŵ� As String
    Dim rsҽ�� As Recordset
    Dim strƵ�� As String
    Dim blnDo As Boolean
    Dim strTmp As String
    
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
    If lngҽ��id = 0 Then
        strSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
            " From ҩƷ�շ���¼ A,סԺ���ü�¼ B,����ҽ����¼ C " & _
            " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
            " And A.����=[2] And A.no=[1] "
        strTmp = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
        strSQL = strSQL & " Union All " & strTmp
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, Int����)
    
        If rsTmp.RecordCount = 0 Then
            rsTmp.Close
            Exit Function
        End If
        
        lngPatiID = rsTmp!����ID
        str�Һŵ� = nvl(rsTmp!�Һŵ�)
        lng��ҳID = rsTmp!��ҳid
    Else
        strSQL = "select A.����id,A.���id,A.��ҳid,A.�շ�ϸĿid,A.����ҽ��,A.��������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ҽ����Ч,A.��ʼִ��ʱ�� ��ʼʱ��,A.ִ����ֹʱ�� ����ʱ��,C.���� �÷�,D.���� ҩƷ����,D.���㵥λ from ����ҽ����¼ A,����ҽ����¼ B,������ĿĿ¼ C,�շ�ϸĿ D where A.���id=B.ҽ��id and B.������Ŀid=C.id and A.�շ�ϸĿid=d.id and A.id=[1]"
        Set rsҽ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��id)
        
        lngPatiID = rsҽ��!����ID
        lng��ҳID = rsҽ��!��ҳid
    End If
    
    '���벡�˾�����Ϣ(PASS��Ҫ�Ļ�������,ͬһ���˿ɲ��ظ�����)
    '-------------------------------------------------------------
    If lngPatiID <> mlngPassPati Then
        If str�Һŵ� <> "" Then               '���ﲡ��
            strSQL = "Select ����ID,Count(Distinct Trunc(�Ǽ�ʱ��)) as ������� From ���˹Һż�¼ Where ��¼����=1 And ��¼״̬=1 And ����ID=[1] Group by ����ID"
            strSQL = "Select D.�������,A.����,A.�Ա�,A.��������," & _
                " C.���� as ������,C.���� as ������,E.��� as ҽ����,E.���� as ҽ����" & _
                " From ������Ϣ A,���˹Һż�¼ B,���ű� C,(" & strSQL & ") D,��Ա�� E" & _
                " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And A.����ID=D.����ID" & _
                " And B.ִ����=E.����(+) And A.����ID=[1] And B.NO=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID, str�Һŵ�)
            If rsTmp.EOF Then
                Screen.MousePointer = 0
                Exit Function
            End If

            Call PassSetPatientInfo(lngPatiID, rsTmp!�������, rsTmp!����, zlStr.nvl(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), zlStr.nvl(rsTmp!ҽ����) & "/" & zlStr.nvl(rsTmp!ҽ����), ""), "")
        Else                                    'סԺ����
            strSQL = _
                " Select A.����,A.�Ա�,A.��������,B.��Ժ����,B.��Ժ����," & _
                " C.���� as ������,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
                " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
                " Where A.����ID=B.����ID And A.��ҳid=B.��ҳid And B.��Ժ����ID=C.ID" & _
                " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID, lng��ҳID)
            If rsTmp.EOF Then
                Screen.MousePointer = 0
                Exit Function
            End If

            Call PassSetPatientInfo(lngPatiID, lng��ҳID, rsTmp!����, zlStr.nvl(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), zlStr.nvl(rsTmp!ҽ����) & "/" & zlStr.nvl(rsTmp!ҽ����), ""), _
                IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy-MM-dd")))
        End If
        mlngPassPati = lngPatiID
    End If

    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        'ȡҩƷ����
         strҩƷ = vsfTrans.TextMatrix(lngRow, vsfTrans.ColIndex("ҩƷ����"))
         lngҩƷid = vsfTrans.TextMatrix(lngRow, vsfTrans.ColIndex("ҩƷID"))
         str������λ = vsfTrans.TextMatrix(lngRow, vsfTrans.ColIndex("������λ"))
         'ȡҩƷ��ҩ;��
         str�÷� = vsfTrans.TextMatrix(lngRow, vsfTrans.ColIndex("�÷�"))

        If InStr(strҩƷ, " ") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, " ") - 1)
        If InStr(strҩƷ, "]") > 0 Then strҩƷ = Mid(strҩƷ, InStr(strҩƷ, "]") + 1, Len(strҩƷ) - InStr(strҩƷ, "]"))
        '�����ѯҩƷ��Ϣ
        Call PassSetQueryDrug(lngҩƷid, strҩƷ, str������λ, str�÷�)

        '���ò˵�����״̬
        Call SetPassMenuState

        AdviceCheckWarn = 1 '��ʾ���Ե����˵�

        Screen.MousePointer = 0: Exit Function
    Else
        With rsҽ��
            '��ҩ��˻���ҩ�о�
            strҩƷ = "": str�÷� = "": strƵ�� = ""
            i = 1
            If !����ҽ�� <> "" Then
                strSQL = "select ��� from ��Ա�� where ����=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "", rsҽ��!����ҽ��)
            End If
            
            blnDo = lngҽ��id <> 0 And !�շ�ϸĿid <> 0
            If blnDo Then
                'ȡҩƷ����
                strҩƷ = !ҩƷ����
                
                'ȡҩƷ��ҩ;��
                str�÷� = !�÷�
                
                'ȡ��ҩƵ��(��/��),��Ϊ������������
                If !�����λ = "��" Then
                    strƵ�� = !Ƶ�ʴ��� & "/" & !Ƶ�ʼ��
                ElseIf !�����λ = "��" Then
                    strƵ�� = !Ƶ�ʴ��� & "/7"
                ElseIf !�����λ = "Сʱ" Then
                    If Val(!Ƶ�ʼ��) <= 24 Then
                        strƵ�� = Format(24 / Val(!Ƶ�ʼ��) * Val(!Ƶ�ʴ���), "0") & "/1"
                    Else
                        strƵ�� = Val(!Ƶ�ʴ���) & "/" & Format(Val(!Ƶ�ʼ��) / 24, "0")
                    End If
                ElseIf !�����λ = "����" Then
                    strƵ�� = Format((24 * 60) / Val(!Ƶ�ʼ��) * Val(!Ƶ�ʴ���), "0") & "/1"
                End If
                
                Call PassSetRecipeInfo(lngҽ��id, !�շ�ϸĿid, strҩƷ, _
                    !��������, !���㵥λ, strƵ��, _
                    Format(!��ʼʱ��, "yyyy-MM-dd"), Format(!����ʱ��, "yyyy-MM-dd"), str�÷�, _
                    !���id, !ҽ����Ч, rsTmp!��� & "\" & !����ҽ��)
            End If
            
            '�޿�����ҩƷ
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) Then
                Screen.MousePointer = 0: Exit Function
            End If
        End With
    End If

    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
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

Private Sub vsfDept_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button = 2 Then
        Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_SortPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub vsfMedis_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long
    
    With Me.vsfMedis
        If Col = .ColIndex("ѡ��") Then
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("���id")) = .TextMatrix(Row, .ColIndex("���id")) And _
                    .TextMatrix(lngRow, .ColIndex("��ҩ��־")) = .TextMatrix(Row, .ColIndex("��ҩ��־")) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = .TextMatrix(Row, .ColIndex("ѡ��"))
                End If
            Next
        End If
    
        Call mfrmPIVCard.ChooseOneRec(.TextMatrix(Row, .ColIndex("���id")), IIf(.TextMatrix(Row, .ColIndex("ѡ��")) = -1, 1, 0))
    
    End With
End Sub

Private Sub vsfMedis_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfMedis.ColIndex("��") And Col <> Me.vsfMedis.ColIndex("ѡ��") Then Cancel = True
    
    If (mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) Then
        If Col = Me.vsfMedis.ColIndex("ѡ��") Then
            If Val(vsfMedis.TextMatrix(Row, vsfMedis.ColIndex("��ҩ��־"))) = 1 Then
                Cancel = True
            End If
        End If
    End If
    
    If ((Val(vsfMedis.TextMatrix(Row, vsfMedis.ColIndex("�����"))) <> 0 And Val(vsfMedis.TextMatrix(Row, vsfMedis.ColIndex("�����"))) <> 3) And mcondition.strTransStep = M_STR_CALSS_AUDIT) Then
        Cancel = True
    ElseIf mcondition.strTransStep = M_STR_CALSS_AUDIT Then
        Cancel = False
    End If
End Sub

Private Sub vsfMedis_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '���ò��ܵ����п����
    With vsfMedis
        If Col = .ColIndex("��") Or _
            Col = .ColIndex("��ǰ��") Or _
            Col = .ColIndex("ѡ��") Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfMedis_DblClick()
    Dim lngRow As Long
    Dim strҽ��ID�� As String
    
    With vsfMedis
        If .Row = 0 Then Exit Sub
        
        'ȡ��ǰһ����ҩ��ҽ��ID��
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("���ID")) = .TextMatrix(.Row, .ColIndex("���ID")) Then
                strҽ��ID�� = strҽ��ID�� & IIf(strҽ��ID�� = "", "", ",") & .TextMatrix(lngRow, .ColIndex("ҽ��id"))
            End If
        Next
        
        If Val(.TextMatrix(.Row, .ColIndex("�����"))) <> 0 And Val(.TextMatrix(.Row, .ColIndex("�����"))) <> 3 And mcondition.strTransStep = M_STR_CALSS_AUDIT Then Exit Sub
        If .Col = .ColIndex("��") Then
            If .TextMatrix(.Row, .ColIndex("��־")) = "0" Then
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("���id")) = .TextMatrix(.Row, .ColIndex("���id")) Then
                        .TextMatrix(lngRow, .ColIndex("��־")) = "1"
                        .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Me.ImgList.ListImages(3).Picture
                        .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                    End If
                Next
                
            ElseIf .TextMatrix(.Row, .ColIndex("��־")) = "1" Then
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("���id")) = .TextMatrix(.Row, .ColIndex("���id")) Then
                        .TextMatrix(lngRow, .ColIndex("��־")) = "2"
                        .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Me.ImgList.ListImages(4).Picture
                        .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                    End If
                Next
                
            ElseIf .TextMatrix(.Row, .ColIndex("��־")) = "2" Then
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("���id")) = .TextMatrix(.Row, .ColIndex("���id")) Then
                        .TextMatrix(lngRow, .ColIndex("��־")) = "0"
                        .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Nothing
                    End If
                Next
            End If
        End If
        
        If .Col = .ColIndex("��") Then
            If IsInString(gstrprivs, "������ҩ���", ";") And Not gobjPass Is Nothing Then
                Call gobjPass.zlPassQueryCheckResult_YF(mlngMode, .TextMatrix(.Row, .ColIndex("סԺ��")), "2", Val(.TextMatrix(.Row, .ColIndex("����ID"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), "", strҽ��ID��)
            End If
        End If
        
        If .Col = .ColIndex("ҩƷ����") Then
            If IsInString(gstrprivs, "������ҩ���", ";") And Not gobjPass Is Nothing Then
                Call gobjPass.zlPassAdviceMainPoint_YF("2", .TextMatrix(.Row, .ColIndex("ҩƷid")), .TextMatrix(.Row, .ColIndex("ҩ��")))
            End If
        End If
        
        If .TextMatrix(.Row, .ColIndex("���ID")) <> "" Then
            Call mfrmPIVCard.ChooseOneRec(.TextMatrix(.Row, .ColIndex("���ID")), .TextMatrix(.Row, .ColIndex("��־")))
        End If
    End With
    
     
End Sub

Private Sub vsfMedis_EnterCell()
    Dim intRow As Integer
    Dim intBegin As Integer
    Dim intEnd As Integer
    Dim strDiag As String
    Dim i As Integer
    
    With Me.vsfMedis
        txtLog.Text = ""
        txtDia.Text = ""
        Me.txtLog.Text = .TextMatrix(.Row, .ColIndex("ҩʦ���ԭ��"))
        If Val(.TextMatrix(.Row, .ColIndex("���ID"))) = 0 Then
            txtLog.Enabled = False
            Me.CmdSave.Enabled = False
            Exit Sub
        End If
        If Me.tabDeptList.Selected.index = 0 Then
            If Me.txtLog.Text = "" Then
                txtLog.Text = "���������ԭ��"
                txtLog.ForeColor = &H80000000
            Else
                txtLog.ForeColor = &H80000001
            End If
            txtLog.Enabled = True
            CmdSave.Enabled = True
        Else
            txtLog.ForeColor = &H80000001
        End If
        If mintBeginRow <> 0 And mintBeginRow <= .rows - 1 Then
            If Val(.TextMatrix(mintBeginRow, .ColIndex("������"))) = 1 Then
                .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &H80000005
            Else
                .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &HC0FFC0
            End If
        End If
        
        For intRow = IIf(.Row > 8, .Row - 8, 1) To IIf(.Row + 8 > .rows - 1, .rows - 1, .Row + 8)
            If Val(.TextMatrix(.Row, .ColIndex("���ID"))) = Val(.TextMatrix(intRow, .ColIndex("���ID"))) And .Row > intRow And intBegin = 0 Then
                intBegin = intRow
            ElseIf .Row < intRow And Val(.TextMatrix(.Row, .ColIndex("���ID"))) = Val(.TextMatrix(intRow, .ColIndex("���ID"))) Then
                intEnd = intRow
            End If
        Next
    
        intRow = 0
        mintBeginRow = IIf(intBegin = 0, .Row, intBegin)
        mintEndRow = IIf(intEnd = 0, .Row, intEnd)
'        .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &HFFE8D0
        
        .Redraw = flexRDNone

        '��ʼ�����
        For intRow = 0 To .rows - 1
            .CellBorderRange intRow, 0, intRow, .Cols - 1, vbBlue, 0, 0, 0, 0, 0, 0
        Next
        
        intBegin = 0
        intEnd = 0
        '����ѡ���еĿ��Χ
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, .ColIndex("���ID")) = .TextMatrix(.Row, .ColIndex("���ID")) Then
                If intBegin = 0 Then intBegin = intRow
                intEnd = intRow
            End If
        Next
        
        '�Ե�ǰѡ�е��и������
        If intBegin = intEnd Then
            '��ֻ��һ��
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 2, 0, 0
        ElseIf intBegin + 1 = intEnd Then
            '��ֻ��2��
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 0, 0, 0     '�ϲ���
            .CellBorderRange intEnd, 0, intEnd, .Cols - 1, vbBlue, 2, 0, 2, 2, 0, 0         '�²���
        Else
            '��3�м�����
            For intRow = intBegin + 1 To intEnd - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, vbBlue, 2, 0, 2, 0, 0, 0     '�м䲿��
            Next
            
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 0, 0, 0     '�ϲ���
            .CellBorderRange intEnd, 0, intEnd, .Cols - 1, vbBlue, 2, 0, 2, 2, 0, 0         '�²���
        End If
        
        'ȥ��ѡ��ʱ�ı���ɫ
        .BackColorSel = .Cell(flexcpBackColor, .Row, 1)
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        
        .Redraw = flexRDDirect
        
        intRow = 0
        
        strDiag = RecipeSendWork_GetDiagnosis(2, Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))))
        
        If InStr(1, strDiag, "��ҽ��Ժ���") >= 1 Then
            strDiag = Mid(strDiag, InStr(1, strDiag, "��ҽ��Ժ���") + 7)
            
            If InStr(1, strDiag, "|") >= 1 Then
                strDiag = Mid(strDiag, 1, InStr(1, strDiag, "|") - 1)
            End If
            
            txtDia.Text = ""
            If strDiag <> "" Then
                strDiag = strDiag & ";"
                For i = 0 To UBound(Split(strDiag, ";"))
                    If Split(strDiag, ";")(i) <> "" Then
                        If InStr(1, txtDia.Text & "��", "��" & Split(strDiag, ";")(i) & "��") < 1 Then
                            txtDia.Text = IIf(txtDia.Text = "", " ��", txtDia.Text & " ��") & Split(strDiag, ";")(i)
                        End If
                    End If
                Next
            End If
        End If
        
        
        If Not gobjPass Is Nothing Then Call gobjPass.zlPassSetDrug_YF(.TextMatrix(.Row, .ColIndex("ҩƷid")), .TextMatrix(.Row, .ColIndex("ҩ��")))
        If Not gobjPass Is Nothing Then Call gobjPass.zlPassClearLight_YF
    End With
End Sub

Private Sub vsfMedis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim lngPatiID As Long
    Dim lng��ҳID As Long
    Dim str����� As String
    Dim lngҽ��id As Long
    
    If Button = 2 Then
        If Me.cbsMain Is Nothing Then Exit Sub
        
        If IsInString(gstrprivs, "������ҩ���", ";") And vsfMedis.MouseCol = vsfMedis.ColIndex("��") And mParams.intShowPass <> 2 And Not gobjPass Is Nothing Then
            '���Pass״̬
            lngҽ��id = Val(vsfMedis.TextMatrix(vsfMedis.MouseRow, vsfMedis.ColIndex("ҽ��id")))
            str����� = vsfMedis.TextMatrix(vsfMedis.MouseRow, vsfMedis.ColIndex("����"))
            lngPatiID = Val(vsfMedis.TextMatrix(vsfMedis.MouseRow, vsfMedis.ColIndex("����id")))
            lng��ҳID = Val(vsfMedis.TextMatrix(vsfMedis.MouseRow, vsfMedis.ColIndex("��ҳid")))

            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_PASS)
            
            objPopup.Visible = True
            
            Call gobjPass.zlPASSPopupCommandBars_YF(mlngMode, objPopup.CommandBar, mconMenu_PASS, lngPatiID, lng��ҳID, "", str�����, lngҽ��id)
            
            objPopup.CommandBar.ShowPopup
        Else
            Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_OperPopup)
            If Not objPopup Is Nothing Then
                For Each cbrControl In objPopup.CommandBar.Controls
                    
                    If cbrControl.Id = conMenu_Oper_Look Then
                        cbrControl.Visible = True
                    Else
                        cbrControl.Visible = False
                    End If
                Next
                    
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub vsfMedis_RowColChange()
    With vsfMedis
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
    End With
End Sub

Private Sub vsfMsg_DblClick()
    Dim i As Integer
    
    With Me.vsfMsg
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("ʱ��")) = "" Then Exit Sub
        
        If DateDiff("D", mdateToday, Format(.TextMatrix(.Row, .ColIndex("ִ��ʱ��")), "yyyy-mm-dd hh:mm:ss")) > 2 Then
            Me.cboʱ�䷶Χ.ListIndex = 3
            
            If Format(.TextMatrix(.Row, .ColIndex("ִ��ʱ��")), "yyyy-mm-dd hh:mm:ss") > Me.Dtp����ʱ��.Value Then
                Me.Dtp����ʱ��.Value = Format(.TextMatrix(.Row, .ColIndex("ִ��ʱ��")), "yyyy-mm-dd hh:mm:ss")
            ElseIf Format(.TextMatrix(.Row, .ColIndex("ִ��ʱ��")), "yyyy-mm-dd hh:mm:ss") < Me.Dtp��ʼʱ��.Value Then
                Me.Dtp��ʼʱ��.Value = Format(.TextMatrix(.Row, .ColIndex("ִ��ʱ��")), "yyyy-mm-dd hh:mm:ss")
            End If
        Else
            Me.cboʱ�䷶Χ.ListIndex = DateDiff("d", mdateToday, Format(.TextMatrix(.Row, .ColIndex("ִ��ʱ��")), "yyyy-mm-dd hh:mm:ss"))
        End If
        
        If (Val(Me.cboʱ�䷶Χ.Tag) = 3 And Me.cboʱ�䷶Χ.ListIndex < 3) Or (Val(Me.cboʱ�䷶Χ.Tag) < 3 And Me.cboʱ�䷶Χ.ListIndex = 3) Then
            Call ResizeConditionArea
        End If
        
        Me.cboʱ�䷶Χ.Tag = Me.cboʱ�䷶Χ.ListIndex
        If .TextMatrix(.Row, .ColIndex("����")) = "ҽ������" Then
            Me.tabDeptList.Item(1).Selected = True
            Me.tbcLook.Item(5).Selected = True
        ElseIf .TextMatrix(.Row, .ColIndex("����")) = "��������" Then
            Me.tabDeptList.Item(0).Selected = True
            Me.tabWork.Item(4).Selected = True
        ElseIf .TextMatrix(.Row, .ColIndex("����")) = "���ε���" Then
            Me.tabDeptList.Item(0).Selected = True
            Me.tabWork.Item(1).Selected = True
        End If
        
        Call RefreshDeptList(Me.tabDeptList.Selected.index)
        
        For i = 1 To Me.vsfDept(0).rows - 1
            If vsfDept(0).TextMatrix(i, vsfDept(0).ColIndex("����id")) = .TextMatrix(.Row, .ColIndex("���˲���id")) Then
                vsfDept(0).TextMatrix(i, vsfDept(0).ColIndex("ѡ��")) = -1
                Call vsfDept_AfterEdit(0, i, .ColIndex("ѡ��"))
            End If
        Next
        
        Call RefreshDetailList(Me.tabDeptList.Selected.index)

        .RemoveItem (.Row)
        lblMsgComment.Caption = "��Ϣ����(" & vsfMsg.rows - 1 & ")"
        
    End With
    
End Sub

Private Sub vsfSumDrug_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfSumDrug
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsfSumDrug_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfSumDrug
        If Col = .ColIndex("���") Then
            .Col = .ColIndex("�Ƿ���")
            .Sort = Order
        End If
    End With
End Sub

Private Sub vsfSumDrug_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '���ò��ܵ����п����
    With vsfSumDrug
        If Col = .ColIndex("���") Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfTrans_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInput As String
    Dim blnNext As Boolean
    Dim strLabel As String
    Dim int��� As Integer
    Dim i As Integer
    
    With vsfTrans
        If Row = 0 Then Exit Sub
        
        If Col = .ColIndex("ѡ��") Then
            Call UpdateExeSign(Val(.TextMatrix(Row, .ColIndex("��ҩID"))), IIf(.TextMatrix(Row, .ColIndex("ѡ��")) = -1, 1, 0))
            
            DoEvents
            
            strLabel = .TextMatrix(Row, .ColIndex("ƿǩ��"))
            
            If Val(.TextMatrix(Row, .ColIndex("ѡ��"))) = -1 Then
                If InStr(1, mstrLastLabel, strLabel) = 0 Then
                    mstrLastLabel = IIf(mstrLastLabel = "", "", mstrLastLabel & ",") & strLabel
                End If
            Else
                mstrLastLabel = Replace(mstrLastLabel, strLabel & ",", "")
                mstrLastLabel = Replace(mstrLastLabel, strLabel, "")
            End If
            
            Call mfrmPIVCard.ChooseOneRec(.TextMatrix(Row, .ColIndex("��ҩID")), IIf(.TextMatrix(Row, .ColIndex("ѡ��")) = -1, 1, 0))
        End If
        
        If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            If Col = .ColIndex("��ҩ����") Then
                .ColComboList(.ColIndex("��ҩ����")) = ""
                If mrsTrans Is Nothing Then Exit Sub
                
                If .TextMatrix(Row, .ColIndex("��ҩ����")) = .TextMatrix(Row, .ColIndex("ԭ����")) Then
                    mfrmPIVCard.Changebatch Val(.TextMatrix(Row, .ColIndex("��ҩID"))), .TextMatrix(Row, .ColIndex("��ҩ����"))
                    Exit Sub
                End If
                
                If MsgBox("�Ƿ�ȷ�ϰ�������[" & .TextMatrix(Row, .ColIndex("ԭ����")) & "]����Ϊ[" & .TextMatrix(Row, .ColIndex("��ҩ����")) & "]��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    .TextMatrix(Row, .ColIndex("��ҩ����")) = .TextMatrix(Row, .ColIndex("ԭ����"))
                    
                    For i = Row - 1 To 0 Step -1
                        If .TextMatrix(Row, .ColIndex("��ҩID")) = .TextMatrix(i, .ColIndex("��ҩID")) Then
                            .TextMatrix(i, .ColIndex("��ҩ����")) = .TextMatrix(Row, .ColIndex("ԭ����"))
                        End If
                    Next
                    
                    For i = Row + 1 To .rows - 1
                        If .TextMatrix(Row, .ColIndex("��ҩID")) = .TextMatrix(i, .ColIndex("��ҩID")) Then
                            .TextMatrix(i, .ColIndex("��ҩ����")) = .TextMatrix(Row, .ColIndex("ԭ����"))
                        End If
                    Next
                    Exit Sub
                End If
                
                .TextMatrix(Row, .ColIndex("ԭ����")) = .TextMatrix(Row, .ColIndex("��ҩ����"))
                mfrmPIVCard.Changebatch Val(.TextMatrix(Row, .ColIndex("��ҩID"))), .TextMatrix(Row, .ColIndex("��ҩ����"))
                
                
                int��� = Mid(mstr���, InStr(mstr���, "," & .TextMatrix(Row, .ColIndex("��ҩ����")) & ",") + Len("," & .TextMatrix(Row, .ColIndex("��ҩ����")) & ","), 1)
                If int��� = 1 Then int��� = 2
                .TextMatrix(Row, .ColIndex("�Ƿ���")) = int���
                .Cell(flexcpPicture, Row, .ColIndex("���"), Row, .ColIndex("���")) = IIf(int��� = 2, picPacker(2).Picture, Nothing)
                .Cell(flexcpPictureAlignment, Row, .ColIndex("���"), Row, .ColIndex("���")) = flexPicAlignCenterCenter
                .Cell(flexcpForeColor, Row, .ColIndex("��ҩ����")) = vbBlue
                
                mrsTrans.Filter = "��ҩID=" & Val(.TextMatrix(Row, .ColIndex("��ҩID")))
                Do While Not mrsTrans.EOF
                    mrsTrans!��ҩ���� = .TextMatrix(Row, .ColIndex("��ҩ����"))
                    mrsTrans!�Ƿ��� = IIf(int��� <> 0, 2, 0)
                    mrsTrans.Update
                    mrsTrans.MoveNext
                Loop
                
                DoEvents
                
                strInput = .TextMatrix(Row, .ColIndex("��ҩID")) & "," & Left(.TextMatrix(Row, .ColIndex("��ҩ����")), 1) & ":"
                
                On Error GoTo errHandle
                
                If strInput <> "" Then
                    gstrSQL = "Zl_��Һ��ҩ��¼_����("
                    '��ҩID,����
                    gstrSQL = gstrSQL & "'" & strInput & "'"
                    '�Ƿ��������
                    gstrSQL = gstrSQL & ",1"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
                    
                    gstrSQL = "Zl_��Һ��ҩ��¼_���("
                    '��ҩID,���
                    gstrSQL = gstrSQL & "'" & .TextMatrix(Row, .ColIndex("��ҩID")) & "," & int��� & "'"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
                End If
            End If
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfTrans_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfTrans
        If Col = .ColIndex("���") Then
            .Col = .ColIndex("�Ƿ���")
            .Sort = Order
        End If
    End With
End Sub

Private Sub vsfTrans_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfTrans
        If Row = 0 Then Exit Sub
        If Val(.TextMatrix(Row, .ColIndex("��ҩID"))) = 0 Then Exit Sub
        
        If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            If Col = .ColIndex("��ҩ����") Then
                If mParams.bln�������� = False Then Exit Sub
                .ColComboList(.ColIndex("��ҩ����")) = mParams.strBatchList
            End If
        End If
    End With
End Sub

Private Sub vsfTrans_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    '���ò����ƶ�����
    With vsfTrans
        If Col = .ColIndex("��") Then
            Position = .ColIndex("��")
        End If
        
        If Col = .ColIndex("ѡ��") Then
            Position = .ColIndex("ѡ��")
        End If
        
        If Col = .ColIndex("��ӡ") Then
            Position = .ColIndex("��ӡ")
        End If
        
        If Col = .ColIndex("���") Then
            Position = .ColIndex("���")
        End If
        
        If Col = .ColIndex("��ҩ����") Then
            Position = .ColIndex("��ҩ����")
        End If
        
        If (Col <> .ColIndex("��") And Position = .ColIndex("��")) Or _
            (Col <> .ColIndex("ѡ��") And Position = .ColIndex("ѡ��")) Or _
            (Col <> .ColIndex("��ӡ") And Position = .ColIndex("��ӡ")) Or _
            (Col <> .ColIndex("���") And Position = .ColIndex("���")) Or _
            (Col <> .ColIndex("��ҩ����") And Position = .ColIndex("��ҩ����")) Then
            Position = Col
        End If
    End With
End Sub

Private Sub vsfTrans_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '���ò��ܵ����п����
    With vsfTrans
        If Col = .ColIndex("��") Or _
            Col = .ColIndex("��ǰ��") Or _
            Col = .ColIndex("ѡ��") Or Col = .ColIndex("���") Or _
            Col = .ColIndex("��ҩ����") Or Col = .ColIndex("��ӡ") Then
            Cancel = True
        End If
    End With
End Sub
Private Sub vsfTrans_DblClick()
    Dim strInput As String
    Dim intFirst As Integer
    Dim lngRow As Long
    Dim strҽ��ID�� As String
    
    With vsfTrans
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("��ҩID")) = "" Then Exit Sub
        If mrsTrans Is Nothing Then Exit Sub
        
        'ȡ��ǰһ����ҩ��ҽ��ID��
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("��ҩID")) = .TextMatrix(.Row, .ColIndex("��ҩID")) Then
                strҽ��ID�� = strҽ��ID�� & IIf(strҽ��ID�� = "", "", ",") & .TextMatrix(lngRow, .ColIndex("��Ӧҽ��ID"))
            End If
        Next
        
        Select Case .Col
            Case .ColIndex("���")
                If mcondition.strTransStep <> M_STR_CALSS_PREPARE And mcondition.strTransStep <> M_STR_CALSS_DOSAGE Then Exit Sub
                If mParams.bln������� = False Then Exit Sub
                
                If MsgBox("�Ƿ����Ϊ" & IIf(Val(.TextMatrix(.Row, .ColIndex("�Ƿ���"))) = 0, """���""", """�����""") & "״̬��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
                If Val(.TextMatrix(.Row, .ColIndex("�Ƿ���"))) = 0 Then
                    .TextMatrix(.Row, .ColIndex("�Ƿ���")) = 2
                Else
                    .TextMatrix(.Row, .ColIndex("�Ƿ���")) = 0
                End If
                
                '������Һ(���)ͼ��
                .Col = .ColIndex("���")
                .CellPicture = IIf(.TextMatrix(.Row, .ColIndex("�Ƿ���")) = 2, picPacker(2).Picture, Nothing)
                .CellPictureAlignment = flexPicAlignCenterCenter
                
                mrsTrans.Filter = "��ҩID=" & Val(.TextMatrix(.Row, .ColIndex("��ҩID")))
                Do While Not mrsTrans.EOF
                    intFirst = intFirst + 1
                    mrsTrans!�Ƿ��� = Val(.TextMatrix(.Row, .ColIndex("�Ƿ���")))
                    
                    If mcondition.strTransStep = M_STR_CALSS_DOSAGE And intFirst = 1 And .TextMatrix(.Row, .ColIndex("�Ƿ���")) > 0 Then
                        mintCountPack = mintCountPack + IIf(IIf(IsNull(mrsTrans!��ҩʱ��), "", Format(mrsTrans!��ҩʱ��, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!���ʱ��), "", Format(mrsTrans!���ʱ��, "YYYY-MM-DD HH:MM:SS")), 0, 1)
                    Else
                        If IIf(IsNull(mrsTrans!��ҩʱ��), "", Format(mrsTrans!��ҩʱ��, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!���ʱ��), "", Format(mrsTrans!���ʱ��, "YYYY-MM-DD HH:MM:SS")) Then
                            mintCountPack = mintCountPack - 1
                        End If
                    End If
                    
                    mrsTrans!���ʱ�� = IIf(.TextMatrix(.Row, .ColIndex("�Ƿ���")) = 0, "", Sys.Currentdate)
                    mrsTrans.Update
                    mrsTrans.MoveNext
                Loop
                
                Call GetCount
                
                mfrmPIVCard.PackCard Val(.TextMatrix(.Row, .ColIndex("��ҩID"))), .TextMatrix(.Row, .ColIndex("�Ƿ���"))
                
                DoEvents
                
                strInput = .TextMatrix(.Row, .ColIndex("��ҩID")) & "," & .TextMatrix(.Row, .ColIndex("�Ƿ���"))
                
                On Error GoTo errHandle
                
                If strInput <> "" Then
                    gstrSQL = "Zl_��Һ��ҩ��¼_���("
                    '��ҩID,���
                    gstrSQL = gstrSQL & "'" & strInput & "'"
                    '�ֹ��������
                    gstrSQL = gstrSQL & ",1"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
                End If
            Case .ColIndex("��")
                .TextMatrix(.Row, .ColIndex("�Ƿ�����")) = IIf(.TextMatrix(.Row, .ColIndex("�Ƿ�����")) = "1", 0, 1)
                .Cell(flexcpPicture, .Row, .ColIndex("��"), .Row, .ColIndex("��")) = IIf(.TextMatrix(.Row, .ColIndex("�Ƿ�����")) = "1", Me.ImgList.ListImages(5).Picture, Me.ImgList.ListImages(6).Picture)
                .Cell(flexcpPictureAlignment, .Row, .ColIndex("��"), .Row, .ColIndex("��")) = flexPicAlignCenterCenter
                    
                 Call SetLock(.TextMatrix(.Row, .ColIndex("�Ƿ�����")), .TextMatrix(.Row, .ColIndex("��ҩid")), True)
            Case .ColIndex("�����")
                If IsInString(gstrprivs, "������ҩ���", ";") And Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassQueryCheckResult_YF(mlngMode, .TextMatrix(.Row, .ColIndex("סԺ��")), "2", Val(.TextMatrix(.Row, .ColIndex("����ID"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), "", strҽ��ID��)
                End If
            Case .ColIndex("ҩƷ����")
                If IsInString(gstrprivs, "������ҩ���", ";") And Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassAdviceMainPoint_YF("2", .TextMatrix(.Row, .ColIndex("ҩƷid")), Mid(.TextMatrix(.Row, .ColIndex("ҩƷ����")), InStr(.TextMatrix(.Row, .ColIndex("ҩƷ����")), "]") + 1))
                End If
            Case .ColIndex("��")
                If .TextMatrix(.Row, .ColIndex("��־")) = "0" Then
                    For lngRow = 1 To .rows - 1
                        If .TextMatrix(lngRow, .ColIndex("��ҩID")) = .TextMatrix(.Row, .ColIndex("��ҩID")) Then
                            .TextMatrix(lngRow, .ColIndex("��־")) = "1"
                            .TextMatrix(lngRow, .ColIndex("ѡ��")) = 1
                            .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Me.ImgList.ListImages(3).Picture
                            .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                        End If
                    Next
                ElseIf .TextMatrix(.Row, .ColIndex("��־")) = "1" Then
                    For lngRow = 1 To .rows - 1
                        If .TextMatrix(lngRow, .ColIndex("��ҩID")) = .TextMatrix(.Row, .ColIndex("��ҩID")) Then
                            If mPrives.bln���ʾܾ� Then
                                .TextMatrix(lngRow, .ColIndex("��־")) = "2"
                                .TextMatrix(lngRow, .ColIndex("ѡ��")) = 1
                                .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Me.ImgList.ListImages(4).Picture
                                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = flexPicAlignCenterCenter
                            Else
                                .TextMatrix(lngRow, .ColIndex("��־")) = "0"
                                .TextMatrix(lngRow, .ColIndex("ѡ��")) = 0
                                .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Nothing
                            End If
                            Call UpdateExeSign(lngRow, .TextMatrix(lngRow, .ColIndex("��־")))
                        End If
                    Next
                ElseIf .TextMatrix(.Row, .ColIndex("��־")) = "2" Then
                    For lngRow = 1 To .rows - 1
                        If .TextMatrix(lngRow, .ColIndex("��ҩID")) = .TextMatrix(.Row, .ColIndex("��ҩID")) Then
                            .TextMatrix(lngRow, .ColIndex("��־")) = "0"
                            .TextMatrix(lngRow, .ColIndex("ѡ��")) = 0
                            .Cell(flexcpPicture, lngRow, .ColIndex("��"), lngRow, .ColIndex("��")) = Nothing
                            
                            Call UpdateExeSign(lngRow, .TextMatrix(lngRow, .ColIndex("��־")))
                        End If
                    Next
                End If
                
                '�������ݼ�ִ�б�־
                Call UpdateExeSign(.TextMatrix(.Row, .ColIndex("��ҩID")), .TextMatrix(.Row, .ColIndex("��־")))
                
                If .TextMatrix(.Row, .ColIndex("��ҩID")) <> "" Then
                    Call mfrmPIVCard.ChooseOneRec(.TextMatrix(.Row, .ColIndex("��ҩID")), .TextMatrix(.Row, .ColIndex("��־")))
                End If
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfTrans_EnterCell()
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim rstemp As Recordset
    Dim intRow As Integer
    Dim lng��ҩid As Long
    Dim intBegin As Integer
    Dim intEnd As Integer
    
    With vsfTrans
        If .Row = 0 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("��ҩID"))) = 0 Then Exit Sub
        
        If mintBeginRow <> 0 And mintBeginRow <= .rows - 1 Then
            If mintEndRow > .rows - 1 Then mintEndRow = .rows - 1
            If Val(.TextMatrix(mintBeginRow, .ColIndex("������"))) = 1 Then
                .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &H80000005
            Else
                .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &HC0FFC0
            End If
        End If
        
        For intRow = IIf(.Row > 8, .Row - 8, 1) To IIf(.Row + 8 > .rows - 1, .rows - 1, .Row + 8)
            If Val(.TextMatrix(.Row, .ColIndex("��ҩID"))) = Val(.TextMatrix(intRow, .ColIndex("��ҩID"))) And .Row > intRow And intBegin = 0 Then
                intBegin = intRow
            ElseIf .Row < intRow And Val(.TextMatrix(.Row, .ColIndex("��ҩID"))) = Val(.TextMatrix(intRow, .ColIndex("��ҩID"))) Then
                intEnd = intRow
            End If
        Next
        
        intRow = 0
        mintBeginRow = IIf(intBegin = 0, .Row, intBegin)
        mintEndRow = IIf(intEnd = 0, .Row, intEnd)
'        .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &HFFE8D0
        
        .Redraw = flexRDNone

        '��ʼ�����
        For intRow = 0 To .rows - 1
            .CellBorderRange intRow, 0, intRow, .Cols - 1, vbBlue, 0, 0, 0, 0, 0, 0
        Next
        
        intBegin = 0
        intEnd = 0
        '����ѡ���еĿ��Χ
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, .ColIndex("��ҩID")) = .TextMatrix(.Row, .ColIndex("��ҩID")) Then
                If intBegin = 0 Then intBegin = intRow
                intEnd = intRow
            End If
        Next
        
        '�Ե�ǰѡ�е��и������
        If intBegin = intEnd Then
            '��ֻ��һ��
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 2, 0, 0
        ElseIf intBegin + 1 = intEnd Then
            '��ֻ��2��
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 0, 0, 0     '�ϲ���
            .CellBorderRange intEnd, 0, intEnd, .Cols - 1, vbBlue, 2, 0, 2, 2, 0, 0         '�²���
        Else
            '��3�м�����
            For intRow = intBegin + 1 To intEnd - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, vbBlue, 2, 0, 2, 0, 0, 0     '�м䲿��
            Next
            
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 0, 0, 0     '�ϲ���
            .CellBorderRange intEnd, 0, intEnd, .Cols - 1, vbBlue, 2, 0, 2, 2, 0, 0         '�²���
        End If
        
        'ȥ��ѡ��ʱ�ı���ɫ
        .BackColorSel = .Cell(flexcpBackColor, .Row, 1)
        
        .Redraw = flexRDDirect
        .Editable = flexEDNone
        
        intRow = 0
        
        Select Case .Col
            Case .ColIndex("ѡ��")
                .Editable = flexEDKbdMouse
            Case .ColIndex("��ҩ����")
                If mcondition.strTransStep = M_STR_CALSS_PREPARE And mParams.bln�������� = True Then
                    .Editable = flexEDKbdMouse
                End If
        End Select
        
        If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            '��ȡ����
            Set rstemp = PIVA_�Ѱ�ҩ��Һ��(mcondition.lngCenterID, CDate(.TextMatrix(.Row, .ColIndex("ִ��ʱ��"))), _
                Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))))
            
            If chkSure(1).Value = 1 Then
                rstemp.Filter = "����״̬<>1"
            End If
            
            rstemp.Sort = "��ҩid"
            With Me.VSFLook
                .rows = 1
                .rows = rstemp.RecordCount + 1
                .RowHeight(0) = 250
                .MergeCells = flexMergeFree
                Do While Not rstemp.EOF
                    intRow = intRow + 1
                    
                    
                    If lng��ҩid <> rstemp!��ҩid Then
                        lng��ҩid = rstemp!��ҩid
                        If lng��ҩid <> 0 Then
                            .rows = .rows + 1
                            .Cell(flexcpText, intRow, 0, intRow, .Cols - 1) = 0
                            .RowHidden(intRow) = True
                            intRow = intRow + 1
                        End If
                    Else
                        .MergeCol(.ColIndex("����")) = True
                        .MergeCol(.ColIndex("��ҩ����")) = True
                        .MergeCol(.ColIndex("��ҩ��")) = True
                        .MergeCol(.ColIndex("��ҩʱ��")) = True
                        .MergeCol(.ColIndex("ƿǩ��")) = True
                        .MergeCol(.ColIndex("ҽ������ʱ��")) = True
                        .MergeCol(.ColIndex("ִ��ʱ��")) = True
                        .MergeCol(.ColIndex("���")) = True
                        .MergeCol(.ColIndex("����״̬")) = True
                        .MergeCol(.ColIndex("NO")) = True
                    End If
                    
                    .RowHeight(intRow) = 250
                    .TextMatrix(intRow, .ColIndex("����")) = IIf(zlStr.nvl(rstemp!��ҩ����) = "", "", zlStr.nvl(rstemp!��ҩ����) & "#")
                    .TextMatrix(intRow, .ColIndex("ҩƷ����")) = rstemp!ͨ����
                    .TextMatrix(intRow, .ColIndex("���")) = rstemp!���
                    .TextMatrix(intRow, .ColIndex("����")) = FormatEx(rstemp!����, 2) & rstemp!������λ
                    .TextMatrix(intRow, .ColIndex("����")) = FormatEx(rstemp!����, 2) & rstemp!��λ
                    .TextMatrix(intRow, .ColIndex("ִ��ʱ��")) = rstemp!ִ��ʱ��
                    .TextMatrix(intRow, .ColIndex("ƿǩ��")) = rstemp!ƿǩ��
                    .TextMatrix(intRow, .ColIndex("��ҩid")) = rstemp!��ҩid
                    .TextMatrix(intRow, .ColIndex("��ҩ��")) = rstemp!������Ա
                    .TextMatrix(intRow, .ColIndex("��ҩʱ��")) = rstemp!����ʱ��
                    .TextMatrix(intRow, .ColIndex("��ҩ����")) = zlStr.nvl(rstemp!��ҩ����, " ")
                    .TextMatrix(intRow, .ColIndex("ҽ������ʱ��")) = rstemp!ҽ������ʱ��
                    .TextMatrix(intRow, .ColIndex("����״̬")) = IIf(rstemp!����״̬ = 1, "��ȷ��", IIf(rstemp!����״̬ = 2, "�Ѱ�ҩ", IIf(rstemp!����״̬ = 4, "����ҩ", "�ѷ���")))
                    .TextMatrix(intRow, .ColIndex("���")) = " "
                    .TextMatrix(intRow, .ColIndex("NO")) = rstemp!NO
                    
'                    .CellPicture = picPacker(Val(.TextMatrix(intRow, .ColIndex("�Ƿ���")))).Picture
                    .Cell(flexcpPicture, intRow, .ColIndex("���"), intRow, .ColIndex("���")) = IIf(Val(rstemp!�Ƿ���) = 0, Nothing, picPacker(Val(rstemp!�Ƿ���)).Picture)
                    .Cell(flexcpPictureAlignment, intRow, .ColIndex("���"), intRow, .ColIndex("���")) = flexPicAlignCenterCenter
    
                    rstemp.MoveNext
                Loop
            End With
        End If
    End With
    
End Sub

Private Sub vsfTrans_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim Int���� As Integer
    Dim strNo As String
    
    If Button = 2 Then
        If Me.cbsMain Is Nothing Then Exit Sub
        
        If mParams.intShowPass = 1 And IsInString(gstrprivs, "������ҩ���", ";") And vsfTrans.MouseCol = vsfTrans.ColIndex("�����") Then
            'PASSϵͳ�����˵�
'            If vsfTrans.TextMatrix(vsfTrans.MouseRow, vsfTrans.ColIndex("NO")) = "" Then Exit Sub
'            Int���� = Val(vsfTrans.TextMatrix(vsfTrans.MouseRow, vsfTrans.ColIndex("����")))
'            strNo = vsfTrans.TextMatrix(vsfTrans.MouseRow, vsfTrans.ColIndex("NO"))
'
'            '���Pass״̬
'            If AdviceCheckWarn(Int����, strNo, 0, vsfTrans.MouseRow) >= 0 Then
'                Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_PASS)
'                If Not objPopup Is Nothing Then
'                    objPopup.CommandBar.ShowPopup
'                End If
'
'                Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_PASS_Item + 10, , True)
'                If Not objPopup Is Nothing Then objPopup.Visible = False
'            End If
        Else
            '�Ҽ������˵�
            With vsfTrans
                If .Row = 0 Or .Col > .ColIndex("סԺ��") Then Exit Sub
                If Val(.TextMatrix(.Row, .ColIndex("��ҩID"))) = 0 Then Exit Sub
            End With
            
            Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_OperPopup)
            If Not objPopup Is Nothing Then
                For Each cbrControl In objPopup.CommandBar.Controls
                    cbrControl.Visible = True
                    If mcondition.strTransStep <> M_STR_CALSS_PREPARE Then
                        If cbrControl.Id = conMenu_Oper_DelBatch Then
                            cbrControl.Visible = False
                        End If
                    End If
                    
                    If cbrControl.Id = conMenu_Oper_PrintLabel Then
                        cbrControl.Visible = mParams.blnƿǩ�ֹ���ӡ
                    ElseIf cbrControl.Id = conMenu_Oper_Bag Then
                        cbrControl.Visible = mParams.bln�������
                    End If
                    
                    If cbrControl.Id = conMenu_Oper_Look Then
                        cbrControl.Visible = False
                    End If
                Next
                    
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub vsfTrans_RowColChange()
    '�ƶ���һ���ı�ǵ���ǰ�У�
    With vsfTrans
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
        
    End With
    
    
End Sub

Public Sub ChooseType(ByVal intIndex As Integer, ByVal intValue As Integer)
    Me.chkType(intIndex).Value = intValue
    chkType_Click (intIndex)
End Sub

Public Sub chkAllClick(ByVal intValue As Integer)
    mint��־ = 1
    Me.chkAll.Value = intValue
    Chk_all
End Sub

Public Sub Get��ҩid(ByVal str��ҩid As String)
    mstr��ҩid = str��ҩid
End Sub

Public Sub CheckOne(ByVal str��ҩid As String, ByVal intValue As Integer)
    Dim i As Integer
    
    If mcondition.strTransStep = M_STR_CALSS_AUDIT Then
        With vsfMedis
            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("���id")) = str��ҩid Then
                    .TextMatrix(i, .ColIndex("��־")) = intValue
                    If intValue = 1 Then
                        .Cell(flexcpPicture, i, .ColIndex("��"), i, .ColIndex("��")) = Me.ImgList.ListImages(3).Picture
                        .Cell(flexcpPictureAlignment, i, .ColIndex("��"), i, .ColIndex("��")) = flexPicAlignCenterCenter
                    ElseIf intValue = 2 Then
                        .Cell(flexcpPicture, i, .ColIndex("��"), i, .ColIndex("��")) = Me.ImgList.ListImages(4).Picture
                        .Cell(flexcpPictureAlignment, i, .ColIndex("��"), i, .ColIndex("��")) = flexPicAlignCenterCenter
                    Else
                        .Cell(flexcpPicture, i, .ColIndex("��"), i, .ColIndex("��")) = Nothing
                    End If
                End If
            Next
        End With
    ElseIf mcondition.strTransStep > M_STR_CALSS_AUDIT And mcondition.strTransStep < M_STR_CALSS_PREPARE Then
        With vsfMedis
            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("���id")) = str��ҩid Then
                    .TextMatrix(i, .ColIndex("ѡ��")) = intValue
                End If
            Next
        End With
    Else
        With Me.vsfTrans
            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("��ҩid")) = str��ҩid Then
                    .TextMatrix(i, .ColIndex("ѡ��")) = IIf(intValue <> 0, -1, 0)
                End If
            Next
        End With
        UpdateExeSign -1, Me.tabDeptList.Selected.index
    End If
    
    
    
    
End Sub

Public Sub PackMain(ByVal str��ҩid As String, ByVal intValue As Integer)
    Dim i As Integer
    
    With Me.vsfTrans
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("��ҩID")) = str��ҩid Then
                If intValue = 1 Then
                    .TextMatrix(i, .ColIndex("�Ƿ���")) = 2
                Else
                    .TextMatrix(i, .ColIndex("�Ƿ���")) = 0
                End If
                
                '������Һ(���)ͼ��
                .Cell(flexcpPicture, i, .ColIndex("���"), i, .ColIndex("���")) = IIf(intValue = 2, picPacker(2).Picture, Nothing)
                .Cell(flexcpPictureAlignment, i, .ColIndex("���"), i, .ColIndex("���")) = flexPicAlignCenterCenter
                Exit For
            End If
        Next
    End With
End Sub

Public Sub ChangeBatchMain(ByVal lng��ҩid As Long, ByVal str���� As String)
    Dim i As Integer
    
    With Me.vsfTrans
        For i = 1 To .rows - 1
            If Val(.TextMatrix(i, .ColIndex("��ҩID"))) = lng��ҩid Then
                .TextMatrix(i, .ColIndex("��ҩ����")) = str����
                mrsTrans.Filter = "��ҩID=" & Val(.TextMatrix(i, .ColIndex("��ҩID")))
                Do While Not mrsTrans.EOF
                    mrsTrans!��ҩ���� = .TextMatrix(i, .ColIndex("��ҩ����"))
                    mrsTrans.Update
                    mrsTrans.MoveNext
                Loop
                Exit For
            End If
        Next
    End With
End Sub

Public Sub SetTxtFind(ByVal strText As String, ByVal IntKeyAscii As Integer)
    Me.txtFindItem.Text = strText
    txtFinditem_KeyPress IntKeyAscii
End Sub
Private Function InitMedi(ByVal intType As Integer, ByVal strIDS As String, ByVal str��ҩ���� As String) As Recordset
    Dim rstemp As Recordset
    Dim lngҽ���� As Long
    Dim lngҽ��id As Long
    Dim bln���� As Boolean
    Dim i As Integer
    Dim arrExecute As Variant
    Dim rsMedi As Recordset
    
    '��Һ����¼��
    Set mrsMedi = New ADODB.Recordset
    With mrsMedi
        If .State = 1 Then .Close
        
        '�ü�¼��Ӧ����Һ��ҩ��¼��Ϣ
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "id", adDouble, 18, adFldIsNullable
        .Fields.Append "���id", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҳid", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "����ҽ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "���˲���id", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "���˿���id", adLongVarChar, 18, adFldIsNullable
        
        '��Һ��ҩ��¼��Ӧ��ҩƷ��Ϣ
        .Fields.Append "ҩƷ����", adLongVarChar, 50, adFldIsNullable   '����+ͨ����/��Ʒ��
        .Fields.Append "���", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "ִ��ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�Ƿ�Ƥ��", adDouble, 1, adFldIsNullable
        .Fields.Append "�����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��˱�־", adDouble, 1, adFldIsNullable
        .Fields.Append "ҩʦ���ԭ��", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "ҩƷid", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҽ����Ч", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ҩ;��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ִ������", adDouble, 1, adFldIsNullable
        .Fields.Append "ִ�б��", adDouble, 1, adFldIsNullable
        
        'ҽ����Ϣ
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ƥ�Խ��", adLongVarChar, 20, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    arrExecute = GetArrayByStr(strIDS, 3950, ",")
    For i = 0 To UBound(arrExecute)
        Set rsMedi = Piva_GetMedi(intType, CStr(arrExecute(i)), mParams.intCheck)
        With rsMedi
            Do While Not .EOF
                mrsMedi.AddNew
                If lngҽ��id <> !���id Then
                    lngҽ��id = !���id
                    lngҽ���� = lngҽ���� + 1
                    bln���� = False
                End If
                
                If !��ҩ���� = str��ҩ���� Then bln���� = True
                
                mrsMedi!��� = lngҽ����
                mrsMedi!Id = !Id
                mrsMedi!���id = !���id
                mrsMedi!סԺ�� = !סԺ��
                mrsMedi!���� = !����
                mrsMedi!�Ա� = !�Ա�
                mrsMedi!���� = !����
                mrsMedi!���� = !����
                mrsMedi!����ID = !����ID
                mrsMedi!��ҳid = !��ҳid
                mrsMedi!���� = !��������
                mrsMedi!����ҽ�� = !����ҽ��
                mrsMedi!ҩƷ���� = !ҩƷ����
                mrsMedi!��� = zlStr.nvl(!���)
                mrsMedi!���� = nvl(!��������, 0)
                mrsMedi!��λ = !���㵥λ
                mrsMedi!Ƶ�� = !ִ��Ƶ��
                mrsMedi!ҩ��ID = !ҩ��ID
                mrsMedi!ҩƷID = !ҩƷID
                mrsMedi!���˲���ID = !���˲���ID
                mrsMedi!���˿���id = !���˿���id
                mrsMedi!ִ��ʱ�� = !ִ��ʱ�䷽��
                mrsMedi!�Ƿ�Ƥ�� = !�Ƿ�Ƥ��
                mrsMedi!����ʱ�� = !����ʱ��
                mrsMedi!Ƥ�Խ�� = !Ƥ�Խ��
                mrsMedi!����� = zlStr.nvl(!�����, 0)
                mrsMedi!��˱�־ = !��˱�־
                mrsMedi!ҩʦ���ԭ�� = !ҩʦ���ԭ��
                mrsMedi!��ҩ���� = !��ҩ����
                mrsMedi!ҽ����Ч = !ҽ����Ч
                mrsMedi!��ҩ;�� = !��ҩ;��
                mrsMedi!ִ������ = nvl(!ִ������, 0)
                mrsMedi!ִ�б�� = nvl(!ִ�б��, 0)
                mrsMedi.Update
                .MoveNext
                
                If .EOF Then
                    If bln���� = False And lngҽ���� <> 0 And str��ҩ���� <> "" Then
                        mrsMedi.Filter = "���=" & lngҽ����
                        Do While Not mrsMedi.EOF
                            mrsMedi.Delete adAffectCurrent
                            mrsMedi.MoveNext
                        Loop
                    End If
                Else
                
                    If lngҽ��id <> !���id Then
                        If bln���� = False And lngҽ���� <> 0 And str��ҩ���� <> "" Then
                            mrsMedi.Filter = "���=" & lngҽ����
                            Do While Not mrsMedi.EOF
                                mrsMedi.Delete adAffectCurrent
                                mrsMedi.MoveNext
                            Loop
                        End If
                    End If
                End If
            Loop
        End With
    Next
End Function

Private Sub LoadVsfMedi(ByVal strIDS As String, Optional ByVal blnFilter As Boolean)
    Dim rstemp As Recordset
    Dim i As Long
    Dim lngҽ��id As Long
    Dim j As Integer
    Dim lng����id As Long
    Dim intBlackColor As Integer
    Dim intType As Integer
    Dim dateCurrent As Date
    Dim strFilter As String
    Dim str��ҩ����  As String
    
    mintBeginRow = 0
    mintEndRow = 0
        
    If mcondition.strTransStep = M_STR_CALSS_AUDIT Then
        intType = 0
    ElseIf mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
        intType = 1
    ElseIf mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
        intType = 2
    End If
    
    str��ҩ���� = Me.cboType.Text
    If Me.cboType.ListIndex = 0 Then str��ҩ���� = ""
    If Not blnFilter Then
        Call InitMedi(intType, strIDS, str��ҩ����)
    End If
    
    For i = 0 To Me.ImgResult.count - 1
        If Me.chkResult(i).Value = 1 Then
            strFilter = IIf(strFilter = "", "�����=" & i, strFilter & " or �����=" & i)
        End If
    Next
    
    If mrsMedi Is Nothing Then Exit Sub
    mrsMedi.Filter = strFilter
    With Me.vsfMedis
        If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
            .ColHidden(.ColIndex("��")) = True
            .ColHidden(.ColIndex("ѡ��")) = False
        ElseIf mcondition.strTransStep = M_STR_CALSS_AUDIT Then
            .ColHidden(.ColIndex("��")) = False
            .ColHidden(.ColIndex("ѡ��")) = True
        End If
        
        .rows = 1
        .rows = 2
        If mrsMedi.RecordCount = 0 Then Exit Sub
        
        dateCurrent = Sys.Currentdate
        
        .Redraw = flexRDNone
        
        .MergeCells = flexMergeFree
        .rows = mrsMedi.RecordCount + 1
        
        i = 1
        mrsMedi.MoveFirst
        Do While Not mrsMedi.EOF
            If lng����id <> Val(mrsMedi!����ID) Then
                lng����id = Val(mrsMedi!����ID)
                If intBlackColor <> 1 Then
                    intBlackColor = 1
                Else
                    intBlackColor = 2
                End If
            End If
            
            If lngҽ��id <> Val(mrsMedi!���id) Then
                If lngҽ��id <> 0 Then
                    .rows = .rows + 1
                    .RowHidden(i) = True
                    For j = 0 To .Cols - 1
                        .TextMatrix(i, j) = "00"
                    Next
                    i = i + 1
                End If
                lngҽ��id = Val(mrsMedi!���id)
            Else
                .MergeCol(.ColIndex("ѡ��")) = True
                .MergeCol(.ColIndex("��")) = True
                .MergeCol(.ColIndex("ҽ��id")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("�Ա�")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("סԺ��")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("����ҽ��")) = True
                .MergeCol(.ColIndex("������")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("����ID")) = True
                .MergeCol(.ColIndex("����ID")) = True
                .MergeCol(.ColIndex("����ID")) = True
                .MergeCol(.ColIndex("��ҳID")) = True
                .MergeCol(.ColIndex("����")) = True
                .MergeCol(.ColIndex("ҩƷid")) = True
                .MergeCol(.ColIndex("ҩ��")) = True
            End If
            
            .TextMatrix(i, .ColIndex("��־")) = "0"
            .TextMatrix(i, .ColIndex("��")) = " "
            .TextMatrix(i, .ColIndex("�����")) = Val(mrsMedi!��˱�־)
            .TextMatrix(i, .ColIndex("������")) = intBlackColor
            .TextMatrix(i, .ColIndex("����id")) = Val(mrsMedi!����ID)
            .TextMatrix(i, .ColIndex("��ҳid")) = Val(mrsMedi!��ҳid)
            .TextMatrix(i, .ColIndex("���id")) = Val(mrsMedi!���id)
            .TextMatrix(i, .ColIndex("ҽ��id")) = Val(mrsMedi!Id)
            .TextMatrix(i, .ColIndex("����")) = zlStr.nvl(mrsMedi!����)
            .TextMatrix(i, .ColIndex("�Ա�")) = zlStr.nvl(mrsMedi!�Ա�)
            .TextMatrix(i, .ColIndex("����")) = zlStr.nvl(mrsMedi!����)
            .TextMatrix(i, .ColIndex("����")) = IIf(zlStr.nvl(mrsMedi!����) = "", "<��>", mrsMedi!����)
            .TextMatrix(i, .ColIndex("סԺ��")) = zlStr.nvl(mrsMedi!סԺ��, " ")
            .TextMatrix(i, .ColIndex("����")) = zlStr.nvl(mrsMedi!����, " ")
            .TextMatrix(i, .ColIndex("����ID")) = zlStr.nvl(mrsMedi!���˲���ID, " ")
            .TextMatrix(i, .ColIndex("����ID")) = zlStr.nvl(mrsMedi!���˿���id, " ")
            .TextMatrix(i, .ColIndex("����")) = zlStr.nvl(mrsMedi!����)
            .TextMatrix(i, .ColIndex("����ҽ��")) = zlStr.nvl(mrsMedi!����ҽ��)
            .TextMatrix(i, .ColIndex("ҩƷ����")) = zlStr.nvl(mrsMedi!ҩƷ����) & IIf(zlStr.nvl(mrsMedi!���) = "", "", "��" & zlStr.nvl(mrsMedi!���))
            .TextMatrix(i, .ColIndex("ҩʦ���ԭ��")) = zlStr.nvl(mrsMedi!ҩʦ���ԭ��)
            .TextMatrix(i, .ColIndex("����")) = nvl(mrsMedi!�����)
            .TextMatrix(i, .ColIndex("ҩƷid")) = Val(nvl(mrsMedi!ҩƷID))
            .TextMatrix(i, .ColIndex("ҩ��")) = nvl(mrsMedi!ҩƷ����)
            .TextMatrix(i, .ColIndex("��Ч")) = nvl(mrsMedi!ҽ����Ч)
            .TextMatrix(i, .ColIndex("��ҩ;��")) = nvl(mrsMedi!��ҩ;��)
            
            
            If Val(mrsMedi!��˱�־) <> 0 And Val(mrsMedi!��˱�־) <> 3 And mcondition.strTransStep = M_STR_CALSS_AUDIT Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &HFF0000
            End If
            
            
            '���ú�����ҩ��־ (PASS)
            If Not gobjPass Is Nothing Then
                .Cell(flexcpPicture, i, .ColIndex("��"), i, .ColIndex("��")) = gobjPass.zlPassSetWarnLight_YF(nvl(mrsMedi!�����, 0))
                .Cell(flexcpPictureAlignment, i, .ColIndex("��"), i, .ColIndex("��")) = flexPicAlignCenterCenter
            End If

            '��ʾ[�Ա�ҩ]��־
            If mrsMedi!ִ������ = 5 And mrsMedi!ִ�б�� = 0 Then
                .Cell(flexcpPicture, i, .ColIndex("ҩƷ����"), i, .ColIndex("ҩƷ����")) = Me.ImgPro.ListImages("�Ա�ҩ").Picture
                .Cell(flexcpPictureAlignment, i, .ColIndex("ҩƷ����"), i, .ColIndex("ҩƷ����")) = flexPicAlignLeftCenter
            End If

            If mrsMedi!�Ƿ�Ƥ�� = 1 Then
                .TextMatrix(i, .ColIndex("Ƥ")) = GetƤ�Խ��(Val(mrsMedi!����ID), Val(mrsMedi!ҩ��ID), dateCurrent, mrsMedi!����ʱ��, mrsMedi!��ҳid)
            End If
            
            .TextMatrix(i, .ColIndex("���")) = zlStr.nvl(mrsMedi!���)
            .TextMatrix(i, .ColIndex("����")) = Format(zlStr.nvl(mrsMedi!����), "#####0.00000;-#####0.00000; ;")
            .TextMatrix(i, .ColIndex("��λ")) = zlStr.nvl(mrsMedi!��λ)
            .TextMatrix(i, .ColIndex("Ƶ��")) = zlStr.nvl(mrsMedi!Ƶ��)
            .TextMatrix(i, .ColIndex("ִ��ʱ��")) = zlStr.nvl(mrsMedi!ִ��ʱ��)
            
            If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
                .TextMatrix(i, .ColIndex("��ҩ��־")) = IIf(CheckIs��ҩ(Val(mrsMedi!���id)) = True, 1, 0)
            End If
            If Val(.TextMatrix(i, .ColIndex("��ҩ��־"))) = 1 Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
            
            .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = IIf(intBlackColor = 1, &H80000005, &HC0FFC0)
            i = i + 1
            mrsMedi.MoveNext
        Loop
        
        .Cell(flexcpFontBold, 0, .ColIndex("��"), 0, .ColIndex("��")) = True
        .Cell(flexcpForeColor, 0, .ColIndex("��"), 0, .ColIndex("��")) = vbBlue
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Function CheckIs��ҩ(ByVal lngҽ��id As Long) As Boolean
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 1 from ��Һ��ҩ��¼ A,����ҽ����¼ B where A.ҽ��id=B.id and A.����״̬>1 and A.����״̬<>12 and b.id=[1]"
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIs��ҩ", lngҽ��id)
    
    CheckIs��ҩ = (Not rstemp.EOF)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mobjMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    '1.������Ϣ����Ϣ���ͺ�����ҵ��Լ��
    '2.���ݿͻ������������ж��Ƿ�����Ч��Ϣ
    Dim i As Integer
    Const CST_INT_MSGREFRESHINTERVAL As Integer = 1
    Const CST_STR_MSGCODE As String = "ZLHIS_CIS_003,ZLHIS_CIS_013,ZLHIS_CIS_008"
    
    '��Ϣ����Ϊ��ʱ�˳�
    If mobjMipModule Is Nothing Then Exit Sub
    
    '��Ϣ��������ʧ��ʱ��������Ϣ
    If mobjMipModule.IsConnect = False Then Exit Sub
        
    '�����Һ�������Ľ��յ���Ϣ����
    If InStr("," & CST_STR_MSGCODE & ",", "," & strMsgItemIdentity & ",") = 0 Then Exit Sub

    '���ݿͻ������������ж��Ƿ�����Ч��Ϣ
    Call IsValidMsg(strMsgItemIdentity, strMsgContent)
    
    
    '������յ���Ч��Ϣʱ�������±��
    If Not mrsMsg Is Nothing Then
    lblMsgComment.Caption = "��Ϣ����(" & mrsMsg.RecordCount & ")"
    mrsMsg.MoveFirst
    With Me.vsfMsg
        .rows = 1
        .rows = mrsMsg.RecordCount + 1
        .RowHeight(0) = 300
        
        For i = 1 To mrsMsg.RecordCount
            .RowHeight(i) = 300
            .TextMatrix(i, .ColIndex("ʱ��")) = mrsMsg!ʱ��
            .TextMatrix(i, .ColIndex("����")) = mrsMsg!����
            .TextMatrix(i, .ColIndex("����")) = mrsMsg!����
            .TextMatrix(i, .ColIndex("����")) = mrsMsg!����
            .TextMatrix(i, .ColIndex("ִ��ʱ��")) = mrsMsg!ִ��ʱ��
'            .TextMatrix(i, .ColIndex("����״̬")) = mrsMsg!����״̬
            .TextMatrix(i, .ColIndex("���˲���id")) = mrsMsg!����ID
            .TextMatrix(i, .ColIndex("����id")) = mrsMsg!����ID
            
            mrsMsg.MoveNext
        Next
    End With
    End If
End Sub

Private Sub IsValidMsg(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    Dim rsMsg As Recordset
    Dim strTemp As String
    Dim strSQL As String
    Dim rstemp As Recordset
    Dim blnNext As Boolean
    Dim str�������� As String
    Dim str�������� As String
    Dim strʱ�� As String
    Dim str���� As String
    Dim objXML As New zl9ComLib.clsXML
    
    On Error GoTo ErrHand
    
    If objXML Is Nothing Then Exit Sub

    '��XML�ļ�
    objXML.OpenXMLDocument strMsgContent
    
    If strMsgItemIdentity = "ZLHIS_CIS_003" Then
'        str���� = "ҽ������"
'
'        If objXML.GetMultiNodeRecord("cancel_order", rsMsg) = False Then Exit Sub
'        If rsMsg Is Nothing Then Exit Sub
'        If rsMsg.RecordCount = 0 Then Exit Sub
'
'        '��ȡҽ��ID,����ҽ��ID�Ƿ������ҩ��¼
'        If objXML.GetSingleNodeValue("order_id", strTemp, xsString) = False Then Exit Sub
'        If objXML.GetSingleNodeValue("cancel_time", strʱ��, xsString) = False Then Exit Sub
'
'        strSQL = "select A.ID,A.ҽ��id,A.����,A.�Ա�,A.����,A.����,A.ִ��ʱ��,A.����״̬,B.ID ����ID,B.���� ��������,C.ID ����ID,C.���� �������� from ��Һ��ҩ��¼ A ,���ű� B,���ű� C where B.id=A.���˲���ID And  C.id=A.���˿���ID And A.ҽ��id=[1] and A.����ID=[2]"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "IsValidMsg", strTemp, mParams.lng��������)
'        If Not rsTemp.EOF Then blnNext = True
        
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_013" Then
        str���� = "��������"
    
        If objXML.GetMultiNodeRecord("cancel_reqeust", rsMsg) = False Then Exit Sub
        If rsMsg Is Nothing Then Exit Sub
        If rsMsg.RecordCount = 0 Then Exit Sub
        
        '��ȡ��ҩID,������ҩ��¼�Ƿ�Ϊ��ǰ���ŵ����������¼
        If objXML.GetSingleNodeValue("transfusion_id", strTemp, xsString) = False Then Exit Sub
        If objXML.GetSingleNodeValue("request_time", strʱ��, xsString) = False Then Exit Sub
        
        strSQL = "select A.ID,A.ҽ��id,A.����,A.�Ա�,A.����,A.����,A.ִ��ʱ��,A.����״̬,B.ID ����ID,B.���� ��������,C.ID ����ID,C.���� �������� from ��Һ��ҩ��¼ A ,���ű� B,���ű� C where B.id=A.���˲���ID And  C.id=A.���˿���ID And A.Id=[1] and A.����ID=[2] "
        Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "IsValidMsg", strTemp, mParams.lng��������)
        If Not rstemp.EOF Then blnNext = True
        
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_008" Then
        str���� = "���ε���"
        strʱ�� = Now
    
        If objXML.GetMultiNodeRecord("transfusion_info", rsMsg) = False Then Exit Sub
        If rsMsg Is Nothing Then Exit Sub
        If rsMsg.RecordCount = 0 Then Exit Sub
        
        '��ȡ��ҩID,������ҩ��¼�Ƿ�Ϊ��ǰ���ŵ���Һ��ҩ��¼
        If objXML.GetSingleNodeValue("transfusion_id", strTemp, xsString) = False Then Exit Sub
        
        strSQL = "select A.ID,A.ҽ��id,A.����,A.�Ա�,A.����,A.����,A.ִ��ʱ��,A.����״̬,B.ID ����ID,B.���� ��������,C.ID ����ID,C.���� �������� from ��Һ��ҩ��¼ A ,���ű� B,���ű� C where B.id=A.���˲���ID And  C.id=A.���˿���ID And A.Id=[1] and A.����ID=[2]"
        Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "IsValidMsg", strTemp, mParams.lng��������)
        If Not rstemp.EOF Then blnNext = True
        
    End If
    
    
    '���������������������ݼ��ͽ�������
    If blnNext Then
        Call mobjMipModule.ShowMessage(strMsgItemIdentity, "������" & str���� & ",�����Աע��鿴", str���� & "����", "���ѹ�����Ա", "����id=1234|����id=344899")
    
        If mrsMsg Is Nothing Then Call InitMsgRs
        With mrsMsg
            .AddNew
            !��ҩid = rstemp!Id
            !���� = rstemp!���� & " " & rstemp!�Ա� & " " & rstemp!���� & " " & rstemp!�������� & " " & rstemp!����          '�������Ա����䣬���ң���λ
            !���� = rstemp!��������
            !ʱ�� = strʱ��
            !���� = str����
            !ҽ��id = rstemp!ҽ��id
            !ִ��ʱ�� = rstemp!ִ��ʱ��
            !����״̬ = rstemp!����״̬
            !����ID = rstemp!����ID
            !����ID = rstemp!����ID
            
            .Update
        End With
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitMsgRs()
    Set mrsMsg = New ADODB.Recordset
    With mrsMsg
        If .State = 1 Then .Close
        
        '�ü�¼��Ӧ����Ϣ��Ϣ
        .Fields.Append "��ҩID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҽ��ID", adDouble, 3, adFldIsNullable
        .Fields.Append "ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ִ��ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����״̬", adDouble, 3, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub SendMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ϣ���ʹ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, objDrugXML As New clsXML, objCheckXML As New clsXML
    Dim objTemp As clsXML, str�շ�ʱ�� As String
    Dim rstemp As ADODB.Recordset, int���� As Integer
    Dim blnֱ���շ� As Boolean, p As Long
    Dim lngDrug As Long, lngCheck As Long, blnAddBill As Boolean, blnHaveCheck As Boolean, blnHaveDrug As Boolean
    On Error GoTo errHandle
    Dim i As Integer
    
    If mobjMipModule Is Nothing Then Exit Sub
'    If mobjMipModule.IsConnect = False Then Exit Sub
        
    
'    transfuse_order ҽ����Ϣ
'    order_id ҽ��id
'    order_reason ���ԭ��
'    send_serial ���ͺ�
'    in_patient ������Ϣ
'    patient_id ����id
'    patient_name ����
'    in_number סԺ��
'    patient_clinic ������Ϣ
'    clinic_id ��ҳid
'    clinic_area_id ���ﲡ��id
'    clinic_dept_id �������id
'    clinic_bed ���ﲡ��


    objDrugXML.ClearXmlText
    objCheckXML.ClearXmlText
    
    With mrsSendMsg
        If .RecordCount = 0 Then Exit Sub
        
        .MoveFirst
        For i = 1 To .RecordCount
            '��ѯ�ܾ�����
            gstrSQL = "Select ���id From ����ҽ����¼ Where ID = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�ܾ�����", !ҽ��id)
            
            gstrSQL = "Select Max(ҩʦ���ԭ��) As ���ԭ�� From ����ҽ����¼ Where ���id = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�ܾ�����", rstemp!���id)
        
            'ҽ����Ϣ
            Call objDrugXML.AppendNode("transfuse_order")
                Call objDrugXML.AppendData("order_id", !ҽ��id)
                Call objDrugXML.AppendData("send_serial", !���ͺ�)
                Call objDrugXML.AppendData("order_reason", rstemp!���ԭ��)
            
            '������Ϣ
            Call objDrugXML.AppendNode("in_patient")
                Call objDrugXML.AppendData("patient_id", !����ID)
                Call objDrugXML.AppendData("patient_name", !����)
                Call objDrugXML.AppendData("in_number", !סԺ��)
            Call objDrugXML.AppendNode("in_patient", True)
            
            '������Ϣ
            Call objDrugXML.AppendNode("patient_clinic")
                Call objDrugXML.AppendData("clinic_id", !��ҳid)
                Call objDrugXML.AppendData("clinic_area_id", !����ID)
                Call objDrugXML.AppendData("clinic_dept_id", !����ID)
                Call objDrugXML.AppendData("clinic_bed", IIf(!���� = "<��>", "", Replace(zlStr.nvl(!����, ""), "��", "")))
            Call objDrugXML.AppendNode("patient_clinic", True)
            
            
            Call objDrugXML.AppendNode("transfuse_order", True)
            '������Ϣ
'            Call zlDebugWriteFile(objDrugXML.XmlText)
            Call mobjMipModule.CommitMessage("ZLHIS_TRANSFUSION_001", objDrugXML.XmlText)
            Call zlDatabase.SendMsg("ZLHIS_TRANSFUSION_001", objDrugXML.XmlText)
            objDrugXML.ClearXmlText: objCheckXML.ClearXmlText
            Set objDrugXML = Nothing: Set objCheckXML = Nothing
            
            .MoveNext
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub
 
 
 
Private Sub InitSendMsgRs()
    Set mrsSendMsg = New ADODB.Recordset
    With mrsSendMsg
        If .State = 1 Then .Close
        
        '�ü�¼��Ӧ����Ϣ��Ϣ
        .Fields.Append "ҽ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "���ͺ�", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ҳid", adDouble, 18, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "����id", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub












