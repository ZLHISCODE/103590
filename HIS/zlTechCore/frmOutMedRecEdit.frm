VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmOutMedRecEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ҳ"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmOutMedRecEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraInfo 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   0
      Left            =   195
      TabIndex        =   34
      Top             =   465
      Width           =   6480
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��"
         Height          =   255
         Index           =   5
         Left            =   6060
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   2265
         Width           =   285
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   900
         MaxLength       =   30
         TabIndex        =   14
         Top             =   2235
         Width           =   5475
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   2
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   495
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��"
         Height          =   255
         Index           =   6
         Left            =   6060
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   2625
         Width           =   285
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "��"
         Height          =   255
         Index           =   9
         Left            =   6060
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   3345
         Width           =   285
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   4905
         MaxLength       =   6
         TabIndex        =   23
         Top             =   3675
         Width           =   1470
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   900
         MaxLength       =   20
         TabIndex        =   22
         Top             =   3675
         Width           =   3090
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   9
         Left            =   900
         MaxLength       =   50
         TabIndex        =   20
         Top             =   3315
         Width           =   5475
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   4905
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2955
         Width           =   1470
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   900
         MaxLength       =   20
         TabIndex        =   18
         Top             =   2955
         Width           =   3090
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   900
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2595
         Width           =   5475
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   900
         MaxLength       =   18
         TabIndex        =   13
         Top             =   1875
         Width           =   5475
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   6
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1365
         Width           =   2355
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   5
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1365
         Width           =   2355
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   4
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1005
         Width           =   2355
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   3
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1005
         Width           =   2355
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   495
         Width           =   615
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   135
         Width           =   1305
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   3030
         MaxLength       =   5
         TabIndex        =   6
         Top             =   495
         Width           =   675
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   900
         MaxLength       =   20
         TabIndex        =   1
         Top             =   135
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         BackColor       =   &H8000000F&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   1200
      End
      Begin MSMask.MaskEdBox txt����ʱ�� 
         Height          =   300
         Left            =   1950
         TabIndex        =   5
         Top             =   495
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt�������� 
         Height          =   300
         Left            =   900
         TabIndex        =   4
         Top             =   495
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   56
         Top             =   555
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   51
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʽ"
         Height          =   180
         Index           =   5
         Left            =   4380
         TabIndex        =   50
         Top             =   555
         Width           =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -105
         X2              =   7245
         Y1              =   1755
         Y2              =   1755
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   -150
         X2              =   7200
         Y1              =   1770
         Y2              =   1770
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   -15
         X2              =   7335
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   -60
         X2              =   7290
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�ʱ�"
         Height          =   180
         Index           =   17
         Left            =   4095
         TabIndex        =   49
         Top             =   3735
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   180
         Index           =   16
         Left            =   120
         TabIndex        =   48
         Top             =   3735
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ"
         Height          =   180
         Index           =   15
         Left            =   120
         TabIndex        =   47
         Top             =   3375
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         Height          =   180
         Index           =   14
         Left            =   4095
         TabIndex        =   46
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   45
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   44
         Top             =   2655
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   43
         Top             =   1935
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   42
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   7
         Left            =   3555
         TabIndex        =   41
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   180
         Index           =   9
         Left            =   3555
         TabIndex        =   40
         Top             =   1425
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   180
         Index           =   8
         Left            =   120
         TabIndex        =   39
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   2610
         TabIndex        =   38
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   1
         Left            =   2610
         TabIndex        =   37
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   36
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   35
         Top             =   195
         Width           =   540
      End
   End
   Begin VB.Frame fraInfo 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   1
      Left            =   195
      TabIndex        =   52
      Top             =   465
      Width           =   6480
      Begin VB.CheckBox chkEdit 
         Caption         =   "����"
         Height          =   195
         Index           =   0
         Left            =   5505
         TabIndex        =   30
         Top             =   3195
         Width           =   750
      End
      Begin VB.CommandButton cmdMakeLog 
         Height          =   255
         Left            =   1260
         Picture         =   "frmOutMedRecEdit.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "����������ɾ���ժҪ(F12)"
         Top             =   3135
         Width           =   345
      End
      Begin VB.TextBox txtEdit 
         Height          =   660
         Index           =   12
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   3405
         Width           =   6405
      End
      Begin VB.OptionButton optInput 
         Caption         =   "������ϱ�׼����(&1)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   0
         Left            =   2295
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1230
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.OptionButton optInput 
         Caption         =   "���ݼ�����������(&2)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   1
         Left            =   4365
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1230
         Width           =   2010
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
         Height          =   960
         Left            =   30
         TabIndex        =   27
         Top             =   1440
         Width           =   6405
         _cx             =   11298
         _cy             =   1693
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
         FixedCols       =   1
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOutMedRecEdit.frx":0102
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   115
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
      Begin VSFlex8Ctl.VSFlexGrid vsAller 
         Height          =   960
         Left            =   30
         TabIndex        =   24
         Top             =   225
         Width           =   6405
         _cx             =   11298
         _cy             =   1693
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOutMedRecEdit.frx":019E
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
      Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
         Height          =   735
         Left            =   30
         TabIndex        =   28
         Top             =   2400
         Width           =   6405
         _cx             =   11298
         _cy             =   1296
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
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOutMedRecEdit.frx":01EF
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   115
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
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   " ����ժҪ "
         Height          =   180
         Index           =   20
         Left            =   285
         TabIndex        =   55
         Top             =   3195
         Width           =   900
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000014&
         X1              =   75
         X2              =   5300
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000010&
         X1              =   75
         X2              =   5300
         Y1              =   3285
         Y2              =   3285
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   " ��ϼ�¼ "
         Height          =   180
         Index           =   19
         Left            =   285
         TabIndex        =   54
         Top             =   1230
         Width           =   900
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000014&
         X1              =   75
         X2              =   6400
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000010&
         X1              =   75
         X2              =   6400
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   " ������¼ "
         Height          =   180
         Index           =   18
         Left            =   285
         TabIndex        =   53
         Top             =   15
         Width           =   900
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000014&
         X1              =   75
         X2              =   6400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   75
         X2              =   6400
         Y1              =   105
         Y2              =   105
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4110
      TabIndex        =   32
      ToolTipText     =   "�ȼ���F2"
      Top             =   4740
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5280
      TabIndex        =   33
      Top             =   4740
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tbsInfo 
      Height          =   4515
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   7964
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������Ϣ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������Ϣ"
            ImageVarType    =   2
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
End
Attribute VB_Name = "frmOutMedRecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnDiagnose As Boolean
Private mlng����ID As Long
Private mstr�Һŵ� As String
Private mlng�Һ�ID As Long
Private mint���� As Integer
Private mbln��ҽ As Boolean
Private mstrLike As String
Private mint���� As Integer
Private mblnChange As Boolean
Private mblnOK As Boolean

Private Enum TXT_ENUM
    txt���� = 0
    txt����� = 1
    'txt�������� = 2
    txt���� = 3
    txt���֤�� = 4
    txt�����ص� = 5
    txt������λ = 6
    txt��λ�绰 = 7
    txt��λ�ʱ� = 8
    txt��ͥ��ַ = 9
    txt��ͥ�绰 = 10
    txt��ͥ�ʱ� = 11
    txt����ժҪ = 12
End Enum
Private Enum CBO_ENUM
    cbo�Ա� = 0
    cbo���� = 1
    cbo���� = 2
    cbo���� = 3
    cbo���� = 4
    cbo���� = 5
    cboְҵ = 6
End Enum
Private Enum CHK_ENUM
    chk���� = 0
End Enum
Private Enum COL_ENUM
    col���� = 0
    col��� = 1
    col���� = 2
    col���ID = 3
    col����ID = 4
    col֤��ID = 5
End Enum

Public Function ShowMe(frmParent As Object, ByVal str�Һŵ� As String, Optional blnDiagnose As Boolean) As Boolean
'������blnDiagnose=�Ƿ����������д���
'���أ�blnDiagnose=�Ƿ���д�˲��˵����
    mblnDiagnose = blnDiagnose
    mstr�Һŵ� = str�Һŵ�
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    blnDiagnose = mblnDiagnose
    ShowMe = mblnOK
End Function

Private Function InitMedData() As Boolean
'���ܣ���ʼ���༭�����ͱ�Ҫ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    Call zlControl.CboSetHeight(cboEdit(cbo����), cboEdit(cbo����).Height * 16)
    Call zlControl.CboSetHeight(cboEdit(cbo����), cboEdit(cbo����).Height * 16)
    Call zlControl.CboSetHeight(cboEdit(cboְҵ), cboEdit(cboְҵ).Height * 16)
    vsDiagXY.MergeCol(0) = True
    vsDiagZY.MergeCol(0) = True
    
    Call SetCboFromList(Array("��", "��", "��"), Array(cbo����), 0)
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From �Ա� Order by ����", Array(cbo�Ա�))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ҽ�Ƹ��ʽ Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ���� Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ���� Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ����״�� Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ְҵ Order by ����", Array(cboְҵ))
    
    optInput(0).TabStop = False: optInput(1).TabStop = False 'Ҫǿ�д���ִ��һ��
    
    InitMedData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadMedRec() As Boolean
'���ܣ���ȡ������ҳ�ĸ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngRow As Long
    
    On Error GoTo errH
    
    '������Ϣ
    strSQL = "Select A.����ID,B.ID as �Һ�ID,B.ִ�в���ID as ����ID,B.ժҪ,B.����," & _
        " A.����,A.�����,A.����,A.�Ա�,A.����,A.��������,A.ҽ�Ƹ��ʽ," & _
        " A.����,A.����,A.����״��,A.ְҵ,A.���֤��,A.�����ص�," & _
        " A.��ͥ��ַ,A.��ͥ�绰,A.�����ʱ�,A.������λ,A.��λ�绰,A.��λ�ʱ�" & _
        " From ������Ϣ A,���˹Һż�¼ B Where A.����ID=B.����ID And B.NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
    If rsTmp.EOF Then Exit Function
    
    mlng����ID = rsTmp!����ID
    mlng�Һ�ID = rsTmp!�Һ�ID
    mint���� = Nvl(rsTmp!����, 0)
    mbln��ҽ = Have��������(rsTmp!����ID, "��ҽ��")
        
    txtEdit(txt����).Text = rsTmp!����
    Call GetCboIndex(cboEdit(cbo�Ա�), Nvl(rsTmp!�Ա�))
    txtEdit(txt�����).Text = Nvl(rsTmp!�����)
    
    If Not IsNull(rsTmp!��������) Then
        txt��������.Text = Format(rsTmp!��������, "yyyy-MM-dd")
        If Format(rsTmp!��������, "HH:mm") <> "00:00" Then
            txt����ʱ��.Text = Format(rsTmp!��������, "HH:mm")
        End If
    End If
    
    Call LoadOldData(Nvl(rsTmp!����))
    
    Call txt��������_Validate(False)
    
    If IsNumeric(txtEdit(txt����).Text) Then
         If Val(txtEdit(txt����).Text) <> CLng(txtEdit(txt����).Text) Then
            cboEdit(cbo����).ListIndex = 2                    '����Ϊ��λ
            txtEdit(txt����).Text = CLng(Val(txtEdit(txt����).Text) * 365)
        End If
    End If
    
    Call GetCboIndex(cboEdit(cbo����), Nvl(rsTmp!ҽ�Ƹ��ʽ))
    Call GetCboIndex(cboEdit(cbo����), Nvl(rsTmp!����))
    Call GetCboIndex(cboEdit(cbo����), Nvl(rsTmp!����))
    Call GetCboIndex(cboEdit(cbo����), Nvl(rsTmp!����״��))
    Call GetCboIndex(cboEdit(cboְҵ), Nvl(rsTmp!ְҵ))
    txtEdit(txt���֤��).Text = Nvl(rsTmp!���֤��)
    txtEdit(txt�����ص�).Text = Nvl(rsTmp!�����ص�)
    txtEdit(txt������λ).Text = Nvl(rsTmp!������λ)
    txtEdit(txt��λ�绰).Text = Nvl(rsTmp!��λ�绰)
    txtEdit(txt��λ�ʱ�).Text = Nvl(rsTmp!��λ�ʱ�)
    txtEdit(txt��ͥ��ַ).Text = Nvl(rsTmp!��ͥ��ַ)
    txtEdit(txt��ͥ�绰).Text = Nvl(rsTmp!��ͥ�绰)
    txtEdit(txt��ͥ�ʱ�).Text = Nvl(rsTmp!�����ʱ�)
    txtEdit(txt����ժҪ).Text = Nvl(rsTmp!ժҪ)
    chkEdit(chk����).Value = Nvl(rsTmp!����, 0)
    
    '������Ϣ:���ιҺŵ�,������
    strSQL = "Select ��¼��Դ,��¼ʱ��,ҩ��ID,ҩ���� From ���˹�����¼ A" & _
        " Where ���=1 And ����ID=[1] And ��ҳID=[2]" & _
        " And Not Exists(Select ҩ��ID From ���˹�����¼" & _
            " Where (Nvl(ҩ��ID,0)=Nvl(A.ҩ��ID,0) Or Nvl(ҩ����,'Null')=Nvl(A.ҩ����,'Null'))" & _
            " And Nvl(���,0)=0 And ��¼ʱ��>=A.��¼ʱ�� And ����ID=[1] And ��ҳID=[2])" & _
        " Order by ��¼ʱ��,ҩ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��¼��Դ=3" '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
        With vsAller
            .Rows = rsTmp.RecordCount + 2 '�̶���+����
            For i = 1 To rsTmp.RecordCount
                '������Դ�Ŀ������ظ�
                lngRow = -1
                If Not IsNull(rsTmp!ҩ��ID) Then
                    lngRow = .FindRow(CLng(rsTmp!ҩ��ID))
                ElseIf Not IsNull(rsTmp!ҩ����) Then
                    lngRow = .FindRow(CStr(rsTmp!ҩ����), , 1)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(Nvl(rsTmp!ҩ��ID, 0))
                    .TextMatrix(i, 0) = Format(rsTmp!��¼ʱ��, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, 0) = Format(rsTmp!��¼ʱ��, "yyyy-MM-dd HH:mm:ss") '���ڱ���
                    .TextMatrix(i, 1) = Nvl(rsTmp!ҩ����)
                    .Cell(flexcpData, i, 1) = .TextMatrix(i, 1) '��������ָ�
                End If
                rsTmp.MoveNext
            Next
        End With
    End If
    vsAller.Row = 1: vsAller.Col = 1
    
    '�����Ϣ:���ιҺŵ�
    strSQL = "Select ��¼��Դ,�������,����ID,���ID,֤��ID,�������,�Ƿ����� From ������ϼ�¼" & _
        " Where ��¼��Դ IN(1,3) And ������� IN(1,11)" & _
        " And ȡ��ʱ�� is Null And ����ID=[1] And ��ҳID=[2]" & _
        " Order by �������,��ϴ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    If Not rsTmp.EOF Then
        '��ҽ���
        rsTmp.Filter = "�������=1 And ��¼��Դ=3" '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "�������=1 And ��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
        With vsDiagXY
            .Rows = rsTmp.RecordCount + 2
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col���) = Nvl(rsTmp!�������)
                .Cell(flexcpData, i, col���) = .TextMatrix(i, col���)
                .TextMatrix(i, col����) = IIF(Nvl(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                .TextMatrix(i, col���ID) = Nvl(rsTmp!���ID, 0)
                .TextMatrix(i, col����ID) = Nvl(rsTmp!����id, 0)
                rsTmp.MoveNext
            Next
            .Cell(flexcpText, .FixedRows, col����, .Rows - 1, col����) = "��ҽ"
            .Cell(flexcpForeColor, .FixedRows, col����, .Rows - 1, col����) = vbRed
        End With
        '��ҽ���
        If mbln��ҽ Then
            rsTmp.Filter = "�������=11 And ��¼��Դ=3"
            If rsTmp.EOF Then rsTmp.Filter = "�������=11 And ��¼��Դ<>3"
            With vsDiagZY
                .Rows = rsTmp.RecordCount + 1
                For i = 0 To rsTmp.RecordCount - 1
                    .TextMatrix(i, col���) = Nvl(rsTmp!�������)
                    .Cell(flexcpData, i, col���) = .TextMatrix(i, col���)
                    .TextMatrix(i, col����) = IIF(Nvl(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                    .TextMatrix(i, col���ID) = Nvl(rsTmp!���ID, 0)
                    .TextMatrix(i, col����ID) = Nvl(rsTmp!����id, 0)
                    .TextMatrix(i, col֤��ID) = Nvl(rsTmp!֤��ID, 0)
                    rsTmp.MoveNext
                Next
                .Cell(flexcpText, .FixedRows, col����, .Rows - 1, col����) = "��ҽ"
                .Cell(flexcpForeColor, .FixedRows, col����, .Rows - 1, col����) = vbRed
            End With
        End If
    End If
    vsDiagXY.Row = vsDiagXY.FixedRows: vsDiagXY.Col = 0: vsDiagXY.Col = col���
    vsDiagZY.Row = vsDiagZY.FixedRows: vsDiagZY.Col = 0: vsDiagZY.Col = col���
        
    If Not mbln��ҽ Then
        vsDiagZY.Visible = False
        vsDiagXY.Height = vsDiagZY.Top + vsDiagZY.Height - vsDiagXY.Top
        vsDiagXY.ColHidden(0) = True
        vsDiagXY.ColWidth(1) = vsDiagXY.ColWidth(1) + vsDiagXY.ColWidth(0)
    End If
    
    LoadMedRec = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CalcHowOld() As Long
'���ܣ����ݵ�ǰ�ĳ������ں����䵥λ����������ֵ
'���أ�-1��ʾδ����
    Dim datBase As Date, lngTmp As Long
        
    CalcHowOld = -1
    If Not (IsDate(txt��������.Text) And IsNumeric(txtEdit(txt����).Text)) Then Exit Function
    
    datBase = zlDatabase.Currentdate
    lngTmp = DateDiff("yyyy", CDate(txt��������.Text), datBase)
    
    'ֻ��������������м��
    If cboEdit(cbo����).ListIndex = 0 Then
        If Format(datBase, "MMdd") < Format(txt��������.Text, "MMdd") Then
            lngTmp = lngTmp - 1
        End If
        CalcHowOld = lngTmp
    End If
End Function

Private Function CheckMedRec(Optional blnDiagnose As Boolean) As Boolean
'���ܣ������ҳ�������ݺϷ���
'���أ�blnDiagnose=�Ƿ���д�����
    Dim objTmp As Object, curDate As Date
    Dim arrInfo() As Variant, arrName As Variant
    Dim str���֤ As String, str�������� As String, lng�Ա� As Long
    Dim i As Long, j As Long
    
    blnDiagnose = False
    curDate = zlDatabase.Currentdate
    
    '����Ҫ��������ݼ��
    '-----------------------------------------------------------------------------------------
    arrInfo = Array(txt����, txt����)
    arrName = Array("����", "����")
    For i = 0 To UBound(arrInfo)
        If txtEdit(arrInfo(i)).Enabled And Not txtEdit(arrInfo(i)).Locked And txtEdit(arrInfo(i)).Text = "" Then
            Call ShowMessage(txtEdit(arrInfo(i)), "�������벡�˵�" & arrName(i) & "��")
            Exit Function
        End If
    Next
    
    Select Case cboEdit(cbo����).Text
        Case "��"
            If Val(txtEdit(txt����).Text) > 200 Then
                MsgBox "����ֵ�������������Ƿ���ȷ��", vbInformation, gstrSysName
                txtEdit(txt����).SetFocus: Exit Function
            End If
        Case "��"
            If Val(txtEdit(txt����).Text) > 2400 Then
                MsgBox "����ֵ�������������Ƿ���ȷ��", vbInformation, gstrSysName
                txtEdit(txt����).SetFocus: Exit Function
            End If
        Case "��"
            If Val(txtEdit(txt����).Text) > 73000 Then
                MsgBox "����ֵ�������������Ƿ���ȷ��", vbInformation, gstrSysName
                txtEdit(txt����).SetFocus: Exit Function
            End If
        Case Else
            Exit Function
    End Select
    If Not IsDate(txt��������.Text) Then
        Call ShowMessage(txt��������, "�������벡�˵ĳ������ڡ�")
        Exit Function
    ElseIf txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        Call ShowMessage(txt����ʱ��, "��������ȷ�Ĳ��˳���ʱ�䡣")
        Exit Function
    End If
    
    i = CalcHowOld
    If i <> -1 And i <> Val(txtEdit(txt����).Text) Then
        If ShowMessage(txt��������, "����ͳ������ڲ�һ�£�" & txt��������.Text & "��������Ӧ����" & i & cboEdit(cbo����).Text & "��" & _
            vbCrLf & vbCrLf & "���������������ڵ���ȷ�ԣ�Ҫ������", True) = vbNo Then
            Exit Function
        End If
    End If
    
    arrInfo = Array(cbo����, cbo�Ա�)
    arrName = Array("���ʽ", "�Ա�")
    For i = 0 To UBound(arrInfo)
        If cboEdit(arrInfo(i)).Enabled And Not cboEdit(arrInfo(i)).Locked And cboEdit(arrInfo(i)).ListIndex = -1 Then
            Call ShowMessage(cboEdit(arrInfo(i)), "�������벡�˵�" & arrName(i) & "��")
            Exit Function
        End If
    Next
    
    '��Ŀ����ĳ��ȼ��
    '-----------------------------------------------------------------------------------------
    For Each objTmp In txtEdit
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "�������ݹ��������顣(����Ŀ������� " & objTmp.MaxLength & " ���ַ��� " & objTmp.MaxLength \ 2 & " ������)")
                Exit Function
            End If
        End If
    Next
    
    '�������ݵ���Ч�Լ��
    '-----------------------------------------------------------------------------------------
    '�������ڱ������ڵ�ǰʱ��
    If Format(txt��������.Text, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
        Call ShowMessage(txt��������, "�������ڲ�Ӧ�ñȵ�ǰ���ڻ���")
        Exit Function
    End If

    '15������ӦΪδ��
    If Not (cboEdit(cbo����).Text = "" Or cboEdit(cbo����).ListIndex = -1) Then
        If DateDiff("yyyy", CDate(txt��������.Text), curDate) < 15 Then
            If InStr(cboEdit(cbo����).Text, "�ѻ�") > 0 _
                Or InStr(cboEdit(cbo����).Text, "ɥż") > 0 Or InStr(cboEdit(cbo����).Text, "���") > 0 Then
                Call ShowMessage(cboEdit(cbo����), "����״����Ϣ��д���ԡ�")
                Exit Function
            End If
        End If
    End If
            
    '���֤������
    '�����֤�Ž�����֤
    str���֤ = txtEdit(txt���֤��).Text
    If str���֤ <> "" Then
        If Len(str���֤) <> 15 And Len(str���֤) <> 18 Then
            Call ShowMessage(txtEdit(txt���֤��), "���֤����ĳ��Ȳ���ȷ��ӦΪ15λ��18λ��")
            Exit Function
        End If

        If Len(str���֤) = 15 Then
            str�������� = Mid(str���֤, 7, 6)
            str�������� = Format(GetFullDate(str��������), "yyyy-MM-dd")
            lng�Ա� = Val(Right(str���֤, 1))
        Else
            str�������� = Mid(str���֤, 7, 8)
            str�������� = Format(GetFullDate(str��������), "yyyy-MM-dd")
            lng�Ա� = Val(Mid(str���֤, 17, 1))
        End If
        If Not IsDate(str��������) Then
            If ShowMessage(txtEdit(txt���֤��), "���֤�����еĳ���������Ϣ����ȷ���Ƿ������", True) = vbNo Then Exit Function
        Else
            If Format(str��������, "yyyy-MM-dd") <> Format(txt��������.Text, "yyyy-MM-dd") Then
                If ShowMessage(txtEdit(txt���֤��), "���֤�����еĳ���������Ϣ�벡�˵ĳ������ڲ������Ƿ������", True) = vbNo Then Exit Function
            End If
        End If
        If (lng�Ա� Mod 2 = 1 And InStr(cboEdit(cbo�Ա�).Text, "Ů") > 0) Or (lng�Ա� Mod 2 = 0 And InStr(cboEdit(cbo�Ա�).Text, "��") > 0) Then
            If ShowMessage(txtEdit(txt���֤��), "���֤�����е��Ա���Ϣ�벡�˵��Ա𲻷����Ƿ������", True) = vbNo Then Exit Function
        End If
    End If
    
    '��ϱ��ļ��
    '-----------------------------------------------------------------------------------------
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col���)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, col���)) > 150 Then
                    .Row = i: .Col = col���
                    Call ShowMessage(vsDiagXY, "�������̫����ֻ����150���ַ���75�����֡�")
                    Exit Function
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, col���)) <> "" Then
                        If .TextMatrix(j, col���) = .TextMatrix(i, col���) Then
                            .Row = i: .Col = col���
                            Call ShowMessage(vsDiagXY, "���ִ���������ͬ�������Ϣ��")
                            Exit Function
                        ElseIf Val(.TextMatrix(i, col���ID)) <> 0 Then
                            If Val(.TextMatrix(j, col���ID)) = Val(.TextMatrix(i, col���ID)) Then
                                .Row = i: .Col = col���
                                Call ShowMessage(vsDiagXY, "���ִ���������ͬ�������Ϣ��")
                                Exit Function
                            End If
                        ElseIf Val(.TextMatrix(i, col����ID)) <> 0 Then
                            If Val(.TextMatrix(j, col����ID)) = Val(.TextMatrix(i, col����ID)) Then
                                .Row = i: .Col = col���
                                Call ShowMessage(vsDiagXY, "���ִ���������ͬ�������Ϣ��")
                                Exit Function
                            End If
                        End If
                    End If
                Next
                blnDiagnose = True
            End If
        Next
    End With
        
    If mbln��ҽ Then
        With vsDiagZY
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col���)) <> "" Then
                    If zlCommFun.ActualLen(.TextMatrix(i, col���)) > 150 Then
                        .Row = i: .Col = col���
                        Call ShowMessage(vsDiagZY, "�������̫����ֻ����150���ַ���75�����֡�")
                        Exit Function
                    End If
                    For j = i + 1 To .Rows - 1
                        If Trim(.TextMatrix(j, col���)) <> "" Then
                            If .TextMatrix(j, col���) = .TextMatrix(i, col���) Then
                                .Row = i: .Col = col���
                                Call ShowMessage(vsDiagZY, "���ִ���������ͬ�������Ϣ��")
                                Exit Function
                            ElseIf Val(.TextMatrix(i, col���ID)) <> 0 Then
                                '����ҽ��ϴ�֤��,�����޶�Ӧ֤��ID,���ID����ͬ
'                                If Val(.TextMatrix(j, col���ID)) & "," & Val(.TextMatrix(j, col֤��ID)) _
'                                    = Val(.TextMatrix(i, col���ID)) & "," & Val(.TextMatrix(i, col֤��ID)) Then
'                                    .Row = i: .Col = col���
'                                    Call ShowMessage(vsDiagZY, "���ִ���������ͬ�������Ϣ��")
'                                    Exit Function
'                                End If
                            ElseIf Val(.TextMatrix(i, col����ID)) <> 0 Then
                                If Val(.TextMatrix(j, col����ID)) = Val(.TextMatrix(i, col����ID)) Then
                                    .Row = i: .Col = col���
                                    Call ShowMessage(vsDiagZY, "���ִ���������ͬ�������Ϣ��")
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                    blnDiagnose = True
                End If
            Next
        End With
    End If
    
    '����ҩ������
    With vsAller
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, 1)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, 1)) > 60 Then
                    .Row = i: .Col = 1
                    Call ShowMessage(vsAller, "����ҩ����̫����ֻ����60���ַ���30�����֡�")
                    Exit Function
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, 1)) <> "" Then
                        If .TextMatrix(j, 1) = .TextMatrix(i, 1) Then
                            .Row = i: .Col = 1
                            Call ShowMessage(vsAller, "���ִ���������ͬ�Ĺ���ҩ�")
                            Exit Function
                        ElseIf .RowData(i) <> 0 Then
                            If .RowData(j) = .RowData(i) Then
                                .Row = i: .Col = 1
                                Call ShowMessage(vsAller, "���ִ���������ͬ�Ĺ���ҩ�")
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End With
    
    CheckMedRec = True
End Function

Private Function SaveMedRec() As Boolean
'���ܣ�����������ҳ�ĸ�����Ϣ
    Dim arrSQL As Variant, i As Integer
    Dim curDate As Date, intIdx As Integer
    Dim str���� As String
    
    arrSQL = Array()
    curDate = zlDatabase.Currentdate
    
    If IsDate(txt����ʱ��.Text) Then
        str���� = "To_Date('" & Format(txt��������.Text & " " & txt����ʱ��.Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
    Else
        str���� = "To_Date('" & Format(txt��������.Text, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    End If
    
    '������Ϣ
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_������Ϣ_��ҳ����(" & _
        mlng����ID & "," & Val(txtEdit(txt�����).Text) & ",'" & txtEdit(txt����).Text & "'," & _
        "'" & NeedName(cboEdit(cbo�Ա�).Text) & "','" & txtEdit(txt����).Text & cboEdit(cbo����).Text & "'," & _
        str���� & ",'" & txtEdit(txt�����ص�).Text & "','" & txtEdit(txt���֤��).Text & "'," & _
        "'" & NeedName(cboEdit(cbo����).Text) & "','" & NeedName(cboEdit(cbo����).Text) & "'," & _
        "'" & NeedName(cboEdit(cbo����).Text) & "','" & NeedName(cboEdit(cboְҵ).Text) & "'," & _
        "'" & NeedName(cboEdit(cbo����).Text) & "','" & txtEdit(txt��ͥ��ַ).Text & "'," & _
        "'" & txtEdit(txt��ͥ�绰).Text & "','" & txtEdit(txt��ͥ�ʱ�).Text & "'," & _
        "'" & txtEdit(txt������λ).Text & "','" & txtEdit(txt��λ�绰).Text & "'," & _
        "'" & txtEdit(txt��λ�ʱ�).Text & "',Null,Null,Null,Null,'" & mstr�Һŵ� & "'," & _
        chkEdit(chk����).Value & ",'" & txtEdit(txt����ժҪ).Text & "')"
    
    '����ҩ��
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_���˹�����¼_Delete(" & mlng����ID & "," & mlng�Һ�ID & ",3)"
    With vsAller
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, 1)) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = _
                    "zl_���˹�����¼_Insert(" & mlng����ID & "," & mlng�Һ�ID & "," & _
                    "3,Null," & ZVal(.RowData(i)) & ",'" & .TextMatrix(i, 1) & "',1," & _
                    "To_Date('" & .Cell(flexcpData, i, 0) & "','YYYY-MM-DD HH24:MI:SS'))"
            End If
        Next
    End With
    
    '��ϼ�¼
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_Delete(" & mlng����ID & "," & mlng�Һ�ID & ",3,Null,'1')"
    With vsDiagXY
        intIdx = 0
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col���)) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3," & _
                    " Null,1," & ZVal(.TextMatrix(i, col����ID)) & "," & ZVal(.TextMatrix(i, col���ID)) & ",Null," & _
                    "'" & .TextMatrix(i, col���) & "',Null,Null," & IIF(.TextMatrix(i, col����) = "", 0, 1) & "," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null," & intIdx & ")"
            End If
        Next
    End With
    
    If mbln��ҽ Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_Delete(" & mlng����ID & "," & mlng�Һ�ID & ",3,Null,'11')"
        With vsDiagZY
            intIdx = 0
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col���)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                    arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng�Һ�ID & ",3," & _
                        "Null,11," & ZVal(.TextMatrix(i, col����ID)) & "," & ZVal(.TextMatrix(i, col���ID)) & "," & _
                        ZVal(.TextMatrix(i, col֤��ID)) & ",'" & .TextMatrix(i, col���) & "',Null,Null," & _
                        IIF(.TextMatrix(i, col����) = "", 0, 1) & ",To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null," & intIdx & ")"
                End If
            Next
        End With
    End If
    
    '�ύ����
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    mblnChange = False
    SaveMedRec = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadOldData(strOld As String)
'����:�����ݿ��б�������䰴���Ƶĸ�ʽ���ص�����
    Dim strTmp As Long
    
    If InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            txtEdit(txt����).Text = strTmp
            If cboEdit(cbo����).ListCount > 0 Then cboEdit(cbo����).ListIndex = 0
        Else
            txtEdit(txt����).Text = strOld
            cboEdit(cbo����).ListIndex = -1
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            txtEdit(txt����).Text = strTmp
            If cboEdit(cbo����).ListCount > 1 Then cboEdit(cbo����).ListIndex = 1
        Else
            txtEdit(txt����).Text = strOld
            cboEdit(cbo����).ListIndex = -1
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            txtEdit(txt����).Text = strTmp
            If cboEdit(cbo����).ListCount > 2 Then cboEdit(cbo����).ListIndex = 2
        Else
            txtEdit(txt����).Text = strOld
            cboEdit(cbo����).ListIndex = -1
        End If
    ElseIf IsNumeric(strOld) Then
        txtEdit(txt����).Text = strOld
        If cboEdit(cbo����).ListCount > 0 Then cboEdit(cbo����).ListIndex = 0
    Else
        txtEdit(txt����).Text = strOld
        cboEdit(cbo����).ListIndex = -1
    End If
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'���ܣ���ʾ��ʾ��Ϣ����λ��������Ŀ��
    Dim lngColor As Long
    
    tbsInfo.Tabs(objTmp.Container.Index + 1).Selected = True
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
    objTmp.SetFocus
    Me.Refresh
End Function

Private Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'���ܣ���ָ������װ��ָ��ComboBox
'������arrList=List String����
'      arrCboIdx=ComboBox��������,���ComboBoxʱ,װ��������ͬ
'      intDefaut=ȱʡ����
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arrCboIdx)
        cboEdit(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            cboEdit(arrCboIdx(i)).AddItem arrList(j)
        Next
        cboEdit(arrCboIdx(i)).ListIndex = intDefault 'ȱʡΪδѡ��
    Next
End Sub

Private Sub SetCboFromSQL(ByVal strSQL As String, ByVal arrCboIdx As Variant)
'���ܣ���ָ������Դ�е�����װ��ָ��������һ������ComboBox
'������strSQL=����"ID,����,����,ȱʡ��־"�ֶ�
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, j As Long
    
    '���ԭ������
    For i = 0 To UBound(arrCboIdx)
        cboEdit(arrCboIdx(i)).Clear
    Next
    
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    'װ������
    For i = 1 To rsTmp.RecordCount
        For j = 0 To UBound(arrCboIdx)
            If IsNull(rsTmp!����) Then
                cboEdit(arrCboIdx(j)).AddItem rsTmp!����
            Else
                cboEdit(arrCboIdx(j)).AddItem rsTmp!���� & "-" & Chr(13) & rsTmp!����
            End If
            cboEdit(arrCboIdx(j)).ItemData(cboEdit(arrCboIdx(j)).NewIndex) = Nvl(rsTmp!ID, 0)
            If Nvl(rsTmp!ȱʡ��־, 0) = 1 Then
                Call zlControl.CboSetIndex(cboEdit(arrCboIdx(j)).Hwnd, cboEdit(arrCboIdx(j)).NewIndex)
            End If
        Next
        rsTmp.MoveNext
    Next
    '��ȱʡʱ,Ϊδѡ��
End Sub

Private Sub cboEdit_Click(Index As Integer)
    Dim strTmp As String
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    
    If Index = cbo���� Then
        '���ݳ���������������
'        If Not mblnChange Then Exit Sub
'        If IsDate(txt��������.Text) And cboEdit(cbo����).ListIndex <> -1 Then
'            strTmp = cboEdit(cbo����).Text
'            strTmp = Switch(strTmp = "��", "yyyy", strTmp = "��", "m", strTmp = "��", "d")
'
'            txtEdit(txt����).Text = DateDiff(strTmp, txt��������.Text, zlDatabase.Currentdate)
'            If strTmp = "d" And txtEdit(txt����).Text = "0" Then txtEdit(txt����).Text = "1"
'        End If
    End If
End Sub

Private Sub cboEdit_GotFocus(Index As Integer)
    If cboEdit(Index).Style = 0 Then
        Call zlControl.TxtSelAll(cboEdit(Index))
    End If
End Sub

Private Sub cboEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cboEdit(Index).Hwnd, KeyAscii)
        If lngIdx = -1 And cboEdit(Index).ListCount > 0 Then lngIdx = 0
        cboEdit(Index).ListIndex = lngIdx
    End If
End Sub

Private Sub cboEdit_LostFocus(Index As Integer)
    Dim strTmp As String, lngTmp As Long
    Dim datTemp As Date, datBase As Date
    
    On Local Error Resume Next
    
    If Index = cbo���� Then
        If IsNumeric(txtEdit(txt����).Text) Then
            'And Between(Val(txtEdit(txt����).Text), 0, 200) Then
            'txt��������.Text = Year(zlDatabase.Currentdate) - Int(txtEdit(txt����).Text) & "-01-01"
            
            If Len(txtEdit(txt����).Text) > txtEdit(txt����).MaxLength Then Exit Sub  '��ǰ����ĳ����Ĳ���
    
            Select Case cboEdit(cbo����).Text
                Case "��"
                    If Val(txtEdit(txt����).Text) > 200 Then
                        MsgBox "����ֵ�������������Ƿ���ȷ��", vbInformation, gstrSysName
                        txtEdit(txt����).SetFocus: Exit Sub
                    End If
                Case "��"
                    If Val(txtEdit(txt����).Text) > 2400 Then
                        MsgBox "����ֵ�������������Ƿ���ȷ��", vbInformation, gstrSysName
                        txtEdit(txt����).SetFocus: Exit Sub
                    End If
                Case "��"
                    If Val(txtEdit(txt����).Text) > 73000 Then
                        MsgBox "����ֵ�������������Ƿ���ȷ��", vbInformation, gstrSysName
                        txtEdit(txt����).SetFocus: Exit Sub
                    End If
                Case Else
                    Exit Sub
            End Select
            
            If txtEdit(txt����).Text = "0" And cboEdit(cbo����).Text = "��" Then  '����һ�찴һ����
                txtEdit(txt����).Text = 1
            End If
            
            If Not IsDate(txt��������.Text) Then
                '���������������,��,������������䵹�������ͬ��,�򲻸ı��������(����ı�����ĳ�����)
                datBase = zlDatabase.Currentdate
                
                If IsDate(txt��������.Text) Then
                    If strTmp = "��" Then
                        datTemp = DateAdd("yyyy", txtEdit(txt����).Text * -1, datBase)
                        If Year(txt��������.Text) = Year(datTemp) Then Exit Sub
                    ElseIf strTmp = "��" Then
                        datTemp = DateAdd("m", txtEdit(txt����).Text * -1, datBase)
                        If Year(txt��������.Text) = Year(datTemp) And Month(txt��������.Text) = Month(datTemp) Then Exit Sub
                    End If
                End If
                
                If Val(txtEdit(txt����).Text) < 1 Then
                    strTmp = "d"
                    datTemp = DateAdd(strTmp, txtEdit(txt����).Text * 365 * -1, datBase)
                Else
                    strTmp = Switch(strTmp = "��", "yyyy", strTmp = "��", "m", strTmp = "��", "d")
                    datTemp = DateAdd(strTmp, txtEdit(txt����).Text * -1, datBase)
                End If
                txt��������.Text = Format(datTemp, "yyyy-MM-dd")    '��������ǰ������º���
            Else
                lngTmp = CalcHowOld
                If lngTmp <> -1 And lngTmp <> Val(txtEdit(txt����).Text) Then
                    If MsgBox("����ͳ������ڲ�һ�£�" & txt��������.Text & "��������Ӧ����" & lngTmp & cboEdit(cbo����).Text & "��" & _
                        vbCrLf & vbCrLf & "���������������ڵ���ȷ�ԣ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        txtEdit(txt����).SetFocus: Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub chkEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click(Index As Integer)
'˵����ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    'ʹ��Lock�ķ�ʽ,������Enabled�ķ�ʽ
    If Not cmdEdit(Index).Enabled Or txtEdit(Index).Locked Then
        txtEdit(Index).SetFocus: Exit Sub
    End If
    
    Select Case Index
        Case txt�����ص�, txt��ͥ��ַ
            'ѡ���������
            strSQL = "Select Rownum as ID,����,����,���� From ���� Order by ����"
            vPoint = GetCoordPos(txtEdit(Index).Container.Hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""����""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!����
                txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt������λ
            'ѡ��λ��Ϣ
            strSQL = "Select ID,�ϼ�ID,ĩ��,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ��" & _
                " From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            vPoint = GetCoordPos(txtEdit(Index).Container.Hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "��Լ��λ", , , , , , True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""��Լ��λ""���ݣ����ȵ���Լ��λ���������á�", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!���� & IIF(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                If txtEdit(txt��λ�绰).Text = "" Then
                    txtEdit(txt��λ�绰).Text = Nvl(rsTmp!�绰)
                End If
                txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
    End Select
End Sub

Private Sub cmdMakeLog_Click()
    Dim strLog As String, i As Long
    
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col���) <> "" Then
                strLog = strLog & "��" & .TextMatrix(i, col���) & IIF(.TextMatrix(i, col����) <> "", "(��)", "")
            End If
        Next
    End With
    With vsDiagZY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col���) <> "" Then
                strLog = strLog & "��" & .TextMatrix(i, col���) & IIF(.TextMatrix(i, col����) <> "", "(��)", "")
            End If
        Next
    End With
    If strLog <> "" Then
        txtEdit(txt����ժҪ).Text = Mid(strLog, 2)
    End If
    txtEdit(txt����ժҪ).SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim blnDiagnose As Boolean
    
    If Not CheckMedRec(blnDiagnose) Then Exit Sub
    If mblnDiagnose And Not blnDiagnose Then
        If MsgBox("���˵������Ϣ��û�����룬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    If Not SaveMedRec Then Exit Sub
        
    mblnDiagnose = blnDiagnose
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnDiagnose Then
        On Error Resume Next
        vsDiagXY.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdMakeLog_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
    mstrLike = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    mint���� = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", 0)) '����ƥ�䷽ʽ��0-ƴ��,1-���
    optInput(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����������", 0))).Value = True
    
    '���������Դ
    If gint�����Դ > 1 Then
        optInput(0).Enabled = False
        optInput(1).Enabled = False
        If gint�����Դ = 2 Then
            optInput(0).Value = True
        ElseIf gint�����Դ = 3 Then
            optInput(1).Value = True
        End If
    End If
    
    If Not InitMedData Then Unload Me: Exit Sub
    If Not LoadMedRec Then Unload Me: Exit Sub
    
    tbsInfo.Tabs(2).Selected = True
    Call tbsInfo_Click
    
    mblnChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("����رմ��壬�������ĸ��Ľ����ᱣ�档Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����������", IIF(optInput(0).Value, 0, 1)
End Sub

Private Sub optInput_LostFocus(Index As Integer)
    optInput(0).TabStop = False: optInput(1).TabStop = False 'Ҫǿ�д���ִ��һ��
End Sub

Private Sub tbsInfo_Click()
    Dim i As Integer
    
    For i = 0 To fraInfo.UBound
        If i = tbsInfo.SelectedItem.Index - 1 Then
            fraInfo(i).Visible = True
            fraInfo(i).ZOrder
        Else
            fraInfo(i).Visible = False
        End If
    Next
End Sub

Private Sub tbsInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtEdit(Index))
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = txt�����ص� Or Index = txt��ͥ��ַ) And txtEdit(Index).Text <> "" Then
            '�����������
            strSQL = "Select Rownum as ID,����,����,���� From ���� " & _
                " Where (Upper(����) Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtEdit(Index).Container.Hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", mstrLike & UCase(txtEdit(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!����
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt������λ And txtEdit(Index).Text <> "" Then
            '���빤����λ
            strSQL = "Select ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " And (Upper(����) Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtEdit(Index).Container.Hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������λ", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", mstrLike & UCase(txtEdit(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!���� & IIF(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                If txtEdit(txt��λ�绰).Text = "" Then
                    txtEdit(txt��λ�绰).Text = Nvl(rsTmp!�绰)
                End If
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("*") Then
        'ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
        KeyAscii = 0
        Call cmdEdit_Click(Index)
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        '�ǿ��ư���
        
        '�������볤��
        If txtEdit(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtEdit(Index).Text) > txtEdit(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '������������
        Select Case Index
            Case txt����
                strMask = "1234567890"
            'Case txt��������
                'strMask = "1234567890-"
            Case txt���֤��
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Case txt��ͥ�绰, txt��λ�绰
                strMask = "1234567890-()"
            Case txt��ͥ�ʱ�, txt��λ�ʱ�
                strMask = "1234567890"
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txt��������_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt��������_GotFocus()
    Call zlControl.TxtSelAll(txt��������)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt��������_Validate(Cancel As Boolean)
    Dim datBase As Date, lngTmp As Long
    
    On Local Error Resume Next
    
    If IsDate(txt��������.Text) Then
        'If txtEdit(txt����).Text = "" Then
            'strTmp = Get����ֵ(CDate(txt��������.Text))
            'If strTmp <> "" Then txtEdit(txt����).Text = strTmp
            
            datBase = zlDatabase.Currentdate
            lngTmp = Val(Format(datBase, "yyyy")) - Val(Format(CDate(txt��������.Text), "yyyy"))
            
            If lngTmp > 1 Then '2������
                'δ������
                If Format(datBase, "MMdd") < Format(txt��������.Text, "MMdd") Then
                    lngTmp = lngTmp - 1
                End If
                txtEdit(txt����).Text = lngTmp
                cboEdit(cbo����).ListIndex = 0
            Else
                '2�����°��¼�
                lngTmp = Val(Format(datBase, "MM")) - Val(Format(CDate(txt��������.Text), "MM")) + IIF(lngTmp = 1, 12, 0)
                
                If lngTmp > 1 Then '��
                   txtEdit(txt����).Text = lngTmp
                   cboEdit(cbo����).ListIndex = 1
                Else
                    '2�����°����
                    lngTmp = Val(datBase - CDate(txt��������.Text))
                    txtEdit(txt����).Text = IIF(lngTmp = 0, 1, lngTmp)   '����һ����һ��
                    cboEdit(cbo����).ListIndex = 2
                End If
            End If
        'End If
    Else
        txt��������.Text = "____-__-__"
        txt����ʱ��.Text = "__:__"
        Cancel = True
    End If
End Sub

Private Sub txt����ʱ��_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt����ʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ʱ��)
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not IsDate(txt��������.Text) Then
        KeyAscii = 0: txt����ʱ��.Text = "__:__"
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.Text = "__:__": Cancel = True
    End If
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsAller
        If Col = 1 Then
            If .TextMatrix(Row, Col) <> .Cell(flexcpData, Row, Col) Then
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            End If
            Call vsAller_AfterRowColChange(-1, -1, Row, Col)
        End If
    End With
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAller
        If NewCol = 1 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsAller_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 Then Cancel = True
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int�Ա� As Integer
    
    With vsAller
        If cboEdit(cbo�Ա�).Text Like "*��*" Then
            int�Ա� = 1
        ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
            int�Ա� = 2
        End If
        
        strSQL = _
            " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
            " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
            " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
            " Select ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
            " NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ��" & _
            " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3)" & _
            " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
            " Union All" & _
            " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����,A.����," & _
            " A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
            " From ������ĿĿ¼ A,ҩƷ���� B" & _
            " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
            IIF(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[1])", "") & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����ҩ��", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int�Ա�)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û��ҩƷ���ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call SetAllerInput(Row, rsTmp)
            Call AllerEnterNextCell
        End If
    End With
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsAller
        If KeyCode = vbKeyF4 Then
            If .Col = 1 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 1) <> "" Then
                If MsgBox("ȷʵҪ������й���ҩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsAller_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    With vsAller
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = 1 Then
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
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim StrInput As String, vPoint As POINTAPI
    Dim int�Ա�  As Integer
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        With vsAller
            If Col = 1 And .EditText <> "" Then
                StrInput = UCase(.EditText)
                If cboEdit(cbo�Ա�).Text Like "*��*" Then
                    int�Ա� = 1
                ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
                    int�Ա� = 2
                End If
                strSQL = _
                    " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ," & _
                    " B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                    " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
                    " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.ID=C.������ĿID" & _
                    " And (A.���� Like [1] Or A.���� Like [2] Or C.���� Like [2] Or C.���� Like [2])" & _
                    IIF(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[3])", "") & _
                    Decode(mint����, 0, " And C.����=[4]", 1, " And C.����=[4]", "") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " Order by A.����"
                
                vPoint = GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҩ��", False, "", "", False, _
                    False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    StrInput & "%", mstrLike & StrInput & "%", int�Ա�, mint���� + 1)
                If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    Call vsAller_AfterRowColChange(Row, Col, Row, Col)
                    .SetFocus: Exit Sub
                Else
                    Call SetAllerInput(Row, rsTmp)
                    .EditText = .TextMatrix(Row, Col)
                End If
                Call AllerEnterNextCell
            End If
        End With
    End If
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAller.EditSelStart = 0
    vsAller.EditSelLength = zlCommFun.ActualLen(vsAller.EditText)
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagXY
        If Col = col��� Then
            If .TextMatrix(Row, Col) <> .Cell(flexcpData, Row, Col) Then
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            End If
            Call vsDiagXY_AfterRowColChange(-1, -1, Row, Col)
        End If
    End With
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDiagXY
        If Not DiagCellEditable(vsDiagXY, NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            If NewCol = col��� Then
                .ComboList = "..."
            Else
                .ComboList = ""
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsDiagZY.ColWidth(Col) = vsDiagXY.ColWidth(Col)
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str�Ա� As String
    
    With vsDiagXY
        If optInput(0).Value Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            strSQL = _
                " Select 0 As ĩ��,NULL||ID As ID,�ϼ�ID," & _
                " -NULL as ��ĿID,����,����,Null As ˵��,Null As ����" & _
                " From ������Ϸ��� Where ���=1" & _
                " Start With �ϼ�ID Is Null Connect By Prior ID=�ϼ�ID" & _
                " Union All" & _
                " Select 1 As ĩ��,A.ID||'0'||B.����ID as ID,B.����ID As �ϼ�ID," & _
                " A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                " From �������Ŀ¼ A,����������� B" & _
                " Where A.ID=B.���ID And A.���=1"
        Else
            If cboEdit(cbo�Ա�).Text Like "*��*" Then
                str�Ա� = "��"
            ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
                str�Ա� = "Ů"
            End If
            'D-ICD-10��������
            strSQL = _
                " Select 0 as ĩ��,ID,�ϼ�ID,-NULL as ��ĿID,���||LPAD(���,3,'0') as ����," & _
                " NULL as ����,����,����,NULL as ˵�� From �����������" & _
                " Where ���='D' Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                " Union ALL " & _
                " Select 1 as ĩ��,ID,����ID as �ϼ�ID,ID as ��ĿID,����,����,����,����,˵��" & _
                " From ��������Ŀ¼ Where ���='D'" & _
                IIF(str�Ա� <> "", " And (�Ա�����=[1] Or �Ա����� is NULL)", "")
        End If
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, IIF(optInput(0).Value, "�������", "��������"), _
            False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, str�Ա�)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û��" & IIF(optInput(0).Value, "�������", "��������") & "���ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call XYSetDiagInput(Row, rsTmp)
            Call DiagEnterNextCell(vsDiagXY)
        End If
    End With
End Sub

Private Sub vsDiagXY_DblClick()
    Call vsDiagXY_KeyPress(32)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsDiagXY
        If KeyCode = vbKeyF4 Then
            If .Col = col��� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col���) <> "" Then
                If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDiagXY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    With vsDiagXY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell(vsDiagXY)
        ElseIf KeyAscii = 32 And (.Col = col����) Then
            If DiagCellEditable(vsDiagXY, .Row, .Col) Then
                KeyAscii = 0
                If .Col = col���� Then
                    .TextMatrix(.Row, .Col) = IIF(.TextMatrix(.Row, .Col) = "", "��", "")
                End If
            End If
        Else
            If .Col = col��� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagXY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str�Ա� As String, StrInput As String
    Dim vPoint As POINTAPI
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        With vsDiagXY
            If Col = col��� Then
                If .EditText <> "" Then
                    StrInput = UCase(.EditText)
                    If optInput(0).Value Then
                        '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                        If zlCommFun.IsCharChinese(StrInput) Then
                            strSQL = "B.���� Like [2]" '���뺺��ʱ,ֻƥ������
                        Else
                            strSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                        End If
                        strSQL = _
                            " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                            " From �������Ŀ¼ A,������ϱ��� B" & _
                            " Where A.ID=B.���ID And A.���=1" & _
                            Decode(mint����, 0, " And B.����=[4]", 1, " And B.����=[4]", "") & _
                            " And (" & strSQL & ")" & _
                            " Order by A.����"
                    Else
                        If cboEdit(cbo�Ա�).Text Like "*��*" Then
                            str�Ա� = "��"
                        ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
                            str�Ա� = "Ů"
                        End If
                        'D-ICD-10��������
                        If zlCommFun.IsCharChinese(StrInput) Then
                            strSQL = "���� Like [2]" '���뺺��ʱ,ֻƥ������
                        Else
                            strSQL = "���� Like [1] Or ���� Like [2] Or ���� Like [2]"
                        End If
                        strSQL = _
                            " Select ID,ID as ��ĿID,����,����,����,����,˵��" & _
                            " From ��������Ŀ¼ Where ���='D'" & _
                            IIF(str�Ա� <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
                            " And (" & strSQL & ")" & _
                            " Order by ����"
                    End If
                    vPoint = GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(optInput(0).Value, "�������", "��������"), _
                        False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                        StrInput & "%", mstrLike & StrInput & "%", str�Ա�, mint���� + 1)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                        Call vsDiagXY_AfterRowColChange(Row, Col, Row, Col)
                        .SetFocus: Exit Sub
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And (gint������� = 2 Or gint������� = 3 And mint���� <> 0) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                            Call vsDiagXY_AfterRowColChange(Row, Col, Row, Col)
                            .SetFocus: Exit Sub
                        End If
                    
                        Call XYSetDiagInput(Row, rsTmp)
                        .EditText = .TextMatrix(Row, Col)
                    End If
                    Call DiagEnterNextCell(vsDiagXY)
                End If
            End If
        End With
    End If
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagXY.EditSelStart = 0
    vsDiagXY.EditSelLength = zlCommFun.ActualLen(vsDiagXY.EditText)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not DiagCellEditable(vsDiagXY, Row, Col) Then
        Cancel = True
    ElseIf Col = col���� Then
        Cancel = True '��ֱ�ӱ༭
    End If
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagZY
        If Col = col��� Then
            If .TextMatrix(Row, Col) <> .Cell(flexcpData, Row, Col) Then
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            End If
            Call vsDiagZY_AfterRowColChange(-1, -1, Row, Col)
        End If
    End With
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDiagZY
        If Not DiagCellEditable(vsDiagZY, NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            If NewCol = col��� Then
                .ComboList = "..."
            Else
                .ComboList = ""
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsDiagXY.ColWidth(Col) = vsDiagZY.ColWidth(Col)
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str�Ա� As String
    
    With vsDiagZY
        If optInput(0).Value Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            strSQL = _
                " Select 0 As ĩ��,NULL||ID As ID,�ϼ�ID," & _
                " -NULL as ��ĿID,����,����,Null As ˵��,Null As ����" & _
                " From ������Ϸ��� Where ���=2" & _
                " Start With �ϼ�ID Is Null Connect By Prior ID=�ϼ�ID" & _
                " Union All" & _
                " Select 1 As ĩ��,A.ID||'0'||B.����ID as ID,B.����ID As �ϼ�ID," & _
                " A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                " From �������Ŀ¼ A,����������� B" & _
                " Where A.ID=B.���ID And A.���=2"
        Else
            If cboEdit(cbo�Ա�).Text Like "*��*" Then
                str�Ա� = "��"
            ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
                str�Ա� = "Ů"
            End If
            'B-��ҽ��������
            strSQL = _
                " Select 0 as ĩ��,ID,�ϼ�ID,-NULL as ��ĿID,���||LPAD(���,3,'0') as ����," & _
                " NULL as ����,����,����,NULL as ˵�� From �����������" & _
                " Where ���='B'" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                " Union ALL " & _
                " Select 1 as ĩ��,ID,����ID as �ϼ�ID,ID as ��ĿID,����,����,����,����,˵��" & _
                " From ��������Ŀ¼ Where ���='B'" & _
                IIF(str�Ա� <> "", " And (�Ա�����=[1] Or �Ա����� is NULL)", "")
        End If
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, IIF(optInput(0).Value, "�������", "��������"), _
            False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, str�Ա�)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û��" & IIF(optInput(0).Value, "�������", "��������") & "���ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call ZYSetDiagInput(Row, rsTmp)
            Call DiagEnterNextCell(vsDiagZY)
        End If
    End With
End Sub

Private Sub vsDiagZY_DblClick()
    Call vsDiagZY_KeyPress(32)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsDiagZY
        If KeyCode = vbKeyF4 Then
            If .Col = col��� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col���) <> "" Then
                If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDiagZY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    With vsDiagZY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell(vsDiagZY)
        ElseIf KeyAscii = 32 And (.Col = col����) Then
            If DiagCellEditable(vsDiagZY, .Row, .Col) Then
                KeyAscii = 0
                If .Col = col���� Then
                    .TextMatrix(.Row, .Col) = IIF(.TextMatrix(.Row, .Col) = "", "��", "")
                End If
            End If
        Else
            If .Col = col��� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagZY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim StrInput As String, vPoint As POINTAPI
    Dim str�Ա� As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        With vsDiagZY
            If Col = col��� Then
                If .EditText <> "" Then
                    StrInput = UCase(.EditText)
                    If optInput(0).Value Then
                        '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                        If zlCommFun.IsCharChinese(StrInput) Then
                            strSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
                        Else
                            strSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                        End If
                        strSQL = _
                            " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                            " From �������Ŀ¼ A,������ϱ��� B" & _
                            " Where A.ID=B.���ID And A.���=2" & _
                            Decode(mint����, 0, " And B.����=[4]", 1, " And B.����=[4]", "") & _
                            " And (" & strSQL & ")" & _
                            " Order by A.����"
                    Else
                        If cboEdit(cbo�Ա�).Text Like "*��*" Then
                            str�Ա� = "��"
                        ElseIf cboEdit(cbo�Ա�).Text Like "*Ů*" Then
                            str�Ա� = "Ů"
                        End If
                        'B-��ҽ��������
                        If zlCommFun.IsCharChinese(StrInput) Then
                            strSQL = "���� Like [2]" '���뺺��ʱֻƥ������
                        Else
                            strSQL = "���� Like [1] Or ���� Like [2] Or ���� Like [2]"
                        End If
                        strSQL = _
                            " Select ID,ID as ��ĿID,����,����,����,����,˵��" & _
                            " From ��������Ŀ¼" & _
                            " Where ���='B'" & _
                            IIF(str�Ա� <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
                            " And (" & strSQL & ")" & _
                            " Order by ����"
                    End If
                    vPoint = GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(optInput(0).Value, "�������", "��������"), False, "", "", False, False, True, _
                        vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%", str�Ա�, mint���� + 1)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                        Call vsDiagZY_AfterRowColChange(Row, Col, Row, Col)
                        .SetFocus: Exit Sub
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And (gint������� = 2 Or gint������� = 3 And mint���� <> 0) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                            Call vsDiagZY_AfterRowColChange(Row, Col, Row, Col)
                            .SetFocus: Exit Sub
                        End If
                    
                        Call ZYSetDiagInput(Row, rsTmp)
                        .EditText = .TextMatrix(Row, Col)
                    End If
                    Call DiagEnterNextCell(vsDiagZY)
                End If
            End If
        End With
    End If
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagZY.EditSelStart = 0
    vsDiagZY.EditSelLength = zlCommFun.ActualLen(vsDiagZY.EditText)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not DiagCellEditable(vsDiagZY, Row, Col) Then
        Cancel = True
    ElseIf Col = col���� Then
        Cancel = True '��ֱ�ӱ༭
    End If
End Sub

Private Function DiagCellEditable(objGrid As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With objGrid
        '�������������
        If .TextMatrix(lngRow, col���) = "" Then
            If lngCol = col���� Then
                Exit Function
            End If
        End If
    End With
    DiagCellEditable = True
End Function

Private Sub AllerEnterNextCell()
    Dim i As Long, j As Long
    
    With vsAller
        If .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub DiagEnterNextCell(objGrid As VSFlexGrid)
    Dim i As Long, j As Long
    
    With objGrid
        '����һ��Ԫ��ʼѭ������
        For i = .Row To .Rows - 1
            For j = IIF(i = .Row, .Col + 1, col���) To col����
                If DiagCellEditable(objGrid, i, j) Then Exit For
            Next
            If j <= col���� Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub SetAllerInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ��������ҩ�������
    Dim strSQL As String, curDate As Date
    
    With vsAller
        If Not rsInput Is Nothing Then
            .RowData(lngRow) = CLng(rsInput!ID)
            .TextMatrix(lngRow, 1) = Nvl(rsInput!����)
        Else
            .RowData(lngRow) = 0
            .TextMatrix(lngRow, 1) = .EditText
        End If
        .Cell(flexcpData, lngRow, 1) = .TextMatrix(lngRow, 1)
        
        curDate = zlDatabase.Currentdate
        .TextMatrix(lngRow, 0) = Format(curDate, "yyyy-MM-dd HH:mm")
        .Cell(flexcpData, lngRow, 0) = Format(curDate, "yyyy-MM-dd HH:mm:ss")
        
        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
        mblnChange = True
    End With
End Sub

Private Sub XYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    With vsDiagXY
        If Not rsInput Is Nothing Then
            .TextMatrix(lngRow, col���) = IIF(Not IsNull(rsInput!����), "(" & rsInput!���� & ")", "") & Nvl(rsInput!����)
            .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
            '�������ȷ������,����ݼ���ȷ�����
            If optInput(0).Value Then
                .TextMatrix(lngRow, col���ID) = rsInput!��ĿID
                .TextMatrix(lngRow, col����ID) = ""
                strSQL = "Select ����ID as ID From ������϶��� Where ���ID=[1]"
            Else
                .TextMatrix(lngRow, col����ID) = rsInput!��ĿID
                .TextMatrix(lngRow, col���ID) = ""
                strSQL = "Select ���ID as ID From ������϶��� Where ����ID=[1]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!��ĿID))
            If Not rsTmp.EOF Then
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col����ID) = Nvl(rsTmp!ID)
                Else
                    .TextMatrix(lngRow, col���ID) = Nvl(rsTmp!ID)
                End If
            End If
        Else
            .TextMatrix(lngRow, col���) = .EditText
            .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
            .TextMatrix(lngRow, col���ID) = ""
            .TextMatrix(lngRow, col����ID) = ""
        End If
        
        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, col����) = "��ҽ"
            .Cell(flexcpForeColor, .FixedRows, col����, .Rows - 1, col����) = vbRed
        End If
        mblnChange = True
    End With
End Sub

Private Sub ZYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, str���� As String
    
    With vsDiagZY
        If Not rsInput Is Nothing Then
            If Not IsNull(rsInput!����) Then
                str���� = "(" & rsInput!���� & ")"
            End If
            .TextMatrix(lngRow, col���) = Nvl(rsInput!����)
            
            '�������ȷ������,����ݼ���ȷ�����
            If optInput(0).Value Then
                .TextMatrix(lngRow, col���ID) = rsInput!��ĿID
                .TextMatrix(lngRow, col����ID) = ""
                strSQL = "Select ����ID as ID From ������϶��� Where ���ID=[1]"
            Else
                .TextMatrix(lngRow, col����ID) = rsInput!��ĿID
                .TextMatrix(lngRow, col���ID) = ""
                strSQL = "Select ���ID as ID From ������϶��� Where ����ID=[1]"
            End If
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!��ĿID))
            If Not rsTmp.EOF Then
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col����ID) = Nvl(rsTmp!ID)
                Else
                    .TextMatrix(lngRow, col���ID) = Nvl(rsTmp!ID)
                End If
            End If
            
            '��ҽ���ݼ�����ϲο�ȡ֤��
            If Val(.TextMatrix(lngRow, col���ID)) <> 0 Then
                strSQL = "Select Distinct ֤����� as ID,֤��ID,֤������" & _
                    " From ������ϲο�" & _
                    " Where ���ID=[1] And ֤������ is Not NULL" & _
                    " Order by ֤�����"
                vPoint = GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = Nothing
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                    vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, Val(Val(.TextMatrix(lngRow, col���ID))))
                If Not rsTmp Is Nothing Then
                    .TextMatrix(lngRow, col֤��ID) = Nvl(rsTmp!֤��ID)
                    .TextMatrix(lngRow, col���) = Nvl(rsTmp!֤������) & .TextMatrix(lngRow, col���)
                End If
            End If
            .TextMatrix(lngRow, col���) = str���� & .TextMatrix(lngRow, col���)
            .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
        Else
            .TextMatrix(lngRow, col���) = .EditText
            .Cell(flexcpData, lngRow, col���) = .TextMatrix(lngRow, col���)
            .TextMatrix(lngRow, col���ID) = ""
            .TextMatrix(lngRow, col����ID) = ""
            .TextMatrix(lngRow, col֤��ID) = ""
        End If
        
        '����ǳ�Ժ���,ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, col����) = "��ҽ"
            .Cell(flexcpForeColor, .FixedRows, col����, .Rows - 1, col����) = vbRed
        End If
        mblnChange = True
    End With
End Sub
