VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmArchiveOutMedRec 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "������ҳ"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   7905
   Icon            =   "frmArchiveOutMedRec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   8235
      Left            =   615
      TabIndex        =   25
      Top             =   150
      Width           =   6570
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   20
         Left            =   4635
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   24
         Top             =   7875
         Width           =   1725
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   19
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   540
         Width           =   1140
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   18
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   9
         Top             =   1410
         Width           =   2310
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   17
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   8
         Top             =   1410
         Width           =   2295
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   16
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   7
         Top             =   1050
         Width           =   2310
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   15
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   6
         Top             =   1050
         Width           =   2295
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   14
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   3
         Top             =   540
         Width           =   1575
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   13
         Left            =   3090
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   1
         Top             =   180
         Width           =   1230
      End
      Begin VB.CheckBox chkEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����(&R)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   22
         Top             =   7890
         Width           =   930
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   12
         Left            =   315
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   7290
         Width           =   6045
      End
      Begin VB.CheckBox chkEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��Ⱦ���ϴ�(&U)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   23
         Top             =   7890
         Width           =   1470
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   18
         Top             =   4080
         Width           =   3030
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2280
         Width           =   5400
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   4965
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   17
         Top             =   3720
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   16
         Top             =   3720
         Width           =   3030
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3360
         Width           =   5400
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   4965
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   14
         Top             =   3000
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   13
         Top             =   3000
         Width           =   3030
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   6
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   12
         Top             =   2640
         Width           =   5400
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   10
         Top             =   1920
         Width           =   5400
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   3090
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   540
         Width           =   1230
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   0
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1140
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiag 
         Height          =   915
         Left            =   135
         TabIndex        =   20
         Top             =   5940
         Width           =   6225
         _cx             =   10980
         _cy             =   1614
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
         BackColorFixed  =   -2147483643
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmArchiveOutMedRec.frx":000C
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsAller 
         Height          =   915
         Left            =   135
         TabIndex        =   19
         Top             =   4665
         Width           =   6225
         _cx             =   10980
         _cy             =   1614
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
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   8421504
         GridColorFixed  =   8421504
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmArchiveOutMedRec.frx":0072
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
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   20
         X1              =   4560
         X2              =   6365
         Y1              =   8070
         Y2              =   8070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   19
         X1              =   135
         X2              =   6365
         Y1              =   7785
         Y2              =   7785
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   18
         X1              =   4890
         X2              =   6360
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   17
         X1              =   4890
         X2              =   6360
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   16
         X1              =   885
         X2              =   3990
         Y1              =   4275
         Y2              =   4275
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   15
         X1              =   885
         X2              =   3990
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   14
         X1              =   885
         X2              =   3990
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   13
         X1              =   885
         X2              =   6360
         Y1              =   3555
         Y2              =   3555
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   12
         X1              =   885
         X2              =   6360
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   11
         X1              =   885
         X2              =   6360
         Y1              =   2475
         Y2              =   2475
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   10
         X1              =   885
         X2              =   6360
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   3990
         X2              =   6360
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   3990
         X2              =   6360
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   885
         X2              =   3255
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   885
         X2              =   3255
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   5160
         X2              =   6360
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   5160
         X2              =   6360
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   3030
         X2              =   4315
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   3030
         X2              =   4315
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   885
         X2              =   2550
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   885
         X2              =   2550
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ժҪ��"
         Height          =   180
         Index           =   20
         Left            =   120
         TabIndex        =   48
         Top             =   7005
         Width           =   900
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������¼��"
         Height          =   180
         Index           =   18
         Left            =   120
         TabIndex        =   47
         Top             =   4455
         Width           =   900
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϼ�¼��"
         Height          =   180
         Index           =   19
         Left            =   120
         TabIndex        =   46
         Top             =   5730
         Width           =   900
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   21
         Left            =   3810
         TabIndex        =   45
         Top             =   7890
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�໤��"
         Height          =   180
         Index           =   22
         Left            =   300
         TabIndex        =   44
         Top             =   4095
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   555
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�ʱ�"
         Height          =   180
         Index           =   17
         Left            =   4095
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   195
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmArchiveOutMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mlng�Һ�ID As Long
Private mblnMoved As Boolean
Private mblnCheck As Boolean

Private Enum TXT_ENUM
    txt���� = 0
    txt�Ա� = 13
    txt���� = 3
    txt���� = 15
    txt���� = 16
    txt���� = 17
    txtְҵ = 18
    txt����� = 1
    txt�໤�� = 2
    txt�������� = 14
    txt���֤�� = 4
    txt�����ص� = 5
    txt������λ = 6
    txt��λ�绰 = 7
    txt��λ�ʱ� = 8
    txt��ͥ��ַ = 9
    txt��ͥ�绰 = 10
    txt��ͥ�ʱ� = 11
    txt����ժҪ = 12
    txt���ʽ = 19
    txt����ʱ�� = 20
End Enum
Private Enum CHK_ENUM
    chk���� = 0
    chk��Ⱦ���ϴ� = 1
End Enum
Private Enum COL_ENUM
    col���� = 0
    col��� = 1
    col���� = 2
End Enum

Public Function zlRefresh(ByVal lng����ID As Long, ByVal lng�Һ�id As Long, ByVal blnMoved As Boolean) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
    mlng����ID = lng����ID: mlng�Һ�ID = lng�Һ�id: mblnMoved = blnMoved
    zlRefresh = LoadMedRec
End Function

Private Function LoadMedRec() As Boolean
'���ܣ���ȡ������ҳ�ĸ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngRow As Long, bln��ҽ As Boolean
    
    mblnCheck = True
    
    On Error GoTo errH
    
    '������Ϣ
    strSQL = "Select B.ִ�в���ID as ����ID,B.ժҪ,B.����," & _
        " B.��Ⱦ���ϴ�,B.����ʱ��,A.����,A.�����,A.����,A.�Ա�,A.����,A.��������,A.ҽ�Ƹ��ʽ," & _
        " A.����,A.����,A.����״��,A.ְҵ,A.���֤��,A.�����ص�,A.�໤��,A.��ͥ��ַ,A.��ͥ�绰," & _
        " A.��ͥ��ַ�ʱ�,A.������λ,A.��λ�绰,A.��λ�ʱ�" & _
        " From ������Ϣ A,���˹Һż�¼ B Where A.����ID=B.����ID And B.ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�Һ�ID)
    If rsTmp.EOF Then Exit Function
    
    bln��ҽ = Have��������(rsTmp!����ID, "��ҽ��")
        
    txtEdit(txt����).Text = NVL(rsTmp!����)
    txtEdit(txt�Ա�).Text = NVL(rsTmp!�Ա�)
    txtEdit(txt����).Text = NVL(rsTmp!����)
    txtEdit(txt�����).Text = NVL(rsTmp!�����)
    
    txtEdit(txt��������).Text = Format(rsTmp!��������, "yyyy-MM-dd")
    If Format(rsTmp!��������, "HH:mm") <> "00:00" Then
        txtEdit(txt��������).Text = Format(rsTmp!��������, "yyyy-MM-dd HH:mm")
    End If
    
    txtEdit(txt���ʽ) = NVL(rsTmp!ҽ�Ƹ��ʽ)
    txtEdit(txt����) = NVL(rsTmp!����)
    txtEdit(txt����) = NVL(rsTmp!����)
    txtEdit(txt����) = NVL(rsTmp!����״��)
    txtEdit(txtְҵ) = NVL(rsTmp!ְҵ)
    txtEdit(txt�໤��).Text = NVL(rsTmp!�໤��)
    txtEdit(txt���֤��).Text = NVL(rsTmp!���֤��)
    txtEdit(txt�����ص�).Text = NVL(rsTmp!�����ص�)
    txtEdit(txt������λ).Text = NVL(rsTmp!������λ)
    txtEdit(txt��λ�绰).Text = NVL(rsTmp!��λ�绰)
    txtEdit(txt��λ�ʱ�).Text = NVL(rsTmp!��λ�ʱ�)
    txtEdit(txt��ͥ��ַ).Text = NVL(rsTmp!��ͥ��ַ)
    txtEdit(txt��ͥ�绰).Text = NVL(rsTmp!��ͥ�绰)
    txtEdit(txt��ͥ�ʱ�).Text = NVL(rsTmp!��ͥ��ַ�ʱ�)
    txtEdit(txt����ժҪ).Text = NVL(rsTmp!ժҪ)
    chkEdit(chk����).Value = NVL(rsTmp!����, 0)
    chkEdit(chk��Ⱦ���ϴ�).Value = NVL(rsTmp!��Ⱦ���ϴ�, 0)

    txtEdit(txt����ʱ��).Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd")
    If Format(rsTmp!����ʱ��, "HH:mm") <> "00:00" Then
        txtEdit(txt����ʱ��).Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
    End If
    
    '������Ϣ:���ιҺŵ�,������
    strSQL = "Select ��¼��Դ,Decode(����ʱ��,Null ,��¼ʱ��,����ʱ��) as ����ʱ��,ҩ��ID,ҩ���� From ���˹�����¼ A" & _
        " Where ���=1 And ����ID=[1] And ��ҳID=[2]" & _
        " And Not Exists(Select ҩ��ID From ���˹�����¼" & _
            " Where (Nvl(ҩ��ID,0)=Nvl(A.ҩ��ID,0) Or Nvl(ҩ����,'Null')=Nvl(A.ҩ����,'Null'))" & _
            " And Nvl(���,0)=0 And ��¼ʱ��>=A.��¼ʱ�� And ����ID=[1] And ��ҳID=[2])" & _
        " Order by ����ʱ��,ҩ����"
    If mblnMoved Then
        strSQL = Replace(strSQL, "���˹�����¼", "H���˹�����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    
    vsAller.Rows = vsAller.FixedRows
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��¼��Դ=3" '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
        With vsAller
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                '������Դ�Ŀ������ظ�
                lngRow = -1
                If Not IsNull(rsTmp!ҩ��ID) Then
                    lngRow = .FindRow(CLng(rsTmp!ҩ��ID))
                ElseIf Not IsNull(rsTmp!ҩ����) Then
                    lngRow = .FindRow(CStr(rsTmp!ҩ����), , 1)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(NVL(rsTmp!ҩ��ID, 0))
                    .TextMatrix(i, 0) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                    .TextMatrix(i, 1) = NVL(rsTmp!ҩ����)
                End If
                rsTmp.MoveNext
            Next
            .Row = 1: .Col = 1
        End With
    End If
    
    '�����Ϣ:���ιҺŵ�
    strSQL = "Select ��¼��Դ,�������,����ID,���ID,֤��ID,�������,�Ƿ����� From ������ϼ�¼" & _
        " Where ��¼��Դ IN(1,3) And ������� IN(1,11)" & _
        " And ȡ��ʱ�� is Null And ����ID=[1] And ��ҳID=[2]" & _
        " Order by �������,��ϴ���"
    If mblnMoved Then
        strSQL = Replace(strSQL, "������ϼ�¼", "H������ϼ�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng�Һ�ID)
    
    vsDiag.Rows = vsDiag.FixedRows
    If Not rsTmp.EOF Then
        With vsDiag
            '��ҽ���
            rsTmp.Filter = "�������=1 And ��¼��Դ=3" '��ҳ������д��
            If rsTmp.EOF Then rsTmp.Filter = "�������=1 And ��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
            .Rows = .Rows + rsTmp.RecordCount
            Do While Not rsTmp.EOF
                .TextMatrix(.Rows - 1, col����) = "��ҽ"
                .TextMatrix(.Rows - 1, col���) = NVL(rsTmp!�������)
                .TextMatrix(.Rows - 1, col����) = IIf(NVL(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                rsTmp.MoveNext
            Loop
            
            '��ҽ���
            rsTmp.Filter = "�������=11 And ��¼��Դ=3"
            If rsTmp.EOF Then rsTmp.Filter = "�������=11 And ��¼��Դ<>3"
            If rsTmp.EOF Then .ColHidden(col����) = True
            .Rows = .Rows + rsTmp.RecordCount
            Do While Not rsTmp.EOF
                .TextMatrix(.Rows - 1, col����) = "��ҽ"
                .TextMatrix(.Rows - 1, col���) = NVL(rsTmp!�������)
                .TextMatrix(.Rows - 1, col����) = IIf(NVL(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                rsTmp.MoveNext
            Loop
            
            .Cell(flexcpForeColor, .FixedRows, col����, .Rows - 1, col����) = vbRed
            .Row = .FixedRows: .Col = col���
        End With
    End If
    
    mblnCheck = False
    LoadMedRec = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub chkEdit_Click(Index As Integer)
    If Not mblnCheck Then
        mblnCheck = True
        chkEdit(Index).Value = IIf(chkEdit(Index).Value = 1, 0, 1)
        mblnCheck = False
    End If
End Sub

Private Sub Form_Load()
    Me.BackColor = fraBack.BackColor
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraBack.Top = 0
    fraBack.Left = 0
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtEdit(Index))
End Sub
