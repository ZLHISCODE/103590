VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmAppforBill 
   Caption         =   "�������뵥"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14805
   Icon            =   "frmAppforBill.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   14805
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   150
      Width           =   165
   End
   Begin VB.PictureBox picRight 
      Height          =   8205
      Left            =   10950
      ScaleHeight     =   8145
      ScaleWidth      =   3765
      TabIndex        =   5
      Top             =   510
      Width           =   3825
      Begin VB.PictureBox picDel 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   3510
         MouseIcon       =   "frmAppforBill.frx":6852
         MousePointer    =   99  'Custom
         Picture         =   "frmAppforBill.frx":6B5C
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   29
         Top             =   30
         Width           =   255
      End
      Begin VB.PictureBox picAdd 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   3150
         MouseIcon       =   "frmAppforBill.frx":755E
         MousePointer    =   99  'Custom
         Picture         =   "frmAppforBill.frx":7868
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   28
         Top             =   30
         Width           =   255
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFList 
         Height          =   3585
         Left            =   30
         TabIndex        =   30
         Top             =   330
         Width           =   3735
         _cx             =   6588
         _cy             =   6324
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16706793
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         Editable        =   0
         ShowComboButton =   0
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
      Begin VSFlex8Ctl.VSFlexGrid VSFSeled 
         Height          =   3735
         Left            =   30
         TabIndex        =   31
         Top             =   4320
         Width           =   3735
         _cx             =   6588
         _cy             =   6588
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16706793
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         Editable        =   0
         ShowComboButton =   0
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ��(˫��ȡ��ѡ��)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   33
         Top             =   4050
         Width           =   2865
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "�ղؼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   32
         Top             =   30
         Width           =   855
      End
   End
   Begin VB.PictureBox picLeft 
      Height          =   8205
      Left            =   0
      ScaleHeight     =   8145
      ScaleWidth      =   10815
      TabIndex        =   4
      Top             =   480
      Width           =   10875
      Begin VB.PictureBox picyj 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   0
         ScaleHeight     =   585
         ScaleWidth      =   10935
         TabIndex        =   6
         Top             =   7380
         Width           =   10935
         Begin VB.TextBox txtAppforAdvice 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   150
            Width           =   1275
         End
         Begin VB.PictureBox picDate 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   10290
            MouseIcon       =   "frmAppforBill.frx":826A
            MousePointer    =   99  'Custom
            Picture         =   "frmAppforBill.frx":83BC
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   9
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtAppFordate 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7860
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "2011-10-16 20:30:30"
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox txtAppForDept 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   150
            Width           =   1515
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   300
            TabIndex        =   13
            Top             =   150
            Width           =   855
         End
         Begin VB.Line Line10 
            X1              =   1200
            X2              =   2505
            Y1              =   450
            Y2              =   450
         End
         Begin VB.Line Line11 
            X1              =   0
            X2              =   10800
            Y1              =   30
            Y2              =   30
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6690
            TabIndex        =   12
            Top             =   120
            Width           =   1125
         End
         Begin VB.Line Line12 
            X1              =   7680
            X2              =   10140
            Y1              =   450
            Y2              =   450
         End
         Begin VB.Label lblAppForDept 
            BackStyle       =   0  'Transparent
            Caption         =   "�������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   11
            Top             =   150
            Width           =   1095
         End
         Begin VB.Line Line13 
            X1              =   4290
            X2              =   6120
            Y1              =   450
            Y2              =   450
         End
      End
      Begin VB.PictureBox picym 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3675
         Left            =   240
         ScaleHeight     =   3675
         ScaleWidth      =   10965
         TabIndex        =   14
         Top             =   300
         Width           =   10965
         Begin VB.PictureBox picDisTwo 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   300
            ScaleHeight     =   525
            ScaleWidth      =   10245
            TabIndex        =   41
            Top             =   1260
            Width           =   10245
            Begin VB.TextBox txtDiagnose 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   690
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   90
               Width           =   8160
            End
            Begin VB.PictureBox picFindDiagnose 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   9690
               MouseIcon       =   "frmAppforBill.frx":8DBE
               MousePointer    =   99  'Custom
               Picture         =   "frmAppforBill.frx":8F10
               ScaleHeight     =   360
               ScaleWidth      =   360
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   60
               Width           =   360
            End
            Begin VB.Line Line 
               X1              =   690
               X2              =   9660
               Y1              =   375
               Y2              =   375
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "���:"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   0
               TabIndex        =   44
               Top             =   90
               Width           =   765
            End
         End
         Begin VB.PictureBox picDiagnose 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   300
            ScaleHeight     =   1935
            ScaleWidth      =   10605
            TabIndex        =   40
            Top             =   1230
            Width           =   10605
            Begin XtremeDockingPane.DockingPane dkpDiagnose 
               Left            =   390
               Top             =   240
               _Version        =   589884
               _ExtentX        =   450
               _ExtentY        =   423
               _StockProps     =   0
            End
         End
         Begin VB.TextBox txtFind 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   990
            TabIndex        =   20
            Top             =   3240
            Width           =   3300
         End
         Begin VB.TextBox txtID 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txtAge 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5460
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   870
            Width           =   795
         End
         Begin VB.TextBox txtSex 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3330
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   870
            Width           =   555
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   870
            Width           =   1275
         End
         Begin VB.CheckBox chkEmergency 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   9600
            TabIndex        =   15
            Top             =   390
            Width           =   795
         End
         Begin VB.Line Line8 
            X1              =   -90
            X2              =   10980
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������뵥"
            BeginProperty Font 
               Name            =   "����"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4200
            TabIndex        =   26
            Top             =   210
            Width           =   1875
         End
         Begin VB.Line Line4 
            X1              =   7950
            X2              =   10470
            Y1              =   1170
            Y2              =   1170
         End
         Begin VB.Label lblID 
            BackStyle       =   0  'Transparent
            Caption         =   "�����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6900
            TabIndex        =   25
            Top             =   870
            Width           =   1005
         End
         Begin VB.Line Line3 
            X1              =   5460
            X2              =   6210
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4740
            TabIndex        =   24
            Top             =   870
            Width           =   765
         End
         Begin VB.Line Line2 
            X1              =   3330
            X2              =   3840
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2640
            TabIndex        =   23
            Top             =   870
            Width           =   765
         End
         Begin VB.Line Line1 
            X1              =   990
            X2              =   2295
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   300
            TabIndex        =   22
            Top             =   870
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   300
            TabIndex        =   21
            Top             =   3240
            Width           =   630
         End
         Begin VB.Line Line6 
            X1              =   990
            X2              =   4290
            Y1              =   3450
            Y2              =   3450
         End
      End
      Begin VB.PictureBox picItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   180
         ScaleHeight     =   2895
         ScaleWidth      =   10695
         TabIndex        =   34
         Top             =   4440
         Width           =   10695
         Begin VB.PictureBox picItemRight 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3195
            Left            =   4170
            ScaleHeight     =   3195
            ScaleWidth      =   5865
            TabIndex        =   38
            Top             =   150
            Width           =   5865
            Begin VSFlex8Ctl.VSFlexGrid VSFItem 
               Height          =   1920
               Index           =   0
               Left            =   1350
               TabIndex        =   39
               Top             =   390
               Width           =   3915
               _cx             =   6906
               _cy             =   3387
               Appearance      =   2
               BorderStyle     =   0
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
               BackColor       =   12648447
               ForeColor       =   -2147483640
               BackColorFixed  =   12648447
               ForeColorFixed  =   -2147483630
               BackColorSel    =   12648447
               ForeColorSel    =   0
               BackColorBkg    =   12648447
               BackColorAlternate=   12648447
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483635
               FloodColor      =   192
               SheetBorder     =   16777215
               FocusRect       =   0
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   0
               GridLinesFixed  =   0
               GridLineWidth   =   0
               Rows            =   3
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   0
               RowHeightMin    =   350
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   ""
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
               Editable        =   0
               ShowComboButton =   0
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
         Begin VB.Frame fraWE 
            Height          =   3375
            Left            =   3930
            MousePointer    =   9  'Size W E
            TabIndex        =   35
            Top             =   90
            Width           =   75
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfScrollLeft 
            Height          =   4185
            Left            =   30
            TabIndex        =   36
            Top             =   90
            Width           =   3825
            _cx             =   6747
            _cy             =   7382
            Appearance      =   2
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   0
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   100
            RowHeightMax    =   100
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
            Begin VB.PictureBox picItemLeft 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   4035
               Left            =   120
               ScaleHeight     =   4035
               ScaleWidth      =   3555
               TabIndex        =   37
               Top             =   60
               Width           =   3555
            End
         End
      End
      Begin XtremeSuiteControls.TabControl TabcrlPage 
         Height          =   8025
         Left            =   30
         TabIndex        =   27
         Top             =   -30
         Width           =   11010
         _Version        =   589884
         _ExtentX        =   19420
         _ExtentY        =   14155
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin VB.OptionButton optGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "�ι�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   3660
      TabIndex        =   3
      Top             =   60
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox rtfAppend 
      Height          =   315
      Left            =   2730
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmAppforBill.frx":9F92
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "����˵��:�������ѡ�������ȡ������;�Ҽ���ѡ��ɼ��걾��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   8745
      Width           =   6435
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAppforBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnShow As Boolean                         '�����Ƿ���ʾ
Private mstrReturnSel As String                     '����ѡ��
Private mlngModifyAppForNO As Long                  '����ID����ѡ����ǰ�����뵥���ݣ��޸��á�
Private mstrAdvice As String                        '����ҽ���ִ�
Private mblnCancel As Boolean                       '������ȡ����ť
Private mlngPatientID As Long                       '����ID
Private mlngPatientPage As Long                     '��ҳID
Private mintPatientType As Integer                  '������Դ
Private mvar����ID As Variant                       '��ҳID��Һŵ���
Private mstrDiagnose As String                      '���id
Private mstrDiagnoseTxt As String                   '�������
Private mintBaby As Integer                         'Ӥ��
Private mstrAdvItem As String                       'ҽ�������
Private mlngApplyBillType As Long                   '���뵥���� 0=������ѡ����ϸ  1=����ѡ����ϸ
Private mstrAppend As String                        '��ѡ�����������
Private mstrSplieListTag As String                  '�ָ���
Private mstrSplieItemTag As String                  '�ָ���
Private mstrSplieColTag As String                   '�ָ���
Private mstrinData As String
Private mstrItemCode As String
Private mlngItem As Long                            'ѡ�е���Ŀҳ
Private mlngRow As Long                             '֮ǰѡ�е���
Private mlngCol As Long                             '֮ǰѡ�е���
Private mBlnShowDiagnose As Boolean                 '�ж�ʱ���ǿ�����ʾ���ѡ�����İ汾
Private mstrTreVsf As String                        '�����������ڵ�VSF
Private mstrBabyZK As String                        'Ӥ��ת��ʱ�䣬��������ת�ƺ��Ӥ��ֻ�ܲ�¼ҽ��
Private mlngAppForDeptID As Long                    '�������ID
Private mintPatientSex As String                    '�����Ա�

Private Type Sizetype
'  �ؼ���������
   wp As Single
   hp As Single
   tp As Single
   lp As Single
   X1 As Single
   X2 As Single
   Y1 As Single
   Y2 As Single
End Type

Private Size() As Sizetype '�ؼ���������
Private mobjMedRecPage As Object                    'zlMedRecPage����
Private mobjfrmDockDiagEdit As Object               '���ѡ����

Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Private Const FCONTROL = 8                  'ctrl��ϼ�

Private Sub Sizeinit()
'��ʼ���ؼ�����
  Dim i As Integer
  On Error Resume Next
  For i = 0 To Controls.Count - 1
    With Size(i)
        If Not Controls(i).Name Like "Line*" Then
            .wp = Controls(i).Width / Me.ScaleWidth
            .hp = Controls(i).Height / Me.ScaleHeight
            .lp = Controls(i).Left / Me.ScaleWidth
            .tp = Controls(i).Top / Me.ScaleHeight
        Else
            .X1 = Controls(i).X1 / Me.ScaleWidth
            .X2 = Controls(i).X2 / Me.ScaleWidth
            .Y1 = Controls(i).Y1 / Me.ScaleHeight
            .Y2 = Controls(i).Y2 / Me.ScaleHeight
        End If
    End With

  Next i
End Sub
Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ConMenu_Browse_SelAll                  'ȫѡ
        Call SelorClearAll(1)
    Case ConMenu_Browse_ClsAll                  'ȫ��
        Call SelorClearAll(2)
    Case ConMenu_Appfro_ModifyItem              '����

        mstrReturnSel = GetSelVal()
        Call CheckAndSaveDiagnose
        If mstrReturnSel <> "���ܱ���" Then
            Unload Me
        Else
            mstrReturnSel = ""
        End If
    Case ConMenu_Appfor_ClincHelp               '���Ʋο�
        Call ShowClincHelp
    Case ConMenu_Appfro_Exit                    '�˳�
        mblnCancel = True
        Unload Me
    End Select
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-12-17
'��    ��:  ȫѡ/ȫ��
'��    ��:
'           intType 1=ȫѡ��2=ȫ��
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Sub SelorClearAll(ByVal intType As Integer)
    Dim intRow As Integer
    Dim intCol As Integer

    With vsfItem(mlngItem)
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                If .TextMatrix(intRow, intCol) <> "" Then
                    Call GetItems(.Index, 1, intRow, intCol, intType)
                    Call SetColWith(vsfItem(.Index))
                End If
            Next
        Next

    End With
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/19
'��    ��:��鲢���������Ϣ
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Private Sub CheckAndSaveDiagnose()
    If mBlnShowDiagnose = True Then
        Call mobjMedRecPage.CheckData
        Call mobjMedRecPage.SaveData(0, mstrDiagnose)
    End If
End Sub

Private Sub cbrthis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    On Error Resume Next
    With Me.picLeft
        .Left = Left
        .Top = Top
        .Height = Bottom - Top - 10
        .Width = Right - Me.picRight.Width
    End With
    With Me.picRight
        .Left = Right - .Width
        .Top = picLeft.Top
        .Height = picLeft.Height
    End With
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Appfor_ClincHelp       '���Ʋο�
            Control.Visible = VerCompare(gSysInfo.VersionHIS, "10.35.120") <> -1
    End Select
End Sub

Private Sub chkEmergency_Click()
    Dim intIndex As Integer
    intIndex = Me.TabcrlPage.Selected.Index
    Me.picTab(intIndex).Tag = Me.chkEmergency.value
End Sub

Private Sub dkpDiagnose_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If mobjfrmDockDiagEdit Is Nothing Then Exit Sub
    Select Case Item.ID
        Case 100
            Item.Handle = mobjfrmDockDiagEdit.hWnd
    End Select
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    '��ȡ�ؼ�����
    On Error Resume Next
    For i = 0 To Controls.Count - 1
        If Controls(i).Visible = True Then
            If TypeName(Controls(i)) = "TextBox" Or TypeName(Controls(i)) = "Label" _
                Or TypeName(Controls(i)) = "PictureBox" Or TypeName(Controls(i)) = "CheckBox" Then

                Controls(i).Left = Size(i).lp * Me.ScaleWidth
                Controls(i).Width = Size(i).wp * Me.ScaleWidth
            End If
        End If
    Next i

    For i = 0 To Controls.Count - 1
        If Controls(i).Visible = True Then
            If Controls(i).Name Like "Line*" Then
                Controls(i).X1 = Size(i).X1 * Me.ScaleWidth
                Controls(i).X2 = Size(i).X2 * Me.ScaleWidth
                Controls(i).ZOrder
            End If
        End If
    Next i
End Sub

Private Sub Form_Activate()
    If mblnShow = False Then
        mblnShow = True
        '��ʼ�����ҳ��
        CheckVersionHIS
        If mBlnShowDiagnose = True Then
            Call LoadDiagnoseInfoPage
        End If
        
        Call LoadDate
        Call LoadKey
    End If
End Sub

Private Function CheckVersionHIS()
    If VerCompare(gSysInfo.VersionHIS, "10.35.110") <> -1 Then
        mBlnShowDiagnose = True
    Else
        mBlnShowDiagnose = False
    End If
End Function

'
Private Sub Form_Load()
'���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True    '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '�˵�����
    Me.cbrthis.ActiveMenuBar.Title = "�˵�"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_ModifyItem, "����")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_SelAll, "ȫѡ")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_ClsAll, "ȫ��")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ClincHelp, "���Ʋο�")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Exit, "�˳�")
        cbrControl.BeginGroup = True
    End With

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next

    '�����
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, vbKeyA, ConMenu_Browse_SelAll
        .Add FCONTROL, vbKeyD, ConMenu_Browse_ClsAll
        .Add FCONTROL, vbKeyS, ConMenu_Appfro_ModifyItem
        .Add FCONTROL, vbKeyH, ConMenu_Appfor_ClincHelp
        .Add FCONTROL, vbKeyQ, ConMenu_Appfro_Exit
    End With

    With Me.TabcrlPage
        .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = True
        .PaintManager.BoldSelected = True
    End With
    With VSFSeled
        .Rows = 0
        .Cols = 3
        .ColKey(0) = "����"
        .ColKey(1) = "��Ŀ": .ColHidden(1) = True
        .ColKey(2) = "�걾": .ColHidden(2) = True
    End With
    With vsfList
        .Rows = 0
        .Cols = 2
        .ColKey(0) = "����"
        .ColKey(1) = "��Ŀ": .ColHidden(1) = True
    End With

    '�ָ���
    mstrSplieColTag = "<Split A>"
    mstrSplieItemTag = "<Split B>"
    mstrSplieListTag = "<Split C>"
    mstrAdvItem = ""

End Sub
'
Private Sub LoadDiagnoseInfoPage()
          Dim objPanle As Pane
          Dim strDiagnoseID As String
          Dim strDiagnoseStr As String


1         On Error GoTo LoadDiagnoseInfoPage_Error
            
2         If mobjMedRecPage Is Nothing Then
3             Set mobjMedRecPage = CreateObject("zlMedRecPage.clsDiagEdit")
              strDiagnoseID = mstrDiagnose
4             Call mobjMedRecPage.InitDiagEdit(gcnHisOracle, 100, IIf(mintPatientType = 2, 1261, 1260), , 1)
5             If mobjMedRecPage.ShowDiagEdit(Me, mlngModifyAppForNO, mlngPatientID, mlngPatientPage, mintPatientType, Val(txtAppForDept.Tag), _
                                      strDiagnoseID, strDiagnoseStr, 9, , mobjfrmDockDiagEdit) = True Then

6             End If
7         End If

8         Set objPanle = dkpDiagnose.CreatePane(100, picDiagnose.Width, picDiagnose.Height, DockTopOf, Nothing)
9         objPanle.Options = PaneNoCaption '�Ƿ���Ը���
10        dkpDiagnose.Options.ThemedFloatingFrames = True
11        dkpDiagnose.Options.UseSplitterTracker = False 'ʵʱ�϶�
12        dkpDiagnose.Options.AlphaDockingContext = True
13        dkpDiagnose.Options.CloseGroupOnButtonClick = True
14        dkpDiagnose.Options.HideClient = True


15        Exit Sub
LoadDiagnoseInfoPage_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(LoadDiagnoseInfoPage)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear

End Sub
'
Private Sub Form_Unload(Cancel As Integer)
    mblnShow = False
    Call SaveKey
    mlngItem = 0
    mlngCol = 0
    mlngRow = 0
  mlngModifyAppForNO = 0                  '����ID����ѡ����ǰ�����뵥���ݣ��޸��á�
  mstrAdvice = ""                          '����ҽ���ִ�
'  mblnCancel = False                          '������ȡ����ť
  mlngPatientID = 0                       '����ID
  mlngPatientPage = 0                     '��ҳID
  mintPatientType = 0                     '������Դ
  mvar����ID = 0                          '��ҳID��Һŵ���
  mstrDiagnoseTxt = ""                     '�������
  mintBaby = 0                            'Ӥ��
  mstrAdvItem = ""                         'ҽ�������
  mlngApplyBillType = 0                   '���뵥���� 0=������ѡ����ϸ  1=����ѡ����ϸ
  mstrAppend = ""                          '��ѡ�����������
  mstrSplieListTag = ""                      '�ָ���
  mstrSplieItemTag = ""                      '�ָ���
  mstrSplieColTag = ""                       '�ָ���
  mstrinData = ""
  mstrItemCode = ""
  mvar����ID = Null
  Set mobjMedRecPage = Nothing
  If mBlnShowDiagnose = True Then
    If Not mobjfrmDockDiagEdit Is Nothing Then Unload mobjfrmDockDiagEdit
    Set mobjfrmDockDiagEdit = Nothing
  End If
  mstrTreVsf = ""
  mstrBabyZK = ""
  mlngAppForDeptID = 0
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2020-01-09
'��    ��:
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Sub LoadDate()
      '���ܶ�������
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsType As ADODB.Recordset
          Dim rsGroupTmp As ADODB.Recordset
          Dim intloop As Integer, intGroupCount As Integer
          Dim intCol As Integer, strItemNO As String
          Dim intCols As Integer
          Dim intRow As Integer
          Dim strPatientType As String    '�������
          Dim intHaveGroup As Integer
          Dim rsGroup As ADODB.Recordset
          Dim blnFind As Boolean

1         On Error GoTo LoadDate_Error

2         TabcrlPage.Visible = False
3         Select Case mintPatientType
          Case 1
4             strPatientType = "3,1"
5         Case 2
6             strPatientType = "3,2"
7         Case 3
8             strPatientType = "1"
9         Case 4
10            strPatientType = "4"
11        End Select

          '��ʼ������(���Ҿ���ID��סԺʱΪ��ҳID������ʱΪ�Һŵ���)
12        If mintPatientType <> 2 Then
13            strSQL = "select NO from ���˹Һż�¼ where id = [1] "
14            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "��ȡ����NO", mlngPatientPage)
15            If rsTmp.RecordCount > 0 Then
16                mvar����ID = rsTmp(0)
17            End If
18        Else
19            mvar����ID = mlngPatientPage
20        End If

21        intCols = 3
22        intCol = 1
23        intRow = 1

24        If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
25            strSQL = "select id,����,����,Ĭ��ִ��С��,ִ��С��,��ɫ,�Ƿ��������뵥 from �������뵥 where ����id = [1] order by ���� "
26            Set rsType = ComOpenSQL(Sel_Lis_DB, strSQL, "���������Ŀ", Val(txtAppForDept.Tag))
27            If rsType.RecordCount = 0 Then
28                strSQL = "select id,����,����,Ĭ��ִ��С��,ִ��С��,��ɫ,�Ƿ��������뵥 from �������뵥 where ����id is null order by ���� "
29                Set rsType = ComOpenSQL(Sel_Lis_DB, strSQL, "���������Ŀ")
30            End If
31        Else
32            strSQL = "select id,����,����,Ĭ��ִ��С��,ִ��С��,��ɫ from �������뵥 where ����id = [1] order by ���� "
33            Set rsType = ComOpenSQL(Sel_Lis_DB, strSQL, "���������Ŀ", Val(txtAppForDept.Tag))
34            If rsType.RecordCount = 0 Then
35                strSQL = "select id,����,����,Ĭ��ִ��С��,ִ��С��,��ɫ from �������뵥 where ����id is null order by ���� "
36                Set rsType = ComOpenSQL(Sel_Lis_DB, strSQL, "���������Ŀ")
37            End If
38        End If

          '��ѯվ���Ӧ��ִ��С��
39        If gUserInfo.NodeNo <> "-" Then
40            strSQL = "Select Distinct ���� From ����С���¼ Where վ�� = [1] or վ�� is null"
41        Else
42            strSQL = "Select Distinct ���� From ����С���¼"
43        End If
44        Set rsGroup = ComOpenSQL(Sel_Lis_DB, strSQL, "����С���¼", gUserInfo.NodeNo)

45        Do Until rsType.EOF
              '�����ǰ������û�е�ǰվ���ִ��С�飬����ʾ�÷���
46            If Not rsGroup Is Nothing Then
47                If rsGroup.RecordCount > 0 Then
48                    rsGroup.MoveFirst
49                End If
50                blnFind = False
51                Do While Not rsGroup.EOF
52                    If rsType("Ĭ��ִ��С��") & "" <> "" Then
53                        If rsType("Ĭ��ִ��С��") & "" = rsGroup("����") & "" Then
54                            blnFind = True
55                        End If
56                    ElseIf InStr("," & rsType("ִ��С��") & ",", "," & rsGroup("����") & ",") > 0 Then
57                        blnFind = True
58                    End If
59                    rsGroup.MoveNext
60                Loop
61            End If

              '���ط�ҳ
62            If blnFind Then
63                If intloop > 0 Then
64                    Load picTab(intloop)
65                    picTab(intloop).BackColor = rsType("��ɫ")
66                Else
67                    SetParent picym.hWnd, picTab(intloop).hWnd
68                    SetParent picyj.hWnd, picTab(intloop).hWnd
69                    SetParent picItem.hWnd, picTab(intloop).hWnd
70                    picTab(intloop).BackColor = rsType("��ɫ")
71                End If

72                picTab(intloop).Visible = False
73                picym.Visible = True
74                picyj.Visible = True
75                Call TabcrlPage.InsertItem(intloop, rsType("����") & "", picTab(intloop).hWnd, 0)

                  '����
76                strSQL = "Select distinct ID, ����, ���� From �������뵥���� where ���뵥id =[1]  order by ID "
77                Set rsGroupTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�������뵥����", Val(rsType("id")))
78                intHaveGroup = 0
79                Do Until rsGroupTmp.EOF
80                    strItemNO = blnGroupItem(Val(rsType("id")), strPatientType, Val(rsGroupTmp("id")))
81                    If strItemNO <> "" Then
82                        intHaveGroup = intHaveGroup + 1
83                        If optGroup.Count - 1 > 0 Or intGroupCount > 0 Or optGroup(0).Caption = "δ������Ŀ" Then
84                            intGroupCount = optGroup.Count
85                            Load optGroup(intGroupCount)
86                            With optGroup(intGroupCount)
87                                .Caption = rsGroupTmp("����")
88                                .Tag = rsGroupTmp("id") & "|" & rsType("����")
89                                .Visible = False
90                            End With
91                            SetParent optGroup(intGroupCount).hWnd, picItemLeft.hWnd

92                            Load vsfItem(intGroupCount)
93                            vsfItem(intGroupCount).Tag = rsType("����")
94                            vsfItem(intGroupCount).Visible = False

95                            Call getVSFItem(rsType, strItemNO, intGroupCount, Val(rsGroupTmp("id")), intGroupCount)
96                            optGroup(optGroup.Count - 1).Tag = optGroup(optGroup.Count - 1).Tag & "|2"
97                        Else
98                            intGroupCount = intGroupCount + 1
99                            With optGroup(0)
100                               .Caption = rsGroupTmp("����")
101                               .Tag = rsGroupTmp("id") & "|" & rsType("����")
102                           End With
103                           SetParent optGroup(0).hWnd, picItemLeft.hWnd
104                           Call miveOptGroup(intloop, intGroupCount - intHaveGroup, optGroup.Count - 1, intHaveGroup, Val(rsType("id")), strItemNO)
105                           Call getVSFItem(rsType, strItemNO, 0, Val(rsGroupTmp("id")), 0)

106                           optGroup(optGroup.Count - 1).Tag = optGroup(optGroup.Count - 1).Tag & "|2"
107                       End If
108                   End If

109                   rsGroupTmp.MoveNext
110               Loop

111               If intloop = 0 And intGroupCount = 0 Then
112                   vsfItem(0).Tag = rsType("����")
113               Else
114                   Load optGroup(optGroup.Count)
115                   Load vsfItem(vsfItem.Count)
116                   vsfItem(vsfItem.Count - 1).Tag = rsType("����")
117                   intGroupCount = intGroupCount + 1
118                   If intGroupCount = 2 And intHaveGroup = 1 Then intGroupCount = 1
119               End If

                  'δ������Ŀ
120               optGroup(optGroup.Count - 1).Caption = "δ������Ŀ"
121               optGroup(optGroup.Count - 1).Tag = "δ������Ŀ|" & rsType("����")

122               SetParent optGroup(optGroup.Count - 1).hWnd, picItemLeft.hWnd
123               strItemNO = blnGroupItem(Val(rsType("id")), strPatientType, 0)
124               Call miveOptGroup(intloop, intGroupCount - intHaveGroup, optGroup.Count - 1, intHaveGroup, Val(rsType("id")), strItemNO)
125               Call getVSFItem(rsType, strItemNO, optGroup.Count - 1, 0, intGroupCount)

126               If intHaveGroup = 0 Then
127                   optGroup(optGroup.Count - 1).Tag = optGroup(optGroup.Count - 1).Tag & "|1"
128               Else
129                   optGroup(optGroup.Count - 1).Tag = optGroup(optGroup.Count - 1).Tag & "|2"
130               End If

131               intloop = intloop + 1
132           End If

133           rsType.MoveNext
134       Loop

135       If intloop = 0 Then
136           picTab(intloop).Visible = True
137           vsfItem(intloop).Visible = True
138           picym.Visible = True
139           picyj.Visible = True
140           Call TabcrlPage.InsertItem(intloop, " " & "", picTab(intloop).hWnd, 0)
141       End If
142       Call SetColour(Me.TabcrlPage.Selected.Index)

143       TabcrlPage_SelectedChanged TabcrlPage.Item(0)
144       Call GetModifyItem(mlngModifyAppForNO, mlngPatientID, "")
145       TabcrlPage.Visible = True
146       If mstrinData <> "" Then
147           Call GetModifySelect(mstrinData)
148       ElseIf mstrItemCode <> "" Then
149           ChoseItem mstrItemCode
150       End If


151       Exit Sub
LoadDate_Error:
152       Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(LoadDate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
153       Err.Clear
End Sub

Private Function blnGroupItem(ByVal lngTypeid As Long, ByVal strPatientType As String, ByVal lngGroupTmp As Long) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItemNO As String

1         On Error GoTo blnGroupItem_Error

2         If lngGroupTmp = 0 Then
3             If gUserInfo.NodeNo <> "-" Then
4                 strSQL = "Select B.Id, B.����, B.����,b.���Ʊ���,b.����걾,c.��ɫ,d.���� С�����,d.���� С������,c.ִ��С��,d.HIS���ű���  " & _
                          "  From �������뵥��ϸ A, ���������Ŀ B,�������뵥 c,����С���¼ d" & _
                          "  Where a.���뵥id = c.id and c.Ĭ��ִ��С�� = d.����(+)  and A.���id = B.Id and b.ͣ������ is null And A.���뵥id = [1] " & _
                          "  And b.���Ʊ��� is not null and  (b.վ��= [3] or b.վ�� is null ) and (d.վ��=[3] or d.վ�� is null) and a.����id is null" & vbNewLine & _
                          "  order by a.����˳��, b.���� "
5             Else
6                 strSQL = "Select B.Id, B.����, B.����,b.���Ʊ���,b.����걾,c.��ɫ,d.���� С�����,d.���� С������,c.ִ��С��,d.HIS���ű���  " & _
                          "  From �������뵥��ϸ A, ���������Ŀ B,�������뵥 c,����С���¼ d " & _
                          "  Where a.���뵥id = c.id and c.Ĭ��ִ��С�� = d.����(+)  and A.���id = B.Id and b.ͣ������ is null And A.���뵥id = [1] " & _
                          "  And b.���Ʊ��� is not null and a.����id  is null  order by a.����˳��, b.���� "
7             End If
8         Else
9             If gUserInfo.NodeNo <> "-" Then
10                strSQL = "Select B.Id, B.����, B.����,b.���Ʊ���,b.����걾,c.��ɫ,d.���� С�����,d.���� С������,c.ִ��С��,d.HIS���ű���  " & _
                          "  From �������뵥��ϸ A, ���������Ŀ B,�������뵥 c,����С���¼ d,�������뵥���� F" & _
                          "  Where a.���뵥id = c.id and c.Ĭ��ִ��С�� = d.����(+) and a.���뵥id= f.���뵥id and f.id=a.����id  and A.���id = B.Id and b.ͣ������ is null And A.���뵥id = [1] " & _
                          "  And b.���Ʊ��� is not null  and  (b.վ��= [3] or b.վ�� is null ) and (d.վ��=[3] or d.վ�� is null) and f.id= [4]" & vbNewLine & _
                          "  order by a.����˳��, b.���� "
11            Else
12                strSQL = "Select B.Id, B.����, B.����,b.���Ʊ���,b.����걾,c.��ɫ,d.���� С�����,d.���� С������,c.ִ��С��,d.HIS���ű���  " & _
                          "  From �������뵥��ϸ A, ���������Ŀ B,�������뵥 c,����С���¼ d ,�������뵥���� F" & _
                          "  Where a.���뵥id = c.id and c.Ĭ��ִ��С�� = d.����(+) and a.���뵥id= f.���뵥id and f.id=a.����id  and A.���id = B.Id and b.ͣ������ is null And A.���뵥id = [1] " & _
                          "  And b.���Ʊ��� is not null and f.id =[4] order by a.����˳��,b.���� "
13            End If
14        End If
15        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���������Ŀ", lngTypeid, strPatientType, gUserInfo.NodeNo, lngGroupTmp)
16        Do Until rsTmp.EOF
17            strItemNO = strItemNO & "," & rsTmp("���Ʊ���")
18            rsTmp.MoveNext
19        Loop
20        If strItemNO <> "" Then strItemNO = Mid(strItemNO, 2)
21        strSQL = "SELECT ����  from  ������ĿĿ¼ where  ���� in  (Select * From Table(Cast(f_str2list([1]) As Zltools.t_Strlist)))   and ������� in  (Select * From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))) "
22        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���������Ŀ", strItemNO, strPatientType)
23        strItemNO = ""
24        Do Until rsTmp.EOF
25            strItemNO = strItemNO & "," & rsTmp("����")
26            rsTmp.MoveNext
27        Loop
28        If strItemNO <> "" Then strItemNO = Mid(strItemNO, 2)
29        blnGroupItem = strItemNO


30        Exit Function
blnGroupItem_Error:
31        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(blnGroupItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
32        Err.Clear

End Function

Private Sub getVSFItem(ByVal rsType As ADODB.Recordset, ByVal strItemNO As String, ByVal intGroupCount As Integer, ByVal lngGroupTmp As Long, _
                       ByVal intloop As Long)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset, rsList As ADODB.Recordset
          Dim rsDept As ADODB.Recordset
          Dim rsItem As ADODB.Recordset
          Dim lngColWidth As Long
          Dim intCols As Integer
          Dim intCol As Integer
          Dim intRow As Integer
          Dim lngExecDeptID As Long
          Dim strExecDept As String
          Dim blnTre As Boolean           '�Ƿ��������������뵥
          Dim strTre As String            '������Ŀʱ�䷽��

          'ȡ�õ�ǰ���뵥��Ĭ����Ŀִ�п���
1         On Error GoTo getVSFItem_Error

2         If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
3             blnTre = IIf(Val(rsType("�Ƿ��������뵥") & "") = 1, True, False)
4         End If

5         If gUserInfo.NodeNo <> "-" Then
6             strSQL = "select id,����,HIS���ű��� from ����С���¼ where ���� in (Select * From Table(Cast(f_str2list([1]) As Zltools.t_Strlist))) and (վ��=[2] or վ�� is null)"
7         Else
8             strSQL = "select id,����,HIS���ű��� from ����С���¼ where ���� in (Select * From Table(Cast(f_str2list([1]) As Zltools.t_Strlist))) "
9         End If
10        If rsType("Ĭ��ִ��С��") & "" <> "" Then
11            Set rsDept = ComOpenSQL(Sel_Lis_DB, strSQL, "ѡ��ִ��С��", CStr(rsType("Ĭ��ִ��С��") & ""), gUserInfo.NodeNo)
12        Else
13            Set rsDept = ComOpenSQL(Sel_Lis_DB, strSQL, "ѡ��ִ��С��", CStr(rsType("ִ��С��") & ""), gUserInfo.NodeNo)
14        End If

15        Do Until rsDept.EOF
16            strSQL = "select id,���� from ���ű� where ���� = [1] and (վ��='" & gUserInfo.NodeNo & "' or վ�� is null)"
17            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ѡ����", rsDept("HIS���ű���") & "")
18            If rsTmp.RecordCount > 0 Then
19                lngExecDeptID = rsTmp("ID") & ""
20                strExecDept = rsTmp("����") & ""
21                Exit Do
22            Else
                  '                MsgBox "û������HIS���Ҷ������벻��ȷ���뵽���������м��!", vbInformation, "ѡ��ִ�п���"
                  '                Exit Sub
23            End If
24            rsDept.MoveNext
25        Loop
26        If lngExecDeptID = 0 Then
27            optGroup(intGroupCount).Tag = optGroup(intGroupCount).Tag & "|0"
28            MsgBox "û������HIS���Ҷ������벻��ȷ���뵽���������м��!", vbInformation, "ѡ��ִ�п���"
29            Exit Sub
30        End If

31        With vsfItem(intGroupCount)
32            If lngGroupTmp = 0 Then
33                If gUserInfo.NodeNo <> "-" Then
34                    strSQL = "Select B.Id, B.����, B.����,���Ʊ���,b.����걾,c.��ɫ,d.���� С�����,d.���� С������,c.ִ��С��,d.HIS���ű���,b.΢��������  " & _
                               "  From �������뵥��ϸ A, ���������Ŀ B,�������뵥 c,����С���¼ d" & _
                               "  Where a.���뵥id = c.id and c.Ĭ��ִ��С�� = d.����(+)  and A.���id = B.Id and b.ͣ������ is null And A.���뵥id = [1] " & _
                               "  And b.���Ʊ��� in   (Select * From Table(Cast(f_str2list([2]) As Zltools.t_Strlist)))  and  (b.վ��= [3] or b.վ�� is null )" & _
                               "  and (d.վ��=[3] or d.վ�� is null) and a.����id is null and (nvl(b.�����Ա�,0)=[5] or nvl(b.�����Ա�,0)=0)" & vbNewLine & _
                               "  order by a.����˳��, b.���� "
35                Else
36                    strSQL = "Select B.Id, B.����, B.����,���Ʊ���,b.����걾,c.��ɫ,d.���� С�����,d.���� С������,c.ִ��С��,d.HIS���ű���,b.΢��������  " & _
                               "  From �������뵥��ϸ A, ���������Ŀ B,�������뵥 c,����С���¼ d " & _
                               "  Where a.���뵥id = c.id and c.Ĭ��ִ��С�� = d.����(+)  and A.���id = B.Id and b.ͣ������ is null And A.���뵥id = [1] " & _
                               "  And b.���Ʊ��� in  (Select * From Table(Cast(f_str2list([2]) As Zltools.t_Strlist)))   and a.����id  is null and (nvl(b.�����Ա�,0)=[5] or nvl(b.�����Ա�,0)=0) order by a.����˳��, b.���� "
37                End If
38            Else
39                If gUserInfo.NodeNo <> "-" Then
40                    strSQL = "Select B.Id, B.����, B.���� ,���Ʊ���,b.����걾,c.��ɫ,d.���� С�����,d.���� С������,c.ִ��С��,d.HIS���ű���,b.΢��������  " & _
                               "  From �������뵥��ϸ A, ���������Ŀ B,�������뵥 c,����С���¼ d,�������뵥���� F" & _
                               "  Where a.���뵥id = c.id and c.Ĭ��ִ��С�� = d.����(+) and a.���뵥id= f.���뵥id and f.id=a.����id  and A.���id = B.Id  and b.ͣ������ is null And A.���뵥id = [1] " & _
                               "  And b.���Ʊ��� in  (Select * From Table(Cast(f_str2list([2]) As Zltools.t_Strlist)))   and  (b.վ��= [3] or b.վ�� is null )" & _
                               "   and (d.վ��=[3] or d.վ�� is null) and f.id= [4] and (nvl(b.�����Ա�,0)=[5] or nvl(b.�����Ա�,0)=0) " & vbNewLine & _
                               "  order by a.����˳��, b.���� "
41                Else
42                    strSQL = "Select B.Id, B.����, B.����,���Ʊ���,b.����걾,c.��ɫ,d.���� С�����,d.���� С������,c.ִ��С��,d.HIS���ű���,b.΢�������� " & _
                               "  From �������뵥��ϸ A, ���������Ŀ B,�������뵥 c,����С���¼ d,�������뵥���� F" & _
                               "  Where a.���뵥id = c.id and c.Ĭ��ִ��С�� = d.����(+) and a.���뵥id= f.���뵥id and f.id=a.����id  and A.���id = B.Id and b.ͣ������ is null And A.���뵥id = [1] " & _
                               "  And b.���Ʊ��� in  (Select * From Table(Cast(f_str2list([2]) As Zltools.t_Strlist)))   and f.id =[4] and (nvl(b.�����Ա�,0)=[5] or nvl(b.�����Ա�,0)=0) order by a.����˳��,b.���� "
43                End If
44            End If

45            Set rsItem = ComOpenSQL(Sel_Lis_DB, strSQL, "���������Ŀ", Val(rsType("id")), strItemNO, gUserInfo.NodeNo, lngGroupTmp, mintPatientSex)
46            optGroup(intGroupCount).Tag = optGroup(intGroupCount).Tag & "|" & rsItem.RecordCount
47            If blnTre = False Then
48                If mlngApplyBillType = 0 Then
49                    lngColWidth = .Width / 3 - 50
50                    .Cols = 3: intCols = 2
51                    .ColKey(0) = "��Ŀ1"    ': .ColWidth(0) = lngColWidth
52                    .ColKey(1) = "��Ŀ2"    ': .ColWidth(1) = lngColWidth
53                    .ColKey(2) = "��Ŀ3"
54                    intCol = 0
55                    intRow = 0
56                    Do Until rsItem.EOF

57                        .TextMatrix(intRow, intCol) = rsItem("����")
58                        .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = 2
                          '�����ʽ����<��ĿID,�걾,��Ŀ��,�ڼ�ҳ,�ɿƿ���ID,ִ�п���ID,ҽ������,����ID,�ɼ���������,ִ�п�������,�ɼ�ID,�ɼ���,���С�
59                        .Cell(flexcpData, intRow, intCol, intRow, intCol) = rsItem("���Ʊ���") & mstrSplieColTag & rsItem("����걾") & mstrSplieColTag & _
                                                                              rsItem("����") & mstrSplieColTag & intloop & mstrSplieColTag & mstrSplieColTag & lngExecDeptID & _
                                                                              mstrSplieColTag & mstrSplieColTag & rsItem("���Ʊ���") & mstrSplieColTag & mstrSplieColTag & strExecDept & _
                                                                              mstrSplieColTag & mstrSplieColTag & mstrSplieColTag & mstrSplieColTag & rsItem("ID")

60                        If intCol >= intCols Then
61                            intRow = intRow + 1
62                            .Rows = intRow + 1
63                            intCol = 0
64                        Else
65                            intCol = intCol + 1
66                        End If
                          '������ɫ(�ȼ�¼����һ������)
                          '                VSFItem(intGroupCount).BackColor = Val(rsItem("��ɫ") & "")
67                        rsItem.MoveNext
68                    Loop
69                Else
70                    lngColWidth = .Width / 3 - 100
71                    .Rows = 0
72                    .Cols = 4
                      '            .AllowUserResizing = flexResizeBoth
73                    .ExtendLastCol = True
74                    .OutlineBar = flexOutlineBarComplete
75                    .Cols = 3: intCols = 2
76                    .ColKey(0) = "��Ŀ1"    ': .ColWidth(0) = lngColWidth
77                    .ColKey(1) = "��Ŀ2"    ': .ColWidth(1) = lngColWidth
78                    .ColKey(2) = "��Ŀ3"
79                    intCol = 0
80                    intRow = 0

81                    Do Until rsItem.EOF

82                        .AddItem rsItem("����")
83                        .IsSubtotal(.Rows - 1) = True
84                        .Cell(flexcpChecked, .Rows - 1, 0, .Rows - 1, 0) = 2
                          '�����ʽ����<��ĿID,�걾,��Ŀ����,�ڼ�ҳ,�ɿƿ���ID,ִ�п���ID,ҽ������,����ID,�ɼ���������,ִ�п�������,�ɼ�ID,�ɼ���,����>
85                        .Cell(flexcpData, .Rows - 1, 0, .Rows - 1, 0) = rsItem("���Ʊ���") & mstrSplieColTag & rsItem("����걾") & mstrSplieColTag & _
                                                                          rsItem("����") & mstrSplieColTag & intloop & mstrSplieColTag & mstrSplieColTag & lngExecDeptID & _
                                                                          mstrSplieColTag & mstrSplieColTag & rsItem("���Ʊ���") & mstrSplieColTag & mstrSplieColTag & strExecDept & _
                                                                          mstrSplieColTag & mstrSplieColTag & mstrSplieColTag & mstrSplieColTag & rsItem("ID")

86                        If Val(rsItem("΢��������") & "") = 1 Then
87                            If gUserInfo.NodeNo <> "-" Then
88                                strSQL = "Select Distinct c.���� ���Ʊ���, c.������ ����, a.����걾" & vbNewLine & _
                                           "   From ���������Ŀ A, �������ϸ�� B, ����ϸ����¼ C" & vbNewLine & _
                                           "   Where a.Id = b.���id(+) And b.ϸ��id = c.Id(+) And a.���Ʊ��� = [1] and (a.վ��=[2] or a.վ�� is null)"
89                            Else
90                                strSQL = "Select Distinct c.���� ���Ʊ���, c.������ ����, a.����걾" & vbNewLine & _
                                           "   From ���������Ŀ A, �������ϸ�� B, ����ϸ����¼ C" & vbNewLine & _
                                           "   Where a.Id = b.���id(+) And b.ϸ��id = c.Id(+) And a.���Ʊ��� = [1] "
91                            End If
92                        Else
93                            If gUserInfo.NodeNo <> "-" Then
94                                strSQL = "Select Distinct c.ָ����� ���Ʊ���, c.������ ����, a.����걾, c.�������" & vbNewLine & _
                                           "   From ���������Ŀ A, �������ָ�� B, ����ָ�� C" & vbNewLine & _
                                           "   Where a.Id = b.���id(+) And b.��Ŀid = c.Id(+) And a.���Ʊ��� = [1] and (a.վ��=[2] or a.վ�� is null) order by c.�������"
95                            Else
96                                strSQL = "Select Distinct c.ָ����� ���Ʊ���, c.������ ����, a.����걾, c.�������" & vbNewLine & _
                                           "   From ���������Ŀ A, �������ָ�� B, ����ָ�� C" & vbNewLine & _
                                           "   Where a.Id = b.���id(+) And b.��Ŀid = c.Id(+) And a.���Ʊ��� = [1] order by c.�������"
97                            End If
98                        End If
99                        Set rsList = ComOpenSQL(Sel_Lis_DB, strSQL, "������Ŀ��ϸ", rsItem("���Ʊ���") & "", gUserInfo.NodeNo)


100                       If rsList.RecordCount > 0 Then
101                           .Rows = .Rows + 1
102                       End If
103                       intCol = 0
104                       Do Until rsList.EOF
105                           If rsList("����") & "" <> "" Then
106                               .TextMatrix(.Rows - 1, intCol) = rsList("����") & ""
                                  '�����ʽ����<��ĿID,�걾,��Ŀ����,�ڼ�ҳ,�ɿƿ���ID,ִ�п���ID,ҽ������,����ID,�ɼ���������,ִ�п�������,�ɼ�ID,�ɼ���,����>
107                               .Cell(flexcpData, .Rows - 1, intCol, .Rows - 1, intCol) = rsList("���Ʊ���") & mstrSplieColTag & rsList("����걾") & _
                                                                                            mstrSplieColTag & rsList("����") & mstrSplieColTag & intloop & mstrSplieColTag & mstrSplieColTag & lngExecDeptID & _
                                                                                            mstrSplieColTag & mstrSplieColTag & rsItem("���Ʊ���") & mstrSplieColTag & mstrSplieColTag & strExecDept & _
                                                                                            mstrSplieColTag & mstrSplieColTag & mstrSplieColTag & mstrSplieColTag & rsItem("ID")

108                               .IsCollapsed(.Rows - 1) = flexOutlineCollapsed
109                               If intCol >= intCols Then
110                                   .Rows = .Rows + 1
111                                   intCol = 0
112                               Else
113                                   intCol = intCol + 1
114                               End If
115                           End If
116                           rsList.MoveNext
117                       Loop
118                       If intCol = 0 Then
119                           .Rows = .Rows - 1
120                       End If
                          '������ɫ(�ȼ�¼����һ������)
                          '                VSFItem(intGroupCount).BackColor = Val(rsItem("��ɫ") & "")
121                       rsItem.MoveNext
122                   Loop
123               End If
124           Else
                  '��������
125               mstrTreVsf = mstrTreVsf & "," & intGroupCount
126               lngColWidth = .Width / 3 - 100
127               .Rows = 0
128               .Cols = 4
                  '            .AllowUserResizing = flexResizeBoth
129               .ExtendLastCol = True
130               .OutlineBar = flexOutlineBarComplete
131               .Cols = 3: intCols = 2
132               .ColKey(0) = "��Ŀ1"    ': .ColWidth(0) = lngColWidth
133               .ColKey(1) = "��Ŀ2"    ': .ColWidth(1) = lngColWidth
134               .ColKey(2) = "��Ŀ3"
135               intCol = 0
136               intRow = 0

137               Do Until rsItem.EOF
138                   .AddItem rsItem("����")
139                   .IsSubtotal(.Rows - 1) = True
140                   .Cell(flexcpChecked, .Rows - 1, 0, .Rows - 1, 0) = 2
141                   .Cell(flexcpData, .Rows - 1, 0, .Rows - 1, 0) = rsItem("���Ʊ���") & mstrSplieColTag & rsItem("����걾") & mstrSplieColTag & _
                                                                      rsItem("����") & mstrSplieColTag & intloop & mstrSplieColTag & mstrSplieColTag & lngExecDeptID & _
                                                                      mstrSplieColTag & mstrSplieColTag & rsItem("���Ʊ���") & mstrSplieColTag & mstrSplieColTag & strExecDept & _
                                                                      mstrSplieColTag & mstrSplieColTag & mstrSplieColTag & mstrSplieColTag & rsItem("ID")


142                   strSQL = "Select c.������ || '(' || b.����ʱ�� || ')' ָ��, b.id, b.����ʱ��" & vbCrLf & _
                               "   From �������ָ�� A, ��������ʱ�䷽�� B, ����ָ�� C" & vbCrLf & _
                               "   Where a.��Ŀid = b.��Ŀid And a.��ĿID = c.id And b.��Ŀid = c.id And a.���id = [1] and nvl(�Ƿ�ͣ��,0)=0"
143                   Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��������ʱ�䷽��", Val(rsItem("ID") & ""))
144                   If rsTmp.RecordCount > 0 Then
145                       .Rows = .Rows + 1
146                   End If
147                   intCol = 0
148                   Do While Not rsTmp.EOF

                          '                    strTre = strTre & "<Split2>" & rsTmp("ID") & "<Split3>" & rsTmp("����ʱ��")
149                       .TextMatrix(.Rows - 1, intCol) = rsTmp("ָ��")
150                       .Cell(flexcpChecked, .Rows - 1, intCol, .Rows - 1, intCol) = 2
151                       .Cell(flexcpData, .Rows - 1, intCol, .Rows - 1, intCol) = rsItem("���Ʊ���") & mstrSplieColTag & rsTmp("ID") & mstrSplieColTag & rsTmp("����ʱ��")
152                       If intCol >= intCols Then
153                           .Rows = .Rows + 1
154                           intCol = 0
155                       Else
156                           intCol = intCol + 1
157                       End If
158                       rsTmp.MoveNext
159                   Loop
160                   If intCol = 0 Then
161                       .Rows = .Rows - 1
162                   End If
163                   rsItem.MoveNext
164               Loop
165           End If
166           .AutoSizeMode = flexAutoSizeColWidth
167           Call SetColWith(vsfItem(intGroupCount))

168       End With


169       Exit Sub
getVSFItem_Error:
170       Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(getVSFItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
171       Err.Clear

End Sub


Private Sub SetColour(intIndex As Integer)
    '������ɫ
    Dim lngColour As Long
    Dim i As Integer

    lngColour = picTab(intIndex).BackColor


    picDate.BackColor = lngColour


    TxtName.BackColor = lngColour
    txtSex.BackColor = lngColour
    txtAge.BackColor = lngColour
    txtID.BackColor = lngColour
    txtFind.BackColor = lngColour

    txtAppforAdvice.BackColor = lngColour
    txtAppFordate.BackColor = lngColour
    txtAppForDept.BackColor = lngColour
    chkEmergency.BackColor = lngColour


    picym.BackColor = lngColour
    picyj.BackColor = lngColour
    picDisTwo.BackColor = lngColour
    txtDiagnose.BackColor = lngColour
    picTab(intIndex).BackColor = lngColour
    For i = 0 To optGroup.Count - 1
        optGroup.Item(i).BackColor = lngColour

        vsfItem.Item(i).BackColor = lngColour
        vsfItem(i).BackColorAlternate = lngColour
        vsfItem(i).BackColorBkg = lngColour
        vsfItem(i).BackColorFixed = lngColour
        vsfItem(i).BackColorSel = lngColour
    Next
    picItemLeft.BackColor = lngColour
    picItemRight.BackColor = lngColour
    fraWE.BackColor = lngColour

    '����ҽ���������ɫ
    If mBlnShowDiagnose = True Then
        Call mobjMedRecPage.SetFrmColor(lngColour)
    End If
End Sub



Private Sub miveOptGroup(ByVal intIndex As Integer, ByVal intCont As Integer, ByVal intTag As Integer, intAdoCount As Integer, _
                        Optional lngTypeid As Long, Optional strItemNO As String)
    Dim i As Integer

    If optGroup.Count - 1 > 0 Then
        For i = intCont To optGroup.Count - 1
            With optGroup.Item(i)
                .Left = 10
                If i = intCont Then
                    .Top = 100
                Else
                    .Top = optGroup.Item(i - 1).Top + optGroup.Item(i - 1).Height + 100
                End If
            End With
        Next
    End If
    '��ʼ���ؼ���С
    ReDim Size(0 To Controls.Count - 1)
    Call Sizeinit
End Sub

Private Sub fraWE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button <> 1 Then Exit Sub
    With Me.fraWE
        If .Left + X < 2000 Or picItem.Width - (.Left + X) < 2000 Then Exit Sub
        .Left = .Left + X
        .Tag = .Left
    End With
    With Me.vsfScrollLeft
        .Width = Me.fraWE.Left
        Me.picItemLeft.Width = .Width
    End With

    With Me.picItemRight
        .Left = Me.fraWE.Left + Me.fraWE.Width
        .Width = Me.picItem.Width - .Left
    End With

End Sub

Private Sub optGroup_Click(Index As Integer)
    Dim i As Integer
    Dim strTag As String
    On Error Resume Next
    vsfItem(mlngItem).Cell(flexcpFontBold, mlngRow, mlngCol, mlngRow, mlngCol) = False
    On Error GoTo 0
    mlngItem = Index

    mlngCol = 0
    mlngRow = 0
    txtFind.Text = ""
    If optGroup(Index).Tag <> "" Then
        strTag = Split(optGroup(Index).Tag, "|")(1)
        For i = 0 To vsfItem.Count - 1
            If i = Index Then
                vsfItem(Index).Visible = True

                If InStr(optGroup(Index).Tag, "<Split1>") = 0 Then
                    optGroup(Index).Tag = optGroup(Index).Tag & "|<Split1>" & Index
                End If
            Else
                vsfItem(i).Visible = False
                If strTag = Split(optGroup(i).Tag, "|")(1) Then
                    If InStr(optGroup(i).Tag, "<Split1>") > 0 Then
                        optGroup(i).Tag = Mid(optGroup(i).Tag, 1, InStr(optGroup(i).Tag, "<Split1>") - 1)
                    End If
                End If
            End If
        Next
    End If
    Call SetColWith(vsfItem(Index))
End Sub

Private Sub optGroup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    optGroup(Index).ToolTipText = optGroup(Index).Caption
End Sub

Private Sub picAdd_Click()
    Dim strName As String
    Dim strItems As String
    strName = frmAppforBillSaveAs.ShowMe(Me)
    If Not mobjfrmDockDiagEdit Is Nothing Then Call EnableWindow(mobjfrmDockDiagEdit.hWnd, True)   'ǿ���������������༭
    If strName = "" Then Exit Sub
    strItems = GetSelItem(1)
    If strItems = "" Then Exit Sub
    'д���б���
    With Me.vsfList
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("����")) = strName
        .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = strItems
    End With
End Sub

Private Sub picDate_Click()
    Dim strData As String
    strData = frmDateSel.ShowMe(Me)
    Call EnableWindow(mobjfrmDockDiagEdit.hWnd, True)   'ǿ���������������༭
    If strData <> "" Then
        txtAppFordate.Text = strData
    End If
End Sub

Private Sub picDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDate.BorderStyle = 1
End Sub

Private Sub picDate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDate.BorderStyle = 0
End Sub

Private Sub picDel_Click()
    With Me.vsfList
        If .Row < 0 Then Exit Sub
        Call DelKey(.TextMatrix(.Row, .ColIndex("����")))
        Call .RemoveItem(.Row)
    End With
End Sub

Private Sub picFindDiagnose_Click()
          '�����ѡ����
          Dim objSelDlg As Object
          Dim strDiagnoseID As String
          Dim strDiagnoseStr As String

1         On Error GoTo picDiagnose_Click_Error

2         If objSelDlg Is Nothing Then
3             Set objSelDlg = CreateObject("zlMedRecPage.clsDiagEdit")
4             Call objSelDlg.InitDiagEdit(gcnHisOracle, 100, IIf(mintPatientType = 2, 1261, 1260))
5             strDiagnoseID = mstrDiagnose
6             If objSelDlg.ShowDiagEdit(Me, mlngModifyAppForNO, mlngPatientID, mlngPatientPage, mintPatientType, txtAppForDept.Tag, _
                                      strDiagnoseID, strDiagnoseStr, 9) = True Then
                                      
7                 txtDiagnose.Text = strDiagnoseStr
8                 mstrDiagnose = strDiagnoseID
9                 mstrDiagnoseTxt = strDiagnoseStr
10            End If
11        End If


12        Exit Sub
picDiagnose_Click_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(picDiagnose_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
14        Err.Clear
End Sub

Private Sub picItem_Resize()
    On Error Resume Next
    With Me.fraWE
        .Top = -200
        .Height = Me.picItem.Height + 500
        If .Left > 0 Then .Tag = .Left
    End With
    With Me.vsfScrollLeft
        .Left = 0
        .Top = 0
        .Width = IIf(fraWE.Left < 0, 0, fraWE.Left)
        .Height = picItem.Height
    End With
    With Me.picItemLeft
        .Left = 0
        .Top = 0
        .Width = Me.vsfScrollLeft.Width
        .Height = Me.vsfScrollLeft.Height
    End With
    With Me.picItemRight
        .Left = Me.fraWE.Left + Me.fraWE.Width
        .Top = Me.vsfScrollLeft.Top
        .Width = Me.picItem.Width - .Left
        .Height = Me.vsfScrollLeft.Height
    End With
End Sub

Private Sub picItemLeft_Resize()
    Dim i As Integer

    On Error Resume Next
    For i = 0 To optGroup.Count - 1
        optGroup(i).Width = Me.picItemLeft.Width - 100
    Next


End Sub

Private Sub picItemRight_Resize()
    On Error Resume Next
    With Me.vsfItem(mlngItem)
        .Left = 0
        .Top = 0
        .Width = Me.picItemRight.Width
        .Height = Me.picItemRight.Height
        Call SetColWith(vsfItem(mlngItem))
    End With

End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With TabcrlPage
        .Left = 0
        .Top = 0
        .Height = Me.picLeft.Height
        .Width = Me.picLeft.Width
    End With
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With VSFSeled
        .Height = Me.picRight.Height - .Top
        .Width = Me.picRight.Width - 100
    End With
    vsfList.Width = VSFSeled.Width
End Sub

Private Sub picTab_Click(Index As Integer)

    Call picTab_Resize(Index)
End Sub

Private Sub picTab_Resize(Index As Integer)
    Dim intGroup As Integer
    Dim BlnShow As Boolean
    intGroup = 2
    SetParent picym.hWnd, picTab(Index).hWnd
    SetParent picyj.hWnd, picTab(Index).hWnd
    SetParent picItem.hWnd, picTab(Index).hWnd
    BlnShow = True
    On Error Resume Next
    With Me.picym
        .Top = 0
        .Left = 0
        .Width = picTab(Index).ScaleWidth
        .Visible = True
        If mBlnShowDiagnose = False Then
            picDisTwo.Visible = True
            picDiagnose.Visible = False
            Label1.Top = picDisTwo.Top + picDisTwo.Height + 50
            txtFind.Top = Label1.Top
            Line6.Y1 = txtFind.Top + txtFind.Height + 10
            Line6.Y2 = txtFind.Top + txtFind.Height + 10
            picym.Height = 2235
        Else
            picDisTwo.Visible = False
            picDiagnose.Visible = True
        End If
    End With
    With Me.picyj
        .Top = picTab.Item(Index).ScaleHeight - picyj.Height - 10
        .Left = 0
        .Width = picym.Width
        .Visible = True
    End With
    With Me.picItem
        .Top = picym.Height + 100
        .Left = 250
        .Width = picym.Width - .Left * 2
        .Height = picyj.Top - .Top
        .Visible = True
    End With

End Sub

Private Sub TabcrlPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    lblTitle.Caption = Item.Caption & "���뵥"
    On Error Resume Next
    vsfItem(mlngItem).Cell(flexcpFontBold, mlngRow, mlngCol, mlngRow, mlngCol) = False
    On Error GoTo 0
    mlngCol = 0
    mlngRow = 0
    txtFind.Text = ""
    Call SetColour(Item.Index)
    Call picTab_Click(Item.Index)
    chkEmergency.value = Val(picTab(Item.Index).Tag)
    Call getTabPageSelec(Item)
End Sub

Private Sub getTabPageSelec(ByVal Item As XtremeSuiteControls.ITabControlItem)
          Dim i As Integer
          Dim J As Integer
          Dim varItemCount As Variant
          Dim strBH As String
          Dim intFirstGroup As Integer
          Dim lngMaxHeightLeft As Long                '�������б�������߶�

1         On Error GoTo getTabPageSelec_Error

2         lngMaxHeightLeft = optGroup.Count * (optGroup(0).Height + 100)
3         For i = 0 To optGroup.Count - 1
4             varItemCount = Split(optGroup.Item(i).Tag, "|")
5             If InStr(optGroup(i).Tag, Item.Caption) > 0 Then
6                 J = J + 1
7                 If J = 1 Then intFirstGroup = i
8                 optGroup(i).Visible = True
9                 If lngMaxHeightLeft < optGroup.Item(i).Top + optGroup.Item(i).Height Then
10                    lngMaxHeightLeft = optGroup.Item(i).Top + optGroup.Item(i).Height
11                End If
      '            VSFItem(i).Visible = True
12                If UBound(Split(optGroup(i).Tag, "<Split1>")) = 1 Then
13                    If InStr(optGroup(i).Tag, "δ������Ŀ") > 0 Then
14                        If varItemCount(3) = 1 Then
15                            optGroup.Item(i).Visible = False
16                            lngMaxHeightLeft = lngMaxHeightLeft - optGroup(0).Height - 100
17                        Else
18                            If varItemCount(2) = 0 Then
19                                optGroup.Item(i).Visible = False
20                                lngMaxHeightLeft = lngMaxHeightLeft - optGroup(0).Height - 100
21                            End If
22                        End If
23                    Else

24                    End If
25                    optGroup.Item(i).value = True
26                    vsfItem.Item(i).Visible = True
27                    strBH = "��ѡ��"

28                Else

29                    If InStr(optGroup(i).Tag, "δ������Ŀ") > 0 Then
30                        If varItemCount(3) = 1 Then
31                            optGroup.Item(i).Visible = False
32                            lngMaxHeightLeft = lngMaxHeightLeft - optGroup(0).Height - 100
33                        Else
34                            If varItemCount(2) = 0 Then
35                                optGroup.Item(i).Visible = False
36                                lngMaxHeightLeft = lngMaxHeightLeft - optGroup(0).Height - 100
37                            End If
38                        End If
39                    End If
40                End If
41            Else
42                optGroup(i).Visible = False
43                vsfItem(i).Visible = False
44                lngMaxHeightLeft = lngMaxHeightLeft - optGroup(0).Height - 100
45            End If
46        Next

47        If strBH <> "��ѡ��" Then
48            optGroup.Item(intFirstGroup).value = True
49            vsfItem.Item(intFirstGroup).Visible = True
50        End If

          '������������
51        With Me.vsfScrollLeft
52            If lngMaxHeightLeft = 0 Then
53               Me.fraWE.Left = -Me.fraWE.Width
54               Call picItem_Resize
55            Else
56                 Me.fraWE.Left = Val(fraWE.Tag)
57                 Call picItem_Resize
58                .Rows = lngMaxHeightLeft / 105
59                .RowHeight(-1) = 100
60            End If
61        End With
62        If lngMaxHeightLeft > picItemLeft.Height Then
63            picItemLeft.Height = lngMaxHeightLeft
64        End If


65        Exit Sub
getTabPageSelec_Error:
66        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(getTabPageSelec)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
67        Err.Clear
End Sub

Private Function GetSelVal() As String
          Dim intCol As Integer
          Dim intRow As Integer
          Dim intTab As Integer
          Dim strData As String
          Dim astrItem() As String
          Dim strTre As String        '��������ʱ�䷽��
          Dim lngRowTre As Long
          Dim lngColTre As Long
          Dim intPage As Integer

          '���ظ�ʽ:<�������1,ִ�п���1,����ʱ��1,�걾1,����,����,�Ƿ�֢,�ɼ�id,������Ŀid1;�������2,ִ�п���2,����ʱ��2,�걾2,����,����,�Ƿ�֢,�ɼ�id,������Ŀid1;.....>
          '�����ʽ����:<��ĿID,�걾,��Ŀ��,�ڼ�ҳ,�ɿƿ���ID,ִ�п���ID,ҽ������,����ID,�ɼ���������,ִ�п�������,�ɼ�ID,�ɼ���,���С�
1         On Error GoTo GetSelVal_Error

          'ת��Ӥ��ֻ�ܲ�¼ת��֮ǰ��ҽ��
2         If mstrBabyZK <> "" Then
3             If DateDiff("n", CDate(mstrBabyZK), CDate(Format(txtAppFordate.Text, "yyyy-mm-dd hh:mm:ss"))) > 1 Then
4                 MsgBox "��ǰӤ������" & mstrBabyZK & "ת�ƣ�ֻ�ܲ�¼ת��ǰ��ҽ��,���޸�����ʱ����ת��ʱ��֮ǰ"
5                 GetSelVal = "���ܱ���"
6                 txtAppFordate.SetFocus
7                 Exit Function
8             End If
9         End If

10        For intTab = 0 To Me.vsfItem.Count - 1
11            With Me.vsfItem(intTab)
12                For intRow = 0 To .Rows - 1
13                    For intCol = 0 To .Cols - 1
14                        If .TextMatrix(intRow, intCol) <> "" Then
15                            If mlngApplyBillType = 1 Then
16                                If .IsSubtotal(intRow) = True Then
17                                    If .Cell(flexcpChecked, intRow, 0) = 1 Then
18                                        astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
19                                        If Val(astrItem(4)) = 0 Then
20                                            MsgBox "��������Ϊ��ʱ���ܱ���", vbInformation, "���뵥"
21                                            GetSelVal = "���ܱ���"
22                                            Exit Function
23                                        End If

24                                        If Val(astrItem(5)) = 0 Then
25                                            MsgBox "ִ�п���Ϊ��ʱ���ܱ���", vbInformation, "���뵥"
26                                            GetSelVal = "���ܱ���"
27                                            Exit Function
28                                        End If

29                                        If Val(astrItem(10)) = 0 Then
30                                            MsgBox "ִ�вɼ���ʽΪ��ʱ���ܱ���", vbInformation, "���뵥"
31                                            GetSelVal = "���ܱ���"
32                                            Exit Function
33                                        End If

34                                        intPage = GetPicTabPage(.Tag)
35                                        strData = strData & mstrSplieItemTag & astrItem(4) & mstrSplieColTag & astrItem(5) & mstrSplieColTag & _
                                                    txtAppFordate.Text & mstrSplieColTag & astrItem(1) & mstrSplieColTag & _
                                                    astrItem(6) & mstrSplieColTag & astrItem(12) & mstrSplieColTag & Val(picTab(intPage).Tag) & _
                                                    mstrSplieColTag & astrItem(10) & mstrSplieColTag & GetDiagnosisItemID(astrItem(0))

                                          '��ȡ������Ŀ��ʱ�䷽��
36                                        If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Then
                                              'ѭ��ʱ�䷽���������Ƕ��ж���
37                                            For lngRowTre = intRow + 1 To .Rows - 1
38                                                If .IsSubtotal(lngRowTre) Then Exit For        '������һ�����θ��ڵ�ʱ����ʾ�Ѿ�ѭ������һ����Ŀ�ˣ��˳�ѭ��
39                                                For lngColTre = 0 To .Cols - 1
40                                                    If .Cell(flexcpChecked, lngRowTre, lngColTre) = 1 Then
41                                                        astrItem = Split(.Cell(flexcpData, lngRowTre, lngColTre, lngRowTre, lngColTre), mstrSplieColTag)
42                                                        strTre = strTre & "<split2>" & astrItem(1) & "<split3>" & astrItem(2)
43                                                    End If
44                                                Next
45                                            Next
                                              
46                                            If strTre <> "" Then strData = strData & mstrSplieColTag & "1" & strTre
47                                            strTre = ""
48                                        End If
49                                    End If
50                                End If
51                            Else
                                  '��ȡ������Ŀ��ʱ�䷽��
52                                If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Then
53                                    If .IsSubtotal(intRow) = True Then
54                                        If .Cell(flexcpChecked, intRow, 0) = 1 Then
55                                            astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
56                                            If Val(astrItem(4)) = 0 Then
57                                                MsgBox "��������Ϊ��ʱ���ܱ���", vbInformation, "���뵥"
58                                                GetSelVal = "���ܱ���"
59                                                Exit Function
60                                            End If

61                                            If Val(astrItem(5)) = 0 Then
62                                                MsgBox "ִ�п���Ϊ��ʱ���ܱ���", vbInformation, "���뵥"
63                                                GetSelVal = "���ܱ���"
64                                                Exit Function
65                                            End If

66                                            If Val(astrItem(10)) = 0 Then
67                                                MsgBox "ִ�вɼ���ʽΪ��ʱ���ܱ���", vbInformation, "���뵥"
68                                                GetSelVal = "���ܱ���"
69                                                Exit Function
70                                            End If

71                                            intPage = GetPicTabPage(.Tag)
72                                            strData = strData & mstrSplieItemTag & astrItem(4) & mstrSplieColTag & astrItem(5) & mstrSplieColTag & _
                                                        txtAppFordate.Text & mstrSplieColTag & astrItem(1) & mstrSplieColTag & _
                                                        astrItem(6) & mstrSplieColTag & astrItem(12) & mstrSplieColTag & Val(picTab(intPage).Tag) & _
                                                        mstrSplieColTag & astrItem(10) & mstrSplieColTag & GetDiagnosisItemID(astrItem(0))

                                              'ѭ��ʱ�䷽���������Ƕ��ж���
73                                            For lngRowTre = intRow + 1 To .Rows - 1
74                                                If .IsSubtotal(lngRowTre) Then Exit For        '������һ�����θ��ڵ�ʱ����ʾ�Ѿ�ѭ������һ����Ŀ�ˣ��˳�ѭ��
75                                                For lngColTre = 0 To .Cols - 1
76                                                    If .Cell(flexcpChecked, lngRowTre, lngColTre) = 1 Then
77                                                        astrItem = Split(.Cell(flexcpData, lngRowTre, lngColTre, lngRowTre, lngColTre), mstrSplieColTag)
78                                                        strTre = strTre & "<split2>" & astrItem(1) & "<split3>" & astrItem(2)
79                                                    End If
80                                                Next
81                                            Next
                                              
82                                            If strTre <> "" Then strData = strData & mstrSplieColTag & "1" & strTre
83                                            strTre = ""
84                                        End If
85                                    End If
86                                Else
87                                    If .Cell(flexcpChecked, intRow, intCol) = 1 Then
88                                        astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
89                                        If Val(astrItem(4)) = 0 Then
90                                            MsgBox "��������Ϊ��ʱ���ܱ���", vbInformation, "���뵥"
91                                            GetSelVal = "���ܱ���"
92                                            Exit Function
93                                        End If

94                                        If Val(astrItem(5)) = 0 Then
95                                            MsgBox "ִ�п���Ϊ��ʱ���ܱ���", vbInformation, "���뵥"
96                                            GetSelVal = "���ܱ���"
97                                            Exit Function
98                                        End If

99                                        If Val(astrItem(10)) = 0 Then
100                                           MsgBox "ִ�вɼ���ʽΪ��ʱ���ܱ���", vbInformation, "���뵥"
101                                           GetSelVal = "���ܱ���"
102                                           Exit Function
103                                       End If
                                          
104                                       intPage = GetPicTabPage(.Tag)
105                                       strData = strData & mstrSplieItemTag & astrItem(4) & mstrSplieColTag & astrItem(5) & mstrSplieColTag & _
                                                    txtAppFordate.Text & mstrSplieColTag & astrItem(1) & mstrSplieColTag & _
                                                    astrItem(6) & mstrSplieColTag & astrItem(12) & mstrSplieColTag & Val(picTab(intPage).Tag) & _
                                                    mstrSplieColTag & astrItem(10) & mstrSplieColTag & GetDiagnosisItemID(astrItem(0))
106                                   End If
107                               End If
108                           End If
109                       End If
110                   Next
111               Next
112           End With
113       Next
          
114       If strData <> "" Then
115           GetSelVal = Mid(strData, Len(mstrSplieColTag) + 1)
116           mblnCancel = False
117       Else
118           mblnCancel = True
119       End If


120       Exit Function
GetSelVal_Error:
121       Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetSelVal)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
122       Err.Clear
End Function

Private Function GetPicTabPage(ByVal vsfName As String)
    '���ݱ�����ƻ�ȡ�������ҳǩ����
    Dim i As Integer
    For i = 0 To TabcrlPage.ItemCount - 1
        If TabcrlPage.Item(i).Caption = vsfName Then
            GetPicTabPage = i
            Exit Function
        End If
    Next
End Function

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim lngCol As Long
    
    If vsfItem(mlngItem).Rows > 0 Then
        lngCol = mlngCol
        lngRow = mlngRow
        
        If KeyAscii = 13 Then
            Call CheckBlnName(mlngItem, lngRow, lngCol)
            mlngCol = lngCol
            mlngRow = lngRow
        Else
            vsfItem(mlngItem).Cell(flexcpFontBold, mlngRow, mlngCol, mlngRow, mlngCol) = False
            mlngCol = 0
            mlngRow = 0
        End If
    End If
End Sub

Private Sub VSFItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With Me.vsfItem(Index)
        Call GetItems(Index, Button, .MouseRow, .MouseCol)
        Call SetColWith(vsfItem(Index))
    End With
    If Not mobjfrmDockDiagEdit Is Nothing Then Call EnableWindow(mobjfrmDockDiagEdit.hWnd, True)   'ǿ���������������༭
End Sub

Private Sub VSFItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With Me.vsfItem(Index)
        If .MouseCol >= 0 And .MouseRow >= 0 And .Rows > 0 Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub GetItems(ByVal Index As Integer, ByVal Button As Integer, ByVal intRow As Integer, ByVal intCol As Integer, Optional ByVal intSelAll As Integer)
          Dim intSel As Integer
          Dim vsfCol As Integer
          Dim intSelUP As Integer
          Dim intRowUp As Integer
          Dim astrItem() As String
          Dim astrCol() As String
          Dim strTmp As String
          Dim strSample As String
          Dim lngGetSampleDept As Long
          Dim strGetSampleDept As String
          Dim lngGetExecDept As Long
          Dim strGetExecDept As String
          Dim astrSampleType() As String
          Dim lngPreViewRow As Long
          Dim lngPreViewCol As Long

1         On Error GoTo GetItems_Error

2         If Button = 1 Then
3             If mlngApplyBillType = 1 Then
4                 With vsfItem(Index)
5                     If intRow < 0 Then Exit Sub
6                     If intCol < 0 Then Exit Sub
7                     If .TextMatrix(intRow, intCol) = "" Then Exit Sub
8                     If intSelAll = 0 Then
9                         intSel = .Cell(flexcpChecked, intRow, intCol)
10                        If intSel = 1 Then
11                            intSel = 2
12                        Else
13                            intSel = 1
14                        End If
15                    Else
16                        intSel = intSelAll
17                    End If
18                    If InStr(mstrTreVsf & ",", "," & Index & ",") > 0 Then
19                        .Cell(flexcpChecked, intRow, intCol) = intSel
20                        Call CheckAllItem(vsfItem(Index), intRow, intCol, intSel, lngPreViewRow, lngPreViewCol)
21                        If Not .IsSubtotal(intRow) Then
22                            intRow = lngPreViewRow
23                            intCol = lngPreViewCol
24                            intSel = .Cell(flexcpChecked, intRow, intCol)
25                        End If
26                    Else
27                        If .IsSubtotal(intRow) = False Then Exit Sub    '���������뵥�����ϸ��Ŀʱ������Ӧ
28                        .Cell(flexcpChecked, intRow, intCol) = intSel
29                    End If

30                    If .IsSubtotal(intRow) = True Then
                          '��ʾ��ȡ����ʾ�걾
31                        astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
32                        strTmp = .TextMatrix(intRow, 0)
33                        strTmp = Replace(strTmp, "(" & astrItem(1) & ")", "")
34                        If intSel = 2 Then
35                            .TextMatrix(intRow, 0) = astrItem(2)
36                        Else
37                            If astrItem(1) <> "" Then
38                                .TextMatrix(intRow, 0) = astrItem(2) & "(" & astrItem(1) & ")"
39                            End If
40                        End If

                          'ѡ��ʱ
41                        If intSel = 1 Then

42                            Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                        Val(txtAppForDept.Tag), lngGetSampleDept, strGetSampleDept, Val(txtID.Tag)

43                            Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                        Val(txtAppForDept.Tag), lngGetExecDept, strGetExecDept, Val(txtID.Tag)

44                            If lngGetSampleDept <> 0 Then
45                                .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 4, CStr(lngGetSampleDept)), mstrSplieColTag)
46                                astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
47                                .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 8, strGetSampleDept), mstrSplieColTag)
48                                astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
49                            End If

50                            If lngGetExecDept <> 0 Then
51                                .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 5, CStr(lngGetExecDept)), mstrSplieColTag)
52                                astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
53                                .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 9, CStr(strGetExecDept)), mstrSplieColTag)
54                                astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
55                            End If

56                            astrSampleType = Split(GetSampleType(astrItem(0)), ",")
57                            If UBound(astrSampleType) > 0 Then

58                                .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 10, CStr(astrSampleType(0))), mstrSplieColTag)
59                                astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

60                                .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 11, CStr(astrSampleType(2))), mstrSplieColTag)
61                                astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

62                            End If

                              'д��ҽ������
63                            strTmp = Init���븽��(astrItem(0))
64                            If VerCompare(gSysInfo.VersionLIS, "10.35.140") = -1 Then
65                                If strTmp <> "" Then
66                                    If strTmp = "δ���" Then
67                                        strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                                    mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                                    CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                                    Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))
68                                    Else
69                                        .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, strTmp), mstrSplieColTag)
70                                        astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
71                                    End If
72                                End If
                                  '���������ʱ�����û�е���ѡ��Ҫ�صĴ��壬�򵯳�
73                            Else
74                                If strTmp = "" Or strTmp <> "δ���" Then
75                                    If CheckRefItem(Val(astrItem(13))) Then
76                                        strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                                    mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                                    CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                                    Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))
77                                    ElseIf strTmp <> "" Then
78                                        .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, strTmp), mstrSplieColTag)
79                                        astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
80                                    End If
81                                Else
82                                    strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                                mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                                CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                                Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))

83                                End If
84                            End If

85                            If strSample <> "" Then
86                                astrCol = Split(strSample, mstrSplieColTag)
87                                If astrCol(0) <> astrItem(1) Then
88                                    .TextMatrix(intRow, intCol) = astrItem(2) & "(" & astrCol(0) & ")"
89                                    .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 1, astrCol(0)), mstrSplieColTag)
90                                    astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
91                                End If
92                                If astrCol(1) <> "" Then
93                                    .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, astrCol(1)), mstrSplieColTag)
94                                    astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
95                                    If mstrAdvItem <> "" Then
96                                        mstrAdvItem = astrCol(1) & "<Split1>" & mstrAdvItem
97                                    Else
98                                        mstrAdvItem = mstrAdvItem & astrCol(1)
99                                    End If
100                               End If

101                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 4, astrCol(2)), mstrSplieColTag)
102                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

103                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 8, astrCol(3)), mstrSplieColTag)
104                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

105                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 10, astrCol(4)), mstrSplieColTag)
106                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

107                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 11, astrCol(5)), mstrSplieColTag)
108                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

109                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 5, astrCol(6)), mstrSplieColTag)
110                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

111                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 9, astrCol(7)), mstrSplieColTag)
112                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

113                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 12, astrCol(8)), mstrSplieColTag)
114                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

                                  '100                           If strTmp <> "" And strTmp <> "δ���" Then
                                  '101                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, strTmp), mstrSplieColTag)
                                  '102                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
                                  '103                           End If

115                           End If
116                       End If
117                   Else
                          'ѡ���ϼ�
118                       If intSel = 1 Then
                              'ѡ��
119                           Do While .IsSubtotal(intRow) = False
120                               intRow = intRow - 1
121                           Loop
122                           .Cell(flexcpChecked, intRow, 0, intRow, 0) = intSel
                              '��ʾ��ȡ����ʾ�걾
123                           astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
124                           strTmp = .TextMatrix(intRow, 0)
125                           strTmp = Replace(strTmp, "(" & astrItem(1) & ")", "")
126                           If astrItem(1) <> "" Then
127                               .TextMatrix(intRow, 0) = astrItem(2) & "(" & astrItem(1) & ")"
128                           End If

129                           Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                        Val(txtAppForDept.Tag), lngGetSampleDept, strGetSampleDept, Val(txtID.Tag)

130                           Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                        Val(txtAppForDept.Tag), lngGetExecDept, strGetExecDept, Val(txtID.Tag)

131                           If lngGetSampleDept <> 0 Then
132                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 4, CStr(lngGetSampleDept)), mstrSplieColTag)
133                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
134                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 8, strGetSampleDept), mstrSplieColTag)
135                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
136                           End If

137                           If lngGetExecDept <> 0 Then
138                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 5, CStr(lngGetExecDept)), mstrSplieColTag)
139                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
140                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 9, CStr(strGetExecDept)), mstrSplieColTag)
141                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
142                           End If

143                           astrSampleType = Split(GetSampleType(astrItem(0)), ",")
144                           If UBound(astrSampleType) > 0 Then

145                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 10, CStr(astrSampleType(0))), mstrSplieColTag)
146                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

147                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 11, CStr(astrSampleType(2))), mstrSplieColTag)
148                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

149                           End If
                              'д��ҽ������
150                           strTmp = Init���븽��(astrItem(0))
151                           If VerCompare(gSysInfo.VersionLIS, "10.35.140") = -1 Then
152                               If strTmp <> "" Then
153                                   If strTmp = "δ���" Then
154                                       strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                                    mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                                    CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                                    Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))
155                                   Else
156                                       .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, strTmp), mstrSplieColTag)
157                                       astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
158                                   End If
159                               End If
                                  '���������ʱ�����û�е���ѡ��Ҫ�صĴ��壬�򵯳�
160                           Else
161                               If strTmp = "" Or strTmp <> "δ���" Then
162                                   If CheckRefItem(Val(astrItem(13))) Then
163                                       strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                                    mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                                    CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                                    Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))

164                                   ElseIf strTmp <> "" Then
165                                       .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, strTmp), mstrSplieColTag)
166                                       astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
167                                   End If
168                               Else
169                                   strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                                mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                                CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                                Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))
170                               End If
171                           End If
172                           If strSample <> "" Then
173                               astrCol = Split(strSample, mstrSplieColTag)
174                               If astrCol(0) <> astrItem(1) Then
175                                   .TextMatrix(intRow, intCol) = astrItem(2) & "(" & astrCol(0) & ")"
176                                   .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 1, astrCol(0)), mstrSplieColTag)
177                                   astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
178                               End If
179                               If astrCol(1) <> "" Then
180                                   .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, astrCol(1)), mstrSplieColTag)
181                                   astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
182                                   If mstrAdvItem <> "" Then
183                                       mstrAdvItem = astrCol(1) & "<Split1>" & mstrAdvItem
184                                   Else
185                                       mstrAdvItem = mstrAdvItem & astrCol(1)
186                                   End If
187                               End If

188                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 4, astrCol(2)), mstrSplieColTag)
189                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

190                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 8, astrCol(3)), mstrSplieColTag)
191                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

192                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 10, astrCol(4)), mstrSplieColTag)
193                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

194                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 11, astrCol(5)), mstrSplieColTag)
195                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

196                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 5, astrCol(6)), mstrSplieColTag)
197                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

198                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 9, astrCol(7)), mstrSplieColTag)
199                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

200                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 12, astrCol(8)), mstrSplieColTag)
201                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)


                                  '179                           If strTmp <> "" And strTmp <> "δ���" Then
                                  '180                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, strTmp), mstrSplieColTag)
                                  '181                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
                                  '182                           End If
202                           End If
203                       Else
                              'ȡ��ѡ��
204                           Do While .IsSubtotal(intRow) = False
205                               intRow = intRow - 1
206                           Loop
                              '�õ��ϼ�(�����ϼ���ʼ���)
207                           intRowUp = intRow
208                           intRow = intRow + 1
209                           Do While .IsSubtotal(intRow) = False
210                               For vsfCol = 0 To .Cols - 1
211                                   If .Cell(flexcpChecked, intRow, vsfCol, intRow, vsfCol) = 1 Then
212                                       intSelUP = 1
213                                   End If
214                               Next
215                               intRow = intRow + 1
216                               If intRow = .Rows Then
217                                   Exit Do
218                               End If
219                           Loop
220                           If intSelUP = 0 Then
221                               .Cell(flexcpChecked, intRowUp, 0, intRowUp, 0) = 2
                                  '��ʾ��ȡ����ʾ�걾
222                               astrItem = Split(.Cell(flexcpData, intRowUp, 0, intRowUp, 0), mstrSplieColTag)
223                               strTmp = .TextMatrix(intRowUp, 0)
224                               strTmp = Replace(strTmp, "(" & astrItem(1) & ")", "")
225                               .TextMatrix(intRowUp, 0) = astrItem(2)
226                           End If
227                       End If
228                   End If
229               End With
230           Else
231               With vsfItem(Index)
232                   If intRow < 0 Then Exit Sub
233                   If intCol < 0 Then Exit Sub
234                   If .TextMatrix(intRow, intCol) = "" Then Exit Sub
235                   If intSelAll = 0 Then
236                       intSel = .Cell(flexcpChecked, intRow, intCol)
237                       If intSel = 1 Then
238                           intSel = 2
239                       Else
240                           intSel = 1
241                       End If
242                   Else
243                       intSel = intSelAll
244                   End If
245                   If InStr(mstrTreVsf & ",", "," & Index & ",") > 0 Then
246                       .Cell(flexcpChecked, intRow, intCol) = intSel
247                       Call CheckAllItem(vsfItem(Index), intRow, intCol, intSel, lngPreViewRow, lngPreViewCol)
248                       If Not .IsSubtotal(intRow) Then
249                           intRow = lngPreViewRow
250                           intCol = lngPreViewCol
251                           intSel = .Cell(flexcpChecked, intRow, intCol)
252                       End If
253                   End If

254                   astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
255                   strTmp = .TextMatrix(intRow, intCol)
256                   If intSel = 2 Then
257                       strTmp = astrItem(2)
258                       .Cell(flexcpChecked, intRow, intCol) = intSel
259                       .TextMatrix(intRow, intCol) = strTmp
260                   Else
261                       strTmp = astrItem(2) & "(" & astrItem(1) & ")"
262                       .Cell(flexcpChecked, intRow, intCol) = intSel
263                       .TextMatrix(intRow, intCol) = strTmp

264                       Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                    Val(txtAppForDept.Tag), lngGetSampleDept, strGetSampleDept, Val(txtID.Tag)

265                       Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                    Val(txtAppForDept.Tag), lngGetExecDept, strGetExecDept, Val(txtID.Tag)

266                       If lngGetSampleDept <> 0 Then
267                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 4, CStr(lngGetSampleDept)), mstrSplieColTag)
268                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
269                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 8, strGetSampleDept), mstrSplieColTag)
270                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
271                       End If

272                       If lngGetExecDept <> 0 Then
273                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 5, CStr(lngGetExecDept)), mstrSplieColTag)
274                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

275                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 9, CStr(strGetExecDept)), mstrSplieColTag)
276                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
277                       End If

278                       astrSampleType = Split(GetSampleType(astrItem(0)), ",")
279                       If UBound(astrSampleType) > 0 Then

280                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 10, CStr(astrSampleType(0))), mstrSplieColTag)
281                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

282                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 11, CStr(astrSampleType(2))), mstrSplieColTag)
283                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

284                       End If

                          'д��ҽ������
285                       strTmp = Init���븽��(astrItem(0))
286                       If VerCompare(gSysInfo.VersionLIS, "10.35.140") = -1 Then
287                           If strTmp <> "" Then
288                               If strTmp = "δ���" Then
289                                   strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                                mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                                CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                                Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))
290                               Else
291                                   .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, strTmp), mstrSplieColTag)
292                                   astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
293                               End If
294                           End If
                              '���������ʱ�����û�е���ѡ��Ҫ�صĴ��壬�򵯳�
295                       Else
296                           If strTmp = "" Or strTmp <> "δ���" Then
297                               If CheckRefItem(Val(astrItem(13))) Then
298                                   strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                                mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                                CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                                Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))
299                               ElseIf strTmp <> "" Then
300                                   .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, strTmp), mstrSplieColTag)
301                                   astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
302                               End If
303                           Else
304                               strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                            mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                            CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                            Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))
305                           End If
306                       End If

307                       If strSample <> "" Then
308                           astrCol = Split(strSample, mstrSplieColTag)
309                           If astrCol(0) <> astrItem(1) Then
310                               .TextMatrix(intRow, intCol) = astrItem(2) & "(" & astrCol(0) & ")"
311                               .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 1, astrCol(0)), mstrSplieColTag)
312                               astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
313                           End If
314                           If astrCol(1) <> "" Then
315                               .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 6, astrCol(1)), mstrSplieColTag)
316                               astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
317                               If mstrAdvItem <> "" Then
318                                   mstrAdvItem = astrCol(1) & "<Split1>" & mstrAdvItem
319                               Else
320                                   mstrAdvItem = mstrAdvItem & astrCol(1)
321                               End If
322                           End If

323                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 4, astrCol(2)), mstrSplieColTag)
324                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

325                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 8, astrCol(3)), mstrSplieColTag)
326                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

327                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 10, astrCol(4)), mstrSplieColTag)
328                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

329                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 11, astrCol(5)), mstrSplieColTag)
330                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

331                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 5, astrCol(6)), mstrSplieColTag)
332                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

333                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 9, astrCol(7)), mstrSplieColTag)
334                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

335                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 12, astrCol(8)), mstrSplieColTag)
336                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
                              '
                              '302                       If strTmp <> "" And strTmp <> "δ���" Then
                              '303                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 6, strTmp), mstrSplieColTag)
                              '304                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
                              '305                       End If
337                       End If
338                   End If
339               End With
340           End If

              'д��ѡ��
341           Call WriterSelVSF
342           CheckBlnChose Index
343       Else
344           With Me.vsfItem(Index)
345               intRow = .MouseRow
346               intCol = .MouseCol
347               If intRow < 0 Then Exit Sub
348               If intCol < 0 Then Exit Sub
349               If .TextMatrix(intRow, intCol) = "" Then Exit Sub
350               If .IsSubtotal(intRow) = False And InStr(mstrTreVsf & ",", "," & Index & ",") > 0 Then Exit Sub    '���������뵥�����ϸ��Ŀʱ������Ӧ
351               If mlngApplyBillType = 1 Then
                      '��ʾ��ϸ
352                   If .IsSubtotal(intRow) = True Then
353                       astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
354                       strTmp = .TextMatrix(intRow, 0)
355                       strTmp = Replace(strTmp, "(" & astrItem(1) & ")", "")
                          '<��ĿID,�걾,��Ŀ��,�ڼ�ҳ,�ɿƿ���ID,ִ�п���ID,ҽ������,����ID,�ɼ���������,ִ�п�������,�ɼ�ID,�ɼ�����
356                       strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                    mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                    CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                    Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))

357                       If strSample <> "" Then
358                           astrCol = Split(strSample, mstrSplieColTag)
359                           If astrCol(0) <> astrItem(1) Then
360                               .TextMatrix(intRow, intCol) = astrItem(2) & "(" & astrCol(0) & ")"
361                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 1, astrCol(0)), mstrSplieColTag)
362                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
363                           End If
364                           If astrCol(1) <> "" Then
365                               .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 6, astrCol(1)), mstrSplieColTag)
366                               astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
367                               If mstrAdvItem <> "" Then
368                                   mstrAdvItem = astrCol(1) & "<Split1>" & mstrAdvItem
369                               Else
370                                   mstrAdvItem = mstrAdvItem & astrCol(1)
371                               End If
372                           End If

373                           .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 4, astrCol(2)), mstrSplieColTag)
374                           astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

375                           .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 8, astrCol(3)), mstrSplieColTag)
376                           astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

377                           .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 10, astrCol(4)), mstrSplieColTag)
378                           astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

379                           .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 11, astrCol(5)), mstrSplieColTag)
380                           astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

381                           .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 5, astrCol(6)), mstrSplieColTag)
382                           astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

383                           .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 9, astrCol(7)), mstrSplieColTag)
384                           astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)

385                           .Cell(flexcpData, intRow, 0, intRow, 0) = Join(ReplaceArrayVal(astrItem, 12, astrCol(8)), mstrSplieColTag)
386                           astrItem = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieColTag)
387                       End If
388                   End If
389               Else
                      '����ʾ��ϸ
390                   astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
391                   strTmp = .TextMatrix(intRow, intCol)
392                   strTmp = Replace(strTmp, "(" & astrItem(1) & ")", "")
393                   strSample = frmAppforBillSelSample.ShowMe(Me, astrItem(0), astrItem(1), mlngPatientID, mvar����ID, mstrDiagnoseTxt, _
                                                                mintBaby, mintPatientType, "", astrItem(6), CLng(Val(astrItem(4))), CStr(astrItem(8)), CLng(Val(astrItem(10))), CStr(astrItem(11)), _
                                                                CLng(astrItem(5)), CStr(astrItem(9)), CStr(astrItem(12)), mlngAppForDeptID, Val(astrItem(13)), Trim(Me.txtSex.Text), _
                                                                Trim(Me.txtAge.Text), Trim(Me.txtAppForDept.Text))
394                   If strSample <> "" Then
395                       astrCol = Split(strSample, mstrSplieColTag)
396                       If astrCol(0) <> astrItem(1) Then
397                           .TextMatrix(intRow, intCol) = astrItem(2) & "(" & astrCol(0) & ")"
398                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 1, astrCol(0)), mstrSplieColTag)
399                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
400                       End If
401                       If astrCol(1) <> "" Then
402                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 6, astrCol(1)), mstrSplieColTag)
403                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
404                           If mstrAdvItem <> "" Then
405                               mstrAdvItem = astrCol(1) & "<Split1>" & mstrAdvItem
406                           Else
407                               mstrAdvItem = mstrAdvItem & astrCol(1)
408                           End If
409                       End If

410                       .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 4, astrCol(2)), mstrSplieColTag)
411                       astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

412                       .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 8, astrCol(3)), mstrSplieColTag)
413                       astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

414                       .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 10, astrCol(4)), mstrSplieColTag)
415                       astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

416                       .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 11, astrCol(5)), mstrSplieColTag)
417                       astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

418                       .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 5, astrCol(6)), mstrSplieColTag)
419                       astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

420                       .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 9, astrCol(7)), mstrSplieColTag)
421                       astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)

422                       .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 12, astrCol(8)), mstrSplieColTag)
423                       astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
424                   End If
425               End If
426           End With
427       End If


428       Exit Sub
GetItems_Error:
429       Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetItems)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
430       Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-03-14
'��    ��:  ��ѡ����������Ŀʱ��Ĭ�Ϲ�ѡ����ʱ�䷽��
'��    ��:
'           objVSF      ��ǰ��ʾ��VSF
'           lngRow      ��ǰ��ѡ����
'           lngCol      ��ǰ��ѡ����
'           intSel      ��ѡ״̬
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Private Sub CheckAllItem(objVSF As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByVal intSel As Integer, Optional lngPreViewRow As Long, Optional lngPreViewCol As Long)
          Dim strCode As String
          Dim blnChk As Boolean
          Dim lngGroupRow As Long
          Dim i As Integer
          Dim J As Integer

1         On Error GoTo CheckAllItem_Error

2         With objVSF
3             strCode = Split(.Cell(flexcpData, lngRow, lngCol, lngRow, lngCol), mstrSplieColTag)(0)
4             If .IsSubtotal(lngRow) Then
5                 For i = 1 To .Rows - 1
6                     If Not .IsSubtotal(i) Then
7                         For J = 0 To .Cols - 1
8                             If .Cell(flexcpData, i, J, i, J) <> "" Then
9                                 If Split(.Cell(flexcpData, i, J, i, J), mstrSplieColTag)(0) = strCode Then
10                                    If lngPreViewRow <> 0 Then
11                                        If lngPreViewRow = Split(.Cell(flexcpData, i, J, i, J), mstrSplieColTag)(1) Then
12                                            .Cell(flexcpChecked, i, J) = intSel
13                                        End If
14                                    Else
15                                        .Cell(flexcpChecked, i, J) = intSel
16                                    End If
17                                End If
18                            End If
19                        Next
20                    End If
21                Next
22            Else
23                For i = 0 To .Rows - 1
24                    If .IsSubtotal(i) Then
25                        If Split(.Cell(flexcpData, i, 0, i, 0), mstrSplieColTag)(0) = strCode Then
26                            If .IsSubtotal(i) = True Then
27                                lngGroupRow = i
28                                Exit For
29                            End If
30                        End If
31                    End If
32                Next
33                For i = lngGroupRow + 1 To .Rows - 1
34                    If Not .IsSubtotal(i) Then
35                        For J = 0 To .Cols - 1
36                            If .Cell(flexcpData, i, J, i, J) <> "" Then
37                                If Split(.Cell(flexcpData, i, J, i, J), mstrSplieColTag)(0) = strCode Then
38                                    If .Cell(flexcpChecked, i, J) = 1 Then blnChk = True
39                                End If
40                            End If
41                        Next
42                    End If
43                Next
44                If blnChk = False Then
45                    .Cell(flexcpChecked, lngGroupRow, 0) = 2
46                Else
47                    .Cell(flexcpChecked, lngGroupRow, 0) = 1
48                End If
49                lngPreViewRow = lngGroupRow
50                lngPreViewCol = 0
51            End If
52        End With


53        Exit Sub
CheckAllItem_Error:
54        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(CheckAllItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
55        Err.Clear
End Sub

Private Sub CheckBlnChose(Index As Integer)
          Dim i As Integer
          Dim intC As Integer

1         On Error GoTo CheckBlnChose_Error

2         With vsfItem(Index)
3             For i = .FixedRows To .Rows - 1
4                 For intC = .FixedCols To .Cols - 1
5                     If .Cell(flexcpChecked, i, intC) = 1 Then
6                         optGroup(Index).Font.Bold = True
7                         optGroup(Index).Font.Underline = True
8                         Exit Sub
9                     End If
10                Next
11            Next
12            optGroup(Index).Font.Bold = False
13            optGroup(Index).Font.Underline = False
14        End With


15        Exit Sub
CheckBlnChose_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(CheckBlnChose)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear

End Sub

Private Sub CheckBlnName(ByVal Index As Integer, lngFindRow As Long, lngFindCol As Long)
          Dim i As Integer
          Dim intC As Integer
          Dim strPY As String
          Dim strFind As String
          Dim strBH As String
          Dim strBegin As String

1         On Error GoTo CheckBlnName_Error

2         With vsfItem(Index)
              '���֮ǰ��ѡ��
3             If .Cell(flexcpFontBold, lngFindRow, lngFindCol, lngFindRow, lngFindCol) = False Then
                  '���ҵ����

4                 strBH = "xxx"
5                 lngFindRow = 0
6                 lngFindCol = 0
7             Else
8                 .Cell(flexcpFontBold, lngFindRow, lngFindCol, lngFindRow, lngFindCol) = False
9                 If lngFindRow = 0 And lngFindCol = 0 Then
10                    strBH = "***"
11                End If
12            End If
13            For i = .FixedRows To .Rows - 1
14                For intC = .FixedCols To .Cols - 1
15                    If i = lngFindRow And intC = lngFindCol Then
16                        strBegin = "Begin"
17                        If strBH = "xxx" Then
18                            strPY = GetPYCode(.Cell(flexcpText, i, intC))
19                            strFind = GetPYCode(txtFind.Text)
20                            If UCase(strPY) Like "*" & UCase(strFind) & "*" Then
              '                    If lngFindRow <= -1 Then
21                                    lngFindRow = i
22                                    lngFindCol = intC
23                                    .Cell(flexcpFontBold, i, intC, i, intC) = True
24                                    Call .Select(i, intC)
25                                    Call .ShowCell(i, intC)
26                                    Exit Sub
              '                    End If
27                            End If
28                        End If
29                    Else
30                        If strBegin = "Begin" Then
31                            strPY = GetPYCode(.Cell(flexcpText, i, intC))
32                            strFind = GetPYCode(txtFind.Text)
33                            If UCase(strPY) Like "*" & UCase(strFind) & "*" Then
34                                    lngFindRow = i
35                                    lngFindCol = intC
36                                    .Cell(flexcpFontBold, i, intC, i, intC) = True
37                                    Call .Select(i, intC)
38                                    Call .ShowCell(i, intC)
39                                    Exit Sub
40                            End If
41                        End If
42                    End If
43                Next
44            Next
45        End With


46        Exit Sub
CheckBlnName_Error:
47        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(CheckBlnName)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
48        Err.Clear

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



Private Sub vsfList_DblClick()
    With Me.vsfList
        If .Row < 0 Then Exit Sub
        Call WriterSelHistoryListVSF(.Row)
        Call WriterSelVSF
    End With
End Sub

Private Sub vsfScrollLeft_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    With Me.vsfScrollLeft
        Me.picItemLeft.Top = -.TopRow * 100
    End With
End Sub
'
'Private Sub vsfScrolRight_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    With Me.vsfScrolRight
'        Me.picItemRight.Top = -.TopRow * 100
'    End With
'End Sub

Private Sub VSFSeled_DblClick()
    With Me.VSFSeled
        If .Row < 0 Then Exit Sub
        Call .RemoveItem(.Row)
        Call WriterSelListVSF(0)
        Call WriterSelVSF
    End With

End Sub

Private Sub SaveKey()
    '����       ���泣������
    Dim intRow As Integer
    With Me.vsfList
        For intRow = 0 To .Rows - 1
            Save_AppForCollect gUserInfo.ID, .TextMatrix(intRow, .ColIndex("����")), .TextMatrix(intRow, .ColIndex("��Ŀ"))
        Next
    End With
End Sub

Private Sub DelKey(strName As String)
          '����       ɾ���ղ�
          Dim strSQL As String
1         On Error GoTo DelKey_Error

2         strSQL = "Zl_�������뵥�ղ�_del(" & gUserInfo.ID & ",'" & strName & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "���뵥�ղ�ɾ��")


4         Exit Sub
DelKey_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(DelKey)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
6         Err.Clear

End Sub

Private Sub LoadKey()
'          ���� ������������
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strName As String

1         On Error GoTo LoadKey_Error

2         strSQL = "Select ��Աid, ����, ���, ���� From �������뵥�ղ� Where ��Աid = [1] Order By ����, ��� "

3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���뵥�ղ�", gUserInfo.ID)

4         strName = ""
5         With vsfList
6             .Rows = 0
7             Do Until rsTmp.EOF
8                 If strName <> rsTmp("����") Then
9                     .Rows = .Rows + 1
10                    .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����")
11                    .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = rsTmp("����")
12                Else
13                    .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) & rsTmp("����")
14                End If
15                strName = rsTmp("����")
16                rsTmp.MoveNext
17            Loop
18        End With

19        Exit Sub
LoadKey_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(LoadKey)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
21        Err.Clear

End Sub

Public Function ShowMe(frmMain As Object, lngModifyAppForNO As Long, lngPatientID As Long, intBaby As Byte, _
                       lngPatientPage As Long, strPatientName As String, strPatientSex As String, _
                       strPatientAge As String, intPatientType As Integer, strOutPatientsNO As String, _
                       strInPatientsNO As String, lngPhysicalExamination As Double, strDiagnose As String, _
                       strAppForMan As String, lngAppForDeptID As Long, strAppForDeptName As String, _
                       lngPatientDeptID As Long, strPatientDeptName As String, _
                       Optional blnCancel As Boolean, Optional strErr As String, Optional strInData As String, _
                       Optional strItemCode As String) As String
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '����       ��ʾ���뵥
      '����           frmMain                         ������
      '               lngModifyAppForNO               �������뵥��,0=������>0�޸�
      '               lngPatientID                    ����ID
      '               intBaby                         Ӥ��
      '               lngPatientPage                  ������ҳ
      '               strPatientName                  ��������
      '               strPatientSex                   �����Ա�
      '               strPatientAge                   ��������
      '               intPatientType                  ������Դ
      '               strOutPatientsNO                �����
      '               strInPatientsNO                 סԺ��
      '               lngPhysicalExamination          ������(����)
      '               strDiagnose                     ���(���һ�����)
      '               strAppForMan                    ������
      '               lngAppForDeptID                 �������ID
      '               strAppForDeptName               ���������
      '               lngPatientDeptID                ���˿���ID
      '               strPatientDeptName              ���˿�����
      '               blnCancel                       ����True=������ȡ����ť��False=���������� ��������˷��ؽ��Ϊ��ʱҲ����Ϊ=True
      '               strErr                          �д�����Ϣʱ���ش�����Ϣ
      '����       ��������
      '           ��ʽ:
      '           ��ʽ:<�������1,ִ�п���1,����ʱ��1,������Ŀ����1,�걾1;�������2,ִ�п���2,����ʱ��2,������Ŀ����2,�걾2;.....>
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim intBabyNo As Integer  'Ӥ�����
          Dim i As Integer
1         On Error GoTo showMe_Error

2         mstrReturnSel = ""
3         mblnCancel = True

4         mlngAppForDeptID = lngAppForDeptID
5         mlngModifyAppForNO = lngModifyAppForNO
6         mstrinData = strInData
7         mstrItemCode = strItemCode
          'mstrAdvice = strAdvice
8         mlngPatientID = lngPatientID
9         TxtName.Tag = lngPatientID
10        mstrDiagnose = strDiagnose
11        mintBaby = intBaby
12        If intBaby > 0 Then
              'txtName.Text = txtName.Text & "(Ӥ��" & intBaby & ")"
              '����ΪӤ������
13            Set rsTmp = GetBabyInfor(lngPatientID, lngPatientPage, intBaby)
14            If Not rsTmp Is Nothing Then
15                TxtName.Text = NVL(rsTmp("Ӥ������")) & "(Ӥ��)"
16                txtSex.Text = NVL(rsTmp("Ӥ���Ա�"))
17                txtAge.Text = NVL(rsTmp("����"))
18                intBabyNo = Val(rsTmp("Ӥ�����") & "")
19            End If
20        Else
21            TxtName.Text = strPatientName
22            txtSex.Text = strPatientSex
23            txtAge.Text = strPatientAge
24        End If
          '���没���Ա��ڶ�ȡ�����Ŀʱ�������������Ŀ�������Ա�
25        Select Case txtSex.Text
          Case "��"
26            mintPatientSex = 1
27        Case "Ů"
28            mintPatientSex = 2
29        Case Else
30            mintPatientSex = 0
31        End Select

32        mlngPatientPage = lngPatientPage
33        lblName.Tag = lngPatientPage
34        txtSex.Tag = lngPatientDeptID
35        mintPatientType = intPatientType
36        Select Case intPatientType
          Case 1          '����'
37            lblID.Caption = "�����:"
38            txtID.Text = strOutPatientsNO
39        Case 2          'סԺ'
40            lblID.Caption = "סԺ��:"

41            txtID.Text = strInPatientsNO & IIf(intBaby > 0, "(ĸ��)", "")
42        Case 4          '���'
43            lblID.Caption = "������:"
44            txtID.Text = lngPhysicalExamination
45        Case Else
46        End Select
47        txtID.Tag = intPatientType
48        txtDiagnose.Text = strDiagnose
49        txtAppforAdvice.Text = strAppForMan
50        strSQL = "select �������� from ��������˵�� where ����id=[1]"
51        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "��ȡ��������", lngAppForDeptID)
52        If rsTmp.RecordCount > 0 Then
53            For i = 1 To rsTmp.RecordCount
54                If rsTmp("��������") & "" = "�ٴ�" Then
55                    txtAppForDept.Tag = lngPatientDeptID
56                    txtAppForDept.Text = strPatientDeptName
57                    Exit For
58                Else
59                    txtAppForDept.Tag = lngAppForDeptID
60                    txtAppForDept.Text = strAppForDeptName
61                End If
62                rsTmp.MoveNext
63            Next
64        Else
65            txtAppForDept.Tag = lngAppForDeptID
66            txtAppForDept.Text = strAppForDeptName
67        End If

68        If intBaby > 0 And VerCompare(gSysInfo.VersionHIS, "10.35.120") <> -1 Then   '������-1���Ž��д������ݵ���ZLHIS���̡�
              'Ӥ������ж�ת��ʱ��
69            If Not GetAppForDate(lngPatientID, lngPatientPage, intBabyNo, txtAppFordate) Then
70                blnCancel = True
71                Unload Me
72                Exit Function
73            End If
74        Else
75            txtAppFordate = Currentdate
76        End If

          '����ʾ��ϸģʽ
77        mlngApplyBillType = ComGetPara(Sel_Lis_DB, "���뵥��ʾ�����ϸ", gSysInfo.SysNo, gSysInfo.ModlNo, 0)
78        Me.Show vbModal, frmMain
79        strDiagnose = mstrDiagnose
80        blnCancel = mblnCancel
81        ShowMe = mstrReturnSel


82        Exit Function
showMe_Error:
83        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(showMe)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
84        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-03-21
'��    ��:  ��ȡת��Ӥ��������ʱ��
'��    ��:
'           lngPaitID       ����ID
'           intPage         ��ҳID
'           intNo           Ӥ�����
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Private Function GetAppForDate(ByVal lngPaitID As Long, ByVal intPage As Integer, ByVal intNO As Integer, objTxt As TextBox) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim dteNow As Date

1         On Error GoTo GetAppForDate_Error

2         dteNow = Currentdate
3         strSQL = "Select b.��Ժ���� as ��ֹʱ�� " & vbNewLine & _
                   "From ������������¼ A, ������ҳ B" & vbNewLine & _
                   "Where a.Ӥ������id = b.����id(+) And a.Ӥ����ҳid = b.��ҳid(+) And a.����id = [1] And a.��ҳid = [2] And a.��� = [3]"
4         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, lngPaitID, intPage, intNO)
5         If rsTmp.EOF Then
6             objTxt = dteNow
7             GetAppForDate = True
8             Exit Function
9         End If
10        If rsTmp("��ֹʱ��") & "" = "" Then
11            objTxt = dteNow
12            GetAppForDate = True
13            Exit Function
14        End If
15        mstrBabyZK = Format(rsTmp("��ֹʱ��"), "yyyy-mm-dd hh:mm:ss")
16        If dteNow > CDate(Format(rsTmp("��ֹʱ��"), "yyyy-mm-dd hh:mm:ss")) Then
17            If MsgBox("��ǰӤ������" & Format(rsTmp("��ֹʱ��"), "yyyy-mm-dd hh:mm:ss") & "ת�ƣ�ֻ�ܲ�¼ת��ǰ��ҽ�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
18                objTxt.Text = DateAdd("n", -1, CDate(Format(rsTmp("��ֹʱ��"), "yyyy-mm-dd hh:mm:ss")))
19                GetAppForDate = True
20            Else
21                GetAppForDate = False
22            End If
23        Else
24            objTxt.Text = dteNow
25            GetAppForDate = True
26        End If


27        Exit Function
GetAppForDate_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetAppForDate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
29        Err.Clear

End Function

Private Function GetDiagnosisItemID(strNO As String) As Long
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����               ͨ��������Ŀ������ȡ������ĿID
          '����               strNO ������Ŀ����
          '����               ȡ����������ĿID
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset

1         On Error GoTo GetDiagnosisItemID_Error

2         strSQL = "select id from ������ĿĿ¼ where ���� = [1] "
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ȡ������Ŀid", strNO)
4         If rsTmp.RecordCount > 0 Then
5             GetDiagnosisItemID = rsTmp("id")
6         End If


7         Exit Function
GetDiagnosisItemID_Error:
8         Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetDiagnosisItemID)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
9         Err.Clear

End Function

Private Function GetSampleType(strNO As String) As String
      '          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '          ���� ͨ��������ĿID�ɼ���ʽID
      '          ����               strNO ������Ŀ����
      '          ���� ȡ����������ĿID
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim strPatientType As String

1         On Error GoTo GetSampleType_Error

2         Select Case mintPatientType
          Case 1
3             strPatientType = "3,1"
4         Case 2
5             strPatientType = "3,2"
6         Case 3
7             strPatientType = "1"
8         Case 4
9             strPatientType = "4"
10        End Select

11        strSQL = "Select /*+ rule */ Distinct c.id, c.����, c.����" & vbNewLine & _
                 "   From ������ĿĿ¼ a, �����÷����� b, ������ĿĿ¼ c, �������ÿ��� D" & vbNewLine & _
                 "   Where a.id = b.��Ŀid And b.�÷�id = c.id And c.id = d.��Ŀid And a.���� = [1] And d.����id = [2]" & vbNewLine & _
                 " and c.������� in  (Select * From Table(Cast(F_Num2list([3]) As Zltools.T_Numlist))) " & vbNewLine & _
                   "order by c.���� "
12        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ȡ������Ŀid", strNO, mlngAppForDeptID, strPatientType)

13        If rsTmp.RecordCount <= 0 Then
14            strSQL = "Select /*+ rule */ Distinct c.id, c.����, c.����" & vbNewLine & _
                     "   From ������ĿĿ¼ a, �����÷����� b, ������ĿĿ¼ c" & vbNewLine & _
                     "   Where a.id = b.��Ŀid And b.�÷�id = c.id And a.���� = [1] " & vbNewLine & _
                     " and c.������� in  (Select * From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))) " & vbNewLine & _
                       "order by c.���� "
15            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ȡ������Ŀid", strNO, strPatientType)
16        End If

17        If rsTmp.RecordCount > 0 Then
18            GetSampleType = rsTmp("id") & "," & rsTmp("����") & "," & rsTmp("����")
19        End If


20        Exit Function
GetSampleType_Error:
21        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetSampleType)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
22        Err.Clear
          '
End Function

Private Function ReplaceArrayVal(astrVal() As String, intIndex As Integer, strVal As String) As String()
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����               �滻�����е�һ��Index�е�ֵ
          '����               astrVal                     ���������
          '                   intIndex                    Ҫ�޸ĵ�λ��
          '                   strVal                      Ҫ�޸ĵ�ֵ
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim intloop As Integer
          Dim strTmp As String
1         On Error GoTo ReplaceArrayVal_Error

2         For intloop = 0 To UBound(astrVal)
3             If intloop = intIndex Then
4                 strTmp = strTmp & mstrSplieColTag & strVal
5             Else
6                 strTmp = strTmp & mstrSplieColTag & astrVal(intloop)
7             End If
8         Next
9         If strTmp <> "" Then
10            strTmp = Mid$(strTmp, Len(mstrSplieColTag) + 1)
11        End If
12        ReplaceArrayVal = Split(strTmp, mstrSplieColTag)


13        Exit Function
ReplaceArrayVal_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(ReplaceArrayVal)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
15        Err.Clear

End Function

Private Function GetSelItem(intType As Integer) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����           ȡ�Ѿ�ѡ�����Ŀ
          '����           intType = 1 �����뵥���б���ѡȡ(VSFItem)
          '                       = 2 ����ѡ���б���ѡ��(VSFSeled)
          Dim intTab As Integer
          Dim intRow As Integer
          Dim intCol As Integer
          Dim strItem As String
          Dim astrLis() As String
          Dim astrCol() As String
          Dim intloop As Integer

1         On Error GoTo GetSelItem_Error

2         If intType = 1 Then
              'VSFItem
3             For intTab = 0 To Me.vsfItem.Count - 1
4                 With Me.vsfItem(intTab)
5                     For intRow = 0 To .Rows - 1
6                         For intCol = 0 To .Cols - 1
7                             If .Cell(flexcpChecked, intRow, intCol) = 1 Then
8                                 If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Or mlngApplyBillType = 1 Then
9                                     If .IsSubtotal(intRow) = True Then
10                                        strItem = strItem & mstrSplieListTag & .Cell(flexcpData, intRow, intCol)
11                                    End If
12                                Else
13                                    strItem = strItem & mstrSplieListTag & .Cell(flexcpData, intRow, intCol)
14                                End If
15                            End If
16                        Next
17                    Next
18                End With
19            Next
20        End If

21        If intType = 2 Then
              'VSFSeled
22            With Me.VSFSeled
23                For intRow = 0 To .Rows - 1
24                    astrLis = Split(.Cell(flexcpData, intRow, 0, intRow, 0), mstrSplieItemTag)
25                    For intloop = 0 To UBound(astrLis)
26                        astrCol = Split(astrLis(intloop), mstrSplieColTag)
27                        strItem = strItem & mstrSplieListTag & astrLis(intloop)
28                    Next
29                Next
30            End With
31        End If

32        If intType = 3 Then
              'VSFSeled
33            With Me.vsfList
34                For intRow = 0 To .Rows - 1
35                    astrLis = .Cell(flexcpData, intRow, 0, intRow, 0)
36                Next
37            End With
38        End If

39        If strItem <> "" Then
40            GetSelItem = Mid$(strItem, Len(mstrSplieListTag) + 1)
41        End If


42        Exit Function
GetSelItem_Error:
43        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetSelItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
44        Err.Clear
End Function

Private Function WriterSelVSF() As Boolean
          Dim strItem As String
          Dim astrItem() As String
          Dim astrLis() As String
          Dim astrCol() As String
          Dim intloop As Integer

          '�����ѡ��
1         On Error GoTo WriterSelVSF_Error

2         With Me.VSFSeled
3             .Rows = 0
4             .Cols = 1
5         End With

6         strItem = GetSelItem(1)
7         If strItem = "" Then Exit Function

8         astrItem = Split(strItem, mstrSplieListTag)

9         With Me.VSFSeled
10            .Rows = 0
11            .Cols = 1
12            For intloop = 0 To UBound(astrItem)
13                astrLis = Split(astrItem(intloop), mstrSplieItemTag)
14                astrCol = Split(astrLis(0), mstrSplieColTag)
15                .Rows = .Rows + 1
16                If astrCol(1) <> "" Then
17                    .TextMatrix(.Rows - 1, 0) = astrCol(2) & "(" & astrCol(1) & ")"
18                Else
19                    .TextMatrix(.Rows - 1, 0) = astrCol(2)
20                End If
21                .Cell(flexcpData, .Rows - 1, 0, .Rows - 1, 0) = astrItem(intloop)
22            Next
23        End With


24        Exit Function
WriterSelVSF_Error:
25        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(WriterSelVSF)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
26        Err.Clear
End Function

Private Function WriterSelListVSF(intType As Integer) As Boolean
          Dim strItem As String
          Dim intTab As Integer
          Dim intRow As Integer
          Dim intCol As Integer
          Dim astrItem() As String
          Dim astrCol() As String
          Dim astrLis() As String
          Dim intloop As Integer
          Dim strPID As String

          '��պ���װ��
1         On Error GoTo WriterSelListVSF_Error

2         For intTab = 0 To Me.vsfItem.Count - 1
3             With Me.vsfItem(intTab)
4                 For intRow = 0 To .Rows - 1
5                     If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Or mlngApplyBillType = 1 Then
6                         If .IsSubtotal(intRow) Then
7                             For intCol = 0 To .Cols - 1
8                                 If .TextMatrix(intRow, intCol) <> "" Then
9                                     .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = 2
10                                    If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Then
11                                        Call CheckAllItem(vsfItem(intTab), intRow, intCol, 2)
12                                    End If
13                                End If
14                            Next
15                        End If
16                    Else
17                        For intCol = 0 To .Cols - 1
18                            If .TextMatrix(intRow, intCol) <> "" Then
19                                .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = 2
20                            End If
21                        Next
22                    End If
23                Next
24            End With
              
25            CheckBlnChose intTab
26        Next
          
27        Select Case intType
              Case 0
28                strItem = GetSelItem(2)
29            Case 1
30                strItem = GetSelItem(1)
31            Case 2
32                strItem = GetSelItem(3)
33        End Select

34        If strItem = "" Then Exit Function

35        astrItem = Split(strItem, mstrSplieListTag)

          'д��ѡ��
36        For intloop = 0 To UBound(astrItem)
37            For intTab = 0 To Me.vsfItem.Count - 1
38                With Me.vsfItem(intTab)
39                    For intRow = 0 To .Rows - 1
40                        For intCol = 0 To .Cols - 1
41                            If .TextMatrix(intRow, intCol) <> "" Then
42                                astrCol = Split(astrItem(intloop), mstrSplieColTag)
43                                astrLis = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
                                  
44                                If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Or mlngApplyBillType = 1 Then
45                                    If .IsSubtotal(intRow) = True Then
46                                        strPID = astrLis(7)
47                                        If astrCol(0) = astrLis(0) And astrCol(3) = astrLis(3) And astrCol(7) = strPID Then
48                                            .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = 1
49                                            .TextMatrix(intRow, intCol) = astrLis(2) & "(" & astrLis(1) & ")"
50                                            If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Then
51                                                Call CheckAllItem(vsfItem(intTab), intRow, intCol, 1)
52                                            End If
53                                        End If
54                                    End If
55                                Else
56                                    strPID = astrLis(7)
57                                    If astrCol(0) = astrLis(0) And astrCol(3) = astrLis(3) And astrCol(7) = strPID Then
58                                        .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = 1
59                                        .TextMatrix(intRow, intCol) = astrLis(2) & "(" & astrLis(1) & ")"
60                                    End If
61                                End If
62                            End If
63                        Next
64                    Next
65                End With
                  
66                CheckBlnChose intTab
67            Next
68        Next


69        Exit Function
WriterSelListVSF_Error:
70        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(WriterSelListVSF)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
71        Err.Clear
End Function

Private Function WriterSelHistoryListVSF(intGetRow As Integer, Optional strData As String) As Boolean
          Dim strItem As String
          Dim intTab As Integer
          Dim intRow As Integer
          Dim intCol As Integer
          Dim intItem As Integer
          Dim astrItem() As String
          Dim astrItem1() As String
          Dim astrCol1() As String
          Dim astrLis() As String
          Dim astrLis1() As String
          Dim intloop As Integer
          Dim lngGetSampleDept As Long
          Dim strGetSampleDept As String
          Dim lngGetExecDept As Long
          Dim strGetExecDept As String
          Dim astrSampleType() As String
          Dim intPage As Integer

          '��պ���װ��
1         On Error GoTo WriterSelHistoryListVSF_Error

2         For intTab = 0 To Me.vsfItem.Count - 1
3             With Me.vsfItem(intTab)
4                 For intRow = 0 To .Rows - 1
5                     If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Or mlngApplyBillType = 1 Then
6                         If .IsSubtotal(intRow) Then
7                             For intCol = 0 To .Cols - 1
8                                 If .TextMatrix(intRow, intCol) <> "" Then
9                                     .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = 2
10                                    If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Then
11                                        Call CheckAllItem(vsfItem(intTab), intRow, intCol, 2)
12                                    End If
13                                End If
14                            Next
15                        End If
16                    Else
17                        For intCol = 0 To .Cols - 1
18                            If .TextMatrix(intRow, intCol) <> "" Then
19                                .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = 2
20                            End If
21                        Next
22                    End If
23                Next
24            End With
25        Next

26        If strData = "" Then
27            With Me.vsfList
28                strItem = .TextMatrix(intGetRow, 1)
29            End With
30        Else
31            strItem = strData
32        End If

33        If strItem = "" Then Exit Function

34        astrLis1 = Split(strItem, mstrSplieListTag)
          
          'д��ѡ��
35        For intTab = 0 To Me.vsfItem.Count - 1
36            For intloop = 0 To UBound(astrLis1)
37                With Me.vsfItem(intTab)
38                    For intRow = 0 To .Rows - 1
39                        If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Or mlngApplyBillType = 1 Then
40                            If .IsSubtotal(intRow) Then
41                                For intCol = 0 To .Cols - 1
42                                    If .TextMatrix(intRow, intCol) <> "" Then
43                                        astrItem1 = Split(astrLis1(intloop), mstrSplieItemTag)
44                                        astrLis = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
45                                        For intItem = 0 To UBound(astrItem1)
46                                            astrCol1 = Split(astrItem1(intItem), mstrSplieColTag)
47                                            If astrCol1(0) = astrLis(0) And astrCol1(3) = astrLis(3) And astrCol1(7) = astrLis(7) Then
48                                                astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
49                                                .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = 1
50                                                If InStr(mstrTreVsf & ",", "," & intTab & ",") > 0 Then
51                                                    If UBound(astrCol1) >= 14 Then
                                                          '���޸Ľ���������Ŀ��ѡ����ϸ
52                                                        Call CheckAllItem(vsfItem(intTab), intRow, intCol, 1, Val(astrCol1(14)))
53                                                    Else
                                                          '�ղ�˫��������Ŀ��ȫѡ��ϸ
54                                                        Call CheckAllItem(vsfItem(intTab), intRow, intCol, 1)
55                                                    End If
56                                                End If
                                                  
57                                                If InStr(.TextMatrix(intRow, intCol), "(" & astrItem(1) & ")") <= 0 Then
58                                                    .TextMatrix(intRow, intCol) = .TextMatrix(intRow, intCol) & "(" & astrItem(1) & ")"
59                                                End If
                                                  
60                                                Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                                            Val(txtAppForDept.Tag), lngGetSampleDept, strGetSampleDept, Val(txtID.Tag)
                                                  
61                                                Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                                            Val(txtAppForDept.Tag), lngGetExecDept, strGetExecDept, Val(txtID.Tag)
                                                  
62                                                If lngGetSampleDept <> 0 Then
63                                                    .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 4, IIf(Val(astrCol1(4)) <> 0, astrCol1(4), CStr(lngGetSampleDept))), mstrSplieColTag)
64                                                    astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
65                                                    .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 8, IIf(Val(astrCol1(4)) <> 0, astrCol1(8), strGetSampleDept)), mstrSplieColTag)
66                                                    astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
67                                                End If
                                                  
68                                                If lngGetExecDept <> 0 Then
69                                                    .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 5, IIf(Val(astrCol1(5)) <> 0, astrCol1(5), CStr(lngGetExecDept))), mstrSplieColTag)
70                                                    astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
71                                                    .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 9, IIf(Val(astrCol1(5)) <> 0, astrCol1(9), CStr(strGetExecDept))), mstrSplieColTag)
72                                                    astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
73                                                End If
                                                  
74                                                astrSampleType = Split(GetSampleType(astrItem(0)), ",")
75                                                If UBound(astrSampleType) > 0 Then
76                                                    .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 10, IIf(Val(astrCol1(10)) <> 0, astrCol1(10), CStr(astrSampleType(0)))), mstrSplieColTag)
77                                                    astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
78                                                    .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 11, IIf(Val(astrCol1(10)) <> 0, astrCol1(11), CStr(astrSampleType(2)))), mstrSplieColTag)
79                                                    astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
80                                                End If
                                                  
81                                                .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 12, IIf(astrCol1(12) <> "", astrCol1(12), "")), mstrSplieColTag)
82                                                If astrCol1(13) = "1" Then 'ȷ����Ŀ����ҳǩ�Ľ�����־
83                                                    intPage = GetPicTabPage(.Tag)
84                                                    picTab(intPage).Tag = astrCol1(13)
85                                                    If TabcrlPage.Selected.Index = intPage Then
86                                                        chkEmergency.value = Val(picTab(intPage).Tag)
87                                                    End If
88                                                End If
89                                            End If
90                                        Next
91                                    End If
92                                Next
93                            End If
94                        Else
95                            For intCol = 0 To .Cols - 1
96                                If .TextMatrix(intRow, intCol) <> "" Then
97                                    astrItem1 = Split(astrLis1(intloop), mstrSplieItemTag)
98                                    astrLis = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
99                                    For intItem = 0 To UBound(astrItem1)
100                                       astrCol1 = Split(astrItem1(intItem), mstrSplieColTag)
101                                       If astrCol1(0) = astrLis(0) And astrCol1(3) = astrLis(3) And astrCol1(7) = astrLis(7) Then
102                                           astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
103                                           .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = 1
                                              
104                                           If InStr(.TextMatrix(intRow, intCol), "(" & astrItem(1) & ")") <= 0 Then
105                                               .TextMatrix(intRow, intCol) = .TextMatrix(intRow, intCol) & "(" & astrItem(1) & ")"
106                                           End If
                                              
107                                           Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                                        Val(txtAppForDept.Tag), lngGetSampleDept, strGetSampleDept, Val(txtID.Tag)

108                                           Get����ִ�п��� Val(TxtName.Tag), Val(lblName.Tag), GetDiagnosisItemID(astrItem(0)), Val(txtSex.Tag), _
                                                        Val(txtAppForDept.Tag), lngGetExecDept, strGetExecDept, Val(txtID.Tag)

109                                           If lngGetSampleDept <> 0 Then
110                                               .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 4, IIf(Val(astrCol1(4)) <> 0, astrCol1(4), CStr(lngGetSampleDept))), mstrSplieColTag)
111                                               astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
112                                               .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 8, IIf(Val(astrCol1(4)) <> 0, astrCol1(8), strGetSampleDept)), mstrSplieColTag)
113                                               astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
114                                           End If

115                                           If lngGetExecDept <> 0 Then
116                                               .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 5, IIf(Val(astrCol1(5)) <> 0, astrCol1(5), CStr(lngGetExecDept))), mstrSplieColTag)
117                                               astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
118                                               .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 9, IIf(Val(astrCol1(5)) <> 0, astrCol1(9), CStr(strGetExecDept))), mstrSplieColTag)
119                                               astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
120                                           End If

121                                           astrSampleType = Split(GetSampleType(astrItem(0)), ",")
122                                           If UBound(astrSampleType) > 0 Then
123                                               .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 10, IIf(Val(astrCol1(10)) <> 0, astrCol1(10), CStr(astrSampleType(0)))), mstrSplieColTag)
124                                               astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
125                                               .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 11, IIf(Val(astrCol1(10)) <> 0, astrCol1(11), CStr(astrSampleType(2)))), mstrSplieColTag)
126                                               astrItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
127                                           End If
                                              
128                                           .Cell(flexcpData, intRow, intCol, intRow, intCol) = Join(ReplaceArrayVal(astrItem, 12, IIf(astrCol1(12) <> "", astrCol1(12), "")), mstrSplieColTag)
129                                           If astrCol1(13) = "1" Then 'ȷ����Ŀ����ҳǩ�Ľ�����־
130                                               intPage = GetPicTabPage(.Tag)
131                                               picTab(intPage).Tag = astrCol1(13)
132                                               If TabcrlPage.Selected.Index = intPage Then
133                                                   chkEmergency.value = Val(picTab(intPage).Tag)
134                                               End If
135                                           End If
136                                       End If
137                                   Next
138                               End If
139                           Next
140                       End If
141                   Next
142               End With
143           Next
              
144           CheckBlnChose intTab
145       Next


146       Exit Function
WriterSelHistoryListVSF_Error:
147       Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(WriterSelHistoryListVSF)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
End Function

Private Function Init���븽��(strSampleNO As String) As String
      '���ܣ���ȡ��Ŀ�ĵ������븽��
      '���أ���Ӧ�ĵ��ݶ��������븽��ʱ����True
          Dim strSQL As String, lngIdx As Long
          Dim arrData As Variant, strData As String
          Dim strNoneAppend As String, strHaveAppend As String
          Dim arrSub As Variant, i As Long
          Dim rsAppend As New ADODB.Recordset
          Dim lngEnd As Long
          Dim lngBegin As Long
          Dim strAppend As String
          Dim rsTmp As ADODB.Recordset
          Dim lng�Һ�ID As Long

1         On Error GoTo Init���븽��_Error

2         rtfAppend.Text = "": rtfAppend.SelStart = 0

          'ͨ���Һŵ���ѯ�Һ�ID
3         If mintPatientType = 1 Then
4             strSQL = "Select ID From ���˹Һż�¼ Where no = [1]"
5             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�Һ�ID", CStr(mvar����ID))
6             If Not rsTmp.EOF Then
7                 lng�Һ�ID = Val(rsTmp("ID") & "")
8             End If
9         Else
10            lng�Һ�ID = mvar����ID
11        End If

12        strSQL = "Select C.��Ŀ,C.����,C.Ҫ��ID,C.����,d.������,E.id " & _
                   " From ��������Ӧ�� A,�����ļ��б� B,�������ݸ��� C,����������Ŀ D,������ĿĿ¼ E" & _
                   " Where a.������ĿID = E.id and E.����=[1] And A.Ӧ�ó���=[2]" & _
                   " And A.�����ļ�ID=B.ID And B.����=7 And B.ID=C.�ļ�ID And c.Ҫ��id=d.id(+)" & _
                   " Order by C.����"


13        Set rsAppend = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, strSampleNO, 2)
14        If Not rsAppend.EOF Then
15            arrData = Split(mstrAppend, "<Split1>")
16            With rtfAppend
17                Do While Not rsAppend.EOF
                      'ȷ����������
18                    strData = ""
                      '��ȡ�°没���е����븽��
19                    If intEMR_Setup = 1 Then
20                        If Not gobjEmrInterface.IsInited Or gobjEmrInterface.IsOffline Then

21                        Else
22                            On Error Resume Next
23                            strData = gobjEmrInterface.GetOrderInspectInfoEX(mintPatientType, mlngPatientID, lng�Һ�ID, rsAppend("������") & "")
24                            If Err.Description <> "" Then
25                                Err.Clear: On Error GoTo Init���븽��_Error
26                                strData = gobjEmrInterface.GetOrderInspectInfo(mlngPatientID, rsAppend("������") & "")
27                            End If
28                        End If

29                    Else
                          '����°没����û��,��ȥ��ȡ�ϰ岡��
30                        If mstrAppend <> "" Then
                              '�޸�ʱ������ԭ������
31                            For i = 0 To UBound(arrData)
32                                arrSub = Split(arrData(i), "<Split2>")
33                                If arrSub(0) = rsAppend!��Ŀ Then
34                                    strData = arrSub(3)
35                                    If strData = "" And UBound(arrSub) >= 4 Then
                                          '���Ը��ƻ���ײ�����ҽ�������޸�ʱ�����븽��ҲҪȡȱʡֵ
36                                        If Val(arrSub(4)) = 1 Then
37                                            If Not IsNull(rsAppend!����) Then
38                                                strData = rsAppend!����
39                                            ElseIf mlngPatientID <> 0 Then
40                                                strData = GetAppendItemValue(rsAppend!��Ŀ, NVL(rsAppend!Ҫ��ID, 0), mlngPatientID, mvar����ID, _
                                                                               mstrDiagnoseTxt, mintBaby, mstrAdvItem)
41                                            End If
42                                        End If
43                                    End If

                                      '���ڵĸ���
44                                    strHaveAppend = strHaveAppend & "," & arrSub(0)
45                                    strNoneAppend = Replace(strNoneAppend & ",", "," & arrSub(0) & ",", ",")
46                                    If Right(strNoneAppend, 1) = "," Then strNoneAppend = Left(strNoneAppend, Len(strNoneAppend) - 1)
47                                ElseIf InStr(strNoneAppend & ",", "," & arrSub(0) & ",") = 0 _
                                         And InStr(strHaveAppend & ",", "," & arrSub(0) & ",") = 0 Then
48                                    strNoneAppend = strNoneAppend & "," & arrSub(0)    '�ȼǵ�û�еĸ�����
49                                End If
50                            Next
51                        Else
                              '����ʱ��ʹ��Ԥ�������ݻ�Ӳ�����������ȡ
52                            If Not IsNull(rsAppend!����) Then
53                                strData = rsAppend!����
54                            ElseIf mlngPatientID <> 0 Then
55                                strData = GetAppendItemValue(rsAppend!��Ŀ, NVL(rsAppend!Ҫ��ID, 0), mlngPatientID, mvar����ID, _
                                                               mstrDiagnoseTxt, mintBaby, mstrAdvItem)
56                            End If
57                        End If
58                    End If

                      '��������ʾ��RTF��:�����ı����һ��λ�ò���ֱ��¼�뺺��,����ȶ��һ���������Ŀո�
59                    .SelText = IIf(.Text = "", "", vbCrLf) & rsAppend!��Ŀ & "�� " & strData
60                    lngIdx = .Find(rsAppend!��Ŀ & "��", , , rtfNoHighlight Or rtfMatchCase)
61                    If lngIdx <> -1 Then
62                        .SelStart = lngIdx
63                        .SelLength = Len(rsAppend!��Ŀ & "��")
64                        .SelBold = True
65                        .SelIndent = 100
66                        .SelProtected = True
67                    End If
68                    .SelStart = Len(.Text)

69                    rsAppend.MoveNext
70                Loop

                  '��궨λ�ڵ�һ�����븽��
71                rsAppend.MoveFirst
72                lngIdx = .Find(rsAppend!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
73                If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsAppend!��Ŀ & "��") + 1

74            End With


75            rsAppend.MoveFirst
76            For i = 1 To rsAppend.RecordCount
77                strData = ""
78                lngEnd = -1
79                lngBegin = rtfAppend.Find(rsAppend!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
80                If lngBegin <> -1 Then
81                    lngBegin = lngBegin + Len(rsAppend!��Ŀ & "��")
82                    If i = rsAppend.RecordCount Then
83                        lngEnd = Len(rtfAppend.Text)
84                    Else
85                        rsAppend.MoveNext
86                        lngEnd = rtfAppend.Find(vbCrLf & rsAppend!��Ŀ & "��", lngBegin, , rtfNoHighlight Or rtfMatchCase)
87                        rsAppend.MovePrevious
88                    End If
89                End If
90                If lngBegin <> -1 And lngEnd <> -1 Then
                      'MID��������1Ϊ������rtf����0Ϊ����
91                    lngBegin = lngBegin + 1
92                    lngEnd = lngEnd + 1
93                    strData = Mid(rtfAppend.Text, lngBegin, lngEnd - lngBegin)
                      'ȥ��Ϊ��������ı����һ��λ�ò���ֱ��¼�뺺������ӵĿո�
94                    If Left(strData, 1) = " " Then strData = Mid(strData, 2)
95                    If Right(strData, 1) = " " Then strData = Left(strData, Len(strData) - 1)

96                    If Trim(strData) = "" And NVL(rsAppend!����, 0) = 1 Then
                          'MsgBox "���ݸ���""" & rsAppend!��Ŀ & """������û����д��", vbInformation, "LI���뵥"
                          '                    strSample = frmAppforBillSelSample.ShowMe(Me, strSampleNO, "", mlngPatientID, mvar����ID, mstrDiagnose, _
                                               mintBaby, mintPatientType, mstrAdvItem, "", 0, "", 0, "", 0, "", "")
97                        If Mid(rtfAppend.Text, lngBegin, 1) = " " Then
98                            rtfAppend.SelStart = lngBegin
99                        Else
100                           rtfAppend.SelStart = lngBegin - 1
101                       End If
102                       Init���븽�� = "δ���"
103                       On Error Resume Next
104                       rtfAppend.SetFocus: Exit Function
105                   ElseIf ActualLen(strData) > 4000 Then
106                       MsgBox "���ݸ���""" & rsAppend!��Ŀ & """�����ݹ������������2000�����ֻ�4000���ַ���", vbInformation, "LI���뵥"
107                       If Mid(rtfAppend.Text, lngBegin, 1) = " " Then
108                           rtfAppend.SelStart = lngBegin
109                       Else
110                           rtfAppend.SelStart = lngBegin - 1
111                       End If
112                       If rtfAppend.SelText = " " Then rtfAppend.SelStart = lngBegin
113                       On Error Resume Next
114                       rtfAppend.SetFocus: Exit Function
115                   End If
116               End If

                  'û���������ݵĸ���Ҳ�����˱���
117               strAppend = strAppend & "<Split1>" & rsAppend!��Ŀ & "<Split2>" & NVL(rsAppend!����, 0) & "<Split2>" & NVL(rsAppend!Ҫ��ID) & "<Split2>" & strData

118               rsAppend.MoveNext
119           Next
120           strAppend = Mid(strAppend, Len("<Split1>") + 1)
121           Init���븽�� = strAppend

122       End If

          '�Ѳ����ڵ�������Ŀ��ʾ
123       If strNoneAppend <> "" Then
124           MsgBox "���¸�������Ŀ��Ӧ�ĵ����������Ѳ����ڣ�" & vbCrLf & Mid(strNoneAppend, 2), vbInformation, "100"
125       End If


126       Exit Function
Init���븽��_Error:
127       Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(Init���븽��)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
128       Err.Clear
End Function

Private Function Save_AppForCollect(lngUserID As Long, strName As String, strData As String) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����          ���浱ǰ�û������뵥�ղ�
          '               lngUserID �û�ID
          '               strName   ����
          '               strData   ����
          '����           ����ʹ�÷ָ����ִ�
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim strTag As String
          Dim astrItem() As String
          Dim lngLoop As Long

1         On Error GoTo Save_AppForCollect_Error

2         strTag = "<split D>"

3         astrItem = Split(GetStrLenSeparate(strData, 2000, strTag), strTag)

4         For lngLoop = 0 To UBound(astrItem)
5             strSQL = "Zl_�������뵥�ղ�_Insert(" & lngUserID & ",'" & strName & "','" & lngLoop + 1 & "','" & astrItem(lngLoop) & "')"
6             Call ComExecuteProc(Sel_Lis_DB, strSQL, "���뵥�ղر���")
7         Next


8         Exit Function
Save_AppForCollect_Error:
9         Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(Save_AppForCollect)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
10        Err.Clear

End Function


Private Function GetStrLenSeparate(strData As String, lngLen As Long, strTag As String) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����              ��ָ���ִ���ָ���ĳ��Ƚ��зָ�
          '����               strData Ҫ�ָ����ִ�
          '                   lngLen  �ָ�����
          '                   strTag  �ָ���
          '����               �ָ��õ��ִ�
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim lngBegin As Long
          Dim strlr As String

1         On Error GoTo GetStrLenSeparate_Error

2         While lngBegin < Len(strData)

3             If Len(strData) - lngBegin <= lngLen Then
4                 strlr = strlr & strTag & Mid$(strData, lngBegin + 1, lngLen)
5                 lngBegin = Len(strData)
6             Else
7                 strlr = strlr & strTag & Mid$(strData, lngBegin + 1, lngLen)
8                 lngBegin = lngBegin + lngLen
9             End If
10        Wend

11        If strlr <> "" Then
12            strlr = Mid$(strlr, Len(strTag) + 1)
13            GetStrLenSeparate = strlr
14        End If


15        Exit Function
GetStrLenSeparate_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetStrLenSeparate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear
End Function

Private Function GetModifyItem(lngAppforNO As Long, lngPatient As Long, strDiagnose As String) As String
          '����           ��ȡ�ϴ��޸ĵ����뵥��Ŀ�����޸�
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim strData As String
          Dim intPage As Integer
          Dim lngParent As Long
          Dim strItem As String
          Dim strDiagnoseInfo As String
          Dim intPicIndex As Integer


1         On Error GoTo GetModifyItem_Error

2         If lngAppforNO = 0 Then Exit Function


          '�����ʽ����<��ĿID,�걾,��Ŀ��,�ڼ�ҳ,�ɿƿ���ID,ִ�п���ID,ҽ������,����ID,�ɼ���������,ִ�п�������,�ɼ�ID,�ɼ���,����,ʱ��id��

3         strSQL = "Select  B.����,a.������־,a.�걾��λ,a.���ID,c.id ִ�п���ID,c.���� ִ�п�������,a.ҽ������,b.���� ��Ŀ��,a.Ƥ�Խ�� ʱ��id " & vbNewLine & _
                  "From ����ҽ����¼ A, ������ĿĿ¼ B,���ű� C" & vbNewLine & _
                  "Where A.������Ŀid = B.Id And ���id Is Not Null and a.ִ�п���ID = C.id  And ������� = [1] and a.����id = [2] order by a.���id,a.���"
4         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ȡ������Ŀid", lngAppforNO, lngPatient)
5         Do Until rsTmp.EOF
6             intPage = getAppPage(rsTmp("����"))
7             intPicIndex = getAppPicIndex(rsTmp("����"))
8             If intPage <> -1 Then

9                 If lngParent <> rsTmp("���ID") Then
10                    strItem = rsTmp("����")
11                    strData = strData & mstrSplieListTag & rsTmp("����") & mstrSplieColTag & rsTmp("�걾��λ") & mstrSplieColTag & rsTmp("��Ŀ��") & mstrSplieColTag & intPage & _
                              mstrSplieColTag & getSampleDept(rsTmp("���ID"), 0) & mstrSplieColTag & rsTmp("ִ�п���ID") & mstrSplieColTag & _
                              GetAppend(rsTmp("���ID")) & mstrSplieColTag & strItem & mstrSplieColTag & getSampleDept(rsTmp("���ID"), 1) & mstrSplieColTag & _
                              rsTmp("ִ�п�������") & mstrSplieColTag & getSampleDept(rsTmp("���ID"), 2) & mstrSplieColTag & getSampleDept(rsTmp("���ID"), 3) & mstrSplieColTag & _
                              GetEntrust(rsTmp("���ID")) & mstrSplieColTag & rsTmp("������־") & mstrSplieColTag & rsTmp("ʱ��id")
12                Else
13                    strData = strData & mstrSplieItemTag & rsTmp("����") & mstrSplieColTag & rsTmp("�걾��λ") & mstrSplieColTag & rsTmp("��Ŀ��") & mstrSplieColTag & intPage & _
                              mstrSplieColTag & getSampleDept(rsTmp("���ID"), 0) & mstrSplieColTag & rsTmp("ִ�п���ID") & mstrSplieColTag & _
                              GetAppend(rsTmp("���ID")) & mstrSplieColTag & strItem & mstrSplieColTag & getSampleDept(rsTmp("���ID"), 1) & mstrSplieColTag & _
                              rsTmp("ִ�п�������") & mstrSplieColTag & getSampleDept(rsTmp("���ID"), 2) & mstrSplieColTag & getSampleDept(rsTmp("���ID"), 3) & mstrSplieColTag & _
                              GetEntrust(rsTmp("���ID")) & mstrSplieColTag & rsTmp("������־") & mstrSplieColTag & rsTmp("ʱ��id")
14                End If

15                lngParent = rsTmp("���ID")
16            End If
17            rsTmp.MoveNext
18        Loop

19        If strData <> "" Then
20            strData = Mid$(strData, Len(mstrSplieListTag) + 1)
21            Call WriterSelHistoryListVSF(0, strData)
22            TabcrlPage.Item(intPicIndex).Selected = True
23            Call WriterSelVSF
24        End If

          '������ʱ��ȡ���
25        If strDiagnose <> "" Then
26            strSQL = "Select �������" & vbNewLine & _
                      "From ������ϼ�¼" & vbNewLine & _
                      "Where ID In (Select * From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)))"

27            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ȡ�����Ϣ", strDiagnose)
28            Do Until rsTmp.EOF
29                strDiagnoseInfo = strDiagnoseInfo & "," & rsTmp("�������")
30                rsTmp.MoveNext
31            Loop
32            If strDiagnoseInfo <> "" Then
33                strDiagnoseInfo = Mid(strDiagnoseInfo, 2)
34            End If
35        End If


36        Exit Function
GetModifyItem_Error:
37        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetModifyItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
38        Err.Clear
End Function


Private Function GetModifySelect(strInData As String)
          '����           ��ȡ�ϴ��޸ĵ����뵥��Ŀ�����޸�(ֱ�Ӵ����ַ�)
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim rsBH As New ADODB.Recordset
          Dim strData As String, i As Integer
          Dim intPage As Integer, lngItemid As Long
          Dim strCaij As String, strKeShi As String
          Dim strItem As String, strCaijiName As String
          Dim strDiagnoseInfo As String, strShijianID As String
          Dim varData As Variant
          Dim varItem As Variant
          Dim intPicIndex As Integer

           '����ĸ�ʽ���������id,ִ�п���id,����ʱ��1,�걾1,����,����,�Ƿ�֢,�ɼ�id,������Ŀid1
1         On Error GoTo GetModifySelect_Error

2         If strInData = "" Then Exit Function
3         varData = Split(strInData, mstrSplieItemTag)

4         For i = LBound(varData) To UBound(varData)
          '�����ʽ����<��ĿID,�걾,��Ŀ��,�ڼ�ҳ,�ɿƿ���ID,ִ�п���ID,ҽ������,����ID,�ɼ���������,ִ�п�������,�ɼ�ID,�ɼ���,���С�
5             varItem = Split(varData(i), mstrSplieColTag)
6             lngItemid = varItem(8)
7             strSQL = "Select  B.����,b.���� ��Ŀ�� " & vbNewLine & _
                  "From ������ĿĿ¼ B" & vbNewLine & _
                  "Where  b.id = [1]"
8             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ȡ������Ŀid", lngItemid)


9             strSQL = "select c.���� from ������ĿĿ¼ c where C.id = [1] "
10            Set rsBH = ComOpenSQL(Sel_His_DB, strSQL, "ȡ������Ŀid", CLng(varItem(7)))
11            strCaijiName = rsBH("����")
12            strSQL = "select c.���� from ���ű� c where C.id = [1] "
13            Set rsBH = ComOpenSQL(Sel_His_DB, strSQL, "ȡ������Ŀid", CLng(varItem(0)))
14            strCaij = rsBH("����")
15            strSQL = "select c.���� from ���ű� c where C.id = [1] "
16            Set rsBH = ComOpenSQL(Sel_His_DB, strSQL, "ȡ������Ŀid", CLng(varItem(1)))
17            strKeShi = rsBH("����")
18            If UBound(varItem) = 8 Then
19              strShijianID = varItem(8)
20            End If
21            intPage = getAppPage(rsTmp("����"))
22            intPicIndex = getAppPicIndex(rsTmp("����"))
23            If intPage <> -1 Then

      '            If strBJ = "" Then
24                    strItem = rsTmp("����")
      '                strdata = strdata & mstrSplieListTag & rsTmp("����") & mstrSplieColTag & rsTmp("�걾��λ") & mstrSplieColTag & rsTmp("��Ŀ��") & mstrSplieColTag & intPage & _
      '                        mstrSplieColTag & getSampleDept(rsTmp("���ID"), 0) & mstrSplieColTag & rsTmp("ִ�п���ID") & mstrSplieColTag & _
      '                        GetAppend(rsTmp("���ID")) & mstrSplieColTag & strItem & mstrSplieColTag & getSampleDept(rsTmp("���ID"), 1) & mstrSplieColTag & _
      '                        rsTmp("ִ�п�������") & mstrSplieColTag & getSampleDept(rsTmp("���ID"), 2) & mstrSplieColTag & getSampleDept(rsTmp("���ID"), 3) & mstrSplieColTag & _
      '                        GetEntrust(rsTmp("���ID")) & mstrSplieColTag & rsTmp("������־")
      '            Else
25                    strData = strData & mstrSplieItemTag & rsTmp("����") & mstrSplieColTag & varItem(3) & mstrSplieColTag & rsTmp("��Ŀ��") & mstrSplieColTag & intPage & _
                              mstrSplieColTag & varItem(0) & mstrSplieColTag & varItem(1) & mstrSplieColTag & _
                             varItem(4) & mstrSplieColTag & strItem & mstrSplieColTag & strCaij & mstrSplieColTag & _
                             strKeShi & mstrSplieColTag & varItem(7) & mstrSplieColTag & strCaijiName & mstrSplieColTag & _
                              varItem(5) & mstrSplieColTag & varItem(6) & mstrSplieColTag & strShijianID
      '            End If

      '            strBJ = "BJ"
26            End If
27        Next

28        If strData <> "" Then
29            strData = Mid$(strData, Len(mstrSplieItemTag) + 1)
30            Call WriterSelHistoryListVSF(0, strData)
31            TabcrlPage.Item(intPicIndex).Selected = True
32            Call WriterSelVSF
33        End If
34        If strInData <> "" And InStr(strInData, "���뵥���<Split2>0<Split2>") > 0 Then
35            strDiagnoseInfo = Split(Split(strInData, "<Split A>")(4), "���뵥���<Split2>0<Split2><Split2>")(1)
36        End If


37        Exit Function
GetModifySelect_Error:
38        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetModifySelect)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
39        Err.Clear

End Function

Private Function ChoseItem(strNO As String)
          '����           ������Ŀ����õ���Ŀ����ҳ��Index
          Dim intPage As Integer
          Dim intRow As Integer
          Dim intCol As Integer
          Dim aItem() As String
          Dim intPicIndex As Integer
1         On Error GoTo ChoseItem_Error

2         intPicIndex = getAppPicIndex(strNO)
3         For intPage = 0 To Me.vsfItem.Count - 1
4             With Me.vsfItem(intPage)
5                 For intRow = 0 To .Rows - 1
6                     For intCol = 0 To .Cols - 1
7                         If .Cell(flexcpData, intRow, intCol, intRow, intCol) <> "" Then
8                             aItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
9                             If aItem(0) = strNO Then
10                                Call GetItems(intPage, 1, intRow, intCol)
11                                TabcrlPage.Item(intPicIndex).Selected = True
12                                Call WriterSelVSF
13                                Exit Function
14                            End If
15                        End If
16                    Next
17                Next
18            End With
19        Next


20        Exit Function
ChoseItem_Error:
21        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(ChoseItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
22        Err.Clear

End Function


Private Function getAppPage(strNO As String) As Integer
          '����           ������Ŀ����õ���Ŀ����ҳ��Index
          Dim intPage As Integer
          Dim intRow As Integer
          Dim intCol As Integer
          Dim aItem() As String



1         On Error GoTo getAppPage_Error

2         For intPage = 0 To Me.vsfItem.Count - 1
3             With Me.vsfItem(intPage)
4                 For intRow = 0 To .Rows - 1
5                     For intCol = 0 To .Cols - 1
6                         If .Cell(flexcpData, intRow, intCol, intRow, intCol) <> "" Then
7                             aItem = Split(.Cell(flexcpData, intRow, intCol, intRow, intCol), mstrSplieColTag)
8                             If aItem(0) = strNO Then
9                                 getAppPage = intPage
10                                Exit Function
11                            End If
12                        End If
13                    Next
14                Next
15            End With
16        Next
17        getAppPage = -1


18        Exit Function
getAppPage_Error:
19        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(getAppPage)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
20        Err.Clear
End Function
Private Function getAppPicIndex(strNO As String) As Integer
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim i As Integer

1         On Error GoTo getAppPicIndex_Error
2         If gUserInfo.NodeNo <> "-" Then
3             strSQL = "select t.���� from �������뵥 t ,�������뵥��ϸ  a,���������Ŀ b " & vbNewLine & _
                          "where  a.���뵥id= t.id and a.���id =b.id and b.���Ʊ��� = [1] and (b.վ��=[2] or b.վ�� is null)"
4         Else
5             strSQL = "select t.���� from �������뵥 t ,�������뵥��ϸ  a,���������Ŀ b " & vbNewLine & _
                          "where  a.���뵥id= t.id and a.���id =b.id and b.���Ʊ��� = [1]"
6         End If
7         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "ҽ������", strNO, gUserInfo.NodeNo)
8         If rsTmp.RecordCount > 0 Then
9             For i = 0 To TabcrlPage.ItemCount - 1
10                If TabcrlPage.Item(i).Caption = rsTmp("����") Then
11                    getAppPicIndex = i
12                End If
13            Next
14        End If


15        Exit Function
getAppPicIndex_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(getAppPicIndex)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear

End Function
Private Function GetEntrust(lngID As Long) As String
          '����       ͨ��ҽ��ID��ȡ����
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset

1         On Error GoTo GetEntrust_Error

2         strSQL = "select a.ҽ������ from ����ҽ����¼ a where a.id = [1] "
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ҽ������", lngID)

4         If rsTmp.RecordCount > 0 Then
5             GetEntrust = rsTmp("ҽ������") & ""
6         Else
7             GetEntrust = ""
8         End If



9         Exit Function
GetEntrust_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(GetEntrust)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear

End Function
Private Function getSampleDept(lngID As Long, intType As Integer) As String
          '����       ͨ��ҽ��ID�ɼ�����ID�����ƺ���ĿID
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset

1         On Error GoTo getSampleDept_Error

2         strSQL = "select b.id,b.���� �ɼ�����,������ĿID,c.���� from ����ҽ����¼ a,���ű� b,������ĿĿ¼ c where a.������Ŀid = c.id and a.ִ�п���ID = b.id and a.id = [1] "
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�ɼ�����id", lngID)
4         If rsTmp.RecordCount > 0 Then
5             If intType = 0 Then
6                 getSampleDept = rsTmp("id")
7             ElseIf intType = 1 Then
8                 getSampleDept = rsTmp("�ɼ�����")
9             ElseIf intType = 2 Then
10                getSampleDept = rsTmp("������ĿID")
11            Else
12                getSampleDept = rsTmp("����")
13            End If
14        Else
15            If intType = 0 Then
16                getSampleDept = 0
17            Else
18                getSampleDept = ""
19            End If
20        End If


21        Exit Function
getSampleDept_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(getSampleDept)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear
End Function

Private Function GetAppend(lngID As Long) As String
    '����       ͨ��ҽ��IDȡҽ������
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strTag1 As String
    Dim strTag2 As String

    strTag1 = "<Split1>"
    strTag2 = "<Split2>"


    '��ʽ="��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>..."

    strSQL = "select ҽ��id,��Ŀ,����,����,Ҫ��ID,���� from ����ҽ������ a where  a.ҽ��id = [1] "
    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ҽ������", lngID)

    Do Until rsTmp.EOF
        If GetAppend = "" Then
            GetAppend = rsTmp("��Ŀ") & strTag2 & rsTmp("����") & strTag2 & rsTmp("Ҫ��ID") & strTag2 & rsTmp("����")
        Else
            GetAppend = GetAppend & strTag1 & rsTmp("��Ŀ") & strTag2 & rsTmp("����") & strTag2 & rsTmp("Ҫ��ID") & strTag2 & rsTmp("����")
        End If
        rsTmp.MoveNext
    Loop
'    If GetAppend <> "" Then
'        GetAppend = Mid$(GetAppend, Len(strTag1) + 1)
'    End If
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/1/3
'��    ��:�����п�
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub SetColWith(ByVal objVSF As VSFlexGrid)
    Dim lngColWidth As Long

    With objVSF
        .Width = Me.picItemRight.Width
        lngColWidth = .Width / 3 - 100
        .AutoSize 0, .Cols - 1
        If .ColWidth(0) + .ColWidth(1) + .ColWidth(2) < .Width Then
            .ColWidth(0) = lngColWidth
            .ColWidth(1) = lngColWidth
        End If
    End With
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-04-19
'��    ��:  ��ʾ���Ʋο�
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Private Sub ShowClincHelp()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim objAdvice As Object
          Dim strItemID As String
          Dim lngIndex As Long
          Dim strItemCode As String
          Dim lngRow As Long
          Dim lngCol As Long
          Dim i As Integer
          Dim blnContinue As Boolean

          '��ȡ��ǰ��ʾ��vsf
1         On Error GoTo ShowClincHelp_Error

2         For i = 0 To Me.vsfItem.Count - 1
3             If Me.vsfItem(i).Visible = True Then lngIndex = i
4         Next

          '��ȡ��ǰVSF�е����Ʊ���
5         With Me.vsfItem(lngIndex)
6             For lngRow = 0 To .Rows - 1
7                 For lngCol = 0 To .Cols - 1
8                     If .Cell(flexcpData, lngRow, lngCol) <> "" Then
9                         strItemCode = strItemCode & "," & Split(.Cell(flexcpData, lngRow, lngCol), mstrSplieColTag)(0)
10                    End If
11                Next
12            Next
13            If Left(strItemCode, 1) = "," Then strItemCode = Mid(strItemCode, 2)
14        End With

          'ͨ�����Ʊ����ѯ������ĿID
15        strSQL = "Select /*+cardinality(b,10)*/" & vbCrLf & _
                   "f_List2str(Cast(Collect(a.ID || '') As t_Strlist)) ID" & vbCrLf & _
                   " From ������ĿĿ¼ A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                   " Where A.���� = B.Column_Value"
16        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", strItemCode)
17        If Not rsTmp.EOF Then strItemID = rsTmp("ID") & ""

          '�ȵ���plugin�еĽӿڣ��ӿڵ���ʧ���ٵ���zlPublicAdvice�еĽӿ�
18        If VerCompare(gSysInfo.VersionHIS, "10.35.130") <> -1 Then
19            If CreatePlugInOK(2500, 2) Then
20                On Error Resume Next
21                blnContinue = gobjPlugIn.ShowClinicHelp(Me.hWnd, 0, mintPatientType, mlngPatientID, mlngPatientPage, strItemID)
22                Call zlPlugInErrH(Err, "ExecuteFunc")
23                Err.Clear: On Error GoTo 0
24            End If
25        End If


          '���ýӿ�
26        If Not blnContinue Then
27            If objAdvice Is Nothing Then
28                Set objAdvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
29                If Not objAdvice Is Nothing Then
30                    On Error Resume Next
31                    Call objAdvice.ShowClincHelp(1, Me, 0, False, strItemID)
32                    If Err.Number = 438 Then
33                        MsgBox "HIS�汾����", vbInformation, gSysInfo.AppName
34                        Exit Sub
35                    End If
36                End If
37            End If
38        End If


39        Exit Sub
ShowClincHelp_Error:
40        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(ShowClincHelp)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
41        Err.Clear
End Sub


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-09-09
'��    ��:  ����Ƿ������Ҫ¼���Ҫ��
'��    ��:
'           lngGroupID          ���ID
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Function CheckRefItem(ByVal lngGroupId As Long) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo CheckRefItem_Error

3         strSQL = "Select Distinct d.Ҫ����" & vbCrLf & _
                   " From �������ָ�� A, ����ָ��ο���Χ B, ����ο�Ҫ�ض��� C, ����ָ��ο�Ҫ�� D" & vbCrLf & _
                   " Where a.��Ŀid = b.ָ��id And b.id = c.�ο�id And c.Ҫ��id = d.id And a.���id = [1]" & vbCrLf & _
                   " And d.�����ֶ��� Is Null And c.�������� Is Not Null"
4         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ָ��ο�Ҫ��", lngGroupId)
5         Do While Not rsTmp.EOF
6             If InStr(",����ʱ��,", "," & rsTmp("Ҫ����") & ",") = 0 Then    '����ʱ���ڳ����������⴦������Ҫҽ��������д
'7                 mrsReference.Filter = "Ҫ��ID=" & rsTmp("ID") & " and �������� <>''"
'8                 If rsTmp.RecordCount > 0 Then
9                     CheckRefItem = True
10                    Exit Function
'11                End If
12            End If
13            rsTmp.MoveNext
14        Loop


15        Exit Function
CheckRefItem_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBill", "ִ��(CheckRefItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear
End Function



