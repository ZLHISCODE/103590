VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Object = "*\A..\ZlPatiAddress\ZlPatiAddress.vbp"
Begin VB.Form frmHosReg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˵Ǽ�"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHosReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicHealth 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9750
      Left            =   15240
      ScaleHeight     =   9750
      ScaleMode       =   0  'User
      ScaleWidth      =   18813.26
      TabIndex        =   188
      Top             =   0
      Width           =   15015
      Begin VB.Frame fraCertificate 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   105
         Left            =   1230
         TabIndex        =   217
         Top             =   6255
         Width           =   13530
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   105
            Left            =   0
            TabIndex        =   219
            Top             =   6195
            Width           =   13530
         End
      End
      Begin VB.CommandButton cmdMedicalWarning 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   14595
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   30
         Width           =   390
      End
      Begin VB.ComboBox cboBloodType 
         Height          =   360
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   193
         Top             =   45
         Width           =   1410
      End
      Begin VB.ComboBox cboBH 
         Height          =   360
         Left            =   3135
         Style           =   2  'Dropdown List
         TabIndex        =   194
         Top             =   45
         Width           =   1410
      End
      Begin VB.TextBox txtMedicalWarning 
         Height          =   360
         Left            =   5625
         Locked          =   -1  'True
         TabIndex        =   195
         Top             =   45
         Width           =   9375
      End
      Begin VB.TextBox txtOtherWaring 
         Height          =   360
         Left            =   1665
         MaxLength       =   100
         TabIndex        =   197
         Top             =   495
         Width           =   13335
      End
      Begin VB.Frame frameLinkMan 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   105
         Left            =   1380
         TabIndex        =   192
         Top             =   4560
         Width           =   13530
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   105
         Left            =   1215
         TabIndex        =   191
         Top             =   7860
         Width           =   13695
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   105
         Left            =   1185
         TabIndex        =   190
         Top             =   2730
         Width           =   13725
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   105
         Left            =   1260
         TabIndex        =   189
         Top             =   1005
         Width           =   13710
      End
      Begin VSFlex8Ctl.VSFlexGrid vsLinkMan 
         Height          =   1305
         Left            =   150
         TabIndex        =   200
         Top             =   4710
         Width           =   14775
         _cx             =   26061
         _cy             =   2302
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   360
         RowHeightMax    =   360
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsOtherInfo 
         Height          =   1460
         Left            =   150
         TabIndex        =   202
         Top             =   8170
         Width           =   14775
         _cx             =   26061
         _cy             =   2575
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   360
         RowHeightMax    =   360
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInoculate 
         Height          =   1305
         Left            =   150
         TabIndex        =   199
         Top             =   2985
         Width           =   14820
         _cx             =   26141
         _cy             =   2302
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   360
         RowHeightMax    =   360
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDrug 
         Height          =   1245
         Left            =   150
         TabIndex        =   198
         Top             =   1275
         Width           =   14820
         _cx             =   26141
         _cy             =   2196
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   360
         RowHeightMax    =   360
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCertificate 
         Height          =   1215
         Left            =   165
         TabIndex        =   201
         Top             =   6510
         Width           =   14820
         _cx             =   26141
         _cy             =   2143
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   360
         RowHeightMax    =   360
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblCertificate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "֤����Ϣ"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -255
         TabIndex        =   218
         Top             =   6195
         Width           =   1860
      End
      Begin VB.Label lblBloodType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Ѫ��"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   405
         TabIndex        =   210
         Top             =   90
         Width           =   1020
      End
      Begin VB.Label lblRH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "RH"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2475
         TabIndex        =   209
         Top             =   90
         Width           =   885
      End
      Begin VB.Label lblMedicalWarning 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "ҽѧ��ʾ"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4155
         TabIndex        =   208
         Top             =   90
         Width           =   1860
      End
      Begin VB.Label lblOtherWaring 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "����ҽѧ��ʾ"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -105
         TabIndex        =   207
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label lblLinkman 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "��ϵ����Ϣ"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -165
         TabIndex        =   206
         Top             =   4410
         Width           =   1860
      End
      Begin VB.Label lblOtherInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "������Ϣ"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   -285
         TabIndex        =   205
         Top             =   7845
         Width           =   1860
      End
      Begin VB.Label lblInoculate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "�������"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -330
         TabIndex        =   204
         Top             =   2685
         Width           =   1860
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "������Ӧ"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -315
         TabIndex        =   203
         Top             =   975
         Width           =   1860
      End
   End
   Begin VB.PictureBox PicBaseInfo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9795
      Left            =   300
      ScaleHeight     =   9795
      ScaleWidth      =   15120
      TabIndex        =   100
      Top             =   285
      Width           =   15120
      Begin VB.PictureBox pic�ſ� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   870
         Left            =   0
         ScaleHeight     =   870
         ScaleWidth      =   15120
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   8880
         Width           =   15120
         Begin VB.Frame fra�ſ� 
            Caption         =   "��������Ϣ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   825
            Left            =   30
            TabIndex        =   180
            Top             =   0
            Width           =   15000
            Begin VB.TextBox txtAudi 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   6360
               MaxLength       =   10
               PasswordChar    =   "*"
               TabIndex        =   90
               Top             =   400
               Width           =   1750
            End
            Begin VB.ComboBox cbo�������� 
               Height          =   360
               Left            =   12645
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Top             =   400
               Width           =   1845
            End
            Begin VB.TextBox txt���� 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H00C00000&
               Height          =   360
               Left            =   9480
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   400
               Width           =   1695
            End
            Begin VB.TextBox txtPass 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   4005
               MaxLength       =   10
               PasswordChar    =   "*"
               TabIndex        =   89
               Top             =   400
               Width           =   1750
            End
            Begin VB.TextBox txt���� 
               BackColor       =   &H00EBFFFF&
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   1590
               PasswordChar    =   "*"
               TabIndex        =   88
               Top             =   400
               Width           =   1750
            End
            Begin VB.CheckBox chk���� 
               Caption         =   "����"
               Height          =   360
               Left            =   11505
               TabIndex        =   92
               Top             =   400
               Width           =   900
            End
            Begin MSComctlLib.TabStrip tabCardMode 
               Height          =   315
               Left            =   120
               TabIndex        =   181
               Top             =   0
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               Style           =   2
               TabFixedHeight  =   526
               HotTracking     =   -1  'True
               Separators      =   -1  'True
               TabMinWidth     =   882
               _Version        =   393216
               BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
                  NumTabs         =   2
                  BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "�����շ�(&1)"
                     Key             =   "CardFee"
                     ImageVarType    =   2
                  EndProperty
                  BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "�󶨿���(&2)"
                     Key             =   "CardBind"
                     ImageVarType    =   2
                  EndProperty
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label lbl������ 
               Height          =   255
               Left            =   12420
               TabIndex        =   214
               Top             =   0
               Width           =   1575
            End
            Begin VB.Label lbl��֤ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��֤"
               Height          =   240
               Left            =   5835
               TabIndex        =   185
               Top             =   460
               Width           =   480
            End
            Begin VB.Label lbl��� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   8880
               TabIndex        =   184
               Top             =   460
               Width           =   480
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   240
               Left            =   3435
               TabIndex        =   183
               Top             =   465
               Width           =   480
            End
            Begin VB.Label lbl���� 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "����"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   930
               TabIndex        =   182
               Top             =   450
               Width           =   510
            End
         End
      End
      Begin VB.PictureBox picԤ�� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1170
         Left            =   0
         ScaleHeight     =   1170
         ScaleWidth      =   15120
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   7725
         Width           =   15120
         Begin VB.Frame fraԤ�� 
            Caption         =   "��סԺԤ����Ϣ��"
            ForeColor       =   &H00C00000&
            Height          =   1160
            Left            =   30
            TabIndex        =   169
            Top             =   0
            Width           =   15000
            Begin VB.CheckBox chk��λ�ɿ� 
               Caption         =   "��λ�ɿ�"
               Height          =   360
               Left            =   12480
               TabIndex        =   84
               Top             =   375
               Width           =   1320
            End
            Begin VB.TextBox txt�ʺ� 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   9480
               MaxLength       =   50
               TabIndex        =   87
               Top             =   735
               Width           =   5025
            End
            Begin VB.TextBox txtԤ���� 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00EBFFFF&
               ForeColor       =   &H00C00000&
               Height          =   360
               Left            =   3750
               MaxLength       =   12
               TabIndex        =   81
               Top             =   375
               Width           =   1335
            End
            Begin VB.ComboBox cboԤ������ 
               Height          =   360
               Left            =   6330
               Style           =   2  'Dropdown List
               TabIndex        =   82
               Top             =   375
               Width           =   1770
            End
            Begin VB.TextBox txt������� 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   9480
               MaxLength       =   30
               TabIndex        =   83
               Top             =   375
               Width           =   2445
            End
            Begin VB.TextBox txtFact 
               Height          =   360
               Left            =   1590
               MaxLength       =   50
               TabIndex        =   80
               Top             =   375
               Width           =   1470
            End
            Begin VB.TextBox txt�ɿλ 
               Height          =   360
               Left            =   1590
               MaxLength       =   50
               TabIndex        =   85
               Top             =   735
               Width           =   2745
            End
            Begin VB.TextBox txt������ 
               Height          =   360
               Left            =   5280
               MaxLength       =   50
               TabIndex        =   86
               Top             =   735
               Width           =   2805
            End
            Begin VB.Label lblYBMoney 
               AutoSize        =   -1  'True
               Caption         =   "�����ʻ����:"
               ForeColor       =   &H00C00000&
               Height          =   240
               Left            =   2055
               TabIndex        =   178
               Top             =   0
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.Label lblNote 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ժҪ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   825
               TabIndex        =   177
               Top             =   1605
               Width           =   480
            End
            Begin VB.Label lblMoney 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   3180
               TabIndex        =   176
               Top             =   435
               Width           =   480
            End
            Begin VB.Label lblCode 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�������"
               Height          =   240
               Left            =   8400
               TabIndex        =   175
               Top             =   435
               Width           =   960
            End
            Begin VB.Label lblStyle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ɿʽ"
               Height          =   240
               Left            =   5205
               TabIndex        =   174
               Top             =   435
               Width           =   960
            End
            Begin VB.Label lblFact 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ʵ��Ʊ��"
               Height          =   240
               Left            =   510
               TabIndex        =   173
               Top             =   435
               Width           =   960
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ɿλ"
               Height          =   240
               Left            =   510
               TabIndex        =   172
               Top             =   795
               Width           =   960
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               Height          =   240
               Left            =   4440
               TabIndex        =   171
               Top             =   795
               Width           =   720
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ʺ�"
               Height          =   240
               Left            =   8880
               TabIndex        =   170
               Top             =   795
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox pic��Ժ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2355
         Left            =   0
         ScaleHeight     =   2355
         ScaleWidth      =   15120
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   5385
         Width           =   15120
         Begin VB.Frame fra��Ժ 
            Caption         =   "��סԺ��Ϣ��"
            ForeColor       =   &H00C00000&
            Height          =   2325
            Left            =   0
            TabIndex        =   151
            Top             =   0
            Width           =   15000
            Begin VB.CommandButton cmdת�� 
               Caption         =   "��"
               Height          =   300
               Left            =   14160
               TabIndex        =   222
               TabStop         =   0   'False
               Top             =   1580
               Width           =   300
            End
            Begin VB.TextBox txtת�� 
               Height          =   360
               Left            =   12720
               TabIndex        =   221
               Top             =   1550
               Width           =   1785
            End
            Begin VB.TextBox txtTimes 
               Height          =   360
               Left            =   12480
               MaxLength       =   3
               TabIndex        =   66
               Tag             =   "1"
               Text            =   "1"
               Top             =   345
               Width           =   460
            End
            Begin VB.ComboBox cbo��Ժ���� 
               Height          =   360
               Left            =   9480
               Style           =   2  'Dropdown List
               TabIndex        =   74
               Top             =   1125
               Width           =   2445
            End
            Begin VB.CheckBox chk����Ժ 
               Caption         =   "����Ժ"
               Height          =   360
               Left            =   12120
               TabIndex        =   70
               ToolTipText     =   "�ٴ���ס��ͬ���ƿ�Ŀ������ٴ�����"
               Top             =   735
               Width           =   1095
            End
            Begin VB.ComboBox cbo��Ժ���� 
               Height          =   360
               Left            =   5520
               TabIndex        =   64
               Top             =   330
               Width           =   2565
            End
            Begin VB.ComboBox cbo��Ժ���� 
               Height          =   360
               Left            =   5520
               Style           =   2  'Dropdown List
               TabIndex        =   73
               Top             =   1125
               Width           =   2565
            End
            Begin VB.ComboBox cbo��Ժ��ʽ 
               Height          =   360
               Left            =   9480
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   1550
               Width           =   2445
            End
            Begin VB.ComboBox cbo��λ 
               Height          =   360
               ItemData        =   "frmHosReg.frx":0442
               Left            =   9480
               List            =   "frmHosReg.frx":0444
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   345
               Width           =   2445
            End
            Begin VB.CheckBox chk����Ժת�� 
               Caption         =   "����Ժת��"
               Height          =   360
               Left            =   12120
               TabIndex        =   75
               Top             =   1125
               Width           =   1680
            End
            Begin VB.ComboBox cbo����ҽʦ 
               Height          =   360
               IMEMode         =   2  'OFF
               Left            =   5520
               TabIndex        =   68
               Top             =   720
               Width           =   2565
            End
            Begin VB.CheckBox chk��� 
               Caption         =   "�Ƿ����"
               Height          =   360
               Left            =   13200
               TabIndex        =   71
               Top             =   735
               Width           =   1380
            End
            Begin VB.TextBox txt��ע 
               Height          =   360
               Left            =   9480
               MaxLength       =   100
               TabIndex        =   79
               Top             =   1905
               Width           =   5025
            End
            Begin VB.TextBox txt��ҽ��� 
               Height          =   360
               Left            =   1590
               MaxLength       =   200
               TabIndex        =   78
               Top             =   1905
               Width           =   6495
            End
            Begin VB.TextBox txt������� 
               Height          =   360
               Left            =   1590
               MaxLength       =   200
               TabIndex        =   76
               Top             =   1550
               Width           =   6495
            End
            Begin VB.ComboBox cbo����ȼ� 
               Height          =   360
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   735
               Width           =   2565
            End
            Begin VB.ComboBox cboסԺĿ�� 
               Height          =   360
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Top             =   1125
               Width           =   2565
            End
            Begin VB.ComboBox cbo��Ժ���� 
               Height          =   360
               Left            =   1590
               TabIndex        =   63
               Text            =   "cbo��Ժ����"
               Top             =   345
               Width           =   2565
            End
            Begin MSMask.MaskEdBox txt��Ժʱ�� 
               Height          =   360
               Left            =   9480
               TabIndex        =   69
               Top             =   735
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   635
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   16
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "yyyy-MM-dd hh:mm"
               Mask            =   "####-##-## ##:##"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtPages 
               Enabled         =   0   'False
               Height          =   360
               Left            =   12480
               MaxLength       =   3
               TabIndex        =   213
               Tag             =   "1"
               Text            =   "1"
               Top             =   345
               Visible         =   0   'False
               Width           =   460
            End
            Begin VB.Label lblInFrom 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ת��"
               Enabled         =   0   'False
               Height          =   240
               Left            =   12120
               TabIndex        =   215
               Top             =   1610
               Width           =   600
            End
            Begin VB.Label lblBedInfo 
               AutoSize        =   -1  'True
               Caption         =   "������Ժ����λ��Ϣ"
               ForeColor       =   &H00C00000&
               Height          =   240
               Left            =   1560
               TabIndex        =   167
               Top             =   0
               Width           =   2160
            End
            Begin VB.Label lblTimes 
               Caption         =   "��      ��סԺ"
               Height          =   255
               Left            =   12120
               TabIndex        =   166
               Top             =   405
               Width           =   1785
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ����"
               Height          =   240
               Left            =   8400
               TabIndex        =   165
               Top             =   1185
               Width           =   960
            End
            Begin VB.Label lbl��Ժ���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ����"
               Height          =   240
               Left            =   4380
               TabIndex        =   164
               Top             =   405
               Width           =   960
            End
            Begin VB.Label lbl��ҽ��� 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ҽ���"
               Height          =   240
               Left            =   510
               TabIndex        =   163
               Top             =   1965
               Width           =   960
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ע"
               Height          =   240
               Left            =   8880
               TabIndex        =   162
               Top             =   1965
               Width           =   480
            End
            Begin VB.Label lbl������� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�������"
               Height          =   240
               Left            =   510
               TabIndex        =   160
               Top             =   1610
               Width           =   960
            End
            Begin VB.Label lbl��Ժ���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ����"
               Height          =   240
               Left            =   510
               TabIndex        =   159
               Top             =   405
               Width           =   960
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ����"
               Height          =   240
               Left            =   4380
               TabIndex        =   158
               Top             =   1185
               Width           =   960
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ��ʽ"
               Height          =   240
               Left            =   8400
               TabIndex        =   157
               Top             =   1610
               Width           =   960
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺĿ��"
               Height          =   240
               Left            =   510
               TabIndex        =   156
               Top             =   1185
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ȼ�"
               Height          =   240
               Left            =   510
               TabIndex        =   155
               Top             =   795
               Width           =   960
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժʱ��"
               Height          =   240
               Left            =   8400
               TabIndex        =   154
               Top             =   795
               Width           =   960
            End
            Begin VB.Label lbl��λ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ����"
               Height          =   240
               Left            =   8400
               TabIndex        =   153
               Top             =   405
               Width           =   960
            End
            Begin VB.Label lbl����ҽʦ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ҽʦ"
               Height          =   240
               Left            =   4380
               TabIndex        =   152
               Top             =   795
               Width           =   960
            End
         End
      End
      Begin VB.PictureBox pic���� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5400
         Left            =   0
         ScaleHeight     =   5400
         ScaleWidth      =   15120
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   15
         Width           =   15120
         Begin VB.Frame fra���� 
            Caption         =   "��������Ϣ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   5340
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   15000
            Begin VB.PictureBox pic���� 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   45
               ScaleHeight     =   360
               ScaleWidth      =   14700
               TabIndex        =   114
               Top             =   4905
               Width           =   14700
               Begin VB.TextBox txt������ 
                  Alignment       =   1  'Right Justify
                  ForeColor       =   &H00C00000&
                  Height          =   360
                  Left            =   4605
                  MaxLength       =   10
                  TabIndex        =   59
                  Top             =   30
                  Width           =   1305
               End
               Begin VB.TextBox txt������ 
                  Height          =   360
                  Left            =   1530
                  MaxLength       =   100
                  TabIndex        =   57
                  Top             =   30
                  Width           =   1605
               End
               Begin VB.CheckBox chkUnlimit 
                  Caption         =   "����"
                  Height          =   255
                  Left            =   3210
                  TabIndex        =   58
                  ToolTipText     =   "���޵�����ʱ�������õ���ʱ��"
                  Top             =   90
                  Width           =   795
               End
               Begin VB.TextBox txtReason 
                  Height          =   360
                  Left            =   12285
                  MaxLength       =   50
                  TabIndex        =   62
                  Top             =   30
                  Width           =   2145
               End
               Begin VB.CheckBox chk��ʱ���� 
                  Caption         =   "��ʱ����"
                  Height          =   360
                  Left            =   9780
                  TabIndex        =   61
                  Top             =   30
                  Width           =   1280
               End
               Begin MSComCtl2.DTPicker dtp����ʱ�� 
                  Height          =   360
                  Left            =   7065
                  TabIndex        =   60
                  Top             =   30
                  Width           =   2610
                  _ExtentX        =   4604
                  _ExtentY        =   635
                  _Version        =   393216
                  Enabled         =   0   'False
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
                  CheckBox        =   -1  'True
                  CustomFormat    =   "yyyy-MM-dd HH:mm"
                  Format          =   93323267
                  CurrentDate     =   38915.6041666667
               End
               Begin VB.Label lbl������ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���"
                  Height          =   240
                  Left            =   4050
                  TabIndex        =   118
                  Top             =   90
                  Width           =   480
               End
               Begin VB.Label lbl������ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "������"
                  Height          =   240
                  Left            =   660
                  TabIndex        =   117
                  Top             =   90
                  Width           =   720
               End
               Begin VB.Label lbl����ʱ�� 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����ʱ��"
                  Height          =   240
                  Left            =   6015
                  TabIndex        =   116
                  Top             =   90
                  Width           =   960
               End
               Begin VB.Label lbl����ԭ�� 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����ԭ��"
                  Height          =   240
                  Left            =   11190
                  TabIndex        =   115
                  Top             =   90
                  Width           =   960
               End
            End
            Begin VB.ComboBox cboIDNumber 
               Height          =   360
               Left            =   10080
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   630
               Width           =   1350
            End
            Begin VB.CommandButton cmd���� 
               Caption         =   "��"
               Height          =   300
               Left            =   14160
               TabIndex        =   34
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   2610
               Width           =   300
            End
            Begin VB.CommandButton cmd���� 
               Caption         =   "��"
               Height          =   300
               Left            =   10965
               TabIndex        =   40
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   3000
               Width           =   300
            End
            Begin VB.CommandButton cmd�����ص� 
               Caption         =   "��"
               Height          =   300
               Left            =   5610
               TabIndex        =   37
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   3000
               Width           =   315
            End
            Begin VB.CommandButton cmd���ڵ�ַ 
               Caption         =   "��"
               Height          =   300
               Left            =   5625
               TabIndex        =   30
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   2610
               Width           =   300
            End
            Begin VB.CommandButton cmd��ͥ��ַ 
               Caption         =   "��"
               Height          =   300
               Left            =   5625
               TabIndex        =   25
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   2220
               Width           =   300
            End
            Begin VB.CommandButton cmdName 
               Caption         =   "��"
               Height          =   300
               Left            =   9150
               TabIndex        =   212
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ���F3"
               Top             =   270
               Width           =   300
            End
            Begin VB.CommandButton cmdSelectNO 
               Caption         =   "��"
               Height          =   300
               Left            =   5625
               TabIndex        =   211
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ�:F8 ȱ��ѡ��"
               Top             =   270
               Width           =   300
            End
            Begin VB.TextBox txt���ڵ�ַ 
               Height          =   360
               Left            =   1590
               MaxLength       =   100
               TabIndex        =   29
               Top             =   2580
               Width           =   4335
            End
            Begin VB.TextBox txt���ڵ�ַ�ʱ� 
               Height          =   360
               Left            =   9555
               MaxLength       =   6
               TabIndex        =   32
               Top             =   2580
               Width           =   1725
            End
            Begin VB.TextBox txt���� 
               Height          =   360
               Left            =   11805
               MaxLength       =   50
               TabIndex        =   33
               Top             =   2580
               Width           =   2685
            End
            Begin VB.TextBox txt���� 
               Height          =   360
               Left            =   7770
               MaxLength       =   50
               TabIndex        =   39
               Top             =   2970
               Width           =   3525
            End
            Begin VB.TextBox txt�����ص� 
               Height          =   360
               Left            =   1590
               MaxLength       =   100
               TabIndex        =   36
               Top             =   2970
               Width           =   4350
            End
            Begin VB.PictureBox picUnUseful 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1545
               Left            =   45
               ScaleHeight     =   1545
               ScaleWidth      =   14925
               TabIndex        =   104
               Tag             =   "0"
               Top             =   3360
               Width           =   14925
               Begin VB.TextBox txtMobile 
                  Height          =   360
                  Left            =   12720
                  MaxLength       =   20
                  TabIndex        =   48
                  Top             =   390
                  Width           =   1725
               End
               Begin VB.TextBox txt��ϵ�����֤�� 
                  Height          =   360
                  IMEMode         =   3  'DISABLE
                  Left            =   1530
                  MaxLength       =   18
                  TabIndex        =   53
                  Top             =   1170
                  Width           =   4365
               End
               Begin VB.TextBox txtLinkManInfo 
                  BackColor       =   &H80000004&
                  Enabled         =   0   'False
                  Height          =   360
                  Left            =   3870
                  MaxLength       =   100
                  TabIndex        =   50
                  Top             =   780
                  Width           =   2025
               End
               Begin VB.CommandButton cmd������λ 
                  Caption         =   "��"
                  Height          =   300
                  Left            =   5565
                  TabIndex        =   43
                  TabStop         =   0   'False
                  ToolTipText     =   "�ȼ���F3"
                  Top             =   30
                  Width           =   315
               End
               Begin VB.CommandButton cmd��ϵ�˵�ַ 
                  Caption         =   "��"
                  Height          =   300
                  Left            =   14145
                  TabIndex        =   55
                  TabStop         =   0   'False
                  ToolTipText     =   "�ȼ���F3"
                  Top             =   1200
                  Width           =   300
               End
               Begin VB.TextBox txt��ϵ�˵�ַ 
                  Height          =   360
                  Left            =   7725
                  MaxLength       =   100
                  TabIndex        =   54
                  Top             =   1170
                  Width           =   6750
               End
               Begin VB.TextBox txt������λ 
                  Height          =   360
                  Left            =   1545
                  MaxLength       =   100
                  TabIndex        =   42
                  Top             =   0
                  Width           =   4350
               End
               Begin VB.TextBox txt��λ������ 
                  Height          =   360
                  Left            =   1545
                  MaxLength       =   50
                  TabIndex        =   46
                  Top             =   390
                  Width           =   4350
               End
               Begin VB.TextBox txt��λ�ʺ� 
                  Height          =   360
                  Left            =   7725
                  MaxLength       =   50
                  TabIndex        =   47
                  Top             =   390
                  Width           =   3525
               End
               Begin VB.TextBox txt��ϵ������ 
                  Height          =   360
                  Left            =   7725
                  MaxLength       =   64
                  TabIndex        =   51
                  Top             =   780
                  Width           =   3525
               End
               Begin VB.ComboBox cbo��ϵ�˹�ϵ 
                  Height          =   360
                  Left            =   1545
                  Style           =   2  'Dropdown List
                  TabIndex        =   49
                  Top             =   780
                  Width           =   2310
               End
               Begin VB.TextBox txt��ϵ�˵绰 
                  Height          =   360
                  Left            =   12720
                  MaxLength       =   20
                  TabIndex        =   52
                  Top             =   780
                  Width           =   1725
               End
               Begin VB.TextBox txt��λ�ʱ� 
                  Height          =   360
                  Left            =   12720
                  MaxLength       =   6
                  TabIndex        =   45
                  Top             =   0
                  Width           =   1725
               End
               Begin VB.TextBox txt��λ�绰 
                  Height          =   360
                  Left            =   7725
                  MaxLength       =   20
                  TabIndex        =   44
                  Top             =   0
                  Width           =   3525
               End
               Begin ZlPatiAddress.PatiAddress PatiAddress 
                  Height          =   360
                  Index           =   5
                  Left            =   7725
                  TabIndex        =   56
                  Tag             =   "��ϵ�˵�ַ"
                  Top             =   1170
                  Visible         =   0   'False
                  Width           =   6750
                  _ExtentX        =   11906
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "����"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MaxLength       =   100
               End
               Begin VB.Label lblMobile 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�ֻ���"
                  Height          =   240
                  Left            =   11970
                  TabIndex        =   220
                  Top             =   450
                  Width           =   720
               End
               Begin VB.Label lbl��ϵ�����֤ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ϵ�����֤"
                  Height          =   240
                  Left            =   0
                  TabIndex        =   216
                  Top             =   1230
                  Width           =   1440
               End
               Begin VB.Label lbl��λ������ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��λ������"
                  Height          =   240
                  Left            =   240
                  TabIndex        =   113
                  Top             =   450
                  Width           =   1200
               End
               Begin VB.Label lbl��λ�ʺ� 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��λ�ʺ�"
                  Height          =   240
                  Left            =   6600
                  TabIndex        =   112
                  Top             =   450
                  Width           =   960
               End
               Begin VB.Label lbl��ϵ������ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ϵ������"
                  Height          =   240
                  Left            =   6360
                  TabIndex        =   111
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label lbl��ϵ�˹�ϵ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ϵ�˹�ϵ"
                  Height          =   240
                  Left            =   210
                  TabIndex        =   110
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label lbl��ϵ�˵�ַ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ϵ�˵�ַ"
                  Height          =   240
                  Left            =   6360
                  TabIndex        =   109
                  Top             =   1230
                  Width           =   1200
               End
               Begin VB.Label lbl��ϵ�˵绰 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ϵ�˵绰"
                  Height          =   240
                  Left            =   11490
                  TabIndex        =   108
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label lbl������λ 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "������λ"
                  Height          =   240
                  Left            =   480
                  TabIndex        =   107
                  Top             =   60
                  Width           =   960
               End
               Begin VB.Label lbl��λ�ʱ� 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��λ�ʱ�"
                  Height          =   240
                  Left            =   11730
                  TabIndex        =   106
                  Top             =   60
                  Width           =   960
               End
               Begin VB.Label lbl��λ�绰 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��λ�绰"
                  Height          =   240
                  Left            =   6600
                  TabIndex        =   105
                  Top             =   60
                  Width           =   960
               End
            End
            Begin VB.TextBox txt���֤�� 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   7755
               TabIndex        =   6
               Top             =   630
               Width           =   2340
            End
            Begin VB.ComboBox cbo��� 
               Height          =   360
               Left            =   12765
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   1410
               Width           =   1725
            End
            Begin VB.ComboBox cbo���� 
               Height          =   360
               Left            =   10080
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   1410
               Width           =   1725
            End
            Begin VB.ComboBox cbo���� 
               Height          =   360
               Left            =   7755
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   1410
               Width           =   1725
            End
            Begin VB.ComboBox cbo�������� 
               Height          =   360
               Left            =   12765
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   2970
               Width           =   1740
            End
            Begin VB.CommandButton cmdYB 
               Caption         =   "��֤"
               Height          =   350
               Left            =   9555
               TabIndex        =   3
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ�:F12(ҽ��������֤)"
               Top             =   245
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.TextBox txtסԺ�� 
               Height          =   360
               Left            =   4050
               MaxLength       =   18
               TabIndex        =   1
               Top             =   240
               Width           =   1875
            End
            Begin VB.TextBox txtPatient 
               BackColor       =   &H00EBFFFF&
               ForeColor       =   &H00C00000&
               Height          =   360
               Left            =   1590
               TabIndex        =   0
               ToolTipText     =   "�����벡�˱�ʶ����������,ֱ�ӻس��Ǽ��²���,��λ�ȼ�:F11"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txt���� 
               Height          =   360
               Left            =   7740
               MaxLength       =   64
               TabIndex        =   2
               ToolTipText     =   "���벡������,��ֱ�ӻس���֤ҽ������,����ǲ�����ǰ�Ĳ���,���ڲ������������"
               Top             =   240
               Width           =   1725
            End
            Begin VB.ComboBox cbo�ѱ� 
               Height          =   360
               Left            =   10065
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   1020
               Width           =   1725
            End
            Begin VB.ComboBox cboְҵ 
               Height          =   360
               Left            =   12765
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1800
               Width           =   1725
            End
            Begin VB.ComboBox cboѧ�� 
               Height          =   360
               Left            =   7755
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   1800
               Width           =   1725
            End
            Begin VB.ComboBox cbo����״�� 
               Height          =   360
               Left            =   10080
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   1800
               Width           =   1725
            End
            Begin VB.TextBox txt��ͥ��ַ 
               Height          =   360
               Left            =   1590
               MaxLength       =   100
               TabIndex        =   24
               Top             =   2190
               Width           =   4335
            End
            Begin VB.TextBox txt��ͥ��ַ�ʱ� 
               Height          =   360
               Left            =   12765
               MaxLength       =   6
               TabIndex        =   28
               Top             =   2190
               Width           =   1725
            End
            Begin VB.TextBox txt��ͥ�绰 
               Height          =   360
               Left            =   9555
               MaxLength       =   20
               TabIndex        =   27
               Top             =   2190
               Width           =   1725
            End
            Begin VB.TextBox txt���� 
               Height          =   360
               IMEMode         =   2  'OFF
               Left            =   4290
               TabIndex        =   12
               Top             =   1020
               Width           =   915
            End
            Begin VB.ComboBox cbo�Ա� 
               Height          =   360
               ItemData        =   "frmHosReg.frx":0446
               Left            =   7755
               List            =   "frmHosReg.frx":0448
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   1020
               Width           =   1725
            End
            Begin VB.TextBox txtҽ���� 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   1590
               MaxLength       =   30
               TabIndex        =   5
               Top             =   630
               Width           =   4335
            End
            Begin VB.ComboBox cbo���䵥λ 
               Height          =   360
               Left            =   5235
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   1020
               Width           =   705
            End
            Begin VB.CommandButton cmdTurn 
               Caption         =   "�������תסԺ(&T)"
               Height          =   350
               Left            =   10200
               TabIndex        =   4
               TabStop         =   0   'False
               ToolTipText     =   "�ȼ�:F12(ҽ��������֤)"
               Top             =   245
               Visible         =   0   'False
               Width           =   2160
            End
            Begin VB.TextBox txt���� 
               BackColor       =   &H80000004&
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   1590
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   161
               TabStop         =   0   'False
               Top             =   1410
               Width           =   4335
            End
            Begin VB.TextBox txt����֤�� 
               Height          =   360
               Left            =   1590
               MaxLength       =   20
               TabIndex        =   20
               Top             =   1800
               Width           =   4335
            End
            Begin VB.ComboBox cboҽ�Ƹ��� 
               Height          =   360
               Left            =   12765
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   1020
               Width           =   1725
            End
            Begin VB.TextBox txt֧������ 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   11925
               MaxLength       =   20
               PasswordChar    =   "*"
               TabIndex        =   8
               Top             =   630
               Width           =   975
            End
            Begin VB.TextBox txt��֤���� 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   13485
               MaxLength       =   20
               PasswordChar    =   "*"
               TabIndex        =   9
               Top             =   630
               Width           =   975
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   30
               Left            =   7365
               TabIndex        =   103
               Top             =   555
               Width           =   30
               _ExtentX        =   53
               _ExtentY        =   53
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin zlIDKind.IDKindNew IDKind 
               Height          =   360
               Left            =   825
               TabIndex        =   187
               ToolTipText     =   "��ݼ�F4"
               Top             =   240
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   635
               Appearance      =   2
               IDKindStr       =   $"frmHosReg.frx":044A
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
               DefaultCardType =   "0"
               BackColor       =   -2147483633
            End
            Begin MSCommLib.MSComm com 
               Left            =   12960
               Top             =   0
               _ExtentX        =   1005
               _ExtentY        =   1005
               _Version        =   393216
               DTREnable       =   -1  'True
            End
            Begin MSMask.MaskEdBox txt����ʱ�� 
               Height          =   360
               Left            =   2955
               TabIndex        =   11
               Top             =   1020
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   635
               _Version        =   393216
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txt�������� 
               Bindings        =   "frmHosReg.frx":052D
               Height          =   360
               Left            =   1590
               TabIndex        =   10
               Top             =   1020
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   635
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "YYYY-MM-DD"
               Mask            =   "####-##-##"
               PromptChar      =   "_"
            End
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   360
               Index           =   1
               Left            =   1590
               TabIndex        =   38
               Tag             =   "�����ص�"
               Top             =   2970
               Visible         =   0   'False
               Width           =   4350
               _ExtentX        =   7673
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Items           =   3
               MaxLength       =   100
            End
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   360
               Index           =   2
               Left            =   11805
               TabIndex        =   35
               Tag             =   "����"
               Top             =   2580
               Visible         =   0   'False
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Items           =   2
               MaxLength       =   100
            End
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   360
               Index           =   3
               Left            =   1590
               TabIndex        =   26
               Tag             =   "��סַ"
               Top             =   2190
               Visible         =   0   'False
               Width           =   6270
               _ExtentX        =   11060
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   100
            End
            Begin ZlPatiAddress.PatiAddress PatiAddress 
               Height          =   360
               Index           =   4
               Left            =   1590
               TabIndex        =   31
               Tag             =   "���ڵ�ַ"
               Top             =   2580
               Visible         =   0   'False
               Width           =   6270
               _ExtentX        =   11060
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   100
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   240
               Left            =   11295
               TabIndex        =   149
               Top             =   2640
               Width           =   480
            End
            Begin VB.Label lbl���ڵ�ַ�ʱ� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ڵ�ַ�ʱ�"
               Height          =   240
               Left            =   7995
               TabIndex        =   148
               Top             =   2640
               Width           =   1440
            End
            Begin VB.Label lbl���ڵ�ַ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ڵ�ַ"
               Height          =   240
               Left            =   510
               TabIndex        =   147
               Top             =   2640
               Width           =   960
            End
            Begin VB.Label lblPatiColor 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   14520
               TabIndex        =   146
               Top             =   3000
               Width           =   300
            End
            Begin VB.Label lbl���֤�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���֤��"
               Height          =   240
               Left            =   6675
               TabIndex        =   145
               Top             =   690
               Width           =   960
            End
            Begin VB.Label lbl��� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���"
               Height          =   240
               Left            =   12255
               TabIndex        =   144
               Top             =   1470
               Width           =   480
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   240
               Left            =   9525
               TabIndex        =   143
               Top             =   1470
               Width           =   480
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   240
               Left            =   7125
               TabIndex        =   142
               Top             =   1470
               Width           =   480
            End
            Begin VB.Label lbl�����ص� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����ص�"
               Height          =   240
               Left            =   510
               TabIndex        =   141
               Top             =   3030
               Width           =   960
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   240
               Left            =   7125
               TabIndex        =   140
               Top             =   3030
               Width           =   480
            End
            Begin VB.Label lblPatiType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               Height          =   240
               Left            =   11775
               TabIndex        =   139
               Top             =   3030
               Width           =   960
            End
            Begin VB.Label lblUnUseful 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "������"
               ForeColor       =   &H00FF0000&
               Height          =   855
               Left            =   45
               TabIndex        =   138
               Top             =   2490
               Width           =   300
            End
            Begin VB.Label lblסԺ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺ��"
               Height          =   240
               Left            =   3255
               TabIndex        =   137
               Top             =   300
               Width           =   720
            End
            Begin VB.Label lbl����ID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   465
               TabIndex        =   136
               Top             =   300
               Width           =   240
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   240
               Left            =   7155
               TabIndex        =   135
               Top             =   300
               Width           =   480
            End
            Begin VB.Label lbl�Ա� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�Ա�"
               Height          =   240
               Left            =   7125
               TabIndex        =   134
               Top             =   1080
               Width           =   480
            End
            Begin VB.Label lbl�ѱ� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ѱ�"
               Height          =   240
               Left            =   9540
               TabIndex        =   133
               Top             =   1080
               Width           =   480
            End
            Begin VB.Label lblְҵ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ְҵ"
               Height          =   240
               Left            =   12255
               TabIndex        =   132
               Top             =   1860
               Width           =   480
            End
            Begin VB.Label lblѧ�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ѧ��"
               Height          =   240
               Left            =   7125
               TabIndex        =   131
               Top             =   1860
               Width           =   480
            End
            Begin VB.Label lbl����״�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   240
               Left            =   9555
               TabIndex        =   130
               Top             =   1860
               Width           =   480
            End
            Begin VB.Label lbl��ͥ��ַ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��סַ"
               Height          =   240
               Left            =   750
               TabIndex        =   129
               Top             =   2250
               Width           =   720
            End
            Begin VB.Label lbl��ͥ�绰 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ͥ�绰"
               Height          =   240
               Left            =   8475
               TabIndex        =   128
               Top             =   2250
               Width           =   960
            End
            Begin VB.Label lbl��ͥ��ַ�ʱ� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ͥ��ַ�ʱ�"
               Height          =   240
               Left            =   11295
               TabIndex        =   127
               Top             =   2250
               Width           =   1440
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               Height          =   240
               Left            =   510
               TabIndex        =   126
               Top             =   1080
               Width           =   960
            End
            Begin VB.Label lbl���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   240
               Left            =   3765
               TabIndex        =   125
               Top             =   1080
               Width           =   480
            End
            Begin VB.Label lblҽ���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽ����"
               Height          =   240
               Left            =   750
               TabIndex        =   124
               Top             =   690
               Width           =   720
            End
            Begin VB.Label lbl���� 
               Alignment       =   1  'Right Justify
               Caption         =   "��������"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   675
               TabIndex        =   123
               Top             =   1470
               Width           =   825
            End
            Begin VB.Label lbl����֤�� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����֤��"
               Height          =   240
               Left            =   510
               TabIndex        =   122
               Top             =   1860
               Width           =   960
            End
            Begin VB.Label lblҽ�Ƹ��� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ѷ�ʽ"
               Height          =   240
               Left            =   11775
               TabIndex        =   121
               Top             =   1080
               Width           =   960
            End
            Begin VB.Label lbl֧������ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   240
               Left            =   11445
               TabIndex        =   120
               Top             =   690
               Width           =   480
            End
            Begin VB.Label lbl��֤���� 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��֤"
               Height          =   240
               Left            =   12975
               TabIndex        =   119
               Top             =   690
               Width           =   480
            End
         End
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   15120
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   10170
      Width           =   15120
      Begin VB.CommandButton cmdSurety 
         Caption         =   "������Ϣ(&S)"
         Height          =   400
         Left            =   1515
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   90
         Width           =   1845
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   400
         Left            =   255
         TabIndex        =   96
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   400
         Left            =   13440
         TabIndex        =   95
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   400
         Left            =   12240
         TabIndex        =   94
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdDeposit 
         Caption         =   "Ԥ������ȡ(&D)"
         Height          =   400
         Left            =   3510
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   90
         Width           =   1845
      End
   End
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   5155
      Left            =   0
      TabIndex        =   186
      Top             =   0
      Width           =   15075
      _Version        =   589884
      _ExtentX        =   26591
      _ExtentY        =   9093
      _StockProps     =   64
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "��ӡ������(&S)"
      Visible         =   0   'False
      Begin VB.Menu mnu������ҳ 
         Caption         =   "������ҳ(&1)"
      End
      Begin VB.Menu mnuԤ�����վ� 
         Caption         =   "Ԥ�����վ�(&2)"
      End
   End
End
Attribute VB_Name = "frmHosReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit 'Ҫ���������
Public mstrPrivs As String
Public mlngModul As Long
Public mbytMode As Byte '�룺0-�����Ǽ�,1-ԤԼ�Ǽ�,2-����ԤԼ   '���ÿ����Ժʹ����סԺ��,��ԤԼʱ����,����ʱ�ٲ���(��Ϊҽ�������ԤԼ)
Public mbytKind As Byte '�룺0=סԺ��Ժ�Ǽ�,1-�������۵Ǽ�,2-סԺ���۵Ǽ�
Public mbytInState As Byte '�룺0=����,1=�޸�,2=����
'�룺Ҫ���ģ��޸ģ����յĲ���ID����ҳID(ԤԼ��Ϊ0)
Public mlng����ID As Long
Private mlng�Һ�ID As Long              'ԤԼ���Ĳ��˽��պ�ش�����״̬ʱ��
Public mlng��ҳID As Long
'Private mstrԤ��NO As String
Private mrsInfo As ADODB.Recordset '������Ϣ
Private mrsPatiReg As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mrsUnit As ADODB.Recordset
Private mrsDept As ADODB.Recordset
Private mrsUnitDept As ADODB.Recordset  '�������Ҷ�Ӧ
Private mrsInputSet  As ADODB.Recordset '���������  �ֶ�����:������Ŀ,��ֹ¼��,��������,������,�ؼ���,�ؼ��±�

Private mblnICCard As Boolean 'IC������,Ҫͬʱ��д������Ϣ��IC���ֶ�
Private mblnOneCard As Boolean      '�Ƿ�������һ��ͨ�ӿ�,��ģʽ�£�Ʊ���ϸ����Ʊ�ŷ�Χ��ķ�����󶨿����շ�

Private mblnAuto As Boolean '
Private mblnUnload As Boolean
Private mlngԤ������ID As Long
Private mblnChange As Boolean
Private mbln�Ƿ�ɨ�����֤ As Boolean

Private mblnPrepayPrint As Boolean    '�Ƿ��ӡԤ����
Private mblnFPagePrint As Boolean   '�Ƿ��ӡ������ҳ
Private mblnWristletPrint As Boolean    '�Ƿ��ӡ�������
Private mdat�ϴε�������ʱ�� As Date '�޸ĵǼ���Ϣʱ,�ϴ�ʱ�޵����ĵ���ʱ��
Private mstrNOS As String   'ѡ��ת��ĵ���,Ʊ��,����ID,����(��ҽ��Ϊ��):H0000001,F000023,81235,901;H0000002,F000045,81263,901;...
Private mobjKeyboard As Object
 
Private mblnHaveAdvice As Boolean   '��ǰ�����Ƿ����ҽ��

Private mstrPatiPlus    As String     '�ӱ���Ϣ:��Ϣ��1:��Ϣֵ1,��Ϣ��2:��Ϣֵ2
Private mblnEMPI As Boolean               'T-�ҵ�EMPI����,F-δ�ҵ�EMPI����
Private mblnAppoint As Boolean              'T-ԤԼ���Ĳ���ֱ�� ��Ժ���
Private mstrAppointBed As String            'ԤԼ��λ

'ҽ������---------------
Private mintInsure As Integer
Private mstrYBPati As String
Private mcurYBMoney As Currency '�����ʻ����
'����Ϊ�ϲ������Ƕ�Ӧ�ļ�¼����
Private mintInsureBak As Integer
Private mstrYBPatiBak As String
Private mcurYBMoneyBak As Currency '�����ʻ����
Private mbytKindBak As Byte
Private mbln�մ� As Boolean

Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML


Private Enum EState
    E���� = 0
    E�޸� = 1
    E���� = 2
End Enum
Private Enum EMode
    E�����Ǽ� = 0
    EԤԼ�Ǽ� = 1
    E����ԤԼ = 2
End Enum
Private Enum EKind
    EסԺ��Ժ�Ǽ� = 0
    E�������۵Ǽ� = 1
    EסԺ���۵Ǽ� = 2
End Enum

'-----------------------------------------------------------------
'��Ʊ���
Private mFactProperty As Ty_FactProperty
'-----------------------------------------------------------------
'ҽ�ƿ����
'Private mobjSquareCard As Object
Private mblnClickSquareCtrl As Boolean
Private mblnStartFactUseType As Boolean '�Ƿ����õ���ص���������
Private mbytPrepayType As Byte '0-����סԺ;1-����;2-סԺ
Private mblnNotClick As Boolean
Private mblnIdNotClick  As Boolean
Private mblnICNotClick As Boolean
Private mblnCheckPatiCard As Boolean

Private mobjSquare As Object 'ҽ�ƿ�����
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private Type Ty_CardProperty
       lng�����ID As Long
       str������  As String
       lng���ų��� As Long
       lng���㷽ʽ As String
       bln���ƿ� As Boolean
       bln�ϸ���� As Boolean
       lng����ID As Long
       lng�������� As Long
       bln��� As Boolean
       bln�ظ����� As Boolean
       bln���￨ As Boolean
       str�������� As String
       int���볤�� As Integer
       int���볤������ As Integer
       int������� As Integer
       blnOneCard As Boolean '  '�Ƿ�������һ��ͨ�ӿ�,��ģʽ�£�Ʊ���ϸ����Ʊ�ŷ�Χ��ķ�����󶨿����շ�
       rs���� As ADODB.Recordset
       dblӦ�ս�� As Double
       dblʵ�ս�� As Double
       bln�Ƿ��ƿ� As Boolean '�����:56599
       bln�Ƿ񷢿� As Boolean
       bln�Ƿ�д�� As Boolean
       bln�Ƿ�Ժ�ⷢ��  As Boolean
       lng�������� As Long '0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0 �����:57326
       str�������� As String
       byt�������� As Byte
       str�ض���Ŀ As String
End Type
Private mCurSendCard As Ty_CardProperty
Private mcolPrepayPayMode As Collection   'Ԥ����֧����ʽ
Private mcolCardPayMode As Collection   '���￨֧����ʽ

Private Type Ty_PayMoney
    lngҽ�ƿ����ID As Long
    bln���ѿ� As Boolean
    str���㷽ʽ As String
    str���� As String
    strˢ������ As String
    strˢ������ As String
    strno As String
    lngID As Long 'Ԥ��ID
    lng����ID As Long
End Type

Private mCurPrepay As Ty_PayMoney
Private mCurCardPay As Ty_PayMoney
Private mstrPassWord As String
Private mblnɨ�����֤ǩԼ As Boolean '���ݲ��������еġ�ɨ�����֤ǩԼ��ȡֵ
Private mstrȱʡ�ѱ� As String
'����� :56599
Private Type Ty_PageHeight
    ���� As Long
    �������� As Long
End Type
Private mPageHeight As Ty_PageHeight
Private mstrPriceGrade As String, mstrPrePriceGrade As String
Private mobjPublicExpense As Object  '���ù�������
Private mintPriceGradeStartType As Integer

Private mstrCboSplit As String
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const C_ColumHeader = "����ҩ��,1,5000,1;������ӳ,4,3000,1;����ҩ��ID,1,100,0" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_InoculateHeader = "��������,4,3500,1;��������,4,3500,1;��������,4,3500,1;��������,4,3500,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_LinkManColumHeader = "��ϵ������,4,3000,1;��ϵ�˹�ϵ,4,3000,1;��ϵ�˹�ϵ��ע,4,2000,1;��ϵ�����֤��,4,3000,1;��ϵ�˵绰,4,3000,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_OtherInfoColumHeader = "��Ϣ��,4,3600,1;��Ϣֵ,4,3600,1;��Ϣ��,4,3600,1;��Ϣֵ,4,3600,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_CertificateHeader = "֤������,4,3500,1;֤������,4,3500,1;֤������,4,3500,1;֤������,4,3500,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_BH = "��,��,����,δ��"
'C_��������Ƹ�ʽ:������,�ؼ�(�ؼ�1,�ؼ�2,...)|����2,�ؼ�|...
Private Const C_��������� = "����,cbo����|����,cbo����|ѧ��,cboѧ��|����״��,cbo����״��|ְҵ,cboְҵ|���,cbo���|��������,txt��������,txt����ʱ��|����֤��,txt����֤��" & _
                        "|���֤��,txt���֤��,cboIDNumber|�����ص�,txt�����ص�,PatiAddress(1)|��סַ,txt��ͥ��ַ,PatiAddress(3)|��ͥ��ַ�ʱ�,txt��ͥ��ַ�ʱ�|��ͥ�绰,txt��ͥ�绰|��ϵ������,txt��ϵ������|��ϵ�˹�ϵ,cbo��ϵ�˹�ϵ,txtLinkManInfo" & _
                        "|���ڵ�ַ,txt���ڵ�ַ,PatiAddress(4)|���ڵ�ַ�ʱ�,txt���ڵ�ַ�ʱ�|����,txt����|��ϵ�˵�ַ,txt��ϵ�˵�ַ,PatiAddress(5)|��ϵ�˵绰,txt��ϵ�˵绰|��ϵ�����֤��,txt��ϵ�����֤��" & _
                        "|������λ,txt������λ|��λ�绰,txt��λ�绰|��λ�ʱ�,txt��λ�ʱ�|��λ������,txt��λ������|��λ�ʺ�,txt��λ�ʺ�|����,txt����,PatiAddress(2)"
Private Const C_COLOR_UNEnabled = &H80000004 '��ֹ¼����ɫ
Private Const C_COLOR_Enabled = &H80000005 '����ֹ¼����ʾ��ɫ

Private mdicҽ�ƿ����� As New Dictionary
Private mbln������󶨿� As Boolean
Private mbln�Ƿ���ʾԤ�� As Boolean
Private mbln�Ƿ���ʾ�ſ� As Boolean
Private marrAddress(0 To 4) As String     '�弶�ṹ����ַȱʡֵ
Private mstrFirstCode As String '��һ��֤�����͵ı���
'-----------------------------------------------------------------
Private mintIDKind As String

Private Sub cbo��������_Click()
    If cbo��������.ListCount > 0 And cbo��������.ListIndex <> -1 Then
        lblPatiColor.BackColor = zlDatabase.GetPatiColor(zlCommFun.GetNeedName(cbo��������.Text))
        txt����.ForeColor = lblPatiColor.BackColor
    End If
End Sub
 

Private Sub cbo��������_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    '�����:48352
    With mCurCardPay
            .lngҽ�ƿ����ID = 0
            .bln���ѿ� = False
            .str���㷽ʽ = ""
            .str���� = ""
     End With
    '0=����,1=�޸�,2=�鿴
    If mbytInState = 2 Then Exit Sub
    Call SetCardVaribles(False)
    '130245,�л����㷽ʽ��ͬ�����¿����ID
    If mblnNotClick = True Then Exit Sub
    Call Local���㷽ʽ(mCurCardPay.lngҽ�ƿ����ID, True)
End Sub

Private Sub cbo��ϵ�˹�ϵ_Click()
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ�˹�ϵ") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ�˹�ϵ")) = zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text)
    End If
    If zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text) = "����" Then
        txtLinkManInfo.Enabled = True: txtLinkManInfo.BackColor = &H80000005
    Else
        txtLinkManInfo.Enabled = False: txtLinkManInfo.Text = "": txtLinkManInfo.BackColor = &H80000004
    End If
End Sub

Private Sub cbo��ϵ�˹�ϵ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo��ϵ�˹�ϵ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo��ϵ�˹�ϵ.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo��ϵ�˹�ϵ.ListIndex = lngIdx
End Sub

Private Sub cbo��Ժ����_Validate(Cancel As Boolean)
    '����27370 by lesfeng 2010-01-26
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSQL As String, intIdx As Long, i As Long
    Dim lngDept As Long
    Dim intFlag As Integer
    
    strInput = UCase(cbo��Ժ����.Text)
    intFlag = -1
    If Trim(strInput) = "��ȷ������" Then Cancel = False: Exit Sub
    If gbln��ѡ���� Then
        Set rsTmp = InputDept(Me, fra��Ժ, cbo��Ժ����, "����", IIf(mbytKind = EKind.E�������۵Ǽ�, "1", "2") & ",3", strInput, blnCancel, intFlag, 0)
    Else
        If cbo��Ժ����.ListIndex >= 0 Then lngDept = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
        mrsUnitDept.Filter = "����ID=" & lngDept
        If mrsUnitDept.RecordCount > 0 Then
            intFlag = 2
        Else
            lngDept = 0
        End If
        Set rsTmp = InputDept(Me, fra��Ժ, cbo��Ժ����, "����", IIf(mbytKind = EKind.E�������۵Ǽ�, "1", "2") & ",3", strInput, blnCancel, intFlag, lngDept)
    End If
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cbo��Ժ����, rsTmp!ID)
        If intIdx <> -1 Then
            cbo��Ժ����.ListIndex = intIdx
'        Else
'            cbo��Ժ����.AddItem Nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����, cbo��Ժ����.ListCount - 1
'            cbo��Ժ����.ItemData(cbo��Ժ����.NewIndex) = rsTmp!ID
'            cbo��Ժ����.ListIndex = cbo��Ժ����.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ����Ժ������", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub cbo��Ժ��ʽ_Click()
    If zlCommFun.GetNeedName(cbo��Ժ��ʽ.Text) = "ת��" Then
        cmdת��.Enabled = True: cmdת��.BackColor = &H80000005
        txtת��.Enabled = True: txtת��.BackColor = &H80000005
        lblInFrom.Enabled = True
    Else
        cmdת��.Enabled = False:  cmdת��.BackColor = &H80000004
        txtת��.Enabled = False: txtת��.Text = "": txtת��.BackColor = &H80000004
        lblInFrom.Enabled = False
    End If
End Sub

'Private Sub cbo��Ժ����_GotFocus()
'    '����27370 by lesfeng 2010-01-26
''    If cbo��Ժ����.Style = 0 Then
''        Call zlcontrol.TxtSelAll(cbo��Ժ����)
''    End If
'    With cbo��Ժ����
'        .SelStart = 0
'        .SelLength = Len(.Text)
'    End With
'End Sub

Private Sub cbo��Ժ����_Validate(Cancel As Boolean)
    '����27370 by lesfeng 2010-01-26
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSQL As String, intIdx As Long, i As Long
    Dim lngUnit As Long
    
    strInput = UCase(cbo��Ժ����.Text)
    If gbln��ѡ���� Then
         If cbo��Ժ����.ListIndex >= 0 Then lngUnit = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
        Set rsTmp = InputDept(Me, fra��Ժ, cbo��Ժ����, "�ٴ�", IIf(mbytKind = EKind.E�������۵Ǽ�, "1", "2") & ",3", strInput, blnCancel, 1, lngUnit)
    Else
        Set rsTmp = InputDept(Me, fra��Ժ, cbo��Ժ����, "�ٴ�", IIf(mbytKind = EKind.E�������۵Ǽ�, "1", "2") & ",3", strInput, blnCancel, -1, 0)
    End If
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cbo��Ժ����, rsTmp!ID)
        If intIdx <> -1 Then
            cbo��Ժ����.ListIndex = intIdx
'        Else
'            cbo��Ժ����.AddItem Nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����, cbo��Ժ����.ListCount - 1
'            cbo��Ժ����.ItemData(cbo��Ժ����.NewIndex) = rsTmp!ID
'            cbo��Ժ����.ListIndex = cbo��Ժ����.NewIndex
        End If
    Else
        If cbo��Ժ����.ListIndex = -1 And cbo��Ժ����.ListCount = 0 Then
        Else
            If Not blnCancel Then
                MsgBox "δ�ҵ���Ӧ����Ժ���ҡ�", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Sub cboҽ�Ƹ���_Click()
    On Error GoTo ErrHandler
    If mintPriceGradeStartType < 2 Then Exit Sub
    Call mobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cboҽ�Ƹ���.Text), , , mstrPriceGrade)
    If mstrPrePriceGrade = mstrPriceGrade Then Exit Sub
    mstrPrePriceGrade = mstrPriceGrade

    If mCurSendCard.str�ض���Ŀ <> "" Then
        Set mCurSendCard.rs���� = zlGetSpecialItemFee(mCurSendCard.str�ض���Ŀ, mstrPriceGrade)
    Else
        Set mCurSendCard.rs���� = Nothing
    End If
    
    Call LoadCardFee
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCardFee()
    '����:���ؿ���
    On Error GoTo errHandle
    If mCurSendCard.rs���� Is Nothing Then
        txt����.Text = "": txt����.Tag = ""
        Exit Sub
    End If
    If mCurSendCard.rs����.RecordCount = 0 Then
        txt����.Text = "": txt����.Tag = ""
        Exit Sub
    End If
    
    With mCurSendCard.rs����
        txt����.Text = Format(IIf(Nvl(!�Ƿ���, 0) = 1, Val(Nvl(!ȱʡ�۸�)), Val(Nvl(!�ּ�))), "0.00")
        If Nvl(!�Ƿ���, 0) <> 1 And Nvl(!���ηѱ�, 0) <> 1 Then
            txt����.Text = Format(GetActualMoney(zlStr.NeedName(cbo�ѱ�.Text), !������ĿID, Val(txt����.Text), !�շ�ϸĿID), "0.00")
        End If
        txt����.Tag = txt����.Text  '���ֲ���
        txt����.Locked = Nvl(!�Ƿ���, 0) <> 1
        txt����.TabStop = Nvl(!�Ƿ���, 0) = 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub chk��λ�ɿ�_Click()
    If chk��λ�ɿ�.Value = 1 Then
        txt�ɿλ.Enabled = True
        txt�ɿλ.BackColor = &H80000005
    Else
        txt�ɿλ.Text = ""
        txt�ɿλ.Enabled = False
        txt�ɿλ.BackColor = Me.BackColor
    End If
End Sub

Private Sub cmdDeposit_Click()
    Dim strCommon As String, intAtom As Integer
            
    On Error Resume Next
    If gobjPatient Is Nothing Then
        Set gobjPatient = CreateObject("zl9Patient.clsPatient")
        If gobjPatient Is Nothing Then Exit Sub
    End If
    
    Err.Clear: On Error GoTo 0
    
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    Call gobjPatient.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 0, mlng����ID, mlng��ҳID, 0, 0)
    Call GlobalDeleteAtom(intAtom)
    If gbln��ԺԤ�� Then
        If gblnPrepayStrict Then
            mlngԤ������ID = CheckUsedBill(2, IIf(mlngԤ������ID > 0, mlngԤ������ID, mFactProperty.lngShareUseID), , 2)
            If mlngԤ������ID <= 0 Then
                Select Case mlngԤ������ID
                    Case 0 '����ʧ��
                    Case -1
                        MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,������Ժʱ����ͬʱ��Ԥ���" & _
                            "��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -2
                        MsgBox "���صĹ���Ʊ���Ѿ�����,������Ժʱ����ͬʱ��Ԥ���" & _
                            "��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End Select
                txtFact.Text = ""
            Else
                txtFact.Text = GetNextBill(mlngԤ������ID)
            End If
        Else
            '��ɢ��ȡ��һ������
            txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngModul, "")))
        End If
    End If
End Sub

Private Sub cmdName_Click()
    Dim rsTmp As ADODB.Recordset
    '��ȡ������Ϣ
    Set rsTmp = GetPatientByName(txt����.Text)
    Call MergePatient(rsTmp, 0)
End Sub

Private Sub cmdSelectNO_Click()
    Dim strno As String
    
    Call frmNOSelect.ShowMe(Me, strno)
    If strno <> "" Then txtסԺ��.Text = strno
    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
End Sub

Private Sub cmdSurety_Click()
    frmSurety.mlng����ID = 0
    frmSurety.mbln��Ժ���� = True
    frmSurety.mstrPrivs = mstrPrivs
    frmSurety.Show 1, Me
End Sub

Private Sub cmdTurn_Click()
    Call frmChargeTurn.ShowMe(Me, Val(txtPatient.Text), mstrNOS, , , mstrPrivs, mlngModul)
End Sub

Private Sub cmd���ڵ�ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt���ڵ�ַ, True)
    If Not rsTmp Is Nothing Then
        txt���ڵ�ַ.Text = rsTmp!����
        txt���ڵ�ַ.SelStart = Len(txt���ڵ�ַ.Text)
        txt���ڵ�ַ.SetFocus
    End If
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt����, True)
    If Not rsTmp Is Nothing Then
        txt����.Text = rsTmp!����
        txt����.SelStart = Len(txt����.Text)
        txt����.SetFocus
    Else
        zlControl.TxtSelAll txt����
        txt����.SetFocus
    End If
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt����, True)
    If Not rsTmp Is Nothing Then
        txt����.Text = rsTmp!����
        txt����.SelStart = Len(txt����.Text)
        txt����.SetFocus
    Else
        zlControl.TxtSelAll txt����
        txt����.SetFocus
    End If
End Sub

Private Sub cmdת��_Click()
    Dim vPoint As POINTAPI
    On Error GoTo errH
    vPoint = GetCoordPos(txtת��.Container.hWnd, txtת��.Left, txtת��.Top)
    Call Getҽ�ƻ���(txtת��, Me, 2, "ҽ�ƻ���", "�ֵ������", vPoint, False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    '�����:53408
    '�����:53408
    mblnɨ�����֤ǩԼ = IIf(zlDatabase.GetPara("ɨ�����֤ǩԼ", glngSys, glngModul) = "1", 1, 0) = "1"
    If mCurSendCard.str������ Like "*�������֤*" Then
        lbl����.Enabled = False: txt����.Enabled = False
        lbl����.Enabled = False: txtPass.Enabled = False
        lbl��֤.Enabled = False: txtAudi.Enabled = False
    End If
    Call Show�󶨿ؼ�(False)
    If gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
    End If
    
    SetCardEditEnabled
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
          If mobjICCard Is Nothing Then
              Set mobjICCard = New clsICCard
              Call mobjICCard.SetParent(Me.hWnd)
              Set mobjICCard.gcnOracle = gcnOracle
          End If
          If Not mobjICCard Is Nothing Then
              txtPatient.Text = mobjICCard.Read_Card()
              If txtPatient.Text <> "" Then
                  Call txtPatient_KeyPress(vbKeyReturn)
                  mblnICCard = True
              End If
          End If
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, False, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    '�����:56599
    If strOutPatiInforXML <> "" Then Call LoadPati(strOutPatiInforXML)
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '�Ƿ�������ʾ
    'txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55571:������,2012-11-12
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And (Not mblnIdNotClick And Not mblnICNotClick) Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub
Private Sub lblUnUseful_Click()
    Dim lngTop As Long
     '51167,������,2012-07-09,����"��ϵ�����֤��"
    lblUnUseful.Appearance = IIf(lblUnUseful.Appearance = 0, 1, 0)
    picUnUseful.Tag = IIf(lblUnUseful.Appearance = 0, 1, 0)
    If lblUnUseful.Appearance = 0 Then
        fra����.Height = fra����.Height - picUnUseful.Height - IIf(pic����.Visible = False, pic����.Height, 0)
        pic����.Height = pic����.Height - picUnUseful.Height - IIf(pic����.Visible = False, pic����.Height, 0)
        pic����.Top = picUnUseful.Top
        Me.Height = Me.Height - picUnUseful.Height - IIf(pic����.Visible = False, pic����.Height, 0)
        tbcPage.Height = picCmd.Top
        picUnUseful.Visible = False
        lblUnUseful.Caption = "����ʾ"
    ElseIf lblUnUseful.Appearance = 1 Then
        fra����.Height = fra����.Height + picUnUseful.Height + IIf(pic����.Visible = False, pic����.Height, 0)
        pic����.Height = pic����.Height + picUnUseful.Height + IIf(pic����.Visible = False, pic����.Height, 0)
        pic����.Top = pic����.Top + picUnUseful.Height + 35
        Me.Height = Me.Height + picUnUseful.Height + IIf(pic����.Visible = False, pic����.Height, 0)
        tbcPage.Height = picCmd.Top
        picUnUseful.Visible = True
        lblUnUseful.Caption = "������"
    End If
    pic��Ժ.Top = pic����.Top + pic����.Height
    lngTop = pic��Ժ.Top + pic��Ժ.Height
    If mbln�Ƿ���ʾԤ�� Then
        picԤ��.Top = lngTop
        lngTop = picԤ��.Top + picԤ��.Height
    End If
    pic�ſ�.Top = lngTop
            
'    pic��Ժ.Top = pic����.Top + pic����.Height
'    picԤ��.Top = pic��Ժ.Top + pic��Ժ.Height
'    pic�ſ�.Top = picԤ��.Top + picԤ��.Height
    lblUnUseful.ForeColor = &HFF0000
    mPageHeight.���� = Me.Height
End Sub

Private Sub lbl����_Click()
    Dim strExpand As String, strOutCardNO As String, strOutPatiInforXML As String

    If mCurSendCard.bln���￨ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt����.Text = mobjICCard.Read_Card()
            If txt����.Text <> "" Then mblnICCard = True
        End If
        Exit Sub
    End If
    If (Mid(mCurSendCard.str��������, 3, 1) = 0 And Mid(mCurSendCard.str��������, 4, 1) = 0) Or mCurSendCard.lng�����ID = 0 Then Exit Sub
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

    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, mCurSendCard.lng�����ID, False, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txt����.Text = strOutCardNO
    If txt����.Text <> "" Then
        '�����:56599
       If strOutPatiInforXML <> "" Then Call LoadPati(strOutPatiInforXML)
       Call CheckFreeCard(txt����.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    End If
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNO As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt����.Text = strCardNO
    If txt����.Text <> "" Then
        '�����:56599
       If strXmlCardInfor <> "" Then Call LoadPati(strXmlCardInfor)
       Call CheckFreeCard(txt����.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    End If
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    Dim lngPreIDKind As Long, lngIndex As Long
    
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnICNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("IC����")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strCardNO
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text <> "" Then
            Call mobjICCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
        End If
        
        IDKind.IDKind = lngPreIDKind
        mblnICNotClick = False
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    Dim lngIndex As Long
    Dim blnǩԼ As Boolean
    Dim strErrMsg As String
'    '�����:53408
'    mbln�Ƿ�ɨ�����֤ = True
'
'    txt���֤��.Text = strID
'    If mCurSendCard.str������ = "�������֤" Then
'        txt����.Text = strID
'        Exit Sub
'    End If
    
    mbln�Ƿ�ɨ�����֤ = False
    
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        mblnIdNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("���֤��")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        
        '57945:������,2013-10-30,��ȡ���֤�еĵ�ַӦ�÷ŵ����ڵ�ַ�����Ǽ�ͥ��ַ
        If mrsInfo Is Nothing Then
            lngIndex = IDKind.GetKindIndex("����")
            If lngIndex >= 0 Then IDKind.IDKind = lngIndex
            txtPatient.Text = ""
            Call txtPatient_KeyPress(vbKeyReturn)
            txt����.Text = strName
            Call cbo.Locate(cbo�Ա�, strSex)
            Call cbo.Locate(cbo����, strNation)
            txt��������.Text = Format(datBirthDay, "yyyy-MM-dd")
            txt���֤��.Text = strID
        End If
        '101692�²���ֱ����ȡ;�Ѿ��������˻��ڵ�ַΪ��ʱ,�����֤��ȡ
        If mrsInfo Is Nothing Or (Not mrsInfo Is Nothing And Trim(txt���ڵ�ַ.Text) = "") Then
            txt���ڵ�ַ.Text = strAddress
            If gbln���ýṹ����ַ Then
                PatiAddress(E_IX_���ڵ�ַ).Value = strAddress
            End If
        End If
        IDKind.IDKind = lngPreIDKind
        mblnIdNotClick = False
        
        If (mCurSendCard.str������ = "�������֤" Or mblnɨ�����֤ǩԼ) Then
            blnǩԼ = �Ƿ��Ѿ�ǩԼ(Trim(strID))
            '���û��ǩԼ,������� �Ա�,���յ����
            If Not blnǩԼ And Not mrsInfo Is Nothing Then
                  If Nvl(mrsInfo!����) <> Trim(strName) Or Nvl(mrsInfo!�Ա�) <> strSex Or Format(txt��������.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
            
                      If Nvl(mrsInfo!����) <> Trim(strName) Then
                           strErrMsg = strErrMsg & "," & "����"
                      End If
                      If Nvl(mrsInfo!�Ա�) <> strSex Then
                           strErrMsg = strErrMsg & "," & "�Ա�"
                      End If
                      If Format(txt��������.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
                          strErrMsg = strErrMsg & "," & "��������"
                      End If
                      strErrMsg = Mid(strErrMsg, 2)
                      strErrMsg = "��ǰ������Ϣ�����֤�ϵ�[" & strErrMsg & "]����Ϣ��һ��!" & vbCrLf & "���ܽ������֤ǩԼ����!"
                      Call MsgBox(strErrMsg, vbQuestion, Me.Caption)
                       mbln�Ƿ�ɨ�����֤ = False
                  Else
                       mbln�Ƿ�ɨ�����֤ = True
                  End If
            ElseIf Not blnǩԼ Then
                mbln�Ƿ�ɨ�����֤ = True
            End If
            
        End If
    End If
    
    

    If Me.ActiveControl Is txt���֤�� Then
        
        If txt����.Text <> "" And cbo�Ա�.ListCount <> 0 And txt��������.Text <> "" Then
            If strName <> txt����.Text Or strSex <> Split(cbo�Ա�.Text, "-")(1) Or txt��������.Text <> Format(datBirthDay, "yyyy-MM-dd") Then
                    MsgBox "���֤��Ϣ��ҺŲ�����Ϣ��һ��,���ܽ���ǩԼ������", vbInformation, gstrSysName
                    Exit Sub
            End If
        Else
             MsgBox "�󶨶������֤ʱ,������Ϣ������Ϊ�գ�", vbInformation, gstrSysName
             Exit Sub
        End If
        
        If �Ƿ��Ѿ�ǩԼ(Trim(strID)) Then
            MsgBox "���֤����Ϊ:" & strID & "�Ѿ�ǩԼ�����ظ�ǩԼ��", vbOKOnly + vbInformation, gstrSysName
            txt���֤��.SetFocus
            Exit Sub
        Else
            txt���֤��.Text = strID
            mbln�Ƿ�ɨ�����֤ = True
        End If
        
    End If
    
    Call Show�󶨿ؼ�(mbln�Ƿ�ɨ�����֤ And mblnɨ�����֤ǩԼ)
End Sub

Public Sub Show�󶨿ؼ�(blnShow As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ���ʾ������
    '���:blnShow �Ƿ���ʾ������
    '����:����
    '����:2012-09-04 15:53:14
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    lbl֧������.Enabled = blnShow: txt֧������.Enabled = blnShow
    lbl��֤����.Enabled = blnShow: txt��֤����.Enabled = blnShow
    If blnShow = False Then
        txt֧������.Text = "": txt��֤����.Text = "": txt��֤����.Tag = ""
    End If
    
End Sub

Private Sub cbo����ҽʦ_Validate(Cancel As Boolean)
    Dim strDoctor As String
    Dim blnFinded As Boolean
    
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSQL As String, intIdx As Long, i As Long
    Dim lngUnit As Long
    
    If mbytMode = 1 Then Exit Sub
    
    '����27370 by lesfeng 2010-02-03
    strInput = UCase(cbo����ҽʦ.Text)
    Set rsTmp = InputDoctors(Me, fra��Ժ, cbo����ҽʦ, 0, "1,2,3", strInput, blnCancel, "")

    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cbo����ҽʦ, rsTmp!ID)
        If intIdx <> -1 Then
            cbo����ҽʦ.ListIndex = intIdx
'        Else
'            cbo����ҽʦ.AddItem Nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����, cbo��Ժ����.ListCount - 1
'            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
'            cbo����ҽʦ.ListIndex = cbo����ҽʦ.NewIndex
        End If
    Else
        Call zlControl.TxtSelAll(cbo����ҽʦ)
        If Not blnCancel Then
            cbo����ҽʦ.Text = ""
'            MsgBox "δ�ҵ���Ӧ��ҽ����", vbInformation, gstrSysName
        End If
'        Cancel = True: Exit Sub
    End If
    
   
'    If cbo����ҽʦ.Locked Then Exit Sub
'    If cbo����ҽʦ.ListCount = 0 Then cbo����ҽʦ.Text = "": Exit Sub
'
'    strDoctor = cbo����ҽʦ.Text
'
'    If mrsDoctor.State = 1 Then
'        If mrsDoctor.RecordCount = 0 Then cbo����ҽʦ.Text = "": Exit Sub
'        mrsDoctor.MoveFirst
'        For i = 1 To mrsDoctor.RecordCount
'            If UCase(strDoctor) = mrsDoctor!��� Or strDoctor = mrsDoctor!���� Or UCase(strDoctor) = mrsDoctor!���� Or strDoctor = mrsDoctor!���� & "-" & mrsDoctor!���� Then
'                strDoctor = mrsDoctor!ID
'                blnFinded = True
'                Exit For
'            End If
'            mrsDoctor.MoveNext
'        Next
'       If Not blnFinded Then Call zlCommFun.PressKey(vbKeyF4)
'    End If
'
'    If blnFinded Then
'        If Not Cbo.Locate(cbo����ҽʦ, strDoctor, True) Then
'            Call zlcontrol.TxtSelAll(cbo����ҽʦ)

'            Cancel = True
'        End If
'    Else
'        Call zlcontrol.TxtSelAll(cbo����ҽʦ)
'        Cancel = mrsDoctor.State = 1 And txtPatient.Text <> ""   'û������ʱ�����뿪����
'        If Not Cancel Then cbo����ҽʦ.Text = ""
'    End If
End Sub

Private Sub cbo���䵥λ_LostFocus()
    '68489:������,2013-12-06,����Ϊ���򲻽��г����շ���
    Dim strBirth As String
    Dim strMsg As String
    Dim lngTmp As Long
    
    If Trim(txt����.Text) = "" Then Exit Sub
    If Not CheckOldData(txt����, cbo���䵥λ) Then Exit Sub
    
    If Not IsDate(txt��������.Text) Then
        mblnChange = False
        Call ReCalcBirthDay(strMsg)
        mblnChange = True
        If InStr(1, strMsg, "|") > 0 Then
            lngTmp = Val(Split(strMsg, "|")(0)) '1��ֹ,0��ʾ
            strMsg = Split(strMsg, "|")(1)
            If lngTmp = 1 Then
                MsgBox strMsg, vbInformation, gstrSysName
                If CanFocus(txt����) = True Then txt����.SetFocus: Exit Sub
            End If
        End If
    End If
    Call ReLoadCardFee
End Sub

Private Sub cbo��Ժ����_Click()
    Dim lngDepID As Long
    Dim rsDiagnosis As ADODB.Recordset
    
    If cbo��Ժ����.ListIndex <> -1 Then
        If mbytInState <> EState.E���� Then Call LoadDept(1)

        cbo��λ.TabStop = (cbo��Ժ����.Text = "��ȷ������")
        '107823��ʾ����������
        If cbo��Ժ����.ListIndex <> -1 Then
            lngDepID = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
            If (mbytMode = EMode.E����ԤԼ Or mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0 And mbytInState = EState.E����) And Me.Visible = True Then
                Set rsDiagnosis = GetDiagnosticInfo(mlng����ID, mlng��ҳID, "1,11", "3", lngDepID)
                If Not rsDiagnosis Is Nothing Then
                    rsDiagnosis.Filter = "�������=1"
                    If Not rsDiagnosis.EOF Then
                        txt�������.Text = Nvl(rsDiagnosis!�������): txt�������.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl�������.Tag = txt�������.Text
                    Else
                        txt�������.Text = ""
                    End If
                    If txt��ҽ���.Enabled Then
                        rsDiagnosis.Filter = "�������=11"
                        If Not rsDiagnosis.EOF Then
                            txt��ҽ���.Text = Nvl(rsDiagnosis!�������): txt��ҽ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                        End If
                    Else
                        txt��ҽ���.Text = ""
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cboԤ������_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long, strInfo As String
    

    With mCurPrepay
        .lngҽ�ƿ����ID = 0
        .bln���ѿ� = False
        .str���㷽ʽ = ""
        .str���� = ""
    End With
    '130245,�л����㷽ʽ��ͬ�����¿����ID
    If mbytInState <> 2 Then
        Call SetCardVaribles(True)
    End If
    If mblnNotClick = True Then Exit Sub
    '��֧Ʊ���ֽ�������,����������
    If InStr(cboԤ������.Text, "֧Ʊ") > 0 Then
        If Not mrsInfo Is Nothing And IsNumeric(txtPatient.Tag) Then
            strInfo = GetLastInfo(CLng(txtPatient.Tag))
            If strInfo <> "" Then
                txt�ɿλ.Text = IIf(Split(strInfo, "|")(0) = "", txt�ɿλ.Text, Split(strInfo, "|")(0))
                txt������.Text = IIf(Split(strInfo, "|")(1) = "", txt������.Text, Split(strInfo, "|")(1))
                txt�ʺ�.Text = IIf(Split(strInfo, "|")(2) = "", txt�ʺ�.Text, Split(strInfo, "|")(2))
            End If
        End If
    Else
        txt�ɿλ.Text = ""
        txt������.Text = ""
        txt�ʺ�.Text = ""
    End If
    
    If is�����ʻ�(cboԤ������) Then
        txt�ɿλ.BackColor = Me.BackColor
        txt������.BackColor = Me.BackColor
        txt�ʺ�.BackColor = Me.BackColor
        
        txt�ɿλ.Enabled = False
        txt������.Enabled = False
        txt�ʺ�.Enabled = False
    Else
        txt�ɿλ.BackColor = &H80000005
        txt������.BackColor = &H80000005
        txt�ʺ�.BackColor = &H80000005
        
        txt�ɿλ.Enabled = True
        txt������.Enabled = True
        txt�ʺ�.Enabled = True
    End If
    
    '54979:������,2012-10-22
    If txt�ɿλ.Text <> "" And txt�ɿλ.Enabled = True Then
        chk��λ�ɿ�.Value = 1
        If txt�ɿλ.Enabled = False Then Call chk��λ�ɿ�_Click
    Else
        chk��λ�ɿ�.Value = 0
        If txt�ɿλ.Enabled = True Then Call chk��λ�ɿ�_Click
    End If
    
    '0=����,1=�޸�,2=�鿴
    If mbytInState = 2 Then Exit Sub
    Call Local���㷽ʽ(mCurPrepay.lngҽ�ƿ����ID, False)
End Sub

Private Function is�����ʻ�(cbo As Object) As Boolean
    If cbo.ListIndex <> -1 Then
        If cbo.ItemData(cbo.ListIndex) = 3 Then
            is�����ʻ� = True
        End If
    End If
End Function

Private Sub cbo��λ_Click()
    cbo.SetListWidth cbo��λ.hWnd, cbo��λ.width * 2.9
    If cbo��λ.Text = "�����䴲λ" Then
        chk���.TabStop = False
    Else
        chk���.TabStop = True
    End If
    If mblnAppoint Then cbo��λ.Tag = Trim(Split(Trim(cbo��λ.Text), " ")(0))
End Sub

Private Sub LoadDept(ByVal bytType As Byte)
'���ܣ����ݿ��Ҽ��ز���������ݲ������ؿ��ң���󣬼��ض�Ӧ�Ĵ�λ
'������bytType=0-���ݿ��Ҽ��ز���,1-���ݲ������ؿ���
    Dim lngDept As Long, lngUnit As Long
    Dim strFilter As String, i As Long
    
    If cbo��Ժ����.ListIndex >= 0 Then lngUnit = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    If cbo��Ժ����.ListIndex >= 0 Then lngDept = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    
    If gbln��ѡ���� And bytType = 1 Then
        '���ݲ������ؿ���
        mrsUnitDept.Filter = "����ID=" & lngUnit
        For i = 1 To mrsUnitDept.RecordCount
            strFilter = strFilter & IIf(strFilter = "", "", " Or ") & "ID=" & mrsUnitDept!����ID
            mrsUnitDept.MoveNext
        Next
        
        '*********************************************************
        '���� 25682 by lesfeng 2009-10-12 b
        If strFilter = "" Then
            cbo��Ժ����.Clear
        Else
            mrsDept.Filter = strFilter
            Call CboLoadData(cbo��Ժ����, mrsDept, True)
        End If
        '���� 25682 by lesfeng 2009-10-12 e
        
        i = cbo.FindIndex(cbo��Ժ����, lngUnit)
        If i = -1 Then
            i = cbo.FindIndex(cbo��Ժ����, lngDept)
            If i = -1 Then i = 0
        End If
        cbo.SetIndex cbo��Ժ����.hWnd, i
        cbo��Ժ����.TabStop = (cbo��Ժ����.ListCount > 1)
        '����27370 by lesfeng 2010-01-26
        cbo��Ժ����.SelLength = 0
    ElseIf Not gbln��ѡ���� And bytType = 0 Then
        '���ݿ��Ҽ��ز���
        mrsUnitDept.Filter = "����ID=" & lngDept
        For i = 1 To mrsUnitDept.RecordCount
            strFilter = strFilter & IIf(strFilter = "", "", " Or ") & "ID=" & mrsUnitDept!����ID
            mrsUnitDept.MoveNext
        Next
        mrsUnit.Filter = strFilter
        
        cbo��Ժ����.Clear
        cbo��Ժ����.AddItem "��ȷ������"
        cbo��Ժ����.ItemData(cbo��Ժ����.NewIndex) = 0
        Call CboLoadData(cbo��Ժ����, mrsUnit, False)
        
        i = cbo.FindIndex(cbo��Ժ����, lngUnit)
        If i = -1 Then i = 0
        cbo.SetIndex cbo��Ժ����.hWnd, i
        cbo��Ժ����.TabStop = (cbo��Ժ����.ListCount > 1)
        '����27370 by lesfeng 2010-01-26
        cbo��Ժ����.SelLength = 0
    End If
    
    '����26779 by lesfeng 2009-12-10
    lngUnit = 0
    lngDept = 0
    If cbo��Ժ����.ListIndex >= 0 Then lngUnit = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    If cbo��Ժ����.ListIndex >= 0 Then lngDept = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    '���ش�λ
    If gbln��Ժ��� And mbytMode <> EMode.EԤԼ�Ǽ� And mbytInState = EState.E���� Then
        Call LoadBed(zlCommFun.GetNeedName(cbo�Ա�.Text), lngDept, lngUnit)
    End If
    
    Call LoadBedInfo(lngDept, lngUnit)
End Sub

Private Sub cbo��Ժ����_Click()
    Dim strDoctors As String, i As Long, lngDepID As Long
    Dim rsDiagnosis As ADODB.Recordset
    
    If cbo��Ժ����.ListIndex <> -1 Then
        lngDepID = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
        
        '�ÿ��Ҷ�Ӧ�Ĳ���,��λ
        If mbytInState <> EState.E���� Then Call LoadDept(0)
        
        '�Ƿ�����ҽ��
        If mbytMode <> 1 Then txt��ҽ���.Enabled = (InStr(1, "," & GetDepCharacter(lngDepID) & ",", ",��ҽ��,") > 0)
        txt��ҽ���.ToolTipText = "ֻ�е���Ժ���ҵ�����Ϊ��ҽ��ʱ������������ҽ���!"
        
        '�Ƿ�����Ժ
        If mbytInState = 0 And Not mrsInfo Is Nothing Then
            chk����Ժ.Value = IIf(CheckReIN(mrsInfo!����ID, lngDepID), 1, 0)
        End If
        
        '107823��ʾ����������
        If (mbytMode = EMode.E����ԤԼ Or mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0 And mbytInState = EState.E����) And Me.Visible = True Then
            Set rsDiagnosis = GetDiagnosticInfo(mlng����ID, mlng��ҳID, "1,11", "3", lngDepID)
            If Not rsDiagnosis Is Nothing Then
                rsDiagnosis.Filter = "�������=1"
                If Not rsDiagnosis.EOF Then
                    txt�������.Text = Nvl(rsDiagnosis!�������): txt�������.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl�������.Tag = txt�������.Text
                Else
                    txt�������.Text = ""
                End If
                If txt��ҽ���.Enabled Then
                    rsDiagnosis.Filter = "�������=11"
                    If Not rsDiagnosis.EOF Then
                        txt��ҽ���.Text = Nvl(rsDiagnosis!�������): txt��ҽ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                    End If
                Else
                    txt��ҽ���.Text = ""
                End If
            End If
        End If
    Else
        txt��ҽ���.Enabled = False
        txt��ҽ���.ToolTipText = "ֻ�е���Ժ���ҵ�����Ϊ��ҽ��ʱ������������ҽ���!"
    End If
End Sub

Private Sub cbo�Ա�_Click()
    Dim lngDept As Long, lngUnit As Long
    
    If Not cbo�Ա�.Visible Then Exit Sub
    
    If cbo��Ժ����.ListIndex >= 0 Then lngUnit = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    If cbo��Ժ����.ListIndex >= 0 Then lngDept = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    Call LoadBed(zlCommFun.GetNeedName(cbo�Ա�.Text), lngDept, lngUnit)
    Call ReLoadCardFee
End Sub

Private Sub chkUnlimit_Click()
     '���޵�����������õ���ʱ��,���Ҳ�������ʱ����
     
    dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm"
    If chkUnlimit.Value = 1 And IsNull(dtp����ʱ��.Value) Then
        dtp����ʱ��.Value = DateAdd("d", 3, CDate(txt��Ժʱ��.Text))
    End If
    
    chk��ʱ����.Enabled = Not (chkUnlimit.Value = 1)
    txt������.Enabled = Not (chkUnlimit.Value = 1)
    If chkUnlimit.Value = 1 Then
        txt������.Text = "999999999":  txt������.BackColor = vbInactiveCaptionText
    Else
        txt������.Text = "": txt������.BackColor = vbWhite
    End If
End Sub


Private Sub chk����_Click()
    If chk����.Value = Checked Then
        cbo��������.Enabled = False
        If Visible Then cmdOK.SetFocus
    Else
        cbo��������.Enabled = True
        cbo��������.SetFocus
    End If
End Sub

Private Sub chk��ʱ����_Click()
    If chk��ʱ����.Value = 1 Then
        '��ʱ���޶�,����������ʱ����
        dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm"
        dtp����ʱ��.Value = Null
        chkUnlimit.Value = 0        'ֵ�ı�ʱ����ʽ����click�¼�
    End If
    chkUnlimit.Enabled = Not (chk��ʱ����.Value = 1)
    dtp����ʱ��.Enabled = Not (chk��ʱ����.Value = 1)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Function Get������(lng����ID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select ���� From ���ղ��� Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then Get������ = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckReIN(ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    strSQL = "Select ����id" & vbNewLine & _
            " From ������ҳ a" & vbNewLine & _
            " Where ����id = [1] And Nvl(a.��ҳid, 0) <> 0 And Exists" & vbNewLine & _
            "       (Select 1" & vbNewLine & _
            "            From �ٴ����� b" & vbNewLine & _
            "            Where b.����id = a.��Ժ����id And b.�������� = (Select �������� From �ٴ����� Where ����id = [2] And Rownum = 1))"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng����ID)
    CheckReIN = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdYB_Click()
    Dim lng����ID As Long, lng����ID As Long
    Dim objCurrent As Object, strTxt As String, arrTxt As Variant
    Dim i As Long, blnDo As Boolean, arrPati As Variant
    Dim objcbo As ComboBox
    
    If (mbytMode = EMode.E����ԤԼ Or mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0) Then
        lng����ID = mlng����ID
    ElseIf Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If MsgBox("��ǰ�Ѿ�����һ������,�Ƿ�Ҫ�Ըò��˵���ݽ�����֤��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                lng����ID = mrsInfo!����ID
            End If
        End If
    End If
    
    'ҽ���Ķ�
    mstrYBPati = gclsInsure.Identify(1, lng����ID, mintInsure)
    mstrYBPatiBak = mstrYBPati '�Զ�����ҽ����Ϣ���б��ݣ��Ա�����ԤԼ���˺ϲ���ָ�
    mintInsureBak = mintInsure
    If mstrYBPati <> "" Then
        arrPati = Split(mstrYBPati, ";")
        '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID,...
        If UBound(arrPati) >= 8 Then
            If Val(arrPati(8)) > 0 Then
                txtPatient.Text = "-" & Val(arrPati(8))
                blnDo = txtPatient.Locked
                txtPatient.Locked = False
                Call txtPatient_KeyPress(13)
                txtPatient.Locked = blnDo
                If mstrYBPati = "" Then txt����.SetFocus: Exit Sub  '������Ϊ��������ѡ�����˳���,������clearcard
            ElseIf mrsInfo Is Nothing Then
                If txtPatient.Tag = "" Then '�����δ����
                    txtPatient.Text = zlDatabase.GetNextNo(1) '�²���ID
                    txtPatient.Tag = txtPatient.Text
                    If txtסԺ��.Visible And mbytKind = EKind.EסԺ��Ժ�Ǽ� Then
                        txtסԺ��.Text = zlDatabase.GetNextNo(2)
                    ElseIf txtסԺ��.Visible And mbytKind = EKind.EסԺ���۵Ǽ� Then
                        txtסԺ��.Text = zlDatabase.GetNextNo(6)
                    End If
                End If
            End If
        End If
        
        txtҽ����.Text = arrPati(1)
        txtҽ����.Locked = True
        
        txt����.Text = arrPati(3)
        cbo�Ա�.ListIndex = GetCboIndex(cbo�Ա�, CStr(arrPati(4)))
        If IsDate(arrPati(5)) Then
            txt��������.Text = Format(arrPati(5), "yyyy-MM-dd")
            Call txt��������_LostFocus
        End If
        txt���֤��.Text = arrPati(6)
        txt������λ.Text = arrPati(7)
       
        '���ղ�����Ϊ��Ժ���
        If UBound(arrPati) >= 14 Then
            If Val(arrPati(14)) > 0 Then
                lng����ID = Val(arrPati(14))
                
                If txt�������.Text = "" And Not RequestCode Then
                    txt�������.Text = Get������(lng����ID)
                End If
            End If
        End If
        
        '��ȡ�����ʻ����
        mcurYBMoney = gclsInsure.SelfBalance(Val(arrPati(8)), CStr(arrPati(1)), 20, , mintInsure)
        mcurYBMoneyBak = mcurYBMoney
        lblYBMoney.Caption = "�����ʻ���" & Format(mcurYBMoney, "0.00")
        lblYBMoney.Visible = True
        
        'ҽ�Ƹ��ʽȱʡ=������ҽ�Ʊ���
        For i = 0 To cboҽ�Ƹ���.ListCount
            If InStr(cboҽ�Ƹ���.List(i), Chr(&HD)) > 0 Then cboҽ�Ƹ���.ListIndex = i: Exit For
        Next
        
        If Not IsDate(txt��������.Text) Then
            txt��������.SetFocus
        Else
            strTxt = "txt����,cbo�Ա�,cbo�ѱ�,cbo����,cbo����,cboѧ��,cbo����״��,cboְҵ,cbo���," & _
                     "txt���֤��,txt�����ص�,txt��ͥ��ַ,txt��ͥ��ַ�ʱ�,txt��ͥ�绰,txt���ڵ�ַ,txt���ڵ�ַ�ʱ�,txt������λ,txt��λ�绰,txt��λ�ʱ�," & _
                     "txt��λ������,txt��λ�ʺ�,txt��ϵ������,cbo��ϵ�˹�ϵ,txt��ϵ�˵�ַ,txt��ϵ�˵绰,txt��ϵ�����֤��,txt������,txt������"
            arrTxt = Split(strTxt, ",")
            i = 0
            For i = 0 To UBound(arrTxt)
                For Each objCurrent In Me.Controls
                    If objCurrent.Name = arrTxt(i) Then
                        blnDo = False
                        If TypeOf objCurrent Is TextBox Then
                            If Trim(objCurrent.Text) = "" And objCurrent.Enabled = True Then blnDo = True
                        ElseIf TypeOf objCurrent Is ComboBox Then
                            Set objcbo = objCurrent
                            If objcbo.ListIndex = -1 And objCurrent.Enabled = True Then blnDo = True
                        End If
                        If blnDo Then
                            If objCurrent.TabStop Then
                                Call SetChargeTurn
                                If objCurrent.Visible Then objCurrent.SetFocus
                                Exit Sub
                            End If
                        End If
                        GoTo exitHandle
                    End If
                Next
exitHandle:
            Next
        End If
        Call SetChargeTurn
        If CanFocus(cbo��Ժ����) Then cbo��Ժ����.SetFocus
    Else
        txt����.SetFocus
    End If
End Sub

Private Sub SetChargeTurn()
    Dim dat��Ժʱ�� As Date
    
    '�������תסԺ���
    dat��Ժʱ�� = CDate(txt��Ժʱ��.Text)
    If frmChargeTurn.CheckExistTurn(Val(txtPatient.Text), dat��Ժʱ��) Then
        MsgBox "�ò����Ѵ�������תסԺ�ĵ���!" & vbCrLf & _
                "��Ժʱ�佫���̶�Ϊ��Щ���ݵ������ʱ�䡣", vbInformation, Me.Caption
        txt��Ժʱ��.Text = Format(dat��Ժʱ��, "yyyy-MM-dd HH:mm")
        txt��Ժʱ��.Enabled = False
    End If
    '����:33635
    If mstrYBPati <> "" Then
        cmdTurn.Visible = True
    Else
        cmdTurn.Visible = InStr(1, mstrPrivs, ";�������תסԺ;") > 0 And mbytKind = EסԺ��Ժ�Ǽ� And mbytMode <> 1
    End If
End Sub

Private Sub cmd��λ��ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & _
            " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID", _
            2, "��λ", , txt������λ.Text)
    If Not rsTmp Is Nothing Then
        txt������λ.Tag = rsTmp!ID
        txt������λ.Text = rsTmp!����
        txt������λ.SelStart = Len(txt������λ.Text)
        txt��λ�绰.Text = Trim(rsTmp!�绰 & "")
        txt��λ������.Text = Trim(rsTmp!�������� & "")
        txt��λ�ʺ�.Text = Trim(rsTmp!�ʺ� & "")
        txt������λ.SetFocus
    End If
End Sub

Private Sub dtp����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        If dtp����ʱ��.CheckBox Then
            KeyAscii = 0
            If IsNull(dtp����ʱ��.Value) Then
                dtp����ʱ��.Value = DateAdd("d", 3, zlDatabase.Currentdate)
            Else
                dtp����ʱ��.Value = Null
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim cboTmp As ComboBox, lngIdx As Long
    
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
        If Me.ActiveControl.Name = txt�������.Name Then
            If InStr(":��;��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc("'") Then
        If Not (Me.ActiveControl Is txt������� Or Me.ActiveControl Is txt��ҽ���) Then KeyAscii = 0      '��������п�����'��
    ElseIf KeyAscii >= 32 And TypeName(Me.ActiveControl) = "ComboBox" Then
        Set cboTmp = Me.ActiveControl
        If cboTmp.Style = 2 Then   'Ŀǰcbo����ҽʦ����
            lngIdx = cbo.MatchIndex(cboTmp.hWnd, KeyAscii, 0.8)
            If lngIdx = -1 And cboTmp.ListCount > 0 Then lngIdx = 0
            cboTmp.ListIndex = lngIdx
        End If
    End If
    
    '��ϵ�˹�ϵ˵����ת�벻����¼�붺�ź�ð��,��Ϊ �ö���mstrPatiPlus�� �ķָ��� ����ð�źͶ���
    If Me.ActiveControl Is txtת�� Or Me.ActiveControl Is txtLinkManInfo Then
        If InStr(":��,��", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub Form_Load()
    mblnChange = True
'    Call zlCardSquareObject
    
    With mPageHeight
        .���� = Me.Height
        .�������� = Me.Height
    End With
    Call CreateObjectKeyboard
    Call CreatePublicExpenseObject(mlngModul)
    mstrPrePriceGrade = ""
    '��ʼ��
    If Not gobjSquare Is Nothing Then Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    
    'ҽ������
    mintInsure = 0
    mintInsureBak = mintInsure
    mstrYBPati = ""
    mcurYBMoney = 0
    mdat�ϴε�������ʱ�� = CDate("0:00:00")
    mstrNOS = ""
    
    '����26779 by lesfeng 2009-12-10
    lblBedInfo.Caption = ""
        
    mblnUnload = False
    gblnOK = False
        
    '����27356 by lesfeng 2010-01-13
    If InStr(mstrPrivs, "�󶨿���") = 0 Then
        tabCardMode.Tabs.Remove ("CardBind")
'        tabCardMode.Tabs("CardBind").Selected = True
'        tabCardMode.Tabs("CardBind").Caption = "�󶨿���"
        tabCardMode.width = tabCardMode.width / 2
    End If
    If mbytMode = 2 Then mblnUnload = Not isValid(mlng����ID)
    Call InitDicts
    If Not InitData Then mblnUnload = True
    If mblnUnload Then Unload Me: Exit Sub
    '�����:56599
    Call InitFace
    Call InitTabPage
    '����27370 by lesfeng 2010-01-26
    cbo��Ժ����.SelLength = 0
    cbo��Ժ����.SelLength = 0
    If mblnUnload Then Unload Me: Exit Sub

    mFactProperty = zl_GetInvoicePreperty(mlngModul, 2, 2)
    
    If gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
    End If
    '����д������
    Call zlCreateSquare
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, P������Ժ����, mstrPrivs)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
End Sub

Private Sub Init������Ϣ(dat��Ժʱ�� As Date)
    
    txt������.Text = ""
    If mbytInState <> 2 Then chkUnlimit.Enabled = True
    chkUnlimit.Value = 0     '���ֵ�б仯,����ʽ����click�¼�
    txt������.Text = ""
    
    If mbytInState <> 2 Then dtp����ʱ��.Enabled = True
    dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm"    '����checkbox�ɼ���
    
    If mbytInState = 0 And mbytMode <> EMode.EԤԼ�Ǽ� Then
        '����ʱ,����ʱ�䲻��С����Ժʱ��(�޸�ʱ�ڶ�ȡ��Ƭ���޸���Ժʱ��ʱ��������)
        dtp����ʱ��.MinDate = dat��Ժʱ��
        dtp����ʱ��.Value = DateAdd("d", 3, dat��Ժʱ��)
    End If
    dtp����ʱ��.Value = Null
    
    If mbytInState <> 2 Then chk��ʱ����.Enabled = True
    chk��ʱ����.Value = 0
    txtReason.Text = ""
End Sub

Private Sub InitFace()
    Dim blnHaveCard As Boolean, dat��Ժʱ�� As Date
    Dim lngTmp As Long, blnԤ�� As Boolean, bln�ſ� As Boolean
    Dim strסԺ�� As String
    
    Call InitvsDrug
    Call InitVsInoculate
    Call InitVsOtherInfo
    Call InitCertificate
    Call InitCombox
    '���ýṹ����ַ����
    Call InitStructAddress

    If mbytInState <> E���� Then
        txt����.MaxLength = GetColumnLength("������Ϣ", "����")
        txt����.MaxLength = GetColumnLength("������Ϣ", "����")
        txtסԺ��.MaxLength = GetColumnLength("������Ϣ", "סԺ��")
    End If
    
    '�������
    If mbytMode = E�޸� Then
        If mbytKind = EסԺ��Ժ�Ǽ� Then
            Caption = "ԤԼ��Ժ�Ǽ�"
        ElseIf mbytKind = E�������۵Ǽ� Then
            Caption = "ԤԼ��������"
        ElseIf mbytKind = EסԺ��Ժ�Ǽ� Then
            Caption = "ԤԼסԺ����"
        End If
    ElseIf mbytMode = 2 Then
        If mbytKind = EסԺ��Ժ�Ǽ� Then
            Caption = "����סԺ����"
        ElseIf mbytKind = E�������۵Ǽ� Then
            Caption = "������������"
        ElseIf mbytKind = EסԺ���۵Ǽ� Then
            Caption = "����סԺ����"
        End If
    Else
        If mbytKind = EסԺ��Ժ�Ǽ� Then
            Caption = "������Ժ�Ǽ�"
        ElseIf mbytKind = E�������۵Ǽ� Then
            Caption = "�������۵Ǽ�"
        ElseIf mbytKind = EסԺ���۵Ǽ� Then
            Caption = "סԺ���۵Ǽ�"
        End If
    End If
    Me.Tag = Me.Caption
    mbytKindBak = mbytKind
    
    Call InitInputTabStop
    
    If InStr(mstrPrivs, "��Լ���˵Ǽ�") = 0 Then
        txt������λ.Enabled = False
        txt������λ.BackColor = &H8000000F
        txt��λ�绰.Enabled = False
        txt��λ�绰.BackColor = &H8000000F
        txt��λ�ʱ�.Enabled = False
        txt��λ�ʱ�.BackColor = &H8000000F
        txt��λ������.Enabled = False
        txt��λ������.BackColor = &H8000000F
        txt��λ�ʺ�.Enabled = False
        txt��λ�ʺ�.BackColor = &H8000000F
        cmd������λ.Visible = False
    End If
    
    
    'ҽ��:1.�����ӻ�Ȩ��,2.ԤԼ�Ǽ�,3.��������,4.����ִ�еǼ�
    'ҽ��һԺ��ԤԼ����ҽ���鿨 �����޸�����ȥ��  And mbytMode <> 1
    cmdYB.Visible = InStr(mstrPrivs, "���ղ��˵Ǽ�") > 0 And mbytKind <> E�������۵Ǽ� And mbytInState = 0
    cmdTurn.Visible = InStr(1, mstrPrivs, ";�������תסԺ;") > 0 And mbytKind = EסԺ��Ժ�Ǽ� And mbytMode <> 1
    txtTimes.Visible = mbytMode <> 1 And mbytKind = EסԺ��Ժ�Ǽ� 'ԤԼ�Ǽ�ʱ�����صǼ�ʱ,סԺ����Ϊ��
    lblTimes.Visible = mbytMode <> 1 And mbytKind = EסԺ��Ժ�Ǽ�
    cmdName.Visible = mbytMode = 2
    txtTimes.Enabled = (InStr(1, mstrPrivs, "�޸�סԺ����") > 0 And mbytInState = 0)   '�޸�ʱ������ģ���Ϊ�����Ѳ���סԺһ�η��ã�Ԥ������￨
        
    IDKind.Enabled = False
    If mbytMode = 0 Or mbytMode = 1 Then
        If mbytInState = 0 Then
            Set mobjIDCard = New clsIDCard
            Set mobjICCard = New clsICCard
            Call mobjIDCard.SetParent(Me.hWnd)
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
            IDKind.Enabled = True
        End If
    End If
    
        
    'ԤԼ�Ǽ�ʱ����д������
    If mbytMode = 1 Then
        txtҽ����.Enabled = False
        txtҽ����.BackColor = Me.BackColor
        txt�������.Enabled = False
        txt�������.BackColor = Me.BackColor
        txt��ҽ���.Enabled = False
        txt��ҽ���.BackColor = Me.BackColor
    End If
    
    'סԺ��
    If mbytKind = E�������۵Ǽ� Then     '��������
        lblסԺ��.Caption = "�����"
        cmdSelectNO.Visible = False
        lbl����.Left = lbl�ѱ�.Left
        txt����.Left = cbo�ѱ�.Left
        txt����.width = cbo�ѱ�.width
        cmdName.Left = txt����.Left + txt����.width - cmdName.width - 20
        
        cmdYB.Visible = False
    ElseIf mbytKind = EסԺ���۵Ǽ� Then     'סԺ����
        lblסԺ��.Caption = "���ۺ�"
        txtסԺ��.TabStop = False
        txtסԺ��.Locked = True
        cmdSelectNO.Visible = False
    End If
    
    If InStr(mstrPrivs, "�޸�סԺ��") = 0 Then
        txtסԺ��.Locked = True
        txtסԺ��.TabStop = False
        txtסԺ��.BackColor = Me.BackColor
        cmdSelectNO.Visible = False
    End If
    If mbytInState = EState.E���� Then cmdSelectNO.Visible = False
    
    If InStr(mstrPrivs, "�޸���Ժ����") = 0 Then
        txt��Ժʱ��.Enabled = False
    End If
        
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    mblnChange = False: cbo���䵥λ.ListIndex = 0: cbo���䵥λ.Tag = txt����.Text & "_" & cbo���䵥λ.Text: mblnChange = True
    
    '����,��Ժ�Ǽǻ���ԺԤԼ����
    If mbytInState = 0 Then dat��Ժʱ�� = zlDatabase.Currentdate           '����ʱ,��������ʱ�䲻��С����Ժʱ��
        
    '������Ϣ
    If mbytInState = 2 Or (mbytMode <> 1 And InStr(mstrPrivs, "������Ϣ") > 0 And gbln����) Then
        Call Init������Ϣ(dat��Ժʱ��)
    End If
    '51167,������,2012-07-09,����"��ϵ�����֤��"
    
    'ԤԼ�Ǽǲ�֧�ֵǼǵ�����Ϣ(��Ϊû����ҳID)
    If mbytMode = 1 Or mbytInState <> 2 And InStr(mstrPrivs, "������Ϣ") = 0 Then
        pic����.Visible = False
        fra����.Height = fra����.Height - pic����.Height
        pic����.Height = pic����.Height - pic����.Height
        Me.Height = Me.Height - pic����.Height
    Else
        If mbytInState <> 2 And Not gbln���� Then
            txt������.Enabled = False:        txt������.BackColor = Me.BackColor
            txt������.Enabled = False:        txt������.BackColor = Me.BackColor
            txtReason.Enabled = False:        txtReason.BackColor = Me.BackColor
            chkUnlimit.Enabled = False:       chk��ʱ����.Enabled = False
            lbl����ʱ��.Enabled = False:      dtp����ʱ��.Enabled = False
        End If
    End If
    
    If InStr(mstrPrivs, "������Ϣ") = 0 Then cmdSurety.Visible = False

    '���������
    If gbln��ѡ���� Then
        lngTmp = lbl��Ժ����.Left
        lbl��Ժ����.Left = lbl��Ժ����.Left
        lbl��Ժ����.Left = lngTmp
        
        lngTmp = cbo��Ժ����.Left
        cbo��Ժ����.Left = cbo��Ժ����.Left
        cbo��Ժ����.Left = lngTmp
        
        lngTmp = cbo��Ժ����.TabIndex
        cbo��Ժ����.TabIndex = cbo��Ժ����.TabIndex
        cbo��Ժ����.TabIndex = lngTmp
    End If
    Call cbo.SetListWidth(cbo��Ժ����.hWnd, cbo.ListWidth(cbo��Ժ����.hWnd) * 1.2)
    
    If Not (gbln��Ժ��� And mbytMode <> EMode.EԤԼ�Ǽ�) Or mbytInState = EState.E�޸� Then
        lbl��λ.Visible = False
        cbo��λ.Visible = False
        chk���.Visible = False
    End If
    
    Select Case mbytInState         '0=����,1=�޸�,2=����
        Case E����
           mFactProperty = zl_GetInvoicePreperty(mlngModul, 2, 2)
            If Not gobjSquare Is Nothing Then
                If Not gobjSquare.objSquareCard Is Nothing Then IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
            End If
        
            If InStr(mstrPrivs, "�����ҽ������") = 0 Then
                txtPatient.TabStop = False
                txtסԺ��.TabStop = False
            End If
            
            Call InitSendCardPreperty
            chk����.Value = IIf(gbln���� = True, 1, 0)
            
            '��ȡԤ����Ʊ��
            Call GetFact(True)
            
            '����ԤԼ�Ǽ�ʱ������п��˾Ͳ������ٷ���
            blnHaveCard = False
            If mbytMode = EMode.E����ԤԼ Then
                blnHaveCard = PatiHaveCard(mlng����ID)
            End If
            
            'ԤԼ�Ǽ�ʱ������λ����,�����䲡��
            blnԤ�� = gbln��ԺԤ�� And (cboԤ������.ListCount > 0) And (gblnPrepayStrict And mlngԤ������ID > 0 Or Not gblnPrepayStrict)
            '76824�����ϴ���2014/8/19��ҽ�ƿ���������
            bln�ſ� = (gbln��Ժ���� And (mCurSendCard.bln�ϸ���� And mCurSendCard.lng����ID > 0 Or Not mCurSendCard.bln�ϸ����) And Not blnHaveCard _
                    Or mCurSendCard.blnOneCard And mCurSendCard.bln�ϸ����) And mCurSendCard.lng�����ID <> 0
            
            Call HideCard(blnԤ��, bln�ſ�)
            If mbytMode = EMode.E����ԤԼ Then
                txtPatient.Locked = True
                txtPatient.TabStop = False

                '��ʾ������Ϣ
                If Not ReadPatiReg(mlng����ID, mlng��ҳID) Then
                    MsgBox "������ȷ��ȡ�ò��˵ĵǼǼ�¼��", vbInformation, gstrSysName
                    mblnUnload = True: Exit Sub
                End If
                
                '50511,������,2013-11-04,ֻ�о��е�������ҽʦȨ�޲����޸�����ҽʦ
                If InStr(mstrPrivs, ";��������ҽʦ;") = 0 And cbo����ҽʦ.ListIndex <> -1 Then
                    cbo����ҽʦ.Enabled = False
                End If
                
                '���֮ǰû��סԺ�Ż�ÿ��סԺ������סԺ��,����ΪסԺ���ˣ����Զ������µ�סԺ��
                '���� 27063 by lesfeng 2009-12-25 ԤԼ�Ǽ�תסԺ���˱���ԭסԺ��(ȡ��gblnÿ��סԺ��סԺ���ж�)
'                If mbytKind = EKind.EסԺ��Ժ�Ǽ� And (Trim(txtסԺ��.Text) = "" Or gblnÿ��סԺ��סԺ��) Then txtסԺ��.Text = zlDatabase.GetNextNo(2)
                '85510:LPF,2015-06-19,ԤԼ�Ǽ�סԺ�Ų�������ҽ���Ǽ���Ժ����,��ҽ���Ǽǲ���סԺ��ʱ��������дסԺҵ�����:
                'ԭ���߼��ж�:If mbytKind = EKind.EסԺ��Ժ�Ǽ� And (Trim(txtסԺ��.Text) = "") Then txtסԺ��.Text = zlDatabase.GetNextNo(2)
                '��Ժ����ԤԼ�Ǽǻ���ݲ���"ÿ��סԺ��סԺ��"����סԺ��,��ҽ���Ǽ�Ŀǰֻ���Բ�����Ϣ��סԺ��Ϊ׼����(���ַ�ʽ������סԺ�ſ��ܾͲ���ȷ)
                '�����Ҫ�����´���
                '1:gblnÿ��סԺ��סԺ��=TRUE,�������סԺ�ţ��������е�סԺ���Ƿ��ظ�������ظ����������ɡ�
                '2:gblnÿ��סԺ��סԺ��=FALSE,���סԺ��Ϊ��,��ʹ����ʷסԺ��(���һ��סԺ�Ų�Ϊ��)����������ʷסԺ���������ɡ�
                If mbytKind = EKind.EסԺ��Ժ�Ǽ� Then
                    If gblnÿ��סԺ��סԺ�� = True Then
                        If Trim(txtסԺ��.Text) <> "" Then
                            If CheckByPatiNO(mlng����ID, mlng��ҳID, 0, Trim(txtסԺ��.Text)) = True Then txtסԺ��.Text = ""
                        End If
                    Else
                        If Trim(txtסԺ��.Text) = "" Then
                            strסԺ�� = ""
                            If CheckByPatiNO(mlng����ID, mlng��ҳID, 1, strסԺ��) = True Then txtסԺ��.Text = strסԺ��
                        End If
                    End If
                    If Trim(txtסԺ��.Text) = "" Then txtסԺ��.Text = zlDatabase.GetNextNo(2)
                ElseIf mbytKind = EסԺ���۵Ǽ� Then
                    If Trim(txtסԺ��.Text) = "" Then txtסԺ��.Text = zlDatabase.GetNextNo(6)
                End If
            Else
                txt��Ժʱ��.Text = Format(dat��Ժʱ��, "yyyy-MM-dd HH:mm")
            End If
            '89980���˽ṹ�� ������������ȱʡֵ
            If gbln���ýṹ����ַ Then
                Call LoadStructAddressDef(marrAddress)
                Call SetStrutAddress(2)
            End If

        Case E�޸�    '�޸�
            If Not gobjSquare Is Nothing Then
                If Not gobjSquare.objSquareCard Is Nothing Then IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
            End If
            '�����ٴ��޸Ĳ�����Ϣ
            txtPatient.Locked = True
            txtPatient.TabStop = False
            
            '�ڹ����嵥����������,����Ʋ��˲�׼�޸���ס��Ϣ(������,ת����)
            Call HideCard(False, False)
            
            '65781:������,2013-11-15,���������ҽ���������޸��������Ա����䡢��������
            If HavedDirections(mlng����ID, mlng��ҳID) Then
                mblnHaveAdvice = True
                txt����.Locked = True
                txt����.BackColor = &H80000016
                txt��������.Enabled = False
                txt��������.BackColor = txt����.BackColor
                txt����ʱ��.Enabled = False
                txt����ʱ��.BackColor = txt����.BackColor
                txt����.Locked = True
                txt����.BackColor = txt����.BackColor
                cbo���䵥λ.Locked = True
                cbo���䵥λ.BackColor = txt����.BackColor
                cbo�Ա�.Locked = True
                cbo�Ա�.BackColor = txt����.BackColor
            Else
                mblnHaveAdvice = False
            End If
            
            '�������ò����޸Ĳ�������
            If HavedInCost(mlng����ID, mlng��ҳID) Then
                txt����.Locked = True
            End If
            
            If Not ReadPatiReg(mlng����ID, mlng��ҳID) Then
                MsgBox "������ȷ��ȡ�ò��˵ĵǼǼ�¼��", vbInformation, gstrSysName
                mblnUnload = True: Exit Sub
            End If
            
            '50511,������,2013-11-04,ֻ�о��е�������ҽʦȨ�޲����޸�����ҽʦ
            If InStr(mstrPrivs, ";��������ҽʦ;") = 0 And cbo����ҽʦ.ListIndex <> -1 Then
                cbo����ҽʦ.Enabled = False
            End If
            '101160
            Call EMPI_LoadPati
            
        Case E����   '����
            Call HideCard(False, False)
            Call SetStrutAddress
            'pic����.Enabled = False
            IDKind.Enabled = False
            txtPatient.Locked = True
            txtסԺ��.Locked = True
            cmdSelectNO.Enabled = False
            txt����.Locked = True
            cmdName.Enabled = False
            cmdYB.Enabled = False
            cmdTurn.Enabled = False
            txtҽ����.Locked = True
            txt����.Locked = True
            txt��������.Enabled = False
            txt����ʱ��.Enabled = False
            txt����.Locked = True
            cbo���䵥λ.Locked = True
            cbo�Ա�.Locked = True
            cbo�ѱ�.Locked = True
            cboҽ�Ƹ���.Locked = True
            txt���֤��.Locked = True
            cbo����.Locked = True
            cbo����.Locked = True
            cbo���.Locked = True
            txt����֤��.Locked = True
            cboѧ��.Locked = True
            cbo����״��.Locked = True
            cboְҵ.Locked = True
            txt��ͥ��ַ.Locked = True
            txt��ͥ�绰.Locked = True
            txt��ͥ��ַ�ʱ�.Locked = True
            txt���ڵ�ַ.Locked = True
            txt���ڵ�ַ�ʱ�.Locked = True
            txt����.Locked = True
            txt�����ص�.Locked = True
            txt����.Locked = True
            cmd����.Enabled = False
            cbo��������.Locked = True
            txt������λ.Locked = True
            txt��λ�绰.Locked = True
            txt��λ�ʱ�.Locked = True
            txt��λ������.Locked = True
            txt��λ�ʺ�.Locked = True
            txt��ϵ������.Locked = True
            txt��ϵ�˵�ַ.Locked = True
            txt��ϵ�˵绰.Locked = True
            cbo��ϵ�˹�ϵ.Locked = True
            txtLinkManInfo.Locked = True
            cmdת��.Enabled = False
            txtת��.Locked = True
            txt��ϵ�����֤��.Locked = True
            txt������.Locked = True
            chkUnlimit.Enabled = False
            txt������.Locked = True
            dtp����ʱ��.Enabled = False
            chk��ʱ����.Enabled = False
            txtReason.Locked = True
            txtMobile.Locked = True
            pic��Ժ.Enabled = False
            
            cmd���ڵ�ַ.Visible = False
            cmd����.Visible = False
            cmd������λ.Visible = False
            cmd�����ص�.Visible = False
            cmd��ͥ��ַ.Visible = False
            cmd��ϵ�˵�ַ.Visible = False
            cbo����ҽʦ.Enabled = False
            cbo��������.Enabled = False
            
            cboBloodType.Locked = True
            cboBH.Locked = True
            txtMedicalWarning.Locked = True
            txtOtherWaring.Locked = True
            cmdMedicalWarning.Visible = False
            cboIDNumber.Locked = True
            
            cmdOK.Visible = False
            cmdCancel.Caption = "�˳�(&X)"
            If Not ReadPatiReg(mlng����ID, mlng��ҳID) Then
                MsgBox "������ȷ��ȡ�ò��˵ĵǼǼ�¼��", vbInformation, gstrSysName
                mblnUnload = True: Exit Sub
            End If
            
    End Select
    'Ԥ�����տ���Ƿ���Ч
    If InStr(GetPrivFunc(glngSys, 1103), "Ԥ���տ�") = 0 And InStr(GetPrivFunc(glngSys, 1103), "���տ���ȡ") = 0 Then
        cmdDeposit.Visible = False
        '88434 1������ʱ���ж�Ԥ����Ƭ�Ƿ���Ч,�޸ĺͲ���ʱȱʡ���ɼ���2�����ǰ��������֧�Ѿ�����Ԥ�����ɼ�,�����ظ�����
        If mbytInState = 0 And mbln�Ƿ���ʾԤ�� Then
            Call HideCard(False)
        End If
    End If
    
    If InStr(mstrPrivs, "������������") = 0 Then
        cbo��������.Enabled = False
    End If

    Call SetCenter(Me)
    mPageHeight.���� = Me.Height
End Sub

Private Function PatiHaveCard(ByVal lng����ID As Long) As Boolean
'���ܣ��ж�ָ�������Ƿ��о��￨
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ���￨�� From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then
        PatiHaveCard = Not IsNull(rsTmp!���￨��)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub HideCard(Optional blnԤ�� As Boolean = True, Optional bln�ſ� As Boolean = True)
    If Not blnԤ�� Then
        mbln�Ƿ���ʾԤ�� = False
        picԤ��.Visible = False
        Me.Height = Me.Height - picԤ��.Height
    Else
        mbln�Ƿ���ʾԤ�� = True
    End If
    If Not bln�ſ� Then
        mbln�Ƿ���ʾ�ſ� = False
        pic�ſ�.Visible = False
        Me.Height = Me.Height - pic�ſ�.Height
    Else
        mbln�Ƿ���ʾ�ſ� = True
    End If
End Sub

Private Sub InitDicts()
    Dim i As Integer
    
    mstrȱʡ�ѱ� = zlDatabase.GetPara("ȱʡ�ѱ�", glngSys, mlngModul, , InStr(1, mstrPrivs, ";��������;") > 0)
    Call ReadDict("�Ա�", cbo�Ա�)
    Call ReadDict("�ѱ�", cbo�ѱ�)
    Call ReadDict("����", cbo����)
    Call ReadDict("����", cbo����)
    Call ReadDict("ѧ��", cboѧ��)
    Call ReadDict("����״��", cbo����״��)
    Call ReadDict("ְҵ", cboְҵ)
    Call ReadDict("���", cbo���)
    Call ReadDict("����ϵ", cbo��ϵ�˹�ϵ)
    
    Call ReadDict("����", cbo��Ժ����)
    Call ReadDict("��Ժ��ʽ", cbo��Ժ��ʽ)
    Call ReadDict("��Ժ����", cbo��Ժ����)  '���˺�:2007/09/13
    Call ReadDict("סԺĿ��", cboסԺĿ��)
     Call ReadDict("���֤δ¼ԭ��", cboIDNumber)
   
    Call ReadDict("ҽ�Ƹ��ʽ", cboҽ�Ƹ���, "ҽ�Ƹ��ʽ")
    
    Call ReadDict("��������", cbo��������, "��������")
    If mbytInState = 0 Then
        Call Load֧����ʽ
    End If
End Sub

Private Function ReadDict(strDict As String, cboInput As ComboBox, Optional strClass As String) As Boolean
'���ܣ���ʼ��ָ���ʵ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngMaxW As Long
    Dim strTemp As String
    Dim strȱʡ�ѱ� As String, blnFee As Boolean
    
    On Error GoTo errH
    strȱʡ�ѱ� = mstrȱʡ�ѱ�
    
    'by lesfeng 2010-01-12 �����Ż�
    If strDict = "���㷽ʽ" Then
        If strClass = "���￨" Then
            strTemp = "1,2"
        ElseIf strClass = "Ԥ����" Then
            If mbytMode = 1 Then
                strTemp = "1,2,8" 'ԤԼ�Ǽ�ʱ
            Else
                If InStr(mstrPrivs, "���ղ��˵Ǽ�") > 0 Then
                    strTemp = "1,2,3,5,8"
                Else
                    strTemp = "1,2,5,8"
                End If
            End If
        End If
'        strSQL = "Select Nvl(A.ȱʡ��־,0) as ȱʡ,B.����,B.����,Nvl(B.����,1) as ����" & _
'            " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
'            " Where A.���㷽ʽ=B.���� And A.Ӧ�ó���='" & strClass & "'" & _
'            " And Nvl(B.����,1) IN(" & strTemp & ") Order by B.����"
        strSQL = "Select Nvl(A.ȱʡ��־,0) as ȱʡ,B.����,B.����,Nvl(B.����,1) as ����" & _
            " From ���㷽ʽӦ�� A,���㷽ʽ B,Table(Cast(f_Num2List([1]) As zlTools.t_Numlist)) C " & _
            " Where A.���㷽ʽ=B.���� And A.Ӧ�ó���=[2]" & _
            " And (B.���� = C.Column_Value or B.���� is null) Order by B.����"
    ElseIf strDict = "���" Then
        strSQL = "Select ����,����,����,Nvl(���ȼ�,0) as ȱʡ From " & strDict & " Order by ����"
    ElseIf strDict = "��������" Then
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ,��ɫ From �������� Order by ����"
    ElseIf strDict = "ҽ�Ƹ��ʽ" Then
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ,�Ƿ�ҽ�� From ҽ�Ƹ��ʽ Order by ����"
    ElseIf strDict = "�ѱ�" Then
        '���ǽ��޳������Ψһ����Ŀ(������ȱʡ�ѱ�),������Ч�ڼ估����
        If mbytKind = E�������۵Ǽ� Then
            strTemp = "1,3" '�������۵Ǽ�
        Else
            strTemp = "2,3" 'סԺ��Ժ��סԺ���۵Ǽ�
        End If
        strSQL = "Select A.����,A.����,A.����,Nvl(A.ȱʡ��־,0) as ȱʡ From �ѱ� A,Table(Cast(f_Num2List([1]) As zlTools.t_Numlist)) B " & _
                 " Where (A.������� = B.Column_Value or A.������� is null) And A.����=1 And Nvl(A.���޳���,0)=0 And  " & _
                 " (a.��Ч��ʼ Is Null And a.��Ч���� Is Null Or Trunc(Sysdate) Between a.��Ч��ʼ And a.��Ч����) Order by A.����"
                 
'        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ�" & _
'            " Where ����=1 And Nvl(���޳���,0)=0 And Nvl(�������,3) IN(" & strTemp & ")" & _
'                " And  Sysdate Between NVL(��Ч��ʼ,Sysdate-1) and NVL(��Ч����,Sysdate+1)" & _
'            " Order by ����"
    Else
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTemp, strClass)
    cboInput.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If strClass = "ҽ�Ƹ��ʽ" Then
                cboInput.AddItem rsTmp!���� & "-" & rsTmp!���� & IIf(strClass = "ҽ�Ƹ��ʽ" And Val(Nvl(rsTmp!�Ƿ�ҽ��)) = 1, Chr(&HD), "")
            ElseIf strDict = "ְҵ" Then
                cboInput.AddItem rsTmp!���� & "-" & Chr(&HA) & rsTmp!����
            Else
                cboInput.AddItem rsTmp!���� & "-" & rsTmp!����
            End If
            
            If rsTmp!ȱʡ = 1 Then
                cboInput.ListIndex = cboInput.NewIndex
                cboInput.ItemData(cboInput.NewIndex) = 1
            End If
            If strDict = "�ѱ�" And strȱʡ�ѱ� = "" & rsTmp!���� Then
                strȱʡ�ѱ� = rsTmp!���� & "-" & rsTmp!����
                blnFee = True
            End If
            
            Select Case strClass
                Case "Ԥ����"
                    cboInput.ItemData(cboInput.NewIndex) = rsTmp!����
            End Select
            If TextWidth(cboInput.List(cboInput.NewIndex) & "��") > lngMaxW Then lngMaxW = TextWidth(cboInput.List(cboInput.NewIndex) & "��")
            rsTmp.MoveNext
        Next
        '69489
        If strDict = "�ѱ�" And blnFee = True Then
            For i = 0 To cboInput.ListCount - 1
                cboInput.ItemData(i) = 0
                If strȱʡ�ѱ� = cboInput.List(i) Then
                    cboInput.ListIndex = i
                End If
            Next i
            If cboInput.ListIndex > 0 Then cboInput.ItemData(cboInput.ListIndex) = 1
        End If
    ElseIf strDict = "���㷽ʽ" Then
        If strClass = "Ԥ����" Then
            MsgBox "û������Ԥ������㷽ʽ,������Ժʱ���ܽ�Ԥ���" & vbCrLf & _
                "Ҫʹ����ԺԤ��,���ȵ����㷽ʽ���������á�", vbInformation, gstrSysName
        Else
            MsgBox "û�����þ��￨���㷽ʽ,������Ժʱֻ�ܼ��ʷ�����" & vbCrLf & _
                "Ҫʹ�ý��㷢��,���ȵ����㷽ʽ���������á�", vbInformation, gstrSysName
            chk����.Value = 1: chk����.Enabled = False: cbo��������.Enabled = False
        End If
    End If
    ReadDict = True
    If cbo.ListWidth(cboInput.hWnd) < lngMaxW Then cbo.SetListWidth cboInput.hWnd, lngMaxW
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    mblnICCard = False
    
    '82401:���ϴ�,2015/3/11,�������Ƿ����
    If mbytInState = 0 And pic�ſ�.Visible Then
        zlDatabase.SetPara "����ģʽ", tabCardMode.SelectedItem.Key, glngSys, mlngModul
    End If
    
    Call zlCommFun.OpenIme
    mbytMode = 0
    mbytInState = 0
    mbytKind = 0
    mlng����ID = 0
    mlng��ҳID = 0
    mlngԤ������ID = 0
    Set mrsInfo = Nothing
    Set mrsDoctor = Nothing
    
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
    
    If gblnLED Then
        zl9LedVoice.DisplayPatient "": zl9LedVoice.Reset com
    End If
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    If Not mobjKeyboard Is Nothing Then
        Set mobjKeyboard = Nothing
    End If
    
    If Not mobjSquare Is Nothing Then Set mobjSquare = Nothing
    If Not mobjCommEvents Is Nothing Then Set mobjCommEvents = Nothing
    
    If Not mdicҽ�ƿ����� Is Nothing Then
        Set mdicҽ�ƿ����� = Nothing
    End If
    
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
'    Call zlCardSquareObject(True)
    Set gobjPatient = Nothing
    '�����:56599
    mbln������󶨿� = False
    Set mrsInputSet = Nothing
End Sub


Private Function InitData() As Boolean
'���ܣ���ʼ������ȼ�����Ժ���ҡ���Ժ����������ҽʦ����Ϣ
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim strDeptIDs As String
    
    If cbo�ѱ�.ListCount = 0 Then
        MsgBox "û�����÷ѱ���Ϣ,���ȵ��ѱ�ȼ����������ã�", vbExclamation, gstrSysName
        Exit Function
    ElseIf cbo�ѱ�.ListIndex = -1 Then
        cbo�ѱ�.ListIndex = 0
    End If
    
    '����ȼ�(ȱʡ��һ�����������)
    Set rsTmp = GetNurseGrade
    If rsTmp.RecordCount > 0 Then
        cbo����ȼ�.Clear
        cbo����ȼ�.AddItem ""   '��һ��Ϊ��,ReadPatiReg��������
        cbo����ȼ�.ItemData(cbo����ȼ�.NewIndex) = 0
        
        Call CboLoadData(cbo����ȼ�, rsTmp, False)
        If cbo����ȼ�.ListIndex = -1 Then cbo����ȼ�.ListIndex = 0
    Else
        MsgBox "û�����û���ȼ������ȵ�����ȼ������г�ʼ��", vbInformation, gstrSysName
        Exit Function
    End If
    
       
    '��ȡ����������ҽʦ�б�
    Set mrsDoctor = GetDoctorOrNurse(0)
    For i = 1 To mrsDoctor.RecordCount
        cbo����ҽʦ.AddItem mrsDoctor!���� & "-" & mrsDoctor!����
        cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = mrsDoctor!ID
        mrsDoctor.MoveNext
    Next
    
    '����۲��ҵĴ�λӦ��û�й̶�����,��������ʱ��������,ͬ���Կ��Ҷ���λ������
    '94400
    If mbytMode = EMode.EԤԼ�Ǽ� And InStr(mstrPrivs, ";ȫԺԤԼ;") = 0 Then
        strDeptIDs = GetDeptOrUnitByUser()
    End If
    Set mrsDept = GetDepartments("�ٴ�", IIf(mbytKind = EKind.E�������۵Ǽ�, "1", "2") & ",3", , True, strDeptIDs)
    If mrsDept.RecordCount = 0 Then
        MsgBox "û�����÷�����" & IIf(mbytKind = EKind.E�������۵Ǽ�, "����", "סԺ") & "�Ŀ��ҵĴ�λ��", vbInformation, gstrSysName
        Exit Function
    End If
    Set mrsUnit = GetDepartments("����", IIf(mbytKind = EKind.E�������۵Ǽ�, "1", "2") & ",3", , True, strDeptIDs)
    If mrsUnit.RecordCount = 0 Then
        MsgBox "û�����÷�����" & IIf(mbytKind = EKind.E�������۵Ǽ�, "����", "סԺ") & "�Ĳ����Ĵ�λ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ȡ�������Ҷ�Ӧ
    Set mrsUnitDept = GetUnitDept
    If mrsUnitDept.RecordCount = 0 Then
        MsgBox "û�����ò������Ҷ�Ӧ��ϵ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
        
    If gbln��ѡ���� Then
        Call CboLoadData(cbo��Ժ����, mrsUnit, True)
        If cbo��Ժ����.ListCount > 0 Then cbo��Ժ����.ListIndex = 0 '����Click�¼�,���ؿ��ҡ���λ����
    Else
        Call CboLoadData(cbo��Ժ����, mrsDept, True)
        If cbo��Ժ����.ListCount > 0 Then cbo��Ժ����.ListIndex = 0 '����Click�¼�,���ز�������λ����
    End If
    
    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
    mintIDKind = Val(mintIDKind)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
         
    
    InitData = True
End Function

Private Sub ClearCard(Optional blnKeepOther As Boolean, Optional blnKeepRec As Boolean)
'���ܣ������Ժ�Ǽǿ�
'������blnKeepOther=�Ƿ����ſ���Ԥ������Ϣ
'      blnKeepRec=�Ƿ����Ѷ�ȡ�Ĳ�����Ϣ����
    Dim lngUnit As Long, lngDept As Long
    Dim strȱʡ�ɿʽ As String
    
    If Not blnKeepRec Then
        Set mrsInfo = Nothing
        txtPatient.Text = "": txtPatient.Tag = ""
        txt����.Locked = False
        cbo�Ա�.Locked = False
        txt����.Locked = False
        cbo���䵥λ.Locked = False
    End If
    
    If gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    txtסԺ��.Locked = mbytKind = EסԺ���۵Ǽ�
    If mbytInState = EState.E���� And (mbytMode = EMode.E�����Ǽ� Or mbytMode = EMode.EԤԼ�Ǽ�) And mlng����ID <> 0 Then
        If mbytMode = EMode.E�����Ǽ� Then mbytKind = mbytKindBak
        txtPatient.Locked = False: txtPatient.TabStop = Not (InStr(mstrPrivs, "�����ҽ������") = 0)
        '66333:������,2013-10-10,�������صǼǺ�lblסԺ��.Caption = "�����"
        If mbytKind = E�������۵Ǽ� Then     '��������
'            txtסԺ��.Locked = True
'            lblסԺ��.Visible = False
'            txtסԺ��.Visible = False
            lblסԺ��.Caption = "�����"
            cmdSelectNO.Visible = False
            cmdYB.Visible = False
        ElseIf mbytKind = EסԺ���۵Ǽ� Then     'סԺ����(סԺ���ۺŲ����޸ģ�ÿ���µǼ�ʱ�������ۺŹ����Զ�����)
            txtסԺ��.TabStop = False
            txtסԺ��.Locked = True
            cmdSelectNO.Visible = False
            lblסԺ��.Caption = "���ۺ�"
        End If
        
        mlng����ID = 0: mlng��ҳID = 0
        Me.Caption = Me.Tag
    End If
    
    mblnEMPI = False
    txt����.Text = ""
    txtҽ����.Text = ""
    txtҽ����.Locked = False
    If mbytMode <> EMode.EԤԼ�Ǽ� And mbytKind = EKind.EסԺ��Ժ�Ǽ� Then
        txtTimes.Text = "1": txtTimes.Tag = 1
    Else
        txtTimes.Text = "": txtTimes.Tag = ""
    End If
    txtPages.Text = "1"
    
    txtסԺ��.Text = ""
    txt����.Text = ""
    txt����.Text = "": Call txt����_Validate(False): cbo���䵥λ.Tag = txt����.Text & "_" & cbo���䵥λ.Text
    txt��������.Text = "____-__-__"
    txt����ʱ��.Text = "__:__"
    txt���֤��.Text = ""
    txtMobile.Text = ""
    txt����֤��.Text = ""
    txt�����ص�.Text = ""
    txt��ͥ��ַ.Text = ""
    txt��ͥ��ַ�ʱ�.Text = ""
    txt��ͥ�绰.Text = ""
    txt���ڵ�ַ.Text = ""
    txt���ڵ�ַ�ʱ�.Text = ""
    txt����.Text = ""
    txt����.Text = ""
    txt��ϵ������.Text = ""
    txt��ϵ�˵�ַ.Text = ""
    txt��ϵ�˵绰.Text = ""
    txt��ϵ�����֤��.Text = ""
    txtLinkManInfo.Text = ""
    txt������λ.Text = "": txt������λ.Tag = ""
    txt������λ.Text = ""
    txt��λ�绰.Text = ""
    txt��λ�ʱ�.Text = ""
    txt��λ������.Text = ""
    txt��λ�ʺ�.Text = ""
    txt��ע.Text = ""
    '�����:53408
    txt֧������.Text = ""
    txt��֤����.Text = ""
    txt��֤����.Tag = ""
    txt֧������.Enabled = False
    txt��֤����.Enabled = False
    lbl֧������.Enabled = False
    lbl��֤����.Enabled = False
    
    
    txt�������.Text = "": txt�������.Tag = "": lbl�������.Tag = ""
    txt��ҽ���.Text = "": txt��ҽ���.Tag = "": lbl��ҽ���.Tag = ""
    
    chk����Ժת��.Value = 0
    chk���.Value = 0
    
    '73420:������,2014-06-09
    If InStr(mstrPrivs, "�޸�סԺ��") = 0 Then
        txtסԺ��.Locked = True
        txtסԺ��.TabStop = False
        txtסԺ��.BackColor = Me.BackColor
        cmdSelectNO.Visible = False
    End If
    
    cboIDNumber.ListIndex = -1 'ȱʡ
    cboIDNumber.Enabled = True
    cbo��ϵ�˹�ϵ.ListIndex = -1
    
    Call SetCboDefault(cbo�Ա�)
    Call SetCboDefault(cbo�ѱ�)
    Call SetCboDefault(cbo����)
    Call SetCboDefault(cbo����)
    Call SetCboDefault(cboѧ��)
    Call SetCboDefault(cbo����״��)
    Call SetCboDefault(cboְҵ)
    Call SetCboDefault(cbo���)
    Call SetCboDefault(cbo��Ժ����)
    Call SetCboDefault(cbo��Ժ��ʽ)
    Call SetCboDefault(cbo��Ժ����) '���˺�:2007/09/13
    Call SetCboDefault(cboסԺĿ��)
    Call SetCboDefault(cboҽ�Ƹ���)
    Call SetCboDefault(cbo��������)
    
    strȱʡ�ɿʽ = zlDatabase.GetPara("ȱʡ�ɿʽ", glngSys, mlngModul)
    'Ԥ������ͷ������� ȱʡֵ ���������Tag��,��ItemdataΪ ��������,�ʲ���SetCboDefault
    If strȱʡ�ɿʽ = "" Then
        If cboԤ������.ListCount > 0 Then cboԤ������.ListIndex = Val(cboԤ������.Tag)
    Else
        Call zlControl.CboLocate(cboԤ������, strȱʡ�ɿʽ, False)
    End If
    
    '����ȡ���ô�λ
    If cbo��Ժ����.ListIndex >= 0 Then lngUnit = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    If cbo��Ժ����.ListIndex >= 0 Then lngDept = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    Call LoadBed(zlCommFun.GetNeedName(cbo�Ա�.Text), lngDept, lngUnit)
    
    txt����.TabStop = True
    
    '��Ժ��Ϣ
    txt��Ժʱ��.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If InStr(mstrPrivs, "�޸���Ժ����") > 0 Then txt��Ժʱ��.Enabled = True
    
    
    If Not blnKeepOther Then
        '�ſ���Ϣ
        txt����.Text = ""
        txtPass.Text = ""
        txtAudi.Text = ""
        
        chk����.Value = IIf(gbln���� = True, 1, 0)
        If strȱʡ�ɿʽ = "" Then
            If cbo��������.ListCount > 0 Then cbo��������.ListIndex = Val(cbo��������.Tag)
        Else
            Call zlControl.CboLocate(cbo��������, strȱʡ�ɿʽ, False)
        End If
        
        'Ԥ����Ϣ
        txtԤ����.Text = ""
        txt�ɿλ.Text = ""
        txt�ʺ�.Text = ""
        txt������.Text = ""
        txt�������.Text = ""
    End If
    
    'ҽ���Ķ�
    txt����.ForeColor = lblPatiColor.BackColor
    mstrNOS = ""
    mintInsure = 0
    mstrYBPati = ""
    mcurYBMoney = 0
    mintInsureBak = 0
    mstrYBPatiBak = ""
    mcurYBMoneyBak = 0
    lblYBMoney.Caption = "�����ʻ����:"
    lblYBMoney.Visible = False
    chk����Ժ.Value = 0
    cmdTurn.Visible = InStr(1, mstrPrivs, ";�������תסԺ;") > 0 And mbytKind = EסԺ��Ժ�Ǽ� And mbytMode <> 1 '33635
    If InStr(mstrPrivs, "������Ϣ") > 0 And gbln���� Then Call Init������Ϣ(CDate(txt��Ժʱ��.Text))
    cmdName.Visible = mbytMode = 2
    txtTimes.Visible = mbytMode <> 1 And mbytKind = EסԺ��Ժ�Ǽ� 'ԤԼ�Ǽ�ʱ�����صǼ�ʱ,סԺ����Ϊ��
    lblTimes.Visible = mbytMode <> 1 And mbytKind = EסԺ��Ժ�Ǽ�
    txtTimes.Enabled = (InStr(1, mstrPrivs, "�޸�סԺ����") > 0 And mbytInState = 0)   '�޸�ʱ������ģ���Ϊ�����Ѳ���סԺһ�η��ã�Ԥ������￨
    
    '�����:56599
    Call Clear��������
    If gbln���ýṹ����ַ Then
        Call SetStrutAddress(1)
        Call SetStrutAddress(2)
    End If
End Sub

Private Sub cmd�����ص�_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt�����ص�, True)
    If Not rsTmp Is Nothing Then
        txt�����ص�.Text = rsTmp!����
        txt�����ص�.SelStart = Len(txt�����ص�.Text)
        txt�����ص�.SetFocus
    End If
End Sub

Private Sub cmd������λ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetOrgAddress(Me, txt������λ, True)
    If Not rsTmp Is Nothing Then
        txt������λ.Tag = rsTmp!ID
        txt������λ.Text = rsTmp!����
        txt������λ.SelStart = Len(txt������λ.Text)
        txt��λ�绰.Text = Trim(rsTmp!�绰 & "")
        txt��λ������.Text = Trim(rsTmp!�������� & "")
        txt��λ�ʺ�.Text = Trim(rsTmp!�ʺ� & "")
        
        txt������λ.SetFocus
    End If
End Sub

Private Sub cmd��ͥ��ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt��ͥ��ַ, True)
    If Not rsTmp Is Nothing Then
        txt��ͥ��ַ.Text = rsTmp!����
        txt��ͥ��ַ.SelStart = Len(txt��ͥ��ַ.Text)
        txt��ͥ��ַ.SetFocus
    End If
End Sub

Private Sub cmd��ϵ�˵�ַ_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt��ϵ�˵�ַ, True)
    If Not rsTmp Is Nothing Then
        txt��ϵ�˵�ַ.Text = rsTmp!����
        txt��ϵ�˵�ַ.SelStart = Len(txt��ϵ�˵�ַ.Text)
        txt��ϵ�˵�ַ.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim obj As Control
    
    Select Case KeyCode
        Case vbKeyF3
            If ActiveControl.Name = txt�����ص�.Name Then
                cmd�����ص�_Click
            ElseIf ActiveControl.Name = txt��ͥ��ַ.Name Then
                cmd��ͥ��ַ_Click
            ElseIf ActiveControl.Name = txt��ϵ�˵�ַ.Name Then
                cmd��ϵ�˵�ַ_Click
            ElseIf ActiveControl.Name = txt������λ.Name Then
                cmd������λ_Click
            ElseIf ActiveControl.Name = txt����.Name Then
                cmd����_Click
            End If
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC����")
                If intIndex < 0 Then Exit Sub
                IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyF8
            If cmdSelectNO.Enabled And cmdSelectNO.Visible Then cmdSelectNO_Click
'        Case vbKeyF10
'            If cmdSetup.Enabled And cmdSetup.Visible Then cmdSetup_Click
        Case vbKeyF11
            If mbytInState = 0 Then txtPatient.SetFocus
        Case vbKeyF12
            If cmdYB.Enabled And cmdYB.Visible Then cmdYB_Click
        Case vbKeyReturn
            Set obj = Me.ActiveControl
            If obj.Name = "cbo�Ա�" Then
                If cbo�Ա�.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo�ѱ�" Then
                If cbo�ѱ�.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo��������" Then
                If cbo��������.ListIndex <> -1 Then cmdOK.SetFocus
            ElseIf InStr(1, ",txt����,txt�����ص�,txt���ڵ�ַ,txt��ͥ��ַ,txt��ϵ�˵�ַ,txt������λ,txtԤ����,txtPatient,txt����," & _
                "txtסԺ��,txt����,txt����,txt�������,txt��ҽ���,txtPass,txtAudi,txt����,vsDrug,vsInoculate,vsLinkMan,vsOtherInfo,vsCertificate,PatiAddress,", "," & obj.Name & ",") <= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                 
        End If
    End Select
End Sub

Private Sub PatiAddress_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True) '���������뷨
End Sub

Private Sub PatiAddress_LostFocus(Index As Integer)
'����:
    Select Case Index
    
    Case E_IX_��סַ
        txt��ͥ��ַ.Text = PatiAddress(Index).Value
    Case E_IX_�����ص�
        txt�����ص�.Text = PatiAddress(Index).Value
    Case E_IX_���ڵ�ַ
        txt���ڵ�ַ.Text = PatiAddress(Index).Value
    Case E_IX_����
        txt����.Text = PatiAddress(Index).Value
    Case E_IX_��ϵ�˵�ַ
        txt��ϵ�˵�ַ.Text = PatiAddress(Index).Value
    End Select
    Call zlCommFun.OpenIme '�ر��������뷨
End Sub

Private Sub PatiAddress_SetInput(Index As Integer, ByVal intLevel As Integer, rsReturn As ADODB.Recordset)
    '���ܣ������벡�˽ṹ����ַ��ʱ��,�����ʱ�
    If (Not rsReturn Is Nothing) And intLevel = 2 Then
        If Index = 3 Then
            txt��ͥ��ַ�ʱ�.Text = rsReturn!�ʱ� & ""
        End If
        If Index = 4 Then
            txt���ڵ�ַ�ʱ�.Text = rsReturn!�ʱ� & ""
        End If
    End If
End Sub

Private Sub PatiAddress_Validate(Index As Integer, Cancel As Boolean)
    Dim lngLen As Long
    
    lngLen = PatiAddress(Index).MaxLength
    If LenB(StrConv(PatiAddress(Index).Value, vbFromUnicode)) > lngLen Then
        MsgBox PatiAddress(Index).Tag & "ֻ�������� " & lngLen & " ���ַ��� " & lngLen \ 2 & " �����֣�", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub tabCardMode_Click()
    If tabCardMode.SelectedItem.Key = "CardFee" Then
        lbl���.Visible = True
        txt����.Visible = True
        chk����.Visible = True
        cbo��������.Visible = True
    Else
        lbl���.Visible = False
        txt����.Visible = False
        chk����.Visible = False
        cbo��������.Visible = False
    End If
End Sub


Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    Call OpenPassKeyboard(txtAudi, True)
End Sub

Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mCurSendCard.int������� = 1 Then
            Call zlControl.TxtCheckKeyPress(txtAudi, KeyAscii, m����ʽ)
        End If
    End If

    If KeyAscii = vbKeyReturn Then
        If txtPass.Text <> txtAudi.Text Then
            MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
            zlControl.TxtSelAll txtAudi
            txtAudi.SetFocus
        Else
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub txtAudi_LostFocus()
    Call ClosePassKeyboard(txtAudi)

End Sub

Private Sub txtAudi_Validate(Cancel As Boolean)
    Select Case mCurSendCard.int���볤������
        Case 0
        Case 1
            If Len(txtAudi.Text) <> mCurSendCard.int���볤�� Then
                MsgBox "ע��:" & vbCrLf & "ȷ�������������" & mCurSendCard.int���볤�� & "λ", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtAudi.Text) < Abs(mCurSendCard.int���볤������) Then
                MsgBox "ע��:" & vbCrLf & "ȷ�����������" & Abs(mCurSendCard.int���볤������) & "λ����.", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Exit Sub
             End If
        End Select
End Sub

Private Sub txtLinkManInfo_GotFocus()
    zlControl.TxtSelAll txtLinkManInfo
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtLinkManInfo_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtLinkManInfo_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ�˹�ϵ��ע") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ�˹�ϵ��ע")) = txtLinkManInfo.Text
    End If
End Sub

Private Sub txtMedicalWarning_GotFocus()
    zlControl.TxtSelAll txtMedicalWarning
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOtherWaring_GotFocus()
    zlControl.TxtSelAll txtOtherWaring
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtOtherWaring_KeyPress(KeyAscii As Integer)
    If InStr("'|?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    CheckInputLen txtOtherWaring, KeyAscii
End Sub

Private Sub txtOtherWaring_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtPass_Validate(Cancel As Boolean)
   Select Case mCurSendCard.int���볤������
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCurSendCard.int���볤�� Then
                MsgBox "ע��:" & vbCrLf & "�����������" & mCurSendCard.int���볤�� & "λ", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCurSendCard.int���볤������) Then
                MsgBox "ע��:" & vbCrLf & "�����������" & Abs(mCurSendCard.int���볤������) & "λ����.", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        End Select
End Sub
Private Sub txtFact_GotFocus()
    zlControl.TxtSelAll txtFact
End Sub

Private Sub txtFact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtFact.Text) = txtFact.MaxLength And KeyAscii <> 8 And txtFact.SelLength <> Len(txtFact) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mCurSendCard.int������� = 1 Then
            Call zlControl.TxtCheckKeyPress(txtPass, KeyAscii, m����ʽ)
        End If
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            If chk����.Visible And chk����.Enabled And txt����.Locked Then
                chk����.SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtPass_LostFocus()
    Call ClosePassKeyboard(txtPass)
End Sub
Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub


Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.ActiveFastKey
End Sub

Private Sub txtReason_GotFocus()
    zlControl.TxtSelAll txtReason
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If InStr("'|?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    CheckInputLen txtReason, KeyAscii
End Sub

Private Sub txtReason_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtTimes_GotFocus()
    zlControl.TxtSelAll txtTimes
End Sub

Private Sub txtTimes_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If InStr(1, "0123456789", Chr(KeyAscii)) <= 0 And KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub
Private Sub txtTimes_Validate(Cancel As Boolean)
    txtTimes.Text = Val(txtTimes.Text)
    If Val(txtTimes.Text) < Val(txtTimes.Tag) Then
        txtTimes.Text = txtTimes.Tag
        Cancel = True
    End If
End Sub

Private Sub txt��ע_GotFocus()
    Call zlControl.TxtSelAll(txt��ע)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ע_KeyPress(KeyAscii As Integer)
    CheckInputLen txt��ע, KeyAscii
End Sub

Private Sub txt��ע_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '����27554 by lesfeng 2010-01-19 lngTXTProc �޸�ΪglngTXTProc
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt�����ص�_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��������_Change()
    Dim str�������� As String
    
    If IsDate(txt��������.Text) And mblnChange Then
        mblnChange = False
        txt��������.Text = Format(CDate(txt��������.Text), "yyyy-mm-dd") '0002-02-02�Զ�ת��Ϊ2002-02-02,����,��������2002,ʵ��ֵȴ��0002
        mblnChange = True
        If txt����ʱ��.Text = "__:__" Then
            str�������� = Format(txt��������.Text, "YYYY-MM-DD")
        Else
            str�������� = Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS")
        End If
        txt����.Text = ReCalcOld(CDate(str��������), cbo���䵥λ, , , CDate(txt��Ժʱ��.Text))
        cbo���䵥λ.Tag = txt����.Text & "_" & cbo���䵥λ.Text
    End If
End Sub

Private Sub txt��������_Validate(Cancel As Boolean)
    If IsDate(txt��������.Text) And IsDate(txt��Ժʱ��.Text) Then
        If CDate(txt��������.Text) > CDate(txt��Ժʱ��.Text) Then Call zlControl.TxtSelAll(txt��������): Cancel = True
    End If
End Sub

Private Sub txt����ʱ��_Change()
    Dim str�������� As String
    
    If IsDate(txt����ʱ��.Text) And IsDate(txt��������.Text) And mblnChange Then
        str�������� = Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS")
        txt����.Text = ReCalcOld(CDate(str��������), cbo���䵥λ, , , CDate(txt��Ժʱ��.Text))
        cbo���䵥λ.Tag = txt����.Text & "_" & cbo���䵥λ.Text
    End If
End Sub

Private Sub txt����ʱ��_GotFocus()
    Call OS.OpenImeByName
    zlControl.TxtSelAll txt����ʱ��
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt��������.Text) Then
        KeyAscii = 0
        txt����ʱ��.Text = "__:__"
    End If
End Sub


Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.SetFocus
        Cancel = True
    End If
End Sub


Private Sub txt��λ������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt������_LostFocus()
    If IsNumeric(txt������.Text) Then
        txt������.Text = Format(txt������.Text, "0.00")
    Else
        txt������.Text = ""
    End If
    
    Call zlCommFun.OpenIme
End Sub

Private Sub txt������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt������λ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt���ڵ�ַ_Change()
    If txt���ڵ�ַ.Text = "" Then txt���ڵ�ַ.Tag = ""
End Sub

Private Sub txt���ڵ�ַ_GotFocus()
    zlControl.TxtSelAll txt���ڵ�ַ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt���ڵ�ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt���ڵ�ַ.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt���ڵ�ַ)
            If Not rsTmp Is Nothing Then
                txt���ڵ�ַ.Text = rsTmp!����
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt���ڵ�ַ, KeyAscii
    End If
End Sub

Private Sub txt���ڵ�ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt���ڵ�ַ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt���ڵ�ַ�ʱ�
End Sub

Private Sub txt���ڵ�ַ�ʱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If ((Not IsNumeric(txt���ڵ�ַ�ʱ�.Text)) Or Len(txt���ڵ�ַ�ʱ�.Text) > 6 Or InStr(txt���ڵ�ַ�ʱ�.Text, ".") > 0) And txt���ڵ�ַ�ʱ�.Text <> "" Then
            Call SelectYouBian(txt���ڵ�ַ�ʱ�)
        End If
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text <> "" Then
            Set rsTmp = GetArea(Me, txt����)
            If Not rsTmp Is Nothing Then
                txt����.Text = rsTmp!����
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                zlControl.TxtSelAll txt����
                txt����.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt����, KeyAscii
    End If
End Sub

Private Sub txt����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��ͥ��ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt�ɿλ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt����_Change()
    SetCardEditEnabled
End Sub

Private Sub txt����_LostFocus()
    Call SetBrushCardObject(False)
End Sub

Private Sub txt������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��ϵ�˵�ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��ϵ�˵绰_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ�˵绰") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ�˵绰")) = txt��ϵ�˵绰.Text
    End If
End Sub

Private Sub txt��ϵ�����֤��_GotFocus()
    zlControl.TxtSelAll txt��ϵ�����֤��
End Sub

Private Sub txt��ϵ�����֤��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt��ϵ�����֤��_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ�����֤��") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ�����֤��")) = txt��ϵ�����֤��.Text
    End If
End Sub

Private Sub txt��ϵ������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��ϵ������_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("��ϵ������") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("��ϵ������")) = txt��ϵ������.Text
        If vsLinkMan.Rows = vsLinkMan.FixedRows + 1 And txt��ϵ������.Text <> "" Then
            vsLinkMan.Rows = vsLinkMan.Rows + 1
        End If
    End If
End Sub

Private Sub txt�������_LostFocus()
    If Not RequestCode Then
        Call zlCommFun.OpenIme
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo���䵥λ.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
        End If
        If cbo���䵥λ.Visible And Not IsNumeric(txt����.Text) And Me.ActiveControl.Name = "txt����" Then Call zlCommFun.PressKey(vbKeyTab)  'Ŀ���ǲ��������䵥λ

    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_LostFocus()
    If cbo���䵥λ.Tag <> txt����.Text & "_" & cbo���䵥λ.Text Then
        cbo���䵥λ_LostFocus
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        If Not InStr(Trim(txt����.Text), "Լ") > 0 And Trim(txt����.Text) <> "����" Then
            cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False
            txt��������.Enabled = True
            txt����ʱ��.Enabled = True
        ElseIf InStr(Trim(txt����.Text), "Լ") > 0 Or Trim(txt����.Text) = "����" Then
            If Trim(txt��������.Text) = "____-__-__" Then
                txt��������.Enabled = False
                txt����ʱ��.Enabled = False
            End If
            cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False
        End If
    ElseIf cbo���䵥λ.Visible = False Or txt��������.Enabled = False Then
        cbo���䵥λ.ListIndex = 0: cbo���䵥λ.Visible = True
        txt��������.Enabled = True
        txt����ʱ��.Enabled = True
    Else
        txt��������.Enabled = True
        txt����ʱ��.Enabled = True
    End If
    Call ReLoadCardFee
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text <> "" Then
            Set rsTmp = GetArea(Me, txt����)
            If Not rsTmp Is Nothing Then
                txt����.Text = rsTmp!����
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                zlControl.TxtSelAll txt����
                txt����.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt����, KeyAscii
    End If
End Sub

Private Sub txt����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��Ժʱ��_LostFocus()
    If Not IsDate(txt��Ժʱ��.Text) Then
        txt��Ժʱ��.SetFocus
    ElseIf dtp����ʱ��.Enabled Then
        If mbytInState = 0 And Not IsNull(dtp����ʱ��.Value) Then
            dtp����ʱ��.MinDate = CDate("1900-01-01")   '����ʱ����һ��С��ֵ,����ֵ�����
            
            If dtp����ʱ��.Value < CDate(txt��Ժʱ��.Text) Then
                dtp����ʱ��.Value = DateAdd("d", 3, CDate(txt��Ժʱ��.Text))
                MsgBox "��ǰ���õĵ�������ʱ��С����Ժʱ��,�ѵ���Ϊ��Ժʱ���3��!", vbInformation, gstrSysName
            End If
            
            '����ʱ�޲���С����Ժʱ��
            dtp����ʱ��.MinDate = CDate(txt��Ժʱ��.Text)
        ElseIf mbytInState = 1 Then
        
            If Not IsNull(dtp����ʱ��.Value) Then
                dtp����ʱ��.MinDate = CDate("1900-01-01")   '����ʱ����һ��С��ֵ,����ֵ�����
                '����ʱ�޲���С����Ժʱ��
                If dtp����ʱ��.Value < CDate(txt��Ժʱ��.Text) And txt������.Enabled Then
                    dtp����ʱ��.Value = DateAdd("d", 3, CDate(txt��Ժʱ��.Text))
                    MsgBox "��ǰ���õĵ�������ʱ��С����Ժʱ��,�ѵ���Ϊ��Ժʱ���3��!", vbInformation, gstrSysName
                End If
                dtp����ʱ��.MinDate = CDate(txt��Ժʱ��.Text)
            End If
        End If
    End If
End Sub

Private Sub txt�������_GotFocus()
    zlControl.TxtSelAll txt�������
    If Not RequestCode Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '����25785 by lesfeng 2009-10-20 ������������¼�����
            '************************************************
            If gint����������� = 1 Then
                strInput = UCase(txt�������.Text)
                strSex = zlCommFun.GetNeedName(cbo�Ա�.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "���� Like [2] or '('||����||')'||���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(gbytCode = 0, "����", "�����") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as ��ĿID,����,����,����," & IIf(gbytCode = 0, "����", "����� as ����") & ",˵��" & _
                        " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                        " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by ����"
                '����27613 by lesfeng 2010-01-21
                '����¼��ʱ�ж��ƥ��(����)������ѡ��,���ּ���ĸ�����ѡ��
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "D", strSex, gbytCode + 1)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                        End If
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fra��Ժ.hWnd, txt�������.Left, txt�������.Top)
                    strInput = UCase(txt�������.Text)
                    strSex = zlCommFun.GetNeedName(cbo�Ա�.Text)
                    lngTxtHeight = txt�������.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight, 1)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '���ݿ���ֻ��һ��ƥ����Ŀ�����Ը�ƥ�����ĿΪ׼
                    txt�������.Tag = rsTmp!ID
                    txt�������.Text = "(" & rsTmp!���� & ")" & rsTmp!���� '
                    lbl�������.Tag = txt�������.Text '���ڻָ���ʾ
                Else
                    '���������ƥ����Ŀʱ���������Ϊ׼
                    txt�������.Tag = ""
                    lbl�������.Tag = txt�������.Text '���ڻָ���ʾ
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt�������.Text = lbl�������.Tag And txt�������.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt�������.Text = "" Then
            txt�������.Tag = "": lbl�������.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            vPoint = zlControl.GetCoordPos(fra��Ժ.hWnd, txt�������.Left, txt�������.Top)
            strInput = UCase(txt�������.Text)
            strSex = zlCommFun.GetNeedName(cbo�Ա�.Text)
            lngTxtHeight = txt�������.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight, 1)
            
            If Not rsTmp Is Nothing Then
                txt�������.Tag = rsTmp!ID
                txt�������.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                lbl�������.Tag = txt�������.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If lbl�������.Tag <> "" Then txt�������.Text = lbl�������.Tag
                Call txt�������_GotFocus
                txt�������.SetFocus
            End If
        End If
    Else
        CheckInputLen txt�������, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt�������_Validate(Cancel As Boolean)
    If Val(txt�������.Tag) > 0 And txt�������.Text <> lbl�������.Tag Then
        txt�������.Text = lbl�������.Tag
    ElseIf Val(txt�������.Tag) = 0 And RequestCode Then
        txt�������.Text = ""
    End If
End Sub

Private Sub txt���֤��_LostFocus()
    '�����:53408
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
    '����81342
    If Trim(txt���֤��.Text) = "" And cboIDNumber.Visible Then
        cboIDNumber.Enabled = True
        cboIDNumber.SetFocus
    Else
        cboIDNumber.Enabled = False
        cboIDNumber.ListIndex = -1
    End If
    Call ReLoadCardFee
End Sub

Private Sub txt��֤����_GotFocus()
    Call zlControl.TxtSelAll(txt��֤����)
    Call OpenPassKeyboard(txt��֤����, False)
End Sub

Private Sub txt��֤����_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCurSendCard.int������� = 1)
End Sub

Private Sub txt��֤����_LostFocus()
    Call ClosePassKeyboard(txt��֤����)
End Sub

Private Sub txt֧������_GotFocus()
    Call zlControl.TxtSelAll(txt֧������)
    Call OpenPassKeyboard(txt֧������, False)
End Sub

Private Sub txt֧������_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCurSendCard.int������� = 1)
End Sub

Private Sub txt֧������_LostFocus()
    Call ClosePassKeyboard(txt֧������)
End Sub

Private Sub txt��ҽ���_GotFocus()
    zlControl.TxtSelAll txt��ҽ���
    If Not RequestCode Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub
Private Sub CheckInputPassWord(KeyAscii As Integer, Optional ByVal blnOnlyNum As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2011-07-07 00:40:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 8 Or KeyAscii = 13 Then Exit Sub
    If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If blnOnlyNum Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
       If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
            If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                 If InStr(1, "!@#$%^&*()_+-=><?,:;~`./", Asc(KeyAscii)) = 0 Then KeyAscii = 0
            End If
       End If
    End If
End Sub
Private Sub txt��ҽ���_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '����25785 by lesfeng 2009-10-20 ������������¼�����
            '************************************************
            If gint����������� = 1 Then
                strInput = UCase(txt��ҽ���.Text)
                strSex = zlCommFun.GetNeedName(cbo�Ա�.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "���� Like [2] or '('||����||')'||���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(gbytCode = 0, "����", "�����") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as ��ĿID,����,����,����," & IIf(gbytCode = 0, "����", "����� as ����") & ",˵��" & _
                        " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                        " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by ����"
                
                '����27613 by lesfeng 2010-01-21
                '����¼��ʱ�ж��ƥ��(����)������ѡ��,���ּ���ĸ�����ѡ��
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "B", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fra��Ժ.hWnd, txt��ҽ���.Left, txt��ҽ���.Top)
                    strInput = UCase(txt��ҽ���.Text)
                    strSex = zlCommFun.GetNeedName(cbo�Ա�.Text)
                    lngTxtHeight = txt��ҽ���.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight, 1)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '���ݿ���ֻ��һ��ƥ����Ŀ�����Ը�ƥ�����ĿΪ׼
                    txt��ҽ���.Tag = rsTmp!ID
                    txt��ҽ���.Text = "(" & rsTmp!���� & ")" & rsTmp!���� '
                    lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                Else
                    '���������ƥ����Ŀʱ���������Ϊ׼
                    txt��ҽ���.Tag = ""
                    lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��ҽ���.Text = lbl��ҽ���.Tag And txt��ҽ���.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��ҽ���.Text = "" Then
            txt��ҽ���.Tag = "": lbl��ҽ���.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            vPoint = zlControl.GetCoordPos(fra��Ժ.hWnd, txt��ҽ���.Left, txt��ҽ���.Top)
            strInput = UCase(txt��ҽ���.Text)
            strSex = zlCommFun.GetNeedName(cbo�Ա�.Text)
            lngTxtHeight = txt��ҽ���.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight, 1)
            If Not rsTmp Is Nothing Then
                txt��ҽ���.Tag = rsTmp!ID
                txt��ҽ���.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ�����ҽ�������롣", vbInformation, gstrSysName
                End If
                If lbl��ҽ���.Tag <> "" Then txt��ҽ���.Text = lbl��ҽ���.Tag
                Call txt��ҽ���_GotFocus
                txt��ҽ���.SetFocus
            End If
        End If
    Else
        CheckInputLen txt��ҽ���, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt��ҽ���_LostFocus()
    If Not RequestCode Then
        Call zlCommFun.OpenIme
    End If
End Sub

Private Sub txt��ҽ���_Validate(Cancel As Boolean)
    If Val(txt��ҽ���.Tag) > 0 And txt��ҽ���.Text <> lbl��ҽ���.Tag Then
        txt��ҽ���.Text = lbl��ҽ���.Tag
    ElseIf Val(txt��ҽ���.Tag) = 0 And RequestCode Then
        txt��ҽ���.Text = ""
    End If
End Sub

Private Sub txt���֤��_Change()
    Dim strBirthDay  As String
    Dim strAge As String
    Dim strSex As String
    Dim strErrInfo As String
    
    If mblnChange Then
        If CreatePublicPatient() Then
            If gobjPublicPatient.CheckPatiIdcard(Trim(txt���֤��.Text), strBirthDay, strAge, strSex, strErrInfo) Then
                If IsDate(strBirthDay) Then
                    txt��������.Enabled = True
                    txt����ʱ��.Enabled = True
                End If
                If txt��������.Enabled = True Then txt��������.Text = strBirthDay
                If cbo�Ա�.Enabled Then Call cbo.Locate(cbo�Ա�, strSex, False)
            End If
        End If
    End If
End Sub

Private Sub txt����_LostFocus()
    Call zlCommFun.OpenIme
    txt����.Text = Trim(txt����.Text)
End Sub

Private Sub txtҽ����_GotFocus()
    Call zlControl.TxtSelAll(txtҽ����)
End Sub

Private Sub txtҽ����_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    '�������ַ�
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ~!@#$%^&*()_+|-=\[]{}<>,./" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    
    'ҽ�Ƹ��ʽȱʡ=������ҽ�Ʊ���
    If txtҽ����.Text <> "" Then
        For i = 0 To cboҽ�Ƹ���.ListCount
            If InStr(cboҽ�Ƹ���.List(i), Chr(&HD)) > 0 Then cboҽ�Ƹ���.ListIndex = i: Exit For
        Next
    End If
End Sub

Private Sub txtԤ����_LostFocus()
    '65965:������,2013-09-24,����Ԥ����ʾǧλλ��ʽ
    If IsNumeric(txtԤ����.Text) Then
        txtԤ����.Text = Format(StrToNum(txtԤ����.Text), "##,##0.00;-##,##0.00; ;")
    Else
        txtԤ����.Text = ""
    End If
    If txtԤ����.MaxLength > 12 Then txtԤ����.MaxLength = 12
    
    If gblnLED Then
        '#22 1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
        '#23 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
        zl9LedVoice.Speak "#22 " & StrToNum(txtԤ����.Text)
    End If
End Sub

Private Sub txt�ɿλ_GotFocus()
    If IsNumeric(txtԤ����.Text) And txt�ɿλ.Text = "" Then
        txt�ɿλ.Text = txt������λ.Text
    End If
    zlControl.TxtSelAll txt�ɿλ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt�ɿλ_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�ɿλ, KeyAscii
End Sub

Private Sub txt��λ������_KeyPress(KeyAscii As Integer)
    CheckInputLen txt��λ������, KeyAscii
End Sub

Private Sub txt��λ�ʺ�_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt��λ�ʺ�, KeyAscii
End Sub

Private Sub txt������_GotFocus()
    zlControl.TxtSelAll txt������
End Sub

Private Sub txt������_GotFocus()
    zlControl.TxtSelAll txt������
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt������, KeyAscii
End Sub

Private Sub txt������λ_Change()
    If txt������λ.Text = "" Then txt������λ.Tag = ""
End Sub

Private Sub txt�������_GotFocus()
    zlControl.TxtSelAll txt�������
End Sub

Private Sub txtԤ����_GotFocus()
    If IsNumeric(txtԤ����.Text) Then
        txtԤ����.Text = StrToNum(txtԤ����.Text)
    Else
        txtԤ����.Text = ""
    End If
    txtԤ����.SelStart = 0: txtԤ����.SelLength = Len(txtԤ����.Text)
End Sub

Private Sub txtԤ����_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii <> 13 Then
        If InStr(txtԤ����.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        '65965:������,2013-09-24,����Ԥ����ʾǧλλ��ʽ
        If (txtԤ����.Text <> "" And txtԤ����.SelLength <> Len(Format(StrToNum(txtԤ����.Text), "##,##0.00;-##,##0.00; ;"))) And _
            (Len(Format(StrToNum(txtԤ����.Text), "##,##0.00;-##,##0.00; ;")) >= txtԤ����.MaxLength) And _
            InStr(Chr(8), Chr(KeyAscii)) = 0 Then
            If txtԤ����.SelLength > 0 And txtԤ����.SelLength <= txtԤ����.MaxLength Then
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf IsNumeric(txtԤ����.Text) Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        '����ȡԤ����,ֱ������
        txtԤ����.Text = ""
        If pic�ſ�.Visible Then
            If Not mrsInfo Is Nothing Then
                If Not IsNull(mrsInfo!���￨��) Then
                    cmdOK.SetFocus
                Else
                    txt����.SetFocus
                End If
            Else
                txt����.SetFocus
            End If
        Else
            cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If gblnSeekName Then Call zlCommFun.OpenIme(True)
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjIDCard.SetEnabled (True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjICCard.SetEnabled (True)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String, curFee As Currency, cur����δ�� As Currency
    Dim i As Integer, blnICCard As Boolean
    Dim blnCard As Boolean
    Dim strסԺ�� As String
    
    If txtPatient.Locked Then Exit Sub
    '�����ַ�������Form_KeyPress�н���
    
    'ֱ�����벡����Ϣ,���²��˱���,�����ԭ�ȵĲ�����Ϣ
    If KeyAscii = 13 Then
        If Trim(txtPatient.Text) = "" Then
            Call ClearCard(True) 'ֻ���������Ϣ
            '�����µĲ���ID��סԺ��
            txtPatient.Text = zlDatabase.GetNextNo(1)
            txtPatient.Tag = txtPatient.Text
            
            '���۲��˲��Զ�����סԺ��
            If mbytKind = EסԺ��Ժ�Ǽ� Then
                txtסԺ��.Text = zlDatabase.GetNextNo(2)
                If Not txtסԺ��.Locked Then
                    txtסԺ��.SetFocus
                Else
                    txt����.SetFocus
                End If
            ElseIf mbytKind = E�������۵Ǽ� Then
                'txtסԺ��.Locked = True
                txtסԺ��.Text = zlDatabase.GetNextNo(3)
                mblnAuto = True
                If Not txtסԺ��.Locked Then
                    txtסԺ��.SetFocus
                Else
                    txt����.SetFocus
                End If
'                txt����.SetFocus
            ElseIf mbytKind = EסԺ���۵Ǽ� Then 'סԺ���ۺŲ����޸ģ�ÿ���µǼ�ʱ�������ۺŹ����Զ�����
                txtסԺ��.Text = zlDatabase.GetNextNo(6)
                txt����.SetFocus
            Else
                txt����.SetFocus
            End If
            Exit Sub
        ElseIf txtPatient.Text = txtPatient.Tag Then
            If Not txtסԺ��.Locked Then
                txtסԺ��.SetFocus
            Else
                txt����.SetFocus
            End If
            Exit Sub
        End If
    End If

    If IDKind.GetCurCard.���� Like "����*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        If blnCard And IDKind.ShowPassText Then txtPatient.PasswordChar = "*"
    ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    End If
    '55571:������,2012-11-12
    txtPatient.IMEMode = 0
    
    On Error GoTo errHandle
    
    'ˢ����ϻ���������س�
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txtPatient.Text <> "" Then
        
        '37662
        If Not InStr(gstrPrivs, "�޸Ĳ�����Ϣ") > 0 Then
            txt����.Locked = True
            cbo�Ա�.Locked = True
            txt����.Locked = True
            cbo���䵥λ.Locked = True
        End If
    
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        If IDKind.GetCurCard.���� Like "IC��*" And IDKind.GetCurCard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        mblnICCard = blnICCard
        
        
        '��ȡ������Ϣ
        If GetPatient(IDKind.GetCurCard, txtPatient.Text, blnCard) Then
            Led��ӭ��Ϣ
            
            If Not isValid(mrsInfo!����ID) Then txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
            '���￨������:3��
            If gblnCheckPass And (blnCard Or blnICCard) Then
                If zlCommFun.VerifyPassWord(Me, mstrPassWord) = False Then
                    txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
                End If
            End If
            
            '�Ѿ���ʽ�Ǽ�(ԤԼ�����ڲ�����Ϣ��û���ǰ����)
            If Not IsNull(mrsInfo!��ǰ����id) Then
                MsgBox """" & mrsInfo!���� & """�Ѿ��Ǽ�Ϊ" & Decode(mrsInfo!��������, 0, "��Ժ", 1, "��������", 2, "סԺ����") & "���ˡ�", vbInformation, gstrSysName
                txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
            End If
            
            'סԺ�Ǽǽ�����ղ���
            If Not IsNull(mrsInfo!��ҳID) And mbytInState = EState.E���� And mbytMode <> EMode.E����ԤԼ Then 'û��ס��Ժ�Ĳ��˵���ҳidΪ��(��Ϊ�����������Ӳ�ѯ)
                If mrsInfo!��ҳID = 0 Then '�Ѿ�ԤԼ�Ĳ���(û���ṩ����ԤԼ)
                    If mbytMode = EMode.EԤԼ�Ǽ� Or mbytMode = EMode.E�����Ǽ� And mbytKind <> EKind.EסԺ��Ժ�Ǽ� Then
                        MsgBox """" & mrsInfo!���� & """�Ѿ�ԤԼ�Ǽǡ�", vbInformation, gstrSysName
                        txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
                    Else
                        strTmp = ""
                        If InStr(mstrPrivs, "����ԤԼ") = 0 Then MsgBox "��û�С�����ԤԼ����Ȩ�ޣ� ���ܽ���ԤԼ���ˣ�", vbInformation, gstrSysName: Exit Sub
                        If InStr(mstrPrivs, "����סԺԤԼ") = 0 And mrsInfo!�������� = 0 Then
                            MsgBox "��û�С�����סԺԤԼ����Ȩ�ޣ� ���ܽ���סԺԤԼ���ˣ�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        If InStr(mstrPrivs, "������������ԤԼ") = 0 And mrsInfo!�������� = 1 Then
                            MsgBox "��û�С�������������ԤԼ����Ȩ�ޣ� ���ܽ�����������ԤԼ���ˣ�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        If InStr(mstrPrivs, "����סԺ����ԤԼ") = 0 And mrsInfo!�������� = 2 Then
                            MsgBox "��û�С�����סԺ����ԤԼ����Ȩ�ޣ� ���ܽ���סԺ����ԤԼ���ˣ�", vbInformation, gstrSysName
                            Exit Sub
                        End If

                        If InStr(mstrPrivs, "�������۵Ǽ�") = 0 And InStr(mstrPrivs, "סԺ���۵Ǽ�") = 0 Then
                            If InStr(mstrPrivs, "����סԺԤԼ") = 0 Then
                                MsgBox "��û���㹻���û�Ȩ�ޣ� ���ܽ���ԤԼ���ˣ�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If MsgBox("Ҫ��""" & mrsInfo!���� & """����ΪסԺ������?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then strTmp = "סԺ����"
                        ElseIf InStr(mstrPrivs, "סԺ���۵Ǽ�") = 0 Then
                            If InStr(mstrPrivs, "������������ԤԼ") <> 0 And InStr(mstrPrivs, "����סԺԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0),��������(&1)"
                            ElseIf InStr(mstrPrivs, "������������ԤԼ") <> 0 Then
                                strTmp = "!��������(&0)"
                            ElseIf InStr(mstrPrivs, "����סԺԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0)"
                            Else
                                MsgBox "��û���㹻���û�Ȩ�ޣ� ���ܽ���ԤԼ���ˣ�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            strTmp = zlCommFun.ShowMsgBox("ԤԼ����", "Ҫ��""" & mrsInfo!���� & """����Ϊ", strTmp, Me, vbQuestion)
                        ElseIf InStr(mstrPrivs, "�������۵Ǽ�") = 0 Then
                            If InStr(mstrPrivs, "����סԺ����ԤԼ") <> 0 And InStr(mstrPrivs, "����סԺԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0),סԺ����(&1)"
                            ElseIf InStr(mstrPrivs, "����סԺ����ԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0)"
                            ElseIf InStr(mstrPrivs, "����סԺԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0)"
                            Else
                                MsgBox "��û���㹻���û�Ȩ�ޣ� ���ܽ���ԤԼ���ˣ�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            strTmp = zlCommFun.ShowMsgBox("ԤԼ����", "Ҫ��""" & mrsInfo!���� & """����Ϊ", strTmp, Me, vbQuestion)
                        Else
                            If InStr(mstrPrivs, "����סԺ����ԤԼ") <> 0 And InStr(mstrPrivs, "����סԺԤԼ") <> 0 And InStr(mstrPrivs, "������������ԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0),��������(&1),סԺ����(&2)"
                            ElseIf InStr(mstrPrivs, "����סԺ����ԤԼ") <> 0 And InStr(mstrPrivs, "����סԺԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0),סԺ����(&1)"
                            ElseIf InStr(mstrPrivs, "����סԺ����ԤԼ") <> 0 And InStr(mstrPrivs, "������������ԤԼ") <> 0 Then
                                strTmp = "!��������(&0),סԺ����(&1)"
                            ElseIf InStr(mstrPrivs, "����סԺԤԼ") <> 0 And InStr(mstrPrivs, "������������ԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0),��������(&1)"
                            ElseIf InStr(mstrPrivs, "������������ԤԼ") <> 0 Then
                                strTmp = "!��������(&0)"
                            ElseIf InStr(mstrPrivs, "����סԺԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0)"
                            ElseIf InStr(mstrPrivs, "����סԺ����ԤԼ") <> 0 Then
                                strTmp = "!סԺ����(&0)"
                            Else
                                MsgBox "��û���㹻���û�Ȩ�ޣ� ���ܽ���ԤԼ���ˣ�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            strTmp = zlCommFun.ShowMsgBox("ԤԼ����", "Ҫ��""" & mrsInfo!���� & """����Ϊ", strTmp, Me, vbQuestion)
                        End If
                        
                        If strTmp = "" Then txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
                        
                        mbytKind = Switch(strTmp = "סԺ����", 0, strTmp = "��������", E�������۵Ǽ�, strTmp = "סԺ����", EסԺ���۵Ǽ�)
                        
                        cmdName.Visible = True
                        cmdTurn.Visible = InStr(1, mstrPrivs, ";�������תסԺ;") > 0 And mbytKind = EסԺ��Ժ�Ǽ� And mbytMode <> 1
                        
                        txtTimes.Visible = mbytMode <> 1 And mbytKind = EסԺ��Ժ�Ǽ� 'ԤԼ�Ǽ�ʱ�����صǼ�ʱ,סԺ����Ϊ��
                        lblTimes.Visible = mbytMode <> 1 And mbytKind = EסԺ��Ժ�Ǽ�
                        txtTimes.Enabled = (InStr(1, mstrPrivs, "�޸�סԺ����") > 0 And mbytInState = 0)   '�޸�ʱ������ģ���Ϊ�����Ѳ���סԺһ�η��ã�Ԥ������￨
                            
                        If Not InitData Then Unload Me: Exit Sub
                        Me.Caption = "����" & strTmp
                        mlng����ID = mrsInfo!����ID: mlng��ҳID = 0
                        Call zlCommFun.PressKey(vbKeyTab)
                        txtPatient.Locked = True: txtPatient.TabStop = False
                        
                        If mbytKind = E�������۵Ǽ� Then     '��������
'                            lblסԺ��.Visible = False
'                            txtסԺ��.Visible = False
                            lblסԺ��.Caption = "�����"
                            cmdSelectNO.Visible = False
                            cmdYB.Visible = False
                        ElseIf mbytKind = EסԺ���۵Ǽ� Then     'סԺ����
                            lblסԺ��.Caption = "���ۺ�"
                            txtסԺ��.Locked = True
                            cmdSelectNO.Visible = False
                        End If
                                                
                        If Not ReadPatiReg(mrsInfo!����ID, 0) Then
                            MsgBox "������ȷ��ȡԤԼ����""" & mrsInfo!���� & """�ĵǼǼ�¼��", vbInformation, gstrSysName
                            Call ClearCard
                            Exit Sub
                        End If
                        
                         '���֮ǰû��סԺ�Ż�ÿ��סԺ������סԺ��,����ΪסԺ���ˣ����Զ������µ�סԺ��
                        '���� 27063 by lesfeng 2009-12-25 ԤԼ�Ǽ�תסԺ���˱���ԭסԺ��(ȡ��gblnÿ��סԺ��סԺ���ж�)
        '                If mbytKind = EKind.EסԺ��Ժ�Ǽ� And (Trim(txtסԺ��.Text) = "" Or gblnÿ��סԺ��סԺ��) Then txtסԺ��.Text = zlDatabase.GetNextNo(2)
                        '85510:LPF,2015-06-19,ԤԼ�Ǽ�סԺ�Ų�������ҽ���Ǽ���Ժ����,��ҽ���Ǽǲ���סԺ��ʱ��������дסԺҵ�����:
                        'ԭ���߼��ж�:If mbytKind = EKind.EסԺ��Ժ�Ǽ� And (Trim(txtסԺ��.Text) = "") Then txtסԺ��.Text = zlDatabase.GetNextNo(2)
                        '��Ժ����ԤԼ�Ǽǻ���ݲ���"ÿ��סԺ��סԺ��"����סԺ��,��ҽ���Ǽ�Ŀǰֻ���Բ�����Ϣ��סԺ��Ϊ׼����(���ַ�ʽ������סԺ�ſ��ܾͲ���ȷ)
                        '�����Ҫ�����´���
                        '1:gblnÿ��סԺ��סԺ��=TRUE,�������סԺ�ţ��������е�סԺ���Ƿ��ظ�������ظ����������ɡ�
                        '2:gblnÿ��סԺ��סԺ��=FALSE,���סԺ��Ϊ��,��ʹ����ʷסԺ��(���һ��סԺ�Ų�Ϊ��)����������ʷסԺ���������ɡ�
                        If mbytKind = EKind.EסԺ��Ժ�Ǽ� Then
                            If gblnÿ��סԺ��סԺ�� = True Then
                                If Trim(txtסԺ��.Text) <> "" Then
                                    If CheckByPatiNO(mrsInfo!����ID, 0, 0, Trim(txtסԺ��.Text)) = True Then txtסԺ��.Text = ""
                                End If
                            Else
                                If Trim(txtסԺ��.Text) = "" Then
                                    strסԺ�� = ""
                                    If CheckByPatiNO(mrsInfo!����ID, 0, 1, strסԺ��) = True Then txtסԺ��.Text = strסԺ��
                                End If
                            End If
                            If Trim(txtסԺ��.Text) = "" Then txtסԺ��.Text = zlDatabase.GetNextNo(2)
                        ElseIf mbytKind = EסԺ���۵Ǽ� Then
                            If Trim(txtסԺ��.Text) = "" Then txtסԺ��.Text = zlDatabase.GetNextNo(6)
                        End If
                
                        Exit Sub
                    End If
                End If
            End If
            
            
            '����������
            strTmp = inBlackList(mrsInfo!����ID)
            If strTmp <> "" Then
                If MsgBox("����""" & mrsInfo!���� & """�����ⲡ�������С�" & vbCrLf & vbCrLf & "ԭ��" & vbCrLf & vbCrLf & "����" & strTmp & vbCrLf & vbCrLf & "Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Call ClearCard(True): txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.SetFocus: Exit Sub
                End If
            End If
            
            '���˷����������
            curFee = GetPatientUnBalance(mrsInfo!����ID, cur����δ��)
            If cur����δ�� <> 0 Or curFee <> 0 Then
                strTmp = ""
                If cur����δ�� <> 0 Then strTmp = "�������" & Format(cur����δ��, "0.00")
                If curFee <> 0 Then strTmp = strTmp & IIf(strTmp = "", "", ",") & "סԺ����" & Format(curFee, "0.00")
                                
                strTmp = "���ѣ�""" & mrsInfo!���� & """��δ����" & strTmp
                If mbytMode = EMode.E����ԤԼ Then
                    MsgBox strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName
                Else
                    If MsgBox(strTmp & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call ClearCard(True): txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.SetFocus: Exit Sub
                    End If
                End If
            End If
            
            
            '����Ƿ���Ӧ�տ�
            strTmp = "Select Zl_Patientdue([1]) ʣ��Ӧ�� From dual"
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "��ȡӦ�տ�", CLng(mrsInfo!����ID))
            If Not rsTmp.EOF Then
                If Nvl(rsTmp!ʣ��Ӧ��, 0) > 0 Then
                    If mbytMode = EMode.E����ԤԼ Then
                        MsgBox "�ò������� " & rsTmp!ʣ��Ӧ�� & "Ԫ Ӧ�տ�δ�ɣ�", vbInformation, gstrSysName
                    Else
                        If MsgBox("�ò������� " & rsTmp!ʣ��Ӧ�� & "Ԫ Ӧ�տ�δ�ɣ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Call ClearCard(True): txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.SetFocus: Exit Sub
                        End If
                    End If
                End If
            End If
            
            '---------------------------------------------------------------------------------------
            'If mstrYBPati <> "" Then txt����.ForeColor = vbRed
            
            '������ϼ�¼
            If mstrYBPati = "" Then
                Call ClearCard(True, True)
            ElseIf RequestCode Then
                If Val(txt�������.Tag) = 0 Then
                    txt�������.Text = "": txt�������.Tag = "": lbl�������.Tag = ""
                End If
                If Val(txt��ҽ���.Tag) = 0 Then
                    txt��ҽ���.Text = "": txt��ҽ���.Tag = "": lbl��ҽ���.Tag = ""
                End If
            End If
            
            Set rsTmp = GetDiagnosticInfo(mrsInfo!����ID, 0, "1,11", "3")
            If Not rsTmp Is Nothing Then
                rsTmp.Filter = "�������=1"
                If Not rsTmp.EOF Then
                    txt�������.Text = Nvl(rsTmp!�������): txt�������.Tag = Nvl(rsTmp!����ID, rsTmp!���ID & ";"): lbl�������.Tag = txt�������.Text
                End If
                
                rsTmp.Filter = "�������=11"
                If Not rsTmp.EOF Then
                    txt��ҽ���.Text = Nvl(rsTmp!�������): txt��ҽ���.Tag = Nvl(rsTmp!����ID, rsTmp!���ID & ";"): lbl��ҽ���.Tag = txt�������.Text
                End If
            End If
            txt��Ժʱ��.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
            If Not IsNull(mrsInfo!���￨��) Then txt����.TabStop = False
            '��д������Ϣ
            If Not FuncPlugPovertyInfo(Val(mrsInfo!����ID)) Then Exit Sub
            Call FillPatient
            'EMPI
            Call EMPI_LoadPati(1)
            '���¿���
            Call ReLoadCardFee(True)
            cbo��������.Enabled = InStr(mstrPrivs, "������������") > 0
            If mbytInState = 0 And cbo��Ժ����.ListIndex >= 0 Then
                chk����Ժ.Value = IIf(CheckReIN(mrsInfo!����ID, Val(cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex))), 1, 0)
            End If
            If CanFocus(cbo��Ժ����) Then cbo��Ժ����.SetFocus
        ElseIf (blnCard Or blnICCard) And pic�ſ�.Visible Then  '���¿�
            MsgBox "�ÿ�û�н���,����Ϊ�¿��Ǽ�,�����벡��������", vbInformation, gstrSysName
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
            txt����.Text = txtPatient.Text
            txtPatient.Text = zlDatabase.GetNextNo(1)
            txtPatient.Tag = txtPatient.Text
            txt��Ժʱ��.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            If mbytKind = EסԺ��Ժ�Ǽ� Then
                txtסԺ��.Text = zlDatabase.GetNextNo(2)
            ElseIf mbytKind = EסԺ���۵Ǽ� Then
                txtסԺ��.Text = zlDatabase.GetNextNo(6)
            End If
            
            Call CheckFreeCard(txt����.Text)
            txt����.Locked = False
            cbo�Ա�.Locked = False
            txt����.Locked = False
            cbo���䵥λ.Locked = False
            txt����.SetFocus
        ElseIf Not IDKind.GetCurCard.���� = "���֤��" Then
            MsgBox "û���ҵ�ָ���Ĳ��ˡ�", vbInformation, gstrSysName
            Call zlControl.TxtSelAll(txtPatient)
            txtPatient.SetFocus
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPatientUnBalance(ByVal lng����ID As Long, ByRef cur����δ�� As Currency) As Currency
'���ܣ���ȡָ������δ�����,�������δ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��Դ;��, Sum(���) ��� From ����δ����� Where ����id=[1] and ��Դ;�� in(1,2) Group By ��Դ;��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��Դ;��=1"
        If rsTmp.RecordCount > 0 Then cur����δ�� = Val("" & rsTmp!���)
        rsTmp.Filter = "��Դ;��=2"
        If rsTmp.RecordCount > 0 Then GetPatientUnBalance = Val("" & rsTmp!���)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPatient_LostFocus()
    '���������ȷ��,����ʾ���ƻ�,��ָ�
    If txtPatient.Tag <> "" Then
        txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
        txtPatient.Text = txtPatient.Tag
    End If
    If gblnSeekName Then Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
End Sub

Private Sub txt�����ص�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt�����ص�.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt�����ص�)
            If Not rsTmp Is Nothing Then
                txt�����ص�.Text = rsTmp!����
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt�����ص�, KeyAscii
    End If
End Sub

Private Sub txt��λ�绰_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��λ�ʱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If ((Not IsNumeric(txt��λ�ʱ�.Text)) Or Len(txt��λ�ʱ�.Text) > 6 Or InStr(txt��λ�ʱ�.Text, ".") > 0) And txt��λ�ʱ�.Text <> "" Then
            Call SelectYouBian(txt��λ�ʱ�)
        End If
    End If
End Sub

Private Sub txt������λ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt������λ.Text <> "" Then
            Set rsTmp = GetOrgAddress(Me, txt������λ)
            If Not rsTmp Is Nothing Then
                txt������λ.Text = rsTmp!����
                txt������λ.Tag = rsTmp!ID
                txt��λ�绰.Text = Trim(rsTmp!�绰 & "")
                txt��λ������.Text = Trim(rsTmp!�������� & "")
                txt��λ�ʺ�.Text = Trim(rsTmp!�ʺ� & "")
            Else
                txt������λ.Tag = ""
            End If
        Else
            txt������λ.Tag = ""
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt������λ, KeyAscii
    End If
End Sub

Private Sub txt��ͥ��ַ�ʱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If ((Not IsNumeric(txt��ͥ��ַ�ʱ�.Text)) Or Len(txt��ͥ��ַ�ʱ�.Text) > 6 Or InStr(txt��ͥ��ַ�ʱ�.Text, ".") > 0) And txt��ͥ��ַ�ʱ�.Text <> "" Then
            Call SelectYouBian(txt��ͥ��ַ�ʱ�)
        End If
    End If
End Sub

Private Sub txt��ͥ��ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��ͥ��ַ.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt��ͥ��ַ)
            If Not rsTmp Is Nothing Then
                txt��ͥ��ַ.Text = rsTmp!����
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt��ͥ��ַ, KeyAscii
    End If
End Sub

Private Sub txt��ͥ�绰_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�������, KeyAscii
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If txt����.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not mCurSendCard.rs���� Is Nothing Then
            If mCurSendCard.rs����!�Ƿ��� = 1 Then
                If mCurSendCard.rs����!�ּ� <> 0 And Abs(CCur(txt����.Text)) > Abs(mCurSendCard.rs����!�ּ�) Then
                    MsgBox "" & mCurSendCard.str������ & "������ֵ���ܴ�������޼ۣ�" & Format(Abs(mCurSendCard.rs����!�ּ�), "0.00"), vbInformation, gstrSysName
                    txt����.SetFocus: Call zlControl.TxtSelAll(txt����): Exit Sub
                End If
                If mCurSendCard.rs����!ԭ�� <> 0 And Abs(CCur(txt����.Text)) < Abs(mCurSendCard.rs����!ԭ��) Then
                    MsgBox "" & mCurSendCard.str������ & "������ֵ����С������޼ۣ�" & Format(Abs(mCurSendCard.rs����!ԭ��), "0.00"), vbInformation, gstrSysName
                    txt����.SetFocus: Call zlControl.TxtSelAll(txt����): Exit Sub
                End If
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr(txt����.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii <> 13 Then
        If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        If Len(txt����.Text) = mCurSendCard.lng���ų��� - 1 And KeyAscii <> 8 Then
            txt����.Text = txt����.Text & Chr(KeyAscii)
            
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf txt����.Text = "" Then
        KeyAscii = 0: cmdOK.SetFocus  '������,ֱ������
    Else
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim lngBindPatientID As Long '�󶨿��Ĳ���ID
    Dim lng�䶯���� As Long '��Ƭ���ı䶯���� 11-�󶨿�,1-����
    txt����.Text = Trim(txt����.Text)
    Call ReLoadCardFee
    Call CheckFreeCard(txt����.Text)
    If mCurSendCard.lng���ų��� = Len(Trim(txt����.Text)) Then
        '���Ƿ��Ѿ��󶨻��߷���
        If WhetherTheCardBinding(Trim(txt����.Text), mCurSendCard.lng�����ID, lngBindPatientID) Then
            
            If mCurSendCard.bln���ƿ� And mCurSendCard.bln�ظ����� And lngBindPatientID > 0 Then
            
                lng�䶯���� = GetCardLastChangeType(Trim(txt����.Text), mCurSendCard.lng�����ID, lngBindPatientID)
                If lng�䶯���� = 11 Then
                    '����ǰ�
                    If MsgBox("����Ϊ��" & txt����.Text & "����{" & mCurSendCard.str������ & "}�Ŀ��Ѿ��벡�˱�ʶΪ��" & lngBindPatientID & "���Ľ����˰󶨣�" & vbCrLf & "�Ƿ�ȡ���ÿ��İ�?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                        Cancel = True
                        txt����.Text = ""
                        Exit Sub
                    End If
                    If BlandCancel(mCurSendCard.lng�����ID, Trim(txt����.Text), lngBindPatientID) Then
                        Exit Sub
                    End If
                End If
                
            End If
            
            MsgBox "�ÿ����Ѿ�����,���ܼ���.", vbInformation, gstrSysName
            Cancel = True
            txt����.Text = ""
            Exit Sub
            
        End If
    End If
End Sub

Private Sub CheckFreeCard(ByVal strCard As String)
'���ܣ���һ��ͨģʽ�µĿ��ţ��ϸ����Ʊ��ʱ������Ƿ���Ʊ�����÷�Χ�ڣ���Χ֮��Ŀ����շ�
    
    If txt����.Visible = False Then Exit Sub
    
    If Not mCurSendCard.rs���� Is Nothing And Val(txt����.Text) = 0 Then  '�Ȼָ�
        txt����.Text = Format(IIf(mCurSendCard.rs����!�Ƿ��� = 1, mCurSendCard.rs����!ȱʡ�۸�, mCurSendCard.rs����!�ּ�), "0.00")
    End If
    If mblnOneCard And mCurSendCard.bln�ϸ���� Then
        mCurSendCard.lng����ID = CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), strCard)
        If mCurSendCard.lng����ID <= 0 Then txt����.Text = "0.00"
    End If
    If Not mCurSendCard.rs���� Is Nothing And Val(txt����.Text) <> 0 Then
        If mCurSendCard.rs����!�Ƿ��� = 0 Then
            txt����.Text = Format(GetActualMoney(zlCommFun.GetNeedName(cbo�ѱ�.Text), mCurSendCard.rs����!������ĿID, mCurSendCard.rs����!�ּ�, mCurSendCard.rs����!�շ�ϸĿID), "0.00")
        End If
    End If
End Sub


Private Sub cbo�ѱ�_Click()
    Call LoadCardFee
End Sub

Private Sub txt������_GotFocus()
    If IsNumeric(txtԤ����.Text) And txt������.Text = "" Then
        txt������.Text = txt��λ������.Text
    End If
    zlControl.TxtSelAll txt������
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt������, KeyAscii
End Sub

Private Sub txt��ϵ�˵�ַ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��ϵ�˵�ַ.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt��ϵ�˵�ַ)
            If Not rsTmp Is Nothing Then
                txt��ϵ�˵�ַ.Text = rsTmp!����
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt��ϵ�˵�ַ, KeyAscii
    End If
End Sub

Private Sub txt��ϵ�˵绰_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��ϵ������_KeyPress(KeyAscii As Integer)
    CheckInputLen txt��ϵ������, KeyAscii
End Sub

Private Sub txt��Ժʱ��_GotFocus()
    Call OS.OpenImeByName
    Call zlControl.TxtSelAll(txt��Ժʱ��)
End Sub

Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    '�����:53408
    mbln�Ƿ�ɨ�����֤ = False

    Call Show�󶨿ؼ�(mbln�Ƿ�ɨ�����֤ And mblnɨ�����֤ǩԼ)
    
    If zl��ǰ�û����֤�Ƿ��(Val(IIf(Trim(txtPatient.Text) = "", "0", Trim(CStr(txtPatient.Tag))))) = True Then
            MsgBox "��ǰ�û������֤���Ѿ��󶨣��������޸������֤��", vbInformation, gstrSysName
            KeyAscii = 0
    End If
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_GotFocus()
    Call zlCommFun.OpenIme
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt��������_GotFocus()
    Call OS.OpenImeByName
    zlControl.TxtSelAll txt��������
End Sub

Private Sub txt���֤��_GotFocus()
    zlControl.TxtSelAll txt���֤��
    '�����:53408
    If mblnɨ�����֤ǩԼ = True Then
        OpenIDCard
    End If
End Sub

Private Sub txt�����ص�_GotFocus()
    zlControl.TxtSelAll txt�����ص�
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ͥ��ַ_GotFocus()
    zlControl.TxtSelAll txt��ͥ��ַ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ͥ��ַ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt��ͥ��ַ�ʱ�
End Sub

Private Sub txt��ͥ�绰_GotFocus()
    zlControl.TxtSelAll txt��ͥ�绰
End Sub

Private Sub txt��ϵ������_GotFocus()
    zlControl.TxtSelAll txt��ϵ������
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ϵ�˵�ַ_GotFocus()
    zlControl.TxtSelAll txt��ϵ�˵�ַ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ϵ�˵绰_GotFocus()
    zlControl.TxtSelAll txt��ϵ�˵绰
End Sub

Private Sub txt������λ_GotFocus()
    zlControl.TxtSelAll txt������λ
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��λ�绰_GotFocus()
    zlControl.TxtSelAll txt��λ�绰
End Sub

Private Sub txt��λ�ʱ�_GotFocus()
    zlControl.TxtSelAll txt��λ�ʱ�
End Sub

Private Sub txt��λ������_GotFocus()
    zlControl.TxtSelAll txt��λ������
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call SetBrushCardObject(True)
End Sub
Private Sub OpenIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����֤������
    '����:����
    '����:2012-08-31 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��ʼ���Կ�����
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '�򿪶�����
    mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    Call OpenPassKeyboard(txtPass, False)
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt��λ�ʺ�_GotFocus()
    zlControl.TxtSelAll txt��λ�ʺ�
End Sub

Private Sub cmdCancel_Click()
    Select Case mbytInState
        Case 0
            If mbytMode <> EMode.E����ԤԼ And (txtPatient.Tag <> "" Or txt����.Text <> "" Or txtסԺ��.Text <> "") Then
                If MsgBox("ȷ��Ҫ�����ǰ������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ClearCard
                    '84577
                    If tbcPage.Selected.Caption = "����" Then
                        If txtPatient.Enabled Then txtPatient.SetFocus
                    Else
                        tbcPage.Item(0).Selected = True
                    End If
                End If
                Exit Sub
            ElseIf gblnOK Then
                If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            Unload Me
        Case 1
            If MsgBox("ȷʵҪ�����޸��˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Unload Me
        Case 2
            Unload Me
    End Select
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
'���ܣ���ȡ������Ϣ
'˵������ȡʧ��ʱ��mrsInfo = Nothing
    Dim lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
     
    On Error GoTo errH

    If blnCard = True And objCard.���� Like "����*" Then   'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then
            '���ֻ��Ų�ѯ
            If IDKind.IsMobileNo(strInput) = False Then GoTo NotFoundPati:
            If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then Exit Function
        End If
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSQL = " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = " And A.�����=[1]"
    Else
        Select Case objCard.����
            Case "����"
                If Not gblnSeekName Then
                    MsgBox "��ˢ��������[-����ID]��[+סԺ��]��[*�����]�ȷ�ʽ��ȡ���˵���Ϣ��", vbInformation, gstrSysName
                    txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                Else
                    'ͨ������ģ�����Ҳ���(�������벡�˱�ʶʱ)
                    strPati = " Select 1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����," & _
                        " A.סԺ��,A.�����,A.סԺ����,trunc(C.��Ժ����,'dd') as ��Ժ����,trunc(C.��Ժ����,'dd') as ��Ժ����,A.��������,A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,zl_PatiType(A.����ID) ��������" & _
                        " From ������Ϣ A,���ű� B,������ҳ C" & _
                        " Where A.ͣ��ʱ�� is NULL And A.����ID=C.����ID(+) And Nvl(A.��ҳID,0)=C.��ҳID(+) And A.��ǰ����ID=B.ID(+) And Rownum<101" & _
                        " And A.���� Like [1]" & IIf(gintNameDays = 0, "", " And (A.�Ǽ�ʱ��>Trunc(Sysdate-[2]) Or A.����ʱ��>Trunc(Sysdate-[2]))")
                    strPati = strPati & " Union ALL " & _
                            "Select 0,0,-NULL,'[�²���]',NULL,NULL,-NULL,-NULL,-NULL,To_Date(NULL),To_Date(NULL),To_Date(NULL),NULL,NULL,NULL,NULL,'��ͨ����' From Dual"
                    strPati = strPati & " Order by ����ID,����,��Ժ���� Desc"
                    
                    vRect = GetControlRect(txtPatient.hWnd)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", gintNameDays)
                                
                    'ֻ��һ������ʱ,blncancel����false,��ȡ������Ҳ��һ��
                    If Not blnCancel Then
                        If rsTmp!ID = 0 Then '�����²���
                            strPati = txtPatient.Text
                            txtPatient.Text = ""
                            txtPatient_KeyPress (13)
                            txt����.Text = strPati
                            Exit Function
                        Else '�Բ���ID��ȡ
                            strInput = rsTmp!����ID
                            strSQL = " And A.����ID=[2]"
                        End If
                    Else
                        Call zlControl.TxtSelAll(txtPatient)
                        txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                    End If
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strSQL = " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = Val(objCard.�ӿ����)
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    '��Ժ֮ǰ�ĵ�����Ϣ��Ч,��Ժ���˵ĵ����Ա�����Ժ�Ǽǵ�Ϊ׼
    'ԤԼ�Ǽǲ���д"������Ϣ.��ǰ����ID,סԺ����"��,�����۵Ǽ�Ҫ
    '�����ԤԼ�ľ�Ҫ��ԤԼ��¼,��������һ��סԺ�ļ�¼
    '60500:������,2013-05-09,���������һ������,������ϢסԺ��Ϊ��
    If mbytInState = 0 And gblnÿ��סԺ��סԺ�� = False Then
            strPati = "Nvl(a.סԺ��," & vbNewLine & _
            "            (SELECT סԺ��" & vbNewLine & _
            "             FROM ������ҳ" & vbNewLine & _
            "             WHERE ����id = a.����id AND" & vbNewLine & _
            "                   ��ҳid = (SELECT MAX(��ҳid) FROM ������ҳ WHERE ����id = a.����id AND סԺ�� IS NOT NULL))) סԺ��,"
    Else
        strPati = "A.סԺ��,"
    End If
    '65973:������,2013-09-29,�µǼ���ȡҽ�Ƹ��ʽ
    strSQL = "Select A.����id, B.��ҳid, A.סԺ����, A.���￨��, A.����֤��, A.�����,B.���ۺ�," & strPati & "A.����, A.�Ա�,A.����, C.���� ��������," & vbNewLine & _
            "       Nvl(A.�ѱ�, B.�ѱ�) As �ѱ�, A.����, Nvl(B.����, A.����) ����, A.����, A.����, A.ѧ��, A.����״��, A.ְҵ, A.���, A.���֤��,A.�ֻ���, A.����֤��," & vbNewLine & _
            "       A.��������, A.�����ص�, A.��ͥ��ַ, A.��ͥ�绰, A.��ͥ��ַ�ʱ�, A.���ڵ�ַ, A.���ڵ�ַ�ʱ�, A.��ϵ�˹�ϵ, A.��ϵ������, A.��ϵ�˵�ַ,A.��ϵ�����֤��," & vbNewLine & _
            "       A.��ϵ�˵绰, A.������λ, A.��ͬ��λid, A.��λ�绰, A.��λ�ʱ�, A.��λ������, A.��λ�ʺ�, A.��Ժʱ��,Nvl(A.ҽ�Ƹ��ʽ, B.ҽ�Ƹ��ʽ) As ҽ�Ƹ��ʽ," & vbNewLine & _
            "       A.��ǰ����id, A.ҽ����, Nvl(B.����, A.����) As ����, Nvl(B.��������, 0) ��������,zl_PatiType(A.����ID) ��������,A.��ҳID �������" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ������� C" & vbNewLine & _
            "Where A.ͣ��ʱ�� Is Null And A.���� = C.���(+) And A.����id = B.����id(+) And A.��ҳid = B.��ҳid(+) And Not Exists" & vbNewLine & _
            " (Select 1 From ������ҳ Z Where Z.����id = A.����id And Z.��ҳid = 0)" & strSQL & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select A.����id, B.��ҳid, A.סԺ����, A.���￨��, A.����֤��, A.�����,B.���ۺ�," & strPati & "NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�, NVL(B.����,A.����) ����, C.���� ��������," & vbNewLine & _
            "       Nvl(A.�ѱ�, B.�ѱ�) As �ѱ�, A.����, Nvl(B.����, A.����) ����, A.����, A.����, A.ѧ��, A.����״��, A.ְҵ, A.���, A.���֤��,A.�ֻ���, A.����֤��," & vbNewLine & _
            "       A.��������, A.�����ص�, A.��ͥ��ַ, A.��ͥ�绰, A.��ͥ��ַ�ʱ�, A.���ڵ�ַ, A.���ڵ�ַ�ʱ�, A.��ϵ�˹�ϵ, A.��ϵ������, A.��ϵ�˵�ַ,A.��ϵ�����֤��," & vbNewLine & _
            "       A.��ϵ�˵绰, A.������λ, A.��ͬ��λid, A.��λ�绰, A.��λ�ʱ�, A.��λ������, A.��λ�ʺ�, A.��Ժʱ��,Nvl(A.ҽ�Ƹ��ʽ, B.ҽ�Ƹ��ʽ) As ҽ�Ƹ��ʽ," & vbNewLine & _
            "       A.��ǰ����id, A.ҽ����, Nvl(B.����, A.����) As ����, Nvl(B.��������, 0) ��������,zl_PatiType(A.����ID) ��������,A.��ҳID �������" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ������� C" & vbNewLine & _
            "Where A.ͣ��ʱ�� Is Null And A.���� = C.���(+) And A.����id = B.����id And B.��ҳid = 0" & strSQL
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then
        Set mrsInfo = Nothing: Exit Function
    End If
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = Nothing
End Function


Private Function GetMaxMinPage(lng����ID As Long, Optional blnMin As Boolean) As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select " & IIf(blnMin, "min", "max") & "(b.��ҳid) ��ҳid," & IIf(blnMin, "min", "max") & "(a.��ҳID) סԺ���� From ������Ϣ A,������ҳ B Where A.����ID = B.����ID And A.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If IsNull(rsTmp!��ҳID) And IsNull(rsTmp!סԺ����) Then
        GetMaxMinPage = -1
    Else
        GetMaxMinPage = IIf(IsNull(rsTmp!��ҳID) Or Nvl("" & rsTmp!��ҳID = 0), Val("" & rsTmp!סԺ����), Val("" & rsTmp!��ҳID))
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetMaxInHosTimes(lng����ID As Long) As Long
'����:��ȡ�������סԺ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    strSQL = "Select NVL(Max(סԺ����),0) סԺ���� From ������Ϣ where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    GetMaxInHosTimes = Val(rsTmp!סԺ����)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FillPatient()
'���ܣ���������ʱ,����mrsInfo�еĲ�����Ϣ��д������Ϣ��Ƭ
    txtPatient.Text = mrsInfo!����ID: txtPatient.Tag = mrsInfo!����ID
    
    If mbytKind = E�������۵Ǽ� Then
        If IsNull(mrsInfo!�����) Then
            txtסԺ��.Text = zlDatabase.GetNextNo(3)  'סԺ��סԺ��ģʽ��,ԤԼʱ������,����ʱ����
            mblnAuto = True
        Else
            txtסԺ��.Text = mrsInfo!�����
            txtסԺ��.Locked = True
        End If
    ElseIf mbytKind = EסԺ���۵Ǽ� Then
         txtסԺ��.Text = zlDatabase.GetNextNo(6) '���۵Ǽ�ʼ�ղ����µ����ۺ�
    Else
        If IsNull(mrsInfo!סԺ��) Or gblnÿ��סԺ��סԺ�� And mbytMode <> EMode.EԤԼ�Ǽ� Then
            If txtסԺ��.Visible And mbytKind = EKind.EסԺ��Ժ�Ǽ� Then txtסԺ��.Text = zlDatabase.GetNextNo(2)  'סԺ��סԺ��ģʽ��,ԤԼʱ������,����ʱ����
        Else
            txtסԺ��.Text = mrsInfo!סԺ��
        End If
    End If
    txtҽ����.Text = Nvl(mrsInfo!ҽ����)
    txtҽ����.Locked = Not IsNull(mrsInfo!����)
    txt����.Text = "" & mrsInfo!��������
    
    txt����.Text = mrsInfo!����
    
    If IsNull(mrsInfo!��ҳID) And IsNull(mrsInfo!�������) Then
        txtPages.Text = "1"
    Else
        If mbytMode = EMode.E����ԤԼ Or (mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0) Then
            txtPages.Text = GetMaxMinPage(mrsInfo!����ID) + 1
        Else
            txtPages.Text = Val(IIf(IsNull(mrsInfo!��ҳID) Or Val("" & mrsInfo!��ҳID) = 0, Val("" & mrsInfo!�������), Val("" & mrsInfo!��ҳID))) + 1
        End If
    End If
    If mbytInState = E���� And mbytKind = EסԺ��Ժ�Ǽ� And mbytMode <> EMode.EԤԼ�Ǽ� Then
        txtTimes.Text = GetMaxInHosTimes(mrsInfo!����ID) + 1
    Else
        txtTimes.Text = Nvl(mrsInfo!סԺ����)
    End If
    txtTimes.Tag = txtTimes.Text
    '65973:������,2013-09-29,���ҽ�Ƹ��ʽ
    cbo�Ա�.ListIndex = GetCboIndex(cbo�Ա�, IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�), mstrYBPati <> "")
    cbo�ѱ�.ListIndex = GetCboIndex(cbo�ѱ�, IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�), mstrYBPati <> "")
    cboҽ�Ƹ���.ListIndex = GetCboIndex(cboҽ�Ƹ���, IIf(IsNull(mrsInfo!ҽ�Ƹ��ʽ), "", mrsInfo!ҽ�Ƹ��ʽ), mstrYBPati <> "")
    cbo����.ListIndex = GetCboIndex(cbo����, IIf(IsNull(mrsInfo!����), "", mrsInfo!����), mstrYBPati <> "")
    cbo����.ListIndex = GetCboIndex(cbo����, IIf(IsNull(mrsInfo!����), "", mrsInfo!����), mstrYBPati <> "")
    cboѧ��.ListIndex = GetCboIndex(cboѧ��, IIf(IsNull(mrsInfo!ѧ��), "", mrsInfo!ѧ��), mstrYBPati <> "")
    cbo����״��.ListIndex = GetCboIndex(cbo����״��, IIf(IsNull(mrsInfo!����״��), "", mrsInfo!����״��), mstrYBPati <> "")
    cboְҵ.ListIndex = GetCboIndex(cboְҵ, IIf(IsNull(mrsInfo!ְҵ), "", mrsInfo!ְҵ), mstrYBPati <> "")
    cbo���.ListIndex = GetCboIndex(cbo���, IIf(IsNull(mrsInfo!���), "", mrsInfo!���), mstrYBPati <> "")
    cbo��ϵ�˹�ϵ.ListIndex = GetCboIndex(cbo��ϵ�˹�ϵ, IIf(IsNull(mrsInfo!��ϵ�˹�ϵ), "", mrsInfo!��ϵ�˹�ϵ), mstrYBPati <> "")
    If mstrYBPati <> "" Then cbo��������.ListIndex = GetCboIndex(cbo��������, Nvl(mrsInfo!��������, "��ͨ����"), True)  'ҽ����֤����
    '����27676 by lesfeng 2010-01-26 �����Ա𡢷ѱ𡢹��������塢ѧ��������״����ְҵ�����
    If cbo�Ա�.ListIndex = -1 Then Call SetCboDefault(cbo�Ա�)
    If cbo�ѱ�.ListIndex = -1 Then Call SetCboDefault(cbo�ѱ�)
    If cboҽ�Ƹ���.ListIndex = -1 Then Call SetCboDefault(cboҽ�Ƹ���)
    If cbo����.ListIndex = -1 Then Call SetCboDefault(cbo����)
    If cbo����.ListIndex = -1 Then Call SetCboDefault(cbo����)
    If cboѧ��.ListIndex = -1 Then Call SetCboDefault(cboѧ��)
    If cbo����״��.ListIndex = -1 Then Call SetCboDefault(cbo����״��)
    If cboְҵ.ListIndex = -1 Then Call SetCboDefault(cboְҵ)
    If cbo���.ListIndex = -1 Then Call SetCboDefault(cbo���)
    
    Call LoadOldData("" & mrsInfo!����, txt����, cbo���䵥λ)
    mblnChange = False
    txt��������.Text = Format(IIf(IsNull(mrsInfo!��������), "____-__-__", mrsInfo!��������), "YYYY-MM-DD")
    mblnChange = True
    
    If Not IsNull(mrsInfo!��������) Then
        If mbytInState <> 2 And mbytInState <> 1 Then txt����.Text = ReCalcOld(CDate(Format(mrsInfo!��������, "YYYY-MM-DD HH:MM:SS")), cbo���䵥λ, Val(mrsInfo!����ID), , CDate(txt��Ժʱ��.Text)) '���ݳ���������������
        If CDate(txt��������.Text) - CDate(mrsInfo!��������) <> 0 Then
            mblnChange = False
            txt����ʱ��.Text = Format(mrsInfo!��������, "HH:MM")
            mblnChange = True
        End If
    Else
        txt����ʱ��.Text = "__:__"
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
    
    cbo���䵥λ.Tag = txt����.Text & "_" & cbo���䵥λ.Text
    
    mblnChange = False
    txt���֤��.Text = "" & mrsInfo!���֤��
    mblnChange = True
    txt����֤��.Text = "" & mrsInfo!����֤��
    txt����.Text = Nvl(mrsInfo!����)
    txt��ͥ�绰.Text = IIf(IsNull(mrsInfo!��ͥ�绰), "", mrsInfo!��ͥ�绰)
    txt��ͥ��ַ�ʱ�.Text = IIf(IsNull(mrsInfo!��ͥ��ַ�ʱ�), "", mrsInfo!��ͥ��ַ�ʱ�)
    txt���ڵ�ַ�ʱ�.Text = IIf(IsNull(mrsInfo!���ڵ�ַ�ʱ�), "", mrsInfo!���ڵ�ַ�ʱ�)
    txt��ϵ������.Text = IIf(IsNull(mrsInfo!��ϵ������), "", mrsInfo!��ϵ������)
    txt��ϵ�˵绰.Text = IIf(IsNull(mrsInfo!��ϵ�˵绰), "", mrsInfo!��ϵ�˵绰)
    txt��ϵ�����֤��.Text = IIf(IsNull(mrsInfo!��ϵ�����֤��), "", mrsInfo!��ϵ�����֤��)
    txt������λ.Text = IIf(IsNull(mrsInfo!������λ), "", mrsInfo!������λ)
    txt������λ.Tag = IIf(IsNull(mrsInfo!��ͬ��λID), "", mrsInfo!��ͬ��λID)
    txt��λ�绰.Text = IIf(IsNull(mrsInfo!��λ�绰), "", mrsInfo!��λ�绰)
    txt��λ�ʱ�.Text = IIf(IsNull(mrsInfo!��λ�ʱ�), "", mrsInfo!��λ�ʱ�)
    txt��λ������.Text = IIf(IsNull(mrsInfo!��λ������), "", mrsInfo!��λ������)
    txt��λ�ʺ�.Text = IIf(IsNull(mrsInfo!��λ�ʺ�), "", mrsInfo!��λ�ʺ�)
    txtMobile.Text = "" & mrsInfo!�ֻ���
    
    If gbln���ýṹ����ַ Then
        Call ReadStructAddress(CLng(Nvl(mrsInfo!����ID, 0)), CLng(Nvl(mrsInfo!��ҳID, 0)), PatiAddress)
        txt�����ص�.Text = PatiAddress(E_IX_�����ص�).Value
        txt����.Text = PatiAddress(E_IX_����).Value
        txt��ͥ��ַ.Text = PatiAddress(E_IX_��סַ).Value
        txt���ڵ�ַ.Text = PatiAddress(E_IX_���ڵ�ַ).Value
        txt��ϵ�˵�ַ.Text = PatiAddress(E_IX_��ϵ�˵�ַ).Value
    Else
        txt�����ص�.Text = IIf(IsNull(mrsInfo!�����ص�), "", mrsInfo!�����ص�)
        txt����.Text = Nvl(mrsInfo!����)
        txt��ͥ��ַ.Text = IIf(IsNull(mrsInfo!��ͥ��ַ), "", mrsInfo!��ͥ��ַ)
        txt���ڵ�ַ.Text = IIf(IsNull(mrsInfo!���ڵ�ַ), "", mrsInfo!���ڵ�ַ)
        txt��ϵ�˵�ַ.Text = IIf(IsNull(mrsInfo!��ϵ�˵�ַ), "", mrsInfo!��ϵ�˵�ַ)
    End If
    '�����:56599
    Call Load�����������Ϣ(Val(Nvl(mrsInfo!����ID, 0)))
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        'ҽ���Ķ�
        If txt����.Text = "" And cmdYB.Enabled And cmdYB.Visible Then
            Call cmdYB_Click
            Call EMPI_LoadPati
            Call ReLoadCardFee
            Exit Sub
        End If
        
        
        If mbytInState = 0 Then
            If txt����.Text = "" Then
                If Not mrsInfo Is Nothing Then
                    txt����.Text = mrsInfo!���� '������Ϊ�����,�ֲ��޸�,���Զ��ָ�
                    Call zlCommFun.PressKey(vbKeyTab)
                Else
                    MsgBox "�������벡��������", vbInformation, gstrSysName
                    txt����.SetFocus
                End If
            Else
                If Not mrsInfo Is Nothing Then
                    Call zlCommFun.PressKey(vbKeyTab) '�޸�����
                Else
                    If txtPatient.Tag = "" And InStr(mstrPrivs, "�����ҽ������") > 0 Then '�����δ����
                        txtPatient.Text = zlDatabase.GetNextNo(1) '�²���ID
                        txtPatient.Tag = txtPatient.Text
                        '93974
                        txt��Ժʱ��.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                        If txtסԺ��.Text = "" And txtסԺ��.Visible Then
                            If mbytKind = EסԺ��Ժ�Ǽ� Then
                                txtסԺ��.Text = zlDatabase.GetNextNo(2)
                            ElseIf mbytKind = EסԺ���۵Ǽ� Then
                                txtסԺ��.Text = zlDatabase.GetNextNo(6)
                            ElseIf mbytKind = E�������۵Ǽ� Then
                                txtסԺ��.Text = zlDatabase.GetNextNo(3)
                            End If
                        End If
                    End If
                    Call EMPI_LoadPati(1)  '�µǼ�
                    Call ReLoadCardFee(True)
                    Call zlCommFun.PressKey(vbKeyTab)
                End If
            End If
        Else
            If txt����.Text = "" Then
                MsgBox "�������벡��������", vbInformation, gstrSysName
                txt����.SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    Else
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            Call CheckInputLen(txt����, KeyAscii)
        End If
    End If
End Sub

Private Sub txt�ʺ�_GotFocus()
    If IsNumeric(txtԤ����.Text) And txt�ʺ�.Text = "" Then
        txt�ʺ�.Text = txt��λ�ʺ�.Text
    End If
    zlControl.TxtSelAll txt�ʺ�
End Sub

Private Sub txt�ʺ�_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt�ʺ�, KeyAscii
End Sub

Private Sub txtסԺ��_GotFocus()
    zlControl.TxtSelAll txtסԺ��
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If mbytInState = 0 Then
            If txtסԺ��.Text = "" Then
                If mbytKind = E�������۵Ǽ� Then
                    If Not mrsInfo Is Nothing Then
                        txtסԺ��.Text = zlDatabase.GetNextNo(3)
                        mblnAuto = True
                        txt����.SetFocus
                    ElseIf Not txtסԺ��.Locked Then
                        MsgBox "�������벡������ţ�", vbInformation, gstrSysName
                        txtסԺ��.SetFocus
                    Else
                        Call zlCommFun.PressKey(vbKeyTab)
                    End If
                ElseIf mbytKind = EסԺ���۵Ǽ� Then
                    txtסԺ��.Text = zlDatabase.GetNextNo(6)
                    txt����.SetFocus
                Else
                    If Not mrsInfo Is Nothing Then
                        If Nvl(mrsInfo!סԺ��, 0) = 0 Then '������Ϊ�����,�ֲ��޸�,���Զ��ָ�,(ҽ����֤��û��סԺ��,��Ҫ��������)
                            txtסԺ��.Text = zlDatabase.GetNextNo(2)
                        Else
                            txtסԺ��.Text = mrsInfo!סԺ��
                        End If
                        txt����.SetFocus
                    ElseIf Not txtסԺ��.Locked Then
                        MsgBox "�������벡��סԺ�ţ�", vbInformation, gstrSysName
                        txtסԺ��.SetFocus
                    Else
                        Call zlCommFun.PressKey(vbKeyTab)
                    End If
                End If
            Else
                If Not mrsInfo Is Nothing Then
                    txt����.SetFocus  '�޸�סԺ��
                Else
                    If txtPatient.Tag = "" And InStr(mstrPrivs, "�����ҽ������") > 0 Then '�����δ����
                        txtPatient.Text = zlDatabase.GetNextNo(1) '�²���ID
                        txtPatient.Tag = txtPatient.Text
                    End If
                    txt����.SetFocus
                End If
                Call txtסԺ��_Validate(False)
            End If
        Else
            If txtסԺ��.Text = "" Then
                If mbytKind = EסԺ��Ժ�Ǽ� And Not txtסԺ��.Locked Then
                    MsgBox "�������벡��סԺ�ţ�", vbInformation, gstrSysName
                    txtסԺ��.SetFocus
                ElseIf mbytKind = EסԺ���۵Ǽ� Then
                    txtסԺ��.Text = zlDatabase.GetNextNo(6)
                    txt����.SetFocus
                ElseIf mbytKind = E�������۵Ǽ� Then
                    Call zlCommFun.PressKey(vbKeyTab)
                End If
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        Else
            mblnAuto = False
        End If
    End If
End Sub

Private Sub InitInputTabStop()
'���ܣ����ݱ������ù��Ҫ��λ��������Ŀ
    Dim i As Integer, j As Integer
    Dim strPara As String
    Dim arrTmp As Variant
    Dim arrSubTmp As Variant
    Dim strInputItem As String
    Dim objTmp As Object
    Dim intBegin As Integer
    Dim intEnd As Integer
    Dim strItem As String
    
    '����ֵ:������Ŀ,��ֹ¼��(0\1),��������(0\1),������(0\1) '����,0,1,1|����,0,1,1|...
    strPara = zlDatabase.GetPara("���������", glngSys, mlngModul)
    Set mrsInputSet = Rec.CopyNew(Nothing, , , Array("������Ŀ", adVarChar, 50, Empty, "��ֹ¼��", adInteger, 1, Empty, "��������", adInteger, 1, Empty, "������", adInteger, 1, Empty, "�ؼ���", adVarChar, 50, Empty, "�ؼ��±�", adInteger, 2, Empty))
    '
    '1)����Ҫ�������ƵĿؼ����ü�¼����¼����
    arrTmp = Split(C_���������, "|")
    For i = LBound(arrTmp) To UBound(arrTmp)
        arrSubTmp = Split(arrTmp(i), ",")
        strInputItem = arrSubTmp(0)      '������Ŀ
        'һ��������Ŀ���ܿ��ƶ���ؼ�:���� �������� ����� txt��������,txt����ʱ��
        For j = LBound(arrSubTmp) + 1 To UBound(arrSubTmp)
            mrsInputSet.AddNew
            mrsInputSet!������Ŀ = strInputItem
            strItem = arrSubTmp(j)
            intBegin = InStr(strItem, "(")
            If intBegin > 0 Then
                intEnd = InStr(strItem, ")")
                mrsInputSet!�ؼ��� = Mid(strItem, 1, intBegin - 1)
                mrsInputSet!�ؼ��±� = Val(Mid(strItem, intBegin + 1, intEnd - intBegin + 1))
            Else
                mrsInputSet!�ؼ��� = strItem
            End If
            mrsInputSet!������ = 1 'ȱʡ����Ϊ1-������
            mrsInputSet.Update
        Next
    Next
    
    If strPara <> "" Then
        '2�����������õ�ֵ���µ���¼����
        arrTmp = Split(strPara, "|")
        For i = LBound(arrTmp) To UBound(arrTmp)
            arrSubTmp = Split(arrTmp(i), ",")
            mrsInputSet.Filter = "������Ŀ ='" & arrSubTmp(0) & "'"
            For j = 1 To mrsInputSet.RecordCount
                mrsInputSet!��ֹ¼�� = Val(arrSubTmp(1))
                mrsInputSet!�������� = Val(arrSubTmp(2))
                mrsInputSet!������ = Val(arrSubTmp(3))
                mrsInputSet.Update
                mrsInputSet.MoveNext
            Next
        Next
    End If
    mrsInputSet.Filter = "" '
  
    For i = 1 To mrsInputSet.RecordCount
        Set objTmp = CallByName(Me, mrsInputSet!�ؼ��� & "", VbGet)
        If Not IsNull(mrsInputSet!�ؼ��±�) Then
            Set objTmp = objTmp(mrsInputSet!�ؼ��±�)
        End If
        '��ֹ¼��
        objTmp.Enabled = Val(mrsInputSet!��ֹ¼�� & "") = 0
        objTmp.BackColor = IIf(objTmp.Enabled, C_COLOR_Enabled, C_COLOR_UNEnabled)
        '����Ƿ����
        objTmp.TabStop = Val(mrsInputSet!������ & "")
        mrsInputSet.MoveNext
    Next
End Sub

Private Sub txt��������_LostFocus()
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
        txt��������.SetFocus
    End If
End Sub
'����26779 by lesfeng 2009-12-10
Private Sub LoadBedInfo(lng����ID As Long, Optional lng����ID As Long)
'���ܣ����ң�������ʾ������Ժ��������λ��
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer, strTmp As String
    Dim strSqlTmp1 As String, strSqlTmp2 As String
    Dim strSqlTmp3 As String, strSqlTmp4 As String
    
    Dim intFlag As Integer
    
    On Error GoTo errHandle
    intFlag = 0
    If lng����ID = lng����ID Then
        intFlag = 1
    Else
        intFlag = 2
    End If
    strSqlTmp1 = " And B.����ID=[2]"
    strSqlTmp2 = " And A.����id = [2]"
    strSqlTmp3 = " And B.����ID=[1]"
    strSqlTmp4 = " And A.����id = [1]"
    strSQL = " Select Sum(������Ժ) as ������Ժ, Sum(�����մ�) as �����մ�,Sum(������Ժ) As ������Ժ, Sum(���ҿմ�) As ���ҿմ�" & _
             " From ( Select Count(A.����id) As ������Ժ, 0 As �����մ�,0 As ������Ժ,0 As ���ҿմ�" & _
             "          From ������Ϣ A,��Ժ���� B" & _
             "         Where A.����ID=B.����ID " & strSqlTmp1 & _
             "         Union All " & _
             "        Select 0 As ������Ժ, Count(A.����) As �����մ� ,0 As ������Ժ,0 As ���ҿմ�" & _
             "          From ��λ״����¼ A" & _
             "        Where A.��λ���� <> '�Ǳ�' and A.��λ���� <> '�໤' And A.״̬ = '�մ�'" & strSqlTmp2 & _
             "         Union All " & _
             "       Select 0 As ������Ժ, 0 As �����մ�,Count(A.����id)  As ������Ժ,0 As ���ҿմ�" & _
             "          From ������Ϣ A,��Ժ���� B" & _
             "         Where A.����ID =B.����ID " & strSqlTmp3 & _
             "         Union All " & _
             "        Select 0 As ������Ժ, 0 As �����մ� ,0 As ������Ժ,Count(A.����) As ���ҿմ�" & _
             "          From ��λ״����¼ A" & _
             "        Where A.��λ���� <> '�Ǳ�' and A.��λ���� <> '�໤' And A.״̬ = '�մ�'" & strSqlTmp4 & ") "
   
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng����ID)
    '���� 27097 by lesfeng 2009-12-25 ���ǲ�ȷ�����һ��߲�����Ӧ���ų����������
    If Not rsTemp.EOF Then
        If intFlag = 1 Then
            If gbln��ѡ���� Then
                If InStr(1, cbo��Ժ����.Text, "-") > 0 Then
                    strTmp = Split(cbo��Ժ����.Text, "-")(1)
                Else
                    strTmp = Trim(cbo��Ժ����.Text)
                End If
            Else
                If InStr(1, cbo��Ժ����.Text, "-") > 0 Then
                    strTmp = Split(cbo��Ժ����.Text, "-")(1)
                Else
                    strTmp = Trim(cbo��Ժ����.Text)
                End If
            End If
            strTmp = strTmp & "����Ժ���� " & IIf(IsNull(rsTemp!������Ժ), 0, rsTemp!������Ժ) & "����λ�� " & IIf(IsNull(rsTemp!���ҿմ�), 0, rsTemp!���ҿմ�)
        Else
            If gbln��ѡ���� Then
                If InStr(1, cbo��Ժ����.Text, "-") > 0 Then
                    strTmp = Split(cbo��Ժ����.Text, "-")(1)
                Else
                    strTmp = Trim(cbo��Ժ����.Text)
                End If
                strTmp = strTmp & "����Ժ���� " & IIf(IsNull(rsTemp!������Ժ), 0, rsTemp!������Ժ) & "����λ�� " & IIf(IsNull(rsTemp!�����մ�), 0, rsTemp!�����մ�)
                
                If lng����ID <> 0 Then
                    If InStr(1, cbo��Ժ����.Text, "-") > 0 Then
                        strTmp = strTmp & "," & Split(cbo��Ժ����.Text, "-")(1)
                    Else
                        strTmp = strTmp & "," & Trim(cbo��Ժ����.Text)
                    End If
                    strTmp = strTmp & "����Ժ���� " & IIf(IsNull(rsTemp!������Ժ), 0, rsTemp!������Ժ) & "����λ�� " & IIf(IsNull(rsTemp!���ҿմ�), 0, rsTemp!���ҿմ�)
                End If
            Else
                If InStr(1, cbo��Ժ����.Text, "-") > 0 Then
                    strTmp = Split(cbo��Ժ����.Text, "-")(1)
                Else
                    strTmp = Trim(cbo��Ժ����.Text)
                End If
                strTmp = strTmp & "����Ժ����" & IIf(IsNull(rsTemp!������Ժ), 0, rsTemp!������Ժ) & "����λ��" & IIf(IsNull(rsTemp!���ҿմ�), 0, rsTemp!���ҿմ�)
                
                If lng����ID <> 0 Then
                    If InStr(1, cbo��Ժ����.Text, "-") > 0 Then
                        strTmp = strTmp & "," & Split(cbo��Ժ����.Text, "-")(1)
                    Else
                        strTmp = strTmp & "," & Trim(cbo��Ժ����.Text)
                    End If
                    strTmp = strTmp & "����Ժ���� " & IIf(IsNull(rsTemp!������Ժ), 0, rsTemp!������Ժ) & "����λ�� " & IIf(IsNull(rsTemp!�����մ�), 0, rsTemp!�����մ�)
                End If

            End If
        End If
    Else
        strTmp = ""
    End If

    lblBedInfo.Caption = strTmp
    If gbln���ƿմ� Then
        If Val(rsTemp!���ҿմ�) = 0 And Val(rsTemp!�����մ�) = 0 Then
            mbln�մ� = True
                Else
                        mbln�մ� = False
        End If
    Else
        mbln�մ� = False
    End If
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub LoadBed(str�Ա� As String, lng����ID As Long, Optional lng����ID As Long)
'���ܣ����ݵ�ǰ�����Ա𣬿��ң��������ؿ��õĲ���
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strTmp As String, strPreBed As String
    Dim blnFind As Boolean
    
    If Not (gbln��Ժ��� And mbytMode <> EMode.EԤԼ�Ǽ� And mbytInState = EState.E����) Then Exit Sub
    
    If cbo��λ.ListCount > 1 And InStr(Trim(cbo��λ.Text), " ") > 1 Then strPreBed = Trim(Mid(Trim(cbo��λ.Text), 1, InStr(Trim(cbo��λ.Text), " ") - 1))
    cbo��λ.Clear: cbo��λ.Tag = ""
    cbo��λ.AddItem "�����䴲λ"
    If lng����ID <> 0 Then
        cbo��λ.AddItem "�����ͥ����"
    End If
    cbo��λ.ListIndex = 0
        
    '��λҪ��ʱȡ����ʹ�û���
    Set rsTmp = GetFreeBeds(lng����ID, lng����ID, str�Ա�)
    For i = 1 To rsTmp.RecordCount
        cbo��λ.AddItem " " & rsTmp!���� & Space(10 - Len(rsTmp!����)) & " " & rsTmp!�Ա���� & IIf(IsNull(rsTmp!�����), "", " ����:" & rsTmp!����� & "|") & _
            IIf(IsNull(rsTmp!�����) Or ((Not IsNull(rsTmp!�����)) And Trim(Nvl(rsTmp!�Ա�) = "")), "", "(" & Nvl(rsTmp!�Ա�) & ")") & _
            Space(15 - Len(IIf(IsNull(rsTmp!�����), "", " ����:" & IsNull(rsTmp!�����))) - Len(IIf(IsNull(rsTmp!�����) Or ((Not IsNull(rsTmp!�����)) And Trim(Nvl(rsTmp!�Ա�) = "")), "", "(" & Nvl(rsTmp!�Ա�) & ")"))) & _
            Nvl(rsTmp!��λ�ȼ�)
        If rsTmp!���� = strPreBed And Not blnFind Then cbo��λ.ListIndex = cbo��λ.NewIndex: cbo��λ.Tag = rsTmp!����
        If mblnAppoint And rsTmp!���� = mstrAppointBed Then
            cbo��λ.ListIndex = cbo��λ.NewIndex: blnFind = True
            cbo��λ.Tag = rsTmp!����
        End If
        rsTmp.MoveNext
    Next
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc(".") And InStr(txt������.Text, ".") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Function Check������Ϣ() As Boolean
    Check������Ϣ = True
    
    If txt������.Tag <> "" Then
    '�޸�ʱ����ɾ��,Ҫɾ���͵�������Ϣ������ȥɾ��
        If Trim(txt������.Text) = "" Then
            MsgBox "�޸ĵǼ���Ϣʱ������ɾ���Ѿ����ڵĵ�����Ϣ!", vbInformation, gstrSysName
            If txt������.Enabled Then txt������.SetFocus
            Check������Ϣ = False
            Exit Function
        End If
    End If
    
    If Not IsNumeric(txt������.Text) And Trim(txt������.Text) <> "" Then
        MsgBox "��������ȷ�ĵ�����,������Ҫ������ֵ!", vbInformation, gstrSysName
        If txt������.Enabled Then txt������.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
    If IsNumeric(txt������.Text) And Trim(txt������.Text) = "" Then
        MsgBox "�����뵣��������,�����˲���Ϊ��!", vbInformation, gstrSysName
        If txt������.Enabled Then txt������.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
    
    'ֻҪ���뵣����,��ѡ���˵���ʱ��,��ѡ������ʱ����,�ͱ�ʾҪ¼�뵣����Ϣ
    If Trim(txt������.Text) <> "" Or Not IsNull(dtp����ʱ��.Value) Or chk��ʱ����.Value = 1 Then
        If Val(txt������.Text) = 0 Then
            MsgBox "�����뵣����,�������Ϊ��!", vbInformation, gstrSysName
            If txt������.Enabled Then txt������.SetFocus
            Check������Ϣ = False
            Exit Function
        End If
    End If
    
    '����ʱ�޲���С����Ժʱ��
    If Not IsNull(dtp����ʱ��.Value) And dtp����ʱ��.Enabled Then
        If dtp����ʱ��.Value < CDate(txt��Ժʱ��.Text) Then
            MsgBox "��������ʱ�䲻��������Ϊ��Ժʱ��֮ǰ!!", vbInformation, gstrSysName
            If dtp����ʱ��.Enabled Then dtp����ʱ��.SetFocus
            Check������Ϣ = False
            Exit Function
        End If
    End If
    
    If chk��ʱ����.Value = 1 Then
        If Not IsNull(dtp����ʱ��.Value) Or chkUnlimit.Value = 1 Then
            MsgBox "��ʱ�������������õ���ʱ�޻��޵�����!", vbInformation, gstrSysName
            If chk��ʱ����.Enabled Then chk��ʱ����.SetFocus
            Check������Ϣ = False
            Exit Function
        End If
    End If
    
    If zlCommFun.ActualLen(Trim(txtReason.Text)) > 50 Then
        MsgBox "����ԭ�������������� 25 �����ֻ� 50 ���ַ���", vbInformation, gstrSysName
        txtReason.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
End Function

Private Function CanFocus(ctlError As Control) As Boolean
    CanFocus = ctlError.Enabled And ctlError.Visible
End Function


Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim rsSimilar As New ADODB.Recordset
    Dim blnOk As Boolean, strSimilar As String, strInfo As String
    Dim lng�ӿڱ�� As Long, strBalanceInfor As String
    Dim i As Long, lng����ID As Long, blnErr As Boolean
    Dim lngTmp As Long, strno As String
    Dim blnTmp As Boolean   '�Ƿ���Ϊ����ű�ռ�ö���������
    Dim bln������Ϣ����, blnMod   As Boolean
    Dim str�������� As String, str���� As String, strAge As String, str�Ա� As String, strErrInfo As String
    Dim strMsg As String
    Dim objTmp As Object
    
    '�����:56599
    tbcPage.Item(0).Selected = True
     
    '65965:������,2013-09-24,����Ԥ����ʾǧλλ��ʽ
    If Not CheckFormInput(Me, "txt�������,txt��ҽ���", "txtԤ����") Then Exit Sub
    
    '�Ϸ��Լ��
    '�����:53408
    If IIf(zlDatabase.GetPara("ɨ�����֤ǩԼ", glngSys, glngModul) = "1", 1, 0) = 0 And ((mCurSendCard.str������ = "�������֤" And Trim(txt����.Text) <> "") Or Trim(txt֧������.Text) <> "") Then
         MsgBox "��û��Ȩ�޽���ǩԼ����,�뵽�������������á�ɨ�����֤ǩԼ����", vbOKOnly + vbInformation, gstrSysName
         txt����.Text = ""
         txtPass.Text = ""
         txtAudi.Text = ""
         If txt����.Visible = True Then txt����.SetFocus
         Exit Sub
    End If
    
    If ((Not IsNumeric(txt���ڵ�ַ�ʱ�.Text)) Or Len(txt���ڵ�ַ�ʱ�.Text) > 6 Or InStr(txt���ڵ�ַ�ʱ�.Text, ".") > 0) And txt���ڵ�ַ�ʱ�.Text <> "" Then
        MsgBox "�ʱ��ʽ����,��������ȷ���ʱ�!" & vbCrLf & "����ȷ�ʱ��ʽΪ��λ�����ֱ��롿", vbInformation, gstrSysName
        If CanFocus(txt���ڵ�ַ�ʱ�) = True Then txt���ڵ�ַ�ʱ�.SetFocus: Exit Sub
    End If
    If ((Not IsNumeric(txt��λ�ʱ�.Text)) Or Len(txt��λ�ʱ�.Text) > 6 Or InStr(txt��λ�ʱ�.Text, ".") > 0) And txt��λ�ʱ�.Text <> "" Then
        MsgBox "�ʱ��ʽ����,��������ȷ���ʱ�!" & vbCrLf & "����ȷ�ʱ��ʽΪ��λ�����ֱ��롿", vbInformation, gstrSysName
        If CanFocus(txt��λ�ʱ�) = True Then txt��λ�ʱ�.SetFocus: Exit Sub
    End If
    If ((Not IsNumeric(txt��ͥ��ַ�ʱ�.Text)) Or Len(txt��ͥ��ַ�ʱ�.Text) > 6 Or InStr(txt��ͥ��ַ�ʱ�.Text, ".") > 0) And txt��ͥ��ַ�ʱ�.Text <> "" Then
        MsgBox "�ʱ��ʽ����,��������ȷ���ʱ�!" & vbCrLf & "����ȷ�ʱ��ʽΪ��λ�����ֱ��롿", vbInformation, gstrSysName
        If CanFocus(txt��ͥ��ַ�ʱ�) = True Then txt��ͥ��ַ�ʱ�.SetFocus: Exit Sub
    End If
    
    If Trim(txt֧������.Text) <> "" And Trim(txt���֤��.Text) <> "" Then
           If �Ƿ��Ѿ�ǩԼ(txt���֤��.Text) Then
                 MsgBox "���֤����Ϊ:" & txt���֤��.Text & "�Ѿ�ǩԼ�����ظ�ǩԼ��", vbOKOnly + vbInformation, gstrSysName
                 txt֧������.Text = ""
                 If txt֧������.Visible = True Then txt֧������.SetFocus
                 Exit Sub
           End If
    End If
    
    If mbln�Ƿ�ɨ�����֤ = False And mCurSendCard.str������ = "�������֤" And txt����.Text <> "" Then
            MsgBox "�����ֻ֤����ˢ���ķ�ʽ���У��������ֶ��������֤���а�!", vbOKOnly + vbInformation, gstrSysName
            txt����.Text = ""
            txtPass.Text = ""
            txtAudi.Text = ""
            txt֧������.Text = ""
            txt��֤����.Text = ""
            If txt����.Visible = True Then txt����.SetFocus
            Exit Sub
    End If
    
    If mbln�Ƿ�ɨ�����֤ = False And mCurSendCard.str������ <> "�������֤" And txt֧������.Text <> "" Then
            MsgBox "�����ֻ֤����ˢ���ķ�ʽ���У��������ֶ��������֤���а�!", vbOKOnly + vbInformation, gstrSysName
            txt���֤��.Text = ""
            txt֧������.Text = ""
            txt��֤����.Text = ""
            If txt���֤��.Visible = True Then
                If CanFocus(txt���֤��) = True Then txt���֤��.SetFocus
            End If
        Exit Sub
    End If
    
    If Trim(txt֧������.Text) <> Trim(txt��֤����.Text) And (Trim(txt֧������.Text) <> "" Or Trim(txt��֤����.Text) <> "") Then
        MsgBox "������������벻һ��,����������", vbOKOnly + vbInformation, gstrSysName
        txt֧������.Text = "": txt��֤����.Text = ""
        If txt֧������.Visible = True Then txt֧������.SetFocus
        Exit Sub
    End If
    
    
    If txtPatient.Tag = "" Then
        MsgBox "����ȷ����Ժ���ˣ�", vbInformation, gstrSysName
        If Not txtPatient.TabStop Then
            txt����.SetFocus
        Else
            txtPatient.SetFocus
        End If
        Exit Sub
    End If
    If Trim(txtסԺ��.Text) = "" And mbytKind = EסԺ��Ժ�Ǽ� And mbytMode <> 1 Then  'סԺ�����²���,��������û��סԺ��
        MsgBox "�������벡��סԺ�ţ�", vbInformation, gstrSysName
        txtסԺ��.SetFocus: Exit Sub
    End If
    
    If txtTimes.Visible And txtTimes.Enabled Then
        If Not IsNumeric(txtTimes.Text) Then
            MsgBox "סԺ�������������֣�", vbInformation, gstrSysName
            txtTimes.SetFocus: Exit Sub
        End If
        If Val(txtTimes.Text) < Val(txtTimes.Tag) Then
            MsgBox "סԺ��������С���Ѵ��ڵĴ�����", vbInformation, gstrSysName
            txtTimes.SetFocus: Exit Sub
        End If
        If Val(txtTimes.Text) = 0 And mbytMode <> EMode.EԤԼ�Ǽ� And mbytKind = EסԺ��Ժ�Ǽ� Then
            MsgBox "סԺ��������Ϊ�㣡", vbInformation, gstrSysName
            txtTimes.SetFocus: Exit Sub
        End If
    End If
    
    If Trim(txt����.Text) = "" Then
        MsgBox "�������벡�˵�������", vbInformation, gstrSysName
        If CanFocus(txt����) = True Then txt����.SetFocus: Exit Sub
    End If
    If cbo�Ա�.ListIndex = -1 Then
        MsgBox "����ȷ�����˵��Ա�", vbInformation, gstrSysName
        If CanFocus(cbo�Ա�) = True Then cbo�Ա�.SetFocus: Exit Sub
    End If
    If txt��������.Enabled Then
        If Not IsDate(txt��������.Text) Then
            MsgBox "������ȷ���벡�˵ĳ������ڣ�", vbInformation, gstrSysName
            If CanFocus(txt��������) = True Then txt��������.SetFocus: Exit Sub
        End If
    End If
    If Trim(txt����.Text) = "" Then
        MsgBox "�������벡�˵����䣡", vbInformation, gstrSysName
        If CanFocus(txt����) = True Then txt����.SetFocus: Exit Sub
    End If
    
    '80505  ����"���������"ָ�������������Ŀ���
    mrsInputSet.Filter = "": blnTmp = False       '
    For i = 1 To mrsInputSet.RecordCount
        '����������Ŀ���
        If Val(mrsInputSet!�������� & "") = 1 Then
            Set objTmp = CallByName(Me, mrsInputSet!�ؼ��� & "", VbGet)
            If Not IsNull(mrsInputSet!�ؼ��±�) Then
                Set objTmp = objTmp(mrsInputSet!�ؼ��±�) '�ؼ�����
            End If
            blnTmp = False
            If objTmp.Enabled = True And objTmp.Visible Then
                If UCase(TypeName(objTmp)) = UCase("TextBox") Then
                    If Trim(objTmp.Text) = "" Then blnTmp = True
                ElseIf UCase(TypeName(objTmp)) = UCase("ComboBox") Then
                    If objTmp.ListIndex = -1 Then blnTmp = True
                ElseIf UCase(TypeName(objTmp)) = UCase("MaskEdBox") Then
                    If mrsInputSet!������Ŀ & "" = "��������" Then
                        blnTmp = False  '�������ں����������,�˴��ݲ����
                    Else
                        If Trim(objTmp.Text) = "" Then blnTmp = True
                    End If
                ElseIf UCase(TypeName(objTmp)) = UCase("PatiAddress") Then
                    If Trim(objTmp.Value) = "" Or objTmp.CheckNullValue() <> "" Then blnTmp = True
                End If
                If blnTmp Then
                    MsgBox "�������벡�˵�" & mrsInputSet!������Ŀ & "��", vbInformation, gstrSysName
                    If CanFocus(objTmp) = True Then objTmp.SetFocus
                    Exit Sub
                End If
            End If
        Else
            '���ڷǱ����������Ŀ�ṹ����ַ����һ��¼��һ���־�Ҫ���������¼�롣
            Set objTmp = CallByName(Me, mrsInputSet!�ؼ��� & "", VbGet)
            If Not IsNull(mrsInputSet!�ؼ��±�) Then
                Set objTmp = objTmp(mrsInputSet!�ؼ��±�) '�ؼ�����
            End If
            
            If objTmp.Enabled = True And objTmp.Visible Then
                If UCase(TypeName(objTmp)) = UCase("PatiAddress") Then
                    If Trim(objTmp.Value) <> "" And objTmp.CheckNullValue() <> "" Then
                        MsgBox "���˵�" & mrsInputSet!������Ŀ & "¼�벻����,������¼�����ɾ����¼�����ݡ�", vbInformation, gstrSysName
                        If CanFocus(objTmp) = True Then objTmp.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        mrsInputSet.MoveNext
    Next
    
    '76409,������,2014-08-06,����Ϸ��Լ��
    If txt����.Locked = False Then
        str���� = txt����.Text
        If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
        If str���� Like "Լ*" Then
            str���� = str���� & cbo���䵥λ.Text
        End If
        If IsDate(txt��������.Text) Then
            If txt����ʱ��.Text = "__:__" Then
                str�������� = Format(txt��������.Text, "YYYY-MM-DD")
            Else
                str�������� = Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS")
            End If
            strInfo = CheckAge(str����, str��������, CDate(txt��Ժʱ��.Text))
        Else
            strInfo = CheckAge(str����)
        End If
        If InStr(1, strInfo, "|") > 0 Then
            lngTmp = Val(Split(strInfo, "|")(0)) '1��ֹ,0��ʾ
            strInfo = Split(strInfo, "|")(1)
            If lngTmp = 1 Then
                MsgBox strInfo, vbInformation, gstrSysName
                If CanFocus(txt����) = True Then txt����.SetFocus: Exit Sub
            Else
                If MsgBox(strInfo & vbCrLf & vbCrLf & "���������������ڵ���ȷ�ԣ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If CanFocus(txt����) = True Then txt����.SetFocus: Exit Sub
                End If
            End If
        End If
    End If

    str�������� = ""
    '--81012,��ΰ��,2014-12-22,�������֤�Գ�������\����\�Ա� �ļ��
    If Trim(zlCommFun.GetNeedName(cbo����.Text)) = "�й�" Then
        If Not CheckLen(txt���֤��, 18) Then Exit Sub
        lngTmp = LenB(StrConv(Trim(txt���֤��.Text), vbFromUnicode))
        If lngTmp > 0 Then
            If CreatePublicPatient() Then
                If gobjPublicPatient.CheckPatiIdcard(Trim(txt���֤��.Text), str��������, strAge, str�Ա�, strErrInfo, CDate(txt��Ժʱ��.Text)) Then
                    '���޻�����Ϣ����Ȩ��
                    bln������Ϣ���� = InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";������Ϣ����;") > 0 And mbytInState = 1 And mblnHaveAdvice
                    strMsg = ""
                    '��������
                    If Trim(txt��������.Text) <> "____-__-__" Then
                        If CDate(Format(str��������, "YYYY-MM-DD")) <> CDate(Format(txt��������.Text, "YYYY-MM-DD")) Then
                            strMsg = "���֤�����еĳ�������[" & str�������� & "]�Ͳ��˳�������[" & Format(txt��������.Text, "YYYY-MM-DD") & "]��һ��"
                            '����
                            str���� = txt����.Text
                            If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
                            If str���� <> strAge Then
                                strMsg = strMsg & vbCrLf & "���֤�����е�����[" & strAge & "]�Ͳ�������[" & str���� & "]��һ��"
                            End If
                        End If
                    End If
                    '�Ա�
                    If InStr(cbo�Ա�.Text, str�Ա�) = 0 Then
                        strMsg = IIf(strMsg <> "", strMsg & vbCrLf, "") & "���֤�����е��Ա�[" & str�Ա� & "]�Ͳ����Ա�[" & zlCommFun.GetNeedName(cbo�Ա�.Text) & "]��һ��"
                    End If
                    
                    If mbytInState = 1 And mblnHaveAdvice Then
                        If strMsg <> "" Then
                            If MsgBox(strMsg & ",�Ƿ������" & vbCrLf & IIf(bln������Ϣ����, "ѡ���ǡ�,�����֤����Ϣ�滻���˵���Ϣ�����ҵ�����ݡ�", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If CanFocus(txt���֤��) = True Then txt���֤��.SetFocus: Exit Sub
                            Else
                                blnMod = True
                            End If
                        End If
                    Else
                        If strMsg <> "" Then
                            If MsgBox(strMsg & ",�Ƿ������" & vbCrLf & "ѡ���ǡ�,�����֤����Ϣ�滻���˵���Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If CanFocus(txt���֤��) = True Then txt���֤��.SetFocus: Exit Sub
                            Else
                                If CDate(Format(str��������, "YYYY-MM-DD")) <> CDate(Format(txt��������.Text, "YYYY-MM-DD")) Then
                                    txt��������.Text = str��������
                                    If mblnChange = False Then
                                        Call LoadOldData(strAge, txt����, cbo���䵥λ)
                                    End If
                                End If
                                Call cbo.Locate(cbo�Ա�, str�Ա�, False)
                            End If
                        End If
                    End If
                Else
                    MsgBox strErrInfo, vbInformation + vbOKOnly, gstrSysName
                    If CanFocus(txt���֤��) = True Then txt���֤��.SetFocus: Exit Sub
                End If
                
            End If
        End If
    End If
    
    If cbo�ѱ�.ListIndex = -1 Then
        MsgBox "����ȷ�����˷ѱ�", vbInformation, gstrSysName
        cbo�ѱ�.SetFocus: Exit Sub
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ�����˹�����", vbInformation, gstrSysName
        If CanFocus(cbo����) = True Then cbo����.SetFocus: Exit Sub
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ���������壡", vbInformation, gstrSysName
        If CanFocus(cbo����) = True Then cbo����.SetFocus: Exit Sub
    End If
    If cbo��������.ListIndex = -1 Then
        MsgBox "����ȷ���������ͣ�", vbInformation, gstrSysName
        If cbo��������.Enabled Then
            cbo��������.SetFocus
        End If
        Exit Sub
    End If
    
    '������Ϣ���
    If txt������.Visible And txt������.Enabled Then
        If Not Check������Ϣ Then Exit Sub
    End If
    
    If cbo��Ժ����.ListIndex = -1 Then
        MsgBox "����ȷ��������Ժ���ң�", vbInformation, gstrSysName
        If CanFocus(cbo��Ժ����) Then cbo��Ժ����.SetFocus: Exit Sub
    End If
    If cbo��Ժ����.ListIndex = -1 And cbo��Ժ����.Visible And gbln��ѡ���� Then
        MsgBox "����ȷ��������Ժ������", vbInformation, gstrSysName
        If CanFocus(cbo��Ժ����) Then cbo��Ժ����.SetFocus: Exit Sub
    End If
    
    If mbln�մ� Then
        MsgBox zlCommFun.GetNeedName(cbo��Ժ����.Text) & "��û�пմ�λ��������������һ�תԺ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If cbo����ȼ�.ListIndex = -1 Then
        MsgBox "����ȷ������ȼ���", vbInformation, gstrSysName
        cbo����ȼ�.SetFocus: Exit Sub
    End If
    
    If cbo����ҽʦ.ListIndex = -1 And mbytMode <> 1 Then
        MsgBox "����ȷ������ҽʦ��", vbInformation, gstrSysName
        cbo����ҽʦ.SetFocus: Exit Sub
    End If
    
    If cbo��Ժ����.ListIndex = -1 Then
        MsgBox "����ȷ��������Ժ������", vbInformation, gstrSysName
        cbo��Ժ����.SetFocus: Exit Sub
    End If
    If cbo��Ժ��ʽ.ListIndex = -1 Then
        MsgBox "����ȷ��������Ժ��ʽ��", vbInformation, gstrSysName
        cbo��Ժ��ʽ.SetFocus: Exit Sub
    End If
    '���˺�:2007/09/13
    If cbo��Ժ����.ListIndex = -1 Then
        MsgBox "����ȷ��������Ժ���ԣ�", vbInformation, gstrSysName
        cbo��Ժ����.SetFocus: Exit Sub
    End If
    
    If cboסԺĿ��.ListIndex = -1 Then
        MsgBox "����ȷ������סԺĿ�ģ�", vbInformation, gstrSysName
        cboסԺĿ��.SetFocus: Exit Sub
    End If
    If Not IsDate(txt��Ժʱ��.Text) Then
        MsgBox "����������ȷ�Ĳ�����Ժʱ�䣡", vbInformation, gstrSysName
        txt��Ժʱ��.SetFocus: Exit Sub
    End If
    
     '��ϵ�˼��
    If Trim(txt��ϵ������.Text) = "" And (cbo��ϵ�˹�ϵ.ListIndex >= 0 Or Trim(txt��ϵ�˵绰.Text) <> "" Or Trim(txt��ϵ�˵�ַ.Text) <> "" Or Trim(txt��ϵ�����֤��.Text) <> "") Then
        MsgBox "����¼����ϵ������!", vbInformation, gstrSysName
        If txt��ϵ������.Enabled And txt��ϵ������.Visible Then txt��ϵ������.SetFocus: Exit Sub
    End If
    
    '78877,84014 �������ں���Ժʱ��ǰ���Ѿ����п�ֵ���
    If txt����ʱ��.Text = "__:__" Then
        str�������� = Format(txt��������.Text, "YYYY-MM-DD")
    Else
        str�������� = Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS")
    End If
    
    If txt��������.Enabled Then
        If CDate(str��������) > CDate(txt��Ժʱ��.Text) Then
            MsgBox "���˳�������[" & str�������� & "]����С�ڲ�����Ժʱ��[" & Format(txt��Ժʱ��.Text, "YYYY-MM-DD HH:MM") & "]��", vbInformation, gstrSysName
            txt��������.SetFocus: Exit Sub
        End If
    End If
    
    '�ѱ����ÿ���
    If cbo��Ժ����.ListIndex <> -1 Then
        If Not Check�ѱ����ÿ���(zlCommFun.GetNeedName(cbo�ѱ�.Text), Val(cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex))) Then
            MsgBox "��ǰ�ѱ�Բ��˿��Ҳ�����,������ѡ��ѱ�!", vbInformation, gstrSysName
            cbo�ѱ�.SetFocus: Exit Sub
        End If
    End If

    
    '��Ժʱ��
    If Not mrsInfo Is Nothing Then
        If CDate(txt��Ժʱ��.Text) < IIf(IsNull(mrsInfo!��Ժʱ��), #1/1/1900#, mrsInfo!��Ժʱ��) Then
            MsgBox "������Ժʱ�䲻��С�ڲ����ϴγ�Ժʱ��[" & Format(IIf(IsNull(mrsInfo!��Ժʱ��), #1/1/1900#, mrsInfo!��Ժʱ��), "yyyy-MM-dd") & "]��", vbInformation, gstrSysName
            txt��Ժʱ��.SetFocus: Exit Sub
        End If
    ElseIf mbytInState = EState.E�޸� And txt��Ժʱ��.Tag <> "" Then
        If CDate(txt��Ժʱ��.Text) < CDate(txt��Ժʱ��.Tag) Then
            MsgBox "������Ժʱ�䲻��С�ڲ����ϴγ�Ժʱ��[" & Format(txt��Ժʱ��.Tag, "yyyy-MM-dd HH:mm:ss") & "]��", vbInformation, gstrSysName
            txt��Ժʱ��.SetFocus: Exit Sub
        End If
    End If
        
    '�������
    If mintInsure <> 0 And mstrYBPati <> "" And mbytMode <> 1 Then
        If gclsInsure.GetCapability(support����¼��������, Val(txtPatient.Tag), mintInsure) Then
            If txt�������.Text = "" Then
                MsgBox "����д�ò��˵�������ϣ�", vbInformation, gstrSysName
                txt�������.SetFocus: Exit Sub
            End If
        End If
    ElseIf InStr(mstrPrivs, "�����ҽ������") = 0 Then
        MsgBox "��û��Ȩ�޶Է�ҽ�����˽��еǼ�.", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�ֻ��źϷ��Լ��
    If Trim(txtMobile.Text) <> "" Then
        If CheckMobile(Trim(txtMobile.Text), Val(txtPatient.Tag)) Then
            If MsgBox("�����еĲ�����Ϣ�д�����ͬ���ֻ���:" & Trim(txtMobile.Text) & vbCrLf & "�Ƿ�����¼�룿", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                If txtMobile.Enabled And txtMobile.Visible Then txtMobile.SetFocus: Exit Sub
            End If
        End If
    End If
    
    '���ȼ��
    
    If Not CheckTextLength("����", txt����) Then Exit Sub
    If Not CheckTextLength("����", txt����) Then Exit Sub
    If Not CheckOldData(txt����, cbo���䵥λ) Then Exit Sub
    
    '64701:������,2013-10-31,�޸ĳ�����ַ��������100���ַ���50������
    If Not gbln���ýṹ����ַ Then
        If Not CheckLen(txt��ͥ��ַ, 100) Then Exit Sub
        If Not CheckLen(txt�����ص�, 100) Then Exit Sub
        If Not CheckLen(txt���ڵ�ַ, 100) Then Exit Sub
        If Not CheckLen(txt��ϵ�˵�ַ, 100) Then Exit Sub
    End If
    If Not CheckLen(txt���ڵ�ַ�ʱ�, 6) Then Exit Sub
    If Not CheckLen(txt��ͥ��ַ�ʱ�, 6) Then Exit Sub
    If Not CheckLen(txt��ͥ�绰, 20) Then Exit Sub
    If Not CheckLen(txt��ϵ������, 64) Then Exit Sub
    If Not CheckLen(txt��ϵ�˵绰, 20) Then Exit Sub
    If Not CheckLen(txt��ϵ�����֤��, 18) Then Exit Sub
    If Not CheckLen(txtLinkManInfo, 100) Then Exit Sub
    If Not CheckLen(txt������λ, txt������λ.MaxLength) Then Exit Sub
    If Not CheckLen(txt��λ�绰, 20) Then Exit Sub
    If Not CheckLen(txtMobile, 20) Then Exit Sub
    If Not CheckLen(txt��λ�ʱ�, 6) Then Exit Sub
    If Not CheckLen(txt��λ������, 50) Then Exit Sub
    If Not CheckLen(txt��λ�ʺ�, 50) Then Exit Sub
    If Not CheckLen(txt������, 64) Then Exit Sub
    If Not CheckLen(txt�������, txt�������.MaxLength) Then Exit Sub
    If Not CheckLen(txt��ҽ���, txt��ҽ���.MaxLength) Then Exit Sub
    If Not CheckLen(txt����, CInt(mCurSendCard.lng���ų���)) Then Exit Sub
    If Not CheckLen(txtPass, 10) Then Exit Sub
    If Not CheckLen(txt�ɿλ, 50) Then Exit Sub
    If Not CheckLen(txt������, 50) Then Exit Sub
    If Not CheckLen(txt�ʺ�, 50) Then Exit Sub
    If Not CheckLen(txt�������, 30) Then Exit Sub
    If Not CheckLen(txt��ע, txt��ע.MaxLength) Then Exit Sub
    If zlStr.NeedName(cbo��Ժ��ʽ.Text) = "ת��" Then
        If Not zlControl.TxtCheckInput(txtת��, "ת��", 100) Then Exit Sub
    End If
    
    '104238:���ϴ���2017/2/15����鿨���Ƿ����㷢����������
    If txt����.Text <> "" And Len(txt����.Text) <> mCurSendCard.lng���ų��� And Not mCurSendCard.bln�ϸ���� Then
        Select Case mCurSendCard.byt��������
            Case 0
                MsgBox "����Ŀ���С��" & mCurSendCard.str������ & "�趨�Ŀ��ų��ȣ����������룡", vbExclamation, gstrSysName
                If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                Exit Sub
            Case 2
                If MsgBox("����Ŀ���С��" & mCurSendCard.str������ & "�趨�Ŀ��ų��ȣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt����.Visible And txt����.Enabled Then txt����.SetFocus
                    Exit Sub
                End If
        End Select
    End If
    
    '�����ӱ���(����/�޸�)
    mstrPatiPlus = ""
    'ת���������
    mstrPatiPlus = mstrPatiPlus & "," & "��Ժת��:" & Trim(zlStr.NeedName(txtת��.Text))
    '��ϵ�˹�ϵ������ϵ����˵��
    mstrPatiPlus = mstrPatiPlus & "," & "��ϵ�˸�����Ϣ:" & Trim(txtLinkManInfo.Text)
    '���֤��
    If Trim(zlCommFun.GetNeedName(cbo����.Text)) = "�й�" Then
        mstrPatiPlus = mstrPatiPlus & "," & "���֤��״̬:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
        mstrPatiPlus = mstrPatiPlus & "," & "�⼮���֤��:"
    Else
        If txt���֤��.Text <> "" Then
            mstrPatiPlus = mstrPatiPlus & "," & "�⼮���֤��:" & txt���֤��.Text
            mstrPatiPlus = mstrPatiPlus & "," & "���֤��״̬:"
            txt���֤��.Text = ""
        Else
            mstrPatiPlus = mstrPatiPlus & "," & "���֤��״̬:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
            mstrPatiPlus = mstrPatiPlus & "," & "�⼮���֤��:"
        End If
    End If
    If mstrPatiPlus <> "" Then mstrPatiPlus = Mid(mstrPatiPlus, 2)
    
    'ԤԼ���Ĳ��˴��ż��
    If mblnAppoint And mstrAppointBed <> cbo��λ.Tag And gbln��Ժ��� And mbytMode = EMode.E����ԤԼ Then
        If MsgBox("ԤԼ��λ��" & mstrAppointBed & "���뵱ǰ��λ��" & cbo��λ.Tag & "������ͬ���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If CanFocus(cbo��λ) Then cbo��λ.SetFocus
            Exit Sub
        End If
    End If
    '��۽����
    '���˺�:    '29134
    '82401:���ϴ�,2015/3/11,�������Ƿ����
    If mbytInState = 0 And pic�ſ�.Visible And txt����.Text <> "" Then
        If tabCardMode.SelectedItem.Key = "CardFee" And Not mCurSendCard.rs���� Is Nothing Then
            If mCurSendCard.rs����!�Ƿ��� Then
                If mCurSendCard.rs����!�ּ� <> 0 And Abs(CCur(txt����.Text)) > Abs(mCurSendCard.rs����!�ּ�) Then
                    MsgBox "" & mCurSendCard.str������ & "������ֵ���ܴ�������޼ۣ�" & Format(Abs(mCurSendCard.rs����!�ּ�), "0.00"), vbInformation, gstrSysName
                    txt����.SetFocus: Exit Sub
                End If
                If mCurSendCard.rs����!ԭ�� <> 0 And Abs(CCur(txt����.Text)) < Abs(mCurSendCard.rs����!ԭ��) Then
                    MsgBox "" & mCurSendCard.str������ & "������ֵ����С������޼ۣ�" & Format(Abs(mCurSendCard.rs����!ԭ��), "0.00"), vbInformation, gstrSysName
                    txt����.SetFocus: Exit Sub
                End If
            End If
        End If
    End If
    
    If pic�ſ�.Visible And txt����.Text <> "" Then
        Select Case mCurSendCard.int���볤������
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCurSendCard.int���볤�� Then
                MsgBox "ע��:" & vbCrLf & "�����������" & mCurSendCard.int���볤�� & "λ", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCurSendCard.int���볤������) Then
                MsgBox "ע��:" & vbCrLf & "�����������" & Abs(mCurSendCard.int���볤������) & "λ����.", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        End Select
    
        If txtPass.Text <> txtAudi.Text Then
            MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
            txtPass.Text = "": txtAudi.Text = ""
            txtPass.SetFocus: Exit Sub
        End If
        
    End If
    
    '���㷽ʽ
    If IsNumeric(txtԤ����.Text) And cboԤ������.Visible And cboԤ������.Enabled And cboԤ������.ListIndex = -1 Then
        MsgBox "��ȷ������Ԥ������㷽ʽ��", vbInformation, gstrSysName
        cboԤ������.SetFocus: Exit Sub
    End If
    If Trim(txt����.Text) <> "" And cbo��������.Visible And cbo��������.Enabled And cbo��������.ListIndex = -1 Then
        MsgBox "��ȷ������" & mCurSendCard.str������ & "���㷽ʽ��", vbInformation, gstrSysName
        cbo��������.SetFocus: Exit Sub
    End If
    
    '63246:������,2013-07-03
    If CheckPatiCard = False Then Exit Sub
    
    If mbytInState = 0 Then
        'ҽ���Ķ�
        If mintInsure <> 0 And mstrYBPati <> "" And mbytMode <> 1 Then
            If is�����ʻ�(cboԤ������) Then
                If IsNumeric(txtԤ����.Text) Then
                    If CCur(StrToNum(txtԤ����.Text)) > mcurYBMoney Then
                        MsgBox "ҽ�������ʻ�ת����ܴ������:" & Format(mcurYBMoney, "0.00"), vbInformation, gstrSysName
                        txtԤ����.SetFocus: Exit Sub
                    End If
                End If
            End If
        ElseIf mstrYBPati = "" And IsNumeric(txtԤ����.Text) And is�����ʻ�(cboԤ������) Then
            MsgBox "��ҽ�����˲���ʹ�ø����ʻ����ʣ�", vbInformation, gstrSysName
            cboԤ������.SetFocus: Exit Sub
        End If
    
        'Ʊ����ؼ��
        mblnPrepayPrint = False
        If IsNumeric(txtԤ����.Text) Then
        'If zlSquareSimulation(lng�ӿڱ��, strBalanceInfor) = False Then Exit Sub
        
            mblnPrepayPrint = True
            '����Ƿ��ӡƱ��
            If gbytPrepayPrint = 0 Then
                mblnPrepayPrint = False
            Else
                If gbytPrepayPrint = 2 Then
                    If MsgBox("�Ƿ��ӡԤ����Ʊ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        mblnPrepayPrint = False
                    End If
                End If
            End If
            
            If mblnPrepayPrint Then
                If gblnPrepayStrict Then
                    If Trim(txtFact.Text) = "" Then
                        MsgBox "��������һ����Ч��Ԥ��Ʊ�ݺ��룡", vbInformation, gstrSysName
                        txtFact.SetFocus: Exit Sub
                    End If
                    mlngԤ������ID = CheckUsedBill(2, IIf(mlngԤ������ID > 0, mlngԤ������ID, mFactProperty.lngShareUseID), txtFact.Text, 2)
                    If mlngԤ������ID <= 0 Then
                        Select Case mlngԤ������ID
                            Case 0 '����ʧ��
                            Case -1
                                MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                            Case -2
                                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                            Case -3
                                MsgBox "Ʊ�ݺ��벻�ڵ�ǰ��Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                                txtFact.SetFocus
                        End Select
                        Exit Sub
                    End If
                Else
                    If Len(txtFact.Text) <> gbytPrepayLen And txtFact.Text <> "" Then
                        MsgBox "Ԥ��Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytPrepayLen & " λ��", vbInformation, gstrSysName
                        txtFact.SetFocus: Exit Sub
                    End If
                End If
            End If
        End If
        
        If txt����.Text <> "" And pic�ſ�.Visible Then
            '����ǰ�����￨�Ƿ��У��Ƿ��ڷ�Χ��
            If mCurSendCard.bln�ϸ���� Then
                mCurSendCard.lng����ID = CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), txt����.Text, mCurSendCard.lng�����ID)
                If mCurSendCard.bln���￨ Then
                    blnErr = mCurSendCard.lng����ID <= 0 And Not mCurSendCard.blnOneCard
                Else
                    blnErr = mCurSendCard.lng����ID <= 0
                End If
                If blnErr Then
                    Select Case mCurSendCard.lng����ID
                        Case 0 '����ʧ��
                        Case -1
                            MsgBox "����û�����ü����õ�" & mCurSendCard.str������ & ",�����ڱ������ù������λ�����һ��" & mCurSendCard.str������ & "��", vbExclamation, gstrSysName
                        Case -2
                            MsgBox "���ع��õ�" & mCurSendCard.str������ & "������,���������ñ��ع���" & mCurSendCard.str������ & "���λ�����һ��" & mCurSendCard.str������ & "��", vbExclamation, gstrSysName
                        Case -3
                            MsgBox "����" & mCurSendCard.str������ & "�Ų�����Ч��Χ��,�����Ƿ���ȷˢ����", vbExclamation, gstrSysName
                            txt����.SetFocus
                    End Select
                    Exit Sub
                End If
            End If
        End If
        
        If mrsInfo Is Nothing Then
            '65689:������,2013-10-30,���ڶ����ͬ���ˣ��ṩѡ����������Աѡ��
            If Not (mbytMode = EMode.E����ԤԼ Or mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0) Then
                '������Ʋ�����Ϣ(����֮ǰ���,����������ظ���Ϣ������)
                Set rsSimilar = SimilarIDs(zlCommFun.GetNeedName(cbo����.Text), zlCommFun.GetNeedName(cbo����), CDate(IIf(IsDate(txt��������.Text), txt��������.Text, #1/1/1900#)), zlCommFun.GetNeedName(cbo�Ա�), txt����.Text, txt���֤��.Text)
                If Not rsSimilar Is Nothing Then
                    If gblnPatiByID And Trim(txt���֤��.Text) <> "" Then
                        '110541 ͬһ���ֻ֤�ܶ�Ӧһ����������;���øò�����ͨ�����֤���ҵ��ѽ�������ʱ����ѡ���
                        rsSimilar.Filter = "���֤�� ='" & Trim(txt���֤��.Text) & "'"
                        If rsSimilar.RecordCount > 0 Then
                            strSimilar = "�����еĲ�����Ϣ�з���" & rsSimilar.RecordCount & "�����֤����ͬ�ĵĲ��ˡ�" & vbCrLf & vbCrLf & _
                                "��ȡ���еĲ�����Ϣ��ѡ���˺�[˫��]����[ȷ��]��"
                            If Not CreatePublicPatient() Then Exit Sub
                            If gobjPublicPatient.ShowSelect(rsSimilar, "ID", "����ѡ��", strSimilar, , , "0|800|1200|800|800|1500|1000", True) Then
                                txtPatient.Text = "-" & rsSimilar!����ID
                                txtPatient.SetFocus
                                Call txtPatient_KeyPress(13)
                                Exit Sub
                            End If
                        End If
                    End If
                    rsSimilar.Filter = ""
                    If rsSimilar.RecordCount > 1 Then
                        strSimilar = "�����еĲ�����Ϣ�з���" & rsSimilar.RecordCount & "����Ϣ���ƵĲ���(����,����,�Ա�,����,����������ͬ�����֤����ͬ)" & vbCrLf & vbCrLf & _
                            "��ȡ���еĲ�����Ϣ��ѡ���˺�[˫��]����[ȷ��],�Ǽ�Ϊ�²�������[ȡ��]"
                        If Not CreatePublicPatient() Then Exit Sub
                        blnOk = gobjPublicPatient.ShowSelect(rsSimilar, "ID", "����ѡ��", strSimilar, , , "0|800|1200|800|800|1500|1000")
                        If blnOk = True Then
                            txtPatient.Text = "-" & rsSimilar!����ID
                            txtPatient.SetFocus
                            Call txtPatient_KeyPress(13)
                            Exit Sub
                        Else
                            MsgBox "�ò��˵����Ƽ�¼�����ڲ�����Ϣ������ʹ��""�ϲ�""���ܴ���", vbInformation, gstrSysName
                        End If
                    ElseIf rsSimilar.RecordCount = 1 Then
                        strSimilar = "ID:" & rsSimilar!����ID & ",�����:" & Nvl(rsSimilar!�����, "��") & ",סԺ��:" & Nvl(rsSimilar!סԺ��, "��") & ",���֤��:" & rsSimilar!���֤�� & ",��ַ:" & rsSimilar!��ַ & ",�Ǽ�����:" & rsSimilar!�Ǽ�ʱ��
                        If MsgBox("�����еĲ�����Ϣ�з��� 1 ����Ϣ���ƵĲ���(����,����,�Ա�,����,����������ͬ�����֤����ͬ): " & vbCrLf & vbCrLf & _
                            strSimilar & vbCrLf & vbCrLf & "�Ǽ�Ϊ�²�����ѡ��[��],��ȡ���еĲ�����Ϣ��ѡ��[��]��", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            txtPatient.Text = "-" & Mid(Split(strSimilar, ",")(0), 4)
                            txtPatient.SetFocus
                            Call txtPatient_KeyPress(13)
                            Exit Sub
                        Else
                            MsgBox "�ò��˵����Ƽ�¼�����ڲ�����Ϣ������ʹ��""�ϲ�""���ܴ���", vbInformation, gstrSysName
                        End If
                    End If
                End If
                
                '����ID���:�Զ��滻�µ�
                Do While ExistInPatiID(CLng(txtPatient.Tag))
                    txtPatient.Text = zlDatabase.GetNextNo(1)
                    txtPatient.Tag = txtPatient.Text
                Loop
            End If
        End If
        
        If txtסԺ��.Visible And (mbytKind = EסԺ��Ժ�Ǽ�) Then
            If mrsInfo Is Nothing Then
                lng����ID = IIf(mbytMode = EMode.E����ԤԼ Or mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0, Val(txtPatient.Tag), 0)
            Else
                lng����ID = mrsInfo!����ID
            End If
            '����29449 by lesfeng 2010-05-05
            Dim blnTrue As Boolean
            blnTrue = False
            If mbytMode = EMode.EԤԼ�Ǽ� Then blnTrue = True
            '60500:������,2013-05-09
            If ExistInPatiNO(txtסԺ��.Text, lng����ID, blnTrue) Then
                strno = zlDatabase.GetNextNo(2)
                If Val(txtסԺ��.Text) = Val(strno) Then
                    MsgBox "��ǰסԺ�ź��Զ���ȡ����סԺ���ظ�,���ֹ��޸�סԺ�ţ�", vbInformation, gstrSysName
                Else
                    MsgBox "��ǰסԺ���ѱ�ʹ��,���Զ���ȡһ���µ�סԺ��,��ȷ�ϣ�", vbInformation, gstrSysName
                    txtסԺ��.Text = strno
                End If
                txtסԺ��.SetFocus: Exit Sub
            End If
        End If
        
        If txtסԺ��.Visible And mbytKind = E�������۵Ǽ� Then
            gstrSQL = "Select 1 From ������Ϣ Where �����=[1] And ����ID<>[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtסԺ��.Text, Val(txtPatient.Tag))
            If rsTmp.RecordCount > 0 Then
                If Not mblnAuto Then
                    MsgBox "��ǰ������ѱ�ʹ��,���Զ���ȡһ���µ������,��ȷ�ϣ�", vbInformation, gstrSysName
                    txtסԺ��.Text = zlDatabase.GetNextNo(3)
                    mblnAuto = True
                    txtסԺ��.SetFocus: Exit Sub
                Else
                    blnTmp = True
                    txtסԺ��.Text = Val(txtסԺ��.Text) + 1
                    mblnAuto = True
                End If
            End If
        End If
        
        If txtסԺ��.Visible And mbytKind = EסԺ���۵Ǽ� Then
            gstrSQL = "Select 1 From ������ҳ Where ���ۺ�=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtסԺ��.Text)
            If rsTmp.RecordCount > 0 Then
                MsgBox "��ǰ���ۺ��ѱ�ʹ��,���Զ���ȡһ���µ����ۺ�,��ȷ�ϣ�", vbInformation, gstrSysName
                txtסԺ��.Text = zlDatabase.GetNextNo(6)
                txtסԺ��.SetFocus: Exit Sub
            End If
        End If
        '�����:51072
        If Len(Trim(txtPass.Text)) <= 0 And Len(Trim(txt����.Text)) > 0 Then 'û����������
            If zl_Get����Ĭ�Ϸ������� = False Then Exit Sub
        End If
        
        If CheckBrushCard = False Then Exit Sub
        
        '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
        If IsCertificateCard(Val(txtPatient.Tag)) = False Then Exit Sub
        '
        '�����¼�¼(����|�޸Ĳ�����Ϣ����Ժ��¼��Ԥ����¼(IFҪ����)���ſ���¼(IFҪ����))
        cmdOK.Enabled = False
        If Not SavePatiNew(mrsInfo Is Nothing And mlng����ID = 0, lng�ӿڱ��, strBalanceInfor) Then
            cmdOK.Enabled = True: Exit Sub
        End If
        
        '�������۵Ǽ�ʱ��ʾ��Ϣ
        If blnTmp And mbytKind = E�������۵Ǽ� Then MsgBox "��ǰ������ѱ�ʹ�ã�ϵͳ�Զ�Ϊ���������µ�����š�" & txtסԺ��.Text & "��", vbInformation, gstrSysName
        gblnOK = True
        
        '��ӡԤ�����վ�
        If mblnPrepayPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mCurPrepay.strno, 2)
        End If
        
        '��ӡ������ҳ:ԤԼ�Ǽǲ���ӡ
        If InStr(mstrPrivs, "��ҳ��ӡ") > 0 Then
            If mbytMode <> 1 Then
                mblnFPagePrint = True
                If gbytFPagePrint = 0 Then
                    mblnFPagePrint = False
                Else
                    If gbytFPagePrint = 2 Then
                        If MsgBox("�Ƿ��ӡ������ҳ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            mblnFPagePrint = False
                        End If
                    End If
                End If
                
                If mblnFPagePrint Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131", Me, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, 2)
                End If
            End If
        End If
        
        '��ӡ�������
        If InStr(mstrPrivs, "�����ӡ") Then
            mblnWristletPrint = True
            If gbytWristletPrint = 0 Then
                mblnWristletPrint = False
            Else
                If gbytWristletPrint = 2 Then
                    If MsgBox("�Ƿ��ӡ���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        mblnWristletPrint = False
                    End If
                End If
            End If
            
            If mblnWristletPrint Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131_1", Me, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, 2)
            End If
        End If
        
        'Ʊ����ش���
        '�µ�һ��Ԥ�����
        If mblnPrepayPrint Then
            If gblnPrepayStrict Then
                If mbytMode <> EMode.E����ԤԼ Then '�ⲿ���ý���ʱ���ٲ����º�
                    mlngԤ������ID = CheckUsedBill(2, IIf(mlngԤ������ID > 0, mlngԤ������ID, mFactProperty.lngShareUseID), , 2)
                    If mlngԤ������ID <= 0 Then
                        Select Case mlngԤ������ID
                            Case 0 '����ʧ��
                            Case -1
                                MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                            Case -2
                                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                        End Select
                        txtFact.Text = ""
                    Else
                        '�ϸ�ȡ��һ������
                        txtFact.Text = GetNextBill(mlngԤ������ID)
                    End If
                End If
            Else
                '��ɢ��ȡ��һ������
                zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", txtFact.Text, glngSys, mlngModul
                txtFact.Text = zlCommFun.IncStr(txtFact.Text)
            End If
        End If
        If mbytMode <> EMode.E����ԤԼ And txt����.Text <> "" And pic�ſ�.Visible Then
            If mCurSendCard.bln�ϸ���� Then
                mCurSendCard.lng����ID = CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), , mCurSendCard.lng�����ID)
                If mCurSendCard.lng����ID <= 0 Then
                    Select Case mCurSendCard.lng����ID
                        Case 0 '����ʧ��
                        Case -1
                            MsgBox "����û�����ü����õ�" & mCurSendCard.str������ & ",�����ڱ������ù������λ�����һ��" & mCurSendCard.str������ & "��", vbExclamation, gstrSysName
                        Case -2
                            MsgBox "���ع��õ�" & mCurSendCard.str������ & "������,���������ñ��ع���" & mCurSendCard.str������ & "���λ�����һ��" & mCurSendCard.str������ & "��", vbExclamation, gstrSysName
                    End Select
                End If
            End If
        End If
                
        cmdOK.Enabled = True
        If mbytMode = EMode.E����ԤԼ Then
            '������˳�
            gblnOK = True: Unload Me: Exit Sub
        Else
            '����������һ��������Ϣ
            mblnICCard = False  '���ܷ���clearcard��,��Ϊ�����ȶ����ٲ������
            Call ClearCard
            If Not mCurSendCard.rs���� Is Nothing Then
                txt����.Text = Format(IIf(mCurSendCard.rs����!�Ƿ��� = 1, mCurSendCard.rs����!ȱʡ�۸�, mCurSendCard.rs����!�ּ�), "0.00")
            End If
            
            txtPatient.SetFocus
        End If
    ElseIf mbytInState = 1 Then
        'סԺ�ż��
        If txtסԺ��.Visible And mbytKind = EסԺ��Ժ�Ǽ� And txtסԺ��.Text <> txtסԺ��.Tag Then
            If ExistInPatiNO(txtסԺ��.Text, mlng����ID, True) Then
                MsgBox "��ǰסԺ���ѱ�ʹ��,���Զ���ȡһ���µ�סԺ��,��ȷ�ϣ�", vbInformation, gstrSysName
                txtסԺ��.Text = zlDatabase.GetNextNo(2)
                txtסԺ��.SetFocus: Exit Sub
            End If
        End If
        
        '����ż��
        If txtסԺ��.Visible And mbytKind = E�������۵Ǽ� And txtסԺ��.Text <> txtסԺ��.Tag Then
            gstrSQL = "Select 1 From ������Ϣ Where �����=[1] And ����ID<>[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtסԺ��.Text, mlng����ID)
            If rsTmp.RecordCount > 0 And Not mblnAuto Then
                MsgBox "��ǰ������ѱ�ʹ��,���Զ���ȡһ���µ������,��ȷ�ϣ�", vbInformation, gstrSysName
                txtסԺ��.Text = zlDatabase.GetNextNo(3)
                mblnAuto = True
                
                txtסԺ��.SetFocus: Exit Sub
            End If
        End If
        
        '����ż��
        If txtסԺ��.Visible And mbytKind = EסԺ���۵Ǽ� And txtסԺ��.Text <> txtסԺ��.Tag Then
            gstrSQL = "Select 1 From ������ҳ Where ���ۺ�=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtסԺ��.Text)
            If rsTmp.RecordCount > 0 Then
                MsgBox "��ǰ���ۺ��ѱ�ʹ��,���Զ���ȡһ���µ����ۺ�,��ȷ�ϣ�", vbInformation, gstrSysName
                txtסԺ��.Text = zlDatabase.GetNextNo(6)
                txtסԺ��.SetFocus: Exit Sub
            End If
        End If
        
        '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
        If IsCertificateCard(mlng����ID) = False Then Exit Sub
        
        '�����޸�(��Ժ��¼)
        cmdOK.Enabled = False
        If Not SavePatiModi Then
            cmdOK.Enabled = True: Exit Sub
        Else
            '������Ϣ����ɹ���,ͬ���޸Ĳ��˻�����Ϣ
            If bln������Ϣ���� And blnMod Then
                strErrInfo = ""
                Call gobjPublicPatient.SavePatiBaseInfo(mlng����ID, mlng��ҳID, Trim(txt����.Text), str�Ա�, strAge, str��������, Me.Caption, IIf(mlng��ҳID <> 0, 2, 1), strErrInfo, True, True)
                If strErrInfo <> "" Then
                    MsgBox strErrInfo, vbInformation + vbOKOnly, Me.Caption
                End If
            End If
        End If
        cmdOK.Enabled = True
        gblnOK = True: Unload Me: Exit Sub
    End If
End Sub

Private Function SavePatiNew(bln�²��� As Boolean, ByVal lng���㿨�ӿ� As Long, ByVal strBalancelInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ������µĲ�����Ժ�Ǽ�(��������Ϣ����Ժ��Ϣ��Ԥ������￨)
    '��Σ�lng���㿨�ӿ�-���㿨�ӿڱ��(0-��ʾ��ͨ����)
    '         strBalancelInfor-ģ�����������Ϣ
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-07-09 17:21:15
    '˵����
    '----------------------------------------------------------------------------------------------------------------------
    Dim strPati As String, strDeposit As String, strSQLCard As String, bytMode As Byte
    Dim strSurety As String, str������ As String, str����ʱ�� As String
    Dim strInsure As String, lng����ID As Long
    Dim lng����ID As Long, lng��ҳID As Long, lng����ID As Long, lng����ID As Long, lngԤ��ID As Long, lng�䶯ID As Long
    Dim strCard As String, strICCard As String, strno As String, strDepositNO As String, strSQL As String, blnTrans As Boolean, blnInRange As Boolean
    Dim lng��ҽ����ID As Long, lng��ҽ����ID As Long
    Dim lng��ҽ���ID As Long, lng��ҽ���ID As Long
    Dim str�������� As String, str���� As String
    Dim str���� As String, str��λ�ȼ� As String
    Dim str����� As String, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim blnNotCommit As Boolean
    Dim bln����תסԺ As Boolean '38069
    Dim bln�����ʻ���Ԥ�� As Boolean    '38069
    Dim cllUpdate As Collection, cllThreeInsert As Collection, cllPro As Collection, cll������ As Collection
    Dim Curdate As Date
    Dim lngInHosTimes  As Long
    Dim i As Long, lngRet As Long
    Dim arrTmp  As Variant
    Dim arrSQL As Variant
    Dim strErr As String
    
    arrSQL = Array()
    
    If cbo��Ժ����.Visible And cbo��Ժ����.ListIndex <> -1 Then
        lng����ID = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    End If
    If cbo��Ժ����.ListIndex <> -1 Then
        lng����ID = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    End If
        
    If cbo��λ.Visible And cbo��λ.ListIndex > 0 Then       '0-���ִ�,1-��ͥ����
        If cbo��λ.ListIndex = 1 Then
            str���� = "��ͥ����"
        Else
            str���� = Trim(Mid(Trim(cbo��λ.Text), 1, InStr(Trim(cbo��λ.Text), " ") - 1))
            If InStr(Trim(cbo��λ.Text), " ����") <> 0 Then
                If InStr(Trim(cbo��λ.Text), "|") - InStr(Trim(cbo��λ.Text), "����:") - 3 > 0 Then
                    str����� = Mid(Trim(cbo��λ.Text), InStr(Trim(cbo��λ.Text), "����:") + 3, InStr(Trim(cbo��λ.Text), "|") - InStr(Trim(cbo��λ.Text), "����:") - 3)
                End If
                strSQL = "Select �Ա� From ������Ϣ A,��λ״����¼ B  Where A.����ID = b.����id And b.����ID Is Not Null And ����ID = [1] And ����� =[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, str�����)
                
                Do While Not rsTmp.EOF
                    If Mid(Trim(cbo�Ա�.Text), 3) <> rsTmp!�Ա� Then
                        If (MsgBox("ָ����λ���ڷ��������Ů��ס������Ƿ������ס��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
                            Exit Do
                        Else
                            Exit Function
                            If CanFocus(cbo��λ) Then cbo��λ.SetFocus
                        End If
                    End If
                    rsTmp.MoveNext
                Loop
            End If
        End If
    Else
        str���� = "-1"    'תΪ��
    End If
    If cbo����ȼ�.ListIndex <> -1 Then lng����ID = cbo����ȼ�.ItemData(cbo����ȼ�.ListIndex) '���û��ѡ,��Ϊ0,�洢�����лᴦ��Ϊ��
    
    If InStr(1, txt�������.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt�������.Tag)
    Else
        lng��ҽ���ID = Val(txt�������.Tag)
    End If
    If InStr(1, txt��ҽ���.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt��ҽ���.Tag)
    Else
        lng��ҽ���ID = Val(txt��ҽ���.Tag)
    End If
    
    str������ = Replace(Trim(txt������.Text), "'", "''")
    lng����ID = Val(txtPatient.Tag)
    
    lngInHosTimes = Val(txtTimes.Text)
    If mbytMode = EMode.EԤԼ�Ǽ� Then
        lng��ҳID = 0
    Else
        lng��ҳID = IIf(lngInHosTimes > Val("" & txtPages.Text), lngInHosTimes, Val("" & txtPages.Text))
    End If
    
    If mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0 Then
        bytMode = 2
    Else
        bytMode = mbytMode
    End If
    
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    '102232�½��������������Fȡ������
    If bln�²��� Then
        If txt����ʱ�� = "__:__" Then
            str�������� = IIf(IsDate(txt��������.Text), Format(txt��������.Text, "YYYY-MM-DD HH:MM:SS"), "")
        Else
            str�������� = IIf(IsDate(txt��������.Text), Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS"), "")
        End If
        strSQL = "<XM>" & Trim(txt����.Text) & "</XM><XB>" & zlCommFun.GetNeedName(cbo�Ա�.Text) & "</XB><NL>" & str���� & "</NL>" & vbNewLine & _
                "<CSRQ>" & str�������� & "</CSRQ><YBH>" & txtҽ����.Text & "</YBH><SFZH>" & txt���֤��.Text & "</SFZH>"
        If Not FuncPlugPovertyInfo(0, strSQL) Then Exit Function
    End If
    
    If txt����ʱ�� = "__:__" Then
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If

    strCard = UCase(txt����.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    bln����תסԺ = False: bln�����ʻ���Ԥ�� = False
    If (mintInsure <> 0 Or InStr(1, mstrPrivs, ";�������תסԺ;") > 0) And mstrNOS <> "" Then
        bln����תסԺ = True
    End If
    
    Curdate = zlDatabase.Currentdate

    strPati = "zl_��Ժ������ҳ_Insert(" & _
        bytMode & "," & mbytKind & "," & lng����ID & "," & IIf(txtסԺ��.Visible And txtסԺ��.Text <> "", txtסԺ��.Text, "NULL") & "," & _
        "'" & txtҽ����.Text & "','" & txt����.Text & "','" & zlCommFun.GetNeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
        "'" & zlCommFun.GetNeedName(cbo�ѱ�.Text) & "'," & str�������� & "," & _
        "'" & zlCommFun.GetNeedName(cbo����.Text) & "','" & zlCommFun.GetNeedName(cbo����.Text) & "','" & zlCommFun.GetNeedName(cboѧ��.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cbo����״��.Text) & "','" & zlCommFun.GetNeedName(cboְҵ.Text) & "','" & zlCommFun.GetNeedName(cbo���.Text) & "'," & _
        "'" & txt���֤��.Text & "','" & txt�����ص�.Text & "','" & txt��ͥ��ַ.Text & "','" & txt��ͥ��ַ�ʱ�.Text & "'," & _
        "'" & txt��ͥ�绰.Text & "','" & txt���ڵ�ַ.Text & "','" & txt���ڵ�ַ�ʱ�.Text & "','" & txt��ϵ������.Text & "','" & zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text) & "'," & _
        "'" & txt��ϵ�˵�ַ.Text & "','" & txt��ϵ�˵绰.Text & "','" & txt������λ.Text & "'," & Val(txt������λ.Tag) & "," & _
        "'" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "','" & txt��λ������.Text & "','" & txt��λ�ʺ�.Text & "'," & _
        "'" & str������ & "'," & Val(txt������.Text) & "," & IIf(str������ = "", "null", chk��ʱ����.Value) & "," & _
        ZVal(lng����ID) & "," & lng����ID & ",'" & zlCommFun.GetNeedName(cbo��Ժ����.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cbo��Ժ��ʽ.Text) & "','" & zlCommFun.GetNeedName(cboסԺĿ��.Text) & "'," & chk����Ժת��.Value & "," & _
        "'" & zlCommFun.GetNeedName(cbo����ҽʦ.Text) & "','" & zlCommFun.GetNeedName(txt����.Text) & "','" & zlCommFun.GetNeedName(txt����.Text) & "'," & _
        "To_Date('" & Format(txt��Ժʱ��.Text, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
        chk���.Value & "," & IIf(str���� = "-1", "NULL", "'" & str���� & "'") & ",'" & zlCommFun.GetNeedName(Replace(cboҽ�Ƹ���.Text, Chr(&HD), "")) & "'," & _
        ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt�������.Text, "'", "''") & "'," & _
        ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt��ҽ���.Text, "'", "''") & "'," & _
        IIf(mintInsure <> 0 And mstrYBPati <> "" And bln����תסԺ = False, mintInsure, "NULL") & ",'" & UserInfo.��� & "'," & _
        "'" & UserInfo.���� & "'," & IIf(bln�²���, 1, 0) & ",'" & txt��ע.Text & "'," & _
        ZVal(lng����ID) & "," & chk����Ժ.Value & ",'" & zlCommFun.GetNeedName(cbo��Ժ����.Text) & "'," & lng��ҳID & "," & IIf(lngInHosTimes = 0, "NULL", lngInHosTimes) & ",'" & _
        Trim(txt����֤��.Text) & "','" & zlCommFun.GetNeedName(cbo��������.Text) & "','" & txt��ϵ�����֤��.Text & "','" & Trim(txtMobile.Text) & "')"
    '������ҳ�ӱ���Ϣ����
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",��ϵ�˸�����Ϣ,��Ժת��,���֤��״̬,�⼮���֤��,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_������ҳ�ӱ�_��ҳ����(" & lng����ID & "," & lng��ҳID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "')"
            End If
            If InStr(",��ϵ�˸�����Ϣ,���֤��״̬,�⼮���֤��,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & lng����ID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
            End If
        Next
    End If
    
    '���ؽṹ����ַSQL
    If gbln���ýṹ����ַ Then
        Call CreateStructAddressSQL(lng����ID, lng��ҳID, arrSQL, PatiAddress)
    End If
    
    '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
    If lng����ID > 0 Then Call AddCertificate(lng����ID, arrSQL, Curdate)

    'û��Ȩ�޻�ԤԼ�Ǽ�ʱ���ɼ�,���ز�������Ϊ�������ϢʱΪ����
    If txt������.Visible And txt������.Enabled And str������ <> "" Then
        str����ʱ�� = "null"
        If Not IsNull(dtp����ʱ��.Value) Then str����ʱ�� = "To_Date('" & Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSurety = "zl_���˵�����¼_insert(" & lng����ID & "," & lng��ҳID & ",'" & str������ & "'," & _
            Val(txt������.Text) & "," & chk��ʱ����.Value & ",'" & Trim(txtReason.Text) & "',NULL," & str����ʱ�� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    End If
    '69231,������,2014-01-07 14:42:55,����ʱǿ�Ƹ��¿���������
    Call SetCardVaribles(False)
    '���ӷ�����¼
    Call AddCardDataSQL(lng����ID, lng��ҳID, lng����ID, lng����ID, Curdate, strSQLCard)
    '�����:57326
    If mbln������󶨿� Then
        If Check��������(lng����ID, mCurSendCard.lng�����ID) = False Then
            txt����.Text = "": txtPass.Text = "": txtAudi.Text = "": txt����.Text = ""
            Exit Function
        End If
        '�����㷽ʽ��Ϣ�Ƿ�Ϸ�
        If cbo��������.ItemData(cbo��������.ListIndex) = 8 And mCurCardPay.lngҽ�ƿ����ID = 0 Then
            MsgBox "��ǰ�������㷽ʽ�����쳣���޷�ʹ�øý��㷽ʽ�������Ƿ�������Ӧ�豸�������Ա��ϵ!", vbInformation + vbOKOnly
            Exit Function
        End If
    End If
    
    '����Ԥ����¼
    Call AddDepositSQL(lng����ID, lng��ҳID, lng����ID, Curdate, bln�����ʻ���Ԥ��, strDeposit)
    '���Ԥ�����㷽ʽ��Ϣ�Ƿ�Ϸ�
    If IsNumeric(txtԤ����.Text) And fraԤ��.Visible Then
        If cboԤ������.ItemData(cboԤ������.ListIndex) = 8 And mCurPrepay.lngҽ�ƿ����ID = 0 Then
            MsgBox "��ǰԤ�����㷽ʽ�����쳣���޷�ʹ�øý��㷽ʽ�������Ƿ�������Ӧ�豸�������Ա��ϵ!", vbInformation + vbOKOnly
            Exit Function
        End If
    End If
    
    
    '��һ��:����HIS��Ժ�ǼǺ�Ԥ����
    '����:31635
    blnNotCommit = False
    On Error GoTo errH
    Set cllUpdate = New Collection
    Set cllThreeInsert = New Collection
    Set cllPro = New Collection
    Set cll������ = New Collection
      
    gcnOracle.BeginTrans: blnTrans = True
    '���˲�����Ϣ
    zlDatabase.ExecuteProcedure strPati, Me.Caption
    '������ҳ�ӱ���Ϣ\�ṹ����ַ
    For i = LBound(arrSQL) To UBound(arrSQL)
         zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    '������Ϣ
    If strSurety <> "" Then zlDatabase.ExecuteProcedure strSurety, Me.Caption
    '��Ժ����
    If strSQLCard <> "" Then zlDatabase.ExecuteProcedure strSQLCard, Me.Caption
    '�����֤
    If txt֧������.Visible = True And txt֧������.Text <> "" Then
        If zl�����֤(cllPro) = False Then Exit Function
        zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
    End If
    '�����:56599
    '���벡�˽�������Ϣ
    If lng����ID > 0 Then Call Add�����������Ϣ(lng����ID, cll������)
    zlExecuteProcedureArrAy cll������, Me.Caption, True, True
    
    '��ԺԤ����
    If strDeposit <> "" And (bln����תסԺ = False Or bln�����ʻ���Ԥ�� = False) Then zlDatabase.ExecuteProcedure strDeposit, Me.Caption
    '��Ժ����һ�μ���ķ���,�������۲��˲�����
    '36454,������,2012-09-06,gbln���ü���=True��ʾ����Ժδ��Ƶ��ã�False��ʾ����סʱ����
    If mbytMode <> 1 And mbytKind <> E�������۵Ǽ� And lng����ID <> 0 And IIf(gbln���ü��� = True, True, str���� <> "-1") Then
        strSQL = "ZL_סԺһ�η���_Insert(" & lng����ID & "," & lng��ҳID & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    If bln����תסԺ = False Then
        '������תסԺ����ʱ,�ȵ�ҽ��,������ͨ��������Ժ,Ȼ��ת����,Ȼ���ҽ����ʽ����
        If zlInsureComeInSwap(lng����ID, lng��ҳID, lngԤ��ID, strDeposit, bytMode, True) = False Then
             gcnOracle.RollbackTrans: Exit Function
        End If
        blnNotCommit = True
    End If
    '֧������
    If Not zlInterfacePrayMoney(cllUpdate, cllThreeInsert) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    '������������
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    '101160 EMPI��ͷ����ҽԺ
    If Not EMPI_AddORUpdatePati(lng����ID, lng��ҳID, strErr) Then
        gcnOracle.RollbackTrans
        MsgBox strErr, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    If mblnAppoint Then
        '����ԤԼϵͳ�ӿ�{"�Һ�id_In": "�Һ�ID","״̬_In": "״̬" ---�ѽ��գ�δ���գ����˳�}
        Call Sys.NewSystemSvr("ԤԼ����", "��ס����סȡ��", "{""�Һ�id_In"": """ & mlng�Һ�ID & """,""״̬_In"": ""�ѽ���""}", "")
    End If
    Err = 0: On Error Resume Next
    '��Ժ����ɹ���ʼ������Ϣ
    If mclsMipModule.IsConnect = True And (Not mbytMode = EMode.EԤԼ�Ǽ�) Then
        '��ȡ�䶯ID
        If str���� = -1 Or str���� = "��ͥ����" Then
            strSQL = " Select ID,'' ����  From  ���˱䶯��¼ where ��ʼԭ��=1 And ����ID=[1] And ��ҳID=[2]"
        Else
            strSQL = " Select A.ID,B.����  From  ���˱䶯��¼ A,�շ���ĿĿ¼ B" & _
                " where A.��ʼԭ��=1 And A.��λ�ȼ�id=B.id(+) And A.����ID=[1] And A.��ҳID=[2]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", lng����ID, lng��ҳID)
        lng�䶯ID = rsTmp!ID
        str��λ�ȼ� = rsTmp!����
        
        mclsXML.ClearXmlText '��������е�XML
        
        '--������Ϣ��װ
        '������Ϣ
        mclsXML.AppendNode "in_patient"
        'patient_id      ����id  1   N
        mclsXML.appendData "patient_id", lng����ID, xsNumber  '����ID
        'page_id     ��ҳid  1   N
        mclsXML.appendData "page_id", lng��ҳID, xsNumber '��ҳID
        'patient_name        ����    1   S
        mclsXML.appendData "patient_name", txt����.Text, xsString '����
        'patient_sex     �Ա�    0..1    S
        mclsXML.appendData "patient_sex", zlCommFun.GetNeedName(cbo�Ա�.Text), xsString '�Ա�
        'in_number       סԺ��  1   S
        mclsXML.appendData "in_number", IIf(txtסԺ��.Visible And txtסԺ��.Text <> "", txtסԺ��.Text, "NULL"), xsString 'סԺ��
        mclsXML.AppendNode "in_patient", True
        
        If str���� = "-1" Then '��ͨ��Ժ�Ǽ�
            'סԺ��Ϣ
            mclsXML.AppendNode "in_hospital"
            'change_id       �䶯id  1   N
            mclsXML.appendData "change_id", lng�䶯ID, xsNumber '�䶯ID
            'in_date     ��Ժʱ��    1   s
            mclsXML.appendData "in_date", Format(txt��Ժʱ��.Text, "yyyy-MM-dd HH:mm:ss"), xsString '��Ժ����
            'in_area_id      ��Ժ����id  0..1    N
            'in_area_title       ��Ժ����    0..1    S
            If lng����ID > 0 Then
                mclsXML.appendData "in_area_id", lng����ID, xsNumber '��Ժ����ID
                mclsXML.appendData "in_area_title", cbo��Ժ����.Text, xsString  '��Ժ����
            End If
            'in_dept_id      ��Ժ����id  1   N
            mclsXML.appendData "in_dept_id", lng����ID, xsNumber '��Ժ����id
            'in_dept_title       ��Ժ����    1   S
            mclsXML.appendData "in_dept_title", cbo��Ժ����.Text, xsString  '��Ժ����
            mclsXML.AppendNode "in_hospital", True
            '�ύ��Ϣ��ZLHIS����̨��Ϣ����
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_001", mclsXML.XmlText
        Else  '��Ժ���
            'סԺ��Ϣ
            mclsXML.AppendNode "in_hospital"
            'in_date     ��Ժʱ��    1   s
            mclsXML.appendData "in_date", Format(txt��Ժʱ��.Text, "yyyy-MM-dd HH:mm:ss"), xsString '��Ժ����
            'in_area_id      ��Ժ����id  0..1    N
            mclsXML.appendData "in_area_id", lng����ID, xsNumber '��Ժ����ID
            'in_area_title       ��Ժ����    0..1    S
            mclsXML.appendData "in_area_title", cbo��Ժ����.Text, xsString  '��Ժ����
            'in_dept_id      ��Ժ����id  1   N
            mclsXML.appendData "in_dept_id", lng����ID, xsNumber '��Ժ����id
            'in_dept_title       ��Ժ����    1   S
            mclsXML.appendData "in_dept_title", cbo��Ժ����.Text, xsString  '��Ժ����
            mclsXML.appendData "in_again", chk����Ժ.Value, xsNumber
            mclsXML.AppendNode "in_hospital", True
            '��ס���
            mclsXML.AppendNode "dept_arrange"
            'change_id       �䶯id  1   N
            mclsXML.appendData "change_id", lng�䶯ID, xsNumber '�䶯ID
            'in_room     ��ס����    0..1    S
            mclsXML.appendData "in_room", IIf(str���� = "��ͥ����", "", str�����), xsString
            'in_bed      ��ס����    1   S
            mclsXML.appendData "in_bed", IIf(str���� = "��ͥ����", "", str����), xsString
            'in_tendgrade        ����ȼ�    0..1    S
            If cbo����ȼ�.ListIndex <> -1 Then
                mclsXML.appendData "in_tendgrade", cbo����ȼ�.Text, xsString
            Else
                mclsXML.appendData "in_tendgrade", "", xsString
            End If
            'in_bedgrade     ��λ�ȼ�    0..1    S
            mclsXML.appendData "in_bedgrade", IIf(str���� = "��ͥ����", "", str��λ�ȼ�), xsString
            'in_doctor       סԺҽʦ    0..1    S
            mclsXML.appendData "in_doctor", "", xsString
            'duty_nurse      ���λ�ʿ    0..1    S
            mclsXML.appendData "duty_nurse", "", xsString
            mclsXML.AppendNode "dept_arrange", True
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_002", mclsXML.XmlText
        End If
    End If
    If Err <> 0 Then Err.Clear
    
    '������ҽӿ�
    If CreatePlugInOK(glngModul) Then
        On Error Resume Next
        Call gobjPlugIn.InPatiCheckInAfter(lng����ID, lng��ҳID)
        Call zlPlugInErrH(Err, "InPatiCheckInAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    
    Err = 0: On Error GoTo errH
   '�����:56599
   'д��
   If mbln������󶨿� And mCurSendCard.bln�Ƿ�д�� Then WriteCard (lng����ID)
    
    Err = 0: On Error Resume Next:
    zlExecuteProcedureArrAy cllThreeInsert, Me.Caption
    If Err <> 0 Then
        Err = 0: On Error GoTo 0
    End If
    
    Err = 0: On Error GoTo errH
   '�ڶ���:�������תסԺ
    If bln����תסԺ Then
        If Not frmChargeTurn.ExecuteTurn(Me, mlngModul, mstrPrivs, mstrNOS, txtסԺ��.Text, lng��ҳID, CDate(txt��Ժʱ��.Text), lng����ID, lng����ID) Then
            MsgBox "ע��:" & "   δִ��ҽ����Ժ����,��HIS��Ժ�ɹ�,�벹����Ժ�Ǽ�!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        gcnOracle.BeginTrans
        blnTrans = True
        '��ԺԤ����
        If strDeposit <> "" And bln�����ʻ���Ԥ�� Then zlDatabase.ExecuteProcedure strDeposit, Me.Caption
        If mintInsure <> 0 And mstrYBPati <> "" And bytMode <> 1 Then
            strSQL = "Zl_������ҳ_ҽ������(" & lng����ID & "," & lng��ҳID & "," & mintInsure & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        '������:����ҽ��
        'ҽ������ͳһ����
        'ԤԼʱ��ͨ��ҽ������֤��ȡ������Ϣ����������ҽ������
        If zlInsureComeInSwap(lng����ID, lng��ҳID, lngԤ��ID, strDeposit, bytMode) = False Then
             gcnOracle.RollbackTrans
            MsgBox "ע��:" & "   ҽ����Ժ����ʧ��,��HIS��Ժ����ɹ�,�벹��ҽ����Ժ�Ǽ�!", vbInformation + vbOKOnly, gstrSysName
            mlng����ID = lng����ID
            mlng��ҳID = lng��ҳID
            SavePatiNew = True
            Exit Function
        End If
        blnNotCommit = True
        gcnOracle.CommitTrans: blnTrans = False
    End If
    '����:31635
    If mintInsure > 0 And mbytMode <> 1 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ComeInSwap, True, mintInsure)
    Dim strOut As String
    Call zlExcuteUploadSwap(lng����ID, strOut, mobjICCard) '�����˵�������һ��ͨ�ϴ�����
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    SavePatiNew = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    '����:31635
    If mintInsure > 0 And mbytMode <> 1 And blnNotCommit Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ComeInSwap, False, mintInsure)
    Call SaveErrLog
    Exit Function
End Function

Private Sub SetCardVaribles(ByVal blnPrepay As Boolean)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:���ý����������
    '���:blnPrepay-�Ƿ�Ԥ���������
    '����:������
    '����:2014-01-07
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    
    If blnPrepay = True Then
        With cboԤ������
        If .ListIndex = -1 Then Exit Sub
        lngIndex = .ListIndex + 1
        End With
        '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
        If Not mcolPrepayPayMode Is Nothing Then
            With mCurPrepay
                    .lngҽ�ƿ����ID = Val(mcolPrepayPayMode(lngIndex)(3))
                    .bln���ѿ� = Val(mcolPrepayPayMode(lngIndex)(5)) = 1
                    .str���㷽ʽ = Trim(mcolPrepayPayMode(lngIndex)(6))
                    .str���� = Trim(mcolPrepayPayMode(lngIndex)(1))
             End With
        End If
    Else
        With cbo��������
            If .ListIndex = -1 Then Exit Sub
            lngIndex = .ListIndex + 1
        End With
        If Not mcolCardPayMode Is Nothing Then
            With mCurCardPay
                .lngҽ�ƿ����ID = Val(mcolCardPayMode(lngIndex)(3))
                .bln���ѿ� = Val(mcolCardPayMode(lngIndex)(5)) = 1
                .str���㷽ʽ = Trim(mcolCardPayMode(lngIndex)(6))
                .str���� = Trim(mcolCardPayMode(lngIndex)(1))
             End With
         End If
     End If
End Sub

Private Function zlInsureComeInSwap(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lngԤ��ID As Long, ByVal strDeposit As String, ByVal bytMode As Byte, Optional blnMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ����Ժ�ӿ�
    '���:�����ʻ�תԤ��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-17 10:40:59
    '����:38069
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not (mintInsure <> 0 And mstrYBPati <> "" And bytMode <> 1) Then
        '��ҽ��,����true
        zlInsureComeInSwap = True: Exit Function
    End If
    
    '��Ժ��֤
    'mstrYBPati=
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    '9����;10.˳���;11��Ա���;12�ʻ����;13��ǰ״̬;14����ID;15��ְ(0,1);16����֤��;17�����;18�Ҷȼ�
    'ҽ����Ժ��֤
    If Not gclsInsure.ComeInSwap(lng����ID, lng��ҳID, CStr(Split(mstrYBPati, ";")(1)), mintInsure) Then
        If blnMsg Then
            MsgBox "ע��:" & vbCrLf & "   ҽ����Ժ����ʧ��!", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    End If
    '��ԺԤ����
    If strDeposit <> "" And is�����ʻ�(cboԤ������) Then
        If Not gclsInsure.TransferSwap(lngԤ��ID, CCur(StrToNum(txtԤ����.Text)), mintInsure) Then
            Exit Function
        End If
    End If
    zlInsureComeInSwap = True
End Function


Private Function SavePatiModi() As Boolean
'���ܣ������µĲ�����Ժ�Ǽ�(��������Ϣ����Ժ��Ϣ��Ԥ������￨)
    Dim lng�ֲ���ID As Long, lngԭ����ID As Long
    Dim strSQL As String, strMoney As String
    Dim strSurety As String, str������ As String, str����ʱ�� As String
    Dim lng����ID As Long, blnTrans As Boolean
    Dim lng��ҽ����ID As Long, lng��ҽ����ID As Long, lng����ID As Long
    Dim lng��ҽ���ID As Long, lng��ҽ���ID As Long
    Dim str�������� As String, str���� As String
    Dim cll������ As Collection '�����:56599
    Dim i As Long
    Dim arrTmp  As Variant
    Dim arrSQL As Variant
    Dim strErr As String

    arrSQL = Array()
    
    If cbo����ȼ�.ListIndex <> -1 Then
        lng����ID = cbo����ȼ�.ItemData(cbo����ȼ�.ListIndex)
    End If
    
    If cbo��Ժ����.ListIndex <> -1 Then lng����ID = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    lngԭ����ID = Val(cbo��Ժ����.Tag)
    If cbo��Ժ����.ListIndex <> -1 Then lng�ֲ���ID = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    
    If InStr(1, txt�������.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt�������.Tag)
    Else
        lng��ҽ���ID = Val(txt�������.Tag)
    End If
    If InStr(1, txt��ҽ���.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt��ҽ���.Tag)
    Else
        lng��ҽ���ID = Val(txt��ҽ���.Tag)
    End If
    
    str������ = Replace(Trim(txt������.Text), "'", "''")
    '˵��:��ʱ������Ϣ�н�����ĵ�����Ϣ�ǴӲ�����Ϣ�ж�����,��Ϊ����Ժ�ǼǺ���ܵ�������ѷ����˱仯
    If txt����ʱ�� = "__:__" Then
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
    
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    
    strSQL = "zl_��Ժ������ҳ_UPDATE(" & mbytMode & "," & _
        mlng����ID & "," & IIf(txtסԺ��.Text = "", "NULL", txtסԺ��.Text) & ",'" & txtҽ����.Text & "'," & _
        "'" & txt����.Text & "','" & zlCommFun.GetNeedName(cbo�Ա�.Text) & "','" & str���� & "','" & zlCommFun.GetNeedName(cbo�ѱ�.Text) & "'," & _
        str�������� & "," & _
        "'" & zlCommFun.GetNeedName(cbo����.Text) & "','" & zlCommFun.GetNeedName(cbo����.Text) & "','" & zlCommFun.GetNeedName(cboѧ��.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cbo����״��.Text) & "','" & zlCommFun.GetNeedName(cboְҵ.Text) & "','" & zlCommFun.GetNeedName(cbo���.Text) & "'," & _
        "'" & txt���֤��.Text & "','" & txt�����ص�.Text & "','" & txt��ͥ��ַ.Text & "'," & _
        "'" & txt��ͥ��ַ�ʱ�.Text & "','" & txt��ͥ�绰.Text & "','" & txt���ڵ�ַ.Text & "','" & txt���ڵ�ַ�ʱ�.Text & "','" & txt��ϵ������.Text & "'," & _
        "'" & zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text) & "','" & txt��ϵ�˵�ַ.Text & "'," & _
        "'" & txt��ϵ�˵绰.Text & "','" & txt������λ.Text & "'," & Val(txt������λ.Tag) & "," & _
        "'" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "','" & txt��λ������.Text & "'," & _
        "'" & txt��λ�ʺ�.Text & "','" & txt������.Tag & "'," & Val(txt������.Tag) & "," & IIf(chk��ʱ����.Tag = "", "null", chk��ʱ����.Tag) & "," & _
        mlng��ҳID & "," & ZVal(lng����ID) & "," & lng����ID & ",'" & zlCommFun.GetNeedName(cbo��Ժ����.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cbo��Ժ��ʽ.Text) & "','" & zlCommFun.GetNeedName(cboסԺĿ��.Text) & "'," & _
        chk����Ժת��.Value & ",'" & zlCommFun.GetNeedName(cbo����ҽʦ.Text) & "','" & zlCommFun.GetNeedName(txt����.Text) & "','" & zlCommFun.GetNeedName(txt����.Text) & "'," & _
        "To_Date('" & Format(txt��Ժʱ��.Text, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & zlCommFun.GetNeedName(Replace(cboҽ�Ƹ���.Text, Chr(&HD), "")) & "'," & _
        ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt�������.Text, "'", "''") & "'," & _
        ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt��ҽ���.Text, "'", "''") & "'," & _
        "'" & UserInfo.��� & "','" & UserInfo.���� & "','" & txt��ע.Text & "'," & ZVal(lng�ֲ���ID) & "," & chk����Ժ.Value & ",'" & _
        zlCommFun.GetNeedName(cbo��Ժ����.Text) & "','" & Trim(txt����֤��.Text) & "','" & zlCommFun.GetNeedName(cbo��������.Text) & _
        "','" & txt��ϵ�����֤��.Text & "','" & Trim(txtMobile.Text) & "')"
    
    '������ҳ�ӱ���Ϣ����
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",��ϵ�˸�����Ϣ,��Ժת��,���֤��״̬,�⼮���֤��,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "')"
            End If
            If InStr(",��ϵ�˸�����Ϣ,���֤��״̬,�⼮���֤��,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
            End If
        Next
    End If
    
    If txt������.Visible And txt������.Enabled And str������ <> "" Then
        'û��Ȩ��ʱ���ɼ�,���ز�������Ϊ�������ϢʱΪ����,�Լ��޸ĵĵ�����¼ʱ�޹���ʱ����
        str����ʱ�� = "null"
        If Not IsNull(dtp����ʱ��.Value) Then str����ʱ�� = "To_Date('" & Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        
        If Trim(txt������.Tag) = "" Then    '֮ǰ�Ǽ�ʱû�е���
            strSurety = "zl_���˵�����¼_insert(" & mlng����ID & "," & mlng��ҳID & ",'" & str������ & "'," & _
            Val(txt������.Text) & "," & chk��ʱ����.Value & ",'" & Trim(txtReason.Text) & "',NULL," & str����ʱ�� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            strSurety = "zl_���˵�����¼_update(" & mlng����ID & "," & mlng��ҳID & ",'" & str������ & "'," & _
                Val(txt������.Text) & "," & chk��ʱ����.Value & ",'" & Trim(txtReason.Text) & "',NULL," & str����ʱ�� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "',To_Date('" & dtp����ʱ��.Tag & "','yyyy-mm-dd hh24:mi:ss'))"
        End If
    End If
    '���ؽṹ����ַSQL
    If gbln���ýṹ����ַ Then
        Call CreateStructAddressSQL(mlng����ID, mlng��ҳID, arrSQL, PatiAddress, 1)
    End If
    
    '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
    If mlng����ID > 0 Then Call AddCertificate(mlng����ID, arrSQL, zlDatabase.Currentdate)
    
    On Error GoTo errH
    gcnOracle.BeginTrans
        blnTrans = True
        '�޸���Ժ��Ϣǰ����һ�Լ���ķ���(�����ڸ��Ĳ���ǰ����)
        If lng�ֲ���ID <> lngԭ����ID And mbytMode <> 1 And mbytKind <> E�������۵Ǽ� Then
            strMoney = "ZL_סԺһ�η���_Delete(" & mlng����ID & "," & mlng��ҳID & ")"
            zlDatabase.ExecuteProcedure strMoney, Me.Caption
        End If
        
        '�޸���Ժ��Ϣ
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        '������ҳ�ӱ���Ϣ
        For i = LBound(arrSQL) To UBound(arrSQL)
             zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        '�޸ĵ�����Ϣ
        If strSurety <> "" Then zlDatabase.ExecuteProcedure strSurety, Me.Caption
        '�����:56599
        '���벡�˽�������Ϣ
        Set cll������ = New Collection
        If mlng����ID > 0 Then Call Add�����������Ϣ(mlng����ID, cll������)
        zlExecuteProcedureArrAy cll������, Me.Caption, True, True
        
        '�޸����²���һ�μ���ķ���
        '36454,������,2012-09-06,gbln���ü���=True��ʾ����Ժδ��Ƶ��ã�False��ʾ����סʱ����
        If lng�ֲ���ID <> lngԭ����ID And mbytMode <> 1 And mbytKind <> E�������۵Ǽ� And gbln���ü��� = True Then
            strMoney = "ZL_סԺһ�η���_Insert(" & mlng����ID & "," & mlng��ҳID & ")"
            zlDatabase.ExecuteProcedure strMoney, Me.Caption
        End If
        '101160EMPI
        If Not EMPI_AddORUpdatePati(mlng����ID, mlng��ҳID, strErr) Then
            gcnOracle.RollbackTrans
            MsgBox strErr, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    gcnOracle.CommitTrans: blnTrans = False
    '����96847��118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng����ID, mlng��ҳID) <> 1 Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    SavePatiModi = True
    '�����:56599
    'д��
    If mbln������󶨿� And mCurSendCard.bln�Ƿ�д�� Then WriteCard (mlng����ID)
    
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadSurety(lng����ID As Long, lng��ҳID As Long, dat��Ժʱ�� As Date)
'����:��Ժ�Ǽǵ��޸ĺͲ鿴(����ԤԼ��ԤԼ����)���ص�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim Dat��Сʱ�� As Date
    
    On Error GoTo errH
    dtp����ʱ��.MinDate = dat��Ժʱ��
    
    '87466,LPF,������Ϣ��ȡ�����������ɾ����־=1��;����������Ϣ����ƴ�ӣ��벡����Ϣ�洢�ĵ�����Ϣ����һ��
    strSQL = "SELECT ������, Decode(������, 999999999, '����', To_Char(������, '999999990.00')) AS ������, ��������, ����ԭ��, " & vbNewLine & _
            "       To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��,To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') �Ǽ�ʱ��" & vbNewLine & _
            "FROM ���˵�����¼" & vbNewLine & _
            "WHERE ����id = [1] AND ��ҳid = [2] AND (����ʱ�� is null or ����ʱ��>sysdate) And ɾ����־ = 1"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 1 Then
        '������Ч������¼��Ҫ������Ϣ���޸�
        dtp����ʱ��.Value = Null
        txt������ = ""
        Do Until rsTmp.EOF
            If "" & rsTmp!������ = "����" Then
                txt������ = Format(Val(txt������.Text) + 999999999, "0.00")
            Else
                txt������ = Format(Val(txt������.Text) + Nvl(rsTmp!������, 0), "0.00")
            End If
            If Nvl(dtp����ʱ��.Value, "3000-01-01 00:00:00") > Nvl(rsTmp!����ʱ��, "3000-01-01 00:00:00") Then
                dtp����ʱ��.Value = Nvl(rsTmp!����ʱ��, "3000-01-01 00:00:00")
            End If
            txt������ = IIf(txt������ = "", "", txt������ & ",") & rsTmp!������
            rsTmp.MoveNext
        Loop
        'txt������ = "���˵���"
        txt������.Enabled = False: txt������.BackColor = Me.BackColor
        chkUnlimit.Enabled = False
        txt������.Enabled = False: txt������.BackColor = Me.BackColor
        dtp����ʱ��.Enabled = False
        chk��ʱ����.Enabled = False
        txtReason.Enabled = False
    ElseIf rsTmp.RecordCount = 1 Then
        '�޸ĵ������һ����Ч�ĵ�����¼
        txt������.Text = "" & rsTmp!������
        chkUnlimit.Value = IIf("" & rsTmp!������ = "����", 1, 0)   'ֵ��ͬʱ����ʽ����click�¼�
        If chkUnlimit.Value = 1 Then
            txt������ = "999999999"
        Else
            txt������ = "" & rsTmp!������
        End If
        dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm"
        If IsDate("" & rsTmp!����ʱ��) Then '��ʱ��϶�����С����Ժʱ��
            dtp����ʱ��.Value = CDate(rsTmp!����ʱ��)
        Else
            dtp����ʱ��.Value = Null
        End If
        dtp����ʱ��.Tag = rsTmp!�Ǽ�ʱ��
        txtReason.Text = Nvl(rsTmp!����ԭ��)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatiOutDate(ByVal lng����ID As Long) As Date
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Max(��Ժ����) ��Ժ���� From ������ҳ Where ����ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If rsTmp.RecordCount > 0 Then
        If Not IsNull(rsTmp!��Ժ����) Then GetPatiOutDate = rsTmp!��Ժ����
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function ReadPatiReg(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'����:��ȡ������Ժ�ǼǼ�¼����ʾ
'����:mbytInState-0,1,2���п��ܵ��ñ�����:�Ǽ��޸�,�Ǽǲ鿴,ԤԼ����
    Dim rsTmp As ADODB.Recordset
    Dim rsDiagnosis As ADODB.Recordset
    Dim rsPlus As ADODB.Recordset '�����ӱ���Ϣֵ
    Dim DatOut As Date
    Dim lngIdx As Long
    Dim strPlus As String   '��¼�ӱ���Ϣ��
    Dim i As Long
    Dim arrTmp As Variant
    
    On Error GoTo errH
       
    gstrSQL = _
        " Select A.����ID,A.���￨��,A.�����,B.סԺ��,B.���ۺ�,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,H.���� ��������,B.�ѱ�," & _
        "   A.סԺ����,A.����,A.����,A.ѧ��,A.����״��,A.ְҵ,A.���,A.���֤��,A.�ֻ���,A.����֤��,A.��������,A.�����ص�,A.��ͥ��ַ," & _
        "   A.��ͥ�绰,A.��ͥ��ַ�ʱ�, A.���ڵ�ַ, A.���ڵ�ַ�ʱ�, A.����, A.��ϵ�˹�ϵ,A.��ϵ������,A.��ϵ�˵�ַ,A.��ϵ�˵绰,A.��ϵ�����֤��," & _
        "   A.������λ,A.��ͬ��λID,A.��λ�绰,A.��λ�ʱ�,A.��λ������,A.��λ�ʺ�,A.������,A.������,A.��������," & _
        "   B.����,Nvl(A.ҽ����,F.��Ϣֵ) as ҽ����,B.��Ժ��ʽ,b.��Ժ����,B.��Ժ����,B.��Ժ����,B.סԺĿ��,B.��Ժ����,B.����ҽʦ,Nvl(B.����, A.����) ����,B.ҽ�Ƹ��ʽ," & _
        "   Nvl(B.�Ƿ����,0) as �Ƿ����,Nvl(B.����Ժת��,0) as ����Ժת��,C.���� as ��Ժ����,B.��Ժ����ID," & _
        "   G.���� as ��Ժ����,B.��Ժ����ID,D.���� as ����ȼ�,B.��ע,B.����Ժ,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) ��������, B.�Һ�ID " & _
        " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ D,������ҳ�ӱ� F,���ű� G,������� H" & _
        " Where B.����ID=A.����ID And B.��Ժ����ID=C.ID And B.��Ժ����ID=G.ID(+) And B.����ȼ�ID=D.ID(+) And A.����=H.���(+)" & _
        " And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+) And F.��Ϣ��(+)='ҽ����'" & _
        " And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng��ҳID)
    If rsTmp.EOF Then Exit Function
    Set mrsPatiReg = rsTmp.Clone
    
    If Not FuncPlugPovertyInfo(Val(rsTmp!����ID)) Then Exit Function
    
    txtPatient.Text = rsTmp!����ID
    txtPatient.Tag = rsTmp!����ID
    mlng�Һ�ID = Nvl(rsTmp!�Һ�ID, 0)
    txtסԺ��.Text = Decode(mbytKind, E�������۵Ǽ�, Nvl(rsTmp!�����), EסԺ���۵Ǽ�, Nvl(rsTmp!���ۺ�), Nvl(rsTmp!סԺ��))
    
    txtסԺ��.Tag = txtסԺ��.Text
    txt����.Text = rsTmp!����
    
    If (mbytMode = EMode.E����ԤԼ Or mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0) And mbytInState = EState.E���� Then 'ԤԼ����ʱ,�������ҳIDΪ0
        txtPages.Text = GetMaxMinPage(lng����ID) + 1
    Else
        txtPages.Text = lng��ҳID
    End If
    
    'ԤԼ���Ĳ��˲�����༭��Ժ��������Ժ����
    mstrAppointBed = "": mblnAppoint = False
    If (mbytInState = EState.E�޸� And mbytMode = EMode.EԤԼ�Ǽ�) Or (mbytMode = EMode.E����ԤԼ) Then
        mblnAppoint = IsAppointPati(mlng�Һ�ID, mstrAppointBed) 'T-ԤԼ���Ĳ���
        cbo��Ժ����.Enabled = Not mblnAppoint
        cbo��Ժ����.Enabled = Not mblnAppoint
    End If
    
    If mbytInState = EState.E���� And mbytKind = EKind.EסԺ��Ժ�Ǽ� And mbytMode <> EMode.EԤԼ�Ǽ� Then
        txtTimes.Text = GetMaxInHosTimes(lng����ID) + 1
    Else
        txtTimes.Text = "" & rsTmp!סԺ����
    End If
    txtTimes.Tag = txtTimes.Text
    
    txtҽ����.Text = Nvl(rsTmp!ҽ����)
    txtҽ����.Locked = Not IsNull(rsTmp!����)
    txt����.Text = "" & rsTmp!��������
    
    cbo�Ա�.ListIndex = GetCboIndex(cbo�Ա�, IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�))
    If cbo�Ա�.ListIndex = -1 Then Call SetCboDefault(cbo�Ա�)
    Call LoadOldData("" & rsTmp!����, txt����, cbo���䵥λ)
    mblnChange = False
    txt��������.Text = Format(IIf(IsNull(rsTmp!��������), "____-__-__", rsTmp!��������), "YYYY-MM-DD")
    If rsTmp!���� Like "Լ*" Or Trim(Nvl(rsTmp!����)) = "����" Then
        If "" & rsTmp!�������� = "____-__-__" Then
            txt��������.Enabled = False
            txt����ʱ��.Enabled = False
        End If
    Else
        txt��������.Enabled = True
        txt����ʱ��.Enabled = True
    End If
    mblnChange = True
    
    txt��Ժʱ��.Text = Format(IIf(lng��ҳID = 0 And mbytInState <> 1, zlDatabase.Currentdate, Nvl(rsTmp!��Ժ����, "")), "yyyy-MM-dd HH:mm")
    If lng��ҳID > 1 And mbytInState = EState.E�޸� Then
        DatOut = GetPatiOutDate(lng����ID) '�ϴγ�Ժʱ��
        If DatOut <> CDate(0) Then txt��Ժʱ��.Tag = Format(DatOut, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If Not IsNull(rsTmp!��������) Then
        If mbytInState <> 2 Then txt����.Text = ReCalcOld(CDate(Format(rsTmp!��������, "YYYY-MM-DD HH:MM:SS")), cbo���䵥λ, Val(rsTmp!����ID), , CDate(txt��Ժʱ��.Text)) '���ݳ���������������
        If CDate(txt��������.Text) - CDate(rsTmp!��������) <> 0 Then
            mblnChange = False
            txt����ʱ��.Text = Format(rsTmp!��������, "HH:MM")
            mblnChange = True
        End If
    Else
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
    cbo���䵥λ.Tag = txt����.Text & "_" & cbo���䵥λ.Text  '���ڱ�������Ƿ�䶯
    
    mblnChange = False          '�޸ĺͲ鿴ʱ,���֤����������ڶ���
    txt���֤��.Text = "" & rsTmp!���֤��
    mblnChange = True
    cboIDNumber.Enabled = txt���֤��.Text = ""
    txt����֤��.Text = "" & rsTmp!����֤��
     
    
    cbo�ѱ�.ListIndex = GetCboIndex(cbo�ѱ�, IIf(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�))
    If cbo�ѱ�.ListIndex = -1 Then Call SetCboDefault(cbo�ѱ�)
    If mbytInState = EState.E�޸� Then If Not IsNull(rsTmp!��Ժ����) Then cbo�ѱ�.Enabled = False
    
    cbo����.ListIndex = GetCboIndex(cbo����, IIf(IsNull(rsTmp!����), "", rsTmp!����))
    If cbo����.ListIndex = -1 Then Call SetCboDefault(cbo����)
    
    cbo����.ListIndex = GetCboIndex(cbo����, IIf(IsNull(rsTmp!����), "", rsTmp!����))
    If cbo����.ListIndex = -1 Then Call SetCboDefault(cbo����)
    
    cboѧ��.ListIndex = GetCboIndex(cboѧ��, IIf(IsNull(rsTmp!ѧ��), "", rsTmp!ѧ��))
    If cboѧ��.ListIndex = -1 And Not IsNull(rsTmp!ѧ��) Then
        cboѧ��.AddItem rsTmp!ѧ��, 0: cboѧ��.ListIndex = 0
    End If
    
    cbo����״��.ListIndex = GetCboIndex(cbo����״��, IIf(IsNull(rsTmp!����״��), "", rsTmp!����״��))
    If cbo����״��.ListIndex = -1 And Not IsNull(rsTmp!����״��) Then
        cbo����״��.AddItem rsTmp!����״��, 0: cbo����״��.ListIndex = 0
    End If
    
    cboְҵ.ListIndex = GetCboIndex(cboְҵ, IIf(IsNull(rsTmp!ְҵ), "", rsTmp!ְҵ))
    If cboְҵ.ListIndex = -1 And Not IsNull(rsTmp!ְҵ) Then
        cboְҵ.AddItem rsTmp!ְҵ, 0: cboְҵ.ListIndex = 0
    End If
    
    cbo���.ListIndex = GetCboIndex(cbo���, IIf(IsNull(rsTmp!���), "", rsTmp!���))
    If cbo���.ListIndex = -1 And Not IsNull(rsTmp!���) Then
        cbo���.AddItem rsTmp!���, 0: cbo���.ListIndex = 0
    End If
    
    txt����.Text = Nvl(rsTmp!����)
    cbo��������.ListIndex = GetCboIndex(cbo��������, Nvl(rsTmp!��������))
             
    txt��ͥ�绰.Text = IIf(IsNull(rsTmp!��ͥ�绰), "", rsTmp!��ͥ�绰)
    txt��ͥ��ַ�ʱ�.Text = IIf(IsNull(rsTmp!��ͥ��ַ�ʱ�), "", rsTmp!��ͥ��ַ�ʱ�)
    txt���ڵ�ַ�ʱ�.Text = IIf(IsNull(rsTmp!���ڵ�ַ�ʱ�), "", rsTmp!���ڵ�ַ�ʱ�)
    txt��ϵ������.Text = IIf(IsNull(rsTmp!��ϵ������), "", rsTmp!��ϵ������)
    
    cbo��ϵ�˹�ϵ.ListIndex = GetCboIndex(cbo��ϵ�˹�ϵ, IIf(IsNull(rsTmp!��ϵ�˹�ϵ), "", rsTmp!��ϵ�˹�ϵ))
    If Not cbo��ϵ�˹�ϵ.ListIndex = -1 And Not IsNull(rsTmp!��ϵ�˹�ϵ) Then
        cbo��ϵ�˹�ϵ.AddItem rsTmp!��ϵ�˹�ϵ, 0: cbo��ϵ�˹�ϵ.ListIndex = 0
    End If
    '��¼�´ӱ���Ϣ��
    If zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text) = "����" Then strPlus = strPlus & "," & "��ϵ�˸�����Ϣ"
    txt��ϵ�˵绰.Text = IIf(IsNull(rsTmp!��ϵ�˵绰), "", rsTmp!��ϵ�˵绰)
    txt��ϵ�����֤��.Text = IIf(IsNull(rsTmp!��ϵ�����֤��), "", rsTmp!��ϵ�����֤��)
    txt������λ.Text = IIf(IsNull(rsTmp!������λ), "", rsTmp!������λ)
    txt������λ.Tag = IIf(IsNull(rsTmp!��ͬ��λID), "", rsTmp!��ͬ��λID)
    txt��λ�绰.Text = IIf(IsNull(rsTmp!��λ�绰), "", rsTmp!��λ�绰)
    txt��λ�ʱ�.Text = IIf(IsNull(rsTmp!��λ�ʱ�), "", rsTmp!��λ�ʱ�)
    txt��λ������.Text = IIf(IsNull(rsTmp!��λ������), "", rsTmp!��λ������)
    txt��λ�ʺ�.Text = IIf(IsNull(rsTmp!��λ�ʺ�), "", rsTmp!��λ�ʺ�)
    txt��ע.Text = Nvl(rsTmp!��ע)
    txtMobile.Text = rsTmp!�ֻ��� & ""
    
    If gbln���ýṹ����ַ Then
        Call ReadStructAddress(lng����ID, lng��ҳID, PatiAddress)
        txt�����ص�.Text = PatiAddress(E_IX_�����ص�).Value
        txt����.Text = PatiAddress(E_IX_����).Value
        txt��ͥ��ַ.Text = PatiAddress(E_IX_��סַ).Value
        txt���ڵ�ַ.Text = PatiAddress(E_IX_���ڵ�ַ).Value
        txt��ϵ�˵�ַ.Text = PatiAddress(E_IX_��ϵ�˵�ַ).Value
    Else
        txt�����ص�.Text = IIf(IsNull(rsTmp!�����ص�), "", rsTmp!�����ص�)
        txt����.Text = Nvl(rsTmp!����)
        txt��ͥ��ַ.Text = IIf(IsNull(rsTmp!��ͥ��ַ), "", rsTmp!��ͥ��ַ)
        txt��ͥ��ַ.ToolTipText = txt��ͥ��ַ.Text
        txt���ڵ�ַ.Text = IIf(IsNull(rsTmp!���ڵ�ַ), "", rsTmp!���ڵ�ַ)
        txt��ϵ�˵�ַ.Text = IIf(IsNull(rsTmp!��ϵ�˵�ַ), "", rsTmp!��ϵ�˵�ַ)
        txt��ϵ�˵�ַ.ToolTipText = txt��ϵ�˵�ַ.Text
    End If

    '������Ϣ(ԤԼ���䵣����Ϣ,ԤԼ�������������)
    If mbytMode = 0 And mlng����ID <> 0 Then
        If mbytInState = 1 Then
            txt������.Tag = "" & rsTmp!������   '����ԭ������ص�������Ϣ��,��Ϊ���ܴ����ѵ��ڵĵ���,�Ͳ������޸�
            txt������.Tag = "" & rsTmp!������
            chk��ʱ����.Tag = "" & rsTmp!��������
        End If
        Call LoadSurety(lng����ID, lng��ҳID, rsTmp!��Ժ����)
    End If
    
    '��Ժ��Ϣ
    If gbln��ѡ���� Then    '(ֻӰ���޸�ʱ)
        '����29007 by lesfeng 2010-04-12
        If IsNull(rsTmp!��Ժ����) And Not IsNull(rsTmp!��Ժ����) Then
            mrsUnitDept.Filter = "����ID=" & Val(rsTmp!��Ժ����ID) & " and ����ID=" & Val(rsTmp!��Ժ����ID)
            If mrsUnitDept.RecordCount > 0 Then
                lngIdx = cbo.FindIndex(cbo��Ժ����, mrsUnitDept!����ID)
                If lngIdx <> -1 Then
                    cbo��Ժ����.ListIndex = lngIdx
                End If
            Else
                mrsUnitDept.Filter = "����ID=" & Val(rsTmp!��Ժ����ID)
                If mrsUnitDept.RecordCount > 0 Then
                    lngIdx = cbo.FindIndex(cbo��Ժ����, mrsUnitDept!����ID)
                    If lngIdx <> -1 Then
                        cbo��Ժ����.ListIndex = lngIdx
                    End If
                End If
            End If
        Else
            cbo��Ժ����.ListIndex = GetCboIndex(cbo��Ժ����, "" & rsTmp!��Ժ����)
        End If
        '----------------------------------
        If cbo��Ժ����.ListIndex = -1 Then
            If Not IsNull(rsTmp!��Ժ����) And mbytInState = EState.E���� Then
                cbo��Ժ����.AddItem rsTmp!��Ժ����
                cbo��Ժ����.ItemData(cbo��Ժ����.NewIndex) = Nvl(rsTmp!��Ժ����ID, 0)
                cbo��Ժ����.ListIndex = cbo��Ժ����.NewIndex
            Else
                If cbo��Ժ����.ListCount > 0 Then cbo��Ժ����.ListIndex = 0 '��һ���ǲ�ȷ������
            End If
        End If
        cbo��Ժ����.ListIndex = GetCboIndex(cbo��Ժ����, rsTmp!��Ժ����)
        If cbo��Ժ����.ListIndex = -1 And mbytInState = EState.E���� Then
            cbo��Ժ����.AddItem rsTmp!��Ժ����, 0
            cbo��Ժ����.ItemData(cbo��Ժ����.NewIndex) = Nvl(rsTmp!��Ժ����ID, 0)
            cbo��Ժ����.ListIndex = 0
        End If
    Else
        cbo��Ժ����.ListIndex = GetCboIndex(cbo��Ժ����, rsTmp!��Ժ����)
        If cbo��Ժ����.ListIndex = -1 And mbytInState = EState.E���� Then
            cbo��Ժ����.AddItem rsTmp!��Ժ����, 0
            cbo��Ժ����.ItemData(cbo��Ժ����.NewIndex) = Nvl(rsTmp!��Ժ����ID, 0)
            cbo��Ժ����.ListIndex = 0
        End If
        cbo��Ժ����.ListIndex = GetCboIndex(cbo��Ժ����, "" & rsTmp!��Ժ����)
        If cbo��Ժ����.ListIndex = -1 Then
            If Not IsNull(rsTmp!��Ժ����) And mbytInState = EState.E���� Then
                cbo��Ժ����.AddItem rsTmp!��Ժ����
                cbo��Ժ����.ItemData(cbo��Ժ����.NewIndex) = Nvl(rsTmp!��Ժ����ID, 0)
                cbo��Ժ����.ListIndex = cbo��Ժ����.NewIndex
            Else
                If cbo��Ժ����.ListCount > 0 Then cbo��Ժ����.ListIndex = 0
            End If
        End If
    End If
    
    If gbln��Ժ��� And mbytMode <> EMode.EԤԼ�Ǽ� And mbytInState = EState.E���� Then
        cbo��λ.ListIndex = GetCboIndex(cbo��λ, Nvl(rsTmp!��Ժ����))
        If cbo��λ.ListIndex = -1 And Not IsNull(rsTmp!��Ժ����) Then    '����д��ţ��ǲ������޸ĵ�
            cbo��λ.AddItem Nvl(rsTmp!��Ժ����), 0
            cbo��λ.ListIndex = 0
        End If
    End If
   
    '��¼ԭʼֵ
    If cbo��Ժ����.ListIndex <> -1 And mbytInState = EState.E�޸� Then
        cbo��Ժ����.Tag = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    End If
    If cbo��Ժ����.ListIndex <> -1 And mbytInState = EState.E�޸� Then
        cbo��Ժ����.Tag = cbo��Ժ����.ItemData(cbo��Ժ����.ListIndex)
    End If
    
    cbo��Ժ����.ListIndex = GetCboIndex(cbo��Ժ����, IIf(IsNull(rsTmp!��Ժ����), "", rsTmp!��Ժ����))
    If cbo��Ժ����.ListIndex = -1 Then Call SetCboDefault(cbo��Ժ����)
        
    cbo��Ժ��ʽ.ListIndex = GetCboIndex(cbo��Ժ��ʽ, IIf(IsNull(rsTmp!��Ժ��ʽ), "", rsTmp!��Ժ��ʽ))
    If cbo��Ժ��ʽ.ListIndex = -1 Then Call SetCboDefault(cbo��Ժ��ʽ)
    '��¼�´ӱ���Ϣ��
    If zlCommFun.GetNeedName(cbo��Ժ��ʽ.Text) = "ת��" Then strPlus = strPlus & "," & "��Ժת��"
    
    '���˺�:2007/09/13
    cbo��Ժ����.ListIndex = GetCboIndex(cbo��Ժ����, IIf(IsNull(rsTmp!��Ժ����), "", rsTmp!��Ժ����))
    If cbo��Ժ����.ListIndex = -1 Then Call SetCboDefault(cbo��Ժ����)
    
    cboסԺĿ��.ListIndex = GetCboIndex(cboסԺĿ��, IIf(IsNull(rsTmp!סԺĿ��), "", rsTmp!סԺĿ��))
    If cboסԺĿ��.ListIndex = -1 Then Call SetCboDefault(cboסԺĿ��)
    
    cboҽ�Ƹ���.ListIndex = GetCboIndex(cboҽ�Ƹ���, IIf(IsNull(rsTmp!ҽ�Ƹ��ʽ), "", rsTmp!ҽ�Ƹ��ʽ), , True)
    If cboҽ�Ƹ���.ListIndex = -1 Then Call SetCboDefault(cboҽ�Ƹ���)
            
            
            
    If IsNull(rsTmp!����ȼ�) Then
        If cbo����ȼ�.ListCount = 0 Then cbo����ȼ�.AddItem "": cbo����ȼ�.ItemData(cbo����ȼ�.NewIndex) = 0    '����ʱ
        cbo����ȼ�.ListIndex = 0 'װ��ʱ,��һ���ǿ�
    Else
        cbo����ȼ�.ListIndex = GetCboIndex(cbo����ȼ�, rsTmp!����ȼ�)
        If cbo����ȼ�.ListIndex = -1 Then
            cbo����ȼ�.AddItem rsTmp!����ȼ�
            cbo����ȼ�.ListIndex = cbo����ȼ�.NewIndex
        End If
    End If
    
    cbo����ҽʦ.ListIndex = GetCboIndex(cbo����ҽʦ, IIf(IsNull(rsTmp!����ҽʦ), "", rsTmp!����ҽʦ))
    If cbo����ҽʦ.ListIndex = -1 And Not IsNull(rsTmp!����ҽʦ) Then
        cbo����ҽʦ.AddItem rsTmp!����ҽʦ, 0: cbo����ҽʦ.ListIndex = 0
    End If
    
        
    chk����Ժ.Value = Val("" & rsTmp!����Ժ)
    chk����Ժת��.Value = rsTmp!����Ժת��
    chk���.Value = rsTmp!�Ƿ����
    
    
    '��ʾ����������
    Set rsDiagnosis = GetDiagnosticInfo(lng����ID, lng��ҳID, "1,11", IIf(mbytMode = EMode.E����ԤԼ Or mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0 And mbytInState = EState.E����, "3", "2"), Val(rsTmp!��Ժ����ID & ""))
    If Not rsDiagnosis Is Nothing Then
        rsDiagnosis.Filter = "�������=1"
        If Not rsDiagnosis.EOF Then
            txt�������.Text = Nvl(rsDiagnosis!�������): txt�������.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl�������.Tag = txt�������.Text
        End If
        
        rsDiagnosis.Filter = "�������=11"
        If Not rsDiagnosis.EOF Then
            txt��ҽ���.Text = Nvl(rsDiagnosis!�������): txt��ҽ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
        End If
    End If
     
    If Not IsNull(rsTmp!����) Then
        If mstrYBPati = "" Then mstrYBPati = "��ҽ��"         '����,�޸�,�鿴���ܵ���,ֻ��Ϊ�˱�ʶ�Ƿ�ҽ������
    End If
    '�����:56599
    Call Load�����������Ϣ(lng����ID)
    '�����ӱ���Ϣ
    If strPlus <> "" Then
        strPlus = Mid(strPlus, 2)
        arrTmp = Split(strPlus, ",")
        Set rsPlus = Get������ҳ�ӱ�(lng����ID, lng��ҳID, strPlus)
        
        If rsPlus.RecordCount > 0 Then
            rsPlus.Filter = "��Ϣ��='��ϵ�˸�����Ϣ'"
            If Not rsPlus.EOF Then txtLinkManInfo.Text = rsPlus!��Ϣֵ & ""
            rsPlus.Filter = "��Ϣ��='��Ժת��'"
            If Not rsPlus.EOF Then txtת��.Text = rsPlus!��Ϣֵ & ""
        End If
    End If
    
     '������Ϣ�ӱ�
    If txt���֤��.Text = "" Then
        Set rsPlus = Get������Ϣ�ӱ�(lng����ID, "���֤��״̬")
        rsPlus.Filter = "��Ϣ��='���֤��״̬'"
        If Not rsPlus.EOF Then
            If Not IsNull(rsPlus!��Ϣֵ) Then
                cbo.Locate cboIDNumber, zlCommFun.GetNeedName(rsPlus!��Ϣֵ)
            End If
        End If
        If Trim(zlCommFun.GetNeedName(cbo����.Text)) <> "�й�" And Trim(txt���֤��.Text) = "" Then
            If Trim(zlCommFun.GetNeedName(cboIDNumber.Text)) = "" Then
                 Set rsPlus = Get������Ϣ�ӱ�(lng����ID, "�⼮���֤��")
                rsPlus.Filter = "��Ϣ��='�⼮���֤��'"
                If Not rsPlus.EOF Then
                    If Not IsNull(rsPlus!��Ϣֵ) Then
                        txt���֤��.Text = "" & rsPlus!��Ϣֵ
                    End If
                End If
            End If
        End If
    End If
    ReadPatiReg = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Function RequestCode() As Boolean
    RequestCode = gint����������� = 2 Or (gint����������� = 3 And mstrYBPati <> "")
End Function
''''
''''
''''
''''Private Function zlSquareSimulation(ByRef lngOut�ӿڱ�� As Long, ByRef strOutBalanceInfor As String) As Boolean
''''    ------------------------------------------------------------------------------------------------------------------------
''''����:     ���п�������㽻��
''''���:
''''����:      lngOut�ӿڱ�� -�ӿڱ��
''''             strBalanceInfor -���ؽ��㽻��
''''����:     �ɹ� (��ǽ��㿨����), ����true, ���򷵻�False
''''����:     ���˺�
''''    ���ڣ�2010-07-09 16:55:19
''''˵��:
''''    ------------------------------------------------------------------------------------------------------------------------
''''    Dim i As Long
''''    Dim strBlanceInfor As String, varData As Variant, blnHave���㷽ʽ As Boolean, lng�ӿڱ�� As Long
''''    strOutBalanceInfor = ""
''''    lngOut�ӿڱ�� = 0: strOutBalanceInfor = ""
''''    If cboԤ������.ItemData(cboԤ������.ListIndex) <> 8 Then    '�ǽ��㿨����Ϊtrue
''''        zlSquareSimulation = True
''''        Exit Function
''''    End If
''''    If Not mtySquareCard.blnExistsObjects Or mobjSquareCard Is Nothing Then
''''        MsgBox "ע��:" & vbCrLf & "    ���㿨���㲿��������,�����ý��㿨���ʽ�Ԥ��,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
''''        Exit Function
''''    End If
''''
''''    zlSimulationBrushCard(ByVal frmMain As Object, ByVal Dblˢ����� As Double, _
''''        ByRef lng�ӿڱ�� As Long, ByRef strBlanceInfor As String) As Boolean
''''        '------------------------------------------------------------------------------------------------------------------------
''''        '���ܣ�ѡ��ָ��������
''''        '��Σ�frmMain HIS���� ���õ�������
''''        '         Dblˢ����� HIS���� ����Ԥ�������еĽ��
''''        '         Lng�ӿڱ��          HIS������
''''        '���Σ�Lng�ӿڱ�� ����    �Ժ��ֽ��㿨����
''''        '         strBlanceInfor  ����    ��||�ָ�: �ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע
''''        '���أ�
''''        '���ƣ����˺�
''''        '���ڣ�2010-06-18 11:33:22
''''        '˵������Ԥ�������������Ԥ��ʱ�����ȷ����ťʱ����(����ǰ����)
''''        '------------------------------------------------------------------------------------------------------------------------
''''    ģ�����
''''     If mobjSquareCard.zlSimulationBrushCard(Me, Val(StrToNum(txtԤ����.Text)), lng�ӿڱ��, strBlanceInfor) = False Then
''''          Exit Function
''''     End If
''''    strBlanceInfor:�ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע
''''    varData = Split(strBlanceInfor, "||")
''''    If Trim(strBlanceInfor) = "" Then
''''           MsgBox "ע��:" & vbCrLf & "    ���صĽ�����Ϣ��ʽ����,����POS�ӿڿ�����ϵ!", vbInformation + vbDefaultButton1 + vbOKOnly
''''           Exit Function
''''    End If
''''
''''    blnHave���㷽ʽ = False
''''
''''    With cboԤ������
''''       For i = 0 To .ListCount - 1
''''            If NeedName(.List(i)) = CStr(varData(2)) Then
''''                    blnHave���㷽ʽ = True:
''''                  If i <> .ListIndex Then .ListIndex = i
''''                  Exit For
''''            End If
''''       Next
''''        If Val(varData(3)) <= 0 Then
''''                MsgBox "ע��:" & vbCrLf & "    �����㷵�صĽ������С�ڵ�����,����!", vbInformation + vbDefaultButton1 + vbOKOnly
''''                Exit Function
''''        End If
''''        If Round(Val(varData(3)), 3) <> Round(Val(StrToNum(txtԤ����.Text)), 3) Then
''''            txtԤ����.Text = Format(Val(varData(3)), "0.00")
''''        End If
''''
''''        If CStr(varData(2)) = "" Then
''''                MsgBox "ע��:" & vbCrLf & "    �����㷵�صĽ��㷽ʽΪ����,����!", vbInformation + vbDefaultButton1 + vbOKOnly
''''                Exit Function
''''        End If
''''        If blnHave���㷽ʽ = False Then
''''            MsgBox "ע��:" & vbCrLf & "    �����㷵�صĽ��㷽ʽ����ȷ,������:" & varData(2) & vbCrLf & _
''''                "     ��δ����Ӧ�ó���,���ڽ��㷽ʽ������!", vbInformation + vbDefaultButton1 + vbOKOnly
''''            Exit Function
''''        End If
''''    End With
''''    strOutBalanceInfor = strBlanceInfor: lngOut�ӿڱ�� = lng�ӿڱ��
''''    zlSquareSimulation = True
''''End Function
'''Private Function zlSequareBlanceToDeposit(ByVal lngԤ��ID As Long, ByVal lng�ӿڱ�� As Long, strBlanceInfor As String) As Boolean
'''    '---------------------------------------------------------------------------------------------------------------------------------------------
'''    '����:���㿨�Ľ���
'''    '����:�ɹ�,����true,���򷵻�False
'''    '����:���˺�
'''    '����:2010-02-08 16:40:12
'''    '---------------------------------------------------------------------------------------------------------------------------------------------
'''    Dim rsSquare As ADODB.Recordset
'''    If mbytInState <> 0 Then GoTo goEnd:
'''
'''    '���˺�:
'''    If Not mtySquareCard.blnExistsObjects Then GoTo goEnd:
'''    If mobjSquareCard Is Nothing Then GoTo goEnd:
'''    '    zlBrushCardToDeposit(ByVal lngԤ��ID As Long, ByVal lng���㿨 As Long, ByRef strBlanceInfor As String) As Boolean
'''    '    '------------------------------------------------------------------------------------------------------------------------
'''    '    '���ܣ�ˢ����Ԥ������
'''    '    '��Σ� lngԤ��ID-Ԥ��ID
'''    '    '           lng���㿨-���㿨���
'''    '    '���Σ�strBlanceInfor-����ˢ����Ϣ:
'''    '    '         ��||�ָ�: �ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע
'''    '    '���أ��ɹ�����true,���򷵻�False
'''    '    '���ƣ����˺�
'''    '    '���ڣ�2010-06-18 11:33:22
'''    '    '˵������Ԥ�������������Ԥ��ʱ�����ȷ����ťʱ����(�����е���)
'''    '    '          ����һ��Ҫ������ȷ,�������ֳ������
'''    '    '------------------------------------------------------------------------------------------------------------------------
'''     If mobjSquareCard.zlBrushCardToDeposit(lngԤ��ID, lng�ӿڱ��, strBlanceInfor) = False Then
'''          Exit Function
'''     End If
'''goEnd:
'''    zlSequareBlanceToDeposit = True
'''    Exit Function
'''End Function
 

Private Sub txtסԺ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtסԺ��.Locked = True Then
        glngTXTProc = GetWindowLong(txtסԺ��.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtסԺ��.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtסԺ��_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtסԺ��.Locked = True Then
        Call SetWindowLong(txtסԺ��.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtסԺ��_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If mbytKind = E�������۵Ǽ� Or mbytKind = EסԺ���۵Ǽ� Then Exit Sub
    
    strSQL = "Select ����ID,סԺ��,����,���֤�� From ������Ϣ Where סԺ�� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Trim(txtסԺ��.Text)))
    Call MergePatient(rsTmp, 1)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetMergePatiInfo(lng����ID As Long) As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '��ҳID=0ʱ(����NULL)����ʾԤԼ��Ժ
    strSQL = _
        " Select A.����ID,Decode(B.����ID,NULL,NULL,Nvl(B.��ҳID,0)) as ��ҳID," & _
        " A.����,B.סԺ��,B.��Ժ����,B.��Ժ����" & _
        " From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID(+) And A.����ID=[1]" & _
        " Order by Nvl(B.��ҳID,0)"
    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then Set GetMergePatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub RestoreYB()
    Dim lng����ID As Long, lng����ID As Long
    Dim objCurrent As Object, strTxt As String, arrTxt As Variant
    Dim i As Long, blnDo As Boolean, arrPati As Variant
    Dim objcbo As ComboBox
    
    If (mbytMode = EMode.E����ԤԼ Or mbytMode = EMode.E�����Ǽ� And mlng����ID <> 0) Then
        lng����ID = mlng����ID
    ElseIf Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If MsgBox("��ǰ�Ѿ�����һ������,�Ƿ�Ҫ�Ըò��˵���ݽ�����֤��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                lng����ID = mrsInfo!����ID
            End If
        End If
    End If
    
    'ҽ���Ķ�
    mintInsure = mintInsureBak
    mstrYBPati = mstrYBPatiBak
    If mstrYBPati <> "" Then
        arrPati = Split(mstrYBPati, ";")
        '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID,...
        If UBound(arrPati) >= 8 Then
            If Val(arrPati(8)) > 0 Then
                txtPatient.Locked = txtPatient.Locked
                If mstrYBPati = "" Then txt����.SetFocus: Exit Sub  '������Ϊ��������ѡ�����˳���,������clearcard
            ElseIf mrsInfo Is Nothing Then
                If txtPatient.Tag = "" Then '�����δ����
                    txtPatient.Text = zlDatabase.GetNextNo(1) '�²���ID
                    txtPatient.Tag = txtPatient.Text
                    If txtסԺ��.Visible And mbytKind = EKind.EסԺ��Ժ�Ǽ� Then
                        txtסԺ��.Text = zlDatabase.GetNextNo(2)
                    ElseIf txtסԺ��.Visible And mbytKind = EKind.EסԺ���۵Ǽ� Then
                        txtסԺ��.Text = zlDatabase.GetNextNo(6)
                    End If
                End If
            End If
        End If
        
        txtҽ����.Text = arrPati(1)
        txtҽ����.Locked = True
        
        txt����.Text = arrPati(3)
        cbo�Ա�.ListIndex = GetCboIndex(cbo�Ա�, CStr(arrPati(4)))
        If IsDate(arrPati(5)) Then
            txt��������.Text = Format(arrPati(5), "yyyy-MM-dd")
            Call txt��������_LostFocus
        End If
        txt���֤��.Text = arrPati(6)
        txt������λ.Text = arrPati(7)
       
        '���ղ�����Ϊ��Ժ���
        If UBound(arrPati) >= 14 Then
            If Val(arrPati(14)) > 0 Then
                lng����ID = Val(arrPati(14))
                
                If txt�������.Text = "" And Not RequestCode Then
                    txt�������.Text = Get������(lng����ID)
                End If
            End If
        End If
        
        '��ȡ�����ʻ����
        mcurYBMoney = mcurYBMoneyBak
        lblYBMoney.Caption = "�����ʻ���" & Format(mcurYBMoney, "0.00")
        lblYBMoney.Visible = True
        
        'ҽ�Ƹ��ʽȱʡ=������ҽ�Ʊ���
        For i = 0 To cboҽ�Ƹ���.ListCount
            If InStr(cboҽ�Ƹ���.List(i), Chr(&HD)) > 0 Then cboҽ�Ƹ���.ListIndex = i: Exit For
        Next
       
        If Not IsDate(txt��������.Text) Then
            txt��������.SetFocus
        Else
            strTxt = "txt����,cbo�Ա�,cbo�ѱ�,cbo����,cbo����,cboѧ��,cbo����״��,cboְҵ,cbo���," & _
                     "txt���֤��,txt�����ص�,txt��ͥ��ַ,txt��ͥ��ַ�ʱ�,txt��ͥ�绰,txt���ڵ�ַ,txt���ڵ�ַ�ʱ�,txt������λ,txt��λ�绰,txt��λ�ʱ�," & _
                     "txt��λ������,txt��λ�ʺ�,txt��ϵ������,cbo��ϵ�˹�ϵ,txt��ϵ�˵�ַ,txt��ϵ�˵绰,txt��ϵ�����֤��,txt������,txt������"
            arrTxt = Split(strTxt, ",")
            i = 0
            For i = 0 To UBound(arrTxt)
                For Each objCurrent In Me.Controls
                    If objCurrent.Name = arrTxt(i) Then
                        blnDo = False
                        If TypeOf objCurrent Is TextBox Then
                            If Trim(objCurrent.Text) = "" And objCurrent.Enabled = True Then blnDo = True
                        ElseIf TypeOf objCurrent Is ComboBox Then
                            Set objcbo = objCurrent
                            If objcbo.ListIndex = -1 And objCurrent.Enabled = True Then blnDo = True
                        End If
                        If blnDo Then
                            If objCurrent.TabStop Then
                                Call SetChargeTurn
                                objCurrent.SetFocus
                                Exit Sub
                            End If
                        End If
                        GoTo exitHandle
                    End If
                Next
exitHandle:
            Next
        End If
        Call SetChargeTurn
        If CanFocus(cbo��Ժ����) Then cbo��Ժ����.SetFocus
    Else
        txt����.SetFocus
    End If
End Sub

Private Function GetPatientByName(ByVal strInput As String) As ADODB.Recordset
'���ܣ���ȡ������Ϣ
'˵������ȡʧ��ʱ��mrsInfo = Nothing
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH

    'ͨ������ģ�����Ҳ���(�������벡�˱�ʶʱ)
    strPati = " Select 1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����," & _
        " C.סԺ��,A.�����,A.סԺ����,trunc(C.��Ժ����,'dd') as ��Ժ����,trunc(C.��Ժ����,'dd') as ��Ժ����,A.��������,A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,zl_PatiType(A.����ID) ��������" & _
        " From ������Ϣ A,���ű� B,������ҳ C" & _
        " Where A.ͣ��ʱ�� is NULL And A.����ID=C.����ID(+) And Nvl(A.��ҳID,0)=C.��ҳID(+) And A.��ǰ����ID=B.ID(+) And Rownum<101" & _
        " And A.���� Like [1]" & IIf(gintNameDays = 0, "", " And (A.�Ǽ�ʱ��>Trunc(Sysdate-[2]) Or A.����ʱ��>Trunc(Sysdate-[2]))") & " And A.����ID <> [3] And a.��ҳID Is Not Null And C.��ҳID(+)<>0 "
    strPati = strPati & " Union ALL " & _
                            "Select 0,0,-NULL,'[��ǰ����]',NULL,NULL,-NULL,-NULL,-NULL,To_Date(NULL),To_Date(NULL),To_Date(NULL),NULL,NULL,NULL,NULL,'��ͨ����' From Dual"
    strPati = strPati & " Order by ����ID,����,��Ժ���� Desc"
    
    vRect = GetControlRect(txt����.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txt����.Height, blnCancel, False, True, strInput, gintNameDays, Val(txtPatient.Tag))
                
    'ֻ��һ������ʱ,blncancel����false,��ȡ������Ҳ��һ��
    If Not blnCancel Then
        If rsTmp!ID = 0 Then Exit Function
    Else
        Call zlControl.TxtSelAll(txt����)
        txt����.SetFocus: Exit Function
    End If
    
    Set GetPatientByName = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = Nothing
End Function

Private Sub MergePatient(ByVal rsTmp As ADODB.Recordset, ByVal bytMode As Byte)
    'bytMode = 0 ,ͨ�� cmdName ����, bytMode = 1 ͨ����֤סԺ�� ����
    Dim rsPatiS As ADODB.Recordset
    Dim rsPatiO As ADODB.Recordset
    Dim strSQL As String, Curdate As Date
    Dim i As Integer, j As Integer
    Dim str�ϲ�ԭ�� As String, strInfo As String

    If rsTmp Is Nothing Then Exit Sub
    If mrsInfo Is Nothing And mrsPatiReg Is Nothing Then Exit Sub
    If rsTmp.RecordCount = 0 Then Exit Sub
    If Nvl(rsTmp!����ID, 0) = Val(txtPatient.Text) Then Exit Sub
    If rsTmp!���� = Trim(txt����.Text) Then
        '45976:������,2012-09-21,���֤�Ų�ͬ���������ʾ��
        If Trim(Nvl(rsTmp!���֤��)) <> Trim(txt���֤��.Text) Then
            strInfo = "���������ظ������֤�Ų�ͬ���Ƿ�Ըò��˽��кϲ�?" & vbCrLf & _
                "Ҫ�������˵����֤�ţ�" & Trim(Nvl(rsTmp!���֤��)) & vbCrLf & _
                "Ҫ�ϲ����˵����֤�ţ�" & Trim(txt���֤��.Text)
        Else
            strInfo = "�������������֤���ظ�,�Ƿ�Ըò��˽��кϲ�?"
        End If
        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '���ҽ�������Ƿ����δ�����
            If ExistFeeInsurePatient(Val(txtPatient.Text)) Then
                MsgBox "��ҽ�����˴���δ�����,���Ƚ�����ٺϲ���", vbExclamation, gstrSysName: Exit Sub
            End If

            If ExistFeeInsurePatient(Val(rsTmp!����ID)) Then
                MsgBox "�����ҵ���ҽ�����˴���δ�����,���Ƚ�����ٺϲ���", vbExclamation, gstrSysName: Exit Sub
            End If

            Set rsPatiS = GetMergePatiInfo(Val(txtPatient.Text))
            Set rsPatiO = GetMergePatiInfo(Val(rsTmp!����ID))


            'AB��ס��Ժ
            If Not IsNull(rsPatiS!��ҳID) And Nvl(rsPatiS!��ҳID, 0) <> 0 And Not IsNull(rsPatiO!��ҳID) And Nvl(rsPatiO!��ҳID, 0) <> 0 Then
                '1.��סԺ����Ժ,������(�Ⱥ�סԺ����Ϊ����Ժ-��Ժ,��Ժ-��Ժ����������Ժ-��Ժ,��Ժ-��Ժ)
                '��Ϊ�����˺ϲ���,���򲻶��⴦���Զ���Ժ������Ժ
                rsPatiS.MoveLast
                rsPatiO.MoveLast
                If rsPatiS!��Ժ���� <= rsPatiO!��Ժ���� Then
                    If IsNull(rsPatiS!��Ժ����) Then
                        MsgBox "����:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    If IsNull(rsPatiO!��Ժ����) Then
                        MsgBox "����:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If

                '2.ʱ�佻����ʾ�Ƿ����
                Curdate = zlDatabase.Currentdate
                rsPatiS.MoveFirst
                For i = 1 To rsPatiS.RecordCount
                    rsPatiO.MoveFirst
                    For j = 1 To rsPatiO.RecordCount
                        If Not (rsPatiO!��Ժ���� >= IIf(IsNull(rsPatiS!��Ժ����), Curdate, rsPatiS!��Ժ����) Or _
                            IIf(IsNull(rsPatiO!��Ժ����), Curdate, rsPatiO!��Ժ����) <= rsPatiS!��Ժ����) Then
                            MsgBox "���ֲ���:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]�� " & rsPatiS!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiS!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiS!��Ժ����), Curdate, rsPatiS!��Ժ����), "yyyy-MM-dd") & vbCrLf & _
                            "�벡��:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]�ĵ� " & rsPatiO!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiO!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiO!��Ժ����), Curdate, rsPatiO!��Ժ����), "yyyy-MM-dd") & _
                            vbCrLf & "���ཻ�棬���ܽ��кϲ���", _
                            vbInformation, gstrSysName
                            Exit Sub
                        End If
                        rsPatiO.MoveNext
                    Next
                    rsPatiS.MoveNext
                Next
            End If

            '�ϲ�ԭ��
            str�ϲ�ԭ�� = "[ϵͳԭ��]����ԤԼ��Ժ������Ҫ�����¾ɵ����ϲ���"

            Screen.MousePointer = 11
            DoEvents
            On Error GoTo errHandle
            strSQL = "zl_������Ϣ_MERGE(" & Val(rsPatiS!����ID) & "," & Val(rsPatiO!����ID) & ",'" & str�ϲ�ԭ�� & "','" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
            Screen.MousePointer = 0

            '�ϲ���Ӧֻʣһ������
            strSQL = "Select ����ID From ������Ϣ Where ����ID IN([1],[2])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsPatiS!����ID), Val(rsPatiO!����ID))

            mlng����ID = rsTmp!����ID
            txtPatient.Locked = False
            txtPatient.Text = "-" & mlng����ID
            Call txtPatient_KeyPress(13)
            RestoreYB
        Else
            If bytMode = 1 Then
                txtסԺ��.Locked = (InStr(mstrPrivs, "�޸�סԺ��") = 0)
                If Not mrsInfo Is Nothing Then
                    txtסԺ��.Text = IIf(Nvl(mrsInfo!סԺ��) = "", zlDatabase.GetNextNo(2), Nvl(mrsInfo!סԺ��))
                ElseIf Not mrsPatiReg Is Nothing Then
                    txtסԺ��.Text = IIf(Nvl(mrsPatiReg!סԺ��) = "", zlDatabase.GetNextNo(2), Nvl(mrsPatiReg!סԺ��))
                Else
                    txtסԺ��.Text = zlDatabase.GetNextNo(2)
                End If
            End If
        End If
    Else
        If bytMode = 1 Then
            MsgBox "�������סԺ���ѱ����ˡ�" & rsTmp!���� & "��ռ�ã�", vbInformation, gstrSysName
            txtסԺ��.Locked = (InStr(mstrPrivs, "�޸�סԺ��") = 0)
            If Not mrsInfo Is Nothing Then
                txtסԺ��.Text = IIf(Nvl(mrsInfo!סԺ��) = "", zlDatabase.GetNextNo(2), Nvl(mrsInfo!סԺ��))
            ElseIf Not mrsPatiReg Is Nothing Then
                txtסԺ��.Text = IIf(Nvl(mrsPatiReg!סԺ��) = "", zlDatabase.GetNextNo(2), Nvl(mrsPatiReg!סԺ��))
            Else
                txtסԺ��.Text = zlDatabase.GetNextNo(2)
            End If
        End If
    End If
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function isValid(ByVal lng����ID As Long) As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select ����ID,��ҳID,��������,��Ժ����,��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    While Not rsTmp.EOF
        If Nvl(rsTmp!��������, 0) = 1 And Not IsNull(rsTmp!��Ժ����) And IsNull(rsTmp!��Ժ����) Then
            MsgBox "���������۲�����δ��Ժ�����������ԤԼ��", vbInformation, gstrSysName
            isValid = False
            Exit Function
        End If
        rsTmp.MoveNext
    Wend
    isValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function


'Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:������رս��㿨����
'    '���:blnClosed:�رն���
'    '����:���˺�
'    '����:2010-01-05 14:51:23
'    '����:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strExpend As String
'    '0=����,1=�޸�,2=�鿴
'   If mbytInState = 2 Then Exit Sub
'
'    'ֻ��:ִ�л��˷�ʱ,�ſ��ܹܽ��㿨��
'    If blnClosed Then
'       If Not mobjSquareCard Is Nothing Then
'            Call mobjSquareCard.CloseWindows
'            Set mobjSquareCard = Nothing
'        End If
'        Exit Sub
'    End If
'    '��������
'    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
'    Err = 0: On Error Resume Next
'    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
'    If Err <> 0 Then
'        Err = 0: On Error GoTo 0:      Exit Sub
'    End If
'
'    '��װ�˽��㿨�Ĳ���
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    '����:zlInitComponents (��ʼ���ӿڲ���)
'    '    ByVal frmMain As Object, _
'    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
'    '        ByVal cnOracle As ADODB.Connection, _
'    '        Optional blnDeviceSet As Boolean = False, _
'    '        Optional strExpand As String
'    '����:
'    '����:   True:���óɹ�,False:����ʧ��
'    '����:���˺�
'    '����:2009-12-15 15:16:22
'    'HIS����˵��.
'    '   1.���������շ�ʱ���ñ��ӿ�
'    '   2.����סԺ����ʱ���ñ��ӿ�
'    '   3.����Ԥ����ʱ
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    If mobjSquareCard.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
'         '��ʼ�������ɹ�,����Ϊ�����ڴ���
'         Exit Sub
'    End If
'End Sub


Private Sub InitSendCardPreperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ˢ������
    '����:���˺�
    '����:2011-07-25 11:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strSQL As String, blnBoundCard As Boolean
    Dim rsTemp As ADODB.Recordset, str���� As String, varData As Variant, i As Long
    Dim varTemp  As Variant
    Dim blnNotBind As Boolean
    On Error GoTo errHandle
    
    Set mCurSendCard.rs���� = Nothing
    
    If gbln��Ժ���� = False Then
'        fra�ſ�.Visible = False
'        Me.Height = Me.Height - fra�ſ�.Height
        Exit Sub
    End If
    '76824�����ϴ���2014/8/19��ҽ�ƿ���������
    '85565:���ϴ�,2015/7/27,��������
     lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, 0))
     If lngCardTypeID = 0 Then mCurSendCard.lng�����ID = 0: Exit Sub
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strSQL = "" & _
    "   Select Id, ����, ����, ����, ǰ׺�ı�, ���ų���, ȱʡ��־, �Ƿ�̶�, �Ƿ��ϸ����, " & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����, nvl(�Ƿ�����ʻ�,0) as �Ƿ�����ʻ�, " & _
    "           nvl(�Ƿ�ȫ��,0) as �Ƿ�ȫ��,nvl(�Ƿ��ظ�ʹ��,0) as �Ƿ��ظ�ʹ�� , " & _
    "           nvl(���볤��,10) as ���볤��,nvl(���볤������,0) as ���볤������,nvl(�������,0) as �������," & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����,����, ��ע, �ض���Ŀ, ���㷽ʽ, �Ƿ�����, ��������," & _
    "           nvl(�Ƿ񷢿�,0) as �Ƿ񷢿�,nvl(�Ƿ��ƿ�,0) as �Ƿ��ƿ�,nvl(�Ƿ�д��,0) as �Ƿ�д��, " & _
    "           nvl(��������,0) as ��������,nvl(��������,0) as ��������,nvl(��������,0) as �������� " & _
    "    From ҽ�ƿ���� A" & _
    "    Where nvl(�Ƿ�����,0)=1 And (ID=[1] or nvl(ȱʡ��־,0)=1)" & _
    "    Order by ����"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardTypeID)
    If rsTemp.EOF Then mCurSendCard.lng�����ID = 0: Exit Sub
    If rsTemp.RecordCount >= 2 Then
        rsTemp.Filter = "ID=" & lngCardTypeID
        If rsTemp.EOF Then rsTemp.Filter = 0
    End If
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
        With mCurSendCard
            .lng�����ID = Val(Nvl(rsTemp!ID))
            .str������ = Nvl(rsTemp!����)
            .lng���ų��� = Val(Nvl(rsTemp!���ų���))
            .lng���㷽ʽ = Trim(Nvl(rsTemp!���㷽ʽ))
            .bln���ƿ� = Val(Nvl(rsTemp!�Ƿ�����)) = 1
            .bln�ϸ���� = Val(Nvl(rsTemp!�Ƿ��ϸ����)) = 1
            .bln�ظ����� = Val(Nvl(rsTemp!�Ƿ��ظ�ʹ��)) = 1
            .str�������� = Nvl(rsTemp!��������)
            .int���볤�� = Val(Nvl(rsTemp!���볤��))
            .int���볤������ = Val(Nvl(rsTemp!���볤������))
            .int������� = Val(Nvl(rsTemp!�������))
            .bln���￨ = .str������ = "���￨" And Val(Nvl(rsTemp!�Ƿ�̶�)) = 1
            '�����:56599
            .bln�Ƿ��ƿ� = Val(Nvl(rsTemp!�Ƿ��ƿ�)) = 1
            .bln�Ƿ񷢿� = Val(Nvl(rsTemp!�Ƿ񷢿�)) = 1
            .bln�Ƿ�д�� = Val(Nvl(rsTemp!�Ƿ�д��)) = 1
            .bln�Ƿ�Ժ�ⷢ�� = (InStr(mstrPrivs, ";��������;") > 0 And .bln���ƿ� = False And .bln�Ƿ񷢿� = True) '�����:56599
            .lng�������� = Val(Nvl(rsTemp!��������)) '�����:57326
            .str�������� = Nvl(rsTemp!��������, "1000")
            .byt�������� = Val(Nvl(rsTemp!��������))
            '76824�����ϴ���2014/8/19��ҽ�ƿ���������
            lbl������.Caption = .str������
            lbl������.width = LenB(lbl������.Caption) * 120
            .blnOneCard = False
            .str�ض���Ŀ = Trim(Nvl(rsTemp!�ض���Ŀ))
            If .str�ض���Ŀ <> "" Then
                Set .rs���� = zlGetSpecialItemFee(.str�ض���Ŀ, mstrPriceGrade)
                If .bln���￨ Then .blnOneCard = GetOneCard.RecordCount > 0
            Else
                Set .rs���� = Nothing
            End If
            str���� = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, mlngModul, "0")
            '����ID,�����ID|...
             .lng�������� = 0
            varData = Split(str����, "|")
            For i = 0 To UBound(varData)
                 varTemp = Split(varData(i), ",")
                 If Val(varTemp(0)) <> 0 Then
                    If Val(varTemp(1)) = .lng�����ID Then
                        .lng�������� = Val(varTemp(0)): Exit For
                    End If
                 End If
            Next
           txt����.PasswordChar = IIf(.str�������� <> "", "*", "")
           txt����.MaxLength = .lng���ų���
        End With
    End If
    
    If mCurSendCard.rs���� Is Nothing Then
        tabCardMode.Tabs.Remove ("CardFee")
        blnBoundCard = InStr(mstrPrivs, ";�󶨿���;") > 0
        '�ް󶨿�Ȩ��
        pic�ſ�.Visible = blnBoundCard
        If Not blnBoundCard Then
            Me.Height = Me.Height - pic�ſ�.Height
        Else
            tabCardMode.Tabs("CardBind").Selected = True
            tabCardMode.Tabs("CardBind").Caption = "�󶨿���"
            tabCardMode.width = tabCardMode.width / 2
        End If
        Exit Sub
    End If
    
    blnNotBind = mCurSendCard.bln���ƿ� And (Not mCurSendCard.bln�ظ����� Or mCurSendCard.bln�ϸ����)
    
    Call LoadCardFee
    
    '���û�а󶨿�Ȩ��,���ش���ʱ,�Ѿ��Ƴ��˰󶨿���
    blnBoundCard = Not InStr(mstrPrivs, ";�󶨿���;") > 0
    If Not blnBoundCard Then
        If zlDatabase.GetPara("����ģʽ", glngSys, mlngModul, "CardFee") = "CardFee" Then
            tabCardMode.Tabs("CardFee").Selected = True
        ElseIf Not blnNotBind Then
            tabCardMode.Tabs("CardBind").Selected = True
        End If
    End If
    
 
    '�󶨿�,���û��Ȩ�����ڴ������ʱ,���Ѿ�ɾ��
    '�����:56599
    If (mCurSendCard.bln�Ƿ�Ժ�ⷢ�� Or blnNotBind) And Not blnBoundCard Then
        tabCardMode.Tabs.Remove ("CardBind")
        If tabCardMode.Tabs.Count > 0 Then
            tabCardMode.Tabs("CardFee").Selected = True
            tabCardMode.Tabs("CardFee").Caption = "�շѷ���"
            tabCardMode.width = tabCardMode.width / 2
        Else
            pic�ſ�.Visible = False
            Me.Height = Me.Height - pic�ſ�.Height
        End If
    ElseIf mCurSendCard.bln���ƿ� = False And mCurSendCard.bln�Ƿ񷢿� = False Then
        tabCardMode.Tabs.Remove ("CardFee")
        If tabCardMode.Tabs.Count > 0 Then
            tabCardMode.Tabs("CardBind").Selected = True
            tabCardMode.Tabs("CardBind").Caption = "�󶨿���"
            tabCardMode.width = tabCardMode.width / 2
        Else
            pic�ſ�.Visible = False
            Me.Height = Me.Height - pic�ſ�.Height
        End If
    End If
        
    If mCurSendCard.bln�ϸ���� Then
        '���￨���ü��
        mCurSendCard.lng����ID = CheckUsedBill(5, IIf(mCurSendCard.lng����ID > 0, mCurSendCard.lng����ID, mCurSendCard.lng��������), , mCurSendCard.lng�����ID)
        If mCurSendCard.lng����ID <= 0 Then
            Select Case mCurSendCard.lng����ID
                Case 0 '����ʧ��
                Case -1
'                    MsgBox "��û�����û��õ�" & mCurSendCard.str������ & ",���ܷ��ţ�" & vbCrLf & _
'                        "�����ڱ������ù������λ�����һ���¿�! ", vbExclamation, gstrSysName
                Case -2
'                    MsgBox "���ع��õ�" & mCurSendCard.str������ & "������,���ܷ��ţ�" & vbCrLf & _
'                        "���������ñ��ع��ÿ����λ�����һ���¿���", vbExclamation, gstrSysName
            End Select
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
Private Sub GetFact(Optional blnFirst As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ͬ���ķ�Ʊ
    '����:���˺�
    '����:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gbln��ԺԤ�� = False Then Exit Sub
    
    If gblnPrepayStrict = False Then
        '��ɢ��ȡ��һ������
        txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngModul, "")))
        Exit Sub
    End If
    '�ϸ�:     ȡ��һ������
    mlngԤ������ID = CheckUsedBill(2, IIf(mlngԤ������ID > 0, mlngԤ������ID, mFactProperty.lngShareUseID), , 2)
    If mlngԤ������ID <= 0 Then
        Select Case mlngԤ������ID
            Case 0 '����ʧ��
'            Case -1
'                MsgBox "��û�����û��õ�Ԥ��Ʊ��,�Ǽǲ�����Ϣʱ����ͬʱ��Ԥ���" & _
'                    "��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
'            Case -2
'                MsgBox "���صĹ���Ʊ���Ѿ�����,�Ǽǲ�����Ϣʱ����ͬʱ��Ԥ���" & _
'                    "��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
        End Select
    Else
        txtFact.Text = GetNextBill(mlngԤ������ID)
    End If
End Sub
Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, strTemp As String
    Dim strȱʡԤ���ʽ As String, blnHavePrivs As Boolean
    
    On Error GoTo errHandle
    
    '���㷽ʽ:���ò�ѯ��ҽ�ƿ�����ʱ��һ��ֻ֧��Ԥ����,�����ڴ��յ����
    If mbytMode = 1 Then
        strTemp = "1,2,7,8" 'ԤԼ�Ǽ�ʱ
    Else
        strTemp = "1,2,5,7,8" & IIf(InStr(mstrPrivs, ";���ղ��˵Ǽ�;") > 0, ",3", "")
    End If
    
    strSQL = _
        "Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó��� ='Ԥ����'  And B.����=A.���㷽ʽ  " & _
        "           And Nvl(B.����,1) In(" & strTemp & ")" & _
        " Order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolPrepayPayMode = New Collection
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType: varData = Split(strPayType, ";")
    With cboԤ������
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            
            If Not blnFind And InStr(",7,8,", "," & Nvl(rsTemp!����) & ",") = 0 Then
                .AddItem Nvl(rsTemp!����)
                mcolPrepayPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
                If rsTemp!ȱʡ = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                'If mstrȱʡ���㷽ʽ = Nvl(rsTemp!����) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            '���㷽ʽ���������豸���������˵Ľ��㷽ʽ����Ч
            rsTemp.Filter = "���� ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                If InStr(1, varData(i), "|") <> 0 Then
                    varTemp = Split(varData(i), "|")
                    mcolPrepayPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    'If mstrȱʡ���㷽ʽ = varTemp(1) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                    j = j + 1
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If cboԤ������.ListCount = 0 Then
        MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnload = True: Exit Sub
    End If
    '�����:48352
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    strȱʡԤ���ʽ = zlDatabase.GetPara("ȱʡ�ɿʽ", glngSys, mlngModul, , blnHavePrivs)
    If strȱʡԤ���ʽ <> "" Then
        For i = 0 To cboԤ������.ListCount
            If cboԤ������.List(i) = strȱʡԤ���ʽ Then
                cboԤ������.ListIndex = i
            End If
        Next
    End If
    
    
    strSQL = _
    "Select B.����,B.����,Nvl(B.����,1) as ����,Nvl(A.ȱʡ��־,0) as ȱʡ" & _
    " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
    " Where A.Ӧ�ó��� ='���￨'  And B.����=A.���㷽ʽ  " & _
    "           And Nvl(B.����,1) In(1,2,7,8)" & _
    " Order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolCardPayMode = New Collection
    With cbo��������
        mblnNotClick = True
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            
            If Not blnFind And InStr(",7,8,", "," & Nvl(rsTemp!����) & ",") = 0 Then
                .AddItem Nvl(rsTemp!����)
                mcolCardPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
                If rsTemp!ȱʡ = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                 If cbo��������.List(j) = strȱʡԤ���ʽ Then
                    cbo��������.ListIndex = j:  .Tag = j
                 End If
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            '���㷽ʽ���������豸���������˵Ľ��㷽ʽ����Ч
            rsTemp.Filter = "���� ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                If InStr(1, varData(i), "|") <> 0 Then
                    varTemp = Split(varData(i), "|")
                    mcolCardPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    j = j + 1
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        mblnNotClick = False
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Local���㷽ʽ(ByVal lng�����ID As Long, Optional blnԤ�� As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��λ���㷽ʽ
    '����:���˺�
    '����:2011-07-26 15:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPayMoney As Collection, cboPay As ComboBox
    Dim i As Long
    If mblnNotClick Then Exit Sub
    If blnԤ�� Then
       Set cllPayMoney = mcolPrepayPayMode
        Set cboPay = cboԤ������
    Else
       Set cllPayMoney = mcolCardPayMode
        Set cboPay = cbo��������
    End If
    If cllPayMoney Is Nothing Then Exit Sub
    With cboPay
        If .ListIndex >= 0 Then
            If blnԤ�� Then
                If .ItemData(.ListIndex) >= 0 Then Exit Sub
            End If
        End If
        mblnNotClick = True
        For i = 0 To .ListCount - 1
            ''��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
            If Val(cllPayMoney(i + 1)(3)) = lng�����ID Then
                .ListIndex = i: Exit For
            End If
        Next
        mblnNotClick = False
    End With
End Sub
Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        '58322
        If .State = adStateOpen Then .Close
        .Fields.Append "�շ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
        .ActiveConnection = Nothing
        If mCurPrepay.lngҽ�ƿ����ID <> 0 Then
            .AddNew
            !�շ���� = "Ԥ��"
            !��� = StrToNum(txtԤ����.Text)
            .Update
        End If
        If mCurCardPay.lngҽ�ƿ����ID <> 0 And Trim(txt����) <> "" _
            And cbo��������.Enabled And cbo��������.Visible Then
            .AddNew
            !�շ���� = mCurSendCard.rs����!�շ����
            !��� = StrToNum(txt����.Text)
            .Update
        End If
    End With
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AddCardDataSQL(ByVal lng����ID As Long, ByVal lng��ҳID As Long, lng����ID As Long, lng����ID As Long, ByVal dtCurdate As Date, ByRef strOutSQL As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���￨���Ŵ���
    '���:lng����ID
    '����:���˺�
    '����:2011-07-07 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt�������� As Byte, strno As String, strPassWord As String, strSQL As String
    Dim strԭ���� As String, str���� As String, strCard As String, str�䶯ԭ�� As String
    Dim strICCard As String, lngBrushCardTypeID As Long, str���㷽ʽ As String, strBrushCardNo As String
    Dim bln���ѿ� As Boolean, blnInRange As Boolean   '��Χ�ڵĿ�
    Dim lngIndex As Long, byt�䶯���� As Byte, lng����ID As Long
    
    strCard = UCase(txt����.Text): strICCard = IIf(mblnICCard, strCard, "")
    If Not ((strCard <> "" Or strICCard <> "") And pic�ſ�.Visible = True) Then Exit Sub
    
    '�����:56599
    mbln������󶨿� = True
     
    lng����ID = 0: blnInRange = True
    If mCurSendCard.blnOneCard And mCurSendCard.bln�ϸ���� Then blnInRange = mCurSendCard.lng����ID > 0
    
    If blnInRange And tabCardMode.SelectedItem.Key = "CardFee" Then
        blnInRange = True
        byt�������� = 0: byt�䶯���� = 1
    Else
        blnInRange = False
        byt�䶯���� = 11: byt�������� = 0
    End If
    str�䶯ԭ�� = "������Ժ�ǼǷ���"
    strPassWord = zlCommFun.zlStringEncode(Trim(txtPass.Text))
    If blnInRange = False Then
          'Zl_ҽ�ƿ��䶯_Insert
           strSQL = "Zl_ҽ�ƿ��䶯_Insert("
          '      �䶯����_In   Number,
          '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
          strSQL = strSQL & "" & byt�䶯���� & ","
          '      ����id_In     סԺ���ü�¼.����id%Type,
          strSQL = strSQL & "" & lng����ID & ","
          '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
          strSQL = strSQL & "" & mCurSendCard.lng�����ID & ","
          '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'" & strԭ���� & "',"
          '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
          '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
          strSQL = strSQL & "'" & str�䶯ԭ�� & "',"
          '      ����_In       ������Ϣ.����֤��%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
          strSQL = strSQL & "'" & UserInfo.���� & "',"
          '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic����_In     ������Ϣ.Ic����%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
          strSQL = strSQL & "NULL)"
    Else
        '103980:���ϴ�,2017/1/19,���淢����������
        str���� = Trim(txt����.Text)
        If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text

        strno = zlDatabase.GetNextNo(16)  'ҽ�ƿ�
        If chk����.Value = 0 Then
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        End If
        mCurCardPay.strno = strno
        mCurCardPay.lng����ID = lng����ID
        strSQL = zlGetSaveCardFeeSQL(mCurSendCard.lng�����ID, byt��������, strno, lng����ID, lng��ҳID, lng����ID, lng����ID, Val(txtסԺ��.Text), _
         zlCommFun.GetNeedName(cbo�ѱ�.Text), "", Trim(txt����.Text), zlCommFun.GetNeedName(cbo�Ա�.Text), str����, _
        strCard, strPassWord, str�䶯ԭ��, IIf(mCurSendCard.bln��� = False, mCurSendCard.dblӦ�ս��, Val(txt����.Text)), Val(txt����.Text), IIf(chk����.Value = 0, mCurCardPay.str���㷽ʽ, ""), _
        dtCurdate, mCurSendCard.lng����ID, mCurSendCard.rs����, strICCard, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, lng����ID)
    End If
    strOutSQL = strSQL
 End Sub
 
 Private Sub AddDepositSQL(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal dtDate As Date, ByRef bln�����ʻ���Ԥ�� As Boolean, strOutSQL As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ�����SQL
    '����:���˺�
    '����:2011-07-26 18:26:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strno As String, strSQL As String, i As Integer, lngԤ��ID As Long
    Dim dblMoney As Double
     If Not (IsNumeric(txtԤ����.Text) And fraԤ��.Visible) Then Exit Sub
     
    '����Ԥ�����¼
    strno = zlDatabase.GetNextNo(11)
    lngԤ��ID = zlDatabase.GetNextId("����Ԥ����¼")
    mCurPrepay.strno = strno
    mCurPrepay.lngID = lngԤ��ID
    dblMoney = StrToNum(txtԤ����.Text)
    bln�����ʻ���Ԥ�� = is�����ʻ�(cboԤ������) And mintInsure <> 0 And mstrYBPati <> "" And mbytMode <> 1
    
    'Zl_����Ԥ����¼_Insert
    strSQL = "Zl_����Ԥ����¼_Insert("
    '  Id_In         ����Ԥ����¼.ID%Type,
    strSQL = strSQL & "" & lngԤ��ID & ","
    '  ���ݺ�_In     ����Ԥ����¼.NO%Type,
    strSQL = strSQL & "'" & strno & "',"
    '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "" & IIf(mblnPrepayPrint, "'" & txtFact.Text & "'", "Null") & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
    strSQL = strSQL & "" & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & IIf(lng����ID = 0, "NULL", lng����ID) & ","
    '  ���_In       ����Ԥ����¼.���%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & mCurPrepay.str���㷽ʽ & "',"
    '  �������_In   ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "'" & txt�������.Text & "',"
    '  �ɿλ_In   ����Ԥ����¼.�ɿλ%Type,
    If bln�����ʻ���Ԥ�� Then
        strSQL = strSQL & "'" & mintInsure & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt�ɿλ.Text) & "',"
    End If
    '  ��λ������_In ����Ԥ����¼.��λ������%Type,
    If bln�����ʻ���Ԥ�� Then
        strSQL = strSQL & "'" & Split(mstrYBPati, ";")(2) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt������.Text) & "',"
    End If
    '  ��λ�ʺ�_In   ����Ԥ����¼.��λ�ʺ�%Type,
    If bln�����ʻ���Ԥ�� Then
        strSQL = strSQL & "'" & Split(mstrYBPati, ";")(1) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt�ʺ�.Text) & "',"
    End If
    '  ժҪ_In       ����Ԥ����¼.ժҪ%Type,
    strSQL = strSQL & "'��ԺԤ��',"
    '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(mlngԤ������ID = 0, "NULL", mlngԤ������ID) & ","
    '  Ԥ�����_In   ����Ԥ����¼.Ԥ�����%Type := Null,
    strSQL = strSQL & "" & 2 & ","
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPrepay.lngҽ�ƿ����ID = 0 Or mCurPrepay.bln���ѿ�, "NULL", mCurPrepay.lngҽ�ƿ����ID) & ","
   '  ���㿨���_in ����Ԥ����¼.���㿨���%type:=NULL,
    strSQL = strSQL & "" & IIf(mCurPrepay.lngҽ�ƿ����ID = 0 Or Not mCurPrepay.bln���ѿ�, "NULL", mCurPrepay.lngҽ�ƿ����ID) & ","
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPrepay.strˢ������ = "", "NULL", "'" & mCurPrepay.strˢ������ & "'") & ","
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  ������λ_In   ����Ԥ����¼.������λ%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null
    '108001:���ϴ���2017/5/8����ʽ��Ԥ��ʱ��Ϊ24Сʱ��
    strSQL = strSQL & "to_date('" & Format(dtDate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '   ��������_In Integer:=0 :0-������Ԥ��;1-��Ϊ���۵�
    strSQL = strSQL & "0 )"
   strOutSQL = strSQL
End Sub
Private Function CheckBrushCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str���� As String
    Dim dblMoney As Double, bln�������� As Boolean
    Dim dblThreeMoney As Double, tyCurThreePay As Ty_PayMoney
    
    Dim strˢ������ As String, strˢ������ As String
    Dim blnTemp As Boolean
    
    On Error GoTo errHandle
    
    dblMoney = 0: dblThreeMoney = 0
    '58322
    If cboԤ������.Visible Then
        If cboԤ������.ListIndex >= 0 Then
            bln�������� = cboԤ������.ItemData(cboԤ������.ListIndex) = -1
            If bln�������� Then dblThreeMoney = dblThreeMoney + StrToNum(txtԤ����.Text)
        End If
        dblMoney = dblMoney + StrToNum(txtԤ����.Text)
    End If
    If cbo��������.Visible And cbo��������.Enabled And Trim(txt����) <> "" Then
        If cbo��������.ListIndex >= 0 Then
            blnTemp = cbo��������.ItemData(cbo��������.ListIndex) = -1
            If blnTemp Then dblThreeMoney = dblThreeMoney + StrToNum(txt����.Text)
            bln�������� = bln�������� Or blnTemp
        End If
        dblMoney = dblMoney + StrToNum(txt����.Text)
    End If
    If Not bln�������� Then CheckBrushCard = True: Exit Function
    If mCurPrepay.lngҽ�ƿ����ID <> 0 Then
       tyCurThreePay = mCurPrepay
    Else
       tyCurThreePay = mCurCardPay
    End If
    
    If (mCurPrepay.lngҽ�ƿ����ID <> mCurCardPay.lngҽ�ƿ����ID Or _
        mCurPrepay.bln���ѿ� <> mCurCardPay.bln���ѿ�) _
        And mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurPrepay.lngҽ�ƿ����ID <> 0 Then
        MsgBox "����ͬʱʹ�����ֲ�ͬ����֧����ʽ,���ܼ���!", vbOKOnly + vbInformation, gstrSysName
        If cboԤ������.Enabled And cboԤ������.Visible Then cboԤ������.SetFocus: Exit Function
        If cbo��������.Enabled And cbo��������.Visible Then cbo��������.SetFocus
        Exit Function
    End If
    Call zlGetClassMoney(rsMoney)
    
     '����ˢ������
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln���ѿ� As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl��� As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String) As Boolean
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
   '58322
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, tyCurThreePay.lngҽ�ƿ����ID, tyCurThreePay.bln���ѿ�, _
    txt����.Text, zlCommFun.GetNeedName(cbo�Ա�.Text), str����, dblThreeMoney, tyCurThreePay.strˢ������, tyCurThreePay.strˢ������, False, True, False) = False Then Exit Function
    
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, tyCurThreePay.lngҽ�ƿ����ID, _
    tyCurThreePay.bln���ѿ�, tyCurThreePay.strˢ������, dblThreeMoney, "", "") = False Then Exit Function
    mCurCardPay.strˢ������ = tyCurThreePay.strˢ������
    mCurCardPay.strˢ������ = tyCurThreePay.strˢ������
    mCurPrepay.strˢ������ = tyCurThreePay.strˢ������
    mCurPrepay.strˢ������ = tyCurThreePay.strˢ������
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlInterfacePrayMoney(ByRef cllPro As Collection, ByRef cllThreeSwap As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�֧�����
    '����:cllPro-�޸�������������
    '        cll��������-����������������
    '����:֧���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim dblMoney As Double
    If mCurCardPay.lngҽ�ƿ����ID = 0 And mCurPrepay.lngҽ�ƿ����ID = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cbo��������.ItemData(cbo��������.ListIndex) <> -1 _
        And cboԤ������.ItemData(cboԤ������.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ��
    '���:frmMain-���õ�������
    '        lngModule-����ģ���
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-֧�����
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�����ID As Long, bln���ѿ� As Boolean, strCardNO As String
    
    dblMoney = 0
    If mCurCardPay.lngҽ�ƿ����ID <> 0 And cbo��������.Enabled And cbo��������.Visible Then
        dblMoney = Val(txt����.Text)
        lng�����ID = mCurCardPay.lngҽ�ƿ����ID
        bln���ѿ� = mCurCardPay.bln���ѿ�
        strCardNO = mCurCardPay.strˢ������
    End If
    If mCurPrepay.lngҽ�ƿ����ID <> 0 And cboԤ������.Visible Then
        dblMoney = dblMoney + StrToNum(txtԤ����.Text)
        If lng�����ID <> mCurPrepay.lngҽ�ƿ����ID And lng�����ID <> 0 Then
            MsgBox "������ѡ���֧����ʽ��Ԥ������ѡ���֧����ʽ��һ��!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        lng�����ID = mCurPrepay.lngҽ�ƿ����ID
        bln���ѿ� = mCurPrepay.bln���ѿ�
        strCardNO = mCurPrepay.strˢ������
    End If
    If lng�����ID = 0 Then Exit Function


    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, lng�����ID, bln���ѿ�, strCardNO, mCurCardPay.lng����ID, mCurPrepay.strno, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '����������������
     If mCurCardPay.lngҽ�ƿ����ID <> 0 And mCurCardPay.lng����ID <> 0 And cbo��������.Visible Then
        If Not mCurCardPay.bln���ѿ� Then
            Call zlAddUpdateSwapSQL(False, mCurCardPay.lng����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, mCurCardPay.lng����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap)
    End If
    If mCurPrepay.lngҽ�ƿ����ID <> 0 And cboԤ������.Visible And StrToNum(txtԤ����.Text) <> 0 Then
        Call zlAddUpdateSwapSQL(True, mCurPrepay.lngID, mCurPrepay.lngҽ�ƿ����ID, mCurPrepay.bln���ѿ�, mCurPrepay.strˢ������, strSwapGlideNO, strSwapMemo, cllPro)
        Call zlAddThreeSwapSQLToCollection(True, mCurPrepay.lngID, mCurPrepay.lngҽ�ƿ����ID, mCurPrepay.bln���ѿ�, mCurPrepay.strˢ������, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Led��ӭ��Ϣ()
    Dim strInfo As String, lngPatient As Long
    'LED��ʼ��
    If gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
        End If
        strInfo = Trim(txtPatient.Text)
        If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!�Ա� & " " & mrsInfo!����: lngPatient = Val("" & mrsInfo!����ID)
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub

Private Function zl_Get����Ĭ�Ϸ�������() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ĭ�Ϸ�������
    '����:�Ƿ������������
    '����:����
    '����:2012-07-06 15:53:14
    '�����:51072
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardType As clsCard
    Dim msgResult As VbMsgBoxResult
    Dim arr() As String
    arr = zl_Getҽ�ƿ�����(mCurSendCard.lng�����ID)
    If Val(arr(2)) = 0 Then '������
        Select Case Val(arr(1))
            Case 0 '������
                zl_Get����Ĭ�Ϸ������� = True
                Exit Function
            Case 1 'δ��������
               msgResult = MsgBox("δ�������뽫��Ӱ���ʻ���ʹ�ð�ȫ,�Ƿ������", vbQuestion + vbYesNo, gstrSysName)
               zl_Get����Ĭ�Ϸ������� = IIf(msgResult = vbYes, True, False)
               Exit Function
            Case 2 'Ϊ�����ֹ
                 MsgBox "δ���뿨����,���ܽ��з�����", vbExclamation, gstrSysName
                zl_Get����Ĭ�Ϸ������� = False
                Exit Function
        End Select
    ElseIf Val(arr(2)) = 1 Then 'ȱʡ���֤��Nλ
        If Len(Trim(txt���֤��.Text)) > 0 Or Len(Trim(txt��ϵ�����֤��.Text)) > 0 Then '���������֤����ϵ�����֤��
            If Len(Trim(txt���֤��.Text)) > 0 Then '�����֤���������֤
                   txtPass.Text = Right(Trim(txt���֤��.Text), Val(arr(0)))
            Else '������ô��������֤��Ϊ����
                   txtPass.Text = Right(Trim(txt��ϵ�����֤��.Text), Val(arr(0)))
            End If
        Else '���֤����ϵ�����֤��û����
            Select Case Val(arr(1))
                Case 0 '������
                    zl_Get����Ĭ�Ϸ������� = True
                    Exit Function
                Case 1 'δ��������
                    msgResult = MsgBox("δ�������뽫��Ӱ���ʻ���ʹ�ð�ȫ,�Ƿ������", vbQuestion + vbYesNo, gstrSysName)
                    zl_Get����Ĭ�Ϸ������� = IIf(msgResult = vbYes, True, False)
                    Exit Function
                Case 2 'Ϊ�����ֹ
                    MsgBox "δ���뿨����,���ܽ��з�����", vbExclamation, gstrSysName
                    zl_Get����Ĭ�Ϸ������� = False
                    Exit Function
            End Select
        End If
    End If
    zl_Get����Ĭ�Ϸ������� = True
End Function
Private Function zl�����֤(colPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ĭ�Ϸ�������
    '����:�Ƿ������������
    '����:����
    '����:2012-07-06 15:53:14
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Trim(txt֧������.Text) <> Trim(txt��֤����.Text) Then
        MsgBox "������������벻һ��,����������", vbOKOnly + vbInformation, gstrSysName
        txt֧������.Text = "": txt��֤����.Text = ""
        If txt֧������.Visible = True Then txt֧������.SetFocus
        Exit Function
    End If
    If Trim(txt֧������.Text) <> "" Then
       If �Ƿ��Ѿ�ǩԼ(Trim(txt���֤��.Text)) Then
             MsgBox "���֤����Ϊ:" & txt���֤��.Text & "�Ѿ�ǩԼ�����ظ�ǩԼ��", vbOKOnly + vbInformation, gstrSysName
             txt֧������.Text = "": txt��֤����.Text = ""
             If txt֧������.Visible = True Then txt֧������.SetFocus
             Exit Function
       End If
    End If
    AddSQL�󶨿� Trim(txtPatient.Text), Getҽ�ƿ����ID("�������֤"), Trim(txt���֤��.Text), zlCommFun.zlStringEncode(Trim(txt֧������.Text)), zlDatabase.Currentdate, False, colPro
    
    zl�����֤ = True
End Function
Private Sub InitTabPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ҳ�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
        
    Set objItem = tbcPage.InsertItem(1, "����", PicBaseInfo.hWnd, 0)
    objItem.Tag = mPageHeight.����
    
    Set objItem = tbcPage.InsertItem(2, "��������", PicHealth.hWnd, 0)
    objItem.Tag = mPageHeight.��������
    
    PicBaseInfo.Enabled = False
    PicHealth.Enabled = False
    With tbcPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = False
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .Item(0).Selected = True
    End With
    PicBaseInfo.Enabled = True
    PicHealth.Enabled = True

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetColumHeader(ByRef vsGrid As VSFlexGrid, ByVal strHead As String, Optional ByVal lngNo As Long = 0)
    '���ܣ���ʼvsFlexGrid
    '           ��һ�̶��У���ʼ����ֻ��һ�м�¼���޹̶��С�
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'vsGrid:    Ҫ��ʼ���Ŀؼ�

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1
        If lngNo = 0 Then
            .FixedCols = 0
            .Cols = .FixedCols + UBound(arrHead) + 1
            .Rows = .FixedRows + 1
        Else
            .FixedCols = 1
            .Cols = .FixedCols + UBound(arrHead)
            .Rows = .FixedRows + 1
        End If

        For i = 0 To UBound(arrHead)
            If .FixedCols > 0 Then
                .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            Else
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            End If
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               'Ϊ��֧��zl9PrintMode
                If .FixedCols > 0 Then
                    .ColHidden(i) = Val(Split(arrHead(i), ",")(3)) = 0
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(2))
                    .colAlignment(i) = Val(Split(arrHead(i), ",")(1))
                    .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = Val(Split(arrHead(i), ",")(1))
                Else
                    .ColHidden(.FixedCols + i) = Val(Split(arrHead(i), ",")(3)) = 0
                    .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                    .colAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
'                    .ColData
                    'Ϊ��֧��zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  'Ϊ��֧��zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
                End If
            End If
        Next
        
        '�̶������־���
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
    End With
End Sub

Private Sub ComboBox(objcbo As ComboBox, strSet As String)
    Dim varTemp As Variant
    Dim i As Long
    varTemp = Split(strSet, ",")
    With objcbo
        For i = LBound(varTemp) To UBound(varTemp)
            .AddItem varTemp(i)
        Next
    End With
    If objcbo.ListCount <> 0 Then objcbo.ListIndex = 0
End Sub
Private Sub InitCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ComBox�ؼ�
    '����:56599
    '����:2012-12-07 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '68192:������,2013-12-02,Ѫ�Ͷ�ȡ�����ֵ䡢RHȱʡĬ��ֵΪ��
    Call ReadDict("Ѫ��", cboBloodType)
    ComboBox cboBH, C_BH
    If cboBH.ListCount > 0 Then cboBH.ListIndex = -1
End Sub

Private Sub InitVsOtherInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo ErrHand
    
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ����ϵ Order by ����"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "����ϵ")
    With rsTemp
        Do While Not rsTemp.EOF
            strTmp = strTmp & "|" & Nvl(rsTemp!����)
        rsTemp.MoveNext
        Loop
    End With
    If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
    
    With vsLinkMan
        '��ʼ���б�����
        SetColumHeader vsLinkMan, C_LinkManColumHeader
        .Editable = IIf(mbytInState = 2, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
        If strTmp <> "" Then .ColComboList(.ColIndex("��ϵ�˹�ϵ")) = strTmp
    End With
    
    With vsOtherInfo
        '������ͷ
        SetColumHeader vsOtherInfo, C_OtherInfoColumHeader
        .Editable = IIf(mbytInState = 2, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InitvsDrug()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsDrug
        '��ʼ���б�����
        SetColumHeader vsDrug, C_ColumHeader
        .Editable = IIf(mbytInState = 2, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
        .ColComboList(0) = "..."
    End With
End Sub

Private Sub InitVsInoculate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsInoculate
        '��ʼ���б�����
        SetColumHeader vsInoculate, C_InoculateHeader
         vsInoculate.Editable = IIf(mbytInState = 2, flexEDNone, flexEDKbdMouse)
        '����ѡ��ť
        .ColDataType(0) = flexDTDate
        .ColEditMask(0) = "####-##-##"
        .ColDataType(2) = flexDTDate
        .ColEditMask(2) = "####-##-##"
        .SelectionMode = flexSelectionFree
    End With

End Sub

Private Sub txtת��_GotFocus()
    zlControl.TxtSelAll txtת��
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtת��_KeyPress(KeyAscii As Integer)
    Dim vPoint As POINTAPI
    On Error GoTo errH
    If KeyAscii = 13 Then
        KeyAscii = 0
        vPoint = GetCoordPos(txtת��.Container.hWnd, txtת��.Left, txtת��.Top)
        Call GetSpcҽ�ƻ���(txtת��, Me, "ҽ�ƻ���", False, False, False, True, vPoint)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtת��_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtת��_Validate(Cancel As Boolean)
    Dim vPoint As POINTAPI
    vPoint = GetCoordPos(txtת��.Container.hWnd, txtת��.Left, txtת��.Top)
    Call GetSpcҽ�ƻ���(txtת��, Me, "ҽ�ƻ���", False, False, False, True, vPoint)
    Exit Sub
End Sub

Private Sub vsCertificate_GotFocus()
    If mblnCheckPatiCard = False Then
        vsCertificate.Col = vsCertificate.FixedCols
        vsCertificate.Row = vsCertificate.FixedRows
    End If
    mblnCheckPatiCard = False
End Sub

Private Sub vsDrug_EnterCell()
    If vsDrug.Col = vsDrug.FixedCols Then
        vsDrug.ColComboList(vsDrug.Col) = "..."
    End If
End Sub

Private Sub vsDrug_GotFocus()
    If mblnCheckPatiCard = False Then
        vsDrug.Col = vsDrug.FixedCols
        vsDrug.Row = vsDrug.FixedRows
    End If
    mblnCheckPatiCard = False
End Sub

Private Sub vsDrug_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsDrug_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim strInput As String, strFilter As String
    Dim rsTemp As Recordset
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandl
    
    If Not Col = vsDrug.FixedCols Then Exit Sub

    strInput = Trim(vsDrug.EditText)
    
    If strInput <> "" Then
        If zlCommFun.IsCharAlpha(strInput) Then
            strFilter = " And zlspellcode(A.����) like [1]"
            strInput = UCase(strInput)
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strFilter = " And A.���� like [1]"
        Else
            strFilter = " And A.���� like [1]"
        End If
    End If
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select Distinct A.ID,A.����," & _
        " A.����,zlspellcode(A.����) ����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & strFilter

    '��ȡ��ǰ�������ֵ
    vRect = GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    strInput = gstrLike & Trim(strInput) & "%"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҩ��", False, "����ҩ��ѡ����", "��������ҩƷ��ѡ��һ����Ϊ���˹���ҩ��", False, False, True, vRect.Left, vRect.Top, 0, blnCancel, False, True, strInput)

    If Not rsTemp Is Nothing And blnCancel = False Then
        If rsTemp.RecordCount > 0 Then
            vsDrug.EditText = Nvl(rsTemp!����)
            vsDrug.TextMatrix(Row, Col) = Nvl(rsTemp!����)
            vsDrug.TextMatrix(Row, 2) = Nvl(rsTemp!ID)
            If vsDrug.Rows - 1 = Row Then vsDrug.Rows = vsDrug.Rows + 1
        End If
    Else
        vsDrug.EditText = vsDrug.TextMatrix(Row, Col)
    End If
    
    Exit Sub
ErrHandl:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



Private Sub vsInoculate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col <> 1 And Col <> 3 Then
        If vsInoculate.TextMatrix(Row, Col) = "____-__-__" Then vsInoculate.TextMatrix(Row, Col) = ""
    End If
End Sub

Private Sub vsInoculate_GotFocus()
    If mblnCheckPatiCard = False Then
        vsInoculate.Col = vsInoculate.FixedCols
        vsInoculate.Row = vsInoculate.FixedRows
    End If
    mblnCheckPatiCard = False
End Sub

Private Sub VsInoculate_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����:56599
    Dim intRow As Integer
    
    With vsInoculate
        If KeyCode = vbKeyDelete And .Row >= .FixedRows And mbytInState <> 2 Then
            intRow = .Row
            If .Col > .FixedCols + 1 Then
                .TextMatrix(intRow, .FixedCols + 2) = ""
                .TextMatrix(intRow, .FixedCols + 3) = ""
            Else
                If .Rows = .FixedRows + 1 Then
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                Else
                    Call .RemoveItem(.Row)
                    If intRow >= .Rows Then
                        .Row = .Rows - 1
                    Else
                        .Row = intRow
                    End If
                    .Col = .FixedCols
                End If
            End If
        ElseIf KeyCode = vbKeyReturn And .Row >= .FixedRows Then
            If ((.TextMatrix(.Row, .FixedCols) = "" And .Col = .FixedCols) Or (.Col = .FixedCols + 2 And .TextMatrix(.Row, .FixedCols + 2) = "") Or .Col = .FixedCols + 3) And .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
               Call MoveNextCell(vsInoculate)
            End If
        End If
    End With
End Sub

Private Sub vsDrug_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�����:56599
    If Col = 1 Then  '������Ӧ�б༭ʱ���ж��Ƿ�����������100
        With vsDrug
           If LenB(StrConv(.TextMatrix(Row, Col), vbFromUnicode)) > 100 Then
                MsgBox "������Ӧ�����ַ���������ַ���100,������ַ������Զ��س���", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = StrConv(MidB(StrConv(.TextMatrix(Row, Col), vbFromUnicode), 1, 100), vbUnicode)
           End If
        End With
    End If
End Sub

Private Sub vsDrug_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '�����:56599
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim rsTemp As Recordset
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandl:
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select ID,nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����,NULL ����," & _
        " NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ��" & _
        " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Union All" & _
        " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����," & _
        " A.����,zlspellcode(A.����) ����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"

    '��ȡ��ǰ�������ֵ
    vRect = GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����ҩ��", False, "����ҩ��ѡ����", "��������ҩƷ��ѡ��һ����Ϊ���˹���ҩ��", False, False, True, vRect.Left, vRect.Top, 0, blnCancel, False, True)

    If Not rsTemp Is Nothing And blnCancel = False Then
        vsDrug.TextMatrix(Row, Col) = rsTemp!����
        vsDrug.TextMatrix(Row, 2) = rsTemp!ID
        If vsDrug.Rows - 1 = Row Then vsDrug.Rows = vsDrug.Rows + 1
    End If
    Exit Sub
ErrHandl:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����:56599
    Dim intRow As Integer
    With vsDrug
        If KeyCode <> vbKeyReturn And KeyCode <> vbKeyDelete And .ColComboList(.Col) = "..." Then
            .ColComboList(.Col) = ""
        End If
        If KeyCode = vbKeyDelete And .Row >= .FixedRows And mbytInState <> 2 Then
            intRow = .Row
            If .Rows = .FixedRows + 1 Then
                vsDrug.TextMatrix(1, 0) = "": vsDrug.Cell(flexcpData, 1, 0) = "": vsDrug.TextMatrix(1, 1) = ""
            Else
                Call vsDrug.RemoveItem(.Row)
                If intRow >= .Rows Then
                    .Row = .Rows - 1
                Else
                    .Row = intRow
                End If
            End If
        ElseIf KeyCode = vbKeyReturn And .Row >= .FixedRows Then
            If .TextMatrix(.Row, .FixedCols) = "" And .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                Call MoveNextCell(vsDrug)
            End If
        End If
    End With
End Sub

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function
Private Sub cmdMedicalWarning_Click()
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim strTemp As String
    
    vRect = GetControlRect(txtMedicalWarning.hWnd)
    strSQL = "" & _
    "       Select ���� as ID,����,���� From ҽѧ��ʾ Where ���� Not Like '����%'"
    Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ҽѧ��ʾ", False, txtMedicalWarning.Text, "", False, False, False, vRect.Left, vRect.Top - 180, 500, True, False, True)
    If Not rsTemp Is Nothing Then
        While rsTemp.EOF = False
          strTemp = strTemp & ";" & rsTemp!����
          rsTemp.MoveNext
        Wend
    Else
        If cmdMedicalWarning.Enabled And cmdMedicalWarning.Visible Then cmdMedicalWarning.SetFocus: Exit Sub
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then txtMedicalWarning.Text = strTemp
    If txtOtherWaring.Enabled And txtOtherWaring.Visible Then txtOtherWaring.SetFocus
End Sub

Private Sub SetDrugAllergy(str����ҩ�� As String, str������Ӧ As String, Optional lng����ID = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���ҩ��
    '����:56599
    '����:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long

    With vsDrug
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str����ҩ�� Then
                    .TextMatrix(i, 1) = str������Ӧ
                    If lng����ID <> 0 Then .TextMatrix(i, 2) = lng����ID
                    Exit Sub
                End If

            Next
        End If
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str����ҩ��
        .TextMatrix(.Rows - 1, 1) = str������Ӧ
        If lng����ID <> 0 Then .TextMatrix(.Rows - 1, 2) = lng����ID
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetInoculate(str�������� As String, str�������� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ý������
    '����:56599
    '����:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    '68192:������,2013-12-02
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If Format(.TextMatrix(i, j - 1), "YYYY-MM-DD") = Format(str��������, "YYYY-MM-DD") Then
                        .TextMatrix(i, j) = str��������
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = Format(str��������, "YYYY-MM-DD")
                .TextMatrix(.Rows - 1, j + 1) = str��������
                If j = 2 Then .Rows = .Rows + 1
                Exit Sub
            End If
        Next
    End With
End Sub
Private Sub SetLinkInfo(str���� As String, str��ϵ As String, str�绰 As String, str���֤�� As String, Optional str������Ϣ As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϵ�������Ϣ
    '����:56599
    '����:2012-12-12 09:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    With vsLinkMan
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("��ϵ������")) = str���� And .TextMatrix(i, .ColIndex("��ϵ�����֤��")) = str���֤�� Then
                    .TextMatrix(i, .ColIndex("��ϵ�˹�ϵ")) = str��ϵ: .TextMatrix(i, .ColIndex("��ϵ�˵绰")) = str�绰
                    If i = 1 Then
                        txt��ϵ�����֤��.Text = str���֤��
                        txt��ϵ������.Text = str����
                        For j = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                            If zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.List(j)) = str��ϵ Then cbo��ϵ�˹�ϵ.ListIndex = j
                        Next
                        txt��ϵ�˵绰.Text = str�绰
                        txtLinkManInfo.Text = str������Ϣ
                    End If
                    Exit Sub
                End If
            Next
        End If
        
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("��ϵ������")) = str����
        .TextMatrix(.Rows - 1, .ColIndex("��ϵ�˹�ϵ")) = str��ϵ
        .TextMatrix(.Rows - 1, .ColIndex("��ϵ�˵绰")) = str�绰
        .TextMatrix(.Rows - 1, .ColIndex("��ϵ�����֤��")) = str���֤��
        .TextMatrix(.Rows - 1, .ColIndex("��ϵ�˹�ϵ��ע")) = str������Ϣ
        If .Rows - 1 = 1 Then
            txt��ϵ�����֤��.Text = str���֤��
            txt��ϵ������.Text = str����
            For j = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                If zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.List(j)) = str��ϵ Then cbo��ϵ�˹�ϵ.ListIndex = j
            Next
            txt��ϵ�˵绰.Text = str�绰
            txtLinkManInfo.Text = str������Ϣ
        End If
        .Rows = .Rows + 1
    End With
End Sub

Private Sub SetOtherInfo(str��Ϣ�� As String, str��Ϣֵ As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:56599
    '����:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str��Ϣ�� Then
                        .TextMatrix(i, j + 1) = str��Ϣֵ
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str��Ϣ��
                .TextMatrix(.Rows - 1, j + 1) = str��Ϣֵ
                If j = 2 Then .Rows = .Rows + 1
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub Load�����������Ϣ(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز��˽�������Ϣ
    '���:lng����ID - ����ID
    '����:56599
    '����:2012-12-12 14:55:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs����ҩ�� As Recordset
    Dim rs���߼�¼ As Recordset
    Dim rsABOѪ�� As Recordset
    Dim rsRH As Recordset
    Dim rsҽѧ��ʾ As Recordset
    Dim rs����ҽѧ��ʾ As Recordset
    Dim rs������Ϣ As Recordset
    Dim rs��ϵ�� As Recordset
    Dim rs������Ϣ As Recordset
    Dim strҽѧ��ʾ As String
    Dim str��ϵ������ As String
    Dim str��ϵ�˹�ϵ As String
    Dim str��ϵ�˸�����Ϣ As String
    Dim str��ϵ�˵绰 As String
    Dim str��ϵ�����֤�� As String
    Dim lng��ϵ������ As Long
    Dim i As Long
    On Error GoTo ErrHandl:

    '��ȡ����ҩ��
    strSQL = "" & _
    "   Select ����ID,����ҩ��ID,����ҩ��,������Ӧ From ���˹���ҩ�� Where ����ID=[1]"
    Set rs����ҩ�� = zlDatabase.OpenSQLRecord(strSQL, "���˹���ҩ��", lng����ID)
    While rs����ҩ��.EOF = False
        SetDrugAllergy Nvl(rs����ҩ��!����ҩ��), Nvl(rs����ҩ��!������Ӧ), Nvl(rs����ҩ��!����ҩ��ID, 0)
        rs����ҩ��.MoveNext
    Wend
    '��ȡ���߼�¼
    strSQL = "" & _
    "   Select ����ID,����ʱ��,�������� From �������߼�¼ Where ����ID=[1]"
    Set rs���߼�¼ = zlDatabase.OpenSQLRecord(strSQL, "�������߼�¼", lng����ID)
    While rs���߼�¼.EOF = False
        SetInoculate Format(Nvl(rs���߼�¼!����ʱ��), "YYYY-MM-DD"), Nvl(rs���߼�¼!��������)
        rs���߼�¼.MoveNext
    Wend
    'Ѫ��
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='Ѫ��'"
    Set rsABOѪ�� = zlDatabase.OpenSQLRecord(strSQL, "ABOѪ��", lng����ID)
    While rsABOѪ��.EOF = False
        For i = 0 To cboBloodType.ListCount - 1
            If cboBloodType.List(i) = Nvl(rsABOѪ��!��Ϣֵ) Then cboBloodType.ListIndex = i
        Next
        rsABOѪ��.MoveNext
    Wend
    'RH
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='RH'"
    Set rsRH = zlDatabase.OpenSQLRecord(strSQL, "RH", lng����ID)
    While rsRH.EOF = False
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = Nvl(rsRH!��Ϣֵ) Then cboBH.ListIndex = i
        Next
        rsRH.MoveNext
    Wend
    'ҽѧ��ʾ
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='ҽѧ��ʾ' "
    Set rsҽѧ��ʾ = zlDatabase.OpenSQLRecord(strSQL, "ҽѧ��ʾ", lng����ID)
    While rsҽѧ��ʾ.EOF = False
        strҽѧ��ʾ = strҽѧ��ʾ & ";" & Nvl(rsҽѧ��ʾ!��Ϣֵ)
        rsҽѧ��ʾ.MoveNext
    Wend
    If strҽѧ��ʾ <> "" Then strҽѧ��ʾ = Mid(strҽѧ��ʾ, 2)
    txtMedicalWarning.Text = strҽѧ��ʾ
    txtMedicalWarning.Tag = txtMedicalWarning.Text
    '����ҽѧ��ʾ
    strSQL = "" & _
    "  Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ��='����ҽѧ��ʾ' "
    Set rs����ҽѧ��ʾ = zlDatabase.OpenSQLRecord(strSQL, "����ҽѧ��ʾ", lng����ID)
    While rs����ҽѧ��ʾ.EOF = False
        txtOtherWaring.Text = Nvl(rs����ҽѧ��ʾ!��Ϣֵ)
        rs����ҽѧ��ʾ.MoveNext
    Wend
    '��ϵ�������Ϣ
    'ȡ������Ϣ���е���ϵ����Ϣ
    strSQL = "Select A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵绰, A.��ϵ�����֤��, B.��Ϣֵ As ��ϵ�˸�����Ϣ" & vbNewLine & _
            "From ������Ϣ A, ������Ϣ�ӱ� B" & vbNewLine & _
            "Where a.����id = b.����id(+) And a.����id = [1] And Not a.��ϵ������ Is Null And b.��Ϣ��(+) = '��ϵ�˸�����Ϣ'"
    Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "������Ϣ��ϵ����Ϣ", lng����ID)
        If rs������Ϣ.EOF = False Then
        txt��ϵ�����֤��.Text = Nvl(rs������Ϣ!��ϵ�����֤��)
        txt��ϵ������.Text = Nvl(rs������Ϣ!��ϵ������)
        For i = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
            If zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.List(i)) = Nvl(rs������Ϣ!��ϵ�˹�ϵ) Then cbo��ϵ�˹�ϵ.ListIndex = i
        Next
        txt��ϵ�˵绰.Text = Nvl(rs������Ϣ!��ϵ�˵绰)
        txtLinkManInfo.Text = Nvl(rs������Ϣ!��ϵ�˸�����Ϣ)
        
        SetLinkInfo Nvl(rs������Ϣ!��ϵ������), Nvl(rs������Ϣ!��ϵ�˹�ϵ), Nvl(rs������Ϣ!��ϵ�˵绰), Nvl(rs������Ϣ!��ϵ�����֤��), Nvl(rs������Ϣ!��ϵ�˸�����Ϣ)
    End If
    'ȡ������Ϣ�ӱ��е���ϵ����Ϣ
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ�� like '��ϵ��%' order by ��Ϣ�� Asc"
    Set rs��ϵ�� = zlDatabase.OpenSQLRecord(strSQL, "��ϵ�������Ϣ", lng����ID)
    If rs��ϵ��.EOF = False Then
        rs��ϵ��.Filter = "��Ϣ�� like '��ϵ������%'"
        lng��ϵ������ = rs��ϵ��.RecordCount
        rs��ϵ��.Filter = ""
        For i = 2 To lng��ϵ������ + 1
            While rs��ϵ��.EOF = False
                Select Case Nvl(rs��ϵ��!��Ϣ��)
                    Case "��ϵ������" & i
                        str��ϵ������ = Nvl(rs��ϵ��!��Ϣֵ)
                    Case "��ϵ�˹�ϵ" & i
                        str��ϵ�˹�ϵ = Nvl(rs��ϵ��!��Ϣֵ)
                    Case "��ϵ�˸�����Ϣ" & i
                        str��ϵ�˸�����Ϣ = Nvl(rs��ϵ��!��Ϣֵ)
                    Case "��ϵ�˵绰" & i
                        str��ϵ�˵绰 = Nvl(rs��ϵ��!��Ϣֵ)
                    Case "��ϵ�����֤��" & i
                        str��ϵ�����֤�� = Nvl(rs��ϵ��!��Ϣֵ)
                End Select
                rs��ϵ��.MoveNext
            Wend
            SetLinkInfo str��ϵ������, str��ϵ�˹�ϵ, str��ϵ�˵绰, str��ϵ�����֤��, str��ϵ�˸�����Ϣ
            rs��ϵ��.MoveFirst
        Next
    End If
    '������Ϣ
    strSQL = "" & _
    "   Select ����ID,����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ��Ϣ�� Not in ('Ѫ��','RH','ҽѧ��ʾ','����ҽѧ��ʾ','���֤��״̬','�⼮���֤��') And ��Ϣ�� Not like '��ϵ��%'"
    Set rs������Ϣ = zlDatabase.OpenSQLRecord(strSQL, "��ϵ��������Ϣ", lng����ID)
    '�����:115886,����,2017/11/24,�Һ���ȡ�ò�����Ϣʱ�����򱨴�
    While rs������Ϣ.EOF = False
        If Nvl(rs������Ϣ!��Ϣ��) <> "" Then
            SetOtherInfo Nvl(rs������Ϣ!��Ϣ��), Nvl(rs������Ϣ!��Ϣֵ)
        End If
        rs������Ϣ.MoveNext
    Wend
    
    '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
    Call LoadCertificate(lng����ID)
    Exit Sub
ErrHandl:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Add�����������Ϣ(ByVal lng����ID As Long, ByRef colPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ݴ���
    '���:
    '����:56599
    '����:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim varKey As Variant
    '����ҩ��
    With vsDrug
        If .Rows > 1 Then
            '����ò������м�¼
            strSQL = " Zl_���˹���ҩ��_Delete(" & lng����ID & ")"
            zlAddArray colPro, strSQL
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '���˹���ҩ��
                    strSQL = "Zl_���˹���ҩ��_Update("
                    '����ID_In ���˹���ҩ��.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '����ҩ��ID_In ���˹���ҩ��.����ҩ��ID%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 2) = "", "", .TextMatrix(i, 2)) & "',"
                    '����ҩ��_In  ���˹���ҩ��.����ҩ��%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '������Ӧ_In ���˹�����Ӧ.������Ӧ%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '������Ϣ
    With vsInoculate
        If .Rows > 1 Then
            '����ò������м�¼
            strSQL = " Zl_�������߼�¼_Delete(" & lng����ID & ")"
            zlAddArray colPro, strSQL
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    '���˹���ҩ��
                    strSQL = "Zl_�������߼�¼_Update("
                    '����ID_In �������߼�¼.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '����ʱ��_In �������߼�¼.����ʱ��%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 0) = "", "''", "to_date('" & .TextMatrix(i, 0) & "','yyyy-mm-dd')") & ","
                    '��������_In  �������߼�¼.��������%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"
                    zlAddArray colPro, strSQL
                    Debug.Print strSQL
                End If
                If .TextMatrix(i, 3) <> "" Then
                    '���˹���ҩ��
                    strSQL = "Zl_�������߼�¼_Update("
                    '����ID_In �������߼�¼.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '����ʱ��_In �������߼�¼.����ʱ��%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 2) = "", "''", "to_date('" & .TextMatrix(i, 2) & "','yyyy-mm-dd')") & ","
                    '��������_In  �������߼�¼.��������%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "''", .TextMatrix(i, 3)) & "')"
                    zlAddArray colPro, strSQL
                    Debug.Print strSQL
                End If
            Next
        End If
    End With
    '������Ϣ
    'ABOѪ��
    '������Ϣ�ӱ�
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'Ѫ��',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & cboBloodType.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'RH
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'RH',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & cboBH.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'ҽѧ��ʾ
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'ҽѧ��ʾ',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & txtMedicalWarning.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '����ҽѧ��ʾ
    strSQL = "Zl_������Ϣ�ӱ�_Update("
    '����ID_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "" & lng����ID & ","
    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
    strSQL = strSQL & "'����ҽѧ��ʾ',"
    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
    strSQL = strSQL & "'" & txtOtherWaring.Text & "',"
    '����Id_In ������Ϣ�ӱ�.����Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    
    '��ϵ�������Ϣ
    '��ϵ�������Ϣ
    With vsLinkMan
        If .Rows > 2 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then '��ϵ����������Ϊ��
                    For j = 0 To .Cols - 1
                        strSQL = "Zl_������Ϣ�ӱ�_Update("
                        '����ID_In ������Ϣ�ӱ�.����Id%Type
                        strSQL = strSQL & "" & lng����ID & ","
                        '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                        strSQL = strSQL & "'" & .TextMatrix(0, j) & i & "',"
                        '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                        strSQL = strSQL & "'" & IIf(.TextMatrix(i, j) = "", "", .TextMatrix(i, j)) & "',"
                        '����Id_In ������Ϣ�ӱ�.����Id%Type
                        strSQL = strSQL & "'')"

                        zlAddArray colPro, strSQL
                    Next
                End If
            Next
        End If
    End With
    '������Ϣ
     With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    strSQL = "Zl_������Ϣ�ӱ�_Update("
                    '����ID_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 0) & "',"
                    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "',"
                    '����Id_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 2) <> "" Then
                    strSQL = "Zl_������Ϣ�ӱ�_Update("
                    '����ID_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "" & lng����ID & ","
                    '��Ϣ��_In ������Ϣ�ӱ�.��Ϣ��%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 2) & "',"
                    '��Ϣֵ_In ������Ϣ�ӱ�.��Ϣֵ%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "", .TextMatrix(i, 3)) & "',"
                    '����Id_In ������Ϣ�ӱ�.����Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
     End With
     'ҽ�ƿ�����
     If Not mdicҽ�ƿ����� Is Nothing And Trim(txt����.Text) <> "" Then
        For Each varKey In mdicҽ�ƿ�����.Keys
            strSQL = "Zl_����ҽ�ƿ�����_Update("
            strSQL = strSQL & lng����ID & ","
            strSQL = strSQL & mCurSendCard.lng�����ID & ","
            strSQL = strSQL & "'" & Trim(txt����.Text) & "',"
            strSQL = strSQL & "'" & varKey & "',"
            strSQL = strSQL & "'" & mdicҽ�ƿ�����(varKey) & "')"
            zlAddArray colPro, strSQL
        Next
     End If
End Sub

Private Function CheckPatiCard() As Boolean
'���ܣ���鲡�˽�����Ƭ¼��������Ƿ�Ϸ�
'63246:������,2013-07-03
    Dim intLen As Integer, i As Integer, j As Integer
    
    intLen = 100
    If LenB(StrConv(txtMedicalWarning.Text, vbFromUnicode)) > intLen Then
        tbcPage.Item(1).Selected = True
        MsgBox "ҽѧ��ʾֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�", vbInformation, gstrSysName
        If txtMedicalWarning.Enabled And txtMedicalWarning.Visible Then txtMedicalWarning.SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtOtherWaring.Text, vbFromUnicode)) > intLen Then
        tbcPage.Item(1).Selected = True
        MsgBox "����ҽѧ��ʾֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�", vbInformation, gstrSysName
        If txtOtherWaring.Enabled And txtOtherWaring.Visible Then txtOtherWaring.SetFocus
        Exit Function
    End If
    
    mblnCheckPatiCard = True
    '����ҩ��
    With vsDrug
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    intLen = 60
                    If LenB(StrConv(.TextMatrix(i, 0), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "����ҩ��ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���1��", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "������Ӧֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���2��", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
    
    '������Ϣ
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    If Not IsDate(.TextMatrix(i, 0)) Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "����ʱ�䲻����Ч�����ڸ�ʽ��" & vbCrLf & "������:��" & i & "�С���1��", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    
                    intLen = 200
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��������ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���2��", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
                If .TextMatrix(i, 3) <> "" Then
                    If Not IsDate(.TextMatrix(i, 2)) Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "����ʱ�䲻����Ч�����ڸ�ʽ��" & vbCrLf & "������:��" & i & "�С���3��", vbInformation, gstrSysName
                        Call .Select(i, 2, i, 2)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    
                    intLen = 200
                    If LenB(StrConv(.TextMatrix(i, 3), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��������ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���4��", vbInformation, gstrSysName
                        Call .Select(i, 3, i, 3)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
    
    '��ϵ�˵�ַ
    With vsLinkMan
        intLen = 100
        If .Rows > 2 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    For j = 0 To .Cols - 1
                        If .ColIndex("��ϵ������") = j Then
                            intLen = 64
                        ElseIf .ColIndex("��ϵ�����֤��") = j Then
                            intLen = 18
                        ElseIf .ColIndex("��ϵ�˵绰") = j Then
                            intLen = 20
                        Else
                            intLen = 100
                        End If
                        If LenB(StrConv(.TextMatrix(i, j), vbFromUnicode)) > intLen Then
                            tbcPage.Item(1).Selected = True
                            MsgBox .TextMatrix(0, j) & "ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "��", vbInformation, gstrSysName
                            Call .Select(i, j, i, j)
                            .TopRow = i
                            If .Enabled = True And .Visible = True Then .SetFocus
                            Exit Function
                        End If
                    Next
                End If
            Next
        End If
    End With
    
    '������Ϣ
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    intLen = 20
                    If LenB(StrConv(.TextMatrix(i, 0), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��Ϣ��ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���1��", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��Ϣֵֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���2��", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
                If .TextMatrix(i, 2) <> "" Then
                    intLen = 20
                    If LenB(StrConv(.TextMatrix(i, 2), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��Ϣ��ֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���3��", vbInformation, gstrSysName
                        Call .Select(i, 2, i, 2)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 3), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "��Ϣֵֻ�������� " & intLen & " ���ַ��� " & intLen \ 2 & " �����֣�" & vbCrLf & "������:��" & i & "�С���4��", vbInformation, gstrSysName
                        Call .Select(i, 3, i, 3)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
     End With
     
     mblnCheckPatiCard = False
     tbcPage.Item(0).Selected = True
     CheckPatiCard = True
End Function

Private Function LoadPati(ByVal strPatiXML As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ,��ȡ������Ϣ
    '����:���˺�
    '����:2011-09-08 21:52:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long '�����:56599
    Dim str����ҩ�� As String, str������Ӧ As String '�����:56599
    Dim str�������� As String, str�������� As String '�����:56599
    Dim strABOѪ�� As String '�����:56599
    Dim str��Ϣ�� As String, str��Ϣֵ As String '�����:56599
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode '�����:56599
    Dim str���� As String, str��ϵ As String, str�绰 As String, str���֤�� As String, str��ַ As String '�����:56599
    On Error GoTo errHandle

    If strPatiXML = "" Then Exit Function
    
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strPatiXML, False) = False Then Exit Function
    '    ��ʶ    ��������    ����    ����    ˵��
    '    ����    Varchar2    20
    Call zlXML_GetNodeValue("����", , strValue)
    '    ����    Varchar2    64
    Call zlXML_GetNodeValue("����", , strValue)
    txtPatient.Text = strValue
    '    �Ա�    Varchar2    4
    Call zlXML_GetNodeValue("�Ա�", , strValue)
    If strValue <> "" Then
        Call cbo.Locate(cbo�Ա�, strValue)
        If cbo�Ա�.ListIndex = -1 Then
            cbo�Ա�.AddItem strValue
            cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
        End If
    End If
    '    ����    Varchar2    10
    Call zlXML_GetNodeValue("����", , strValue)
    If strValue <> "" Then
        Call LoadOldData(strValue, txt����, cbo���䵥λ)
    End If
    '    ��������    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("��������", , strValue)
    
    txt��������.Text = Format(IIf(strValue = "", "____-__-__", strValue), "YYYY-MM-DD")
    If strValue <> "" Then
        txt����.Text = ReCalcOld(CDate(Format(strValue, "YYYY-MM-DD HH:MM:SS")), cbo���䵥λ, , , CDate(txt��Ժʱ��.Text))   '�޸ĵ�ʱ��,���ݳ���������������
        If CDate(txt��������.Text) - CDate(strValue) <> 0 Then
            mblnChange = False
            txt����ʱ��.Text = Format(strValue, "HH:MM")
            mblnChange = True
        End If
    Else
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
    '    �����ص�    Varchar2    50
    Call zlXML_GetNodeValue("�����ص�", , strValue)
    '    ���֤��    VARCHAR2    18
    Call zlXML_GetNodeValue("���֤��", , strValue)
    If strValue <> "" Then txt���֤��.Text = strValue
    '    ����֤��    Varchar2    20
    Call zlXML_GetNodeValue("����֤��", , strValue)
    If strValue <> "" Then txt����֤��.Text = strValue
    '    ְҵ    Varchar2    80
    Call zlXML_GetNodeValue("ְҵ", , strValue)
    If strValue <> "" Then
        cboְҵ.ListIndex = GetCboIndex(cboְҵ, strValue, , , mstrCboSplit)
        If cboְҵ.ListIndex = -1 Then
            cboְҵ.AddItem strValue, 0
            cboְҵ.ListIndex = cboְҵ.NewIndex
        End If
    End If
    '    ����    Varchar2    20
    Call zlXML_GetNodeValue("����", , strValue)
    cbo����.ListIndex = GetCboIndex(cbo����, strValue)
     If cbo����.ListIndex = -1 And strValue <> "" Then
         cbo����.AddItem strValue, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue)
    cbo����.ListIndex = GetCboIndex(cbo����, strValue)
     If cbo����.ListIndex = -1 And strValue <> "" Then
         cbo����.AddItem strValue, 0
         cbo����.ListIndex = cbo����.NewIndex
     End If
    '    ѧ��    Varchar2    10
    Call zlXML_GetNodeValue("ѧ��", , strValue)
    cboѧ��.ListIndex = GetCboIndex(cboѧ��, strValue)
    If cboѧ��.ListIndex = -1 And strValue <> "" Then
        cboѧ��.AddItem strValue, 0
        cboѧ��.ListIndex = cboѧ��.NewIndex
    End If
    '    ����״��    Varchar2    4
    Call zlXML_GetNodeValue("����״��", , strValue)
    cbo����״��.ListIndex = GetCboIndex(cbo����״��, strValue)
     If cbo����״��.ListIndex = -1 And strValue <> "" Then
         cbo����״��.AddItem strValue, 0
         cbo����״��.ListIndex = cbo����״��.NewIndex
     End If
    '    ����    Varchar2    30
    Call zlXML_GetNodeValue("����", , strValue)
    txt����.Text = strValue
    '    ��ͥ��ַ    Varchar2    50
    Call zlXML_GetNodeValue("��ͥ��ַ", , strValue)
    txt��ͥ��ַ.Text = strValue
    
        '    ���ڵ�ַ    Varchar2    50
    Call zlXML_GetNodeValue("���ڵ�ַ", , strValue)
    txt���ڵ�ַ.Text = strValue
    If gbln���ýṹ����ַ Then PatiAddress(E_IX_���ڵ�ַ).Value = strValue
    
    If gbln���ýṹ����ַ Then PatiAddress(E_IX_��סַ).Value = strValue
    '    ��ͥ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��ͥ�绰", , strValue)
    txt��ͥ�绰.Text = strValue
    '    ��ͥ��ַ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��ͥ��ַ�ʱ�", , strValue)
    txt��ͥ��ַ�ʱ�.Text = strValue
    '    ������λ    Varchar2    100
    Call zlXML_GetNodeValue("������λ", , strValue)
    txt������λ.Text = strValue
    lbl������λ.Tag = ""
    '    ��λ�绰    Varchar2    20
    Call zlXML_GetNodeValue("��λ�绰", , strValue)
    txt��λ�绰.Text = strValue
    '�ֻ���
    Call zlXML_GetNodeValue("�ֻ���", , strValue)
    txtMobile.Text = strValue
    '    ��λ�ʱ�    Varchar2    6
    Call zlXML_GetNodeValue("��λ�ʱ�", , strValue)
   txt��λ�ʱ�.Text = strValue
    '    ��λ������  Varchar2    50
    Call zlXML_GetNodeValue("��λ������", , strValue)
   txt��λ������.Text = strValue
    '    ��λ�ʺ�    Varchar2    50
    Call zlXML_GetNodeValue("��λ�ʺ�", , strValue)
   txt��λ�ʺ�.Text = strValue
    '�����:56599
    '�������
    Call zlXML_GetRows("ҩ������", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("ҩ������", i, str����ҩ��)
        Call zlXML_GetNodeValue("ҩ�ﷴӦ", i, str������Ӧ)
        SetDrugAllergy str����ҩ��, str������Ӧ
    Next
    lngCount = 0
    '���߼�¼
    Call zlXML_GetRows("��������", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("��������", i, str��������)
        Call zlXML_GetNodeValue("����ʱ��", i, str��������)
        SetInoculate str��������, str��������
    Next
    lngCount = 0
    'ABOѪ��
    Call zlXML_GetNodeValue("ABOѪ��", , strABOѪ��)
    If strABOѪ�� <> "" Then
        For i = 0 To cboBloodType.ListCount - 1
            If cboBloodType.List(i) = strABOѪ�� Then cboBloodType.ListIndex = i
        Next
    End If
    'RH
    Call zlXML_GetNodeValue("RH", , strValue)
    If strValue <> "" Then
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = strValue Then cboBH.ListIndex = i
        Next
    End If
    'ҽѧ��ʾ
    strValue = ""
    Set xmlChildNodes = zlXML_GetChildNodes("�ٴ�������Ϣ")
    If Not xmlChildNodes Is Nothing Then
        If xmlChildNodes.length > 0 Then
            For i = 0 To xmlChildNodes.length - 1
                Set xmlChildNode = xmlChildNodes(i)
                If xmlChildNode.Text = "1" Then
                    strValue = strValue & ";" & Replace(xmlChildNode.nodeName, "��־", "")
                End If
            Next
        End If
    End If
    If strValue <> "" Then txtMedicalWarning.Text = Mid(strValue, 2)
   
    
    '����ҽѧ��ʾ
    Call zlXML_GetNodeValue("����ҽѧ��ʾ", , strValue)
    If strValue <> "" Then txtOtherWaring.Text = strValue
    '��ϵ��Ϣ
    '    ��ϵ�˵�ַ  Varchar2    50
    Call zlXML_GetNodeValue("��ϵ�˵�ַ", , str��ַ)
    txt��ϵ�˵�ַ.Text = str��ַ
    If gbln���ýṹ����ַ Then PatiAddress(E_IX_��ϵ�˵�ַ).Value = str��ַ
     '    ��ϵ������  Varchar2    64
    Call zlXML_GetNodeValue("��ϵ������", , str����)
    '    ��ϵ�˹�ϵ  Varchar2    30
    Call zlXML_GetNodeValue("��ϵ�˹�ϵ", , str��ϵ)
    '    ��ϵ�˵绰  Varchar2    20
    Call zlXML_GetNodeValue("��ϵ�˵绰", , str�绰)
    '    ��ϵ�����֤ Varchar2   20
    Call zlXML_GetNodeValue("��ϵ�����֤��", , str���֤��)
    SetLinkInfo str����, str��ϵ, str�绰, str���֤��
    
    Call zlXML_GetRows("��ϵ��Ϣ", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("��ϵ��Ϣ", "����", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "����", i, j, str����)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "��ϵ", i, j, str��ϵ)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "�绰", i, j, str�绰)
                Call zlXML_GetChildNodeValue("��ϵ��Ϣ", "���֤��", i, j, str���֤��)
                SetLinkInfo str����, str��ϵ, str�绰, str���֤��
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '������Ϣ
    '�����������
    Call zlXML_GetNodeValue("�����������", , strValue)
    SetOtherInfo "�����������", strValue
    
    '��ũ��֤��
    Call zlXML_GetNodeValue("��ũ��֤��", , strValue)
    SetOtherInfo "��ũ��֤��", strValue

    '����֤��
    Call zlXML_GetRows("����֤��", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("����֤��", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("����֤��", "��Ϣֵ", i, j, str��Ϣֵ)
                SetOtherInfo str��Ϣ��, str��Ϣֵ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '������Ϣ
    Call zlXML_GetRows("������Ϣ", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("������Ϣ", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("������Ϣ", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("������Ϣ", "��Ϣֵ", i, j, str��Ϣֵ)
                SetOtherInfo str��Ϣ��, str��Ϣֵ
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    'ҽ�ƿ�����
    Call zlXML_GetRows("ҽ�ƿ�����", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("ҽ�ƿ�����", "��Ϣ��", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("ҽ�ƿ�����", "��Ϣ��", i, j, str��Ϣ��)
                Call zlXML_GetChildNodeValue("ҽ�ƿ�����", "��Ϣֵ", i, j, str��Ϣֵ)
                If mdicҽ�ƿ�����.Exists(str��Ϣ��) Then
                    mdicҽ�ƿ�����.Item(str��Ϣ��) = str��Ϣֵ
                Else
                    mdicҽ�ƿ�����.Add str��Ϣ��, str��Ϣֵ
                End If
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    
    LoadPati = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetCboIndex(cbo As ComboBox, strFind As String, _
    Optional blnKeep As Boolean, _
    Optional blnLike As Boolean, Optional strSplit As String = "-") As Long
'���ܣ����ַ�����ComboBox�в�������
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '�Ⱦ�ȷ����
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), strSplit) > 0 Then
            If zlCommFun.GetNeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '���ģ������
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Public Sub Clear��������()
    '---------------------------------------------------------------------------------------------------------------------------------------------
'����:�жϵ�ǰ�Ƿ�Ϊ�������� (���Ƿ����������ǰ󶨿�����)
'���:
'����:56599
'����:2012-12-25 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    'Ѫ��
    Call SetCboDefault(cboBloodType)
    'RH
    If cboBH.ListCount > 0 Then cboBH.ListIndex = -1
    'ҽѧ��ʾ
    txtMedicalWarning.Text = ""
    '����ҽѧ��ʾ
    txtOtherWaring.Text = ""
    '��ϵ����Ϣ
    With vsLinkMan
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
    End With
    '�������
    With vsInoculate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '������Ϣ
    With vsOtherInfo
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""

    End With
    '֤����Ϣ
    With vsCertificate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '������Ӧ
    With vsDrug
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
    End With
End Sub

Public Function zlCreateSquare() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ�ƿ�����
    '����:���ϴ�
    '����:2016/6/21 11:57:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If Not mobjSquare Is Nothing Then zlCreateSquare = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjSquare = CreateObject("zl9CardSquare.clsCardsquare")
    If Err <> 0 Then Err = 0: Exit Function
    Call mobjSquare.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend)
    '��ʼ�������ɹ�,����Ϊ�����ڴ���
    zlCreateSquare = True
End Function

Private Function WriteCard(lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:д��
    '���:lng����ID - ����ID
    '����:����
    '����:56599
    '����:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    If mobjSquare Is Nothing Then
       If Not zlCreateSquare() Then Exit Function
    End If
    If mobjSquare Is Nothing Then Exit Function
    WriteCard = mobjSquare.zlBandCardArfter(Me, mlngModul, mCurSendCard.lng�����ID, lng����ID, strExpend)
    Exit Function
ErrHandl:
    WriteCard = False
    If ErrCenter() = 1 Then Resume
End Function
Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim lngTop As Long
    '�����:56599
    Select Case Item.Caption
        Case "����"
            Me.Height = mPageHeight.����
            pic��Ժ.Top = pic����.Top + pic����.Height
            lngTop = pic��Ժ.Top + pic��Ժ.Height
            If mbln�Ƿ���ʾԤ�� Then
                picԤ��.Top = lngTop
                lngTop = picԤ��.Top + picԤ��.Height
            End If
            If mbytInState = 1 Or (mbytInState = 0 And mbytMode = 2 And mbytKind <> EסԺ���۵Ǽ�) Then
                If txtסԺ��.Enabled And txtסԺ��.Visible Then txtסԺ��.SetFocus
            ElseIf mbytInState = 0 And mbytMode = 2 And mbytKind = EסԺ���۵Ǽ� Then
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Else
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            End If
            pic�ſ�.Top = lngTop
        Case "��������"
            Me.Height = mPageHeight.��������
            If cboBloodType.Enabled And cboBloodType.Visible Then cboBloodType.SetFocus
    End Select
    tbcPage.Height = picCmd.Top
    tbcPage.width = Me.width - 90
    Call SetCenter(Me)
End Sub

Private Sub SetCardEditEnabled()
    '���þ��￨�༭����
    Dim blnEdit As Boolean
    blnEdit = Trim(txt����.Text) <> ""
    
    txtPass.Enabled = blnEdit: txtAudi.Enabled = blnEdit
    lbl����.Enabled = txtPass.Enabled: lbl��֤.Enabled = blnEdit
    
    txt����.Enabled = blnEdit: lbl���.Enabled = blnEdit
    chk����.Enabled = blnEdit
    cbo��������.Enabled = chk����.Value = 0 And blnEdit
End Sub
Private Function Check��������(lng����ID As Long, lng�����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ����Ƿ����Ʋ��˵ķ�������
    '���:lng����ID - ����ID;lng�����ID  - ҽ�ƿ������ID
    '����:����
    '����:57326
    '����:2013-01-30 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
        strSQL = "Select Count(1) as ���� From ����ҽ�ƿ���Ϣ Where ״̬=0 And ����ID=[1] And �����ID=[2] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng�����ID)
        If Val(Nvl(rsTemp!����)) <= 0 Then Check�������� = True: Exit Function
        Select Case mCurSendCard.lng��������
            Case 0 '������
                Check�������� = True
            Case 1 'ͬһ������ֻ����һ�ſ�
                MsgBox "�ò����Ѿ�����" & mCurSendCard.str������ & ",�����ڽ��з�������!", vbInformation + vbOKOnly
                Check�������� = False
            Case 2 'ͬһ�������������ſ�,����Ҫ����
               Check�������� = MsgBox("�ò����Ѿ�����" & mCurSendCard.str������ & ",�Ƿ�Ҫ���з�������?", vbQuestion + vbYesNo) = vbYes
        End Select
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function WhetherTheCardBinding(ByVal str���� As String, ByVal lng����� As Long, Optional ByRef lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:��ȡָ�������Ƿ��Ѿ�����
'���:str���ţ����� ��lng����𣺿���� , lngPatientID :����ID
'����:True :�Ѿ�����;False:δ����
'����:
'����:
'�����:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl
    strSQL = "" & _
           "   Select ����ID From ����ҽ�ƿ���Ϣ Where ����=[1]  And �����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Һ�", str����, lng�����)
    WhetherTheCardBinding = rsTemp.RecordCount > 0

    If rsTemp.RecordCount > 0 Then
        lngPatientID = Val(Nvl(rsTemp!����ID))
    End If

    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function GetCardLastChangeType(ByVal str���� As String, ByVal lng����� As Long, ByVal lngPaitentID As Long) As Long
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:��ȡ�����ı䶯����
'���:str���ţ����� ��lng����𣺿���� , lngPatientID :����ID
'����:0-δ�ҵ������Ϣ   1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
'����:��⸣
'����:2013-2-4 17:36:33
'�����:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "     Select �䶯���" & vbNewLine & _
           "    From (With ҽ�ƿ��䶯 As (Select ����id, ID, �䶯���, �䶯ʱ�� " & vbNewLine & _
           "                              From ����ҽ�ƿ��䶯 Bd" & vbNewLine & _
           "                              Where Bd.���� = [2] And �����id = [1] And ����id = [3])" & vbNewLine & _
           "           Select A.�䶯���" & vbNewLine & _
           "           From ҽ�ƿ��䶯 A, (Select Max(�䶯ʱ��) As �䶯ʱ�� From ҽ�ƿ��䶯 C) B" & vbNewLine & _
           "           Where A.�䶯ʱ�� = B.�䶯ʱ��) A"
    On Error GoTo ErrHand
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����䶯��Ϣ", lng�����, str����, lngPaitentID)
    If Not rsTmp Is Nothing Then
        If rsTmp.RecordCount > 0 Then
            GetCardLastChangeType = Val(Nvl(rsTmp!�䶯���))
        End If
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Private Function BlandCancel(ByVal lngCardTypeID As Long, ByVal strCardNO As String, ByVal lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:ȡ���󶨿�
'���:intType:0-��ǰ����;1-��ǰ���;2-��ǰ��������
'����:ȡ���ɹ�,����true,���򷵻�False
'����:���˺�
'����:2011-07-29 11:18:05
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Curdate As Date
    Dim strSQL As String, strPassWord As String

    On Error GoTo errHandle

    Curdate = zlDatabase.Currentdate

    'Zl_ҽ�ƿ��䶯_Insert
    strSQL = "Zl_ҽ�ƿ��䶯_Insert("
    '      �䶯����_In   Number,
    '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
    strSQL = strSQL & "" & 14 & ","
    '      ����id_In     סԺ���ü�¼.����id%Type,
    strSQL = strSQL & "" & lngPatientID & ","
    '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
    strSQL = strSQL & "" & lngCardTypeID & ","
    '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "NULL,"
    '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & strCardNO & "'" & ","
    '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
    strSQL = strSQL & "'���ظ������Զ�ȡ��ԭ������Ϣ',"
    '      ����_In       ������Ϣ.����֤��%Type,
    strSQL = strSQL & "NULL,"
    '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
    strSQL = strSQL & "NULL,"
    '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic����_In     ������Ϣ.Ic����%Type := Null,
    strSQL = strSQL & "NULL,"
    '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
    strSQL = strSQL & "NULL)"

     
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    BlandCancel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub MoveNextCell(ByVal VsfData As VSFlexGrid, Optional ByVal blnNext As Boolean = True)

    Dim intRow As Integer
    If blnNext = True Then
toMoveNextCol:
        If VsfData.Col < VsfData.Cols - 1 Then
            VsfData.Col = VsfData.Col + 1
            If VsfData.ColWidth(VsfData.Col) = 0 Or VsfData.ColHidden(VsfData.Col) Then
                GoTo toMoveNextCol
            End If
        Else
toMoveNextRow:
            '������һ��
            intRow = 1
            If VsfData.Row + intRow < VsfData.Rows Then
                VsfData.Row = VsfData.Row + intRow
            End If
            If VsfData.RowHidden(VsfData.Row) Then
                If VsfData.Row < VsfData.Rows - 1 Then
                    GoTo toMoveNextRow
                Else
                    For intRow = VsfData.Rows - 1 To VsfData.FixedRows Step -1
                        If VsfData.RowHidden(intRow) = False Then
                            VsfData.Row = intRow
                            Exit For
                        End If
                    Next intRow
                End If
            End If
            VsfData.Col = VsfData.FixedCols
        End If
    Else
toMovePrevCol:
        If VsfData.Col > VsfData.FixedCols Then
            VsfData.Col = VsfData.Col - 1
            If VsfData.ColWidth(VsfData.Col) = 0 Or VsfData.ColHidden(VsfData.Col) Then GoTo toMovePrevCol
        Else
toMovePrevRow:
'            '������һ��
            intRow = 1
            If VsfData.Row - intRow >= VsfData.FixedRows Then
                VsfData.Row = VsfData.Row - intRow
            End If
            If VsfData.RowHidden(VsfData.Row) Then
                If VsfData.Row > VsfData.FixedRows Then
                    GoTo toMovePrevRow
                Else
                    For intRow = VsfData.FixedRows To VsfData.Rows - 1
                        If VsfData.RowHidden(intRow) = False Then
                            VsfData.Row = intRow
                            Exit For
                        End If
                    Next intRow
                End If
            End If
            VsfData.Col = VsfData.FixedCols
        End If
    End If

    If VsfData.ColIsVisible(VsfData.Col) = False Then
        VsfData.LeftCol = VsfData.Col
    End If
    If VsfData.RowIsVisible(VsfData.Row) = False Then
        VsfData.TopRow = VsfData.Row
    End If
End Sub

Private Sub vsInoculate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsInoculate_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strCurDate As String
    If Col = 1 Or Col = 3 Then '���������б༭ʱ���ж��Ƿ�����������200
        With vsInoculate
           If LenB(StrConv(.EditText, vbFromUnicode)) > 200 Then
                If MsgBox("�������������ַ���������ַ���200,�����Ƿ񽫶�����ַ������Զ��س���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    .EditText = StrConv(MidB(StrConv(.EditText, vbFromUnicode), 1, 200), vbUnicode)
                Else
                    Cancel = True
                End If
           End If
        End With
    Else
        With vsInoculate
            If IsDate(Format(.EditText, "YYYY-MM-DD")) = False And .EditText <> "    -  -  " Then
                 MsgBox "��������ڸ�ʽ���Ի�����ȷ�����ڣ����飡", vbInformation, gstrSysName
                 Cancel = True
            ElseIf .EditText = "    -  -  " Then
                 .EditText = ""
            Else
                If .EditText <> "" Then
                    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
                    If Format(.EditText, "YYYY-MM-DD") > strCurDate Then
                        MsgBox "�������ڲ��ܴ��ڷ�����ϵͳʱ��[" & strCurDate & "],���飡", vbInformation, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = Format(.EditText, "YYYY-MM-DD")
                    
                    If Col = 2 And vsInoculate.Rows - 1 = Row Then
                        vsInoculate.Rows = vsInoculate.Rows + 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsLinkMan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsLinkMan
        If Row = .Rows - 1 And Col = .FixedCols And .TextMatrix(Row, Col) <> "" Then
            .Rows = .Rows + 1
        End If
        If Col = .ColIndex("��ϵ�˹�ϵ") Then
            If zlCommFun.GetNeedName(.TextMatrix(Row, Col)) = "����" Then
                .Cell(flexcpBackColor, Row, .ColIndex("��ϵ�˹�ϵ��ע")) = &H80000005
            Else
                .TextMatrix(Row, .ColIndex("��ϵ�˹�ϵ��ע")) = ""
                .Cell(flexcpBackColor, Row, .ColIndex("��ϵ�˹�ϵ��ע")) = &HE0E0E0
            End If
        End If
    End With
End Sub

Private Sub vsLinkMan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsLinkMan
        If .Col = .ColIndex("��ϵ�˹�ϵ��ע") Then
            If zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("��ϵ�˹�ϵ"))) = "����" Then
                Cancel = False
            Else
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub vsLinkMan_GotFocus()
    If mblnCheckPatiCard = False Then
        vsLinkMan.Col = vsLinkMan.FixedCols
        vsLinkMan.Row = vsLinkMan.FixedRows
    End If
    mblnCheckPatiCard = False
End Sub

Private Sub vsLinkMan_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer, i As Integer
    With vsLinkMan
        If KeyCode = vbKeyDelete And .Row >= .FixedRows And mbytInState <> 2 Then
            intRow = .Row
            If .Rows = .FixedRows + 1 Then
                .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
            Else
                Call .RemoveItem(.Row)
                If intRow >= .Rows Then
                    .Row = .Rows - 1
                Else
                    .Row = intRow
                End If
                .Col = .FixedCols
            End If
            If .Rows <= .FixedRows Then .Rows = .FixedRows + 1
            txt��ϵ������.Text = .TextMatrix(.FixedRows, .ColIndex("��ϵ������"))
            For i = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                If zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.List(i)) = .TextMatrix(.FixedRows, .ColIndex("��ϵ�˹�ϵ")) Then
                    Exit For
                End If
            Next
            If i < cbo��ϵ�˹�ϵ.ListCount Then
                cbo��ϵ�˹�ϵ.ListIndex = i
            Else
                cbo��ϵ�˹�ϵ.ListIndex = -1
            End If
            
            txt��ϵ�����֤��.Text = .TextMatrix(.FixedRows, .ColIndex("��ϵ�����֤��"))
            txt��ϵ�˵绰.Text = .TextMatrix(.FixedRows, .ColIndex("��ϵ�˵绰"))
            txtLinkManInfo.Text = .TextMatrix(.FixedRows, .ColIndex("��ϵ�˹�ϵ��ע"))
            
        ElseIf KeyCode = vbKeyReturn And .Row >= .FixedRows Then
            If .TextMatrix(.Row, .FixedCols) = "" And .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                Call MoveNextCell(vsLinkMan)
            End If
        End If
    End With
End Sub

Private Sub vsLinkMan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsLinkMan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsLinkMan
        If KeyAscii = vbKeyReturn Then Exit Sub
        If Col = .ColIndex("��ϵ�����֤��") Then
            If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
                KeyAscii = 0
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        ElseIf Col = .ColIndex("��ϵ�˵绰") Then
            If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        ElseIf Col = .ColIndex("��ϵ�˹�ϵ��ע") Then
            If InStr(":��,��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        End If
    End With
End Sub

Private Sub vsLinkMan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Integer
    
    With vsLinkMan
        If Not Row = .FixedRows Then Exit Sub
        Select Case Col
            Case .ColIndex("��ϵ������")
                txt��ϵ������.Text = Trim(.EditText)
            Case .ColIndex("��ϵ�˹�ϵ")
                For i = 0 To cbo��ϵ�˹�ϵ.ListCount - 1
                    If zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.List(i)) = Trim(.EditText) Then Exit For
                Next
                If i < cbo��ϵ�˹�ϵ.ListCount Then
                    cbo��ϵ�˹�ϵ.ListIndex = i
                Else
                    cbo��ϵ�˹�ϵ.ListIndex = -1
                End If
                
            Case .ColIndex("��ϵ�����֤��")
                txt��ϵ�����֤��.Text = Trim(.EditText)
            Case .ColIndex("��ϵ�˵绰")
                txt��ϵ�˵绰.Text = Trim(.EditText)
            Case .ColIndex("��ϵ�˹�ϵ��ע")
                txtLinkManInfo.Text = Trim(.EditText)
        End Select
    End With
End Sub

Private Sub vsOtherInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 And vsOtherInfo.Rows - 1 = Row And vsOtherInfo.TextMatrix(Row, Col) <> "" Then
        vsOtherInfo.Rows = vsOtherInfo.Rows + 1
    End If
End Sub

Private Sub vsOtherInfo_GotFocus()
    If mblnCheckPatiCard = False Then
        vsOtherInfo.Col = vsOtherInfo.FixedCols
        vsOtherInfo.Row = vsOtherInfo.FixedRows
    End If
    mblnCheckPatiCard = False
End Sub

Private Sub vsOtherInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    
    With vsOtherInfo
        If KeyCode = vbKeyDelete And .Row >= .FixedRows And mbytInState <> 2 Then
            intRow = .Row
            If .Col > .FixedCols + 1 Then
                .TextMatrix(intRow, .FixedCols + 2) = ""
                .TextMatrix(intRow, .FixedCols + 3) = ""
            Else
                If .Rows = .FixedRows + 1 Then
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                Else
                    Call .RemoveItem(.Row)
                    If intRow >= .Rows Then
                        .Row = .Rows - 1
                    Else
                        .Row = intRow
                    End If
                    .Col = .FixedCols
                End If
            End If
        ElseIf KeyCode = vbKeyReturn And .Row >= .FixedRows Then
            If ((.TextMatrix(.Row, .FixedCols) = "" And .Col = .FixedCols) Or (.Col = .FixedCols + 2 And .TextMatrix(.Row, .FixedCols + 2) = "") Or .Col = .FixedCols + 3) And .Row = .Rows - 1 Then
                If mbytInState = 2 Then
                    If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
                Else
                    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
                End If
            Else
               Call MoveNextCell(vsOtherInfo)
            End If
        End If
    End With
End Sub

Private Sub vsOtherInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Function CheckByPatiNO(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bytMode As Byte, ByRef strno As String) As Boolean
'���:
'       lngPatiID:����ID
'       bytMode:0:���סԺ���Ƿ���֮ǰ�Ѿ�ʹ��,1:��ȡ���˱���סԺǰ�����һ�ε�סԺ��
'       strNo:bytMode=0,Ҫ����סԺ��,bytMode=1,���ص�סԺ��
'����:bytMode=0,�Ѿ�ʹ�÷���TRUE,bytMode=1,������ʷסԺ����סԺ�Ų�Ϊ�գ�����TRUE
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    If bytMode = 1 Then
        If lngPageID = 0 Then 'ԤԼ�Ǽ�
            gstrSQL = "Select סԺ�� From ������ҳ Where ����id = [1] And Nvl(��ҳid, 0) <> [2] And סԺ�� Is Not Null Order By ��ҳid Desc"
        Else
            gstrSQL = "Select סԺ�� From ������ҳ Where ����id = [1] And ��ҳid < [2] And סԺ�� Is Not Null Order By ��ҳid Desc"
        End If
    Else
        gstrSQL = "Select ����ID from ������ҳ where ����ID=[1] and nvl(��ҳID,0)<>[2] and סԺ��=[3] and rownum<2"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺ����ȡ", lngPatiID, lngPageID, Val(strno))
    If bytMode = 0 Then
        CheckByPatiNO = rsTemp.RecordCount > 0
    ElseIf bytMode = 1 Then
        If rsTemp.RecordCount > 0 Then strno = rsTemp!סԺ�� & ""
        CheckByPatiNO = strno <> ""
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub InitStructAddress()
'����:�����Ƿ����ýṹ����ַ��������
    Dim i As Long
    Dim lngLeft As Long
    
    If gbln���ýṹ����ַ Then
        For i = PatiAddress.LBound To PatiAddress.UBound
            If i = E_IX_��סַ Or i = E_IX_���ڵ�ַ Or i = E_IX_��ϵ�˵�ַ Then
                PatiAddress(i).Items = Five
            End If
            PatiAddress(i).TextBackColor = &H80000005
            PatiAddress(i).Visible = True
            PatiAddress(i).ShowTown = gbln��ʾ����
        Next
        For i = LBound(marrAddress) To UBound(marrAddress)
            marrAddress(i) = ""
        Next
        txt��ͥ��ַ.Visible = False
        cmd��ͥ��ַ.Visible = False
        txt�����ص�.Visible = False
        cmd�����ص�.Visible = False
        txt���ڵ�ַ.Visible = False
        cmd���ڵ�ַ.Visible = False
        txt����.Visible = False
        cmd����.Visible = False
        txt��ϵ�˵�ַ.Visible = False
        cmd��ϵ�˵�ַ.Visible = False
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
             PatiAddress(i).Visible = False
        Next
        
        txt��ͥ��ַ.Visible = True
        cmd��ͥ��ַ.Visible = True
        txt�����ص�.Visible = True
        cmd�����ص�.Visible = True
        txt���ڵ�ַ.Visible = True
        cmd���ڵ�ַ.Visible = True
        txt����.Visible = True
        cmd����.Visible = True
        txt��ϵ�˵�ַ.Visible = True
        cmd��ϵ�˵�ַ.Visible = True
        
        '�������
        lngLeft = lblѧ��.Left + lblѧ��.width
        lbl��ͥ�绰.Left = lngLeft - lbl��ͥ�绰.width
        lbl���ڵ�ַ�ʱ�.Left = lngLeft - lbl���ڵ�ַ�ʱ�.width
        lngLeft = cboѧ��.Left
        txt��ͥ�绰.Left = lngLeft
        txt���ڵ�ַ�ʱ�.Left = lngLeft
    End If
End Sub

Private Sub SetStrutAddress(Optional ByVal bytFunc As Byte)
'����:89980���˽ṹ��
'bytFunc=1 �������
'       =2 ���û��ڵ�ַ�ͼ�ͥ��ַ��ȱʡֵ
    Dim i As Long
    
    If bytFunc = 2 Then
        txt��ͥ��ַ.Text = marrAddress(0) & marrAddress(1) & marrAddress(2) & marrAddress(3) & marrAddress(4)
        txt���ڵ�ַ.Text = marrAddress(0) & marrAddress(1) & marrAddress(2) & marrAddress(3) & marrAddress(4)
        Call PatiAddress(E_IX_��סַ).LoadStructAdress(marrAddress(0), marrAddress(1), marrAddress(2), marrAddress(3), marrAddress(4))
        Call PatiAddress(E_IX_���ڵ�ַ).LoadStructAdress(marrAddress(0), marrAddress(1), marrAddress(2), marrAddress(3), marrAddress(4))
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
            If bytFunc = 1 Then
                PatiAddress(i).Value = ""
            Else
                PatiAddress(i).Enabled = (mbytInState <> EState.E����)
            End If
        Next
    End If
End Sub

Private Sub ReCalcBirthDay(Optional ByRef strMsg As String)
    Dim strBirth As String
    
    If CreatePublicPatient() Then
        If gobjPublicPatient.ReCalcBirthDay(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, Trim(cbo���䵥λ.Text), ""), strBirth, Format(txt��Ժʱ��.Text, "YYYY-MM-DD HH:MM"), strMsg) Then
            If txt��������.Enabled Then txt��������.Text = Format(strBirth, "YYYY-MM-DD")
            If txt����ʱ��.Enabled Then
                strBirth = Format(strBirth, "HH:MM")
                txt����ʱ��.Text = IIf(strBirth = "00:00", "__:__", strBirth)
            End If
            cbo���䵥λ.Tag = txt����.Text & "_" & cbo���䵥λ.Text
        End If
    End If
End Sub

Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ���ӿ�
    '����: true-�ɹ���false-ʧ��
    '����:���ϴ�
    '����:2016/6/20 13:54:56
    '����:97634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    If txt����.Locked Then Exit Function
    If mobjSquare Is Nothing Then
       If Not zlCreateSquare() Then Exit Function
    End If
    If mobjSquare Is Nothing Then Exit Function
    If mCurSendCard.lng�����ID = 0 Or Val(mCurSendCard.str��������) < 99 Then Exit Function
    If mobjSquare.zlSetBrushCardObject(mCurSendCard.lng�����ID, IIf(blnComm, txt����, Nothing), strExpend) Then
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New clsCommEvents
        Call mobjSquare.zlInitEvents(Me.hWnd, mobjCommEvents)
    End If
End Function

Private Sub EMPI_LoadPati(Optional ByVal lngFunc As Long = 0)
'����:��EMPI�������Ĳ�����Ϣ���µ�����
'lngFunc=0 ���²�����Ϣ;1-���ݷ��صĲ���ID���¼��ز��˻�����Ϣ�����
    Dim rsPatiIn As ADODB.Recordset
    Dim rsPatiOut As ADODB.Recordset
    Dim str�������� As String
    Dim blnRet As Boolean
    Static blnOpen As Boolean
    
    If blnOpen Then Exit Sub
    If CreatePlugInOK(glngModul) Then
        '��֯���˻�����Ϣ
        Set rsPatiIn = New ADODB.Recordset
        With rsPatiIn.Fields
            .Append "����ID", adBigInt
            .Append "��ҳID", adBigInt
            .Append "�Һ�ID", adBigInt
            '-------------------------------
            .Append "�����", adVarChar, 18
            .Append "סԺ��", adVarChar, 18
            .Append "ҽ����", adVarChar, 30
            .Append "���֤��", adVarChar, 18
            .Append "����֤��", adVarChar, 20
            .Append "����", adVarChar, 100
            .Append "�Ա�", adVarChar, 4
            .Append "��������", adVarChar, 20 '���ڸ�ʽ��YYYY-MM-DD HH:MM:SS
            .Append "�����ص�", adVarChar, 100
            .Append "����", adVarChar, 30
            .Append "����", adVarChar, 20
            .Append "ѧ��", adVarChar, 10
            .Append "ְҵ", adVarChar, 80
            .Append "������λ", adVarChar, 100
            .Append "����", adVarChar, 30
            .Append "����״��", adVarChar, 4
            .Append "��ͥ�绰", adVarChar, 20
            .Append "��ϵ�˵绰", adVarChar, 20
            .Append "��λ�绰", adVarChar, 20
            .Append "��ͥ��ַ", adVarChar, 100
            .Append "��ͥ��ַ�ʱ�", adVarChar, 6
            .Append "���ڵ�ַ", adVarChar, 100
            .Append "���ڵ�ַ�ʱ�", adVarChar, 6
            .Append "��λ�ʱ�", adVarChar, 6
            .Append "��ϵ�˵�ַ", adVarChar, 100
            .Append "��ϵ�˹�ϵ", adVarChar, 30
            .Append "��ϵ������", adVarChar, 64
        End With
        rsPatiIn.CursorLocation = adUseClient
        rsPatiIn.LockType = adLockOptimistic
        rsPatiIn.CursorType = adOpenStatic
        rsPatiIn.Open
        
        If txt����ʱ�� = "__:__" Then
            str�������� = IIf(IsDate(txt��������.Text), Format(txt��������.Text, "YYYY-MM-DD"), "")
        Else
            str�������� = IIf(IsDate(txt��������.Text), Format(txt��������.Text & " " & txt����ʱ��.Text, "YYYY-MM-DD HH:MM:SS"), "")
        End If
 
        With rsPatiIn
            .AddNew
            !����ID = CLng(txtPatient.Tag)
            !��ҳID = CLng(txtPages.Text)
            !סԺ�� = Trim(txtסԺ��.Text)
            !ҽ���� = Trim(txtҽ����.Text)
            '-Ҫ���µ��ֶ�--------------------------------------------
            !���֤�� = Trim(txt���֤��.Text)
            !����֤�� = Trim(txt����֤��.Text)
            !���� = Trim(txt����.Text)
            !�Ա� = zlCommFun.GetNeedName(cbo�Ա�.Text)
            !�������� = str�������� '���ڸ�ʽ��YYYY-MM-DD HH:MM:SS
            !�����ص� = Trim(txt�����ص�.Text)
            !���� = zlCommFun.GetNeedName(cbo����.Text)
            !���� = zlCommFun.GetNeedName(cbo����.Text)
            !ѧ�� = zlCommFun.GetNeedName(cboѧ��.Text)
            !ְҵ = zlCommFun.GetNeedName(cboְҵ.Text)
            !������λ = Trim(txt������λ.Text)
            !����״�� = zlCommFun.GetNeedName(cbo����״��.Text)
            !��ͥ�绰 = Trim(txt��ͥ�绰.Text)
            !��ϵ�˵绰 = Trim(txt��ϵ�˵绰.Text)
            !��λ�绰 = Trim(txt��λ�绰.Text)
            !��ͥ��ַ = Trim(txt��ͥ��ַ.Text)
            !��ͥ��ַ�ʱ� = Trim(txt��ͥ��ַ�ʱ�.Text)
            !���ڵ�ַ = Trim(txt���ڵ�ַ.Text)
            !���ڵ�ַ�ʱ� = Trim(txt���ڵ�ַ�ʱ�.Text)
            !��λ�ʱ� = Trim(txt��λ�ʱ�.Text)
            !��ϵ�˵�ַ = Trim(txt��ϵ�˵�ַ.Text)
            !��ϵ�˹�ϵ = zlCommFun.GetNeedName(cbo��ϵ�˹�ϵ.Text)
            !��ϵ������ = Trim(txt��ϵ������.Text)
            .Update
            '-------------------------------------------------------
        End With
        
        '���ò�ѯ�ӿ�
        On Error Resume Next
        blnRet = gobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModul, rsPatiIn, rsPatiOut)
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: On Error GoTo 0
        If Not blnRet Then Exit Sub
        If rsPatiOut Is Nothing Then Exit Sub
        If rsPatiOut.RecordCount = 0 Then Exit Sub
        '�ҵ����ˣ����������µ���Ϣ���µ�����
        With rsPatiOut
            mblnEMPI = True     '���ڱ���ҵ���������
            '104916 ֻ��������,�ӿڵ����������������Ϣ�ҵ�HIS����IDʱ�������½�����
            If mbytInState = E���� And CLng(txtPatient.Tag) <> CLng(!����ID & "") And CLng(!����ID & "") <> 0 And lngFunc = 1 Then
                ClearCard
                txtPatient.Text = "-" & !����ID
                blnOpen = True
                Call txtPatient_KeyPress(vbKeyReturn)
                blnOpen = False
                If txtPatient.Text = "" Then Exit Sub
            End If
            Call cbo.Locate(cbo�Ա�, !�Ա� & "")
            Call cbo.Locate(cbo����, !���� & "")
            Call cbo.Locate(cbo����, !���� & "")
            Call cbo.Locate(cboѧ��, !ѧ�� & "")
            Call cbo.SeekIndex(cboְҵ, !ְҵ & "")  '���������ַ�
            Call cbo.Locate(cbo����״��, !����״�� & "")
            Call cbo.Locate(cbo��ϵ�˹�ϵ, !��ϵ�˹�ϵ & "")
            
            If IsDate(!�������� & "") Then
                txt��������.Text = Format(CDate(!�������� & ""), "YYYY-MM-DD")
                txt����ʱ��.Text = IIf(Format(CDate(!�������� & ""), "HH:MM") = "00:00", "__:__", Format(CDate(!�������� & ""), "HH:MM"))
            End If
            
            If gbln���ýṹ����ַ Then
                PatiAddress(E_IX_�����ص�).Value = !�����ص� & ""
                PatiAddress(E_IX_��סַ).Value = !��ͥ��ַ & ""
                PatiAddress(E_IX_���ڵ�ַ).Value = !���ڵ�ַ & ""
                PatiAddress(E_IX_��ϵ�˵�ַ).Value = !��ϵ�˵�ַ & ""
            End If
            txtҽ����.Text = !ҽ���� & ""
            txt�����ص�.Text = !�����ص� & ""
            txt��ͥ��ַ.Text = !��ͥ��ַ & ""
            txt���ڵ�ַ.Text = !���ڵ�ַ & ""
            txt��ϵ�˵�ַ.Text = !��ϵ�˵�ַ & ""
            txt���֤��.Text = !���֤�� & ""
            txt����֤��.Text = !����֤�� & ""
            txt����.Text = !���� & ""
            txt������λ.Text = !������λ & ""
            txt��ͥ�绰.Text = !��ͥ�绰 & ""
            txt��ϵ�˵绰.Text = !��ϵ�˵绰 & ""
            txt��λ�绰.Text = !��λ�绰 & ""
            txt��ͥ��ַ�ʱ�.Text = !��ͥ��ַ�ʱ� & ""
            txt���ڵ�ַ�ʱ�.Text = !���ڵ�ַ�ʱ� & ""
            txt��λ�ʱ�.Text = !��λ�ʱ� & ""
            txt��ϵ������.Text = !��ϵ������ & ""
        End With
    End If
End Sub

Private Function EMPI_AddORUpdatePati(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef strErr As String) As Boolean
'����:���ӻ����EMPI������Ϣ
    Dim lngRet  As Long
    Dim strPlugErr As String
    Dim strTmp As String
    
    lngRet = 1 'Ĭ�ϳɹ� ���� �ϰ�zlPlug����֧�ִ˽ӿڴ����:438
    If CreatePlugInOK(glngModul) Then
        If Not mblnEMPI Then
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=�ɹ�;0-ʧ��
            Call zlPlugInErrH(Err, "EMPI_AddPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "��EMPIƽ̨����������Ϣʧ�ܣ�"
        Else
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=�ɹ�;0-ʧ��
            Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "��EMPIƽ̨���²�����Ϣʧ�ܣ�"
        End If
        If strPlugErr <> "" Then
            strErr = strTmp & vbCrLf & strPlugErr
             Exit Function
        ElseIf lngRet = 0 Then
            strErr = strTmp & vbCrLf & strErr
            Exit Function
        End If
    End If
    
    EMPI_AddORUpdatePati = True
End Function

Private Sub vsCertificate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long
    If Row < 1 Or Col < 0 Then Exit Sub
    '�����:90875

    With vsCertificate
        If Col = 1 Or Col = 3 Then '֤�����벻�ܳ���30
            If Len(.TextMatrix(Row, Col)) > 30 Then
                 MsgBox "֤�������ַ���������ַ���30,������ַ������Զ��س���", vbInformation, gstrSysName
                 .TextMatrix(Row, Col) = Mid(.TextMatrix(Row, Col), 1, 30)
            End If
            If Col = 3 And .Rows - 1 = Row And .TextMatrix(Row, Col) <> "" Then
                .Rows = .Rows + 1
            End If
        ElseIf Col = 0 Or Col = 2 Then '����Ƿ�ѡ�����ظ���֤������
            For lngRow = 1 To .Rows - 1
                For lngCol = 0 To .Cols - 1 Step 2
                    If (lngRow <> Row Or lngCol <> Col) And .TextMatrix(lngRow, lngCol) = .TextMatrix(Row, Col) And .TextMatrix(Row, Col) <> "" Then
                        MsgBox .TextMatrix(lngRow, lngCol) & "�Ѵ��ڣ������ظ�ѡ��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        .Select Row, Col
                        Exit Sub
                    End If
                Next
            Next
        End If
    End With
End Sub
Private Sub vsCertificate_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����:90875
    If KeyCode = 27 And vsCertificate.Rows = 2 Then
        If vsCertificate.TextMatrix(1, 3) <> "" Then
            vsCertificate.TextMatrix(1, 2) = "": vsCertificate.TextMatrix(1, 3) = ""
        Else
            vsCertificate.TextMatrix(1, 0) = "": vsCertificate.TextMatrix(1, 1) = ""
        End If
    End If
    If KeyCode = 27 And vsCertificate.Rows > 2 Then 'Esc
        If vsCertificate.TextMatrix(vsCertificate.Rows - 1, 2) <> "" Or vsCertificate.TextMatrix(vsCertificate.Rows - 1, 3) <> "" Then
            vsCertificate.TextMatrix(vsCertificate.Rows - 1, 2) = "": vsCertificate.TextMatrix(vsCertificate.Rows - 1, 3) = ""
        Else
            vsCertificate.Rows = vsCertificate.Rows - 1
        End If
    End If
End Sub

Private Sub vsCertificate_KeyPress(KeyAscii As Integer)
    '78408:���ϴ�,2014/10/9,�����ת
    If KeyAscii = 13 Then
        If vsCertificate.Col = 3 And vsCertificate.Rows - 1 = vsCertificate.Row Then
            zlCommFun.PressKey vbKeyTab
        ElseIf vsCertificate.Col = 3 Then
            vsCertificate.Col = 0: vsCertificate.Row = vsCertificate.Row + 1
            zlCommFun.PressKey vbKeyReturn
        Else
            zlCommFun.PressKey vbKeyRight
        End If
    End If
End Sub

Private Sub InitCertificate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:90875
    '����:2015/12/17 16:59:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    Dim strSQL As String, rsTemp As ADODB.Recordset, str��ϵ As String, i As Integer
    With vsCertificate
    '��ʼ���б�����
        .Editable = IIf(mbytInState = 2, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
    '������ͷ
        SetColumHeader vsCertificate, C_CertificateHeader
    '��������Ϣ
        strSQL = "Select ����,ȱʡ��־ from ֤������  Where  ���� Not Like '����%' and ���� Not Like '%���֤'" & vbNewLine & _
                " And Not ���� in (Select ���� from  ҽ�ƿ���� Where Nvl(�Ƿ�֤��,0)=0 or Nvl(�Ƿ�����,0)=0)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTemp.RecordCount = 0 Then .Editable = flexEDNone: Exit Sub
        Do While Not rsTemp.EOF
            str��ϵ = str��ϵ & "|" & Nvl(rsTemp!����)
            rsTemp.MoveNext
        Loop
        str��ϵ = Mid(str��ϵ, 2)
        If str��ϵ <> "" Then .ColComboList(0) = str��ϵ: .ColComboList(2) = str��ϵ
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub LoadCertificate(ByVal lng����ID As Long)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:���ز��˵�֤����Ϣ������
    '����:���ϴ�
    'ʱ��:2015/12/17 17:37:27
    '����:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    
    On Error GoTo ErrHand
    strSQL = "Select  A.����,A.ID,B.���� from ҽ�ƿ���� A, ����ҽ�ƿ���Ϣ B " & _
            "Where A.ID= B.�����ID And A.�Ƿ�����=1 And A.�Ƿ�֤��=1 And B.״̬=0  And  B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    With vsCertificate
        .Clear 1
        .Rows = 2
        lngRow = 1: lngCol = 0
        While Not rsTemp.EOF
            .TextMatrix(lngRow, lngCol) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, lngCol + 1) = Nvl(rsTemp!����)
            lngCol = lngCol + 2
            If lngCol > 2 Then .Rows = .Rows + 1: lngRow = lngRow + 1: lngCol = 0
            rsTemp.MoveNext
        Wend
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub AddCardTypeSQL(ByVal intOper As Integer, ByVal lng�����ID As Long, ByVal strCode As String, ByVal strȫ�� As String, ByVal str���� As String, _
                           ByVal lng���ų��� As Long, ByRef strSQL As String)

    ' Zl_ҽ�ƿ����_Update
    strSQL = "Zl_ҽ�ƿ����_Update("
    '  Id_In           In ҽ�ƿ����.ID%Type,
    strSQL = strSQL & "" & lng�����ID & ","
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & strȫ�� & "',"
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ǰ׺�ı�_In     In ҽ�ƿ����.ǰ׺�ı�%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  ���ų���_In     In ҽ�ƿ����.���ų���%Type,
    strSQL = strSQL & "" & lng���ų��� & ","
    '  ȱʡ��־_In     In ҽ�ƿ����.ȱʡ��־%Type,
    strSQL = strSQL & "" & 0 & ","
    '  �Ƿ�̶�_In     In ҽ�ƿ����.�Ƿ�̶�%Type,
    strSQL = strSQL & "1,"
    '  �Ƿ��ϸ����_In In ҽ�ƿ����.�Ƿ��ϸ����%Type,
    strSQL = strSQL & "" & 0 & ","
    '  �Ƿ�����_In     In ҽ�ƿ����.�Ƿ�����%Type,
    strSQL = strSQL & "" & 0 & ","
    '  �Ƿ�����ʻ�_In In ҽ�ƿ����.�Ƿ�����ʻ�%Type,
    strSQL = strSQL & "" & 0 & ","
    '  �Ƿ�ȫ��_In     In ҽ�ƿ����.�Ƿ�ȫ��%Type,
    strSQL = strSQL & "0,"
    '  ����_In         In ҽ�ƿ����.����%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  ��ע_In         In ҽ�ƿ����.��ע%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  �ض���Ŀ_In     In ҽ�ƿ����.�ض���Ŀ%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '    �շ�ϸĿid_In   In �շ���ĿĿ¼.ID%Type,
    strSQL = strSQL & "" & "0" & ","
    '  ���㷽ʽ_In     In ҽ�ƿ����.���㷽ʽ%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  �Ƿ�����_In     In ҽ�ƿ����.�Ƿ�����%Type,
    strSQL = strSQL & "1,"
    '  ��������_In     In ҽ�ƿ����.��������%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  �Ƿ��ظ�ʹ��_In In ҽ�ƿ����.�Ƿ��ظ�ʹ��%Type,
    strSQL = strSQL & "" & 1 & ","
    '���볤��_In     In ҽ�ƿ����.���볤��%Type,
    strSQL = strSQL & "" & 10 & ","
    '���볤������_In In ҽ�ƿ����.���볤������%Type,
    strSQL = strSQL & "" & 0 & ","
    '�������_In     In ҽ�ƿ����.�������%Type,
    strSQL = strSQL & "" & 0 & ","
    strSQL = strSQL & "" & 1 & ","
    '  ������ʽ_In     In Integer := 0
    strSQL = strSQL & "" & intOper & ","
    '�Ƿ�ģ������_In     In ҽ�ƿ����.�Ƿ�ģ������%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '�����:51072
    '������������_In     In ҽ�ƿ����.������������%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '�Ƿ�ȱʡ����_In     In ҽ�ƿ����.�Ƿ�ȱʡ����%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '�����:56508
    '�Ƿ��ƿ�_In
    strSQL = strSQL & "" & 0 & ","
    '�Ƿ񷢿�_In
    strSQL = strSQL & "" & 0 & ","
    '�Ƿ�д��_In
    strSQL = strSQL & "" & 0 & ","
    '�����:57697
    '����_In
    strSQL = strSQL & "" & 0 & ","
    '�����:57326
    strSQL = strSQL & "" & 1 & ","
    '77872,���ϴ�,2014/12/3:�Ƿ�֧��ת�ʼ�����
    '�Ƿ�ת�ʼ�����_In  In ҽ�ƿ����.�Ƿ�ת�ʼ�����%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '��������_In       In ҽ�ƿ����.��������%Type := '1000',
    strSQL = strSQL & "" & "1000" & ","
    '���̿��Ʒ�ʽ_In   In ҽ�ƿ����.���̿��Ʒ�ʽ%Type := 0,
    strSQL = strSQL & "" & 0 & ","
    '90875:���ϴ�,2015/12/16,����ҽ�ƿ�֤������
    '�Ƿ�֤��_In  In ҽ�ƿ����.�Ƿ�֤��%Type:=0
    strSQL = strSQL & "" & 1 & ")"
End Sub

Private Sub AddCertificate(ByVal lng����ID As Long, ByRef arrSQL As Variant, ByVal dtCurdate As Date)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:����֤��������Ϣ������ǵ�һ�ν��������
    '����:���ϴ�
    'ʱ��:2015/12/17 17:37:27
    '����:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, rsPatiCard As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    Dim lngID As Long, strCode As String
    Dim colPro As Collection
    
    On Error GoTo ErrHand
    Set colPro = New Collection
    '�󶨿�ǰҪ�жϿ�����Ƿ����
    strSQL = "Select B.ID,B.����,B.���ų���,B.����,A.����,A.����ID,Decode(A.���� ,NULL,1,0) as ��ʶ from ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B " & _
            "Where A.�����ID(+)=B.ID And B.�Ƿ�֤��=1 And A.״̬(+)=0 And A.����ID(+)=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    Set rsPatiCard = zlDatabase.CopyNewRec(rsTemp)
    With vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    lngID = 0: strCode = ""
                    rsTemp.Filter = "����='" & .TextMatrix(lngRow, lngCol) & "'"
                    If rsTemp.RecordCount = 0 Then
                        lngID = zlDatabase.GetNextId("ҽ�ƿ����")
                        If mstrFirstCode = "" Then
                            strCode = zlDatabase.GetMax("ҽ�ƿ����", "����", 4)
                            mstrFirstCode = strCode
                        Else
                            strCode = CStr(Val(mstrFirstCode) + 1)
                            strCode = Format(strCode, String(4, "0"))
                            mstrFirstCode = strCode
                        End If
                        Call AddCardTypeSQL(0, lngID, strCode, .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), strSQL)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    ElseIf Len(.TextMatrix(lngRow, lngCol + 1)) > Val(Nvl(rsTemp!���ų���)) Then
                        Call AddCardTypeSQL(1, Val(Nvl(rsTemp!ID)), Nvl(rsTemp!����), .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), strSQL)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    End If
                    
                    '����֤������
                    rsPatiCard.Filter = "����='" & .TextMatrix(lngRow, lngCol) & "' And ����='" & .TextMatrix(lngRow, lngCol + 1) & "'"
                    If rsPatiCard.RecordCount = 0 Then
                        'Zl_ҽ�ƿ��䶯_Insert
                         strSQL = "Zl_ҽ�ƿ��䶯_Insert("
                        '      �䶯����_In   Number,
                        '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
                        strSQL = strSQL & "" & 11 & ","
                        '      ����id_In     סԺ���ü�¼.����id%Type,
                        strSQL = strSQL & "" & lng����ID & ","
                        '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
                        strSQL = strSQL & "" & IIf(lngID = 0, rsTemp!ID, lngID) & ","
                        '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
                        strSQL = strSQL & "'" & Trim(.TextMatrix(lngRow, lngCol + 1)) & "',"
                        '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
                        '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
                        strSQL = strSQL & "'" & "֤������" & "',"
                        '      ����_In       ������Ϣ.����֤��%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSQL = strSQL & "'" & UserInfo.���� & "',"
                        '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
                        strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                        '      Ic����_In     ������Ϣ.Ic����%Type := Null,
                        strSQL = strSQL & "'" & "" & "',"
                        '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
                        strSQL = strSQL & "NULL)"
                    
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    Else
                        rsPatiCard!��ʶ = 1
                        rsPatiCard.Update
                    End If
                End If
            Next
        Next
    End With
    mstrFirstCode = ""
    
    '�����б���û��֤���ţ�Ҫ�����
    rsPatiCard.Filter = "��ʶ=0"
    If rsPatiCard.RecordCount > 0 Then
        rsPatiCard.MoveFirst
        Do While Not rsPatiCard.EOF
            'Zl_ҽ�ƿ��䶯_Insert
             strSQL = "Zl_ҽ�ƿ��䶯_Insert("
            '      �䶯����_In   Number,
            '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
            strSQL = strSQL & "" & 14 & ","
            '      ����id_In     סԺ���ü�¼.����id%Type,
            strSQL = strSQL & "" & lng����ID & ","
            '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
            strSQL = strSQL & "" & rsPatiCard!ID & ","
            '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
            strSQL = strSQL & "'" & rsPatiCard!���� & "',"
            '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
            '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
            strSQL = strSQL & "'" & "֤����ȡ����" & "',"
            '      ����_In       ������Ϣ.����֤��%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
            strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '      Ic����_In     ������Ϣ.Ic����%Type := Null,
            strSQL = strSQL & "'" & "" & "',"
            '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
            strSQL = strSQL & "NULL)"
        
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
            rsPatiCard.MoveNext
        Loop
    End If
    rsPatiCard.Close
    Exit Sub
ErrHand:
    rsPatiCard.Close
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function IsCertificateCard(ByVal lng����ID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:֤��������
    '����:���ϴ�
    'ʱ��:2015/12/17 17:37:27
    '����:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long, str֤������ As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strCardName As String
    
    On Error GoTo ErrHand
    With vsCertificate
        '��������Ƿ�����
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    MsgBox "��ѡ�񿨺�" & .TextMatrix(lngRow, lngCol + 1) & "��֤������", vbInformation, gstrSysName
                    .Select lngRow, lngCol
                    Exit Function
                End If
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    strSQL = "Select 1 from ����ҽ�ƿ���Ϣ A,ҽ�ƿ���� B " & _
                            "Where A.�����ID=B.ID And B.����=[1] And B.�Ƿ�֤��=1 And A.����=[2] And  A.����ID<>[3]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .TextMatrix(lngRow, lngCol), Trim(.TextMatrix(lngRow, lngCol + 1)), lng����ID)
                    If rsTmp.RecordCount > 0 Then
                        MsgBox .TextMatrix(lngRow, lngCol) & ":" & .TextMatrix(lngRow, lngCol + 1) & "���ڱ�ʹ��,����!", vbInformation, gstrSysName
                        .Select lngRow, lngCol
                        Exit Function
                    End If
                    str֤������ = str֤������ & ",'" & .TextMatrix(lngRow, lngCol) & "'"
                End If
            Next
        Next
        
        '���֤�������Ƿ����֤����ҽ�ƿ�����ظ����ظ��򲻱�����Ϣ
        str֤������ = Mid(str֤������, 2)
        If str֤������ = "" Then IsCertificateCard = True: Exit Function
        strSQL = "Select ���� From ҽ�ƿ���� where ���� in (" & str֤������ & ") And Nvl(�Ƿ�֤��,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                strCardName = strCardName & "," & Nvl(rsTmp!����)
            Loop
            
            strCardName = Mid(strCardName, 2)
            MsgBox "ҽ�ƿ����" & strCardName & "�������ظ�,���ܼ�����ӡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    IsCertificateCard = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function FuncPlugPovertyInfo(ByVal lngPatiID As Long, Optional ByVal strXMLPati As String) As Boolean
'����:'102232 zlPlugInչʾ���˷�ƶ��Ϣ
'�����HIS��������:����ֵΪT�������,����ֵΪF-��ֹ����
'�����HISδ�������²���,����ʱ����T-������,F-��ֹ����,����ս���
'δ���ò������,���������������ýӿ� ȱʡ����T-������ز��˼������²��ˡ�
    Dim blnRet As Boolean
    
    blnRet = True
    If CreatePlugInOK(glngModul) And mbytInState <> EState.E���� Then
        On Error Resume Next
        blnRet = gobjPlugIn.PatiValiedCheck(glngSys, glngModul, 2, lngPatiID, 0, strXMLPati) 'T=�ɹ�;F-ʧ��
        Call zlPlugInErrH(Err, "PatiValiedCheck")
        If blnRet = False And Err.Number <> 438 Then
            blnRet = False
            Call ClearCard
        Else
            blnRet = True
        End If
        Err.Clear: On Error GoTo 0
    End If
    FuncPlugPovertyInfo = blnRet
End Function

Private Function CheckMobile(ByVal strMobile As String, ByVal lngPatiID As Long) As Boolean
'����:��鵱ǰ�ֻ����Ƿ����
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "SELECT 1 FROM ������Ϣ Where �ֻ��� = [1] And ����ID <> [2] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ֻ���", strMobile, lngPatiID)
    If Not rsTemp Is Nothing Then
        CheckMobile = rsTemp.EOF = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub SelectYouBian(objTextBox As TextBox)
    '���ܣ��ʱ�ѡ����
    Dim strInput As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI

    strInput = objTextBox.Text
    If strInput <> "" Then
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = strSQL & " And A.���� Like [1] "
        Else
            strSQL = strSQL & " And A.���� Like [1] "
        End If
    Else
        Exit Sub
    End If
    strSQL = "Select Rownum as ID,����,����,�ʱ�  From ���� A " & _
             "Where �ʱ� is not null " & strSQL & " Order by ����"
    vPoint = GetCoordPos(objTextBox.hWnd, 0, 0)
    Set rsTmp = zlDatabase.ShowSQLSelect(objTextBox.Parent, strSQL, 0, "�ʱ�", False, "", "", False, _
        False, True, vPoint.X, vPoint.Y, objTextBox.Height, False, False, False, UCase(strInput) & "%")
    If Not rsTmp Is Nothing Then
        objTextBox.Text = rsTmp!�ʱ� & ""
    End If
End Sub


Private Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Sub CreatePublicExpenseObject(ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ò���
    '���:
    '����:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If mobjPublicExpense Is Nothing Then
        Set mobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If mobjPublicExpense Is Nothing Then Exit Sub
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If mobjPublicExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    mintPriceGradeStartType = mobjPublicExpense.zlGetPriceGradeStartType()
    If mintPriceGradeStartType = 0 Then Exit Sub
    '��ȡվ��۸�ȼ�
    Call mobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", , , mstrPriceGrade)
End Sub

Private Sub ReLoadCardFee(Optional ByVal blnFeedName As Boolean)
    '118124:���ϴ���2018/1/18����ȡ����
    Dim lng����ID As Long, lng�շ�ϸĿID As Long
    Dim strSQL As String, str���� As String
    Dim rsTmp As ADODB.Recordset
    
    If mCurSendCard.rs���� Is Nothing Then Exit Sub
    If mCurSendCard.rs����.RecordCount = 0 Then Exit Sub
    If mCurSendCard.lng�����ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(txt����.Text) = "" Then Exit Sub
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
        lng����ID = mrsInfo!����ID
    End If
    If blnFeedName = False And lng����ID <> 0 Then Exit Sub
    
    str���� = Trim(txt����.Text)
    If IsNumeric(str����) Then str���� = str���� & cbo���䵥λ.Text
    mCurSendCard.rs����.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as �շ�ϸĿID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����", mlngModul, mCurSendCard.lng�����ID, Trim(txt����.Text), lng����ID, _
                Trim(txtPatient.Text), zlStr.NeedName(cbo�Ա�.Text), str����, txt���֤��.Text, Val(Nvl(mCurSendCard.rs����!�շ�ϸĿID)))
    If rsTmp.EOF Then Exit Sub
    
    lng�շ�ϸĿID = Val(Nvl(rsTmp!�շ�ϸĿID))
    Set rsTmp = zlGetSpecialItemFee(mCurSendCard.str�ض���Ŀ, mstrPriceGrade, lng�շ�ϸĿID)
    If Not rsTmp Is Nothing Then Set mCurSendCard.rs���� = rsTmp
    
    With mCurSendCard.rs����
        txt����.Text = Format(IIf(Val(Nvl(!�Ƿ���)) = 1, Val(Nvl(!ȱʡ�۸�)), Val(Nvl(!�ּ�))), "0.00")
        txt����.Tag = txt����.Text  '���ֲ���
        txt����.Locked = Not (Val(Nvl(!�Ƿ���)) = 1)
        txt����.TabStop = (Val(Nvl(!�Ƿ���)) = 1)
        
        If mCurSendCard.rs����!�Ƿ��� = 0 And Val(txt����.Text) <> 0 Then
            txt����.Text = Format(GetActualMoney(zlStr.NeedName(cbo�ѱ�.Text), mCurSendCard.rs����!������ĿID, mCurSendCard.rs����!�ּ�, mCurSendCard.rs����!�շ�ϸĿID), "0.00")
        End If
    End With
End Sub
