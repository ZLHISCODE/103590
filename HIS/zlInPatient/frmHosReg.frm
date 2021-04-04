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
   Caption         =   "病人登记"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "宋体"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicHealth 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
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
            Name            =   "宋体"
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
               Name            =   "宋体"
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
         Caption         =   "…"
         BeginProperty Font 
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
            Name            =   "宋体"
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
         Caption         =   "证件信息"
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
         Caption         =   "血型"
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
         Caption         =   "医学警示"
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
         Caption         =   "其他医学警示"
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
         Caption         =   "联系人信息"
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
         Caption         =   "其他信息"
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
         Caption         =   "接种情况"
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
         Caption         =   "过敏反应"
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
         Name            =   "宋体"
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
      Begin VB.PictureBox pic磁卡 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
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
         Begin VB.Frame fra磁卡 
            Caption         =   "【发卡信息】"
            BeginProperty Font 
               Name            =   "宋体"
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
            Begin VB.ComboBox cbo发卡结算 
               Height          =   360
               Left            =   12645
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Top             =   400
               Width           =   1845
            End
            Begin VB.TextBox txt卡额 
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
            Begin VB.TextBox txt卡号 
               BackColor       =   &H00EBFFFF&
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   1590
               PasswordChar    =   "*"
               TabIndex        =   88
               Top             =   400
               Width           =   1750
            End
            Begin VB.CheckBox chk记帐 
               Caption         =   "记帐"
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
                     Caption         =   "发卡收费(&1)"
                     Key             =   "CardFee"
                     ImageVarType    =   2
                  EndProperty
                  BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "绑定卡号(&2)"
                     Key             =   "CardBind"
                     ImageVarType    =   2
                  EndProperty
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label lbl卡名称 
               Height          =   255
               Left            =   12420
               TabIndex        =   214
               Top             =   0
               Width           =   1575
            End
            Begin VB.Label lbl验证 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "验证"
               Height          =   240
               Left            =   5835
               TabIndex        =   185
               Top             =   460
               Width           =   480
            End
            Begin VB.Label lbl金额 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "金额"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   8880
               TabIndex        =   184
               Top             =   460
               Width           =   480
            End
            Begin VB.Label lbl密码 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "密码"
               Height          =   240
               Left            =   3435
               TabIndex        =   183
               Top             =   465
               Width           =   480
            End
            Begin VB.Label lbl卡号 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "卡号"
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   930
               TabIndex        =   182
               Top             =   450
               Width           =   510
            End
         End
      End
      Begin VB.PictureBox pic预交 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
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
         Begin VB.Frame fra预交 
            Caption         =   "【住院预交信息】"
            ForeColor       =   &H00C00000&
            Height          =   1160
            Left            =   30
            TabIndex        =   169
            Top             =   0
            Width           =   15000
            Begin VB.CheckBox chk单位缴款 
               Caption         =   "单位缴款"
               Height          =   360
               Left            =   12480
               TabIndex        =   84
               Top             =   375
               Width           =   1320
            End
            Begin VB.TextBox txt帐号 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   9480
               MaxLength       =   50
               TabIndex        =   87
               Top             =   735
               Width           =   5025
            End
            Begin VB.TextBox txt预交额 
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
            Begin VB.ComboBox cbo预交结算 
               Height          =   360
               Left            =   6330
               Style           =   2  'Dropdown List
               TabIndex        =   82
               Top             =   375
               Width           =   1770
            End
            Begin VB.TextBox txt结算号码 
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
            Begin VB.TextBox txt缴款单位 
               Height          =   360
               Left            =   1590
               MaxLength       =   50
               TabIndex        =   85
               Top             =   735
               Width           =   2745
            End
            Begin VB.TextBox txt开户行 
               Height          =   360
               Left            =   5280
               MaxLength       =   50
               TabIndex        =   86
               Top             =   735
               Width           =   2805
            End
            Begin VB.Label lblYBMoney 
               AutoSize        =   -1  'True
               Caption         =   "个人帐户余额:"
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
               Caption         =   "摘要"
               BeginProperty Font 
                  Name            =   "宋体"
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
               Caption         =   "金额"
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
               Caption         =   "结算号码"
               Height          =   240
               Left            =   8400
               TabIndex        =   175
               Top             =   435
               Width           =   960
            End
            Begin VB.Label lblStyle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "缴款方式"
               Height          =   240
               Left            =   5205
               TabIndex        =   174
               Top             =   435
               Width           =   960
            End
            Begin VB.Label lblFact 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "实际票号"
               Height          =   240
               Left            =   510
               TabIndex        =   173
               Top             =   435
               Width           =   960
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "缴款单位"
               Height          =   240
               Left            =   510
               TabIndex        =   172
               Top             =   795
               Width           =   960
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "开户行"
               Height          =   240
               Left            =   4440
               TabIndex        =   171
               Top             =   795
               Width           =   720
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "帐号"
               Height          =   240
               Left            =   8880
               TabIndex        =   170
               Top             =   795
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox pic入院 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
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
         Begin VB.Frame fra入院 
            Caption         =   "【住院信息】"
            ForeColor       =   &H00C00000&
            Height          =   2325
            Left            =   0
            TabIndex        =   151
            Top             =   0
            Width           =   15000
            Begin VB.CommandButton cmd转入 
               Caption         =   "…"
               Height          =   300
               Left            =   14160
               TabIndex        =   222
               TabStop         =   0   'False
               Top             =   1580
               Width           =   300
            End
            Begin VB.TextBox txt转入 
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
            Begin VB.ComboBox cbo入院属性 
               Height          =   360
               Left            =   9480
               Style           =   2  'Dropdown List
               TabIndex        =   74
               Top             =   1125
               Width           =   2445
            End
            Begin VB.CheckBox chk再入院 
               Caption         =   "再入院"
               Height          =   360
               Left            =   12120
               TabIndex        =   70
               ToolTipText     =   "再次入住相同诊疗科目编码的临床科室"
               Top             =   735
               Width           =   1095
            End
            Begin VB.ComboBox cbo入院病区 
               Height          =   360
               Left            =   5520
               TabIndex        =   64
               Top             =   330
               Width           =   2565
            End
            Begin VB.ComboBox cbo入院病况 
               Height          =   360
               Left            =   5520
               Style           =   2  'Dropdown List
               TabIndex        =   73
               Top             =   1125
               Width           =   2565
            End
            Begin VB.ComboBox cbo入院方式 
               Height          =   360
               Left            =   9480
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   1550
               Width           =   2445
            End
            Begin VB.ComboBox cbo床位 
               Height          =   360
               ItemData        =   "frmHosReg.frx":0442
               Left            =   9480
               List            =   "frmHosReg.frx":0444
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   345
               Width           =   2445
            End
            Begin VB.CheckBox chk二级院转入 
               Caption         =   "二级院转入"
               Height          =   360
               Left            =   12120
               TabIndex        =   75
               Top             =   1125
               Width           =   1680
            End
            Begin VB.ComboBox cbo门诊医师 
               Height          =   360
               IMEMode         =   2  'OFF
               Left            =   5520
               TabIndex        =   68
               Top             =   720
               Width           =   2565
            End
            Begin VB.CheckBox chk陪伴 
               Caption         =   "是否陪伴"
               Height          =   360
               Left            =   13200
               TabIndex        =   71
               Top             =   735
               Width           =   1380
            End
            Begin VB.TextBox txt备注 
               Height          =   360
               Left            =   9480
               MaxLength       =   100
               TabIndex        =   79
               Top             =   1905
               Width           =   5025
            End
            Begin VB.TextBox txt中医诊断 
               Height          =   360
               Left            =   1590
               MaxLength       =   200
               TabIndex        =   78
               Top             =   1905
               Width           =   6495
            End
            Begin VB.TextBox txt门诊诊断 
               Height          =   360
               Left            =   1590
               MaxLength       =   200
               TabIndex        =   76
               Top             =   1550
               Width           =   6495
            End
            Begin VB.ComboBox cbo护理等级 
               Height          =   360
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   735
               Width           =   2565
            End
            Begin VB.ComboBox cbo住院目的 
               Height          =   360
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Top             =   1125
               Width           =   2565
            End
            Begin VB.ComboBox cbo入院科室 
               Height          =   360
               Left            =   1590
               TabIndex        =   63
               Text            =   "cbo入院科室"
               Top             =   345
               Width           =   2565
            End
            Begin MSMask.MaskEdBox txt入院时间 
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
                  Name            =   "宋体"
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
               Caption         =   "转入"
               Enabled         =   0   'False
               Height          =   240
               Left            =   12120
               TabIndex        =   215
               Top             =   1610
               Width           =   600
            End
            Begin VB.Label lblBedInfo 
               AutoSize        =   -1  'True
               Caption         =   "科室在院及床位信息"
               ForeColor       =   &H00C00000&
               Height          =   240
               Left            =   1560
               TabIndex        =   167
               Top             =   0
               Width           =   2160
            End
            Begin VB.Label lblTimes 
               Caption         =   "第      次住院"
               Height          =   255
               Left            =   12120
               TabIndex        =   166
               Top             =   405
               Width           =   1785
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院属性"
               Height          =   240
               Left            =   8400
               TabIndex        =   165
               Top             =   1185
               Width           =   960
            End
            Begin VB.Label lbl入院病区 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院病区"
               Height          =   240
               Left            =   4380
               TabIndex        =   164
               Top             =   405
               Width           =   960
            End
            Begin VB.Label lbl中医诊断 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "中医诊断"
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
               Caption         =   "备注"
               Height          =   240
               Left            =   8880
               TabIndex        =   162
               Top             =   1965
               Width           =   480
            End
            Begin VB.Label lbl门诊诊断 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊诊断"
               Height          =   240
               Left            =   510
               TabIndex        =   160
               Top             =   1610
               Width           =   960
            End
            Begin VB.Label lbl入院科室 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院科室"
               Height          =   240
               Left            =   510
               TabIndex        =   159
               Top             =   405
               Width           =   960
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院病况"
               Height          =   240
               Left            =   4380
               TabIndex        =   158
               Top             =   1185
               Width           =   960
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院方式"
               Height          =   240
               Left            =   8400
               TabIndex        =   157
               Top             =   1610
               Width           =   960
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院目的"
               Height          =   240
               Left            =   510
               TabIndex        =   156
               Top             =   1185
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "护理等级"
               Height          =   240
               Left            =   510
               TabIndex        =   155
               Top             =   795
               Width           =   960
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院时间"
               Height          =   240
               Left            =   8400
               TabIndex        =   154
               Top             =   795
               Width           =   960
            End
            Begin VB.Label lbl床位 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院病床"
               Height          =   240
               Left            =   8400
               TabIndex        =   153
               Top             =   405
               Width           =   960
            End
            Begin VB.Label lbl门诊医师 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "门诊医师"
               Height          =   240
               Left            =   4380
               TabIndex        =   152
               Top             =   795
               Width           =   960
            End
         End
      End
      Begin VB.PictureBox pic病人 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
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
         Begin VB.Frame fra病人 
            Caption         =   "【基本信息】"
            BeginProperty Font 
               Name            =   "宋体"
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
            Begin VB.PictureBox pic担保 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "宋体"
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
               Begin VB.TextBox txt担保额 
                  Alignment       =   1  'Right Justify
                  ForeColor       =   &H00C00000&
                  Height          =   360
                  Left            =   4605
                  MaxLength       =   10
                  TabIndex        =   59
                  Top             =   30
                  Width           =   1305
               End
               Begin VB.TextBox txt担保人 
                  Height          =   360
                  Left            =   1530
                  MaxLength       =   100
                  TabIndex        =   57
                  Top             =   30
                  Width           =   1605
               End
               Begin VB.CheckBox chkUnlimit 
                  Caption         =   "不限"
                  Height          =   255
                  Left            =   3210
                  TabIndex        =   58
                  ToolTipText     =   "不限担保额时必须设置担保时限"
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
               Begin VB.CheckBox chk临时担保 
                  Caption         =   "临时担保"
                  Height          =   360
                  Left            =   9780
                  TabIndex        =   61
                  Top             =   30
                  Width           =   1280
               End
               Begin MSComCtl2.DTPicker dtp担保时限 
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
                     Name            =   "宋体"
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
               Begin VB.Label lbl担保额 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "金额"
                  Height          =   240
                  Left            =   4050
                  TabIndex        =   118
                  Top             =   90
                  Width           =   480
               End
               Begin VB.Label lbl担保人 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "担保人"
                  Height          =   240
                  Left            =   660
                  TabIndex        =   117
                  Top             =   90
                  Width           =   720
               End
               Begin VB.Label lbl担保时限 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "到期时间"
                  Height          =   240
                  Left            =   6015
                  TabIndex        =   116
                  Top             =   90
                  Width           =   960
               End
               Begin VB.Label lbl担保原因 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "担保原因"
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
            Begin VB.CommandButton cmd籍贯 
               Caption         =   "…"
               Height          =   300
               Left            =   14160
               TabIndex        =   34
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   2610
               Width           =   300
            End
            Begin VB.CommandButton cmd区域 
               Caption         =   "…"
               Height          =   300
               Left            =   10965
               TabIndex        =   40
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   3000
               Width           =   300
            End
            Begin VB.CommandButton cmd出生地点 
               Caption         =   "…"
               Height          =   300
               Left            =   5610
               TabIndex        =   37
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   3000
               Width           =   315
            End
            Begin VB.CommandButton cmd户口地址 
               Caption         =   "…"
               Height          =   300
               Left            =   5625
               TabIndex        =   30
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   2610
               Width           =   300
            End
            Begin VB.CommandButton cmd家庭地址 
               Caption         =   "…"
               Height          =   300
               Left            =   5625
               TabIndex        =   25
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   2220
               Width           =   300
            End
            Begin VB.CommandButton cmdName 
               Caption         =   "…"
               Height          =   300
               Left            =   9150
               TabIndex        =   212
               TabStop         =   0   'False
               ToolTipText     =   "热键：F3"
               Top             =   270
               Width           =   300
            End
            Begin VB.CommandButton cmdSelectNO 
               Caption         =   "…"
               Height          =   300
               Left            =   5625
               TabIndex        =   211
               TabStop         =   0   'False
               ToolTipText     =   "热键:F8 缺号选择"
               Top             =   270
               Width           =   300
            End
            Begin VB.TextBox txt户口地址 
               Height          =   360
               Left            =   1590
               MaxLength       =   100
               TabIndex        =   29
               Top             =   2580
               Width           =   4335
            End
            Begin VB.TextBox txt户口地址邮编 
               Height          =   360
               Left            =   9555
               MaxLength       =   6
               TabIndex        =   32
               Top             =   2580
               Width           =   1725
            End
            Begin VB.TextBox txt籍贯 
               Height          =   360
               Left            =   11805
               MaxLength       =   50
               TabIndex        =   33
               Top             =   2580
               Width           =   2685
            End
            Begin VB.TextBox txt区域 
               Height          =   360
               Left            =   7770
               MaxLength       =   50
               TabIndex        =   39
               Top             =   2970
               Width           =   3525
            End
            Begin VB.TextBox txt出生地点 
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
                  Name            =   "宋体"
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
               Begin VB.TextBox txt联系人身份证号 
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
               Begin VB.CommandButton cmd工作单位 
                  Caption         =   "…"
                  Height          =   300
                  Left            =   5565
                  TabIndex        =   43
                  TabStop         =   0   'False
                  ToolTipText     =   "热键：F3"
                  Top             =   30
                  Width           =   315
               End
               Begin VB.CommandButton cmd联系人地址 
                  Caption         =   "…"
                  Height          =   300
                  Left            =   14145
                  TabIndex        =   55
                  TabStop         =   0   'False
                  ToolTipText     =   "热键：F3"
                  Top             =   1200
                  Width           =   300
               End
               Begin VB.TextBox txt联系人地址 
                  Height          =   360
                  Left            =   7725
                  MaxLength       =   100
                  TabIndex        =   54
                  Top             =   1170
                  Width           =   6750
               End
               Begin VB.TextBox txt工作单位 
                  Height          =   360
                  Left            =   1545
                  MaxLength       =   100
                  TabIndex        =   42
                  Top             =   0
                  Width           =   4350
               End
               Begin VB.TextBox txt单位开户行 
                  Height          =   360
                  Left            =   1545
                  MaxLength       =   50
                  TabIndex        =   46
                  Top             =   390
                  Width           =   4350
               End
               Begin VB.TextBox txt单位帐号 
                  Height          =   360
                  Left            =   7725
                  MaxLength       =   50
                  TabIndex        =   47
                  Top             =   390
                  Width           =   3525
               End
               Begin VB.TextBox txt联系人姓名 
                  Height          =   360
                  Left            =   7725
                  MaxLength       =   64
                  TabIndex        =   51
                  Top             =   780
                  Width           =   3525
               End
               Begin VB.ComboBox cbo联系人关系 
                  Height          =   360
                  Left            =   1545
                  Style           =   2  'Dropdown List
                  TabIndex        =   49
                  Top             =   780
                  Width           =   2310
               End
               Begin VB.TextBox txt联系人电话 
                  Height          =   360
                  Left            =   12720
                  MaxLength       =   20
                  TabIndex        =   52
                  Top             =   780
                  Width           =   1725
               End
               Begin VB.TextBox txt单位邮编 
                  Height          =   360
                  Left            =   12720
                  MaxLength       =   6
                  TabIndex        =   45
                  Top             =   0
                  Width           =   1725
               End
               Begin VB.TextBox txt单位电话 
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
                  Tag             =   "联系人地址"
                  Top             =   1170
                  Visible         =   0   'False
                  Width           =   6750
                  _ExtentX        =   11906
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
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
                  Caption         =   "手机号"
                  Height          =   240
                  Left            =   11970
                  TabIndex        =   220
                  Top             =   450
                  Width           =   720
               End
               Begin VB.Label lbl联系人身份证 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "联系人身份证"
                  Height          =   240
                  Left            =   0
                  TabIndex        =   216
                  Top             =   1230
                  Width           =   1440
               End
               Begin VB.Label lbl单位开户行 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "单位开户行"
                  Height          =   240
                  Left            =   240
                  TabIndex        =   113
                  Top             =   450
                  Width           =   1200
               End
               Begin VB.Label lbl单位帐号 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "单位帐号"
                  Height          =   240
                  Left            =   6600
                  TabIndex        =   112
                  Top             =   450
                  Width           =   960
               End
               Begin VB.Label lbl联系人姓名 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "联系人姓名"
                  Height          =   240
                  Left            =   6360
                  TabIndex        =   111
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label lbl联系人关系 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "联系人关系"
                  Height          =   240
                  Left            =   210
                  TabIndex        =   110
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label lbl联系人地址 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "联系人地址"
                  Height          =   240
                  Left            =   6360
                  TabIndex        =   109
                  Top             =   1230
                  Width           =   1200
               End
               Begin VB.Label lbl联系人电话 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "联系人电话"
                  Height          =   240
                  Left            =   11490
                  TabIndex        =   108
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label lbl工作单位 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "工作单位"
                  Height          =   240
                  Left            =   480
                  TabIndex        =   107
                  Top             =   60
                  Width           =   960
               End
               Begin VB.Label lbl单位邮编 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "单位邮编"
                  Height          =   240
                  Left            =   11730
                  TabIndex        =   106
                  Top             =   60
                  Width           =   960
               End
               Begin VB.Label lbl单位电话 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "单位电话"
                  Height          =   240
                  Left            =   6600
                  TabIndex        =   105
                  Top             =   60
                  Width           =   960
               End
            End
            Begin VB.TextBox txt身份证号 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   7755
               TabIndex        =   6
               Top             =   630
               Width           =   2340
            End
            Begin VB.ComboBox cbo身份 
               Height          =   360
               Left            =   12765
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   1410
               Width           =   1725
            End
            Begin VB.ComboBox cbo民族 
               Height          =   360
               Left            =   10080
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   1410
               Width           =   1725
            End
            Begin VB.ComboBox cbo国籍 
               Height          =   360
               Left            =   7755
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   1410
               Width           =   1725
            End
            Begin VB.ComboBox cbo病人类型 
               Height          =   360
               Left            =   12765
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   2970
               Width           =   1740
            End
            Begin VB.CommandButton cmdYB 
               Caption         =   "验证"
               Height          =   350
               Left            =   9555
               TabIndex        =   3
               TabStop         =   0   'False
               ToolTipText     =   "热键:F12(医保病人验证)"
               Top             =   245
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.TextBox txt住院号 
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
               ToolTipText     =   "请输入病人标识或姓名查找,直接回车登记新病人,定位热键:F11"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txt姓名 
               Height          =   360
               Left            =   7740
               MaxLength       =   64
               TabIndex        =   2
               ToolTipText     =   "输入病人姓名,或直接回车验证医保病人,如果是查找以前的病人,请在病人输入框输入"
               Top             =   240
               Width           =   1725
            End
            Begin VB.ComboBox cbo费别 
               Height          =   360
               Left            =   10065
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   1020
               Width           =   1725
            End
            Begin VB.ComboBox cbo职业 
               Height          =   360
               Left            =   12765
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1800
               Width           =   1725
            End
            Begin VB.ComboBox cbo学历 
               Height          =   360
               Left            =   7755
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   1800
               Width           =   1725
            End
            Begin VB.ComboBox cbo婚姻状况 
               Height          =   360
               Left            =   10080
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   1800
               Width           =   1725
            End
            Begin VB.TextBox txt家庭地址 
               Height          =   360
               Left            =   1590
               MaxLength       =   100
               TabIndex        =   24
               Top             =   2190
               Width           =   4335
            End
            Begin VB.TextBox txt家庭地址邮编 
               Height          =   360
               Left            =   12765
               MaxLength       =   6
               TabIndex        =   28
               Top             =   2190
               Width           =   1725
            End
            Begin VB.TextBox txt家庭电话 
               Height          =   360
               Left            =   9555
               MaxLength       =   20
               TabIndex        =   27
               Top             =   2190
               Width           =   1725
            End
            Begin VB.TextBox txt年龄 
               Height          =   360
               IMEMode         =   2  'OFF
               Left            =   4290
               TabIndex        =   12
               Top             =   1020
               Width           =   915
            End
            Begin VB.ComboBox cbo性别 
               Height          =   360
               ItemData        =   "frmHosReg.frx":0446
               Left            =   7755
               List            =   "frmHosReg.frx":0448
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   1020
               Width           =   1725
            End
            Begin VB.TextBox txt医保号 
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   1590
               MaxLength       =   30
               TabIndex        =   5
               Top             =   630
               Width           =   4335
            End
            Begin VB.ComboBox cbo年龄单位 
               Height          =   360
               Left            =   5235
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   1020
               Width           =   705
            End
            Begin VB.CommandButton cmdTurn 
               Caption         =   "门诊费用转住院(&T)"
               Height          =   350
               Left            =   10200
               TabIndex        =   4
               TabStop         =   0   'False
               ToolTipText     =   "热键:F12(医保病人验证)"
               Top             =   245
               Visible         =   0   'False
               Width           =   2160
            End
            Begin VB.TextBox txt险类 
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
            Begin VB.TextBox txt其他证件 
               Height          =   360
               Left            =   1590
               MaxLength       =   20
               TabIndex        =   20
               Top             =   1800
               Width           =   4335
            End
            Begin VB.ComboBox cbo医疗付款 
               Height          =   360
               Left            =   12765
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   1020
               Width           =   1725
            End
            Begin VB.TextBox txt支付密码 
               BeginProperty Font 
                  Name            =   "宋体"
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
            Begin VB.TextBox txt验证密码 
               BeginProperty Font 
                  Name            =   "宋体"
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
               ToolTipText     =   "快捷键F4"
               Top             =   240
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   635
               Appearance      =   2
               IDKindStr       =   $"frmHosReg.frx":044A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontSize        =   12
               FontName        =   "宋体"
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
            Begin MSMask.MaskEdBox txt出生时间 
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
                  Name            =   "宋体"
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
            Begin MSMask.MaskEdBox txt出生日期 
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
                  Name            =   "宋体"
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
               Tag             =   "出生地点"
               Top             =   2970
               Visible         =   0   'False
               Width           =   4350
               _ExtentX        =   7673
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
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
               Tag             =   "籍贯"
               Top             =   2580
               Visible         =   0   'False
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
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
               Tag             =   "现住址"
               Top             =   2190
               Visible         =   0   'False
               Width           =   6270
               _ExtentX        =   11060
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
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
               Tag             =   "户口地址"
               Top             =   2580
               Visible         =   0   'False
               Width           =   6270
               _ExtentX        =   11060
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   100
            End
            Begin VB.Label lbl籍贯 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "籍贯"
               Height          =   240
               Left            =   11295
               TabIndex        =   149
               Top             =   2640
               Width           =   480
            End
            Begin VB.Label lbl户口地址邮编 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口地址邮编"
               Height          =   240
               Left            =   7995
               TabIndex        =   148
               Top             =   2640
               Width           =   1440
            End
            Begin VB.Label lbl户口地址 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "户口地址"
               Height          =   240
               Left            =   510
               TabIndex        =   147
               Top             =   2640
               Width           =   960
            End
            Begin VB.Label lblPatiColor 
               BeginProperty Font 
                  Name            =   "宋体"
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
            Begin VB.Label lbl身份证号 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份证号"
               Height          =   240
               Left            =   6675
               TabIndex        =   145
               Top             =   690
               Width           =   960
            End
            Begin VB.Label lbl身份 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份"
               Height          =   240
               Left            =   12255
               TabIndex        =   144
               Top             =   1470
               Width           =   480
            End
            Begin VB.Label lbl民族 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "民族"
               Height          =   240
               Left            =   9525
               TabIndex        =   143
               Top             =   1470
               Width           =   480
            End
            Begin VB.Label lbl国籍 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "国籍"
               Height          =   240
               Left            =   7125
               TabIndex        =   142
               Top             =   1470
               Width           =   480
            End
            Begin VB.Label lbl出生地点 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生地点"
               Height          =   240
               Left            =   510
               TabIndex        =   141
               Top             =   3030
               Width           =   960
            End
            Begin VB.Label lbl区域 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "区域"
               Height          =   240
               Left            =   7125
               TabIndex        =   140
               Top             =   3030
               Width           =   480
            End
            Begin VB.Label lblPatiType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "病人类型"
               Height          =   240
               Left            =   11775
               TabIndex        =   139
               Top             =   3030
               Width           =   960
            End
            Begin VB.Label lblUnUseful 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "▼隐藏"
               ForeColor       =   &H00FF0000&
               Height          =   855
               Left            =   45
               TabIndex        =   138
               Top             =   2490
               Width           =   300
            End
            Begin VB.Label lbl住院号 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号"
               Height          =   240
               Left            =   3255
               TabIndex        =   137
               Top             =   300
               Width           =   720
            End
            Begin VB.Label lbl病人ID 
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
            Begin VB.Label lbl姓名 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名"
               Height          =   240
               Left            =   7155
               TabIndex        =   135
               Top             =   300
               Width           =   480
            End
            Begin VB.Label lbl性别 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "性别"
               Height          =   240
               Left            =   7125
               TabIndex        =   134
               Top             =   1080
               Width           =   480
            End
            Begin VB.Label lbl费别 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "费别"
               Height          =   240
               Left            =   9540
               TabIndex        =   133
               Top             =   1080
               Width           =   480
            End
            Begin VB.Label lbl职业 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "职业"
               Height          =   240
               Left            =   12255
               TabIndex        =   132
               Top             =   1860
               Width           =   480
            End
            Begin VB.Label lbl学历 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "学历"
               Height          =   240
               Left            =   7125
               TabIndex        =   131
               Top             =   1860
               Width           =   480
            End
            Begin VB.Label lbl婚姻状况 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "婚姻"
               Height          =   240
               Left            =   9555
               TabIndex        =   130
               Top             =   1860
               Width           =   480
            End
            Begin VB.Label lbl家庭地址 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "现住址"
               Height          =   240
               Left            =   750
               TabIndex        =   129
               Top             =   2250
               Width           =   720
            End
            Begin VB.Label lbl家庭电话 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭电话"
               Height          =   240
               Left            =   8475
               TabIndex        =   128
               Top             =   2250
               Width           =   960
            End
            Begin VB.Label lbl家庭地址邮编 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "家庭地址邮编"
               Height          =   240
               Left            =   11295
               TabIndex        =   127
               Top             =   2250
               Width           =   1440
            End
            Begin VB.Label lbl出生日期 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "出生日期"
               Height          =   240
               Left            =   510
               TabIndex        =   126
               Top             =   1080
               Width           =   960
            End
            Begin VB.Label lbl年龄 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄"
               Height          =   240
               Left            =   3765
               TabIndex        =   125
               Top             =   1080
               Width           =   480
            End
            Begin VB.Label lbl医保号 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "医保号"
               Height          =   240
               Left            =   750
               TabIndex        =   124
               Top             =   690
               Width           =   720
            End
            Begin VB.Label lbl险类 
               Alignment       =   1  'Right Justify
               Caption         =   "险类名称"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   675
               TabIndex        =   123
               Top             =   1470
               Width           =   825
            End
            Begin VB.Label lbl其它证件 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "其他证件"
               Height          =   240
               Left            =   510
               TabIndex        =   122
               Top             =   1860
               Width           =   960
            End
            Begin VB.Label lbl医疗付款 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "付费方式"
               Height          =   240
               Left            =   11775
               TabIndex        =   121
               Top             =   1080
               Width           =   960
            End
            Begin VB.Label lbl支付密码 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "密码"
               Height          =   240
               Left            =   11445
               TabIndex        =   120
               Top             =   690
               Width           =   480
            End
            Begin VB.Label lbl验证密码 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "验证"
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
         Name            =   "宋体"
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
         Caption         =   "担保信息(&S)"
         Height          =   400
         Left            =   1515
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   90
         Width           =   1845
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   400
         Left            =   255
         TabIndex        =   96
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   400
         Left            =   13440
         TabIndex        =   95
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   400
         Left            =   12240
         TabIndex        =   94
         Top             =   90
         Width           =   1100
      End
      Begin VB.CommandButton cmdDeposit 
         Caption         =   "预交款收取(&D)"
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
      Caption         =   "打印机设置(&S)"
      Visible         =   0   'False
      Begin VB.Menu mnu病案首页 
         Caption         =   "病案首页(&1)"
      End
      Begin VB.Menu mnu预交款收据 
         Caption         =   "预交款收据(&2)"
      End
   End
End
Attribute VB_Name = "frmHosReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit '要求变量声明
Public mstrPrivs As String
Public mlngModul As Long
Public mbytMode As Byte '入：0-正常登记,1-预约登记,2-接收预约   '如果每次入院使用新住院号,则预约时不变,接收时再产生(因为医嘱会产生预约)
Public mbytKind As Byte '入：0=住院入院登记,1-门诊留观登记,2-住院留观登记
Public mbytInState As Byte '入：0=新增,1=修改,2=查阅
'入：要查阅，修改，接收的病人ID、主页ID(预约的为0)
Public mlng病人ID As Long
Private mlng挂号ID As Long              '预约中心病人接收后回传接收状态时用
Public mlng主页ID As Long
'Private mstr预交NO As String
Private mrsInfo As ADODB.Recordset '病人信息
Private mrsPatiReg As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mrsUnit As ADODB.Recordset
Private mrsDept As ADODB.Recordset
Private mrsUnitDept As ADODB.Recordset  '病区科室对应
Private mrsInputSet  As ADODB.Recordset '输入项控制  字段名称:输入项目,禁止录入,必须输入,光标进入,控件名,控件下标

Private mblnICCard As Boolean 'IC卡发卡,要同时填写病人信息的IC卡字段
Private mblnOneCard As Boolean      '是否启用了一卡通接口,此模式下，票号严格管理，票号范围外的发卡或绑定卡不收费

Private mblnAuto As Boolean '
Private mblnUnload As Boolean
Private mlng预交领用ID As Long
Private mblnChange As Boolean
Private mbln是否扫描身份证 As Boolean

Private mblnPrepayPrint As Boolean    '是否打印预交款
Private mblnFPagePrint As Boolean   '是否打印病案主页
Private mblnWristletPrint As Boolean    '是否打印病人腕带
Private mdat上次担保到期时间 As Date '修改登记信息时,上次时限担保的到期时间
Private mstrNOS As String   '选择转入的单据,票据,结帐ID,险类(非医保为零):H0000001,F000023,81235,901;H0000002,F000045,81263,901;...
Private mobjKeyboard As Object
 
Private mblnHaveAdvice As Boolean   '当前病人是否存在医嘱

Private mstrPatiPlus    As String     '从表信息:信息名1:信息值1,信息名2:信息值2
Private mblnEMPI As Boolean               'T-找到EMPI病人,F-未找到EMPI病人
Private mblnAppoint As Boolean              'T-预约中心病人直接 入院入科
Private mstrAppointBed As String            '预约床位

'医保变量---------------
Private mintInsure As Integer
Private mstrYBPati As String
Private mcurYBMoney As Currency '个人帐户余额
'以下为合并病人是对应的记录变量
Private mintInsureBak As Integer
Private mstrYBPatiBak As String
Private mcurYBMoneyBak As Currency '个人帐户余额
Private mbytKindBak As Byte
Private mbln空床 As Boolean

Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML


Private Enum EState
    E新增 = 0
    E修改 = 1
    E查阅 = 2
End Enum
Private Enum EMode
    E正常登记 = 0
    E预约登记 = 1
    E接收预约 = 2
End Enum
Private Enum EKind
    E住院入院登记 = 0
    E门诊留观登记 = 1
    E住院留观登记 = 2
End Enum

'-----------------------------------------------------------------
'发票相关
Private mFactProperty As Ty_FactProperty
'-----------------------------------------------------------------
'医疗卡相关
'Private mobjSquareCard As Object
Private mblnClickSquareCtrl As Boolean
Private mblnStartFactUseType As Boolean '是否启用的相关的门诊类别的
Private mbytPrepayType As Byte '0-门诊住院;1-门诊;2-住院
Private mblnNotClick As Boolean
Private mblnIdNotClick  As Boolean
Private mblnICNotClick As Boolean
Private mblnCheckPatiCard As Boolean

Private mobjSquare As Object '医疗卡部件
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private Type Ty_CardProperty
       lng卡类别ID As Long
       str卡名称  As String
       lng卡号长度 As Long
       lng结算方式 As String
       bln自制卡 As Boolean
       bln严格控制 As Boolean
       lng领用ID As Long
       lng共用批次 As Long
       bln变价 As Boolean
       bln重复利用 As Boolean
       bln就诊卡 As Boolean
       str卡号密文 As String
       int密码长度 As Integer
       int密码长度限制 As Integer
       int密码规则 As Integer
       blnOneCard As Boolean '  '是否启用了一卡通接口,此模式下，票号严格管理，票号范围外的发卡或绑定卡不收费
       rs卡费 As ADODB.Recordset
       dbl应收金额 As Double
       dbl实收金额 As Double
       bln是否制卡 As Boolean '问题号:56599
       bln是否发卡 As Boolean
       bln是否写卡 As Boolean
       bln是否院外发卡  As Boolean
       lng发卡性质 As Long '0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0 问题号:57326
       str读卡性质 As String
       byt发卡控制 As Byte
       str特定项目 As String
End Type
Private mCurSendCard As Ty_CardProperty
Private mcolPrepayPayMode As Collection   '预交款支付方式
Private mcolCardPayMode As Collection   '就诊卡支付方式

Private Type Ty_PayMoney
    lng医疗卡类别ID As Long
    bln消费卡 As Boolean
    str结算方式 As String
    str名称 As String
    str刷卡卡号 As String
    str刷卡密码 As String
    strno As String
    lngID As Long '预交ID
    lng结帐ID As Long
End Type

Private mCurPrepay As Ty_PayMoney
Private mCurCardPay As Ty_PayMoney
Private mstrPassWord As String
Private mbln扫描身份证签约 As Boolean '根据参数设置中的“扫描身份证签约”取值
Private mstr缺省费别 As String
'问题号 :56599
Private Type Ty_PageHeight
    基本 As Long
    健康档案 As Long
End Type
Private mPageHeight As Ty_PageHeight
Private mstrPriceGrade As String, mstrPrePriceGrade As String
Private mobjPublicExpense As Object  '费用公共部件
Private mintPriceGradeStartType As Integer

Private mstrCboSplit As String
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const C_ColumHeader = "过敏药物,1,5000,1;过敏反映,4,3000,1;过敏药物ID,1,100,0" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_InoculateHeader = "接种日期,4,3500,1;接种名称,4,3500,1;接种日期,4,3500,1;接种名称,4,3500,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_LinkManColumHeader = "联系人姓名,4,3000,1;联系人关系,4,3000,1;联系人关系备注,4,2000,1;联系人身份证号,4,3000,1;联系人电话,4,3000,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_OtherInfoColumHeader = "信息名,4,3600,1;信息值,4,3600,1;信息名,4,3600,1;信息值,4,3600,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_CertificateHeader = "证件类型,4,3500,1;证件号码,4,3500,1;证件类型,4,3500,1;证件号码,4,3500,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_BH = "阴,阳,不详,未查"
'C_输入项控制格式:参数名,控件(控件1,控件2,...)|参数2,控件|...
Private Const C_输入项控制 = "国籍,cbo国籍|民族,cbo民族|学历,cbo学历|婚姻状况,cbo婚姻状况|职业,cbo职业|身份,cbo身份|出生日期,txt出生日期,txt出生时间|其他证件,txt其他证件" & _
                        "|身份证号,txt身份证号,cboIDNumber|出生地点,txt出生地点,PatiAddress(1)|现住址,txt家庭地址,PatiAddress(3)|家庭地址邮编,txt家庭地址邮编|家庭电话,txt家庭电话|联系人姓名,txt联系人姓名|联系人关系,cbo联系人关系,txtLinkManInfo" & _
                        "|户口地址,txt户口地址,PatiAddress(4)|户口地址邮编,txt户口地址邮编|区域,txt区域|联系人地址,txt联系人地址,PatiAddress(5)|联系人电话,txt联系人电话|联系人身份证号,txt联系人身份证号" & _
                        "|工作单位,txt工作单位|单位电话,txt单位电话|单位邮编,txt单位邮编|单位开户行,txt单位开户行|单位帐号,txt单位帐号|籍贯,txt籍贯,PatiAddress(2)"
Private Const C_COLOR_UNEnabled = &H80000004 '禁止录入颜色
Private Const C_COLOR_Enabled = &H80000005 '不禁止录入显示颜色

Private mdic医疗卡属性 As New Dictionary
Private mbln发卡或绑定卡 As Boolean
Private mbln是否显示预交 As Boolean
Private mbln是否显示磁卡 As Boolean
Private marrAddress(0 To 4) As String     '五级结构化地址缺省值
Private mstrFirstCode As String '第一种证件类型的编码
'-----------------------------------------------------------------
Private mintIDKind As String

Private Sub cbo病人类型_Click()
    If cbo病人类型.ListCount > 0 And cbo病人类型.ListIndex <> -1 Then
        lblPatiColor.BackColor = zlDatabase.GetPatiColor(zlCommFun.GetNeedName(cbo病人类型.Text))
        txt姓名.ForeColor = lblPatiColor.BackColor
    End If
End Sub
 

Private Sub cbo发卡结算_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    '问题号:48352
    With mCurCardPay
            .lng医疗卡类别ID = 0
            .bln消费卡 = False
            .str结算方式 = ""
            .str名称 = ""
     End With
    '0=新增,1=修改,2=查看
    If mbytInState = 2 Then Exit Sub
    Call SetCardVaribles(False)
    '130245,切换结算方式，同步更新卡类别ID
    If mblnNotClick = True Then Exit Sub
    Call Local结算方式(mCurCardPay.lng医疗卡类别ID, True)
End Sub

Private Sub cbo联系人关系_Click()
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("联系人关系") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("联系人关系")) = zlCommFun.GetNeedName(cbo联系人关系.Text)
    End If
    If zlCommFun.GetNeedName(cbo联系人关系.Text) = "其他" Then
        txtLinkManInfo.Enabled = True: txtLinkManInfo.BackColor = &H80000005
    Else
        txtLinkManInfo.Enabled = False: txtLinkManInfo.Text = "": txtLinkManInfo.BackColor = &H80000004
    End If
End Sub

Private Sub cbo联系人关系_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo联系人关系.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo联系人关系.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo联系人关系.ListIndex = lngIdx
End Sub

Private Sub cbo入院病区_Validate(Cancel As Boolean)
    '问题27370 by lesfeng 2010-01-26
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSQL As String, intIdx As Long, i As Long
    Dim lngDept As Long
    Dim intFlag As Integer
    
    strInput = UCase(cbo入院病区.Text)
    intFlag = -1
    If Trim(strInput) = "不确定病区" Then Cancel = False: Exit Sub
    If gbln先选病区 Then
        Set rsTmp = InputDept(Me, fra入院, cbo入院病区, "护理", IIf(mbytKind = EKind.E门诊留观登记, "1", "2") & ",3", strInput, blnCancel, intFlag, 0)
    Else
        If cbo入院科室.ListIndex >= 0 Then lngDept = cbo入院科室.ItemData(cbo入院科室.ListIndex)
        mrsUnitDept.Filter = "科室ID=" & lngDept
        If mrsUnitDept.RecordCount > 0 Then
            intFlag = 2
        Else
            lngDept = 0
        End If
        Set rsTmp = InputDept(Me, fra入院, cbo入院病区, "护理", IIf(mbytKind = EKind.E门诊留观登记, "1", "2") & ",3", strInput, blnCancel, intFlag, lngDept)
    End If
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cbo入院病区, rsTmp!ID)
        If intIdx <> -1 Then
            cbo入院病区.ListIndex = intIdx
'        Else
'            cbo入院病区.AddItem Nvl(rsTmp!编码) & "-" & Chr(13) & rsTmp!名称, cbo入院病区.ListCount - 1
'            cbo入院病区.ItemData(cbo入院病区.NewIndex) = rsTmp!ID
'            cbo入院病区.ListIndex = cbo入院病区.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的入院病区。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub cbo入院方式_Click()
    If zlCommFun.GetNeedName(cbo入院方式.Text) = "转入" Then
        cmd转入.Enabled = True: cmd转入.BackColor = &H80000005
        txt转入.Enabled = True: txt转入.BackColor = &H80000005
        lblInFrom.Enabled = True
    Else
        cmd转入.Enabled = False:  cmd转入.BackColor = &H80000004
        txt转入.Enabled = False: txt转入.Text = "": txt转入.BackColor = &H80000004
        lblInFrom.Enabled = False
    End If
End Sub

'Private Sub cbo入院科室_GotFocus()
'    '问题27370 by lesfeng 2010-01-26
''    If cbo入院科室.Style = 0 Then
''        Call zlcontrol.TxtSelAll(cbo入院科室)
''    End If
'    With cbo入院科室
'        .SelStart = 0
'        .SelLength = Len(.Text)
'    End With
'End Sub

Private Sub cbo入院科室_Validate(Cancel As Boolean)
    '问题27370 by lesfeng 2010-01-26
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSQL As String, intIdx As Long, i As Long
    Dim lngUnit As Long
    
    strInput = UCase(cbo入院科室.Text)
    If gbln先选病区 Then
         If cbo入院病区.ListIndex >= 0 Then lngUnit = cbo入院病区.ItemData(cbo入院病区.ListIndex)
        Set rsTmp = InputDept(Me, fra入院, cbo入院科室, "临床", IIf(mbytKind = EKind.E门诊留观登记, "1", "2") & ",3", strInput, blnCancel, 1, lngUnit)
    Else
        Set rsTmp = InputDept(Me, fra入院, cbo入院科室, "临床", IIf(mbytKind = EKind.E门诊留观登记, "1", "2") & ",3", strInput, blnCancel, -1, 0)
    End If
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cbo入院科室, rsTmp!ID)
        If intIdx <> -1 Then
            cbo入院科室.ListIndex = intIdx
'        Else
'            cbo入院科室.AddItem Nvl(rsTmp!编码) & "-" & Chr(13) & rsTmp!名称, cbo入院科室.ListCount - 1
'            cbo入院科室.ItemData(cbo入院科室.NewIndex) = rsTmp!ID
'            cbo入院科室.ListIndex = cbo入院科室.NewIndex
        End If
    Else
        If cbo入院科室.ListIndex = -1 And cbo入院科室.ListCount = 0 Then
        Else
            If Not blnCancel Then
                MsgBox "未找到对应的入院科室。", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Sub cbo医疗付款_Click()
    On Error GoTo ErrHandler
    If mintPriceGradeStartType < 2 Then Exit Sub
    Call mobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cbo医疗付款.Text), , , mstrPriceGrade)
    If mstrPrePriceGrade = mstrPriceGrade Then Exit Sub
    mstrPrePriceGrade = mstrPriceGrade

    If mCurSendCard.str特定项目 <> "" Then
        Set mCurSendCard.rs卡费 = zlGetSpecialItemFee(mCurSendCard.str特定项目, mstrPriceGrade)
    Else
        Set mCurSendCard.rs卡费 = Nothing
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
    '功能:加载卡费
    On Error GoTo errHandle
    If mCurSendCard.rs卡费 Is Nothing Then
        txt卡额.Text = "": txt卡额.Tag = ""
        Exit Sub
    End If
    If mCurSendCard.rs卡费.RecordCount = 0 Then
        txt卡额.Text = "": txt卡额.Tag = ""
        Exit Sub
    End If
    
    With mCurSendCard.rs卡费
        txt卡额.Text = Format(IIf(Nvl(!是否变价, 0) = 1, Val(Nvl(!缺省价格)), Val(Nvl(!现价))), "0.00")
        If Nvl(!是否变价, 0) <> 1 And Nvl(!屏蔽费别, 0) <> 1 Then
            txt卡额.Text = Format(GetActualMoney(zlStr.NeedName(cbo费别.Text), !收入项目ID, Val(txt卡额.Text), !收费细目ID), "0.00")
        End If
        txt卡额.Tag = txt卡额.Text  '保持不变
        txt卡额.Locked = Nvl(!是否变价, 0) <> 1
        txt卡额.TabStop = Nvl(!是否变价, 0) = 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub chk单位缴款_Click()
    If chk单位缴款.Value = 1 Then
        txt缴款单位.Enabled = True
        txt缴款单位.BackColor = &H80000005
    Else
        txt缴款单位.Text = ""
        txt缴款单位.Enabled = False
        txt缴款单位.BackColor = Me.BackColor
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
    
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    Call gobjPatient.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 0, mlng病人ID, mlng主页ID, 0, 0)
    Call GlobalDeleteAtom(intAtom)
    If gbln入院预交 Then
        If gblnPrepayStrict Then
            mlng预交领用ID = CheckUsedBill(2, IIf(mlng预交领用ID > 0, mlng预交领用ID, mFactProperty.lngShareUseID), , 2)
            If mlng预交领用ID <= 0 Then
                Select Case mlng预交领用ID
                    Case 0 '操作失败
                    Case -1
                        MsgBox "你没有自用和共用的预交票据,病人入院时不能同时缴预交款！" & _
                            "请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                    Case -2
                        MsgBox "本地的共用票据已经用完,病人入院时不能同时缴预交款！" & _
                            "请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End Select
                txtFact.Text = ""
            Else
                txtFact.Text = GetNextBill(mlng预交领用ID)
            End If
        Else
            '松散：取下一个号码
            txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngModul, "")))
        End If
    End If
End Sub

Private Sub cmdName_Click()
    Dim rsTmp As ADODB.Recordset
    '读取病人信息
    Set rsTmp = GetPatientByName(txt姓名.Text)
    Call MergePatient(rsTmp, 0)
End Sub

Private Sub cmdSelectNO_Click()
    Dim strno As String
    
    Call frmNOSelect.ShowMe(Me, strno)
    If strno <> "" Then txt住院号.Text = strno
    If txt姓名.Enabled And txt姓名.Visible Then txt姓名.SetFocus
End Sub

Private Sub cmdSurety_Click()
    frmSurety.mlng病人ID = 0
    frmSurety.mbln在院病人 = True
    frmSurety.mstrPrivs = mstrPrivs
    frmSurety.Show 1, Me
End Sub

Private Sub cmdTurn_Click()
    Call frmChargeTurn.ShowMe(Me, Val(txtPatient.Text), mstrNOS, , , mstrPrivs, mlngModul)
End Sub

Private Sub cmd户口地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt户口地址, True)
    If Not rsTmp Is Nothing Then
        txt户口地址.Text = rsTmp!名称
        txt户口地址.SelStart = Len(txt户口地址.Text)
        txt户口地址.SetFocus
    End If
End Sub

Private Sub cmd籍贯_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt籍贯, True)
    If Not rsTmp Is Nothing Then
        txt籍贯.Text = rsTmp!名称
        txt籍贯.SelStart = Len(txt籍贯.Text)
        txt籍贯.SetFocus
    Else
        zlControl.TxtSelAll txt籍贯
        txt籍贯.SetFocus
    End If
End Sub

Private Sub cmd区域_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt区域, True)
    If Not rsTmp Is Nothing Then
        txt区域.Text = rsTmp!名称
        txt区域.SelStart = Len(txt区域.Text)
        txt区域.SetFocus
    Else
        zlControl.TxtSelAll txt区域
        txt区域.SetFocus
    End If
End Sub

Private Sub cmd转入_Click()
    Dim vPoint As POINTAPI
    On Error GoTo errH
    vPoint = GetCoordPos(txt转入.Container.hWnd, txt转入.Left, txt转入.Top)
    Call Get医疗机构(txt转入, Me, 2, "医疗机构", "字典管理工具", vPoint, False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    '问题号:53408
    '问题号:53408
    mbln扫描身份证签约 = IIf(zlDatabase.GetPara("扫描身份证签约", glngSys, glngModul) = "1", 1, 0) = "1"
    If mCurSendCard.str卡名称 Like "*二代身份证*" Then
        lbl卡号.Enabled = False: txt卡号.Enabled = False
        lbl密码.Enabled = False: txtPass.Enabled = False
        lbl验证.Enabled = False: txtAudi.Enabled = False
    End If
    Call Show绑定控件(False)
    If gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""    '双屏显示窗体必须在当前窗口显示之后调用显示才能移动窗体
    End If
    
    SetCardEditEnabled
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
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
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub

    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, False, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    '问题号:56599
    If strOutPatiInforXML <> "" Then Call LoadPati(strOutPatiInforXML)
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '是否密文显示
    'txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55571:刘鹏飞,2012-11-12
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And (Not mblnIdNotClick And Not mblnICNotClick) Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub
Private Sub lblUnUseful_Click()
    Dim lngTop As Long
     '51167,刘鹏飞,2012-07-09,增加"联系人身份证号"
    lblUnUseful.Appearance = IIf(lblUnUseful.Appearance = 0, 1, 0)
    picUnUseful.Tag = IIf(lblUnUseful.Appearance = 0, 1, 0)
    If lblUnUseful.Appearance = 0 Then
        fra病人.Height = fra病人.Height - picUnUseful.Height - IIf(pic担保.Visible = False, pic担保.Height, 0)
        pic病人.Height = pic病人.Height - picUnUseful.Height - IIf(pic担保.Visible = False, pic担保.Height, 0)
        pic担保.Top = picUnUseful.Top
        Me.Height = Me.Height - picUnUseful.Height - IIf(pic担保.Visible = False, pic担保.Height, 0)
        tbcPage.Height = picCmd.Top
        picUnUseful.Visible = False
        lblUnUseful.Caption = "▼显示"
    ElseIf lblUnUseful.Appearance = 1 Then
        fra病人.Height = fra病人.Height + picUnUseful.Height + IIf(pic担保.Visible = False, pic担保.Height, 0)
        pic病人.Height = pic病人.Height + picUnUseful.Height + IIf(pic担保.Visible = False, pic担保.Height, 0)
        pic担保.Top = pic担保.Top + picUnUseful.Height + 35
        Me.Height = Me.Height + picUnUseful.Height + IIf(pic担保.Visible = False, pic担保.Height, 0)
        tbcPage.Height = picCmd.Top
        picUnUseful.Visible = True
        lblUnUseful.Caption = "▲隐藏"
    End If
    pic入院.Top = pic病人.Top + pic病人.Height
    lngTop = pic入院.Top + pic入院.Height
    If mbln是否显示预交 Then
        pic预交.Top = lngTop
        lngTop = pic预交.Top + pic预交.Height
    End If
    pic磁卡.Top = lngTop
            
'    pic入院.Top = pic病人.Top + pic病人.Height
'    pic预交.Top = pic入院.Top + pic入院.Height
'    pic磁卡.Top = pic预交.Top + pic预交.Height
    lblUnUseful.ForeColor = &HFF0000
    mPageHeight.基本 = Me.Height
End Sub

Private Sub lbl卡号_Click()
    Dim strExpand As String, strOutCardNO As String, strOutPatiInforXML As String

    If mCurSendCard.bln就诊卡 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt卡号.Text = mobjICCard.Read_Card()
            If txt卡号.Text <> "" Then mblnICCard = True
        End If
        Exit Sub
    End If
    If (Mid(mCurSendCard.str读卡性质, 3, 1) = 0 And Mid(mCurSendCard.str读卡性质, 4, 1) = 0) Or mCurSendCard.lng卡类别ID = 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\

    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, mCurSendCard.lng卡类别ID, False, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txt卡号.Text = strOutCardNO
    If txt卡号.Text <> "" Then
        '问题号:56599
       If strOutPatiInforXML <> "" Then Call LoadPati(strOutPatiInforXML)
       Call CheckFreeCard(txt卡号.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    End If
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNO As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt卡号.Text = strCardNO
    If txt卡号.Text <> "" Then
        '问题号:56599
       If strXmlCardInfor <> "" Then Call LoadPati(strXmlCardInfor)
       Call CheckFreeCard(txt卡号.Text)
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
    Else
        If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
    End If
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    Dim lngPreIDKind As Long, lngIndex As Long
    
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnICNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("IC卡号")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strCardNO
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text <> "" Then
            Call mobjICCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
        End If
        
        IDKind.IDKind = lngPreIDKind
        mblnICNotClick = False
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    Dim lngIndex As Long
    Dim bln签约 As Boolean
    Dim strErrMsg As String
'    '问题号:53408
'    mbln是否扫描身份证 = True
'
'    txt身份证号.Text = strID
'    If mCurSendCard.str卡名称 = "二代身份证" Then
'        txt卡号.Text = strID
'        Exit Sub
'    End If
    
    mbln是否扫描身份证 = False
    
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        mblnIdNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("身份证号")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        
        '57945:刘鹏飞,2013-10-30,读取身份证中的地址应该放到户口地址而不是家庭地址
        If mrsInfo Is Nothing Then
            lngIndex = IDKind.GetKindIndex("姓名")
            If lngIndex >= 0 Then IDKind.IDKind = lngIndex
            txtPatient.Text = ""
            Call txtPatient_KeyPress(vbKeyReturn)
            txt姓名.Text = strName
            Call cbo.Locate(cbo性别, strSex)
            Call cbo.Locate(cbo民族, strNation)
            txt出生日期.Text = Format(datBirthDay, "yyyy-MM-dd")
            txt身份证号.Text = strID
        End If
        '101692新病人直接提取;已经建档病人户口地址为空时,从身份证获取
        If mrsInfo Is Nothing Or (Not mrsInfo Is Nothing And Trim(txt户口地址.Text) = "") Then
            txt户口地址.Text = strAddress
            If gbln启用结构化地址 Then
                PatiAddress(E_IX_户口地址).Value = strAddress
            End If
        End If
        IDKind.IDKind = lngPreIDKind
        mblnIdNotClick = False
        
        If (mCurSendCard.str卡名称 = "二代身份证" Or mbln扫描身份证签约) Then
            bln签约 = 是否已经签约(Trim(strID))
            '如果没有签约,检查姓名 性别,生日等情况
            If Not bln签约 And Not mrsInfo Is Nothing Then
                  If Nvl(mrsInfo!姓名) <> Trim(strName) Or Nvl(mrsInfo!性别) <> strSex Or Format(txt出生日期.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
            
                      If Nvl(mrsInfo!姓名) <> Trim(strName) Then
                           strErrMsg = strErrMsg & "," & "姓名"
                      End If
                      If Nvl(mrsInfo!性别) <> strSex Then
                           strErrMsg = strErrMsg & "," & "性别"
                      End If
                      If Format(txt出生日期.Text, "yyyy-MM-dd") <> Format(datBirthDay, "yyyy-MM-dd") Then
                          strErrMsg = strErrMsg & "," & "出生日期"
                      End If
                      strErrMsg = Mid(strErrMsg, 2)
                      strErrMsg = "当前病人信息与身份证上的[" & strErrMsg & "]等信息不一致!" & vbCrLf & "不能进行身份证签约操作!"
                      Call MsgBox(strErrMsg, vbQuestion, Me.Caption)
                       mbln是否扫描身份证 = False
                  Else
                       mbln是否扫描身份证 = True
                  End If
            ElseIf Not bln签约 Then
                mbln是否扫描身份证 = True
            End If
            
        End If
    End If
    
    

    If Me.ActiveControl Is txt身份证号 Then
        
        If txt姓名.Text <> "" And cbo性别.ListCount <> 0 And txt出生日期.Text <> "" Then
            If strName <> txt姓名.Text Or strSex <> Split(cbo性别.Text, "-")(1) Or txt出生日期.Text <> Format(datBirthDay, "yyyy-MM-dd") Then
                    MsgBox "身份证信息与挂号病人信息不一致,不能进行签约操作！", vbInformation, gstrSysName
                    Exit Sub
            End If
        Else
             MsgBox "绑定二代身份证时,病人信息不允许为空！", vbInformation, gstrSysName
             Exit Sub
        End If
        
        If 是否已经签约(Trim(strID)) Then
            MsgBox "身份证号码为:" & strID & "已经签约不能重复签约！", vbOKOnly + vbInformation, gstrSysName
            txt身份证号.SetFocus
            Exit Sub
        Else
            txt身份证号.Text = strID
            mbln是否扫描身份证 = True
        End If
        
    End If
    
    Call Show绑定控件(mbln是否扫描身份证 And mbln扫描身份证签约)
End Sub

Public Sub Show绑定控件(blnShow As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否显示绑定密码
    '入参:blnShow 是否显示绑定密码
    '编制:王吉
    '日期:2012-09-04 15:53:14
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    lbl支付密码.Enabled = blnShow: txt支付密码.Enabled = blnShow
    lbl验证密码.Enabled = blnShow: txt验证密码.Enabled = blnShow
    If blnShow = False Then
        txt支付密码.Text = "": txt验证密码.Text = "": txt验证密码.Tag = ""
    End If
    
End Sub

Private Sub cbo门诊医师_Validate(Cancel As Boolean)
    Dim strDoctor As String
    Dim blnFinded As Boolean
    
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSQL As String, intIdx As Long, i As Long
    Dim lngUnit As Long
    
    If mbytMode = 1 Then Exit Sub
    
    '问题27370 by lesfeng 2010-02-03
    strInput = UCase(cbo门诊医师.Text)
    Set rsTmp = InputDoctors(Me, fra入院, cbo门诊医师, 0, "1,2,3", strInput, blnCancel, "")

    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cbo门诊医师, rsTmp!ID)
        If intIdx <> -1 Then
            cbo门诊医师.ListIndex = intIdx
'        Else
'            cbo门诊医师.AddItem Nvl(rsTmp!简码) & "-" & Chr(13) & rsTmp!姓名, cbo入院病区.ListCount - 1
'            cbo门诊医师.ItemData(cbo门诊医师.NewIndex) = rsTmp!ID
'            cbo门诊医师.ListIndex = cbo门诊医师.NewIndex
        End If
    Else
        Call zlControl.TxtSelAll(cbo门诊医师)
        If Not blnCancel Then
            cbo门诊医师.Text = ""
'            MsgBox "未找到对应的医生。", vbInformation, gstrSysName
        End If
'        Cancel = True: Exit Sub
    End If
    
   
'    If cbo门诊医师.Locked Then Exit Sub
'    If cbo门诊医师.ListCount = 0 Then cbo门诊医师.Text = "": Exit Sub
'
'    strDoctor = cbo门诊医师.Text
'
'    If mrsDoctor.State = 1 Then
'        If mrsDoctor.RecordCount = 0 Then cbo门诊医师.Text = "": Exit Sub
'        mrsDoctor.MoveFirst
'        For i = 1 To mrsDoctor.RecordCount
'            If UCase(strDoctor) = mrsDoctor!编号 Or strDoctor = mrsDoctor!姓名 Or UCase(strDoctor) = mrsDoctor!简码 Or strDoctor = mrsDoctor!简码 & "-" & mrsDoctor!姓名 Then
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
'        If Not Cbo.Locate(cbo门诊医师, strDoctor, True) Then
'            Call zlcontrol.TxtSelAll(cbo门诊医师)

'            Cancel = True
'        End If
'    Else
'        Call zlcontrol.TxtSelAll(cbo门诊医师)
'        Cancel = mrsDoctor.State = 1 And txtPatient.Text <> ""   '没有数据时允许离开焦点
'        If Not Cancel Then cbo门诊医师.Text = ""
'    End If
End Sub

Private Sub cbo年龄单位_LostFocus()
    '68489:刘鹏飞,2013-12-06,年龄为空则不进行出生日反算
    Dim strBirth As String
    Dim strMsg As String
    Dim lngTmp As Long
    
    If Trim(txt年龄.Text) = "" Then Exit Sub
    If Not CheckOldData(txt年龄, cbo年龄单位) Then Exit Sub
    
    If Not IsDate(txt出生日期.Text) Then
        mblnChange = False
        Call ReCalcBirthDay(strMsg)
        mblnChange = True
        If InStr(1, strMsg, "|") > 0 Then
            lngTmp = Val(Split(strMsg, "|")(0)) '1禁止,0提示
            strMsg = Split(strMsg, "|")(1)
            If lngTmp = 1 Then
                MsgBox strMsg, vbInformation, gstrSysName
                If CanFocus(txt年龄) = True Then txt年龄.SetFocus: Exit Sub
            End If
        End If
    End If
    Call ReLoadCardFee
End Sub

Private Sub cbo入院病区_Click()
    Dim lngDepID As Long
    Dim rsDiagnosis As ADODB.Recordset
    
    If cbo入院病区.ListIndex <> -1 Then
        If mbytInState <> EState.E查阅 Then Call LoadDept(1)

        cbo床位.TabStop = (cbo入院病区.Text = "不确定病区")
        '107823显示病人诊断情况
        If cbo入院科室.ListIndex <> -1 Then
            lngDepID = cbo入院科室.ItemData(cbo入院科室.ListIndex)
            If (mbytMode = EMode.E接收预约 Or mbytMode = EMode.E正常登记 And mlng病人ID <> 0 And mbytInState = EState.E新增) And Me.Visible = True Then
                Set rsDiagnosis = GetDiagnosticInfo(mlng病人ID, mlng主页ID, "1,11", "3", lngDepID)
                If Not rsDiagnosis Is Nothing Then
                    rsDiagnosis.Filter = "诊断类型=1"
                    If Not rsDiagnosis.EOF Then
                        txt门诊诊断.Text = Nvl(rsDiagnosis!诊断描述): txt门诊诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl门诊诊断.Tag = txt门诊诊断.Text
                    Else
                        txt门诊诊断.Text = ""
                    End If
                    If txt中医诊断.Enabled Then
                        rsDiagnosis.Filter = "诊断类型=11"
                        If Not rsDiagnosis.EOF Then
                            txt中医诊断.Text = Nvl(rsDiagnosis!诊断描述): txt中医诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                        End If
                    Else
                        txt中医诊断.Text = ""
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cbo预交结算_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long, strInfo As String
    

    With mCurPrepay
        .lng医疗卡类别ID = 0
        .bln消费卡 = False
        .str结算方式 = ""
        .str名称 = ""
    End With
    '130245,切换结算方式，同步更新卡类别ID
    If mbytInState <> 2 Then
        Call SetCardVaribles(True)
    End If
    If mblnNotClick = True Then Exit Sub
    '无支票这种结算性质,所以用名称
    If InStr(cbo预交结算.Text, "支票") > 0 Then
        If Not mrsInfo Is Nothing And IsNumeric(txtPatient.Tag) Then
            strInfo = GetLastInfo(CLng(txtPatient.Tag))
            If strInfo <> "" Then
                txt缴款单位.Text = IIf(Split(strInfo, "|")(0) = "", txt缴款单位.Text, Split(strInfo, "|")(0))
                txt开户行.Text = IIf(Split(strInfo, "|")(1) = "", txt开户行.Text, Split(strInfo, "|")(1))
                txt帐号.Text = IIf(Split(strInfo, "|")(2) = "", txt帐号.Text, Split(strInfo, "|")(2))
            End If
        End If
    Else
        txt缴款单位.Text = ""
        txt开户行.Text = ""
        txt帐号.Text = ""
    End If
    
    If is个人帐户(cbo预交结算) Then
        txt缴款单位.BackColor = Me.BackColor
        txt开户行.BackColor = Me.BackColor
        txt帐号.BackColor = Me.BackColor
        
        txt缴款单位.Enabled = False
        txt开户行.Enabled = False
        txt帐号.Enabled = False
    Else
        txt缴款单位.BackColor = &H80000005
        txt开户行.BackColor = &H80000005
        txt帐号.BackColor = &H80000005
        
        txt缴款单位.Enabled = True
        txt开户行.Enabled = True
        txt帐号.Enabled = True
    End If
    
    '54979:刘鹏飞,2012-10-22
    If txt缴款单位.Text <> "" And txt缴款单位.Enabled = True Then
        chk单位缴款.Value = 1
        If txt缴款单位.Enabled = False Then Call chk单位缴款_Click
    Else
        chk单位缴款.Value = 0
        If txt缴款单位.Enabled = True Then Call chk单位缴款_Click
    End If
    
    '0=新增,1=修改,2=查看
    If mbytInState = 2 Then Exit Sub
    Call Local结算方式(mCurPrepay.lng医疗卡类别ID, False)
End Sub

Private Function is个人帐户(cbo As Object) As Boolean
    If cbo.ListIndex <> -1 Then
        If cbo.ItemData(cbo.ListIndex) = 3 Then
            is个人帐户 = True
        End If
    End If
End Function

Private Sub cbo床位_Click()
    cbo.SetListWidth cbo床位.hWnd, cbo床位.width * 2.9
    If cbo床位.Text = "不分配床位" Then
        chk陪伴.TabStop = False
    Else
        chk陪伴.TabStop = True
    End If
    If mblnAppoint Then cbo床位.Tag = Trim(Split(Trim(cbo床位.Text), " ")(0))
End Sub

Private Sub LoadDept(ByVal bytType As Byte)
'功能：根据科室加载病区，或根据病区加载科室，最后，加载对应的床位
'参数：bytType=0-根据科室加载病区,1-根据病区加载科室
    Dim lngDept As Long, lngUnit As Long
    Dim strFilter As String, i As Long
    
    If cbo入院病区.ListIndex >= 0 Then lngUnit = cbo入院病区.ItemData(cbo入院病区.ListIndex)
    If cbo入院科室.ListIndex >= 0 Then lngDept = cbo入院科室.ItemData(cbo入院科室.ListIndex)
    
    If gbln先选病区 And bytType = 1 Then
        '根据病区加载科室
        mrsUnitDept.Filter = "病区ID=" & lngUnit
        For i = 1 To mrsUnitDept.RecordCount
            strFilter = strFilter & IIf(strFilter = "", "", " Or ") & "ID=" & mrsUnitDept!科室ID
            mrsUnitDept.MoveNext
        Next
        
        '*********************************************************
        '问题 25682 by lesfeng 2009-10-12 b
        If strFilter = "" Then
            cbo入院科室.Clear
        Else
            mrsDept.Filter = strFilter
            Call CboLoadData(cbo入院科室, mrsDept, True)
        End If
        '问题 25682 by lesfeng 2009-10-12 e
        
        i = cbo.FindIndex(cbo入院科室, lngUnit)
        If i = -1 Then
            i = cbo.FindIndex(cbo入院科室, lngDept)
            If i = -1 Then i = 0
        End If
        cbo.SetIndex cbo入院科室.hWnd, i
        cbo入院科室.TabStop = (cbo入院科室.ListCount > 1)
        '问题27370 by lesfeng 2010-01-26
        cbo入院科室.SelLength = 0
    ElseIf Not gbln先选病区 And bytType = 0 Then
        '根据科室加载病区
        mrsUnitDept.Filter = "科室ID=" & lngDept
        For i = 1 To mrsUnitDept.RecordCount
            strFilter = strFilter & IIf(strFilter = "", "", " Or ") & "ID=" & mrsUnitDept!病区ID
            mrsUnitDept.MoveNext
        Next
        mrsUnit.Filter = strFilter
        
        cbo入院病区.Clear
        cbo入院病区.AddItem "不确定病区"
        cbo入院病区.ItemData(cbo入院病区.NewIndex) = 0
        Call CboLoadData(cbo入院病区, mrsUnit, False)
        
        i = cbo.FindIndex(cbo入院病区, lngUnit)
        If i = -1 Then i = 0
        cbo.SetIndex cbo入院病区.hWnd, i
        cbo入院病区.TabStop = (cbo入院病区.ListCount > 1)
        '问题27370 by lesfeng 2010-01-26
        cbo入院病区.SelLength = 0
    End If
    
    '问题26779 by lesfeng 2009-12-10
    lngUnit = 0
    lngDept = 0
    If cbo入院病区.ListIndex >= 0 Then lngUnit = cbo入院病区.ItemData(cbo入院病区.ListIndex)
    If cbo入院科室.ListIndex >= 0 Then lngDept = cbo入院科室.ItemData(cbo入院科室.ListIndex)
    '加载床位
    If gbln入院入科 And mbytMode <> EMode.E预约登记 And mbytInState = EState.E新增 Then
        Call LoadBed(zlCommFun.GetNeedName(cbo性别.Text), lngDept, lngUnit)
    End If
    
    Call LoadBedInfo(lngDept, lngUnit)
End Sub

Private Sub cbo入院科室_Click()
    Dim strDoctors As String, i As Long, lngDepID As Long
    Dim rsDiagnosis As ADODB.Recordset
    
    If cbo入院科室.ListIndex <> -1 Then
        lngDepID = cbo入院科室.ItemData(cbo入院科室.ListIndex)
        
        '该科室对应的病区,床位
        If mbytInState <> EState.E查阅 Then Call LoadDept(0)
        
        '是否是中医科
        If mbytMode <> 1 Then txt中医诊断.Enabled = (InStr(1, "," & GetDepCharacter(lngDepID) & ",", ",中医科,") > 0)
        txt中医诊断.ToolTipText = "只有当入院科室的性质为中医科时才允许输入中医诊断!"
        
        '是否再入院
        If mbytInState = 0 And Not mrsInfo Is Nothing Then
            chk再入院.Value = IIf(CheckReIN(mrsInfo!病人ID, lngDepID), 1, 0)
        End If
        
        '107823显示病人诊断情况
        If (mbytMode = EMode.E接收预约 Or mbytMode = EMode.E正常登记 And mlng病人ID <> 0 And mbytInState = EState.E新增) And Me.Visible = True Then
            Set rsDiagnosis = GetDiagnosticInfo(mlng病人ID, mlng主页ID, "1,11", "3", lngDepID)
            If Not rsDiagnosis Is Nothing Then
                rsDiagnosis.Filter = "诊断类型=1"
                If Not rsDiagnosis.EOF Then
                    txt门诊诊断.Text = Nvl(rsDiagnosis!诊断描述): txt门诊诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl门诊诊断.Tag = txt门诊诊断.Text
                Else
                    txt门诊诊断.Text = ""
                End If
                If txt中医诊断.Enabled Then
                    rsDiagnosis.Filter = "诊断类型=11"
                    If Not rsDiagnosis.EOF Then
                        txt中医诊断.Text = Nvl(rsDiagnosis!诊断描述): txt中医诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                    End If
                Else
                    txt中医诊断.Text = ""
                End If
            End If
        End If
    Else
        txt中医诊断.Enabled = False
        txt中医诊断.ToolTipText = "只有当入院科室的性质为中医科时才允许输入中医诊断!"
    End If
End Sub

Private Sub cbo性别_Click()
    Dim lngDept As Long, lngUnit As Long
    
    If Not cbo性别.Visible Then Exit Sub
    
    If cbo入院病区.ListIndex >= 0 Then lngUnit = cbo入院病区.ItemData(cbo入院病区.ListIndex)
    If cbo入院科室.ListIndex >= 0 Then lngDept = cbo入院科室.ItemData(cbo入院科室.ListIndex)
    Call LoadBed(zlCommFun.GetNeedName(cbo性别.Text), lngDept, lngUnit)
    Call ReLoadCardFee
End Sub

Private Sub chkUnlimit_Click()
     '不限担保额必须设置担保时间,并且不能是临时担保
     
    dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm"
    If chkUnlimit.Value = 1 And IsNull(dtp担保时限.Value) Then
        dtp担保时限.Value = DateAdd("d", 3, CDate(txt入院时间.Text))
    End If
    
    chk临时担保.Enabled = Not (chkUnlimit.Value = 1)
    txt担保额.Enabled = Not (chkUnlimit.Value = 1)
    If chkUnlimit.Value = 1 Then
        txt担保额.Text = "999999999":  txt担保额.BackColor = vbInactiveCaptionText
    Else
        txt担保额.Text = "": txt担保额.BackColor = vbWhite
    End If
End Sub


Private Sub chk记帐_Click()
    If chk记帐.Value = Checked Then
        cbo发卡结算.Enabled = False
        If Visible Then cmdOK.SetFocus
    Else
        cbo发卡结算.Enabled = True
        cbo发卡结算.SetFocus
    End If
End Sub

Private Sub chk临时担保_Click()
    If chk临时担保.Value = 1 Then
        '限时或不限额,不适用于临时担保
        dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm"
        dtp担保时限.Value = Null
        chkUnlimit.Value = 0        '值改变时有隐式调用click事件
    End If
    chkUnlimit.Enabled = Not (chk临时担保.Value = 1)
    dtp担保时限.Enabled = Not (chk临时担保.Value = 1)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Function Get病种名(lng病种ID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select 名称 From 保险病种 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病种ID)
    
    If Not rsTmp.EOF Then Get病种名 = rsTmp!名称
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckReIN(ByVal lng病人ID As Long, ByVal lng科室ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    strSQL = "Select 病人id" & vbNewLine & _
            " From 病案主页 a" & vbNewLine & _
            " Where 病人id = [1] And Nvl(a.主页id, 0) <> 0 And Exists" & vbNewLine & _
            "       (Select 1" & vbNewLine & _
            "            From 临床部门 b" & vbNewLine & _
            "            Where b.部门id = a.出院科室id And b.工作性质 = (Select 工作性质 From 临床部门 Where 部门id = [2] And Rownum = 1))"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng科室ID)
    CheckReIN = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdYB_Click()
    Dim lng病人ID As Long, lng病种ID As Long
    Dim objCurrent As Object, strTxt As String, arrTxt As Variant
    Dim i As Long, blnDo As Boolean, arrPati As Variant
    Dim objcbo As ComboBox
    
    If (mbytMode = EMode.E接收预约 Or mbytMode = EMode.E正常登记 And mlng病人ID <> 0) Then
        lng病人ID = mlng病人ID
    ElseIf Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If MsgBox("当前已经输入一个病人,是否要以该病人的身份进行验证？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                lng病人ID = mrsInfo!病人ID
            End If
        End If
    End If
    
    '医保改动
    mstrYBPati = gclsInsure.Identify(1, lng病人ID, mintInsure)
    mstrYBPatiBak = mstrYBPati '对读到的医保信息进行备份，以便门诊预约病人合并后恢复
    mintInsureBak = mintInsure
    If mstrYBPati <> "" Then
        arrPati = Split(mstrYBPati, ";")
        '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID,...
        If UBound(arrPati) >= 8 Then
            If Val(arrPati(8)) > 0 Then
                txtPatient.Text = "-" & Val(arrPati(8))
                blnDo = txtPatient.Locked
                txtPatient.Locked = False
                Call txtPatient_KeyPress(13)
                txtPatient.Locked = blnDo
                If mstrYBPati = "" Then txt姓名.SetFocus: Exit Sub  '可能因为余额不足提醒选择了退出等,调用了clearcard
            ElseIf mrsInfo Is Nothing Then
                If txtPatient.Tag = "" Then '如果尚未产生
                    txtPatient.Text = zlDatabase.GetNextNo(1) '新病人ID
                    txtPatient.Tag = txtPatient.Text
                    If txt住院号.Visible And mbytKind = EKind.E住院入院登记 Then
                        txt住院号.Text = zlDatabase.GetNextNo(2)
                    ElseIf txt住院号.Visible And mbytKind = EKind.E住院留观登记 Then
                        txt住院号.Text = zlDatabase.GetNextNo(6)
                    End If
                End If
            End If
        End If
        
        txt医保号.Text = arrPati(1)
        txt医保号.Locked = True
        
        txt姓名.Text = arrPati(3)
        cbo性别.ListIndex = GetCboIndex(cbo性别, CStr(arrPati(4)))
        If IsDate(arrPati(5)) Then
            txt出生日期.Text = Format(arrPati(5), "yyyy-MM-dd")
            Call txt出生日期_LostFocus
        End If
        txt身份证号.Text = arrPati(6)
        txt工作单位.Text = arrPati(7)
       
        '保险病种作为入院诊断
        If UBound(arrPati) >= 14 Then
            If Val(arrPati(14)) > 0 Then
                lng病种ID = Val(arrPati(14))
                
                If txt门诊诊断.Text = "" And Not RequestCode Then
                    txt门诊诊断.Text = Get病种名(lng病种ID)
                End If
            End If
        End If
        
        '获取个人帐户余额
        mcurYBMoney = gclsInsure.SelfBalance(Val(arrPati(8)), CStr(arrPati(1)), 20, , mintInsure)
        mcurYBMoneyBak = mcurYBMoney
        lblYBMoney.Caption = "个人帐户余额：" & Format(mcurYBMoney, "0.00")
        lblYBMoney.Visible = True
        
        '医疗付款方式缺省=社会基本医疗保险
        For i = 0 To cbo医疗付款.ListCount
            If InStr(cbo医疗付款.List(i), Chr(&HD)) > 0 Then cbo医疗付款.ListIndex = i: Exit For
        Next
        
        If Not IsDate(txt出生日期.Text) Then
            txt出生日期.SetFocus
        Else
            strTxt = "txt年龄,cbo性别,cbo费别,cbo国藉,cbo民族,cbo学历,cbo婚姻状况,cbo职业,cbo身份," & _
                     "txt身份证号,txt出生地点,txt家庭地址,txt家庭地址邮编,txt家庭电话,txt户口地址,txt户口地址邮编,txt工作单位,txt单位电话,txt单位邮编," & _
                     "txt单位开户行,txt单位帐号,txt联系人姓名,cbo联系人关系,txt联系人地址,txt联系人电话,txt联系人身份证号,txt担保人,txt担保额"
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
        If CanFocus(cbo入院科室) Then cbo入院科室.SetFocus
    Else
        txt姓名.SetFocus
    End If
End Sub

Private Sub SetChargeTurn()
    Dim dat入院时间 As Date
    
    '门诊费用转住院检查
    dat入院时间 = CDate(txt入院时间.Text)
    If frmChargeTurn.CheckExistTurn(Val(txtPatient.Text), dat入院时间) Then
        MsgBox "该病人已存在门诊转住院的单据!" & vbCrLf & _
                "入院时间将被固定为这些单据的最大发生时间。", vbInformation, Me.Caption
        txt入院时间.Text = Format(dat入院时间, "yyyy-MM-dd HH:mm")
        txt入院时间.Enabled = False
    End If
    '问题:33635
    If mstrYBPati <> "" Then
        cmdTurn.Visible = True
    Else
        cmdTurn.Visible = InStr(1, mstrPrivs, ";门诊费用转住院;") > 0 And mbytKind = E住院入院登记 And mbytMode <> 1
    End If
End Sub

Private Sub cmd单位地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From  合约单位" & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID", _
            2, "单位", , txt工作单位.Text)
    If Not rsTmp Is Nothing Then
        txt工作单位.Tag = rsTmp!ID
        txt工作单位.Text = rsTmp!名称
        txt工作单位.SelStart = Len(txt工作单位.Text)
        txt单位电话.Text = Trim(rsTmp!电话 & "")
        txt单位开户行.Text = Trim(rsTmp!开户银行 & "")
        txt单位帐号.Text = Trim(rsTmp!帐号 & "")
        txt工作单位.SetFocus
    End If
End Sub

Private Sub dtp担保时限_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        If dtp担保时限.CheckBox Then
            KeyAscii = 0
            If IsNull(dtp担保时限.Value) Then
                dtp担保时限.Value = DateAdd("d", 3, zlDatabase.Currentdate)
            Else
                dtp担保时限.Value = Null
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim cboTmp As ComboBox, lngIdx As Long
    
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
        If Me.ActiveControl.Name = txt门诊诊断.Name Then
            If InStr(":：;；", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc("'") Then
        If Not (Me.ActiveControl Is txt门诊诊断 Or Me.ActiveControl Is txt中医诊断) Then KeyAscii = 0      '诊断内容中可能有'号
    ElseIf KeyAscii >= 32 And TypeName(Me.ActiveControl) = "ComboBox" Then
        Set cboTmp = Me.ActiveControl
        If cboTmp.Style = 2 Then   '目前cbo门诊医师除外
            lngIdx = cbo.MatchIndex(cboTmp.hWnd, KeyAscii, 0.8)
            If lngIdx = -1 And cboTmp.ListCount > 0 Then lngIdx = 0
            cboTmp.ListIndex = lngIdx
        End If
    End If
    
    '联系人关系说明或转入不允许录入逗号和冒号,因为 该对象（mstrPatiPlus） 的分隔符 包含冒号和逗号
    If Me.ActiveControl Is txt转入 Or Me.ActiveControl Is txtLinkManInfo Then
        If InStr(":：,，", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub Form_Load()
    mblnChange = True
'    Call zlCardSquareObject
    
    With mPageHeight
        .基本 = Me.Height
        .健康档案 = Me.Height
    End With
    Call CreateObjectKeyboard
    Call CreatePublicExpenseObject(mlngModul)
    mstrPrePriceGrade = ""
    '初始化
    If Not gobjSquare Is Nothing Then Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    
    '医保处理
    mintInsure = 0
    mintInsureBak = mintInsure
    mstrYBPati = ""
    mcurYBMoney = 0
    mdat上次担保到期时间 = CDate("0:00:00")
    mstrNOS = ""
    
    '问题26779 by lesfeng 2009-12-10
    lblBedInfo.Caption = ""
        
    mblnUnload = False
    gblnOK = False
        
    '问题27356 by lesfeng 2010-01-13
    If InStr(mstrPrivs, "绑定卡号") = 0 Then
        tabCardMode.Tabs.Remove ("CardBind")
'        tabCardMode.Tabs("CardBind").Selected = True
'        tabCardMode.Tabs("CardBind").Caption = "绑定卡号"
        tabCardMode.width = tabCardMode.width / 2
    End If
    If mbytMode = 2 Then mblnUnload = Not isValid(mlng病人ID)
    Call InitDicts
    If Not InitData Then mblnUnload = True
    If mblnUnload Then Unload Me: Exit Sub
    '问题号:56599
    Call InitFace
    Call InitTabPage
    '问题27370 by lesfeng 2010-01-26
    cbo入院科室.SelLength = 0
    cbo入院病区.SelLength = 0
    If mblnUnload Then Unload Me: Exit Sub

    mFactProperty = zl_GetInvoicePreperty(mlngModul, 2, 2)
    
    If gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
    End If
    '创建写卡对象
    Call zlCreateSquare
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, P病人入院管理, mstrPrivs)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
End Sub

Private Sub Init担保信息(dat入院时间 As Date)
    
    txt担保人.Text = ""
    If mbytInState <> 2 Then chkUnlimit.Enabled = True
    chkUnlimit.Value = 0     '如果值有变化,则隐式调用click事件
    txt担保额.Text = ""
    
    If mbytInState <> 2 Then dtp担保时限.Enabled = True
    dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm"    '设置checkbox可见性
    
    If mbytInState = 0 And mbytMode <> EMode.E预约登记 Then
        '新增时,到期时间不能小于入院时间(修改时在读取卡片和修改入院时间时设置限制)
        dtp担保时限.MinDate = dat入院时间
        dtp担保时限.Value = DateAdd("d", 3, dat入院时间)
    End If
    dtp担保时限.Value = Null
    
    If mbytInState <> 2 Then chk临时担保.Enabled = True
    chk临时担保.Value = 0
    txtReason.Text = ""
End Sub

Private Sub InitFace()
    Dim blnHaveCard As Boolean, dat入院时间 As Date
    Dim lngTmp As Long, bln预交 As Boolean, bln磁卡 As Boolean
    Dim str住院号 As String
    
    Call InitvsDrug
    Call InitVsInoculate
    Call InitVsOtherInfo
    Call InitCertificate
    Call InitCombox
    '重置结构化地址界面
    Call InitStructAddress

    If mbytInState <> E查阅 Then
        txt姓名.MaxLength = GetColumnLength("病人信息", "姓名")
        txt年龄.MaxLength = GetColumnLength("病人信息", "年龄")
        txt住院号.MaxLength = GetColumnLength("病人信息", "住院号")
    End If
    
    '窗体标题
    If mbytMode = E修改 Then
        If mbytKind = E住院入院登记 Then
            Caption = "预约入院登记"
        ElseIf mbytKind = E门诊留观登记 Then
            Caption = "预约门诊留观"
        ElseIf mbytKind = E住院入院登记 Then
            Caption = "预约住院留观"
        End If
    ElseIf mbytMode = 2 Then
        If mbytKind = E住院入院登记 Then
            Caption = "接收住院病人"
        ElseIf mbytKind = E门诊留观登记 Then
            Caption = "接收门诊留观"
        ElseIf mbytKind = E住院留观登记 Then
            Caption = "接收住院留观"
        End If
    Else
        If mbytKind = E住院入院登记 Then
            Caption = "病人入院登记"
        ElseIf mbytKind = E门诊留观登记 Then
            Caption = "门诊留观登记"
        ElseIf mbytKind = E住院留观登记 Then
            Caption = "住院留观登记"
        End If
    End If
    Me.Tag = Me.Caption
    mbytKindBak = mbytKind
    
    Call InitInputTabStop
    
    If InStr(mstrPrivs, "合约病人登记") = 0 Then
        txt工作单位.Enabled = False
        txt工作单位.BackColor = &H8000000F
        txt单位电话.Enabled = False
        txt单位电话.BackColor = &H8000000F
        txt单位邮编.Enabled = False
        txt单位邮编.BackColor = &H8000000F
        txt单位开户行.Enabled = False
        txt单位开户行.BackColor = &H8000000F
        txt单位帐号.Enabled = False
        txt单位帐号.BackColor = &H8000000F
        cmd工作单位.Visible = False
    End If
    
    
    '医保:1.无连接或权限,2.预约登记,3.门诊留观,4.不是执行登记
    '医大附一院：预约允许医保验卡 所以修改条件去掉  And mbytMode <> 1
    cmdYB.Visible = InStr(mstrPrivs, "保险病人登记") > 0 And mbytKind <> E门诊留观登记 And mbytInState = 0
    cmdTurn.Visible = InStr(1, mstrPrivs, ";门诊费用转住院;") > 0 And mbytKind = E住院入院登记 And mbytMode <> 1
    txtTimes.Visible = mbytMode <> 1 And mbytKind = E住院入院登记 '预约登记时或留关登记时,住院次数为零
    lblTimes.Visible = mbytMode <> 1 And mbytKind = E住院入院登记
    cmdName.Visible = mbytMode = 2
    txtTimes.Enabled = (InStr(1, mstrPrivs, "修改住院次数") > 0 And mbytInState = 0)   '修改时不允许改，因为可能已产生住院一次费用，预交款，就诊卡
        
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
    
        
    '预约登记时不填写的内容
    If mbytMode = 1 Then
        txt医保号.Enabled = False
        txt医保号.BackColor = Me.BackColor
        txt门诊诊断.Enabled = False
        txt门诊诊断.BackColor = Me.BackColor
        txt中医诊断.Enabled = False
        txt中医诊断.BackColor = Me.BackColor
    End If
    
    '住院号
    If mbytKind = E门诊留观登记 Then     '门诊留观
        lbl住院号.Caption = "门诊号"
        cmdSelectNO.Visible = False
        lbl姓名.Left = lbl费别.Left
        txt姓名.Left = cbo费别.Left
        txt姓名.width = cbo费别.width
        cmdName.Left = txt姓名.Left + txt姓名.width - cmdName.width - 20
        
        cmdYB.Visible = False
    ElseIf mbytKind = E住院留观登记 Then     '住院留观
        lbl住院号.Caption = "留观号"
        txt住院号.TabStop = False
        txt住院号.Locked = True
        cmdSelectNO.Visible = False
    End If
    
    If InStr(mstrPrivs, "修改住院号") = 0 Then
        txt住院号.Locked = True
        txt住院号.TabStop = False
        txt住院号.BackColor = Me.BackColor
        cmdSelectNO.Visible = False
    End If
    If mbytInState = EState.E查阅 Then cmdSelectNO.Visible = False
    
    If InStr(mstrPrivs, "修改入院日期") = 0 Then
        txt入院时间.Enabled = False
    End If
        
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    mblnChange = False: cbo年龄单位.ListIndex = 0: cbo年龄单位.Tag = txt年龄.Text & "_" & cbo年龄单位.Text: mblnChange = True
    
    '新增,入院登记或入院预约接收
    If mbytInState = 0 Then dat入院时间 = zlDatabase.Currentdate           '新增时,担保到期时间不能小于入院时间
        
    '担保信息
    If mbytInState = 2 Or (mbytMode <> 1 And InStr(mstrPrivs, "担保信息") > 0 And gbln担保) Then
        Call Init担保信息(dat入院时间)
    End If
    '51167,刘鹏飞,2012-07-09,增加"联系人身份证号"
    
    '预约登记不支持登记担保信息(因为没有主页ID)
    If mbytMode = 1 Or mbytInState <> 2 And InStr(mstrPrivs, "担保信息") = 0 Then
        pic担保.Visible = False
        fra病人.Height = fra病人.Height - pic担保.Height
        pic病人.Height = pic病人.Height - pic担保.Height
        Me.Height = Me.Height - pic担保.Height
    Else
        If mbytInState <> 2 And Not gbln担保 Then
            txt担保人.Enabled = False:        txt担保人.BackColor = Me.BackColor
            txt担保额.Enabled = False:        txt担保额.BackColor = Me.BackColor
            txtReason.Enabled = False:        txtReason.BackColor = Me.BackColor
            chkUnlimit.Enabled = False:       chk临时担保.Enabled = False
            lbl担保时限.Enabled = False:      dtp担保时限.Enabled = False
        End If
    End If
    
    If InStr(mstrPrivs, "担保信息") = 0 Then cmdSurety.Visible = False

    '病区与科室
    If gbln先选病区 Then
        lngTmp = lbl入院科室.Left
        lbl入院科室.Left = lbl入院病区.Left
        lbl入院病区.Left = lngTmp
        
        lngTmp = cbo入院科室.Left
        cbo入院科室.Left = cbo入院病区.Left
        cbo入院病区.Left = lngTmp
        
        lngTmp = cbo入院科室.TabIndex
        cbo入院科室.TabIndex = cbo入院病区.TabIndex
        cbo入院病区.TabIndex = lngTmp
    End If
    Call cbo.SetListWidth(cbo入院科室.hWnd, cbo.ListWidth(cbo入院科室.hWnd) * 1.2)
    
    If Not (gbln入院入科 And mbytMode <> EMode.E预约登记) Or mbytInState = EState.E修改 Then
        lbl床位.Visible = False
        cbo床位.Visible = False
        chk陪伴.Visible = False
    End If
    
    Select Case mbytInState         '0=新增,1=修改,2=查阅
        Case E新增
           mFactProperty = zl_GetInvoicePreperty(mlngModul, 2, 2)
            If Not gobjSquare Is Nothing Then
                If Not gobjSquare.objSquareCard Is Nothing Then IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
            End If
        
            If InStr(mstrPrivs, "允许非医保病人") = 0 Then
                txtPatient.TabStop = False
                txt住院号.TabStop = False
            End If
            
            Call InitSendCardPreperty
            chk记帐.Value = IIf(gbln记账 = True, 1, 0)
            
            '获取预交发票号
            Call GetFact(True)
            
            '接收预约登记时，如果有卡了就不可以再发卡
            blnHaveCard = False
            If mbytMode = EMode.E接收预约 Then
                blnHaveCard = PatiHaveCard(mlng病人ID)
            End If
            
            '预约登记时不处理床位分配,不分配病区
            bln预交 = gbln入院预交 And (cbo预交结算.ListCount > 0) And (gblnPrepayStrict And mlng预交领用ID > 0 Or Not gblnPrepayStrict)
            '76824，李南春，2014/8/19，医疗卡发卡处理
            bln磁卡 = (gbln入院发卡 And (mCurSendCard.bln严格控制 And mCurSendCard.lng领用ID > 0 Or Not mCurSendCard.bln严格控制) And Not blnHaveCard _
                    Or mCurSendCard.blnOneCard And mCurSendCard.bln严格控制) And mCurSendCard.lng卡类别ID <> 0
            
            Call HideCard(bln预交, bln磁卡)
            If mbytMode = EMode.E接收预约 Then
                txtPatient.Locked = True
                txtPatient.TabStop = False

                '显示接收信息
                If Not ReadPatiReg(mlng病人ID, mlng主页ID) Then
                    MsgBox "不能正确读取该病人的登记记录！", vbInformation, gstrSysName
                    mblnUnload = True: Exit Sub
                End If
                
                '50511,刘鹏飞,2013-11-04,只有具有调整门诊医师权限才能修改门诊医师
                If InStr(mstrPrivs, ";调整门诊医师;") = 0 And cbo门诊医师.ListIndex <> -1 Then
                    cbo门诊医师.Enabled = False
                End If
                
                '如果之前没有住院号或每次住院产生新住院号,接收为住院病人，则自动分配新的住院号
                '问题 27063 by lesfeng 2009-12-25 预约登记转住院病人保留原住院号(取消gbln每次住院新住院号判断)
'                If mbytKind = EKind.E住院入院登记 And (Trim(txt住院号.Text) = "" Or gbln每次住院新住院号) Then txt住院号.Text = zlDatabase.GetNextNo(2)
                '85510:LPF,2015-06-19,预约登记住院号产生规则（医嘱登记入院处理,因医嘱登记插入住院号时不可能重写住院业务规则）:
                '原有逻辑判断:If mbytKind = EKind.E住院入院登记 And (Trim(txt住院号.Text) = "") Then txt住院号.Text = zlDatabase.GetNextNo(2)
                '入院管理，预约登记会根据参数"每次住院新住院号"生成住院号,而医嘱登记目前只是以病人信息的住院号为准插入(这种方式产生的住院号可能就不正确)
                '因此需要做如下处理
                '1:gbln每次住院新住院号=TRUE,如果存在住院号，则检查已有的住院号是否重复，如果重复则重新生成。
                '2:gbln每次住院新住院号=FALSE,如果住院号为空,则使用历史住院号(最后一次住院号不为空)，不存在历史住院则重新生成。
                If mbytKind = EKind.E住院入院登记 Then
                    If gbln每次住院新住院号 = True Then
                        If Trim(txt住院号.Text) <> "" Then
                            If CheckByPatiNO(mlng病人ID, mlng主页ID, 0, Trim(txt住院号.Text)) = True Then txt住院号.Text = ""
                        End If
                    Else
                        If Trim(txt住院号.Text) = "" Then
                            str住院号 = ""
                            If CheckByPatiNO(mlng病人ID, mlng主页ID, 1, str住院号) = True Then txt住院号.Text = str住院号
                        End If
                    End If
                    If Trim(txt住院号.Text) = "" Then txt住院号.Text = zlDatabase.GetNextNo(2)
                ElseIf mbytKind = E住院留观登记 Then
                    If Trim(txt住院号.Text) = "" Then txt住院号.Text = zlDatabase.GetNextNo(6)
                End If
            Else
                txt入院时间.Text = Format(dat入院时间, "yyyy-MM-dd HH:mm")
            End If
            '89980病人结构化 新增病人设置缺省值
            If gbln启用结构化地址 Then
                Call LoadStructAddressDef(marrAddress)
                Call SetStrutAddress(2)
            End If

        Case E修改    '修改
            If Not gobjSquare Is Nothing Then
                If Not gobjSquare.objSquareCard Is Nothing Then IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
            End If
            '可以再次修改病人信息
            txtPatient.Locked = True
            txtPatient.TabStop = False
            
            '在管理清单界面已限制,已入科病人不准修改入住信息(换床了,转科了)
            Call HideCard(False, False)
            
            '65781:刘鹏飞,2013-11-15,如果产生了医嘱则不允许修改姓名、性别、年龄、出生日期
            If HavedDirections(mlng病人ID, mlng主页ID) Then
                mblnHaveAdvice = True
                txt姓名.Locked = True
                txt姓名.BackColor = &H80000016
                txt出生日期.Enabled = False
                txt出生日期.BackColor = txt姓名.BackColor
                txt出生时间.Enabled = False
                txt出生时间.BackColor = txt姓名.BackColor
                txt年龄.Locked = True
                txt年龄.BackColor = txt姓名.BackColor
                cbo年龄单位.Locked = True
                cbo年龄单位.BackColor = txt姓名.BackColor
                cbo性别.Locked = True
                cbo性别.BackColor = txt姓名.BackColor
            Else
                mblnHaveAdvice = False
            End If
            
            '发生费用不能修改病人姓名
            If HavedInCost(mlng病人ID, mlng主页ID) Then
                txt姓名.Locked = True
            End If
            
            If Not ReadPatiReg(mlng病人ID, mlng主页ID) Then
                MsgBox "不能正确读取该病人的登记记录！", vbInformation, gstrSysName
                mblnUnload = True: Exit Sub
            End If
            
            '50511,刘鹏飞,2013-11-04,只有具有调整门诊医师权限才能修改门诊医师
            If InStr(mstrPrivs, ";调整门诊医师;") = 0 And cbo门诊医师.ListIndex <> -1 Then
                cbo门诊医师.Enabled = False
            End If
            '101160
            Call EMPI_LoadPati
            
        Case E查阅   '查阅
            Call HideCard(False, False)
            Call SetStrutAddress
            'pic病人.Enabled = False
            IDKind.Enabled = False
            txtPatient.Locked = True
            txt住院号.Locked = True
            cmdSelectNO.Enabled = False
            txt姓名.Locked = True
            cmdName.Enabled = False
            cmdYB.Enabled = False
            cmdTurn.Enabled = False
            txt医保号.Locked = True
            txt险类.Locked = True
            txt出生日期.Enabled = False
            txt出生时间.Enabled = False
            txt年龄.Locked = True
            cbo年龄单位.Locked = True
            cbo性别.Locked = True
            cbo费别.Locked = True
            cbo医疗付款.Locked = True
            txt身份证号.Locked = True
            cbo国籍.Locked = True
            cbo民族.Locked = True
            cbo身份.Locked = True
            txt其他证件.Locked = True
            cbo学历.Locked = True
            cbo婚姻状况.Locked = True
            cbo职业.Locked = True
            txt家庭地址.Locked = True
            txt家庭电话.Locked = True
            txt家庭地址邮编.Locked = True
            txt户口地址.Locked = True
            txt户口地址邮编.Locked = True
            txt籍贯.Locked = True
            txt出生地点.Locked = True
            txt区域.Locked = True
            cmd区域.Enabled = False
            cbo病人类型.Locked = True
            txt工作单位.Locked = True
            txt单位电话.Locked = True
            txt单位邮编.Locked = True
            txt单位开户行.Locked = True
            txt单位帐号.Locked = True
            txt联系人姓名.Locked = True
            txt联系人地址.Locked = True
            txt联系人电话.Locked = True
            cbo联系人关系.Locked = True
            txtLinkManInfo.Locked = True
            cmd转入.Enabled = False
            txt转入.Locked = True
            txt联系人身份证号.Locked = True
            txt担保人.Locked = True
            chkUnlimit.Enabled = False
            txt担保额.Locked = True
            dtp担保时限.Enabled = False
            chk临时担保.Enabled = False
            txtReason.Locked = True
            txtMobile.Locked = True
            pic入院.Enabled = False
            
            cmd户口地址.Visible = False
            cmd籍贯.Visible = False
            cmd工作单位.Visible = False
            cmd出生地点.Visible = False
            cmd家庭地址.Visible = False
            cmd联系人地址.Visible = False
            cbo门诊医师.Enabled = False
            cbo病人类型.Enabled = False
            
            cboBloodType.Locked = True
            cboBH.Locked = True
            txtMedicalWarning.Locked = True
            txtOtherWaring.Locked = True
            cmdMedicalWarning.Visible = False
            cboIDNumber.Locked = True
            
            cmdOK.Visible = False
            cmdCancel.Caption = "退出(&X)"
            If Not ReadPatiReg(mlng病人ID, mlng主页ID) Then
                MsgBox "不能正确读取该病人的登记记录！", vbInformation, gstrSysName
                mblnUnload = True: Exit Sub
            End If
            
    End Select
    '预交款收款按键是否有效
    If InStr(GetPrivFunc(glngSys, 1103), "预交收款") = 0 And InStr(GetPrivFunc(glngSys, 1103), "代收款收取") = 0 Then
        cmdDeposit.Visible = False
        '88434 1）新增时才判断预交卡片是否有效,修改和查阅时缺省不可见。2）如果前面新增分支已经设置预交不可见,无需重复设置
        If mbytInState = 0 And mbln是否显示预交 Then
            Call HideCard(False)
        End If
    End If
    
    If InStr(mstrPrivs, "调整病人类型") = 0 Then
        cbo病人类型.Enabled = False
    End If

    Call SetCenter(Me)
    mPageHeight.基本 = Me.Height
End Sub

Private Function PatiHaveCard(ByVal lng病人ID As Long) As Boolean
'功能：判断指定病人是否有就诊卡
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 就诊卡号 From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    If Not rsTmp.EOF Then
        PatiHaveCard = Not IsNull(rsTmp!就诊卡号)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub HideCard(Optional bln预交 As Boolean = True, Optional bln磁卡 As Boolean = True)
    If Not bln预交 Then
        mbln是否显示预交 = False
        pic预交.Visible = False
        Me.Height = Me.Height - pic预交.Height
    Else
        mbln是否显示预交 = True
    End If
    If Not bln磁卡 Then
        mbln是否显示磁卡 = False
        pic磁卡.Visible = False
        Me.Height = Me.Height - pic磁卡.Height
    Else
        mbln是否显示磁卡 = True
    End If
End Sub

Private Sub InitDicts()
    Dim i As Integer
    
    mstr缺省费别 = zlDatabase.GetPara("缺省费别", glngSys, mlngModul, , InStr(1, mstrPrivs, ";参数设置;") > 0)
    Call ReadDict("性别", cbo性别)
    Call ReadDict("费别", cbo费别)
    Call ReadDict("国籍", cbo国籍)
    Call ReadDict("民族", cbo民族)
    Call ReadDict("学历", cbo学历)
    Call ReadDict("婚姻状况", cbo婚姻状况)
    Call ReadDict("职业", cbo职业)
    Call ReadDict("身份", cbo身份)
    Call ReadDict("社会关系", cbo联系人关系)
    
    Call ReadDict("病情", cbo入院病况)
    Call ReadDict("入院方式", cbo入院方式)
    Call ReadDict("入院属性", cbo入院属性)  '刘兴宏:2007/09/13
    Call ReadDict("住院目的", cbo住院目的)
     Call ReadDict("身份证未录原因", cboIDNumber)
   
    Call ReadDict("医疗付款方式", cbo医疗付款, "医疗付款方式")
    
    Call ReadDict("病人类型", cbo病人类型, "病人类型")
    If mbytInState = 0 Then
        Call Load支付方式
    End If
End Sub

Private Function ReadDict(strDict As String, cboInput As ComboBox, Optional strClass As String) As Boolean
'功能：初始化指定词典
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngMaxW As Long
    Dim strTemp As String
    Dim str缺省费别 As String, blnFee As Boolean
    
    On Error GoTo errH
    str缺省费别 = mstr缺省费别
    
    'by lesfeng 2010-01-12 性能优化
    If strDict = "结算方式" Then
        If strClass = "就诊卡" Then
            strTemp = "1,2"
        ElseIf strClass = "预交款" Then
            If mbytMode = 1 Then
                strTemp = "1,2,8" '预约登记时
            Else
                If InStr(mstrPrivs, "保险病人登记") > 0 Then
                    strTemp = "1,2,3,5,8"
                Else
                    strTemp = "1,2,5,8"
                End If
            End If
        End If
'        strSQL = "Select Nvl(A.缺省标志,0) as 缺省,B.编码,B.名称,Nvl(B.性质,1) as 性质" & _
'            " From 结算方式应用 A,结算方式 B" & _
'            " Where A.结算方式=B.名称 And A.应用场合='" & strClass & "'" & _
'            " And Nvl(B.性质,1) IN(" & strTemp & ") Order by B.编码"
        strSQL = "Select Nvl(A.缺省标志,0) as 缺省,B.编码,B.名称,Nvl(B.性质,1) as 性质" & _
            " From 结算方式应用 A,结算方式 B,Table(Cast(f_Num2List([1]) As zlTools.t_Numlist)) C " & _
            " Where A.结算方式=B.名称 And A.应用场合=[2]" & _
            " And (B.性质 = C.Column_Value or B.性质 is null) Order by B.编码"
    ElseIf strDict = "身份" Then
        strSQL = "Select 编码,名称,简码,Nvl(优先级,0) as 缺省 From " & strDict & " Order by 编码"
    ElseIf strDict = "病人类型" Then
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省,颜色 From 病人类型 Order by 编码"
    ElseIf strDict = "医疗付款方式" Then
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省,是否医保 From 医疗付款方式 Order by 编码"
    ElseIf strDict = "费别" Then
        '不是仅限初诊身份唯一性项目(包含了缺省费别),不管有效期间及科室
        If mbytKind = E门诊留观登记 Then
            strTemp = "1,3" '门诊留观登记
        Else
            strTemp = "2,3" '住院入院或住院留观登记
        End If
        strSQL = "Select A.编码,A.名称,A.简码,Nvl(A.缺省标志,0) as 缺省 From 费别 A,Table(Cast(f_Num2List([1]) As zlTools.t_Numlist)) B " & _
                 " Where (A.服务对象 = B.Column_Value or A.服务对象 is null) And A.属性=1 And Nvl(A.仅限初诊,0)=0 And  " & _
                 " (a.有效开始 Is Null And a.有效结束 Is Null Or Trunc(Sysdate) Between a.有效开始 And a.有效结束) Order by A.编码"
                 
'        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别" & _
'            " Where 属性=1 And Nvl(仅限初诊,0)=0 And Nvl(服务对象,3) IN(" & strTemp & ")" & _
'                " And  Sysdate Between NVL(有效开始,Sysdate-1) and NVL(有效结束,Sysdate+1)" & _
'            " Order by 编码"
    Else
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTemp, strClass)
    cboInput.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If strClass = "医疗付款方式" Then
                cboInput.AddItem rsTmp!编码 & "-" & rsTmp!名称 & IIf(strClass = "医疗付款方式" And Val(Nvl(rsTmp!是否医保)) = 1, Chr(&HD), "")
            ElseIf strDict = "职业" Then
                cboInput.AddItem rsTmp!编码 & "-" & Chr(&HA) & rsTmp!名称
            Else
                cboInput.AddItem rsTmp!编码 & "-" & rsTmp!名称
            End If
            
            If rsTmp!缺省 = 1 Then
                cboInput.ListIndex = cboInput.NewIndex
                cboInput.ItemData(cboInput.NewIndex) = 1
            End If
            If strDict = "费别" And str缺省费别 = "" & rsTmp!名称 Then
                str缺省费别 = rsTmp!编码 & "-" & rsTmp!名称
                blnFee = True
            End If
            
            Select Case strClass
                Case "预交款"
                    cboInput.ItemData(cboInput.NewIndex) = rsTmp!性质
            End Select
            If TextWidth(cboInput.List(cboInput.NewIndex) & "字") > lngMaxW Then lngMaxW = TextWidth(cboInput.List(cboInput.NewIndex) & "字")
            rsTmp.MoveNext
        Next
        '69489
        If strDict = "费别" And blnFee = True Then
            For i = 0 To cboInput.ListCount - 1
                cboInput.ItemData(i) = 0
                If str缺省费别 = cboInput.List(i) Then
                    cboInput.ListIndex = i
                End If
            Next i
            If cboInput.ListIndex > 0 Then cboInput.ItemData(cboInput.ListIndex) = 1
        End If
    ElseIf strDict = "结算方式" Then
        If strClass = "预交款" Then
            MsgBox "没有设置预交款结算方式,病人入院时不能缴预交款！" & vbCrLf & _
                "要使用入院预交,请先到结算方式管理中设置。", vbInformation, gstrSysName
        Else
            MsgBox "没有设置就诊卡结算方式,病人入院时只能记帐发卡！" & vbCrLf & _
                "要使用结算发卡,请先到结算方式管理中设置。", vbInformation, gstrSysName
            chk记帐.Value = 1: chk记帐.Enabled = False: cbo发卡结算.Enabled = False
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
    
    '82401:李南春,2015/3/11,检查对象是否存在
    If mbytInState = 0 And pic磁卡.Visible Then
        zlDatabase.SetPara "发卡模式", tabCardMode.SelectedItem.Key, glngSys, mlngModul
    End If
    
    Call zlCommFun.OpenIme
    mbytMode = 0
    mbytInState = 0
    mbytKind = 0
    mlng病人ID = 0
    mlng主页ID = 0
    mlng预交领用ID = 0
    Set mrsInfo = Nothing
    Set mrsDoctor = Nothing
    
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
    
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
    
    If Not mdic医疗卡属性 Is Nothing Then
        Set mdic医疗卡属性 = Nothing
    End If
    
    '卸载消息对象
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
    '问题号:56599
    mbln发卡或绑定卡 = False
    Set mrsInputSet = Nothing
End Sub


Private Function InitData() As Boolean
'功能：初始化护理等级、入院科室、入院病区、门诊医师等信息
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim strDeptIDs As String
    
    If cbo费别.ListCount = 0 Then
        MsgBox "没有设置费别信息,请先到费别等级设置中设置！", vbExclamation, gstrSysName
        Exit Function
    ElseIf cbo费别.ListIndex = -1 Then
        cbo费别.ListIndex = 0
    End If
    
    '护理等级(缺省第一个或基本护理)
    Set rsTmp = GetNurseGrade
    If rsTmp.RecordCount > 0 Then
        cbo护理等级.Clear
        cbo护理等级.AddItem ""   '第一个为空,ReadPatiReg中有依赖
        cbo护理等级.ItemData(cbo护理等级.NewIndex) = 0
        
        Call CboLoadData(cbo护理等级, rsTmp, False)
        If cbo护理等级.ListIndex = -1 Then cbo护理等级.ListIndex = 0
    Else
        MsgBox "没有设置护理等级，请先到护理等级设置中初始！", vbInformation, gstrSysName
        Exit Function
    End If
    
       
    '读取并加载门诊医师列表
    Set mrsDoctor = GetDoctorOrNurse(0)
    For i = 1 To mrsDoctor.RecordCount
        cbo门诊医师.AddItem mrsDoctor!简码 & "-" & mrsDoctor!姓名
        cbo门诊医师.ItemData(cbo门诊医师.NewIndex) = mrsDoctor!ID
        mrsDoctor.MoveNext
    Next
    
    '门诊观察室的床位应该没有固定科室,但现在暂时这样处理,同样以科室定床位及病区
    '94400
    If mbytMode = EMode.E预约登记 And InStr(mstrPrivs, ";全院预约;") = 0 Then
        strDeptIDs = GetDeptOrUnitByUser()
    End If
    Set mrsDept = GetDepartments("临床", IIf(mbytKind = EKind.E门诊留观登记, "1", "2") & ",3", , True, strDeptIDs)
    If mrsDept.RecordCount = 0 Then
        MsgBox "没有设置服务于" & IIf(mbytKind = EKind.E门诊留观登记, "门诊", "住院") & "的科室的床位！", vbInformation, gstrSysName
        Exit Function
    End If
    Set mrsUnit = GetDepartments("护理", IIf(mbytKind = EKind.E门诊留观登记, "1", "2") & ",3", , True, strDeptIDs)
    If mrsUnit.RecordCount = 0 Then
        MsgBox "没有设置服务于" & IIf(mbytKind = EKind.E门诊留观登记, "门诊", "住院") & "的病区的床位！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '读取病区科室对应
    Set mrsUnitDept = GetUnitDept
    If mrsUnitDept.RecordCount = 0 Then
        MsgBox "没有设置病区科室对应关系,请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If
        
    If gbln先选病区 Then
        Call CboLoadData(cbo入院病区, mrsUnit, True)
        If cbo入院病区.ListCount > 0 Then cbo入院病区.ListIndex = 0 '调用Click事件,加载科室、床位内容
    Else
        Call CboLoadData(cbo入院科室, mrsDept, True)
        If cbo入院科室.ListCount > 0 Then cbo入院科室.ListIndex = 0 '调用Click事件,加载病区、床位内容
    End If
    
    Call GetRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
    mintIDKind = Val(mintIDKind)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
         
    
    InitData = True
End Function

Private Sub ClearCard(Optional blnKeepOther As Boolean, Optional blnKeepRec As Boolean)
'功能：清除入院登记卡
'参数：blnKeepOther=是否保留磁卡和预交的信息
'      blnKeepRec=是否保留已读取的病人信息对象
    Dim lngUnit As Long, lngDept As Long
    Dim str缺省缴款方式 As String
    
    If Not blnKeepRec Then
        Set mrsInfo = Nothing
        txtPatient.Text = "": txtPatient.Tag = ""
        txt姓名.Locked = False
        cbo性别.Locked = False
        txt年龄.Locked = False
        cbo年龄单位.Locked = False
    End If
    
    If gblnLED Then
        zl9LedVoice.DisplayPatient ""
    End If
    
    txt住院号.Locked = mbytKind = E住院留观登记
    If mbytInState = EState.E新增 And (mbytMode = EMode.E正常登记 Or mbytMode = EMode.E预约登记) And mlng病人ID <> 0 Then
        If mbytMode = EMode.E正常登记 Then mbytKind = mbytKindBak
        txtPatient.Locked = False: txtPatient.TabStop = Not (InStr(mstrPrivs, "允许非医保病人") = 0)
        '66333:刘鹏飞,2013-10-10,门诊留关登记后lbl住院号.Caption = "门诊号"
        If mbytKind = E门诊留观登记 Then     '门诊留观
'            txt住院号.Locked = True
'            lbl住院号.Visible = False
'            txt住院号.Visible = False
            lbl住院号.Caption = "门诊号"
            cmdSelectNO.Visible = False
            cmdYB.Visible = False
        ElseIf mbytKind = E住院留观登记 Then     '住院留观(住院留观号不能修改，每次新登记时按照留观号规则自动产生)
            txt住院号.TabStop = False
            txt住院号.Locked = True
            cmdSelectNO.Visible = False
            lbl住院号.Caption = "留观号"
        End If
        
        mlng病人ID = 0: mlng主页ID = 0
        Me.Caption = Me.Tag
    End If
    
    mblnEMPI = False
    txt险类.Text = ""
    txt医保号.Text = ""
    txt医保号.Locked = False
    If mbytMode <> EMode.E预约登记 And mbytKind = EKind.E住院入院登记 Then
        txtTimes.Text = "1": txtTimes.Tag = 1
    Else
        txtTimes.Text = "": txtTimes.Tag = ""
    End If
    txtPages.Text = "1"
    
    txt住院号.Text = ""
    txt姓名.Text = ""
    txt年龄.Text = "": Call txt年龄_Validate(False): cbo年龄单位.Tag = txt年龄.Text & "_" & cbo年龄单位.Text
    txt出生日期.Text = "____-__-__"
    txt出生时间.Text = "__:__"
    txt身份证号.Text = ""
    txtMobile.Text = ""
    txt其他证件.Text = ""
    txt出生地点.Text = ""
    txt家庭地址.Text = ""
    txt家庭地址邮编.Text = ""
    txt家庭电话.Text = ""
    txt户口地址.Text = ""
    txt户口地址邮编.Text = ""
    txt籍贯.Text = ""
    txt区域.Text = ""
    txt联系人姓名.Text = ""
    txt联系人地址.Text = ""
    txt联系人电话.Text = ""
    txt联系人身份证号.Text = ""
    txtLinkManInfo.Text = ""
    txt工作单位.Text = "": txt工作单位.Tag = ""
    txt工作单位.Text = ""
    txt单位电话.Text = ""
    txt单位邮编.Text = ""
    txt单位开户行.Text = ""
    txt单位帐号.Text = ""
    txt备注.Text = ""
    '问题号:53408
    txt支付密码.Text = ""
    txt验证密码.Text = ""
    txt验证密码.Tag = ""
    txt支付密码.Enabled = False
    txt验证密码.Enabled = False
    lbl支付密码.Enabled = False
    lbl验证密码.Enabled = False
    
    
    txt门诊诊断.Text = "": txt门诊诊断.Tag = "": lbl门诊诊断.Tag = ""
    txt中医诊断.Text = "": txt中医诊断.Tag = "": lbl中医诊断.Tag = ""
    
    chk二级院转入.Value = 0
    chk陪伴.Value = 0
    
    '73420:刘鹏飞,2014-06-09
    If InStr(mstrPrivs, "修改住院号") = 0 Then
        txt住院号.Locked = True
        txt住院号.TabStop = False
        txt住院号.BackColor = Me.BackColor
        cmdSelectNO.Visible = False
    End If
    
    cboIDNumber.ListIndex = -1 '缺省
    cboIDNumber.Enabled = True
    cbo联系人关系.ListIndex = -1
    
    Call SetCboDefault(cbo性别)
    Call SetCboDefault(cbo费别)
    Call SetCboDefault(cbo国籍)
    Call SetCboDefault(cbo民族)
    Call SetCboDefault(cbo学历)
    Call SetCboDefault(cbo婚姻状况)
    Call SetCboDefault(cbo职业)
    Call SetCboDefault(cbo身份)
    Call SetCboDefault(cbo入院病况)
    Call SetCboDefault(cbo入院方式)
    Call SetCboDefault(cbo入院属性) '刘兴宏:2007/09/13
    Call SetCboDefault(cbo住院目的)
    Call SetCboDefault(cbo医疗付款)
    Call SetCboDefault(cbo病人类型)
    
    str缺省缴款方式 = zlDatabase.GetPara("缺省缴款方式", glngSys, mlngModul)
    '预交结算和发卡结算 缺省值 存放在属性Tag中,而Itemdata为 结算性质,故不用SetCboDefault
    If str缺省缴款方式 = "" Then
        If cbo预交结算.ListCount > 0 Then cbo预交结算.ListIndex = Val(cbo预交结算.Tag)
    Else
        Call zlControl.CboLocate(cbo预交结算, str缺省缴款方式, False)
    End If
    
    '重新取可用床位
    If cbo入院病区.ListIndex >= 0 Then lngUnit = cbo入院病区.ItemData(cbo入院病区.ListIndex)
    If cbo入院科室.ListIndex >= 0 Then lngDept = cbo入院科室.ItemData(cbo入院科室.ListIndex)
    Call LoadBed(zlCommFun.GetNeedName(cbo性别.Text), lngDept, lngUnit)
    
    txt卡号.TabStop = True
    
    '入院信息
    txt入院时间.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If InStr(mstrPrivs, "修改入院日期") > 0 Then txt入院时间.Enabled = True
    
    
    If Not blnKeepOther Then
        '磁卡信息
        txt卡号.Text = ""
        txtPass.Text = ""
        txtAudi.Text = ""
        
        chk记帐.Value = IIf(gbln记账 = True, 1, 0)
        If str缺省缴款方式 = "" Then
            If cbo发卡结算.ListCount > 0 Then cbo发卡结算.ListIndex = Val(cbo发卡结算.Tag)
        Else
            Call zlControl.CboLocate(cbo发卡结算, str缺省缴款方式, False)
        End If
        
        '预交信息
        txt预交额.Text = ""
        txt缴款单位.Text = ""
        txt帐号.Text = ""
        txt开户行.Text = ""
        txt结算号码.Text = ""
    End If
    
    '医保改动
    txt姓名.ForeColor = lblPatiColor.BackColor
    mstrNOS = ""
    mintInsure = 0
    mstrYBPati = ""
    mcurYBMoney = 0
    mintInsureBak = 0
    mstrYBPatiBak = ""
    mcurYBMoneyBak = 0
    lblYBMoney.Caption = "个人帐户余额:"
    lblYBMoney.Visible = False
    chk再入院.Value = 0
    cmdTurn.Visible = InStr(1, mstrPrivs, ";门诊费用转住院;") > 0 And mbytKind = E住院入院登记 And mbytMode <> 1 '33635
    If InStr(mstrPrivs, "担保信息") > 0 And gbln担保 Then Call Init担保信息(CDate(txt入院时间.Text))
    cmdName.Visible = mbytMode = 2
    txtTimes.Visible = mbytMode <> 1 And mbytKind = E住院入院登记 '预约登记时或留关登记时,住院次数为零
    lblTimes.Visible = mbytMode <> 1 And mbytKind = E住院入院登记
    txtTimes.Enabled = (InStr(1, mstrPrivs, "修改住院次数") > 0 And mbytInState = 0)   '修改时不允许改，因为可能已产生住院一次费用，预交款，就诊卡
    
    '问题号:56599
    Call Clear健康档案
    If gbln启用结构化地址 Then
        Call SetStrutAddress(1)
        Call SetStrutAddress(2)
    End If
End Sub

Private Sub cmd出生地点_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt出生地点, True)
    If Not rsTmp Is Nothing Then
        txt出生地点.Text = rsTmp!名称
        txt出生地点.SelStart = Len(txt出生地点.Text)
        txt出生地点.SetFocus
    End If
End Sub

Private Sub cmd工作单位_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetOrgAddress(Me, txt工作单位, True)
    If Not rsTmp Is Nothing Then
        txt工作单位.Tag = rsTmp!ID
        txt工作单位.Text = rsTmp!名称
        txt工作单位.SelStart = Len(txt工作单位.Text)
        txt单位电话.Text = Trim(rsTmp!电话 & "")
        txt单位开户行.Text = Trim(rsTmp!开户银行 & "")
        txt单位帐号.Text = Trim(rsTmp!帐号 & "")
        
        txt工作单位.SetFocus
    End If
End Sub

Private Sub cmd家庭地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt家庭地址, True)
    If Not rsTmp Is Nothing Then
        txt家庭地址.Text = rsTmp!名称
        txt家庭地址.SelStart = Len(txt家庭地址.Text)
        txt家庭地址.SetFocus
    End If
End Sub

Private Sub cmd联系人地址_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetAddress(Me, txt联系人地址, True)
    If Not rsTmp Is Nothing Then
        txt联系人地址.Text = rsTmp!名称
        txt联系人地址.SelStart = Len(txt联系人地址.Text)
        txt联系人地址.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim obj As Control
    
    Select Case KeyCode
        Case vbKeyF3
            If ActiveControl.Name = txt出生地点.Name Then
                cmd出生地点_Click
            ElseIf ActiveControl.Name = txt家庭地址.Name Then
                cmd家庭地址_Click
            ElseIf ActiveControl.Name = txt联系人地址.Name Then
                cmd联系人地址_Click
            ElseIf ActiveControl.Name = txt工作单位.Name Then
                cmd工作单位_Click
            ElseIf ActiveControl.Name = txt区域.Name Then
                cmd区域_Click
            End If
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC卡号")
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
            If obj.Name = "cbo性别" Then
                If cbo性别.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo费别" Then
                If cbo费别.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
            ElseIf obj.Name = "cbo发卡结算" Then
                If cbo发卡结算.ListIndex <> -1 Then cmdOK.SetFocus
            ElseIf InStr(1, ",txt卡号,txt出生地点,txt户口地址,txt家庭地址,txt联系人地址,txt工作单位,txt预交额,txtPatient,txt姓名," & _
                "txt住院号,txt籍贯,txt区域,txt门诊诊断,txt中医诊断,txtPass,txtAudi,txt卡额,vsDrug,vsInoculate,vsLinkMan,vsOtherInfo,vsCertificate,PatiAddress,", "," & obj.Name & ",") <= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                 
        End If
    End Select
End Sub

Private Sub PatiAddress_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True) '打开中文输入法
End Sub

Private Sub PatiAddress_LostFocus(Index As Integer)
'功能:
    Select Case Index
    
    Case E_IX_现住址
        txt家庭地址.Text = PatiAddress(Index).Value
    Case E_IX_出生地点
        txt出生地点.Text = PatiAddress(Index).Value
    Case E_IX_户口地址
        txt户口地址.Text = PatiAddress(Index).Value
    Case E_IX_籍贯
        txt籍贯.Text = PatiAddress(Index).Value
    Case E_IX_联系人地址
        txt联系人地址.Text = PatiAddress(Index).Value
    End Select
    Call zlCommFun.OpenIme '关闭中文输入法
End Sub

Private Sub PatiAddress_SetInput(Index As Integer, ByVal intLevel As Integer, rsReturn As ADODB.Recordset)
    '功能：在输入病人结构化地址的时候,加载邮编
    If (Not rsReturn Is Nothing) And intLevel = 2 Then
        If Index = 3 Then
            txt家庭地址邮编.Text = rsReturn!邮编 & ""
        End If
        If Index = 4 Then
            txt户口地址邮编.Text = rsReturn!邮编 & ""
        End If
    End If
End Sub

Private Sub PatiAddress_Validate(Index As Integer, Cancel As Boolean)
    Dim lngLen As Long
    
    lngLen = PatiAddress(Index).MaxLength
    If LenB(StrConv(PatiAddress(Index).Value, vbFromUnicode)) > lngLen Then
        MsgBox PatiAddress(Index).Tag & "只允许输入 " & lngLen & " 个字符或 " & lngLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub tabCardMode_Click()
    If tabCardMode.SelectedItem.Key = "CardFee" Then
        lbl金额.Visible = True
        txt卡额.Visible = True
        chk记帐.Visible = True
        cbo发卡结算.Visible = True
    Else
        lbl金额.Visible = False
        txt卡额.Visible = False
        chk记帐.Visible = False
        cbo发卡结算.Visible = False
    End If
End Sub


Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    Call OpenPassKeyboard(txtAudi, True)
End Sub

Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mCurSendCard.int密码规则 = 1 Then
            Call zlControl.TxtCheckKeyPress(txtAudi, KeyAscii, m数字式)
        End If
    End If

    If KeyAscii = vbKeyReturn Then
        If txtPass.Text <> txtAudi.Text Then
            MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
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
    Select Case mCurSendCard.int密码长度限制
        Case 0
        Case 1
            If Len(txtAudi.Text) <> mCurSendCard.int密码长度 Then
                MsgBox "注意:" & vbCrLf & "确认密码必须输入" & mCurSendCard.int密码长度 & "位", vbOKOnly + vbInformation
                If txtAudi.Enabled Then txtAudi.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtAudi.Text) < Abs(mCurSendCard.int密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "确密码必须输入" & Abs(mCurSendCard.int密码长度限制) & "位以上.", vbOKOnly + vbInformation
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
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("联系人关系备注") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("联系人关系备注")) = txtLinkManInfo.Text
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
   Select Case mCurSendCard.int密码长度限制
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCurSendCard.int密码长度 Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & mCurSendCard.int密码长度 & "位", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCurSendCard.int密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & Abs(mCurSendCard.int密码长度限制) & "位以上.", vbOKOnly + vbInformation
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
        If mCurSendCard.int密码规则 = 1 Then
            Call zlControl.TxtCheckKeyPress(txtPass, KeyAscii, m数字式)
        End If
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            If chk记帐.Visible And chk记帐.Enabled And txt卡额.Locked Then
                chk记帐.SetFocus
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

Private Sub txt备注_GotFocus()
    Call zlControl.TxtSelAll(txt备注)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt备注_KeyPress(KeyAscii As Integer)
    CheckInputLen txt备注, KeyAscii
End Sub

Private Sub txt备注_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '问题27554 by lesfeng 2010-01-19 lngTXTProc 修改为glngTXTProc
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt出生地点_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt出生日期_Change()
    Dim str出生日期 As String
    
    If IsDate(txt出生日期.Text) And mblnChange Then
        mblnChange = False
        txt出生日期.Text = Format(CDate(txt出生日期.Text), "yyyy-mm-dd") '0002-02-02自动转换为2002-02-02,否则,看到的是2002,实际值却是0002
        mblnChange = True
        If txt出生时间.Text = "__:__" Then
            str出生日期 = Format(txt出生日期.Text, "YYYY-MM-DD")
        Else
            str出生日期 = Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS")
        End If
        txt年龄.Text = ReCalcOld(CDate(str出生日期), cbo年龄单位, , , CDate(txt入院时间.Text))
        cbo年龄单位.Tag = txt年龄.Text & "_" & cbo年龄单位.Text
    End If
End Sub

Private Sub txt出生日期_Validate(Cancel As Boolean)
    If IsDate(txt出生日期.Text) And IsDate(txt入院时间.Text) Then
        If CDate(txt出生日期.Text) > CDate(txt入院时间.Text) Then Call zlControl.TxtSelAll(txt出生日期): Cancel = True
    End If
End Sub

Private Sub txt出生时间_Change()
    Dim str出生日期 As String
    
    If IsDate(txt出生时间.Text) And IsDate(txt出生日期.Text) And mblnChange Then
        str出生日期 = Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS")
        txt年龄.Text = ReCalcOld(CDate(str出生日期), cbo年龄单位, , , CDate(txt入院时间.Text))
        cbo年龄单位.Tag = txt年龄.Text & "_" & cbo年龄单位.Text
    End If
End Sub

Private Sub txt出生时间_GotFocus()
    Call OS.OpenImeByName
    zlControl.TxtSelAll txt出生时间
End Sub

Private Sub txt出生时间_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt出生日期.Text) Then
        KeyAscii = 0
        txt出生时间.Text = "__:__"
    End If
End Sub


Private Sub txt出生时间_Validate(Cancel As Boolean)
    If txt出生时间.Text <> "__:__" And Not IsDate(txt出生时间.Text) Then
        txt出生时间.SetFocus
        Cancel = True
    End If
End Sub


Private Sub txt单位开户行_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt担保额_LostFocus()
    If IsNumeric(txt担保额.Text) Then
        txt担保额.Text = Format(txt担保额.Text, "0.00")
    Else
        txt担保额.Text = ""
    End If
    
    Call zlCommFun.OpenIme
End Sub

Private Sub txt担保人_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt工作单位_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt户口地址_Change()
    If txt户口地址.Text = "" Then txt户口地址.Tag = ""
End Sub

Private Sub txt户口地址_GotFocus()
    zlControl.TxtSelAll txt户口地址
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt户口地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt户口地址.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt户口地址)
            If Not rsTmp Is Nothing Then
                txt户口地址.Text = rsTmp!名称
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt户口地址, KeyAscii
    End If
End Sub

Private Sub txt户口地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt户口地址邮编_GotFocus()
    zlControl.TxtSelAll txt户口地址邮编
End Sub

Private Sub txt户口地址邮编_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If ((Not IsNumeric(txt户口地址邮编.Text)) Or Len(txt户口地址邮编.Text) > 6 Or InStr(txt户口地址邮编.Text, ".") > 0) And txt户口地址邮编.Text <> "" Then
            Call SelectYouBian(txt户口地址邮编)
        End If
    End If
End Sub

Private Sub txt籍贯_GotFocus()
    zlControl.TxtSelAll txt籍贯
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt籍贯_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt籍贯.Text <> "" Then
            Set rsTmp = GetArea(Me, txt籍贯)
            If Not rsTmp Is Nothing Then
                txt籍贯.Text = rsTmp!名称
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                zlControl.TxtSelAll txt籍贯
                txt区域.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt籍贯, KeyAscii
    End If
End Sub

Private Sub txt籍贯_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt家庭地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt缴款单位_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt卡号_Change()
    SetCardEditEnabled
End Sub

Private Sub txt卡号_LostFocus()
    Call SetBrushCardObject(False)
End Sub

Private Sub txt开户行_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt联系人地址_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt联系人电话_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("联系人电话") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("联系人电话")) = txt联系人电话.Text
    End If
End Sub

Private Sub txt联系人身份证号_GotFocus()
    zlControl.TxtSelAll txt联系人身份证号
End Sub

Private Sub txt联系人身份证号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt联系人身份证号_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("联系人身份证号") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("联系人身份证号")) = txt联系人身份证号.Text
    End If
End Sub

Private Sub txt联系人姓名_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt联系人姓名_Validate(Cancel As Boolean)
    If vsLinkMan.Rows > vsLinkMan.FixedRows And vsLinkMan.ColIndex("联系人姓名") >= 0 Then
        vsLinkMan.TextMatrix(vsLinkMan.FixedRows, vsLinkMan.ColIndex("联系人姓名")) = txt联系人姓名.Text
        If vsLinkMan.Rows = vsLinkMan.FixedRows + 1 And txt联系人姓名.Text <> "" Then
            vsLinkMan.Rows = vsLinkMan.Rows + 1
        End If
    End If
End Sub

Private Sub txt门诊诊断_LostFocus()
    If Not RequestCode Then
        Call zlCommFun.OpenIme
    End If
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo年龄单位.Visible = False And IsNumeric(txt年龄.Text) Then
            Call txt年龄_Validate(False)
        End If
        If cbo年龄单位.Visible And Not IsNumeric(txt年龄.Text) And Me.ActiveControl.Name = "txt年龄" Then Call zlCommFun.PressKey(vbKeyTab)  '目的是不经过年龄单位

    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt年龄_LostFocus()
    If cbo年龄单位.Tag <> txt年龄.Text & "_" & cbo年龄单位.Text Then
        cbo年龄单位_LostFocus
    End If
End Sub

Private Sub txt年龄_Validate(Cancel As Boolean)
    If Not IsNumeric(txt年龄.Text) And Trim(txt年龄.Text) <> "" Then
        If Not InStr(Trim(txt年龄.Text), "约") > 0 And Trim(txt年龄.Text) <> "不详" Then
            cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False
            txt出生日期.Enabled = True
            txt出生时间.Enabled = True
        ElseIf InStr(Trim(txt年龄.Text), "约") > 0 Or Trim(txt年龄.Text) = "不详" Then
            If Trim(txt出生日期.Text) = "____-__-__" Then
                txt出生日期.Enabled = False
                txt出生时间.Enabled = False
            End If
            cbo年龄单位.ListIndex = -1: cbo年龄单位.Visible = False
        End If
    ElseIf cbo年龄单位.Visible = False Or txt出生日期.Enabled = False Then
        cbo年龄单位.ListIndex = 0: cbo年龄单位.Visible = True
        txt出生日期.Enabled = True
        txt出生时间.Enabled = True
    Else
        txt出生日期.Enabled = True
        txt出生时间.Enabled = True
    End If
    Call ReLoadCardFee
End Sub

Private Sub txt区域_GotFocus()
    zlControl.TxtSelAll txt区域
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt区域_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt区域.Text <> "" Then
            Set rsTmp = GetArea(Me, txt区域)
            If Not rsTmp Is Nothing Then
                txt区域.Text = rsTmp!名称
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                zlControl.TxtSelAll txt区域
                txt区域.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt区域, KeyAscii
    End If
End Sub

Private Sub txt区域_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt入院时间_LostFocus()
    If Not IsDate(txt入院时间.Text) Then
        txt入院时间.SetFocus
    ElseIf dtp担保时限.Enabled Then
        If mbytInState = 0 And Not IsNull(dtp担保时限.Value) Then
            dtp担保时限.MinDate = CDate("1900-01-01")   '先暂时设置一个小的值,否则赋值会出错
            
            If dtp担保时限.Value < CDate(txt入院时间.Text) Then
                dtp担保时限.Value = DateAdd("d", 3, CDate(txt入院时间.Text))
                MsgBox "当前设置的担保到期时间小于入院时间,已调整为入院时间后3天!", vbInformation, gstrSysName
            End If
            
            '担保时限不能小于入院时间
            dtp担保时限.MinDate = CDate(txt入院时间.Text)
        ElseIf mbytInState = 1 Then
        
            If Not IsNull(dtp担保时限.Value) Then
                dtp担保时限.MinDate = CDate("1900-01-01")   '先暂时设置一个小的值,否则赋值会出错
                '担保时限不能小于入院时间
                If dtp担保时限.Value < CDate(txt入院时间.Text) And txt担保额.Enabled Then
                    dtp担保时限.Value = DateAdd("d", 3, CDate(txt入院时间.Text))
                    MsgBox "当前设置的担保到期时间小于入院时间,已调整为入院时间后3天!", vbInformation, gstrSysName
                End If
                dtp担保时限.MinDate = CDate(txt入院时间.Text)
            End If
        End If
    End If
End Sub

Private Sub txt门诊诊断_GotFocus()
    zlControl.TxtSelAll txt门诊诊断
    If Not RequestCode Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub

Private Sub txt门诊诊断_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '问题25785 by lesfeng 2009-10-20 处理允许自由录入规则
            '************************************************
            If gint门诊诊断输入 = 1 Then
                strInput = UCase(txt门诊诊断.Text)
                strSex = zlCommFun.GetNeedName(cbo性别.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "名称 Like [2] or '('||编码||')'||名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gbytCode = 0, "简码", "五笔码") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gbytCode = 0, "简码", "五笔码 as 简码") & ",说明" & _
                        " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                        " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by 编码"
                '问题27613 by lesfeng 2010-01-21
                '自由录入时有多个匹配(汉字)不进行选择,数字及字母则进行选择
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "D", strSex, gbytCode + 1)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                        End If
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fra入院.hWnd, txt门诊诊断.Left, txt门诊诊断.Top)
                    strInput = UCase(txt门诊诊断.Text)
                    strSex = zlCommFun.GetNeedName(cbo性别.Text)
                    lngTxtHeight = txt门诊诊断.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight, 1)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '数据库中只有一个匹配项目，则以该匹配的项目为准
                    txt门诊诊断.Tag = rsTmp!ID
                    txt门诊诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称 '
                    lbl门诊诊断.Tag = txt门诊诊断.Text '用于恢复显示
                Else
                    '多项或者无匹配项目时才以输入的为准
                    txt门诊诊断.Tag = ""
                    lbl门诊诊断.Tag = txt门诊诊断.Text '用于恢复显示
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt门诊诊断.Text = lbl门诊诊断.Tag And txt门诊诊断.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt门诊诊断.Text = "" Then
            txt门诊诊断.Tag = "": lbl门诊诊断.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            vPoint = zlControl.GetCoordPos(fra入院.hWnd, txt门诊诊断.Left, txt门诊诊断.Top)
            strInput = UCase(txt门诊诊断.Text)
            strSex = zlCommFun.GetNeedName(cbo性别.Text)
            lngTxtHeight = txt门诊诊断.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight, 1)
            
            If Not rsTmp Is Nothing Then
                txt门诊诊断.Tag = rsTmp!ID
                txt门诊诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                lbl门诊诊断.Tag = txt门诊诊断.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If lbl门诊诊断.Tag <> "" Then txt门诊诊断.Text = lbl门诊诊断.Tag
                Call txt门诊诊断_GotFocus
                txt门诊诊断.SetFocus
            End If
        End If
    Else
        CheckInputLen txt门诊诊断, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt门诊诊断_Validate(Cancel As Boolean)
    If Val(txt门诊诊断.Tag) > 0 And txt门诊诊断.Text <> lbl门诊诊断.Tag Then
        txt门诊诊断.Text = lbl门诊诊断.Tag
    ElseIf Val(txt门诊诊断.Tag) = 0 And RequestCode Then
        txt门诊诊断.Text = ""
    End If
End Sub

Private Sub txt身份证号_LostFocus()
    '问题号:53408
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
    '问题81342
    If Trim(txt身份证号.Text) = "" And cboIDNumber.Visible Then
        cboIDNumber.Enabled = True
        cboIDNumber.SetFocus
    Else
        cboIDNumber.Enabled = False
        cboIDNumber.ListIndex = -1
    End If
    Call ReLoadCardFee
End Sub

Private Sub txt验证密码_GotFocus()
    Call zlControl.TxtSelAll(txt验证密码)
    Call OpenPassKeyboard(txt验证密码, False)
End Sub

Private Sub txt验证密码_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCurSendCard.int密码规则 = 1)
End Sub

Private Sub txt验证密码_LostFocus()
    Call ClosePassKeyboard(txt验证密码)
End Sub

Private Sub txt支付密码_GotFocus()
    Call zlControl.TxtSelAll(txt支付密码)
    Call OpenPassKeyboard(txt支付密码, False)
End Sub

Private Sub txt支付密码_KeyPress(KeyAscii As Integer)
    Call CheckInputPassWord(KeyAscii, mCurSendCard.int密码规则 = 1)
End Sub

Private Sub txt支付密码_LostFocus()
    Call ClosePassKeyboard(txt支付密码)
End Sub

Private Sub txt中医诊断_GotFocus()
    zlControl.TxtSelAll txt中医诊断
    If Not RequestCode Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub
Private Sub CheckInputPassWord(KeyAscii As Integer, Optional ByVal blnOnlyNum As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查密码输入
    '编制:刘兴洪
    '日期:2011-07-07 00:40:53
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
Private Sub txt中医诊断_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '问题25785 by lesfeng 2009-10-20 处理允许自由录入规则
            '************************************************
            If gint门诊诊断输入 = 1 Then
                strInput = UCase(txt中医诊断.Text)
                strSex = zlCommFun.GetNeedName(cbo性别.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "名称 Like [2] or '('||编码||')'||名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gbytCode = 0, "简码", "五笔码") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gbytCode = 0, "简码", "五笔码 as 简码") & ",说明" & _
                        " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                        " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by 编码"
                
                '问题27613 by lesfeng 2010-01-21
                '自由录入时有多个匹配(汉字)不进行选择,数字及字母则进行选择
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "B", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fra入院.hWnd, txt中医诊断.Left, txt中医诊断.Top)
                    strInput = UCase(txt中医诊断.Text)
                    strSex = zlCommFun.GetNeedName(cbo性别.Text)
                    lngTxtHeight = txt中医诊断.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight, 1)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '数据库中只有一个匹配项目，则以该匹配的项目为准
                    txt中医诊断.Tag = rsTmp!ID
                    txt中医诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称 '
                    lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                Else
                    '多项或者无匹配项目时才以输入的为准
                    txt中医诊断.Tag = ""
                    lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt中医诊断.Text = lbl中医诊断.Tag And txt中医诊断.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt中医诊断.Text = "" Then
            txt中医诊断.Tag = "": lbl中医诊断.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            vPoint = zlControl.GetCoordPos(fra入院.hWnd, txt中医诊断.Left, txt中医诊断.Top)
            strInput = UCase(txt中医诊断.Text)
            strSex = zlCommFun.GetNeedName(cbo性别.Text)
            lngTxtHeight = txt中医诊断.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight, 1)
            If Not rsTmp Is Nothing Then
                txt中医诊断.Tag = rsTmp!ID
                txt中医诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的中医疾病编码。", vbInformation, gstrSysName
                End If
                If lbl中医诊断.Tag <> "" Then txt中医诊断.Text = lbl中医诊断.Tag
                Call txt中医诊断_GotFocus
                txt中医诊断.SetFocus
            End If
        End If
    Else
        CheckInputLen txt中医诊断, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt中医诊断_LostFocus()
    If Not RequestCode Then
        Call zlCommFun.OpenIme
    End If
End Sub

Private Sub txt中医诊断_Validate(Cancel As Boolean)
    If Val(txt中医诊断.Tag) > 0 And txt中医诊断.Text <> lbl中医诊断.Tag Then
        txt中医诊断.Text = lbl中医诊断.Tag
    ElseIf Val(txt中医诊断.Tag) = 0 And RequestCode Then
        txt中医诊断.Text = ""
    End If
End Sub

Private Sub txt身份证号_Change()
    Dim strBirthDay  As String
    Dim strAge As String
    Dim strSex As String
    Dim strErrInfo As String
    
    If mblnChange Then
        If CreatePublicPatient() Then
            If gobjPublicPatient.CheckPatiIdcard(Trim(txt身份证号.Text), strBirthDay, strAge, strSex, strErrInfo) Then
                If IsDate(strBirthDay) Then
                    txt出生日期.Enabled = True
                    txt出生时间.Enabled = True
                End If
                If txt出生日期.Enabled = True Then txt出生日期.Text = strBirthDay
                If cbo性别.Enabled Then Call cbo.Locate(cbo性别, strSex, False)
            End If
        End If
    End If
End Sub

Private Sub txt姓名_LostFocus()
    Call zlCommFun.OpenIme
    txt姓名.Text = Trim(txt姓名.Text)
End Sub

Private Sub txt医保号_GotFocus()
    Call zlControl.TxtSelAll(txt医保号)
End Sub

Private Sub txt医保号_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    '允许输字符
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ~!@#$%^&*()_+|-=\[]{}<>,./" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    
    '医疗付款方式缺省=社会基本医疗保险
    If txt医保号.Text <> "" Then
        For i = 0 To cbo医疗付款.ListCount
            If InStr(cbo医疗付款.List(i), Chr(&HD)) > 0 Then cbo医疗付款.ListIndex = i: Exit For
        Next
    End If
End Sub

Private Sub txt预交额_LostFocus()
    '65965:刘鹏飞,2013-09-24,处理预交显示千位位格式
    If IsNumeric(txt预交额.Text) Then
        txt预交额.Text = Format(StrToNum(txt预交额.Text), "##,##0.00;-##,##0.00; ;")
    Else
        txt预交额.Text = ""
    End If
    If txt预交额.MaxLength > 12 Then txt预交额.MaxLength = 12
    
    If gblnLED Then
        '#22 1234.56   --预收一千二百三十四点五六元 Y
        '#23 1234.56   --找零一千二百三十四点五六元 Z
        zl9LedVoice.Speak "#22 " & StrToNum(txt预交额.Text)
    End If
End Sub

Private Sub txt缴款单位_GotFocus()
    If IsNumeric(txt预交额.Text) And txt缴款单位.Text = "" Then
        txt缴款单位.Text = txt工作单位.Text
    End If
    zlControl.TxtSelAll txt缴款单位
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt缴款单位_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt缴款单位, KeyAscii
End Sub

Private Sub txt单位开户行_KeyPress(KeyAscii As Integer)
    CheckInputLen txt单位开户行, KeyAscii
End Sub

Private Sub txt单位帐号_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt单位帐号, KeyAscii
End Sub

Private Sub txt担保额_GotFocus()
    zlControl.TxtSelAll txt担保额
End Sub

Private Sub txt担保人_GotFocus()
    zlControl.TxtSelAll txt担保人
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt担保人_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt担保人, KeyAscii
End Sub

Private Sub txt工作单位_Change()
    If txt工作单位.Text = "" Then txt工作单位.Tag = ""
End Sub

Private Sub txt结算号码_GotFocus()
    zlControl.TxtSelAll txt结算号码
End Sub

Private Sub txt预交额_GotFocus()
    If IsNumeric(txt预交额.Text) Then
        txt预交额.Text = StrToNum(txt预交额.Text)
    Else
        txt预交额.Text = ""
    End If
    txt预交额.SelStart = 0: txt预交额.SelLength = Len(txt预交额.Text)
End Sub

Private Sub txt预交额_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii <> 13 Then
        If InStr(txt预交额.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        '65965:刘鹏飞,2013-09-24,处理预交显示千位位格式
        If (txt预交额.Text <> "" And txt预交额.SelLength <> Len(Format(StrToNum(txt预交额.Text), "##,##0.00;-##,##0.00; ;"))) And _
            (Len(Format(StrToNum(txt预交额.Text), "##,##0.00;-##,##0.00; ;")) >= txt预交额.MaxLength) And _
            InStr(Chr(8), Chr(KeyAscii)) = 0 Then
            If txt预交额.SelLength > 0 And txt预交额.SelLength <= txt预交额.MaxLength Then
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf IsNumeric(txt预交额.Text) Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        '不收取预交款,直接跳过
        txt预交额.Text = ""
        If pic磁卡.Visible Then
            If Not mrsInfo Is Nothing Then
                If Not IsNull(mrsInfo!就诊卡号) Then
                    cmdOK.SetFocus
                Else
                    txt卡号.SetFocus
                End If
            Else
                txt卡号.SetFocus
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
    Dim strTmp As String, curFee As Currency, cur门诊未结 As Currency
    Dim i As Integer, blnICCard As Boolean
    Dim blnCard As Boolean
    Dim str住院号 As String
    
    If txtPatient.Locked Then Exit Sub
    '特殊字符过滤在Form_KeyPress中进行
    
    '直接输入病人信息,以新病人保存,并清除原先的病人信息
    If KeyAscii = 13 Then
        If Trim(txtPatient.Text) = "" Then
            Call ClearCard(True) '只清除病人信息
            '产生新的病人ID及住院号
            txtPatient.Text = zlDatabase.GetNextNo(1)
            txtPatient.Tag = txtPatient.Text
            
            '留观病人不自动生成住院号
            If mbytKind = E住院入院登记 Then
                txt住院号.Text = zlDatabase.GetNextNo(2)
                If Not txt住院号.Locked Then
                    txt住院号.SetFocus
                Else
                    txt姓名.SetFocus
                End If
            ElseIf mbytKind = E门诊留观登记 Then
                'txt住院号.Locked = True
                txt住院号.Text = zlDatabase.GetNextNo(3)
                mblnAuto = True
                If Not txt住院号.Locked Then
                    txt住院号.SetFocus
                Else
                    txt姓名.SetFocus
                End If
'                txt姓名.SetFocus
            ElseIf mbytKind = E住院留观登记 Then '住院留观号不能修改，每次新登记时按照留观号规则自动产生
                txt住院号.Text = zlDatabase.GetNextNo(6)
                txt姓名.SetFocus
            Else
                txt姓名.SetFocus
            End If
            Exit Sub
        ElseIf txtPatient.Text = txtPatient.Tag Then
            If Not txt住院号.Locked Then
                txt住院号.SetFocus
            Else
                txt姓名.SetFocus
            End If
            Exit Sub
        End If
    End If

    If IDKind.GetCurCard.名称 Like "姓名*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        If blnCard And IDKind.ShowPassText Then txtPatient.PasswordChar = "*"
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    End If
    '55571:刘鹏飞,2012-11-12
    txtPatient.IMEMode = 0
    
    On Error GoTo errHandle
    
    '刷卡完毕或输入号码后回车
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txtPatient.Text <> "" Then
        
        '37662
        If Not InStr(gstrPrivs, "修改病人信息") > 0 Then
            txt姓名.Locked = True
            cbo性别.Locked = True
            txt年龄.Locked = True
            cbo年龄单位.Locked = True
        End If
    
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        If IDKind.GetCurCard.名称 Like "IC卡*" And IDKind.GetCurCard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        mblnICCard = blnICCard
        
        
        '读取病人信息
        If GetPatient(IDKind.GetCurCard, txtPatient.Text, blnCard) Then
            Led欢迎信息
            
            If Not isValid(mrsInfo!病人ID) Then txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
            '就诊卡密码检查:3次
            If gblnCheckPass And (blnCard Or blnICCard) Then
                If zlCommFun.VerifyPassWord(Me, mstrPassWord) = False Then
                    txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
                End If
            End If
            
            '已经正式登记(预约病人在病人信息中没有填当前科室)
            If Not IsNull(mrsInfo!当前科室id) Then
                MsgBox """" & mrsInfo!姓名 & """已经登记为" & Decode(mrsInfo!病人性质, 0, "入院", 1, "门诊留观", 2, "住院留观") & "病人。", vbInformation, gstrSysName
                txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
            End If
            
            '住院登记界面接收病人
            If Not IsNull(mrsInfo!主页ID) And mbytInState = EState.E新增 And mbytMode <> EMode.E接收预约 Then '没有住过院的病人的主页id为空(因为是两表外连接查询)
                If mrsInfo!主页ID = 0 Then '已经预约的病人(没有提供留观预约)
                    If mbytMode = EMode.E预约登记 Or mbytMode = EMode.E正常登记 And mbytKind <> EKind.E住院入院登记 Then
                        MsgBox """" & mrsInfo!姓名 & """已经预约登记。", vbInformation, gstrSysName
                        txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
                    Else
                        strTmp = ""
                        If InStr(mstrPrivs, "接收预约") = 0 Then MsgBox "您没有【接收预约】的权限， 不能接收预约病人！", vbInformation, gstrSysName: Exit Sub
                        If InStr(mstrPrivs, "接收住院预约") = 0 And mrsInfo!病人性质 = 0 Then
                            MsgBox "您没有【接收住院预约】的权限， 不能接收住院预约病人！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        If InStr(mstrPrivs, "接收门诊留观预约") = 0 And mrsInfo!病人性质 = 1 Then
                            MsgBox "您没有【接收门诊留观预约】的权限， 不能接收门诊留观预约病人！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        If InStr(mstrPrivs, "接收住院留观预约") = 0 And mrsInfo!病人性质 = 2 Then
                            MsgBox "您没有【接收住院留观预约】的权限， 不能接收住院留观预约病人！", vbInformation, gstrSysName
                            Exit Sub
                        End If

                        If InStr(mstrPrivs, "门诊留观登记") = 0 And InStr(mstrPrivs, "住院留观登记") = 0 Then
                            If InStr(mstrPrivs, "接收住院预约") = 0 Then
                                MsgBox "您没有足够的用户权限， 不能接收预约病人！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If MsgBox("要将""" & mrsInfo!姓名 & """接收为住院病人吗?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then strTmp = "住院病人"
                        ElseIf InStr(mstrPrivs, "住院留观登记") = 0 Then
                            If InStr(mstrPrivs, "接收门诊留观预约") <> 0 And InStr(mstrPrivs, "接收住院预约") <> 0 Then
                                strTmp = "!住院病人(&0),门诊留观(&1)"
                            ElseIf InStr(mstrPrivs, "接收门诊留观预约") <> 0 Then
                                strTmp = "!门诊留观(&0)"
                            ElseIf InStr(mstrPrivs, "接收住院预约") <> 0 Then
                                strTmp = "!住院病人(&0)"
                            Else
                                MsgBox "您没有足够的用户权限， 不能接收预约病人！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            strTmp = zlCommFun.ShowMsgBox("预约接收", "要将""" & mrsInfo!姓名 & """接收为", strTmp, Me, vbQuestion)
                        ElseIf InStr(mstrPrivs, "门诊留观登记") = 0 Then
                            If InStr(mstrPrivs, "接收住院留观预约") <> 0 And InStr(mstrPrivs, "接收住院预约") <> 0 Then
                                strTmp = "!住院病人(&0),住院留观(&1)"
                            ElseIf InStr(mstrPrivs, "接收住院留观预约") <> 0 Then
                                strTmp = "!住院留观(&0)"
                            ElseIf InStr(mstrPrivs, "接收住院预约") <> 0 Then
                                strTmp = "!住院病人(&0)"
                            Else
                                MsgBox "您没有足够的用户权限， 不能接收预约病人！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            strTmp = zlCommFun.ShowMsgBox("预约接收", "要将""" & mrsInfo!姓名 & """接收为", strTmp, Me, vbQuestion)
                        Else
                            If InStr(mstrPrivs, "接收住院留观预约") <> 0 And InStr(mstrPrivs, "接收住院预约") <> 0 And InStr(mstrPrivs, "接收门诊留观预约") <> 0 Then
                                strTmp = "!住院病人(&0),门诊留观(&1),住院留观(&2)"
                            ElseIf InStr(mstrPrivs, "接收住院留观预约") <> 0 And InStr(mstrPrivs, "接收住院预约") <> 0 Then
                                strTmp = "!住院病人(&0),住院留观(&1)"
                            ElseIf InStr(mstrPrivs, "接收住院留观预约") <> 0 And InStr(mstrPrivs, "接收门诊留观预约") <> 0 Then
                                strTmp = "!门诊留观(&0),住院留观(&1)"
                            ElseIf InStr(mstrPrivs, "接收住院预约") <> 0 And InStr(mstrPrivs, "接收门诊留观预约") <> 0 Then
                                strTmp = "!住院病人(&0),门诊留观(&1)"
                            ElseIf InStr(mstrPrivs, "接收门诊留观预约") <> 0 Then
                                strTmp = "!门诊留观(&0)"
                            ElseIf InStr(mstrPrivs, "接收住院预约") <> 0 Then
                                strTmp = "!住院病人(&0)"
                            ElseIf InStr(mstrPrivs, "接收住院留观预约") <> 0 Then
                                strTmp = "!住院留观(&0)"
                            Else
                                MsgBox "您没有足够的用户权限， 不能接收预约病人！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            strTmp = zlCommFun.ShowMsgBox("预约接收", "要将""" & mrsInfo!姓名 & """接收为", strTmp, Me, vbQuestion)
                        End If
                        
                        If strTmp = "" Then txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Sub
                        
                        mbytKind = Switch(strTmp = "住院病人", 0, strTmp = "门诊留观", E门诊留观登记, strTmp = "住院留观", E住院留观登记)
                        
                        cmdName.Visible = True
                        cmdTurn.Visible = InStr(1, mstrPrivs, ";门诊费用转住院;") > 0 And mbytKind = E住院入院登记 And mbytMode <> 1
                        
                        txtTimes.Visible = mbytMode <> 1 And mbytKind = E住院入院登记 '预约登记时或留关登记时,住院次数为零
                        lblTimes.Visible = mbytMode <> 1 And mbytKind = E住院入院登记
                        txtTimes.Enabled = (InStr(1, mstrPrivs, "修改住院次数") > 0 And mbytInState = 0)   '修改时不允许改，因为可能已产生住院一次费用，预交款，就诊卡
                            
                        If Not InitData Then Unload Me: Exit Sub
                        Me.Caption = "接收" & strTmp
                        mlng病人ID = mrsInfo!病人ID: mlng主页ID = 0
                        Call zlCommFun.PressKey(vbKeyTab)
                        txtPatient.Locked = True: txtPatient.TabStop = False
                        
                        If mbytKind = E门诊留观登记 Then     '门诊留观
'                            lbl住院号.Visible = False
'                            txt住院号.Visible = False
                            lbl住院号.Caption = "门诊号"
                            cmdSelectNO.Visible = False
                            cmdYB.Visible = False
                        ElseIf mbytKind = E住院留观登记 Then     '住院留观
                            lbl住院号.Caption = "留观号"
                            txt住院号.Locked = True
                            cmdSelectNO.Visible = False
                        End If
                                                
                        If Not ReadPatiReg(mrsInfo!病人ID, 0) Then
                            MsgBox "不能正确读取预约病人""" & mrsInfo!姓名 & """的登记记录！", vbInformation, gstrSysName
                            Call ClearCard
                            Exit Sub
                        End If
                        
                         '如果之前没有住院号或每次住院产生新住院号,接收为住院病人，则自动分配新的住院号
                        '问题 27063 by lesfeng 2009-12-25 预约登记转住院病人保留原住院号(取消gbln每次住院新住院号判断)
        '                If mbytKind = EKind.E住院入院登记 And (Trim(txt住院号.Text) = "" Or gbln每次住院新住院号) Then txt住院号.Text = zlDatabase.GetNextNo(2)
                        '85510:LPF,2015-06-19,预约登记住院号产生规则（医嘱登记入院处理,因医嘱登记插入住院号时不可能重写住院业务规则）:
                        '原有逻辑判断:If mbytKind = EKind.E住院入院登记 And (Trim(txt住院号.Text) = "") Then txt住院号.Text = zlDatabase.GetNextNo(2)
                        '入院管理，预约登记会根据参数"每次住院新住院号"生成住院号,而医嘱登记目前只是以病人信息的住院号为准插入(这种方式产生的住院号可能就不正确)
                        '因此需要做如下处理
                        '1:gbln每次住院新住院号=TRUE,如果存在住院号，则检查已有的住院号是否重复，如果重复则重新生成。
                        '2:gbln每次住院新住院号=FALSE,如果住院号为空,则使用历史住院号(最后一次住院号不为空)，不存在历史住院则重新生成。
                        If mbytKind = EKind.E住院入院登记 Then
                            If gbln每次住院新住院号 = True Then
                                If Trim(txt住院号.Text) <> "" Then
                                    If CheckByPatiNO(mrsInfo!病人ID, 0, 0, Trim(txt住院号.Text)) = True Then txt住院号.Text = ""
                                End If
                            Else
                                If Trim(txt住院号.Text) = "" Then
                                    str住院号 = ""
                                    If CheckByPatiNO(mrsInfo!病人ID, 0, 1, str住院号) = True Then txt住院号.Text = str住院号
                                End If
                            End If
                            If Trim(txt住院号.Text) = "" Then txt住院号.Text = zlDatabase.GetNextNo(2)
                        ElseIf mbytKind = E住院留观登记 Then
                            If Trim(txt住院号.Text) = "" Then txt住院号.Text = zlDatabase.GetNextNo(6)
                        End If
                
                        Exit Sub
                    End If
                End If
            End If
            
            
            '黑名单提醒
            strTmp = inBlackList(mrsInfo!病人ID)
            If strTmp <> "" Then
                If MsgBox("病人""" & mrsInfo!姓名 & """在特殊病人名单中。" & vbCrLf & vbCrLf & "原因：" & vbCrLf & vbCrLf & "　　" & strTmp & vbCrLf & vbCrLf & "要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Call ClearCard(True): txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.SetFocus: Exit Sub
                End If
            End If
            
            '病人费用余额提醒
            curFee = GetPatientUnBalance(mrsInfo!病人ID, cur门诊未结)
            If cur门诊未结 <> 0 Or curFee <> 0 Then
                strTmp = ""
                If cur门诊未结 <> 0 Then strTmp = "门诊费用" & Format(cur门诊未结, "0.00")
                If curFee <> 0 Then strTmp = strTmp & IIf(strTmp = "", "", ",") & "住院费用" & Format(curFee, "0.00")
                                
                strTmp = "提醒：""" & mrsInfo!姓名 & """有未结清" & strTmp
                If mbytMode = EMode.E接收预约 Then
                    MsgBox strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName
                Else
                    If MsgBox(strTmp & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call ClearCard(True): txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.SetFocus: Exit Sub
                    End If
                End If
            End If
            
            
            '检查是否有应收款
            strTmp = "Select Zl_Patientdue([1]) 剩余应收 From dual"
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "提取应收款", CLng(mrsInfo!病人ID))
            If Not rsTmp.EOF Then
                If Nvl(rsTmp!剩余应收, 0) > 0 Then
                    If mbytMode = EMode.E接收预约 Then
                        MsgBox "该病人尚有 " & rsTmp!剩余应收 & "元 应收款未缴！", vbInformation, gstrSysName
                    Else
                        If MsgBox("该病人尚有 " & rsTmp!剩余应收 & "元 应收款未缴！要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Call ClearCard(True): txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.SetFocus: Exit Sub
                        End If
                    End If
                End If
            End If
            
            '---------------------------------------------------------------------------------------
            'If mstrYBPati <> "" Then txt姓名.ForeColor = vbRed
            
            '病人诊断记录
            If mstrYBPati = "" Then
                Call ClearCard(True, True)
            ElseIf RequestCode Then
                If Val(txt门诊诊断.Tag) = 0 Then
                    txt门诊诊断.Text = "": txt门诊诊断.Tag = "": lbl门诊诊断.Tag = ""
                End If
                If Val(txt中医诊断.Tag) = 0 Then
                    txt中医诊断.Text = "": txt中医诊断.Tag = "": lbl中医诊断.Tag = ""
                End If
            End If
            
            Set rsTmp = GetDiagnosticInfo(mrsInfo!病人ID, 0, "1,11", "3")
            If Not rsTmp Is Nothing Then
                rsTmp.Filter = "诊断类型=1"
                If Not rsTmp.EOF Then
                    txt门诊诊断.Text = Nvl(rsTmp!诊断描述): txt门诊诊断.Tag = Nvl(rsTmp!疾病ID, rsTmp!诊断ID & ";"): lbl门诊诊断.Tag = txt门诊诊断.Text
                End If
                
                rsTmp.Filter = "诊断类型=11"
                If Not rsTmp.EOF Then
                    txt中医诊断.Text = Nvl(rsTmp!诊断描述): txt中医诊断.Tag = Nvl(rsTmp!疾病ID, rsTmp!诊断ID & ";"): lbl中医诊断.Tag = txt门诊诊断.Text
                End If
            End If
            txt入院时间.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
            If Not IsNull(mrsInfo!就诊卡号) Then txt卡号.TabStop = False
            '填写病人信息
            If Not FuncPlugPovertyInfo(Val(mrsInfo!病人ID)) Then Exit Sub
            Call FillPatient
            'EMPI
            Call EMPI_LoadPati(1)
            '更新卡费
            Call ReLoadCardFee(True)
            cbo病人类型.Enabled = InStr(mstrPrivs, "调整病人类型") > 0
            If mbytInState = 0 And cbo入院科室.ListIndex >= 0 Then
                chk再入院.Value = IIf(CheckReIN(mrsInfo!病人ID, Val(cbo入院科室.ItemData(cbo入院科室.ListIndex))), 1, 0)
            End If
            If CanFocus(cbo入院科室) Then cbo入院科室.SetFocus
        ElseIf (blnCard Or blnICCard) And pic磁卡.Visible Then  '发新卡
            MsgBox "该卡没有建档,将作为新卡登记,请输入病人姓名。", vbInformation, gstrSysName
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
            txt卡号.Text = txtPatient.Text
            txtPatient.Text = zlDatabase.GetNextNo(1)
            txtPatient.Tag = txtPatient.Text
            txt入院时间.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            If mbytKind = E住院入院登记 Then
                txt住院号.Text = zlDatabase.GetNextNo(2)
            ElseIf mbytKind = E住院留观登记 Then
                txt住院号.Text = zlDatabase.GetNextNo(6)
            End If
            
            Call CheckFreeCard(txt卡号.Text)
            txt姓名.Locked = False
            cbo性别.Locked = False
            txt年龄.Locked = False
            cbo年龄单位.Locked = False
            txt姓名.SetFocus
        ElseIf Not IDKind.GetCurCard.名称 = "身份证号" Then
            MsgBox "没有找到指定的病人。", vbInformation, gstrSysName
            Call zlControl.TxtSelAll(txtPatient)
            txtPatient.SetFocus
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPatientUnBalance(ByVal lng病人ID As Long, ByRef cur门诊未结 As Currency) As Currency
'功能：获取指定病人未结费用,不含体检未结费用
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 来源途径, Sum(金额) 金额 From 病人未结费用 Where 病人id=[1] and 来源途径 in(1,2) Group By 来源途径"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "来源途径=1"
        If rsTmp.RecordCount > 0 Then cur门诊未结 = Val("" & rsTmp!金额)
        rsTmp.Filter = "来源途径=2"
        If rsTmp.RecordCount > 0 Then GetPatientUnBalance = Val("" & rsTmp!金额)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPatient_LostFocus()
    '如果病人已确定,而显示被破坏,则恢复
    If txtPatient.Tag <> "" Then
        txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
        txtPatient.Text = txtPatient.Tag
    End If
    If gblnSeekName Then Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
End Sub

Private Sub txt出生地点_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt出生地点.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt出生地点)
            If Not rsTmp Is Nothing Then
                txt出生地点.Text = rsTmp!名称
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt出生地点, KeyAscii
    End If
End Sub

Private Sub txt单位电话_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt单位邮编_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If ((Not IsNumeric(txt单位邮编.Text)) Or Len(txt单位邮编.Text) > 6 Or InStr(txt单位邮编.Text, ".") > 0) And txt单位邮编.Text <> "" Then
            Call SelectYouBian(txt单位邮编)
        End If
    End If
End Sub

Private Sub txt工作单位_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt工作单位.Text <> "" Then
            Set rsTmp = GetOrgAddress(Me, txt工作单位)
            If Not rsTmp Is Nothing Then
                txt工作单位.Text = rsTmp!名称
                txt工作单位.Tag = rsTmp!ID
                txt单位电话.Text = Trim(rsTmp!电话 & "")
                txt单位开户行.Text = Trim(rsTmp!开户银行 & "")
                txt单位帐号.Text = Trim(rsTmp!帐号 & "")
            Else
                txt工作单位.Tag = ""
            End If
        Else
            txt工作单位.Tag = ""
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt工作单位, KeyAscii
    End If
End Sub

Private Sub txt家庭地址邮编_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If ((Not IsNumeric(txt家庭地址邮编.Text)) Or Len(txt家庭地址邮编.Text) > 6 Or InStr(txt家庭地址邮编.Text, ".") > 0) And txt家庭地址邮编.Text <> "" Then
            Call SelectYouBian(txt家庭地址邮编)
        End If
    End If
End Sub

Private Sub txt家庭地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt家庭地址.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt家庭地址)
            If Not rsTmp Is Nothing Then
                txt家庭地址.Text = rsTmp!名称
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt家庭地址, KeyAscii
    End If
End Sub

Private Sub txt家庭电话_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt结算号码_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt结算号码, KeyAscii
End Sub

Private Sub txt卡额_KeyPress(KeyAscii As Integer)
    If txt卡额.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not mCurSendCard.rs卡费 Is Nothing Then
            If mCurSendCard.rs卡费!是否变价 = 1 Then
                If mCurSendCard.rs卡费!现价 <> 0 And Abs(CCur(txt卡额.Text)) > Abs(mCurSendCard.rs卡费!现价) Then
                    MsgBox "" & mCurSendCard.str卡名称 & "金额绝对值不能大于最高限价：" & Format(Abs(mCurSendCard.rs卡费!现价), "0.00"), vbInformation, gstrSysName
                    txt卡额.SetFocus: Call zlControl.TxtSelAll(txt卡额): Exit Sub
                End If
                If mCurSendCard.rs卡费!原价 <> 0 And Abs(CCur(txt卡额.Text)) < Abs(mCurSendCard.rs卡费!原价) Then
                    MsgBox "" & mCurSendCard.str卡名称 & "金额绝对值不能小于最低限价：" & Format(Abs(mCurSendCard.rs卡费!原价), "0.00"), vbInformation, gstrSysName
                    txt卡额.SetFocus: Call zlControl.TxtSelAll(txt卡额): Exit Sub
                End If
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr(txt卡额.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt卡号_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii <> 13 Then
        If InStr(":：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        If Len(txt卡号.Text) = mCurSendCard.lng卡号长度 - 1 And KeyAscii <> 8 Then
            txt卡号.Text = txt卡号.Text & Chr(KeyAscii)
            
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf txt卡号.Text = "" Then
        KeyAscii = 0: cmdOK.SetFocus  '不发卡,直接跳过
    Else
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt卡号_Validate(Cancel As Boolean)
    Dim lngBindPatientID As Long '绑定卡的病人ID
    Dim lng变动类型 As Long '卡片最后的变动类型 11-绑定卡,1-发卡
    txt卡号.Text = Trim(txt卡号.Text)
    Call ReLoadCardFee
    Call CheckFreeCard(txt卡号.Text)
    If mCurSendCard.lng卡号长度 = Len(Trim(txt卡号.Text)) Then
        '卡是否已经绑定或者发卡
        If WhetherTheCardBinding(Trim(txt卡号.Text), mCurSendCard.lng卡类别ID, lngBindPatientID) Then
            
            If mCurSendCard.bln自制卡 And mCurSendCard.bln重复利用 And lngBindPatientID > 0 Then
            
                lng变动类型 = GetCardLastChangeType(Trim(txt卡号.Text), mCurSendCard.lng卡类别ID, lngBindPatientID)
                If lng变动类型 = 11 Then
                    '如果是绑定
                    If MsgBox("卡号为【" & txt卡号.Text & "】的{" & mCurSendCard.str卡名称 & "}的卡已经与病人标识为【" & lngBindPatientID & "】的进行了绑定！" & vbCrLf & "是否取消该卡的绑定?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                        Cancel = True
                        txt卡号.Text = ""
                        Exit Sub
                    End If
                    If BlandCancel(mCurSendCard.lng卡类别ID, Trim(txt卡号.Text), lngBindPatientID) Then
                        Exit Sub
                    End If
                End If
                
            End If
            
            MsgBox "该卡号已经被绑定,不能继续.", vbInformation, gstrSysName
            Cancel = True
            txt卡号.Text = ""
            Exit Sub
            
        End If
    End If
End Sub

Private Sub CheckFreeCard(ByVal strCard As String)
'功能：对一卡通模式下的卡号，严格控制票号时，检查是否在票据领用范围内，范围之外的卡不收费
    
    If txt卡额.Visible = False Then Exit Sub
    
    If Not mCurSendCard.rs卡费 Is Nothing And Val(txt卡额.Text) = 0 Then  '先恢复
        txt卡额.Text = Format(IIf(mCurSendCard.rs卡费!是否变价 = 1, mCurSendCard.rs卡费!缺省价格, mCurSendCard.rs卡费!现价), "0.00")
    End If
    If mblnOneCard And mCurSendCard.bln严格控制 Then
        mCurSendCard.lng领用ID = CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), strCard)
        If mCurSendCard.lng领用ID <= 0 Then txt卡额.Text = "0.00"
    End If
    If Not mCurSendCard.rs卡费 Is Nothing And Val(txt卡额.Text) <> 0 Then
        If mCurSendCard.rs卡费!是否变价 = 0 Then
            txt卡额.Text = Format(GetActualMoney(zlCommFun.GetNeedName(cbo费别.Text), mCurSendCard.rs卡费!收入项目ID, mCurSendCard.rs卡费!现价, mCurSendCard.rs卡费!收费细目ID), "0.00")
        End If
    End If
End Sub


Private Sub cbo费别_Click()
    Call LoadCardFee
End Sub

Private Sub txt开户行_GotFocus()
    If IsNumeric(txt预交额.Text) And txt开户行.Text = "" Then
        txt开户行.Text = txt单位开户行.Text
    End If
    zlControl.TxtSelAll txt开户行
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt开户行_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt开户行, KeyAscii
End Sub

Private Sub txt联系人地址_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt联系人地址.Text <> "" Then
            Set rsTmp = GetAddress(Me, txt联系人地址)
            If Not rsTmp Is Nothing Then
                txt联系人地址.Text = rsTmp!名称
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt联系人地址, KeyAscii
    End If
End Sub

Private Sub txt联系人电话_KeyPress(KeyAscii As Integer)
    If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt联系人姓名_KeyPress(KeyAscii As Integer)
    CheckInputLen txt联系人姓名, KeyAscii
End Sub

Private Sub txt入院时间_GotFocus()
    Call OS.OpenImeByName
    Call zlControl.TxtSelAll(txt入院时间)
End Sub

Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    '问题号:53408
    mbln是否扫描身份证 = False

    Call Show绑定控件(mbln是否扫描身份证 And mbln扫描身份证签约)
    
    If zl当前用户身份证是否绑定(Val(IIf(Trim(txtPatient.Text) = "", "0", Trim(CStr(txtPatient.Tag))))) = True Then
            MsgBox "当前用户的身份证号已经绑定，不允许修改其身份证号", vbInformation, gstrSysName
            KeyAscii = 0
    End If
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt年龄_GotFocus()
    Call zlCommFun.OpenIme
    zlControl.TxtSelAll txt年龄
End Sub

Private Sub txt出生日期_GotFocus()
    Call OS.OpenImeByName
    zlControl.TxtSelAll txt出生日期
End Sub

Private Sub txt身份证号_GotFocus()
    zlControl.TxtSelAll txt身份证号
    '问题号:53408
    If mbln扫描身份证签约 = True Then
        OpenIDCard
    End If
End Sub

Private Sub txt出生地点_GotFocus()
    zlControl.TxtSelAll txt出生地点
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt家庭地址_GotFocus()
    zlControl.TxtSelAll txt家庭地址
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt家庭地址邮编_GotFocus()
    zlControl.TxtSelAll txt家庭地址邮编
End Sub

Private Sub txt家庭电话_GotFocus()
    zlControl.TxtSelAll txt家庭电话
End Sub

Private Sub txt联系人姓名_GotFocus()
    zlControl.TxtSelAll txt联系人姓名
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt联系人地址_GotFocus()
    zlControl.TxtSelAll txt联系人地址
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt联系人电话_GotFocus()
    zlControl.TxtSelAll txt联系人电话
End Sub

Private Sub txt工作单位_GotFocus()
    zlControl.TxtSelAll txt工作单位
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt单位电话_GotFocus()
    zlControl.TxtSelAll txt单位电话
End Sub

Private Sub txt单位邮编_GotFocus()
    zlControl.TxtSelAll txt单位邮编
End Sub

Private Sub txt单位开户行_GotFocus()
    zlControl.TxtSelAll txt单位开户行
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt卡号_GotFocus()
    zlControl.TxtSelAll txt卡号
    Call SetBrushCardObject(True)
End Sub
Private Sub OpenIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开身份证读卡器
    '编制:王吉
    '日期:2012-08-31 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '初始化对卡对象
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '打开读卡器
    mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    Call OpenPassKeyboard(txtPass, False)
End Sub

Private Sub txt卡额_GotFocus()
    zlControl.TxtSelAll txt卡额
End Sub

Private Sub txt单位帐号_GotFocus()
    zlControl.TxtSelAll txt单位帐号
End Sub

Private Sub cmdCancel_Click()
    Select Case mbytInState
        Case 0
            If mbytMode <> EMode.E接收预约 And (txtPatient.Tag <> "" Or txt姓名.Text <> "" Or txt住院号.Text <> "") Then
                If MsgBox("确定要清除当前病人信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ClearCard
                    '84577
                    If tbcPage.Selected.Caption = "基本" Then
                        If txtPatient.Enabled Then txtPatient.SetFocus
                    Else
                        tbcPage.Item(0).Selected = True
                    End If
                End If
                Exit Sub
            ElseIf gblnOK Then
                If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            Unload Me
        Case 1
            If MsgBox("确实要放弃修改退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Unload Me
        Case 2
            Unload Me
    End Select
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
'功能：读取病人信息
'说明：提取失败时，mrsInfo = Nothing
    Dim lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
     
    On Error GoTo errH

    If blnCard = True And objCard.名称 Like "姓名*" Then   '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then
            '按手机号查询
            If IDKind.IsMobileNo(strInput) = False Then GoTo NotFoundPati:
            If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Function
        End If
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strSQL = " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = " And A.门诊号=[1]"
    Else
        Select Case objCard.名称
            Case "姓名"
                If Not gblnSeekName Then
                    MsgBox "请刷卡或输入[-病人ID]、[+住院号]、[*门诊号]等方式提取病人的信息。", vbInformation, gstrSysName
                    txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                Else
                    '通过姓名模糊查找病人(允许输入病人标识时)
                    strPati = " Select 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
                        " A.住院号,A.门诊号,A.住院次数,trunc(C.入院日期,'dd') as 入院日期,trunc(C.出院日期,'dd') as 出院日期,A.出生日期,A.身份证号,A.手机号,A.家庭地址,A.工作单位,zl_PatiType(A.病人ID) 病人类型" & _
                        " From 病人信息 A,部门表 B,病案主页 C" & _
                        " Where A.停用时间 is NULL And A.病人ID=C.病人ID(+) And Nvl(A.主页ID,0)=C.主页ID(+) And A.当前科室ID=B.ID(+) And Rownum<101" & _
                        " And A.姓名 Like [1]" & IIf(gintNameDays = 0, "", " And (A.登记时间>Trunc(Sysdate-[2]) Or A.就诊时间>Trunc(Sysdate-[2]))")
                    strPati = strPati & " Union ALL " & _
                            "Select 0,0,-NULL,'[新病人]',NULL,NULL,-NULL,-NULL,-NULL,To_Date(NULL),To_Date(NULL),To_Date(NULL),NULL,NULL,NULL,NULL,'普通病人' From Dual"
                    strPati = strPati & " Order by 排序ID,姓名,入院日期 Desc"
                    
                    vRect = GetControlRect(txtPatient.hWnd)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", gintNameDays)
                                
                    '只有一行数据时,blncancel返回false,按取消返回也是一样
                    If Not blnCancel Then
                        If rsTmp!ID = 0 Then '当作新病人
                            strPati = txtPatient.Text
                            txtPatient.Text = ""
                            txtPatient_KeyPress (13)
                            txt姓名.Text = strPati
                            Exit Function
                        Else '以病人ID读取
                            strInput = rsTmp!病人ID
                            strSQL = " And A.病人ID=[2]"
                        End If
                    Else
                        Call zlControl.TxtSelAll(txtPatient)
                        txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                    End If
                End If
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = " And A.医保号=[2]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = Val(objCard.接口序号)
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    '入院之前的担保信息无效,在院病人的担保以本次入院登记的为准
    '预约登记不填写"病人信息.当前科室ID,住院次数"等,而留观登记要
    '如果有预约的就要读预约记录,否则读最近一次住院的记录
    '60500:刘鹏飞,2013-05-09,如果病人上一次留观,病人信息住院号为空
    If mbytInState = 0 And gbln每次住院新住院号 = False Then
            strPati = "Nvl(a.住院号," & vbNewLine & _
            "            (SELECT 住院号" & vbNewLine & _
            "             FROM 病案主页" & vbNewLine & _
            "             WHERE 病人id = a.病人id AND" & vbNewLine & _
            "                   主页id = (SELECT MAX(主页id) FROM 病案主页 WHERE 病人id = a.病人id AND 住院号 IS NOT NULL))) 住院号,"
    Else
        strPati = "A.住院号,"
    End If
    '65973:刘鹏飞,2013-09-29,新登记提取医疗付款方式
    strSQL = "Select A.病人id, B.主页id, A.住院次数, A.就诊卡号, A.卡验证码, A.门诊号,B.留观号," & strPati & "A.姓名, A.性别,A.年龄, C.名称 险类名称," & vbNewLine & _
            "       Nvl(A.费别, B.费别) As 费别, A.国籍, Nvl(B.区域, A.区域) 区域, A.籍贯, A.民族, A.学历, A.婚姻状况, A.职业, A.身份, A.身份证号,A.手机号, A.其他证件," & vbNewLine & _
            "       A.出生日期, A.出生地点, A.家庭地址, A.家庭电话, A.家庭地址邮编, A.户口地址, A.户口地址邮编, A.联系人关系, A.联系人姓名, A.联系人地址,A.联系人身份证号," & vbNewLine & _
            "       A.联系人电话, A.工作单位, A.合同单位id, A.单位电话, A.单位邮编, A.单位开户行, A.单位帐号, A.出院时间,Nvl(A.医疗付款方式, B.医疗付款方式) As 医疗付款方式," & vbNewLine & _
            "       A.当前科室id, A.医保号, Nvl(B.险类, A.险类) As 险类, Nvl(B.病人性质, 0) 病人性质,zl_PatiType(A.病人ID) 病人类型,A.主页ID 就诊次数" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 保险类别 C" & vbNewLine & _
            "Where A.停用时间 Is Null And A.险类 = C.序号(+) And A.病人id = B.病人id(+) And A.主页id = B.主页id(+) And Not Exists" & vbNewLine & _
            " (Select 1 From 病案主页 Z Where Z.病人id = A.病人id And Z.主页id = 0)" & strSQL & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select A.病人id, B.主页id, A.住院次数, A.就诊卡号, A.卡验证码, A.门诊号,B.留观号," & strPati & "NVL(B.姓名,A.姓名) 姓名, NVL(B.性别,A.性别) 性别, NVL(B.年龄,A.年龄) 年龄, C.名称 险类名称," & vbNewLine & _
            "       Nvl(A.费别, B.费别) As 费别, A.国籍, Nvl(B.区域, A.区域) 区域, A.籍贯, A.民族, A.学历, A.婚姻状况, A.职业, A.身份, A.身份证号,A.手机号, A.其他证件," & vbNewLine & _
            "       A.出生日期, A.出生地点, A.家庭地址, A.家庭电话, A.家庭地址邮编, A.户口地址, A.户口地址邮编, A.联系人关系, A.联系人姓名, A.联系人地址,A.联系人身份证号," & vbNewLine & _
            "       A.联系人电话, A.工作单位, A.合同单位id, A.单位电话, A.单位邮编, A.单位开户行, A.单位帐号, A.出院时间,Nvl(A.医疗付款方式, B.医疗付款方式) As 医疗付款方式," & vbNewLine & _
            "       A.当前科室id, A.医保号, Nvl(B.险类, A.险类) As 险类, Nvl(B.病人性质, 0) 病人性质,zl_PatiType(A.病人ID) 病人类型,A.主页ID 就诊次数" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 保险类别 C" & vbNewLine & _
            "Where A.停用时间 Is Null And A.险类 = C.序号(+) And A.病人id = B.病人id And B.主页id = 0" & strSQL
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then
        Set mrsInfo = Nothing: Exit Function
    End If
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = Nothing
End Function


Private Function GetMaxMinPage(lng病人ID As Long, Optional blnMin As Boolean) As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select " & IIf(blnMin, "min", "max") & "(b.主页id) 主页id," & IIf(blnMin, "min", "max") & "(a.主页ID) 住院次数 From 病人信息 A,病案主页 B Where A.病人ID = B.病人ID And A.病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    If IsNull(rsTmp!主页ID) And IsNull(rsTmp!住院次数) Then
        GetMaxMinPage = -1
    Else
        GetMaxMinPage = IIf(IsNull(rsTmp!主页ID) Or Nvl("" & rsTmp!主页ID = 0), Val("" & rsTmp!住院次数), Val("" & rsTmp!主页ID))
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetMaxInHosTimes(lng病人ID As Long) As Long
'功能:获取病人最大住院次数
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    strSQL = "Select NVL(Max(住院次数),0) 住院次数 From 病人信息 where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    GetMaxInHosTimes = Val(rsTmp!住院次数)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FillPatient()
'功能：新增功能时,根据mrsInfo中的病人信息填写病人信息卡片
    txtPatient.Text = mrsInfo!病人ID: txtPatient.Tag = mrsInfo!病人ID
    
    If mbytKind = E门诊留观登记 Then
        If IsNull(mrsInfo!门诊号) Then
            txt住院号.Text = zlDatabase.GetNextNo(3)  '住院新住院号模式下,预约时不产生,接收时产生
            mblnAuto = True
        Else
            txt住院号.Text = mrsInfo!门诊号
            txt住院号.Locked = True
        End If
    ElseIf mbytKind = E住院留观登记 Then
         txt住院号.Text = zlDatabase.GetNextNo(6) '留观登记始终产生新的留观号
    Else
        If IsNull(mrsInfo!住院号) Or gbln每次住院新住院号 And mbytMode <> EMode.E预约登记 Then
            If txt住院号.Visible And mbytKind = EKind.E住院入院登记 Then txt住院号.Text = zlDatabase.GetNextNo(2)  '住院新住院号模式下,预约时不产生,接收时产生
        Else
            txt住院号.Text = mrsInfo!住院号
        End If
    End If
    txt医保号.Text = Nvl(mrsInfo!医保号)
    txt医保号.Locked = Not IsNull(mrsInfo!险类)
    txt险类.Text = "" & mrsInfo!险类名称
    
    txt姓名.Text = mrsInfo!姓名
    
    If IsNull(mrsInfo!主页ID) And IsNull(mrsInfo!就诊次数) Then
        txtPages.Text = "1"
    Else
        If mbytMode = EMode.E接收预约 Or (mbytMode = EMode.E正常登记 And mlng病人ID <> 0) Then
            txtPages.Text = GetMaxMinPage(mrsInfo!病人ID) + 1
        Else
            txtPages.Text = Val(IIf(IsNull(mrsInfo!主页ID) Or Val("" & mrsInfo!主页ID) = 0, Val("" & mrsInfo!就诊次数), Val("" & mrsInfo!主页ID))) + 1
        End If
    End If
    If mbytInState = E新增 And mbytKind = E住院入院登记 And mbytMode <> EMode.E预约登记 Then
        txtTimes.Text = GetMaxInHosTimes(mrsInfo!病人ID) + 1
    Else
        txtTimes.Text = Nvl(mrsInfo!住院次数)
    End If
    txtTimes.Tag = txtTimes.Text
    '65973:刘鹏飞,2013-09-29,添加医疗付款方式
    cbo性别.ListIndex = GetCboIndex(cbo性别, IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别), mstrYBPati <> "")
    cbo费别.ListIndex = GetCboIndex(cbo费别, IIf(IsNull(mrsInfo!费别), "", mrsInfo!费别), mstrYBPati <> "")
    cbo医疗付款.ListIndex = GetCboIndex(cbo医疗付款, IIf(IsNull(mrsInfo!医疗付款方式), "", mrsInfo!医疗付款方式), mstrYBPati <> "")
    cbo国籍.ListIndex = GetCboIndex(cbo国籍, IIf(IsNull(mrsInfo!国籍), "", mrsInfo!国籍), mstrYBPati <> "")
    cbo民族.ListIndex = GetCboIndex(cbo民族, IIf(IsNull(mrsInfo!民族), "", mrsInfo!民族), mstrYBPati <> "")
    cbo学历.ListIndex = GetCboIndex(cbo学历, IIf(IsNull(mrsInfo!学历), "", mrsInfo!学历), mstrYBPati <> "")
    cbo婚姻状况.ListIndex = GetCboIndex(cbo婚姻状况, IIf(IsNull(mrsInfo!婚姻状况), "", mrsInfo!婚姻状况), mstrYBPati <> "")
    cbo职业.ListIndex = GetCboIndex(cbo职业, IIf(IsNull(mrsInfo!职业), "", mrsInfo!职业), mstrYBPati <> "")
    cbo身份.ListIndex = GetCboIndex(cbo身份, IIf(IsNull(mrsInfo!身份), "", mrsInfo!身份), mstrYBPati <> "")
    cbo联系人关系.ListIndex = GetCboIndex(cbo联系人关系, IIf(IsNull(mrsInfo!联系人关系), "", mrsInfo!联系人关系), mstrYBPati <> "")
    If mstrYBPati <> "" Then cbo病人类型.ListIndex = GetCboIndex(cbo病人类型, Nvl(mrsInfo!病人类型, "普通病人"), True)  '医保验证才用
    '问题27676 by lesfeng 2010-01-26 处理性别、费别、国籍、民族、学历、婚姻状况、职业、身份
    If cbo性别.ListIndex = -1 Then Call SetCboDefault(cbo性别)
    If cbo费别.ListIndex = -1 Then Call SetCboDefault(cbo费别)
    If cbo医疗付款.ListIndex = -1 Then Call SetCboDefault(cbo医疗付款)
    If cbo国籍.ListIndex = -1 Then Call SetCboDefault(cbo国籍)
    If cbo民族.ListIndex = -1 Then Call SetCboDefault(cbo民族)
    If cbo学历.ListIndex = -1 Then Call SetCboDefault(cbo学历)
    If cbo婚姻状况.ListIndex = -1 Then Call SetCboDefault(cbo婚姻状况)
    If cbo职业.ListIndex = -1 Then Call SetCboDefault(cbo职业)
    If cbo身份.ListIndex = -1 Then Call SetCboDefault(cbo身份)
    
    Call LoadOldData("" & mrsInfo!年龄, txt年龄, cbo年龄单位)
    mblnChange = False
    txt出生日期.Text = Format(IIf(IsNull(mrsInfo!出生日期), "____-__-__", mrsInfo!出生日期), "YYYY-MM-DD")
    mblnChange = True
    
    If Not IsNull(mrsInfo!出生日期) Then
        If mbytInState <> 2 And mbytInState <> 1 Then txt年龄.Text = ReCalcOld(CDate(Format(mrsInfo!出生日期, "YYYY-MM-DD HH:MM:SS")), cbo年龄单位, Val(mrsInfo!病人ID), , CDate(txt入院时间.Text)) '根据出生日期重算年龄
        If CDate(txt出生日期.Text) - CDate(mrsInfo!出生日期) <> 0 Then
            mblnChange = False
            txt出生时间.Text = Format(mrsInfo!出生日期, "HH:MM")
            mblnChange = True
        End If
    Else
        txt出生时间.Text = "__:__"
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
    
    cbo年龄单位.Tag = txt年龄.Text & "_" & cbo年龄单位.Text
    
    mblnChange = False
    txt身份证号.Text = "" & mrsInfo!身份证号
    mblnChange = True
    txt其他证件.Text = "" & mrsInfo!其他证件
    txt区域.Text = Nvl(mrsInfo!区域)
    txt家庭电话.Text = IIf(IsNull(mrsInfo!家庭电话), "", mrsInfo!家庭电话)
    txt家庭地址邮编.Text = IIf(IsNull(mrsInfo!家庭地址邮编), "", mrsInfo!家庭地址邮编)
    txt户口地址邮编.Text = IIf(IsNull(mrsInfo!户口地址邮编), "", mrsInfo!户口地址邮编)
    txt联系人姓名.Text = IIf(IsNull(mrsInfo!联系人姓名), "", mrsInfo!联系人姓名)
    txt联系人电话.Text = IIf(IsNull(mrsInfo!联系人电话), "", mrsInfo!联系人电话)
    txt联系人身份证号.Text = IIf(IsNull(mrsInfo!联系人身份证号), "", mrsInfo!联系人身份证号)
    txt工作单位.Text = IIf(IsNull(mrsInfo!工作单位), "", mrsInfo!工作单位)
    txt工作单位.Tag = IIf(IsNull(mrsInfo!合同单位ID), "", mrsInfo!合同单位ID)
    txt单位电话.Text = IIf(IsNull(mrsInfo!单位电话), "", mrsInfo!单位电话)
    txt单位邮编.Text = IIf(IsNull(mrsInfo!单位邮编), "", mrsInfo!单位邮编)
    txt单位开户行.Text = IIf(IsNull(mrsInfo!单位开户行), "", mrsInfo!单位开户行)
    txt单位帐号.Text = IIf(IsNull(mrsInfo!单位帐号), "", mrsInfo!单位帐号)
    txtMobile.Text = "" & mrsInfo!手机号
    
    If gbln启用结构化地址 Then
        Call ReadStructAddress(CLng(Nvl(mrsInfo!病人ID, 0)), CLng(Nvl(mrsInfo!主页ID, 0)), PatiAddress)
        txt出生地点.Text = PatiAddress(E_IX_出生地点).Value
        txt籍贯.Text = PatiAddress(E_IX_籍贯).Value
        txt家庭地址.Text = PatiAddress(E_IX_现住址).Value
        txt户口地址.Text = PatiAddress(E_IX_户口地址).Value
        txt联系人地址.Text = PatiAddress(E_IX_联系人地址).Value
    Else
        txt出生地点.Text = IIf(IsNull(mrsInfo!出生地点), "", mrsInfo!出生地点)
        txt籍贯.Text = Nvl(mrsInfo!籍贯)
        txt家庭地址.Text = IIf(IsNull(mrsInfo!家庭地址), "", mrsInfo!家庭地址)
        txt户口地址.Text = IIf(IsNull(mrsInfo!户口地址), "", mrsInfo!户口地址)
        txt联系人地址.Text = IIf(IsNull(mrsInfo!联系人地址), "", mrsInfo!联系人地址)
    End If
    '问题号:56599
    Call Load健康卡相关信息(Val(Nvl(mrsInfo!病人ID, 0)))
End Sub

Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        '医保改动
        If txt姓名.Text = "" And cmdYB.Enabled And cmdYB.Visible Then
            Call cmdYB_Click
            Call EMPI_LoadPati
            Call ReLoadCardFee
            Exit Sub
        End If
        
        
        If mbytInState = 0 Then
            If txt姓名.Text = "" Then
                If Not mrsInfo Is Nothing Then
                    txt姓名.Text = mrsInfo!姓名 '对于人为的清除,又不修改,则自动恢复
                    Call zlCommFun.PressKey(vbKeyTab)
                Else
                    MsgBox "必须输入病人姓名！", vbInformation, gstrSysName
                    txt姓名.SetFocus
                End If
            Else
                If Not mrsInfo Is Nothing Then
                    Call zlCommFun.PressKey(vbKeyTab) '修改姓名
                Else
                    If txtPatient.Tag = "" And InStr(mstrPrivs, "允许非医保病人") > 0 Then '如果尚未产生
                        txtPatient.Text = zlDatabase.GetNextNo(1) '新病人ID
                        txtPatient.Tag = txtPatient.Text
                        '93974
                        txt入院时间.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                        If txt住院号.Text = "" And txt住院号.Visible Then
                            If mbytKind = E住院入院登记 Then
                                txt住院号.Text = zlDatabase.GetNextNo(2)
                            ElseIf mbytKind = E住院留观登记 Then
                                txt住院号.Text = zlDatabase.GetNextNo(6)
                            ElseIf mbytKind = E门诊留观登记 Then
                                txt住院号.Text = zlDatabase.GetNextNo(3)
                            End If
                        End If
                    End If
                    Call EMPI_LoadPati(1)  '新登记
                    Call ReLoadCardFee(True)
                    Call zlCommFun.PressKey(vbKeyTab)
                End If
            End If
        Else
            If txt姓名.Text = "" Then
                MsgBox "必须输入病人姓名！", vbInformation, gstrSysName
                txt姓名.SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    Else
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            Call CheckInputLen(txt姓名, KeyAscii)
        End If
    End If
End Sub

Private Sub txt帐号_GotFocus()
    If IsNumeric(txt预交额.Text) And txt帐号.Text = "" Then
        txt帐号.Text = txt单位帐号.Text
    End If
    zlControl.TxtSelAll txt帐号
End Sub

Private Sub txt帐号_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckInputLen txt帐号, KeyAscii
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If mbytInState = 0 Then
            If txt住院号.Text = "" Then
                If mbytKind = E门诊留观登记 Then
                    If Not mrsInfo Is Nothing Then
                        txt住院号.Text = zlDatabase.GetNextNo(3)
                        mblnAuto = True
                        txt姓名.SetFocus
                    ElseIf Not txt住院号.Locked Then
                        MsgBox "必须输入病人门诊号！", vbInformation, gstrSysName
                        txt住院号.SetFocus
                    Else
                        Call zlCommFun.PressKey(vbKeyTab)
                    End If
                ElseIf mbytKind = E住院留观登记 Then
                    txt住院号.Text = zlDatabase.GetNextNo(6)
                    txt姓名.SetFocus
                Else
                    If Not mrsInfo Is Nothing Then
                        If Nvl(mrsInfo!住院号, 0) = 0 Then '对于人为的清除,又不修改,则自动恢复,(医保验证后，没有住院号,需要重新生成)
                            txt住院号.Text = zlDatabase.GetNextNo(2)
                        Else
                            txt住院号.Text = mrsInfo!住院号
                        End If
                        txt姓名.SetFocus
                    ElseIf Not txt住院号.Locked Then
                        MsgBox "必须输入病人住院号！", vbInformation, gstrSysName
                        txt住院号.SetFocus
                    Else
                        Call zlCommFun.PressKey(vbKeyTab)
                    End If
                End If
            Else
                If Not mrsInfo Is Nothing Then
                    txt姓名.SetFocus  '修改住院号
                Else
                    If txtPatient.Tag = "" And InStr(mstrPrivs, "允许非医保病人") > 0 Then '如果尚未产生
                        txtPatient.Text = zlDatabase.GetNextNo(1) '新病人ID
                        txtPatient.Tag = txtPatient.Text
                    End If
                    txt姓名.SetFocus
                End If
                Call txt住院号_Validate(False)
            End If
        Else
            If txt住院号.Text = "" Then
                If mbytKind = E住院入院登记 And Not txt住院号.Locked Then
                    MsgBox "必须输入病人住院号！", vbInformation, gstrSysName
                    txt住院号.SetFocus
                ElseIf mbytKind = E住院留观登记 Then
                    txt住院号.Text = zlDatabase.GetNextNo(6)
                    txt姓名.SetFocus
                ElseIf mbytKind = E门诊留观登记 Then
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
'功能：根据本地设置光标要定位的输入项目
    Dim i As Integer, j As Integer
    Dim strPara As String
    Dim arrTmp As Variant
    Dim arrSubTmp As Variant
    Dim strInputItem As String
    Dim objTmp As Object
    Dim intBegin As Integer
    Dim intEnd As Integer
    Dim strItem As String
    
    '参数值:输入项目,禁止录入(0\1),必须输入(0\1),光标进入(0\1) '国籍,0,1,1|民族,0,1,1|...
    strPara = zlDatabase.GetPara("输入项控制", glngSys, mlngModul)
    Set mrsInputSet = Rec.CopyNew(Nothing, , , Array("输入项目", adVarChar, 50, Empty, "禁止录入", adInteger, 1, Empty, "必须输入", adInteger, 1, Empty, "光标进入", adInteger, 1, Empty, "控件名", adVarChar, 50, Empty, "控件下标", adInteger, 2, Empty))
    '
    '1)将需要参数控制的控件名用记录集记录下来
    arrTmp = Split(C_输入项控制, "|")
    For i = LBound(arrTmp) To UBound(arrTmp)
        arrSubTmp = Split(arrTmp(i), ",")
        strInputItem = arrSubTmp(0)      '输入项目
        '一个输入项目可能控制多个控件:例如 出生日期 会控制 txt出生日期,txt出生时间
        For j = LBound(arrSubTmp) + 1 To UBound(arrSubTmp)
            mrsInputSet.AddNew
            mrsInputSet!输入项目 = strInputItem
            strItem = arrSubTmp(j)
            intBegin = InStr(strItem, "(")
            If intBegin > 0 Then
                intEnd = InStr(strItem, ")")
                mrsInputSet!控件名 = Mid(strItem, 1, intBegin - 1)
                mrsInputSet!控件下标 = Val(Mid(strItem, intBegin + 1, intEnd - intBegin + 1))
            Else
                mrsInputSet!控件名 = strItem
            End If
            mrsInputSet!光标进入 = 1 '缺省设置为1-光标进入
            mrsInputSet.Update
        Next
    Next
    
    If strPara <> "" Then
        '2）将参数设置的值更新到记录集上
        arrTmp = Split(strPara, "|")
        For i = LBound(arrTmp) To UBound(arrTmp)
            arrSubTmp = Split(arrTmp(i), ",")
            mrsInputSet.Filter = "输入项目 ='" & arrSubTmp(0) & "'"
            For j = 1 To mrsInputSet.RecordCount
                mrsInputSet!禁止录入 = Val(arrSubTmp(1))
                mrsInputSet!必须输入 = Val(arrSubTmp(2))
                mrsInputSet!光标进入 = Val(arrSubTmp(3))
                mrsInputSet.Update
                mrsInputSet.MoveNext
            Next
        Next
    End If
    mrsInputSet.Filter = "" '
  
    For i = 1 To mrsInputSet.RecordCount
        Set objTmp = CallByName(Me, mrsInputSet!控件名 & "", VbGet)
        If Not IsNull(mrsInputSet!控件下标) Then
            Set objTmp = objTmp(mrsInputSet!控件下标)
        End If
        '禁止录入
        objTmp.Enabled = Val(mrsInputSet!禁止录入 & "") = 0
        objTmp.BackColor = IIf(objTmp.Enabled, C_COLOR_Enabled, C_COLOR_UNEnabled)
        '光标是否进入
        objTmp.TabStop = Val(mrsInputSet!光标进入 & "")
        mrsInputSet.MoveNext
    Next
End Sub

Private Sub txt出生日期_LostFocus()
    If txt出生日期.Text <> "____-__-__" And Not IsDate(txt出生日期.Text) Then
        txt出生日期.SetFocus
    End If
End Sub
'问题26779 by lesfeng 2009-12-10
Private Sub LoadBedInfo(lng科室ID As Long, Optional lng病区ID As Long)
'功能：科室，病区显示病人在院人数，床位数
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer, strTmp As String
    Dim strSqlTmp1 As String, strSqlTmp2 As String
    Dim strSqlTmp3 As String, strSqlTmp4 As String
    
    Dim intFlag As Integer
    
    On Error GoTo errHandle
    intFlag = 0
    If lng科室ID = lng病区ID Then
        intFlag = 1
    Else
        intFlag = 2
    End If
    strSqlTmp1 = " And B.病区ID=[2]"
    strSqlTmp2 = " And A.病区id = [2]"
    strSqlTmp3 = " And B.科室ID=[1]"
    strSqlTmp4 = " And A.科室id = [1]"
    strSQL = " Select Sum(病区在院) as 病区在院, Sum(病区空床) as 病区空床,Sum(科室在院) As 科室在院, Sum(科室空床) As 科室空床" & _
             " From ( Select Count(A.病人id) As 病区在院, 0 As 病区空床,0 As 科室在院,0 As 科室空床" & _
             "          From 病人信息 A,在院病人 B" & _
             "         Where A.病人ID=B.病人ID " & strSqlTmp1 & _
             "         Union All " & _
             "        Select 0 As 病区在院, Count(A.床号) As 病区空床 ,0 As 科室在院,0 As 科室空床" & _
             "          From 床位状况记录 A" & _
             "        Where A.床位编制 <> '非编' and A.床位编制 <> '监护' And A.状态 = '空床'" & strSqlTmp2 & _
             "         Union All " & _
             "       Select 0 As 病区在院, 0 As 病区空床,Count(A.病人id)  As 科室在院,0 As 科室空床" & _
             "          From 病人信息 A,在院病人 B" & _
             "         Where A.病人ID =B.病人ID " & strSqlTmp3 & _
             "         Union All " & _
             "        Select 0 As 病区在院, 0 As 病区空床 ,0 As 科室在院,Count(A.床号) As 科室空床" & _
             "          From 床位状况记录 A" & _
             "        Where A.床位编制 <> '非编' and A.床位编制 <> '监护' And A.状态 = '空床'" & strSqlTmp4 & ") "
   
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng科室ID, lng病区ID)
    '问题 27097 by lesfeng 2009-12-25 考虑不确定科室或者病区，应该排除这种情况。
    If Not rsTemp.EOF Then
        If intFlag = 1 Then
            If gbln先选病区 Then
                If InStr(1, cbo入院科室.Text, "-") > 0 Then
                    strTmp = Split(cbo入院科室.Text, "-")(1)
                Else
                    strTmp = Trim(cbo入院科室.Text)
                End If
            Else
                If InStr(1, cbo入院病区.Text, "-") > 0 Then
                    strTmp = Split(cbo入院病区.Text, "-")(1)
                Else
                    strTmp = Trim(cbo入院病区.Text)
                End If
            End If
            strTmp = strTmp & "：在院人数 " & IIf(IsNull(rsTemp!科室在院), 0, rsTemp!科室在院) & "，床位数 " & IIf(IsNull(rsTemp!科室空床), 0, rsTemp!科室空床)
        Else
            If gbln先选病区 Then
                If InStr(1, cbo入院病区.Text, "-") > 0 Then
                    strTmp = Split(cbo入院病区.Text, "-")(1)
                Else
                    strTmp = Trim(cbo入院病区.Text)
                End If
                strTmp = strTmp & "：在院人数 " & IIf(IsNull(rsTemp!病区在院), 0, rsTemp!病区在院) & "，床位数 " & IIf(IsNull(rsTemp!病区空床), 0, rsTemp!病区空床)
                
                If lng科室ID <> 0 Then
                    If InStr(1, cbo入院科室.Text, "-") > 0 Then
                        strTmp = strTmp & "," & Split(cbo入院科室.Text, "-")(1)
                    Else
                        strTmp = strTmp & "," & Trim(cbo入院科室.Text)
                    End If
                    strTmp = strTmp & "：在院人数 " & IIf(IsNull(rsTemp!科室在院), 0, rsTemp!科室在院) & "，床位数 " & IIf(IsNull(rsTemp!科室空床), 0, rsTemp!科室空床)
                End If
            Else
                If InStr(1, cbo入院科室.Text, "-") > 0 Then
                    strTmp = Split(cbo入院科室.Text, "-")(1)
                Else
                    strTmp = Trim(cbo入院科室.Text)
                End If
                strTmp = strTmp & "：在院人数" & IIf(IsNull(rsTemp!科室在院), 0, rsTemp!科室在院) & "，床位数" & IIf(IsNull(rsTemp!科室空床), 0, rsTemp!科室空床)
                
                If lng病区ID <> 0 Then
                    If InStr(1, cbo入院病区.Text, "-") > 0 Then
                        strTmp = strTmp & "," & Split(cbo入院病区.Text, "-")(1)
                    Else
                        strTmp = strTmp & "," & Trim(cbo入院病区.Text)
                    End If
                    strTmp = strTmp & "：在院人数 " & IIf(IsNull(rsTemp!病区在院), 0, rsTemp!病区在院) & "，床位数 " & IIf(IsNull(rsTemp!病区空床), 0, rsTemp!病区空床)
                End If

            End If
        End If
    Else
        strTmp = ""
    End If

    lblBedInfo.Caption = strTmp
    If gbln控制空床 Then
        If Val(rsTemp!科室空床) = 0 And Val(rsTemp!病区空床) = 0 Then
            mbln空床 = True
                Else
                        mbln空床 = False
        End If
    Else
        mbln空床 = False
    End If
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub LoadBed(str性别 As String, lng科室ID As Long, Optional lng病区ID As Long)
'功能：根据当前病人性别，科室，病区加载可用的病床
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strTmp As String, strPreBed As String
    Dim blnFind As Boolean
    
    If Not (gbln入院入科 And mbytMode <> EMode.E预约登记 And mbytInState = EState.E新增) Then Exit Sub
    
    If cbo床位.ListCount > 1 And InStr(Trim(cbo床位.Text), " ") > 1 Then strPreBed = Trim(Mid(Trim(cbo床位.Text), 1, InStr(Trim(cbo床位.Text), " ") - 1))
    cbo床位.Clear: cbo床位.Tag = ""
    cbo床位.AddItem "不分配床位"
    If lng病区ID <> 0 Then
        cbo床位.AddItem "分配家庭病床"
    End If
    cbo床位.ListIndex = 0
        
    '床位要即时取，不使用缓存
    Set rsTmp = GetFreeBeds(lng病区ID, lng科室ID, str性别)
    For i = 1 To rsTmp.RecordCount
        cbo床位.AddItem " " & rsTmp!床号 & Space(10 - Len(rsTmp!床号)) & " " & rsTmp!性别分类 & IIf(IsNull(rsTmp!房间号), "", " 房间:" & rsTmp!房间号 & "|") & _
            IIf(IsNull(rsTmp!房间号) Or ((Not IsNull(rsTmp!房间号)) And Trim(Nvl(rsTmp!性别) = "")), "", "(" & Nvl(rsTmp!性别) & ")") & _
            Space(15 - Len(IIf(IsNull(rsTmp!房间号), "", " 房间:" & IsNull(rsTmp!房间号))) - Len(IIf(IsNull(rsTmp!房间号) Or ((Not IsNull(rsTmp!房间号)) And Trim(Nvl(rsTmp!性别) = "")), "", "(" & Nvl(rsTmp!性别) & ")"))) & _
            Nvl(rsTmp!床位等级)
        If rsTmp!床号 = strPreBed And Not blnFind Then cbo床位.ListIndex = cbo床位.NewIndex: cbo床位.Tag = rsTmp!床号
        If mblnAppoint And rsTmp!床号 = mstrAppointBed Then
            cbo床位.ListIndex = cbo床位.NewIndex: blnFind = True
            cbo床位.Tag = rsTmp!床号
        End If
        rsTmp.MoveNext
    Next
End Sub

Private Sub txt担保额_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc(".") And InStr(txt担保额.Text, ".") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Function Check担保信息() As Boolean
    Check担保信息 = True
    
    If txt担保人.Tag <> "" Then
    '修改时不能删除,要删除就到病人信息管理中去删除
        If Trim(txt担保人.Text) = "" Then
            MsgBox "修改登记信息时不允许删除已经存在的担保信息!", vbInformation, gstrSysName
            If txt担保人.Enabled Then txt担保人.SetFocus
            Check担保信息 = False
            Exit Function
        End If
    End If
    
    If Not IsNumeric(txt担保额.Text) And Trim(txt担保额.Text) <> "" Then
        MsgBox "请输入正确的担保额,担保额要求是数值!", vbInformation, gstrSysName
        If txt担保额.Enabled Then txt担保额.SetFocus
        Check担保信息 = False
        Exit Function
    End If
    
    If IsNumeric(txt担保额.Text) And Trim(txt担保人.Text) = "" Then
        MsgBox "请输入担保人姓名,担保人不能为空!", vbInformation, gstrSysName
        If txt担保人.Enabled Then txt担保人.SetFocus
        Check担保信息 = False
        Exit Function
    End If
    
    
    '只要输入担保人,或选择了担保时限,或选择了临时担保,就表示要录入担保信息
    If Trim(txt担保人.Text) <> "" Or Not IsNull(dtp担保时限.Value) Or chk临时担保.Value = 1 Then
        If Val(txt担保额.Text) = 0 Then
            MsgBox "请输入担保额,担保额不能为零!", vbInformation, gstrSysName
            If txt担保额.Enabled Then txt担保额.SetFocus
            Check担保信息 = False
            Exit Function
        End If
    End If
    
    '担保时限不能小于入院时间
    If Not IsNull(dtp担保时限.Value) And dtp担保时限.Enabled Then
        If dtp担保时限.Value < CDate(txt入院时间.Text) Then
            MsgBox "担保到期时间不允许设置为入院时间之前!!", vbInformation, gstrSysName
            If dtp担保时限.Enabled Then dtp担保时限.SetFocus
            Check担保信息 = False
            Exit Function
        End If
    End If
    
    If chk临时担保.Value = 1 Then
        If Not IsNull(dtp担保时限.Value) Or chkUnlimit.Value = 1 Then
            MsgBox "临时担保不允许设置担保时限或不限担保额!", vbInformation, gstrSysName
            If chk临时担保.Enabled Then chk临时担保.SetFocus
            Check担保信息 = False
            Exit Function
        End If
    End If
    
    If zlCommFun.ActualLen(Trim(txtReason.Text)) > 50 Then
        MsgBox "担保原因过长，最多允许 25 个汉字或 50 个字符。", vbInformation, gstrSysName
        txtReason.SetFocus
        Check担保信息 = False
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
    Dim lng接口编号 As Long, strBalanceInfor As String
    Dim i As Long, lng病人ID As Long, blnErr As Boolean
    Dim lngTmp As Long, strno As String
    Dim blnTmp As Boolean   '是否因为门诊号被占用而重新生成
    Dim bln基本信息调整, blnMod   As Boolean
    Dim str出生日期 As String, str年龄 As String, strAge As String, str性别 As String, strErrInfo As String
    Dim strMsg As String
    Dim objTmp As Object
    
    '问题号:56599
    tbcPage.Item(0).Selected = True
     
    '65965:刘鹏飞,2013-09-24,处理预交显示千位位格式
    If Not CheckFormInput(Me, "txt门诊诊断,txt中医诊断", "txt预交额") Then Exit Sub
    
    '合法性检查
    '问题号:53408
    If IIf(zlDatabase.GetPara("扫描身份证签约", glngSys, glngModul) = "1", 1, 0) = 0 And ((mCurSendCard.str卡名称 = "二代身份证" And Trim(txt卡号.Text) <> "") Or Trim(txt支付密码.Text) <> "") Then
         MsgBox "您没有权限进行签约操作,请到参数设置中设置【扫描身份证签约】！", vbOKOnly + vbInformation, gstrSysName
         txt卡号.Text = ""
         txtPass.Text = ""
         txtAudi.Text = ""
         If txt卡号.Visible = True Then txt卡号.SetFocus
         Exit Sub
    End If
    
    If ((Not IsNumeric(txt户口地址邮编.Text)) Or Len(txt户口地址邮编.Text) > 6 Or InStr(txt户口地址邮编.Text, ".") > 0) And txt户口地址邮编.Text <> "" Then
        MsgBox "邮编格式错误,请输入正确的邮编!" & vbCrLf & "【正确邮编格式为六位纯数字编码】", vbInformation, gstrSysName
        If CanFocus(txt户口地址邮编) = True Then txt户口地址邮编.SetFocus: Exit Sub
    End If
    If ((Not IsNumeric(txt单位邮编.Text)) Or Len(txt单位邮编.Text) > 6 Or InStr(txt单位邮编.Text, ".") > 0) And txt单位邮编.Text <> "" Then
        MsgBox "邮编格式错误,请输入正确的邮编!" & vbCrLf & "【正确邮编格式为六位纯数字编码】", vbInformation, gstrSysName
        If CanFocus(txt单位邮编) = True Then txt单位邮编.SetFocus: Exit Sub
    End If
    If ((Not IsNumeric(txt家庭地址邮编.Text)) Or Len(txt家庭地址邮编.Text) > 6 Or InStr(txt家庭地址邮编.Text, ".") > 0) And txt家庭地址邮编.Text <> "" Then
        MsgBox "邮编格式错误,请输入正确的邮编!" & vbCrLf & "【正确邮编格式为六位纯数字编码】", vbInformation, gstrSysName
        If CanFocus(txt家庭地址邮编) = True Then txt家庭地址邮编.SetFocus: Exit Sub
    End If
    
    If Trim(txt支付密码.Text) <> "" And Trim(txt身份证号.Text) <> "" Then
           If 是否已经签约(txt身份证号.Text) Then
                 MsgBox "身份证号码为:" & txt身份证号.Text & "已经签约不能重复签约！", vbOKOnly + vbInformation, gstrSysName
                 txt支付密码.Text = ""
                 If txt支付密码.Visible = True Then txt支付密码.SetFocus
                 Exit Sub
           End If
    End If
    
    If mbln是否扫描身份证 = False And mCurSendCard.str卡名称 = "二代身份证" And txt卡号.Text <> "" Then
            MsgBox "绑定身份证只能以刷卡的方式进行，不允许手动输入身份证进行绑定!", vbOKOnly + vbInformation, gstrSysName
            txt卡号.Text = ""
            txtPass.Text = ""
            txtAudi.Text = ""
            txt支付密码.Text = ""
            txt验证密码.Text = ""
            If txt卡号.Visible = True Then txt卡号.SetFocus
            Exit Sub
    End If
    
    If mbln是否扫描身份证 = False And mCurSendCard.str卡名称 <> "二代身份证" And txt支付密码.Text <> "" Then
            MsgBox "绑定身份证只能以刷卡的方式进行，不允许手动输入身份证进行绑定!", vbOKOnly + vbInformation, gstrSysName
            txt身份证号.Text = ""
            txt支付密码.Text = ""
            txt验证密码.Text = ""
            If txt身份证号.Visible = True Then
                If CanFocus(txt身份证号) = True Then txt身份证号.SetFocus
            End If
        Exit Sub
    End If
    
    If Trim(txt支付密码.Text) <> Trim(txt验证密码.Text) And (Trim(txt支付密码.Text) <> "" Or Trim(txt验证密码.Text) <> "") Then
        MsgBox "两次输入的密码不一致,请重新输入", vbOKOnly + vbInformation, gstrSysName
        txt支付密码.Text = "": txt验证密码.Text = ""
        If txt支付密码.Visible = True Then txt支付密码.SetFocus
        Exit Sub
    End If
    
    
    If txtPatient.Tag = "" Then
        MsgBox "必须确定入院病人！", vbInformation, gstrSysName
        If Not txtPatient.TabStop Then
            txt姓名.SetFocus
        Else
            txtPatient.SetFocus
        End If
        Exit Sub
    End If
    If Trim(txt住院号.Text) = "" And mbytKind = E住院入院登记 And mbytMode <> 1 Then  '住院留观新病人,门诊留观没有住院号
        MsgBox "必须输入病人住院号！", vbInformation, gstrSysName
        txt住院号.SetFocus: Exit Sub
    End If
    
    If txtTimes.Visible And txtTimes.Enabled Then
        If Not IsNumeric(txtTimes.Text) Then
            MsgBox "住院次数必须是数字！", vbInformation, gstrSysName
            txtTimes.SetFocus: Exit Sub
        End If
        If Val(txtTimes.Text) < Val(txtTimes.Tag) Then
            MsgBox "住院次数不能小于已存在的次数！", vbInformation, gstrSysName
            txtTimes.SetFocus: Exit Sub
        End If
        If Val(txtTimes.Text) = 0 And mbytMode <> EMode.E预约登记 And mbytKind = E住院入院登记 Then
            MsgBox "住院次数不能为零！", vbInformation, gstrSysName
            txtTimes.SetFocus: Exit Sub
        End If
    End If
    
    If Trim(txt姓名.Text) = "" Then
        MsgBox "必须输入病人的姓名！", vbInformation, gstrSysName
        If CanFocus(txt姓名) = True Then txt姓名.SetFocus: Exit Sub
    End If
    If cbo性别.ListIndex = -1 Then
        MsgBox "必须确定病人的性别！", vbInformation, gstrSysName
        If CanFocus(cbo性别) = True Then cbo性别.SetFocus: Exit Sub
    End If
    If txt出生日期.Enabled Then
        If Not IsDate(txt出生日期.Text) Then
            MsgBox "必须正确输入病人的出生日期！", vbInformation, gstrSysName
            If CanFocus(txt出生日期) = True Then txt出生日期.SetFocus: Exit Sub
        End If
    End If
    If Trim(txt年龄.Text) = "" Then
        MsgBox "必须输入病人的年龄！", vbInformation, gstrSysName
        If CanFocus(txt年龄) = True Then txt年龄.SetFocus: Exit Sub
    End If
    
    '80505  参数"输入项控制"指定必须输入的项目检查
    mrsInputSet.Filter = "": blnTmp = False       '
    For i = 1 To mrsInputSet.RecordCount
        '必须输入项目检查
        If Val(mrsInputSet!必须输入 & "") = 1 Then
            Set objTmp = CallByName(Me, mrsInputSet!控件名 & "", VbGet)
            If Not IsNull(mrsInputSet!控件下标) Then
                Set objTmp = objTmp(mrsInputSet!控件下标) '控件数组
            End If
            blnTmp = False
            If objTmp.Enabled = True And objTmp.Visible Then
                If UCase(TypeName(objTmp)) = UCase("TextBox") Then
                    If Trim(objTmp.Text) = "" Then blnTmp = True
                ElseIf UCase(TypeName(objTmp)) = UCase("ComboBox") Then
                    If objTmp.ListIndex = -1 Then blnTmp = True
                ElseIf UCase(TypeName(objTmp)) = UCase("MaskEdBox") Then
                    If mrsInputSet!输入项目 & "" = "出生日期" Then
                        blnTmp = False  '出生日期后续单独检查,此处暂不检查
                    Else
                        If Trim(objTmp.Text) = "" Then blnTmp = True
                    End If
                ElseIf UCase(TypeName(objTmp)) = UCase("PatiAddress") Then
                    If Trim(objTmp.Value) = "" Or objTmp.CheckNullValue() <> "" Then blnTmp = True
                End If
                If blnTmp Then
                    MsgBox "必须输入病人的" & mrsInputSet!输入项目 & "！", vbInformation, gstrSysName
                    If CanFocus(objTmp) = True Then objTmp.SetFocus
                    Exit Sub
                End If
            End If
        Else
            '对于非必须输入的项目结构化地址内容一旦录入一部分就要求必须完整录入。
            Set objTmp = CallByName(Me, mrsInputSet!控件名 & "", VbGet)
            If Not IsNull(mrsInputSet!控件下标) Then
                Set objTmp = objTmp(mrsInputSet!控件下标) '控件数组
            End If
            
            If objTmp.Enabled = True And objTmp.Visible Then
                If UCase(TypeName(objTmp)) = UCase("PatiAddress") Then
                    If Trim(objTmp.Value) <> "" And objTmp.CheckNullValue() <> "" Then
                        MsgBox "病人的" & mrsInputSet!输入项目 & "录入不完整,请重新录入或者删除已录入内容。", vbInformation, gstrSysName
                        If CanFocus(objTmp) = True Then objTmp.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        mrsInputSet.MoveNext
    Next
    
    '76409,刘鹏飞,2014-08-06,年龄合法性检查
    If txt年龄.Locked = False Then
        str年龄 = txt年龄.Text
        If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
        If str年龄 Like "约*" Then
            str年龄 = str年龄 & cbo年龄单位.Text
        End If
        If IsDate(txt出生日期.Text) Then
            If txt出生时间.Text = "__:__" Then
                str出生日期 = Format(txt出生日期.Text, "YYYY-MM-DD")
            Else
                str出生日期 = Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS")
            End If
            strInfo = CheckAge(str年龄, str出生日期, CDate(txt入院时间.Text))
        Else
            strInfo = CheckAge(str年龄)
        End If
        If InStr(1, strInfo, "|") > 0 Then
            lngTmp = Val(Split(strInfo, "|")(0)) '1禁止,0提示
            strInfo = Split(strInfo, "|")(1)
            If lngTmp = 1 Then
                MsgBox strInfo, vbInformation, gstrSysName
                If CanFocus(txt年龄) = True Then txt年龄.SetFocus: Exit Sub
            Else
                If MsgBox(strInfo & vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If CanFocus(txt年龄) = True Then txt年龄.SetFocus: Exit Sub
                End If
            End If
        End If
    End If

    str出生日期 = ""
    '--81012,余伟节,2014-12-22,根据身份证对出生日期\年龄\性别 的检查
    If Trim(zlCommFun.GetNeedName(cbo国籍.Text)) = "中国" Then
        If Not CheckLen(txt身份证号, 18) Then Exit Sub
        lngTmp = LenB(StrConv(Trim(txt身份证号.Text), vbFromUnicode))
        If lngTmp > 0 Then
            If CreatePublicPatient() Then
                If gobjPublicPatient.CheckPatiIdcard(Trim(txt身份证号.Text), str出生日期, strAge, str性别, strErrInfo, CDate(txt入院时间.Text)) Then
                    '有无基本信息调整权限
                    bln基本信息调整 = InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";基本信息调整;") > 0 And mbytInState = 1 And mblnHaveAdvice
                    strMsg = ""
                    '出生日期
                    If Trim(txt出生日期.Text) <> "____-__-__" Then
                        If CDate(Format(str出生日期, "YYYY-MM-DD")) <> CDate(Format(txt出生日期.Text, "YYYY-MM-DD")) Then
                            strMsg = "身份证号码中的出生日期[" & str出生日期 & "]和病人出生日期[" & Format(txt出生日期.Text, "YYYY-MM-DD") & "]不一致"
                            '年龄
                            str年龄 = txt年龄.Text
                            If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
                            If str年龄 <> strAge Then
                                strMsg = strMsg & vbCrLf & "身份证号码中的年龄[" & strAge & "]和病人年龄[" & str年龄 & "]不一致"
                            End If
                        End If
                    End If
                    '性别
                    If InStr(cbo性别.Text, str性别) = 0 Then
                        strMsg = IIf(strMsg <> "", strMsg & vbCrLf, "") & "身份证号码中的性别[" & str性别 & "]和病人性别[" & zlCommFun.GetNeedName(cbo性别.Text) & "]不一致"
                    End If
                    
                    If mbytInState = 1 And mblnHaveAdvice Then
                        If strMsg <> "" Then
                            If MsgBox(strMsg & ",是否继续？" & vbCrLf & IIf(bln基本信息调整, "选【是】,用身份证的信息替换病人的信息及相关业务数据。", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If CanFocus(txt身份证号) = True Then txt身份证号.SetFocus: Exit Sub
                            Else
                                blnMod = True
                            End If
                        End If
                    Else
                        If strMsg <> "" Then
                            If MsgBox(strMsg & ",是否继续？" & vbCrLf & "选【是】,用身份证的信息替换病人的信息。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                If CanFocus(txt身份证号) = True Then txt身份证号.SetFocus: Exit Sub
                            Else
                                If CDate(Format(str出生日期, "YYYY-MM-DD")) <> CDate(Format(txt出生日期.Text, "YYYY-MM-DD")) Then
                                    txt出生日期.Text = str出生日期
                                    If mblnChange = False Then
                                        Call LoadOldData(strAge, txt年龄, cbo年龄单位)
                                    End If
                                End If
                                Call cbo.Locate(cbo性别, str性别, False)
                            End If
                        End If
                    End If
                Else
                    MsgBox strErrInfo, vbInformation + vbOKOnly, gstrSysName
                    If CanFocus(txt身份证号) = True Then txt身份证号.SetFocus: Exit Sub
                End If
                
            End If
        End If
    End If
    
    If cbo费别.ListIndex = -1 Then
        MsgBox "必须确定病人费别！", vbInformation, gstrSysName
        cbo费别.SetFocus: Exit Sub
    End If
    If cbo国籍.ListIndex = -1 Then
        MsgBox "必须确定病人国籍！", vbInformation, gstrSysName
        If CanFocus(cbo国籍) = True Then cbo国籍.SetFocus: Exit Sub
    End If
    If cbo民族.ListIndex = -1 Then
        MsgBox "必须确定病人民族！", vbInformation, gstrSysName
        If CanFocus(cbo民族) = True Then cbo民族.SetFocus: Exit Sub
    End If
    If cbo病人类型.ListIndex = -1 Then
        MsgBox "必须确定病人类型！", vbInformation, gstrSysName
        If cbo病人类型.Enabled Then
            cbo病人类型.SetFocus
        End If
        Exit Sub
    End If
    
    '担保信息检查
    If txt担保额.Visible And txt担保额.Enabled Then
        If Not Check担保信息 Then Exit Sub
    End If
    
    If cbo入院科室.ListIndex = -1 Then
        MsgBox "必须确定病人入院科室！", vbInformation, gstrSysName
        If CanFocus(cbo入院科室) Then cbo入院科室.SetFocus: Exit Sub
    End If
    If cbo入院病区.ListIndex = -1 And cbo入院病区.Visible And gbln先选病区 Then
        MsgBox "必须确定病人入院病区！", vbInformation, gstrSysName
        If CanFocus(cbo入院病区) Then cbo入院病区.SetFocus: Exit Sub
    End If
    
    If mbln空床 Then
        MsgBox zlCommFun.GetNeedName(cbo入院科室.Text) & "已没有空床位，请办理到其他科室或转院治疗", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If cbo护理等级.ListIndex = -1 Then
        MsgBox "必须确定护理等级！", vbInformation, gstrSysName
        cbo护理等级.SetFocus: Exit Sub
    End If
    
    If cbo门诊医师.ListIndex = -1 And mbytMode <> 1 Then
        MsgBox "必须确定门诊医师！", vbInformation, gstrSysName
        cbo门诊医师.SetFocus: Exit Sub
    End If
    
    If cbo入院病况.ListIndex = -1 Then
        MsgBox "必须确定病人入院病况！", vbInformation, gstrSysName
        cbo入院病况.SetFocus: Exit Sub
    End If
    If cbo入院方式.ListIndex = -1 Then
        MsgBox "必须确定病人入院方式！", vbInformation, gstrSysName
        cbo入院方式.SetFocus: Exit Sub
    End If
    '刘兴宏:2007/09/13
    If cbo入院属性.ListIndex = -1 Then
        MsgBox "必须确定病人入院属性！", vbInformation, gstrSysName
        cbo入院属性.SetFocus: Exit Sub
    End If
    
    If cbo住院目的.ListIndex = -1 Then
        MsgBox "必须确定病人住院目的！", vbInformation, gstrSysName
        cbo住院目的.SetFocus: Exit Sub
    End If
    If Not IsDate(txt入院时间.Text) Then
        MsgBox "必须输入正确的病人入院时间！", vbInformation, gstrSysName
        txt入院时间.SetFocus: Exit Sub
    End If
    
     '联系人检查
    If Trim(txt联系人姓名.Text) = "" And (cbo联系人关系.ListIndex >= 0 Or Trim(txt联系人电话.Text) <> "" Or Trim(txt联系人地址.Text) <> "" Or Trim(txt联系人身份证号.Text) <> "") Then
        MsgBox "必须录入联系人姓名!", vbInformation, gstrSysName
        If txt联系人姓名.Enabled And txt联系人姓名.Visible Then txt联系人姓名.SetFocus: Exit Sub
    End If
    
    '78877,84014 出生日期和入院时间前面已经进行空值检查
    If txt出生时间.Text = "__:__" Then
        str出生日期 = Format(txt出生日期.Text, "YYYY-MM-DD")
    Else
        str出生日期 = Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS")
    End If
    
    If txt出生日期.Enabled Then
        If CDate(str出生日期) > CDate(txt入院时间.Text) Then
            MsgBox "病人出生日期[" & str出生日期 & "]必须小于病人入院时间[" & Format(txt入院时间.Text, "YYYY-MM-DD HH:MM") & "]！", vbInformation, gstrSysName
            txt出生日期.SetFocus: Exit Sub
        End If
    End If
    
    '费别适用科室
    If cbo入院科室.ListIndex <> -1 Then
        If Not Check费别适用科室(zlCommFun.GetNeedName(cbo费别.Text), Val(cbo入院科室.ItemData(cbo入院科室.ListIndex))) Then
            MsgBox "当前费别对病人科室不适用,请重新选择费别!", vbInformation, gstrSysName
            cbo费别.SetFocus: Exit Sub
        End If
    End If

    
    '入院时间
    If Not mrsInfo Is Nothing Then
        If CDate(txt入院时间.Text) < IIf(IsNull(mrsInfo!出院时间), #1/1/1900#, mrsInfo!出院时间) Then
            MsgBox "病人入院时间不能小于病人上次出院时间[" & Format(IIf(IsNull(mrsInfo!出院时间), #1/1/1900#, mrsInfo!出院时间), "yyyy-MM-dd") & "]！", vbInformation, gstrSysName
            txt入院时间.SetFocus: Exit Sub
        End If
    ElseIf mbytInState = EState.E修改 And txt入院时间.Tag <> "" Then
        If CDate(txt入院时间.Text) < CDate(txt入院时间.Tag) Then
            MsgBox "病人入院时间不能小于病人上次出院时间[" & Format(txt入院时间.Tag, "yyyy-MM-dd HH:mm:ss") & "]！", vbInformation, gstrSysName
            txt入院时间.SetFocus: Exit Sub
        End If
    End If
        
    '门诊诊断
    If mintInsure <> 0 And mstrYBPati <> "" And mbytMode <> 1 Then
        If gclsInsure.GetCapability(support必须录入入出诊断, Val(txtPatient.Tag), mintInsure) Then
            If txt门诊诊断.Text = "" Then
                MsgBox "请填写该病人的门诊诊断！", vbInformation, gstrSysName
                txt门诊诊断.SetFocus: Exit Sub
            End If
        End If
    ElseIf InStr(mstrPrivs, "允许非医保病人") = 0 Then
        MsgBox "你没有权限对非医保病人进行登记.", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '手机号合法性检查
    If Trim(txtMobile.Text) <> "" Then
        If CheckMobile(Trim(txtMobile.Text), Val(txtPatient.Tag)) Then
            If MsgBox("在已有的病人信息中存在相同的手机号:" & Trim(txtMobile.Text) & vbCrLf & "是否重新录入？", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                If txtMobile.Enabled And txtMobile.Visible Then txtMobile.SetFocus: Exit Sub
            End If
        End If
    End If
    
    '长度检查
    
    If Not CheckTextLength("姓名", txt姓名) Then Exit Sub
    If Not CheckTextLength("年龄", txt年龄) Then Exit Sub
    If Not CheckOldData(txt年龄, cbo年龄单位) Then Exit Sub
    
    '64701:刘鹏飞,2013-10-31,修改出生地址最大可输入100个字符，50个汉字
    If Not gbln启用结构化地址 Then
        If Not CheckLen(txt家庭地址, 100) Then Exit Sub
        If Not CheckLen(txt出生地点, 100) Then Exit Sub
        If Not CheckLen(txt户口地址, 100) Then Exit Sub
        If Not CheckLen(txt联系人地址, 100) Then Exit Sub
    End If
    If Not CheckLen(txt户口地址邮编, 6) Then Exit Sub
    If Not CheckLen(txt家庭地址邮编, 6) Then Exit Sub
    If Not CheckLen(txt家庭电话, 20) Then Exit Sub
    If Not CheckLen(txt联系人姓名, 64) Then Exit Sub
    If Not CheckLen(txt联系人电话, 20) Then Exit Sub
    If Not CheckLen(txt联系人身份证号, 18) Then Exit Sub
    If Not CheckLen(txtLinkManInfo, 100) Then Exit Sub
    If Not CheckLen(txt工作单位, txt工作单位.MaxLength) Then Exit Sub
    If Not CheckLen(txt单位电话, 20) Then Exit Sub
    If Not CheckLen(txtMobile, 20) Then Exit Sub
    If Not CheckLen(txt单位邮编, 6) Then Exit Sub
    If Not CheckLen(txt单位开户行, 50) Then Exit Sub
    If Not CheckLen(txt单位帐号, 50) Then Exit Sub
    If Not CheckLen(txt担保人, 64) Then Exit Sub
    If Not CheckLen(txt门诊诊断, txt门诊诊断.MaxLength) Then Exit Sub
    If Not CheckLen(txt中医诊断, txt中医诊断.MaxLength) Then Exit Sub
    If Not CheckLen(txt卡号, CInt(mCurSendCard.lng卡号长度)) Then Exit Sub
    If Not CheckLen(txtPass, 10) Then Exit Sub
    If Not CheckLen(txt缴款单位, 50) Then Exit Sub
    If Not CheckLen(txt开户行, 50) Then Exit Sub
    If Not CheckLen(txt帐号, 50) Then Exit Sub
    If Not CheckLen(txt结算号码, 30) Then Exit Sub
    If Not CheckLen(txt备注, txt备注.MaxLength) Then Exit Sub
    If zlStr.NeedName(cbo入院方式.Text) = "转入" Then
        If Not zlControl.TxtCheckInput(txt转入, "转入", 100) Then Exit Sub
    End If
    
    '104238:李南春，2017/2/15，检查卡号是否满足发卡控制限制
    If txt卡号.Text <> "" And Len(txt卡号.Text) <> mCurSendCard.lng卡号长度 And Not mCurSendCard.bln严格控制 Then
        Select Case mCurSendCard.byt发卡控制
            Case 0
                MsgBox "输入的卡号小于" & mCurSendCard.str卡名称 & "设定的卡号长度，请重新输入！", vbExclamation, gstrSysName
                If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                Exit Sub
            Case 2
                If MsgBox("输入的卡号小于" & mCurSendCard.str卡名称 & "设定的卡号长度，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If txt卡号.Visible And txt卡号.Enabled Then txt卡号.SetFocus
                    Exit Sub
                End If
        End Select
    End If
    
    '病案从表检查(新增/修改)
    mstrPatiPlus = ""
    '转入机构名称
    mstrPatiPlus = mstrPatiPlus & "," & "入院转入:" & Trim(zlStr.NeedName(txt转入.Text))
    '联系人关系其他关系附加说明
    mstrPatiPlus = mstrPatiPlus & "," & "联系人附加信息:" & Trim(txtLinkManInfo.Text)
    '身份证号
    If Trim(zlCommFun.GetNeedName(cbo国籍.Text)) = "中国" Then
        mstrPatiPlus = mstrPatiPlus & "," & "身份证号状态:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
        mstrPatiPlus = mstrPatiPlus & "," & "外籍身份证号:"
    Else
        If txt身份证号.Text <> "" Then
            mstrPatiPlus = mstrPatiPlus & "," & "外籍身份证号:" & txt身份证号.Text
            mstrPatiPlus = mstrPatiPlus & "," & "身份证号状态:"
            txt身份证号.Text = ""
        Else
            mstrPatiPlus = mstrPatiPlus & "," & "身份证号状态:" & Trim(zlCommFun.GetNeedName(cboIDNumber.Text))
            mstrPatiPlus = mstrPatiPlus & "," & "外籍身份证号:"
        End If
    End If
    If mstrPatiPlus <> "" Then mstrPatiPlus = Mid(mstrPatiPlus, 2)
    
    '预约中心病人床号检查
    If mblnAppoint And mstrAppointBed <> cbo床位.Tag And gbln入院入科 And mbytMode = EMode.E接收预约 Then
        If MsgBox("预约床位【" & mstrAppointBed & "】与当前床位【" & cbo床位.Tag & "】不相同，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If CanFocus(cbo床位) Then cbo床位.SetFocus
            Exit Sub
        End If
    End If
    '变价金额检查
    '刘兴洪:    '29134
    '82401:李南春,2015/3/11,检查对象是否存在
    If mbytInState = 0 And pic磁卡.Visible And txt卡号.Text <> "" Then
        If tabCardMode.SelectedItem.Key = "CardFee" And Not mCurSendCard.rs卡费 Is Nothing Then
            If mCurSendCard.rs卡费!是否变价 Then
                If mCurSendCard.rs卡费!现价 <> 0 And Abs(CCur(txt卡额.Text)) > Abs(mCurSendCard.rs卡费!现价) Then
                    MsgBox "" & mCurSendCard.str卡名称 & "金额绝对值不能大于最高限价：" & Format(Abs(mCurSendCard.rs卡费!现价), "0.00"), vbInformation, gstrSysName
                    txt卡额.SetFocus: Exit Sub
                End If
                If mCurSendCard.rs卡费!原价 <> 0 And Abs(CCur(txt卡额.Text)) < Abs(mCurSendCard.rs卡费!原价) Then
                    MsgBox "" & mCurSendCard.str卡名称 & "金额绝对值不能小于最低限价：" & Format(Abs(mCurSendCard.rs卡费!原价), "0.00"), vbInformation, gstrSysName
                    txt卡额.SetFocus: Exit Sub
                End If
            End If
        End If
    End If
    
    If pic磁卡.Visible And txt卡号.Text <> "" Then
        Select Case mCurSendCard.int密码长度限制
        Case 0
        Case 1
            If Len(txtPass.Text) <> mCurSendCard.int密码长度 Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & mCurSendCard.int密码长度 & "位", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        Case Else
            If Len(txtPass.Text) < Abs(mCurSendCard.int密码长度限制) Then
                MsgBox "注意:" & vbCrLf & "密码必须输入" & Abs(mCurSendCard.int密码长度限制) & "位以上.", vbOKOnly + vbInformation
                If txtPass.Enabled Then txtPass.SetFocus
                Exit Sub
             End If
        End Select
    
        If txtPass.Text <> txtAudi.Text Then
            MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
            txtPass.Text = "": txtAudi.Text = ""
            txtPass.SetFocus: Exit Sub
        End If
        
    End If
    
    '结算方式
    If IsNumeric(txt预交额.Text) And cbo预交结算.Visible And cbo预交结算.Enabled And cbo预交结算.ListIndex = -1 Then
        MsgBox "请确定病人预交款结算方式！", vbInformation, gstrSysName
        cbo预交结算.SetFocus: Exit Sub
    End If
    If Trim(txt卡号.Text) <> "" And cbo发卡结算.Visible And cbo发卡结算.Enabled And cbo发卡结算.ListIndex = -1 Then
        MsgBox "请确定病人" & mCurSendCard.str卡名称 & "结算方式！", vbInformation, gstrSysName
        cbo发卡结算.SetFocus: Exit Sub
    End If
    
    '63246:刘鹏飞,2013-07-03
    If CheckPatiCard = False Then Exit Sub
    
    If mbytInState = 0 Then
        '医保改动
        If mintInsure <> 0 And mstrYBPati <> "" And mbytMode <> 1 Then
            If is个人帐户(cbo预交结算) Then
                If IsNumeric(txt预交额.Text) Then
                    If CCur(StrToNum(txt预交额.Text)) > mcurYBMoney Then
                        MsgBox "医保个人帐户转入金额不能大于余额:" & Format(mcurYBMoney, "0.00"), vbInformation, gstrSysName
                        txt预交额.SetFocus: Exit Sub
                    End If
                End If
            End If
        ElseIf mstrYBPati = "" And IsNumeric(txt预交额.Text) And is个人帐户(cbo预交结算) Then
            MsgBox "非医保病人不能使用个人帐户下帐！", vbInformation, gstrSysName
            cbo预交结算.SetFocus: Exit Sub
        End If
    
        '票据相关检查
        mblnPrepayPrint = False
        If IsNumeric(txt预交额.Text) Then
        'If zlSquareSimulation(lng接口编号, strBalanceInfor) = False Then Exit Sub
        
            mblnPrepayPrint = True
            '检查是否打印票据
            If gbytPrepayPrint = 0 Then
                mblnPrepayPrint = False
            Else
                If gbytPrepayPrint = 2 Then
                    If MsgBox("是否打印预交款票据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        mblnPrepayPrint = False
                    End If
                End If
            End If
            
            If mblnPrepayPrint Then
                If gblnPrepayStrict Then
                    If Trim(txtFact.Text) = "" Then
                        MsgBox "必须输入一个有效的预交票据号码！", vbInformation, gstrSysName
                        txtFact.SetFocus: Exit Sub
                    End If
                    mlng预交领用ID = CheckUsedBill(2, IIf(mlng预交领用ID > 0, mlng预交领用ID, mFactProperty.lngShareUseID), txtFact.Text, 2)
                    If mlng预交领用ID <= 0 Then
                        Select Case mlng预交领用ID
                            Case 0 '操作失败
                            Case -1
                                MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                            Case -2
                                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                            Case -3
                                MsgBox "票据号码不在当前有效领用范围内,请重新输入！", vbInformation, gstrSysName
                                txtFact.SetFocus
                        End Select
                        Exit Sub
                    End If
                Else
                    If Len(txtFact.Text) <> gbytPrepayLen And txtFact.Text <> "" Then
                        MsgBox "预交票据号码长度应该为 " & gbytPrepayLen & " 位！", vbInformation, gstrSysName
                        txtFact.SetFocus: Exit Sub
                    End If
                End If
            End If
        End If
        
        If txt卡号.Text <> "" And pic磁卡.Visible Then
            '保存前检查就诊卡是否有，是否在范围内
            If mCurSendCard.bln严格控制 Then
                mCurSendCard.lng领用ID = CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), txt卡号.Text, mCurSendCard.lng卡类别ID)
                If mCurSendCard.bln就诊卡 Then
                    blnErr = mCurSendCard.lng领用ID <= 0 And Not mCurSendCard.blnOneCard
                Else
                    blnErr = mCurSendCard.lng领用ID <= 0
                End If
                If blnErr Then
                    Select Case mCurSendCard.lng领用ID
                        Case 0 '操作失败
                        Case -1
                            MsgBox "你已没有自用及共用的" & mCurSendCard.str卡名称 & ",请先在本地设置共用批次或领用一批" & mCurSendCard.str卡名称 & "！", vbExclamation, gstrSysName
                        Case -2
                            MsgBox "本地共用的" & mCurSendCard.str卡名称 & "已用完,请重新设置本地共用" & mCurSendCard.str卡名称 & "批次或领用一批" & mCurSendCard.str卡名称 & "！", vbExclamation, gstrSysName
                        Case -3
                            MsgBox "该张" & mCurSendCard.str卡名称 & "号不在有效范围内,请检查是否正确刷卡！", vbExclamation, gstrSysName
                            txt卡号.SetFocus
                    End Select
                    Exit Sub
                End If
            End If
        End If
        
        If mrsInfo Is Nothing Then
            '65689:刘鹏飞,2013-10-30,存在多个相同病人，提供选择器供操作员选择
            If Not (mbytMode = EMode.E接收预约 Or mbytMode = EMode.E正常登记 And mlng病人ID <> 0) Then
                '检查相似病人信息(新增之前检查,以免加入了重复信息！！！)
                Set rsSimilar = SimilarIDs(zlCommFun.GetNeedName(cbo国籍.Text), zlCommFun.GetNeedName(cbo民族), CDate(IIf(IsDate(txt出生日期.Text), txt出生日期.Text, #1/1/1900#)), zlCommFun.GetNeedName(cbo性别), txt姓名.Text, txt身份证号.Text)
                If Not rsSimilar Is Nothing Then
                    If gblnPatiByID And Trim(txt身份证号.Text) <> "" Then
                        '110541 同一身份证只能对应一个建档病人;启用该参数且通过身份证号找到已建档病人时弹出选择框
                        rsSimilar.Filter = "身份证号 ='" & Trim(txt身份证号.Text) & "'"
                        If rsSimilar.RecordCount > 0 Then
                            strSimilar = "在已有的病人信息中发现" & rsSimilar.RecordCount & "个身份证号相同的的病人。" & vbCrLf & vbCrLf & _
                                "提取已有的病人信息请选择病人后[双击]或点击[确定]。"
                            If Not CreatePublicPatient() Then Exit Sub
                            If gobjPublicPatient.ShowSelect(rsSimilar, "ID", "病人选择", strSimilar, , , "0|800|1200|800|800|1500|1000", True) Then
                                txtPatient.Text = "-" & rsSimilar!病人ID
                                txtPatient.SetFocus
                                Call txtPatient_KeyPress(13)
                                Exit Sub
                            End If
                        End If
                    End If
                    rsSimilar.Filter = ""
                    If rsSimilar.RecordCount > 1 Then
                        strSimilar = "在已有的病人信息中发现" & rsSimilar.RecordCount & "个信息相似的病人(国籍,民族,性别,姓名,出生日期相同或身份证号相同)" & vbCrLf & vbCrLf & _
                            "提取已有的病人信息请选择病人后[双击]或点击[确定],登记为新病人请点击[取消]"
                        If Not CreatePublicPatient() Then Exit Sub
                        blnOk = gobjPublicPatient.ShowSelect(rsSimilar, "ID", "病人选择", strSimilar, , , "0|800|1200|800|800|1500|1000")
                        If blnOk = True Then
                            txtPatient.Text = "-" & rsSimilar!病人ID
                            txtPatient.SetFocus
                            Call txtPatient_KeyPress(13)
                            Exit Sub
                        Else
                            MsgBox "该病人的相似记录可以在病人信息管理中使用""合并""功能处理！", vbInformation, gstrSysName
                        End If
                    ElseIf rsSimilar.RecordCount = 1 Then
                        strSimilar = "ID:" & rsSimilar!病人ID & ",门诊号:" & Nvl(rsSimilar!门诊号, "无") & ",住院号:" & Nvl(rsSimilar!住院号, "无") & ",身份证号:" & rsSimilar!身份证号 & ",地址:" & rsSimilar!地址 & ",登记日期:" & rsSimilar!登记时间
                        If MsgBox("在已有的病人信息中发现 1 个信息相似的病人(国籍,民族,性别,姓名,出生日期相同或身份证号相同): " & vbCrLf & vbCrLf & _
                            strSimilar & vbCrLf & vbCrLf & "登记为新病人请选择[是],提取已有的病人信息请选择[否]？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            txtPatient.Text = "-" & Mid(Split(strSimilar, ",")(0), 4)
                            txtPatient.SetFocus
                            Call txtPatient_KeyPress(13)
                            Exit Sub
                        Else
                            MsgBox "该病人的相似记录可以在病人信息管理中使用""合并""功能处理！", vbInformation, gstrSysName
                        End If
                    End If
                End If
                
                '病人ID检查:自动替换新的
                Do While ExistInPatiID(CLng(txtPatient.Tag))
                    txtPatient.Text = zlDatabase.GetNextNo(1)
                    txtPatient.Tag = txtPatient.Text
                Loop
            End If
        End If
        
        If txt住院号.Visible And (mbytKind = E住院入院登记) Then
            If mrsInfo Is Nothing Then
                lng病人ID = IIf(mbytMode = EMode.E接收预约 Or mbytMode = EMode.E正常登记 And mlng病人ID <> 0, Val(txtPatient.Tag), 0)
            Else
                lng病人ID = mrsInfo!病人ID
            End If
            '问题29449 by lesfeng 2010-05-05
            Dim blnTrue As Boolean
            blnTrue = False
            If mbytMode = EMode.E预约登记 Then blnTrue = True
            '60500:刘鹏飞,2013-05-09
            If ExistInPatiNO(txt住院号.Text, lng病人ID, blnTrue) Then
                strno = zlDatabase.GetNextNo(2)
                If Val(txt住院号.Text) = Val(strno) Then
                    MsgBox "当前住院号和自动获取的新住院号重复,请手工修改住院号！", vbInformation, gstrSysName
                Else
                    MsgBox "当前住院号已被使用,将自动获取一个新的住院号,请确认！", vbInformation, gstrSysName
                    txt住院号.Text = strno
                End If
                txt住院号.SetFocus: Exit Sub
            End If
        End If
        
        If txt住院号.Visible And mbytKind = E门诊留观登记 Then
            gstrSQL = "Select 1 From 病人信息 Where 门诊号=[1] And 病人ID<>[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt住院号.Text, Val(txtPatient.Tag))
            If rsTmp.RecordCount > 0 Then
                If Not mblnAuto Then
                    MsgBox "当前门诊号已被使用,将自动获取一个新的门诊号,请确认！", vbInformation, gstrSysName
                    txt住院号.Text = zlDatabase.GetNextNo(3)
                    mblnAuto = True
                    txt住院号.SetFocus: Exit Sub
                Else
                    blnTmp = True
                    txt住院号.Text = Val(txt住院号.Text) + 1
                    mblnAuto = True
                End If
            End If
        End If
        
        If txt住院号.Visible And mbytKind = E住院留观登记 Then
            gstrSQL = "Select 1 From 病案主页 Where 留观号=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt住院号.Text)
            If rsTmp.RecordCount > 0 Then
                MsgBox "当前留观号已被使用,将自动获取一个新的留观号,请确认！", vbInformation, gstrSysName
                txt住院号.Text = zlDatabase.GetNextNo(6)
                txt住院号.SetFocus: Exit Sub
            End If
        End If
        '问题号:51072
        If Len(Trim(txtPass.Text)) <= 0 And Len(Trim(txt卡号.Text)) > 0 Then '没有输入密码
            If zl_Get设置默认发卡密码 = False Then Exit Sub
        End If
        
        If CheckBrushCard = False Then Exit Sub
        
        '90875:李南春,2016/11/8,医疗卡证件类型
        If IsCertificateCard(Val(txtPatient.Tag)) = False Then Exit Sub
        '
        '保存新记录(新增|修改病人信息、入院记录、预交记录(IF要且有)、磁卡记录(IF要且有))
        cmdOK.Enabled = False
        If Not SavePatiNew(mrsInfo Is Nothing And mlng病人ID = 0, lng接口编号, strBalanceInfor) Then
            cmdOK.Enabled = True: Exit Sub
        End If
        
        '门诊留观登记时提示信息
        If blnTmp And mbytKind = E门诊留观登记 Then MsgBox "当前门诊号已被使用，系统自动为您生成了新的门诊号【" & txt住院号.Text & "】", vbInformation, gstrSysName
        gblnOK = True
        
        '打印预交款收据
        If mblnPrepayPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mCurPrepay.strno, 2)
        End If
        
        '打印病案主页:预约登记不打印
        If InStr(mstrPrivs, "首页打印") > 0 Then
            If mbytMode <> 1 Then
                mblnFPagePrint = True
                If gbytFPagePrint = 0 Then
                    mblnFPagePrint = False
                Else
                    If gbytFPagePrint = 2 Then
                        If MsgBox("是否打印病案首页？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            mblnFPagePrint = False
                        End If
                    End If
                End If
                
                If mblnFPagePrint Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131", Me, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, 2)
                End If
            End If
        End If
        
        '打印病人腕带
        If InStr(mstrPrivs, "腕带打印") Then
            mblnWristletPrint = True
            If gbytWristletPrint = 0 Then
                mblnWristletPrint = False
            Else
                If gbytWristletPrint = 2 Then
                    If MsgBox("是否打印病人腕带？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        mblnWristletPrint = False
                    End If
                End If
            End If
            
            If mblnWristletPrint Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131_1", Me, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, 2)
            End If
        End If
        
        '票据相关处理
        '新的一张预交款单据
        If mblnPrepayPrint Then
            If gblnPrepayStrict Then
                If mbytMode <> EMode.E接收预约 Then '外部调用接收时不再产生新号
                    mlng预交领用ID = CheckUsedBill(2, IIf(mlng预交领用ID > 0, mlng预交领用ID, mFactProperty.lngShareUseID), , 2)
                    If mlng预交领用ID <= 0 Then
                        Select Case mlng预交领用ID
                            Case 0 '操作失败
                            Case -1
                                MsgBox "你没有自用和共用的预交票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                            Case -2
                                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                        End Select
                        txtFact.Text = ""
                    Else
                        '严格：取下一个号码
                        txtFact.Text = GetNextBill(mlng预交领用ID)
                    End If
                End If
            Else
                '松散：取下一个号码
                zlDatabase.SetPara "当前预交票据号", txtFact.Text, glngSys, mlngModul
                txtFact.Text = zlCommFun.IncStr(txtFact.Text)
            End If
        End If
        If mbytMode <> EMode.E接收预约 And txt卡号.Text <> "" And pic磁卡.Visible Then
            If mCurSendCard.bln严格控制 Then
                mCurSendCard.lng领用ID = CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), , mCurSendCard.lng卡类别ID)
                If mCurSendCard.lng领用ID <= 0 Then
                    Select Case mCurSendCard.lng领用ID
                        Case 0 '操作失败
                        Case -1
                            MsgBox "你已没有自用及共用的" & mCurSendCard.str卡名称 & ",请先在本地设置共用批次或领用一批" & mCurSendCard.str卡名称 & "！", vbExclamation, gstrSysName
                        Case -2
                            MsgBox "本地共用的" & mCurSendCard.str卡名称 & "已用完,请重新设置本地共用" & mCurSendCard.str卡名称 & "批次或领用一批" & mCurSendCard.str卡名称 & "！", vbExclamation, gstrSysName
                    End Select
                End If
            End If
        End If
                
        cmdOK.Enabled = True
        If mbytMode = EMode.E接收预约 Then
            '保存后退出
            gblnOK = True: Unload Me: Exit Sub
        Else
            '保存后继续下一个病人信息
            mblnICCard = False  '不能放在clearcard中,因为可能先读卡再查出病人
            Call ClearCard
            If Not mCurSendCard.rs卡费 Is Nothing Then
                txt卡额.Text = Format(IIf(mCurSendCard.rs卡费!是否变价 = 1, mCurSendCard.rs卡费!缺省价格, mCurSendCard.rs卡费!现价), "0.00")
            End If
            
            txtPatient.SetFocus
        End If
    ElseIf mbytInState = 1 Then
        '住院号检查
        If txt住院号.Visible And mbytKind = E住院入院登记 And txt住院号.Text <> txt住院号.Tag Then
            If ExistInPatiNO(txt住院号.Text, mlng病人ID, True) Then
                MsgBox "当前住院号已被使用,将自动获取一个新的住院号,请确认！", vbInformation, gstrSysName
                txt住院号.Text = zlDatabase.GetNextNo(2)
                txt住院号.SetFocus: Exit Sub
            End If
        End If
        
        '门诊号检查
        If txt住院号.Visible And mbytKind = E门诊留观登记 And txt住院号.Text <> txt住院号.Tag Then
            gstrSQL = "Select 1 From 病人信息 Where 门诊号=[1] And 病人ID<>[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt住院号.Text, mlng病人ID)
            If rsTmp.RecordCount > 0 And Not mblnAuto Then
                MsgBox "当前门诊号已被使用,将自动获取一个新的门诊号,请确认！", vbInformation, gstrSysName
                txt住院号.Text = zlDatabase.GetNextNo(3)
                mblnAuto = True
                
                txt住院号.SetFocus: Exit Sub
            End If
        End If
        
        '门诊号检查
        If txt住院号.Visible And mbytKind = E住院留观登记 And txt住院号.Text <> txt住院号.Tag Then
            gstrSQL = "Select 1 From 病案主页 Where 留观号=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt住院号.Text)
            If rsTmp.RecordCount > 0 Then
                MsgBox "当前留观号已被使用,将自动获取一个新的留观号,请确认！", vbInformation, gstrSysName
                txt住院号.Text = zlDatabase.GetNextNo(6)
                txt住院号.SetFocus: Exit Sub
            End If
        End If
        
        '90875:李南春,2016/11/8,医疗卡证件类型
        If IsCertificateCard(mlng病人ID) = False Then Exit Sub
        
        '保存修改(入院记录)
        cmdOK.Enabled = False
        If Not SavePatiModi Then
            cmdOK.Enabled = True: Exit Sub
        Else
            '病人信息保存成功后,同步修改病人基本信息
            If bln基本信息调整 And blnMod Then
                strErrInfo = ""
                Call gobjPublicPatient.SavePatiBaseInfo(mlng病人ID, mlng主页ID, Trim(txt姓名.Text), str性别, strAge, str出生日期, Me.Caption, IIf(mlng主页ID <> 0, 2, 1), strErrInfo, True, True)
                If strErrInfo <> "" Then
                    MsgBox strErrInfo, vbInformation + vbOKOnly, Me.Caption
                End If
            End If
        End If
        cmdOK.Enabled = True
        gblnOK = True: Unload Me: Exit Sub
    End If
End Sub

Private Function SavePatiNew(bln新病人 As Boolean, ByVal lng结算卡接口 As Long, ByVal strBalancelInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：保存新的病人入院登记(含病人信息、入院信息、预交款、就诊卡)
    '入参：lng结算卡接口-结算卡接口编号(0-表示普通病人)
    '         strBalancelInfor-模拟结算的相关信息
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-07-09 17:21:15
    '说明：
    '----------------------------------------------------------------------------------------------------------------------
    Dim strPati As String, strDeposit As String, strSQLCard As String, bytMode As Byte
    Dim strSurety As String, str担保人 As String, str到期时间 As String
    Dim strInsure As String, lng护级ID As Long
    Dim lng病人ID As Long, lng主页ID As Long, lng病区ID As Long, lng科室ID As Long, lng预交ID As Long, lng变动ID As Long
    Dim strCard As String, strICCard As String, strno As String, strDepositNO As String, strSQL As String, blnTrans As Boolean, blnInRange As Boolean
    Dim lng西医疾病ID As Long, lng中医疾病ID As Long
    Dim lng西医诊断ID As Long, lng中医诊断ID As Long
    Dim str出生日期 As String, str年龄 As String
    Dim str床号 As String, str床位等级 As String
    Dim str房间号 As String, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim blnNotCommit As Boolean
    Dim bln门诊转住院 As Boolean '38069
    Dim bln个人帐户缴预交 As Boolean    '38069
    Dim cllUpdate As Collection, cllThreeInsert As Collection, cllPro As Collection, cll健康卡 As Collection
    Dim Curdate As Date
    Dim lngInHosTimes  As Long
    Dim i As Long, lngRet As Long
    Dim arrTmp  As Variant
    Dim arrSQL As Variant
    Dim strErr As String
    
    arrSQL = Array()
    
    If cbo入院病区.Visible And cbo入院病区.ListIndex <> -1 Then
        lng病区ID = cbo入院病区.ItemData(cbo入院病区.ListIndex)
    End If
    If cbo入院科室.ListIndex <> -1 Then
        lng科室ID = cbo入院科室.ItemData(cbo入院科室.ListIndex)
    End If
        
    If cbo床位.Visible And cbo床位.ListIndex > 0 Then       '0-不分床,1-家庭病床
        If cbo床位.ListIndex = 1 Then
            str床号 = "家庭病床"
        Else
            str床号 = Trim(Mid(Trim(cbo床位.Text), 1, InStr(Trim(cbo床位.Text), " ") - 1))
            If InStr(Trim(cbo床位.Text), " 房间") <> 0 Then
                If InStr(Trim(cbo床位.Text), "|") - InStr(Trim(cbo床位.Text), "房间:") - 3 > 0 Then
                    str房间号 = Mid(Trim(cbo床位.Text), InStr(Trim(cbo床位.Text), "房间:") + 3, InStr(Trim(cbo床位.Text), "|") - InStr(Trim(cbo床位.Text), "房间:") - 3)
                End If
                strSQL = "Select 性别 From 病人信息 A,床位状况记录 B  Where A.病人ID = b.病人id And b.病人ID Is Not Null And 病区ID = [1] And 房间号 =[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病区ID, str房间号)
                
                Do While Not rsTmp.EOF
                    If Mid(Trim(cbo性别.Text), 3) <> rsTmp!性别 Then
                        If (MsgBox("指定床位所在房间存在男女混住情况，是否继续入住？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
                            Exit Do
                        Else
                            Exit Function
                            If CanFocus(cbo床位) Then cbo床位.SetFocus
                        End If
                    End If
                    rsTmp.MoveNext
                Loop
            End If
        End If
    Else
        str床号 = "-1"    '转为空
    End If
    If cbo护理等级.ListIndex <> -1 Then lng护级ID = cbo护理等级.ItemData(cbo护理等级.ListIndex) '如果没有选,则为0,存储过程中会处理为空
    
    If InStr(1, txt门诊诊断.Tag, ";") <= 0 Then
        lng西医疾病ID = Val(txt门诊诊断.Tag)
    Else
        lng西医诊断ID = Val(txt门诊诊断.Tag)
    End If
    If InStr(1, txt中医诊断.Tag, ";") <= 0 Then
        lng中医疾病ID = Val(txt中医诊断.Tag)
    Else
        lng中医诊断ID = Val(txt中医诊断.Tag)
    End If
    
    str担保人 = Replace(Trim(txt担保人.Text), "'", "''")
    lng病人ID = Val(txtPatient.Tag)
    
    lngInHosTimes = Val(txtTimes.Text)
    If mbytMode = EMode.E预约登记 Then
        lng主页ID = 0
    Else
        lng主页ID = IIf(lngInHosTimes > Val("" & txtPages.Text), lngInHosTimes, Val("" & txtPages.Text))
    End If
    
    If mbytMode = EMode.E正常登记 And mlng病人ID <> 0 Then
        bytMode = 2
    Else
        bytMode = mbytMode
    End If
    
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    '102232新建档病人如果返回F取消保存
    If bln新病人 Then
        If txt出生时间 = "__:__" Then
            str出生日期 = IIf(IsDate(txt出生日期.Text), Format(txt出生日期.Text, "YYYY-MM-DD HH:MM:SS"), "")
        Else
            str出生日期 = IIf(IsDate(txt出生日期.Text), Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS"), "")
        End If
        strSQL = "<XM>" & Trim(txt姓名.Text) & "</XM><XB>" & zlCommFun.GetNeedName(cbo性别.Text) & "</XB><NL>" & str年龄 & "</NL>" & vbNewLine & _
                "<CSRQ>" & str出生日期 & "</CSRQ><YBH>" & txt医保号.Text & "</YBH><SFZH>" & txt身份证号.Text & "</SFZH>"
        If Not FuncPlugPovertyInfo(0, strSQL) Then Exit Function
    End If
    
    If txt出生时间 = "__:__" Then
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & " " & txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If

    strCard = UCase(txt卡号.Text)
    strICCard = IIf(mblnICCard, strCard, "")
    
    bln门诊转住院 = False: bln个人帐户缴预交 = False
    If (mintInsure <> 0 Or InStr(1, mstrPrivs, ";门诊费用转住院;") > 0) And mstrNOS <> "" Then
        bln门诊转住院 = True
    End If
    
    Curdate = zlDatabase.Currentdate

    strPati = "zl_入院病案主页_Insert(" & _
        bytMode & "," & mbytKind & "," & lng病人ID & "," & IIf(txt住院号.Visible And txt住院号.Text <> "", txt住院号.Text, "NULL") & "," & _
        "'" & txt医保号.Text & "','" & txt姓名.Text & "','" & zlCommFun.GetNeedName(cbo性别.Text) & "','" & str年龄 & "'," & _
        "'" & zlCommFun.GetNeedName(cbo费别.Text) & "'," & str出生日期 & "," & _
        "'" & zlCommFun.GetNeedName(cbo国籍.Text) & "','" & zlCommFun.GetNeedName(cbo民族.Text) & "','" & zlCommFun.GetNeedName(cbo学历.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cbo婚姻状况.Text) & "','" & zlCommFun.GetNeedName(cbo职业.Text) & "','" & zlCommFun.GetNeedName(cbo身份.Text) & "'," & _
        "'" & txt身份证号.Text & "','" & txt出生地点.Text & "','" & txt家庭地址.Text & "','" & txt家庭地址邮编.Text & "'," & _
        "'" & txt家庭电话.Text & "','" & txt户口地址.Text & "','" & txt户口地址邮编.Text & "','" & txt联系人姓名.Text & "','" & zlCommFun.GetNeedName(cbo联系人关系.Text) & "'," & _
        "'" & txt联系人地址.Text & "','" & txt联系人电话.Text & "','" & txt工作单位.Text & "'," & Val(txt工作单位.Tag) & "," & _
        "'" & txt单位电话.Text & "','" & txt单位邮编.Text & "','" & txt单位开户行.Text & "','" & txt单位帐号.Text & "'," & _
        "'" & str担保人 & "'," & Val(txt担保额.Text) & "," & IIf(str担保人 = "", "null", chk临时担保.Value) & "," & _
        ZVal(lng科室ID) & "," & lng护级ID & ",'" & zlCommFun.GetNeedName(cbo入院病况.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cbo入院方式.Text) & "','" & zlCommFun.GetNeedName(cbo住院目的.Text) & "'," & chk二级院转入.Value & "," & _
        "'" & zlCommFun.GetNeedName(cbo门诊医师.Text) & "','" & zlCommFun.GetNeedName(txt籍贯.Text) & "','" & zlCommFun.GetNeedName(txt区域.Text) & "'," & _
        "To_Date('" & Format(txt入院时间.Text, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
        chk陪伴.Value & "," & IIf(str床号 = "-1", "NULL", "'" & str床号 & "'") & ",'" & zlCommFun.GetNeedName(Replace(cbo医疗付款.Text, Chr(&HD), "")) & "'," & _
        ZVal(lng西医疾病ID) & "," & ZVal(lng西医诊断ID) & ",'" & Replace(txt门诊诊断.Text, "'", "''") & "'," & _
        ZVal(lng中医疾病ID) & "," & ZVal(lng中医诊断ID) & ",'" & Replace(txt中医诊断.Text, "'", "''") & "'," & _
        IIf(mintInsure <> 0 And mstrYBPati <> "" And bln门诊转住院 = False, mintInsure, "NULL") & ",'" & UserInfo.编号 & "'," & _
        "'" & UserInfo.姓名 & "'," & IIf(bln新病人, 1, 0) & ",'" & txt备注.Text & "'," & _
        ZVal(lng病区ID) & "," & chk再入院.Value & ",'" & zlCommFun.GetNeedName(cbo入院属性.Text) & "'," & lng主页ID & "," & IIf(lngInHosTimes = 0, "NULL", lngInHosTimes) & ",'" & _
        Trim(txt其他证件.Text) & "','" & zlCommFun.GetNeedName(cbo病人类型.Text) & "','" & txt联系人身份证号.Text & "','" & Trim(txtMobile.Text) & "')"
    '病案主页从表信息保存
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",联系人附加信息,入院转入,身份证号状态,外籍身份证号,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病案主页从表_首页整理(" & lng病人ID & "," & lng主页ID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "')"
            End If
            If InStr(",联系人附加信息,身份证号状态,外籍身份证号,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & lng病人ID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
            End If
        Next
    End If
    
    '返回结构化地址SQL
    If gbln启用结构化地址 Then
        Call CreateStructAddressSQL(lng病人ID, lng主页ID, arrSQL, PatiAddress)
    End If
    
    '90875:李南春,2016/11/8,医疗卡证件类型
    If lng病人ID > 0 Then Call AddCertificate(lng病人ID, arrSQL, Curdate)

    '没有权限或预约登记时不可见,本地参数设置为不填担保信息时为禁用
    If txt担保人.Visible And txt担保人.Enabled And str担保人 <> "" Then
        str到期时间 = "null"
        If Not IsNull(dtp担保时限.Value) Then str到期时间 = "To_Date('" & Format(dtp担保时限.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strSurety = "zl_病人担保记录_insert(" & lng病人ID & "," & lng主页ID & ",'" & str担保人 & "'," & _
            Val(txt担保额.Text) & "," & chk临时担保.Value & ",'" & Trim(txtReason.Text) & "',NULL," & str到期时间 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    End If
    '69231,刘尔旋,2014-01-07 14:42:55,保存时强制更新卡对象数据
    Call SetCardVaribles(False)
    '增加发卡记录
    Call AddCardDataSQL(lng病人ID, lng主页ID, lng病区ID, lng科室ID, Curdate, strSQLCard)
    '问题号:57326
    If mbln发卡或绑定卡 Then
        If Check发卡性质(lng病人ID, mCurSendCard.lng卡类别ID) = False Then
            txt卡号.Text = "": txtPass.Text = "": txtAudi.Text = "": txt卡额.Text = ""
            Exit Function
        End If
        '检查结算方式信息是否合法
        If cbo发卡结算.ItemData(cbo发卡结算.ListIndex) = 8 And mCurCardPay.lng医疗卡类别ID = 0 Then
            MsgBox "当前发卡结算方式存在异常，无法使用该结算方式，请检查是否启用相应设备或与管理员联系!", vbInformation + vbOKOnly
            Exit Function
        End If
    End If
    
    '增加预交记录
    Call AddDepositSQL(lng病人ID, lng主页ID, lng科室ID, Curdate, bln个人帐户缴预交, strDeposit)
    '检查预交结算方式信息是否合法
    If IsNumeric(txt预交额.Text) And fra预交.Visible Then
        If cbo预交结算.ItemData(cbo预交结算.ListIndex) = 8 And mCurPrepay.lng医疗卡类别ID = 0 Then
            MsgBox "当前预交结算方式存在异常，无法使用该结算方式，请检查是否启用相应设备或与管理员联系!", vbInformation + vbOKOnly
            Exit Function
        End If
    End If
    
    
    '第一步:处理HIS入院登记和预交等
    '问题:31635
    blnNotCommit = False
    On Error GoTo errH
    Set cllUpdate = New Collection
    Set cllThreeInsert = New Collection
    Set cllPro = New Collection
    Set cll健康卡 = New Collection
      
    gcnOracle.BeginTrans: blnTrans = True
    '病人病案信息
    zlDatabase.ExecuteProcedure strPati, Me.Caption
    '病案主页从表信息\结构化地址
    For i = LBound(arrSQL) To UBound(arrSQL)
         zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    '担保信息
    If strSurety <> "" Then zlDatabase.ExecuteProcedure strSurety, Me.Caption
    '入院发卡
    If strSQLCard <> "" Then zlDatabase.ExecuteProcedure strSQLCard, Me.Caption
    '绑定身份证
    If txt支付密码.Visible = True And txt支付密码.Text <> "" Then
        If zl绑定身份证(cllPro) = False Then Exit Function
        zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
    End If
    '问题号:56599
    '填入病人健康卡信息
    If lng病人ID > 0 Then Call Add健康卡相关信息(lng病人ID, cll健康卡)
    zlExecuteProcedureArrAy cll健康卡, Me.Caption, True, True
    
    '入院预交款
    If strDeposit <> "" And (bln门诊转住院 = False Or bln个人帐户缴预交 = False) Then zlDatabase.ExecuteProcedure strDeposit, Me.Caption
    '入院产生一次计算的费用,门诊留观病人不计算
    '36454,刘鹏飞,2012-09-06,gbln费用计算=True表示在入院未入科调用，False表示在入住时调用
    If mbytMode <> 1 And mbytKind <> E门诊留观登记 And lng病区ID <> 0 And IIf(gbln费用计算 = True, True, str床号 <> "-1") Then
        strSQL = "ZL_住院一次费用_Insert(" & lng病人ID & "," & lng主页ID & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    If bln门诊转住院 = False Then
        '非门诊转住院费用时,先调医保,否则按普通病人先入院,然后转费用,然后调医保方式进行
        If zlInsureComeInSwap(lng病人ID, lng主页ID, lng预交ID, strDeposit, bytMode, True) = False Then
             gcnOracle.RollbackTrans: Exit Function
        End If
        blnNotCommit = True
    End If
    '支付交易
    If Not zlInterfacePrayMoney(cllUpdate, cllThreeInsert) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    '修正三方交易
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    '101160 EMPI包头中心医院
    If Not EMPI_AddORUpdatePati(lng病人ID, lng主页ID, strErr) Then
        gcnOracle.RollbackTrans
        MsgBox strErr, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    If mblnAppoint Then
        '更新预约系统接口{"挂号id_In": "挂号ID","状态_In": "状态" ---已接收，未接收，已退出}
        Call Sys.NewSystemSvr("预约中心", "入住或入住取消", "{""挂号id_In"": """ & mlng挂号ID & """,""状态_In"": ""已接收""}", "")
    End If
    Err = 0: On Error Resume Next
    '入院办理成功开始发送消息
    If mclsMipModule.IsConnect = True And (Not mbytMode = EMode.E预约登记) Then
        '提取变动ID
        If str床号 = -1 Or str床号 = "家庭病床" Then
            strSQL = " Select ID,'' 名称  From  病人变动记录 where 开始原因=1 And 病人ID=[1] And 主页ID=[2]"
        Else
            strSQL = " Select A.ID,B.名称  From  病人变动记录 A,收费项目目录 B" & _
                " where A.开始原因=1 And A.床位等级id=B.id(+) And A.病人ID=[1] And A.主页ID=[2]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", lng病人ID, lng主页ID)
        lng变动ID = rsTmp!ID
        str床位等级 = rsTmp!名称
        
        mclsXML.ClearXmlText '清除缓存中的XML
        
        '--进行消息组装
        '病人信息
        mclsXML.AppendNode "in_patient"
        'patient_id      病人id  1   N
        mclsXML.appendData "patient_id", lng病人ID, xsNumber  '病人ID
        'page_id     主页id  1   N
        mclsXML.appendData "page_id", lng主页ID, xsNumber '主页ID
        'patient_name        姓名    1   S
        mclsXML.appendData "patient_name", txt姓名.Text, xsString '姓名
        'patient_sex     性别    0..1    S
        mclsXML.appendData "patient_sex", zlCommFun.GetNeedName(cbo性别.Text), xsString '性别
        'in_number       住院号  1   S
        mclsXML.appendData "in_number", IIf(txt住院号.Visible And txt住院号.Text <> "", txt住院号.Text, "NULL"), xsString '住院号
        mclsXML.AppendNode "in_patient", True
        
        If str床号 = "-1" Then '普通入院登记
            '住院信息
            mclsXML.AppendNode "in_hospital"
            'change_id       变动id  1   N
            mclsXML.appendData "change_id", lng变动ID, xsNumber '变动ID
            'in_date     入院时间    1   s
            mclsXML.appendData "in_date", Format(txt入院时间.Text, "yyyy-MM-dd HH:mm:ss"), xsString '入院日期
            'in_area_id      入院病区id  0..1    N
            'in_area_title       入院病区    0..1    S
            If lng病区ID > 0 Then
                mclsXML.appendData "in_area_id", lng病区ID, xsNumber '入院病区ID
                mclsXML.appendData "in_area_title", cbo入院病区.Text, xsString  '入院病区
            End If
            'in_dept_id      入院科室id  1   N
            mclsXML.appendData "in_dept_id", lng科室ID, xsNumber '入院科室id
            'in_dept_title       入院科室    1   S
            mclsXML.appendData "in_dept_title", cbo入院科室.Text, xsString  '入院科室
            mclsXML.AppendNode "in_hospital", True
            '提交消息到ZLHIS导航台消息中心
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_001", mclsXML.XmlText
        Else  '入院入科
            '住院信息
            mclsXML.AppendNode "in_hospital"
            'in_date     入院时间    1   s
            mclsXML.appendData "in_date", Format(txt入院时间.Text, "yyyy-MM-dd HH:mm:ss"), xsString '入院日期
            'in_area_id      入院病区id  0..1    N
            mclsXML.appendData "in_area_id", lng病区ID, xsNumber '入院病区ID
            'in_area_title       入院病区    0..1    S
            mclsXML.appendData "in_area_title", cbo入院病区.Text, xsString  '入院病区
            'in_dept_id      入院科室id  1   N
            mclsXML.appendData "in_dept_id", lng科室ID, xsNumber '入院科室id
            'in_dept_title       入院科室    1   S
            mclsXML.appendData "in_dept_title", cbo入院科室.Text, xsString  '入院科室
            mclsXML.appendData "in_again", chk再入院.Value, xsNumber
            mclsXML.AppendNode "in_hospital", True
            '入住情况
            mclsXML.AppendNode "dept_arrange"
            'change_id       变动id  1   N
            mclsXML.appendData "change_id", lng变动ID, xsNumber '变动ID
            'in_room     入住病房    0..1    S
            mclsXML.appendData "in_room", IIf(str床号 = "家庭病床", "", str房间号), xsString
            'in_bed      入住病床    1   S
            mclsXML.appendData "in_bed", IIf(str床号 = "家庭病床", "", str床号), xsString
            'in_tendgrade        护理等级    0..1    S
            If cbo护理等级.ListIndex <> -1 Then
                mclsXML.appendData "in_tendgrade", cbo护理等级.Text, xsString
            Else
                mclsXML.appendData "in_tendgrade", "", xsString
            End If
            'in_bedgrade     床位等级    0..1    S
            mclsXML.appendData "in_bedgrade", IIf(str床号 = "家庭病床", "", str床位等级), xsString
            'in_doctor       住院医师    0..1    S
            mclsXML.appendData "in_doctor", "", xsString
            'duty_nurse      责任护士    0..1    S
            mclsXML.appendData "duty_nurse", "", xsString
            mclsXML.AppendNode "dept_arrange", True
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_002", mclsXML.XmlText
        End If
    End If
    If Err <> 0 Then Err.Clear
    
    '调用外挂接口
    If CreatePlugInOK(glngModul) Then
        On Error Resume Next
        Call gobjPlugIn.InPatiCheckInAfter(lng病人ID, lng主页ID)
        Call zlPlugInErrH(Err, "InPatiCheckInAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    
    Err = 0: On Error GoTo errH
   '问题号:56599
   '写卡
   If mbln发卡或绑定卡 And mCurSendCard.bln是否写卡 Then WriteCard (lng病人ID)
    
    Err = 0: On Error Resume Next:
    zlExecuteProcedureArrAy cllThreeInsert, Me.Caption
    If Err <> 0 Then
        Err = 0: On Error GoTo 0
    End If
    
    Err = 0: On Error GoTo errH
   '第二步:门诊费用转住院
    If bln门诊转住院 Then
        If Not frmChargeTurn.ExecuteTurn(Me, mlngModul, mstrPrivs, mstrNOS, txt住院号.Text, lng主页ID, CDate(txt入院时间.Text), lng科室ID, lng病区ID) Then
            MsgBox "注意:" & "   未执行医保入院交易,但HIS入院成功,请补办入院登记!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        gcnOracle.BeginTrans
        blnTrans = True
        '入院预交款
        If strDeposit <> "" And bln个人帐户缴预交 Then zlDatabase.ExecuteProcedure strDeposit, Me.Caption
        If mintInsure <> 0 And mstrYBPati <> "" And bytMode <> 1 Then
            strSQL = "Zl_病案主页_医保更新(" & lng病人ID & "," & lng主页ID & "," & mintInsure & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        '第三步:处理医保
        '医保交易统一处理
        '预约时可通过医保卡验证提取病人信息，但不保存医保交易
        If zlInsureComeInSwap(lng病人ID, lng主页ID, lng预交ID, strDeposit, bytMode) = False Then
             gcnOracle.RollbackTrans
            MsgBox "注意:" & "   医保入院交易失败,但HIS入院办理成功,请补办医保入院登记!", vbInformation + vbOKOnly, gstrSysName
            mlng病人ID = lng病人ID
            mlng主页ID = lng主页ID
            SavePatiNew = True
            Exit Function
        End If
        blnNotCommit = True
        gcnOracle.CommitTrans: blnTrans = False
    End If
    '问题:31635
    If mintInsure > 0 And mbytMode <> 1 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ComeInSwap, True, mintInsure)
    Dim strOut As String
    Call zlExcuteUploadSwap(lng病人ID, strOut, mobjICCard) '发卡了调用宁波一卡通上传功能
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    SavePatiNew = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    '问题:31635
    If mintInsure > 0 And mbytMode <> 1 And blnNotCommit Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ComeInSwap, False, mintInsure)
    Call SaveErrLog
    Exit Function
End Function

Private Sub SetCardVaribles(ByVal blnPrepay As Boolean)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:设置结算对象数据
    '入参:blnPrepay-是否预交结算对象
    '编制:刘尔旋
    '日期:2014-01-07
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    
    If blnPrepay = True Then
        With cbo预交结算
        If .ListIndex = -1 Then Exit Sub
        lngIndex = .ListIndex + 1
        End With
        '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
        If Not mcolPrepayPayMode Is Nothing Then
            With mCurPrepay
                    .lng医疗卡类别ID = Val(mcolPrepayPayMode(lngIndex)(3))
                    .bln消费卡 = Val(mcolPrepayPayMode(lngIndex)(5)) = 1
                    .str结算方式 = Trim(mcolPrepayPayMode(lngIndex)(6))
                    .str名称 = Trim(mcolPrepayPayMode(lngIndex)(1))
             End With
        End If
    Else
        With cbo发卡结算
            If .ListIndex = -1 Then Exit Sub
            lngIndex = .ListIndex + 1
        End With
        If Not mcolCardPayMode Is Nothing Then
            With mCurCardPay
                .lng医疗卡类别ID = Val(mcolCardPayMode(lngIndex)(3))
                .bln消费卡 = Val(mcolCardPayMode(lngIndex)(5)) = 1
                .str结算方式 = Trim(mcolCardPayMode(lngIndex)(6))
                .str名称 = Trim(mcolCardPayMode(lngIndex)(1))
             End With
         End If
     End If
End Sub

Private Function zlInsureComeInSwap(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng预交ID As Long, ByVal strDeposit As String, ByVal bytMode As Byte, Optional blnMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用医保入院接口
    '入参:个人帐户转预交
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-17 10:40:59
    '问题:38069
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not (mintInsure <> 0 And mstrYBPati <> "" And bytMode <> 1) Then
        '非医保,返回true
        zlInsureComeInSwap = True: Exit Function
    End If
    
    '入院验证
    'mstrYBPati=
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    '9中心;10.顺序号;11人员身份;12帐户余额;13当前状态;14病种ID;15在职(0,1);16退休证号;17年龄段;18灰度级
    '医保入院验证
    If Not gclsInsure.ComeInSwap(lng病人ID, lng主页ID, CStr(Split(mstrYBPati, ";")(1)), mintInsure) Then
        If blnMsg Then
            MsgBox "注意:" & vbCrLf & "   医保入院交易失败!", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    End If
    '入院预交款
    If strDeposit <> "" And is个人帐户(cbo预交结算) Then
        If Not gclsInsure.TransferSwap(lng预交ID, CCur(StrToNum(txt预交额.Text)), mintInsure) Then
            Exit Function
        End If
    End If
    zlInsureComeInSwap = True
End Function


Private Function SavePatiModi() As Boolean
'功能：保存新的病人入院登记(含病人信息、入院信息、预交款、就诊卡)
    Dim lng现病区ID As Long, lng原病区ID As Long
    Dim strSQL As String, strMoney As String
    Dim strSurety As String, str担保人 As String, str到期时间 As String
    Dim lng护级ID As Long, blnTrans As Boolean
    Dim lng西医疾病ID As Long, lng中医疾病ID As Long, lng科室ID As Long
    Dim lng西医诊断ID As Long, lng中医诊断ID As Long
    Dim str出生日期 As String, str年龄 As String
    Dim cll健康卡 As Collection '问题号:56599
    Dim i As Long
    Dim arrTmp  As Variant
    Dim arrSQL As Variant
    Dim strErr As String

    arrSQL = Array()
    
    If cbo护理等级.ListIndex <> -1 Then
        lng护级ID = cbo护理等级.ItemData(cbo护理等级.ListIndex)
    End If
    
    If cbo入院科室.ListIndex <> -1 Then lng科室ID = cbo入院科室.ItemData(cbo入院科室.ListIndex)
    lng原病区ID = Val(cbo入院病区.Tag)
    If cbo入院病区.ListIndex <> -1 Then lng现病区ID = cbo入院病区.ItemData(cbo入院病区.ListIndex)
    
    If InStr(1, txt门诊诊断.Tag, ";") <= 0 Then
        lng西医疾病ID = Val(txt门诊诊断.Tag)
    Else
        lng西医诊断ID = Val(txt门诊诊断.Tag)
    End If
    If InStr(1, txt中医诊断.Tag, ";") <= 0 Then
        lng中医疾病ID = Val(txt中医诊断.Tag)
    Else
        lng中医诊断ID = Val(txt中医诊断.Tag)
    End If
    
    str担保人 = Replace(Trim(txt担保人.Text), "'", "''")
    '说明:此时病人信息中将保存的担保信息是从病人信息中读出的,因为在入院登记后可能担保金额已发生了变化
    If txt出生时间 = "__:__" Then
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str出生日期 = IIf(IsDate(txt出生日期.Text), "TO_Date('" & txt出生日期.Text & " " & txt出生时间.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If
    
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    
    strSQL = "zl_入院病案主页_UPDATE(" & mbytMode & "," & _
        mlng病人ID & "," & IIf(txt住院号.Text = "", "NULL", txt住院号.Text) & ",'" & txt医保号.Text & "'," & _
        "'" & txt姓名.Text & "','" & zlCommFun.GetNeedName(cbo性别.Text) & "','" & str年龄 & "','" & zlCommFun.GetNeedName(cbo费别.Text) & "'," & _
        str出生日期 & "," & _
        "'" & zlCommFun.GetNeedName(cbo国籍.Text) & "','" & zlCommFun.GetNeedName(cbo民族.Text) & "','" & zlCommFun.GetNeedName(cbo学历.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cbo婚姻状况.Text) & "','" & zlCommFun.GetNeedName(cbo职业.Text) & "','" & zlCommFun.GetNeedName(cbo身份.Text) & "'," & _
        "'" & txt身份证号.Text & "','" & txt出生地点.Text & "','" & txt家庭地址.Text & "'," & _
        "'" & txt家庭地址邮编.Text & "','" & txt家庭电话.Text & "','" & txt户口地址.Text & "','" & txt户口地址邮编.Text & "','" & txt联系人姓名.Text & "'," & _
        "'" & zlCommFun.GetNeedName(cbo联系人关系.Text) & "','" & txt联系人地址.Text & "'," & _
        "'" & txt联系人电话.Text & "','" & txt工作单位.Text & "'," & Val(txt工作单位.Tag) & "," & _
        "'" & txt单位电话.Text & "','" & txt单位邮编.Text & "','" & txt单位开户行.Text & "'," & _
        "'" & txt单位帐号.Text & "','" & txt担保人.Tag & "'," & Val(txt担保额.Tag) & "," & IIf(chk临时担保.Tag = "", "null", chk临时担保.Tag) & "," & _
        mlng主页ID & "," & ZVal(lng科室ID) & "," & lng护级ID & ",'" & zlCommFun.GetNeedName(cbo入院病况.Text) & "'," & _
        "'" & zlCommFun.GetNeedName(cbo入院方式.Text) & "','" & zlCommFun.GetNeedName(cbo住院目的.Text) & "'," & _
        chk二级院转入.Value & ",'" & zlCommFun.GetNeedName(cbo门诊医师.Text) & "','" & zlCommFun.GetNeedName(txt籍贯.Text) & "','" & zlCommFun.GetNeedName(txt区域.Text) & "'," & _
        "To_Date('" & Format(txt入院时间.Text, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & zlCommFun.GetNeedName(Replace(cbo医疗付款.Text, Chr(&HD), "")) & "'," & _
        ZVal(lng西医疾病ID) & "," & ZVal(lng西医诊断ID) & ",'" & Replace(txt门诊诊断.Text, "'", "''") & "'," & _
        ZVal(lng中医疾病ID) & "," & ZVal(lng中医诊断ID) & ",'" & Replace(txt中医诊断.Text, "'", "''") & "'," & _
        "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & txt备注.Text & "'," & ZVal(lng现病区ID) & "," & chk再入院.Value & ",'" & _
        zlCommFun.GetNeedName(cbo入院属性.Text) & "','" & Trim(txt其他证件.Text) & "','" & zlCommFun.GetNeedName(cbo病人类型.Text) & _
        "','" & txt联系人身份证号.Text & "','" & Trim(txtMobile.Text) & "')"
    
    '病案主页从表信息保存
    If mstrPatiPlus <> "" Then
        arrTmp = Split(mstrPatiPlus, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(",联系人附加信息,入院转入,身份证号状态,外籍身份证号,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病案主页从表_首页整理(" & mlng病人ID & "," & mlng主页ID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "')"
            End If
            If InStr(",联系人附加信息,身份证号状态,外籍身份证号,", "," & Split(arrTmp(i), ":")(0) & ",") > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & mlng病人ID & ",'" & Split(arrTmp(i), ":")(0) & "','" & Split(arrTmp(i), ":")(1) & "','')"
            End If
        Next
    End If
    
    If txt担保人.Visible And txt担保人.Enabled And str担保人 <> "" Then
        '没有权限时不可见,本地参数设置为不填担保信息时为禁用,以及修改的担保记录时限过期时禁用
        str到期时间 = "null"
        If Not IsNull(dtp担保时限.Value) Then str到期时间 = "To_Date('" & Format(dtp担保时限.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        
        If Trim(txt担保人.Tag) = "" Then    '之前登记时没有担保
            strSurety = "zl_病人担保记录_insert(" & mlng病人ID & "," & mlng主页ID & ",'" & str担保人 & "'," & _
            Val(txt担保额.Text) & "," & chk临时担保.Value & ",'" & Trim(txtReason.Text) & "',NULL," & str到期时间 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            strSurety = "zl_病人担保记录_update(" & mlng病人ID & "," & mlng主页ID & ",'" & str担保人 & "'," & _
                Val(txt担保额.Text) & "," & chk临时担保.Value & ",'" & Trim(txtReason.Text) & "',NULL," & str到期时间 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',To_Date('" & dtp担保时限.Tag & "','yyyy-mm-dd hh24:mi:ss'))"
        End If
    End If
    '返回结构化地址SQL
    If gbln启用结构化地址 Then
        Call CreateStructAddressSQL(mlng病人ID, mlng主页ID, arrSQL, PatiAddress, 1)
    End If
    
    '90875:李南春,2016/11/8,医疗卡证件类型
    If mlng病人ID > 0 Then Call AddCertificate(mlng病人ID, arrSQL, zlDatabase.Currentdate)
    
    On Error GoTo errH
    gcnOracle.BeginTrans
        blnTrans = True
        '修改入院信息前作废一性计算的费用(必须在更改病区前作废)
        If lng现病区ID <> lng原病区ID And mbytMode <> 1 And mbytKind <> E门诊留观登记 Then
            strMoney = "ZL_住院一次费用_Delete(" & mlng病人ID & "," & mlng主页ID & ")"
            zlDatabase.ExecuteProcedure strMoney, Me.Caption
        End If
        
        '修改入院信息
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        '病案主页从表信息
        For i = LBound(arrSQL) To UBound(arrSQL)
             zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        '修改担保信息
        If strSurety <> "" Then zlDatabase.ExecuteProcedure strSurety, Me.Caption
        '问题号:56599
        '填入病人健康卡信息
        Set cll健康卡 = New Collection
        If mlng病人ID > 0 Then Call Add健康卡相关信息(mlng病人ID, cll健康卡)
        zlExecuteProcedureArrAy cll健康卡, Me.Caption, True, True
        
        '修改重新产生一次计算的费用
        '36454,刘鹏飞,2012-09-06,gbln费用计算=True表示在入院未入科调用，False表示在入住时调用
        If lng现病区ID <> lng原病区ID And mbytMode <> 1 And mbytKind <> E门诊留观登记 And gbln费用计算 = True Then
            strMoney = "ZL_住院一次费用_Insert(" & mlng病人ID & "," & mlng主页ID & ")"
            zlDatabase.ExecuteProcedure strMoney, Me.Caption
        End If
        '101160EMPI
        If Not EMPI_AddORUpdatePati(mlng病人ID, mlng主页ID, strErr) Then
            gcnOracle.RollbackTrans
            MsgBox strErr, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    gcnOracle.CommitTrans: blnTrans = False
    '新网96847、118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng病人ID, mlng主页ID) <> 1 Then
            MsgBox "当前启用了影像信息系统接口，但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
    End If
    SavePatiModi = True
    '问题号:56599
    '写卡
    If mbln发卡或绑定卡 And mCurSendCard.bln是否写卡 Then WriteCard (mlng病人ID)
    
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadSurety(lng病人ID As Long, lng主页ID As Long, dat入院时间 As Date)
'功能:入院登记的修改和查看(不含预约及预约接收)加载担保信息
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim Dat最小时间 As Date
    
    On Error GoTo errH
    dtp担保时限.MinDate = dat入院时间
    
    '87466,LPF,担保信息提取，添加条件“删除标志=1”;多条担保信息进行拼接，与病人信息存储的担保信息保持一致
    strSQL = "SELECT 担保人, Decode(担保额, 999999999, '不限', To_Char(担保额, '999999990.00')) AS 担保额, 担保性质, 担保原因, " & vbNewLine & _
            "       To_Char(到期时间, 'yyyy-mm-dd hh24:mi:ss') 到期时间,To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss') 登记时间" & vbNewLine & _
            "FROM 病人担保记录" & vbNewLine & _
            "WHERE 病人id = [1] AND 主页id = [2] AND (到期时间 is null or 到期时间>sysdate) And 删除标志 = 1"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 1 Then
        '多条有效担保记录需要病人信息中修改
        dtp担保时限.Value = Null
        txt担保人 = ""
        Do Until rsTmp.EOF
            If "" & rsTmp!担保额 = "不限" Then
                txt担保额 = Format(Val(txt担保额.Text) + 999999999, "0.00")
            Else
                txt担保额 = Format(Val(txt担保额.Text) + Nvl(rsTmp!担保额, 0), "0.00")
            End If
            If Nvl(dtp担保时限.Value, "3000-01-01 00:00:00") > Nvl(rsTmp!到期时间, "3000-01-01 00:00:00") Then
                dtp担保时限.Value = Nvl(rsTmp!到期时间, "3000-01-01 00:00:00")
            End If
            txt担保人 = IIf(txt担保人 = "", "", txt担保人 & ",") & rsTmp!担保人
            rsTmp.MoveNext
        Loop
        'txt担保人 = "多人担保"
        txt担保人.Enabled = False: txt担保人.BackColor = Me.BackColor
        chkUnlimit.Enabled = False
        txt担保额.Enabled = False: txt担保额.BackColor = Me.BackColor
        dtp担保时限.Enabled = False
        chk临时担保.Enabled = False
        txtReason.Enabled = False
    ElseIf rsTmp.RecordCount = 1 Then
        '修改的是最后一条有效的担保记录
        txt担保人.Text = "" & rsTmp!担保人
        chkUnlimit.Value = IIf("" & rsTmp!担保额 = "不限", 1, 0)   '值不同时会隐式调用click事件
        If chkUnlimit.Value = 1 Then
            txt担保额 = "999999999"
        Else
            txt担保额 = "" & rsTmp!担保额
        End If
        dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm"
        If IsDate("" & rsTmp!到期时间) Then '此时间肯定不会小于入院时间
            dtp担保时限.Value = CDate(rsTmp!到期时间)
        Else
            dtp担保时限.Value = Null
        End If
        dtp担保时限.Tag = rsTmp!登记时间
        txtReason.Text = Nvl(rsTmp!担保原因)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatiOutDate(ByVal lng病人ID As Long) As Date
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Max(出院日期) 出院日期 From 病案主页 Where 病人ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    If rsTmp.RecordCount > 0 Then
        If Not IsNull(rsTmp!出院日期) Then GetPatiOutDate = rsTmp!出院日期
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function ReadPatiReg(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能:读取病人入院登记记录并显示
'调用:mbytInState-0,1,2都有可能调用本过程:登记修改,登记查看,预约接收
    Dim rsTmp As ADODB.Recordset
    Dim rsDiagnosis As ADODB.Recordset
    Dim rsPlus As ADODB.Recordset '病案从表信息值
    Dim DatOut As Date
    Dim lngIdx As Long
    Dim strPlus As String   '记录从表信息名
    Dim i As Long
    Dim arrTmp As Variant
    
    On Error GoTo errH
       
    gstrSQL = _
        " Select A.病人ID,A.就诊卡号,A.门诊号,B.住院号,B.留观号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,H.名称 险类名称,B.费别," & _
        "   A.住院次数,A.国籍,A.民族,A.学历,A.婚姻状况,A.职业,A.身份,A.身份证号,A.手机号,A.其他证件,A.出生日期,A.出生地点,A.家庭地址," & _
        "   A.家庭电话,A.家庭地址邮编, A.户口地址, A.户口地址邮编, A.籍贯, A.联系人关系,A.联系人姓名,A.联系人地址,A.联系人电话,A.联系人身份证号," & _
        "   A.工作单位,A.合同单位ID,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.担保人,A.担保额,A.担保性质," & _
        "   B.险类,Nvl(A.医保号,F.信息值) as 医保号,B.入院方式,b.入院属性,B.入院病况,B.入院日期,B.住院目的,B.入院病床,B.门诊医师,Nvl(B.区域, A.区域) 区域,B.医疗付款方式," & _
        "   Nvl(B.是否陪伴,0) as 是否陪伴,Nvl(B.二级院转入,0) as 二级院转入,C.名称 as 入院科室,B.入院科室ID," & _
        "   G.名称 as 入院病区,B.入院病区ID,D.名称 as 护理等级,B.备注,B.再入院,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型, B.挂号ID " & _
        " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 D,病案主页从表 F,部门表 G,保险类别 H" & _
        " Where B.病人ID=A.病人ID And B.入院科室ID=C.ID And B.入院病区ID=G.ID(+) And B.护理等级ID=D.ID(+) And A.险类=H.序号(+)" & _
        " And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+) And F.信息名(+)='医保号'" & _
        " And A.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, lng主页ID)
    If rsTmp.EOF Then Exit Function
    Set mrsPatiReg = rsTmp.Clone
    
    If Not FuncPlugPovertyInfo(Val(rsTmp!病人ID)) Then Exit Function
    
    txtPatient.Text = rsTmp!病人ID
    txtPatient.Tag = rsTmp!病人ID
    mlng挂号ID = Nvl(rsTmp!挂号ID, 0)
    txt住院号.Text = Decode(mbytKind, E门诊留观登记, Nvl(rsTmp!门诊号), E住院留观登记, Nvl(rsTmp!留观号), Nvl(rsTmp!住院号))
    
    txt住院号.Tag = txt住院号.Text
    txt姓名.Text = rsTmp!姓名
    
    If (mbytMode = EMode.E接收预约 Or mbytMode = EMode.E正常登记 And mlng病人ID <> 0) And mbytInState = EState.E新增 Then '预约接收时,传入的主页ID为0
        txtPages.Text = GetMaxMinPage(lng病人ID) + 1
    Else
        txtPages.Text = lng主页ID
    End If
    
    '预约中心病人不允许编辑入院病区及入院科室
    mstrAppointBed = "": mblnAppoint = False
    If (mbytInState = EState.E修改 And mbytMode = EMode.E预约登记) Or (mbytMode = EMode.E接收预约) Then
        mblnAppoint = IsAppointPati(mlng挂号ID, mstrAppointBed) 'T-预约中心病人
        cbo入院病区.Enabled = Not mblnAppoint
        cbo入院科室.Enabled = Not mblnAppoint
    End If
    
    If mbytInState = EState.E新增 And mbytKind = EKind.E住院入院登记 And mbytMode <> EMode.E预约登记 Then
        txtTimes.Text = GetMaxInHosTimes(lng病人ID) + 1
    Else
        txtTimes.Text = "" & rsTmp!住院次数
    End If
    txtTimes.Tag = txtTimes.Text
    
    txt医保号.Text = Nvl(rsTmp!医保号)
    txt医保号.Locked = Not IsNull(rsTmp!险类)
    txt险类.Text = "" & rsTmp!险类名称
    
    cbo性别.ListIndex = GetCboIndex(cbo性别, IIf(IsNull(rsTmp!性别), "", rsTmp!性别))
    If cbo性别.ListIndex = -1 Then Call SetCboDefault(cbo性别)
    Call LoadOldData("" & rsTmp!年龄, txt年龄, cbo年龄单位)
    mblnChange = False
    txt出生日期.Text = Format(IIf(IsNull(rsTmp!出生日期), "____-__-__", rsTmp!出生日期), "YYYY-MM-DD")
    If rsTmp!年龄 Like "约*" Or Trim(Nvl(rsTmp!年龄)) = "不详" Then
        If "" & rsTmp!出生日期 = "____-__-__" Then
            txt出生日期.Enabled = False
            txt出生时间.Enabled = False
        End If
    Else
        txt出生日期.Enabled = True
        txt出生时间.Enabled = True
    End If
    mblnChange = True
    
    txt入院时间.Text = Format(IIf(lng主页ID = 0 And mbytInState <> 1, zlDatabase.Currentdate, Nvl(rsTmp!入院日期, "")), "yyyy-MM-dd HH:mm")
    If lng主页ID > 1 And mbytInState = EState.E修改 Then
        DatOut = GetPatiOutDate(lng病人ID) '上次出院时间
        If DatOut <> CDate(0) Then txt入院时间.Tag = Format(DatOut, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If Not IsNull(rsTmp!出生日期) Then
        If mbytInState <> 2 Then txt年龄.Text = ReCalcOld(CDate(Format(rsTmp!出生日期, "YYYY-MM-DD HH:MM:SS")), cbo年龄单位, Val(rsTmp!病人ID), , CDate(txt入院时间.Text)) '根据出生日期重算年龄
        If CDate(txt出生日期.Text) - CDate(rsTmp!出生日期) <> 0 Then
            mblnChange = False
            txt出生时间.Text = Format(rsTmp!出生日期, "HH:MM")
            mblnChange = True
        End If
    Else
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
    cbo年龄单位.Tag = txt年龄.Text & "_" & cbo年龄单位.Text  '用于标记年龄是否变动
    
    mblnChange = False          '修改和查看时,身份证号与出生日期独立
    txt身份证号.Text = "" & rsTmp!身份证号
    mblnChange = True
    cboIDNumber.Enabled = txt身份证号.Text = ""
    txt其他证件.Text = "" & rsTmp!其他证件
     
    
    cbo费别.ListIndex = GetCboIndex(cbo费别, IIf(IsNull(rsTmp!费别), "", rsTmp!费别))
    If cbo费别.ListIndex = -1 Then Call SetCboDefault(cbo费别)
    If mbytInState = EState.E修改 Then If Not IsNull(rsTmp!入院病床) Then cbo费别.Enabled = False
    
    cbo国籍.ListIndex = GetCboIndex(cbo国籍, IIf(IsNull(rsTmp!国籍), "", rsTmp!国籍))
    If cbo国籍.ListIndex = -1 Then Call SetCboDefault(cbo国籍)
    
    cbo民族.ListIndex = GetCboIndex(cbo民族, IIf(IsNull(rsTmp!民族), "", rsTmp!民族))
    If cbo民族.ListIndex = -1 Then Call SetCboDefault(cbo民族)
    
    cbo学历.ListIndex = GetCboIndex(cbo学历, IIf(IsNull(rsTmp!学历), "", rsTmp!学历))
    If cbo学历.ListIndex = -1 And Not IsNull(rsTmp!学历) Then
        cbo学历.AddItem rsTmp!学历, 0: cbo学历.ListIndex = 0
    End If
    
    cbo婚姻状况.ListIndex = GetCboIndex(cbo婚姻状况, IIf(IsNull(rsTmp!婚姻状况), "", rsTmp!婚姻状况))
    If cbo婚姻状况.ListIndex = -1 And Not IsNull(rsTmp!婚姻状况) Then
        cbo婚姻状况.AddItem rsTmp!婚姻状况, 0: cbo婚姻状况.ListIndex = 0
    End If
    
    cbo职业.ListIndex = GetCboIndex(cbo职业, IIf(IsNull(rsTmp!职业), "", rsTmp!职业))
    If cbo职业.ListIndex = -1 And Not IsNull(rsTmp!职业) Then
        cbo职业.AddItem rsTmp!职业, 0: cbo职业.ListIndex = 0
    End If
    
    cbo身份.ListIndex = GetCboIndex(cbo身份, IIf(IsNull(rsTmp!身份), "", rsTmp!身份))
    If cbo身份.ListIndex = -1 And Not IsNull(rsTmp!身份) Then
        cbo身份.AddItem rsTmp!身份, 0: cbo身份.ListIndex = 0
    End If
    
    txt区域.Text = Nvl(rsTmp!区域)
    cbo病人类型.ListIndex = GetCboIndex(cbo病人类型, Nvl(rsTmp!病人类型))
             
    txt家庭电话.Text = IIf(IsNull(rsTmp!家庭电话), "", rsTmp!家庭电话)
    txt家庭地址邮编.Text = IIf(IsNull(rsTmp!家庭地址邮编), "", rsTmp!家庭地址邮编)
    txt户口地址邮编.Text = IIf(IsNull(rsTmp!户口地址邮编), "", rsTmp!户口地址邮编)
    txt联系人姓名.Text = IIf(IsNull(rsTmp!联系人姓名), "", rsTmp!联系人姓名)
    
    cbo联系人关系.ListIndex = GetCboIndex(cbo联系人关系, IIf(IsNull(rsTmp!联系人关系), "", rsTmp!联系人关系))
    If Not cbo联系人关系.ListIndex = -1 And Not IsNull(rsTmp!联系人关系) Then
        cbo联系人关系.AddItem rsTmp!联系人关系, 0: cbo联系人关系.ListIndex = 0
    End If
    '记录下从表信息名
    If zlCommFun.GetNeedName(cbo联系人关系.Text) = "其他" Then strPlus = strPlus & "," & "联系人附加信息"
    txt联系人电话.Text = IIf(IsNull(rsTmp!联系人电话), "", rsTmp!联系人电话)
    txt联系人身份证号.Text = IIf(IsNull(rsTmp!联系人身份证号), "", rsTmp!联系人身份证号)
    txt工作单位.Text = IIf(IsNull(rsTmp!工作单位), "", rsTmp!工作单位)
    txt工作单位.Tag = IIf(IsNull(rsTmp!合同单位ID), "", rsTmp!合同单位ID)
    txt单位电话.Text = IIf(IsNull(rsTmp!单位电话), "", rsTmp!单位电话)
    txt单位邮编.Text = IIf(IsNull(rsTmp!单位邮编), "", rsTmp!单位邮编)
    txt单位开户行.Text = IIf(IsNull(rsTmp!单位开户行), "", rsTmp!单位开户行)
    txt单位帐号.Text = IIf(IsNull(rsTmp!单位帐号), "", rsTmp!单位帐号)
    txt备注.Text = Nvl(rsTmp!备注)
    txtMobile.Text = rsTmp!手机号 & ""
    
    If gbln启用结构化地址 Then
        Call ReadStructAddress(lng病人ID, lng主页ID, PatiAddress)
        txt出生地点.Text = PatiAddress(E_IX_出生地点).Value
        txt籍贯.Text = PatiAddress(E_IX_籍贯).Value
        txt家庭地址.Text = PatiAddress(E_IX_现住址).Value
        txt户口地址.Text = PatiAddress(E_IX_户口地址).Value
        txt联系人地址.Text = PatiAddress(E_IX_联系人地址).Value
    Else
        txt出生地点.Text = IIf(IsNull(rsTmp!出生地点), "", rsTmp!出生地点)
        txt籍贯.Text = Nvl(rsTmp!籍贯)
        txt家庭地址.Text = IIf(IsNull(rsTmp!家庭地址), "", rsTmp!家庭地址)
        txt家庭地址.ToolTipText = txt家庭地址.Text
        txt户口地址.Text = IIf(IsNull(rsTmp!户口地址), "", rsTmp!户口地址)
        txt联系人地址.Text = IIf(IsNull(rsTmp!联系人地址), "", rsTmp!联系人地址)
        txt联系人地址.ToolTipText = txt联系人地址.Text
    End If

    '担保信息(预约不输担保信息,预约接收无需读担保)
    If mbytMode = 0 And mlng病人ID <> 0 Then
        If mbytInState = 1 Then
            txt担保人.Tag = "" & rsTmp!担保人   '用于原样保存回到病人信息中,因为可能存在已到期的担保,就不允许修改
            txt担保额.Tag = "" & rsTmp!担保额
            chk临时担保.Tag = "" & rsTmp!担保性质
        End If
        Call LoadSurety(lng病人ID, lng主页ID, rsTmp!入院日期)
    End If
    
    '入院信息
    If gbln先选病区 Then    '(只影响修改时)
        '问题29007 by lesfeng 2010-04-12
        If IsNull(rsTmp!入院病区) And Not IsNull(rsTmp!入院科室) Then
            mrsUnitDept.Filter = "科室ID=" & Val(rsTmp!入院科室ID) & " and 病区ID=" & Val(rsTmp!入院科室ID)
            If mrsUnitDept.RecordCount > 0 Then
                lngIdx = cbo.FindIndex(cbo入院病区, mrsUnitDept!病区ID)
                If lngIdx <> -1 Then
                    cbo入院病区.ListIndex = lngIdx
                End If
            Else
                mrsUnitDept.Filter = "科室ID=" & Val(rsTmp!入院科室ID)
                If mrsUnitDept.RecordCount > 0 Then
                    lngIdx = cbo.FindIndex(cbo入院病区, mrsUnitDept!病区ID)
                    If lngIdx <> -1 Then
                        cbo入院病区.ListIndex = lngIdx
                    End If
                End If
            End If
        Else
            cbo入院病区.ListIndex = GetCboIndex(cbo入院病区, "" & rsTmp!入院病区)
        End If
        '----------------------------------
        If cbo入院病区.ListIndex = -1 Then
            If Not IsNull(rsTmp!入院病区) And mbytInState = EState.E查阅 Then
                cbo入院病区.AddItem rsTmp!入院病区
                cbo入院病区.ItemData(cbo入院病区.NewIndex) = Nvl(rsTmp!入院病区ID, 0)
                cbo入院病区.ListIndex = cbo入院病区.NewIndex
            Else
                If cbo入院病区.ListCount > 0 Then cbo入院病区.ListIndex = 0 '第一个是不确定病区
            End If
        End If
        cbo入院科室.ListIndex = GetCboIndex(cbo入院科室, rsTmp!入院科室)
        If cbo入院科室.ListIndex = -1 And mbytInState = EState.E查阅 Then
            cbo入院科室.AddItem rsTmp!入院科室, 0
            cbo入院科室.ItemData(cbo入院科室.NewIndex) = Nvl(rsTmp!入院科室ID, 0)
            cbo入院科室.ListIndex = 0
        End If
    Else
        cbo入院科室.ListIndex = GetCboIndex(cbo入院科室, rsTmp!入院科室)
        If cbo入院科室.ListIndex = -1 And mbytInState = EState.E查阅 Then
            cbo入院科室.AddItem rsTmp!入院科室, 0
            cbo入院科室.ItemData(cbo入院科室.NewIndex) = Nvl(rsTmp!入院科室ID, 0)
            cbo入院科室.ListIndex = 0
        End If
        cbo入院病区.ListIndex = GetCboIndex(cbo入院病区, "" & rsTmp!入院病区)
        If cbo入院病区.ListIndex = -1 Then
            If Not IsNull(rsTmp!入院病区) And mbytInState = EState.E查阅 Then
                cbo入院病区.AddItem rsTmp!入院病区
                cbo入院病区.ItemData(cbo入院病区.NewIndex) = Nvl(rsTmp!入院病区ID, 0)
                cbo入院病区.ListIndex = cbo入院病区.NewIndex
            Else
                If cbo入院病区.ListCount > 0 Then cbo入院病区.ListIndex = 0
            End If
        End If
    End If
    
    If gbln入院入科 And mbytMode <> EMode.E预约登记 And mbytInState = EState.E查阅 Then
        cbo床位.ListIndex = GetCboIndex(cbo床位, Nvl(rsTmp!入院病床))
        If cbo床位.ListIndex = -1 And Not IsNull(rsTmp!入院病床) Then    '如果有床号，是不允许修改的
            cbo床位.AddItem Nvl(rsTmp!入院病床), 0
            cbo床位.ListIndex = 0
        End If
    End If
   
    '记录原始值
    If cbo入院科室.ListIndex <> -1 And mbytInState = EState.E修改 Then
        cbo入院科室.Tag = cbo入院科室.ItemData(cbo入院科室.ListIndex)
    End If
    If cbo入院病区.ListIndex <> -1 And mbytInState = EState.E修改 Then
        cbo入院病区.Tag = cbo入院病区.ItemData(cbo入院病区.ListIndex)
    End If
    
    cbo入院病况.ListIndex = GetCboIndex(cbo入院病况, IIf(IsNull(rsTmp!入院病况), "", rsTmp!入院病况))
    If cbo入院病况.ListIndex = -1 Then Call SetCboDefault(cbo入院病况)
        
    cbo入院方式.ListIndex = GetCboIndex(cbo入院方式, IIf(IsNull(rsTmp!入院方式), "", rsTmp!入院方式))
    If cbo入院方式.ListIndex = -1 Then Call SetCboDefault(cbo入院方式)
    '记录下从表信息名
    If zlCommFun.GetNeedName(cbo入院方式.Text) = "转入" Then strPlus = strPlus & "," & "入院转入"
    
    '刘兴宏:2007/09/13
    cbo入院属性.ListIndex = GetCboIndex(cbo入院属性, IIf(IsNull(rsTmp!入院属性), "", rsTmp!入院属性))
    If cbo入院属性.ListIndex = -1 Then Call SetCboDefault(cbo入院属性)
    
    cbo住院目的.ListIndex = GetCboIndex(cbo住院目的, IIf(IsNull(rsTmp!住院目的), "", rsTmp!住院目的))
    If cbo住院目的.ListIndex = -1 Then Call SetCboDefault(cbo住院目的)
    
    cbo医疗付款.ListIndex = GetCboIndex(cbo医疗付款, IIf(IsNull(rsTmp!医疗付款方式), "", rsTmp!医疗付款方式), , True)
    If cbo医疗付款.ListIndex = -1 Then Call SetCboDefault(cbo医疗付款)
            
            
            
    If IsNull(rsTmp!护理等级) Then
        If cbo护理等级.ListCount = 0 Then cbo护理等级.AddItem "": cbo护理等级.ItemData(cbo护理等级.NewIndex) = 0    '查阅时
        cbo护理等级.ListIndex = 0 '装入时,第一个是空
    Else
        cbo护理等级.ListIndex = GetCboIndex(cbo护理等级, rsTmp!护理等级)
        If cbo护理等级.ListIndex = -1 Then
            cbo护理等级.AddItem rsTmp!护理等级
            cbo护理等级.ListIndex = cbo护理等级.NewIndex
        End If
    End If
    
    cbo门诊医师.ListIndex = GetCboIndex(cbo门诊医师, IIf(IsNull(rsTmp!门诊医师), "", rsTmp!门诊医师))
    If cbo门诊医师.ListIndex = -1 And Not IsNull(rsTmp!门诊医师) Then
        cbo门诊医师.AddItem rsTmp!门诊医师, 0: cbo门诊医师.ListIndex = 0
    End If
    
        
    chk再入院.Value = Val("" & rsTmp!再入院)
    chk二级院转入.Value = rsTmp!二级院转入
    chk陪伴.Value = rsTmp!是否陪伴
    
    
    '显示病人诊断情况
    Set rsDiagnosis = GetDiagnosticInfo(lng病人ID, lng主页ID, "1,11", IIf(mbytMode = EMode.E接收预约 Or mbytMode = EMode.E正常登记 And mlng病人ID <> 0 And mbytInState = EState.E新增, "3", "2"), Val(rsTmp!入院科室ID & ""))
    If Not rsDiagnosis Is Nothing Then
        rsDiagnosis.Filter = "诊断类型=1"
        If Not rsDiagnosis.EOF Then
            txt门诊诊断.Text = Nvl(rsDiagnosis!诊断描述): txt门诊诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl门诊诊断.Tag = txt门诊诊断.Text
        End If
        
        rsDiagnosis.Filter = "诊断类型=11"
        If Not rsDiagnosis.EOF Then
            txt中医诊断.Text = Nvl(rsDiagnosis!诊断描述): txt中医诊断.Tag = Nvl(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
        End If
    End If
     
    If Not IsNull(rsTmp!险类) Then
        If mstrYBPati = "" Then mstrYBPati = "是医保"         '接收,修改,查看功能调用,只是为了标识是否医保病人
    End If
    '问题号:56599
    Call Load健康卡相关信息(lng病人ID)
    '病案从表信息
    If strPlus <> "" Then
        strPlus = Mid(strPlus, 2)
        arrTmp = Split(strPlus, ",")
        Set rsPlus = Get病案主页从表(lng病人ID, lng主页ID, strPlus)
        
        If rsPlus.RecordCount > 0 Then
            rsPlus.Filter = "信息名='联系人附加信息'"
            If Not rsPlus.EOF Then txtLinkManInfo.Text = rsPlus!信息值 & ""
            rsPlus.Filter = "信息名='入院转入'"
            If Not rsPlus.EOF Then txt转入.Text = rsPlus!信息值 & ""
        End If
    End If
    
     '病人信息从表
    If txt身份证号.Text = "" Then
        Set rsPlus = Get病人信息从表(lng病人ID, "身份证号状态")
        rsPlus.Filter = "信息名='身份证号状态'"
        If Not rsPlus.EOF Then
            If Not IsNull(rsPlus!信息值) Then
                cbo.Locate cboIDNumber, zlCommFun.GetNeedName(rsPlus!信息值)
            End If
        End If
        If Trim(zlCommFun.GetNeedName(cbo国籍.Text)) <> "中国" And Trim(txt身份证号.Text) = "" Then
            If Trim(zlCommFun.GetNeedName(cboIDNumber.Text)) = "" Then
                 Set rsPlus = Get病人信息从表(lng病人ID, "外籍身份证号")
                rsPlus.Filter = "信息名='外籍身份证号'"
                If Not rsPlus.EOF Then
                    If Not IsNull(rsPlus!信息值) Then
                        txt身份证号.Text = "" & rsPlus!信息值
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
    RequestCode = gint门诊诊断输入 = 2 Or (gint门诊诊断输入 = 3 And mstrYBPati <> "")
End Function
''''
''''
''''
''''Private Function zlSquareSimulation(ByRef lngOut接口编号 As Long, ByRef strOutBalanceInfor As String) As Boolean
''''    ------------------------------------------------------------------------------------------------------------------------
''''功能:     进行卡虚拟结算交易
''''入参:
''''出参:      lngOut接口编号 -接口编号
''''             strBalanceInfor -返回结算交易
''''返回:     成功 (或非结算卡结算), 返回true, 否则返回False
''''编制:     刘兴洪
''''    日期：2010-07-09 16:55:19
''''说明:
''''    ------------------------------------------------------------------------------------------------------------------------
''''    Dim i As Long
''''    Dim strBlanceInfor As String, varData As Variant, blnHave结算方式 As Boolean, lng接口编号 As Long
''''    strOutBalanceInfor = ""
''''    lngOut接口编号 = 0: strOutBalanceInfor = ""
''''    If cbo预交结算.ItemData(cbo预交结算.ListIndex) <> 8 Then    '非结算卡返回为true
''''        zlSquareSimulation = True
''''        Exit Function
''''    End If
''''    If Not mtySquareCard.blnExistsObjects Or mobjSquareCard Is Nothing Then
''''        MsgBox "注意:" & vbCrLf & "    结算卡结算部件不存在,不能用结算卡性质缴预交,请检查!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
''''        Exit Function
''''    End If
''''
''''    zlSimulationBrushCard(ByVal frmMain As Object, ByVal Dbl刷卡金额 As Double, _
''''        ByRef lng接口编号 As Long, ByRef strBlanceInfor As String) As Boolean
''''        '------------------------------------------------------------------------------------------------------------------------
''''        '功能：选择指定卡类型
''''        '入参：frmMain HIS传入 调用的主窗体
''''        '         Dbl刷卡金额 HIS传入 传入预交界面中的金额
''''        '         Lng接口编号          HIS不传入
''''        '出参：Lng接口编号 传出    以何种结算卡结算
''''        '         strBlanceInfor  传出    用||分隔: 接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注
''''        '返回：
''''        '编制：刘兴洪
''''        '日期：2010-06-18 11:33:22
''''        '说明：在预交款管理中增加预交时，点击确定按钮时调用(事务前调用)
''''        '------------------------------------------------------------------------------------------------------------------------
''''    模拟计算
''''     If mobjSquareCard.zlSimulationBrushCard(Me, Val(StrToNum(txt预交额.Text)), lng接口编号, strBlanceInfor) = False Then
''''          Exit Function
''''     End If
''''    strBlanceInfor:接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注
''''    varData = Split(strBlanceInfor, "||")
''''    If Trim(strBlanceInfor) = "" Then
''''           MsgBox "注意:" & vbCrLf & "    返回的结算信息格式错误,请与POS接口开放联系!", vbInformation + vbDefaultButton1 + vbOKOnly
''''           Exit Function
''''    End If
''''
''''    blnHave结算方式 = False
''''
''''    With cbo预交结算
''''       For i = 0 To .ListCount - 1
''''            If NeedName(.List(i)) = CStr(varData(2)) Then
''''                    blnHave结算方式 = True:
''''                  If i <> .ListIndex Then .ListIndex = i
''''                  Exit For
''''            End If
''''       Next
''''        If Val(varData(3)) <= 0 Then
''''                MsgBox "注意:" & vbCrLf & "    卡结算返回的结算金额不能小于等于零,请检查!", vbInformation + vbDefaultButton1 + vbOKOnly
''''                Exit Function
''''        End If
''''        If Round(Val(varData(3)), 3) <> Round(Val(StrToNum(txt预交额.Text)), 3) Then
''''            txt预交额.Text = Format(Val(varData(3)), "0.00")
''''        End If
''''
''''        If CStr(varData(2)) = "" Then
''''                MsgBox "注意:" & vbCrLf & "    卡结算返回的结算方式为空了,请检查!", vbInformation + vbDefaultButton1 + vbOKOnly
''''                Exit Function
''''        End If
''''        If blnHave结算方式 = False Then
''''            MsgBox "注意:" & vbCrLf & "    卡结算返回的结算方式不正确,不存在:" & varData(2) & vbCrLf & _
''''                "     或未设置应用场合,请在结算方式中设置!", vbInformation + vbDefaultButton1 + vbOKOnly
''''            Exit Function
''''        End If
''''    End With
''''    strOutBalanceInfor = strBlanceInfor: lngOut接口编号 = lng接口编号
''''    zlSquareSimulation = True
''''End Function
'''Private Function zlSequareBlanceToDeposit(ByVal lng预交ID As Long, ByVal lng接口编号 As Long, strBlanceInfor As String) As Boolean
'''    '---------------------------------------------------------------------------------------------------------------------------------------------
'''    '功能:结算卡的结算
'''    '返回:成功,返回true,否则返回False
'''    '编制:刘兴洪
'''    '日期:2010-02-08 16:40:12
'''    '---------------------------------------------------------------------------------------------------------------------------------------------
'''    Dim rsSquare As ADODB.Recordset
'''    If mbytInState <> 0 Then GoTo goEnd:
'''
'''    '刘兴洪:
'''    If Not mtySquareCard.blnExistsObjects Then GoTo goEnd:
'''    If mobjSquareCard Is Nothing Then GoTo goEnd:
'''    '    zlBrushCardToDeposit(ByVal lng预交ID As Long, ByVal lng结算卡 As Long, ByRef strBlanceInfor As String) As Boolean
'''    '    '------------------------------------------------------------------------------------------------------------------------
'''    '    '功能：刷卡存预交交易
'''    '    '入参： lng预交ID-预交ID
'''    '    '           lng结算卡-结算卡编号
'''    '    '出参：strBlanceInfor-返回刷卡信息:
'''    '    '         用||分隔: 接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注
'''    '    '返回：成功返回true,否则返回False
'''    '    '编制：刘兴洪
'''    '    '日期：2010-06-18 11:33:22
'''    '    '说明：在预交款管理中增加预交时，点击确定按钮时调用(事务中调用)
'''    '    '          出参一定要传入正确,否则会出现程序错误
'''    '    '------------------------------------------------------------------------------------------------------------------------
'''     If mobjSquareCard.zlBrushCardToDeposit(lng预交ID, lng接口编号, strBlanceInfor) = False Then
'''          Exit Function
'''     End If
'''goEnd:
'''    zlSequareBlanceToDeposit = True
'''    Exit Function
'''End Function
 

Private Sub txt住院号_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt住院号.Locked = True Then
        glngTXTProc = GetWindowLong(txt住院号.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt住院号.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt住院号_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt住院号.Locked = True Then
        Call SetWindowLong(txt住院号.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt住院号_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If mbytKind = E门诊留观登记 Or mbytKind = E住院留观登记 Then Exit Sub
    
    strSQL = "Select 病人ID,住院号,姓名,身份证号 From 病人信息 Where 住院号 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Trim(txt住院号.Text)))
    Call MergePatient(rsTmp, 1)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetMergePatiInfo(lng病人ID As Long) As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '主页ID=0时(不是NULL)，表示预约入院
    strSQL = _
        " Select A.病人ID,Decode(B.病人ID,NULL,NULL,Nvl(B.主页ID,0)) as 主页ID," & _
        " A.姓名,B.住院号,B.入院日期,B.出院日期" & _
        " From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID(+) And A.病人ID=[1]" & _
        " Order by Nvl(B.主页ID,0)"
    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    If Not rsTmp.EOF Then Set GetMergePatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub RestoreYB()
    Dim lng病人ID As Long, lng病种ID As Long
    Dim objCurrent As Object, strTxt As String, arrTxt As Variant
    Dim i As Long, blnDo As Boolean, arrPati As Variant
    Dim objcbo As ComboBox
    
    If (mbytMode = EMode.E接收预约 Or mbytMode = EMode.E正常登记 And mlng病人ID <> 0) Then
        lng病人ID = mlng病人ID
    ElseIf Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If MsgBox("当前已经输入一个病人,是否要以该病人的身份进行验证？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                lng病人ID = mrsInfo!病人ID
            End If
        End If
    End If
    
    '医保改动
    mintInsure = mintInsureBak
    mstrYBPati = mstrYBPatiBak
    If mstrYBPati <> "" Then
        arrPati = Split(mstrYBPati, ";")
        '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID,...
        If UBound(arrPati) >= 8 Then
            If Val(arrPati(8)) > 0 Then
                txtPatient.Locked = txtPatient.Locked
                If mstrYBPati = "" Then txt姓名.SetFocus: Exit Sub  '可能因为余额不足提醒选择了退出等,调用了clearcard
            ElseIf mrsInfo Is Nothing Then
                If txtPatient.Tag = "" Then '如果尚未产生
                    txtPatient.Text = zlDatabase.GetNextNo(1) '新病人ID
                    txtPatient.Tag = txtPatient.Text
                    If txt住院号.Visible And mbytKind = EKind.E住院入院登记 Then
                        txt住院号.Text = zlDatabase.GetNextNo(2)
                    ElseIf txt住院号.Visible And mbytKind = EKind.E住院留观登记 Then
                        txt住院号.Text = zlDatabase.GetNextNo(6)
                    End If
                End If
            End If
        End If
        
        txt医保号.Text = arrPati(1)
        txt医保号.Locked = True
        
        txt姓名.Text = arrPati(3)
        cbo性别.ListIndex = GetCboIndex(cbo性别, CStr(arrPati(4)))
        If IsDate(arrPati(5)) Then
            txt出生日期.Text = Format(arrPati(5), "yyyy-MM-dd")
            Call txt出生日期_LostFocus
        End If
        txt身份证号.Text = arrPati(6)
        txt工作单位.Text = arrPati(7)
       
        '保险病种作为入院诊断
        If UBound(arrPati) >= 14 Then
            If Val(arrPati(14)) > 0 Then
                lng病种ID = Val(arrPati(14))
                
                If txt门诊诊断.Text = "" And Not RequestCode Then
                    txt门诊诊断.Text = Get病种名(lng病种ID)
                End If
            End If
        End If
        
        '获取个人帐户余额
        mcurYBMoney = mcurYBMoneyBak
        lblYBMoney.Caption = "个人帐户余额：" & Format(mcurYBMoney, "0.00")
        lblYBMoney.Visible = True
        
        '医疗付款方式缺省=社会基本医疗保险
        For i = 0 To cbo医疗付款.ListCount
            If InStr(cbo医疗付款.List(i), Chr(&HD)) > 0 Then cbo医疗付款.ListIndex = i: Exit For
        Next
       
        If Not IsDate(txt出生日期.Text) Then
            txt出生日期.SetFocus
        Else
            strTxt = "txt年龄,cbo性别,cbo费别,cbo国藉,cbo民族,cbo学历,cbo婚姻状况,cbo职业,cbo身份," & _
                     "txt身份证号,txt出生地点,txt家庭地址,txt家庭地址邮编,txt家庭电话,txt户口地址,txt户口地址邮编,txt工作单位,txt单位电话,txt单位邮编," & _
                     "txt单位开户行,txt单位帐号,txt联系人姓名,cbo联系人关系,txt联系人地址,txt联系人电话,txt联系人身份证号,txt担保人,txt担保额"
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
        If CanFocus(cbo入院科室) Then cbo入院科室.SetFocus
    Else
        txt姓名.SetFocus
    End If
End Sub

Private Function GetPatientByName(ByVal strInput As String) As ADODB.Recordset
'功能：读取病人信息
'说明：提取失败时，mrsInfo = Nothing
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH

    '通过姓名模糊查找病人(允许输入病人标识时)
    strPati = " Select 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
        " C.住院号,A.门诊号,A.住院次数,trunc(C.入院日期,'dd') as 入院日期,trunc(C.出院日期,'dd') as 出院日期,A.出生日期,A.身份证号,A.手机号,A.家庭地址,A.工作单位,zl_PatiType(A.病人ID) 病人类型" & _
        " From 病人信息 A,部门表 B,病案主页 C" & _
        " Where A.停用时间 is NULL And A.病人ID=C.病人ID(+) And Nvl(A.主页ID,0)=C.主页ID(+) And A.当前科室ID=B.ID(+) And Rownum<101" & _
        " And A.姓名 Like [1]" & IIf(gintNameDays = 0, "", " And (A.登记时间>Trunc(Sysdate-[2]) Or A.就诊时间>Trunc(Sysdate-[2]))") & " And A.病人ID <> [3] And a.主页ID Is Not Null And C.主页ID(+)<>0 "
    strPati = strPati & " Union ALL " & _
                            "Select 0,0,-NULL,'[当前病人]',NULL,NULL,-NULL,-NULL,-NULL,To_Date(NULL),To_Date(NULL),To_Date(NULL),NULL,NULL,NULL,NULL,'普通病人' From Dual"
    strPati = strPati & " Order by 排序ID,姓名,入院日期 Desc"
    
    vRect = GetControlRect(txt姓名.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txt姓名.Height, blnCancel, False, True, strInput, gintNameDays, Val(txtPatient.Tag))
                
    '只有一行数据时,blncancel返回false,按取消返回也是一样
    If Not blnCancel Then
        If rsTmp!ID = 0 Then Exit Function
    Else
        Call zlControl.TxtSelAll(txt姓名)
        txt姓名.SetFocus: Exit Function
    End If
    
    Set GetPatientByName = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = Nothing
End Function

Private Sub MergePatient(ByVal rsTmp As ADODB.Recordset, ByVal bytMode As Byte)
    'bytMode = 0 ,通过 cmdName 调用, bytMode = 1 通过验证住院号 调用
    Dim rsPatiS As ADODB.Recordset
    Dim rsPatiO As ADODB.Recordset
    Dim strSQL As String, Curdate As Date
    Dim i As Integer, j As Integer
    Dim str合并原因 As String, strInfo As String

    If rsTmp Is Nothing Then Exit Sub
    If mrsInfo Is Nothing And mrsPatiReg Is Nothing Then Exit Sub
    If rsTmp.RecordCount = 0 Then Exit Sub
    If Nvl(rsTmp!病人ID, 0) = Val(txtPatient.Text) Then Exit Sub
    If rsTmp!姓名 = Trim(txt姓名.Text) Then
        '45976:刘鹏飞,2012-09-21,身份证号不同进行相关提示。
        If Trim(Nvl(rsTmp!身份证号)) <> Trim(txt身份证号.Text) Then
            strInfo = "病人姓名重复但身份证号不同，是否对该病人进行合并?" & vbCrLf & _
                "要保留病人的身份证号：" & Trim(Nvl(rsTmp!身份证号)) & vbCrLf & _
                "要合并病人的身份证号：" & Trim(txt身份证号.Text)
        Else
            strInfo = "病人姓名和身份证号重复,是否对该病人进行合并?"
        End If
        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '检查医保病人是否存在未结费用
            If ExistFeeInsurePatient(Val(txtPatient.Text)) Then
                MsgBox "该医保病人存在未结费用,请先结清后再合并！", vbExclamation, gstrSysName: Exit Sub
            End If

            If ExistFeeInsurePatient(Val(rsTmp!病人ID)) Then
                MsgBox "您查找到的医保病人存在未结费用,请先结清后再合并！", vbExclamation, gstrSysName: Exit Sub
            End If

            Set rsPatiS = GetMergePatiInfo(Val(txtPatient.Text))
            Set rsPatiO = GetMergePatiInfo(Val(rsTmp!病人ID))


            'AB都住过院
            If Not IsNull(rsPatiS!主页ID) And Nvl(rsPatiS!主页ID, 0) <> 0 And Not IsNull(rsPatiO!主页ID) And Nvl(rsPatiO!主页ID, 0) <> 0 Then
                '1.先住院的在院,不允许(先后住院可以为：出院-出院,出院-在院；不允许：在院-出院,在院-在院)
                '因为除病人合并外,程序不额外处理自动出院或撤消出院
                rsPatiS.MoveLast
                rsPatiO.MoveLast
                If rsPatiS!入院日期 <= rsPatiO!入院日期 Then
                    If IsNull(rsPatiS!出院日期) Then
                        MsgBox "病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    If IsNull(rsPatiO!出院日期) Then
                        MsgBox "病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If

                '2.时间交叉提示是否继续
                Curdate = zlDatabase.Currentdate
                rsPatiS.MoveFirst
                For i = 1 To rsPatiS.RecordCount
                    rsPatiO.MoveFirst
                    For j = 1 To rsPatiO.RecordCount
                        If Not (rsPatiO!入院日期 >= IIf(IsNull(rsPatiS!出院日期), Curdate, rsPatiS!出院日期) Or _
                            IIf(IsNull(rsPatiO!出院日期), Curdate, rsPatiO!出院日期) <= rsPatiS!入院日期) Then
                            MsgBox "发现病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]第 " & rsPatiS!主页ID & " 次住院的期间" & Format(rsPatiS!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiS!出院日期), Curdate, rsPatiS!出院日期), "yyyy-MM-dd") & vbCrLf & _
                            "与病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]的第 " & rsPatiO!主页ID & " 次住院的期间" & Format(rsPatiO!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiO!出院日期), Curdate, rsPatiO!出院日期), "yyyy-MM-dd") & _
                            vbCrLf & "互相交叉，不能进行合并！", _
                            vbInformation, gstrSysName
                            Exit Sub
                        End If
                        rsPatiO.MoveNext
                    Next
                    rsPatiS.MoveNext
                Next
            End If

            '合并原因
            str合并原因 = "[系统原因]门诊预约入院病人需要进行新旧档案合并。"

            Screen.MousePointer = 11
            DoEvents
            On Error GoTo errHandle
            strSQL = "zl_病人信息_MERGE(" & Val(rsPatiS!病人ID) & "," & Val(rsPatiO!病人ID) & ",'" & str合并原因 & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
            Screen.MousePointer = 0

            '合并后应只剩一个病人
            strSQL = "Select 病人ID From 病人信息 Where 病人ID IN([1],[2])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsPatiS!病人ID), Val(rsPatiO!病人ID))

            mlng病人ID = rsTmp!病人ID
            txtPatient.Locked = False
            txtPatient.Text = "-" & mlng病人ID
            Call txtPatient_KeyPress(13)
            RestoreYB
        Else
            If bytMode = 1 Then
                txt住院号.Locked = (InStr(mstrPrivs, "修改住院号") = 0)
                If Not mrsInfo Is Nothing Then
                    txt住院号.Text = IIf(Nvl(mrsInfo!住院号) = "", zlDatabase.GetNextNo(2), Nvl(mrsInfo!住院号))
                ElseIf Not mrsPatiReg Is Nothing Then
                    txt住院号.Text = IIf(Nvl(mrsPatiReg!住院号) = "", zlDatabase.GetNextNo(2), Nvl(mrsPatiReg!住院号))
                Else
                    txt住院号.Text = zlDatabase.GetNextNo(2)
                End If
            End If
        End If
    Else
        If bytMode = 1 Then
            MsgBox "您输入的住院号已被病人【" & rsTmp!姓名 & "】占用！", vbInformation, gstrSysName
            txt住院号.Locked = (InStr(mstrPrivs, "修改住院号") = 0)
            If Not mrsInfo Is Nothing Then
                txt住院号.Text = IIf(Nvl(mrsInfo!住院号) = "", zlDatabase.GetNextNo(2), Nvl(mrsInfo!住院号))
            ElseIf Not mrsPatiReg Is Nothing Then
                txt住院号.Text = IIf(Nvl(mrsPatiReg!住院号) = "", zlDatabase.GetNextNo(2), Nvl(mrsPatiReg!住院号))
            Else
                txt住院号.Text = zlDatabase.GetNextNo(2)
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

Private Function isValid(ByVal lng病人ID As Long) As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select 病人ID,主页ID,病人性质,入院日期,出院日期 From 病案主页 Where 病人ID=[1] And 主页ID>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    While Not rsTmp.EOF
        If Nvl(rsTmp!病人性质, 0) = 1 And Not IsNull(rsTmp!入院日期) And IsNull(rsTmp!出院日期) Then
            MsgBox "该门诊留观病人尚未出院，不允许接收预约！", vbInformation, gstrSysName
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
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
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

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
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
'    '功能:创建或关闭结算卡对象
'    '入参:blnClosed:关闭对象
'    '编制:刘兴洪
'    '日期:2010-01-05 14:51:23
'    '问题:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strExpend As String
'    '0=新增,1=修改,2=查看
'   If mbytInState = 2 Then Exit Sub
'
'    '只有:执行或退费时,才可能管结算卡的
'    If blnClosed Then
'       If Not mobjSquareCard Is Nothing Then
'            Call mobjSquareCard.CloseWindows
'            Set mobjSquareCard = Nothing
'        End If
'        Exit Sub
'    End If
'    '创建对象
'    '刘兴洪:增加结算卡的结算:执行或退费时
'    Err = 0: On Error Resume Next
'    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
'    If Err <> 0 Then
'        Err = 0: On Error GoTo 0:      Exit Sub
'    End If
'
'    '安装了结算卡的部件
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    '功能:zlInitComponents (初始化接口部件)
'    '    ByVal frmMain As Object, _
'    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
'    '        ByVal cnOracle As ADODB.Connection, _
'    '        Optional blnDeviceSet As Boolean = False, _
'    '        Optional strExpand As String
'    '出参:
'    '返回:   True:调用成功,False:调用失败
'    '编制:刘兴洪
'    '日期:2009-12-15 15:16:22
'    'HIS调用说明.
'    '   1.进入门诊收费时调用本接口
'    '   2.进入住院结帐时调用本接口
'    '   3.进入预交款时
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    If mobjSquareCard.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
'         '初始部件不成功,则作为不存在处理
'         Exit Sub
'    End If
'End Sub


Private Sub InitSendCardPreperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化刷卡属性
    '编制:刘兴洪
    '日期:2011-07-25 11:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strSQL As String, blnBoundCard As Boolean
    Dim rsTemp As ADODB.Recordset, str批次 As String, varData As Variant, i As Long
    Dim varTemp  As Variant
    Dim blnNotBind As Boolean
    On Error GoTo errHandle
    
    Set mCurSendCard.rs卡费 = Nothing
    
    If gbln入院发卡 = False Then
'        fra磁卡.Visible = False
'        Me.Height = Me.Height - fra磁卡.Height
        Exit Sub
    End If
    '76824，李南春，2014/8/19，医疗卡发卡处理
    '85565:李南春,2015/7/27,读卡性质
     lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, 0))
     If lngCardTypeID = 0 Then mCurSendCard.lng卡类别ID = 0: Exit Sub
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strSQL = "" & _
    "   Select Id, 编码, 名称, 短名, 前缀文本, 卡号长度, 缺省标志, 是否固定, 是否严格控制, " & _
    "           nvl(是否自制,0) as 是否自制, nvl(是否存在帐户,0) as 是否存在帐户, " & _
    "           nvl(是否全退,0) as 是否全退,nvl(是否重复使用,0) as 是否重复使用 , " & _
    "           nvl(密码长度,10) as 密码长度,nvl(密码长度限制,0) as 密码长度限制,nvl(密码规则,0) as 密码规则," & _
    "           nvl(是否退现,0) as 是否退现,部件, 备注, 特定项目, 结算方式, 是否启用, 卡号密文," & _
    "           nvl(是否发卡,0) as 是否发卡,nvl(是否制卡,0) as 是否制卡,nvl(是否写卡,0) as 是否写卡, " & _
    "           nvl(发卡性质,0) as 发卡性质,nvl(读卡性质,0) as 读卡性质,nvl(发卡控制,0) as 发卡控制 " & _
    "    From 医疗卡类别 A" & _
    "    Where nvl(是否启用,0)=1 And (ID=[1] or nvl(缺省标志,0)=1)" & _
    "    Order by 编码"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardTypeID)
    If rsTemp.EOF Then mCurSendCard.lng卡类别ID = 0: Exit Sub
    If rsTemp.RecordCount >= 2 Then
        rsTemp.Filter = "ID=" & lngCardTypeID
        If rsTemp.EOF Then rsTemp.Filter = 0
    End If
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
        With mCurSendCard
            .lng卡类别ID = Val(Nvl(rsTemp!ID))
            .str卡名称 = Nvl(rsTemp!名称)
            .lng卡号长度 = Val(Nvl(rsTemp!卡号长度))
            .lng结算方式 = Trim(Nvl(rsTemp!结算方式))
            .bln自制卡 = Val(Nvl(rsTemp!是否自制)) = 1
            .bln严格控制 = Val(Nvl(rsTemp!是否严格控制)) = 1
            .bln重复利用 = Val(Nvl(rsTemp!是否重复使用)) = 1
            .str卡号密文 = Nvl(rsTemp!卡号密文)
            .int密码长度 = Val(Nvl(rsTemp!密码长度))
            .int密码长度限制 = Val(Nvl(rsTemp!密码长度限制))
            .int密码规则 = Val(Nvl(rsTemp!密码规则))
            .bln就诊卡 = .str卡名称 = "就诊卡" And Val(Nvl(rsTemp!是否固定)) = 1
            '问题号:56599
            .bln是否制卡 = Val(Nvl(rsTemp!是否制卡)) = 1
            .bln是否发卡 = Val(Nvl(rsTemp!是否发卡)) = 1
            .bln是否写卡 = Val(Nvl(rsTemp!是否写卡)) = 1
            .bln是否院外发卡 = (InStr(mstrPrivs, ";发卡事务;") > 0 And .bln自制卡 = False And .bln是否发卡 = True) '问题号:56599
            .lng发卡性质 = Val(Nvl(rsTemp!发卡性质)) '问题号:57326
            .str读卡性质 = Nvl(rsTemp!读卡性质, "1000")
            .byt发卡控制 = Val(Nvl(rsTemp!发卡控制))
            '76824，李南春，2014/8/19，医疗卡发卡处理
            lbl卡名称.Caption = .str卡名称
            lbl卡名称.width = LenB(lbl卡名称.Caption) * 120
            .blnOneCard = False
            .str特定项目 = Trim(Nvl(rsTemp!特定项目))
            If .str特定项目 <> "" Then
                Set .rs卡费 = zlGetSpecialItemFee(.str特定项目, mstrPriceGrade)
                If .bln就诊卡 Then .blnOneCard = GetOneCard.RecordCount > 0
            Else
                Set .rs卡费 = Nothing
            End If
            str批次 = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModul, "0")
            '领用ID,卡类别ID|...
             .lng共用批次 = 0
            varData = Split(str批次, "|")
            For i = 0 To UBound(varData)
                 varTemp = Split(varData(i), ",")
                 If Val(varTemp(0)) <> 0 Then
                    If Val(varTemp(1)) = .lng卡类别ID Then
                        .lng共用批次 = Val(varTemp(0)): Exit For
                    End If
                 End If
            Next
           txt卡号.PasswordChar = IIf(.str卡号密文 <> "", "*", "")
           txt卡号.MaxLength = .lng卡号长度
        End With
    End If
    
    If mCurSendCard.rs卡费 Is Nothing Then
        tabCardMode.Tabs.Remove ("CardFee")
        blnBoundCard = InStr(mstrPrivs, ";绑定卡号;") > 0
        '无绑定卡权限
        pic磁卡.Visible = blnBoundCard
        If Not blnBoundCard Then
            Me.Height = Me.Height - pic磁卡.Height
        Else
            tabCardMode.Tabs("CardBind").Selected = True
            tabCardMode.Tabs("CardBind").Caption = "绑定卡号"
            tabCardMode.width = tabCardMode.width / 2
        End If
        Exit Sub
    End If
    
    blnNotBind = mCurSendCard.bln自制卡 And (Not mCurSendCard.bln重复利用 Or mCurSendCard.bln严格控制)
    
    Call LoadCardFee
    
    '如果没有绑定卡权限,加载窗体时,已经移除了绑定卡号
    blnBoundCard = Not InStr(mstrPrivs, ";绑定卡号;") > 0
    If Not blnBoundCard Then
        If zlDatabase.GetPara("发卡模式", glngSys, mlngModul, "CardFee") = "CardFee" Then
            tabCardMode.Tabs("CardFee").Selected = True
        ElseIf Not blnNotBind Then
            tabCardMode.Tabs("CardBind").Selected = True
        End If
    End If
    
 
    '绑定卡,如果没有权限在在窗体加载时,便已经删除
    '问题号:56599
    If (mCurSendCard.bln是否院外发卡 Or blnNotBind) And Not blnBoundCard Then
        tabCardMode.Tabs.Remove ("CardBind")
        If tabCardMode.Tabs.Count > 0 Then
            tabCardMode.Tabs("CardFee").Selected = True
            tabCardMode.Tabs("CardFee").Caption = "收费发卡"
            tabCardMode.width = tabCardMode.width / 2
        Else
            pic磁卡.Visible = False
            Me.Height = Me.Height - pic磁卡.Height
        End If
    ElseIf mCurSendCard.bln自制卡 = False And mCurSendCard.bln是否发卡 = False Then
        tabCardMode.Tabs.Remove ("CardFee")
        If tabCardMode.Tabs.Count > 0 Then
            tabCardMode.Tabs("CardBind").Selected = True
            tabCardMode.Tabs("CardBind").Caption = "绑定卡号"
            tabCardMode.width = tabCardMode.width / 2
        Else
            pic磁卡.Visible = False
            Me.Height = Me.Height - pic磁卡.Height
        End If
    End If
        
    If mCurSendCard.bln严格控制 Then
        '就诊卡领用检查
        mCurSendCard.lng领用ID = CheckUsedBill(5, IIf(mCurSendCard.lng领用ID > 0, mCurSendCard.lng领用ID, mCurSendCard.lng共用批次), , mCurSendCard.lng卡类别ID)
        If mCurSendCard.lng领用ID <= 0 Then
            Select Case mCurSendCard.lng领用ID
                Case 0 '操作失败
                Case -1
'                    MsgBox "你没有自用或共用的" & mCurSendCard.str卡名称 & ",不能发放！" & vbCrLf & _
'                        "请先在本地设置共用批次或领用一批新卡! ", vbExclamation, gstrSysName
                Case -2
'                    MsgBox "本地共用的" & mCurSendCard.str卡名称 & "已用完,不能发放！" & vbCrLf & _
'                        "请重新设置本地共用卡批次或领用一批新卡！", vbExclamation, gstrSysName
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
    '功能:获取不同类别的发票
    '编制:刘兴洪
    '日期:2011-07-19 17:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gbln入院预交 = False Then Exit Sub
    
    If gblnPrepayStrict = False Then
        '松散：取下一个号码
        txtFact.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("当前预交票据号", glngSys, mlngModul, "")))
        Exit Sub
    End If
    '严格:     取下一个号码
    mlng预交领用ID = CheckUsedBill(2, IIf(mlng预交领用ID > 0, mlng预交领用ID, mFactProperty.lngShareUseID), , 2)
    If mlng预交领用ID <= 0 Then
        Select Case mlng预交领用ID
            Case 0 '操作失败
'            Case -1
'                MsgBox "你没有自用或共用的预交票据,登记病人信息时不能同时缴预交款！" & _
'                    "请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
'            Case -2
'                MsgBox "本地的共用票据已经用完,登记病人信息时不能同时缴预交款！" & _
'                    "请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
        End Select
    Else
        txtFact.Text = GetNextBill(mlng预交领用ID)
    End If
End Sub
Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, strTemp As String
    Dim str缺省预交款方式 As String, blnHavePrivs As Boolean
    
    On Error GoTo errHandle
    
    '结算方式:费用查询和医疗卡调用时，一般只支付预交款,不存在代收的情况
    If mbytMode = 1 Then
        strTemp = "1,2,7,8" '预约登记时
    Else
        strTemp = "1,2,5,7,8" & IIf(InStr(mstrPrivs, ";保险病人登记;") > 0, ",3", "")
    End If
    
    strSQL = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合 ='预交款'  And B.名称=A.结算方式  " & _
        "           And Nvl(B.性质,1) In(" & strTemp & ")" & _
        " Order by B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolPrepayPayMode = New Collection
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType: varData = Split(strPayType, ";")
    With cbo预交结算
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            
            If Not blnFind And InStr(",7,8,", "," & Nvl(rsTemp!性质) & ",") = 0 Then
                .AddItem Nvl(rsTemp!名称)
                mcolPrepayPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
                If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                'If mstr缺省结算方式 = Nvl(rsTemp!名称) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            '结算方式中设置且设备配置启用了的结算方式才有效
            rsTemp.Filter = "名称 ='" & Split(varData(i), "|")(6) & "'"
            If Not rsTemp.EOF Then
                If InStr(1, varData(i), "|") <> 0 Then
                    varTemp = Split(varData(i), "|")
                    mcolPrepayPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    'If mstr缺省结算方式 = varTemp(1) Then .ListIndex = .NewIndex: cboStyle.Tag = cboStyle.NewIndex
                    j = j + 1
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    If cbo预交结算.ListCount = 0 Then
        MsgBox "预交场合没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        mblnUnload = True: Exit Sub
    End If
    '问题号:48352
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    str缺省预交款方式 = zlDatabase.GetPara("缺省缴款方式", glngSys, mlngModul, , blnHavePrivs)
    If str缺省预交款方式 <> "" Then
        For i = 0 To cbo预交结算.ListCount
            If cbo预交结算.List(i) = str缺省预交款方式 Then
                cbo预交结算.ListIndex = i
            End If
        Next
    End If
    
    
    strSQL = _
    "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
    " From 结算方式应用 A,结算方式 B" & _
    " Where A.应用场合 ='就诊卡'  And B.名称=A.结算方式  " & _
    "           And Nvl(B.性质,1) In(1,2,7,8)" & _
    " Order by B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolCardPayMode = New Collection
    With cbo发卡结算
        mblnNotClick = True
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            
            If Not blnFind And InStr(",7,8,", "," & Nvl(rsTemp!性质) & ",") = 0 Then
                .AddItem Nvl(rsTemp!名称)
                mcolCardPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
                If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                 If cbo发卡结算.List(j) = str缺省预交款方式 Then
                    cbo发卡结算.ListIndex = j:  .Tag = j
                 End If
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            '结算方式中设置且设备配置启用了的结算方式才有效
            rsTemp.Filter = "名称 ='" & Split(varData(i), "|")(6) & "'"
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


Private Sub Local结算方式(ByVal lng卡类别ID As Long, Optional bln预交 As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定位结算方式
    '编制:刘兴洪
    '日期:2011-07-26 15:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPayMoney As Collection, cboPay As ComboBox
    Dim i As Long
    If mblnNotClick Then Exit Sub
    If bln预交 Then
       Set cllPayMoney = mcolPrepayPayMode
        Set cboPay = cbo预交结算
    Else
       Set cllPayMoney = mcolCardPayMode
        Set cboPay = cbo发卡结算
    End If
    If cllPayMoney Is Nothing Then Exit Sub
    With cboPay
        If .ListIndex >= 0 Then
            If bln预交 Then
                If .ItemData(.ListIndex) >= 0 Then Exit Sub
            End If
        End If
        mblnNotClick = True
        For i = 0 To .ListCount - 1
            ''短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
            If Val(cllPayMoney(i + 1)(3)) = lng卡类别ID Then
                .ListIndex = i: Exit For
            End If
        Next
        mblnNotClick = False
    End With
End Sub
Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        '58322
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
        .ActiveConnection = Nothing
        If mCurPrepay.lng医疗卡类别ID <> 0 Then
            .AddNew
            !收费类别 = "预交"
            !金额 = StrToNum(txt预交额.Text)
            .Update
        End If
        If mCurCardPay.lng医疗卡类别ID <> 0 And Trim(txt卡号) <> "" _
            And cbo发卡结算.Enabled And cbo发卡结算.Visible Then
            .AddNew
            !收费类别 = mCurSendCard.rs卡费!收费类别
            !金额 = StrToNum(txt卡额.Text)
            .Update
        End If
    End With
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AddCardDataSQL(ByVal lng病人ID As Long, ByVal lng主页ID As Long, lng病区ID As Long, lng科室ID As Long, ByVal dtCurdate As Date, ByRef strOutSQL As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:就诊卡发放处理
    '入参:lng病人ID
    '编制:刘兴洪
    '日期:2011-07-07 04:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt操作类型 As Byte, strno As String, strPassWord As String, strSQL As String
    Dim str原卡号 As String, str年龄 As String, strCard As String, str变动原因 As String
    Dim strICCard As String, lngBrushCardTypeID As Long, str结算方式 As String, strBrushCardNo As String
    Dim bln消费卡 As Boolean, blnInRange As Boolean   '范围内的卡
    Dim lngIndex As Long, byt变动类型 As Byte, lng结帐ID As Long
    
    strCard = UCase(txt卡号.Text): strICCard = IIf(mblnICCard, strCard, "")
    If Not ((strCard <> "" Or strICCard <> "") And pic磁卡.Visible = True) Then Exit Sub
    
    '问题号:56599
    mbln发卡或绑定卡 = True
     
    lng结帐ID = 0: blnInRange = True
    If mCurSendCard.blnOneCard And mCurSendCard.bln严格控制 Then blnInRange = mCurSendCard.lng领用ID > 0
    
    If blnInRange And tabCardMode.SelectedItem.Key = "CardFee" Then
        blnInRange = True
        byt操作类型 = 0: byt变动类型 = 1
    Else
        blnInRange = False
        byt变动类型 = 11: byt操作类型 = 0
    End If
    str变动原因 = "病人入院登记发卡"
    strPassWord = zlCommFun.zlStringEncode(Trim(txtPass.Text))
    If blnInRange = False Then
          'Zl_医疗卡变动_Insert
           strSQL = "Zl_医疗卡变动_Insert("
          '      变动类型_In   Number,
          '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
          strSQL = strSQL & "" & byt变动类型 & ","
          '      病人id_In     住院费用记录.病人id%Type,
          strSQL = strSQL & "" & lng病人ID & ","
          '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
          strSQL = strSQL & "" & mCurSendCard.lng卡类别ID & ","
          '      原卡号_In     病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & str原卡号 & "',"
          '      医疗卡号_In   病人医疗卡信息.卡号%Type,
          strSQL = strSQL & "'" & strCard & "',"
          '      变动原因_In   病人医疗卡变动.变动原因%Type,
          '      --变动原因_In:如果密码调整，变动原因为密码.加密的
          strSQL = strSQL & "'" & str变动原因 & "',"
          '      密码_In       病人信息.卡验证码%Type,
          strSQL = strSQL & "'" & strPassWord & "',"
          '      操作员姓名_In 住院费用记录.操作员姓名%Type,
          strSQL = strSQL & "'" & UserInfo.姓名 & "',"
          '      变动时间_In   住院费用记录.登记时间%Type,
          strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
          strSQL = strSQL & "'" & strICCard & "',"
          '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
          strSQL = strSQL & "NULL)"
    Else
        '103980:李南春,2017/1/19,保存发卡病人年龄
        str年龄 = Trim(txt年龄.Text)
        If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text

        strno = zlDatabase.GetNextNo(16)  '医疗卡
        If chk记帐.Value = 0 Then
            lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
        End If
        mCurCardPay.strno = strno
        mCurCardPay.lng结帐ID = lng结帐ID
        strSQL = zlGetSaveCardFeeSQL(mCurSendCard.lng卡类别ID, byt操作类型, strno, lng病人ID, lng主页ID, lng病区ID, lng科室ID, Val(txt住院号.Text), _
         zlCommFun.GetNeedName(cbo费别.Text), "", Trim(txt姓名.Text), zlCommFun.GetNeedName(cbo性别.Text), str年龄, _
        strCard, strPassWord, str变动原因, IIf(mCurSendCard.bln变价 = False, mCurSendCard.dbl应收金额, Val(txt卡额.Text)), Val(txt卡额.Text), IIf(chk记帐.Value = 0, mCurCardPay.str结算方式, ""), _
        dtCurdate, mCurSendCard.lng领用ID, mCurSendCard.rs卡费, strICCard, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, lng结帐ID)
    End If
    strOutSQL = strSQL
 End Sub
 
 Private Sub AddDepositSQL(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, ByVal dtDate As Date, ByRef bln个人帐户缴预交 As Boolean, strOutSQL As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加预交款的SQL
    '编制:刘兴洪
    '日期:2011-07-26 18:26:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strno As String, strSQL As String, i As Integer, lng预交ID As Long
    Dim dblMoney As Double
     If Not (IsNumeric(txt预交额.Text) And fra预交.Visible) Then Exit Sub
     
    '病人预交款记录
    strno = zlDatabase.GetNextNo(11)
    lng预交ID = zlDatabase.GetNextId("病人预交记录")
    mCurPrepay.strno = strno
    mCurPrepay.lngID = lng预交ID
    dblMoney = StrToNum(txt预交额.Text)
    bln个人帐户缴预交 = is个人帐户(cbo预交结算) And mintInsure <> 0 And mstrYBPati <> "" And mbytMode <> 1
    
    'Zl_病人预交记录_Insert
    strSQL = "Zl_病人预交记录_Insert("
    '  Id_In         病人预交记录.ID%Type,
    strSQL = strSQL & "" & lng预交ID & ","
    '  单据号_In     病人预交记录.NO%Type,
    strSQL = strSQL & "'" & strno & "',"
    '  票据号_In     票据使用明细.号码%Type,
    strSQL = strSQL & "" & IIf(mblnPrepayPrint, "'" & txtFact.Text & "'", "Null") & ","
    '  病人id_In     病人预交记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  主页id_In     病人预交记录.主页id%Type,
    strSQL = strSQL & "" & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
    '  科室id_In     病人预交记录.科室id%Type,
    strSQL = strSQL & "" & IIf(lng科室ID = 0, "NULL", lng科室ID) & ","
    '  金额_In       病人预交记录.金额%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "'" & mCurPrepay.str结算方式 & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSQL = strSQL & "'" & txt结算号码.Text & "',"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    If bln个人帐户缴预交 Then
        strSQL = strSQL & "'" & mintInsure & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt缴款单位.Text) & "',"
    End If
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    If bln个人帐户缴预交 Then
        strSQL = strSQL & "'" & Split(mstrYBPati, ";")(2) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt开户行.Text) & "',"
    End If
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    If bln个人帐户缴预交 Then
        strSQL = strSQL & "'" & Split(mstrYBPati, ";")(1) & "',"
    Else
        strSQL = strSQL & "'" & Trim(txt帐号.Text) & "',"
    End If
    '  摘要_In       病人预交记录.摘要%Type,
    strSQL = strSQL & "'入院预交',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & IIf(mlng预交领用ID = 0, "NULL", mlng预交领用ID) & ","
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strSQL = strSQL & "" & 2 & ","
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPrepay.lng医疗卡类别ID = 0 Or mCurPrepay.bln消费卡, "NULL", mCurPrepay.lng医疗卡类别ID) & ","
   '  结算卡序号_in 病人预交记录.结算卡序号%type:=NULL,
    strSQL = strSQL & "" & IIf(mCurPrepay.lng医疗卡类别ID = 0 Or Not mCurPrepay.bln消费卡, "NULL", mCurPrepay.lng医疗卡类别ID) & ","
    '  卡号_In       病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "" & IIf(mCurPrepay.str刷卡卡号 = "", "NULL", "'" & mCurPrepay.str刷卡卡号 & "'") & ","
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '  合作单位_In   病人预交记录.合作单位%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  收款时间_In   病人预交记录.收款时间%Type := Null
    '108001:李南春，2017/5/8，格式化预交时间为24小时制
    strSQL = strSQL & "to_date('" & Format(dtDate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '   操作类型_In Integer:=0 :0-正常缴预交;1-存为划价单
    strSQL = strSQL & "0 )"
   strOutSQL = strSQL
End Sub
Private Function CheckBrushCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset, str年龄 As String
    Dim dblMoney As Double, bln三方结算 As Boolean
    Dim dblThreeMoney As Double, tyCurThreePay As Ty_PayMoney
    
    Dim str刷卡卡号 As String, str刷卡密码 As String
    Dim blnTemp As Boolean
    
    On Error GoTo errHandle
    
    dblMoney = 0: dblThreeMoney = 0
    '58322
    If cbo预交结算.Visible Then
        If cbo预交结算.ListIndex >= 0 Then
            bln三方结算 = cbo预交结算.ItemData(cbo预交结算.ListIndex) = -1
            If bln三方结算 Then dblThreeMoney = dblThreeMoney + StrToNum(txt预交额.Text)
        End If
        dblMoney = dblMoney + StrToNum(txt预交额.Text)
    End If
    If cbo发卡结算.Visible And cbo发卡结算.Enabled And Trim(txt卡号) <> "" Then
        If cbo发卡结算.ListIndex >= 0 Then
            blnTemp = cbo发卡结算.ItemData(cbo发卡结算.ListIndex) = -1
            If blnTemp Then dblThreeMoney = dblThreeMoney + StrToNum(txt卡额.Text)
            bln三方结算 = bln三方结算 Or blnTemp
        End If
        dblMoney = dblMoney + StrToNum(txt卡额.Text)
    End If
    If Not bln三方结算 Then CheckBrushCard = True: Exit Function
    If mCurPrepay.lng医疗卡类别ID <> 0 Then
       tyCurThreePay = mCurPrepay
    Else
       tyCurThreePay = mCurCardPay
    End If
    
    If (mCurPrepay.lng医疗卡类别ID <> mCurCardPay.lng医疗卡类别ID Or _
        mCurPrepay.bln消费卡 <> mCurCardPay.bln消费卡) _
        And mCurCardPay.lng医疗卡类别ID <> 0 And mCurPrepay.lng医疗卡类别ID <> 0 Then
        MsgBox "不能同时使用两种不同类别的支付方式,不能继续!", vbOKOnly + vbInformation, gstrSysName
        If cbo预交结算.Enabled And cbo预交结算.Visible Then cbo预交结算.SetFocus: Exit Function
        If cbo发卡结算.Enabled And cbo发卡结算.Visible Then cbo发卡结算.SetFocus
        Exit Function
    End If
    Call zlGetClassMoney(rsMoney)
    
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String) As Boolean
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
   '58322
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, tyCurThreePay.lng医疗卡类别ID, tyCurThreePay.bln消费卡, _
    txt姓名.Text, zlCommFun.GetNeedName(cbo性别.Text), str年龄, dblThreeMoney, tyCurThreePay.str刷卡卡号, tyCurThreePay.str刷卡密码, False, True, False) = False Then Exit Function
    
    '保存前,一些数据检查
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, tyCurThreePay.lng医疗卡类别ID, _
    tyCurThreePay.bln消费卡, tyCurThreePay.str刷卡卡号, dblThreeMoney, "", "") = False Then Exit Function
    mCurCardPay.str刷卡卡号 = tyCurThreePay.str刷卡卡号
    mCurCardPay.str刷卡密码 = tyCurThreePay.str刷卡密码
    mCurPrepay.str刷卡卡号 = tyCurThreePay.str刷卡卡号
    mCurPrepay.str刷卡密码 = tyCurThreePay.str刷卡密码
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlInterfacePrayMoney(ByRef cllPro As Collection, ByRef cllThreeSwap As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim dblMoney As Double
    If mCurCardPay.lng医疗卡类别ID = 0 And mCurPrepay.lng医疗卡类别ID = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cbo发卡结算.ItemData(cbo发卡结算.ListIndex) <> -1 _
        And cbo预交结算.ItemData(cbo预交结算.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng卡类别ID As Long, bln消费卡 As Boolean, strCardNO As String
    
    dblMoney = 0
    If mCurCardPay.lng医疗卡类别ID <> 0 And cbo发卡结算.Enabled And cbo发卡结算.Visible Then
        dblMoney = Val(txt卡额.Text)
        lng卡类别ID = mCurCardPay.lng医疗卡类别ID
        bln消费卡 = mCurCardPay.bln消费卡
        strCardNO = mCurCardPay.str刷卡卡号
    End If
    If mCurPrepay.lng医疗卡类别ID <> 0 And cbo预交结算.Visible Then
        dblMoney = dblMoney + StrToNum(txt预交额.Text)
        If lng卡类别ID <> mCurPrepay.lng医疗卡类别ID And lng卡类别ID <> 0 Then
            MsgBox "发卡所选择的支付方式与预交款所选择的支付方式不一致!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        lng卡类别ID = mCurPrepay.lng医疗卡类别ID
        bln消费卡 = mCurPrepay.bln消费卡
        strCardNO = mCurPrepay.str刷卡卡号
    End If
    If lng卡类别ID = 0 Then Exit Function


    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, lng卡类别ID, bln消费卡, strCardNO, mCurCardPay.lng结帐ID, mCurPrepay.strno, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '更新三交交易数据
     If mCurCardPay.lng医疗卡类别ID <> 0 And mCurCardPay.lng结帐ID <> 0 And cbo发卡结算.Visible Then
        If Not mCurCardPay.bln消费卡 Then
            Call zlAddUpdateSwapSQL(False, mCurCardPay.lng结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, mCurCardPay.lng结帐ID, mCurCardPay.lng医疗卡类别ID, mCurCardPay.bln消费卡, mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    If mCurPrepay.lng医疗卡类别ID <> 0 And cbo预交结算.Visible And StrToNum(txt预交额.Text) <> 0 Then
        Call zlAddUpdateSwapSQL(True, mCurPrepay.lngID, mCurPrepay.lng医疗卡类别ID, mCurPrepay.bln消费卡, mCurPrepay.str刷卡卡号, strSwapGlideNO, strSwapMemo, cllPro)
        Call zlAddThreeSwapSQLToCollection(True, mCurPrepay.lngID, mCurPrepay.lng医疗卡类别ID, mCurPrepay.bln消费卡, mCurPrepay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Led欢迎信息()
    Dim strInfo As String, lngPatient As Long
    'LED初始化
    If gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.编号 & "号为您服务", mlngModul, gcnOracle
        End If
        strInfo = Trim(txtPatient.Text)
        If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!性别 & " " & mrsInfo!年龄: lngPatient = Val("" & mrsInfo!病人ID)
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub

Private Function zl_Get设置默认发卡密码() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置默认发卡密码
    '返回:是否继续发卡操作
    '编制:王吉
    '日期:2012-07-06 15:53:14
    '问题号:51072
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardType As clsCard
    Dim msgResult As VbMsgBoxResult
    Dim arr() As String
    arr = zl_Get医疗卡类型(mCurSendCard.lng卡类别ID)
    If Val(arr(2)) = 0 Then '无限制
        Select Case Val(arr(1))
            Case 0 '无限制
                zl_Get设置默认发卡密码 = True
                Exit Function
            Case 1 '未输入提醒
               msgResult = MsgBox("未输入密码将会影响帐户的使用安全,是否继续？", vbQuestion + vbYesNo, gstrSysName)
               zl_Get设置默认发卡密码 = IIf(msgResult = vbYes, True, False)
               Exit Function
            Case 2 '为输入禁止
                 MsgBox "未输入卡密码,不能进行发卡！", vbExclamation, gstrSysName
                zl_Get设置默认发卡密码 = False
                Exit Function
        End Select
    ElseIf Val(arr(2)) = 1 Then '缺省身份证后N位
        If Len(Trim(txt身份证号.Text)) > 0 Or Len(Trim(txt联系人身份证号.Text)) > 0 Then '输入了身份证或联系人身份证号
            If Len(Trim(txt身份证号.Text)) > 0 Then '有身份证优先用身份证
                   txtPass.Text = Right(Trim(txt身份证号.Text), Val(arr(0)))
            Else '否则就用代办人身份证作为密码
                   txtPass.Text = Right(Trim(txt联系人身份证号.Text), Val(arr(0)))
            End If
        Else '身份证与联系人身份证都没输入
            Select Case Val(arr(1))
                Case 0 '无限制
                    zl_Get设置默认发卡密码 = True
                    Exit Function
                Case 1 '未输入提醒
                    msgResult = MsgBox("未输入密码将会影响帐户的使用安全,是否继续？", vbQuestion + vbYesNo, gstrSysName)
                    zl_Get设置默认发卡密码 = IIf(msgResult = vbYes, True, False)
                    Exit Function
                Case 2 '为输入禁止
                    MsgBox "未输入卡密码,不能进行发卡！", vbExclamation, gstrSysName
                    zl_Get设置默认发卡密码 = False
                    Exit Function
            End Select
        End If
    End If
    zl_Get设置默认发卡密码 = True
End Function
Private Function zl绑定身份证(colPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置默认发卡密码
    '返回:是否继续发卡操作
    '编制:王吉
    '日期:2012-07-06 15:53:14
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Trim(txt支付密码.Text) <> Trim(txt验证密码.Text) Then
        MsgBox "两次输入的密码不一致,请重新输入", vbOKOnly + vbInformation, gstrSysName
        txt支付密码.Text = "": txt验证密码.Text = ""
        If txt支付密码.Visible = True Then txt支付密码.SetFocus
        Exit Function
    End If
    If Trim(txt支付密码.Text) <> "" Then
       If 是否已经签约(Trim(txt身份证号.Text)) Then
             MsgBox "身份证号码为:" & txt身份证号.Text & "已经签约不能重复签约！", vbOKOnly + vbInformation, gstrSysName
             txt支付密码.Text = "": txt验证密码.Text = ""
             If txt支付密码.Visible = True Then txt支付密码.SetFocus
             Exit Function
       End If
    End If
    AddSQL绑定卡 Trim(txtPatient.Text), Get医疗卡类别ID("二代身份证"), Trim(txt身份证号.Text), zlCommFun.zlStringEncode(Trim(txt支付密码.Text)), zlDatabase.Currentdate, False, colPro
    
    zl绑定身份证 = True
End Function
Private Sub InitTabPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化分页控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
        
    Set objItem = tbcPage.InsertItem(1, "基本", PicBaseInfo.hWnd, 0)
    objItem.Tag = mPageHeight.基本
    
    Set objItem = tbcPage.InsertItem(2, "健康档案", PicHealth.hWnd, 0)
    objItem.Tag = mPageHeight.健康档案
    
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
    '功能：初始vsFlexGrid
    '           有一固定行，初始化后，只有一行记录，无固定列。
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'vsGrid:    要初始化的控件

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
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               '为了支持zl9PrintMode
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
                    '为了支持zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  '为了支持zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
                End If
            End If
        Next
        
        '固定行文字居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
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
    '功能:初始化ComBox控件
    '编制:56599
    '日期:2012-12-07 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '68192:刘鹏飞,2013-12-02,血型读取数据字典、RH缺省默认值为空
    Call ReadDict("血型", cboBloodType)
    ComboBox cboBH, C_BH
    If cboBH.ListCount > 0 Then cboBH.ListIndex = -1
End Sub

Private Sub InitVsOtherInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo ErrHand
    
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 社会关系 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "社会关系")
    With rsTemp
        Do While Not rsTemp.EOF
            strTmp = strTmp & "|" & Nvl(rsTemp!名称)
        rsTemp.MoveNext
        Loop
    End With
    If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
    
    With vsLinkMan
        '初始化列表属性
        SetColumHeader vsLinkMan, C_LinkManColumHeader
        .Editable = IIf(mbytInState = 2, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
        If strTmp <> "" Then .ColComboList(.ColIndex("联系人关系")) = strTmp
    End With
    
    With vsOtherInfo
        '设置列头
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
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsDrug
        '初始化列表属性
        SetColumHeader vsDrug, C_ColumHeader
        .Editable = IIf(mbytInState = 2, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
        .ColComboList(0) = "..."
    End With
End Sub

Private Sub InitVsInoculate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsInoculate
        '初始化列表属性
        SetColumHeader vsInoculate, C_InoculateHeader
         vsInoculate.Editable = IIf(mbytInState = 2, flexEDNone, flexEDKbdMouse)
        '设置选择按钮
        .ColDataType(0) = flexDTDate
        .ColEditMask(0) = "####-##-##"
        .ColDataType(2) = flexDTDate
        .ColEditMask(2) = "####-##-##"
        .SelectionMode = flexSelectionFree
    End With

End Sub

Private Sub txt转入_GotFocus()
    zlControl.TxtSelAll txt转入
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt转入_KeyPress(KeyAscii As Integer)
    Dim vPoint As POINTAPI
    On Error GoTo errH
    If KeyAscii = 13 Then
        KeyAscii = 0
        vPoint = GetCoordPos(txt转入.Container.hWnd, txt转入.Left, txt转入.Top)
        Call GetSpc医疗机构(txt转入, Me, "医疗机构", False, False, False, True, vPoint)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt转入_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt转入_Validate(Cancel As Boolean)
    Dim vPoint As POINTAPI
    vPoint = GetCoordPos(txt转入.Container.hWnd, txt转入.Left, txt转入.Top)
    Call GetSpc医疗机构(txt转入, Me, "医疗机构", False, False, False, True, vPoint)
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
            strFilter = " And zlspellcode(A.名称) like [1]"
            strInput = UCase(strInput)
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strFilter = " And A.名称 like [1]"
        Else
            strFilter = " And A.编码 like [1]"
        End If
    End If
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select Distinct A.ID,A.编码," & _
        " A.名称,zlspellcode(A.名称) 简码,A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类," & _
        " Decode(B.是否新药,1,'√','') as 新药,Decode(B.是否皮试,1,'√','') as 皮试" & _
        " From 诊疗项目目录 A,药品特性 B" & _
        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & strFilter

    '获取当前鼠标坐标值
    vRect = GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    strInput = gstrLike & Trim(strInput) & "%"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "过敏药物", False, "过敏药物选择器", "请从下面的药品中选择一项作为病人过敏药物", False, False, True, vRect.Left, vRect.Top, 0, blnCancel, False, True, strInput)

    If Not rsTemp Is Nothing And blnCancel = False Then
        If rsTemp.RecordCount > 0 Then
            vsDrug.EditText = Nvl(rsTemp!名称)
            vsDrug.TextMatrix(Row, Col) = Nvl(rsTemp!名称)
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
    '问题号:56599
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
    '问题号:56599
    If Col = 1 Then  '过敏反应列编辑时需判断是否字数超过了100
        With vsDrug
           If LenB(StrConv(.TextMatrix(Row, Col), vbFromUnicode)) > 100 Then
                MsgBox "过敏反应输入字符超出最大字符数100,多出的字符将被自动截除！", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = StrConv(MidB(StrConv(.TextMatrix(Row, Col), vbFromUnicode), 1, 100), vbUnicode)
           End If
        End With
    End If
End Sub

Private Sub vsDrug_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '问题号:56599
    Dim strSQL As String
    Dim datCurr As Date
    Dim vRect As RECT
    Dim rsTemp As Recordset
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandl:
    datCurr = zlDatabase.Currentdate
    strSQL = _
        " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL 简码,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL 简码,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL 简码,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试 From Dual Union ALL" & _
        " Select ID,nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称,NULL 简码," & _
        " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 新药,NULL as 皮试" & _
        " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Union All" & _
        " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码," & _
        " A.名称,zlspellcode(A.名称) 简码,A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类," & _
        " Decode(B.是否新药,1,'√','') as 新药,Decode(B.是否皮试,1,'√','') as 皮试" & _
        " From 诊疗项目目录 A,药品特性 B" & _
        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"

    '获取当前鼠标坐标值
    vRect = GetControlRect(vsDrug.hWnd)
    vRect.Top = vRect.Top + (Row - 1) * 300 + 150
    vRect.Left = vRect.Left + 30
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "过敏药物", False, "过敏药物选择器", "请从下面的药品中选择一项作为病人过敏药物", False, False, True, vRect.Left, vRect.Top, 0, blnCancel, False, True)

    If Not rsTemp Is Nothing And blnCancel = False Then
        vsDrug.TextMatrix(Row, Col) = rsTemp!名称
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
    '问题号:56599
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
'功能：获取指定控件在屏幕中的位置(Twip)
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
    "       Select 编码 as ID,名称,简码 From 医学警示 Where 名称 Not Like '其他%'"
    Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "医学警示", False, txtMedicalWarning.Text, "", False, False, False, vRect.Left, vRect.Top - 180, 500, True, False, True)
    If Not rsTemp Is Nothing Then
        While rsTemp.EOF = False
          strTemp = strTemp & ";" & rsTemp!名称
          rsTemp.MoveNext
        Wend
    Else
        If cmdMedicalWarning.Enabled And cmdMedicalWarning.Visible Then cmdMedicalWarning.SetFocus: Exit Sub
    End If
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then txtMedicalWarning.Text = strTemp
    If txtOtherWaring.Enabled And txtOtherWaring.Visible Then txtOtherWaring.SetFocus
End Sub

Private Sub SetDrugAllergy(str过敏药物 As String, str过敏反应 As String, Optional lng过敏ID = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置过敏药物
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long

    With vsDrug
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = str过敏药物 Then
                    .TextMatrix(i, 1) = str过敏反应
                    If lng过敏ID <> 0 Then .TextMatrix(i, 2) = lng过敏ID
                    Exit Sub
                End If

            Next
        End If
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str过敏药物
        .TextMatrix(.Rows - 1, 1) = str过敏反应
        If lng过敏ID <> 0 Then .TextMatrix(.Rows - 1, 2) = lng过敏ID
        .Rows = .Rows + 1
    End With
End Sub
Private Sub SetInoculate(str接种日期 As String, str接种名称 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置接种情况
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    '68192:刘鹏飞,2013-12-02
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If Format(.TextMatrix(i, j - 1), "YYYY-MM-DD") = Format(str接种日期, "YYYY-MM-DD") Then
                        .TextMatrix(i, j) = str接种名称
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = Format(str接种日期, "YYYY-MM-DD")
                .TextMatrix(.Rows - 1, j + 1) = str接种名称
                If j = 2 Then .Rows = .Rows + 1
                Exit Sub
            End If
        Next
    End With
End Sub
Private Sub SetLinkInfo(str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, Optional str附加信息 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置联系人相关信息
    '编制:56599
    '日期:2012-12-12 09:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    With vsLinkMan
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("联系人姓名")) = str姓名 And .TextMatrix(i, .ColIndex("联系人身份证号")) = str身份证号 Then
                    .TextMatrix(i, .ColIndex("联系人关系")) = str关系: .TextMatrix(i, .ColIndex("联系人电话")) = str电话
                    If i = 1 Then
                        txt联系人身份证号.Text = str身份证号
                        txt联系人姓名.Text = str姓名
                        For j = 0 To cbo联系人关系.ListCount - 1
                            If zlCommFun.GetNeedName(cbo联系人关系.List(j)) = str关系 Then cbo联系人关系.ListIndex = j
                        Next
                        txt联系人电话.Text = str电话
                        txtLinkManInfo.Text = str附加信息
                    End If
                    Exit Sub
                End If
            Next
        End If
        
        If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("联系人姓名")) = str姓名
        .TextMatrix(.Rows - 1, .ColIndex("联系人关系")) = str关系
        .TextMatrix(.Rows - 1, .ColIndex("联系人电话")) = str电话
        .TextMatrix(.Rows - 1, .ColIndex("联系人身份证号")) = str身份证号
        .TextMatrix(.Rows - 1, .ColIndex("联系人关系备注")) = str附加信息
        If .Rows - 1 = 1 Then
            txt联系人身份证号.Text = str身份证号
            txt联系人姓名.Text = str姓名
            For j = 0 To cbo联系人关系.ListCount - 1
                If zlCommFun.GetNeedName(cbo联系人关系.List(j)) = str关系 Then cbo联系人关系.ListIndex = j
            Next
            txt联系人电话.Text = str电话
            txtLinkManInfo.Text = str附加信息
        End If
        .Rows = .Rows + 1
    End With
End Sub

Private Sub SetOtherInfo(str信息名 As String, str信息值 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置其他情况
    '编制:56599
    '日期:2012-12-11 09:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) = str信息名 Then
                        .TextMatrix(i, j + 1) = str信息值
                        Exit Sub
                    End If
                Next
            Next
        End If

        If .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
        For j = 0 To .Cols - 1 Step 2
            If .TextMatrix(.Rows - 1, j) = "" And .TextMatrix(.Rows - 1, j + 1) = "" Then
                .TextMatrix(.Rows - 1, j) = str信息名
                .TextMatrix(.Rows - 1, j + 1) = str信息值
                If j = 2 Then .Rows = .Rows + 1
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub Load健康卡相关信息(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人健康卡信息
    '入参:lng病人ID - 病人ID
    '编制:56599
    '日期:2012-12-12 14:55:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs过敏药物 As Recordset
    Dim rs免疫记录 As Recordset
    Dim rsABO血型 As Recordset
    Dim rsRH As Recordset
    Dim rs医学警示 As Recordset
    Dim rs其他医学警示 As Recordset
    Dim rs病人信息 As Recordset
    Dim rs联系人 As Recordset
    Dim rs其他信息 As Recordset
    Dim str医学警示 As String
    Dim str联系人姓名 As String
    Dim str联系人关系 As String
    Dim str联系人附加信息 As String
    Dim str联系人电话 As String
    Dim str联系人身份证号 As String
    Dim lng联系人数量 As Long
    Dim i As Long
    On Error GoTo ErrHandl:

    '获取过敏药物
    strSQL = "" & _
    "   Select 病人ID,过敏药物ID,过敏药物,过敏反应 From 病人过敏药物 Where 病人ID=[1]"
    Set rs过敏药物 = zlDatabase.OpenSQLRecord(strSQL, "病人过敏药物", lng病人ID)
    While rs过敏药物.EOF = False
        SetDrugAllergy Nvl(rs过敏药物!过敏药物), Nvl(rs过敏药物!过敏反应), Nvl(rs过敏药物!过敏药物ID, 0)
        rs过敏药物.MoveNext
    Wend
    '获取免疫记录
    strSQL = "" & _
    "   Select 病人ID,接种时间,接种名称 From 病人免疫记录 Where 病人ID=[1]"
    Set rs免疫记录 = zlDatabase.OpenSQLRecord(strSQL, "病人免疫记录", lng病人ID)
    While rs免疫记录.EOF = False
        SetInoculate Format(Nvl(rs免疫记录!接种时间), "YYYY-MM-DD"), Nvl(rs免疫记录!接种名称)
        rs免疫记录.MoveNext
    Wend
    '血型
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='血型'"
    Set rsABO血型 = zlDatabase.OpenSQLRecord(strSQL, "ABO血型", lng病人ID)
    While rsABO血型.EOF = False
        For i = 0 To cboBloodType.ListCount - 1
            If cboBloodType.List(i) = Nvl(rsABO血型!信息值) Then cboBloodType.ListIndex = i
        Next
        rsABO血型.MoveNext
    Wend
    'RH
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='RH'"
    Set rsRH = zlDatabase.OpenSQLRecord(strSQL, "RH", lng病人ID)
    While rsRH.EOF = False
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = Nvl(rsRH!信息值) Then cboBH.ListIndex = i
        Next
        rsRH.MoveNext
    Wend
    '医学警示
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='医学警示' "
    Set rs医学警示 = zlDatabase.OpenSQLRecord(strSQL, "医学警示", lng病人ID)
    While rs医学警示.EOF = False
        str医学警示 = str医学警示 & ";" & Nvl(rs医学警示!信息值)
        rs医学警示.MoveNext
    Wend
    If str医学警示 <> "" Then str医学警示 = Mid(str医学警示, 2)
    txtMedicalWarning.Text = str医学警示
    txtMedicalWarning.Tag = txtMedicalWarning.Text
    '其他医学警示
    strSQL = "" & _
    "  Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名='其他医学警示' "
    Set rs其他医学警示 = zlDatabase.OpenSQLRecord(strSQL, "其他医学警示", lng病人ID)
    While rs其他医学警示.EOF = False
        txtOtherWaring.Text = Nvl(rs其他医学警示!信息值)
        rs其他医学警示.MoveNext
    Wend
    '联系人相关信息
    '取病人信息表中的联系人信息
    strSQL = "Select A.联系人姓名, A.联系人关系, A.联系人电话, A.联系人身份证号, B.信息值 As 联系人附加信息" & vbNewLine & _
            "From 病人信息 A, 病人信息从表 B" & vbNewLine & _
            "Where a.病人id = b.病人id(+) And a.病人id = [1] And Not a.联系人姓名 Is Null And b.信息名(+) = '联系人附加信息'"
    Set rs病人信息 = zlDatabase.OpenSQLRecord(strSQL, "病人信息联系人信息", lng病人ID)
        If rs病人信息.EOF = False Then
        txt联系人身份证号.Text = Nvl(rs病人信息!联系人身份证号)
        txt联系人姓名.Text = Nvl(rs病人信息!联系人姓名)
        For i = 0 To cbo联系人关系.ListCount - 1
            If zlCommFun.GetNeedName(cbo联系人关系.List(i)) = Nvl(rs病人信息!联系人关系) Then cbo联系人关系.ListIndex = i
        Next
        txt联系人电话.Text = Nvl(rs病人信息!联系人电话)
        txtLinkManInfo.Text = Nvl(rs病人信息!联系人附加信息)
        
        SetLinkInfo Nvl(rs病人信息!联系人姓名), Nvl(rs病人信息!联系人关系), Nvl(rs病人信息!联系人电话), Nvl(rs病人信息!联系人身份证号), Nvl(rs病人信息!联系人附加信息)
    End If
    '取病人信息从表中的联系人信息
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 like '联系人%' order by 信息名 Asc"
    Set rs联系人 = zlDatabase.OpenSQLRecord(strSQL, "联系人相关信息", lng病人ID)
    If rs联系人.EOF = False Then
        rs联系人.Filter = "信息名 like '联系人姓名%'"
        lng联系人数量 = rs联系人.RecordCount
        rs联系人.Filter = ""
        For i = 2 To lng联系人数量 + 1
            While rs联系人.EOF = False
                Select Case Nvl(rs联系人!信息名)
                    Case "联系人姓名" & i
                        str联系人姓名 = Nvl(rs联系人!信息值)
                    Case "联系人关系" & i
                        str联系人关系 = Nvl(rs联系人!信息值)
                    Case "联系人附加信息" & i
                        str联系人附加信息 = Nvl(rs联系人!信息值)
                    Case "联系人电话" & i
                        str联系人电话 = Nvl(rs联系人!信息值)
                    Case "联系人身份证号" & i
                        str联系人身份证号 = Nvl(rs联系人!信息值)
                End Select
                rs联系人.MoveNext
            Wend
            SetLinkInfo str联系人姓名, str联系人关系, str联系人电话, str联系人身份证号, str联系人附加信息
            rs联系人.MoveFirst
        Next
    End If
    '其他信息
    strSQL = "" & _
    "   Select 病人ID,就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 信息名 Not in ('血型','RH','医学警示','其他医学警示','身份证号状态','外籍身份证号') And 信息名 Not like '联系人%'"
    Set rs其他信息 = zlDatabase.OpenSQLRecord(strSQL, "联系人其他信息", lng病人ID)
    '问题号:115886,焦博,2017/11/24,挂号提取该病人信息时，程序报错
    While rs其他信息.EOF = False
        If Nvl(rs其他信息!信息名) <> "" Then
            SetOtherInfo Nvl(rs其他信息!信息名), Nvl(rs其他信息!信息值)
        End If
        rs其他信息.MoveNext
    Wend
    
    '90875:李南春,2016/11/8,医疗卡证件类型
    Call LoadCertificate(lng病人ID)
    Exit Sub
ErrHandl:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Add健康卡相关信息(ByVal lng病人ID As Long, ByRef colPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:健康卡数据处理
    '入参:
    '编制:56599
    '日期:2012-12-13 18:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim varKey As Variant
    '过敏药物
    With vsDrug
        If .Rows > 1 Then
            '清除该病人所有记录
            strSQL = " Zl_病人过敏药物_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSQL
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人过敏药物_Update("
                    '病人ID_In 病人过敏药物.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '过敏药物ID_In 病人过敏药物.过敏药物ID%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 2) = "", "", .TextMatrix(i, 2)) & "',"
                    '过敏药物_In  病人过敏药物.过敏药物%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 0) = "", "", .TextMatrix(i, 0)) & "',"
                    '过敏反应_In 病人过敏反应.过敏反应%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"

                    zlAddArray colPro, strSQL
                End If
            Next
        End If
    End With
    '接种信息
    With vsInoculate
        If .Rows > 1 Then
            '清除该病人所有记录
            strSQL = " Zl_病人免疫记录_Delete(" & lng病人ID & ")"
            zlAddArray colPro, strSQL
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 0) = "", "''", "to_date('" & .TextMatrix(i, 0) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "')"
                    zlAddArray colPro, strSQL
                    Debug.Print strSQL
                End If
                If .TextMatrix(i, 3) <> "" Then
                    '病人过敏药物
                    strSQL = "Zl_病人免疫记录_Update("
                    '病人ID_In 病人免疫记录.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '接种时间_In 病人免疫记录.接种时间%Type
                    strSQL = strSQL & "" & IIf(.TextMatrix(i, 2) = "", "''", "to_date('" & .TextMatrix(i, 2) & "','yyyy-mm-dd')") & ","
                    '接种名称_In  病人免疫记录.接种名称%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "''", .TextMatrix(i, 3)) & "')"
                    zlAddArray colPro, strSQL
                    Debug.Print strSQL
                End If
            Next
        End If
    End With
    '其他信息
    'ABO血型
    '病人信息从表
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'血型',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & cboBloodType.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    'RH
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'RH',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & cboBH.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '医学警示
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & txtMedicalWarning.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    '其他医学警示
    strSQL = "Zl_病人信息从表_Update("
    '病人ID_In 病人信息从表.病人Id%Type
    strSQL = strSQL & "" & lng病人ID & ","
    '信息名_In 病人信息从表.信息名%Type
    strSQL = strSQL & "'其他医学警示',"
    '信息值_In 病人信息从表.信息值%Type
    strSQL = strSQL & "'" & txtOtherWaring.Text & "',"
    '就诊Id_In 病人信息从表.就诊Id%Type
    strSQL = strSQL & "'')"
    zlAddArray colPro, strSQL
    
    '联系人相关信息
    '联系人相关信息
    With vsLinkMan
        If .Rows > 2 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then '联系人姓名不能为空
                    For j = 0 To .Cols - 1
                        strSQL = "Zl_病人信息从表_Update("
                        '病人ID_In 病人信息从表.病人Id%Type
                        strSQL = strSQL & "" & lng病人ID & ","
                        '信息名_In 病人信息从表.信息名%Type
                        strSQL = strSQL & "'" & .TextMatrix(0, j) & i & "',"
                        '信息值_In 病人信息从表.信息值%Type
                        strSQL = strSQL & "'" & IIf(.TextMatrix(i, j) = "", "", .TextMatrix(i, j)) & "',"
                        '就诊Id_In 病人信息从表.就诊Id%Type
                        strSQL = strSQL & "'')"

                        zlAddArray colPro, strSQL
                    Next
                End If
            Next
        End If
    End With
    '其他信息
     With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    strSQL = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 0) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 1) = "", "", .TextMatrix(i, 1)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
                If .TextMatrix(i, 2) <> "" Then
                    strSQL = "Zl_病人信息从表_Update("
                    '病人ID_In 病人信息从表.病人Id%Type
                    strSQL = strSQL & "" & lng病人ID & ","
                    '信息名_In 病人信息从表.信息名%Type
                    strSQL = strSQL & "'" & .TextMatrix(i, 2) & "',"
                    '信息值_In 病人信息从表.信息值%Type
                    strSQL = strSQL & "'" & IIf(.TextMatrix(i, 3) = "", "", .TextMatrix(i, 3)) & "',"
                    '就诊Id_In 病人信息从表.就诊Id%Type
                    strSQL = strSQL & "'')"
                        
                    zlAddArray colPro, strSQL
                End If
            Next
        End If
     End With
     '医疗卡属性
     If Not mdic医疗卡属性 Is Nothing And Trim(txt卡号.Text) <> "" Then
        For Each varKey In mdic医疗卡属性.Keys
            strSQL = "Zl_病人医疗卡属性_Update("
            strSQL = strSQL & lng病人ID & ","
            strSQL = strSQL & mCurSendCard.lng卡类别ID & ","
            strSQL = strSQL & "'" & Trim(txt卡号.Text) & "',"
            strSQL = strSQL & "'" & varKey & "',"
            strSQL = strSQL & "'" & mdic医疗卡属性(varKey) & "')"
            zlAddArray colPro, strSQL
        Next
     End If
End Sub

Private Function CheckPatiCard() As Boolean
'功能：检查病人健康卡片录入的内容是否合法
'63246:刘鹏飞,2013-07-03
    Dim intLen As Integer, i As Integer, j As Integer
    
    intLen = 100
    If LenB(StrConv(txtMedicalWarning.Text, vbFromUnicode)) > intLen Then
        tbcPage.Item(1).Selected = True
        MsgBox "医学警示只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        If txtMedicalWarning.Enabled And txtMedicalWarning.Visible Then txtMedicalWarning.SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtOtherWaring.Text, vbFromUnicode)) > intLen Then
        tbcPage.Item(1).Selected = True
        MsgBox "其他医学警示只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        If txtOtherWaring.Enabled And txtOtherWaring.Visible Then txtOtherWaring.SetFocus
        Exit Function
    End If
    
    mblnCheckPatiCard = True
    '过敏药物
    With vsDrug
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    intLen = 60
                    If LenB(StrConv(.TextMatrix(i, 0), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "过敏药物只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第1列", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "过敏反应只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第2列", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
    
    '接种信息
    With vsInoculate
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    If Not IsDate(.TextMatrix(i, 0)) Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "接种时间不是有效的日期格式！" & vbCrLf & "错误行:第" & i & "行、第1列", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    
                    intLen = 200
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "接种名称只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第2列", vbInformation, gstrSysName
                        Call .Select(i, 1, i, 1)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
                If .TextMatrix(i, 3) <> "" Then
                    If Not IsDate(.TextMatrix(i, 2)) Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "接种时间不是有效的日期格式！" & vbCrLf & "错误行:第" & i & "行、第3列", vbInformation, gstrSysName
                        Call .Select(i, 2, i, 2)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    
                    intLen = 200
                    If LenB(StrConv(.TextMatrix(i, 3), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "接种名称只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第4列", vbInformation, gstrSysName
                        Call .Select(i, 3, i, 3)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                End If
            Next
        End If
    End With
    
    '联系人地址
    With vsLinkMan
        intLen = 100
        If .Rows > 2 Then
            For i = 2 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    For j = 0 To .Cols - 1
                        If .ColIndex("联系人姓名") = j Then
                            intLen = 64
                        ElseIf .ColIndex("联系人身份证号") = j Then
                            intLen = 18
                        ElseIf .ColIndex("联系人电话") = j Then
                            intLen = 20
                        Else
                            intLen = 100
                        End If
                        If LenB(StrConv(.TextMatrix(i, j), vbFromUnicode)) > intLen Then
                            tbcPage.Item(1).Selected = True
                            MsgBox .TextMatrix(0, j) & "只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行", vbInformation, gstrSysName
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
    
    '其他信息
    With vsOtherInfo
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    intLen = 20
                    If LenB(StrConv(.TextMatrix(i, 0), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "信息名只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第1列", vbInformation, gstrSysName
                        Call .Select(i, 0, i, 0)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 1), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "信息值只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第2列", vbInformation, gstrSysName
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
                        MsgBox "信息名只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第3列", vbInformation, gstrSysName
                        Call .Select(i, 2, i, 2)
                        .TopRow = i
                        If .Enabled = True And .Visible = True Then .SetFocus
                        Exit Function
                    End If
                    intLen = 100
                    If LenB(StrConv(.TextMatrix(i, 3), vbFromUnicode)) > intLen Then
                        tbcPage.Item(1).Selected = True
                        MsgBox "信息值只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！" & vbCrLf & "错误行:第" & i & "行、第4列", vbInformation, gstrSysName
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
    '功能:加载病人信息,读取病人信息
    '编制:刘兴洪
    '日期:2011-09-08 21:52:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim i As Long, j As Long, lngCount As Long, lngChildCount As Long '问题号:56599
    Dim str过敏药物 As String, str过敏反应 As String '问题号:56599
    Dim str接种日期 As String, str接种名称 As String '问题号:56599
    Dim strABO血型 As String '问题号:56599
    Dim str信息名 As String, str信息值 As String '问题号:56599
    Dim xmlChildNodes As IXMLDOMNodeList, xmlChildNode As IXMLDOMNode '问题号:56599
    Dim str姓名 As String, str关系 As String, str电话 As String, str身份证号 As String, str地址 As String '问题号:56599
    On Error GoTo errHandle

    If strPatiXML = "" Then Exit Function
    
    If zlXML_Init = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strPatiXML, False) = False Then Exit Function
    '    标识    数据类型    长度    精度    说明
    '    卡号    Varchar2    20
    Call zlXML_GetNodeValue("卡号", , strValue)
    '    姓名    Varchar2    64
    Call zlXML_GetNodeValue("姓名", , strValue)
    txtPatient.Text = strValue
    '    性别    Varchar2    4
    Call zlXML_GetNodeValue("性别", , strValue)
    If strValue <> "" Then
        Call cbo.Locate(cbo性别, strValue)
        If cbo性别.ListIndex = -1 Then
            cbo性别.AddItem strValue
            cbo性别.ListIndex = cbo性别.NewIndex
        End If
    End If
    '    年龄    Varchar2    10
    Call zlXML_GetNodeValue("年龄", , strValue)
    If strValue <> "" Then
        Call LoadOldData(strValue, txt年龄, cbo年龄单位)
    End If
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call zlXML_GetNodeValue("出生日期", , strValue)
    
    txt出生日期.Text = Format(IIf(strValue = "", "____-__-__", strValue), "YYYY-MM-DD")
    If strValue <> "" Then
        txt年龄.Text = ReCalcOld(CDate(Format(strValue, "YYYY-MM-DD HH:MM:SS")), cbo年龄单位, , , CDate(txt入院时间.Text))   '修改的时候,根据出生日期重算年龄
        If CDate(txt出生日期.Text) - CDate(strValue) <> 0 Then
            mblnChange = False
            txt出生时间.Text = Format(strValue, "HH:MM")
            mblnChange = True
        End If
    Else
        mblnChange = False
        Call ReCalcBirthDay
        mblnChange = True
    End If
    '    出生地点    Varchar2    50
    Call zlXML_GetNodeValue("出生地点", , strValue)
    '    身份证号    VARCHAR2    18
    Call zlXML_GetNodeValue("身份证号", , strValue)
    If strValue <> "" Then txt身份证号.Text = strValue
    '    其他证件    Varchar2    20
    Call zlXML_GetNodeValue("其他证件", , strValue)
    If strValue <> "" Then txt其他证件.Text = strValue
    '    职业    Varchar2    80
    Call zlXML_GetNodeValue("职业", , strValue)
    If strValue <> "" Then
        cbo职业.ListIndex = GetCboIndex(cbo职业, strValue, , , mstrCboSplit)
        If cbo职业.ListIndex = -1 Then
            cbo职业.AddItem strValue, 0
            cbo职业.ListIndex = cbo职业.NewIndex
        End If
    End If
    '    民族    Varchar2    20
    Call zlXML_GetNodeValue("民族", , strValue)
    cbo民族.ListIndex = GetCboIndex(cbo民族, strValue)
     If cbo民族.ListIndex = -1 And strValue <> "" Then
         cbo民族.AddItem strValue, 0
         cbo民族.ListIndex = cbo民族.NewIndex
     End If
    '    国籍    Varchar2    30
    Call zlXML_GetNodeValue("国籍", , strValue)
    cbo国籍.ListIndex = GetCboIndex(cbo国籍, strValue)
     If cbo国籍.ListIndex = -1 And strValue <> "" Then
         cbo国籍.AddItem strValue, 0
         cbo国籍.ListIndex = cbo国籍.NewIndex
     End If
    '    学历    Varchar2    10
    Call zlXML_GetNodeValue("学历", , strValue)
    cbo学历.ListIndex = GetCboIndex(cbo学历, strValue)
    If cbo学历.ListIndex = -1 And strValue <> "" Then
        cbo学历.AddItem strValue, 0
        cbo学历.ListIndex = cbo学历.NewIndex
    End If
    '    婚姻状况    Varchar2    4
    Call zlXML_GetNodeValue("婚姻状况", , strValue)
    cbo婚姻状况.ListIndex = GetCboIndex(cbo婚姻状况, strValue)
     If cbo婚姻状况.ListIndex = -1 And strValue <> "" Then
         cbo婚姻状况.AddItem strValue, 0
         cbo婚姻状况.ListIndex = cbo婚姻状况.NewIndex
     End If
    '    区域    Varchar2    30
    Call zlXML_GetNodeValue("区域", , strValue)
    txt区域.Text = strValue
    '    家庭地址    Varchar2    50
    Call zlXML_GetNodeValue("家庭地址", , strValue)
    txt家庭地址.Text = strValue
    
        '    户口地址    Varchar2    50
    Call zlXML_GetNodeValue("户口地址", , strValue)
    txt户口地址.Text = strValue
    If gbln启用结构化地址 Then PatiAddress(E_IX_户口地址).Value = strValue
    
    If gbln启用结构化地址 Then PatiAddress(E_IX_现住址).Value = strValue
    '    家庭电话    Varchar2    20
    Call zlXML_GetNodeValue("家庭电话", , strValue)
    txt家庭电话.Text = strValue
    '    家庭地址邮编    Varchar2    6
    Call zlXML_GetNodeValue("家庭地址邮编", , strValue)
    txt家庭地址邮编.Text = strValue
    '    工作单位    Varchar2    100
    Call zlXML_GetNodeValue("工作单位", , strValue)
    txt工作单位.Text = strValue
    lbl工作单位.Tag = ""
    '    单位电话    Varchar2    20
    Call zlXML_GetNodeValue("单位电话", , strValue)
    txt单位电话.Text = strValue
    '手机号
    Call zlXML_GetNodeValue("手机号", , strValue)
    txtMobile.Text = strValue
    '    单位邮编    Varchar2    6
    Call zlXML_GetNodeValue("单位邮编", , strValue)
   txt单位邮编.Text = strValue
    '    单位开户行  Varchar2    50
    Call zlXML_GetNodeValue("单位开户行", , strValue)
   txt单位开户行.Text = strValue
    '    单位帐号    Varchar2    50
    Call zlXML_GetNodeValue("单位帐号", , strValue)
   txt单位帐号.Text = strValue
    '问题号:56599
    '过敏情况
    Call zlXML_GetRows("药物名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("药物名称", i, str过敏药物)
        Call zlXML_GetNodeValue("药物反应", i, str过敏反应)
        SetDrugAllergy str过敏药物, str过敏反应
    Next
    lngCount = 0
    '免疫记录
    Call zlXML_GetRows("疫苗名称", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("疫苗名称", i, str接种名称)
        Call zlXML_GetNodeValue("接种时间", i, str接种日期)
        SetInoculate str接种日期, str接种名称
    Next
    lngCount = 0
    'ABO血型
    Call zlXML_GetNodeValue("ABO血型", , strABO血型)
    If strABO血型 <> "" Then
        For i = 0 To cboBloodType.ListCount - 1
            If cboBloodType.List(i) = strABO血型 Then cboBloodType.ListIndex = i
        Next
    End If
    'RH
    Call zlXML_GetNodeValue("RH", , strValue)
    If strValue <> "" Then
        For i = 0 To cboBH.ListCount - 1
            If cboBH.List(i) = strValue Then cboBH.ListIndex = i
        Next
    End If
    '医学警示
    strValue = ""
    Set xmlChildNodes = zlXML_GetChildNodes("临床基本信息")
    If Not xmlChildNodes Is Nothing Then
        If xmlChildNodes.length > 0 Then
            For i = 0 To xmlChildNodes.length - 1
                Set xmlChildNode = xmlChildNodes(i)
                If xmlChildNode.Text = "1" Then
                    strValue = strValue & ";" & Replace(xmlChildNode.nodeName, "标志", "")
                End If
            Next
        End If
    End If
    If strValue <> "" Then txtMedicalWarning.Text = Mid(strValue, 2)
   
    
    '其他医学警示
    Call zlXML_GetNodeValue("其他医学警示", , strValue)
    If strValue <> "" Then txtOtherWaring.Text = strValue
    '联系信息
    '    联系人地址  Varchar2    50
    Call zlXML_GetNodeValue("联系人地址", , str地址)
    txt联系人地址.Text = str地址
    If gbln启用结构化地址 Then PatiAddress(E_IX_联系人地址).Value = str地址
     '    联系人姓名  Varchar2    64
    Call zlXML_GetNodeValue("联系人姓名", , str姓名)
    '    联系人关系  Varchar2    30
    Call zlXML_GetNodeValue("联系人关系", , str关系)
    '    联系人电话  Varchar2    20
    Call zlXML_GetNodeValue("联系人电话", , str电话)
    '    联系人身份证 Varchar2   20
    Call zlXML_GetNodeValue("联系人身份证号", , str身份证号)
    SetLinkInfo str姓名, str关系, str电话, str身份证号
    
    Call zlXML_GetRows("联系信息", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("联系信息", "姓名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("联系信息", "姓名", i, j, str姓名)
                Call zlXML_GetChildNodeValue("联系信息", "关系", i, j, str关系)
                Call zlXML_GetChildNodeValue("联系信息", "电话", i, j, str电话)
                Call zlXML_GetChildNodeValue("联系信息", "身份证号", i, j, str身份证号)
                SetLinkInfo str姓名, str关系, str电话, str身份证号
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0

    '其他信息
    '健康档案编号
    Call zlXML_GetNodeValue("健康档案编号", , strValue)
    SetOtherInfo "健康档案编号", strValue
    
    '新农合证号
    Call zlXML_GetNodeValue("新农合证号", , strValue)
    SetOtherInfo "新农合证号", strValue

    '其他证件
    Call zlXML_GetRows("其他证件", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("其他证件", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("其他证件", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("其他证件", "信息值", i, j, str信息值)
                SetOtherInfo str信息名, str信息值
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '其他信息
    Call zlXML_GetRows("其他信息", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("其他信息", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("其他信息", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("其他信息", "信息值", i, j, str信息值)
                SetOtherInfo str信息名, str信息值
            Next
        End If
    Next
    lngCount = 0: lngChildCount = 0
    '医疗卡属性
    Call zlXML_GetRows("医疗卡属性", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetChildRows("医疗卡属性", "信息名", lngChildCount, i)
        If lngChildCount > 0 Then
            For j = 0 To lngChildCount - 1
                Call zlXML_GetChildNodeValue("医疗卡属性", "信息名", i, j, str信息名)
                Call zlXML_GetChildNodeValue("医疗卡属性", "信息值", i, j, str信息值)
                If mdic医疗卡属性.Exists(str信息名) Then
                    mdic医疗卡属性.Item(str信息名) = str信息值
                Else
                    mdic医疗卡属性.Add str信息名, str信息值
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
'功能：由字符串在ComboBox中查找索引
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '先精确查找
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), strSplit) > 0 Then
            If zlCommFun.GetNeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '最后模糊查找
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Public Sub Clear健康档案()
    '---------------------------------------------------------------------------------------------------------------------------------------------
'功能:判断当前是否为卡发操作 (不是发卡操作既是绑定卡操作)
'入参:
'编制:56599
'日期:2012-12-25 14:55:36
'---------------------------------------------------------------------------------------------------------------------------------------------
    '血型
    Call SetCboDefault(cboBloodType)
    'RH
    If cboBH.ListCount > 0 Then cboBH.ListIndex = -1
    '医学警示
    txtMedicalWarning.Text = ""
    '其他医学警示
    txtOtherWaring.Text = ""
    '联系人信息
    With vsLinkMan
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
    End With
    '接种情况
    With vsInoculate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '其他信息
    With vsOtherInfo
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""

    End With
    '证件信息
    With vsCertificate
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
    End With
    '过敏反应
    With vsDrug
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
    End With
End Sub

Public Function zlCreateSquare() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建医疗卡对象
    '编制:李南春
    '日期:2016/6/21 11:57:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If Not mobjSquare Is Nothing Then zlCreateSquare = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjSquare = CreateObject("zl9CardSquare.clsCardsquare")
    If Err <> 0 Then Err = 0: Exit Function
    Call mobjSquare.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend)
    '初始部件不成功,则作为不存在处理
    zlCreateSquare = True
End Function

Private Function WriteCard(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写卡
    '入参:lng病人ID - 病人ID
    '编制:王吉
    '问题:56599
    '日期:2012-12-17 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    On Error GoTo ErrHandl:
    If mobjSquare Is Nothing Then
       If Not zlCreateSquare() Then Exit Function
    End If
    If mobjSquare Is Nothing Then Exit Function
    WriteCard = mobjSquare.zlBandCardArfter(Me, mlngModul, mCurSendCard.lng卡类别ID, lng病人ID, strExpend)
    Exit Function
ErrHandl:
    WriteCard = False
    If ErrCenter() = 1 Then Resume
End Function
Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim lngTop As Long
    '问题号:56599
    Select Case Item.Caption
        Case "基本"
            Me.Height = mPageHeight.基本
            pic入院.Top = pic病人.Top + pic病人.Height
            lngTop = pic入院.Top + pic入院.Height
            If mbln是否显示预交 Then
                pic预交.Top = lngTop
                lngTop = pic预交.Top + pic预交.Height
            End If
            If mbytInState = 1 Or (mbytInState = 0 And mbytMode = 2 And mbytKind <> E住院留观登记) Then
                If txt住院号.Enabled And txt住院号.Visible Then txt住院号.SetFocus
            ElseIf mbytInState = 0 And mbytMode = 2 And mbytKind = E住院留观登记 Then
                If txt姓名.Enabled And txt姓名.Visible Then txt姓名.SetFocus
            Else
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            End If
            pic磁卡.Top = lngTop
        Case "健康档案"
            Me.Height = mPageHeight.健康档案
            If cboBloodType.Enabled And cboBloodType.Visible Then cboBloodType.SetFocus
    End Select
    tbcPage.Height = picCmd.Top
    tbcPage.width = Me.width - 90
    Call SetCenter(Me)
End Sub

Private Sub SetCardEditEnabled()
    '设置就诊卡编辑属性
    Dim blnEdit As Boolean
    blnEdit = Trim(txt卡号.Text) <> ""
    
    txtPass.Enabled = blnEdit: txtAudi.Enabled = blnEdit
    lbl密码.Enabled = txtPass.Enabled: lbl验证.Enabled = blnEdit
    
    txt卡额.Enabled = blnEdit: lbl金额.Enabled = blnEdit
    chk记帐.Enabled = blnEdit
    cbo发卡结算.Enabled = chk记帐.Value = 0 And blnEdit
End Sub
Private Function Check发卡性质(lng病人ID As Long, lng卡类别ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发卡时检查是否限制病人的发卡张数
    '入参:lng病人ID - 病人ID;lng卡类别ID  - 医疗卡的类别ID
    '编制:王吉
    '问题:57326
    '日期:2013-01-30 15:07:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl:
        strSQL = "Select Count(1) as 存在 From 病人医疗卡信息 Where 状态=0 And 病人ID=[1] And 卡类别ID=[2] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng卡类别ID)
        If Val(Nvl(rsTemp!存在)) <= 0 Then Check发卡性质 = True: Exit Function
        Select Case mCurSendCard.lng发卡性质
            Case 0 '不限制
                Check发卡性质 = True
            Case 1 '同一个病人只允许发一张卡
                MsgBox "该病人已经发过" & mCurSendCard.str卡名称 & ",不能在进行发卡操作!", vbInformation + vbOKOnly
                Check发卡性质 = False
            Case 2 '同一个病人允许发多张卡,但需要提醒
               Check发卡性质 = MsgBox("该病人已经发过" & mCurSendCard.str卡名称 & ",是否要进行发卡操作?", vbQuestion + vbYesNo) = vbYes
        End Select
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function WhetherTheCardBinding(ByVal str卡号 As String, ByVal lng卡类别 As Long, Optional ByRef lngPatientID As Long) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:获取指定卡号是否已经发卡
'入参:str卡号：卡号 ，lng卡类别：卡类别 , lngPatientID :病人ID
'返回:True :已经发卡;False:未发卡
'编制:
'日期:
'问题号:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHandl
    strSQL = "" & _
           "   Select 病人ID From 病人医疗卡信息 Where 卡号=[1]  And 卡类别ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "门诊挂号", str卡号, lng卡类别)
    WhetherTheCardBinding = rsTemp.RecordCount > 0

    If rsTemp.RecordCount > 0 Then
        lngPatientID = Val(Nvl(rsTemp!病人ID))
    End If

    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then Resume
End Function

Private Function GetCardLastChangeType(ByVal str卡号 As String, ByVal lng卡类别 As Long, ByVal lngPaitentID As Long) As Long
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:获取卡最后的变动类型
'入参:str卡号：卡号 ，lng卡类别：卡类别 , lngPatientID :病人ID
'返回:0-未找到相关信息   1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
'编制:李光福
'日期:2013-2-4 17:36:33
'问题号:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "     Select 变动类别" & vbNewLine & _
           "    From (With 医疗卡变动 As (Select 病人id, ID, 变动类别, 变动时间 " & vbNewLine & _
           "                              From 病人医疗卡变动 Bd" & vbNewLine & _
           "                              Where Bd.卡号 = [2] And 卡类别id = [1] And 病人id = [3])" & vbNewLine & _
           "           Select A.变动类别" & vbNewLine & _
           "           From 医疗卡变动 A, (Select Max(变动时间) As 变动时间 From 医疗卡变动 C) B" & vbNewLine & _
           "           Where A.变动时间 = B.变动时间) A"
    On Error GoTo ErrHand
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取卡最后变动信息", lng卡类别, str卡号, lngPaitentID)
    If Not rsTmp Is Nothing Then
        If rsTmp.RecordCount > 0 Then
            GetCardLastChangeType = Val(Nvl(rsTmp!变动类别))
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
'功能:取消绑定卡
'入参:intType:0-当前卡号;1-当前类别;2-当前病人所有
'返回:取消成功,返回true,否则返回False
'编制:刘兴洪
'日期:2011-07-29 11:18:05
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Curdate As Date
    Dim strSQL As String, strPassWord As String

    On Error GoTo errHandle

    Curdate = zlDatabase.Currentdate

    'Zl_医疗卡变动_Insert
    strSQL = "Zl_医疗卡变动_Insert("
    '      变动类型_In   Number,
    '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
    strSQL = strSQL & "" & 14 & ","
    '      病人id_In     住院费用记录.病人id%Type,
    strSQL = strSQL & "" & lngPatientID & ","
    '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
    strSQL = strSQL & "" & lngCardTypeID & ","
    '      原卡号_In     病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "NULL,"
    '      医疗卡号_In   病人医疗卡信息.卡号%Type,
    strSQL = strSQL & "'" & strCardNO & "'" & ","
    '      变动原因_In   病人医疗卡变动.变动原因%Type,
    strSQL = strSQL & "'卡重复利用自动取消原卡绑定信息',"
    '      密码_In       病人信息.卡验证码%Type,
    strSQL = strSQL & "NULL,"
    '      操作员姓名_In 住院费用记录.操作员姓名%Type,
    strSQL = strSQL & "NULL,"
    '      变动时间_In   住院费用记录.登记时间%Type,
    strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
    strSQL = strSQL & "NULL,"
    '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
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
            '跳到下一行
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
'            '跳到上一行
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
    If Col = 1 Or Col = 3 Then '接种名称列编辑时需判断是否字数超过了200
        With vsInoculate
           If LenB(StrConv(.EditText, vbFromUnicode)) > 200 Then
                If MsgBox("接种名称输入字符超出最大字符数200,请问是否将多出的字符将被自动截除？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    .EditText = StrConv(MidB(StrConv(.EditText, vbFromUnicode), 1, 200), vbUnicode)
                Else
                    Cancel = True
                End If
           End If
        End With
    Else
        With vsInoculate
            If IsDate(Format(.EditText, "YYYY-MM-DD")) = False And .EditText <> "    -  -  " Then
                 MsgBox "输入的日期格式不对或不是正确的日期，请检查！", vbInformation, gstrSysName
                 Cancel = True
            ElseIf .EditText = "    -  -  " Then
                 .EditText = ""
            Else
                If .EditText <> "" Then
                    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
                    If Format(.EditText, "YYYY-MM-DD") > strCurDate Then
                        MsgBox "接种日期不能大于服务器系统时间[" & strCurDate & "],请检查！", vbInformation, gstrSysName
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
        If Col = .ColIndex("联系人关系") Then
            If zlCommFun.GetNeedName(.TextMatrix(Row, Col)) = "其他" Then
                .Cell(flexcpBackColor, Row, .ColIndex("联系人关系备注")) = &H80000005
            Else
                .TextMatrix(Row, .ColIndex("联系人关系备注")) = ""
                .Cell(flexcpBackColor, Row, .ColIndex("联系人关系备注")) = &HE0E0E0
            End If
        End If
    End With
End Sub

Private Sub vsLinkMan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsLinkMan
        If .Col = .ColIndex("联系人关系备注") Then
            If zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("联系人关系"))) = "其他" Then
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
            txt联系人姓名.Text = .TextMatrix(.FixedRows, .ColIndex("联系人姓名"))
            For i = 0 To cbo联系人关系.ListCount - 1
                If zlCommFun.GetNeedName(cbo联系人关系.List(i)) = .TextMatrix(.FixedRows, .ColIndex("联系人关系")) Then
                    Exit For
                End If
            Next
            If i < cbo联系人关系.ListCount Then
                cbo联系人关系.ListIndex = i
            Else
                cbo联系人关系.ListIndex = -1
            End If
            
            txt联系人身份证号.Text = .TextMatrix(.FixedRows, .ColIndex("联系人身份证号"))
            txt联系人电话.Text = .TextMatrix(.FixedRows, .ColIndex("联系人电话"))
            txtLinkManInfo.Text = .TextMatrix(.FixedRows, .ColIndex("联系人关系备注"))
            
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
        If Col = .ColIndex("联系人身份证号") Then
            If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
                KeyAscii = 0
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
        ElseIf Col = .ColIndex("联系人电话") Then
            If InStr("0123456789()-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        ElseIf Col = .ColIndex("联系人关系备注") Then
            If InStr(":：,，", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        End If
    End With
End Sub

Private Sub vsLinkMan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Integer
    
    With vsLinkMan
        If Not Row = .FixedRows Then Exit Sub
        Select Case Col
            Case .ColIndex("联系人姓名")
                txt联系人姓名.Text = Trim(.EditText)
            Case .ColIndex("联系人关系")
                For i = 0 To cbo联系人关系.ListCount - 1
                    If zlCommFun.GetNeedName(cbo联系人关系.List(i)) = Trim(.EditText) Then Exit For
                Next
                If i < cbo联系人关系.ListCount Then
                    cbo联系人关系.ListIndex = i
                Else
                    cbo联系人关系.ListIndex = -1
                End If
                
            Case .ColIndex("联系人身份证号")
                txt联系人身份证号.Text = Trim(.EditText)
            Case .ColIndex("联系人电话")
                txt联系人电话.Text = Trim(.EditText)
            Case .ColIndex("联系人关系备注")
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
'入参:
'       lngPatiID:病人ID
'       bytMode:0:检查住院号是否病人之前已经使用,1:获取病人本次住院前的最后一次的住院号
'       strNo:bytMode=0,要检查的住院号,bytMode=1,返回的住院号
'返回:bytMode=0,已经使用返回TRUE,bytMode=1,存在历史住院，且住院号不为空，返回TRUE
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    If bytMode = 1 Then
        If lngPageID = 0 Then '预约登记
            gstrSQL = "Select 住院号 From 病案主页 Where 病人id = [1] And Nvl(主页id, 0) <> [2] And 住院号 Is Not Null Order By 主页id Desc"
        Else
            gstrSQL = "Select 住院号 From 病案主页 Where 病人id = [1] And 主页id < [2] And 住院号 Is Not Null Order By 主页id Desc"
        End If
    Else
        gstrSQL = "Select 病人ID from 病案主页 where 病人ID=[1] and nvl(主页ID,0)<>[2] and 住院号=[3] and rownum<2"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院号提取", lngPatiID, lngPageID, Val(strno))
    If bytMode = 0 Then
        CheckByPatiNO = rsTemp.RecordCount > 0
    ElseIf bytMode = 1 Then
        If rsTemp.RecordCount > 0 Then strno = rsTemp!住院号 & ""
        CheckByPatiNO = strno <> ""
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub InitStructAddress()
'功能:根据是否启用结构化地址调整界面
    Dim i As Long
    Dim lngLeft As Long
    
    If gbln启用结构化地址 Then
        For i = PatiAddress.LBound To PatiAddress.UBound
            If i = E_IX_现住址 Or i = E_IX_户口地址 Or i = E_IX_联系人地址 Then
                PatiAddress(i).Items = Five
            End If
            PatiAddress(i).TextBackColor = &H80000005
            PatiAddress(i).Visible = True
            PatiAddress(i).ShowTown = gbln显示乡镇
        Next
        For i = LBound(marrAddress) To UBound(marrAddress)
            marrAddress(i) = ""
        Next
        txt家庭地址.Visible = False
        cmd家庭地址.Visible = False
        txt出生地点.Visible = False
        cmd出生地点.Visible = False
        txt户口地址.Visible = False
        cmd户口地址.Visible = False
        txt籍贯.Visible = False
        cmd籍贯.Visible = False
        txt联系人地址.Visible = False
        cmd联系人地址.Visible = False
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
             PatiAddress(i).Visible = False
        Next
        
        txt家庭地址.Visible = True
        cmd家庭地址.Visible = True
        txt出生地点.Visible = True
        cmd出生地点.Visible = True
        txt户口地址.Visible = True
        cmd户口地址.Visible = True
        txt籍贯.Visible = True
        cmd籍贯.Visible = True
        txt联系人地址.Visible = True
        cmd联系人地址.Visible = True
        
        '界面对齐
        lngLeft = lbl学历.Left + lbl学历.width
        lbl家庭电话.Left = lngLeft - lbl家庭电话.width
        lbl户口地址邮编.Left = lngLeft - lbl户口地址邮编.width
        lngLeft = cbo学历.Left
        txt家庭电话.Left = lngLeft
        txt户口地址邮编.Left = lngLeft
    End If
End Sub

Private Sub SetStrutAddress(Optional ByVal bytFunc As Byte)
'功能:89980病人结构化
'bytFunc=1 清空数据
'       =2 设置户口地址和家庭地址的缺省值
    Dim i As Long
    
    If bytFunc = 2 Then
        txt家庭地址.Text = marrAddress(0) & marrAddress(1) & marrAddress(2) & marrAddress(3) & marrAddress(4)
        txt户口地址.Text = marrAddress(0) & marrAddress(1) & marrAddress(2) & marrAddress(3) & marrAddress(4)
        Call PatiAddress(E_IX_现住址).LoadStructAdress(marrAddress(0), marrAddress(1), marrAddress(2), marrAddress(3), marrAddress(4))
        Call PatiAddress(E_IX_户口地址).LoadStructAdress(marrAddress(0), marrAddress(1), marrAddress(2), marrAddress(3), marrAddress(4))
    Else
        For i = PatiAddress.LBound To PatiAddress.UBound
            If bytFunc = 1 Then
                PatiAddress(i).Value = ""
            Else
                PatiAddress(i).Enabled = (mbytInState <> EState.E查阅)
            End If
        Next
    End If
End Sub

Private Sub ReCalcBirthDay(Optional ByRef strMsg As String)
    Dim strBirth As String
    
    If CreatePublicPatient() Then
        If gobjPublicPatient.ReCalcBirthDay(Trim(txt年龄.Text) & IIf(cbo年龄单位.Visible, Trim(cbo年龄单位.Text), ""), strBirth, Format(txt入院时间.Text, "YYYY-MM-DD HH:MM"), strMsg) Then
            If txt出生日期.Enabled Then txt出生日期.Text = Format(strBirth, "YYYY-MM-DD")
            If txt出生时间.Enabled Then
                strBirth = Format(strBirth, "HH:MM")
                txt出生时间.Text = IIf(strBirth = "00:00", "__:__", strBirth)
            End If
            cbo年龄单位.Tag = txt年龄.Text & "_" & cbo年龄单位.Text
        End If
    End If
End Sub

Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡接口
    '返回: true-成功，false-失败
    '编制:李南春
    '日期:2016/6/20 13:54:56
    '问题:97634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    If txt卡号.Locked Then Exit Function
    If mobjSquare Is Nothing Then
       If Not zlCreateSquare() Then Exit Function
    End If
    If mobjSquare Is Nothing Then Exit Function
    If mCurSendCard.lng卡类别ID = 0 Or Val(mCurSendCard.str读卡性质) < 99 Then Exit Function
    If mobjSquare.zlSetBrushCardObject(mCurSendCard.lng卡类别ID, IIf(blnComm, txt卡号, Nothing), strExpend) Then
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New clsCommEvents
        Call mobjSquare.zlInitEvents(Me.hWnd, mobjCommEvents)
    End If
End Function

Private Sub EMPI_LoadPati(Optional ByVal lngFunc As Long = 0)
'功能:将EMPI返回来的病人信息更新到界面
'lngFunc=0 更新病人信息;1-根据返回的病人ID重新加载病人基础信息后更新
    Dim rsPatiIn As ADODB.Recordset
    Dim rsPatiOut As ADODB.Recordset
    Dim str出生日期 As String
    Dim blnRet As Boolean
    Static blnOpen As Boolean
    
    If blnOpen Then Exit Sub
    If CreatePlugInOK(glngModul) Then
        '组织病人基本信息
        Set rsPatiIn = New ADODB.Recordset
        With rsPatiIn.Fields
            .Append "病人ID", adBigInt
            .Append "主页ID", adBigInt
            .Append "挂号ID", adBigInt
            '-------------------------------
            .Append "门诊号", adVarChar, 18
            .Append "住院号", adVarChar, 18
            .Append "医保号", adVarChar, 30
            .Append "身份证号", adVarChar, 18
            .Append "其他证件", adVarChar, 20
            .Append "姓名", adVarChar, 100
            .Append "性别", adVarChar, 4
            .Append "出生日期", adVarChar, 20 '日期格式：YYYY-MM-DD HH:MM:SS
            .Append "出生地点", adVarChar, 100
            .Append "国籍", adVarChar, 30
            .Append "民族", adVarChar, 20
            .Append "学历", adVarChar, 10
            .Append "职业", adVarChar, 80
            .Append "工作单位", adVarChar, 100
            .Append "邮箱", adVarChar, 30
            .Append "婚姻状况", adVarChar, 4
            .Append "家庭电话", adVarChar, 20
            .Append "联系人电话", adVarChar, 20
            .Append "单位电话", adVarChar, 20
            .Append "家庭地址", adVarChar, 100
            .Append "家庭地址邮编", adVarChar, 6
            .Append "户口地址", adVarChar, 100
            .Append "户口地址邮编", adVarChar, 6
            .Append "单位邮编", adVarChar, 6
            .Append "联系人地址", adVarChar, 100
            .Append "联系人关系", adVarChar, 30
            .Append "联系人姓名", adVarChar, 64
        End With
        rsPatiIn.CursorLocation = adUseClient
        rsPatiIn.LockType = adLockOptimistic
        rsPatiIn.CursorType = adOpenStatic
        rsPatiIn.Open
        
        If txt出生时间 = "__:__" Then
            str出生日期 = IIf(IsDate(txt出生日期.Text), Format(txt出生日期.Text, "YYYY-MM-DD"), "")
        Else
            str出生日期 = IIf(IsDate(txt出生日期.Text), Format(txt出生日期.Text & " " & txt出生时间.Text, "YYYY-MM-DD HH:MM:SS"), "")
        End If
 
        With rsPatiIn
            .AddNew
            !病人ID = CLng(txtPatient.Tag)
            !主页ID = CLng(txtPages.Text)
            !住院号 = Trim(txt住院号.Text)
            !医保号 = Trim(txt医保号.Text)
            '-要更新的字段--------------------------------------------
            !身份证号 = Trim(txt身份证号.Text)
            !其他证件 = Trim(txt其他证件.Text)
            !姓名 = Trim(txt姓名.Text)
            !性别 = zlCommFun.GetNeedName(cbo性别.Text)
            !出生日期 = str出生日期 '日期格式：YYYY-MM-DD HH:MM:SS
            !出生地点 = Trim(txt出生地点.Text)
            !国籍 = zlCommFun.GetNeedName(cbo国籍.Text)
            !民族 = zlCommFun.GetNeedName(cbo民族.Text)
            !学历 = zlCommFun.GetNeedName(cbo学历.Text)
            !职业 = zlCommFun.GetNeedName(cbo职业.Text)
            !工作单位 = Trim(txt工作单位.Text)
            !婚姻状况 = zlCommFun.GetNeedName(cbo婚姻状况.Text)
            !家庭电话 = Trim(txt家庭电话.Text)
            !联系人电话 = Trim(txt联系人电话.Text)
            !单位电话 = Trim(txt单位电话.Text)
            !家庭地址 = Trim(txt家庭地址.Text)
            !家庭地址邮编 = Trim(txt家庭地址邮编.Text)
            !户口地址 = Trim(txt户口地址.Text)
            !户口地址邮编 = Trim(txt户口地址邮编.Text)
            !单位邮编 = Trim(txt单位邮编.Text)
            !联系人地址 = Trim(txt联系人地址.Text)
            !联系人关系 = zlCommFun.GetNeedName(cbo联系人关系.Text)
            !联系人姓名 = Trim(txt联系人姓名.Text)
            .Update
            '-------------------------------------------------------
        End With
        
        '调用查询接口
        On Error Resume Next
        blnRet = gobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModul, rsPatiIn, rsPatiOut)
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: On Error GoTo 0
        If Not blnRet Then Exit Sub
        If rsPatiOut Is Nothing Then Exit Sub
        If rsPatiOut.RecordCount = 0 Then Exit Sub
        '找到病人，将病人最新的信息更新到界面
        With rsPatiOut
            mblnEMPI = True     '用于标记找到建档病人
            '104916 只输入姓名,接口弹出界面输入更多信息找到HIS病人ID时无需再新建病人
            If mbytInState = E新增 And CLng(txtPatient.Tag) <> CLng(!病人ID & "") And CLng(!病人ID & "") <> 0 And lngFunc = 1 Then
                ClearCard
                txtPatient.Text = "-" & !病人ID
                blnOpen = True
                Call txtPatient_KeyPress(vbKeyReturn)
                blnOpen = False
                If txtPatient.Text = "" Then Exit Sub
            End If
            Call cbo.Locate(cbo性别, !性别 & "")
            Call cbo.Locate(cbo民族, !民族 & "")
            Call cbo.Locate(cbo国籍, !国籍 & "")
            Call cbo.Locate(cbo学历, !学历 & "")
            Call cbo.SeekIndex(cbo职业, !职业 & "")  '包含特殊字符
            Call cbo.Locate(cbo婚姻状况, !婚姻状况 & "")
            Call cbo.Locate(cbo联系人关系, !联系人关系 & "")
            
            If IsDate(!出生日期 & "") Then
                txt出生日期.Text = Format(CDate(!出生日期 & ""), "YYYY-MM-DD")
                txt出生时间.Text = IIf(Format(CDate(!出生日期 & ""), "HH:MM") = "00:00", "__:__", Format(CDate(!出生日期 & ""), "HH:MM"))
            End If
            
            If gbln启用结构化地址 Then
                PatiAddress(E_IX_出生地点).Value = !出生地点 & ""
                PatiAddress(E_IX_现住址).Value = !家庭地址 & ""
                PatiAddress(E_IX_户口地址).Value = !户口地址 & ""
                PatiAddress(E_IX_联系人地址).Value = !联系人地址 & ""
            End If
            txt医保号.Text = !医保号 & ""
            txt出生地点.Text = !出生地点 & ""
            txt家庭地址.Text = !家庭地址 & ""
            txt户口地址.Text = !户口地址 & ""
            txt联系人地址.Text = !联系人地址 & ""
            txt身份证号.Text = !身份证号 & ""
            txt其他证件.Text = !其他证件 & ""
            txt姓名.Text = !姓名 & ""
            txt工作单位.Text = !工作单位 & ""
            txt家庭电话.Text = !家庭电话 & ""
            txt联系人电话.Text = !联系人电话 & ""
            txt单位电话.Text = !单位电话 & ""
            txt家庭地址邮编.Text = !家庭地址邮编 & ""
            txt户口地址邮编.Text = !户口地址邮编 & ""
            txt单位邮编.Text = !单位邮编 & ""
            txt联系人姓名.Text = !联系人姓名 & ""
        End With
    End If
End Sub

Private Function EMPI_AddORUpdatePati(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef strErr As String) As Boolean
'功能:增加或更新EMPI病人信息
    Dim lngRet  As Long
    Dim strPlugErr As String
    Dim strTmp As String
    
    lngRet = 1 '默认成功 兼容 老版zlPlug当不支持此接口错误号:438
    If CreatePlugInOK(glngModul) Then
        If Not mblnEMPI Then
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_AddPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=成功;0-失败
            Call zlPlugInErrH(Err, "EMPI_AddPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "向EMPI平台新增病人信息失败！"
        Else
            On Error Resume Next
            lngRet = gobjPlugIn.EMPI_ModifyPatiInfo(glngSys, glngModul, lngPatiID, lngPageID, 0, strErr) '1=成功;0-失败
            Call zlPlugInErrH(Err, "EMPI_ModifyPatiInfo", strPlugErr)
            Err.Clear: On Error GoTo 0
            strTmp = "向EMPI平台更新病人信息失败！"
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
    '问题号:90875

    With vsCertificate
        If Col = 1 Or Col = 3 Then '证件号码不能超过30
            If Len(.TextMatrix(Row, Col)) > 30 Then
                 MsgBox "证件输入字符超出最大字符数30,多出的字符将被自动截除！", vbInformation, gstrSysName
                 .TextMatrix(Row, Col) = Mid(.TextMatrix(Row, Col), 1, 30)
            End If
            If Col = 3 And .Rows - 1 = Row And .TextMatrix(Row, Col) <> "" Then
                .Rows = .Rows + 1
            End If
        ElseIf Col = 0 Or Col = 2 Then '检查是否选择了重复的证件类型
            For lngRow = 1 To .Rows - 1
                For lngCol = 0 To .Cols - 1 Step 2
                    If (lngRow <> Row Or lngCol <> Col) And .TextMatrix(lngRow, lngCol) = .TextMatrix(Row, Col) And .TextMatrix(Row, Col) <> "" Then
                        MsgBox .TextMatrix(lngRow, lngCol) & "已存在，不能重复选择。", vbInformation, gstrSysName
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
    '问题号:90875
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
    '78408:李南春,2014/10/9,光标跳转
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
    '功能:初始化VSGrid控件
    '编制:90875
    '日期:2015/12/17 16:59:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    Dim strSQL As String, rsTemp As ADODB.Recordset, str关系 As String, i As Integer
    With vsCertificate
    '初始化列表属性
        .Editable = IIf(mbytInState = 2, flexEDNone, flexEDKbdMouse)
        .SelectionMode = flexSelectionFree
    '设置列头
        SetColumHeader vsCertificate, C_CertificateHeader
    '设置列信息
        strSQL = "Select 名称,缺省标志 from 证件类型  Where  名称 Not Like '其他%' and 名称 Not Like '%身份证'" & vbNewLine & _
                " And Not 名称 in (Select 名称 from  医疗卡类别 Where Nvl(是否证件,0)=0 or Nvl(是否启用,0)=0)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTemp.RecordCount = 0 Then .Editable = flexEDNone: Exit Sub
        Do While Not rsTemp.EOF
            str关系 = str关系 & "|" & Nvl(rsTemp!名称)
            rsTemp.MoveNext
        Loop
        str关系 = Mid(str关系, 2)
        If str关系 <> "" Then .ColComboList(0) = str关系: .ColComboList(2) = str关系
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub LoadCertificate(ByVal lng病人ID As Long)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人的证件信息到界面
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    
    On Error GoTo ErrHand
    strSQL = "Select  A.名称,A.ID,B.卡号 from 医疗卡类别 A, 病人医疗卡信息 B " & _
            "Where A.ID= B.卡类别ID And A.是否启用=1 And A.是否证件=1 And B.状态=0  And  B.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    With vsCertificate
        .Clear 1
        .Rows = 2
        lngRow = 1: lngCol = 0
        While Not rsTemp.EOF
            .TextMatrix(lngRow, lngCol) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, lngCol + 1) = Nvl(rsTemp!卡号)
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

Private Sub AddCardTypeSQL(ByVal intOper As Integer, ByVal lng卡类别ID As Long, ByVal strCode As String, ByVal str全名 As String, ByVal str短名 As String, _
                           ByVal lng卡号长度 As Long, ByRef strSQL As String)

    ' Zl_医疗卡类别_Update
    strSQL = "Zl_医疗卡类别_Update("
    '  Id_In           In 医疗卡类别.ID%Type,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '  编码_In         In 医疗卡类别.编码%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '  名称_In         In 医疗卡类别.名称%Type,
    strSQL = strSQL & "'" & str全名 & "',"
    '  短名_In         In 医疗卡类别.短名%Type,
    strSQL = strSQL & "'" & str短名 & "',"
    '  前缀文本_In     In 医疗卡类别.前缀文本%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  卡号长度_In     In 医疗卡类别.卡号长度%Type,
    strSQL = strSQL & "" & lng卡号长度 & ","
    '  缺省标志_In     In 医疗卡类别.缺省标志%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否固定_In     In 医疗卡类别.是否固定%Type,
    strSQL = strSQL & "1,"
    '  是否严格控制_In In 医疗卡类别.是否严格控制%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否自制_In     In 医疗卡类别.是否自制%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否存在帐户_In In 医疗卡类别.是否存在帐户%Type,
    strSQL = strSQL & "" & 0 & ","
    '  是否全退_In     In 医疗卡类别.是否全退%Type,
    strSQL = strSQL & "0,"
    '  部件_In         In 医疗卡类别.部件%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  备注_In         In 医疗卡类别.备注%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  特定项目_In     In 医疗卡类别.特定项目%Type,
    strSQL = strSQL & "'" & strCode & "',"
    '    收费细目id_In   In 收费项目目录.ID%Type,
    strSQL = strSQL & "" & "0" & ","
    '  结算方式_In     In 医疗卡类别.结算方式%Type,
    strSQL = strSQL & "'" & "" & "',"
    '  是否启用_In     In 医疗卡类别.是否启用%Type,
    strSQL = strSQL & "1,"
    '  卡号密文_In     In 医疗卡类别.卡号密文%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  是否重复使用_In In 医疗卡类别.是否重复使用%Type,
    strSQL = strSQL & "" & 1 & ","
    '密码长度_In     In 医疗卡类别.密码长度%Type,
    strSQL = strSQL & "" & 10 & ","
    '密码长度限制_In In 医疗卡类别.密码长度限制%Type,
    strSQL = strSQL & "" & 0 & ","
    '密码规则_In     In 医疗卡类别.密码规则%Type,
    strSQL = strSQL & "" & 0 & ","
    strSQL = strSQL & "" & 1 & ","
    '  操作方式_In     In Integer := 0
    strSQL = strSQL & "" & intOper & ","
    '是否模糊查找_In     In 医疗卡类别.是否模糊查找%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '问题号:51072
    '密码输入限制_In     In 医疗卡类别.密码输入限制%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '是否缺省密码_In     In 医疗卡类别.是否缺省密码%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '问题号:56508
    '是否制卡_In
    strSQL = strSQL & "" & 0 & ","
    '是否发卡_In
    strSQL = strSQL & "" & 0 & ","
    '是否写卡_In
    strSQL = strSQL & "" & 0 & ","
    '问题号:57697
    '险类_In
    strSQL = strSQL & "" & 0 & ","
    '问题号:57326
    strSQL = strSQL & "" & 1 & ","
    '77872,李南春,2014/12/3:是否支持转帐及代扣
    '是否转帐及代扣_In  In 医疗卡类别.是否转帐及代扣%Type:=0
    strSQL = strSQL & "" & 0 & ","
    '读卡性质_In       In 医疗卡类别.读卡性质%Type := '1000',
    strSQL = strSQL & "" & "1000" & ","
    '键盘控制方式_In   In 医疗卡类别.键盘控制方式%Type := 0,
    strSQL = strSQL & "" & 0 & ","
    '90875:李南春,2015/12/16,增加医疗卡证件类型
    '是否证件_In  In 医疗卡类别.是否证件%Type:=0
    strSQL = strSQL & "" & 1 & ")"
End Sub

Private Sub AddCertificate(ByVal lng病人ID As Long, ByRef arrSQL As Variant, ByVal dtCurdate As Date)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:建立证件卡类信息，如果是第一次建立卡类别
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, rsPatiCard As ADODB.Recordset
    Dim lngRow As Integer, lngCol As Integer
    Dim lngID As Long, strCode As String
    Dim colPro As Collection
    
    On Error GoTo ErrHand
    Set colPro = New Collection
    '绑定卡前要判断卡类别是否存在
    strSQL = "Select B.ID,B.编码,B.卡号长度,B.名称,A.卡号,A.病人ID,Decode(A.卡号 ,NULL,1,0) as 标识 from 病人医疗卡信息 A,医疗卡类别 B " & _
            "Where A.卡类别ID(+)=B.ID And B.是否证件=1 And A.状态(+)=0 And A.病人ID(+)=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    Set rsPatiCard = zlDatabase.CopyNewRec(rsTemp)
    With vsCertificate
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    lngID = 0: strCode = ""
                    rsTemp.Filter = "名称='" & .TextMatrix(lngRow, lngCol) & "'"
                    If rsTemp.RecordCount = 0 Then
                        lngID = zlDatabase.GetNextId("医疗卡类别")
                        If mstrFirstCode = "" Then
                            strCode = zlDatabase.GetMax("医疗卡类别", "编码", 4)
                            mstrFirstCode = strCode
                        Else
                            strCode = CStr(Val(mstrFirstCode) + 1)
                            strCode = Format(strCode, String(4, "0"))
                            mstrFirstCode = strCode
                        End If
                        Call AddCardTypeSQL(0, lngID, strCode, .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), strSQL)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    ElseIf Len(.TextMatrix(lngRow, lngCol + 1)) > Val(Nvl(rsTemp!卡号长度)) Then
                        Call AddCardTypeSQL(1, Val(Nvl(rsTemp!ID)), Nvl(rsTemp!编码), .TextMatrix(lngRow, lngCol), Left(.TextMatrix(lngRow, lngCol), 1), Len(.TextMatrix(lngRow, lngCol + 1)), strSQL)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    End If
                    
                    '进行证件卡绑定
                    rsPatiCard.Filter = "名称='" & .TextMatrix(lngRow, lngCol) & "' And 卡号='" & .TextMatrix(lngRow, lngCol + 1) & "'"
                    If rsPatiCard.RecordCount = 0 Then
                        'Zl_医疗卡变动_Insert
                         strSQL = "Zl_医疗卡变动_Insert("
                        '      变动类型_In   Number,
                        '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
                        strSQL = strSQL & "" & 11 & ","
                        '      病人id_In     住院费用记录.病人id%Type,
                        strSQL = strSQL & "" & lng病人ID & ","
                        '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
                        strSQL = strSQL & "" & IIf(lngID = 0, rsTemp!ID, lngID) & ","
                        '      原卡号_In     病人医疗卡信息.卡号%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      医疗卡号_In   病人医疗卡信息.卡号%Type,
                        strSQL = strSQL & "'" & Trim(.TextMatrix(lngRow, lngCol + 1)) & "',"
                        '      变动原因_In   病人医疗卡变动.变动原因%Type,
                        '      --变动原因_In:如果密码调整，变动原因为密码.加密的
                        strSQL = strSQL & "'" & "证件卡绑定" & "',"
                        '      密码_In       病人信息.卡验证码%Type,
                        strSQL = strSQL & "'" & "" & "',"
                        '      操作员姓名_In 住院费用记录.操作员姓名%Type,
                        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                        '      变动时间_In   住院费用记录.登记时间%Type,
                        strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                        '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
                        strSQL = strSQL & "'" & "" & "',"
                        '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
                        strSQL = strSQL & "NULL)"
                    
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    Else
                        rsPatiCard!标识 = 1
                        rsPatiCard.Update
                    End If
                End If
            Next
        Next
    End With
    mstrFirstCode = ""
    
    '卡号列表中没有证件号，要解除绑定
    rsPatiCard.Filter = "标识=0"
    If rsPatiCard.RecordCount > 0 Then
        rsPatiCard.MoveFirst
        Do While Not rsPatiCard.EOF
            'Zl_医疗卡变动_Insert
             strSQL = "Zl_医疗卡变动_Insert("
            '      变动类型_In   Number,
            '发卡类型=1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失)
            strSQL = strSQL & "" & 14 & ","
            '      病人id_In     住院费用记录.病人id%Type,
            strSQL = strSQL & "" & lng病人ID & ","
            '      卡类别id_In   病人医疗卡信息.卡类别id%Type,
            strSQL = strSQL & "" & rsPatiCard!ID & ","
            '      原卡号_In     病人医疗卡信息.卡号%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      医疗卡号_In   病人医疗卡信息.卡号%Type,
            strSQL = strSQL & "'" & rsPatiCard!卡号 & "',"
            '      变动原因_In   病人医疗卡变动.变动原因%Type,
            '      --变动原因_In:如果密码调整，变动原因为密码.加密的
            strSQL = strSQL & "'" & "证件卡取消绑定" & "',"
            '      密码_In       病人信息.卡验证码%Type,
            strSQL = strSQL & "'" & "" & "',"
            '      操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '      变动时间_In   住院费用记录.登记时间%Type,
            strSQL = strSQL & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            '      Ic卡号_In     病人信息.Ic卡号%Type := Null,
            strSQL = strSQL & "'" & "" & "',"
            '      挂失方式_In   病人医疗卡变动.挂失方式%Type := Null
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

Public Function IsCertificateCard(ByVal lng病人ID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:证件卡类检查
    '编制:李南春
    '时间:2015/12/17 17:37:27
    '问题:90875
    '-------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long, str证件类型 As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strCardName As String
    
    On Error GoTo ErrHand
    With vsCertificate
        '检查输入是否完整
        For lngRow = 1 To .Rows - 1
            For lngCol = 0 To .Cols - 1 Step 2
                If .TextMatrix(lngRow, lngCol) = "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    MsgBox "请选择卡号" & .TextMatrix(lngRow, lngCol + 1) & "的证件类型", vbInformation, gstrSysName
                    .Select lngRow, lngCol
                    Exit Function
                End If
                If .TextMatrix(lngRow, lngCol) <> "" And .TextMatrix(lngRow, lngCol + 1) <> "" Then
                    strSQL = "Select 1 from 病人医疗卡信息 A,医疗卡类别 B " & _
                            "Where A.卡类别ID=B.ID And B.名称=[1] And B.是否证件=1 And A.卡号=[2] And  A.病人ID<>[3]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .TextMatrix(lngRow, lngCol), Trim(.TextMatrix(lngRow, lngCol + 1)), lng病人ID)
                    If rsTmp.RecordCount > 0 Then
                        MsgBox .TextMatrix(lngRow, lngCol) & ":" & .TextMatrix(lngRow, lngCol + 1) & "正在被使用,请检查!", vbInformation, gstrSysName
                        .Select lngRow, lngCol
                        Exit Function
                    End If
                    str证件类型 = str证件类型 & ",'" & .TextMatrix(lngRow, lngCol) & "'"
                End If
            Next
        Next
        
        '检查证件类型是否与非证件的医疗卡类别重复，重复则不保存信息
        str证件类型 = Mid(str证件类型, 2)
        If str证件类型 = "" Then IsCertificateCard = True: Exit Function
        strSQL = "Select 名称 From 医疗卡类别 where 名称 in (" & str证件类型 & ") And Nvl(是否证件,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                strCardName = strCardName & "," & Nvl(rsTmp!名称)
            Loop
            
            strCardName = Mid(strCardName, 2)
            MsgBox "医疗卡类别【" & strCardName & "】名称重复,不能继续添加。", vbInformation, gstrSysName
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
'功能:'102232 zlPlugIn展示病人扶贫信息
'如果是HIS建档病人:返回值为T允许加载,返回值为F-禁止加载
'如果是HIS未建档的新病人,保存时返回T-允许保存,F-禁止保存,并清空界面
'未启用插件部件,或插件部件不包含该接口 缺省返回T-允许加载病人及保存新病人。
    Dim blnRet As Boolean
    
    blnRet = True
    If CreatePlugInOK(glngModul) And mbytInState <> EState.E查阅 Then
        On Error Resume Next
        blnRet = gobjPlugIn.PatiValiedCheck(glngSys, glngModul, 2, lngPatiID, 0, strXMLPati) 'T=成功;F-失败
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
'功能:检查当前手机号是否存在
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "SELECT 1 FROM 病人信息 Where 手机号 = [1] And 病人ID <> [2] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查手机号", strMobile, lngPatiID)
    If Not rsTemp Is Nothing Then
        CheckMobile = rsTemp.EOF = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub SelectYouBian(objTextBox As TextBox)
    '功能：邮编选择器
    Dim strInput As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI

    strInput = objTextBox.Text
    If strInput <> "" Then
        If zlCommFun.IsCharChinese(strInput) Then
            strSQL = strSQL & " And A.名称 Like [1] "
        Else
            strSQL = strSQL & " And A.简码 Like [1] "
        End If
    Else
        Exit Sub
    End If
    strSQL = "Select Rownum as ID,名称,简码,邮编  From 区域 A " & _
             "Where 邮编 is not null " & strSQL & " Order by 编码"
    vPoint = GetCoordPos(objTextBox.hWnd, 0, 0)
    Set rsTmp = zlDatabase.ShowSQLSelect(objTextBox.Parent, strSQL, 0, "邮编", False, "", "", False, _
        False, True, vPoint.X, vPoint.Y, objTextBox.Height, False, False, False, UCase(strInput) & "%")
    If Not rsTmp Is Nothing Then
        objTextBox.Text = rsTmp!邮编 & ""
    End If
End Sub


Private Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Sub CreatePublicExpenseObject(ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共费用部件
    '入参:
    '编制:
    '日期:
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If mobjPublicExpense Is Nothing Then
        Set mobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If mobjPublicExpense Is Nothing Then Exit Sub
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If mobjPublicExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    mintPriceGradeStartType = mobjPublicExpense.zlGetPriceGradeStartType()
    If mintPriceGradeStartType = 0 Then Exit Sub
    '读取站点价格等级
    Call mobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", , , mstrPriceGrade)
End Sub

Private Sub ReLoadCardFee(Optional ByVal blnFeedName As Boolean)
    '118124:李南春，2018/1/18，获取卡费
    Dim lng病人ID As Long, lng收费细目ID As Long
    Dim strSQL As String, str年龄 As String
    Dim rsTmp As ADODB.Recordset
    
    If mCurSendCard.rs卡费 Is Nothing Then Exit Sub
    If mCurSendCard.rs卡费.RecordCount = 0 Then Exit Sub
    If mCurSendCard.lng卡类别ID = 0 Then Exit Sub
    If Trim(txtPatient.Text) = "" Or Trim(txt卡号.Text) = "" Then Exit Sub
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = mrsInfo!病人ID
    End If
    If blnFeedName = False And lng病人ID <> 0 Then Exit Sub
    
    str年龄 = Trim(txt年龄.Text)
    If IsNumeric(str年龄) Then str年龄 = str年龄 & cbo年龄单位.Text
    mCurSendCard.rs卡费.MoveFirst
    
    strSQL = "Select Zl1_Ex_CardFee([1],[2],[3],[4],[5],[6],[7],[8],[9]) as 收费细目ID From Dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "卡费", mlngModul, mCurSendCard.lng卡类别ID, Trim(txt卡号.Text), lng病人ID, _
                Trim(txtPatient.Text), zlStr.NeedName(cbo性别.Text), str年龄, txt身份证号.Text, Val(Nvl(mCurSendCard.rs卡费!收费细目ID)))
    If rsTmp.EOF Then Exit Sub
    
    lng收费细目ID = Val(Nvl(rsTmp!收费细目ID))
    Set rsTmp = zlGetSpecialItemFee(mCurSendCard.str特定项目, mstrPriceGrade, lng收费细目ID)
    If Not rsTmp Is Nothing Then Set mCurSendCard.rs卡费 = rsTmp
    
    With mCurSendCard.rs卡费
        txt卡额.Text = Format(IIf(Val(Nvl(!是否变价)) = 1, Val(Nvl(!缺省价格)), Val(Nvl(!现价))), "0.00")
        txt卡额.Tag = txt卡额.Text  '保持不变
        txt卡额.Locked = Not (Val(Nvl(!是否变价)) = 1)
        txt卡额.TabStop = (Val(Nvl(!是否变价)) = 1)
        
        If mCurSendCard.rs卡费!是否变价 = 0 And Val(txt卡额.Text) <> 0 Then
            txt卡额.Text = Format(GetActualMoney(zlStr.NeedName(cbo费别.Text), mCurSendCard.rs卡费!收入项目ID, mCurSendCard.rs卡费!现价, mCurSendCard.rs卡费!收费细目ID), "0.00")
        End If
    End With
End Sub
