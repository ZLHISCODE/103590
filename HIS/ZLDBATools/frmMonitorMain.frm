VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMonitorMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "性能监控"
   ClientHeight    =   12555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16575
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12555
   ScaleWidth      =   16575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cboInst 
      Height          =   300
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   510
      Width           =   1215
   End
   Begin VB.PictureBox pctASH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   600
      ScaleHeight     =   3735
      ScaleWidth      =   3735
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   3735
      Begin SHDocVwCtl.WebBrowser webASH 
         Height          =   2295
         Left            =   480
         TabIndex        =   22
         Top             =   720
         Width           =   2175
         ExtentX         =   3836
         ExtentY         =   4048
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "修改缺省目录(&C)"
      Height          =   350
      Left            =   11880
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存至本地(&S)"
      Height          =   350
      Left            =   10200
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdGetReport 
      Caption         =   "查看AWR报告"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   17
      Top             =   485
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox pctLine 
      Height          =   50
      Left            =   9000
      MousePointer    =   7  'Size N S
      ScaleHeight     =   50
      ScaleMode       =   0  'User
      ScaleWidth      =   5055
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox pctAddm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   11640
      ScaleHeight     =   5535
      ScaleWidth      =   3255
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   3255
      Begin VSFlex8Ctl.VSFlexGrid vsfAddm 
         Height          =   3000
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   11895
         _cx             =   20981
         _cy             =   5292
         Appearance      =   2
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   0
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin RichTextLib.RichTextBox rtxAddm 
         Height          =   615
         Left            =   0
         TabIndex        =   13
         Top             =   3360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   1085
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMonitorMain.frx":0000
      End
   End
   Begin VB.PictureBox pctReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5610
      Left            =   5520
      ScaleHeight     =   5610
      ScaleWidth      =   4500
      TabIndex        =   11
      ToolTipText     =   "请在表格中选取两个时间节点作为报告的开始和结束点。"
      Top             =   3480
      Visible         =   0   'False
      Width           =   4500
      Begin SHDocVwCtl.WebBrowser webReport 
         Height          =   1815
         Left            =   480
         TabIndex        =   18
         Top             =   2400
         Width           =   3255
         ExtentX         =   5741
         ExtentY         =   3201
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfReport 
         Height          =   1560
         Left            =   -240
         TabIndex        =   15
         ToolTipText     =   "请在表格中选取两个时间节点作为报告的开始和结束点。"
         Top             =   0
         Width           =   12135
         _cx             =   21405
         _cy             =   2752
         Appearance      =   2
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   26
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   0
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
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H80000009&
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   485
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tabPage 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "数据库时间"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "重做日志"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "AWR报告"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ASH报告"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ADDM报告"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.PictureBox pctDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2535
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   2880
      Width           =   4155
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "2017/8/15至2017/8/22折线图"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   3840
      End
   End
   Begin VB.PictureBox pctData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1455
      ScaleWidth      =   14055
      TabIndex        =   0
      Top             =   960
      Width           =   14055
      Begin VSFlex8Ctl.VSFlexGrid vsfData 
         Height          =   1815
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   12735
         _cx             =   22463
         _cy             =   3201
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
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
      Begin VSFlex8Ctl.VSFlexGrid vsfData 
         Height          =   1815
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   12375
         _cx             =   21828
         _cy             =   3201
         Appearance      =   1
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
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
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   510
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy/MM/dd HH:mm"
      Format          =   115081219
      CurrentDate     =   42961
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Top             =   510
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy/MM/dd HH:mm"
      Format          =   115081219
      CurrentDate     =   42961
   End
   Begin VB.Label lblInst 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "实例"
      Height          =   180
      Left            =   5400
      TabIndex        =   24
      Top             =   570
      Width           =   360
   End
   Begin VB.Label lblPath 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "当前默认路径"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   13680
      TabIndex        =   23
      Top             =   565
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "快照时间范围"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label lblRange 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "至"
      Height          =   180
      Left            =   3240
      TabIndex        =   7
      Top             =   570
      Width           =   180
   End
End
Attribute VB_Name = "frmMonitorMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnTabInit(5) As Boolean      '标记选项卡是否初始化
Private mintActiveTab As Integer   '标记当前活跃的选项卡
Private mrsActiveData(2) As New ADODB.Recordset    '标记当前活跃选项卡对应的数据集,用于即时重绘
Private mstrAWRSelected As String  '记录表格的选中行，用于获取AWR报告 ,记录形式如： _1,_2,_3,
Private mstrActiveFile(3) As String '查看的文件
Private mdblDBID  As Double
Private mstrADDMSelected As String

Private Enum TabNum
    tabDBTime = 0
    tabRedoLog = 1
    tabAWR
    tabASH
    tabAddm
End Enum

Private Enum PicColor
    坐标系Color = &H454545
    dblYellow = &H62C6F0
    选中Color = &H8000000D
    选中字体Color = &H8000000E
    未选中Color = &H80000005
    未选中字体Color = &H80000008
End Enum

Private Const strTblLog = "Date,1,1;Day,1,1;MaxValue,1,1;H1,1,1;H2,1,1;H3,1,1;H4,1,1;H5,1,1;H6,1,1;H7,1,1;H8,1,1;" & _
                                        "H9,1,1;H10,1,1;H11,1,1;H12,1,1;H13,1,1;H14,1,1;H15,1,1;H16,1,1;H17,1,1;H18,1,1;H19,1,1;H20,1,1;H21,1,1;H22,1,1;H23,1,1;H24,1,1"
Private Const strTblReport = "Snap_Id,1,1;Dbid,1,1;Instance_Number,1,1;Begin_Interval_Time,1,1;End_Interval_Time,1,1;Startup_Time,1,1"
Private Const strTblADDM = "Owner,1,1;Task_Id,1,1;Task_Name,1,1;Advisor_Name,1,1;Created,1,1;Description,1,1"


Public Sub ShowMe()
    Me.Show
End Sub

Private Sub cmdGetReport_Click()
    Dim strSql As String, rstmp As ADODB.Recordset, strTmp As String
    Dim intRow1 As Integer, intRow2 As Integer
    Dim lngBid As Long, lngEid As Long
    Dim strETime As String, strBTime As String
    Dim objFile As New FileSystemObject, strFileName As String
    
    '如果没有设置默认路径，需要先进行设置
    If gstrFilePath = "" Then
        cmdPath_Click
    End If
    
    Select Case cmdGetReport.Caption
        Case "查看AWR报告"
            intRow1 = Replace(Split(mstrAWRSelected, ",")(0), "_", "")
            intRow2 = Replace(Split(mstrAWRSelected, ",")(1), "_", "")
            With vsfReport
                If .TextMatrix(intRow1, .ColIndex("Dbid")) <> .TextMatrix(intRow1, .ColIndex("Dbid")) Then
                    MsgBox "请选择相同的DBID。"
                    Exit Sub
                End If
    
                lngBid = .TextMatrix(IIf(intRow1 > intRow2, intRow1, intRow2), .ColIndex("Snap_Id"))
                lngEid = .TextMatrix(IIf(intRow1 < intRow2, intRow1, intRow2), .ColIndex("Snap_Id"))
                strETime = .TextMatrix(IIf(intRow1 < intRow2, intRow1, intRow2), .ColIndex("Begin_Interval_Time"))
                strBTime = .TextMatrix(IIf(intRow1 > intRow2, intRow1, intRow2), .ColIndex("Begin_Interval_Time"))
                strFileName = "AWR" & Format(strBTime, "yyyy-MM-dd_hh") & "-" & Format(strETime, "hh") & ".html"
            End With
             
            On Error GoTo errH
            'DBID INSTANCEID BID EID
            Screen.MousePointer = vbArrowHourglass
            strSql = "Select Output From Table(Dbms_Workload_Repository.Awr_Report_Html([1], [2], [3], [4]))"
            Set rstmp = OpenSQLRecord(strSql, "GetAWR", mdblDBID, cboInst.ItemData(cboInst.ListIndex), lngBid, lngEid)
            
            Do While Not rstmp.EOF
                strTmp = strTmp & rstmp!OutPut & ""
                rstmp.MoveNext
            Loop
            
            objFile.CreateTextFile(gstrFilePath & strFileName).Write strTmp
            webReport.Navigate "file:///" & Replace(gstrFilePath, "\", "/") & strFileName
            mstrActiveFile(0) = "已保存当前文件：" & gstrFilePath & strFileName
            lblPath.Caption = mstrActiveFile(0)
            
        Case "查看ASH报告"
            strFileName = "ASH" & Format(dtpStart.Value, "yyyy-MM-dd_hhmm") & "-" & Format(dtpEnd.Value, "hhmm") & ".html"
            
            On Error GoTo errH
            Screen.MousePointer = vbArrowHourglass
            strSql = "Select Output From Table (Dbms_Workload_Repository.Ash_Report_Html([1], [2], [3],[4]))"
            Set rstmp = OpenSQLRecord(strSql, "GetASH", mdblDBID, cboInst.ItemData(cboInst.ListIndex), _
                                    CDate(Format(dtpStart.Value, "yyyy-MM-dd hh:mm:ss")), CDate(Format(dtpEnd.Value, "yyyy-MM-dd hh:mm:ss")))
            
            Do While Not rstmp.EOF
                strTmp = strTmp & rstmp!OutPut & ""
                rstmp.MoveNext
            Loop
            
            objFile.CreateTextFile(gstrFilePath & strFileName).Write strTmp
            webASH.Navigate "file:///" & Replace(gstrFilePath, "\", "/") & strFileName

            mstrActiveFile(1) = "已保存当前文件：" & gstrFilePath & strFileName
            lblPath.Caption = mstrActiveFile(1)
        Case "查看ADDM报告"
            
            On Error GoTo errH
            Screen.MousePointer = vbArrowHourglass
            strSql = "Select Dbms_Advisor.Get_Task_Report([1], 'TEXT', 'ALL') OutPut From Dual"
            mstrADDMSelected = vsfAddm.TextMatrix(vsfAddm.Row, vsfAddm.ColIndex("Task_Name")) & ""
            If mstrADDMSelected = "Task_Name" Or vsfAddm.Row = 0 Or vsfAddm.Row = -1 Then
                MsgBox "请先选中一个ADDM报告。"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            Set rstmp = OpenSQLRecord(strSql, "GetADDM", mstrADDMSelected)
            
            'ADDM文件名
            mstrADDMSelected = "ADDM" & Format(vsfAddm.TextMatrix(vsfAddm.Row, vsfAddm.ColIndex("Created")), "yyyymmdd") _
                                            & "_" & vsfAddm.TextMatrix(vsfAddm.Row, vsfAddm.ColIndex("Task_ID")) & ".txt"
            
            If rstmp.RecordCount = 0 Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            rtxAddm.Text = rstmp!OutPut
                                                
            cmdSave.Enabled = Not rtxAddm.Text = ""
            
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
errH:
    Screen.MousePointer = vbDefault

    If InStr(Err.Description, "ORA-01843") Then
        MsgBox "所选时间超过快照保留期，无法生成报告。"
        Exit Sub
    End If
    If InStr(Err.Description, "ORA-20019") Then
        MsgBox "请选择StartUpTime相同的快照生成报告。"
        Exit Sub
    End If
    If InStr(Err.Description, "ORA-13605") Then
        MsgBox "当前用户没有查询ADDM报告的权限，请使用SYS用户。"
        Exit Sub
    End If
    
    MsgBox Err.Description
End Sub

Private Sub cmdPath_Click()
    Dim strTmp As String
    
    strTmp = OpenFolder(Me, "请选择默认保存路径", gstrFilePath)
    If strTmp = "" Then
        Exit Sub
    End If
    gstrFilePath = strTmp & IIf(Right(strTmp, 1) <> "\", "\", "")
    Call SaveSetting("zlMonitor", "Setting", "DefaultPath", gstrFilePath)
    lblPath.Caption = "缺省保存路径为：" & strTmp
End Sub

Private Sub cmdRefresh_Click()

    Select Case mintActiveTab
        Case tabDBTime
            lblTitle.Caption = dtpStart.Value & "至" & dtpEnd.Value & "数据库耗时变化情况"
            Call loadData(mintActiveTab)
        Case tabRedoLog
            lblTitle.Caption = dtpStart.Value & "至" & dtpEnd.Value & "重做日志产生数量变化情况"
            Call loadData(mintActiveTab)
        Case tabAWR, tabAddm
            
            If mintActiveTab = tabAWR Then
                mstrAWRSelected = ""
                cmdGetReport.Enabled = False
            End If
            
            Call LoadReport(mintActiveTab)
    End Select

End Sub


Private Sub cmdSave_Click()
    Dim objFile As New FileSystemObject
    
    If gstrFilePath = "" Then
        MsgBox "未设定缺省路径，保存失败。"
        Exit Sub
    End If
    
    objFile.CreateTextFile(gstrFilePath & mstrADDMSelected).Write rtxAddm.Text
    mstrActiveFile(2) = "已保存当前文件：" & gstrFilePath & mstrADDMSelected
    lblPath.Caption = mstrActiveFile(2)
End Sub

Private Sub dtpStart_Change()
    If mintActiveTab <> tabASH Then Exit Sub
    
    On Error Resume Next
    dtpEnd.Value = dtpStart.Value + 1 / 24 / 4
End Sub

Private Sub Form_load()

    dtpStart.Value = Date - 2
    dtpEnd.Value = Date + 1
    lblTitle.Caption = dtpStart.Value & "至" & dtpEnd.Value & "数据库耗时变化情况"
    Call InitCboList
    Call InitTable(vsfData(tabDBTime), strTblLog)
    Call InitTable(vsfData(tabRedoLog), strTblLog)
    Call InitTable(vsfReport, strTblReport)
    Call InitTable(vsfAddm, strTblADDM)
    loadData (mintActiveTab)
    
    webASH.Navigate "about:blank"
    webReport.Navigate "about:blank"
    
    '从注册表读取路径
    gstrFilePath = GetSetting("zlMonitor", "Setting", "DefaultPath")
    
    If gstrFilePath = "" Then
        mstrActiveFile(0) = "": mstrActiveFile(1) = "": mstrActiveFile(2) = ""
        lblPath.Caption = mstrActiveFile(0)
    Else
        mstrActiveFile(0) = "当前文件的保存路径为：" & Mid(gstrFilePath, 1, Len(gstrFilePath) - 1)
        mstrActiveFile(1) = "当前文件的保存路径为：" & Mid(gstrFilePath, 1, Len(gstrFilePath) - 1)
        mstrActiveFile(2) = "当前文件的保存路径为：" & Mid(gstrFilePath, 1, Len(gstrFilePath) - 1)
        lblPath.Caption = mstrActiveFile(0)
    End If
    
End Sub

Private Sub Form_Resize()
    Dim intIndex As Integer
    
    On Error Resume Next
    
    tabPage.Width = Me.ScaleWidth
    pctLine.Width = Me.ScaleWidth: pctLine.Left = 0
    
    cmdRefresh.Left = (IIf(cboInst.Visible, cboInst.Width + cboInst.Left + 45, dtpEnd.Width + dtpEnd.Left + 45))
    cmdGetReport.Left = (IIf(cmdRefresh.Visible, cmdRefresh.Width + cmdRefresh.Left + 45, (IIf(cboInst.Visible, cboInst.Width + cboInst.Left + 45, dtpEnd.Width + dtpEnd.Left + 45))))
    cmdSave.Left = cmdGetReport.Width + cmdGetReport.Left + 45
    cmdPath.Left = IIf(cmdSave.Visible, cmdSave.Width + cmdSave.Left + 45, cmdGetReport.Width + cmdGetReport.Left + 45)
    lblPath.Left = cmdPath.Left + cmdPath.Width + 45
    Select Case mintActiveTab
        Case tabDBTime, tabRedoLog
            pctData.Top = cmdRefresh.Top + cmdRefresh.Height + 75
            pctData.Left = 0
            pctData.Width = Me.ScaleWidth
    
            pctDraw.Top = pctData.Top + pctData.Height
            pctDraw.Height = Me.ScaleHeight - pctData.Top - pctData.Height
            pctDraw.Width = Me.ScaleWidth
            
            DrawPicture pctDraw, mrsActiveData(mintActiveTab), 3, 26, True
        Case tabAWR
            pctReport.Width = Me.ScaleWidth
            pctReport.Top = cmdRefresh.Top + cmdRefresh.Height + 75
            pctReport.Height = Me.ScaleHeight - pctReport.Top
            pctReport.Left = 0
            pctLine.Top = pctReport.Top + vsfReport.Height
        Case tabASH
            pctASH.Width = Me.ScaleWidth
            pctASH.Top = cmdRefresh.Top + cmdRefresh.Height + 75
            pctASH.Height = Me.ScaleHeight - pctASH.Top
            pctASH.Left = 0
        Case tabAddm
            pctAddm.Left = 0
            pctAddm.Top = cmdRefresh.Top + cmdRefresh.Height + 75
            pctAddm.Width = Me.ScaleWidth: pctAddm.Height = Me.ScaleHeight
            pctLine.Top = pctAddm.Top + vsfAddm.Height
    End Select
    
End Sub

Private Sub loadData(Index As Integer)
'功能：根据传入的Tab号加载表格数据
    Dim strSql As String
    Dim i As Integer, j As Integer
    
    On Error GoTo errH
    
    If Not CheckDate Then Exit Sub

    If Index = tabDBTime Then
        '加载DB TIME
        strSql = "Select To_Char(Trunc(t.Begin_Interval_Time), 'yyyy/mm/dd') ""Date"", To_Char(t.Begin_Interval_Time, 'Dy') ""Day""," & vbNewLine & _
                        "       Nvl(Max(t.Value), 0) Maxvalue," & vbNewLine & _
                        "       Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '01', t.Value, 0)) H1,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '02', t.Value, 0)) H2, Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '03', t.Value, 0)) H3," & vbNewLine & _
                        "       Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '04', t.Value, 0)) H4,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '05', t.Value, 0)) H5, Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '06', t.Value, 0)) H6," & vbNewLine & _
                        "       Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '07', t.Value, 0)) H7,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '08', t.Value, 0)) H8,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '09', t.Value, 0)) H9," & vbNewLine & _
                        "       Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '10', t.Value, 0)) H10,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '11', t.Value, 0)) H11,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '12', t.Value, 0)) H12," & vbNewLine & _
                        "       Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '13', t.Value, 0)) H13,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '14', t.Value, 0)) H14,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '15', t.Value, 0)) H15," & vbNewLine & _
                        "       Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '16', t.Value, 0)) H16,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '17', t.Value, 0)) H17,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '18', t.Value, 0)) H18," & vbNewLine & _
                        "       Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '19', t.Value, 0)) H19,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '20', t.Value, 0)) H20,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '21', t.Value, 0)) H21," & vbNewLine & _
                        "       Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '22', t.Value, 0)) H22,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '23', t.Value, 0)) H23,Sum(Decode(To_Char(t.Begin_Interval_Time, 'hh24'), '0', t.Value, 0)) H24" & vbNewLine & _
                        "From" & vbNewLine & _
                        "(Select a.Dbid, b.Startup_Time, b.Begin_Interval_Time," & vbNewLine & _
                        "        To_Char((a.Value - Lag(a.Value) Over(Partition By Trunc(b.Begin_Interval_Time),  b.Startup_Time Order By b.Begin_Interval_Time)) /1000000, '999999.9') As Value" & vbNewLine & _
                        "From Dba_Hist_Sys_Time_Model A, Dba_Hist_Snapshot B" & vbNewLine & _
                        "Where a.Stat_Name = 'DB time' And a.Snap_Id = b.Snap_Id And a.Dbid = b.Dbid And" & vbNewLine & _
                        "        a.Instance_Number = b.Instance_Number And b.Begin_Interval_Time Between [1] And [2] And a.Instance_Number = [3]) T" & vbNewLine & _
                        "Group By Trunc(t.Begin_Interval_Time), To_Char(t.Begin_Interval_Time, 'Dy') Order By Trunc(t.Begin_Interval_Time) Desc "


    Else
        strSql = "Select ""Date"",""Day""," & vbNewLine & _
                        "Greatest(""h24"" ,""h1"",""h2"",""h3"",""h4"",""h5"",""h6"",""h7"" ,""h8"",""h9"" ,""h10"" ,""h11"" ,""h12"" ,""h13"" ,""h14"" ,""h15"" ,""h16"" ,""h17"" ,""h18"" ,""h19"" ,""h20"" ,""h21"" ,""h22"",""h23"" ) MaxValue," & vbNewLine & _
                        """h1"",""h2"",""h3"",""h4"",""h5"",""h6"",""h7"" ,""h8"",""h9"" ,""h10"" ,""h11"" ,""h12"" ,""h13"" ,""h14"" ,""h15"" ,""h16"" ,""h17"" ,""h18"" ,""h19"" ,""h20"" ,""h21"" ,""h22"",""h23"",""h24""" & vbNewLine & _
                        "from" & vbNewLine & _
                        "(Select To_Char(Trunc(First_Time), 'yyyy/mm/dd') ""Date"", To_Char(First_Time, 'Dy') ""Day""," & vbNewLine & _
                        "       Sum(Decode(To_Char(First_Time, 'hh24'), '00', 1, 0)) ""h24"",Sum(Decode(To_Char(First_Time, 'hh24'), '01', 1, 0)) ""h1"",Sum(Decode(To_Char(First_Time, 'hh24'), '02', 1, 0)) ""h2""," & vbNewLine & _
                        "       Sum(Decode(To_Char(First_Time, 'hh24'), '03', 1, 0)) ""h3"",Sum(Decode(To_Char(First_Time, 'hh24'), '04', 1, 0)) ""h4"",Sum(Decode(To_Char(First_Time, 'hh24'), '05', 1, 0)) ""h5""," & vbNewLine & _
                        "       Sum(Decode(To_Char(First_Time, 'hh24'), '06', 1, 0)) ""h6"",Sum(Decode(To_Char(First_Time, 'hh24'), '07', 1, 0)) ""h7"",Sum(Decode(To_Char(First_Time, 'hh24'), '08', 1, 0)) ""h8""," & vbNewLine & _
                        "       Sum(Decode(To_Char(First_Time, 'hh24'), '09', 1, 0)) ""h9"",Sum(Decode(To_Char(First_Time, 'hh24'), '10', 1, 0)) ""h10"",Sum(Decode(To_Char(First_Time, 'hh24'), '11', 1, 0)) ""h11""," & vbNewLine & _
                        "       Sum(Decode(To_Char(First_Time, 'hh24'), '12', 1, 0)) ""h12"",Sum(Decode(To_Char(First_Time, 'hh24'), '13', 1, 0)) ""h13"",Sum(Decode(To_Char(First_Time, 'hh24'), '14', 1, 0)) ""h14""," & vbNewLine & _
                        "       Sum(Decode(To_Char(First_Time, 'hh24'), '15', 1, 0)) ""h15"",Sum(Decode(To_Char(First_Time, 'hh24'), '16', 1, 0)) ""h16"",Sum(Decode(To_Char(First_Time, 'hh24'), '17', 1, 0)) ""h17""," & vbNewLine & _
                        "       Sum(Decode(To_Char(First_Time, 'hh24'), '18', 1, 0)) ""h18"",Sum(Decode(To_Char(First_Time, 'hh24'), '19', 1, 0)) ""h19"",Sum(Decode(To_Char(First_Time, 'hh24'), '20', 1, 0)) ""h20""," & vbNewLine & _
                        "       Sum(Decode(To_Char(First_Time, 'hh24'), '21', 1, 0)) ""h21"",Sum(Decode(To_Char(First_Time, 'hh24'), '22', 1, 0)) ""h22"",Sum(Decode(To_Char(First_Time, 'hh24'), '23', 1, 0)) ""h23""" & vbNewLine & _
                        "From GV$log_History" & vbNewLine & _
                        "Where  First_Time Between [1]  And [2] And Inst_ID =[3] " & vbNewLine & _
                        "Group By Trunc(First_Time), To_Char(First_Time, 'Dy')" & vbNewLine & _
                        ")" & vbNewLine & _
                        "Order By ""Date"" Desc"

    End If
    
    Set mrsActiveData(mintActiveTab) = OpenSQLRecord(strSql, "zlMonitor", CDate(Format(dtpStart.Value, "yyyy-MM-dd hh:mm:ss")) _
                                                                        , CDate(Format(dtpEnd.Value, "yyyy-MM-dd hh:mm:ss")), cboInst.ItemData(cboInst.ListIndex))
    
    If mrsActiveData(mintActiveTab).RecordCount = 0 Then
        lblTitle.Caption = dtpStart.Value & "至" & dtpEnd.Value & "没有产生日志"
    End If
    
    With vsfData(Index)
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + mrsActiveData(mintActiveTab).RecordCount
        Randomize
        i = 0
        Do While Not mrsActiveData(mintActiveTab).EOF
            For j = 0 To 26
                If .ColIndex("Date") = j Then
                    .TextMatrix(i + 1, j) = Format(mrsActiveData(mintActiveTab).Fields(j).Value, "yyyy/MM/dd hh:mm:ss")
                Else
                    .TextMatrix(i + 1, j) = Format(mrsActiveData(mintActiveTab).Fields(j).Value, "0.0")
                End If
                If j = 0 Then
                    .Cell(flexcpData, i + 1, j) = RGB(Int((255 * Rnd) + 1), Int((255 * Rnd) + 1), Int((255 * Rnd) + 1))
                End If
            Next
            i = i + 1
            mrsActiveData(mintActiveTab).MoveNext
        Loop
        
        .AutoSize 0, 26
        .Redraw = flexRDDirect
    End With
    
    DrawPicture pctDraw, mrsActiveData(mintActiveTab), 3, 26, True
    Exit Sub
errH:
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub


Private Sub pctAddm_Resize()
    On Error Resume Next
    vsfAddm.Top = 0: vsfAddm.Left = 0
    vsfAddm.Width = pctAddm.ScaleWidth
    
    rtxAddm.Left = 0: rtxAddm.Top = vsfAddm.Height + 125
    rtxAddm.Width = pctAddm.ScaleWidth
    rtxAddm.Height = pctAddm.ScaleHeight - rtxAddm.Top - 130

End Sub

Private Sub pctASH_resize()
    webASH.Top = 0: webASH.Left = 0
    webASH.Width = pctASH.ScaleWidth
    webASH.Height = pctASH.ScaleHeight
End Sub

Private Sub pctData_Resize()
    vsfData(0).Height = pctData.ScaleHeight
    vsfData(0).Width = pctData.ScaleWidth
    vsfData(1).Height = pctData.ScaleHeight
    vsfData(1).Width = pctData.ScaleWidth
End Sub

Private Sub pctLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    Select Case mintActiveTab
        Case tabAWR
            '防止过度拖拽，控件消失
            If vsfReport.Height + y < 100 Or webReport.Top + y > pctReport.ScaleHeight - 1000 Then Exit Sub
            
            pctLine.Top = pctLine.Top + y
            vsfReport.Height = vsfReport.Height + y
            webReport.Top = webReport.Top + y
            webReport.Height = pctReport.ScaleHeight - webReport.Top
        Case tabAddm
            If vsfAddm.Height + y < 100 Or rtxAddm.Top + y > pctAddm.ScaleHeight - 1000 Then Exit Sub
            pctLine.Top = pctLine.Top + y
            vsfAddm.Height = vsfAddm.Height + y
            rtxAddm.Top = rtxAddm.Top + y
            rtxAddm.Height = pctAddm.ScaleHeight - rtxAddm.Top
    End Select
    Me.Refresh
End Sub

Private Sub pctReport_Resize()
    On Error Resume Next
    vsfReport.Top = 0: vsfReport.Left = 0
    vsfReport.Width = pctReport.ScaleWidth
    webReport.Top = vsfReport.Height + 125
    webReport.Left = 0
    webReport.Width = pctReport.ScaleWidth
    webReport.Height = pctReport.ScaleHeight - webReport.Top
    
End Sub

Private Sub tabPage_Click()
    Dim intIndex As Integer, i As Integer

    intIndex = tabPage.SelectedItem.Index - 1
    If mintActiveTab = intIndex Then Exit Sub
    mintActiveTab = intIndex
    
    Select Case intIndex
        
        Case tabDBTime, tabRedoLog
            '界面控制
            pctData.Visible = True: pctDraw.Visible = True
            
            If intIndex = tabDBTime Then
                lblDate.Caption = "快照时间范围"
                vsfData(tabDBTime).Visible = True
                vsfData(tabRedoLog).Visible = False
            Else
                lblDate.Caption = "日志时间范围"
                vsfData(tabRedoLog).Visible = True
                vsfData(tabDBTime).Visible = False
            End If
            pctReport.Visible = False: pctASH.Visible = False: pctAddm.Visible = False
            pctLine.Visible = False
            cmdGetReport.Visible = False: cmdPath.Visible = False: cmdSave.Visible = False
            lblPath.Visible = False
            cmdRefresh.Visible = True
            lblTitle.Caption = dtpStart.Value & "至" & dtpEnd.Value & IIf(intIndex = tabDBTime, "数据库耗时变化情况", "重做日志产生数量变化情况")
            
            '功能控制
            loadData intIndex
        Case tabAWR
            '界面控制
            cmdGetReport.Visible = True: cmdPath.Visible = True: cmdSave.Visible = True
            pctLine.Visible = True: pctReport.Visible = True
            cmdGetReport.Enabled = Len(mstrAWRSelected) - Len(Replace(mstrAWRSelected, ",", "")) = 2
            lblPath.Visible = True: cmdSave.Visible = False
            cmdRefresh.Visible = True
            
            pctASH.Visible = False: pctAddm.Visible = False
            pctData.Visible = False: pctDraw.Visible = False
            
            '功能控制
            cmdGetReport.Caption = "查看AWR报告"
            lblDate.Caption = "快照时间范围"
            If Not mblnTabInit(intIndex) Then
                Call LoadReport(intIndex)
            End If
            lblPath.Caption = mstrActiveFile(0)
        Case tabASH
            '界面控制
            cmdGetReport.Visible = True: cmdPath.Visible = True: cmdSave.Visible = True
            cmdGetReport.Enabled = True
            pctASH.Visible = True
            lblPath.Visible = True: cmdSave.Visible = False
            cmdRefresh.Visible = False
            
            pctAddm.Visible = False: pctLine.Visible = False
            pctData.Visible = False: pctDraw.Visible = False: pctReport.Visible = False
            '功能控制
            cmdGetReport.Caption = "查看ASH报告"
            lblDate.Caption = "快照时间范围"
            If Not mblnTabInit(intIndex) Then
                Call LoadReport(intIndex)
            End If
            lblPath.Caption = mstrActiveFile(1)
        Case tabAddm
            '界面控制
            cmdGetReport.Visible = True: cmdPath.Visible = True: cmdSave.Visible = True
            cmdGetReport.Enabled = True
            pctAddm.Visible = True
            lblPath.Visible = True
            cmdRefresh.Visible = True
            pctLine.Visible = True
            
            pctASH.Visible = False: pctReport.Visible = False
            pctData.Visible = False: pctDraw.Visible = False
            '功能控制
            cmdGetReport.Caption = "查看ADDM报告"
            lblDate.Caption = "任务时间范围"
            cmdSave.Enabled = Not rtxAddm.Text = ""
            If Not mblnTabInit(intIndex) Then
                Call LoadReport(intIndex)
            End If
            lblPath.Caption = mstrActiveFile(2)
    End Select
    
    If Not mblnTabInit(intIndex) Then
        mblnTabInit(intIndex) = True
    End If
    Call Form_Resize
End Sub

Private Sub LoadReport(intIndex As Integer)
    '功能：根据传入的Index加载不同报告
    '参数：intIndex  :   3-AWR 4-ASH 5-ADDM
    Dim strSql As String, rsData As ADODB.Recordset
    Dim i As Integer
    
    
    If Not CheckDate Then Exit Sub
    On Error GoTo errH:
    Select Case intIndex
        Case tabAWR
            strSql = "Select Begin_Interval_Time, End_Interval_Time, Snap_Id, Dbid, Instance_Number, Snap_Level,Startup_Time" & vbNewLine & _
                            "From Dba_Hist_Snapshot" & vbNewLine & _
                            "Where Begin_Interval_Time Between [1] And [2] And Instance_Number = [3]" & vbNewLine & _
                            "Order By Snap_Id Desc"
            Set rsData = OpenSQLRecord(strSql, "LoadReport", CDate(Format(dtpStart.Value, "yyyy-MM-dd hh:mm:ss")), _
                                CDate(Format(dtpEnd.Value + 1, "yyyy-MM-dd hh:mm:ss")), cboInst.ItemData(cboInst.ListIndex))
    
            With vsfReport
                .Redraw = flexRDNone
                .Rows = .FixedRows
                .Rows = .FixedRows + rsData.RecordCount
                i = 0
                Do While Not rsData.EOF
                    .TextMatrix(i + 1, .ColIndex("Begin_Interval_Time")) = "" & Format(rsData!Begin_Interval_Time, "yyyy-MM-dd hh:mm:ss")
                    .TextMatrix(i + 1, .ColIndex("End_Interval_Time")) = "" & Format(rsData!End_Interval_Time, "yyyy-MM-dd hh:mm:ss")
                    .TextMatrix(i + 1, .ColIndex("Snap_Id")) = rsData!Snap_Id
                    .TextMatrix(i + 1, .ColIndex("Dbid")) = rsData!DBID
                    .TextMatrix(i + 1, .ColIndex("Instance_Number")) = rsData!Instance_Number
                    .TextMatrix(i + 1, .ColIndex("Startup_Time")) = "" & Format(rsData!Startup_Time, "yyyy-MM-dd hh:mm:ss")
                    i = i + 1
                    rsData.MoveNext
                Loop
                
                .AutoSize 0, .Cols - 1
                .Redraw = flexRDDirect
            End With
        Case tabAddm
            strSql = "Select Owner, Task_Id, Task_Name, Advisor_Name, Created, Description" & vbNewLine & _
                            "From Dba_Advisor_Tasks" & vbNewLine & _
                            "Where Advisor_Name = 'ADDM' And Created Between [1] And [2] " & vbNewLine & _
                            "Order By Task_Id Desc"
            Set rsData = OpenSQLRecord(strSql, "LoadReport", CDate(Format(dtpStart.Value, "yyyy-MM-dd hh:mm:ss")), CDate(Format(dtpEnd.Value, "yyyy-MM-dd hh:mm:ss")))
    
            With vsfAddm
                .Redraw = flexRDNone
                .Rows = .FixedRows
                .Rows = .FixedRows + rsData.RecordCount
                i = 0
                Do While Not rsData.EOF
                    .TextMatrix(i + 1, .ColIndex("Owner")) = rsData!Owner
                    .TextMatrix(i + 1, .ColIndex("Task_Id")) = rsData!Task_Id
                    .TextMatrix(i + 1, .ColIndex("Task_Name")) = rsData!Task_Name
                    .TextMatrix(i + 1, .ColIndex("Advisor_Name")) = rsData!Advisor_Name
                    .TextMatrix(i + 1, .ColIndex("Created")) = "" & Format(rsData!Created, "yyyy-MM-dd hh:mm:ss")
                    .TextMatrix(i + 1, .ColIndex("Description")) = rsData!Description
                    i = i + 1
                    rsData.MoveNext
                Loop
                
                .AutoSize 0, .Cols - 1
                .Redraw = flexRDDirect
            End With
    End Select
    
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Function CheckDate() As Boolean
    '功能： 检查选择的日期是否正确
    Dim blnResult As Boolean
    
    blnResult = True
    If CDate(Format(dtpStart.Value, "yyyy-MM-dd hh:mm:ss")) > CDate(Format(dtpEnd.Value, "yyyy-MM-dd hh:mm:ss")) Then
        MsgBox "起始时间不能晚于终止时间,请重新输入。"
        blnResult = False
    End If
    If CDate(Format(dtpEnd.Value, "yyyy-MM-dd hh:mm:ss")) - CDate(Format(dtpStart.Value, "yyyy-MM-dd hh:mm:ss")) + 1 > 15 Then
        MsgBox "最多查询15天的数据,请重新输入。"
        blnResult = False
    End If

    CheckDate = blnResult
End Function

Private Sub vsfReport_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim intSelectCount As Integer
    Dim intRow1  As Integer, intRow2 As Integer
    
    If vsfReport.Redraw = flexRDNone Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    If InStr(1, mstrAWRSelected, NewRow) > 0 Then Exit Sub
    
    intSelectCount = Len(mstrAWRSelected) - Len(Replace(mstrAWRSelected, ",", ""))
    
    With vsfReport
        .Cell(flexcpForeColor, NewRow, 0, NewRow, .Cols - 1) = 选中字体Color
        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = 选中Color
        cmdGetReport.Enabled = False
        Select Case intSelectCount
            
            Case 0
                mstrAWRSelected = mstrAWRSelected & "_" & NewRow & ","
            Case 1
                cmdGetReport.Enabled = True
                mstrAWRSelected = mstrAWRSelected & "_" & NewRow & ","
            Case 2
                '已有两行被选中，则清除原有行的背景效果
                intRow1 = Replace(Split(mstrAWRSelected, ",")(0), "_", "")
                intRow2 = Replace(Split(mstrAWRSelected, ",")(1), "_", "")
                .Cell(flexcpForeColor, intRow1, 0, intRow1, .Cols - 1) = 未选中字体Color
                .Cell(flexcpBackColor, intRow1, 0, intRow1, .Cols - 1) = 未选中Color
                .Cell(flexcpForeColor, intRow2, 0, intRow2, .Cols - 1) = 未选中字体Color
                .Cell(flexcpBackColor, intRow2, 0, intRow2, .Cols - 1) = 未选中Color

                mstrAWRSelected = "_" & NewRow & ","
        End Select
    End With
End Sub


Private Sub InitCboList()
    Dim strSql As String, rstmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
    strSql = "Select INST_ID, DBID From gv$database "
    Set rstmp = OpenSQLRecord(strSql, "InitCboList")
                            
    cboInst.Visible = rstmp.RecordCount > 1
    lblInst.Visible = rstmp.RecordCount > 1
    
    mdblDBID = rstmp!DBID
    i = 0
    Do While Not rstmp.EOF
        cboInst.AddItem rstmp!Inst_ID, i
        cboInst.ItemData(i) = Val(rstmp!Inst_ID)
        i = i + 1
        rstmp.MoveNext
    Loop
    
    '选中第一个
    cboInst.ListIndex = 0
    Exit Sub
errH:
    MsgBox Err.Description
End Sub
