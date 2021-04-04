VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMediCheckList 
   Caption         =   "药品验收管理"
   ClientHeight    =   8160
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11910
   Icon            =   "frmMediCheckList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   11910
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1605
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   7695
      _cx             =   13573
      _cy             =   2831
      Appearance      =   1
      BorderStyle     =   1
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMediCheckList.frx":030A
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6840
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   28
      Top             =   7440
      Width           =   2175
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   30
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   29
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "不合格"
         Height          =   180
         Left            =   360
         TabIndex        =   32
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "合格"
         Height          =   180
         Left            =   1680
         TabIndex        =   31
         Top             =   30
         Width           =   360
      End
   End
   Begin VB.PictureBox pic水平分割 
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   50
      Left            =   3960
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   7815
      TabIndex        =   20
      Top             =   3000
      Width           =   7815
   End
   Begin VB.Frame fraFilter 
      Height          =   6495
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2415
      Begin VB.CheckBox chkVerify 
         Caption         =   "复核时间"
         Height          =   255
         Left            =   90
         TabIndex        =   27
         Top             =   3110
         Width           =   1455
      End
      Begin VB.CheckBox chkCheck 
         Caption         =   "验收时间"
         Height          =   180
         Left            =   90
         TabIndex        =   26
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "…"
         Height          =   300
         Left            =   1760
         TabIndex        =   24
         Top             =   1200
         Width           =   300
      End
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   90
         TabIndex        =   23
         Top             =   1200
         Width           =   1965
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   480
         Width           =   1965
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "过滤"
         Height          =   300
         Left            =   1185
         TabIndex        =   19
         Top             =   5520
         Width           =   855
      End
      Begin VB.ComboBox cboResult 
         Height          =   300
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   4800
         Width           =   1965
      End
      Begin VB.PictureBox picVerify 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   90
         ScaleHeight     =   735
         ScaleWidth      =   2295
         TabIndex        =   13
         Top             =   3720
         Visible         =   0   'False
         Width           =   2295
         Begin MSComCtl2.DTPicker dtpVerifyBegin 
            Height          =   315
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   157810691
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtpVerifyEnd 
            Height          =   315
            Left            =   0
            TabIndex        =   15
            Top             =   360
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   157810691
            CurrentDate     =   36263
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   2040
            TabIndex        =   16
            Top             =   120
            Width           =   180
         End
      End
      Begin VB.ComboBox cboVerifyDate 
         Enabled         =   0   'False
         Height          =   300
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3360
         Width           =   1965
      End
      Begin VB.PictureBox picCheck 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   90
         ScaleHeight     =   735
         ScaleWidth      =   2295
         TabIndex        =   8
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
         Begin MSComCtl2.DTPicker dtpCheckBegin 
            Height          =   315
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   157810691
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtpCheckEnd 
            Height          =   315
            Left            =   0
            TabIndex        =   10
            Top             =   360
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   157810691
            CurrentDate     =   36263
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   2040
            TabIndex        =   11
            Top             =   120
            Width           =   180
         End
      End
      Begin VB.ComboBox cboCheckDate 
         Height          =   300
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Label LblProvider 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "供药单位"
         Height          =   180
         Left            =   90
         TabIndex        =   25
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblStore 
         AutoSize        =   -1  'True
         Caption         =   "验收库房"
         Height          =   180
         Left            =   90
         TabIndex        =   22
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblResult 
         AutoSize        =   -1  'True
         Caption         =   "验收结果"
         Height          =   180
         Left            =   90
         TabIndex        =   18
         Top             =   4560
         Width           =   720
      End
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   3480
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6855
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   840
      Width           =   50
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1376
      BandCount       =   1
      ImageList       =   "ilsCold"
      _CBWidth        =   11895
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinHeight1      =   720
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         Appearance      =   1
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "BillPreview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "新增"
               Key             =   "add"
               Object.ToolTipText     =   "新增"
               Object.Tag             =   "新增"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "update"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "delete"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "复核"
               Key             =   "Verify"
               Object.ToolTipText     =   "复核"
               Object.Tag             =   "复核"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "view"
               Object.ToolTipText     =   "查阅"
               Object.Tag             =   "查阅"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   9840
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":044B
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":066B
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":088B
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":0AA7
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":0CC7
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":0EE7
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":1103
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":1323
            Key             =   "verify"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":7B85
            Key             =   "view"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   10800
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":7D9F
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":7FBF
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":81DF
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":83FB
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":861B
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":883B
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":8A5B
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":8C7B
            Key             =   "verify"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediCheckList.frx":F4DD
            Key             =   "view"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7800
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediCheckList.frx":F6F7
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15928
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   3495
      Left            =   4080
      TabIndex        =   5
      Top             =   3840
      Width           =   7335
      _cx             =   12938
      _cy             =   6165
      Appearance      =   1
      BorderStyle     =   1
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
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediCheckList.frx":FF8B
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.Menu mnufile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
      End
      Begin VB.Menu mnuFileBillPreview 
         Caption         =   "预览(&V)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditADD 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu mnuEditUpdate 
         Caption         =   "修改(&U)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "复核(&O)"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "查阅(&V)"
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&M)..."
         End
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmMediCheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Public Sub showMe(ByVal fraParent As Form)
    Me.Show , fraParent
End Sub

Private Sub cboCheckDate_Click()
    With cboCheckDate
        If .Text = "自定义" Then
            picCheck.Visible = True
        Else
            picCheck.Visible = False
        End If
    End With
End Sub

Private Sub cboVerifyDate_Click()
    With cboVerifyDate
        If .Text = "自定义" Then
            picVerify.Visible = True
        Else
            picVerify.Visible = False
        End If
    End With
End Sub

Private Sub SetCboDate()
    '往cbo控件中添加数据
    With cboCheckDate
        .AddItem "一周内"
        .AddItem "一月内"
        .AddItem "三月内"
        .AddItem "半年内"
        .AddItem "一年内"
        .AddItem "自定义"
        .ListIndex = 0
    End With
    
    With cboVerifyDate
        .AddItem "一周内"
        .AddItem "一月内"
        .AddItem "三月内"
        .AddItem "半年内"
        .AddItem "一年内"
        .AddItem "自定义"
        .ListIndex = 0
    End With
    
    With cboResult
        .AddItem "合格"
        .AddItem "不合格"
        .AddItem "忽略结果"
        .ListIndex = 2
    End With
    
    dtpCheckEnd = Sys.Currentdate
    dtpVerifyEnd = dtpCheckEnd
    dtpCheckBegin = DateAdd("d", -7, dtpCheckEnd)
    dtpVerifyBegin = dtpCheckBegin
End Sub

Private Sub chkCheck_Click()
    If chkCheck.Value = 1 Then
        cboCheckDate.Enabled = True
    Else
        cboCheckDate.Enabled = False
    End If
End Sub

Private Sub chkVerify_Click()
    If chkVerify.Value = 1 Then
        cboVerifyDate.Enabled = True
    Else
        cboVerifyDate.Enabled = False
    End If
End Sub

Private Sub cmdFilter_Click()
    If chkCheck.Value = 0 And chkVerify.Value = 0 Then
        MsgBox "必须选择一个时间范围进行查询！", vbInformation, gstrSysName
        chkCheck.SetFocus
        Exit Sub
    End If
    
    Call GetList
End Sub

Private Sub cmdProvider_Click()
    Dim rsProvider As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo errHandle
    vRect = GetControlRect(txtProvider.hWnd) '获取位置
    dblLeft = vRect.Left
    dblTop = vRect.Top - 700
    
    gstrSQL = "Select id,上级ID,末级,编码,简码,名称 From 供应商 " & _
              "Where (站点 = [1] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
              "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
              "Start with 上级ID is null connect by prior ID =上级ID order by level,ID"
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "供药单位", True, "", "", False, False, _
                        True, dblLeft, dblTop, 1000, blnCancel, False, True, gstrNodeNo)
    If rsProvider Is Nothing Then
        Exit Sub
    Else
        txtProvider.Text = rsProvider!名称
        txtProvider.Tag = rsProvider!id
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If CheckDepend() = False Then Unload Me: Exit Sub
    Call SetCboDate
    
    Call GetDrugDigit(cboStock.ItemData(cboStock.ListIndex), "药品验收管理", 4, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    Call SetMenu
    
    staThis.Panels(2).Picture = picColor
End Sub

Private Sub SetMenu()
    '根据权限设置菜单和工具栏
    If InStr(1, ";" & gstrprivs & ";", ";新增;") = 0 Then
        mnuEditADD.Visible = False
        tlbTool.Buttons("add").Visible = False
    End If
    If InStr(1, ";" & gstrprivs & ";", ";修改;") = 0 Then
        mnuEditUpdate.Visible = False
        tlbTool.Buttons("update").Visible = False
    End If
    If InStr(1, ";" & gstrprivs & ";", ";删除;") = 0 Then
        mnuEditDelete.Visible = False
        tlbTool.Buttons("delete").Visible = False
    End If
    If InStr(1, ";" & gstrprivs & ";", ";审核;") = 0 Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    If mnuEditADD.Visible = False And mnuEditUpdate.Visible = False And mnuEditDelete.Visible = False Then
        tlbTool.Buttons(3).Visible = False
    End If
    If mnuEditVerify.Visible = False Then
        tlbTool.Buttons(9).Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With cbrTool
        .Width = Me.Width
    End With
    
    With fraFilter
        .Move 50, IIf(cbrTool.Bands(1).Visible = True, cbrTool.Top + cbrTool.Height - 30, 30), 2415, Me.Height - cbrTool.Top - cbrTool.Bands(1).Height - staThis.Height - 850
    End With
    
    With picFilter
        .Move fraFilter.Left + fraFilter.Width, fraFilter.Top, picFilter.Width, fraFilter.Height
    End With
    
    With vsfList
        .Move picFilter.Left + picFilter.Width, fraFilter.Top + 85, Me.Width - picFilter.ScaleWidth - picFilter.Left - 280, picFilter.Height / 3
    End With
    
    With pic水平分割
        .Move vsfList.Left, vsfList.Top + vsfList.Height, vsfList.Width
    End With
    
    With vsfDetail
        .Move picFilter.Left + picFilter.Width, vsfList.Top + vsfList.Height, Me.Width - picFilter.ScaleWidth - picFilter.Left - 280, (picFilter.Height / 3) * 2 - 85
    End With
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - .Width - 300
    End With
End Sub

Private Sub SetColor()
    '设置单据颜色
    Dim lngRow As Long
    
    With vsfList
        If .rows = 1 Then Exit Sub
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("验收结果")) = "不合格" Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = picColor1.BackColor
            End If
        Next
    End With
End Sub

Private Sub mnuEditADD_Click()
    If cboStock.ListIndex = -1 Then
        MsgBox "必须选择一个具有药库或者药房性质的部门！", vbInformation, gstrSysName
        Exit Sub
    End If
    frmMediCheckCard.showMe 1, Me, cboStock.ItemData(cboStock.ListIndex), 0
End Sub


Private Sub mnuEditDelete_Click()
    Dim lng验收id As Long
    
    With vsfList
        If .Row > 0 Then
            If .TextMatrix(.Row, .ColIndex("验收id")) <> "" Then
                lng验收id = Val(.TextMatrix(.Row, .ColIndex("验收id")))
            Else
                lng验收id = 0
            End If
        Else
            lng验收id = 0
        End If
    End With
    
    If lng验收id <> 0 Then
        Call DeleteDate(lng验收id)
        Call GetList
    End If
End Sub

Private Sub mnuEditUpdate_Click()
    Dim lng验收id As Long
    
    With vsfList
        If .Row > 0 Then
            If .TextMatrix(.Row, .ColIndex("验收id")) <> "" Then
                lng验收id = Val(.TextMatrix(.Row, .ColIndex("验收id")))
            Else
                lng验收id = 0
            End If
        Else
            lng验收id = 0
        End If
    End With
    
    If lng验收id <> 0 Then
        frmMediCheckCard.showMe 2, Me, cboStock.ItemData(cboStock.ListIndex), lng验收id
        Call GetList
    End If
End Sub

Private Sub mnuEditVerify_Click()
    Dim lng验收id As Long
    
    With vsfList
        If .Row > 0 Then
            If .TextMatrix(.Row, .ColIndex("验收id")) <> "" Then
                lng验收id = Val(.TextMatrix(.Row, .ColIndex("验收id")))
            Else
                lng验收id = 0
            End If
        Else
            lng验收id = 0
        End If
    End With

    If lng验收id <> 0 Then
        frmMediCheckCard.showMe 3, Me, cboStock.ItemData(cboStock.ListIndex), lng验收id
        Call GetList
    End If
End Sub

Private Sub mnuEditView_Click()
    Dim lng验收id As Long
    
    With vsfList
        If .Row > 0 Then
            If .TextMatrix(.Row, .ColIndex("验收id")) <> "" Then
                lng验收id = Val(.TextMatrix(.Row, .ColIndex("验收id")))
            Else
                lng验收id = 0
            End If
        Else
            lng验收id = 0
        End If
    End With
    
    If lng验收id <> 0 Then
        frmMediCheckCard.showMe 4, Me, cboStock.ItemData(cboStock.ListIndex), lng验收id
    End If
End Sub

Private Sub mnuFileBillPreview_Click()
    Call mnuFilePreView_Click
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, "药品验收管理")
End Sub

Private Sub mnuHelpWebForum_Click()
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '工具条索引
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            '取消所有的文本标签显示
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '让所有的文本标签显示。说明：Tag中放的文本标签
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
        
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub

Private Sub picFilter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    With picFilter
        If .Left + x <= 1000 Then Exit Sub
        If .Left + x >= 5000 Then Exit Sub
        .Move .Left + x, .Top, .Width, .Height
    End With
    
    With fraFilter
        .Move .Left, .Top, .Width + x
    End With
    
    With vsfList
        .Move .Left + x, .Top, .Width - x
    End With
    
    With pic水平分割
        .Left = vsfList.Left
        .Width = vsfList.Width
    End With
    
    With vsfDetail
        .Left = vsfList.Left
        .Width = vsfList.Width
    End With
End Sub

Private Sub pic水平分割_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With pic水平分割
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    With vsfList
        .Height = pic水平分割.Top - .Top
    End With
    
    With vsfDetail
        .Top = pic水平分割.Top + pic水平分割.Height + 100
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
    End With
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add" '新增
            Call mnuEditADD_Click
            Call GetList
        Case "update" '修改
            Call mnuEditUpdate_Click
            Call GetList
        Case "delete" '删除
            Call mnuEditDelete_Click
        Case "Verify" '复核
            Call mnuEditVerify_Click
            Call GetList
        Case "view" '查阅
            Call mnuEditView_Click
        Case "help" '帮助
            Call mnuHelpTitle_Click
        Case "exit" '退出
            Unload Me
        Case "BillPreview" '预览
            Call mnuFilePreView_Click
        Case "print" '打印
            Call mnuFilePrint_Click
    End Select
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    Dim lngCurRow As Long
    
    If ActiveControl Is vsfList Then
        lngCurRow = vsfList.Row
        vsfList.Redraw = flexRDNone
        subPrint 1
        vsfList.Row = lngCurRow
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    Else
        lngCurRow = vsfDetail.Row
        vsfDetail.Redraw = flexRDNone
        subPrint 1
        vsfDetail.Row = lngCurRow
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    Dim lngCurRow As Long
    
    If ActiveControl Is vsfList Then
        lngCurRow = vsfList.Row
        vsfList.Redraw = flexRDNone
        subPrint 2
        vsfList.Row = lngCurRow
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    Else
        lngCurRow = vsfDetail.Row
        vsfDetail.Redraw = flexRDNone
        subPrint 2
        vsfDetail.Row = lngCurRow
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "药品验收管理"
        
'    objRow.Add "时间：" & strRange
    objRow.Add "部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.用户姓名
    objRow.Add "打印日期:" & Format(zlDataBase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    If ActiveControl Is vsfList Then
        Set objPrint.Body = vsfList
    Else
        Set objPrint.Body = vsfDetail
    End If
    
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

Private Sub DeleteDate(ByVal lng验收id As Long)
    '删除单据
    If MsgBox("将删除当前选中单据，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        gstrSQL = "Zl_药品验收记录_Delete(" & lng验收id & ")"
        Call zlDataBase.ExecuteProcedure(gstrSQL, "删除验收入库单")
        MsgBox "删除单据成功！", vbInformation, gstrSysName
        Call GetList
    End If
End Sub

Private Function CheckDepend() As Boolean
    Dim rsDepend As New ADODB.Recordset
    Dim strStock As String, strCaption As String
    
    On Error GoTo errHandle
    
    '获取可操作的库房
    strStock = "HIJKLMN"
    
    '如果是药品领用，则检查当前科室是否是领用部门，且允许向库房领药
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = [3] Or a.站点 is Null) And c.工作性质 = b.名称 " _
            & "  AND Instr([2],b.编码,1) > 0 " _
            & "  AND a.id = c.部门id " _
            & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
            & " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])"
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "药品入库验收", UserInfo.用户ID, strStock, gstrNodeNo)
        
    If rsDepend.EOF Then
        MsgBox "你的角色不是库房人员，不能使用药品验收模块！" & vbCrLf & "请在人员管理中至少添加一个具有药库性质，或药房、制剂室性质的部门。", vbInformation, gstrSysName
        If rsDepend.State = 1 Then rsDepend.Close
        Exit Function
    End If
    
    '装入库房数据
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!名称
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = UserInfo.部门ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
        rsDepend.Close
    End With
    
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtProvider_Change()
    If txtProvider.Text = "" Then
        txtProvider.Tag = 0
    End If
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    Dim strProviderText As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    vRect = GetControlRect(txtProvider.hWnd) '获取位置
    dblLeft = vRect.Left
    dblTop = vRect.Top
    
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        gstrSQL = "Select id,编码,名称,简码 From 供应商 " & _
                  "Where (站点 = [2] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                  "  And 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                  "  And (简码 like [1] Or 编码 like [1] or 名称 like [1] )"
'        Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "药品入库验收管理", _
'                            IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
'
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "供药单位", False, "", "", False, False, _
                        True, dblLeft, dblTop, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        If blnCancel Then txtProvider.SetFocus: Exit Sub
        
        If rsProvider Is Nothing Then
            MsgBox "没有您输入的供药单位，请重输！", vbOKOnly + vbInformation, gstrSysName
            txtProvider.SelStart = 0
            txtProvider.SelLength = Len(txtProvider)
            Exit Sub
        Else
            txtProvider.Text = rsProvider!名称
            txtProvider.Tag = rsProvider!id
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetList()
    '加载汇总数据
    Dim rsTemp As ADODB.Recordset
    Dim datTemp As Date
    Dim strCheckBegin As String
    Dim strCheckEnd As String
    Dim strVerifyBegin As String
    Dim strVerifyEnd As String
    Dim datCheckBegin As Date
    Dim datCheckEnd As Date
    Dim datVerifyBegin As Date
    Dim datVerifyEnd As Date
    Dim lng库房ID As Long
    Dim lng供药单位ID As Long
    Dim int合格 As Integer
    
    On Error GoTo errHandle
    
    vsfList.rows = 1
    '库房id
    gstrSQL = " and a.库房id=[1]"
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    
    '供药单位id
    If Val(txtProvider.Tag) <> 0 Then
        gstrSQL = gstrSQL & " and a.供药单位id=[2]"
    End If
    lng供药单位ID = Val(txtProvider.Tag)
    
    datTemp = zlDataBase.Currentdate
    '验收日期
    If chkCheck.Value = 1 Then
        Select Case cboCheckDate.Text
        Case "一周内"
            strCheckBegin = Format(DateAdd("D", -7, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "一月内"
            strCheckBegin = Format(DateAdd("M", -1, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "三月内"
            strCheckBegin = Format(DateAdd("M", -3, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "半年内"
            strCheckBegin = Format(DateAdd("M", -6, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "一年内"
            strCheckBegin = Format(DateAdd("yyyy", -1, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "自定义"
            strCheckBegin = Format(dtpCheckBegin.Value, "yyyy-mm-dd")
            strCheckEnd = Format(dtpCheckEnd.Value, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        End Select
        
        gstrSQL = gstrSQL & " and a.验收日期 between [3] and [4]"
    End If
    
    '复核时间
    If chkVerify.Value = 1 Then
        Select Case cboVerifyDate.Text
        Case "一周内"
            strVerifyBegin = Format(DateAdd("D", -7, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "一月内"
            strVerifyBegin = Format(DateAdd("M", -1, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "三月内"
            strVerifyBegin = Format(DateAdd("M", -3, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "半年内"
            strVerifyBegin = Format(DateAdd("M", -6, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "一年内"
            strVerifyBegin = Format(DateAdd("yyyy", -1, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "自定义"
            strVerifyBegin = Format(dtpVerifyBegin.Value, "yyyy-mm-dd")
            strVerifyEnd = Format(dtpVerifyEnd.Value, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        End Select
        
        gstrSQL = gstrSQL & " and a.复核日期 between [5] and [6]"
    End If
    
    '验收结果
    Select Case cboResult.Text
        Case "合格"
            int合格 = 0
            gstrSQL = gstrSQL & " and a.是否合格 = [7]"
        Case "不合格"
            int合格 = 1
            gstrSQL = gstrSQL & " and a.是否合格 = [7]"
    End Select
        
    gstrSQL = "Select distinct a.Id, a.No, a.库房id, a.供药单位id, a.验收人, a.验收日期, a.复核人, a.复核日期, a.是否合格, c.名称 as 供药单位, a.备注" & vbNewLine & _
                "From 药品验收记录 A, 药品验收明细 B, 供应商 C" & vbNewLine & _
                "Where a.Id = b.验收id And a.供药单位id = c.Id(+)" & gstrSQL
                
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "药品验收入库", _
    lng库房ID, _
    lng供药单位ID, _
    datCheckBegin, _
    datCheckEnd, _
    datVerifyBegin, _
    datVerifyEnd, _
    int合格)
    
    With vsfList
        Do While Not rsTemp.EOF
            .rows = .rows + 1
            .TextMatrix(.rows - 1, .ColIndex("验收id")) = rsTemp!id
            .TextMatrix(.rows - 1, .ColIndex("no")) = rsTemp!NO
            .TextMatrix(.rows - 1, .ColIndex("验收结果")) = IIf(rsTemp!是否合格 = 0, "合格", "不合格")
            .TextMatrix(.rows - 1, .ColIndex("验收人")) = IIf(IsNull(rsTemp!验收人), "", rsTemp!验收人)
            .TextMatrix(.rows - 1, .ColIndex("验收日期")) = IIf(IsNull(rsTemp!验收日期), "", Format(rsTemp!验收日期, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.rows - 1, .ColIndex("复核人")) = IIf(IsNull(rsTemp!复核人), "", rsTemp!复核人)
            .TextMatrix(.rows - 1, .ColIndex("复核日期")) = IIf(IsNull(rsTemp!复核日期), "", Format(rsTemp!复核日期, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.rows - 1, .ColIndex("供药单位")) = IIf(IsNull(rsTemp!供药单位), "", rsTemp!供药单位)
            .TextMatrix(.rows - 1, .ColIndex("备注")) = IIf(IsNull(rsTemp!备注), "", rsTemp!备注)
            
            rsTemp.MoveNext
        Loop
        
        Call SetColor
        If .rows > 1 Then
            .Row = 1
            .SetFocus
            
            Call vsfList_EnterCell
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfList
        If Val(.TextMatrix(NewRow, .ColIndex("验收id"))) = 0 Then
            tlbTool.Buttons("update").Enabled = False
            tlbTool.Buttons("delete").Enabled = False
            tlbTool.Buttons("Verify").Enabled = False
        Else
            If .TextMatrix(NewRow, .ColIndex("复核人")) <> "" Then
                tlbTool.Buttons("delete").Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
                tlbTool.Buttons("update").Enabled = False
            Else
                tlbTool.Buttons("delete").Enabled = True
                tlbTool.Buttons("Verify").Enabled = True
                tlbTool.Buttons("update").Enabled = True
            End If
        End If
        mnuEditUpdate.Enabled = tlbTool.Buttons("update").Enabled
        mnuEditDelete.Enabled = tlbTool.Buttons("delete").Enabled
        mnuEditVerify.Enabled = tlbTool.Buttons("Verify").Enabled
    End With
End Sub

Private Sub vsfList_EnterCell()
    Dim lng验收id As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    With vsfDetail
        .rows = 1
        
        If vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("验收id")) <> "" Then
            lng验收id = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("验收id")))
            
            gstrSQL = "Select b.编码, b.名称, b.规格, c.药库单位, c.药库包装, a.进药日期, e.名称 As 剂型, a.成本价, a.零售价, a.进药数量, a.批号, a.生产日期, a.效期, a.产地, a.批准文号," & vbNewLine & _
                        "       nvl(a.是否合格,0) as 是否合格" & vbNewLine & _
                        "From 药品验收明细 A, 收费项目目录 B, 药品规格 C, 药品特性 D, 药品剂型 E" & vbNewLine & _
                        "Where a.药品id = b.Id And b.Id = c.药品id And c.药名id = d.药名id And d.药品剂型 = e.名称(+) and a.验收id=[1]"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "验收明细查询", lng验收id)
            
            Do While Not rsTemp.EOF
                .rows = .rows + 1
                
                .TextMatrix(.rows - 1, .ColIndex("验收结果")) = IIf(rsTemp!是否合格 = 0, "合格", "不合格")
                .TextMatrix(.rows - 1, .ColIndex("药品名称")) = "[" & rsTemp!编码 & "]" & rsTemp!名称
                .TextMatrix(.rows - 1, .ColIndex("规格")) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                .TextMatrix(.rows - 1, .ColIndex("单位")) = IIf(IsNull(rsTemp!药库单位), "", rsTemp!药库单位)
                .TextMatrix(.rows - 1, .ColIndex("进药日期")) = IIf(IsNull(rsTemp!进药日期), "", Format(rsTemp!进药日期, "yyyy-mm-dd hh:mm:ss"))
                .TextMatrix(.rows - 1, .ColIndex("剂型")) = IIf(IsNull(rsTemp!剂型), "", rsTemp!剂型)
                .TextMatrix(.rows - 1, .ColIndex("成本价")) = IIf(IsNull(rsTemp!成本价), "", zlStr.FormatEx(rsTemp!成本价, mintCostDigit, True, True))
                .TextMatrix(.rows - 1, .ColIndex("零售价")) = IIf(IsNull(rsTemp!零售价), "", zlStr.FormatEx(rsTemp!零售价, mintPriceDigit, True, True))
                .TextMatrix(.rows - 1, .ColIndex("进药数量")) = IIf(IsNull(rsTemp!进药数量), "", zlStr.FormatEx(rsTemp!进药数量, mintNumberDigit, True, True))
                .TextMatrix(.rows - 1, .ColIndex("药品批号")) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                .TextMatrix(.rows - 1, .ColIndex("生产日期")) = IIf(IsNull(rsTemp!生产日期), "", Format(rsTemp!生产日期, "yyyy-mm-dd"))
                .TextMatrix(.rows - 1, .ColIndex("效期")) = IIf(IsNull(rsTemp!效期), "", Format(rsTemp!效期, "yyyy-mm-dd"))
                .TextMatrix(.rows - 1, .ColIndex("产地")) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                .TextMatrix(.rows - 1, .ColIndex("批准文号")) = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
                
                rsTemp.MoveNext
            Loop
        End If
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


