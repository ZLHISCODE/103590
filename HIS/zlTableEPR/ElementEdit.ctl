VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl ElementEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   11940
   ToolboxBitmap   =   "ElementEdit.ctx":0000
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   9570
      ScaleHeight     =   3690
      ScaleWidth      =   2235
      TabIndex        =   15
      Top             =   15
      Visible         =   0   'False
      Width           =   2235
      Begin MSMask.MaskEdBox txtTime 
         Height          =   330
         Left            =   0
         TabIndex        =   24
         Top             =   1815
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.ListBox lstTime 
         Appearance      =   0  'Flat
         Height          =   1290
         Left            =   0
         TabIndex        =   17
         Top             =   2400
         Width           =   2235
      End
      Begin VB.PictureBox picClock 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1770
         Left            =   0
         ScaleHeight     =   1740
         ScaleWidth      =   2205
         TabIndex        =   16
         Top             =   0
         Width           =   2235
         Begin VB.Line linHand 
            BorderWidth     =   2
            X1              =   240
            X2              =   960
            Y1              =   735
            Y2              =   585
         End
         Begin VB.Shape shpCenter 
            Height          =   90
            Left            =   675
            Shape           =   3  'Circle
            Top             =   720
            Width           =   90
         End
         Begin VB.Shape shpDot 
            FillColor       =   &H00FFFFFF&
            Height          =   105
            Index           =   0
            Left            =   690
            Shape           =   3  'Circle
            Top             =   180
            Width           =   135
         End
      End
      Begin VB.Label lblTimeType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "格式类型"
         Height          =   180
         Left            =   0
         TabIndex        =   19
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label lblAmOrPm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上午"
         Height          =   180
         Left            =   1725
         TabIndex        =   18
         Top             =   1890
         Width           =   360
      End
   End
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   5460
      ScaleHeight     =   3690
      ScaleWidth      =   4035
      TabIndex        =   11
      Top             =   15
      Visible         =   0   'False
      Width           =   4035
      Begin VB.ListBox lstDate 
         Appearance      =   0  'Flat
         Height          =   1290
         Left            =   0
         TabIndex        =   12
         Top             =   2400
         Width           =   4035
      End
      Begin MSComCtl2.MonthView mvwDate 
         Height          =   2160
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3810
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   59834369
         TitleForeColor  =   16711680
         CurrentDate     =   38549
         MaxDate         =   401769
         MinDate         =   367
      End
      Begin VB.Label lblDateType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "格式类型"
         Height          =   180
         Left            =   0
         TabIndex        =   14
         Top             =   2190
         Width           =   720
      End
   End
   Begin VB.TextBox txt文本 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   210
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox txt上下1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   -15
      TabIndex        =   5
      Text            =   "99999"
      Top             =   225
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F5BE9E&
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   30
      MousePointer    =   5  'Size
      ScaleHeight     =   105
      ScaleWidth      =   5325
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   5325
      Begin VB.Image imgTitle 
         Height          =   45
         Left            =   1350
         MousePointer    =   5  'Size
         Picture         =   "ElementEdit.ctx":0312
         Top             =   30
         Width           =   2250
      End
   End
   Begin VB.TextBox txt上下2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Text            =   "99999"
      Top             =   225
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   5415
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3360
      Width           =   5415
      Begin VB.Image imgResize 
         Height          =   270
         Left            =   5175
         MousePointer    =   8  'Size NW SE
         Picture         =   "ElementEdit.ctx":0394
         Top             =   0
         Width           =   225
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "ESC 取消退出；回车:保存修改。"
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   60
         Width           =   3810
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg复选 
      Height          =   555
      Left            =   2865
      TabIndex        =   4
      Top             =   225
      Visible         =   0   'False
      Width           =   780
      _cx             =   1376
      _cy             =   979
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ElementEdit.ctx":0736
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComCtl2.UpDown ud上下 
      Height          =   300
      Left            =   645
      TabIndex        =   6
      Top             =   165
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      OrigLeft        =   1395
      OrigTop         =   2220
      OrigRight       =   1635
      OrigBottom      =   2520
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg单选 
      Height          =   570
      Left            =   3705
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   690
      _cx             =   1217
      _cy             =   1005
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16777215
      ForeColorFixed  =   0
      BackColorSel    =   16761024
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ElementEdit.ctx":0773
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   1395
      Left            =   60
      TabIndex        =   21
      Top             =   1350
      Width           =   2595
      _cx             =   4577
      _cy             =   2461
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ElementEdit.ctx":07B0
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
   Begin MSComctlLib.Toolbar cbrThis 
      Height          =   330
      Left            =   1905
      TabIndex        =   22
      Top             =   945
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ButtonWidth     =   1349
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ils16"
      HotImageList    =   "ils16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "搜索"
            Key             =   "搜索"
            Object.ToolTipText     =   "搜索"
            Object.Tag             =   "搜索"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "选中"
            Key             =   "选中"
            Object.ToolTipText     =   "选中"
            Object.Tag             =   "选中"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "取消"
            Object.ToolTipText     =   "取消"
            Object.Tag             =   "取消"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   4215
      Top             =   1245
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
            Picture         =   "ElementEdit.ctx":0885
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ElementEdit.ctx":70E7
            Key             =   "find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ElementEdit.ctx":7481
            Key             =   "cancel"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFind 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   105
      TabIndex        =   20
      Top             =   975
      Width           =   1755
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   220
      Left            =   3240
      TabIndex        =   23
      Top             =   2655
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Shape shpBorderOut 
      BorderColor     =   &H00E09060&
      Height          =   255
      Left            =   2490
      Top             =   210
      Width           =   270
   End
   Begin VB.Label lblDot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   300
      Left            =   675
      TabIndex        =   10
      Top             =   202
      Width           =   105
   End
   Begin VB.Shape shpBorder1 
      BorderColor     =   &H00E09060&
      Height          =   255
      Left            =   1785
      Top             =   210
      Width           =   270
   End
   Begin VB.Shape shpBorder2 
      BorderColor     =   &H00E09060&
      Height          =   255
      Left            =   2160
      Top             =   210
      Width           =   270
   End
   Begin VB.Label lbl单位 
      BackStyle       =   0  'Transparent
      Caption         =   "单位"
      Height          =   210
      Left            =   1320
      TabIndex        =   9
      Top             =   247
      Width           =   480
   End
   Begin VB.Image imgOpt2 
      Height          =   195
      Left            =   5205
      Picture         =   "ElementEdit.ctx":DCE3
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgOpt1 
      Height          =   195
      Left            =   5205
      Picture         =   "ElementEdit.ctx":DF69
      Top             =   210
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "ElementEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const conPI As Double = 3.14159265358979
Public Event pOk()              '保存数据
Public Event pChange()          '复选框数据改变
Public Event pCancel()          '取消修改
Public Event TitleMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event TitleMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event TitleMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public MoveTag As String, MoveOldX As Single, MoveOldY As Single
Public Element As cTabElement
Private mPressKeyAscii As Integer
Private lngX As Long, lngY As Long
Private mblnHandMove As Boolean, mblnHandTime As Boolean, mbEt As Byte
Public Property Get hWnd()
    hWnd = UserControl.hWnd
End Property
'################################################################################################################
'## 功能：  显示诊治要素编辑器
'##
'## 参数：  Ele         :所编辑的诊治要素
'################################################################################################################
Public Sub SetElement(ByRef Ele As cTabElement, ByVal KeyAscii As Integer, Optional ByVal bEditType As Byte = 0)
'功能：显示控件以前赋值要素，该控件会改变要素,并将结果以要素.内空文本方式返回
Dim i As Long, j As Long, T As Variant, strTmp As String, dtInit As Date
    mPressKeyAscii = KeyAscii: MoveTag = "": MoveOldX = 0: MoveOldY = 0: mbEt = bEditType
    Set Element = Ele
    With Element
        If .替换域 = 2 Then         '字典类
            txtFind.Text = Chr(KeyAscii)
            vgdList.Clear: vgdList.Rows = 2
            If txtFind.Text <> "" Then DoFind
        Else
            Select Case .要素表示       '0-文本,1-上下,2-单选,3-复选
            Case 0
                Select Case .要素类型
                    Case 2                      '日期型
                        Dim strMinTime As String, strMaxTime As String, strMinDate As String, strMaxDate As String
                        T = Split(.要素值域, ";")
                        If T(0) = 0 Then
                            strMinDate = Format("1901-01-01", "yyyy-MM-dd"): strMinTime = Format("00:00:00", "HH:mm:ss")
                        Else
                            strMinDate = Format(T(0), "yyyy-MM-dd"): strMinTime = Format(T(0), "HH:mm:ss")
                        End If
                        
                        If T(1) = 0 Then
                            strMaxDate = Format("3000-01-01", "yyyy-MM-dd"): strMaxTime = Format("23:59:59", "HH:mm:ss")
                        Else
                            strMaxDate = Format(T(1), "yyyy-MM-dd"): strMaxTime = Format(T(1), "HH:mm:ss")
                        End If
                        
                        If .输入形态 = 0 Then '弹出形式
                            On Error Resume Next
                            mvwDate.MinDate = Format("1901-01-01", "yyyy-MM-dd"): mvwDate.MaxDate = Format("3000-01-01", "yyyy-MM-dd")
                            Err.Clear
                            mvwDate.MinDate = Format(strMinDate, "yyyy-MM-dd"): mvwDate.MaxDate = Format(strMaxDate, "yyyy-MM-dd")
                            txtTime.Tag = Format(strMinTime, "HH:mm:ss") & "|" & Format(strMaxTime, "HH:mm:ss")
'                            dtTime.MinDate = Format(strMinTime, "HH:mm:ss"): dtTime.MaxDate = Format(strMaxTime, "HH:mm:ss")
                            lstDate.Tag = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\ElementEdit\" & .要素类型 & .要素表示 & .替换域, "\DateType", 0)
                            lstTime.Tag = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\ElementEdit\" & .要素类型 & .要素表示 & .替换域, "\TimeType", 0)
                            If .要素长度 >= 10 Then '日期和长日期
                                If .内容文本 = "" Then
                                    dtInit = Now
                                Else
                                    dtInit = Format(.内容文本, "yyyy-MM-dd")
                                End If
                                If dtInit < mvwDate.MinDate Then dtInit = mvwDate.MinDate
                                If dtInit > mvwDate.MaxDate Then dtInit = mvwDate.MaxDate
                                mvwDate.Value = dtInit
                                Call mvwDate_SelChange(mvwDate.MinDate, mvwDate.MaxDate, False)
                            End If
                            
                            Call DrawWatch
                            If .要素长度 <> 10 Then '时间和长日期
                                If .内容文本 <> "" And Format(.内容文本, "hh:mm:ss") <> CDate("00:00:00") Then
                                    txtTime.Text = Format(.内容文本, "HH:mm:ss")
                                Else
                                    txtTime.Text = "__:__:__"
                                End If
                            End If
                            Err.Clear
                        Else            '展开型
                            txtDate.Tag = strMinDate & "|" & strMaxDate & "|" & strMinTime & "|" & strMaxTime
                            Select Case .要素长度
                                Case 8
                                    txtDate.Format = "HH:mm:ss"
                                    txtDate.Mask = "##:##:##"
                                    txtDate.Text = Format(IIf(Trim(.内容文本) = "", Now, Trim(.内容文本)), "HH:mm:ss")
                                    If CDate(txtDate) < CDate(strMinTime) Then txtDate.Text = strMinTime
                                    If CDate(txtDate) > CDate(strMaxTime) Then txtDate.Text = strMaxTime
                                Case 10
                                    txtDate.Format = "yyyy-MM-dd"
                                    txtDate.Mask = "####-##-##"
                                    txtDate.Text = Format(IIf(Trim(.内容文本) = "", Now, Trim(.内容文本)), "yyyy-MM-dd")
                                    If CDate(txtDate) < CDate(strMinDate) Then txtDate.Text = strMinDate
                                    If CDate(txtDate) > CDate(strMaxDate) Then txtDate.Text = strMaxDate
                                Case 19
                                    txtDate.Format = "yyyy-MM-dd HH:mm:ss"
                                    txtDate.Mask = "####-##-## ##:##:##"
                                    txtDate.Text = Format(IIf(Trim(.内容文本) = "", Now, Trim(.内容文本)), "yyyy-MM-dd HH:mm:ss")
                                    If CDate(txtDate) < CDate(strMinDate & " " & strMinTime) Then txtDate.Text = strMinDate & " " & strMinTime
                                    If CDate(txtDate) > CDate(strMaxDate & " " & strMaxTime) Then txtDate.Text = strMaxDate & " " & strMaxTime
                            End Select
                            txtDate.MaxLength = .要素长度
                        End If
                    Case 3                      '逻辑型
                        strTmp = .内容文本:   T = Split(.要素值域, ";"):      strTmp = IIf(strTmp = "", T(1), strTmp):      .内容文本 = strTmp
                        vfg单选.RowHeightMax = 240:     vfg单选.Cols = 2:       vfg单选.ColWidth(0) = 250
                        vfg单选.ColWidth(1) = IIf(ScaleWidth > 250, ScaleWidth - 250, 250):             vfg单选.Rows = UBound(T) + 1
                        For i = 0 To UBound(T)
                            vfg单选.Cell(flexcpText, i, 1) = T(i)
                            vfg单选.Cell(flexcpPicture, i, 0) = IIf(T(i) = strTmp, imgOpt2.Picture, imgOpt1.Picture)
                        Next i
                    Case Else
                        txt文本.MaxLength = .要素长度:  txt文本 = .内容文本 & Chr(mPressKeyAscii)
                        txt文本.SelStart = 0: txt文本.SelLength = Len(.内容文本): txt文本.Visible = True
                End Select
            Case 1
                T = Split(.要素值域, ";")    '格式:  0;100000
                If UBound(T) < 1 Then
                    ud上下.Min = 0:             ud上下.Max = 999999999
                Else
                    ud上下.Min = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(1)), CLng(T(0)))
                    ud上下.Max = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(0)), CLng(T(1)))
                End If
                txt上下1.Tag = "赋值...":       i = InStr(1, .内容文本, ".")
                If i > 0 Then
                    txt上下1 = Mid(.内容文本, 1, i - 1):    txt上下1.Visible = True
                    txt上下1.SelStart = 0: txt上下1.SelLength = Len(txt上下1)
                    txt上下2 = Mid(.内容文本, i + 1)
                Else
                    txt上下1 = .内容文本:                   txt上下2 = ""
                End If
                txt上下1.Tag = "":                          txt上下1.MaxLength = .要素长度
                lbl单位 = .要素单位
                If Trim(.要素单位) <> "" Then
                    lbl单位.Visible = True
                Else
                    lbl单位.Visible = False
                End If
                If .要素小数 > 0 Then
                    txt上下2.MaxLength = .要素小数:         txt上下2.Visible = True:        lblDot.Visible = True
                Else
                    txt上下2.Visible = False:   lblDot.Visible = False
                End If
            Case 2
                vfg单选.Clear:      vfg单选.FocusRect = flexFocusNone:      vfg单选.Editable = flexEDKbdMouse:      T = Split(.要素值域, ";")
                If .输入形态 = 0 Then
                    strTmp = "、" & .内容文本 & "、"
                Else '展开形式   '○●
                    For i = 1 To UBound(Split(.内容文本, "●"))
                        strTmp = strTmp & "、" & Split(Split(.内容文本, "●")(i), "○")(0)
                    Next
                    strTmp = strTmp & "、"
                End If
                If .输入形态 = 0 Then
                    vfg单选.RowHeightMax = 240:     vfg单选.Cols = 2:       vfg单选.ColWidth(0) = 250
                    vfg单选.ColWidth(1) = IIf(ScaleWidth > 250, ScaleWidth - 250, 250):             vfg单选.Rows = UBound(T) + 1
                    For i = 0 To UBound(T)
                        vfg单选.Cell(flexcpText, i, 1) = T(i)
                        vfg单选.Cell(flexcpPicture, i, 0) = IIf(InStr(strTmp, "、" & T(i) & "、") > 0, imgOpt2.Picture, imgOpt1.Picture)
                    Next i
                Else
                    vfg单选.RowHeightMax = 0:       vfg单选.Rows = 1:               vfg单选.Cols = (UBound(T) + 1) * 2
                    For i = 0 To UBound(T) * 2 Step 2 '每间隔一列为图标
                        vfg单选.Cell(flexcpText, 0, i + 1) = T(i / 2)
                        vfg单选.Cell(flexcpPicture, 0, i) = IIf(InStr(strTmp, "、" & T(i / 2) & "、") > 0, imgOpt2.Picture, imgOpt1.Picture)
                    Next
                End If
            Case 3
                vfg复选.Clear:      vfg复选.Editable = flexEDKbdMouse:      T = Split(.要素值域, ";")
                If .输入形态 = 0 Then
                    strTmp = "、" & .内容文本 & "、"
                Else '展开形式
                    For i = 1 To UBound(Split(.内容文本, "■"))
                        strTmp = strTmp & "、" & Split(Split(.内容文本, "■")(i), "□")(0)
                    Next
                    strTmp = strTmp & "、"
                End If
                If .输入形态 = 0 Then
                    vfg复选.RowHeightMax = 240:         vfg复选.Cols = 1:       vfg复选.Rows = UBound(T) + 1:       vfg复选.ColWidth(0) = 240
                    For i = 0 To UBound(T)
                        vfg复选.Cell(flexcpText, i, 0) = T(i)
                        vfg复选.Cell(flexcpChecked, i, 0) = IIf(InStr(1, strTmp, "、" & vfg复选.Cell(flexcpText, i, 0) & "、") > 0, flexChecked, flexUnchecked)
                    Next
                Else
                    vfg复选.RowHeightMax = 0:           vfg复选.Rows = 1:                   vfg复选.Cols = UBound(T) + 1
                    For i = 0 To UBound(T)
                        vfg复选.Cell(flexcpText, 0, i) = T(i)
                        vfg复选.Cell(flexcpChecked, 0, i) = IIf(InStr(1, strTmp, "、" & vfg复选.Cell(flexcpText, 0, i) & "、") > 0, flexChecked, flexUnchecked)
                    Next
                End If
            End Select
        End If

        If .输入形态 = 0 Then
            Width = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\ElementEdit\" & .要素类型 & .要素表示 & .替换域, "\MainWidth", 2500)
            Height = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\ElementEdit\" & .要素类型 & .要素表示 & .替换域, "\MainHeight", 3870)
        End If
    End With
    
    Call UserControl_Resize
End Sub

Private Sub cbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call DoFind
    Case 2
        vgdList_DblClick
    Case 3
        RaiseEvent pOk
    End Select
End Sub

Private Sub dtTime_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub imgResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgResize.Tag = "Down"
    lngX = x: lngY = y
End Sub

Private Sub imgResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If imgResize.Tag = "Down" Then
        If Width + x - lngX >= 1000 And Width + x - lngX <= 12000 Then
            Width = Width + x - lngX
        End If
        If Height + y - lngY >= 1000 And Height + y - lngY <= 9000 Then
            Height = Height + y - lngY
        End If
        DoEvents
    End If
End Sub

Private Sub imgResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgResize.Tag = ""
'    Call SetCtlFocus
End Sub
Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent TitleMouseDown(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
End Sub
Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent TitleMouseDown(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
End Sub

Private Sub lstDate_DblClick()
    If Element Is Nothing Then Exit Sub
    With Element
        Select Case .要素长度
            Case 19
                .内容文本 = lstDate.Text & " " & lstTime.Text
                RaiseEvent pOk
        End Select
    End With
End Sub

Private Sub lstDate_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Element Is Nothing Then Exit Sub
    
    If lstDate.ListIndex >= 0 Then
        lstDate.Tag = Val(lstDate.ListIndex)
    End If
    
    With Element
        Select Case .要素长度
            Case 10
                .内容文本 = lstDate.Text
                RaiseEvent pOk
            Case 19
                .内容文本 = lstDate.Text & " " & lstTime.Text
        End Select
    End With
End Sub

Private Sub lstTime_DblClick()
    If Element Is Nothing Then Exit Sub
    With Element
        Select Case .要素长度
            Case 19
                .内容文本 = lstDate.Text & " " & lstTime.Text
                RaiseEvent pOk
        End Select
    End With
End Sub

Private Sub lstTime_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Element Is Nothing Then Exit Sub
    
    If lstTime.ListIndex >= 0 Then
        lstTime.Tag = Val(lstTime.ListIndex)
    End If
    With Element
        Select Case .要素长度
            Case 8
                .内容文本 = lstTime.Text
                RaiseEvent pOk
            Case 19
                .内容文本 = lstDate.Text & " " & lstTime.Text
        End Select
    End With
End Sub

Private Sub picClock_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim xLine As Long, yLine As Long
    Dim dblSin As Double, dblValue As Double
    If mblnHandMove = False Then Exit Sub
    xLine = x - (shpCenter.Left + shpCenter.Width / 2)
    yLine = y - (shpCenter.Top + shpCenter.Height / 2)
    
    If xLine = 0 And yLine = 0 Then Exit Sub
    
    dblSin = yLine / Sqr(xLine ^ 2 + yLine ^ 2)
    If dblSin = 1 Then
        dblValue = 6
    ElseIf dblSin = -1 Then
        dblValue = 0
    Else
        If Sgn(xLine) >= 0 Then
            dblValue = Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1) + 3
        Else
            dblValue = 9 - Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1)
        End If
    End If
    If lblAmOrPm.Caption = "下午" And dblValue < 12 Then dblValue = dblValue + 12
    Call SetTimer(dblValue)
End Sub

Private Sub picStatus_Resize()
imgResize.Move picStatus.ScaleWidth - imgResize.Width, 0
lblInfo.Move 80, 0, picStatus.Width - imgResize.Width
End Sub
Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent TitleMouseDown(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
End Sub
'
'Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent TitleMouseMove(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
'End Sub

'Private Sub picTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent TitleMouseUp(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
'End Sub

Private Sub picTitle_Resize()
    imgTitle.Move (picTitle.ScaleWidth - imgTitle.Width) / 2, 30
End Sub
Private Sub txtDate_Validate(Cancel As Boolean)
'txtdate.tag=最小日期|最大日期|最小时间|最大时间
    If Element Is Nothing Then Exit Sub
    With Element
        txtDate.Text = Trim(txtDate.Text)
        If IsDate(txtDate.Text) Then
            Select Case .要素长度
                Case 8
                    If CDate(txtDate.Text) > CDate(Split(txtDate.Tag, "|")(3)) Then txtDate.Text = Split(txtDate.Tag, "|")(3)
                    If CDate(txtDate.Text) < CDate(Split(txtDate.Tag, "|")(2)) Then txtDate.Text = Split(txtDate.Tag, "|")(2)
                Case 10
                    If CDate(txtDate.Text) > CDate(Split(txtDate.Tag, "|")(1)) Then txtDate.Text = Split(txtDate.Tag, "|")(1)
                    If CDate(txtDate.Text) < CDate(Split(txtDate.Tag, "|")(0)) Then txtDate.Text = Split(txtDate.Tag, "|")(0)
                Case 19
                    If CDate(txtDate.Text) > CDate(Split(txtDate.Tag, "|")(1) & " " & Split(txtDate.Tag, "|")(3)) Then txtDate.Text = Split(txtDate.Tag, "|")(1) & " " & Split(txtDate.Tag, "|")(3)
                    If CDate(txtDate.Text) < CDate(Split(txtDate.Tag, "|")(0) & " " & Split(txtDate.Tag, "|")(2)) Then txtDate.Text = Split(txtDate.Tag, "|")(0) & " " & Split(txtDate.Tag, "|")(2)
            End Select
            .内容文本 = Format(txtDate.Text, txtDate.Format)
        Else
            .内容文本 = Format(Now, txtDate.Format)
        End If
    End With
    UserControl_KeyPress vbKeyReturn
End Sub

Private Sub txtDate_ValidationError(InvalidText As String, StartPosition As Integer)
    StartPosition = 0
End Sub

Private Sub txtFind_GotFocus()
    If Element Is Nothing Then Exit Sub
    txtFind.SelStart = 0: txtFind.SelLength = 1000
    lblInfo = Element.要素名称 & " 输入希望查找项目的编码/名称/简码"
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call DoFind
        Exit Sub
    End If
End Sub

Private Sub txtTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call UserControl_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub UserControl_EnterFocus()
    SetCtlFocus
End Sub

Private Sub UserControl_Hide()
    If Not Element Is Nothing Then
        With Element
            If .输入形态 = 0 Then
                SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\ElementEdit\" & .要素类型 & .要素表示 & .替换域, "\MainWidth", Width
                SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\ElementEdit\" & .要素类型 & .要素表示 & .替换域, "\MainHeight", Height
                If .要素类型 = 2 Then  '记录日期字符格式 记录时间字符格式
                    If lstDate.ListIndex >= 0 Then
                        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\ElementEdit\" & .要素类型 & .要素表示 & .替换域, "\DateType", lstDate.ListIndex
                    End If
                    If lstTime.ListIndex >= 0 Then
                        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\ElementEdit\" & .要素类型 & .要素表示 & .替换域, "\TimeType", lstTime.ListIndex
                    End If
                End If
            Else
                If .要素表示 = 3 Then '复选框展开形式,选中是即时生效的
                    UserControl_KeyPress vbKeyReturn
                End If
            End If
        End With
    End If
    Set Element = Nothing
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Element Is Nothing Then Exit Sub
    If KeyCode = vbKeyDelete Then
        If Element.要素类型 = 2 Then '日期型
            Element.内容文本 = ""
            RaiseEvent pOk
        End If
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Element.要素类型 = 0 Then
            '数值型
            Dim T As Variant, dblMax As Double, dblMin As Double
            T = Split(Element.要素值域, ";")    '格式:  0;100000
            If UBound(T) < 1 Then
                dblMin = 0#
                dblMax = 0#
            Else
                dblMin = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(1)), CLng(T(0)))
                dblMax = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(0)), CLng(T(1)))
            End If
            If Element.要素表示 = 0 Then
                '文本表示
                If Trim(txt文本) = "" Then
                    Element.内容文本 = ""
                ElseIf Element.要素值域 <> ";" And Element.要素值域 <> "0;0" And Element.要素值域 <> "" Then
                    If Val(txt文本) > dblMax Then
                        txt文本 = dblMax
                    ElseIf Val(txt文本) < dblMin Then
                        txt文本 = dblMin
                    End If
                    Element.内容文本 = IIf(Element.要素小数 > 0, Format(txt文本, "0." & String(Element.要素小数, "0")), txt文本)
                Else
                    Element.内容文本 = IIf(Element.要素小数 > 0, Format(txt文本, "0." & String(Element.要素小数, "0")), txt文本)
                End If
            ElseIf Element.要素表示 = 1 Then
                '上下表示
                If Trim(Element.内容文本) <> "" And Element.要素值域 <> ";" And Element.要素值域 <> "0;0" Then
                    If Val(Element.内容文本) > dblMax Then
                        Element.内容文本 = dblMax
                    ElseIf Val(Element.内容文本) < dblMin Then
                        Element.内容文本 = dblMin
                    End If
                Else
                    Element.内容文本 = IIf(Element.要素小数 > 0, Format(Element.内容文本, "0." & String(Element.要素小数, "0")), Element.内容文本)
                End If
            End If
        ElseIf Element.要素类型 = 2 Then '日期/时间型要素
            If lstDate.Visible And lstTime.Visible = False Then '日期型
                Element.内容文本 = lstDate.Text
            ElseIf lstDate.Visible = False And lstTime.Visible Then '时间型
                If Not IsDate(txtTime.Text) Then
                    txtTime.Text = Format(Val(Mid(txtTime, 1, 2)) & ":" & Val(Mid(txtTime, 4, 2)) & ":" & Val(Mid(txtTime, 7, 2)), "00:00:00")
                End If
                Element.内容文本 = lstTime.Text
            ElseIf lstDate.Visible And lstTime.Visible Then '日期时间型
                If Not IsDate(txtTime.Text) Then
                    txtTime.Text = Format(Val(Mid(txtTime, 1, 2)) & ":" & Val(Mid(txtTime, 4, 2)) & ":" & Val(Mid(txtTime, 7, 2)), "00:00:00")
                End If
                Element.内容文本 = lstDate.Text & " " & lstTime.Text
            End If
        End If
        If Element.替换域 <> 2 Then '字典项目由选中触发
            RaiseEvent pOk
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        RaiseEvent pCancel
    ElseIf KeyAscii = vbKeySpace Then
        If vfg单选.Visible Then vfg单选_KeyDown KeyAscii, 0
    ElseIf KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Then
        If vfg单选.Visible Then vfg单选_KeyDown KeyAscii, 0
    ElseIf InStr("1234567890", Chr(KeyAscii)) > 0 And (Element.要素表示 = 2 Or Element.要素表示 = 3) Then
        mPressKeyAscii = KeyAscii
        Call UserControl_Show
    End If
End Sub
Private Sub UserControl_Resize()
Dim lX As Long, lY As Long
    On Error Resume Next
    lX = Screen.TwipsPerPixelX:     lY = Screen.TwipsPerPixelY
    txt上下1.Visible = False:       txt上下2.Visible = False:       lblDot.Visible = False:         lbl单位.Visible = False:    shpBorder1.Visible = False
    shpBorder2.Visible = False:     txt文本.Visible = False:        ud上下.Visible = False:         vfg单选.Visible = False:    vfg复选.Visible = False
    picTime.Visible = False:        picDate.Visible = False:        cbrThis.Visible = False:        txtFind.Visible = False:    vgdList.Visible = False
    shpBorder1.BorderWidth = 1:     shpBorder2.BorderWidth = 1:     txtDate.Visible = False
    
    picTitle.Move 60, 60, ScaleWidth - 120
    picStatus.Move lX, ScaleHeight - picStatus.Height - lY, ScaleWidth - lX * 2
    shpBorderOut.Move 0, 0, Width, Height: lblInfo = "ESC 取消退出；回车:保存修改。"
    If Element Is Nothing Then Exit Sub

    With Element
        If .输入形态 = 1 Then
            picTitle.Visible = False: picStatus.Visible = False: shpBorderOut.Visible = False
        Else
            picTitle.Visible = True: picStatus.Visible = True: shpBorderOut.Visible = True
        End If
        If .替换域 = 2 Then
            Dim ltxtWidth As Long, lvgdListHeight As Long
            lblInfo = .要素名称 & " 输入希望查找项目的编码/名称/简码后回车"
            cbrThis.Move ScaleWidth - 80 - cbrThis.Width, picTitle.Height + 120
            ltxtWidth = Width - cbrThis.Width - 240: If ltxtWidth < 100 Then ltxtWidth = 100
            txtFind.Move 80, cbrThis.Top, ltxtWidth, cbrThis.Height
            vgdList.Move 80, txtFind.Top + txtFind.Height + lX * 4, ScaleWidth - 160, IIf(ScaleHeight - picStatus.Height - picTitle.Height - txtFind.Height - 250 < 0, 0, ScaleHeight - picStatus.Height - picTitle.Height - txtFind.Height - 250)
            shpBorder1.Move vgdList.Left - lX, vgdList.Top - lY, vgdList.Width + lX * 3, vgdList.Height + lY * 2
            shpBorder2.Move txtFind.Left - lX, txtFind.Top - lY, txtFind.Width + lX * 2, txtFind.Height + lY * 2
            cbrThis.Visible = True:     txtFind.Visible = True:     vgdList.Visible = True: shpBorder1.Visible = True: shpBorder2.Visible = True
            txtFind.ZOrder 0
            If txtFind.Visible And txtFind.Enabled Then txtFind.SetFocus
        Else
            Select Case .要素表示
            Case 0
                Select Case .要素类型
                    Case 2      '日期型
                        If .输入形态 = 1 Then
                            If .要素长度 < 19 Then txtDate.Width = 1000 Else txtDate.Width = 1800
                            If ScaleWidth <= txtDate.Width + lX * 2 Then txtDate.Left = 0 Else txtDate.Left = (ScaleWidth - txtDate.Width - lX * 2) / 2
                            If ScaleHeight <= txtDate.Height + lY * 2 Then txtDate.Top = 0 Else txtDate.Top = (ScaleHeight - txtDate.Height - lY * 2) / 2
                            shpBorder1.Move txtDate.Left - lX, txtDate.Top - lY, txtDate.Width + lX * 2, txtDate.Height + lY * 2
                            shpBorder1.Visible = True: txtDate.Visible = True
                            If txtDate.Visible And txtDate.Enabled Then txtDate.SetFocus
                            txtDate.SelStart = 0: txtDate.SelLength = Len(txtDate)
                        Else
                            imgResize.Tag = ""
                            Select Case .要素长度
                                Case 8
                                    picTime.Move 80, picTitle + 200, picClock.Width + 80 + lstTime.Width, picClock.Height
                                    txtTime.Move picClock.Width + 80, picClock.Top: lblAmOrPm.Move txtTime.Left + txtTime.Width + 100, txtTime.Top
                                    lblTimeType.Move txtTime.Left, txtTime.Height
                                    lstTime.Move txtTime.Left, lblTimeType.Top + lblTimeType.Height, lstTime.Width, picClock.Height - txtTime.Height - lblTimeType.Height
                                    Width = picTime.Width + 160: Height = picStatus.Height + picTitle.Height + picTime.Height + 200
                                    picTime.Visible = True
                                    If picTime.Visible Then txtTime.SetFocus
                                Case 10
                                    picDate.Move 80, picTitle + 200
                                    Width = picDate.Width + 160: Height = picStatus.Height + picTitle.Height + picDate.Height + 200
                                    picDate.Visible = True
                                    If picDate.Visible Then mvwDate.SetFocus
                                Case Else
                                    picDate.Move 80, picTitle + 200
                                    picTime.Move picDate.Left + picDate.Width + 80, picTitle + 200, picClock.Width, picDate.Height
                                        txtTime.Move picClock.Left, picClock.Height: lblAmOrPm.Move txtTime.Left + txtTime.Width + 100, txtTime.Top
                                        lblTimeType.Move txtTime.Left, txtTime.Height + txtTime.Top
                                        lstTime.Move txtTime.Left, lblTimeType.Top + lblTimeType.Height, lstTime.Width, picTime.Height - picClock.Height - txtTime.Height - lblTimeType.Height
                                    picDate.Visible = True
                                    If picDate.Visible Then txtTime.SetFocus
                                    picTime.Visible = True
                                    Width = picDate.Width + picTime.Width + 240: Height = picStatus.Height + picTitle.Height + picDate.Height + 200
                            End Select
                        End If
                    Case 3      '逻辑型
                        vfg单选.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
                        shpBorder1.Move vfg单选.Left - lX, vfg单选.Top - lY, vfg单选.Width + lX * 3, vfg单选.Height + lY * 2
                        vfg单选.Cell(flexcpAlignment, 0, 1, vfg单选.Rows - 1, 1) = flexAlignLeftCenter: vfg单选.Visible = True
                        shpBorder1.Visible = True:       vfg单选.BackColorSel = &HFFC0C0
                    Case Else
                        txt文本.Move 80, picTitle.Height + 120, ScaleWidth - 160, IIf(ScaleHeight - 200 - picStatus.Height - picTitle.Height < 0, 0, ScaleHeight - 200 - picStatus.Height - picTitle.Height)
                        shpBorder1.Move txt文本.Left - lX, txt文本.Top - lY, txt文本.Width + lX * 2, txt文本.Height + lY * 2
                        txt文本.Visible = True: shpBorder1.Visible = True
                        If txt文本.Visible And txt文本.Enabled Then txt文本.SetFocus
                End Select
            Case 1
                Dim lW1 As Long, lW2 As Long, lW3 As Long, lW4 As Long, lW5 As Long
                If Trim(Element.要素单位) <> "" Then
                    lbl单位.Width = TextWidth(lbl单位) + lX * 6
                    lbl单位.Move ScaleWidth - lbl单位.Width + lX * 3, picTitle.Height + 170
                    lbl单位.Visible = True
                    lW5 = lbl单位.Width
                Else
                    lbl单位.Visible = False
                    lW5 = 0
                End If
                lW4 = ud上下.Width + lX * 4
                ud上下.Move ScaleWidth - lW4 - lW5 + lX * 3, picTitle.Height + 120
                ud上下.Visible = True
                If Element.要素小数 > 0 Then
                    txt上下2.Width = TextWidth(Space(Element.要素小数)) + lX * 4
                    lW3 = txt上下2.Width + lX
                    txt上下2.Move ScaleWidth - lW5 - lW4 - lW3 + lX, picTitle.Height + 170
                    shpBorder2.Move txt上下2.Left - lX, txt上下2.Top - lY - 50, txt上下2.Width + lX * 2, txt上下2.Height + 50 + lY * 2
                    shpBorder2.Visible = True
                    txt上下2.Visible = True
                    lblDot.Width = TextWidth(".") + lX * 2
                    lW2 = lblDot.Width
                    lblDot.Move txt上下2.Left - lW2 + lX * 2, picTitle.Height + 170
                    lblDot.BackStyle = 0
                    lblDot.Visible = True
                Else
                    lW2 = 0
                    lW3 = 0
                    shpBorder2.Visible = False
                    txt上下2.Visible = False
                    lblDot.Visible = False
                End If
                lW1 = TextWidth(txt上下1.Text) + lX * 2
                lW1 = IIf(lW1 < 400, 400, lW1)
                
                If Width < lW1 + lW2 + lW3 + lW4 + lW5 Then Width = lW1 + lW2 + lW3 + lW4 + lW5
                Height = txt上下1.Height + lY * 3 + picStatus.Height + picTitle.Height + 180
                
                txt上下1.Move 80, picTitle.Height + 170, ScaleWidth - lW5 - lW4 - lW3 - lW2 - lX * 4
                shpBorder1.Move txt上下1.Left - lX, txt上下1.Top - lY - 50, txt上下1.Width + lX * 2, txt上下1.Height + 50 + lY * 2
                txt上下1.Visible = True
                shpBorder1.Visible = True
                If txt上下1.Visible And txt上下1.Enabled Then txt上下1.SelStart = 0: txt上下1.SelLength = Len(txt上下1): txt上下1.SetFocus
            Case 2
                If .输入形态 = 0 Then '弹出
                    vfg单选.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
                    shpBorder1.Move vfg单选.Left - lX, vfg单选.Top - lY, vfg单选.Width + lX * 3, vfg单选.Height + lY * 2
                    vfg单选.Cell(flexcpAlignment, 0, 1, vfg单选.Rows - 1, 1) = flexAlignLeftCenter: vfg单选.Visible = True
                    shpBorder1.Visible = True: vfg单选.BackColorSel = &HFFC0C0: vfg单选.HighLight = flexHighlightAlways
                Else                   '展开
                    Dim i As Byte
                    vfg单选.BackColorSel = &HFFFFFF: vfg单选.ForeColorSel = &H0&
                    For i = 0 To vfg单选.Cols - 1
                        If i Mod 2 = 0 Then
                            vfg单选.ColWidth(i) = 250
                        Else
                            vfg单选.ColWidth(i) = IIf((ScaleWidth - 50 - vfg单选.Cols / 2 * 250) / (vfg单选.Cols / 2) > 0, (ScaleWidth - 50 - vfg单选.Cols / 2 * 250) / (vfg单选.Cols / 2), 0)
                        End If
                        vfg单选.Cell(flexcpAlignment, 0, i) = flexAlignLeftCenter
                    Next
                    vfg单选.Move lX, lY, ScaleWidth - lX * 2, ScaleHeight - lY * 2: vfg单选.RowHeight(0) = vfg单选.Height - lX * 2
                    vfg单选.Visible = True: vfg单选.HighLight = flexHighlightNever: vfg单选.Col = 0: vfg单选.Refresh
                End If
                If vfg单选.Visible And vfg单选.Enabled Then vfg单选.SetFocus
            Case 3
                If .输入形态 = 0 Then '弹出
                    vfg复选.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
                    shpBorder1.Move vfg复选.Left - lX, vfg复选.Top - lY, vfg复选.Width + lX * 3, vfg复选.Height + lY * 2
                    vfg复选.Visible = True: shpBorder1.Visible = True: vfg复选.HighLight = flexHighlightAlways: vfg复选.BackColorSel = &HFFC0C0
                Else                '展开
                    vfg复选.Move 0, 0, ScaleWidth, ScaleHeight: vfg复选.RowHeight(0) = ScaleHeight
                    vfg复选.BackColorSel = &HFFFFFF: vfg复选.ForeColorSel = &H0&
                    For i = 0 To vfg复选.Cols - 1
                        vfg复选.ColWidth(i) = ScaleWidth / vfg复选.Cols
                        vfg复选.Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                    Next
                    vfg复选.Visible = True: vfg复选.HighLight = flexHighlightNever: vfg复选.Col = 0: vfg复选.Refresh
                End If
                If vfg复选.Visible And vfg复选.Enabled Then vfg复选.SetFocus
            End Select
        End If
    End With
    Err.Clear
End Sub
Private Sub txt上下1_Change()
    If Element Is Nothing Then Exit Sub
    If txt上下1.Tag = "" Then
        Element.内容文本 = Trim(txt上下1.Text) & IIf(Element.要素小数 > 0, "." & Format(Trim(txt上下2.Text), String(Element.要素小数, "0")), "")
    End If
End Sub

Private Sub txt文本_Change()
    If Element Is Nothing Then Exit Sub
    Element.内容文本 = Trim(txt文本.Text)
End Sub
Private Sub SetCtlFocus()
    '设置控件焦点
    If txt上下1.Visible And txt上下1.Enabled Then
        txt上下1.SetFocus
    ElseIf txt上下2.Visible And txt上下2.Enabled Then
        txt上下2.SetFocus
    ElseIf txt文本.Visible And txt文本.Enabled Then
        txt文本.SetFocus
    ElseIf vfg单选.Visible And vfg单选.Enabled Then
        vfg单选.SetFocus
    ElseIf vfg复选.Visible And vfg复选.Enabled Then
        vfg复选.SetFocus
    ElseIf txtTime.Visible And txtTime.Enabled Then
        txtTime.SetFocus
    ElseIf lstDate.Visible And lstDate.Enabled And lstTime.Visible = False Then
        mvwDate.SetFocus
    End If
End Sub

Private Sub UserControl_Show()
Dim strInput As String
    If Element Is Nothing Then Exit Sub
    With Element
        '单选复选通过输入1234567890直接定位
        If InStr("1234567890", Chr(mPressKeyAscii)) > 0 Then
            Dim PressN As Integer, i As Integer, strValue As String
            PressN = CByte(Chr(mPressKeyAscii))
            Select Case .要素表示
                Case 2
                    vfg单选.Visible = False
                    If .输入形态 = 1 Then '展开型
                        i = PressN * 2 - 1
                        If i < 0 Then i = 0
                        If i > vfg单选.Cols Then i = 1
                        vfg单选.Col = i
                        For i = 0 To vfg单选.Cols - 1 Step 2
                            If i = vfg单选.Col Or i = vfg单选.Col - 1 Then
                                If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                                    vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                                Else
                                    vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                                End If
                            Else
                                vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                            End If
                            
                            If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                                If Trim(vfg单选.Cell(flexcpText, 0, i + 1)) = "自定义" And .动态域 = 1 Then
                                    strInput = MidUni(Trim(InputBox("请录入自定义要素选项", gstrSysName)), 1, 200)
                                    strValue = strValue & IIf(strInput = "", "○", "●") & IIf(strInput = "", "自定义", strInput)
                                Else
                                    strValue = strValue & "●" & vfg单选.Cell(flexcpText, 0, i + 1)
                                End If
                            Else
                                strValue = strValue & "○" & vfg单选.Cell(flexcpText, 0, i + 1)
                            End If
                        Next
                    Else '弹出型
                        i = PressN - 1
                        If i < 0 Then i = 0
                        If i > vfg单选.Rows - 1 Then i = 0
                        vfg单选.Row = i
                        For i = 0 To vfg单选.Rows - 1
                            If i = vfg单选.Row Then
                                If vfg单选.Cell(flexcpPicture, i, 0) = imgOpt2.Picture Then
                                    vfg单选.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                                Else
                                    vfg单选.Cell(flexcpPicture, i, 0) = imgOpt2.Picture
                                End If
                            Else
                                vfg单选.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                            End If
                        Next
                        
                        If vfg单选.Cell(flexcpPicture, vfg单选.Row, 0) = imgOpt2.Picture Then
                            strValue = vfg单选.Cell(flexcpText, vfg单选.Row, 1)
                            If strValue = "自定义" And .动态域 = 1 Then
                                strValue = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                            End If
                        Else
                            strValue = ""
                        End If
                    End If
                    .内容文本 = strValue
                    UserControl_KeyPress vbKeyReturn
                Case 3
                    If .输入形态 = 1 Then
                        vfg复选.Visible = False
                        PressN = PressN - 1
                        If PressN < 0 Or PressN >= vfg复选.Cols Then PressN = 0
                        If vfg复选.Cell(flexcpChecked, 0, PressN) = flexChecked Then
                            vfg复选.Cell(flexcpChecked, 0, PressN) = flexUnchecked
                        Else
                            vfg复选.Cell(flexcpChecked, 0, PressN) = flexChecked
                        End If
        
                        For i = 0 To vfg复选.Cols - 1
                            If vfg复选.Cell(flexcpChecked, 0, i) = flexChecked Then
                                If Trim(vfg复选.Cell(flexcpText, 0, i)) = "自定义" And .动态域 = 1 Then
                                    strInput = MidUni(Trim(InputBox("请录入自定义要素选项", gstrSysName)), 1, 200)
                                    strValue = strValue & IIf(strInput = "", "□", "■") & IIf(strInput = "", "自定义", strInput)
                                Else
                                    strValue = strValue & "■" & vfg复选.Cell(flexcpText, 0, i)
                                End If
                            Else
                                strValue = strValue & "□" & vfg复选.Cell(flexcpText, 0, i)
                            End If
                        Next
                        .内容文本 = strValue
                        UserControl_KeyPress vbKeyReturn
                    Else
                        i = PressN - 1
                        If i < 0 Then i = 0
                        If i >= vfg复选.Rows Then i = 0
                        If vfg复选.Cell(flexcpChecked, i, 0) = flexChecked Then
                            vfg复选.Cell(flexcpChecked, i, 0) = flexUnchecked
                        Else
                            vfg复选.Cell(flexcpChecked, i, 0) = flexChecked
                        End If
                         
                        For i = 0 To vfg复选.Rows - 1
                            If vfg复选.Cell(flexcpChecked, i, 0) = flexChecked Then
                                If vfg复选.Cell(flexcpText, i, 0) = "自定义" And Element.动态域 = 1 Then
                                    strInput = MidUni(Trim(InputBox("请录入自定义要素选项", gstrSysName)), 1, 200)
                                    strValue = strValue & "、" & strInput
                                Else
                                    strValue = strValue & "、" & vfg复选.Cell(flexcpText, i, 0)
                                End If
                            End If
                        Next
                        strValue = Mid(strValue, 2)
                        .内容文本 = strValue
                        RaiseEvent pChange
                    End If
            End Select
        End If
    End With
End Sub
Private Sub txt上下1_GotFocus()
    zlCommFun.OpenIme
    txt上下1.SelStart = 0:              txt上下1.SelLength = Len(txt上下1)
    ud上下.BuddyControl = txt上下1:     ud上下.BuddyProperty = "Text"
End Sub

Private Sub txt上下1_KeyPress(KeyAscii As Integer)
    If InStr("1234567890. " & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = vbKeySpace Or InStr(".", Chr(KeyAscii)) = 1 Then
        KeyAscii = 0
        If txt上下2.Visible And txt上下2.Enabled Then
            txt上下2.SelStart = 0
            txt上下2.SelLength = Len(txt上下2)
            txt上下2.SetFocus
        End If
    End If
End Sub

Private Sub txt上下2_Change()
    If Element Is Nothing Then Exit Sub
    If txt上下1.Tag = "" Then
        If Element.要素小数 > 0 Then
            Dim lngLen As Long, strR As String
            lngLen = Len(Trim(txt上下2))
            If lngLen > Element.要素小数 Then
                strR = Trim(txt上下1.Text) & "." & Trim(txt上下2) & String(Element.要素小数 - Len(Trim(txt上下2)), "0")
            Else
                strR = Trim(txt上下1.Text) & "." & Left(Trim(txt上下2), Element.要素小数)
            End If
        Else
            strR = Trim(txt上下1.Text)
        End If
        Element.内容文本 = IIf(Element.要素小数 > 0, Format(strR, "0." & String(Element.要素小数, "0")), strR)
    End If
End Sub

Private Sub txt上下2_GotFocus()
    zlCommFun.OpenIme
    txt上下2.SelStart = 0:                      txt上下2.SelLength = Len(txt上下2)
    ud上下.BuddyControl = txt上下2:             ud上下.BuddyProperty = "Text"
    ud上下.Move txt上下2.Left + txt上下2.Width + 15
End Sub

Private Sub txt上下2_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txt文本_GotFocus()
    If Element Is Nothing Then Exit Sub
    If Element.要素类型 = 0 Then
        zlCommFun.OpenIme
    End If
End Sub
Private Sub txt文本_KeyPress(KeyAscii As Integer)
    If Element Is Nothing Then Exit Sub
    If Element.要素类型 = 0 Then
        '数值型的控制：只能输入数字（小数点和负号，且小数点只能为1个，不能在开头；负号只能在开始处）
        'Asc(".") = vbKeyDelete = 46
        If Len(txt文本.Text) = 0 And KeyAscii = 46 Then KeyAscii = 0
        If InStr(1, txt文本.Text, ".") <> 0 And KeyAscii = 46 Then
            KeyAscii = 0
        ElseIf InStr(1, txt文本.Text, ".") = 0 And KeyAscii = 46 And txt文本.SelLength = Len(txt文本) And txt文本.SelStart = 0 Then
            KeyAscii = 0
        End If
        If txt文本.Text = "-" And KeyAscii = 46 Then KeyAscii = 0
        If KeyAscii = vbKeyBack Or KeyAscii = 46 Then Exit Sub
        If KeyAscii = Asc("-") Then
            If txt文本.SelStart <> 0 Then KeyAscii = 0
        Else
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub vfg单选_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strValue As String, PressN As Integer, strInput As String

    On Error Resume Next
    If Element Is Nothing Then Exit Sub
    If Not KeyCode = vbKeySpace Then Exit Sub
    '空格选中
    strValue = ""
    If Element.输入形态 = 0 Then '弹出式
        For i = 0 To vfg单选.Rows - 1
            If i = vfg单选.Row Then
                If vfg单选.Cell(flexcpPicture, i, 0) = imgOpt2.Picture Then
                    vfg单选.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                Else
                    vfg单选.Cell(flexcpPicture, i, 0) = imgOpt2.Picture
                End If
            Else
                vfg单选.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
            End If
        Next

        If vfg单选.Cell(flexcpPicture, vfg单选.Row, 0) = imgOpt2.Picture Then
            strValue = vfg单选.Cell(flexcpText, vfg单选.Row, 1)
        Else
            If Element.要素类型 = 3 Then
                strValue = vfg单选.Cell(flexcpText, 1, 1)
                If strValue = "自定义" And Element.动态域 = 1 Then
                    strInput = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                    strValue = strInput
                End If
            Else
                strValue = ""
            End If
        End If
    Else '展开式
        For i = 0 To vfg单选.Cols - 1 Step 2
            If i = vfg单选.Col Or i = vfg单选.Col - 1 Then
                If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                    vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                Else
                    vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                End If
            Else
                vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
            End If
            If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                If Trim(vfg单选.Cell(flexcpText, 0, i + 1)) = "自定义" And Element.动态域 = 1 Then
                    strInput = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                    strValue = strValue & IIf(strInput = "", "○", "●") & IIf(strInput = "", "自定义", strInput)
                Else
                    strValue = strValue & "●" & vfg单选.Cell(flexcpText, 0, i + 1)
                End If
            Else
                strValue = strValue & "○" & vfg单选.Cell(flexcpText, 0, i + 1)
            End If
        Next
    End If
    Element.内容文本 = strValue
    KeyCode = 0
    UserControl_KeyPress vbKeyReturn
    Err.Clear
End Sub

Private Sub vfg单选_KeyPress(KeyAscii As Integer)
Dim i As Long, strValue As String, PressN As Integer, strInput As String
    If Element Is Nothing Then Exit Sub
    If InStr("1234567890", Chr(KeyAscii)) > 0 And Element.输入形态 = 1 Then
        PressN = CByte(Chr(KeyAscii))
        With Element
            i = PressN * 2 - 1
            If i < 0 Then i = 0
            If i > vfg单选.Cols Then i = 1
            vfg单选.Col = i
            For i = 0 To vfg单选.Cols - 1 Step 2
                If i = vfg单选.Col Or i = vfg单选.Col - 1 Then
                    If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                        vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                    Else
                        vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                    End If
                Else
                    vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                End If
    
                If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                    If Trim(vfg单选.Cell(flexcpText, 0, i + 1)) = "自定义" And .动态域 = 1 Then
                        strInput = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                        strValue = strValue & IIf(strInput = "", "○", "●") & IIf(strInput = "", "自定义", strInput)
                    Else
                        strValue = strValue & "●" & vfg单选.Cell(flexcpText, 0, i + 1)
                    End If
                Else
                    strValue = strValue & "○" & vfg单选.Cell(flexcpText, 0, i + 1)
                End If
            Next
            KeyAscii = 0
            Element.内容文本 = strValue
            UserControl_KeyPress vbKeyReturn
         End With
    End If
End Sub


Private Sub vfg单选_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, strValue As String, strInput As String
    If mbEt >= 2 Then Exit Sub
    If Element Is Nothing Then Exit Sub
    strValue = ""

    If Button = vbLeftButton Then
        If Element.输入形态 = 0 Then
            For i = 0 To vfg单选.Rows - 1
                If i = vfg单选.Row Then
                    If vfg单选.Cell(flexcpPicture, i, 0) = imgOpt2.Picture Then
                        vfg单选.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                    Else
                        vfg单选.Cell(flexcpPicture, i, 0) = imgOpt2.Picture
                    End If
                Else
                    vfg单选.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                End If
            Next

            If vfg单选.Cell(flexcpPicture, vfg单选.Row, 0) = imgOpt2.Picture Then
                strValue = vfg单选.Cell(flexcpText, vfg单选.Row, 1)
                If strValue = "自定义" And Element.动态域 = 1 Then
                    strValue = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                End If
            Else
                If Element.要素类型 = 3 Then
                    strValue = vfg单选.Cell(flexcpText, 1, 1)
                Else
                    strValue = ""
                End If
            End If
        Else
            For i = 0 To vfg单选.Cols - 1 Step 2
                If i = vfg单选.Col Or i = vfg单选.Col - 1 Then
                    If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                        vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                    Else
                        vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                    End If
                Else
                    vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                End If
                If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                    If Trim(vfg单选.Cell(flexcpText, 0, i + 1)) = "自定义" And Element.动态域 = 1 Then
                        strInput = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                        strValue = strValue & IIf(strInput = "", "○", "●") & IIf(strInput = "", "自定义", strInput)
                    Else
                        strValue = strValue & "●" & vfg单选.Cell(flexcpText, 0, i + 1)
                    End If
                Else
                    strValue = strValue & "○" & vfg单选.Cell(flexcpText, 0, i + 1)
                End If
            Next
        End If
    End If
    Element.内容文本 = strValue
    UserControl_KeyPress vbKeyReturn
End Sub

Private Sub vfg单选_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, strValue As String, strInput As String
    If mbEt < 2 Then Exit Sub
    If Element Is Nothing Then Exit Sub
    strValue = ""

    If Button = vbLeftButton Then
        If Element.输入形态 = 0 Then
            For i = 0 To vfg单选.Rows - 1
                If i = vfg单选.Row Then
                    If vfg单选.Cell(flexcpPicture, i, 0) = imgOpt2.Picture Then
                        vfg单选.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                    Else
                        vfg单选.Cell(flexcpPicture, i, 0) = imgOpt2.Picture
                    End If
                Else
                    vfg单选.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                End If
            Next

            If vfg单选.Cell(flexcpPicture, vfg单选.Row, 0) = imgOpt2.Picture Then
                strValue = vfg单选.Cell(flexcpText, vfg单选.Row, 1)
                If strValue = "自定义" And Element.动态域 = 1 Then
                    strValue = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                End If
            Else
                If Element.要素类型 = 3 Then
                    strValue = vfg单选.Cell(flexcpText, 1, 1)
                Else
                    strValue = ""
                End If
            End If
        Else
            For i = 0 To vfg单选.Cols - 1 Step 2
                If i = vfg单选.Col Or i = vfg单选.Col - 1 Then
                    If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                        vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                    Else
                        vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                    End If
                Else
                    vfg单选.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                End If
                If vfg单选.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                    If Trim(vfg单选.Cell(flexcpText, 0, i + 1)) = "自定义" And Element.动态域 = 1 Then
                        strInput = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                        strValue = strValue & IIf(strInput = "", "○", "●") & IIf(strInput = "", "自定义", strInput)
                    Else
                        strValue = strValue & "●" & vfg单选.Cell(flexcpText, 0, i + 1)
                    End If
                Else
                    strValue = strValue & "○" & vfg单选.Cell(flexcpText, 0, i + 1)
                End If
            Next
        End If
    End If
    Element.内容文本 = strValue
    UserControl_KeyPress vbKeyReturn
End Sub

Private Sub vfg单选_RowColChange()
    If Element Is Nothing Then Exit Sub
    If Element.输入形态 = 1 Then
        vfg单选.Cell(flexcpBackColor, 0, 0, 0, vfg单选.Cols - 1) = 0
        vfg单选.Cell(flexcpBackColor, 0, vfg单选.Col) = &HFFC0C0
    End If
End Sub

Private Sub vfg复选_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, strValue As String, strInput As String
    If Element Is Nothing Then Exit Sub
    strValue = ""
    With Element
        If .输入形态 = 0 Then
            For i = 0 To vfg复选.Rows - 1
                If vfg复选.Cell(flexcpChecked, i, 0) = flexChecked Then
                    If vfg复选.Cell(flexcpText, i, 0) = "自定义" And .动态域 = 1 Then
                        strInput = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                        strValue = strValue & "、" & IIf(strInput = "", "自定义", strInput)
                    Else
                        strValue = strValue & "、" & vfg复选.Cell(flexcpText, i, 0)
                    End If
                End If
            Next
            strValue = Mid(strValue, 2)
        Else
            For i = 0 To vfg复选.Cols - 1
                If vfg复选.Cell(flexcpChecked, 0, i) = flexChecked Then
                    If Trim(vfg复选.Cell(flexcpText, 0, i)) = "自定义" And .动态域 = 1 Then
                        strInput = MidUni((Trim(InputBox("请录入自定义要素选项", gstrSysName))), 1, 200)
                        strValue = strValue & IIf(strInput = "", "□", "■") & IIf(strInput = "", "自定义", strInput)
                    Else
                        strValue = strValue & "■" & vfg复选.Cell(flexcpText, 0, i)
                    End If
                Else
                    strValue = strValue & "□" & vfg复选.Cell(flexcpText, 0, i)
                End If
            Next
        End If
        .内容文本 = strValue
        RaiseEvent pChange
    End With
End Sub
Private Sub vfg复选_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Element Is Nothing Then Exit Sub
    If Element.输入形态 = 0 Then
        vfg复选.Col = 0
    End If
End Sub

Private Sub vfg单选_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    vfg单选.Col = 0
    Cancel = True
End Sub

Private Sub txtTime_Change()
    If mblnHandTime Then Exit Sub
    Call SetTimer(Val(Mid(txtTime, 1, 2)) + Val(Mid(txtTime, 4, 2)) / 60 + Val(Mid(txtTime, 7, 2)) / 60 / 60)
End Sub


Private Sub lblAmOrPm_Click()
    If Not IsDate(txtTime.Text) Then
        txtTime.SetFocus
        Exit Sub
    End If
    If lblAmOrPm.Caption = "下午" Then
        lblAmOrPm.Caption = "上午"
        txtTime.Text = Format(CDate(txtTime.Text) - 12 / 24, "hh:mm:ss")
    Else
        lblAmOrPm.Caption = "下午"
        txtTime.Text = Format(CDate(txtTime.Text) + 12 / 24, "hh:mm:ss")
    End If
    Call SetTimer(Hour(txtTime.Text) + Minute(txtTime.Text) / 60 + Second(txtTime.Text) / 60 / 60)
End Sub

Private Sub mvwDate_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    Dim intYear As Integer, intMonth As Integer, intDay As Integer
    Dim strYear As String, strMonth As String, strDay As String
    
    If Element Is Nothing Then Exit Sub
    If Element.要素长度 = 8 Then Exit Sub
    intYear = Year(mvwDate.Value)
    intMonth = Month(mvwDate.Value)
    intDay = Day(mvwDate.Value)
    
    strYear = GetChineseNumber(Mid(intYear, 1, 1), True) & GetChineseNumber(Mid(intYear, 2, 1), True) & GetChineseNumber(Mid(intYear, 3, 1), True) & GetChineseNumber(Mid(intYear, 4, 1), True)
    strMonth = GetChineseNumber(intMonth)
    strDay = GetChineseNumber(intDay)
     
    With lstDate
        .Clear
        .AddItem strYear & "年" & strMonth & "月" & strDay & "日"
        .AddItem strYear & "年" & strMonth & "月"
        .AddItem strMonth & "月" & strDay & "日"
        .AddItem intYear & "年" & intMonth & "月" & intDay & "日"
        .AddItem intYear & "年" & intMonth & "月"
        .AddItem intMonth & "月" & intDay & "日"
        .AddItem intYear & "-" & intMonth & "-" & intDay
        .AddItem intMonth & "-" & intDay
        .AddItem WeekdayName(Weekday(mvwDate.Value))
        .AddItem GetSolarTerm(mvwDate.Value)
        If Val(.Tag) >= 0 Then
            .ListIndex = Val(.Tag)
            .TopIndex = .ListIndex
        End If
    End With
    If Element Is Nothing Then Exit Sub
    If Element.要素类型 = 2 Then
        If Element.要素长度 = 10 Then
            Element.内容文本 = lstDate.Text
            RaiseEvent pChange
        ElseIf Element.要素长度 > 10 Then
            Element.内容文本 = lstDate.Text & " " & lstTime.Text
            RaiseEvent pChange
        End If
    End If
End Sub

Private Sub picClock_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnHandMove = True: mblnHandTime = True
End Sub

Private Sub picClock_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim xLine As Long, yLine As Long
    Dim dblSin As Double, dblValue As Double

    mblnHandMove = False
    xLine = x - (shpCenter.Left + shpCenter.Width / 2)
    yLine = y - (shpCenter.Top + shpCenter.Height / 2)
    
    If xLine = 0 And yLine = 0 Then Exit Sub
    
    dblSin = yLine / Sqr(xLine ^ 2 + yLine ^ 2)
    If dblSin = 1 Then
        dblValue = 6
    ElseIf dblSin = -1 Then
        dblValue = 0
    Else
        If Sgn(xLine) >= 0 Then
            dblValue = Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1) + 3
        Else
            dblValue = 9 - Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1)
        End If
    End If
    If lblAmOrPm.Caption = "下午" And dblValue < 12 Then dblValue = dblValue + 12
    Call SetTimer(dblValue)
    mblnHandTime = False
End Sub

Private Function GetSolarTerm(ByVal dtAsk As Date) As String
    '功能：获得指定日期的节气
    '参数：dtAsk，公历日期
    
    Const conYearMinutes As Double = 525948.76   '每年的分钟数，一年实际是365.242194444天，按分钟计算基本能准确
    Dim dtBaseDate As Date
    Dim aryTermName() As String
    Dim aryTermData() As String
    
    Dim dblMinutes As Double
    Dim dtTermDate As Date
    
    dtAsk = Int(dtAsk) + 2 / 24 + 5 / 24 / 60
    dtBaseDate = Format("1900-01-06 2:05:00", "YYYY-MM-DD hh:mm:ss")
    If dtAsk < dtBaseDate Then GetSolarTerm = "": Exit Function
    aryTermName = Split("小寒,大寒,立春,雨水,惊蛰,春分,清明,谷雨,立夏,小满,芒种,夏至,小暑,大暑,立秋,处暑,白露,秋分,寒露,霜降,立冬,小雪,大雪,冬至", ",")
    aryTermData = Split("0,21208,42467,63836,85337,107014,128867,150921,173149,195551,218072,240693,263343,285989,308563,331033,353350,375494,397447,419210,440795,462224,483532,504758", ",")
      
    Dim intCount As Integer
    For intCount = 0 To UBound(aryTermData)
        dblMinutes = conYearMinutes * (Year(dtAsk) - 1900) + CLng(aryTermData(intCount))
        dtTermDate = DateAdd("n", dblMinutes, dtBaseDate)
        
        If DateDiff("d", dtAsk, dtTermDate) >= 0 Then
            Select Case DateDiff("d", dtAsk, dtTermDate)
            Case 0
                GetSolarTerm = aryTermName(intCount)
            Case Is < 8
                GetSolarTerm = aryTermName(intCount) & "前" & DateDiff("d", dtAsk, dtTermDate) & "天"
            Case Else
                If intCount <> 0 Then
                    dblMinutes = conYearMinutes * (Year(dtAsk) - 1900) + CLng(aryTermData(intCount - 1))
                    dtTermDate = DateAdd("n", dblMinutes, dtBaseDate)
                    GetSolarTerm = aryTermName(intCount - 1) & "后" & Abs(DateDiff("d", dtAsk, dtTermDate)) & "天"
                Else
                    dblMinutes = conYearMinutes * (Year(dtAsk) - 1 - 1900) + CLng(aryTermData(UBound(aryTermData)))
                    dtTermDate = DateAdd("n", dblMinutes, dtBaseDate)
                    GetSolarTerm = aryTermName(UBound(aryTermData)) & "后" & Abs(DateDiff("d", dtAsk, dtTermDate)) & "天"
                End If
            End Select
            Exit Function
        ElseIf intCount = UBound(aryTermData) Then
            GetSolarTerm = aryTermName(intCount) & "后" & Abs(DateDiff("d", dtAsk, dtTermDate)) & "天"
            Exit Function
        End If
    Next
    GetSolarTerm = ""
End Function

Private Sub DrawWatch()
'画表盘
    Dim CenterX As Long, CenterY As Long, lngRadii As Long
    CenterX = picClock.ScaleWidth / 2
    CenterY = picClock.ScaleHeight / 2
    If CenterX < CenterY Then
        lngRadii = CenterX - 60
    Else
        lngRadii = CenterY - 60
    End If
    
    Dim intHour As Integer, x As Long, y As Long
    x = CenterX - shpCenter.Width / 2
    y = CenterY - shpCenter.Height / 2
    shpCenter.Move x, y
    
    linHand.X1 = x + shpCenter.Width / 2
    linHand.Y1 = y + shpCenter.Height / 2
    
    For intHour = 0 To 11
        If intHour > shpDot.Count - 1 Then
            Load shpDot(intHour)
        End If
        If intHour Mod 3 = 0 Then
            shpDot(intHour).Width = 60
            shpDot(intHour).Height = 60
        Else
            shpDot(intHour).Width = 45
            shpDot(intHour).Height = 45
        End If
        x = CenterX + lngRadii * Sin(intHour * 30 / 180 * conPI) - shpDot(intHour).Width / 2
        y = CenterY - lngRadii * Cos(intHour * 30 / 180 * conPI) - shpDot(intHour).Height / 2
        shpDot(intHour).Move x, y
        shpDot(intHour).Visible = True
    Next
End Sub

Private Sub SetTimer(ByVal dblTime As Double)
'跟据时间定指针
    Dim CenterX As Long, CenterY As Long, lngRadii As Long
    Dim intCount As Integer

    If Element Is Nothing Then Exit Sub
    CenterX = picClock.ScaleWidth / 2
    CenterY = picClock.ScaleHeight / 2
    If CenterX < CenterY Then
        lngRadii = CenterX - 60
    Else
        lngRadii = CenterY - 60
    End If
    
    If dblTime < 12 Then
        lblAmOrPm.Caption = "上午"
    Else
        lblAmOrPm.Caption = "下午"
    End If
    
    If mblnHandTime = True Then
        txtTime.Text = Format(dblTime / 24, "hh:mm:ss")
    End If
    linHand.X2 = CenterX + lngRadii * Sin(dblTime * 30 / 180 * conPI)
    linHand.Y2 = CenterY - lngRadii * Cos(dblTime * 30 / 180 * conPI)
    
    For intCount = 0 To 11
        If intCount = IIf(dblTime < 12, dblTime, dblTime - 12) Then
            shpDot(intCount).BorderColor = RGB(255, 0, 0)
        Else
            shpDot(intCount).BorderColor = RGB(0, 0, 0)
        End If
    Next

    If mblnHandMove = True Then Exit Sub

    '设置格式
    Dim intHour As Integer, intMinute As Integer, intSecond As Integer
    
    intHour = Val(Mid(txtTime.Text, 1, 2))
    intMinute = Val(Mid(txtTime.Text, 4, 2))
    intSecond = Val(Mid(txtTime.Text, 7, 2))
    
    With lstTime
        .Clear
        .AddItem IIf(intHour < 12, "上午", "下午") & GetChineseNumber(IIf(intHour < 12, intHour, intHour - 12)) & "时" & GetChineseNumber(intMinute) & "分"
        .AddItem GetChineseNumber(intHour) & "时" & GetChineseNumber(intMinute) & "分"
        .AddItem IIf(intHour < 12, "上午", "下午") & IIf(intHour < 12, intHour, intHour - 12) & "时" & intMinute & "分" & intSecond & "秒"
        .AddItem IIf(intHour < 12, "上午", "下午") & IIf(intHour < 12, intHour, intHour - 12) & "时" & intMinute & "分"
        .AddItem intHour & "时" & intMinute & "分" & intSecond & "秒"
        .AddItem intHour & "时" & intMinute & "分"
        .AddItem IIf(intHour < 12, intHour, intHour - 12) & ":" & Format(intMinute, "00") & ":" & Format(intSecond, "00") & IIf(intHour < 12, " AM", " PM")
        .AddItem IIf(intHour < 12, intHour, intHour - 12) & ":" & Format(intMinute, "00") & IIf(intHour < 12, " AM", " PM")
        .AddItem intHour & ":" & Format(intMinute, "00") & ":" & Format(intSecond, "00")
        .AddItem intHour & ":" & Format(intMinute, "00")
        If Val(.Tag) >= 0 Then
            .ListIndex = Val(.Tag)
            .TopIndex = .ListIndex
        End If
    End With

    With Element
        If .要素类型 = 2 Then
            Select Case .要素长度
                Case 8
                    .内容文本 = lstTime.Text
                Case 10
                    .内容文本 = lstDate.Text
                Case 19
                    .内容文本 = lstDate.Text & " " & lstTime.Text
            End Select
            RaiseEvent pChange
        End If
    End With
End Sub

Private Function GetChineseNumber(ByVal bytNumber As Byte, Optional blnZeroCircle As Boolean) As String
    '功能：返回汉字数字
    '参数：
    '   bytNumber,要处理的数字，本函数要求不大于99;
    '   blnZeroCircle,是否以○代表0，否则表现为零
    
    Dim bytBit1 As Byte, bytBit2 As Byte
    Dim strBit1 As String, strBit2 As String
    
    If bytNumber > 99 Then GetChineseNumber = "": Exit Function
    
    bytBit1 = bytNumber \ 10: bytBit2 = bytNumber Mod 10
    
    If bytBit1 = 0 Then
        strBit1 = ""
        If blnZeroCircle = False Then
            strBit2 = Split("零,一,二,三,四,五,六,七,八,九", ",")(bytBit2)
        Else
            strBit2 = Split("○,一,二,三,四,五,六,七,八,九", ",")(bytBit2)
        End If
    Else
        strBit1 = Split(",,二,三,四,五,六,七,八,九", ",")(bytBit1) & "十"
        strBit2 = Split(",一,二,三,四,五,六,七,八,九", ",")(bytBit2)
    End If
    GetChineseNumber = strBit1 & strBit2
End Function
Private Sub DoFind()
Dim lngMatch As Long, rsTemp As ADODB.Recordset, lngCount As Long
    If Element Is Nothing Then Exit Sub
    lngMatch = Val(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0))
    If Trim(txtFind.Text) = "" Then Exit Sub
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 编码, 名称, 简码 From The (Select Cast(Zl_Dic_Search([1], [2], " & lngMatch & ") As " & gstrDbOwner & ".t_Dic_Rowset) From Dual)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "字典", Element.要素名称, Trim(txtFind.Text))
    Set vgdList.DataSource = rsTemp
    With vgdList
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftCenter
        Next
    End With
    If rsTemp.RecordCount = 0 Then
        cbrThis.Buttons(2).Enabled = False
        lblInfo = Element.要素名称 & " 没有匹配的项目"
        txtFind.SelStart = 0: txtFind.SelLength = 1000: If txtFind.Visible And txtFind.Enabled Then txtFind.SetFocus
    Else
        cbrThis.Buttons(2).Enabled = True
        lblInfo = Element.要素名称 & " 请选择希望的项目"
        If vgdList.Visible And vgdList.Enabled Then vgdList.SetFocus
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfg复选_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Element Is Nothing Then Exit Sub
    If Button = vbLeftButton Then
        If Element.输入形态 = 1 Then
            Dim i As Integer, strValue As String
            vfg复选.Cell(flexcpBackColor, 0, 0, 0, vfg复选.Cols - 1) = 0
            vfg复选.Cell(flexcpBackColor, 0, vfg复选.Col) = &HFFC0C0: vfg复选.Refresh
            
            If x > vfg复选.Cell(flexcpLeft, 0, vfg复选.Col) + 200 Then  '选中按扭周围点按AfterEdit事件处理
                If vfg复选.Cell(flexcpChecked, 0, vfg复选.Col) = flexUnchecked Then
                    vfg复选.Cell(flexcpChecked, 0, vfg复选.Col) = flexChecked
                Else
                    vfg复选.Cell(flexcpChecked, 0, vfg复选.Col) = flexUnchecked
                End If
                Call vfg复选_AfterEdit(0, vfg单选.Col)
            End If
        End If
    End If
End Sub

Private Sub vgdList_DblClick()
Dim i As Long, strReturn As String
    If Element Is Nothing Then Exit Sub
    If vgdList.Row <= 0 Then Exit Sub
    strReturn = ""
    For i = 0 To vgdList.Cols - 1
        strReturn = strReturn & ";" & vgdList.TextMatrix(vgdList.Row, i)
    Next
    If Len(strReturn) > 0 Then strReturn = Mid(strReturn, 2)
    If strReturn = ";;" Then Exit Sub
    Element.内容文本 = Split(strReturn, ";")(1)
    
    RaiseEvent pOk
End Sub

Private Sub vgdList_GotFocus()
    shpBorder1.BorderWidth = 2
End Sub

Private Sub vgdList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then vgdList_DblClick
End Sub

Private Sub vgdList_LostFocus()
    shpBorder1.BorderWidth = 1
End Sub


