VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBaby 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "婴儿选择"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5310
   Icon            =   "frmBaby.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5310
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfBaby 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3735
      _cx             =   6588
      _cy             =   4048
      Appearance      =   0
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBaby.frx":151A
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin VB.Label lbltip 
      AutoSize        =   -1  'True
      Caption         =   "请选择一个婴儿进行查看："
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2160
   End
End
Attribute VB_Name = "frmBaby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsData As ADODB.Recordset
Private mlngBaby As Long

Public Sub ShowMe(ByVal rsData As ADODB.Recordset, ByRef lngBaby As Long)
'传入的mrsData肯定是有数据的，故代码没有写无数据的逻辑
    Set mrsData = rsData
    
    Me.Show 1
    lngBaby = mlngBaby
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    mlngBaby = vsfBaby.TextMatrix(vsfBaby.Row, vsfBaby.ColIndex("序号"))
    Unload Me
End Sub

Private Sub Form_Load()
    
    mlngBaby = -1
    ShowData
End Sub

Private Sub ShowData()
    Dim i As Long
    
    With vsfBaby
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + mrsData.RecordCount
        i = .FixedRows
        While Not mrsData.EOF
            .TextMatrix(i, .ColIndex("序号")) = mrsData!序号
            .TextMatrix(i, .ColIndex("姓名")) = mrsData!婴儿姓名
            .TextMatrix(i, .ColIndex("性别")) = mrsData!婴儿性别
            .TextMatrix(i, .ColIndex("出生时间")) = mrsData!出生时间
            i = i + 1
            mrsData.MoveNext
        Wend
        .Row = 1
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfBaby_DblClick()
    Call cmdOK_Click
End Sub
