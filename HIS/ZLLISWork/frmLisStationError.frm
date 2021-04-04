VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLisStationError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "审核结果报告"
   ClientHeight    =   4500
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7095
   Icon            =   "frmLisStationError.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "忽略(&I)"
      Height          =   350
      Left            =   4650
      TabIndex        =   2
      Top             =   4005
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5850
      TabIndex        =   1
      Top             =   4005
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   3075
      Left            =   600
      TabIndex        =   0
      Top             =   795
      Width           =   6465
      _cx             =   11404
      _cy             =   5424
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
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
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "审核报告结果，即病人的两次检验结果在异常范围之内。"
      Height          =   180
      Left            =   720
      TabIndex        =   3
      Top             =   330
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   45
      Picture         =   "frmLisStationError.frx":000C
      Top             =   210
      Width           =   480
   End
End
Attribute VB_Name = "frmLisStationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnIgnore As Boolean

Public Function ShowError(ByVal frmMain As Object, ByVal rs As ADODB.Recordset) As Boolean
    
    mblnIgnore = False
    
    vsf.Cols = 0
    NewColumn vsf, "", 255, 4
    NewColumn vsf, "标本时间", 1080, 1
    NewColumn vsf, "标本序号", 900, 1
    NewColumn vsf, "检验项目", 2400, 1
    NewColumn vsf, "本次结果", 810, 1
    NewColumn vsf, "上次结果", 810, 1
    vsf.FixedCols = 1
    
    vsf.ExtendLastCol = True
    
    vsf.Rows = 2
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
 
    If rs.BOF = False Then
        vsf.TextMatrix(0, 0) = "序号"
        Call FillGrid(vsf, rs)
        vsf.TextMatrix(0, 0) = ""
    Else
        Exit Function
    End If
        
    Me.Show 1, frmMain
    
    ShowError = mblnIgnore
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mblnIgnore = True
    Unload Me
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col = 0)
End Sub

