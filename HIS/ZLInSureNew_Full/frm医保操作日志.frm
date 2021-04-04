VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm医保操作日志 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保操作日志查看窗口"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14340
   Icon            =   "frm医保操作日志.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancle 
      Caption         =   "关闭(&C)"
      Height          =   345
      Left            =   12915
      TabIndex        =   2
      Top             =   8970
      Width           =   1200
   End
   Begin VB.TextBox txtDetail 
      Height          =   8520
      Left            =   6375
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   30
      Width           =   7935
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   8520
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Shift+Delete删除当前行"
      Top             =   30
      Width           =   6330
      _cx             =   11165
      _cy             =   15028
      Appearance      =   2
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm医保操作日志.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   15
      X2              =   15520
      Y1              =   8775
      Y2              =   8775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   0
      X2              =   15520
      Y1              =   8745
      Y2              =   8745
   End
End
Attribute VB_Name = "frm医保操作日志"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr模块                As String
Private mstr功能                As String
Private mstr主键1               As String
Private mstr主键2               As String
Private mstr主键3               As String
Private mstr主键4               As String
Private mrsDetail               As ADODB.Recordset
Dim strSql                      As String

Const strDetail = "Select Decode(类型, 0, '新增', 1, '修改', 2, '删除', 3, '状态') 类型, Decode(日志类型, 1, '说明', 2, '日志') As 日志类型, 日志编码, 日期, 用户, 工作站,数据来源,日志描述" & vbCrLf & _
                  "From 医保操作日志" & vbCrLf & _
                  "Where 模块 = [1] And 功能 = [2] And 主键1 = [3] And 主键2 = [4] And 主键3 = [5] And 主键4 = [6]" & vbCrLf & _
                  "Order By 日期"

Public Property Let str模块(ByVal vstr模块 As String)
    mstr模块 = vstr模块
End Property

Public Property Let str功能(ByVal vstr功能 As String)
    mstr功能 = vstr功能
End Property
 
Public Property Let str主键1(ByVal vstr主键1 As String)
    mstr主键1 = vstr主键1
End Property
 
Public Property Let str主键2(ByVal vstr主键2 As String)
    mstr主键2 = vstr主键2
End Property
 
Public Property Let str主键3(ByVal vstr主键3 As String)
    mstr主键3 = vstr主键3
End Property

Public Property Let str主键4(ByVal vstr主键4 As String)
    mstr主键4 = vstr主键4
End Property

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call mDataload
End Sub

Private Sub mDataload()
On Error GoTo ErrH
    strSql = strDetail
    If mstr主键2 = "" Then strSql = Replace(strSql, " And 主键2 = [4]", "")
    If mstr主键3 = "" Then strSql = Replace(strSql, " And 主键3 = [5]", "")
    If mstr主键4 = "" Then strSql = Replace(strSql, " And 主键4 = [6]", "")
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstr模块, mstr功能, mstr主键1, mstr主键2, mstr主键3, mstr主键4)
    Set vsfDetail.DataSource = mrsDetail
    If vsfDetail.Rows > 1 Then vsfDetail.Row = 1
    Call vsfDetail_RowColChange
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub vsfDetail_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    On Error GoTo ErrH
    Cancel = True
    Exit Sub
ErrH:
    Err.Clear
End Sub
 
Private Sub vsfDetail_CellChanged(ByVal Row As Long, ByVal COL As Long)
    Call vsfDetail_RowColChange
End Sub

Private Sub vsfDetail_Click()
    Call vsfDetail_RowColChange
End Sub

Private Sub vsfDetail_RowColChange()
On Error GoTo ErrH
    If vsfDetail.Row < 1 Or vsfDetail.COL < 1 Then
        txtDetail.Text = ""
    Else
        txtDetail.Text = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("日志描述"))
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub
