VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDrawCondition 
   Caption         =   "申购单导入"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9345
   Icon            =   "frmDrawCondition.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   9345
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   60
      ScaleWidth      =   9135
      TabIndex        =   10
      Top             =   2160
      Width           =   9135
   End
   Begin VB.ComboBox cboDate 
      Height          =   300
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8160
      TabIndex        =   6
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6840
      TabIndex        =   5
      Top             =   5280
      Width           =   1100
   End
   Begin VB.TextBox txtDept 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtRequestDept 
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfInfo 
      Height          =   2805
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   9135
      _cx             =   16113
      _cy             =   4948
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
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrawCondition.frx":000C
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1485
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   9135
      _cx             =   16113
      _cy             =   2619
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrawCondition.frx":0150
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
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      Height          =   180
      Left            =   6240
      TabIndex        =   7
      Top             =   180
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "被申购部门"
      Height          =   180
      Left            =   2880
      TabIndex        =   2
      Top             =   180
      Width           =   900
   End
   Begin VB.Label lblRequestDept 
      AutoSize        =   -1  'True
      Caption         =   "申购部门"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frmDrawCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngDept As Long    '被申购库房id
Private mlngRequest As Long    '申购部门id
Private mstrDept As String  '被申购库房
Private mstrRequest As String '申购部门
Private mintUint As Integer     '显示单位:0-散装单位,1-包装单位
Private mstrNO As String
Private mFMT As g_FmtString

Private Sub cboDate_Click()
    Call GetList
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim lngRow As Long
    
    mstrNO = ""
    With vsfList
        If .Rows > 1 Then
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, 0) = "-1" Then
                    mstrNO = mstrNO & .TextMatrix(lngRow, .ColIndex("no")) & ","
                End If
            Next
        End If
    End With
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitCbo
    
    With mFMT
        .FM_成本价 = GetFmtString(mintUint, g_成本价)
        .FM_金额 = GetFmtString(mintUint, g_金额)
        .FM_零售价 = GetFmtString(mintUint, g_售价)
        .FM_数量 = GetFmtString(mintUint, g_数量)
    End With
    
    txtRequestDept.Text = mstrRequest
    txtDept.Text = mstrDept
    mstrNO = ""
    
    Call GetList
End Sub

Public Function ShowMe(ByVal frmPara As Form, ByVal intUint As Integer, ByVal strDept As String, ByVal lngDept As Long, ByVal strRequest As String, ByVal lngRequest As Long) As String
    mintUint = intUint
    mlngDept = lngDept
    mlngRequest = lngRequest
    mstrDept = strDept
    mstrRequest = strRequest
    ShowMe = ""
    
    Me.Show vbModal, frmPara
    ShowMe = mstrNO
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    With vsfList
        .Move 50, txtRequestDept.Top + txtRequestDept.Height + 100, Me.ScaleWidth - 100, CLng(Me.Height / 4)
    End With
    
    With picSplit
        .Move 50, vsfList.Top + vsfList.Height + 20, vsfList.Width
    End With
    
    CmdCancel.Move Me.ScaleWidth - CmdCancel.Width - 100, Me.ScaleHeight - CmdCancel.Height - 50
    CmdSave.Move CmdCancel.Left - CmdSave.Width - 50, CmdCancel.Top
    
    With vsfInfo
        .Move 50, picSplit.Top + picSplit.Height + 20, picSplit.Width, CmdCancel.Top - vsfInfo.Top + 30
    End With
End Sub

Private Sub InitCbo()
    '初始化下拉列表
    With cboDate
        .AddItem "一星期内"
        .AddItem "一月内"
        .AddItem "三个月内"
        .AddItem "半年内"
        
        .ListIndex = 0
    End With
End Sub

Private Sub GetList()
    Dim rsTemp As ADODB.Recordset
    Dim datBeginDate As Date '开始日期
    Dim dateEndDate As Date '结束日期
    Dim datCurentDate As Date '当前日期
    
    Select Case cboDate.Text
        Case "一星期内"
            datBeginDate = CDate(Format(DateAdd("D", -7, Date), "yyyy-mm-dd") & " 00:00:00")
        Case "一月内"
            datBeginDate = CDate(Format(DateAdd("M", -1, Date), "yyyy-mm-dd") & " 00:00:00")
        Case "三个月内"
            datBeginDate = CDate(Format(DateAdd("M", -3, Date), "yyyy-mm-dd") & " 00:00:00")
        Case "半年内"
            datBeginDate = CDate(Format(DateAdd("M", -6, Date), "yyyy-mm-dd") & " 00:00:00")
    End Select
    dateEndDate = sys.Currentdate
    
    With vsfList
        .Rows = 1
        .ColDataType(0) = flexDTBoolean
        
        gstrSQL = "Select id,NO, 编制人, 编制日期, 审核人, 审核日期" & vbNewLine & _
                    "From 材料采购计划" & vbNewLine & _
                    "Where 库房id = [1] And 部门id = [2] And 审核日期 Between [3] And [4]" & vbNewLine & _
                    "Order By NO Desc, 审核日期 Desc"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "GetList", mlngDept, mlngRequest, datBeginDate, dateEndDate)
        
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("计划id")) = rsTemp!Id
            .TextMatrix(.Rows - 1, .ColIndex("no")) = rsTemp!NO
            .TextMatrix(.Rows - 1, .ColIndex("编制人")) = rsTemp!编制人
            .TextMatrix(.Rows - 1, .ColIndex("编制日期")) = rsTemp!编制日期
            .TextMatrix(.Rows - 1, .ColIndex("审核人")) = rsTemp!审核人
            .TextMatrix(.Rows - 1, .ColIndex("审核日期")) = rsTemp!审核日期
            
            rsTemp.MoveNext
        Loop
    End With
End Sub

Private Sub GetDetails(ByVal lngID As Long)
    Dim rsTemp As ADODB.Recordset
    
    With vsfInfo
        gstrSQL = "Select a.Id, a.No,'[' || d.编码 || ']' || d.名称 || '-' || d.规格 As 编名称, b.材料id, a.审核人, a.审核日期, b.计划数量, b.单价 As 成本价, b.上次供应商, b.上次生产商, d.计算单位, c.包装单位, c.换算系数" & vbNewLine & _
                    "From 材料采购计划 A, 材料计划内容 B, 材料特性 C, 收费项目目录 D" & vbNewLine & _
                    "Where a.Id = b.计划id And b.材料id = c.材料id And c.材料id = d.Id And d.类别 = '4' And a.id=[1]" & vbNewLine & _
                    "Order By NO Desc, 审核日期 Desc"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "InitCrad", lngID)
        
        If rsTemp.RecordCount > 0 Then
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("计划id")) = rsTemp!Id
                .TextMatrix(.Rows - 1, .ColIndex("no")) = rsTemp!NO
                .TextMatrix(.Rows - 1, .ColIndex("材料名称")) = rsTemp!编名称
                .TextMatrix(.Rows - 1, .ColIndex("材料id")) = rsTemp!材料ID
                .TextMatrix(.Rows - 1, .ColIndex("上次供应商")) = IIf(IsNull(rsTemp!上次供应商), "", rsTemp!上次供应商)
                .TextMatrix(.Rows - 1, .ColIndex("上次生产商")) = IIf(IsNull(rsTemp!上次生产商), "", rsTemp!上次生产商)
                
                If mintUint = 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("单位")) = rsTemp!计算单位
                    .TextMatrix(.Rows - 1, .ColIndex("计划数量")) = Format(rsTemp!计划数量, mFMT.FM_数量)
                    .TextMatrix(.Rows - 1, .ColIndex("成本价")) = Format(rsTemp!成本价, mFMT.FM_成本价)
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("单位")) = rsTemp!包装单位
                    .TextMatrix(.Rows - 1, .ColIndex("计划数量")) = Format(rsTemp!计划数量 / rsTemp!换算系数, mFMT.FM_数量)
                    .TextMatrix(.Rows - 1, .ColIndex("成本价")) = Format(rsTemp!成本价 * rsTemp!换算系数, mFMT.FM_成本价)
                End If
                
                rsTemp.MoveNext
            Loop
        End If
    End With
End Sub

Private Sub DeleteDetails(ByVal lngID As Long)
    '清除不要的单据
    Dim lngRow As Long
    
    With vsfInfo
        For lngRow = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(lngRow, .ColIndex("计划id"))) = lngID Then
                .RemoveItem lngRow
            End If
        Next
    End With
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit
        If .Top + Y < 2000 Then Exit Sub
        If .Top + Y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + Y
    End With

    With vsfList
        .Height = picSplit.Top - .Top
    End With
    
    With vsfInfo
        .Top = picSplit.Top + picSplit.Height + 100
        .Height = ScaleHeight - .Top - CmdSave.Height - 50
    End With
    Me.Refresh
End Sub

Private Sub vsfList_DblClick()
    Dim strTemp As String
    Dim lngRow As Long
    Dim blnTemp As Boolean
    Dim lngID As Long
    
    With vsfList
        If .Col = 0 And .Row >= 1 Then
            strTemp = .TextMatrix(.Row, .ColIndex("no"))
            lngID = Val(.TextMatrix(.Row, .ColIndex("计划id")))
            
            If .TextMatrix(.Row, 0) = "-1" Then
                .TextMatrix(.Row, 0) = ""
                blnTemp = False
            Else
                .TextMatrix(.Row, 0) = "-1"
                blnTemp = True
            End If
            
            If blnTemp = True Then
                Call GetDetails(lngID)
            Else
                Call DeleteDetails(lngID)
            End If
        End If
    End With
End Sub

