VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMediContrast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置中选对照药品"
   ClientHeight    =   8415
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "frmMediContrast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic提示 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   720
      ScaleHeight     =   615
      ScaleWidth      =   9255
      TabIndex        =   9
      Top             =   120
      Width           =   9255
      Begin VB.Label lblnote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMediContrast.frx":6852
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   8925
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10335
      TabIndex        =   4
      Top             =   7800
      Width           =   10335
      Begin VB.CommandButton cmdOk 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   7560
         TabIndex        =   7
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   8760
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   5
         Top             =   150
         Width           =   1365
      End
      Begin VB.Label lblFind 
         BackColor       =   &H80000003&
         Caption         =   "查找"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9840
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   9975
      _cx             =   17595
      _cy             =   5741
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediContrast.frx":6914
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
   Begin VSFlex8Ctl.VSFlexGrid vsf对照 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   9975
      _cx             =   17595
      _cy             =   5741
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediContrast.frx":6A4F
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
   Begin VSFlex8Ctl.VSFlexGrid vsf数据汇总 
      Height          =   2055
      Left            =   10800
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   5775
      _cx             =   10186
      _cy             =   3625
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediContrast.frx":6B44
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmMediContrast.frx":6C6D
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMediContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long '记录选择录入后的药品ID
Private mlng中选药品ID As Long
Private mlng对照药品ID As Long
Private Const mlngBorderColor As Long = &H0&    '选中行边框颜色
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' 没选中行边框颜色
Private Const mcstEditColor = &H80000003   '能编辑的颜色
Private Const mcstBachColor = &H80000008          '设置了对照药品的颜色
Private Const mcstNotBachColor = &H8080FF      '没有设置对照药品的颜色
Private mrsFindName As ADODB.Recordset '查询的数据集
Public Sub ShowMe(ByVal objFra As frmMediLists)
    Me.Show vbModal, objFra
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Long, n As Long
    Dim int序号 As Integer
    Dim str对照数据 As String
    
    If vsfList.Rows < 2 Then Exit Sub
    
    With vsf数据汇总
        For i = 1 To vsfList.Rows - 1
            int序号 = 0
            For n = 1 To .Rows - 1
                If Val(vsfList.TextMatrix(i, vsfList.ColIndex("药品ID"))) = Val(.TextMatrix(n, .ColIndex("中选药品ID"))) Then
                    int序号 = int序号 + 1
                    str对照数据 = int序号 & "^" & .TextMatrix(n, .ColIndex("中选药品ID")) & _
                         "^" & .TextMatrix(n, .ColIndex("对照药品ID")) & "|" & str对照数据
                End If
            Next
        Next
    End With

    gstrSql = "Zl_中选药品对照_Update('" & str对照数据 & "')"
    
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Call vsfList_EnterCell
    Call EditColor
    
    MsgBox "保存成功！", vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    Call IniGrid
    Call FillVSF
    Call ShowData
    Call EditColor
End Sub

Private Sub IniGrid()
    With vsfList
        .Editable = flexEDNone
        .Rows = 1
        .ColWidth(0) = 350
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = 400
        .AllowSelection = False '不能多选
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
    End With

    With vsf对照
        .Editable = flexEDNone
        .Rows = 2
        .ColWidth(0) = 350
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = 400
        .AllowSelection = False '不能多选
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .Cell(flexcpBackColor, 1, .ColIndex("对照药品"), 1, .ColIndex("对照药品")) = mcstEditColor
    End With
End Sub

Private Sub FillVSF()
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSql = "Select a.药品id, b.编码, b.名称, b.规格, b.产地 As 生产商, c.名称 As 供应商,n.名称 As 商品名" & vbNewLine & _
                    "From 药品规格 A, 收费项目目录 B, 供应商 C, 收费项目别名 N" & vbNewLine & _
                    "Where a.药品id = b.Id And a.上次供应商id = c.Id(+) And b.Id = n.收费细目id(+)" & vbNewLine & _
                    "      And n.码类(+) = 1 And n.性质(+) = 3 And a.是否带量采购 = 1" & vbNewLine & _
                    "Order By b.编码"


    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "否带量采购药品")
    
    With vsfList
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = .Rows - 1
            .TextMatrix(.Rows - 1, .ColIndex("药品ID")) = rsTemp!药品ID
            .TextMatrix(.Rows - 1, .ColIndex("中选药品")) = "[" & rsTemp!编码 & "]" & rsTemp!名称
            .TextMatrix(.Rows - 1, .ColIndex("商品名")) = NVL(rsTemp!商品名)
            .TextMatrix(.Rows - 1, .ColIndex("规格")) = rsTemp!规格
            .TextMatrix(.Rows - 1, .ColIndex("生产商")) = NVL(rsTemp!生产商)
            .TextMatrix(.Rows - 1, .ColIndex("供应商")) = NVL(rsTemp!供应商)
            
            rsTemp.MoveNext
        Loop
        
    End With
    
    Call VsfRowHeight(vsfList)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowData()
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    vsf数据汇总.Rows = 1
    
    For i = 1 To vsfList.Rows - 1

            gstrSql = "Select b.编码, b.名称, b.规格, b.产地 As 生产商, c.名称 As 供应商, d.序号, d.中选药品id, d.对照药品id, n.名称 As 商品名" & vbNewLine & _
                            "From 药品规格 A, 收费项目目录 B, 供应商 C, 中选药品对照 D, 收费项目别名 N" & vbNewLine & _
                            "Where a.药品id = b.Id And a.上次供应商id = c.Id(+) And b.Id = d.对照药品id" & vbNewLine & _
                            "      And b.Id = n.收费细目id(+) And n.码类(+) = 1 And n.性质(+) = 3 And d.中选药品id = [1]" & vbNewLine & _
                            "Order By d.序号"

            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "对照药品", Val(vsfList.TextMatrix(i, vsfList.ColIndex("药品ID"))))
            
            With vsf数据汇总
        
                Do While Not rsTemp.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("序号")) = rsTemp!序号
                    .TextMatrix(.Rows - 1, .ColIndex("中选药品ID")) = rsTemp!中选药品ID
                    .TextMatrix(.Rows - 1, .ColIndex("对照药品ID")) = rsTemp!对照药品ID
                    .TextMatrix(.Rows - 1, .ColIndex("对照药品")) = "[" & rsTemp!编码 & "]" & rsTemp!名称
                    .TextMatrix(.Rows - 1, .ColIndex("商品名")) = NVL(rsTemp!商品名)
                    .TextMatrix(.Rows - 1, .ColIndex("规格")) = rsTemp!规格
                    .TextMatrix(.Rows - 1, .ColIndex("生产商")) = NVL(rsTemp!生产商)
                    .TextMatrix(.Rows - 1, .ColIndex("供应商")) = NVL(rsTemp!供应商)
                    
                    rsTemp.MoveNext
                Loop
        
            End With
        Next
        
        If vsfList.Rows > 1 Then vsfList.Row = 1

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng对照药品ID = 0
    mlng中选药品ID = 0
    mlngID = 0
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub

    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub vsf对照_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsf对照
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("供应商") Then
                .Col = .Col + 1
            ElseIf .Row <> .Rows - 1 And .Col = .ColIndex("供应商") Then
                .Row = .Row + 1
                .Col = .ColIndex("对照药品")
            ElseIf .Row = .Rows - 1 And .TextMatrix(.Row, 1) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .Row = .Rows - 1
                .Col = .ColIndex("对照药品")
            End If
        ElseIf KeyCode = vbKeyDelete Then
            Call Delete
        End If
    End With
End Sub

Private Sub vsf对照_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then Exit Sub
    If Col = vsf对照.ColIndex("对照药品") Then
        If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub vsf对照_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrHand
    
    vRect = zlControl.GetControlRect(vsf对照.hwnd) '获取位置
    dblLeft = vRect.Left + vsf对照.CellLeft
    dblTop = vRect.Top + vsf对照.CellTop + vsf对照.CellHeight + 3300
    With vsf对照
        mlng对照药品ID = Val(.TextMatrix(.Row, .ColIndex("药品ID")))
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Col = .ColIndex("对照药品") And .EditText = "" Then Exit Sub
        If Col = .ColIndex("对照药品") And InStr(1, .EditText, "[") = 0 Then
        gstrSql = "Select Distinct i.Id, i.编码, i.名称, i.规格, i.产地 As 生产商, c.名称 As 供应商, m.商品名" & vbNewLine & _
                        "From 收费项目目录 I, 收费项目别名 N, 药品规格 A, 供应商 C, (Select 收费细目id, 名称 As 商品名 From 收费项目别名 Where 码类 = 1 And 性质 = 3) M" & vbNewLine & _
                        "Where i.Id = n.收费细目id And i.Id = a.药品id And a.上次供应商id = c.Id(+) And i.Id = m.收费细目id(+) And i.类别 In ('5', '6') And" & vbNewLine & _
                        "      (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                        "      (i.编码 Like [1] Or n.名称 Like [2] Or n.简码 Like [2]) And Nvl(a.是否带量采购, 0) = 0" & vbNewLine & _
                        "Order By i.编码"

            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "对照药品", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True, UCase(.EditText) & "%", gstrMatch & UCase(.EditText) & "%")
            
            If blnCancel = True Then
                Exit Sub
            End If
  
            If rsRecord Is Nothing Then
                MsgBox "没有找到该对照药品！", vbInformation, gstrSysName
                Exit Sub
            Else
                mlngID = rsRecord!ID
                If CheckDub = False Then
                    .EditText = "[" & rsRecord!编码 & "]" & rsRecord!名称
                    .TextMatrix(.Row, .ColIndex("序号")) = .Row
                    .TextMatrix(.Row, .ColIndex("药品ID")) = rsRecord!ID
                    .TextMatrix(.Row, .ColIndex("对照药品")) = "[" & rsRecord!编码 & "]" & rsRecord!名称
                    .TextMatrix(.Row, .ColIndex("商品名")) = NVL(rsRecord!商品名)
                    .TextMatrix(.Row, .ColIndex("规格")) = rsRecord!规格
                    .TextMatrix(.Row, .ColIndex("生产商")) = NVL(rsRecord!生产商)
                    .TextMatrix(.Row, .ColIndex("供应商")) = NVL(rsRecord!供应商)
                    
                    Call UpDate
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Cell(flexcpBackColor, .Row, .ColIndex("对照药品"), .Rows - 1, .ColIndex("对照药品")) = mcstEditColor
                    Call VsfRowHeight(vsf对照)
                Else
                    MsgBox "已经有该药品！", vbInformation, gstrSysName
                End If
            End If
            
        End If
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf对照_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsf对照
        If .Col = .ColIndex("对照药品") Then
            .ColComboList(.ColIndex("对照药品")) = "|..."
        Else
            .ColComboList(.ColIndex("对照药品")) = ""
        End If
    End With
End Sub

Private Sub vsf对照_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsf对照
        .EditSelStart = 0
        .EditSelLength = zlcommfun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsf对照_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsf对照
        .EditMaxLength = 50
    End With
End Sub


Private Sub vsf对照_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = zlControl.GetControlRect(vsf对照.hwnd) '获取位置
    dblLeft = vRect.Left + vsf对照.CellLeft
    dblTop = vRect.Top + vsf对照.CellTop + vsf对照.CellHeight + 3300
    With vsf对照
        mlng对照药品ID = Val(.TextMatrix(.Row, .ColIndex("药品ID")))
        If Col = .ColIndex("对照药品") Then
            gstrSql = "Select i.Id, i.编码, i.名称, i.规格, i.产地 As 生产商, c.名称 As 供应商, n.名称 As 商品名" & vbNewLine & _
                            "From 收费项目目录 I, 药品规格 A, 供应商 C, 收费项目别名 N" & vbNewLine & _
                            "Where i.Id = a.药品id And a.上次供应商id = c.Id(+) And i.类别 In ('5', '6') And i.Id = n.收费细目id(+)" & vbNewLine & _
                            "      And n.码类(+) = 1 And n.性质(+) = 3 And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                            "      And Nvl(a.是否带量采购, 0) = 0 Order By i.编码"

            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "对照药品", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True)

            If rsRecord Is Nothing Then
                Exit Sub
            Else
                mlngID = rsRecord!ID
                If CheckDub = False Then
                    .TextMatrix(.Row, .ColIndex("序号")) = .Row
                    .TextMatrix(.Row, .ColIndex("药品ID")) = rsRecord!ID
                    .TextMatrix(.Row, .ColIndex("对照药品")) = "[" & rsRecord!编码 & "]" & rsRecord!名称
                    .TextMatrix(.Row, .ColIndex("商品名")) = NVL(rsRecord!商品名)
                    .TextMatrix(.Row, .ColIndex("规格")) = rsRecord!规格
                    .TextMatrix(.Row, .ColIndex("生产商")) = NVL(rsRecord!生产商)
                    .TextMatrix(.Row, .ColIndex("供应商")) = NVL(rsRecord!供应商)
                    
                    Call UpDate
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Cell(flexcpBackColor, .Row, .ColIndex("对照药品"), .Rows - 1, .ColIndex("对照药品")) = mcstEditColor
                    Call VsfRowHeight(vsf对照)
                Else
                    MsgBox "已经有该药品！", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
    
End Sub

Private Sub vsf对照_EnterCell()
    
    With vsf对照
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstEditColor Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
        mlng对照药品ID = Val(.TextMatrix(.Row, .ColIndex("药品ID")))
    End With
End Sub

Private Sub vsfList_EnterCell()
    Dim i As Integer
    
    With vsfList
        If Val(.TextMatrix(.Row, .ColIndex("药品ID"))) = 0 Then Exit Sub
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                .CellBorderRange i, 0, i, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            .CellBorderRange .Row, 0, .Row, .Cols - 1, mlngBorderColor, 0, 2, 0, 2, 0, 2
        End If
        mlng中选药品ID = Val(.TextMatrix(.Row, .ColIndex("药品ID")))
        Call ShowGrid(Val(.TextMatrix(.Row, .ColIndex("药品ID"))))
    End With
    
End Sub

Private Function CheckDub() As Boolean
    '检查是否存在该中标单位或是否存在药品
    Dim i As Integer

    With vsf数据汇总
        For i = 1 To .Rows - 1
            If Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("药品ID"))) = Val(.TextMatrix(i, .ColIndex("中选药品ID"))) And _
                .TextMatrix(i, .ColIndex("对照药品id")) = mlngID Then
                CheckDub = True
                Exit Function
            End If
        Next
    End With
    CheckDub = False

End Function

Private Sub Delete()
    Dim i As Integer
    
    With vsf数据汇总
        For i = 1 To .Rows - 1
            If Val(vsf对照.TextMatrix(vsf对照.Row, vsf对照.ColIndex("药品ID"))) = Val(.TextMatrix(i, .ColIndex("对照药品ID"))) And _
                Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("药品ID"))) = Val(.TextMatrix(i, .ColIndex("中选药品ID"))) Then
                .RemoveItem i
                Exit For
            End If
        Next
    End With
    
    With vsf对照
        If .Rows = 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("药品ID"))) = 0 Then Exit Sub
        
        If .Rows - 1 = 1 Then
            For i = 1 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
        Else
            .RemoveItem .Row
            Call vsf_ResetSerial
        End If
    End With
    
End Sub

Private Sub vsf_ResetSerial()
    Dim i As Integer
    With vsf对照
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("序号")) = i
        Next
    End With
End Sub

Private Sub UpDate()
    Dim i As Integer
    
    With vsf数据汇总
        If vsf对照.Row = vsf对照.Rows - 1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("中选药品ID")) = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("药品ID"))
            .TextMatrix(.Rows - 1, .ColIndex("对照药品ID")) = vsf对照.TextMatrix(vsf对照.Rows - 1, vsf对照.ColIndex("药品ID"))
            .TextMatrix(.Rows - 1, .ColIndex("对照药品")) = vsf对照.TextMatrix(vsf对照.Rows - 1, vsf对照.ColIndex("对照药品"))
            .TextMatrix(.Rows - 1, .ColIndex("商品名")) = vsf对照.TextMatrix(vsf对照.Rows - 1, vsf对照.ColIndex("商品名"))
            .TextMatrix(.Rows - 1, .ColIndex("规格")) = vsf对照.TextMatrix(vsf对照.Rows - 1, vsf对照.ColIndex("规格"))
            .TextMatrix(.Rows - 1, .ColIndex("生产商")) = vsf对照.TextMatrix(vsf对照.Rows - 1, vsf对照.ColIndex("生产商"))
            .TextMatrix(.Rows - 1, .ColIndex("供应商")) = vsf对照.TextMatrix(vsf对照.Rows - 1, vsf对照.ColIndex("供应商"))
        Else
            For i = 1 To .Rows - 1
                If mlng对照药品ID = Val(.TextMatrix(i, .ColIndex("对照药品ID"))) And _
                    mlng中选药品ID = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("药品ID"))) Then
                        .TextMatrix(i, .ColIndex("中选药品ID")) = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("药品ID"))
                        .TextMatrix(i, .ColIndex("对照药品ID")) = vsf对照.TextMatrix(vsf对照.Row, vsf对照.ColIndex("药品ID"))
                        .TextMatrix(i, .ColIndex("对照药品")) = vsf对照.TextMatrix(vsf对照.Row, vsf对照.ColIndex("对照药品"))
                        .TextMatrix(i, .ColIndex("商品名")) = vsf对照.TextMatrix(vsf对照.Row, vsf对照.ColIndex("商品名"))
                        .TextMatrix(i, .ColIndex("规格")) = vsf对照.TextMatrix(vsf对照.Row, vsf对照.ColIndex("规格"))
                        .TextMatrix(i, .ColIndex("生产商")) = vsf对照.TextMatrix(vsf对照.Row, vsf对照.ColIndex("生产商"))
                        .TextMatrix(i, .ColIndex("供应商")) = vsf对照.TextMatrix(vsf对照.Row, vsf对照.ColIndex("供应商"))
                    Exit For
                End If
            Next
        End If
    End With
End Sub


Private Sub ShowGrid(ByVal lng药品ID As Long)
    Dim n As Long
    
    With vsf对照
        .Rows = 1
        .Rows = 2
        .Cell(flexcpBackColor, 1, .ColIndex("对照药品"), 1, .ColIndex("对照药品")) = mcstEditColor
        For n = 1 To vsf数据汇总.Rows - 1
            If lng药品ID = Val(vsf数据汇总.TextMatrix(n, vsf数据汇总.ColIndex("中选药品ID"))) Then
                .TextMatrix(.Rows - 1, .ColIndex("序号")) = .Rows - 1
                .TextMatrix(.Rows - 1, .ColIndex("药品ID")) = vsf数据汇总.TextMatrix(n, vsf数据汇总.ColIndex("对照药品ID"))
                .TextMatrix(.Rows - 1, .ColIndex("对照药品")) = vsf数据汇总.TextMatrix(n, vsf数据汇总.ColIndex("对照药品"))
                .TextMatrix(.Rows - 1, .ColIndex("商品名")) = vsf数据汇总.TextMatrix(n, vsf数据汇总.ColIndex("商品名"))
                .TextMatrix(.Rows - 1, .ColIndex("规格")) = vsf数据汇总.TextMatrix(n, vsf数据汇总.ColIndex("规格"))
                .TextMatrix(.Rows - 1, .ColIndex("生产商")) = vsf数据汇总.TextMatrix(n, vsf数据汇总.ColIndex("生产商"))
                .TextMatrix(.Rows - 1, .ColIndex("供应商")) = vsf数据汇总.TextMatrix(n, vsf数据汇总.ColIndex("供应商"))
                
                .Rows = .Rows + 1
                .Cell(flexcpBackColor, 1, .ColIndex("对照药品"), .Rows - 1, .ColIndex("对照药品")) = mcstEditColor
            End If
        Next
    End With
    Call VsfRowHeight(vsf对照)
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim str药名 As String
    Dim lngRow As Long

    '查找药品
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '表示新的查找
        txtFind.Tag = strInput

        gstrSql = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B,药品规格 C " & _
                  "Where a.id=c.药品ID And A.Id =B.收费细目id And A.类别 In ('5','6') " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) and c.是否带量采购=1 " & _
                  "Order By 药品编码 "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSql, "取匹配的药品ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        str药名 = mrsFindName!药品编码 & mrsFindName!通用名

        For lngRow = 1 To vsfList.Rows - 1
            lngFindRow = vsfList.FindRow(str药名, lngRow, CLng(vsfList.ColIndex("中选药品")), True, True)
            If lngFindRow > 0 Then
                vsfList.Row = lngFindRow
                vsfList.TopRow = lngFindRow
                Exit For
            End If
        Next

        If lngFindRow > 0 Then  '查询到数据后就移动下下一条并退出本次查询
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext '未查询到数据则移动到下一条数据集继续查询
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub EditColor()
    Dim i As Long, n As Long
    Dim bln是否设置对照 As Boolean
    
    With vsfList
        For i = 1 To .Rows - 1
            bln是否设置对照 = False
            For n = 1 To vsf数据汇总.Rows - 1
                If Val(.TextMatrix(i, .ColIndex("药品ID"))) = Val(vsf数据汇总.TextMatrix(n, vsf数据汇总.ColIndex("中选药品ID"))) Then
                    bln是否设置对照 = True
                    Exit For
                End If
            Next
            If bln是否设置对照 Then
                .Cell(flexcpForeColor, i, .ColIndex("序号"), i, .ColIndex("供应商")) = mcstBachColor
            Else
                .Cell(flexcpForeColor, i, .ColIndex("序号"), i, .ColIndex("供应商")) = mcstNotBachColor
            End If
        Next
    End With
End Sub

Private Sub VsfRowHeight(ByVal VsfObj As VSFlexGrid)
    Dim i As Long
    With VsfObj
        For i = 1 To .Rows - 1
            .RowHeight(i) = 350
        Next
    End With
End Sub
