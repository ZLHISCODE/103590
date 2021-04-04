VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmMediDosage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "配方原料"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   Icon            =   "frmMediDosage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSpilt 
      BackColor       =   &H80000005&
      Height          =   4455
      Left            =   3480
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4455
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   360
      Width           =   15
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfVariety 
      Height          =   3165
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3015
      _cx             =   5318
      _cy             =   5583
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediDosage.frx":6852
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
   Begin VSFlex8Ctl.VSFlexGrid vsfSpec 
      Height          =   3165
      Left            =   4680
      TabIndex        =   4
      Top             =   600
      Width           =   3015
      _cx             =   5318
      _cy             =   5583
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediDosage.frx":68C7
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
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "以下是规格"
      Height          =   180
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblVar 
      AutoSize        =   -1  'True
      Caption         =   "以下是品种"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmMediDosage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintDosageType As Integer '配方类型
Private mstrName As String  '选取的药品
Private mstrFind As String  '条件查询

'规格
Private Enum menuSpec
    ID = 0
    编码 = 1
    名称 = 2
    规格 = 3
    计算单位 = 4
    产地 = 5
    是否变价 = 6
    费用类型 = 7
    服务对象 = 8
    药名ID = 9
    Cols = 10
End Enum

'品种
Private Enum menuVar
    ID = 0
    编码 = 1
    名称 = 2
    计算单位 = 3
    服务对象 = 4
    Cols = 5
End Enum

Public Sub ShowMe(ByVal intDosageType As Integer, ByVal frmPar As Form, ByVal strFind As String, ByRef strName As String)
    Select Case intDosageType
        Case 0
            mintDosageType = 3  '忽略形态
        Case 1
            mintDosageType = 0  '散装
        Case 2
            mintDosageType = 1  '饮片
        Case 3
            mintDosageType = 2  '免煎剂
    End Select
    
    mstrFind = strFind
    mstrName = ""
    
    Me.Show vbModal, frmPar
    strName = mstrName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub GetVarInfo()
    '得到品种信息
    Dim rsTemp As Recordset
    Dim intRow As Integer
    
    On Error GoTo ErrHand
    
    If mintDosageType <> 3 Then '忽略形态
        gstrSql = " b.中药形态= " & mintDosageType & " and "
    Else
        gstrSql = ""
    End If
    
    If Trim(mstrFind) = "" Then
        gstrSql = "Select a.Id, a.编码, a.名称,a.计算单位, Decode(a.服务对象, 1, '门诊', 2, '住院', 3, '门诊和住院', 4, '体检', '不服务于病人') 服务对象" & vbNewLine & _
            "From 诊疗项目目录 A" & vbNewLine & _
            "Where Exists (Select 1 From 药品规格 B Where " & gstrSql & " a.Id = b.药名id) And a.类别 = '7' And Sysdate Between a.建档时间 And a.撤档时间"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "getVarInfo")
    Else
        gstrSql = "Select Distinct a.Id, a.编码, a.名称, a.计算单位, Decode(a.服务对象, 1, '门诊', 2, '住院', 3, '门诊和住院', 4, '体检', '不服务于病人') 服务对象" & vbNewLine & _
            "From 诊疗项目目录 A, 诊疗项目别名 N" & vbNewLine & _
            "Where Exists (Select 1 From 药品规格 B Where " & gstrSql & " a.Id = b.药名id) And a.Id = n.诊疗项目id And a.类别 = '7' And" & vbNewLine & _
            "      (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "      (a.编码 Like [1] Or n.名称 Like [2] Or n.简码 Like [2])" & vbNewLine & _
            "Order By a.编码"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "getVarInfo", mstrFind & "%", gstrMatch & mstrFind & "%")
    End If
    
    intRow = 1
    vsfVariety.Rows = rsTemp.RecordCount + 1
    Do While Not rsTemp.EOF
        With vsfVariety
            .TextMatrix(intRow, menuVar.ID) = rsTemp!ID
            .TextMatrix(intRow, menuVar.编码) = rsTemp!编码
            .TextMatrix(intRow, menuVar.名称) = IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
            .TextMatrix(intRow, menuVar.计算单位) = IIf(IsNull(rsTemp!计算单位), "", rsTemp!计算单位)
            .TextMatrix(intRow, menuVar.服务对象) = rsTemp!服务对象
            intRow = intRow + 1
            rsTemp.MoveNext
        End With
    Loop
    
    vsfVariety.Cell(flexcpAlignment, 0, 0, 0, vsfVariety.Cols - 1) = flexAlignCenterCenter
    If vsfVariety.Rows > 1 Then
        vsfVariety.Cell(flexcpAlignment, 1, 0, vsfVariety.Rows - 1, vsfVariety.Cols - 1) = flexAlignLeftCenter
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetSpecInfo()
    '得到规格信息
    Dim rsTemp As Recordset
    Dim intRow As Integer
    
    On Error GoTo ErrHand
    
    vsfSpec.Rows = 1
    
    If mintDosageType <> 3 Then '忽略形态
        gstrSql = " and b.中药形态=" & mintDosageType
    End If
    
    gstrSql = "Select a.Id, a.编码, a.名称, a.规格,a.计算单位, a.产地, Decode(a.是否变价, 0, '定价', '时价') As 是否变价, a.费用类型," & vbNewLine & _
        "       Decode(a.服务对象, 1, '门诊', 2, '住院', 3, '门诊和住院', '不服务于病人') 服务对象, b.药名id " & vbNewLine & _
        "From 收费项目目录 A, 药品规格 B" & vbNewLine & _
        "Where a.Id = b.药品id And  b.药名id = [1]" & gstrSql & " and a.类别 = '7' And Sysdate Between a.建档时间 And a.撤档时间" & vbNewLine & _
        "Order By a.Id"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "getVarInfo", vsfVariety.TextMatrix(vsfVariety.Row, menuVar.ID))
    
    intRow = 1
    vsfSpec.Rows = rsTemp.RecordCount + 1
    Do While Not rsTemp.EOF
        With vsfSpec
            .TextMatrix(intRow, menuSpec.ID) = rsTemp!ID
            .TextMatrix(intRow, menuSpec.编码) = rsTemp!编码
            .TextMatrix(intRow, menuSpec.名称) = IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
            .TextMatrix(intRow, menuSpec.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
            .TextMatrix(intRow, menuSpec.计算单位) = IIf(IsNull(rsTemp!计算单位), "", rsTemp!计算单位)
            .TextMatrix(intRow, menuSpec.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
            .TextMatrix(intRow, menuSpec.是否变价) = rsTemp!是否变价
            .TextMatrix(intRow, menuSpec.费用类型) = IIf(IsNull(rsTemp!费用类型), "", rsTemp!费用类型)
            .TextMatrix(intRow, menuSpec.服务对象) = rsTemp!服务对象
            .TextMatrix(intRow, menuSpec.药名ID) = rsTemp!药名ID
            intRow = intRow + 1
            
            rsTemp.MoveNext
        End With
    Loop
    vsfSpec.Cell(flexcpAlignment, 0, 0, 0, vsfSpec.Cols - 1) = flexAlignCenterCenter
    If vsfSpec.Rows > 1 Then
        vsfSpec.Cell(flexcpAlignment, 1, 0, vsfSpec.Rows - 1, vsfSpec.Cols - 1) = flexAlignLeftCenter
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call initControl    '初始化控件大小
    Call InitData   '初始化vsf控件表头、颜色
    Call GetVarInfo '为品种列表赋值
    Call SetCaption '设置标题
End Sub

Private Sub SetCaption()
    '设置标题
    Dim strCaption As String
    
    Select Case mintDosageType
        Case 0
            strCaption = "配方原料(散装)"
        Case 1
            strCaption = "配方原料(饮片)"
        Case 2
            strCaption = "配方原料(免煎剂)"
        Case 3
            strCaption = "配方原料(忽略形态)"
    End Select
    Me.Caption = strCaption
End Sub

Private Sub initControl()
    '初始化控件位置和状态
    Select Case mintDosageType
    Case 0 ' 散装
        lblVar.Left = 50
        vsfVariety.Top = lblVar.Height + lblVar.Top + 100
        vsfVariety.Width = Me.Width / 3 - picSpilt.Width
        vsfVariety.Height = Me.Height - lblVar.Height - lblVar.Top - 500
        vsfVariety.Left = 50
        picSpilt.Left = vsfVariety.Left + vsfVariety.Width
        picSpilt.Top = 0
        picSpilt.Height = Me.Height
        vsfSpec.Width = Me.Width - picSpilt.Width - picSpilt.Left - 100
        vsfSpec.Top = vsfVariety.Top
        vsfSpec.Height = Me.Height - lblVar.Height - lblVar.Top - 500
        vsfSpec.Left = picSpilt.Left + picSpilt.Width
        lblSpec.Left = vsfSpec.Left
    Case 1, 2, 3 '饮片，免煎剂,忽略形态
        vsfSpec.Visible = False
        lblSpec.Visible = False
        lblVar.Left = 50
        vsfVariety.Top = lblVar.Height + lblVar.Top + 100
        vsfVariety.Left = 50
        vsfVariety.Width = Me.ScaleWidth
        vsfVariety.Height = Me.Height - lblVar.Height - lblVar.Top - 500
        picSpilt.Visible = False
    End Select
End Sub

Private Sub picSpilt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        If vsfVariety.Width + x < 200 Then Exit Sub
        If vsfSpec.Width + x < 200 Then Exit Sub
            picSpilt.Left = picSpilt.Left + x
            vsfVariety.Width = vsfVariety.Width + x
            vsfSpec.Width = vsfSpec.Width - x
            vsfSpec.Left = vsfSpec.Left + x
            lblSpec.Left = vsfSpec.Left
    End If
End Sub

Private Sub InitData()
    '初始化界面数据
    Dim intCol As Integer
    Dim intRow As Integer
    
    With vsfSpec
        .SelectionMode = flexSelectionByRow
        .AllowSelection = False '不能多选
        .ExplorerBar = flexExSortShowAndMove '排序和移动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .Cols = menuSpec.Cols
        .Rows = 1
        .TextMatrix(0, menuSpec.ID) = "药品id"
        .TextMatrix(0, menuSpec.编码) = "编码"
        .TextMatrix(0, menuSpec.名称) = "名称"
        .TextMatrix(0, menuSpec.规格) = "规格"
        .TextMatrix(0, menuSpec.计算单位) = "计算单位"
        .TextMatrix(0, menuSpec.产地) = "产地"
        .TextMatrix(0, menuSpec.是否变价) = "是否变价"
        .TextMatrix(0, menuSpec.费用类型) = "费用类型"
        .TextMatrix(0, menuSpec.服务对象) = "服务对象"
        .TextMatrix(0, menuSpec.药名ID) = "药名id"
        
        .ColHidden(menuSpec.ID) = True
        .ColWidth(menuSpec.编码) = 800
        .ColWidth(menuSpec.名称) = 1500
        .ColWidth(menuSpec.规格) = 1000
        .ColWidth(menuSpec.计算单位) = 1000
        .ColWidth(menuSpec.产地) = 1200
        .ColWidth(menuSpec.是否变价) = 850
        .ColWidth(menuSpec.费用类型) = 900
        .ColWidth(menuSpec.服务对象) = 1200
        .ColWidth(menuSpec.药名ID) = 0
        
    End With
    vsfSpec.Cell(flexcpAlignment, 0, 0, 0, vsfSpec.Cols - 1) = flexAlignCenterCenter
    If vsfSpec.Rows > 1 Then
        vsfSpec.Cell(flexcpAlignment, 1, 0, vsfSpec.Rows - 1, vsfSpec.Cols - 1) = flexAlignLeftCenter
    End If
    vsfSpec.Cell(flexcpFontBold, 0, 0, 0, vsfSpec.Cols - 1) = 35
    
    With vsfVariety
        .SelectionMode = flexSelectionByRow
        .AllowSelection = False '不能多选
        .ExplorerBar = flexExSortShowAndMove '排序和移动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .Cols = menuVar.Cols
        .Rows = 1
        .TextMatrix(0, menuVar.ID) = "ID"
        .TextMatrix(0, menuVar.编码) = "编码"
        .TextMatrix(0, menuVar.名称) = "名称"
        .TextMatrix(0, menuVar.计算单位) = "计算单位"
        .TextMatrix(0, menuVar.服务对象) = "服务对象"
        
        .ColHidden(menuVar.ID) = True
        .ColWidth(menuVar.编码) = 800
        .ColWidth(menuVar.名称) = 1500
        .ColWidth(menuVar.计算单位) = 1000
        .ColWidth(menuVar.服务对象) = 1200
    End With
    vsfVariety.Cell(flexcpAlignment, 0, 0, 0, vsfVariety.Cols - 1) = flexAlignCenterCenter
    If vsfVariety.Rows > 1 Then
        vsfVariety.Cell(flexcpAlignment, 1, 0, vsfVariety.Rows - 1, vsfVariety.Cols - 1) = flexAlignLeftCenter
    End If
    vsfVariety.Cell(flexcpFontBold, 0, 0, 0, vsfVariety.Cols - 1) = 35
End Sub

Private Sub vsfSpec_DblClick()
    With vsfSpec
        If Val(.TextMatrix(.Row, menuSpec.ID)) <> 0 Then
            mstrName = .TextMatrix(.Row, menuSpec.药名ID) & "," & .TextMatrix(.Row, menuSpec.ID) & "," & .TextMatrix(.Row, menuSpec.名称) & "(" & .TextMatrix(.Row, menuSpec.规格) & ")" & "," & .TextMatrix(.Row, menuSpec.计算单位)
            Unload Me
        End If
    End With
End Sub

Private Sub vsfSpec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vsfSpec_DblClick
    End If
End Sub

Private Sub vsfVariety_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And Val(vsfVariety.TextMatrix(vsfVariety.Row, menuVar.ID)) <> 0 And mintDosageType = 0 Then '散装才查询规格
        Call GetSpecInfo
    End If
End Sub

Private Sub vsfVariety_DblClick()
    With vsfVariety
        If Val(.TextMatrix(.Row, menuVar.ID)) <> 0 Then
            mstrName = .TextMatrix(.Row, menuVar.ID) & ",0," & .TextMatrix(.Row, menuVar.名称) & "," & .TextMatrix(.Row, menuVar.计算单位)
            Unload Me
        End If
    End With
End Sub

Private Sub vsfVariety_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vsfVariety_DblClick
    End If
End Sub
