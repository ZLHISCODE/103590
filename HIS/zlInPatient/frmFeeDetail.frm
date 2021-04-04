VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFeeDetail 
   BorderStyle     =   0  'None
   Caption         =   "费用明细列表"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   4395
      Left            =   1425
      TabIndex        =   0
      Top             =   990
      Width           =   5850
      _cx             =   10319
      _cy             =   7752
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFeeDetail.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   4
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
Attribute VB_Name = "frmFeeDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngMode As Long, mstrPrivs As String, mblnOriginal As Boolean
Private mbytType As Byte, mstrBalanceID As String

Public Sub ShowMe(ByVal objFont As Object, _
      ByVal lngModule As Long, ByVal strPrivs As String, _
      ByVal bytType As Byte, ByVal strBalanceID As String)
    '-------------------------------------------------------------------------------------------------
    '功能:程序入口,显示单据的明细数据
    '入参:objFont-主窗体字体
    '       lngModule-模块号
    '       strPrivs-权限串
    '　　   bytType:1-全显示;2-只显示未退部分;
    '       strBalanceID -结帐ID
    '编制:刘尔旋
    '日期:2014-06-13
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set vsfDetail.Font = objFont
    Call zlRefresh(bytType, strBalanceID)
End Sub

Public Sub zlRefresh(ByVal bytType As Byte, ByVal strBalanceID As String, Optional blnOriginal As Boolean = True)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '编制:刘尔旋
    '日期:2014-06-13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytType = bytType
    mstrBalanceID = strBalanceID
    mblnOriginal = blnOriginal
    Call ReadListData(bytType)
End Sub

Private Sub SetJZDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant
    
    strHead = "类别,4,800|名称,1,2000|规格,1,1200|数量,7,800|单位,4,800|单价,7,1000|应收金额,7,1000|实收金额,7,1000|执行科室,4,1000|类型,4,1000|说明,1,1800|记录状态,1,0"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .colAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
'        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        .Redraw = True

        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
            If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
        Next i
    End With
    
End Sub

Private Function CheckBalance(lngBalanceID As Long) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = "Select 1 From 病人预交记录 Where 结算序号= [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID)
    CheckBalance = rsTemp.EOF
End Function

Private Function ReadListData(ByVal bytType As Byte) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取相关的明细数据
    '返回:数据获取成功返回true,否则返回False
    '编制:刘尔旋
    '日期:2014-06-13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMain As ADODB.Recordset, rsSub As ADODB.Recordset
    Dim strWithTable As String, strWhere As String, i As Long, str药房单位 As String, str药房包装 As String
    Dim strTable As String, lngMainRow As Long, blnDel As Boolean, bln药房单位 As Boolean
    On Error GoTo errHandle
    blnDel = False
    If bytType = 1 Then
        '收费单
        If CheckBalance(Val(mstrBalanceID)) = True Then
        '10.29以前数据的获取
            strSQL = _
                " Select NO As 单据号, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, " & _
                "       Sum(数量) As 数量, 单价, Sum(应收金额) As 应收金额, Sum(实收金额) As 实收金额, 执行科室, 类型, 说明, 记录状态" & vbNewLine & _
                " From (Select a.结帐ID,D1.名称 as 开单科室,A.开单人,a.No,C.名称 as 类别,Nvl(E.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格," & _
                        IIf(bln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & str药房单位 & ")", "A.计算单位") & " as 单位," & _
                "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
                        IIf(bln药房单位, "/Nvl(X." & str药房包装 & ",1)", "") & ",'9999990.00000') as 数量, " & _
                "       a.费别,To_Char(Sum(A.标准单价)" & _
                        IIf(bln药房单位, "*Nvl(X." & str药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
                "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型," & _
                "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行',9,'异常收费单','第'||ABS(A.执行状态)||'次退费') as 说明," & _
                "       A.记录状态, Nvl(a.价格父号, a.序号) As 序号" & _
                " From  门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 D1,收费项目别名 E,收费项目别名 E1,药品规格 X" & _
                " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
                "       And A.记录性质=1 And A.结帐ID = [1] And A.记录状态" & IIf(blnDel, "=2", " IN(1,3)") & _
                "       And A.收费细目ID=E.收费细目ID(+) And a.开单部门ID=D1.ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(1 = 1, 3, 1) & _
                "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
                " Group by a.结帐id, D1.名称, a.开单人, a.费别,a.No,Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称),E1.名称 , B.规格,A.计算单位,D.名称," & _
                "       Nvl(A.费用类型,B.费用类型),A.执行状态,A.记录状态,X.药品ID )" & _
                " Group By NO, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单价, 执行科室, 类型, 说明, 记录状态" & _
                " Order By 单据号, 序号"
        Else
            strSQL = _
                " Select NO As 单据号, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, " & _
                "       Sum(数量) As 数量, 单价, Sum(应收金额) As 应收金额, Sum(实收金额) As 实收金额, 执行科室, 类型, 说明, 记录状态" & vbNewLine & _
                " From (Select a.结帐ID,D1.名称 as 开单科室,A.开单人,a.No,C.名称 as 类别,Nvl(E.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格," & _
                        IIf(bln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & str药房单位 & ")", "A.计算单位") & " as 单位," & _
                "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
                        IIf(bln药房单位, "/Nvl(X." & str药房包装 & ",1)", "") & ",'9999990.00000') as 数量, " & _
                "       a.费别,To_Char(Sum(A.标准单价)" & _
                        IIf(bln药房单位, "*Nvl(X." & str药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
                "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型," & _
                "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行',9,'异常收费单','第'||ABS(A.执行状态)||'次退费') as 说明," & _
                "       A.记录状态, Nvl(a.价格父号, a.序号) As 序号" & _
                " From  门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 D1,收费项目别名 E,收费项目别名 E1,药品规格 X," & _
                "       (Select Distinct 结帐ID From 病人预交记录 Where 结算序号= [1]) F" & _
                " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
                "       And A.记录性质=1 And A.结帐ID = F.结帐ID And A.记录状态" & IIf(blnDel, "=2", " IN(1,3)") & _
                "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(1 = 1, 3, 1) & _
                "       And A.收费细目ID=E1.收费细目ID(+) And A.开单部门ID=D1.ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
                " Group by a.结帐id, D1.名称, a.开单人, a.费别,a.No,Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称),E1.名称 , B.规格,A.计算单位,D.名称," & _
                "       Nvl(A.费用类型,B.费用类型),A.执行状态,A.记录状态,X.药品ID )" & _
                " Group By NO, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单价, 执行科室, 类型, 说明, 记录状态" & _
                " Order By 单据号, 序号"
        End If
        
        Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrBalanceID)
        Set vsfDetail.DataSource = rsMain
        Call SetDetail
    Else
        '记账单
        strSQL = "" & _
            " Select C.名称 As 类别, Nvl(E.名称, B.名称) As 名称, B.规格, Avg(Nvl(A.付数, 1) * A.数次) As 数量, A.计算单位 As 单位," & vbNewLine & _
            "        Sum(A.标准单价) As 单价, LTrim(To_Char(Sum(A.应收金额), '99999" & gstrDec & "')) As 应收金额," & vbNewLine & _
            "        LTrim(To_Char(Sum(A.实收金额), '99999" & gstrDec & "')) As 实收金额, D.名称 As 执行科室,Nvl(A.费用类型,B.费用类型) As 类型," & vbNewLine & _
            "        Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行',9,'异常收费单','第'||ABS(A.执行状态)||'次退费') as 说明,A.记录状态" & _
            " From 门诊费用记录 A, 收费项目目录 B, 收费项目类别 C, 部门表 D, 收费项目别名 E" & vbNewLine & _
            " Where A.收费细目id = B.ID And A.收费类别 = C.编码 And A.执行部门id = D.ID(+) And A.NO = [1] And A.记录性质 = 2 And" & vbNewLine & _
            "       A.记录状态 In (1,3) And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = 3" & vbNewLine & _
            " Group By Nvl(A.价格父号, A.序号), A.标准单价, C.名称, Nvl(E.名称, B.名称), B.规格, A.计算单位, D.名称, Nvl(A.费用类型,B.费用类型), A.执行状态, A.记录状态" & vbNewLine & _
            " Order By Nvl(A.价格父号, A.序号)"
        Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrBalanceID)
        Set vsfDetail.DataSource = rsMain
        Call SetJZDetail
    End If
    ReadListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub DetailSplitGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对费用列表信息进行分组显示
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo errHandle
    With vsfDetail
        For i = 0 To .Cols - 1
            If i < .ColIndex("类别") And i > .ColIndex("说明") Then
                .ColHidden(i) = True
            End If
        Next
        
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("实收金额"), , &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("应收金额"), , &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("类别")
        .OutlineCol = .ColIndex("类别")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("类别")) = strTemp

                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("单据号"))
                 strTemp = strTemp & Space(2) & "费别:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("费别"))
                 strTemp = strTemp & Space(2) & "开单部门:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单科室"))
                 strTemp = strTemp & Space(2) & "开单人:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单人"))
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("类别"), i, .ColIndex("类别")) = 1
'                 If Val(.TextMatrix(i + 1, .ColIndex("记录状态"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
'                 If Val(.TextMatrix(i + 1, .ColIndex("记录状态"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
'                 If Val(.TextMatrix(i + 1, .ColIndex("记录状态"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                 
                 For j = 0 To .Cols - 1
                    If j < .ColIndex("应收金额") Then
                        If j >= .ColIndex("类别") Then
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = False
                        End If
                    ElseIf .ColIndex("实收金额") = j Then
                        .TextMatrix(i, j) = Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    ElseIf .ColIndex("应收金额") = j Then
                        .TextMatrix(i, j) = " " & Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    End If
                 Next
            Else
                .TextMatrix(i, .ColIndex("单价")) = Format(Val(.TextMatrix(i, .ColIndex("单价"))), gstrDec)
                .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(.TextMatrix(i, .ColIndex("应收金额"))), gstrDec)
                .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(.TextMatrix(i, .ColIndex("实收金额"))), gstrDec)
            End If
        Next
        Call .AutoSize(.ColIndex("类别"))
        Call .AutoSize(.ColIndex("单价"))
        
        For j = 0 To .Cols - 1
            If j < .ColIndex("应收金额") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long, intShowName As Integer
    Dim varData As Variant

    strHead = "单据号,1,0|序号,1,0|开单科室,1,0|开单人,1,0|费别,1,0|类别,4,800|名称,1,2000|商品名,1,2000|规格,1,1200|数量,7,800|单价,7,1000|应收金额,7,1000|实收金额,7,1000|执行科室,4,1000|类型,4,1000|说明,1,1800|记录状态,1,0"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .colAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        'Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        
        '.Row = 0: .Col = 0: .ColSel = .Cols - 1
        .Redraw = True
        If .Rows > 1 Then
            If .TextMatrix(1, .ColIndex("单据号")) <> "" Then Call DetailSplitGroup
        End If
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) = False Then
                If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                .RowHeight(i) = 300
            End If
        Next i
        
        intShowName = Val(zlDatabase.GetPara("药品名称显示"))
        If intShowName <> 2 Then
            .ColHidden(.ColIndex("商品名")) = True
        Else
            .ColHidden(.ColIndex("商品名")) = False
        End If
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With vsfDetail
        .Top = 0
        .Left = 0
        .Height = Me.Height
        .width = Me.width
    End With
End Sub
