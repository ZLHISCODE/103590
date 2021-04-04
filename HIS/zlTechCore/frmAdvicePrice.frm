VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdvicePrice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1290
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   4755
   ControlBox      =   0   'False
   Icon            =   "frmAdvicePrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmAdvicePrice"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   75
      ScaleHeight     =   210
      ScaleWidth      =   4575
      TabIndex        =   1
      Top             =   75
      Width           =   4575
      Begin VB.Label lblClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4350
         TabIndex        =   3
         Top             =   15
         Width           =   210
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊疗计价"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   75
         TabIndex        =   2
         Top             =   15
         Width           =   780
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   900
      Left            =   75
      TabIndex        =   0
      Top             =   330
      Width           =   4575
      _cx             =   8070
      _cy             =   1587
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
      BackColor       =   15659506
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   13811126
      ForeColorSel    =   0
      BackColorBkg    =   15659506
      BackColorAlternate=   15659506
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   15659506
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdvicePrice.frx":000C
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
   Begin VB.Shape Bdr 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   1230
      Left            =   45
      Top             =   45
      Width           =   4665
   End
End
Attribute VB_Name = "frmAdvicePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event PanelHide()
Private COL_序号 As Long
Private COL_相关ID As Long
Private COL_医嘱状态 As Long
Private COL_诊疗类别 As Long
Private COL_诊疗项目ID As Long
Private COL_收费细目ID As Long
Private COL_标本部位 As Long
Private COL_计价特性 As Long
Private COL_执行性质 As Long
Private COL_执行科室ID As Long

Private mfrmParent As Object
Private vsAdvice As VSFlexGrid
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng科室ID As Long
Private mint场合 As Integer

Public Sub HideMe()
    If mlng病人ID <> 0 Then Me.Hide
End Sub

Public Sub ShowMe(frmParent As Object, vsEdit As Object, ByVal lng病人ID As Long, lng主页ID As Long, ByVal lng科室ID As Long, ByVal strCol As String)
'参数：lng主页ID=门诊调用时传入0
    Dim arrCol As Variant
    
    Set mfrmParent = frmParent
    Set vsAdvice = vsEdit
    
    arrCol = Split(strCol, ",")
    COL_序号 = arrCol(0)
    COL_相关ID = arrCol(1)
    COL_医嘱状态 = arrCol(2)
    COL_诊疗类别 = arrCol(3)
    COL_诊疗项目ID = arrCol(4)
    COL_收费细目ID = arrCol(5)
    COL_标本部位 = arrCol(6)
    COL_计价特性 = arrCol(7)
    COL_执行性质 = arrCol(8)
    COL_执行科室ID = arrCol(9)
    
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng科室ID = lng科室ID
    mint场合 = IIF(mlng主页ID = 0, 1, 2)
        
    Call ShowPrice
    Me.Show , frmParent
    
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Function ShowPrice() As Boolean
'功能：读取指定医嘱的计价,并根据当前的诊疗收费关系进行更新
    Dim rs收费细目 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim str收费细目IDs As String
    Dim strSQL As String, i As Long, j As Long
    Dim bln配方行 As Boolean, bln检验行 As Boolean, blnLoad As Boolean
    Dim lng病人科室ID As Long, lng执行科室ID As Long
    Dim dblPrice As Double, lngRow As Long, lngW As Long
    
    Dim strAdvice As String, lngBegin As Long, lngEnd As Long
    
    On Error GoTo errH
        
    With vsAdvice
        lngRow = .Row
        
        '生成病人医嘱记录临时表
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                strAdvice = strAdvice & " Union ALL " & _
                    "Select " & .RowData(i) & " as ID," & Val(.TextMatrix(i, COL_序号)) & " as 序号," & ZVal(.TextMatrix(i, COL_相关ID)) & " as 相关ID," & _
                    Val(.TextMatrix(i, COL_医嘱状态)) & " as 医嘱状态,'" & .TextMatrix(i, COL_诊疗类别) & "' as 诊疗类别," & _
                    Val(.TextMatrix(i, COL_诊疗项目ID)) & " as 诊疗项目ID," & ZVal(.TextMatrix(i, COL_收费细目ID)) & " as 收费细目ID," & _
                    "'" & .TextMatrix(i, COL_标本部位) & "' as 标本部位," & Val(.TextMatrix(i, COL_计价特性)) & " as 计价特性," & _
                    Val(.TextMatrix(i, COL_执行性质)) & " as 执行性质," & ZVal(.TextMatrix(i, COL_执行科室ID), True) & " as 执行科室ID From Dual"
            End If
        Next
        strAdvice = Mid(strAdvice, 12)
    End With
    
    With vsPrice
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If vsAdvice.RowData(lngRow) = 0 Then
            .Redraw = True: ShowPrice = True: Exit Function
        End If
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            bln配方行 = RowIn配方行(lngRow)
            bln检验行 = RowIn检验行(lngRow)
        End If
                                    
        blnLoad = True
        
        '药品的计价
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '中,西成药:可能按规格下医嘱,计算1个药房包装的单价
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,C.ID as 收费细目ID," & _
                " Decode([3],1,B.门诊包装,B.住院包装) as 药房包装,Decode([3],1,B.门诊单位,B.住院单位) as 计算单位," & _
                " 1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*Decode([3],1,B.门诊包装,B.住院包装) as 单价," & _
                " A.执行科室ID,0 as 从项" & _
                " From (" & strAdvice & ") A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.诊疗项目ID=B.药名ID And B.药品ID=C.ID And Nvl(A.执行性质,0)<>5" & _
                " And (A.收费细目ID is NULL Or A.收费细目ID=B.药品ID)" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.服务对象 IN([3],3) And D.收费细目ID=C.ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
                
                '仅一并给药(如果是)的第一成药行才显示给药途径的计价
                blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_相关ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
        ElseIf bln配方行 Then
            '中草药:一定对应有规格记录且填写了收费细目ID
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,C.ID as 收费细目ID," & _
                " Decode([3],1,B.门诊包装,B.住院包装) as 药房包装,Decode([3],1,B.门诊单位,B.住院单位) as 计算单位," & _
                " 1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*Decode([3],1,B.门诊包装,B.住院包装) as 单价," & _
                " A.执行科室ID,0 as 从项" & _
                " From (" & strAdvice & ") A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别='7' And A.相关ID=[1]" & _
                " And A.收费细目ID=B.药品ID And A.收费细目ID=C.ID And C.服务对象 IN([3],3)" & _
                " And D.收费细目ID=C.ID And Nvl(A.执行性质,0)<>5" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        End If
        
        '读取现有计价(取最新价格)：除药品外的计价,包含相关医嘱计价
        '不计价,手工计价的医嘱不读取
        '用Union方式可以利用索引
        If blnLoad Then
            '不是新开的医嘱，根据病人医嘱计价提取
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位," & _
                " B.收费细目ID,1 as 药房包装,C.计算单位,B.数量,Decode(C.是否变价,1,B.单价,Sum(D.现价)) as 单价," & _
                " Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID,Nvl(B.从项,0) as 从项" & _
                " From (" & strAdvice & ") A,病人医嘱计价 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别 Not IN('5','6','7') And A.ID=B.医嘱ID" & _
                " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0)<>5 And B.收费细目ID=C.ID And B.收费细目ID=D.收费细目ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                " And (A.ID=[1] Or A.ID=[2] Or A.相关ID=[1])" & _
                " Group by A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,B.收费细目ID," & _
                " C.计算单位,B.数量,C.是否变价,B.单价,Nvl(B.执行科室ID,A.执行科室ID),Nvl(B.从项,0)"
            '新开的医嘱，根据诊疗收费关系提取(非药变价显示为0)
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,B.收费项目ID," & _
                " 1 as 药房包装,C.计算单位,B.收费数量 as 数量,Decode(C.是否变价,1,0,Sum(D.现价)) as 单价," & _
                " A.执行科室ID,Nvl(B.从属项目,0) as 从项" & _
                " From (" & strAdvice & ") A,诊疗收费关系 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别 Not IN('5','6','7') And A.医嘱状态 IN(1,2) And A.诊疗项目ID=B.诊疗项目ID" & _
                " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0)<>5 And B.收费项目ID=C.ID And B.收费项目ID=D.收费细目ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) And C.服务对象 IN([3],3)" & _
                " And (A.ID=[1] Or A.ID=[2] Or A.相关ID=[1])" & _
                " Group by A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,B.收费项目ID," & _
                " C.计算单位,B.收费数量,C.是否变价,A.执行科室ID,Nvl(B.从属项目,0)"
        End If
        
        '读取诊疗项目信息
        strSQL = "Select /*+ RULE */ A.*,B.名称 as 诊疗项目,C.名称 as 诊疗类别名称" & _
            " From (" & strSQL & ") A,诊疗项目目录 B,诊疗项目类别 C" & _
            " Where A.诊疗项目ID=B.ID And B.类别=C.编码"
        strSQL = strSQL & " Order by 序号,从项"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.RowData(lngRow)), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)), mint场合)
        
        '显示计价内容
        If Not rsTmp.EOF Then
            '确定显示行数
            .Rows = .FixedRows + rsTmp.RecordCount
            
            '获取诊疗项目,收费细目信息
            For i = 1 To rsTmp.RecordCount
                str收费细目IDs = str收费细目IDs & " Union ALL Select " & rsTmp!收费细目ID & " From Dual"
                rsTmp.MoveNext
            Next
            str收费细目IDs = Mid(str收费细目IDs, 12)
                        
            strSQL = "Select A.ID,A.类别,B.名称 as 类别名称,A.编码," & _
                " A.名称,A.规格,A.产地,A.费用类型,A.是否变价" & _
                " From 收费项目目录 A,收费项目类别 B" & _
                " Where A.类别=B.编码 And A.ID IN(" & str收费细目IDs & ")"
            strSQL = "Select A.ID,A.类别,A.类别名称,A.编码,Nvl(B.名称,A.名称) as 名称," & _
                " A.规格,A.产地,A.费用类型,A.是否变价,C.跟踪在用" & _
                " From (" & strSQL & ") A,收费项目别名 B,材料特性 C" & _
                " Where A.ID=C.材料ID(+) And A.ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIF(gbln商品名, 3, 1)
            Call zlDatabase.OpenRecordset(rs收费细目, strSQL, Me.Name) 'IN
            
            '显示每行内容
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rs收费细目.Filter = "ID=" & rsTmp!收费细目ID
                
                '计价医嘱
                If InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, 0) = "药品"
                ElseIf rsTmp!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    .TextMatrix(i, 0) = "给药"
                ElseIf rsTmp!诊疗类别 = "E" And (bln配方行 Or bln检验行) Then
                    If bln检验行 Then
                        .TextMatrix(i, 0) = "采集"
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        .TextMatrix(i, 0) = "煎法"
                    Else
                        .TextMatrix(i, 0) = "用法"
                    End If
                ElseIf Not IsNull(rsTmp!相关ID) Then
                    If rsTmp!诊疗类别 = "C" Then
                        .TextMatrix(i, 0) = "检验"
                    ElseIf rsTmp!诊疗类别 = "D" Then
                        .TextMatrix(i, 0) = "部位"
                    ElseIf rsTmp!诊疗类别 = "F" Then
                        .TextMatrix(i, 0) = "附术"
                    ElseIf rsTmp!诊疗类别 = "G" Then
                        .TextMatrix(i, 0) = "麻醉"
                    End If
                Else
                    .TextMatrix(i, 0) = rsTmp!诊疗类别名称
                End If
                
                '类别
                .TextMatrix(i, 1) = rs收费细目!类别名称
                '收费项目:规格/产地
                .TextMatrix(i, 2) = rs收费细目!名称
                If Not IsNull(rs收费细目!规格) Then
                    .TextMatrix(i, 2) = .TextMatrix(i, 2) & " " & rs收费细目!规格
                End If
                
                '计价数量:药嘱药品为1,非药嘱药品为对应售价数
                '计算单位:药嘱药品为药房单位,非药嘱药品为售价单位
                .TextMatrix(i, 3) = FormatEx(rsTmp!数量, 5) & Nvl(rsTmp!计算单位)
                
                '执行科室
                lng执行科室ID = Nvl(rsTmp!执行科室ID, 0)
                If rs收费细目!类别 = "4" And Nvl(rs收费细目!跟踪在用, 0) = 1 _
                    Or InStr(",5,6,7,", rs收费细目!类别) > 0 And InStr(",5,6,7,", rsTmp!诊疗类别) = 0 Then
                    lng病人科室ID = mlng科室ID
                    lng执行科室ID = Get收费执行科室ID(mlng病人ID, mlng主页ID, rs收费细目!类别, rs收费细目!ID, 4, lng病人科室ID, 0, mint场合, lng执行科室ID)
                End If
                
                '单价处理
                If InStr(",5,6,7,", rs收费细目!类别) > 0 Then
                    If Nvl(rs收费细目!是否变价, 0) = 1 Then
                        '求药品时价
                        If InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                            '药嘱药品计算一个药房包装的时价
                            .TextMatrix(i, 4) = CalcDrugPrice(rs收费细目!ID, lng执行科室ID, Nvl(rsTmp!药房包装, 1))
                            .TextMatrix(i, 4) = FormatEx(Val(.TextMatrix(i, 5)) * Nvl(rsTmp!药房包装, 0), 5)
                        Else
                            '非药嘱药品计算相对售价数量的售价实价
                            .TextMatrix(i, 4) = FormatEx(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, Nvl(rsTmp!数量, 0)), 5)
                        End If
                    Else
                        '药嘱药品为药房单价,非药药品为售价
                        .TextMatrix(i, 4) = FormatEx(Nvl(rsTmp!单价), 5)
                    End If
                ElseIf rs收费细目!类别 = "4" And Nvl(rs收费细目!跟踪在用, 0) = 1 And Nvl(rs收费细目!是否变价, 0) = 1 Then
                    '时价卫材的单价和药品一样计算
                    .TextMatrix(i, 4) = FormatEx(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, Nvl(rsTmp!数量, 0)), 5)
                Else
                    .TextMatrix(i, 4) = FormatEx(Nvl(rsTmp!单价), 5)
                End If
                
                '费用类型
                .TextMatrix(i, 5) = Nvl(rs收费细目!费用类型)
                
                dblPrice = dblPrice + FormatEx(Nvl(rsTmp!数量, 0) * Val(.TextMatrix(i, 4)), 5)
                
                rsTmp.MoveNext
            Next
        End If
        
        '处理表格尺寸
        With vsPrice
            If .Rows < 3 Then .Rows = 3
            Call .AutoSize(0, .Cols - 1)
            For i = 0 To .Cols - 1
                If .ColWidth(i) > 1500 Then
                    .ColWidth(i) = 1500
                Else
                    .ColWidth(i) = .ColWidth(i) - 90
                End If
                lngW = lngW + .ColWidth(i)
            Next
            .Width = lngW + IIF(.Rows > 6, 225, 0)
            .Height = .RowHeight(1) * IIF(.Rows > 6, 6, .Rows)
        End With
        
        .Row = 1: .Col = 0
        .Redraw = True
    End With
    Call SetFormSize
    ShowPrice = True
    Exit Function
errH:
    vsPrice.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    Dim lngS组ID As Long, lngO组ID As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS组ID = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) = 0, .RowData(lngRow), Val(.TextMatrix(lngRow, COL_相关ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_相关ID)))
            If lngO组ID = lngS组ID Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_相关ID)))
            If lngO组ID = lngS组ID Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Function RowIn检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于检验组合中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
            '采集方法行
            If .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
                RowIn检验行 = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "C" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '检验项目行
            RowIn检验行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于中药配方中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
                '用法行
                If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_诊疗类别) = "E" Then
                    RowIn配方行 = True: Exit Function
                End If
            Else
                '煎法行
                If .TextMatrix(lngRow - 1, COL_诊疗类别) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    RowIn配方行 = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "7" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '中药行
            RowIn配方行 = True: Exit Function
        End If
    End With
End Function

Private Sub Form_Load()
    Dim strPos As String
    
    Call FormSetCaption(Me, False, False)

    strPos = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "PricePanePostion", "1600,5500")
    Me.Top = mfrmParent.Top + Val(Split(strPos, ",")(0))
    Me.Left = mfrmParent.Left + Val(Split(strPos, ",")(1))
End Sub

Private Sub SetFormSize()
    LockWindowUpdate Me.Hwnd
    Me.Width = vsPrice.Width + (Bdr.BorderWidth * 15 + 30) * 2
    Me.Height = vsPrice.Height + picTitle.Height + (Bdr.BorderWidth * 15 + 30) * 2 - 15
    
    Bdr.Left = 15
    Bdr.Top = 15
    Bdr.Width = Me.Width - 15
    Bdr.Height = Me.Height - 15
    
    picTitle.Left = Bdr.Left + Bdr.BorderWidth * 15 + 15
    picTitle.Top = Bdr.Top + Bdr.BorderWidth * 15 + 15
    picTitle.Width = Me.Width - picTitle.Left * 2
    
    vsPrice.Left = picTitle.Left
    vsPrice.Top = picTitle.Top + picTitle.Height
    
    Call SetCloseButton(0, True)
    LockWindowUpdate 0
End Sub

Private Sub SetCloseButton(ByVal intState As Integer, Optional ByVal blnSize As Boolean)
'参数：intState=0-正常,1-弹起,2-按下
    If intState = 0 Then
        lblClose.BackColor = picTitle.BackColor
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 0
    ElseIf intState = 1 Then
        lblClose.BackColor = vsPrice.BackColorSel
        lblClose.ForeColor = vbBlack
        lblClose.BorderStyle = 1
    ElseIf intState = 2 Then
        lblClose.BackColor = 11899525
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 1
    End If
    
    If blnSize Then
        lblClose.Width = 210
        lblClose.Height = 195
        lblClose.Left = picTitle.Width - lblClose.Width - 15
        lblClose.Top = (picTitle.Height - lblClose.Height) / 2
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
    End If
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetCloseButton(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngTop As Long, lngLeft As Long
    
    '保存相对于主窗体右上角的位置
    If mfrmParent.WindowState = 0 Then
        lngTop = Me.Top - mfrmParent.Top
        lngLeft = Me.Left - mfrmParent.Left
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "PricePanePostion", lngTop & "," & lngLeft
    End If
    
    mlng病人ID = 0
    mlng主页ID = 0
    mlng科室ID = 0
    Set mfrmParent = Nothing
    Set vsAdvice = Nothing
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call SetCloseButton(2)
    End If
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x >= 0 And y >= 0 And x <= lblClose.Width And y <= lblClose.Height Then
        If Button = 1 Then
            Call SetCloseButton(2)
        Else
            Call SetCloseButton(1)
        End If
    Else
        Call SetCloseButton(1)
    End If
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x >= 0 And y >= 0 And x <= lblClose.Width And y <= lblClose.Height Then
        Me.Hide
        RaiseEvent PanelHide
        If mfrmParent.Visible Then mfrmParent.SetFocus
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
    End If
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
        If mfrmParent.Visible Then mfrmParent.SetFocus
    End If
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetCloseButton(0)
End Sub

Private Sub vsPrice_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Sub vsPrice_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetCloseButton(0)
    With vsPrice
        If .MouseCol = 2 And Between(.MouseRow, .FixedRows, .Rows - 1) Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        Else
            .ToolTipText = ""
        End If
    End With
End Sub
