VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMediCheckCard 
   Caption         =   "药品验收编辑"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "frmMediCheckCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   10830
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      Height          =   315
      IMEMode         =   2  'OFF
      Left            =   9360
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   1425
   End
   Begin VB.TextBox txt备注 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   14
      Top             =   6660
      Width           =   9975
   End
   Begin VB.TextBox txtProvider 
      Enabled         =   0   'False
      Height          =   300
      Left            =   7500
      TabIndex        =   11
      Top             =   660
      Width           =   2895
   End
   Begin VB.CommandButton cmdProvider 
      Caption         =   "…"
      Enabled         =   0   'False
      Height          =   300
      Left            =   10455
      TabIndex        =   10
      Top             =   660
      Width           =   300
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   9
      Top             =   7080
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8400
      TabIndex        =   8
      Top             =   7080
      Width           =   1100
   End
   Begin VB.TextBox txtVerify 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6360
      TabIndex        =   7
      Top             =   6270
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCheck 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Top             =   6270
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfBill 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   10695
      _cx             =   18865
      _cy             =   8916
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediCheckCard.frx":030A
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
   Begin VB.ComboBox cboStock 
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   660
      Width           =   1920
   End
   Begin VB.Label LblVerifyDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "复核日期"
      Height          =   180
      Left            =   7920
      TabIndex        =   20
      Top             =   6300
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TxtVerifyDate 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   300
      Left            =   8730
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label LblCheckDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "验收日期"
      Height          =   180
      Left            =   2400
      TabIndex        =   18
      Top             =   6300
      Width           =   720
   End
   Begin VB.Label TxtCheckDate 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   300
      Left            =   3180
      TabIndex        =   17
      Top             =   6240
      Width           =   1875
   End
   Begin VB.Label LblNo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   8880
      TabIndex        =   16
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl说明 
      AutoSize        =   -1  'True
      Caption         =   "备注"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   360
   End
   Begin VB.Label LblProvider 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "供药单位"
      Height          =   180
      Left            =   6750
      TabIndex        =   12
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblVerify 
      AutoSize        =   -1  'True
      Caption         =   "复核人"
      Height          =   180
      Left            =   5640
      TabIndex        =   6
      Top             =   6330
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "验收人"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   6330
      Width           =   540
   End
   Begin VB.Label lblStore 
      AutoSize        =   -1  'True
      Caption         =   "验收库房"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   720
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "药品验收单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "frmMediCheckCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint编辑状态 As Integer '1-新增 2-修改 3-复核 4-查看
Private mlng验收id As Long  '验收id,修改和查看状态下有该值
Private mstrMatch As String         '匹配方式
Private mlng库房ID As Long
Private mstr验收结论 As String '记录默认验收结论

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Public Sub showMe(ByVal int编辑状态 As Integer, ByVal fraPar As Form, ByVal lng库房ID As Long, ByVal lng验收id As Long)
    mint编辑状态 = int编辑状态
    mlng库房ID = lng库房ID
    mlng验收id = lng验收id
    
    Me.Show vbModal, fraPar
End Sub

'检查数据依赖性
Private Function CheckDepend() As Boolean
    Dim rsDepend As New ADODB.Recordset
    Dim strStock As String, strCaption As String
    
    CheckDepend = False
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
        MsgBox "该人员无药品入库验收”权限，请与管理员联系！", vbInformation, gstrSysName
        If rsDepend.State = 1 Then rsDepend.Close
        Exit Function
    End If
    
    '装入库房数据
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!名称
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = mlng库房ID Then
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

Private Sub cboStock_Click()
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        Call SetSelectorRS(1, "药品入库验收管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0)
    End If
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    '获取可操作的库房
    str工作性质 = "H,I,J,K,L,M,N"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfBill): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfBill, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), str工作性质, True) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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
        vsfBill.SetFocus
        vsfBill.Row = 1
        vsfBill.Col = vsfBill.ColIndex("药品名称")
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CmdSave_Click()
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        If ValidData = False Then Exit Sub
        
        If SaveCard = True Then
            MsgBox "保存成功！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    If mint编辑状态 = 3 Then
        If mlng验收id = 0 Then
            Exit Sub
        End If
        
        If SaveCheck(mlng验收id) = True Then
            MsgBox "复核成功！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
End Sub

Private Function SaveCheck(ByVal lng验收id As Long) As Boolean
    '审核单据
    Dim strVerifyDate As String
    
    On Error GoTo errHandle
    
    strVerifyDate = TxtVerifyDate.Caption
    
    gstrSQL = "Zl_药品验收记录_Verify("
    '验收id_In   In 药品验收记录.Id%Type,
    gstrSQL = gstrSQL & lng验收id & ","
    '复核人_In   In 药品验收记录.复核人%Type,
    gstrSQL = gstrSQL & "'" & txtVerify.Text & "',"
    '复核日期_In In 药品验收记录.复核日期%Type
    gstrSQL = gstrSQL & "to_date('" & strVerifyDate & "','yyyy-mm-dd HH24:MI:SS'))"
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, "SaveCard")
    SaveCheck = True
            
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function SaveCard() As Boolean
    '新增、修改保存
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strNo As String
    Dim lng验收id As Long
    Dim strCheckDate As String
    Dim int合格 As Integer
    Dim bln执行过程 As Boolean
    Dim arrSql As Variant
    Dim str生产日期 As String
    Dim str效期 As String
    Dim str进药日期 As String
        
    arrSql = Array()
        
    On Error GoTo errHandle

    If txtNo.Text = "" Then
        strNo = zlDataBase.GetNextNo(148, Me.cboStock.ItemData(Me.cboStock.ListIndex))
    Else
        strNo = txtNo.Text
    End If
    
    With vsfBill
        If .rows > 1 Then
            If mint编辑状态 = 2 Then '修改
                gstrSQL = "Zl_药品验收记录_Delete(" & mlng验收id & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            Else
                gstrSQL = "Select 药品验收记录_Id.Nextval as id From Dual"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "验收保存")
                mlng验收id = rsTemp!id
            End If
            
            strCheckDate = Format(txtCheckDate.Caption, "yyyy-mm-dd hh:mm:ss")
            '检查是否都合格了 1表示不合格，0表示合格
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("药品id")) <> "" Then
                    If .TextMatrix(lngRow, .ColIndex("验收结果")) = "不合格" Then
                        int合格 = 1
                        Exit For
                    End If
                End If
            Next
            '主表插入
            gstrSQL = "Zl_药品验收记录_Insert ("
            'Id_In         In 药品验收记录.Id%Type,
            gstrSQL = gstrSQL & mlng验收id & ","
            'No_In         In 药品验收记录.No%Type,
            gstrSQL = gstrSQL & "'" & strNo & "',"
            '库房id_In     In 药品验收记录.库房id%Type,
            gstrSQL = gstrSQL & cboStock.ItemData(Me.cboStock.ListIndex) & ","
            '供药单位id_In In 药品验收记录.供药单位id%Type,
            gstrSQL = gstrSQL & Val(txtProvider.Tag) & ","
            '验收人_In     In 药品验收记录.验收人%Type,
            gstrSQL = gstrSQL & "'" & txtCheck.Text & "',"
            '验收日期_In   In 药品验收记录.验收日期%Type,
            gstrSQL = gstrSQL & "to_date('" & strCheckDate & "','yyyy-mm-dd HH24:MI:SS'),"
            '是否合格_In   In 药品验收记录.是否合格%Type,
            gstrSQL = gstrSQL & int合格 & ","
            '备注_in     in 药品验收记录.备注%type
            gstrSQL = gstrSQL & IIf(txt备注.Text = "", "NULL", "'" & txt备注.Text & "'") & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            '次表插入
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("药品id")) <> "" Then
                    str生产日期 = .TextMatrix(lngRow, .ColIndex("生产日期"))
                    str效期 = .TextMatrix(lngRow, .ColIndex("效期"))
                    str进药日期 = .TextMatrix(lngRow, .ColIndex("进药日期"))
                    
                    gstrSQL = "Zl_药品验收明细_Insert ("
                    '验收id_In   In 药品验收明细.验收id%Type,
                    gstrSQL = gstrSQL & mlng验收id & ","
                    '药品id_In   In 药品验收明细.药品id%Type,
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("药品id"))) & ","
                    '成本价_In   In 药品验收明细.成本价%Type,
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("成本价"))) & ","
                    '零售价_In   In 药品验收明细.零售价%Type,
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("零售价"))) & ","
                    '进药数量_In In 药品验收明细.进药数量%Type,
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("进药数量"))) & ","
                    '批号_In     In 药品验收明细.批号%Type,
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("药品批号")) & "',"
                    '生产日期_In In 药品验收明细.生产日期%Type,
                    gstrSQL = gstrSQL & IIf(str生产日期 = "", "NULL", "to_date('" & str生产日期 & "','yyyy-mm-dd')") & ","
                    '效期_In     In 药品验收明细.效期%Type,
                    gstrSQL = gstrSQL & IIf(str效期 = "", "NULL", "to_date('" & str效期 & "','yyyy-mm-dd')") & ","
                    '产地_In     In 药品验收明细.产地%Type,
                    gstrSQL = gstrSQL & IIf(.TextMatrix(lngRow, .ColIndex("产地")) = "", "NULL", "'" & .TextMatrix(lngRow, .ColIndex("产地")) & "'") & ","
                    '批准文号_In In 药品验收明细.批准文号%Type,
                    gstrSQL = gstrSQL & IIf(.TextMatrix(lngRow, .ColIndex("批准文号")) = "", "NULL", "'" & .TextMatrix(lngRow, .ColIndex("批准文号")) & "'") & ","
                    '进药日期_In In 药品验收明细.进药日期%Type,
                    gstrSQL = gstrSQL & IIf(str进药日期 = "", "NULL", "to_date('" & str进药日期 & "','yyyy-mm-dd')") & ","
                    '是否合格_In In 药品验收明细.是否合格%Type
                    gstrSQL = gstrSQL & IIf(.TextMatrix(lngRow, .ColIndex("验收结果")) = "不合格", 1, 0) & ","
                    '验收结论_In In 药品验收明细.验收结论%Type
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("验收结论")) & "')"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
            Next
            
            bln执行过程 = True
            gcnOracle.BeginTrans
            For lngRow = 0 To UBound(arrSql)
                Call zlDataBase.ExecuteProcedure(CStr(arrSql(lngRow)), "SaveCard")
            Next
            gcnOracle.CommitTrans
            SaveCard = True
        Else
            SaveCard = False
            Exit Function
        End If
    End With

    Exit Function
errHandle:
    If bln执行过程 = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ValidData() As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    
    '保存时数据检查
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "请选择一个供应商！", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If
    
    With vsfBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("药品id")) <> "" And Val(.TextMatrix(lngRow, .ColIndex("零售价"))) = 0 Then
                MsgBox "第" & lngRow & "行数据零售价不能为零！", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("零售价")
                .SetFocus
                Exit Function
            End If
            
            If .TextMatrix(lngRow, .ColIndex("药品id")) <> "" And Val(.TextMatrix(lngRow, .ColIndex("进药数量"))) = 0 Then
                MsgBox "第" & lngRow & "行数据进药数量不能为零！", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("进药数量")
                .SetFocus
                Exit Function
            End If
        Next
    End With
    ValidData = True
End Function

Private Sub Form_Load()
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    If CheckDepend = False Then Exit Sub
    mstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")  '匹配方式
    
    Call GetDrugDigit(mlng库房ID, "药品验收管理", 4, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initGrid
    
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        cboStock.Enabled = True
        txtProvider.Enabled = True
        cmdProvider.Enabled = True
        txtCheck.Text = UserInfo.用户姓名
        txt备注.Enabled = True
                
        If mint编辑状态 = 1 Then
            txtNo.Text = zlDataBase.GetNextNo(148, Me.cboStock.ItemData(Me.cboStock.ListIndex))
            txtCheckDate.Caption = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
    End If
    If mint编辑状态 = 3 Then
        txtVerify.Text = UserInfo.用户姓名
        TxtVerifyDate.Caption = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    End If
    
    If mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 4 Then
        Call initCard
        If mint编辑状态 = 3 Then
            CmdSave.Caption = "复核(&O)"
        End If
        If mint编辑状态 = 4 Then
            CmdSave.Visible = False
        End If
    End If
    RestoreWinState Me, App.ProductName, "药品验收入库"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With LblTitle
        .Left = 0
        .Top = 120
        .Width = Me.Width
    End With
    
    With txtNo
        .Move Me.Width - .Width - 300
    End With
    LblNo.Move txtNo.Left - LblNo.Width - 100
    
    With lblStore
        .Move 100, 720
    End With
    
    With cboStock
        .Move lblStore.Left + lblStore.Width + 50, lblStore.Top - 60
    End With
    
    With cmdProvider
        .Move Me.Width - cmdProvider.Width - 300, lblStore.Top - 60
    End With
    
    With txtProvider
        .Move cmdProvider.Left - .Width - 10, lblStore.Top - 60
    End With
    
    With LblProvider
        .Move txtProvider.Left - .Width - 100, lblStore.Top
    End With
    
    With vsfBill
        .Move lblStore.Left, lblStore.Top + lblStore.Height + 100, Me.Width - lblStore.Left - 300, Me.Height - .Top - txtVerify.Height - 1500
    End With
    
    lblCheck.Move vsfBill.Left, vsfBill.Top + vsfBill.Height + 200
    txtCheck.Move lblCheck.Left + lblCheck.Width + 100, lblCheck.Top - 60
    lblCheckDate.Move txtCheck.Left + txtCheck.Width + 100, vsfBill.Top + vsfBill.Height + 200
    txtCheckDate.Move lblCheckDate.Left + lblCheckDate.Width + 100, lblCheck.Top - 60
    
    If mint编辑状态 = 3 Or mint编辑状态 = 4 Then
        lblVerify.Visible = True
        txtVerify.Visible = True
        LblVerifyDate.Visible = True
        TxtVerifyDate.Visible = True
        
        lblVerify.Move txtCheckDate.Left + txtCheckDate.Width + 500, vsfBill.Top + vsfBill.Height + 200
        txtVerify.Move lblVerify.Left + lblVerify.Width + 100, lblCheck.Top - 60
        
        LblVerifyDate.Move txtVerify.Left + txtVerify.Width + 100, lblVerify.Top
        TxtVerifyDate.Move LblVerifyDate.Left + LblVerifyDate.Width + 200, lblVerify.Top - 60
    End If
    
    lbl说明.Move lblCheck.Left, lblCheck.Top + lblCheck.Height + 100
    txt备注.Move txtCheck.Left, lbl说明.Top - 20, vsfBill.Width - lbl说明.Left - 530
    
    CmdCancel.Move Me.Width - CmdCancel.Width - 300, lbl说明.Top + lbl说明.Height + 180
    CmdSave.Move CmdCancel.Left - CmdSave.Width - 200, lbl说明.Top + lbl说明.Height + 180
End Sub

Private Sub initGrid()
    '初始化表格
    With vsfBill
        .ColComboList(.ColIndex("药品名称")) = "|..."
        .ColComboList(.ColIndex("产地")) = "|..."
        .ColComboList(.ColIndex("验收结论")) = "|..."
        .ColDataType(.ColIndex("进药日期")) = flexDTDate
        .ColDataType(.ColIndex("生产日期")) = flexDTDate
        .ColDataType(.ColIndex("效期")) = flexDTDate
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, "药品验收入库"
    Call ReleaseSelectorRS  '卸载数据集
End Sub

Private Sub txtProvider_GotFocus()
    zlControl.TxtSelAll txtProvider
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
    dblTop = vRect.Top - 700
    
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        gstrSQL = "Select id,编码,名称,简码 From 供应商 " & _
                  "Where (站点 = [2] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                  "  And 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                  "  And (简码 like [1] Or 编码 like [1] or 名称 like [1] )"
             
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "供药单位", False, "", "", False, False, _
                        True, dblLeft, dblTop, 1000, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        
        If blnCancel Then txtProvider.SetFocus: Exit Sub
        
        If rsProvider Is Nothing Then
            MsgBox "未匹配到您输入的供药单位", vbOKOnly + vbInformation, gstrSysName
            txtProvider.SelStart = 0
            txtProvider.SelLength = Len(txtProvider)
            Exit Sub
        Else
            txtProvider.Text = rsProvider!名称
            txtProvider.Tag = rsProvider!id
            vsfBill.SetFocus
            vsfBill.Row = 1
            vsfBill.Col = vsfBill.ColIndex("药品名称")
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt备注_GotFocus()
    zlControl.TxtSelAll txt备注
End Sub


Private Sub vsfBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As ADODB.Recordset
    Dim i As Integer
    Dim intRow As Integer
    Dim intOldRow As Integer
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo errHandle
    With vsfBill
        intOldRow = vsfBill.Row
        vRect = GetControlRect(vsfBill.hWnd) '获取位置
        dblLeft = vRect.Left + vsfBill.CellLeft
        dblTop = vRect.Top - vsfBill.Height + vsfBill.CellTop + vsfBill.CellHeight
        
        Select Case .ColKey(Col)
            Case "产地"
                gstrSQL = "Select 编码 as id,名称,简码 From 药品生产商 Order By 编码 "
                    
                Set RecReturn = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
                    True, dblLeft, dblTop, .Height, blnCancel, False, True)
                
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
'                    .TextMatrix(Row, .ColIndex("产地id")) = RecReturn!id
                    .TextMatrix(Row, .ColIndex("产地")) = RecReturn!名称
                End If
            Case "药品名称"
                If grsMaster.State = adStateClosed Then
                    Call SetSelectorRS(1, "药品入库验收管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0)
                End If
                Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0, True, True, True)
                If RecReturn.RecordCount > 0 Then
                    Set RecReturn = CheckRedo(RecReturn) '检查重复记录，将重复的记录过滤掉然后返回过滤后的数据集
                End If
                
                If RecReturn.RecordCount > 0 Then
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        With vsfBill
                            intRow = .Row
                            SetColValue .Row, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                                RecReturn!药品id, _
                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), RecReturn!剂型, _
                                RecReturn!药库单位

                            .Col = .ColIndex("进药日期")
                                                    
                            If (.TextMatrix(intRow, .ColIndex("药品id")) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, .ColIndex("药品id")) <> "" Then
                                .rows = .rows + 1
                            End If
        
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        End With
                    Next
                    vsfBill.Row = intOldRow
                    RecReturn.Close
                End If
            Case "验收结论"
                gstrSQL = "Select 编码 as id, 编码, 名称 as 名称 From 入库验收结论 Order By 编码 "
                    
                Set RecReturn = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "验收结论", False, "", "", False, False, _
                    True, dblLeft, dblTop, .Height, blnCancel, False, True)
                
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
                    .TextMatrix(Row, .ColIndex("验收结论")) = RecReturn!名称
                End If
        End Select
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckRedo(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '功能：将重复的记录过滤掉，并返回过滤后的数据集合
    Dim i As Integer
    Dim strTemp As String
    Dim str药品id As String
    Dim str重复药名 As String
    Dim strDub As String
    Dim strsql As String
    
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If InStr(1, strTemp, rsTemp!药品id) = 0 Then
            strTemp = strTemp & rsTemp!药品id & "|"
        End If
        rsTemp.MoveNext
    Loop
    
    With vsfBill
        For i = 1 To .rows - 1
            If InStr(1, strTemp, .TextMatrix(i, .ColIndex("药品id"))) > 0 And .TextMatrix(i, .ColIndex("药品id")) <> "" Then
                str药品id = str药品id & .TextMatrix(i, .ColIndex("药品id")) & "," & .TextMatrix(i, .ColIndex("药品名称")) & "|"
            End If
        Next
        
        If str药品id <> "" Then   '为过滤数据拼接sql
            strDub = ""
            For i = 0 To UBound(Split(str药品id, "|")) - 1
                strDub = strDub & "药品id<>" & Split(Split(str药品id, "|")(i), ",")(0) & " and "
                If UBound(Split(str重复药名, ",")) <= 2 Then
                    str重复药名 = str重复药名 & Split(Split(str药品id, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        
        If str重复药名 <> "" Then
            MsgBox str重复药名 & "列表中已经含有了！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strsql = strDub
        End If
        rsTemp.Filter = strsql
        Set CheckRedo = rsTemp
    End With
End Function

Private Function SetColValue(ByVal intRow As Integer, _
    ByVal str药品编码 As String, _
    ByVal str通用名 As String, _
    ByVal str商品名 As String, _
    ByVal lng药品ID As Long, _
    ByVal str规格 As String, _
    ByVal str剂型 As String, _
    ByVal str单位 As String) As Boolean
    Dim str药名 As String
    Dim rsTemp As Recordset
    '将选择出来的药品添加到vsf表格
    '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
    On Error GoTo ErrHand
    With vsfBill
        .TextMatrix(intRow, .ColIndex("验收结果")) = "合格"
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = str通用名
        Else
            str药名 = IIf(str商品名 <> "", str商品名, str通用名)
        End If
                
        .TextMatrix(intRow, .ColIndex("药品名称")) = str药品编码 & str药名
        .TextMatrix(intRow, .ColIndex("药品id")) = lng药品ID
        .TextMatrix(intRow, .ColIndex("规格")) = str规格
        .TextMatrix(intRow, .ColIndex("剂型")) = str剂型
        .TextMatrix(intRow, .ColIndex("单位")) = str单位
        
        If mstr验收结论 = "" Then
            gstrSQL = "Select 名称  From 入库验收结论 where 缺省标志=1"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "SetColValue")
            
            If Not rsTemp.EOF Then
                .TextMatrix(intRow, .ColIndex("验收结论")) = IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
                mstr验收结论 = rsTemp!名称
            End If
        Else
            .TextMatrix(intRow, .ColIndex("验收结论")) = mstr验收结论
        End If
    End With
    
    SetColValue = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfBill_DblClick()
    With vsfBill
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            If .Col = .ColIndex("验收结果") And .TextMatrix(.Row, .ColIndex("药品id")) <> "" Then
                If .TextMatrix(.Row, .Col) = "合格" Then
                    .TextMatrix(.Row, .Col) = "不合格"
                Else
                    .TextMatrix(.Row, .Col) = "合格"
                End If
            End If
                
            If Not (.Col = .ColIndex("验收结果") Or .Col = .ColIndex("规格") Or .Col = .ColIndex("剂型") Or .Col = .ColIndex("单位")) Then
                .EditCell
                .EditSelStart = 0
                .EditSelLength = Len(.TextMatrix(.Row, .Col)) * 2
            End If
        End If
    End With
End Sub

Private Sub vsfBill_EnterCell()
    With vsfBill
        If .Col = .ColIndex("验收结果") Or .Col = .ColIndex("规格") Or .Col = .ColIndex("剂型") Or .Col = .ColIndex("单位") Then
            .Editable = flexEDNone
        Else
            If mint编辑状态 = 3 Or mint编辑状态 = 4 Then
                .Editable = flexEDNone
            Else
                .Editable = flexEDKbdMouse
            End If
            If .Col = .ColIndex("药品名称") Or .Col = .ColIndex("产地") Or .Col = .ColIndex("验收结论") Then
                .ColComboList(.Col) = ""
            End If
        End If
    End With
End Sub

Private Sub vsfBill_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If MsgBox("将删除此行，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Shift = 0
        Else
            vsfBill.RemoveItem vsfBill.Row
        End If
    ElseIf KeyCode = vbKeyReturn Then
        With vsfBill
            If .Col <> .Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row = .rows - 1 Then
                    If .TextMatrix(.Row, .ColIndex("药品id")) = "" Then
                        KeyCode = 0
                    Else
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                Else
                    .Row = .Row + 1
                    .Col = 1
                End If
            End If
        End With
    End If
End Sub



Private Sub vsfBill_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With vsfBill
        If .Col = .ColIndex("药品名称") Or .Col = .ColIndex("产地") Or .Col = .ColIndex("验收结论") Then
            .ColComboList(.Col) = "|..."
        End If
    End With
End Sub

Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim RecReturn As ADODB.Recordset
    Dim strkey As String
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intOldRow As Integer
    Dim i As Integer
    Dim intRow As Integer
    Dim intPosition As Integer
    
    On Error GoTo errHandle
    With vsfBill
        intOldRow = .Row
        
        strkey = UCase(.EditText)
                
        Select Case .ColKey(Col)
            Case "药品名称"
                If Trim(strkey) = "" Then Exit Sub
                If KeyAscii <> vbKeyReturn Then Exit Sub
                
                vRect = GetControlRect(vsfBill.hWnd) '获取位置
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top - vsfBill.Height + vsfBill.CellTop + vsfBill.CellHeight
                
                dblTop = dblTop + vsfBill.Height
                If strkey <> "" Then
                    If grsMaster.State = adStateClosed Then '获取数据集
                        Call SetSelectorRS(1, "药品入库验收管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0)
                    End If
                    Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, dblLeft, dblTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0, True, True, True)
                                                        
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckRedo(RecReturn) '检查重复记录 并将重复记录的药品id返回回来
                    End If
                                                        
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        
                        For i = 1 To RecReturn.RecordCount
                            intRow = .Row
                            If SetColValue(.Row, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                                RecReturn!药品id, _
                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), RecReturn!剂型, RecReturn!药库单位) = False Then
                                 KeyAscii = 0
                                 Exit Sub
                             End If
                            .EditText = .TextMatrix(.Row, .Col)
                            If (.TextMatrix(intRow, .ColIndex("药品id")) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, .ColIndex("药品id")) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        KeyAscii = 0
                    End If
                End If
            Case "产地"
                If Trim(strkey) = "" Then Exit Sub
                If KeyAscii <> vbKeyReturn Then Exit Sub
                vRect = GetControlRect(vsfBill.hWnd) '获取位置
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top - vsfBill.Height + vsfBill.CellTop + vsfBill.CellHeight
                
                gstrSQL = "Select 编码 as id,名称,简码" & _
                            " From 药品生产商" & _
                            " where 编码 Like [1] " & _
                            "       Or 名称 Like [2] " & _
                            "       Or 简码 Like [2] Order By 编码 "
                Set RecReturn = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "诊疗项目", False, "", "", False, False, _
                    True, dblLeft, dblTop, .Height, blnCancel, False, True, strkey & "%", mstrMatch & strkey & "%")
                If RecReturn Is Nothing Then
                    .Text = ""
                    Exit Sub
                Else
                    .Text = RecReturn!名称
                    .EditText = RecReturn!名称
                End If
            Case "验收结论"
                If Trim(strkey) = "" Then Exit Sub
                If KeyAscii <> vbKeyReturn Then Exit Sub
                vRect = GetControlRect(vsfBill.hWnd) '获取位置
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top - vsfBill.Height + vsfBill.CellTop + vsfBill.CellHeight
                
                gstrSQL = "Select 编码 as id, 编码, 名称" & _
                            " From 入库验收结论" & _
                            " where 编码 Like [1] " & _
                            "       Or 名称 Like [2] Order By 编码 "
                Set RecReturn = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "验收结论", False, "", "", False, False, _
                    True, dblLeft, dblTop, .Height, blnCancel, False, True, strkey & "%", mstrMatch & strkey & "%")
                If RecReturn Is Nothing Then
                    .Text = ""
                    Exit Sub
                Else
                    .Text = RecReturn!名称
                End If
            Case "进药日期", "生产日期", "效期"
                If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or (KeyAscii = vbKeyDelete Or KeyAscii = vbKeyReturn Or Chr(KeyAscii) = "-")) Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                    End If
                End If
            Case "成本价"
                If Not (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then
                    If InStr(1, strkey, ".") > 0 Then
                        If Chr(KeyAscii) = "." Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                        intPosition = InStr(1, strkey, ".") + 1
                        If Len(Mid(strkey, intPosition)) >= mintCostDigit Then
                            If strkey = .TextMatrix(.Row, .Col) Then
                                strkey = Chr(KeyAscii)
                            Else
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or Chr(KeyAscii) = ".") Then
                        If KeyAscii <> vbKeyBack Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        If Val(strkey + Chr(KeyAscii)) > 99999999 Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                End If
                If KeyAscii = vbKeyReturn Then
                    .EditText = zlStr.FormatEx(strkey, mintCostDigit, True, True)
                    .TextMatrix(.Row, .Col) = .EditText
                End If
            Case "零售价"
                If Not (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then
                    If InStr(1, strkey, ".") > 0 Then
                        If Chr(KeyAscii) = "." Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                        intPosition = InStr(1, strkey, ".") + 1
                        If Len(Mid(strkey, intPosition)) >= mintCostDigit Then
                            If strkey = .TextMatrix(.Row, .Col) Then
                                strkey = Chr(KeyAscii)
                            Else
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or Chr(KeyAscii) = ".") Then
                        If KeyAscii <> vbKeyBack Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        If Val(strkey + Chr(KeyAscii)) > 99999999 Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                End If
                If KeyAscii = vbKeyReturn Then
                    .EditText = zlStr.FormatEx(strkey, mintPriceDigit, True, True)
                    .TextMatrix(.Row, .Col) = .EditText
                End If
            Case "进药数量"
                If Not (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then
                    If InStr(1, strkey, ".") > 0 Then
                        If Chr(KeyAscii) = "." Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                        intPosition = InStr(1, strkey, ".") + 1
                        If Len(Mid(strkey, intPosition)) >= mintCostDigit Then
                            If strkey = .TextMatrix(.Row, .Col) Then
                                strkey = Chr(KeyAscii)
                            Else
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or Chr(KeyAscii) = ".") Then
                        If KeyAscii <> vbKeyBack Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        If Val(strkey + Chr(KeyAscii)) > 99999999 Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                End If
                If KeyAscii = vbKeyReturn Then
                    .EditText = zlStr.FormatEx(strkey, mintNumberDigit, True, True)
                    .TextMatrix(.Row, .Col) = .EditText
                End If
        End Select
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select f.Id, f.No, f.库房id, f.供药单位id, g.名称 As 供应商名称, f.验收人, f.验收日期, f.复核人, f.复核日期, f.备注, b.编码, b.Id As 药品id, b.名称, b.规格," & vbNewLine & _
                "       c.药库单位, c.药库包装, a.进药日期, e.名称 As 剂型, a.成本价, a.零售价, a.进药数量, a.批号, a.生产日期, a.效期, a.产地, a.批准文号, a.验收结论," & vbNewLine & _
                "       Nvl(a.是否合格, 0) As 是否合格" & vbNewLine & _
                "From 药品验收记录 F, 药品验收明细 A, 收费项目目录 B, 药品规格 C, 药品特性 D, 药品剂型 E, 供应商 G" & vbNewLine & _
                "Where f.Id = a.验收id And a.药品id = b.Id And b.Id = c.药品id And c.药名id = d.药名id And d.药品剂型 = e.名称(+) And f.供药单位id = g.Id(+) And" & vbNewLine & _
                "      a.验收id = [1]"

    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "验收明细查询", mlng验收id)
            
    With vsfBill
        Do While Not rsTemp.EOF
            txtNo.Text = rsTemp!NO
            .TextMatrix(.rows - 1, .ColIndex("验收结果")) = IIf(rsTemp!是否合格 = 0, "合格", "不合格")
            .TextMatrix(.rows - 1, .ColIndex("药品id")) = rsTemp!药品id
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
            .TextMatrix(.rows - 1, .ColIndex("验收结论")) = IIf(IsNull(rsTemp!验收结论), "", rsTemp!验收结论)
            
            txtProvider.Text = rsTemp!供应商名称
            txtProvider.Tag = rsTemp!供药单位ID
            
            If mint编辑状态 = 4 Then
                txtCheck.Text = rsTemp!验收人
            End If
            txtCheckDate.Caption = Format(rsTemp!验收日期, "yyyy-mm-dd hh:mm:ss")
            txt备注.Text = IIf(IsNull(rsTemp!备注), "", rsTemp!备注)
            
            If mint编辑状态 = 4 Then
                txtVerify.Text = IIf(IsNull(rsTemp!复核人), "", rsTemp!复核人)
                If IsNull(rsTemp!复核日期) = False Then
                    TxtVerifyDate.Caption = Format(rsTemp!复核日期, "yyyy-mm-dd hh:mm:ss")
                End If
            End If
            
            .rows = .rows + 1
            rsTemp.MoveNext
        Loop
        
        If .rows > 1 Then
            .Row = 1
            .Col = .ColIndex("药品名称")
        End If
    End With
End Sub

Private Sub vsfBill_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsfBill
        If .Col = .ColIndex("药品名称") Or .Col = .ColIndex("产地") Or .Col = .ColIndex("验收结论") Then
            .ColComboList(.Col) = "|..."
        Else
            .ColComboList(.ColIndex("药品名称")) = ""
            .ColComboList(.ColIndex("产地")) = ""
            .ColComboList(.ColIndex("验收结论")) = ""
        End If
    End With
End Sub

Private Sub vsfBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strkey As String
    
    On Error GoTo errHandle
    With vsfBill
        
        strkey = UCase(Trim(.Text))
'        If strkey = "" Then
            strkey = UCase(.EditText)
'        End If
        
        If Trim(strkey) = "" Then Exit Sub
        
        Select Case .ColKey(Col)
            Case "进药日期"
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "对不起，进药日期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(.Row, .ColIndex("进药日期")) = strkey
                Else
                    If Not IsDate(strkey) Then
                        MsgBox "对不起，进药日期必须为日期型如(2015-10-10) 或（20151010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Case "生产日期"
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "对不起，生产日期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(.Row, .ColIndex("生产日期")) = strkey
                Else
                    If Not IsDate(strkey) Then
                        MsgBox "对不起，生产日期必须为日期型如(2015-10-10) 或（20151010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Case "效期"
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "对不起，效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(.Row, .ColIndex("效期")) = strkey
                Else
                    If Not IsDate(strkey) Then
                        MsgBox "对不起，效期必须为日期型如(2015-10-10) 或（20151010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
        End Select
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


