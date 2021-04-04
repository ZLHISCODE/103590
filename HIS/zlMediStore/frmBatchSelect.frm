VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmBatchSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品批量选择"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12045
   DrawStyle       =   4  'Dash-Dot-Dot
   Icon            =   "frmBatchSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15644.17
   ScaleMode       =   0  'User
   ScaleWidth      =   14442.45
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picInit 
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   600
      Width           =   12015
      Begin VB.TextBox txtClass 
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   3495
      End
      Begin VB.TextBox txtPingZhong 
         Height          =   300
         Left            =   6000
         TabIndex        =   3
         Top             =   120
         Width           =   4020
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   2
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton cmdClass 
         Caption         =   "…"
         Height          =   300
         Left            =   4440
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   285
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSelectDrug 
         Height          =   6045
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   11895
         _cx             =   20981
         _cy             =   10663
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
         FormatString    =   $"frmBatchSelect.frx":000C
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
      Begin VB.Label lblPingZhong 
         AutoSize        =   -1  'True
         Caption         =   "品种简码"
         Height          =   180
         Left            =   5160
         TabIndex        =   8
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   6660
         Width           =   360
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         Caption         =   "分类简码"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   4680
      Top             =   6600
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
            Picture         =   "frmBatchSelect.frx":0081
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":061B
            Key             =   "pingzhong"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":6E7D
            Key             =   "规格U"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgList 
      Bindings        =   "frmBatchSelect.frx":7417
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmBatchSelect.frx":742B
   End
End
Attribute VB_Name = "frmBatchSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintUnit As Integer '本模块中设置的显示单位 0-药库单位;1-门诊单位;2-住院单位;3-售价单位
Private Const mlngRowHeight As Long = 300 '表格中各行行高
Private mrsReturn As ADODB.Recordset        '返回选定药品数据
Private mblnOK As Boolean   '记录是否是点击的确定按钮
Private mrsFindName As ADODB.Recordset '记录查询数据集
Private mstrMatch  As String '0-双向匹配 1-单向右匹配
Private mint进入模式  As Integer '0-调价进入，1-零差价批量调价进入


'各单位
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数
Private mstrMoneyFormat As String
Private mintSalePriceDigit As Integer
Private Const MStrCaption As String = "药品批量选择"

'功能按钮
Private Const mconMenu_Save = 100 '添加
Private Const mconMenu_Quit = 101 '取消
Private Const mconMenu_ClearAll = 102 '清空列表
Private Const mconMenu_Find = 103 '查找

Private Enum vsfSelectDrugCol
    药品id = 0
    药品信息 = 1
    药品编码
    商品名
    通用名
    规格
    产地
    单位
    售价单位
    门诊单位
    门诊系数
    住院单位
    住院系数
    药库单位
    药库系数
    类型
    售价
    成本价
    指导批价
    指导售价
    总列数
End Enum

Public Sub ShowME(ByVal frmParent As Form, ByRef rsTemp As ADODB.Recordset, ByRef blnOK As Boolean, Optional int进入模式 As Integer = 0)
    mint进入模式 = int进入模式
    Me.Show vbModal, frmParent
    blnOK = mblnOK
    Set rsTemp = mrsReturn
End Sub

Private Sub initVsflexgrid()
    With vsfSelectDrug
        .Editable = flexEDNone
        .Cols = vsfSelectDrugCol.总列数
        .rows = 1
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '不能多选
        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMove '移动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度

        '设置列宽
        .ColWidth(vsfSelectDrugCol.药品id) = 0
        .ColWidth(vsfSelectDrugCol.药品信息) = 3000
        .ColWidth(vsfSelectDrugCol.药品编码) = 0
        .ColWidth(vsfSelectDrugCol.商品名) = 0
        .ColWidth(vsfSelectDrugCol.通用名) = 0
        .ColWidth(vsfSelectDrugCol.产地) = 1500
        .ColWidth(vsfSelectDrugCol.单位) = 800

        .ColWidth(vsfSelectDrugCol.售价单位) = 0
        .ColWidth(vsfSelectDrugCol.门诊单位) = 0
        .ColWidth(vsfSelectDrugCol.门诊系数) = 0
        .ColWidth(vsfSelectDrugCol.住院单位) = 0
        .ColWidth(vsfSelectDrugCol.住院系数) = 0
        .ColWidth(vsfSelectDrugCol.药库单位) = 0
        .ColWidth(vsfSelectDrugCol.药库系数) = 0

        .ColWidth(vsfSelectDrugCol.类型) = 1000
        .ColWidth(vsfSelectDrugCol.售价) = 1500
        .ColWidth(vsfSelectDrugCol.成本价) = 1500
        .ColWidth(vsfSelectDrugCol.指导批价) = 1500
        .ColWidth(vsfSelectDrugCol.指导售价) = 1500
        '设置列头
        .TextMatrix(0, vsfSelectDrugCol.药品id) = "药品id"
        .TextMatrix(0, vsfSelectDrugCol.药品信息) = "药品"
        .TextMatrix(0, vsfSelectDrugCol.药品编码) = "药品编码"
        .TextMatrix(0, vsfSelectDrugCol.商品名) = "商品名"
        .TextMatrix(0, vsfSelectDrugCol.通用名) = "通用名"
        .TextMatrix(0, vsfSelectDrugCol.规格) = "规格"
        .TextMatrix(0, vsfSelectDrugCol.产地) = "生产商"
        .TextMatrix(0, vsfSelectDrugCol.单位) = "单位"

        .TextMatrix(0, vsfSelectDrugCol.售价单位) = "售价单位"
        .TextMatrix(0, vsfSelectDrugCol.门诊单位) = "门诊单位"
        .TextMatrix(0, vsfSelectDrugCol.门诊系数) = "门诊系数"
        .TextMatrix(0, vsfSelectDrugCol.住院单位) = "住院单位"
        .TextMatrix(0, vsfSelectDrugCol.住院系数) = "住院系数"
        .TextMatrix(0, vsfSelectDrugCol.药库单位) = "药库单位"
        .TextMatrix(0, vsfSelectDrugCol.药库系数) = "药库系数"

        .TextMatrix(0, vsfSelectDrugCol.类型) = "类型"
        .TextMatrix(0, vsfSelectDrugCol.售价) = "售价"
        .TextMatrix(0, vsfSelectDrugCol.成本价) = "成本价"
        .TextMatrix(0, vsfSelectDrugCol.指导批价) = "指导批价"
        .TextMatrix(0, vsfSelectDrugCol.指导售价) = "指导售价"

        .ColAlignment(vsfSelectDrugCol.药品id) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.药品信息) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.药品编码) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.规格) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.产地) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.单位) = flexAlignCenterCenter
        .ColAlignment(vsfSelectDrugCol.类型) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.售价) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.成本价) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.指导批价) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.指导售价) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case mconMenu_Save  '添加
            Call Save
        Case mconMenu_ClearAll  '清空
            Call ClearAll
        Case mconMenu_Find '查找
            txtFind.SetFocus
            If Trim(txtFind.Text) = "" Then Exit Sub
            Call FindGridRow(UCase(Trim(txtFind.Text)))
        Case mconMenu_Quit  '取消
            Call Quit
    End Select
End Sub

Private Sub ClearAll()
    With vsfSelectDrug
        If MsgBox("确定要清空所有已经选择的药品？", vbYesNo, gstrSysName) = vbYes Then
            .rows = 1
        End If
    End With
End Sub

Private Sub cmdClass_Click()
    Dim rsProvider As Recordset
    Dim strsql零差价模式 As String
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txtClass.hWnd)
    On Error GoTo ErrHandle
    
    If mint进入模式 = 1 Then
        strsql零差价模式 = " a.是否零差价管理=1 And "
    Else
        strsql零差价模式 = ""
    End If
    
    '分类
    gstrSQL = "Select Level, ID, 上级id, 编码, 名称, 分类" & vbNewLine & _
                    "From (Select -1 As ID, Null As 上级id, '001' As 编码, '西成药' As 名称, '西成药' As 分类" & vbNewLine & _
                    "       From Dual" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select -2 As ID, Null As 上级id, '002' As 编码, '中成药' As 名称, '中成药' As 分类" & vbNewLine & _
                    "       From Dual" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select -3 As ID, Null As 上级id, '003' As 编码, '中草药' As 名称, '中草药' As 分类" & vbNewLine & _
                    "       From Dual" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select ID, Nvl(上级id, Decode(类型, 1, -1, Decode(类型, 2, -2, -3))) As 上级id, 编码, 名称," & vbNewLine & _
                    "              Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 分类" & vbNewLine & _
                    "       From 诊疗分类目录" & vbNewLine & _
                    "       Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01')" & vbNewLine & _
                    "Start With 上级id Is Null" & vbNewLine & _
                    "Connect By Prior ID = 上级id" & vbNewLine & _
                    "Order By Level"

    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "分类", False, "", "调价选择", False, False, _
    True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider Is Nothing Then
        Exit Sub
    End If
    
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select Distinct a.药品id, c.编码 As 药品编码, c.名称 As 通用名, d.商品名, c.规格, c.是否变价 As 时价, c.产地, c.计算单位 As 售价单位, a.门诊单位, a.门诊包装," & vbNewLine & _
                "                a.住院单位, a.住院包装, a.药库单位, a.药库包装, a.成本价, e.现价, a.指导批发价, a.指导零售价" & vbNewLine & _
                "From 药品规格 A, 诊疗项目目录 B, 收费项目目录 C, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) D, 收费价目 E" & vbNewLine & _
                "Where a.药名id = b.Id And a.药品id = c.Id And c.Id = d.收费细目id(+) And a.药品id = e.收费细目id And Sysdate Between e.执行日期 And" & vbNewLine & _
                "      e.终止日期 And (c.撤档时间 = to_date('3000-01-01','yyyy-mm-dd') or c.撤档时间 is null ) " & GetPriceClassString("E") & _
                "And " & strsql零差价模式 & "b.分类id In (Select ID" & vbNewLine & _
                "                            From 诊疗分类目录" & vbNewLine & _
                "                            Where 类型 In (1, 2, 3) And Nvl(To_Char(撤档时间, 'Yyyy-Mm-Dd'), '3000-01-01') = '3000-01-01'" & vbNewLine & _
                "                            Start With ID = [1]" & vbNewLine & _
                "                            Connect By Prior ID = 上级id)" & vbNewLine & _
                "Order By c.编码"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "分类", rsProvider!id)
    
    If rsTemp.RecordCount = 0 And rsProvider!id > 0 Then
        If mint进入模式 = 1 Then
            MsgBox "没有找到该分类下零差价管理的药品！", vbInformation, gstrSysName
            Exit Sub
        Else
            MsgBox "没有找到该分类下的药品！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
           
    Call GetDetails(rsTemp)
        
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save()
    Dim intRow As Integer
    Set mrsReturn = New ADODB.Recordset

    With mrsReturn
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品编码", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "通用名", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "时价", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 40, adFldIsNullable

        .Fields.Append "售价单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "门诊单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "门诊包装", adDouble, 11, adFldIsNullable
        .Fields.Append "住院单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "住院包装", adDouble, 11, adFldIsNullable
        .Fields.Append "药库单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "药库包装", adDouble, 11, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    With vsfSelectDrug
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, vsfSelectDrugCol.药品id) = "" Then Exit For
            mrsReturn.AddNew
            mrsReturn!药品id = .TextMatrix(intRow, vsfSelectDrugCol.药品id)
            mrsReturn!药品编码 = .TextMatrix(intRow, vsfSelectDrugCol.药品编码)
            mrsReturn!商品名 = .TextMatrix(intRow, vsfSelectDrugCol.商品名)
            mrsReturn!通用名 = .TextMatrix(intRow, vsfSelectDrugCol.通用名)
            mrsReturn!规格 = .TextMatrix(intRow, vsfSelectDrugCol.规格)
            mrsReturn!时价 = IIf(.TextMatrix(intRow, vsfSelectDrugCol.类型) = "时价", 1, 0)
            mrsReturn!产地 = .TextMatrix(intRow, vsfSelectDrugCol.产地)
            mrsReturn!售价单位 = .TextMatrix(intRow, vsfSelectDrugCol.售价单位)
            mrsReturn!门诊单位 = .TextMatrix(intRow, vsfSelectDrugCol.门诊单位)
            mrsReturn!门诊包装 = .TextMatrix(intRow, vsfSelectDrugCol.门诊系数)
            mrsReturn!住院单位 = .TextMatrix(intRow, vsfSelectDrugCol.住院单位)
            mrsReturn!住院包装 = .TextMatrix(intRow, vsfSelectDrugCol.住院系数)
            mrsReturn!药库单位 = .TextMatrix(intRow, vsfSelectDrugCol.药库单位)
            mrsReturn!药库包装 = .TextMatrix(intRow, vsfSelectDrugCol.药库系数)

            mrsReturn.Update
        Next
    End With
    mblnOK = True

    Unload Me
End Sub

Private Sub Quit()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intUnitTemp As Integer
    
    '获取设置的单位
    mintUnit = Val(zlDataBase.GetPara("药品单位", glngSys, 1333, 1))
    Select Case mintUnit
        Case 0 '药库
            intUnitTemp = 4
        Case 1 '住院
            intUnitTemp = 3
        Case 2 '门诊
            intUnitTemp = 2
        Case 3 '售价
            intUnitTemp = 1
    End Select
    '获取各级单位精度
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
    mintNumberDigit = GetDigitTiaoJia(1, 3, intUnitTemp)
    mintMoneyDigit = GetDigitTiaoJia(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    mintSalePriceDigit = GetDigitTiaoJia(1, 2, 1)

    mstrMatch = IIf(zlDataBase.GetPara("输入匹配", , , 0) = "0", "%", "")
    mblnOK = False
    Call initCommandBars
    Call initVsflexgrid
    Call RestoreWinState(Me, App.ProductName, MStrCaption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, MStrCaption)
End Sub

Private Sub txtClass_GotFocus()
    zlControl.TxtSelAll txtClass
End Sub

Private Sub txtClass_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As Recordset
    Dim rsTemp As ADODB.Recordset
    Dim strsql零差价模式 As String
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtClass.hWnd)
    On Error GoTo ErrHandle
    
    If KeyCode = vbKeyReturn Then
    
        If mint进入模式 = 1 Then
            strsql零差价模式 = " a.是否零差价管理=1 And "
        Else
            strsql零差价模式 = ""
        End If
        '分类
        
        If Trim(txtClass.Text) = "" Then Exit Sub
        
        gstrSQL = "Select id,编码,名称" & vbNewLine & _
                    "From 诊疗分类目录" & vbNewLine & _
                    "Where 类型 In (1, 2, 3) And (Sysdate Between 建档时间 And 撤档时间 Or 撤档时间 Is Null) And" & vbNewLine & _
                    "      (编码 Like '" & "%" & UCase(txtClass.Text) & mstrMatch & "' Or 名称 Like '" & "%" & UCase(txtClass.Text) & mstrMatch & "' Or" & vbNewLine & _
                    "       简码 Like '" & "%" & UCase(txtClass.Text) & mstrMatch & "')" & vbNewLine & _
                    "Order By ID"
    
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "分类", False, "", "调价选择", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
        
        If blnCancel = True Then Exit Sub '打开选择器时，点Esc不做以下处理
        
        If rsProvider Is Nothing Then
            MsgBox "没有找到该分类下的药品，请重输！", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            Exit Sub
        End If
        
        gstrSQL = "Select Distinct a.药品id, c.编码 As 药品编码, c.名称 As 通用名, d.商品名, c.规格, c.是否变价 As 时价, c.产地, c.计算单位 As 售价单位, a.门诊单位, a.门诊包装," & vbNewLine & _
                    "                a.住院单位, a.住院包装, a.药库单位, a.药库包装, a.成本价, e.现价, a.指导批发价, a.指导零售价" & vbNewLine & _
                    "From 药品规格 A, 诊疗项目目录 B, 收费项目目录 C, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) D, 收费价目 E" & vbNewLine & _
                    "Where a.药名id = b.Id And a.药品id = c.Id And c.Id = d.收费细目id(+) And a.药品id = e.收费细目id And Sysdate Between e.执行日期 And" & vbNewLine & _
                    "      e.终止日期  And (c.撤档时间 = to_date('3000-01-01','yyyy-mm-dd') or c.撤档时间 is null ) " & GetPriceClassString("E") & _
                    " And " & strsql零差价模式 & "b.分类id In (Select ID" & vbNewLine & _
                    "                            From 诊疗分类目录" & vbNewLine & _
                    "                            Where 类型 In (1, 2, 3) And Nvl(To_Char(撤档时间, 'Yyyy-Mm-Dd'), '3000-01-01') = '3000-01-01'" & vbNewLine & _
                    "                            Start With ID = [1]" & vbNewLine & _
                    "                            Connect By Prior ID = 上级id)" & vbNewLine & _
                    "Order By c.编码"
        
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "分类", rsProvider!id)
        
        If rsTemp.RecordCount = 0 Then
            If mint进入模式 = 1 Then
                MsgBox "没有找到该分类下零差价管理的药品！", vbInformation, gstrSysName
                Exit Sub
            Else
                MsgBox "没有找到该分类下的药品！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        Call GetDetails(rsTemp)
    End If
        
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub

    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim str药名 As String
    Dim lngRow As Long

    '查找药品
    On Error GoTo ErrHandle
    If strInput <> txtFind.Tag Then
        '表示新的查找
        txtFind.Tag = strInput

        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别 In ('5','6','7') " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) " & _
                  "Order By 药品编码 "
        Set mrsFindName = zlDataBase.OpenSQLRecord(gstrSQL, "取匹配的药品ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = mrsFindName!药品编码 & mrsFindName!通用名
        Else
            str药名 = mrsFindName!药品编码 & IIf(IsNull(mrsFindName!商品名), mrsFindName!通用名, mrsFindName!商品名)
        End If

        For lngRow = 1 To vsfSelectDrug.rows - 1
            lngFindRow = vsfSelectDrug.FindRow(str药名, lngRow, CLng(vsfSelectDrugCol.药品信息), True, True)
            If lngFindRow > 0 Then
                vsfSelectDrug.Select lngFindRow, 1, lngFindRow, vsfSelectDrug.Cols - 1
                vsfSelectDrug.TopRow = lngFindRow
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtPingZhong_GotFocus()
    zlControl.TxtSelAll txtPingZhong
End Sub

Private Sub txtPingZhong_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As Recordset
    Dim rsTemp As ADODB.Recordset
    Dim strsql零差价模式 As String
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtPingZhong.hWnd)
    On Error GoTo ErrHandle
    
    If KeyCode = vbKeyReturn Then

        If mint进入模式 = 1 Then
            strsql零差价模式 = " a.是否零差价管理=1 And "
        Else
            strsql零差价模式 = ""
        End If
        
        If Trim(txtPingZhong.Text) = "" Then Exit Sub

        gstrSQL = "Select Distinct a.id,a.编码,a.名称" & vbNewLine & _
                  "  From 诊疗项目目录 A, 诊疗项目别名 B" & vbNewLine & _
                    " Where a.Id = b.诊疗项目id(+) And a.类别 In ('5', '6', '7') And Sysdate Between 建档时间 And 撤档时间 And " & vbNewLine & _
                         " (a.编码 Like '" & "%" & UCase(txtPingZhong.Text) & mstrMatch & "' Or " & vbNewLine & _
                         "a.名称 Like '" & "%" & UCase(txtPingZhong.Text) & mstrMatch & "'  Or " & vbNewLine & _
                         "b.名称 Like '" & "%" & UCase(txtPingZhong.Text) & mstrMatch & "'  Or " & vbNewLine & _
                         "b.简码 Like '" & "%" & UCase(txtPingZhong.Text) & mstrMatch & "' )"
    
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "分类", False, "", "调价选择", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
        
        If blnCancel = True Then Exit Sub '打开选择器时，点Esc不做以下处理
        
        If rsProvider Is Nothing Then
            MsgBox "没有找到该品种下的药品，请重输！", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            Exit Sub
        End If
        
        gstrSQL = "Select Distinct a.药品id, c.编码 As 药品编码, c.名称 As 通用名, d.商品名, c.规格, c.是否变价 As 时价, c.产地, c.计算单位 As 售价单位, a.门诊单位, a.门诊包装," & _
                                  " a.住院单位 , a.住院包装, a.药库单位, a.药库包装, a.成本价, e.现价, a.指导批发价, a.指导零售价" & _
                  " From 药品规格 A, 诊疗项目目录 B, 收费项目目录 C, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) D,收费价目 E" & _
                  " Where a.药名id = b.Id And a.药品id = c.Id And c.Id = d.收费细目id(+) and a.药品id=e.收费细目id and sysdate between e.执行日期 and e.终止日期  And (c.撤档时间 = to_date('3000-01-01','yyyy-mm-dd') or c.撤档时间 is null ) " & _
                  GetPriceClassString("E") & " and " & strsql零差价模式 & "b.id=[1] order by c.编码"
                  
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "品种", rsProvider!id)
               
        If rsTemp.RecordCount = 0 Then
            If mint进入模式 = 1 Then
                MsgBox "没有找到该品种下零差价管理的药品！", vbInformation, gstrSysName
                Exit Sub
            Else
                MsgBox "没有找到该品种下的药品！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
               
        Call GetDetails(rsTemp)
        
    End If

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetDetails(ByVal rsTemp As ADODB.Recordset)
    Dim lngID As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim blnDou As Boolean '重复数据
    Dim dbl换算系数 As Double
    Dim strUnit As String   '单位

    With vsfSelectDrug
        For intRow = 0 To rsTemp.RecordCount - 1
            blnDou = False
            For i = 1 To .rows - 1
                If .TextMatrix(i, vsfSelectDrugCol.药品id) = rsTemp!药品id Then
                    blnDou = True
                End If
            Next
            If blnDou = False Then
                .rows = .rows + 1
                .RowHeight(.rows - 1) = mlngRowHeight

                Select Case mintUnit
                    Case 0
                        dbl换算系数 = rsTemp!药库包装
                        strUnit = rsTemp!药库单位
                    Case 1
                        dbl换算系数 = rsTemp!住院包装
                        strUnit = rsTemp!住院单位
                    Case 2
                        dbl换算系数 = rsTemp!门诊包装
                        strUnit = rsTemp!门诊单位
                    Case 3
                        dbl换算系数 = 1
                        strUnit = rsTemp!售价单位
                End Select

                .TextMatrix(.rows - 1, vsfSelectDrugCol.药品id) = rsTemp!药品id
                If gint药品名称显示 = 1 Then
                    .TextMatrix(.rows - 1, vsfSelectDrugCol.药品信息) = "[" & rsTemp!药品编码 & "]" & IIf(IsNull(rsTemp!商品名), rsTemp!通用名, rsTemp!商品名)
                Else
                    .TextMatrix(.rows - 1, vsfSelectDrugCol.药品信息) = "[" & rsTemp!药品编码 & "]" & rsTemp!通用名
                End If

                .TextMatrix(.rows - 1, vsfSelectDrugCol.药品编码) = rsTemp!药品编码
                .TextMatrix(.rows - 1, vsfSelectDrugCol.商品名) = IIf(IsNull(rsTemp!商品名), "", rsTemp!商品名)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.通用名) = IIf(IsNull(rsTemp!通用名), "", rsTemp!通用名)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.单位) = strUnit

                .TextMatrix(.rows - 1, vsfSelectDrugCol.售价单位) = rsTemp!售价单位
                .TextMatrix(.rows - 1, vsfSelectDrugCol.门诊单位) = rsTemp!门诊单位
                .TextMatrix(.rows - 1, vsfSelectDrugCol.门诊系数) = rsTemp!门诊包装
                .TextMatrix(.rows - 1, vsfSelectDrugCol.住院单位) = rsTemp!住院单位
                .TextMatrix(.rows - 1, vsfSelectDrugCol.住院系数) = rsTemp!住院包装
                .TextMatrix(.rows - 1, vsfSelectDrugCol.药库单位) = rsTemp!药库单位
                .TextMatrix(.rows - 1, vsfSelectDrugCol.药库系数) = rsTemp!药库包装


                .TextMatrix(.rows - 1, vsfSelectDrugCol.类型) = IIf(rsTemp!时价 = 1, "时价", "定价")
                .TextMatrix(.rows - 1, vsfSelectDrugCol.售价) = zlStr.FormatEx(dbl换算系数 * rsTemp!现价, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.成本价) = zlStr.FormatEx(dbl换算系数 * rsTemp!成本价, mintCostDigit, , True)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.指导批价) = zlStr.FormatEx(dbl换算系数 * rsTemp!指导批发价, mintCostDigit, , True)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.指导售价) = zlStr.FormatEx(dbl换算系数 * rsTemp!指导零售价, mintPriceDigit, , True)

            End If
            rsTemp.MoveNext
        Next
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub initCommandBars()
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    Dim cbrControlPopu As CommandBarControl
    Dim lngCount As Integer
    
    With CommandBarsGlobalSettings
        .App = App
        .CompanyName = "重庆中联信息产业有限责任公司" '公司名称
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '设置中文语言资源文件
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '控件整体的颜色方案
    End With

    With cbsMain.Options
        .ShowExpandButtonAlways = False '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True '显示按钮提示
        .AlwaysShowFullMenus = False '不常用的菜单项先隐藏
        .UseFadedIcons = True '图标显示为褪色效果
        .IconsWithShadow = True '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True '工具栏显示为大图标
        .SetIconSize True, 24, 24 '设置大图标的尺寸
        .SetIconSize False, 16, 16 '设置小图标的尺寸
    End With

    With cbsMain
        .VisualTheme = xtpThemeOffice2003 '设置控件显示风格
        .EnableCustomization False '是否允许自定义设置
        Set .Icons = imgList.Icons '设置关联的图标控件
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap '窗体变化时，如果显示不完菜单也不换行
        .ActiveMenuBar.Title = "菜单"
    End With
    
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 1 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '创建工具栏
    Set cbrToolBar = cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    cbrToolBar.ContextMenuPresent = False

    With cbrToolBar
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_ClearAll, "清空")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Find, "查找")
        cbrControl.Visible = False

        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Save, "添加")
        cbrControl.BeginGroup = True
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Quit, "取消")
                
    End With

    For Each cbrControl In cbrToolBar.Controls  '让工具栏中按钮同时显示图标和文字
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F3, mconMenu_Find
    End With

End Sub

