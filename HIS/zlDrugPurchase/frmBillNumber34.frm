VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBillNumber34 
   Caption         =   "配送单号导入"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15855
   Icon            =   "frmBillNumber34.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   15855
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   14640
      TabIndex        =   3
      Top             =   7695
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "生成入库单据(&O)"
      Height          =   350
      Left            =   12960
      TabIndex        =   2
      Top             =   7695
      Width           =   1575
   End
   Begin VB.PictureBox pic发票信息 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   7740
      Width           =   260
   End
   Begin VB.PictureBox pic供应商颜色 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6480
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   7740
      Width           =   260
   End
   Begin VB.Frame fra导入库房 
      Caption         =   "导入库房"
      Height          =   675
      Left            =   9120
      TabIndex        =   10
      Top             =   840
      Width           =   3345
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fra提取数据条件 
      Caption         =   "提取数据条件"
      Height          =   675
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   8505
      Begin VB.CommandButton cmd提取补录发票信息 
         Caption         =   "提取补录发票信息(&P)"
         Height          =   350
         Left            =   6360
         TabIndex        =   16
         Top             =   215
         Width           =   2055
      End
      Begin VB.CommandButton cmd提取数据 
         Caption         =   "提取配送单数据(&T)"
         Height          =   350
         Left            =   4440
         TabIndex        =   7
         Top             =   215
         Width           =   1695
      End
      Begin VB.TextBox txt配送单号 
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   3075
      End
      Begin VB.Label lbl配送单号 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "配送单号"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame Frmline1 
      Height          =   120
      Left            =   120
      TabIndex        =   1
      Top             =   645
      Width           =   15735
   End
   Begin VB.Frame Frmline2 
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   15735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   5685
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   15540
      _cx             =   27411
      _cy             =   10028
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   37
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBillNumber34.frx":6852
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
      ExplorerBar     =   5
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
   Begin VB.Label lbl产生单据数 
      AutoSize        =   -1  'True
      Caption         =   "提示：共有0张单据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   17
      Top             =   7770
      Width           =   1785
   End
   Begin VB.Label lbl发票信息 
      AutoSize        =   -1  'True
      Caption         =   "发票信息不正确"
      Height          =   180
      Left            =   4680
      TabIndex        =   15
      Top             =   7785
      Width           =   1260
   End
   Begin VB.Label lbl供应商颜色 
      AutoSize        =   -1  'True
      Caption         =   "供应商信息不正确"
      Height          =   180
      Left            =   6840
      TabIndex        =   13
      Top             =   7785
      Width           =   1440
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   255
      Picture         =   "frmBillNumber34.frx":6D9C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInfor 
      Caption         =   $"frmBillNumber34.frx":70A6
      Height          =   405
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Width           =   10695
   End
End
Attribute VB_Name = "frmBillNumber34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mCanColColor As Long = &H8080FF           '供应商信息有误的用浅红色标识
Private Const mNoColColor As Long = &H80000012           '供应商信息正确的为黑色
Private Const mFPColColor As Long = &HFF80FF              '发票信息有误的用浅紫色标识
Private mint数据库类型 As Integer              '0:SQLserver数据库；1：Oracle数据库
Private mblnIsConn As Boolean
Private marrSql As Variant

Private Sub cbo库房_Click()
    If Val(cbo库房.ListIndex) <> Val(cbo库房.Tag) And vsfList.Rows > 1 Then
        If MsgBox("如果改变库房，需要重新提取单据内容，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, GSTR_MESSAGE) = vbYes Then
            vsfList.Rows = 1
            lbl产生单据数.Caption = "提示：共有0张单据"
        Else
            cbo库房.ListIndex = Val(cbo库房.Tag)
        End If
    End If
    cbo库房.Tag = Val(cbo库房.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    Dim blnTrans As Boolean
    
    If vsfList.Rows < 2 Then Exit Sub
    If cbo库房.ItemData(cbo库房.ListIndex) = -1 Then
        MsgBox "请选择库房！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    On Error GoTo ErrHand

    If MsgBox("是否确定导入？", vbQuestion + vbYesNo + vbDefaultButton2, GSTR_MESSAGE) = vbNo Then
        Exit Sub
    End If
    
    If cmdOk.Tag = "新增入库单" Then
        Call Save新增入库单
    ElseIf cmdOk.Tag = "补录发票信息" Then
        Call Save未审核发票信息
        Call Save已审核发票信息
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(marrSql)
        Call gobjComLib.zlDatabase.ExecuteProcedure(CStr(marrSql(i)), "保存外购入库单")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    marrSql = Array()
    MsgBox "导入成功！", vbInformation, GSTR_MESSAGE
    vsfList.Rows = 1

    lbl产生单据数.Caption = "提示：共有0张单据"
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = vbDefault
    MsgBox "导入失败，请检查！", vbInformation, GSTR_MESSAGE
End Sub

Private Sub cmd提取补录发票信息_Click()
    Dim strSQL As String
    Dim rs发票信息 As New ADODB.Recordset
    Dim rs补录发票信息 As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If mblnIsConn = False Then
        MsgBox "请先设置中间数据库连接！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

    If Trim(txt配送单号.Text) = "" Then
        MsgBox "请先录入配送单号！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    strSQL = "Select Distinct g.配送单号 , a.No, a.药品id, a.序号, d.编码 As 药品编码,d.名称, d.规格, d.计算单位 as 单位,a.产地 as 生产商,  a.批号 as 批号, a.效期 as 有效日期, a.填写数量 As 数量, a.成本价 as 单价, a.成本金额 As 金额," & vbNewLine & _
                    "                Nvl(a.批次, 0) 批次,d.是否变价, Decode(a.扣率, Null, 0, a.扣率) As 扣率, a.零售价 As 零售价, a.零售金额, a.差价, a.批准文号, c.随货单号, c.发票号," & vbNewLine & _
                    "                c.发票代码, c.发票日期, c.发票金额, a.供药单位id, f.编码 As 供应商编码, f.名称 As 供应商, a.摘要 , a.产品合格证 , a.填制人, a.填制日期,  a.审核人," & vbNewLine & _
                    "                a.审核日期, a.库房id, a.外观, a.验收结论, a.生产日期, a.配药人 As 核查人, a.配药日期 As 核查日期, Nvl(a.用法, 0) As 金额差, a.频次 As 加成率," & vbNewLine & _
                    "                a.对方部门id, a.计划id " & vbNewLine & _
                    "From 药品收发记录 A, 药品规格 B, 收费项目目录 D, 应付记录 C, 供应商 F, 配送单号对照 G " & vbNewLine & _
                    "Where a.药品id = b.药品id And b.药品id = d.Id And a.供药单位id = f.Id And a.库房id=[2] And Substr(f.类型, 1, 1) = 1 And a.Id = c.收发id(+) And" & vbNewLine & _
                    "      c.系统标识(+) = 1 And c.记录性质(+) = 0 And a.记录状态 = 1 And a.单据 = 1 And a.No = g.No And g.单据 = 1" & vbNewLine & _
                    "      And Not Exists (Select 1 From 应付记录 Where ID = c.Id And 付款序号 Is Not Null) And g.配送单号 =[1] order by a.no , a.药品id"

    Set rs补录发票信息 = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询药品信息", txt配送单号.Text, Val(cbo库房.ItemData(cbo库房.ListIndex)))
    

    strSQL = "SELECT FPHM 发票号,PSQYBM 供应商编码,KPRQ 发票日期,GLDJLX 性质," & _
                    "GLMXBH 配送单号,YPBM 药品编码,DJ 发票单价,SL 发票数量,JE 发票金额 " & _
                    "from mid_invoice where  GLDJLX='1' and GLMXBH='" & txt配送单号.Text & "'"
    
    rs发票信息.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
    
'    If mint数据库类型 = 0 Then
'        'SQLserver数据库
'        rs发票信息.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
'    Else
'        'Oracle数据库
'        Set rs发票信息 = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询发票信息")
'    End If
    
    If rs补录发票信息.RecordCount > 0 And rs发票信息.RecordCount > 0 Then
        cmdOk.Caption = "导入发票信息(&O)"
        cmdOk.Tag = "补录发票信息"
        Call DataVsf(rs补录发票信息, rs发票信息)
    Else
        MsgBox "没有查询到发票数据，请检查！", vbInformation, GSTR_MESSAGE
        vsfList.Rows = 1
        lbl产生单据数.Caption = "提示：共有0张单据"
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    If mblnIsConn = False Then
        MsgBox "请先设置中间数据库连接！", vbInformation, GSTR_MESSAGE
    Else
        MsgBox "获取外部数据错误！", vbInformation, GSTR_MESSAGE
    End If
    vsfList.Rows = 1
End Sub

Private Sub cmd提取数据_Click()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim rs发票信息 As New ADODB.Recordset
    Dim rs检查配送单信息 As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If mblnIsConn = False Then
        MsgBox "请先设置中间数据库连接！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

    If Trim(txt配送单号.Text) = "" Then
        MsgBox "请先录入配送单号！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

    strSQL = "Select 1" & vbNewLine & _
                    "From 配送单号对照 A, 药品收发记录 B" & vbNewLine & _
                    "Where a.No = b.No And a.单据 = 1 And a.配送单号 = [1] And Rownum < 2"
                    
    Set rs检查配送单信息 = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询配送单是否已经产生过", Trim(txt配送单号.Text))
        
    If rs检查配送单信息.RecordCount > 0 Then
        If MsgBox("配送单号[" & Trim(txt配送单号.Text) & "]已经产生过入库记录，是否继续提取数据？", vbQuestion + vbYesNo + vbDefaultButton2, GSTR_MESSAGE) = vbNo Then
            Exit Sub
        End If
    End If
    
    strSQL = "select KFBM 库房id,PSQYBM 供应商编码,PSDH 配送单号,YPBM 药品编码,SCPH 批号,SCRQ 生产日期," _
                  & " YXRQ 有效日期,DJ 单价,SL 数量,JLDW 单位,JE 金额,SCQY 生产商 " _
                  & " From MID_DELIVERY_ORDER  where PSDH='" & txt配送单号.Text & "' order by 供应商编码,药品编码"
    
    rsTmp.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
    
'    If mint数据库类型 = 0 Then
'        'SQLserver数据库
'        rsTmp.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
'    Else
'        'Oracle数据库
'        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询药品信息")
'    End If
    
    strSQL = "SELECT FPHM 发票号,PSQYBM 供应商编码,KPRQ 发票日期,GLDJLX 性质," & _
                    "GLMXBH 配送单号,YPBM 药品编码,DJ 发票单价,SL 发票数量,JE 发票金额 " & _
                    "from mid_invoice where  GLDJLX='1' and GLMXBH='" & txt配送单号.Text & "'"
    
    rs发票信息.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
    
'    If mint数据库类型 = 0 Then
'        'SQLserver数据库
'        rs发票信息.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
'    Else
'        'Oracle数据库
'        Set rs发票信息 = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询发票信息")
'    End If
    
    If rsTmp.RecordCount > 0 Then
        cmdOk.Caption = "生成入库单据(&O)"
        cmdOk.Tag = "新增入库单"
        Call DataVsf(rsTmp, rs发票信息)
    Else
        MsgBox "没有查询到数据，请检查！", vbInformation, GSTR_MESSAGE
        vsfList.Rows = 1
        lbl产生单据数.Caption = "提示：共有0张单据"
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    If mblnIsConn = False Then
        MsgBox "请先设置中间数据库连接！", vbInformation, GSTR_MESSAGE
    Else
        MsgBox "获取外部数据错误！", vbInformation, GSTR_MESSAGE
    End If
    vsfList.Rows = 1
End Sub

Private Sub Form_Load()
    vsfList.AllowSelection = False '不能多选
    vsfList.Rows = 1
    Call GetUserNameInfo
    Call SetMedicalWH
    Call ConnectDatabase
    marrSql = Array()
    lbl产生单据数.Caption = "提示：共有0张单据"
End Sub

Public Function GetUserNameInfo() As Boolean
'获取用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = gobjComLib.zlDatabase.GetUserInfo
    
    With rsTmp
        If Not .EOF Then
            glngUserID = IIf(IsNull(!Id), 0, !Id)
            glngDeptID = IIf(IsNull(!部门id), 0, !部门id)
            gstrUserNameNew = IIf(IsNull(!姓名), "", !姓名) '当前用户姓名
            GetUserNameInfo = True
        Else
            glngUserID = 0
            glngDeptID = 0
            gstrUserNameNew = "" '当前用户姓名
        End If
    End With
    rsTmp.Close

    strSQL = "Select 参数号, 参数值, 缺省值 From Zlparameters Where 系统 = [1] And Nvl(私有, 0) = 0 And 模块 Is Null and 参数号=[2] "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "取系统参数", 100, 149)
    With rsTmp
        If Not .EOF Then
            gbyt效期 = IIf(IsNull(rsTmp!参数值), rsTmp!缺省值, rsTmp!参数值)
        Else
            gbyt效期 = 0
        End If
    End With
    
End Function

Private Sub SetMedicalWH()
'设置药库combobox信息，同HIS规则，用户要和HIS的部门权限一样。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i, j As Integer
    Dim strStock As String
    
    If InStr(1, gstrPrivs, "允许药房外购入库") = 0 Then
        strStock = "HIJ"
    Else
        strStock = "HIJKLMN"
    End If
            
    '药库信息
    strSQL = "SELECT DISTINCT a.id, a.名称 " _
            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = '-' Or a.站点 is Null) And c.工作性质 = b.名称 " _
            & "  AND Instr([2],b.编码,1) > 0 " _
            & "  AND a.id = c.部门id " _
            & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
            & IIf(InStr(1, gstrPrivs, "所有库房") > 0, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])") _
            & " order by a.id"
            
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngUserID, strStock)
    
    cbo库房.Clear
    For i = 0 To rsTmp.RecordCount - 1
        cbo库房.AddItem rsTmp!名称
        cbo库房.ItemData(i) = rsTmp!Id
        rsTmp.MoveNext
    Next
    cbo库房.Tag = IIf(gintListIndex = -1, 0, gintListIndex)
    cbo库房.ListIndex = IIf(gintListIndex = -1, 0, gintListIndex)
    rsTmp.Close
End Sub


Private Sub DataVsf(ByVal rsVal As ADODB.Recordset, ByVal rs发票信息 As ADODB.Recordset)
'填充表格数据
    Dim i As Integer
    Dim str供应商信息 As String
    Dim str供应商 As String
    Dim strSQL As String
    Dim lng供应商id As Long
    Dim lng药品ID As Long
    Dim str药品名称 As String
    Dim str药品规格 As String
    Dim rs药品信息 As New ADODB.Recordset
    Dim dbl零售价 As Double
    Dim int变价 As Integer
    Dim int产生单据数量 As Integer
    Dim str供应商id As String
    Dim strNO串 As String
    
    On Error GoTo ErrHand
    
    If cmdOk.Tag = "新增入库单" Then
        strSQL = "Select Distinct a.Id, a.名称, a.规格, a.编码, a.是否变价, c.现价" & vbNewLine & _
                        "From 收费项目目录 A, 药品规格 B, 收费价目 C" & vbNewLine & _
                        "Where a.Id = b.药品id And a.Id = c.收费细目id And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                        "      Sysdate Between c.执行日期 And Nvl(c.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))  And Exists" & vbNewLine & _
                        "     (Select 1 From 收费执行科室 D Where b.药品id = d.收费细目id And d.执行科室id = [1])"
   
        Set rs药品信息 = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询药品信息", Val(cbo库房.ItemData(cbo库房.ListIndex)))
    
        If rs药品信息.RecordCount = 0 Then
            MsgBox "没有查询到药品信息，请检查！", vbInformation, GSTR_MESSAGE
            vsfList.Rows = 1
            lbl产生单据数.Caption = "提示：共有0张单据"
            Exit Sub
        End If
    End If
    
    vsfList.Rows = 1
    With vsfList
        For i = 1 To rsVal.RecordCount
            If cmdOk.Tag = "新增入库单" Then
                If IIf(IsNull(rsVal!生产商), "", rsVal!生产商) <> "" Then
                    Call Add生产商(rsVal!生产商)
                End If
                
                str供应商信息 = Check供应商(rsVal!供应商编码)
                 str供应商 = Split(str供应商信息, "|")(0)
                 lng供应商id = Val(Split(str供应商信息, "|")(1))
                
                rs药品信息.Filter = "编码='" & rsVal!药品编码 & "'"
                If rs药品信息.RecordCount = 0 Then
                    lng药品ID = 0
                Else
                    lng药品ID = rs药品信息!Id
                    str药品名称 = "[" & rs药品信息!编码 & "]" & rs药品信息!名称
                    str药品规格 = rs药品信息!规格
                    dbl零售价 = rs药品信息!现价
                    int变价 = IIf(IsNull(rs药品信息!是否变价), 0, rs药品信息!是否变价)
                End If
            ElseIf cmdOk.Tag = "补录发票信息" Then
                str供应商 = rsVal!供应商
                lng供应商id = Val(rsVal!供药单位ID)
                lng药品ID = Val(rsVal!药品ID)
                str药品名称 = "[" & rsVal!药品编码 & "]" & rsVal!名称
                str药品规格 = rsVal!规格
            End If
            
            rs发票信息.Filter = "供应商编码='" & rsVal!供应商编码 & "' and 药品编码='" & rsVal!药品编码 & "'"
    
            If lng药品ID <> 0 Then
                .Rows = .Rows + 1
                .TextMatrix(i, .ColIndex("行号")) = i
                .TextMatrix(i, .ColIndex("药品ID")) = lng药品ID
                .TextMatrix(i, .ColIndex("药品名称")) = str药品名称
                .TextMatrix(i, .ColIndex("规格")) = str药品规格
                .TextMatrix(i, .ColIndex("生产商")) = NVL(rsVal!生产商)
                .TextMatrix(i, .ColIndex("批号")) = NVL(rsVal!批号)
                .TextMatrix(i, .ColIndex("生产日期")) = Format(NVL(rsVal!生产日期), "yyyy-mm-dd")
                .TextMatrix(i, .ColIndex("有效日期")) = Format(NVL(rsVal!有效日期), "yyyy-mm-dd")
                .TextMatrix(i, .ColIndex("单位")) = NVL(rsVal!单位)
                .TextMatrix(i, .ColIndex("单价")) = Format(NVL(rsVal!单价, 0), "0.0000")
                .TextMatrix(i, .ColIndex("数量")) = Format(NVL(rsVal!数量, 0), "0.00")
                .TextMatrix(i, .ColIndex("金额")) = Format(NVL(rsVal!金额, 0), "0.00")
                .TextMatrix(i, .ColIndex("供应商")) = str供应商
                .TextMatrix(i, .ColIndex("供应商id")) = lng供应商id
                
                .Cell(flexcpForeColor, i, .ColIndex("行号"), i, .ColIndex("发票金额")) = IIf(lng供应商id = 0, mCanColColor, mNoColColor)
                
                If cmdOk.Tag = "新增入库单" Then
                    .TextMatrix(i, .ColIndex("零售价")) = Format(IIf(int变价 = 0, dbl零售价, NVL(rsVal!单价, 0)), "0.0000")
                    .TextMatrix(i, .ColIndex("零售金额")) = Format(.TextMatrix(i, .ColIndex("零售价")) * .TextMatrix(i, .ColIndex("数量")), "0.00")
                    .TextMatrix(i, .ColIndex("加成率")) = Format(IIf(.TextMatrix(i, .ColIndex("单价")) = 0, 0, .TextMatrix(i, .ColIndex("零售价")) / .TextMatrix(i, .ColIndex("单价")) - 1), "0.0000")
                    .TextMatrix(i, .ColIndex("差价")) = Format(.TextMatrix(i, .ColIndex("零售金额")) - .TextMatrix(i, .ColIndex("金额")), "0.00")
                    .ColHidden(.ColIndex("NO")) = True
                    .ColHidden(.ColIndex("审核人")) = True
                    .ColHidden(.ColIndex("审核日期")) = True
                ElseIf cmdOk.Tag = "补录发票信息" Then
                    .TextMatrix(i, .ColIndex("零售价")) = Format(NVL(rsVal!零售价, 0), "0.0000")
                    .TextMatrix(i, .ColIndex("零售金额")) = Format(NVL(rsVal!零售金额, 0), "0.00")
                    .TextMatrix(i, .ColIndex("加成率")) = Format(NVL(rsVal!加成率, 0), "0.0000")
                    .TextMatrix(i, .ColIndex("差价")) = Format(NVL(rsVal!差价, 0), "0.00")
                    .TextMatrix(i, .ColIndex("NO")) = rsVal!NO
                    .TextMatrix(i, .ColIndex("外观")) = NVL(rsVal!外观)
                    .TextMatrix(i, .ColIndex("产品合格证")) = NVL(rsVal!产品合格证)
                    .TextMatrix(i, .ColIndex("核查人")) = NVL(rsVal!核查人)
                    .TextMatrix(i, .ColIndex("核查日期")) = NVL(rsVal!核查日期)
                    .TextMatrix(i, .ColIndex("批次")) = NVL(rsVal!批次, 0)
                    .TextMatrix(i, .ColIndex("批准文号")) = NVL(rsVal!批准文号)
                    .TextMatrix(i, .ColIndex("随货单号")) = NVL(rsVal!随货单号)
                    .TextMatrix(i, .ColIndex("金额差")) = NVL(rsVal!金额差, 0)
                    .TextMatrix(i, .ColIndex("发票代码")) = NVL(rsVal!发票代码)
                    .TextMatrix(i, .ColIndex("计划id")) = NVL(rsVal!计划id, 0)
                    .TextMatrix(i, .ColIndex("验收结论")) = NVL(rsVal!验收结论)
                    .TextMatrix(i, .ColIndex("对方部门ID")) = NVL(rsVal!对方部门ID, 0)
                    .TextMatrix(i, .ColIndex("审核人")) = NVL(rsVal!审核人)
                    .TextMatrix(i, .ColIndex("审核日期")) = NVL(rsVal!审核日期)
                    .TextMatrix(i, .ColIndex("序号")) = NVL(rsVal!序号)
                    .ColHidden(.ColIndex("NO")) = False
                    .ColHidden(.ColIndex("审核人")) = False
                    .ColHidden(.ColIndex("审核日期")) = False
                End If
                '发票信息
                If rs发票信息.RecordCount = 1 Then
                    .TextMatrix(i, .ColIndex("发票号")) = NVL(rs发票信息!发票号)
                    .TextMatrix(i, .ColIndex("发票日期")) = Format(NVL(rs发票信息!发票日期), "yyyy-mm-dd")
                    
                    If Val(NVL(rs发票信息!发票数量, 0)) = Val(.TextMatrix(i, .ColIndex("数量"))) Then
                        .TextMatrix(i, .ColIndex("发票金额")) = Format(NVL(rs发票信息!发票金额, 0), "0.00")
                    Else
                        If Val(NVL(rs发票信息!发票数量, 0)) = 0 Then
                            .TextMatrix(i, .ColIndex("发票金额")) = Format(0, "0.00")
                        Else
                            .TextMatrix(i, .ColIndex("发票金额")) = Format(NVL(rs发票信息!发票金额, 0) / rs发票信息!发票数量 * .TextMatrix(i, .ColIndex("数量")), "0.00")
                        End If
                    End If
                    
                ElseIf rs发票信息.RecordCount > 1 Then
                    .Cell(flexcpForeColor, i, .ColIndex("行号"), i, .ColIndex("发票金额")) = mFPColColor
                End If
                
                If cmdOk.Tag = "补录发票信息" Then
                    If NVL(rsVal!发票号) <> "" Then .TextMatrix(i, .ColIndex("发票号")) = NVL(rsVal!发票号)
                    If NVL(rsVal!发票日期) <> "" Then .TextMatrix(i, .ColIndex("发票日期")) = Format(NVL(rsVal!发票日期), "yyyy-mm-dd")
                    If NVL(rsVal!发票金额) <> "" Then .TextMatrix(i, .ColIndex("发票金额")) = Format(NVL(rsVal!发票金额, 0), "0.00")
                End If

                If cmdOk.Tag = "新增入库单" Then
                    If .Cell(flexcpForeColor, i, .ColIndex("行号"), i, .ColIndex("发票金额")) = mNoColColor Then
                        If InStr(";" & str供应商id & ";", ";" & .TextMatrix(i, .ColIndex("供应商id")) & ";") = 0 Then
                            str供应商id = IIf(str供应商id = "", "", str供应商id & ";") & .TextMatrix(i, .ColIndex("供应商id"))
                            int产生单据数量 = int产生单据数量 + 1
                        End If
                    End If
                Else
                    If .Cell(flexcpForeColor, i, .ColIndex("行号"), i, .ColIndex("发票金额")) = mNoColColor Then
                        If InStr(";" & strNO串 & ";", ";" & .TextMatrix(i, .ColIndex("NO")) & ";") = 0 Then
                            strNO串 = IIf(strNO串 = "", "", strNO串 & ";") & .TextMatrix(i, .ColIndex("NO"))
                            int产生单据数量 = int产生单据数量 + 1
                        End If
                    End If
                End If
                
            End If
            
            rsVal.MoveNext
        Next
        
    End With

    If int产生单据数量 > 0 Then
        lbl产生单据数.Caption = "提示：共有" & int产生单据数量 & "张单据"
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "获取外部数据错误！", vbInformation, GSTR_MESSAGE
    vsfList.Rows = 1
End Sub

Private Sub Add生产商(ByVal str生产商 As String)
    Dim int编码 As Integer
    Dim strCodes As String
    Dim rs生产商 As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand

    strSQL = "select 编码,名称 from 药品生产商 where 名称=[1]"
    Set rs生产商 = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询生产商信息", str生产商)
    If rs生产商.RecordCount = 0 Then
                
        strSQL = "SELECT Nvl(MAX(LENGTH(编码)),2) As Length FROM 药品生产商"
        Set rs生产商 = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "药品生产商编码长度")
        int编码 = rs生产商!length
        
        strSQL = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM 药品生产商"
        Set rs生产商 = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "药品生产商编码")
        strCodes = rs生产商!Code
        
        int编码 = Len(strCodes)
        strCodes = strCodes + 1
        If int编码 >= Len(strCodes) Then
            strCodes = String(int编码 - Len(strCodes), "0") & strCodes
        End If
    
        strSQL = "ZL_药品生产商_INSERT('" & strCodes & "','" & str生产商 & "',zlSpellCode('" & str生产商 & "',10))"
        
        Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "")
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "获取外部数据错误！", vbInformation, GSTR_MESSAGE
End Sub

Private Function Check供应商(ByVal str供应商编码 As String) As String
    Dim rs供应商 As New ADODB.Recordset
    Dim str供应商 As String
    Dim strSQL As String
    Dim lng供应商id As Long
    
    On Error GoTo ErrHand

    strSQL = "Select a.Id, a.编码 ,a.名称" & vbNewLine & _
                    "From 供应商 A" & vbNewLine & _
                    "Where a.末级 = 1 And substr(类型,1,1)=1 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) and a.编码=[1]"
                    
    Set rs供应商 = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询供应商信息", str供应商编码)

    If rs供应商.RecordCount > 0 Then
        Check供应商 = rs供应商!名称 & "|" & rs供应商!Id
    Else
        Check供应商 = "|"
    End If
    
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "获取外部数据错误！", vbInformation, GSTR_MESSAGE
End Function

Private Sub ConnectDatabase()
'用于连接数据库
    Dim str服务器 As String, str数据库 As String, str用户名 As String, str密码 As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = "Select 类型, 设置 From 药品三方事务接口 where 名称='东莞药事管理系统' and 是否启动=1"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    '类型：0--SQLserver数据库；1--Oracle数据库
    'SQLserver数据库   设置：服务器名|数据库名|用户名|密码
    'Oracle数据库        设置：服务器名|数据库名|用户名|密码

    If rsTmp.RecordCount > 0 Then
        str服务器 = Split(rsTmp!设置, "|")(0)
        str数据库 = Split(rsTmp!设置, "|")(1)
        str用户名 = Split(rsTmp!设置, "|")(2)
        str密码 = Split(rsTmp!设置, "|")(3)
    Else
        MsgBox "连接服务器失败，请设置中间数据库的连接！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

'        str服务器 = "秦龙\WINCCPLUSMIG2008"
'        str数据库 = "master"
'        str用户名 = "sa"
'        str密码 = "his"
    
    mint数据库类型 = NVL(rsTmp!类型, 0)
    If mint数据库类型 = 0 Then
        'SQLserver数据库
        mblnIsConn = MSSQLServerOpen(str服务器, str数据库, str用户名, str密码)
    Else
        'Oracle数据库
        mblnIsConn = OraDataOpenTest(str数据库, str用户名, str密码)
    End If

    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "获取外部数据错误！", vbInformation, GSTR_MESSAGE
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Height < 8000 Then Me.Height = 8000
    If Me.Width < 14000 Then Me.Width = 14000
    
    Frmline1.Left = 0
    Frmline1.Width = Me.ScaleWidth
    Frmline2.Left = 0
    Frmline2.Width = Me.ScaleWidth
    Frmline2.Top = Me.ScaleHeight * 25 / 28
    
    vsfList.Left = Me.ScaleHeight / 80
    vsfList.Width = Me.ScaleWidth - Me.ScaleHeight / 40
    vsfList.Height = Frmline2.Top - vsfList.Top - Me.ScaleHeight / 40
    
    cmdCancel.Left = vsfList.Width - cmdCancel.Width + Me.ScaleHeight / 80
    cmdCancel.Top = Frmline2.Top + Me.ScaleHeight / 25
    
    cmdOk.Top = cmdCancel.Top
    cmdOk.Left = cmdCancel.Left - cmdOk.Width - 100

    lbl产生单据数.Top = cmdCancel.Top + 80
    lbl产生单据数.Left = vsfList.Left
    
    lbl供应商颜色.Top = cmdCancel.Top + 100
    lbl供应商颜色.Left = Me.ScaleWidth / 2
    
    pic供应商颜色.Top = cmdCancel.Top + 60
    pic供应商颜色.Left = lbl供应商颜色.Left - pic供应商颜色.Width - 50
    
    lbl发票信息.Top = cmdCancel.Top + 100
    lbl发票信息.Left = pic供应商颜色.Left - lbl发票信息.Width - 500
    
    pic发票信息.Top = cmdCancel.Top + 60
    pic发票信息.Left = lbl发票信息.Left - pic发票信息.Width - 50
End Sub

Private Sub txt配送单号_KeyPress(KeyAscii As Integer)
    If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt配送单号_GotFocus()
    Me.txt配送单号.SelStart = 0: Me.txt配送单号.SelLength = 100
End Sub

Private Sub Save新增入库单()
    Dim i As Integer
    Dim strSQL As String
    Dim strNO As String
    Dim str供应商id As String
    Dim strNO串 As String
    Dim strDate As String
    Dim int序号 As Integer
    Dim lng库房ID As Long
    
    strDate = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    lng库房ID = Val(cbo库房.ItemData(cbo库房.ListIndex))
    
    With vsfList
        For i = 1 To .Rows - 1
            If .Cell(flexcpForeColor, i, .ColIndex("行号"), i, .ColIndex("发票金额")) = mNoColColor Then
                If InStr(";" & str供应商id & ";", ";" & .TextMatrix(i, .ColIndex("供应商id")) & ";") = 0 Then
                    str供应商id = IIf(str供应商id = "", "", str供应商id & ";") & .TextMatrix(i, .ColIndex("供应商id"))
                    strNO = gobjComLib.zlDatabase.GetNextNo(21, lng库房ID)
                    int序号 = 0
                    
                    strSQL = "Zl_配送单号对照_INSERT("
                    '库房id_In
                    strSQL = strSQL & lng库房ID
                    '单据_In
                    strSQL = strSQL & ",1"
                    'No_In
                    strSQL = strSQL & ",'" & strNO & "'"
                    '配送单号_In
                    strSQL = strSQL & ",'" & txt配送单号.Text & "'"
                    strSQL = strSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = strSQL
                    
                End If
                
                int序号 = int序号 + 1
                
                strSQL = "zl_药品外购_INSERT("
                'NO
                strSQL = strSQL & "'" & strNO & "'"
                '序号
                strSQL = strSQL & "," & int序号
                '库房ID
                strSQL = strSQL & "," & lng库房ID
                '对方部门ID
                strSQL = strSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("对方部门ID"))) = 0, "NULL", Val(.TextMatrix(i, .ColIndex("对方部门ID"))))
                '供药单位ID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("供应商id")))
                '药品ID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("药品id")))
                '产地
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("生产商")) & "'"
                '批号
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("批号")) & "'"
                '效期
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("有效日期")) & "','yyyy-mm-dd HH24:MI:SS')"
                '实际数量
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("数量")))
                '成本价
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("单价")))
                '成本金额
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("金额")))
                '扣率
                strSQL = strSQL & "," & 100
                '零售价
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("零售价")))
                '零售金额
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("零售金额")))
                '差价
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("差价")))
                '摘要
                strSQL = strSQL & ",'由配送单号[" & Trim(txt配送单号.Text) & "]导入'"
                '填制人
                strSQL = strSQL & ",'" & gstrUserNameNew & "'"
                '发票号
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("发票号")) & "'"
                '发票日期
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("发票日期")) & "','yyyy-mm-dd HH24:MI:SS')"
                '发票金额
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("发票金额")))
                '填制日期
                strSQL = strSQL & "," & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS')"
                '外观
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("外观")) & "'"
                '产品合格证
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("产品合格证")) & "'"
                '核查人
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("核查人")) & "'"
                '核查日期
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("核查日期")) & "','yyyy-mm-dd HH24:MI:SS')"
                '批次
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("批次")))
                '是否退货
                strSQL = strSQL & "," & 1
                '生产日期
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("生产日期")) & "','yyyy-mm-dd HH24:MI:SS')"
                '批准文号
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("批准文号")) & "'"
                '随货单号
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("随货单号")) & "'"
                '金额差
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("金额差")))
                '加成率
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("加成率")))

                strSQL = strSQL & ")"
                    
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = strSQL
            End If
        Next
    
    End With
    
End Sub

Private Sub Save未审核发票信息()
    Dim i As Integer
    Dim strSQL As String
    Dim strNO As String
    Dim str供应商id As String
    Dim strNO串 As String
    Dim strDate As String
    Dim int序号 As Integer
    Dim lng库房ID As Long
    
    strDate = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    lng库房ID = Val(cbo库房.ItemData(cbo库房.ListIndex))
    
    With vsfList
        For i = 1 To .Rows - 1
            If .Cell(flexcpForeColor, i, .ColIndex("行号"), i, .ColIndex("发票金额")) = mNoColColor And Trim(.TextMatrix(i, .ColIndex("审核人"))) = "" Then

                If InStr(";" & strNO串 & ";", ";" & .TextMatrix(i, .ColIndex("NO")) & ";") = 0 Then
                    strNO串 = IIf(strNO串 = "", "", strNO串 & ";") & .TextMatrix(i, .ColIndex("NO"))
                    int序号 = 0
                    strNO = .TextMatrix(i, .ColIndex("NO"))
                    strSQL = "zl_药品外购_Delete('" & strNO & "')"
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = strSQL
                End If
                
                int序号 = int序号 + 1
                
                strSQL = "zl_药品外购_INSERT("
                'NO
                strSQL = strSQL & "'" & strNO & "'"
                '序号
                strSQL = strSQL & "," & int序号
                '库房ID
                strSQL = strSQL & "," & lng库房ID
                '对方部门ID
                strSQL = strSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("对方部门ID"))) = 0, "NULL", Val(.TextMatrix(i, .ColIndex("对方部门ID"))))
                '供药单位ID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("供应商id")))
                '药品ID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("药品id")))
                '产地
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("生产商")) & "'"
                '批号
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("批号")) & "'"
                '效期
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("有效日期")) & "','yyyy-mm-dd HH24:MI:SS')"
                '实际数量
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("数量")))
                '成本价
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("单价")))
                '成本金额
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("金额")))
                '扣率
                strSQL = strSQL & "," & 100
                '零售价
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("零售价")))
                '零售金额
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("零售金额")))
                '差价
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("差价")))
                '摘要
                strSQL = strSQL & ",'由配送单号[" & Trim(txt配送单号.Text) & "]导入发票信息'"
                '填制人
                strSQL = strSQL & ",'" & gstrUserNameNew & "'"
                '发票号
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("发票号")) & "'"
                '发票日期
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("发票日期")) & "','yyyy-mm-dd HH24:MI:SS')"
                '发票金额
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("发票金额")))
                '填制日期
                strSQL = strSQL & "," & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS')"
                '外观
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("外观")) & "'"
                '产品合格证
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("产品合格证")) & "'"
                '核查人
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("核查人")) & "'"
                '核查日期
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("核查日期")) & "','yyyy-mm-dd HH24:MI:SS')"
                '批次
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("批次")))
                '是否退货
                strSQL = strSQL & "," & 1
                '生产日期
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("生产日期")) & "','yyyy-mm-dd HH24:MI:SS')"
                '批准文号
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("批准文号")) & "'"
                '随货单号
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("随货单号")) & "'"
                '金额差
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("金额差")))
                '加成率
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("加成率")))
                '发票代码
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("发票代码")) & "'"
                '计划id
                strSQL = strSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("计划id"))) = 0, "NULL", Val(.TextMatrix(i, .ColIndex("计划id"))))
                '财务审核
                strSQL = strSQL & "," & 0
                '验收结论
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("验收结论")) & "'"
                strSQL = strSQL & ")"
                    
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = strSQL
            End If
        Next
    
    End With
End Sub

Private Sub Save已审核发票信息()
    Dim i As Integer
    Dim strSQL As String
    
    With vsfList
        For i = 1 To .Rows - 1
            If .Cell(flexcpForeColor, i, .ColIndex("行号"), i, .ColIndex("发票金额")) = mNoColColor And Trim(.TextMatrix(i, .ColIndex("审核人"))) <> "" Then

                strSQL = "zl_药品外购发票信息_UPDATE("
                'NO
                strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("NO")) & "'"
                '序号
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("序号")))
                '发票号
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("发票号")) & "'"
                '发票日期
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("发票日期")) & "','yyyy-mm-dd HH24:MI:SS')"
                '发票金额
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("发票金额")))
                '供药单位ID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("供应商id")))
                '操作标志
                strSQL = strSQL & "," & 1
                strSQL = strSQL & ")"
                    
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = strSQL
                
            End If
        Next
    
    End With
End Sub


