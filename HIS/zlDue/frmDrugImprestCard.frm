VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrugImprestCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "预付款单"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmDrugImprestCard.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.TreeView tvwProvider 
      Height          =   3585
      Left            =   1560
      TabIndex        =   23
      Top             =   1665
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   6324
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgTree"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5400
      TabIndex        =   28
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   6120
      Top             =   3720
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
            Picture         =   "frmDrugImprestCard.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugImprestCard.frx":1D18
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugImprestCard.frx":3A22
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
      Height          =   3375
      Left            =   0
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5953
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fraImprest 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton Cmd供应商 
         Caption         =   "…"
         Height          =   300
         Left            =   4680
         TabIndex        =   27
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox Txt供药单位 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   1230
         Width           =   3225
      End
      Begin VB.TextBox TxtNo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3405
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   10
         Top             =   810
         Width           =   1515
      End
      Begin VB.TextBox Txt付款说明 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   3
         Top             =   3960
         Width           =   3585
      End
      Begin VB.TextBox Txt填制日期 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   4380
         Width           =   1875
      End
      Begin VB.TextBox Txt审核日期 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   4770
         Width           =   1875
      End
      Begin VB.TextBox Txt审核人 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4770
         Width           =   1005
      End
      Begin VB.TextBox Txt填制人 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4380
         Width           =   1005
      End
      Begin ZL9BillEdit.BillEdit mshImprest 
         Height          =   1485
         Left            =   495
         TabIndex        =   2
         Top             =   2385
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2619
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label txt税务号 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1440
         TabIndex        =   26
         Top             =   2070
         Width           =   3450
      End
      Begin VB.Label txt开户行 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1440
         TabIndex        =   25
         Top             =   1830
         Width           =   3450
      End
      Begin VB.Label txt电话地址 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1440
         TabIndex        =   24
         Top             =   1560
         Width           =   3450
      End
      Begin VB.Label Lbl标题 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "预付款通知单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1410
         TabIndex        =   21
         Top             =   360
         Width           =   1890
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3075
         TabIndex        =   20
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Lbl单位名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   19
         Top             =   1290
         Width           =   810
      End
      Begin VB.Label Lbl电话地址 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "地址电话:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   18
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Lbl开户行 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开户银行:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   17
         Top             =   1830
         Width           =   810
      End
      Begin VB.Label Lbl付款说明 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "付款说明:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   495
         TabIndex        =   16
         Top             =   4020
         Width           =   810
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   2265
         TabIndex        =   15
         Top             =   4845
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   525
         TabIndex        =   14
         Top             =   4830
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   2265
         TabIndex        =   13
         Top             =   4440
         Width           =   720
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   525
         TabIndex        =   12
         Top             =   4440
         Width           =   540
      End
      Begin VB.Label lbl税务号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "税务登记号:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   2070
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmDrugImprestCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSuccess As Boolean
Private mstr单据号 As String
Private mblnSave As Boolean
Private mint编辑状态 As Integer
Private mint记录状态 As Integer
Private mblnChange As Boolean
Private mfrmMain As Object
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Dim mstrPrivs As String                     '权限
Private Const mlngModule = 1323

Private Function GetDepend() As Boolean
'--------------------------------------------------------------
'功能：读取供应商、结算方式等数据
'参数：
'返回：取出数据是否成功
'说明：
'--------------------------------------------------------------
    Dim rsDepend As New Recordset
    Dim rs结算方式 As New Recordset
    Dim intLop As Integer
    Dim str权限 As String
    GetDepend = False
    
    str权限 = " and (末级<>1 or (末级=1 " & zl_获取站点限制 & " and " & Get分类权限(gstrPrivs) & "))"
    
    gstrSQL = "" & _
        "   Select ID,上级ID,编码,简码,名称,末级,地址||电话 as 电话地址,开户银行,帐号,税务登记号 " & _
        "   From 供应商 " & _
        "   Where (撤档时间 is null or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01') " & _
        "        " & str权限 & _
        "   Start with 上级ID is Null Connect by prior ID=上级ID"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rsDepend, gstrSQL, "预付款通知单"
    With rsDepend
        '读取供应商信息
        If .EOF Then
            '无供应商信息退出
            ShowMsgbox "供应商的信息不全！"
            Exit Function
        End If
    End With
    'by lesfeng 2009-12-2 性能优化
    gstrSQL = "Select 应用场合,结算方式,缺省标志 From 结算方式应用 Where 应用场合='付货款' Order by 缺省标志 desc"
    zlDatabase.OpenRecordset rs结算方式, gstrSQL, "预付款通知单"
    With rs结算方式
        '读取结算方式
        If .EOF Then
            '无结算方式数据退出
            ShowMsgbox "结算方式应用信息不全！"
            Exit Function
        End If
        '填充结算方式列表
        mshImprest.Clear
        For intLop = 1 To .RecordCount
            mshImprest.AddItem !结算方式
            .MoveNext
        Next
        mshImprest.ListIndex = 0
        
        .Close
    End With
    
    With rsDepend
        '填充供应商数据到TreeView中
        tvwProvider.Nodes.Clear
        tvwProvider.Nodes.Add , , "R", "所有供应商", 1, 1
        tvwProvider.Nodes("R").Tag = 0
        .MoveFirst
        
        Do While Not .EOF
            If IsNull(!上级ID) Then
                If !末级 = 1 Then
                    tvwProvider.Nodes.Add "R", 4, "K_" & !ID, !名称, 3, 3
                Else
                    tvwProvider.Nodes.Add "R", 4, "K_" & !ID, !名称, 2, 2
                End If
            Else
                If !末级 = 1 Then
                    tvwProvider.Nodes.Add "K_" & !上级ID, 4, "K_" & !ID, !名称, 3, 3
                Else
                    tvwProvider.Nodes.Add "K_" & !上级ID, 4, "K_" & !ID, !名称, 2, 2
                End If
            End If
            tvwProvider.Nodes("K_" & !ID).Tag = !末级
            .MoveNext
        Loop
        tvwProvider.Nodes("R").Selected = True
        tvwProvider.Nodes("R").Expanded = True
        
    End With
    GetDepend = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ShowCard(frmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
        Optional int记录状态 As Integer = 1, Optional blnSuccess As Boolean = False)
'--------------------------------------------------------------
'功能：预付款录入及编辑
'参数：FrmMain---------调用窗体
'      str单据号-------预付款单据号,新增预付款时为""
'      int编辑状态-----1、新增预付款，2、编辑，3、审核，4、查看打印
'      int记录状态-----1、正常记录，2、被冲销记录，3、冲销记录
'      BlnSuccess------返回是否成功操作
'返回：
'说明：
'--------------------------------------------------------------
    mblnSave = False
    mblnSuccess = False
    Set mfrmMain = frmMain
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1323)
    
    If Not GetDepend Then Exit Sub
    
    If mint编辑状态 = 1 Then            '新增预付款
        mstr单据号 = NextNo(31)
        TxtNo = mstr单据号
        
    ElseIf mint编辑状态 = 2 Then        '编辑预付款
'        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then        '审核预付款
        'mblnEdit = False
        cmdOk.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then        '查看及打印
        'mblnEdit = False
        cmdOk.Caption = "打印(&P)"
        If InStr(mstrPrivs, "预付款通知单打印") = 0 Then
            cmdOk.Visible = False
        Else
            cmdOk.Visible = True
        End If
    End If
    
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

Private Sub cmdCancel_Click()
    '退出
    Unload Me
End Sub

Private Function ValidData() As Boolean
'--------------------------------------------------------------
'功能：检查输入的数据是否符合要求
'参数：
'返回：True----符合，False----不符合
'说明：
'--------------------------------------------------------------
    Dim intRow As Integer
    
    ValidData = False
    '检查供应商
    If Txt供药单位.Text = "" Then
         ShowMsgbox "对不起，没有供药单位!"
         Txt供药单位.SetFocus
         Exit Function
    End If
    If IIf(Txt供药单位.Tag = "", 0, Txt供药单位.Tag) = 0 Then
        ShowMsgbox "对不起，没有正确选择供药单位，请重新选择!"
         Txt供药单位.SetFocus
         Exit Function
    End If
    
    With mshImprest
        '检查每行数据的输入金额
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                If IIf(.TextMatrix(intRow, 1) = "", 0, .TextMatrix(intRow, 1)) = 0 And intRow <> .Rows - 1 Then
                    ShowMsgbox "对不起，金额必须输入，且不为零"
                    .SetFocus
                    .Row = intRow
                    .MsfObj.TopRow = intRow
                    .Col = 1
                    Exit Function
                End If
                If Val(.TextMatrix(intRow, 1)) > 9999999999999# Then
                    ShowMsgbox "第" & intRow & "行金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！"
                    .SetFocus
                    .Row = intRow
                    .MsfObj.TopRow = intRow
                    .Col = 1
                    Exit Function
                End If
                If Val(.TextMatrix(intRow, 1)) < -1 * 9999999999999# Then
                    ShowMsgbox "第" & intRow & "行金额大于了数据库能够保存的" & vbCrLf & "最大范围-1*9999999999999，请检查！"
                    .SetFocus
                    .Row = intRow
                    .MsfObj.TopRow = intRow
                    .Col = 1
                    Exit Function
                End If
            End If
        Next
    End With
    '检查付款说明
    If LenB(StrConv(Txt付款说明.Text, vbFromUnicode)) > 50 Then
        ShowMsgbox "付款说明的长度超长!(最多为50个字符或25个汉字)"
        Txt付款说明.SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Function SaveCard() As Boolean
'--------------------------------------------------------------
'功能：保存录入的应付款数据
'参数：
'返回：保存是否成功
'说明：
'--------------------------------------------------------------
    Dim intRow As Integer
    Dim strNO_IN As String
    Dim int序号_IN As Integer
    Dim int预付款_IN As Integer
    Dim lng单位ID_IN As Long
    Dim dbl金额_IN As Double
    Dim str结算方式_IN As String
    Dim str结算号码_IN As String
    Dim str填制人_IN As String
    Dim str填制日期_IN As String
    Dim lng付款序号_IN As Long
    Dim str摘要_IN As String
    
    SaveCard = False
    '准备数据
    strNO_IN = TxtNo
    int预付款_IN = 1
    lng单位ID_IN = Txt供药单位.Tag
    str填制人_IN = UserInfo.姓名
    str填制日期_IN = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    str摘要_IN = Txt付款说明
    
    
    On Error GoTo errHandle:
    
    '开始事务
    gcnOracle.BeginTrans
    
    If mint编辑状态 = 2 Then
            gstrSQL = "zl_付款记录_DELETE('" & TxtNo & "')"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End If
        
    '循环保存每行数据
    With mshImprest
        'zl_付款管理_INSERT( /*strNO_IN*/, /*int序号_IN*/, /*int预付款_IN*/, /*lng单位ID_IN*/,
            '/*dbl金额_IN*/, /*str结算方式_IN*/, /*str结算号码_IN*/, /*str填制人_IN*/, /*str填制日期_IN*/,
            '/*lng付款序号_IN*/, /*str摘要_IN*/ );
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And IIf(.TextMatrix(intRow, 1) = "", 0, .TextMatrix(intRow, 1)) <> 0 Then
                int序号_IN = intRow
                dbl金额_IN = .TextMatrix(intRow, 1)
                str结算方式_IN = .TextMatrix(intRow, 0)
                str结算号码_IN = .TextMatrix(intRow, 2)
                
                gstrSQL = "zl_付款管理_INSERT('" & strNO_IN & "'," & int序号_IN & "," & int预付款_IN & "," & lng单位ID_IN _
                    & "," & dbl金额_IN & ",'" & str结算方式_IN & "','" & str结算号码_IN & "','" & str填制人_IN & "',to_date('" _
                    & str填制日期_IN & "','yyyy-mm-dd HH24:MI:SS'),NULL,'" & str摘要_IN & "')"
               zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With
    
    '提交事务
    gcnOracle.CommitTrans
    SaveCard = True
    Exit Function

errHandle:
    gcnOracle.RollbackTrans

    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub cmdHelp_Click()
    '帮助
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    '确定
    Dim blnSuccess As Boolean
    Dim strReg As String
    Select Case mint编辑状态
        Case 1, 2           '新增或编辑
            With mshImprest
                If .TextMatrix(1, 0) = "" Then Exit Sub
                If Not ValidData() Then Exit Sub        '如果录入数据不符合规范则退出模块
                blnSuccess = SaveCard                   '保存数据
                If blnSuccess = False Then Exit Sub     '保存失败则退出模块
                mblnChange = False
                mblnSave = False
                mblnSuccess = True
                If IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                     '打印
                    If InStr(mstrPrivs, "预付款通知单打印") <> 0 Then
                        ReportOpen gcnOracle, glngSys, "zl1_bill_1323_2", Me, "单据编号=" & TxtNo.Text, "记录状态=" & mint记录状态, 2
                    End If
                End If
                
                If mint编辑状态 = 1 Then                '新增时重新开始下一条数据的增加
                    .ClearBill
                    TxtNo = NextNo(31)
                    Txt供药单位.Text = ""
                    Txt供药单位.Tag = 0
                    txt电话地址 = ""
                    Txt付款说明 = ""
                    txt开户行 = ""
                    txt税务号 = ""
                    Txt供药单位.SetFocus
                Else                                    '编辑时退出
                    Unload Me
                End If
                mblnChange = False
                Exit Sub
            End With
        Case 3                                          '审核
            With mshImprest
                If .TextMatrix(1, 0) = "" Then Exit Sub
                If Not ValidData() Then Exit Sub        '检查数据
                blnSuccess = SaveVerify                 '审核
                If blnSuccess = False Then Exit Sub
                
                If IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                     '打印
                    If InStr(mstrPrivs, "预付款通知单打印") <> 0 Then
                        ReportOpen gcnOracle, glngSys, "zl1_bill_1323_2", Me, "单据编号=" & TxtNo.Text, "记录状态=" & mint记录状态, 2
                    End If
                End If
                
                mblnChange = False
                mblnSave = False
                mblnSuccess = True
                Unload Me
                Exit Sub
            End With
        Case 4
            '打印
            ReportOpen gcnOracle, glngSys, "zl1_bill_1323_2", Me, "单据编号=" & TxtNo.Text, "记录状态=" & mint记录状态, 2
    End Select
End Sub

Private Function SaveVerify() As Boolean
'--------------------------------------------------------------
'功能：审核预付款数据
'参数：
'返回：审核是否成功
'说明：
'--------------------------------------------------------------
    Dim intRow As Integer
    Dim strNO_IN As String
    Dim dbl付款金额_IN As Double
    Dim lng单位ID_IN As Long
    Dim str审核人_IN As String
    
    SaveVerify = False
    '准备数据
    strNO_IN = TxtNo
    lng单位ID_IN = Txt供药单位.Tag
    str审核人_IN = UserInfo.姓名
    dbl付款金额_IN = 0
    On Error GoTo errHandle:
    
    With mshImprest
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And IIf(.TextMatrix(intRow, 1) = "", 0, .TextMatrix(intRow, 1)) <> 0 Then
                dbl付款金额_IN = dbl付款金额_IN + Val(.TextMatrix(intRow, 1))
            End If
        Next
    End With
    '填写审核数据
    
    'zl_付款管理_VERIFY( /*strNO_IN*/, /*lng单位ID_IN*/, /*dbl付款金额_IN*/, /*str审核人_IN*/ );
    gstrSQL = "zl_付款管理_VERIFY('" & strNO_IN & "'" & "" & "" & "" _
        & "" & "" & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    
    
    SaveVerify = True
    Exit Function
errHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            '单据已被删除
            ShowMsgbox "该单据已被删除，请检查！"
            Unload Me
            Exit Sub
        Case 3
            '修改的单据已被审核
            ShowMsgbox "该单据已被其他人审核，请检查！"
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim intLop As Integer
    
    TxtNo = mstr单据号
    '初始化单据
    With mshImprest
        .Clear
        .Cols = 3
        .Rows = 2
        
        .TextMatrix(0, 0) = "付款方式"
        .TextMatrix(0, 1) = "付款金额"
        .TextMatrix(0, 2) = "结算号码"
        
        If Not RestoreFlexState(mshImprest, Me.Caption) Then
            .ColWidth(0) = 1000
            .ColWidth(1) = 1500
            .ColWidth(2) = 1800
        End If
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignLeftCenter
        
        .ColData(0) = 3
        .ColData(1) = 4
        .ColData(2) = 4
        
        .PrimaryCol = 0
        
        If mint编辑状态 > 2 Then
            .Active = False
            Txt供药单位.Enabled = False
            Cmd供应商.Enabled = False
            Txt付款说明.Enabled = False
            If mint编辑状态 = 3 Then
                cmdOk.Caption = "审核(&V)"
            Else
                cmdOk.Caption = "打印(&P)"
            End If
        Else
            .Active = True
        End If
    End With
    
    On Error GoTo errHandle
    If mint编辑状态 = 1 Then                '新增时填写
        Txt填制人 = UserInfo.姓名
        Txt填制日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    Else
        '从数据库中获取数据
        Dim rsImprest As New Recordset
        Dim intRecord As Integer
        
        gstrSQL = "" & _
            "   Select  a.序号,a.金额,a.结算方式,a.结算号码,a.摘要,a.填制人,a.填制日期,a.审核人," & _
            "           a.审核日期,b.名称,地址 || 电话 as 电话地址,开户银行,税务登记号,b.id " & _
            "   From 付款记录 a,供应商 b " & _
            "   Where a.单位id=b.id " & _
            "       and no=[1] and 记录状态=[2]"
        
        Set rsImprest = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr单据号, mint记录状态)
        
        If rsImprest.EOF Then
            '无数据
            mintParallelRecord = 2
            Exit Sub
        End If
        
        '填写数据到表单中
        intRecord = rsImprest.RecordCount
        Txt供药单位.Text = rsImprest!名称
        Txt供药单位.Tag = rsImprest!ID
        txt电话地址 = IIf(IsNull(rsImprest!电话地址), "", rsImprest!电话地址)
        txt开户行 = IIf(IsNull(rsImprest!开户银行), "", rsImprest!开户银行)
        txt税务号 = IIf(IsNull(rsImprest!税务登记号), "", rsImprest!税务登记号)
        Txt付款说明.Text = IIf(IsNull(rsImprest!摘要), "", rsImprest!摘要)
        Txt填制人 = rsImprest!填制人
        If mint编辑状态 = 2 Then
            Txt填制人 = UserInfo.姓名
        End If
        Txt填制日期 = Format(rsImprest!填制日期, "yyyy-mm-dd hh:mm:ss")
        Txt审核人 = IIf(IsNull(rsImprest!审核人), "", rsImprest!审核人)
        Txt审核日期 = IIf(IsNull(rsImprest!审核日期), "", Format(rsImprest!审核日期, "yyyy-mm-dd hh:mm:ss"))
        
        If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
            mintParallelRecord = 3
            Exit Sub
        End If
        
        With mshImprest
            For intLop = 1 To intRecord
                .TextMatrix(intLop, 0) = rsImprest!结算方式
                .TextMatrix(intLop, 1) = GetFormat(rsImprest!金额, 2)
                .TextMatrix(intLop, 2) = IIf(IsNull(rsImprest!结算号码), "", rsImprest!结算号码)
                If intLop = .Rows - 1 Then .Rows = .Rows + 1
                rsImprest.MoveNext
            Next
        End With
                
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        Txt供药单位.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If tvwProvider.Visible = True Then
        tvwProvider.Visible = False
        Txt供药单位.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        
        SaveFlexState mshImprest, Me.Caption
        Exit Sub
    End If
    Dim blnYes As Boolean
    
    ShowMsgbox "数据可能已改变，但未存盘，真要退出吗？", True, blnYes
    If blnYes = False Then
        Cancel = 1
        Exit Sub
    End If
    SaveFlexState mshImprest, Me.Caption
End Sub

Private Sub mshImprest_cboClick(ListIndex As Long)
        mshImprest.TextMatrix(mshImprest.Row, mshImprest.Col) = mshImprest.CboText
End Sub

Private Sub mshImprest_EditChange(curText As String)
    With mshImprest
        If .Col <> 0 Then
            .Text = UCase(curText)
            .SelStart = Len(curText)
        End If
    End With
    mblnChange = True
End Sub

Private Sub mshImprest_EnterCell(Row As Long, Col As Long)
    With mshImprest
    Select Case Col
        Case 1
            .TxtCheck = True
            .MaxLength = 16
            .TextMask = ".1234567890-"
        Case 2
            .TxtCheck = True
            .MaxLength = 10
            .ColData(Col) = 4
    End Select
    End With
End Sub

Private Sub mshImprest_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> 13 Then Exit Sub
    
    With mshImprest
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case 1
                If .Row = .Rows - 1 And KeyCode = vbKeyReturn And strKey = "" Then
                    Txt付款说明.SetFocus
                    Cancel = True
                    Exit Sub
                End If
                
                
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    ShowMsgbox "预付金额必须输入！"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    ShowMsgbox "预付金额必须为数字型,请重输！"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        ShowMsgbox "预付金额不能为零,请重输！"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) >= 10 ^ 12 - 1 Then
                        ShowMsgbox "预付金额必须小于" & (10 ^ 12 - 1)
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) < -10 ^ 12 - 1 Then
                        ShowMsgbox "预付金额必须大于" & (-1 * 10 ^ 12 - 1)
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = GetFormat(strKey, 2)
                End If
                    
            Case 2
               
                If KeyCode <> vbKeyReturn Then
                    .ColData(2) = 4
                    .TxtCheck = False
                Else
                    .ColData(2) = 0
                    .TxtCheck = True
                    .TextLen = 10
                End If
                
        End Select
    End With
End Sub

Private Sub mshProvider_DblClick()
    mshProvider_KeyPress 13
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intTemp As Integer
    Dim sngWidth As Single
    
    With mshProvider
        If KeyCode = vbKeyRight Then
            If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
                
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If .LeftCol <> 0 Then
                .LeftCol = .LeftCol - 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyHome Then
            If .LeftCol <> 0 Then
                .LeftCol = 0
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyEnd Then
            For intTemp = .Cols - 1 To 0 Step -1
                sngWidth = sngWidth + .ColWidth(intTemp)
                If sngWidth > .Width Then
                    .LeftCol = intTemp + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub mshProvider_KeyPress(KeyAscii As Integer)
    With mshProvider
        If KeyAscii = 13 Then
            Txt供药单位.Text = .TextMatrix(.Row, 2)
            Txt供药单位.Tag = .TextMatrix(.Row, 0)
            
            txt电话地址 = IIf(IsNull(.TextMatrix(.Row, 4)), "", .TextMatrix(.Row, 4))
            
            txt开户行 = IIf(IsNull(.TextMatrix(.Row, 5)), "", .TextMatrix(.Row, 5))
            txt税务号 = IIf(IsNull(.TextMatrix(.Row, 7)), "", .TextMatrix(.Row, 7))
            
            .Visible = False
            
            mshImprest.SetFocus
        End If
    End With
End Sub

Private Sub mshProvider_LostFocus()
    SaveFlexState mshProvider, Me.Caption
    If mshProvider.Visible Then mshProvider.Visible = False
End Sub
'设置供应商选择器的宽度及相关属性
Private Sub SetProviderWidth(ByVal sngLeft As Single, ByVal sngTop As Single)
    
    With mshProvider
        .Visible = True
        .Left = sngLeft
        .Top = sngTop
'        If RestoreFlexState(mshProvider, Me.Caption) = False Then
            'Select ID,名称,编码,简码,地址||电话 as 电话地址,开户银行,帐号,税务登记号
            
            .ColWidth(0) = 0
            .ColWidth(1) = 1000
            .ColWidth(2) = 2500
            .ColWidth(3) = 1000
            
            .ColWidth(4) = 1500
            .ColWidth(5) = 1500
            .ColWidth(6) = 1000
            .ColWidth(7) = 1000
            
'        End If
        
        .SetFocus
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub tvwProvider_DblClick()
    Dim rsProvider As New Recordset
    
    If tvwProvider.SelectedItem.Children <> 0 Then Exit Sub
    If tvwProvider.SelectedItem.Tag = 0 Then Exit Sub
    
    Txt供药单位 = tvwProvider.SelectedItem
    Txt供药单位.Tag = Mid(tvwProvider.SelectedItem.Key, 3)
    tvwProvider.Tag = "1"
    tvwProvider.Visible = False
    On Error GoTo errHandle
    gstrSQL = "" & _
        "   Select 名称,地址||电话 as 电话地址,开户银行,帐号,税务登记号 " & _
        "   From 供应商  " & _
        "   Where id=[1]"
    Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Txt供药单位.Tag))
    
    With rsProvider
        
        If .EOF Then Exit Sub
        
        Txt供药单位 = !名称
        txt电话地址 = IIf(IsNull(!电话地址), "", !电话地址)
        txt开户行 = IIf(IsNull(!开户银行), "", !开户银行)
        txt税务号 = IIf(IsNull(!税务登记号), "", !税务登记号)
        mshImprest.SetFocus
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Txt付款说明_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt付款说明, KeyAscii, m文本式
End Sub

Private Sub Txt供药单位_GotFocus()
    tvwProvider.Visible = False
End Sub

Private Sub txt供药单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(Txt供药单位)) = "" Then Exit Sub
    Dim strKey As String
    
    Dim rs供应商 As New ADODB.Recordset
    
    Dim str权限 As String
    str权限 = " and " & Get分类权限(gstrPrivs)
    strKey = GetMatchingSting(Txt供药单位.Text, False)
    gstrSQL = "" & _
        "   Select ID,编码,名称,简码,地址||电话 as 电话地址,开户银行,帐号,税务登记号  " & _
        "   From 供应商 " & _
        "   Where " & zl_获取站点限制(False) & "   " & _
        "       and (编码 like upper([1]) Or 名称 like [1] Or 简码 like upper([1])) " & _
        "       And (撤档时间 is null or 撤档时间>=to_date('3000-01-01','yyyy-mm-dd'))  And 末级=1 " & str权限
    On Error GoTo errHandle
    Set rs供应商 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKey)
    
    With rs供应商
    
        If .EOF Then
            ShowMsgbox "不存在该条件的供应商！"
            KeyCode = 0
            Txt供药单位 = ""
            tvwProvider.Tag = "0"
            Exit Sub
        End If
        If .RecordCount > 1 Then
            Set mshProvider.Recordset = rs供应商
            SetProviderWidth Txt供药单位.Left + fraImprest.Left, Txt供药单位.Top + Txt供药单位.Height + fraImprest.Top
            Exit Sub
        Else
            Txt供药单位 = !名称
            Txt供药单位.Tag = !ID
            tvwProvider.Tag = "1"
        End If
    End With
    
    Txt供药单位 = rs供应商!名称
    txt电话地址 = IIf(IsNull(rs供应商!电话地址), "", rs供应商!电话地址)
    txt开户行 = IIf(IsNull(rs供应商!开户银行), "", rs供应商!开户银行)
    txt税务号 = IIf(IsNull(rs供应商!税务登记号), "", rs供应商!税务登记号)
    zlCommFun.PressKey (vbKeyTab)
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Cmd供应商_Click()
    tvwProvider.Visible = tvwProvider.Visible Xor True
    If tvwProvider.Visible Then
        tvwProvider.Top = Txt供药单位.Top + Txt供药单位.Height + fraImprest.Top
        tvwProvider.SetFocus
    End If
End Sub

Private Sub txt付款说明_Change()
    mblnChange = True
End Sub

Private Sub txt付款说明_GotFocus()
    zlCommFun.OpenIme (True)
    With Txt付款说明
        .SelStart = 0
        .SelLength = Len(Txt付款说明.Text)
    End With
End Sub

Private Sub txt付款说明_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt付款说明_LostFocus()
    zlCommFun.OpenIme (False)
End Sub

Private Sub Txt供药单位_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt供药单位, KeyAscii, m文本式
End Sub
