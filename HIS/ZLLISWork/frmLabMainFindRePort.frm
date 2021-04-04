VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#3.4#0"; "zlIDKind.ocx"
Begin VB.Form frmLabMainFindRePort 
   Caption         =   "病人报告查询"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11745
   Icon            =   "frmLabMainFindRePort.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   11745
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptFind 
      Height          =   5355
      Left            =   30
      TabIndex        =   7
      Top             =   690
      Width           =   11655
      _Version        =   589884
      _ExtentX        =   20558
      _ExtentY        =   9446
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "预览(&V)"
      Height          =   345
      Left            =   5940
      TabIndex        =   17
      Top             =   6270
      Width           =   1155
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "已打印"
      Height          =   225
      Index           =   2
      Left            =   2250
      TabIndex        =   15
      Top             =   6090
      Width           =   885
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "已审核"
      Height          =   225
      Index           =   1
      Left            =   1155
      TabIndex        =   14
      Top             =   6090
      Width           =   885
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "已核收"
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   6090
      Width           =   885
   End
   Begin VB.CommandButton cmdUnionPrint 
      Caption         =   "合并打印(&U)"
      Height          =   345
      Left            =   7335
      TabIndex        =   12
      ToolTipText     =   "把多个标本合并为一个报告单进行打印"
      Top             =   6270
      Width           =   1155
   End
   Begin VB.CommandButton cmdSetupPrint 
      Caption         =   "打印设置(&P)"
      Height          =   345
      Left            =   4500
      TabIndex        =   11
      Top             =   6270
      Width           =   1155
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   10260
      TabIndex        =   10
      Top             =   6270
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   345
      Left            =   8805
      TabIndex        =   9
      Top             =   6270
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.CheckBox ch模糊查找 
         Caption         =   "模糊查找"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3090
         TabIndex        =   16
         Top             =   270
         Width           =   1155
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "时间范围"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4290
         TabIndex        =   8
         Top             =   270
         Width           =   1245
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Height          =   375
         Left            =   10170
         TabIndex        =   5
         Top             =   180
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   345
         Left            =   5550
         TabIndex        =   2
         Top             =   210
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   73138179
         CurrentDate     =   39449
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   750
         TabIndex        =   1
         Top             =   210
         Width           =   2265
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   345
         Left            =   7920
         TabIndex        =   4
         Top             =   210
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   73138179
         CurrentDate     =   39449
      End
      Begin zlIDKind.IDKind IDKind 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   217
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   582
         IDKindStr       =   "姓|姓名|0;医|医保号|1;身|身份证号|2;IC|IC卡号|3;门|门诊号|4;就|就诊卡|5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "至"
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
         Left            =   7650
         TabIndex        =   3
         Top             =   270
         Width           =   210
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMainFindRePort.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMainFindRePort.frx":0078
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMainFindRePort.frx":0612
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMainFindRePort.frx":0BAC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLabMainFindRePort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------- 2007-08-17 加入一卡通支持
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private mintUnion As Integer
Private mstrPrivs As String
Private mbln身份证 As Boolean

Private Enum IDKinds
    C0姓名 = 0
    C1医保号 = 1
    C2身份证号 = 2
    C3IC卡号 = 3
    C4门诊号 = 4
    C5就诊卡 = 5
End Enum
Private Enum mCol
    ID
    选择
    姓名
    性别
    年龄
    收费状态
    状态
    门诊号
    住院号
    标本号
    检验项目
    接收人
    接收时间
    样本条码
    费用性质
    开嘱科室
    开嘱医生
    开嘱时间
    发送人
    发送时间
    采样人
    采样时间
    检验人
    检验时间
    审核人
    审核时间
    医嘱id
    标本id
End Enum
Private mblnCard As Boolean '是否刷卡
Private mobjSquareCard As Object                                        '取卡类型
Private mblnShowPwd As Boolean                                          '是否显示密文

Private Sub chkDate_Click()
    Me.DTPBegin.Enabled = Me.chkDate.Value
    Me.DTPEnd.Enabled = Me.chkDate.Value
End Sub

Private Sub chkSelect_Click(Index As Integer)
    Dim intLoop As Integer
    
    With Me.rptFind
        If .Rows.Count = 0 Then Exit Sub
        For intLoop = 0 To .Rows.Count - 1
            '选中已核收
            If .Rows(intLoop).GroupRow = False Then
                If .Rows(intLoop).Record(mCol.状态).Value = "5-已核收" Then
                    .Rows(intLoop).Record(mCol.选择).Checked = (Me.chkSelect(0).Value = 1)
                End If
                '选中已审核
                If .Rows(intLoop).Record(mCol.状态).Value = "6-已审核" Then
                    .Rows(intLoop).Record(mCol.选择).Checked = (Me.chkSelect(1).Value = 1)
                End If
                '选中已打印
                If .Rows(intLoop).Record(mCol.状态).Value = "7-已打印" Then
                    .Rows(intLoop).Record(mCol.选择).Checked = (Me.chkSelect(2).Value = 1)
                End If
            End If
        Next
        .Redraw
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Me.cmdFind.SetFocus
    RefreshData
    Me.txtID.SetFocus
    Me.txtID.SelStart = 0
    Me.txtID.SelLength = Len(Me.txtID.Text)
End Sub

Private Sub cmdPreview_Click()
    Dim intLoop As Long
    With Me.rptFind
        If .FocusedRow Is Nothing Then Exit Sub
        If .FocusedRow.GroupRow = True Then Exit Sub
        intLoop = .FocusedRow.Index
        If .Rows(intLoop).Record(mCol.状态).Value = "5-已核收" Or .Rows(intLoop).Record(mCol.状态).Value = "6-已审核" _
            Or .Rows(intLoop).Record(mCol.状态).Value = "7-已打印" Then
            '预览
            ReportPrint intLoop, False
        End If
    End With
End Sub

Private Sub cmdPrint_Click()
    Dim intLoop As Integer
    Dim blnPrint As Boolean
    Dim strInfo As String
    
    With Me.rptFind
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).GroupRow = False Then
                If .Rows(intLoop).Record(mCol.选择).Checked = True Then
                    If .Rows(intLoop).Record(mCol.状态).Value = "5-已核收" Or .Rows(intLoop).Record(mCol.状态).Value = "6-已审核" _
                        Or .Rows(intLoop).Record(mCol.状态).Value = "7-已打印" Then
                        If Me.rptFind.Rows(intLoop).Record(mCol.审核人).Value = "" Then
                            If InStr(mstrPrivs, "未审核打印") <= 0 Then
                                If strInfo = "" Then
                                    MsgBox "你没有<未审核打印>权限，不能打印未审核单据!"
                                End If
                            Else
                                '打印
                                ReportPrint intLoop, True
                            End If
                        Else
                            '打印
                            ReportPrint intLoop, True
                        End If
                        
                    End If
                End If
            End If
        Next
    End With
    

End Sub

Private Sub cmdSetupPrint_Click()
    '打印设置
    PrintSetup
End Sub

Private Sub cmdUnionPrint_Click()
    '合并打印标本
    Call AllReportPrint
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyF10 Then
'        Call IdKindChange
'    End If
End Sub


Private Sub IdKindChange()
    If Me.ActiveControl Is txtGoto Then
       IDKind.IDKind = IIf(IDKind.IDKind = IDKinds.C5就诊卡, 0, IDKind.IDKind + 1)
    End If
End Sub

Private Sub Form_Load()
    Dim Column As ReportColumn
    
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    mbln身份证 = False
    
    rptFind.AllowColumnRemove = False
    rptFind.ShowItemsInGroups = False
    
    With rptFind.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "拖动列标题到这里,按该列分组..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
    End With
    rptFind.SetImageList ImgList
    With Me.rptFind.Columns
        Set Column = .Add(mCol.ID, "病人信息", 75, True): Column.Visible = False
        Set Column = .Add(mCol.选择, "", 18, False): Column.Icon = 0
        Set Column = .Add(mCol.姓名, "姓名", 60, True): Column.Visible = False
        Set Column = .Add(mCol.性别, "性别", 60, True): Column.Visible = False
        Set Column = .Add(mCol.年龄, "年龄", 60, True): Column.Visible = False
        Set Column = .Add(mCol.收费状态, "收费状态", 75, True): Column.Visible = False
        Set Column = .Add(mCol.状态, "状态", 100, True): Column.Visible = False
        Set Column = .Add(mCol.门诊号, "门诊号", 100, True): Column.Visible = False
        Set Column = .Add(mCol.住院号, "住院号", 100, True): Column.Visible = False
        Set Column = .Add(mCol.标本号, "标本号", 60, True)
        Set Column = .Add(mCol.检验项目, "检验项目", 100, True)
        Set Column = .Add(mCol.接收人, "接收人", 100, True)
        Set Column = .Add(mCol.接收时间, "接收时间", 100, True)
        Set Column = .Add(mCol.样本条码, "样本条码", 100, True)
        Set Column = .Add(mCol.费用性质, "费用性质", 100, True)
        Set Column = .Add(mCol.开嘱科室, "开嘱科室", 100, True)
        Set Column = .Add(mCol.开嘱时间, "开嘱时间", 100, True)
        Set Column = .Add(mCol.发送人, "发送人", 100, True)
        Set Column = .Add(mCol.发送时间, "发送时间", 100, True)
        Set Column = .Add(mCol.采样人, "采样人", 100, True)
        Set Column = .Add(mCol.采样时间, "采样时间", 100, True)
        Set Column = .Add(mCol.检验人, "检验人", 100, True)
        Set Column = .Add(mCol.检验时间, "检验时间", 100, True)
        Set Column = .Add(mCol.审核人, "审核人", 100, True)
        Set Column = .Add(mCol.审核时间, "审核时间", 100, True)
        Set Column = .Add(mCol.医嘱id, "医嘱ID", 100, True)
        Set Column = .Add(mCol.标本id, "标本ID", 100, True)
    End With
    
    Me.DTPEnd.Value = Now
    Me.DTPBegin.Value = Now - 30
    Me.chkDate.Value = zlDatabase.GetPara("frmLabMainFindRePort_使用时间范围", 100, 1208, 0)
    Me.rptFind.LoadSettings zlDatabase.GetPara("frmLabMainFindRePort_rptFind", 100, 1208, "")
    mintUnion = zlDatabase.GetPara("不区分仪器显示核收项目", 100, 1208, 0)
    Me.DTPBegin.Enabled = Me.chkDate.Value
    Me.DTPEnd.Enabled = Me.chkDate.Value
    IDKind.IDKind = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "输入方式", 0))
    
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle, False) = False Then
            MsgBox "IDKind初始化失败!", vbInformation, gstrSysName
        Else
            IDKind.IDKindStr = mobjSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    End If
    
    Call RestoreWinState(Me, App.ProductName)                   '界面恢复
End Sub

Private Sub RefreshData()
    '功能       刷新数据
    Dim strWhere As String
    Dim strFind As String
    Dim rsTmp As New ADODB.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim GroupRow As ReportRow
    Dim blnBarCode As Boolean
    Dim strSQLbak As String
    Dim lng卡类别ID As Long
    Dim lng病人ID As Long
    
    If Trim(Me.txtID.Text) = "" Then Me.txtID.SetFocus: Exit Sub
    
    If mbln身份证 Or IDKind.IDKind = IDKind.GetKindIndex("身份证号") Then
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), txtID, False, lng病人ID) = False Then lng病人ID = 0
        txtID = "-" & lng病人ID
'        strSQL = "select 病人ID from 病人信息 where 身份证号 = [1] "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtID)
'        If Not rsTmp.EOF Then
'            txtID = "-" & rsTmp.Fields("病人ID")
'        End If
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("IC卡号") Then
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), txtID, False, lng病人ID) = False Then lng病人ID = 0
        txtID = "-" & lng病人ID
'        strSQL = "select 病人ID from 病人信息 where IC卡号 = [1] "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtID)
'        If Not rsTmp.EOF Then
'            txtID = "-" & rsTmp.Fields("病人ID")
'        End If
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("就诊卡") Then
        strSQL = "select 病人ID from 病人信息 where 就诊卡号 = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtID)
        If Not rsTmp.EOF Then
            txtID.Tag = txtID.Text
            txtID = "-" & rsTmp.Fields("病人ID")
        End If
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("门诊号") Then
        If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), txtID, False, lng病人ID) = False Then lng病人ID = 0
        txtID = "-" & lng病人ID
'        If InStr("-+*./", Mid(Me.txtID.Text, 1, 1)) <= 0 Then
'            Me.txtID.Text = "*" & Me.txtID.Text
'        End If
    Else
        If Val(IDKind.GetKindItem("卡类别ID")) <> 0 Then
            lng卡类别ID = Val(IDKind.GetKindItem("卡类别ID"))
            If mobjSquareCard.zlGetPatiID(lng卡类别ID, txtID, False, lng病人ID) = False Then lng病人ID = 0
            If lng病人ID = 0 Then lng病人ID = 0
        Else
            lng病人ID = 0
'            If mobjSquareCard.zlGetPatiID(IDKind.GetKindItem("全名"), txtID, False, lng病人ID) = False Then lng病人ID = 0
        End If
        If lng病人ID > 0 Then
            txtID.Tag = txtID.Text
            txtID = "-" & lng病人ID
        End If
    End If
    
    Select Case Mid(Me.txtID, 1, 1)
        Case "-"                                '病人ID
            strWhere = "(select " & gConst_病人信息_列名 & " from 病人信息 a where 病人id = [1]) e "
            strFind = Val(Mid(Me.txtID, 2))
        Case "+"                                '住院号
            strWhere = "(select " & gConst_病人信息_列名 & " from 病人信息 a where 住院号 = [1]) e "
            strFind = Val(Mid(Me.txtID, 2))
        Case "*"                                '门诊号
            strWhere = "(select " & gConst_病人信息_列名 & " from 病人信息 a where 门诊号 = [1]) e "
            strFind = Val(Mid(Me.txtID, 2))
        Case "."                                '挂号单号
            strWhere = "(select b.病人ID,b.姓名,b.性别,b.年龄,b.门诊号,b.住院号 from 病人医嘱记录 a , 病人信息 b 　where  a.病人id = b.病人ID and 挂号单 = [1] ) e "
            strFind = Mid(Me.txtID, 2)
        Case "/"                                '收费单据号
            strWhere = "(select  b.病人ID,b.姓名,b.性别,b.年龄,b.门诊号,b.住院号 from 门诊费用记录 a, 病人信息 b " & _
                       " where No = [1] and a.病人id = b.病人id ) e "
            strFind = zlCommFun.GetFullNO(Mid(txtID, 2))
        Case Else                               '就诊卡和姓名
            strFind = Me.txtID
            If IDKind.IDKind = IDKind.GetKindIndex("姓名") And BlnIsNumber(strFind) Then
                    strWhere = "( select C.* from 病人医嘱记录 a , 病人医嘱发送 b , 病人信息 C " & _
                         " Where a.ID = b.医嘱id And a.病人ID = C.病人ID and  b.样本条码 = [1] ) e "
                    blnBarCode = True
            Else
                If mblnCard Or IDKind.IDKind = IDKind.GetKindIndex("就诊卡") Then
                    strWhere = "(select " & gConst_病人信息_列名 & " from 病人信息 a where 就诊卡号 = [1]) e "
                    strFind = UCase(Me.txtID)
                ElseIf IDKind.IDKind = IDKind.GetKindIndex("门诊号") Then
                    strWhere = "(select " & gConst_病人信息_列名 & " from 病人信息 a where 门诊号 = [1]) e "
                    strFind = Val(Mid(Me.txtID, 2))
                Else
                    If ch模糊查找.Value = 1 Then
                        strWhere = "(select " & gConst_病人信息_列名 & " from 病人信息 a where 姓名 || ''  like '%' || [1] || '%' ) e "
                        chkDate.Enabled = True
                    ElseIf Len(Me.txtID.Text) = 1 Then
                        strWhere = "(select " & gConst_病人信息_列名 & " from 病人信息 a where 姓名 || '' like  [1] || '%' ) e "
                        chkDate.Enabled = True
                    Else
                        strWhere = "(select " & gConst_病人信息_列名 & " from 病人信息 a where 姓名 like   [1] || '%' ) e "
                    End If
                End If
            End If
    End Select
    mblnCard = False
    
    gstrSql = "select /*+ rule */ distinct a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, 住院号, 检验项目, 状态, 接收人, 接收时间, 样本条码, 费用性质," & vbNewLine & _
                "       开嘱科室, 开嘱医生, 开嘱时间,发送人, 发送时间, 采样人, 采样时间, 检验人, 检验时间," & vbNewLine & _
                "       a.审核人, a.审核时间, 医嘱ID, 标本ID," & vbNewLine & _
                "       b.记录性质, b.记录状态, b.门诊标志,婴儿,年龄1,性别1, 标本序号,标本序号显示　" & vbNewLine & _
                "from (select " & vbNewLine & _
                "       distinct e.病人id, e.姓名, e.性别, e.年龄, e.门诊号, e.住院号, a.医嘱内容 as 检验项目," & vbNewLine & _
                "                decode(b.医嘱id, null, '1-未发送', decode(b.采样人, Null, '2-未采样', decode(b.接收人, null, '3-已采样', '4-已接收'))) as 状态," & vbNewLine & _
                "                b.接收人, b.接收时间, b.样本条码," & vbNewLine & _
                "                decode(b.记录性质, 1, '收费', 2, '记帐') as 费用性质, d.名称 as 开嘱科室, a.开嘱医生," & vbNewLine & _
                "                a.开嘱时间, b.发送人, b.发送时间, b.采样人, b.采样时间, '' as 检验人, '' as 检验时间," & vbNewLine & _
                "                '' as 审核人, '' as 审核时间, a.id as 医嘱Id, '' as 标本ID, b.记录性质,a.婴儿,e.性别 as 性别1,e.年龄 as 年龄1, " & vbNewLine & _
                "                '' as 标本序号显示, '' as 标本序号,a.id as 相关医嘱ID " & vbNewLine & _
                "       from 病人医嘱记录 a, 病人医嘱发送 b, 部门表 d " & "," & strWhere & vbNewLine & _
                "       where a.id = b.医嘱id(+) and a.开嘱科室id = d.id and a.病人id = e.病人id and" & vbNewLine & _
                "             b.执行状态 = 0 and a.诊疗类别 = 'C' and a.病人来源 = 2 " & vbNewLine & _
                " " & IIf(blnBarCode = True, " and b.样本条码 = [4] ", " ") & vbNewLine & _
                " " & IIf(Me.chkDate.Value = 0, "", " and a.开嘱时间 between [2] and [3] ") & vbNewLine

    gstrSql = gstrSql & "       union all " & vbNewLine & _
                "       select " & vbNewLine & _
                "       distinct e.病人id, e.姓名, e.性别, e.年龄, e.门诊号, e.住院号, a.检验项目," & vbNewLine & _
                "" & vbNewLine & _
                "                decode(样本状态, 1, '5-已核收', decode(sign(nvl(打印次数, 0)), 1, '7-已打印', '6-已审核')) as 状态," & vbNewLine & _
                "                d.接收人, d.接收时间, d.样本条码," & vbNewLine & _
                "                decode(d.记录性质, 1, '收费', 2, '记帐') as 费用性质, f.名称 as 开嘱科室, b.开嘱医生," & vbNewLine & _
                "                b.开嘱时间, d.发送人, d.发送时间, d.采样人, d.采样时间, a.检验人," & vbNewLine & _
                "                to_char(a.检验时间,'YYYY-MM-DD HH24:MI:SS') as 检验时间, a.审核人, to_char(a.审核时间,'YYYY-MM-DD HH24:MI:SS') as 审核时间," & vbNewLine & _
                "                a.医嘱ID, to_char(a.ID) as 标本ID, d.记录性质,b.婴儿,a.性别 as 性别1,a.年龄 as 年龄1, " & vbNewLine & _
                "                Decode(a.仪器id, Null," & vbNewLine & _
                "                 To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
                "                 a.标本序号) As 标本序号显示, a.标本序号,b.id 相关医嘱ID " & vbNewLine & _
                "       from 检验标本记录 a, 病人医嘱记录 b, 病人医嘱发送 d, 部门表 f " & "," & strWhere & vbNewLine & _
                "       where a.医嘱id = b.相关id and b.相关id = d.医嘱id and b.开嘱科室ID = f.id and" & vbNewLine & _
                "             a.病人id = e.病人id And a.病人来源 = 2 " & vbNewLine & _
                " " & IIf(blnBarCode = True, " And D.样本条码 = [4] ", " ") & vbNewLine & _
                " " & IIf(Me.chkDate.Value = 0, "", " and a.核收时间 between [2] and [3] ") & vbNewLine & _
                ") a, 住院费用记录 b" & vbNewLine & _
                "where a.相关医嘱ID = b.医嘱序号(+) and a.记录性质 = b.记录性质(+) and nvl(b.记录状态,0) in (0,1) "

    strSQLbak = gstrSql
    strSQLbak = Replace$(strSQLbak, "住院费用记录", "门诊费用记录")
    strSQLbak = Replace$(strSQLbak, "a.病人来源 = 2", "a.病人来源 <> 2")
    gstrSql = gstrSql & " union  " & strSQLbak & " order by 病人id, 状态, 开嘱时间, 发送时间 "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strFind, CDate(Format(Me.DTPBegin.Value, "yyyy-mm-dd 00:00:00")), _
                                         CDate(Format(Me.DTPEnd.Value, "yyyy-mm-dd 23:59:59")), strFind)
    If rsTmp.RecordCount > 0 Then Me.txtID.Text = "": Me.txtID.SetFocus
    Me.rptFind.Records.DeleteAll
    Me.rptFind.GroupsOrder.DeleteAll
    
    On Error GoTo errH
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "正在查找数据请等待....", Me
    
    Do Until rsTmp.EOF
        With Me.rptFind
            Set Record = .Records.Add
            For intLoop = 0 To .Columns.Count
                Record.AddItem ""
            Next
        End With
        If Nvl(rsTmp("状态")) = "5-已核收" Or Nvl(rsTmp("状态")) = "6-已审核" Or Nvl(rsTmp("状态")) = "7-已打印" Then
            Record.Item(mCol.选择).HasCheckbox = True
            If blnBarCode = True Then
                Record.Item(mCol.选择).Checked = True
            End If
        Else
            Record.Item(mCol.选择).HasCheckbox = False
        End If
        Record.Item(mCol.ID).Value = Nvl(rsTmp("病人id"))
        Record.Item(mCol.姓名).Value = Nvl(rsTmp("姓名")) & IIf(Nvl(rsTmp("婴儿"), 0) > 0, "(婴儿" & rsTmp("婴儿") & ")", "")
        Record.Item(mCol.性别).Value = IIf(Nvl(rsTmp("婴儿"), 0) = 0, Nvl(rsTmp("性别")), Nvl(rsTmp("性别1")))
        Record.Item(mCol.年龄).Value = IIf(Nvl(rsTmp("婴儿"), 0) = 0, Nvl(rsTmp("年龄")), Nvl(rsTmp("年龄1")))
        '分组信息
        Record.Item(mCol.ID).GroupCaption = "姓名:" & Record.Item(mCol.姓名).Value & " 性别:" & Record.Item(mCol.性别).Value & " 年龄:" & _
                                    Replace(Nvl(Record.Item(mCol.年龄).Value), "婴儿", "") & _
                                    " 门诊号:" & Nvl(rsTmp("门诊号")) & " 住院号" & Nvl(rsTmp("住院号"))
        Record.Item(mCol.门诊号).Value = Nvl(rsTmp("门诊号"))
        Record.Item(mCol.住院号).Value = Nvl(rsTmp("住院号"))
        Record.Item(mCol.检验项目).Value = Nvl(rsTmp("检验项目"))
        '收费状态
        Select Case Nvl(rsTmp("记录状态"))
            Case ""     '未收费
                Record.Item(mCol.收费状态).Value = "未收费"
            Case "0"    '划价单
                If Nvl(rsTmp("门诊标志")) = 1 Then
                    Record.Item(mCol.收费状态).Value = "门诊" & Nvl(rsTmp("费用性质")) & "(划价单)"
                Else
                    Record.Item(mCol.收费状态).Value = "住院" & Nvl(rsTmp("费用性质")) & "(划价单)"
                End If
            Case "1"    '记帐和收费完成
                If Nvl(rsTmp("门诊标志")) = 1 Then
                    Record.Item(mCol.收费状态).Value = "门诊" & Nvl(rsTmp("费用性质")) & IIf(Nvl(rsTmp("费用性质")) = "收费", "(已收费)", "(已记帐)")
                Else
                    Record.Item(mCol.收费状态).Value = "住院" & Nvl(rsTmp("费用性质")) & IIf(Nvl(rsTmp("费用性质")) = "收费", "(已收费)", "(已记帐)")
                End If
        End Select
        Record.Item(mCol.状态).Value = Nvl(rsTmp("状态"))
        Record.Item(mCol.接收人).Value = Nvl(rsTmp("接收人"))
        Record.Item(mCol.接收时间).Value = Format(Nvl(rsTmp("接收时间")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.样本条码).Value = Nvl(rsTmp("样本条码"))
        Record.Item(mCol.费用性质).Value = Nvl(rsTmp("费用性质"))
        Record.Item(mCol.开嘱科室).Value = Nvl(rsTmp("开嘱科室"))
        Record.Item(mCol.开嘱医生).Value = Nvl(rsTmp("开嘱医生"))
        Record.Item(mCol.开嘱时间).Value = Nvl(rsTmp("开嘱时间"))
        Record.Item(mCol.发送人).Value = Nvl(rsTmp("发送人"))
        Record.Item(mCol.发送时间).Value = Format(Nvl(rsTmp("发送时间")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.采样人).Value = Nvl(rsTmp("采样人"))
        Record.Item(mCol.采样时间).Value = Format(Nvl(rsTmp("采样时间")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.检验人).Value = Nvl(rsTmp("检验人"))
        Record.Item(mCol.检验时间).Value = Format(Nvl(rsTmp("检验时间")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.审核人).Value = Nvl(rsTmp("审核人"))
        Record.Item(mCol.审核时间).Value = Format(Nvl(rsTmp("审核时间")), "yyyy-mm-dd hh:mm:ss")
        Record.Item(mCol.医嘱id).Value = Nvl(rsTmp("医嘱ID"))
        Record.Item(mCol.标本id).Value = Nvl(rsTmp("标本ID"))
        Record.Item(mCol.标本号).Value = Nvl(rsTmp("标本序号"))
        Record.Item(mCol.标本号).Caption = Nvl(rsTmp("标本序号显示"))
        rsTmp.MoveNext
    Loop
    For intLoop = 0 To Me.rptFind.Columns.Count - 1
        If Me.rptFind.Columns(intLoop).Caption = "病人信息" Then
            Me.rptFind.GroupsOrder.Add Me.rptFind.Columns(intLoop)
            Me.rptFind.GroupsOrder(0).Editable = True
        End If
    Next
    For intLoop = 0 To Me.rptFind.Columns.Count - 1
        If Me.rptFind.Columns(intLoop).Caption = "状态" Then
            Me.rptFind.GroupsOrder.Add Me.rptFind.Columns(intLoop)
            Me.rptFind.GroupsOrder(1).Editable = True
        End If
    
    Next
    Me.rptFind.Populate
    '
    For Each GroupRow In rptFind.Rows
        If GroupRow.GroupRow = False Then
            If GroupRow.Record(mCol.状态).Value = "1-未发送" Or GroupRow.Record(mCol.状态).Value = "2-未采样" Or _
                GroupRow.Record(mCol.状态).Value = "3-已采样" Or GroupRow.Record(mCol.状态).Value = "4-已接收" Then
                GroupRow.ParentRow.Expanded = False
            End If
        End If
    Next
    
    zlCommFun.StopFlash
    
    If IDKind.IDKind = IDKinds.C5就诊卡 And Me.txtID.Tag <> "" Then
        txtID = txtID.Tag
        txtID.Tag = ""
    End If
    
    Me.MousePointer = 0
    
    Exit Sub
errH:
    zlCommFun.StopFlash
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With Me.Frame1
        .Left = 10
        .Width = Me.ScaleWidth - 10
    End With
    
    With Me.cmdExit
        .Top = Me.ScaleHeight - 500
        .Left = Me.ScaleWidth - .Width - 300
    End With
    
    With Me.cmdPrint
        .Top = Me.cmdExit.Top
        .Left = Me.cmdExit.Left - .Width - 400
    End With
    
    With Me.cmdUnionPrint
        .Top = Me.cmdExit.Top
        .Left = Me.cmdPrint.Left - .Width - 400
    End With
    
    With Me.cmdPreview
        .Top = Me.cmdExit.Top
        .Left = Me.cmdUnionPrint.Left - .Width - 400
    End With
    
    With Me.cmdSetupPrint
        .Top = Me.cmdExit.Top
        .Left = Me.cmdPreview.Left - .Width - 400
    End With
    
    With Me.rptFind
        .Width = Me.ScaleWidth
        .Height = Me.cmdExit.Top - .Top - 150
    End With
    
    Me.chkSelect(0).Top = Me.rptFind.Top + Me.rptFind.Height + 20
    Me.chkSelect(1).Top = Me.rptFind.Top + Me.rptFind.Height + 20
    Me.chkSelect(2).Top = Me.rptFind.Top + Me.rptFind.Height + 20
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Set mobjICCard = Nothing
    Set mobjSquareCard = Nothing
    zlDatabase.SetPara "frmLabMainFindRePort_使用时间范围", Me.chkDate.Value, 100, 1208
    zlDatabase.SetPara "frmLabMainFindRePort_rptFind", Me.rptFind.SaveSettings, 100, 1208
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "输入方式", IDKind.IDKind)

End Sub

Private Sub IDKind_Click()
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand As String, strOutPatiInforXML As String
    If IDKind.IDKind = IDKind.GetKindIndex("IC卡号") Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtID.Text = mobjICCard.Read_Card()
            If txtID.Text <> "" Then Call RefreshData
        End If
    End If
    lng卡类别ID = Val(IDKind.GetKindItem("卡类别ID"))
    If lng卡类别ID = 0 Then Exit Sub
    
    If mobjSquareCard.zlReadCard(Me, glngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtID.Text = strOutCardNO
    If txtID.Text <> "" Then Call txtID_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer)
    mblnShowPwd = Trim(IDKind.GetKindItem(7)) <> ""
    Me.txtID = ""
    If mblnShowPwd = True Then
        Me.txtID.PasswordChar = "*"
    Else
        Me.txtID.PasswordChar = ""
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    mbln身份证 = False
    If Not txtID.Locked And txtID.Text = "" And Me.ActiveControl Is txtID Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKinds.C2身份证号
        txtID.Text = strID
        mbln身份证 = True
        Call RefreshData
        mbln身份证 = False
        IDKind.IDKind = lngPreIDKind
    End If
End Sub

Private Sub rptFind_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    On Error Resume Next
    If Item.Record(mCol.状态).Value = "5-已核收" Or Item.Record(mCol.状态).Value = "6-已审核" Or Item.Record(mCol.状态).Value = "7-已打印" Then
        Item.Record(mCol.选择).Checked = Not Item.Record(mCol.选择).Checked
        Me.rptFind.Redraw
    End If
End Sub

Private Sub rptFind_SelectionChanged()
    If Me.rptFind.FocusedRow Is Nothing Then Me.cmdPrint.Enabled = False: Me.cmdSetupPrint.Enabled = False: Exit Sub
    If Me.rptFind.FocusedRow.GroupRow = True Then Me.cmdPrint.Enabled = False: Me.cmdSetupPrint.Enabled = False: Exit Sub

    If Me.rptFind.FocusedRow.Record(mCol.状态).Value = "5-已核收" Or Me.rptFind.FocusedRow.Record(mCol.状态).Value = "6-已审核" _
       Or Me.rptFind.FocusedRow.Record(mCol.状态).Value = "7-已打印" Then
        Me.cmdPrint.Enabled = True
    Else
        Me.cmdPrint.Enabled = False
    End If
    Me.cmdSetupPrint.Enabled = True
End Sub

Private Sub txtID_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtID.Text = "" And Me.ActiveControl Is txtID)
End Sub

Private Sub txtID_GotFocus()
    If Not mobjIDCard Is Nothing And txtID.Text = "" And Not txtID.Locked Then mobjIDCard.SetEnabled (True)
    txtID.SelStart = 0
    txtID.SelLength = Len(txtID.Text)
    txtID.SetFocus
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'‘’;；:：?？|,，.。""") = True Then KeyAscii = 0
    blnCard = False
    If IDKind.IDKind = IDKind.GetKindIndex("姓名") Then
'        mblnCard = zlCommFun.InputIsCard(txtID, KeyAscii, False)
        If mblnCard = False And KeyAscii = 13 Then
            KeyAscii = 0
            cmdFind_Click
        End If
    End If
    If IDKind.IDKind = IDKind.GetKindIndex("就诊卡") Then
'        Call zlCommFun.InputIsCard(txtID, KeyAscii, True)
        gbytCardNOLen = Val(IDKind.GetKindItem("卡号长度", IDKind.IDKind))
        blnCard = KeyAscii <> 8 And Len(txtID.Text) = gbytCardNOLen - 1 And txtID.SelLength <> Len(txtID.Text)
        If blnCard = True Then
            If KeyAscii <> 13 Then
                Me.txtID = Me.txtID & Chr(KeyAscii)
            End If
            KeyAscii = 0
            cmdFind_Click
        End If
    End If
    If KeyAscii = 13 Or (IDKind.IDKind = IDKind.GetKindIndex("就诊卡") And blnCard = True) Then
        Call cmdFind_Click
    End If
End Sub

Private Sub txtID_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub
Private Sub ReportPrint(ByVal intIndex As Integer, ByVal blnPrint As Boolean)
    '单个报告打印
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lng医嘱ID As Long, lng发送号 As Long, lng病人ID As Long
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    Dim lngKey As Long
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "正在打印请等待...", Me
    
    If Me.rptFind.Rows(intIndex) Is Nothing Then
        zlCommFun.StopFlash
        Me.MousePointer = 0
        Exit Sub
    End If
    
    lng医嘱ID = Val(Me.rptFind.Rows(intIndex).Record(mCol.医嘱id).Value)
    lng病人ID = Val(Me.rptFind.Rows(intIndex).Record(mCol.ID).Value)
    lngKey = Val(Me.rptFind.Rows(intIndex).Record(mCol.标本id).Value)
    
    '生成图形供自定义报表调用
    strSQL = "select id from 检验图像结果 where 标本id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lngKey)
    intLoop = 1
    Do Until rsTmp.EOF
        strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
        Call LoadImageData(App.path, rsTmp("ID"))
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    
    If GetReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lng医嘱ID, _
                        "病人ID=" & lng病人ID, "标本ID=" & lngKey, "多个医嘱=" & lng医嘱ID, "多个标本=" & lngKey, _
                        "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
                        "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
                        "图形9=" & strChart(9), IIf(blnPrint, 2, 1))
    End If
    
    
    On Error GoTo errH
    If blnPrint = True Then
        If Me.rptFind.Rows(intIndex).Record(mCol.状态).Value = "6-已审核" Then
            If mintUnion = 1 Then
                gstrSql = " select id from 检验标本记录 where 医嘱id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng医嘱ID)
                Do Until rsTmp.EOF
                    strSQL = "ZL_检验标本记录_标本质控(" & rsTmp("ID") & ",'',1)"
                    zlDatabase.ExecuteProcedure strSQL, gstrSysName
                    rsTmp.MoveNext
                Loop
            Else
                strSQL = "ZL_检验标本记录_标本质控(" & lngKey & ",'',1)"
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
            End If
        End If
    End If
    Me.MousePointer = 0
    zlCommFun.StopFlash
    
    On Error Resume Next
    '删除图形文件
    For intLoop = 1 To 9
        Kill strChart(intLoop)
    Next
    Exit Sub
errH:
    Me.MousePointer = 0
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(lngKey As Long, Objfrm As Object, strPrivs As String)
    '显示病人病人报告窗体
    mstrPrivs = strPrivs
    Me.Show , Objfrm
    Me.txtID.Text = "-" & lngKey
    Call cmdFind_Click
    Me.txtID.Text = ""
End Sub
Private Sub PrintSetup()
    '打印设置
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lng医嘱ID As Long, lng发送号 As Long, lng病人ID As Long
    Dim strSQL As String
    Dim intLoop  As Integer
    
   
    If Me.rptFind.FocusedRow Is Nothing Then
        MsgBox "选择一个记录后才能设置！", vbInformation, Me.Caption
        Exit Sub
    End If
    lng医嘱ID = Val(rptFind.FocusedRow.Record(mCol.医嘱id).Value)
'    lng病人id = Val(rptList.FocusedRow.Record(mCol.病人ID).Value)
'
'    strsql = "select 发送号 from 病人医嘱发送 a , 病人医嘱记录 b where b.id = a.医嘱id and b.id = [1]"
'    Set rsTmp = zldatabase.OpenSQLRecord(strsql, gstrSysName, lng医嘱ID)
'    If rsTmp.EOF = False Then
'        lng发送号 = Nvl(rsTmp(0))
'    End If
    
    If GetReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        ReportPrintSet gcnOracle, glngSys, strReportCode, Me
        
    End If
End Sub
Private Sub AllReportPrint()
    '单个报告打印
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lng医嘱ID As Long, lng发送号 As Long, lng病人ID As Long
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    Dim lngKey As Long
    Dim str医嘱ID As String                 '医嘱ID，多个医嘱ID使用","分隔。
    Dim str标本ID As String                 '标本ID, 多个标本ID使用","分隔。
    Dim strPrintCode As String              '单据编码
    Dim intItem As Integer
    Dim astrItem() As String
        
    
    If Me.rptFind.Rows.Count = 0 Then Exit Sub
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "正在打印请等待...", Me
    For intLoop = 0 To Me.rptFind.Rows.Count - 1
        If Me.rptFind.Rows(intLoop).GroupRow = False Then
            If Me.rptFind.Rows(intLoop).Record(mCol.选择).Checked = True Then
                If Me.rptFind.Rows(intLoop).Record(mCol.审核人).Value = "" Then
                    If InStr(mstrPrivs, "未审核打印") <= 0 Then
                        MsgBox "你没有<未审核打印>权限，不能打印未审核单据!"
                        Me.MousePointer = 1
                        zlCommFun.StopFlash
                        Exit Sub
                    End If
                End If
                str医嘱ID = str医嘱ID & "," & Me.rptFind.Rows(intLoop).Record(mCol.医嘱id).Value
                str标本ID = str标本ID & "," & Me.rptFind.Rows(intLoop).Record(mCol.标本id).Value
                lng病人ID = Me.rptFind.Rows(intLoop).Record(mCol.ID).Value
            End If
        End If
    Next
    If str医嘱ID <> "" Then
        str医嘱ID = Mid(str医嘱ID, 2)
        lng医嘱ID = Split(str医嘱ID, ",")(0)
    End If
    If str标本ID <> "" Then
        str标本ID = Mid(str标本ID, 2)
        lngKey = Split(str标本ID, ",")(0)
    End If
    
    '有多个格式时得到格式
    frmLabMainPrintFormat.ShowMe Me, str医嘱ID, strPrintCode
    
    '生成图形供自定义报表调用
    astrItem = Split(Mid(str标本ID, 2), ",")
    intLoop = 1
    For intItem = 0 To UBound(astrItem)
        If intLoop >= 9 Then Exit For
        frmLabMain.ReadImageData CLng(astrItem(intItem)), True
        strSQL = "select id from 检验图像结果 where 标本id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, astrItem(intItem))
        Do Until rsTmp.EOF
            If intLoop < 9 Then
                strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                intLoop = intLoop + 1
            End If
            rsTmp.MoveNext
        Loop
    Next
    
    Call ReportOpen(gcnOracle, glngSys, strPrintCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & str医嘱ID, _
                        "病人ID=" & lng病人ID, "标本ID=" & str标本ID, "多个医嘱=" & str医嘱ID, "多个标本=" & str标本ID, _
                        "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
                        "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
                        "图形9=" & strChart(9), 2)


    On Error GoTo errH
    For intLoop = 0 To Me.rptFind.Rows.Count - 1
        If Me.rptFind.Rows(intLoop).GroupRow = False Then
            If Me.rptFind.Rows(intLoop).Record(mCol.状态).Value = "6-已审核" Then
                If mintUnion = 1 Then
                    gstrSql = " select id from 检验标本记录 where 医嘱id = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng医嘱ID)
                    Do Until rsTmp.EOF
                        strSQL = "ZL_检验标本记录_标本质控(" & rsTmp("ID") & ",'',1)"
                        zlDatabase.ExecuteProcedure strSQL, gstrSysName
                        rsTmp.MoveNext
                    Loop
                Else
                    strSQL = "ZL_检验标本记录_标本质控(" & lngKey & ",'',1)"
                    zlDatabase.ExecuteProcedure strSQL, gstrSysName
                End If
            End If
        End If
    Next
    Me.MousePointer = 0
    zlCommFun.StopFlash
    
    On Error Resume Next
    '删除图形文件
    For intLoop = 1 To 9
        Kill strChart(intLoop)
    Next
    Exit Sub
errH:
    Me.MousePointer = 0
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "日期"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "日期时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "可打印字符"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function
