VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmLabBarCodeBatPrint 
   Caption         =   "病区条码打印"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   9390
   Icon            =   "frmLabBarCodeBatPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   9390
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl RptItem 
      Height          =   3615
      Left            =   60
      TabIndex        =   10
      Top             =   2790
      Width           =   9255
      _Version        =   589884
      _ExtentX        =   16325
      _ExtentY        =   6376
      _StockProps     =   0
      BorderStyle     =   3
      MultipleSelection=   0   'False
      SkipGroupsFocus =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdPrintSend 
      Caption         =   "打印送检单(&S)"
      Height          =   350
      Left            =   3075
      TabIndex        =   28
      Top             =   6570
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Frame fraBarCode 
      Caption         =   "使用条码输入"
      Height          =   645
      Left            =   60
      TabIndex        =   24
      Top             =   2100
      Width           =   9285
      Begin VB.TextBox txtBarCode 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1170
         TabIndex        =   25
         Top             =   210
         Width           =   7995
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "请扫描条码"
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "打印设置(&S)"
      Height          =   350
      Left            =   120
      TabIndex        =   22
      Top             =   6570
      Width           =   1395
   End
   Begin VB.CommandButton cmdReturBill 
      Caption         =   "标本送检(&I)"
      Height          =   350
      Left            =   4665
      TabIndex        =   21
      Top             =   6570
      Width           =   1395
   End
   Begin VB.PictureBox picBarCodePrint 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   210
      ScaleHeight     =   405
      ScaleWidth      =   675
      TabIndex        =   18
      Top             =   6540
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7830
      TabIndex        =   2
      Top             =   6570
      Width           =   1395
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   6255
      TabIndex        =   1
      Top             =   6570
      Width           =   1395
   End
   Begin VB.Frame fraFilter 
      Caption         =   "过滤"
      Height          =   2085
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   9285
      Begin VB.CheckBox chkCodePrint 
         Caption         =   "提示条码打印"
         Height          =   180
         Left            =   2700
         TabIndex        =   29
         ToolTipText     =   "完成采集时提示是否需要打印条码"
         Top             =   1710
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.TextBox txtUnit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   27
         Top             =   1050
         Width           =   7845
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "隐藏已打印"
         Height          =   180
         Left            =   1320
         TabIndex        =   20
         Top             =   1710
         Width           =   1275
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "选   择"
         Height          =   180
         Left            =   270
         TabIndex        =   19
         Top             =   1710
         Value           =   1  'Checked
         Width           =   945
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   240
         TabIndex        =   17
         Top             =   1500
         Width           =   8715
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         ItemData        =   "frmLabBarCodeBatPrint.frx":000C
         Left            =   7020
         List            =   "frmLabBarCodeBatPrint.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   660
         Width           =   1905
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1905
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   7830
         TabIndex        =   0
         Top             =   1590
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   675
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   223346691
         CurrentDate     =   39064.0416666667
      End
      Begin VB.ComboBox cboSample 
         Height          =   300
         Left            =   3990
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1905
      End
      Begin VB.ComboBox cboCapture 
         Height          =   300
         Left            =   7020
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   285
         Left            =   3810
         TabIndex        =   7
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   223346691
         CurrentDate     =   39064.0416666667
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "工作单位"
         Height          =   180
         Left            =   270
         TabIndex        =   23
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "执行状态"
         Height          =   180
         Left            =   6210
         TabIndex        =   16
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "-----"
         Height          =   180
         Left            =   3270
         TabIndex        =   15
         Top             =   705
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "科    室"
         Height          =   180
         Left            =   270
         TabIndex        =   14
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "发送时间"
         Height          =   180
         Left            =   270
         TabIndex        =   13
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "标    本"
         Height          =   180
         Left            =   3180
         TabIndex        =   12
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "采集方式"
         Height          =   180
         Left            =   6210
         TabIndex        =   11
         Top             =   300
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ImageListReport 
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBarCodeBatPrint.frx":0010
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabBarCodeBatPrint.frx":007C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLabBarCodeBatPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    类别 = 0
    医嘱id
    相关ID
    选择
    图标
    采集方式
    医嘱内容
    标本
    姓名
    性别
    年龄
    科室
    床号
    标识号
    条码
    病人ID
    管码
    试管颜色
    合并医嘱
    执行科室
    开嘱医生
    开嘱时间
    采样人
    采样时间
    采血量
    试管名称
    紧急
    病人来源
    婴儿
    别名
    条码打印
    送出时间
    诊疗项目ID
    检验执行科室ID
    开嘱科室
End Enum
Dim BlCancel As Boolean                             '当按下"ESC"键时忽略错误
Private mstrPrivs As String                         '权限
Private mintBarCodeFormat As Integer                '条码打印格式 1=39Code 2=128Code
Private mintExecDept As Integer                     '不区分执行科室打印
Private mblnNowConsumption As Boolean                                   '是否立即付款

Private Sub CrateRptHead()
    '功能           初始化列表头
    Dim Column As ReportColumn
    With Me.RptItem.Columns
        
        RptItem.AllowColumnRemove = False
        RptItem.ShowItemsInGroups = False
        Me.RptItem.SetImageList Me.ImageListReport
        With RptItem.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "请选择好过滤条件，单击查找按钮..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mCol.类别, "类别", 120, False): Column.Visible = False
        Set Column = .Add(mCol.选择, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mCol.图标, "试管", 18, False): Column.Icon = 1
        Set Column = .Add(mCol.姓名, "姓名", 60, True)
        Set Column = .Add(mCol.采集方式, "采集方式", 130, True)
        Set Column = .Add(mCol.医嘱内容, "医嘱内容", 100, True)
        Set Column = .Add(mCol.标本, "标本", 45, True)
        Set Column = .Add(mCol.性别, "性别", 45, True)
        Set Column = .Add(mCol.年龄, "年龄", 45, True)
        Set Column = .Add(mCol.科室, "科室", 45, True)
        Set Column = .Add(mCol.床号, "床号", 45, True)
        Set Column = .Add(mCol.标识号, "标识号", 55, True)
        Set Column = .Add(mCol.条码, "条码", 120, True)
        
        Set Column = .Add(mCol.医嘱id, "医嘱ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.相关ID, "相关ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.病人ID, "病人ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.管码, "管码", 65, True): Column.Visible = False
        Set Column = .Add(mCol.合并医嘱, "合并医嘱", 65, True): Column.Visible = False
        Set Column = .Add(mCol.试管颜色, "试管颜色", 75, True): Column.Visible = False
        Set Column = .Add(mCol.执行科室, "执行科室", 65, True): Column.Visible = False
        Set Column = .Add(mCol.开嘱医生, "开嘱医生", 65, True): Column.Visible = False
        Set Column = .Add(mCol.开嘱时间, "开嘱时间", 65, True): Column.Visible = False
        Set Column = .Add(mCol.采样人, "采样人", 65, True): Column.Visible = False
        Set Column = .Add(mCol.采样时间, "采样时间", 65, True): Column.Visible = False
        Set Column = .Add(mCol.采血量, "采血量", 65, True): Column.Visible = False
        Set Column = .Add(mCol.试管名称, "试管名称", 65, True): Column.Visible = False
        Set Column = .Add(mCol.紧急, "紧急", 65, True): Column.Visible = False
        Set Column = .Add(mCol.病人来源, "来源", 65, True): Column.Visible = False
        Set Column = .Add(mCol.婴儿, "婴儿", 65, True): Column.Visible = False
        Set Column = .Add(mCol.条码打印, "打印", 65, True) ': Column.Visible = False
        Set Column = .Add(mCol.送出时间, "送出时间", 130, True)
        Set Column = .Add(mCol.诊疗项目ID, "诊疗项目ID", 130, True): Column.Visible = False
        Set Column = .Add(mCol.检验执行科室ID, "检验执行科室ID", 130, True): Column.Visible = False
        Set Column = .Add(mCol.开嘱科室, "开嘱科室", 130, True): Column.Visible = False
    End With
End Sub

Private Sub cboState_Click()
    Me.chkCodePrint.Visible = False
    Me.cmdPrintSend.Visible = False
    Me.cmdReturBill.Visible = (Me.cboState.Text = "已采样")
    Me.cmdPrint.Visible = True
    
    Select Case cboState.Text
        Case "未绑定"
            Me.cmdPrint.Caption = "生成条码(&B)"
            Me.cmdReturBill.Visible = (InStr(mstrPrivs, "完成采集") > 0)
            Me.cmdReturBill.Caption = "完成采集(&M)"
        Case "已绑定"
            Me.cmdPrint.Visible = (InStr(mstrPrivs, "完成采集") > 0)
            Me.cmdPrint.Caption = "完成采集(&F)"
            Me.cmdReturBill.Visible = True
            Me.cmdReturBill.Caption = "打印条码(&P)"
            Me.chkCodePrint.Visible = True
        Case "已采样"
            Me.cmdPrint.Caption = "打印条码(&P)"
            Me.cmdReturBill.Visible = True
            Me.cmdReturBill.Caption = "标本送检(&I)"
        Case "已送检"
            Me.cmdPrint.Caption = "打印条码(&P)"
            Me.cmdReturBill.Visible = True
            Me.cmdReturBill.Caption = "取消送检(&I)"
            Me.cmdPrintSend.Visible = True
        Case "已执行"
            Me.cmdPrint.Caption = "打印条码(&P)"
    End Select
    
    Call Form_Resize
    Me.TxtBarCode.Tag = ""
    Call ReadData
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkNOComple_Click()
    
End Sub

Private Sub chkPrinted_Click()
    Call cmdSelectAll_Click
End Sub

Private Sub chkPrintNO_Click()
    Call cmdSelectAll_Click
End Sub

Private Sub chkPrint_Click()
    SelectOrCancelReprotCheck Me.RptItem.Records, mCol.选择, Me.chkSelect.Value
End Sub

Private Sub chkSelect_Click()
    SelectOrCancelReprotCheck Me.RptItem.Records, mCol.选择, Me.chkSelect.Value
End Sub

Private Sub cmdCancel_Click()
    BlCancel = True
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    SelectOrCancelReprotCheck Me.RptItem.Records, mCol.选择, False
    Me.RptItem.Populate
End Sub

Private Sub cmdFind_Click()
    Me.TxtBarCode.Tag = ""
    ReadData
    Call cmdSelectAll_Click
End Sub

Private Sub cmdPrint_Click()
    Select Case cboState.Text
        Case "未绑定"
            BarCodeMake 1, True
        Case "已绑定"
            If chkCodePrint.Value = 1 Then
                If MsgBox("是否需要打印条码?", vbYesNo + vbDefaultButton2) = vbYes Then
                    BarCodeMake 2, True
                Else
                    BarCodeMake 2, False
                End If
            Else
                BarCodeMake 2, False
            End If
        Case "已采样", "已送检"
            BarCodeMake 6, True
        Case "已执行"
            BarCodeMake 6, True
    End Select
End Sub

Private Sub cmdPrintSend_Click()
    Dim strName As String
    Dim strID As String
    Dim strTemp As String
    Dim strAdvices As String
    Dim intLoop As Integer
    Dim astrReprot() As String
    
    With Me.RptItem
        If Me.cboState.Text = "已送检" And .Rows.Count > 0 Then
            frmLabSamplingSendInfo.chkPrint.Enabled = False
            If frmLabSamplingSendInfo.ShowME(Me, strName, True) = False Then
                Exit Sub
            End If
            
            For intLoop = 0 To .Rows.Count - 1
                If .Rows(intLoop).Record(mCol.选择).Checked = True Then
                    If Len(strID) >= 3800 Then  '字符串超长分段处理
                        strTemp = strTemp & ";" & Mid(strID, 2)
                        strID = ""
                    Else
                        strID = strID & "," & .Rows(intLoop).Record(mCol.相关ID).Value
                    End If
                End If
            Next
            
            strAdvices = strTemp & ";" & Mid(strID, 2)
            astrReprot = Split(strAdvices, ";")
            For intLoop = 0 To UBound(astrReprot)
                If astrReprot(intLoop) <> "" Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_3", Me, "医嘱字串=" & astrReprot(intLoop), 2)
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdPrintSet_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1211_3", Me
End Sub

Private Sub cmdReturBill_Click()
    Select Case cboState.Text
        Case "未绑定"
            BarCodeMake 3, True
        Case "已绑定"
            BarCodeMake 6, True
        Case "已采样", "已送检"
            SampleSend
        Case "已执行"
            BarCodeMake 6, True
    End Select
End Sub

Private Sub cmdSelectAll_Click()
    SelectOrCancelReprotCheck Me.RptItem.Records, mCol.选择, True
    Me.RptItem.Populate
End Sub

Private Sub Form_Load()
    '创建列表头
    CrateRptHead
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '读入初使化数据
    GetInitData
End Sub
Private Sub ReadData(Optional ByVal strBarCode As String)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim Record As ReportRecord
    Dim Item As ReportRecordItem
    Dim i As Integer
    Dim strUnionID As Long                          '相关Id
    Dim blnShowExec As Boolean                      '是否显示未收费门诊病人
    Dim strSQLbak As String
    Dim strSQLCheck As String
    Dim str医嘱内容 As String
    Dim blnFL As Boolean                            '是否分类显示,存在输血医嘱时,分类显示
    Dim strBooldSql As String                       '输血查询语句
    zlCommFun.ShowFlash "正在查找,请稍候...", Me
    Me.MousePointer = vbHourglass
    
    On Error GoTo errH
    
    strSQL = "Select /*+ rule */ a.类别,A.医嘱id, A.相关id, A.医嘱内容, A.发送时间, A.执行科室, A.姓名, A.性别, A.年龄, A.样本条码, A.标本, A.病人id, A.采集方式, A.管码, A.当前床号, A.标识号," & vbNewLine & _
            "       A.开嘱医生, A.开嘱时间, A.采样人, A.采样时间, A.紧急, A.病人来源, A.婴儿, A.别名, A.组合项目, A.条码打印, A.记录状态, A.记录性质, A.重采, A.标本送出时间," & vbNewLine & _
            "       B.颜色 As 试管颜色, B.采血量, B.名称 As 试管名称,a.诊疗项目ID,a.检验执行科室ID,a.开嘱科室 from  " & vbCrLf & _
             " (Select  distinct decode(d.类别,'K','输血','检验') 类别,B.Id as 医嘱ID,B.相关Id,decode(d.类别,'K',d.名称,b.医嘱内容) as 医嘱内容,m.发送时间,H.名称 as 执行科室,I.姓名,I.性别,I.年龄,m.样本条码,b.标本部位 as 标本, " & vbCrLf & _
             " I.病人ID,decode(d.类别,'K',b.医嘱内容 ,d.名称) As 采集方式,E.试管编码 as  管码,b.诊疗项目ID,b.执行科室ID as 检验执行科室ID, " & vbCrLf & _
             "decode(i.当前床号,null,decode(l.出院病床,null,l.入院病床,l.出院病床),i.当前床号) as 当前床号, " & vbCrLf & _
             "Decode(B.病人来源,1,I.门诊号,2,i.住院号,4,i.门诊号) as 标识号, " & vbCrLf & _
             " b.开嘱医生,b.开嘱时间,m.采样人,m.采样时间,decode(b.紧急标志,1,'紧急','') as 紧急, " & _
             " b.病人来源,b.婴儿,n.名称 as 别名,E.组合项目,c.条码打印,nvl(P.记录状态,0) as 记录状态, " & vbCrLf & _
             " nvl(P.记录性质,0) as 记录性质,nvl(c.重采标本,0) as 重采,c.标本送出时间,Q.名称 as 开嘱科室,decode(d.类别, 'K', M.执行状态,C.执行状态) 执行状态 " & vbCrLf & _
             " From 病人医嘱记录 A, 病人医嘱记录 B,病人医嘱发送 C,诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
             "      部门表 H,病人信息 I,病案主页 L,病人医嘱发送 M, " & vbCrLf & _
             " (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N,住院费用记录 P,部门表 Q " & vbCrLf & _
             " Where A.ID = B.相关id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID And (e.类别 = 'E' Or e.类别 = 'C') And " & vbCrLf & _
             " B.执行科室id = H.ID And d.类别 = 'E' And d.操作类型 = '6'  And " & vbCrLf & _
             " A.病人Id = I.病人ID And a.病人id = l.病人ID(+) and a.主页id = l.主页id(+) and  m.执行部门id + 0 = [1] and M.发送时间 Between [2] And [3] " & vbCrLf & _
             " And A.id = M.医嘱ID And b.id = c.医嘱ID and E.id = N.诊疗项目ID(+)  " & vbCrLf & _
             " And C.医嘱ID = P.医嘱序号(+) and C.记录性质 = Mod(P.记录性质(+),10) and b.开嘱科室ID = Q.ID " & vbCrLf
    
    '标本
    If cboSample.ItemData(cboSample.ListIndex) <> 0 Then
        strSQL = strSQL & " And B.标本部位= [4] "
    End If

    If Me.cboCapture.ItemData(cboCapture.ListIndex) <> 0 Then
        strSQL = strSQL & " And D.名称 = [5] "
    End If
    
    '病人单位
    If Trim(Me.txtUnit.Text) <> "" Then
        strSQL = strSQL & " and I.工作单位 like [6] "
    End If
    
    '条码
    If strBarCode <> "" Then
        strSQL = strSQL & " And m.样本条码 = [7]"
    End If
    
    If Me.cboState = "未绑定" Then
        strSQL = strSQL & " And c.样本条码 is null) a , 采血管类型 b "
    ElseIf Me.cboState = "已绑定" Then
        strSQL = strSQL & " And c.样本条码 is not null and c.采样人 is null) a , 采血管类型 B "
    ElseIf Me.cboState = "已采样" Then
        strSQL = strSQL & " and c.样本条码 is not null and c.采样人 is not null and c.标本送出时间 is null) a,采血管类型 B  "
    ElseIf Me.cboState = "已送检" Then
        strSQL = strSQL & " and c.样本条码 is not null and c.采样人 is not null and c.标本送出时间 is not null) a,采血管类型 B  "
    ElseIf Me.cboState = "已执行" Then
        strSQL = strSQL & ") a,采血管类型 B "
    End If
    strSQL = strSQL & " Where a.管码 = b.编码 "
    If Me.cboState = "已执行" Then
        strSQL = strSQL & " and a.执行状态 IN (1,3)"
    Else
        strSQL = strSQL & " and a.执行状态 IN (0,2)"
    End If
    strSQLbak = strSQL
    
    strSQLCheck = strSQL
    
    If Me.cboState = "未绑定" And mblnNowConsumption = True Then
        '单独处理需要费用确认的内容
        strSQLCheck = Replace$(strSQLCheck, "住院费用记录", "门诊费用记录")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQLCheck, gstrSysName, cboDept.ItemData(cboDept.ListIndex), _
                CDate(Format(DtpBegin.Value, "yyyy-MM-dd hh:mm:00")), _
                CDate(Format(DTPEnd.Value, "yyyy-MM-dd hh:mm:59")), Mid(cboSample.Text, InStr(1, cboSample.Text, "-") + 1), _
                cboCapture.Text, "%" & Me.txtUnit & "%", strBarCode)
        If rsTmp.RecordCount > 0 Then
            rsTmp.filter = "记录状态 = 0 and 病人来源 <> 2 "
            If rsTmp.RecordCount > 0 Then
                MsgBox "记录中有门诊费用未确认的病人，不能进行批量处理!", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
    End If
    
    strBooldSql = GetBooldReadDAtaSql(strBarCode)
    strSQLbak = Replace$(strSQLbak, "住院费用记录", "门诊费用记录")
    strSQL = strSQL & " union all " & strSQLbak
    
    '排序
'    strSQL = strSQL & " order by  病人ID,E.试管编码,b.标本部位,a.医嘱内容,B.相关Id,b.开嘱时间,e.组合项目 "
    strSQL = strSQL & " Order By 类别,病人id, 管码,相关id,执行科室, 标本,紧急, 医嘱内容,  开嘱时间, 组合项目 "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, cboDept.ItemData(cboDept.ListIndex), _
                CDate(Format(DtpBegin.Value, "yyyy-MM-dd hh:mm:00")), _
                CDate(Format(DTPEnd.Value, "yyyy-MM-dd hh:mm:59")), Mid(cboSample.Text, InStr(1, cboSample.Text, "-") + 1), _
                cboCapture.Text, "%" & Me.txtUnit & "%", strBarCode)
    
    If strBarCode = "" Then
        RptItem.Records.DeleteAll
        RptItem.GroupsOrder.DeleteAll
    Else
        If rsTmp.EOF Then
            MsgBox "未查询到条码号为【" & strBarCode & "】的标本！"
            TxtBarCode.SetFocus
            Me.MousePointer = vbDefault
            zlCommFun.StopFlash
            Exit Sub
        Else
            For Each Record In RptItem.Records
                If Record.Item(mCol.条码).Value = strBarCode Then
                    MsgBox "条码号为【" & strBarCode & "】的标本已存在！"
                    TxtBarCode.SetFocus
                    Me.MousePointer = vbDefault
                    zlCommFun.StopFlash
                    Exit Sub
                End If
            Next
        End If
    End If
                    
    With Me.RptItem
        Do Until rsTmp.EOF
            blnShowExec = True
            '跟据权限来判断是否显示未收费的门诊记录
            If InStr(mstrPrivs, "显示划价记录") <= 0 Then
                If Nvl(rsTmp("记录状态"), 0) = 0 And ((Nvl(rsTmp("记录性质"), 0) Mod 10) = 1 Or Nvl(rsTmp("记录性质"), 0) = 0) Then blnShowExec = False
            End If
            
            If blnShowExec = True Then
                If strUnionID <> Nvl(rsTmp("相关ID")) Then
                
                    Set Record = .Records.Add
                    For i = 0 To .Columns.Count
                        Record.AddItem ""
                    Next
                    
                    Set Item = Record(mCol.选择): Item.HasCheckbox = True: Item.Checked = True
                    Record(mCol.图标).BackColor = Nvl(rsTmp("试管颜色"), -1)
                    Record(mCol.标本).Value = Nvl(rsTmp("标本"))
                    Record(mCol.采集方式).Value = Nvl(rsTmp("采集方式"))
                    Record(mCol.科室).Value = Nvl(rsTmp("执行科室"))
                    Record(mCol.执行科室).Value = Nvl(rsTmp("执行科室"))
                    Record(mCol.检验执行科室ID).Value = Nvl(rsTmp("检验执行科室ID"))
                    Record(mCol.年龄).Value = Nvl(rsTmp("年龄"))
                    Record(mCol.条码).Value = Nvl(rsTmp("样本条码"))
                    Record(mCol.姓名).Value = Nvl(rsTmp("姓名")) & IIf(Nvl(rsTmp("婴儿"), "0") = 0, "", "(婴儿" & rsTmp("婴儿") & ")")
                    Record(mCol.性别).Value = Nvl(rsTmp("性别"))
                    Record(mCol.医嘱id).Value = Nvl(rsTmp("医嘱ID"))
                    Record(mCol.相关ID).Value = Nvl(rsTmp("相关ID"))
                    Record(mCol.医嘱内容).Value = Nvl(rsTmp("医嘱内容"))
                    Record(mCol.病人ID).Value = Nvl(rsTmp("病人ID"))
                    Record(mCol.管码).Value = Nvl(rsTmp("管码"))
                    Record(mCol.床号).Value = Nvl(rsTmp("当前床号"))
                    Record(mCol.标识号).Value = Nvl(rsTmp("标识号"))
                    Record(mCol.试管颜色).Value = Nvl(rsTmp("试管颜色"), -1)
                    Record(mCol.开嘱医生).Value = Nvl(rsTmp("开嘱医生"))
                    Record(mCol.开嘱时间).Value = Nvl(rsTmp("开嘱时间"))
                    Record(mCol.采样人).Value = Nvl(rsTmp("采样人"))
                    Record(mCol.采样时间).Value = Nvl(rsTmp("采样时间"))
                    Record(mCol.采血量).Value = Nvl(rsTmp("采血量"))
                    Record(mCol.试管名称).Value = Nvl(rsTmp("试管名称"))
                    Record(mCol.紧急).Value = Nvl(rsTmp("紧急"))
                    Record(mCol.病人来源).Value = Nvl(rsTmp("病人来源"))
                    Record(mCol.婴儿).Value = Nvl(rsTmp("婴儿"), 0)
                    Record(mCol.送出时间).Value = Nvl(rsTmp("标本送出时间"))
                    Record(mCol.开嘱科室).Value = Nvl(rsTmp("开嘱科室"))
                    Record(mCol.诊疗项目ID).Value = Nvl(rsTmp("诊疗项目ID"))
                    Record(mCol.别名).Value = IIf(Trim(Nvl(rsTmp("别名"))) = "", Nvl(rsTmp("医嘱内容")), Nvl(rsTmp("别名")))
                    Record(mCol.条码打印).Value = IIf(Val(Nvl(rsTmp("条码打印"))) = 0, "未打印", "已打印")
                    Record(mCol.类别).Value = Nvl(rsTmp("类别"))
                    If Nvl(rsTmp("类别")) = "输血" Then blnFL = True    '当存在输血医嘱时,需要分类显示,用于提示技师这是输血医嘱,不要当做检验医嘱处理
                    For i = 0 To .Columns.Count
                        Record(i).ForeColor = Nvl(rsTmp("试管颜色"), -1)
                    Next
                    
                    If Nvl(rsTmp("重采")) = 1 Then
                        For i = 0 To .Columns.Count
                            Record(i).Bold = True
                        Next
                    End If
                Else
                    Record(mCol.合并医嘱).Value = Record(mCol.合并医嘱).Value & "," & Nvl(rsTmp("医嘱ID")) & "," & Nvl(rsTmp("相关ID"))
                    
                    str医嘱内容 = Nvl(rsTmp("医嘱内容"))
                    If InStr(";" & Record(mCol.医嘱内容).Value & ";", str医嘱内容) <= 0 Then
                        Record(mCol.医嘱内容).Value = Record(mCol.医嘱内容).Value & ";" & Nvl(rsTmp("医嘱内容"))
                    End If
                        
                    str医嘱内容 = IIf(Trim(Nvl(rsTmp("别名"))) = "", Nvl(rsTmp("医嘱内容")), Nvl(rsTmp("别名")))
                    If InStr(";" & Record(mCol.别名).Value & ";", str医嘱内容) <= 0 Then
                        Record(mCol.别名).Value = Record(mCol.别名).Value & ";" & str医嘱内容
                    End If
                End If
                
                strUnionID = Nvl(rsTmp("相关ID"))
            End If
            
            rsTmp.MoveNext
        Loop
        .Columns(mCol.选择).TreeColumn = False
        
        If blnFL = True Then
            Call .GroupsOrder.Add(.Columns.Column(mCol.类别))
        End If
        .Populate
        If strBarCode <> "" And .Records.Count > 0 Then
            Me.TxtBarCode.Tag = .Records.Count
        End If
    End With
    Call chkPrint_Click
    Me.MousePointer = vbDefault
    zlCommFun.StopFlash
    
    Exit Sub
errH:
    Me.MousePointer = vbDefault
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function GetBooldReadDAtaSql(Optional ByVal strBarCode As String) As String
    Dim strSQL
    Dim strSQLbak As String
    On Error GoTo errH
    
    strSQL = "Select /*+ rule */ a.类别,A.医嘱id, A.相关id, A.医嘱内容, A.发送时间, A.执行科室, A.姓名, A.性别, A.年龄, A.样本条码, A.标本, A.病人id, A.采集方式, A.管码, A.当前床号, A.标识号," & vbNewLine & _
            "       A.开嘱医生, A.开嘱时间, A.采样人, A.采样时间, A.紧急, A.病人来源, A.婴儿, A.别名, A.组合项目, A.条码打印, A.记录状态, A.记录性质, A.重采, A.标本送出时间," & vbNewLine & _
            "       B.颜色 As 试管颜色, B.采血量, B.名称 As 试管名称,a.诊疗项目ID,a.检验执行科室ID,a.开嘱科室 from  " & vbCrLf & _
             " (Select  distinct decode(d.类别,'K','输血','检验') 类别,B.Id as 医嘱ID,B.相关Id,decode(d.类别,'K',d.名称,b.医嘱内容) as 医嘱内容,m.发送时间,H.名称 as 执行科室,I.姓名,I.性别,I.年龄,m.样本条码,b.标本部位 as 标本, " & vbCrLf & _
             " I.病人ID,decode(d.类别,'K',b.医嘱内容 ,d.名称) As 采集方式,E.试管编码 as  管码,b.诊疗项目ID,b.执行科室ID as 检验执行科室ID, " & vbCrLf & _
             "decode(i.当前床号,null,decode(l.出院病床,null,l.入院病床,l.出院病床),i.当前床号) as 当前床号, " & vbCrLf & _
             "Decode(B.病人来源,1,I.门诊号,2,i.住院号,4,i.门诊号) as 标识号, " & vbCrLf & _
             " b.开嘱医生,b.开嘱时间,m.采样人,m.采样时间,decode(b.紧急标志,1,'紧急','') as 紧急, " & _
             " b.病人来源,b.婴儿,n.名称 as 别名,E.组合项目,c.条码打印,nvl(P.记录状态,0) as 记录状态, " & vbCrLf & _
             " nvl(P.记录性质,0) as 记录性质,nvl(c.重采标本,0) as 重采,c.标本送出时间,Q.名称 as 开嘱科室,decode(d.类别, 'K', M.执行状态,C.执行状态) 执行状态 " & vbCrLf & _
             " From 病人医嘱记录 A, 病人医嘱记录 B,病人医嘱发送 C,诊疗项目目录 D, 诊疗项目目录 E, " & vbCrLf & _
             "      部门表 H,病人信息 I,病案主页 L,病人医嘱发送 M, " & vbCrLf & _
             " (select 诊疗项目ID,名称 from 诊疗项目别名 where 性质 = 9 and 码类 = 1 ) N,住院费用记录 P,部门表 Q " & vbCrLf & _
             " Where A.ID = B.相关id And A.诊疗项目id = D.ID And B.诊疗项目id = E.ID And (e.类别 = 'E' Or e.类别 = 'C') And " & vbCrLf & _
             " B.执行科室id = H.ID And d.类别 = 'K'  And  e.操作类型 = '9' And " & vbCrLf & _
             " A.病人Id = I.病人ID And a.病人id = l.病人ID(+) and a.主页id = l.主页id(+) and  c.执行部门id + 0 = [1] and M.发送时间 Between [2] And [3] " & vbCrLf & _
             " And A.id = M.医嘱ID And b.id = c.医嘱ID and E.id = N.诊疗项目ID(+)  " & vbCrLf & _
             " And C.医嘱ID = P.医嘱序号(+) and C.记录性质 = Mod(P.记录性质(+),10) and b.开嘱科室ID = Q.ID " & vbCrLf
    
    '标本
    If cboSample.ItemData(cboSample.ListIndex) <> 0 Then
        strSQL = strSQL & " And decode(d.类别, 'K',[4],B.标本部位)= [4] "
    End If

    If Me.cboCapture.ItemData(cboCapture.ListIndex) <> 0 Then
        strSQL = strSQL & " And decode(d.类别, 'K',E.名称,D.名称) = [5] "
    End If
    
    '病人单位
    If Trim(Me.txtUnit.Text) <> "" Then
        strSQL = strSQL & " and I.工作单位 like [6] "
    End If
    
    '条码
    If strBarCode <> "" Then
        strSQL = strSQL & " And m.样本条码 = [7]"
    End If
    
    If Me.cboState = "未绑定" Then
        strSQL = strSQL & " And c.样本条码 is null) a , 采血管类型 b "
    ElseIf Me.cboState = "已绑定" Then
        strSQL = strSQL & " And c.样本条码 is not null and c.采样人 is null) a , 采血管类型 B "
    ElseIf Me.cboState = "已采样" Then
        strSQL = strSQL & " and c.样本条码 is not null and c.采样人 is not null and c.标本送出时间 is null) a,采血管类型 B  "
    ElseIf Me.cboState = "已送检" Then
        strSQL = strSQL & " and c.样本条码 is not null and c.采样人 is not null and c.标本送出时间 is not null) a,采血管类型 B  "
    ElseIf Me.cboState = "已执行" Then
        strSQL = strSQL & ") a,采血管类型 B "
    End If
    strSQL = strSQL & " Where a.管码 = b.编码 "
    If Me.cboState = "已执行" Then
        strSQL = strSQL & " and a.执行状态 IN (1,3)"
    Else
        strSQL = strSQL & " and a.执行状态 IN (0,2)"
    End If
    strSQLbak = strSQL
    
    strSQLbak = Replace$(strSQLbak, "住院费用记录", "门诊费用记录")
    strSQL = strSQL & " union all " & strSQLbak
    GetBooldReadDAtaSql = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub GetInitData()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngDeptID As Long                       '科室ID
    Dim lngSampleID As Long                     '标本ID
    Dim lngCaptureID As Long                    '采集方式
    Dim intUnionState As Integer                '执行状态
    Dim intSpaceDate As Integer                 '间隔时间
    Dim strNowDate As Date                      '取当前服务器时间
    
    intSpaceDate = DateDiff("d", Me.DTPEnd.Value, Me.DtpBegin.Value)
   
    lngDeptID = zlDatabase.GetPara("frmLabBarCodeBatPrint_科室名称Id", 100, 1208, 0)
    lngSampleID = zlDatabase.GetPara("frmLabBarCodeBatPrint_标本ID", 100, 1208, 0)
    lngCaptureID = zlDatabase.GetPara("frmLabBarCodeBatPrint_采集方法", 100, 1208, 0)
    intUnionState = zlDatabase.GetPara("frmLabBarCodeBatPrint_执行状态", 100, 1208, 0)
'    Me.chkComplete.Value = zlDatabase.GetPara("frmLabBarCodeBatPrint_是否标记为完成", 100, 1208, 0)
    intSpaceDate = zlDatabase.GetPara("frmLabBarCodeBatPrint_间隔时间", 100, 1208, 2)
    
    On Error GoTo errH
    
    '===读入科室
    strSQL = _
            " Select Distinct A.ID,A.编码 || '-' || A.名称 as 名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(1,2,3,4) And B.工作性质 IN('检验','护理')"
            
            
    If InStr(1, mstrPrivs, "所有科室") <= 0 Then
        strSQL = strSQL & " And C.人员id = [1] "
    End If
            
    strSQL = strSQL & " Order by A.编码 || '-' || A.名称"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    cboDept.Clear
    Do Until rsTmp.EOF
        cboDept.AddItem rsTmp("名称")
        cboDept.ItemData(cboDept.NewIndex) = rsTmp("ID")
        If rsTmp("id") = IIf(lngDeptID = 0, UserInfo.部门ID, lngDeptID) Then
            cboDept.ListIndex = cboDept.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cboDept.Text = "" And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    
    '===读入采集方式(加入血库输血采集)
    strSQL = "select ID,名称 from 诊疗项目目录 where (撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL) And 操作类型 in ('6','9') And 类别='E'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    cboCapture.Clear
    cboCapture.AddItem "所有采集方式"
    cboCapture.ItemData(cboCapture.NewIndex) = 0
    Do Until rsTmp.EOF
        cboCapture.AddItem rsTmp("名称")
        cboCapture.ItemData(cboCapture.NewIndex) = rsTmp("ID")
        If lngCaptureID = rsTmp("id") Then
            cboCapture.ListIndex = cboCapture.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cboCapture.Text = "" And cboCapture.ListCount > 0 Then cboCapture.ListIndex = 0
    
    '===读入检验标本
    strSQL = "select 编码,名称 from 诊疗检验标本"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    cboSample.Clear
    cboSample.AddItem "所有标本"
    Do Until rsTmp.EOF
        cboSample.AddItem rsTmp("编码") & "-" & rsTmp("名称")
        cboSample.ItemData(cboSample.NewIndex) = rsTmp("编码")
        If rsTmp("编码") = lngSampleID Then
            cboSample.ListIndex = cboSample.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cboSample.Text = "" And cboSample.ListCount > 0 Then cboSample.ListIndex = 0
    
    '===执行状态
    cboState.Clear
    cboState.AddItem "未绑定"
    cboState.ItemData(cboState.NewIndex) = 0
    cboState.AddItem "已绑定"
    cboState.ItemData(cboState.NewIndex) = 1
    cboState.AddItem "已采样"
    cboState.ItemData(cboState.NewIndex) = 2
    cboState.AddItem "已送检"
    cboState.ItemData(cboState.NewIndex) = 3
    cboState.AddItem "已执行"
    cboState.ItemData(cboState.NewIndex) = 4
    cboState.ListIndex = intUnionState
    
    
    '===时间段
    strNowDate = zlDatabase.Currentdate
    Me.DTPEnd = Format(strNowDate, "yyyy-mm-dd hh:mm")
    Me.DtpBegin.Value = Format(strNowDate - intSpaceDate, "yyyy-mm-dd 00:00")
    
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With fraFilter
        .Width = Me.ScaleWidth - 7 * Screen.TwipsPerPixelX
    End With
    
    With cboCapture
        .Width = fraFilter.Width - .Left - 15 * Screen.TwipsPerPixelX
    End With
    
    With cboState
        .Width = fraFilter.Width - .Left - 15 * Screen.TwipsPerPixelX
    End With
    
    With txtUnit
        .Width = fraFilter.Width - .Left - 15 * Screen.TwipsPerPixelX
    End With
    
    With cmdFind
        .Left = fraFilter.Width - .Width - 15 * Screen.TwipsPerPixelX
    End With
    
    With Frame1
        .Width = fraFilter.Width - .Left - 15 * Screen.TwipsPerPixelX
    End With
    
    If cboState.Text = "已采样" Or cboState.Text = "已送检" Or cboState.Text = "已绑定" Then
        Me.fraBarCode.Visible = True
        Me.fraBarCode.Width = Me.fraFilter.Width
        Me.TxtBarCode.Width = Me.fraBarCode.Width - Me.TxtBarCode.Left - 40
        Me.RptItem.Top = Me.fraBarCode.Top + Me.fraBarCode.Height + 20
    Else
        Me.fraBarCode.Visible = False
        Me.fraBarCode.Width = Me.fraFilter.Width
        Me.TxtBarCode.Width = Me.fraBarCode.Width - Me.TxtBarCode.Left - 40
        Me.RptItem.Top = Me.fraFilter.Top + Me.fraFilter.Height + 20
    End If
    

    
    With RptItem
        .Width = Me.fraFilter.Width
        .Height = Me.ScaleHeight - .Top - Me.cmdCancel.Height - 20 * Screen.TwipsPerPixelY
    End With
    
    With cmdCancel
        .Top = Me.ScaleHeight - .Height - 10 * Screen.TwipsPerPixelY
        .Left = Me.ScaleWidth - .Width - 20 * Screen.TwipsPerPixelX
    End With
    
    With cmdPrint
        .Top = Me.cmdCancel.Top
        .Left = Me.cmdCancel.Left - .Width - 20 * Screen.TwipsPerPixelX
    End With
    
    With cmdReturBill
        .Top = Me.cmdCancel.Top
        .Left = Me.cmdPrint.Left - .Width - 20 * Screen.TwipsPerPixelX
    End With
    
    With cmdPrintSend
        .Top = Me.cmdCancel.Top
        .Left = Me.cmdReturBill.Left - .Width - 20 * Screen.TwipsPerPixelX
    End With
    
    With cmdPrintSet
        .Top = Me.cmdCancel.Top
        .Left = 300
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    Call SaveWinState(Me, App.ProductName)
    
    i = DateDiff("d", Me.DtpBegin.Value, Me.DTPEnd.Value)
    
    zlDatabase.SetPara "frmLabBarCodeBatPrint_科室名称Id", cboDept.ItemData(cboDept.ListIndex), 100, 1208
    zlDatabase.SetPara "frmLabBarCodeBatPrint_标本ID", cboSample.ItemData(cboSample.ListIndex), 100, 1208
    zlDatabase.SetPara "frmLabBarCodeBatPrint_采集方法", cboCapture.ItemData(cboCapture.ListIndex), 100, 1208
    zlDatabase.SetPara "frmLabBarCodeBatPrint_执行状态", cboState.ItemData(cboState.ListIndex), 100, 1208
'    zlDatabase.SetPara "frmLabBarCodeBatPrint_是否标记为完成", chkComplete.Value, 100, 1208
    zlDatabase.SetPara "frmLabBarCodeBatPrint_间隔时间", i, 100, 1208
End Sub

Private Sub RptItem_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim RecordC As ReportRecord
    For Each RecordC In Me.RptItem.Records
        If RecordC(mCol.条码).Value = Item.Record(mCol.条码).Value And Item.Record(mCol.条码).Value <> "" Then
            RecordC(mCol.选择).Checked = Item.Checked
        End If
    Next
End Sub

Private Sub RptItem_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim blSelect As Boolean
    Dim RepRow As ReportRow
    Dim hitColumn As ReportColumn
    With Me.RptItem
        Set hitColumn = .HitTest(X, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "Check" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                blSelect = Not .Records(0).Item(mCol.选择).Checked
                SelectOrCancelReprotCheck .Records, mCol.选择, blSelect
            End If
            .Populate
        End If
    End With
End Sub
Private Sub SelectOrCancelReprotCheck(RepObj As ReportRecords, intFiledCol As Integer, blSelect As Boolean)
    Dim Record As ReportRecord
    For Each Record In RepObj
        Record.Visible = True
        Record.Item(intFiledCol).Checked = blSelect
        If Record.Item(mCol.条码打印).Value = "已打印" And chkPrint.Value = 1 Then
            Record.Visible = False
        End If
    Next
    Me.RptItem.Populate
End Sub

Private Sub RptItem_SelectionChanged()
    With Me.RptItem
        If Not .FocusedRow Is Nothing And .FocusedRow.GroupRow = False Then
            .PaintManager.HighlightBackColor = Val(.FocusedRow.Record(mCol.试管颜色).Value)
            .Populate
        End If
    End With
End Sub

Public Sub ShowME(Objfrm As Object, strPrivs As String, intBarCodeFormat As Integer, intExecDept As Integer, blnNowConsumption As Boolean)
    mstrPrivs = strPrivs
    mintBarCodeFormat = intBarCodeFormat
    mintExecDept = intExecDept
    mblnNowConsumption = blnNowConsumption
    Me.Show , Objfrm
End Sub

Private Function CheckPlugIn(ByVal lngSys As Long, ByVal lngModual As Long, ByVal rsMoneyNow As ADODB.Recordset) As Boolean
'    rsNumber.Fields.Append "类别", adVarChar, 20
'    rsNumber.Fields.Append "管码", adVarChar, 18
'    rsNumber.Fields.Append "相关ID", adBigInt
'    rsNumber.Fields.Append "样本条码", adVarChar, 18
'    rsNumber.Fields.Append "执行科室ID", adVarChar, 18
'    rsNumber.Fields.Append "诊疗项目ID", adVarChar, 18
'    rsNumber.Fields.Append "婴儿", adBigInt
'    rsNumber.Fields.Append "紧急标志", adBigInt
'    rsNumber.Fields.Append "标本", adVarChar, 30
'    rsNumber.Fields.Append "医嘱内容", adVarChar, 500
'    rsNumber.Fields.Append "采集方式", adVarChar, 100
'    rsNumber.Fields.Append "开嘱医生", adVarChar, 50
'    rsNumber.Fields.Append "开嘱时间", adDate
'    rsNumber.Fields.Append "采样人", adVarChar, 50
'    rsNumber.Fields.Append "采样时间", adDate
'    rsNumber.Fields.Append "采血量", adVarChar, 20
'    rsNumber.Fields.Append "试管名称", adVarChar, 50
'    rsNumber.Fields.Append "病人来源", adInteger
'    rsNumber.Fields.Append "医嘱ID串", adVarChar, 500
'    rsNumber.Fields.Append "执行科室", adVarChar, 50
'    rsNumber.Fields.Append "婴儿姓名", adVarChar, 50
'    rsNumber.Fields.Append "婴儿性别", adVarChar, 50
'    rsNumber.Fields.Append "申请科室", adVarChar, 50
    
    Dim blnTmp As Boolean
        On Error Resume Next
        CheckPlugIn = True
        If Not mobjZLIHISPlugIn Is Nothing Then
            blnTmp = mobjZLIHISPlugIn.LisPrintCodeBefore(lngSys, lngModual, rsMoneyNow)
            Call zlPlugInErrH(Err, "LisPrintCodeBefore")
            If Err.Number <> 0 Then
                '接口出错了,继续打印
                blnTmp = True
            End If
        Else
            blnTmp = True
        End If
        CheckPlugIn = blnTmp
    Err.Clear: On Error GoTo 0

End Function

Private Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub BarCodeMake(intMode As Integer, bln条码打印 As Boolean)
    '功能                           写入条码.当没有条码时使用医嘱ID生成条码
    '                               intMode 1=生成 2=完成 3=生成完成 4=送检 5=取消送检 6=只打印
    '                               bln条码打印  true = 打印
    '                               生成条码时，按一个病人同样的标本为单位进行写入
    
    Dim lngPatientID As Long                    '病人ID
    Dim lngloop As Long                         '循环变量
    Dim intLoop As Long                         '循环变量
    Dim strCuvetteNumber As String              '管码
    Dim strUnion As String                      '医嘱ID,发送号,条码 使用"|"分隔
    Dim strNewBarCode As String                 '生成的条码
    Dim varAdvice As Variant                    '合并的医嘱ID
    Dim varItem As Variant                      '分解字串用的临进变量
    Dim strSQL As String                        'SQL语句
    Dim strBarCodeUnion As String               '条码字串
    Dim varBarCodeUnion As Variant              '条码字串分解
    Dim i As Integer                            '循环变量
    Dim intBaby As Integer                      '婴儿 >0 表示婴儿数量
    Dim strSample As String                     '标本
    Dim strAdviceContent As String              '医嘱内容
    Dim lngConnectID As Long                    '相关ID
    Dim varFilter As Variant                    '过滤相同的医嘱内容
    Dim strDept As String                       '执行科室
    Dim str紧急 As String                       '紧急
    Dim rsNumber As ADODB.Recordset
    Dim astrSQL() As String
    Dim blnRollBak As Boolean
    Dim blnPrint As Boolean                     '是否外挂打印
    
    ReDim astrSQL(0)
    
    If RptItem.Records.Count = 0 Then Exit Sub

    '关闭其他按钮
    Me.cmdFind.Enabled = False
    Me.cmdPrint.Enabled = False
    Me.cmdCancel.Enabled = False
    Me.cmdReturBill.Enabled = False
    On Error GoTo errH
    
    BlCancel = False
    
    zlCommFun.ShowFlash "正在打印条码,请稍候...", Me
    Me.MousePointer = vbHourglass
    
    '批量打印条码
    
    InitRecordSet rsNumber
    
    With Me.RptItem
        For lngloop = 0 To .Records.Count - 1
        
            If .Records(lngloop).Item(mCol.选择).Checked = True And .Records(lngloop).Visible = True Then
            
                If BlCancel = True Then Exit Sub                                    '按下"ESC"时退出
                
                Select Case intMode
                    Case 1, 3
                        MakeBarCode rsNumber, .Records(lngloop), 1, mintExecDept
                    Case 2
                        MakeBarCode rsNumber, .Records(lngloop), 3, mintExecDept
                    Case 4, 5, 6
                        MakeBarCode rsNumber, .Records(lngloop), 4, mintExecDept
                End Select
                
            End If
        Next
       
       
    End With
    
    On Error GoTo errH
    
    If rsNumber.RecordCount = 0 Then Exit Sub
    rsNumber.MoveFirst
    Select Case intMode
        Case 1, 3                                   '生成条码和完成
            Do Until rsNumber.EOF
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_条码生成('" & rsNumber("医嘱ID串") & "','" & rsNumber("样本条码") & "')"
                If intMode = 3 Then
                    '执行完成
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0," & IIf(rsNumber("类别") & "" = "输血", 1, 0) & ")"
                End If
                rsNumber.MoveNext
            Loop
        Case 2                                      '完成采集
            Do Until rsNumber.EOF
                '执行完成
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_检验预置条码_采集完成('" & rsNumber("医嘱ID串") & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0," & IIf(rsNumber("类别") & "" = "输血", 1, 0) & ")"
                rsNumber.MoveNext
            Loop
        Case 4                                      '送检
            Do Until rsNumber.EOF
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_Lis预置条码_标本送出('" & rsNumber("医嘱ID串") & "')"
                rsNumber.MoveNext
            Loop
        Case 5                                      '取消送检
            Do Until rsNumber.EOF
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "Zl_Lis预置条码_标本送出('" & rsNumber("医嘱ID串") & "',1)"
                rsNumber.MoveNext
            Loop
    End Select
    
    gcnOracle.BeginTrans
    blnRollBak = True
    
    
    For intLoop = 1 To UBound(astrSQL)
        If astrSQL(intLoop) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        End If
    Next
    gcnOracle.CommitTrans
    blnRollBak = False
    
    
    If intMode = 1 Or intMode = 3 Or intMode = 2 Then
        Call WriterBarCodeToLIS(rsNumber, 3)
    End If
    '打印条码
    If bln条码打印 = True Then
        blnPrint = CheckPlugIn(glngSys, glngModul, rsNumber)
        If blnPrint = True Then
            rsNumber.MoveFirst
            Do Until rsNumber.EOF
                '成生条码到PIC
                
                If mintBarCodeFormat = 1 Then
                    Bar39 Me.picBarCodePrint, 3, Nvl(rsNumber("样本条码")), False, True
                Else
                    Bar128 Me.picBarCodePrint, 3, Nvl(rsNumber("样本条码")), True
                End If
                SavePicture Me.picBarCodePrint.Image, App.path & "\BarCode.Bmp"
                '开始打印
                Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_1", Me, "样本条码=" & Nvl(rsNumber("样本条码")), _
                "项目=" & Replace(Nvl(rsNumber("医嘱内容")), ",", " "), _
                "病人姓名 = " & IIf(Nvl(rsNumber("姓名")) <> "", Nvl(rsNumber("姓名")) & IIf(Nvl(rsNumber("婴儿"), 0) = 0, "", "(婴儿" & Nvl(rsNumber("婴儿")) & ")"), "无"), _
                "性别 = " & IIf(Nvl(rsNumber("性别")) <> "", Nvl(rsNumber("性别")), "无"), _
                "年龄 = " & IIf(Nvl(rsNumber("年龄")) <> "", Nvl(rsNumber("年龄")), "无"), _
                "床号 = " & IIf(Nvl(rsNumber("床号")) <> "", Nvl(rsNumber("床号")), "无"), _
                "标识号 = " & IIf(Nvl(rsNumber("标识号")) <> "", Nvl(rsNumber("标识号")), "无"), _
                "所在科室 = " & IIf(Nvl(rsNumber("开嘱科室")) <> "", Nvl(rsNumber("开嘱科室")), "无"), _
                "采集方式 = " & IIf(Nvl(rsNumber("采集方式")) <> "", Nvl(rsNumber("采集方式")), "无"), _
                "标本 = " & IIf(Nvl(rsNumber("标本")) <> "", Nvl(rsNumber("标本")), "无"), _
                "执行科室 = " & IIf(Nvl(rsNumber("执行科室")) <> "", Nvl(rsNumber("执行科室")), "无"), _
                "开嘱医生 = " & IIf(Nvl(rsNumber("开嘱医生")) <> "", Nvl(rsNumber("开嘱医生")), "无"), _
                "开嘱时间 = " & IIf(Nvl(rsNumber("开嘱时间")) <> "", Nvl(rsNumber("开嘱时间")), "无"), _
                "采样人 = " & IIf(Nvl(rsNumber("采样人")) <> "", Nvl(rsNumber("采样人")), "无"), _
                "采样时间 = " & IIf(Nvl(rsNumber("采样时间")) <> "", Nvl(rsNumber("采样时间")), "无"), _
                "管码 = " & IIf(Nvl(rsNumber("管码")) <> "", Nvl(rsNumber("管码")), "无"), _
                "采血量 = " & IIf(Nvl(rsNumber("采血量")) <> "", Nvl(rsNumber("采血量")), "无"), _
                "试管名称 = " & IIf(Nvl(rsNumber("试管名称")) <> "", Nvl(rsNumber("试管名称")), "无"), _
                "紧急 = " & IIf(Nvl(rsNumber("紧急标志")) <> "", Nvl(rsNumber("紧急标志")), "无"), _
                "病人来源 = " & IIf(Nvl(rsNumber("病人来源")) <> "", Nvl(rsNumber("病人来源")), "无"), _
                "条码图像1=" & App.path & "\BarCode.Bmp", 2)
                '删除条码图像
                Kill App.path & "\BarCode.Bmp"
                strSQL = "Zl_Lis预置条码_条码打印('" & Replace(rsNumber("医嘱ID串"), ",,", ",") & "')"
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
                rsNumber.MoveNext
            Loop
        End If
    End If
    
    
    
    
    
    zlCommFun.StopFlash
    Me.MousePointer = vbDefault
    '恢复正常
    Me.cmdFind.Enabled = True
    Me.cmdPrint.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.cmdReturBill.Enabled = True
    ReadData
    Exit Sub
errH:
    If blnRollBak Then
        gcnOracle.RollbackTrans
        blnRollBak = False
    End If
    Me.cmdFind.Enabled = True
    Me.cmdPrint.Enabled = True
    Me.MousePointer = vbDefault
    zlCommFun.StopFlash
    
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog


End Sub
Private Sub SampleSend()
    '记录送出标本时间
    Dim intLoop As Integer
    Dim strIDs As String
    Dim strAllIDs As String
    Dim intRow As Integer
    Dim blnRollBak As Boolean
    Dim astrSQL() As String
    Dim astrReprot() As String
    Dim strTmp As String
    ReDim astrSQL(0)
    Dim strAdvils As String
    Dim strName As String
    Dim blnPrint As Boolean
    Dim strSQL As String
    Dim rsSampleCode As ADODB.Recordset
    
    If Me.RptItem.Rows.Count = 0 Then Exit Sub
    
    If Me.cboState.Text <> "已送检" Then
        If frmLabSamplingSendInfo.ShowME(Me, strName, blnPrint) = False Then
            Exit Sub
        End If
    End If
    
    '生成发送批号
    strSQL = "select 病人医嘱发送_标本发送批号.NEXTVAL  from dual"
    Set rsSampleCode = zlDatabase.OpenSQLRecord(strSQL, "标本发送批号", "")
    
    With Me.RptItem
    
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).GroupRow = False Then
                If .Rows(intLoop).Record(mCol.选择).Checked = True Then
                    strIDs = strIDs & "," & .Rows(intLoop).Record(mCol.相关ID).Value
                    If Len(strTmp) >= 3800 Then
                        strAllIDs = strAllIDs & strTmp & ";"
                        strTmp = ""
                    Else
                        If strAdvils <> "" Then
                            strTmp = strTmp & strAdvils
                            strAdvils = ""
                        End If
                    End If
                    If intRow = 5 Then
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        strIDs = Replace(Replace(strIDs, "|", ","), ";", ",")
                                           
                        astrSQL(UBound(astrSQL)) = "Zl_Lis预置条码_标本送出('" & Mid(strIDs, 2) & "'" & IIf(Me.cboState.Text = "已送检", ",1", ",0") & _
                                    ",'" & strName & "','" & rsSampleCode(0) & "')"
                        strAdvils = strIDs
                        strIDs = ""
                        intRow = 0
                    End If
                    intRow = intRow + 1
                End If
            End If
        Next
        If strAdvils <> "" Then strTmp = strTmp & strAdvils
        If strIDs <> "" Then strTmp = strTmp & strIDs
        strAllIDs = strAllIDs & strTmp & ";"
        strIDs = Replace(Replace(strIDs, "|", ","), ";", ",")
    End With
    On Error GoTo errH
        
    '保存送出时间
    If strIDs <> "" Then
        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
        astrSQL(UBound(astrSQL)) = "Zl_Lis预置条码_标本送出('" & Mid(strIDs, 2) & "'" & IIf(Me.cboState.Text = "已送检", ",1", ",0") & _
                                ",'" & strName & "','" & rsSampleCode(0) & "')"
    End If
    
    gcnOracle.BeginTrans
    blnRollBak = True
    
    For intLoop = 1 To UBound(astrSQL)
        If astrSQL(intLoop) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        End If
    Next
    gcnOracle.CommitTrans
    blnRollBak = False
    
    If strAllIDs <> "" Then
        strAllIDs = Mid(strAllIDs, 2)
    Else
        strAllIDs = 0
    End If
    astrReprot = Split(strAllIDs, ";")
    For intLoop = 0 To UBound(astrReprot)
        If astrReprot(intLoop) <> "" Then
            '写入送检时间到检验申请单中
            Call WriterSampleSendDateToLIS(astrReprot(intLoop), IIf(Me.cboState.Text = "已送检", "1", "0"), strName)
        End If
    Next
    
    If Me.cboState.Text <> "已送检" Then
'        If MsgBox("是否打印送出清单?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        If blnPrint = True Then
            astrReprot = Split(strAllIDs, ";")
            For intLoop = 0 To UBound(astrReprot)
                If astrReprot(intLoop) <> "" Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_3", Me, "医嘱字串=" & astrReprot(intLoop), 2)
                End If
            Next
        End If
    End If
    ReadData
    Exit Sub
errH:
    If blnRollBak Then
        gcnOracle.RollbackTrans
        blnRollBak = False
    End If
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub TxtBarCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.TxtBarCode.Tag = "" Then
            Me.RptItem.Records.DeleteAll
            Me.RptItem.Populate
        End If
        Call ReadData(Me.TxtBarCode)
        Me.TxtBarCode.Text = ""
        Me.TxtBarCode.SetFocus
    End If
End Sub

Private Sub InitRecordSet(rsNumber As ADODB.Recordset)
    '初始化记录集
    
    '计录试管编码
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "类别", adVarChar, 20
    rsNumber.Fields.Append "管码", adVarChar, 18
    rsNumber.Fields.Append "相关ID", adBigInt
    rsNumber.Fields.Append "样本条码", adVarChar, 18
    rsNumber.Fields.Append "执行科室ID", adVarChar, 18
    rsNumber.Fields.Append "诊疗项目ID", adVarChar, 18
    rsNumber.Fields.Append "婴儿", adBigInt
    rsNumber.Fields.Append "紧急标志", adBigInt
    rsNumber.Fields.Append "标本", adVarChar, 30
    rsNumber.Fields.Append "医嘱内容", adVarChar, 500
    rsNumber.Fields.Append "采集方式", adVarChar, 100
    rsNumber.Fields.Append "开嘱医生", adVarChar, 50
    rsNumber.Fields.Append "开嘱时间", adDate
    rsNumber.Fields.Append "采样人", adVarChar, 50
    rsNumber.Fields.Append "采样时间", adDate
    rsNumber.Fields.Append "采血量", adVarChar, 20
    rsNumber.Fields.Append "试管名称", adVarChar, 50
    rsNumber.Fields.Append "病人来源", adInteger
    rsNumber.Fields.Append "医嘱ID串", adVarChar, 500
    rsNumber.Fields.Append "执行科室", adVarChar, 50
    rsNumber.Fields.Append "病人ID", adVarChar, 18
    rsNumber.Fields.Append "姓名", adVarChar, 50
    rsNumber.Fields.Append "性别", adVarChar, 10
    rsNumber.Fields.Append "年龄", adVarChar, 50
    rsNumber.Fields.Append "床号", adVarChar, 50
    rsNumber.Fields.Append "标识号", adVarChar, 50
    rsNumber.Fields.Append "开嘱科室", adVarChar, 50
    
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
End Sub

Public Function MakeBarCode(rsNumber As ADODB.Recordset, RowRecord As ReportRecord, intMode As Integer, Optional intExecDept As Integer, Optional strBarCode As String) As Boolean
    '功能                   生成条码并记录方便后面保存到数据或打印
    '参数                   用于记录的记录庥
    '                       RowRecord行数据
    '                       '执行科室是否要区分
    '                       Mode =0 绑定条码 =1 生成条码 =2 清除条码 = 3 完成采集 = 4 打印条码或回执单
    '                       strBarCode <> ""时表示使用绑定条码
    Dim strFilter As String
    Dim blnNew As Boolean
    Dim str医嘱内容 As String
    
    blnNew = False
    Select Case intMode
        Case 0                              '绑定
            If rsNumber.RecordCount = 0 Then blnNew = True
        Case 1                              '生成
            strFilter = "病人ID=" & RowRecord.Item(mCol.病人ID).Value & " And 诊疗项目ID=" & Val(RowRecord.Item(mCol.诊疗项目ID).Value)
            rsNumber.filter = strFilter
            If rsNumber.EOF = False Then
                '当诊疗项目相同时新增一个条码
                blnNew = True
            Else
                strFilter = "病人ID=" & RowRecord.Item(mCol.病人ID).Value & _
                      " And 管码='" & RowRecord.Item(mCol.管码).Value & _
                      "' And 婴儿=" & RowRecord.Item(mCol.婴儿).Value & _
                      " And 紧急标志=" & IIf(RowRecord.Item(mCol.紧急).Value = "紧急", 1, 0) & _
                      " And 标本='" & RowRecord.Item(mCol.标本).Value & "'"
                If intExecDept = 1 Then strFilter = strFilter & " And 执行科室id=" & RowRecord.Item(mCol.检验执行科室ID).Value
                rsNumber.filter = strFilter
                If rsNumber.EOF = True Then
                    '生成新条码
                    blnNew = True
                End If
            End If
        Case 2                              '取消条码
            If rsNumber.RecordCount = 0 Then blnNew = True
        
        Case 3, 4                           '用于条码打印
            strFilter = "样本条码='" & RowRecord.Item(mCol.条码).Value & "'"
            rsNumber.filter = strFilter
            If rsNumber.EOF = True Then
                blnNew = True
            End If
    End Select
    If blnNew = True Then
        rsNumber.AddNew
        rsNumber!类别 = RowRecord.Item(mCol.类别).Value
        '绑定和生成条码
        If strBarCode <> "" Then
            rsNumber!样本条码 = strBarCode
        Else
            If intMode = 3 Or intMode = 4 Then
                rsNumber!样本条码 = RowRecord.Item(mCol.条码).Value
            Else
                rsNumber!样本条码 = zlDatabase.GetNextNo(125, Split(RowRecord.Item(mCol.医嘱id).Value, ",")(0))
            End If
        End If
        rsNumber!采集方式 = RowRecord.Item(mCol.采集方式).Value
        rsNumber!标本 = RowRecord.Item(mCol.标本).Value
        rsNumber!执行科室ID = RowRecord.Item(mCol.检验执行科室ID).Value
        rsNumber!开嘱医生 = RowRecord.Item(mCol.开嘱医生).Value
        rsNumber!开嘱时间 = RowRecord.Item(mCol.开嘱时间).Value
        rsNumber!采样人 = RowRecord.Item(mCol.采样人).Value
        If RowRecord.Item(mCol.采样时间).Value <> "" Then
            rsNumber!采样时间 = RowRecord.Item(mCol.采样时间).Value
        End If
        rsNumber!管码 = RowRecord.Item(mCol.管码).Value
        rsNumber!采血量 = RowRecord.Item(mCol.采血量).Value
        rsNumber!试管名称 = RowRecord.Item(mCol.试管名称).Value
        rsNumber!紧急标志 = IIf(RowRecord.Item(mCol.紧急).Value = "紧急", 1, 0)
        rsNumber!病人来源 = RowRecord.Item(mCol.病人来源).Value
        rsNumber!婴儿 = RowRecord.Item(mCol.婴儿).Value
        rsNumber!执行科室 = RowRecord.Item(mCol.执行科室).Value
        rsNumber!医嘱内容 = RowRecord.Item(mCol.别名).Value
        rsNumber!姓名 = RowRecord.Item(mCol.姓名).Value
        rsNumber!性别 = RowRecord.Item(mCol.性别).Value
        rsNumber!年龄 = RowRecord.Item(mCol.年龄).Value
        rsNumber!床号 = RowRecord.Item(mCol.床号).Value
        rsNumber!标识号 = RowRecord.Item(mCol.标识号).Value
        rsNumber!开嘱科室 = RowRecord.Item(mCol.开嘱科室).Value
        rsNumber!病人ID = RowRecord.Item(mCol.病人ID).Value
        rsNumber!诊疗项目ID = Val(RowRecord.Item(mCol.诊疗项目ID).Value)
        rsNumber!医嘱ID串 = Replace(Replace(RowRecord.Item(mCol.医嘱id).Value & "," & _
                            RowRecord.Item(mCol.相关ID).Value & "," & RowRecord.Item(mCol.合并医嘱).Value, ";", ","), ",,", ",")
        rsNumber.Update
    Else
        If rsNumber.RecordCount > 0 Then
            rsNumber.MoveLast
            str医嘱内容 = IIf(Trim(RowRecord.Item(mCol.别名).Value) = "", RowRecord.Item(mCol.医嘱内容).Value, RowRecord.Item(mCol.别名).Value)
            If InStr(";" & rsNumber!医嘱内容 & ";", ";" & str医嘱内容 & ";") <= 0 Then
                rsNumber!医嘱内容 = rsNumber!医嘱内容 & ";" & str医嘱内容
                
                
            End If
            rsNumber!医嘱ID串 = Replace(rsNumber!医嘱ID串 & "," & Replace(RowRecord.Item(mCol.医嘱id).Value & "," & _
                            RowRecord.Item(mCol.相关ID).Value & "," & RowRecord.Item(mCol.合并医嘱).Value, ";", ","), ",,", ",")
            rsNumber.Update
        End If
        
    End If
    rsNumber.filter = ""
End Function

Private Sub txtUnit_GotFocus()
    Me.txtUnit.SelStart = 0
    Me.txtUnit.SelLength = Len(Me.txtUnit)
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    Dim objPoint As POINTAPI
    Dim sglX As Single, sglY As Single
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    
    If KeyAscii = 13 Then
        
        If Len(Me.txtUnit.Text) < 2 Then
            MsgBox "必须输入1位以上的单位名称才能查询", vbInformation, "提示"
            Me.txtUnit.SetFocus
            Exit Sub
        End If
        strSQL = "select /*+ rule */ distinct 病人id id,工作单位 from 病人信息 where 登记时间 <= sysdate - (365/2) and 工作单位 like [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "%" & Me.txtUnit.Text & "%")
        Call ClientToScreen(txtUnit.hWnd, objPoint)
        sglX = objPoint.X * 15 - 30
        sglY = objPoint.Y * 15 + txtUnit.Height
        If frmSelectList.ShowSelect(Me, rsTmp, "工作单位,3000,0,0", sglX, sglY, txtUnit.Width, 2000, Me.Name, "请选择试工作单位") Then
            Me.txtUnit = rsTmp!工作单位
            Me.txtUnit.SelStart = 0
            Me.txtUnit.SelLength = Len(Me.txtUnit)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



