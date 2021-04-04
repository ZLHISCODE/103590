VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm按票据号批量发药 
   Caption         =   "按票据号发药"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   Icon            =   "Frm按票据号批量发药.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9705
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt医保号 
      Height          =   300
      Left            =   2520
      TabIndex        =   13
      Top             =   180
      Width           =   1725
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   7
      Top             =   5790
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   2580
      TabIndex        =   6
      Top             =   5790
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrintSet 
      Caption         =   "设置(&S)"
      Height          =   350
      Left            =   1350
      TabIndex        =   5
      Top             =   5790
      Visible         =   0   'False
      Width           =   1100
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   3165
      Left            =   30
      TabIndex        =   2
      Top             =   2400
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5583
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "单据明细(&D)"
      TabPicture(0)   =   "Frm按票据号批量发药.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Msf待发明细"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "药品汇总(&T)"
      TabPicture(1)   =   "Frm按票据号批量发药.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf待发汇总"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf待发明细 
         Height          =   2745
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4842
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf待发汇总 
         Height          =   2745
         Left            =   -74940
         TabIndex        =   9
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4842
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.TextBox TxtNo 
      Height          =   300
      Left            =   660
      TabIndex        =   0
      Top             =   180
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8220
      TabIndex        =   4
      Top             =   5790
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6930
      TabIndex        =   3
      Top             =   5790
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf待发列表 
      Height          =   1755
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3096
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl医保号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保号"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1920
      TabIndex        =   12
      Top             =   240
      Width           =   540
   End
   Begin VB.Label LblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "未输入任何处方"
      ForeColor       =   &H80000002&
      Height          =   180
      Left            =   4440
      TabIndex        =   11
      Top             =   240
      Width           =   3630
   End
   Begin VB.Label LblNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "票据号"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "Frm按票据号批量发药"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--外部传递参数--
Private mblnModify As Boolean
Private strUnit As String
Private strPrivs As String
Private mint服务对象 As Integer                     '药房的服务对象：1-门诊病人;2-住院病人;3-门诊和住院
Private lng药房ID As Long                           '药房
Private IntSendAfterDosage As Integer               '允许未配药发药
Private Int允许未审核处方发药 As Integer            '允许未审核处方发药
Private mint允许未收费处方发药 As Integer           '允许未收费处方发药
Private IntCheckStock As Integer                    '库存检测
Private Int校验处方 As Integer                      '校验处方
Private Str窗口 As String                           '发药窗口
Private int金额保留位数 As Integer                  '费用金额保留位数
Private int审核划价单 As Integer                    '执行后自动审核划价单
Private mint金额显示 As Integer                     '金额显示方式：0-显示应收金额,1-显示实收金额,2-显示应收和实收金额
Private mstrOpr As String
Private mblnConPacker As Boolean
Private mblnLoadDrug As Boolean
Private mbln启用审方 As Boolean

'--本程序使用变量--
Private RecBill As New ADODB.Recordset              '单据记录
Private RecTotal As New ADODB.Recordset             '汇总数据
Private BlnStartUp As Boolean
Private LngListRow As Long                          '待发列表
Private LngDetailRow As Long                        '待发明细
Private LngTotalRow As Long                         '待发汇总
Private StrBillNo As String                         '汇总单据号
Private strID As String                             '汇总ID

Private LngBillCount As Long
Public str配药人 As String

Private rs序号 As ADODB.Recordset
Private mobjDrugMAC As Object
Private mobjPlugIn As Object             '外挂接口对象
Private mstrDeptNode As String

Public Property Get In_DrugMAC() As Object
    Set In_DrugMAC = mobjDrugMAC
End Property
Public Property Set In_DrugMAC(ByVal objVal As Object)
    Set mobjDrugMAC = objVal
End Property
Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Public Property Get In_启用发药() As Boolean
    In_启用发药 = mblnLoadDrug
End Property

Public Property Let In_启用发药(ByVal vNewValue As Boolean)
    mblnLoadDrug = vNewValue
End Property

Public Property Get In_自动发药() As Boolean
    In_自动发药 = mblnConPacker
End Property

Public Property Let In_自动发药(ByVal vNewValue As Boolean)
    mblnConPacker = vNewValue
End Property
Private Sub GetRecipe(ByVal intType As Integer, ByVal txtInput As TextBox)
    'intType：1－票据号；2－医保号
    Dim blnAdd As Boolean
    Dim strNo As String, IntBill As Integer
    Dim rstemp As New ADODB.Recordset
    Dim strInput As String
    Dim strsql As String
    
    If Trim(txtInput.Text) = "" Then Exit Sub
    strInput = Trim(UCase(txtInput.Text))
    
    If intType = 1 Then
        '根据输入的票据号提取处方
        gstrSQL = "Select Distinct A.No " & _
                 " From 票据打印内容 A,票据使用明细 B " & _
                 " Where A.ID=B.打印ID And A.数据性质=1 " & _
                 " And B.票种=1 And B.号码=[1]"
    Else
        '根据输入的医保号提取处方
        gstrSQL = "Select Distinct B.NO " & _
                " From 病人信息 A, 未发药品记录 B " & _
                " Where A.病人id = B.病人id And B.单据 = 8 And A.医保号 = [1] And B.库房id = [2]"
    End If
    On Error GoTo errHandle
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[根据输入的票据号提取处方]", strInput, lng药房ID)
    
    If rstemp.RecordCount = 0 Then
        MsgBox "没有找到任何处方！", vbInformation, gstrSysName
        GoTo ExitSub
        Exit Sub
    End If
    
    With rstemp
        Do While Not .EOF
            gstrSQL = " Select /*+ Rule*/ Distinct Decode(C.单据,8,'收费',9,'记帐') 类型,C.No,C.单据,A.已收费,Decode(A.配药人,Null,'','部门发药','',A.配药人) 配药人,P.名称 科室,B.姓名,B.标识号 住院号,'' 床号," & _
                " B.开单人 开单医生,B.操作员姓名 填制人,To_Char(C.填制日期,'yyyy-MM-dd') 填制日期,B.记录性质,B.门诊标志, d.病人类型 " & _
                " From 未发药品记录 A,门诊费用记录 B,药品收发记录 C,部门表 P,部门表 S, 病人信息 D " & IIf(Str窗口 = "", "", ",Table(Cast(f_Str2list([3]) As zlTools.t_Strlist)) E ") & IIf(mbln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & _
                " Where C.费用ID=B.ID And B.开单部门ID+0=P.ID And Nvl(C.库房ID,0)+0=S.ID and Nvl(A.库房ID,0)=Nvl(C.库房ID,0) And Mod(C.记录状态,3)=1 And A.No=C.No " & IIf(mbln启用审方, " and b.医嘱序号=k.医嘱id(+) and Q.id(+)=K.审方id and K.最后提交(+)=1 And (b.医嘱序号 is null or nvl(q.审查结果,0) = 1)", "") & _
                " And (C.库房ID+0=[2] OR C.库房ID IS NULL)" & IIf(Str窗口 = "", "", " And (C.发药窗口=E.Column_Value Or C.发药窗口 Is NULL)") & _
                " and Not Exists(select 1 from 药品收发记录 F where F.单据=C.单据 and F.库房id=C.库房id and F.no=C.no and 发药方式=-1) " & _
                " And C.单据 =8 And C.审核人 Is Null And C.单据=A.单据 And C.No=[1] and nvl(C.发药方式,-999)<>-1 And A.病人id=D.病人id(+) "     '增加一个条件，排除已标记为不发药的记录  by lyq 20050416
            
            If mstrDeptNode <> "" Then
                gstrSQL = gstrSQL & " And (P.站点 = [4] Or P.站点 Is Null)"
            End If
            
            If mint服务对象 = 3 Then
                strsql = Replace(gstrSQL, "'' 床号", "B.床号")
                strsql = Replace(strsql, "门诊费用记录", "住院费用记录")
                gstrSQL = gstrSQL & " Union All " & strsql
            ElseIf mint服务对象 = 2 Then
                gstrSQL = Replace(gstrSQL, "'' 床号", "B.床号")
                gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
            End If
            On Error GoTo errHandle
            Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(!NO), lng药房ID, Str窗口, mstrDeptNode)
            
            blnAdd = (RecBill.RecordCount <> 0)
            If blnAdd Then     '找到指定处方
                strNo = RecBill!NO
                IntBill = RecBill!单据
                txtInput.Tag = IntBill
                
                '如果已存在该单据，则退出
                blnAdd = Not SetLocateBill(strNo, False)
                
                '检测合法性
                If blnAdd Then blnAdd = Not (CheckBill(IntBill, strNo, Val(RecBill!记录性质), Val(RecBill!门诊标志)) <> 0)
                If blnAdd Then blnAdd = WriteSendListData()
                If blnAdd Then
                    LngBillCount = LngBillCount + 1
                    LblNote.Caption = IIf(LngBillCount = 0, "未输入任何处方", "已输入" & LngBillCount & "张处方")
                End If
            End If
            .MoveNext
        Loop
    End With
    
    '定位到刚才输入的处方单
    Call SetLocateBill(strNo, True)
    With Msf待发列表
        CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
    End With
    
    mblnModify = True
    If TabShow.Tab = 1 Then Call RefreshData
    Exit Sub
ExitSub:
    With txtInput
        .SelStart = 0
        .SelLength = Len(txtInput.Text)
        .SetFocus
    End With
Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Property Get In_窗口() As String
    In_窗口 = mstrOpr
End Property

Public Property Let In_窗口(ByVal vNewValue As String)
    mstrOpr = vNewValue
End Property

Public Property Get In_权限() As String
    In_权限 = strPrivs
End Property

Public Property Let In_权限(ByVal vNewValue As String)
    strPrivs = vNewValue
End Property

Public Property Get In_服务对象() As Integer
    In_服务对象 = mint服务对象
End Property

Public Property Let In_服务对象(ByVal vNewValue As Integer)
    mint服务对象 = vNewValue
End Property

Public Property Get In_校验处方() As Integer
    In_校验处方 = Int校验处方
End Property

Public Property Let In_校验处方(ByVal vNewValue As Integer)
    Int校验处方 = vNewValue
End Property

Public Property Get In_库存检查() As Integer
    In_库存检查 = IntCheckStock
End Property

Public Property Let In_库存检查(ByVal vNewValue As Integer)
    IntCheckStock = vNewValue
End Property

Public Property Get In_药房ID() As Long
    In_药房ID = lng药房ID
End Property

Public Property Let In_药房ID(ByVal vNewValue As Long)
    lng药房ID = vNewValue
    mstrDeptNode = GetDeptStationNode(lng药房ID)
End Property

Public Property Get In_发药窗口() As String
    In_发药窗口 = Str窗口
End Property

Public Property Let In_发药窗口(ByVal vNewValue As String)
    Str窗口 = vNewValue
End Property

Public Property Get In_允许未配药发药() As Integer
    In_允许未配药发药 = IntSendAfterDosage
End Property

Public Property Let In_允许未配药发药(ByVal vNewValue As Integer)
    IntSendAfterDosage = vNewValue
End Property

Public Property Get IN_允许未审核发药() As Integer
    IN_允许未审核发药 = Int允许未审核处方发药
End Property

Public Property Let IN_允许未审核发药(ByVal vNewValue As Integer)
    Int允许未审核处方发药 = vNewValue
End Property

Public Property Get IN_允许未收费发药() As Integer
    IN_允许未收费发药 = mint允许未收费处方发药
End Property

Public Property Let IN_允许未收费发药(ByVal vNewValue As Integer)
    mint允许未收费处方发药 = vNewValue
End Property

Public Property Get In_金额保留位数() As Integer
    In_金额保留位数 = int金额保留位数
End Property

Public Property Let In_金额保留位数(ByVal vNewValue As Integer)
    int金额保留位数 = vNewValue
End Property

Public Property Get IN_审核划价单() As Integer
    IN_审核划价单 = int审核划价单
End Property

Public Property Let IN_审核划价单(ByVal vNewValue As Integer)
    int审核划价单 = vNewValue
End Property

Private Sub SetFormat(Optional ByVal IntStyle As Integer = 1)
    Dim intCol As Integer
    '设置各列表控件的格式

    Select Case IntStyle
    Case 1
        With Msf待发列表
            .rows = 2
            .Cols = 11
    
            .TextMatrix(0, 0) = "类型"
            .TextMatrix(0, 1) = "NO"
            .TextMatrix(0, 2) = "科室"
            .TextMatrix(0, 3) = "姓名"
            .TextMatrix(0, 4) = "住院号"
            .TextMatrix(0, 5) = "床号"
            .TextMatrix(0, 6) = "收费员"
            .TextMatrix(0, 7) = "开单医生"
            .TextMatrix(0, 8) = "开单日期"
            .TextMatrix(0, 9) = "记录性质"
            .TextMatrix(0, 10) = "门诊标志"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(0) = 500
                .ColWidth(1) = 1000
                .ColWidth(2) = 1200
                .ColWidth(3) = 1000
                .ColWidth(4) = 1000
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 1000
                .ColWidth(8) = 1200
                .ColWidth(9) = 0
                .ColWidth(10) = 0
                
                .Row = 1
                Call RestoreFlexState(Msf待发列表, Me.Name)
                If glngSys \ 100 <> 1 Then
                    .ColWidth(2) = 0
                    .ColWidth(4) = 0
                    .ColWidth(5) = 0
                End If
                .ColWidth(7) = IIf(Int校验处方 = 1, 0, 1000)
            End If
        End With
    Case 2
        With Msf待发明细
            .rows = 2
            .Cols = 8
    
            .TextMatrix(0, 0) = "药品名称"
            .TextMatrix(0, 1) = "商品名"
            .TextMatrix(0, 2) = "规格"
            .TextMatrix(0, 3) = "单位"
            .TextMatrix(0, 4) = "单价"
            .TextMatrix(0, 5) = "数量"
            .TextMatrix(0, 6) = "应收金额"
            .TextMatrix(0, 7) = "实收金额"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
    
            If BlnStartUp = False Then
                .ColWidth(0) = 2000
                .ColWidth(2) = 1500
                .ColWidth(3) = 500
                .ColWidth(4) = 800
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 1000
                
                .Row = 1
                Call RestoreFlexState(Msf待发明细, Me.Name)
                If gint药品名称显示 = 2 Then
                    If .ColWidth(1) = 0 Then .ColWidth(1) = 2000
                Else
                    .ColWidth(1) = 0
                End If
                
                If mint金额显示 = 0 Then
                    .ColWidth(7) = 0
                    If .ColWidth(6) <= 0 Then .ColWidth(6) = 1000
                ElseIf mint金额显示 = 1 Then
                    .ColWidth(6) = 0
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                Else
                    If .ColWidth(6) <= 0 Then .ColWidth(6) = 1000
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                End If
            End If
        End With
    Case 3
        With Msf待发汇总
            .rows = 2
            .Cols = 9
    
            .TextMatrix(0, 0) = "序号"
            .TextMatrix(0, 1) = "药品名称"
            .TextMatrix(0, 2) = "商品名"
            .TextMatrix(0, 3) = "规格"
            .TextMatrix(0, 4) = "单位"
            .TextMatrix(0, 5) = "单价"
            .TextMatrix(0, 6) = "数量"
            .TextMatrix(0, 7) = "应收金额"
            .TextMatrix(0, 8) = "实收金额"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(0) = 500
                .ColWidth(1) = 2000
                .ColWidth(3) = 1500
                .ColWidth(4) = 500
                .ColWidth(5) = 800
                .ColWidth(6) = 800
                .ColWidth(7) = 1000
                .ColWidth(8) = 1000
                
                .Row = 1
                Call RestoreFlexState(Msf待发汇总, Me.Name)
                If gint药品名称显示 = 2 Then
                    If .ColWidth(2) = 0 Then .ColWidth(2) = 2000
                Else
                    .ColWidth(2) = 0
                End If
                
                If mint金额显示 = 0 Then
                    .ColWidth(8) = 0
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                ElseIf mint金额显示 = 1 Then
                    .ColWidth(7) = 0
                    If .ColWidth(8) <= 0 Then .ColWidth(8) = 1000
                Else
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                    If .ColWidth(8) <= 0 Then .ColWidth(8) = 1000
                End If
            End If
        End With
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    '启用电子签名时检查用户是否注册
    If gblnESign处方发药 = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Sub
        End If
    End If
    
    Call RefreshData
    If CheckStock = False Then Exit Sub
    If Not CheckCorrelation Then Exit Sub
    If Not CheckBillOperate Then Exit Sub
    If SendBill = False Then Exit Sub
    
    LngBillCount = 0
    LblNote.Caption = IIf(LngBillCount = 0, "未输入任何处方", "已输入" & LngBillCount & "张处方")
    
    '初始化
    strID = ""
    StrBillNo = ""
    TxtNo.Text = ""
    txt医保号.Text = ""
    
    With Msf待发汇总
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    With Msf待发列表
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    With Msf待发明细
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    CmdOK.Enabled = False
    TxtNo.SetFocus
End Sub

Private Sub cmdPrint_Click()
    Dim HisPrint As New zlPrint1Grd
    Dim HisRow As New zlTabAppRow
    Dim ArrayNo, IntArray As Integer
    Dim LngSelectRow As Long, intCol As Integer
    
    On Error Resume Next
    '取消表格的选择状态
    With Msf待发汇总
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngTotalRow > 0 And LngTotalRow < .rows Then
            .Row = LngTotalRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
    End With
    
    HisPrint.Title = "药品汇总"
    Set HisRow = New zlTabAppRow
    HisRow.Add "日期:" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")
    HisPrint.UnderAppRows.Add HisRow
    
    ArrayNo = Split(StrBillNo, ";")
    
    Set HisRow = New zlTabAppRow
    HisRow.Add "单据号:"
    HisPrint.BelowAppRows.Add HisRow
    For IntArray = 0 To UBound(ArrayNo)
        Set HisRow = New zlTabAppRow
        HisRow.Add Space(10) & ArrayNo(IntArray)
        HisPrint.BelowAppRows.Add HisRow
    Next
    
    Set HisPrint.Body = Msf待发汇总
    Select Case zlPrintAsk(HisPrint)
    Case 1
        zlPrintOrView1Grd HisPrint, 1
    Case 2
        zlPrintOrView1Grd HisPrint, 2
    Case 3
        zlPrintOrView1Grd HisPrint, 3
    End Select
    
    '恢复表格的选择状态
    With Msf待发汇总
        
        LngTotalRow = LngSelectRow
        .Row = LngTotalRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub CmdPrintSet_Click()
    zlPrintSet
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    BlnStartUp = False
    LngBillCount = 0
    
    mint金额显示 = Val(zldatabase.GetPara("金额显示方式", glngSys, 1341, 0))
    mbln启用审方 = ((gtype_UserSysParms.P240_药房处方审查 = 1 Or gtype_UserSysParms.P240_药房处方审查 = 3) And gtype_UserSysParms.P241_处方审查时机 = 2)
    
    strID = ""
    StrBillNo = ""
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
   
    BlnStartUp = True
End Sub

Private Function CheckBillOperate() As Boolean
    Dim n, i As Integer
    Dim Dbl金额 As Double
    
    For n = 1 To Msf待发列表.rows - 1
        If Msf待发列表.TextMatrix(n, 1) <> "" Then
            Msf待发列表.Row = n
            Call Msf待发列表_EnterCell
            DoEvents
            
            Dbl金额 = 0
            
            For i = 1 To Msf待发明细.rows - 2
                Dbl金额 = Dbl金额 + Val(Msf待发明细.TextMatrix(i, 7))
            Next
            
            If CheckBillControl(3, Val(Msf待发列表.RowData(n)), Msf待发列表.TextMatrix(n, 1), Dbl金额) = False Then
                Exit Function
            End If
        End If
    Next
    
    CheckBillOperate = True
End Function
Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 8505 Then Me.Width = 8505
    If Me.Height < 6165 Then Me.Height = 6165
    
    With LblNote
        .Left = Me.ScaleWidth - .Width - 100
    End With
    
    With CmdHelp
        .Top = Me.ScaleHeight - .Height - 100
    End With
    With CmdPrintSet
        .Top = CmdHelp.Top
        .Left = CmdHelp.Left + CmdHelp.Width + 100
    End With
    With CmdPrint
        .Top = CmdHelp.Top
        .Left = CmdPrintSet.Left + CmdPrintSet.Width + 100
    End With
    
    With CmdCancel
        .Top = CmdHelp.Top
        .Left = Me.ScaleWidth - .Width - 100
    End With
    With CmdOK
        .Top = CmdHelp.Top
        .Left = CmdCancel.Left - .Width - 100
    End With
    
    With Msf待发列表
        .Height = (CmdOK.Top - 200 - .Top) / 2
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With TabShow
        .Top = Msf待发列表.Top + Msf待发列表.Height + 100
        .Height = CmdOK.Top - 100 - .Top
        .Width = Msf待发列表.Width
    End With
    With Msf待发汇总
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
    With Msf待发明细
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(Msf待发汇总, Me.Name)
    Call SaveFlexState(Msf待发列表, Me.Name)
    Call SaveFlexState(Msf待发明细, Me.Name)
End Sub

Private Sub Msf待发汇总_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf待发汇总
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngTotalRow > 0 And LngTotalRow < .rows Then
            .Row = LngTotalRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngTotalRow = LngSelectRow
        .Row = LngTotalRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf待发汇总_GotFocus()
    With Msf待发汇总
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf待发汇总_LostFocus()
    With Msf待发汇总
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf待发列表_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf待发列表
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngListRow > 0 And LngListRow < .rows Then
            .Row = LngListRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H80000005
                If intCol <> 3 Then
                    .CellForeColor = &H80000008
                End If
            Next
            .Col = 0
        End If
        
        LngListRow = LngSelectRow
        .Row = LngListRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellBackColor = &H8000000D
            If intCol <> 3 Then
                .CellForeColor = &H80000005
            End If
        Next
        .Col = 0
        .Redraw = True
        
        If Trim(.TextMatrix(.Row, 1)) = "" Then
            With Msf待发明细
                .Clear
                .rows = 2
                Call SetFormat(2)
            End With
            Exit Sub
        End If
        
        '显示单据明细
        Call ReadBillData(.RowData(.Row), .TextMatrix(.Row, 1), Val(.TextMatrix(.Row, 9)), Val(.TextMatrix(.Row, 10)))
    End With
End Sub

Private Sub Msf待发列表_GotFocus()
    With Msf待发列表
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf待发列表_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNo As String, lng单据 As Long
    
    If KeyCode = vbKeyDelete Then
        With Msf待发列表
            lng单据 = Val(.TextMatrix(.Row, 0))
            strNo = .TextMatrix(.Row, 1)
            If .rows - 1 = 1 Then
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
                .TextMatrix(1, 6) = ""
                .TextMatrix(1, 7) = ""
                .TextMatrix(1, 8) = ""
                .TextMatrix(1, 9) = ""
                .TextMatrix(1, 10) = ""
                .RowData(1) = 0
            Else
                If Trim(.TextMatrix(.Row, 1)) <> "" Then .RemoveItem .Row: LngBillCount = LngBillCount - 1
            End If
            
            CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
            LblNote.Caption = IIf(LngBillCount = 0, "未输入任何处方", "已输入" & LngBillCount & "张处方")
        
            '删除该单据
            With rs序号
                If .RecordCount <> 0 Then .MoveFirst
                .Find "单据标识='" & strNo & "|" & lng单据 & "'"
                If Not .EOF Then .Delete
            End With
        End With
        
        Msf待发列表_EnterCell
        mblnModify = True
        If TabShow.Tab = 1 Then Call RefreshData
    End If
End Sub

Private Sub Msf待发列表_LostFocus()
    With Msf待发列表
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf待发明细_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf待发明细
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngDetailRow > 0 And LngDetailRow < .rows Then
            .Row = LngDetailRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngDetailRow = LngSelectRow
        .Row = LngDetailRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf待发明细_GotFocus()
    With Msf待发明细
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf待发明细_LostFocus()
    With Msf待发明细
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case TabShow.Tab
    Case 0
        Msf待发明细.ZOrder
        Msf待发明细_EnterCell
    Case 1
        Call RefreshData
        Msf待发汇总.ZOrder
        Msf待发汇总_EnterCell
    End Select
End Sub

Private Sub TxtNo_GotFocus()
    GetFocus TxtNo
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call GetRecipe(1, TxtNo)
End Sub

Private Function ReadData(ByVal StrQuery As String) As Boolean
    '--读取数据--

'    On Error Resume Next
'    err = 0
    ReadData = False
    On Error GoTo errHandle
    gstrSQL = StrQuery
    With RecBill
        If .State = 1 Then .Close

        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, "ReadData")
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
    End With

    If err <> 0 Then
        MsgBox "读取处方时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    ReadData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadBillData(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer) As Boolean
    Dim IntStyle As Integer
    Dim str序号 As String
    Dim str明细单位串 As String
    '--读取单据内容--
    'BillStyle-单据类型;BIllNO-单据号
    '单位显示根据服务对象来（门诊：门诊单位；住院或住院门诊：住院单位；其它；售价单位）
    On Error Resume Next
    err = 0
    ReadBillData = False
    
    strUnit = GetUnit(lng药房ID, BillStyle, BillNo, int门诊标志)
    Select Case strUnit
    Case "售价单位"
        str明细单位串 = "C.计算单位 单位,B.零售价 单价,B.实际数量*Nvl(B.付数,1) 数量"
    Case "门诊单位"
        str明细单位串 = "D.门诊单位 单位,B.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装) 单价,B.实际数量/Decode(D.门诊包装,Null,1,0,1,D.门诊包装)*Nvl(B.付数,1) 数量"
    Case "住院单位"
        str明细单位串 = "D.住院单位 单位,B.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装) 单价,B.实际数量/Decode(D.住院包装,Null,1,0,1,D.住院包装)*Nvl(B.付数,1) 数量"
    Case "药库单位"
        str明细单位串 = "D.药库单位 单位,B.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装) 单价,B.实际数量/Decode(D.药库包装,Null,1,0,1,D.药库包装)*Nvl(B.付数,1) 数量"
    End Select
    str明细单位串 = str明细单位串 & ",B.零售金额 金额,Nvl(B.付数, 1) * B.实际数量 / (Nvl(F.付数, 1) * F.数次) * F.实收金额 As 实收金额 "
    
    gstrSQL = " SELECT DISTINCT F.序号,'['||C.编码||']'|| C.名称 As 品名,A.名称 As 商品名, " & _
            " DECODE(C.规格,NULL,B.产地,DECODE(B.产地,NULL,C.规格,C.规格||'|'||B.产地)) 规格," & _
            str明细单位串 & _
            " FROM 药品收发记录 B,药品规格 D,收费项目目录 C,收费项目别名 A,门诊费用记录 F" & _
            " WHERE B.药品ID=D.药品ID AND D.药品ID=C.ID" & _
            " AND d.药品ID=A.收费细目ID(+) AND A.性质(+)=3 And B.费用ID=F.ID" & _
            " AND MOD(B.记录状态,3)=1 AND B.NO=[1] AND B.单据=[2] " & _
            " AND (B.库房ID+0=[3] OR B.库房ID IS NULL) " & _
            " And 审核人 Is Null And Nvl(F.费用状态,0)<>1 " & _
            " Order by F.序号"
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        gstrSQL = Replace(gstrSQL, "And Nvl(F.费用状态,0)<>1", "")
    End If
    On Error GoTo errHandle
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, BillNo, BillStyle, lng药房ID)
        
    With RecBill
        str序号 = ""
        Do While Not .EOF
            str序号 = str序号 & "," & !序号
            .MoveNext
        Loop
        If str序号 <> "" Then str序号 = Mid(str序号, 2)
        .MoveFirst
    End With
    
    '将单据信息与明细序号写入内部映射记录集中
    With rs序号
        If .RecordCount <> 0 Then .MoveFirst
        .Find "单据标识='" & BillNo & "|" & BillStyle & "'"
        If str序号 <> "" Then
            If .EOF Then
                .AddNew
                !单据标识 = BillNo & "|" & BillStyle
                !序号 = str序号
                !记录性质 = int记录性质
                !门诊标志 = int门诊标志
                .Update
            End If
        End If
    End With
    
    If WriteDataToBill() = False Then Exit Function

    If err <> 0 Then
        MsgBox "读取处方时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    ReadBillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBill(ByVal IntBillStyle As Integer, ByVal strNo As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer) As Integer
    Dim RecCheck As New ADODB.Recordset

    '--根据将要执行的操作，判断是否允许--
    '返回:
    '0-允许操作
    '1-未配药
    '2-已配药
    '3-已发药
    '4-已删除
    '5-未发药
    On Error GoTo errHandle
    gstrSQL = " Select A.配药人,A.审核人,nvl(B.已收费,0) 已收费, C.操作员姓名 填制人 " & _
            " From 药品收发记录 A,未发药品记录 B, 门诊费用记录 C " & _
            " Where A.No=B.No And A.单据=B.单据 And A.费用id = C.ID And A.审核人 IS Null And mod(A.记录状态,3)=1 And Rownum=1 " & _
            " And A.No=[1] And A.单据=[2]  And (A.库房ID+0=[3] Or A.库房ID Is NULL)"
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lng药房ID)
        
    With RecCheck
        If .EOF Then CheckBill = 4: MsgBox "未找到指定单据,可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!审核人) Then
            CheckBill = 3: MsgBox "该处方已被其它操作员发药，操作被迫中止！", vbInformation, gstrSysName: Exit Function
        End If
        If !已收费 = 0 Then
            CheckBill = 3: MsgBox "该处方还未收费，操作被迫中止！", vbInformation, gstrSysName: Exit Function
        End If
    End With
     
    If mint允许未收费处方发药 = 0 And IntBillStyle = 8 Then
        If RecCheck!已收费 = 0 Then
            MsgBox "该处方还未收费，操作被迫中止！", vbInformation, gstrSysName
            CheckBill = 5
            Exit Function
        End If
    End If
    
    If Int允许未审核处方发药 = 0 And IntBillStyle <> 8 Then
        If IsNull(RecCheck!填制人) Then
            MsgBox "该处方还未审核，操作被迫中止！", vbInformation, gstrSysName
            CheckBill = 5
            Exit Function
        End If
    End If

    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteSendListData() As Boolean
    Dim RecCheck As New ADODB.Recordset
    
    WriteSendListData = False
    
    If IntSendAfterDosage = 0 Then
        If IsNull(RecBill!配药人) Then
            MsgBox "处方" & RecBill!NO & "还未配药，不允许加入发药列表！", vbInformation, gstrSysName
            Exit Function
        End If
        If Trim(RecBill!配药人) = "" Then
            MsgBox "处方" & RecBill!NO & "还未配药，不允许加入发药列表！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mint允许未收费处方发药 = 0 And RecBill!单据 = 8 Then
        If RecBill!已收费 = 0 Then
            MsgBox "处方" & RecBill!NO & "还未收费，不允许加入发药列表！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If Int允许未审核处方发药 = 0 And RecBill!单据 <> 8 Then
        If IsNull(RecBill!填制人) Then
            MsgBox "处方" & RecBill!NO & "还未审核，不允许加入发药列表！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With Msf待发列表
        .Redraw = False
        .TextMatrix(.rows - 1, 0) = RecBill!类型
        .TextMatrix(.rows - 1, 1) = RecBill!NO
        .TextMatrix(.rows - 1, 2) = IIf(IsNull(RecBill!科室), "", RecBill!科室)
        .TextMatrix(.rows - 1, 3) = IIf(IsNull(RecBill!姓名), "", RecBill!姓名)
        .TextMatrix(.rows - 1, 4) = IIf(IsNull(RecBill!住院号), "", RecBill!住院号)
        .TextMatrix(.rows - 1, 5) = IIf(IsNull(RecBill!床号), "", RecBill!床号)
        .TextMatrix(.rows - 1, 6) = IIf(IsNull(RecBill!填制人), "", RecBill!填制人)
        .TextMatrix(.rows - 1, 7) = IIf(IsNull(RecBill!开单医生), "", RecBill!开单医生)
        .TextMatrix(.rows - 1, 8) = IIf(IsNull(RecBill!填制日期), "", RecBill!填制日期)
        .TextMatrix(.rows - 1, 9) = RecBill!记录性质
        .TextMatrix(.rows - 1, 10) = RecBill!门诊标志
        .RowData(.rows - 1) = RecBill!单据
        
        .Row = .rows - 1
        .Col = 3
        .CellForeColor = zldatabase.GetPatiColor(IIf(IsNull(RecBill!病人类型), "", RecBill!病人类型))

        .rows = .rows + 1
        .RowData(.rows - 1) = 0
        .Redraw = True
    End With
    WriteSendListData = True
End Function

Private Function RefreshData() As Boolean
    Dim intRow As Integer, intRows As Integer
    Dim arrID
    Dim StrNoThis As String, IntBillThis As Integer
    Dim str汇总单位串 As String
    If mblnModify = False Then Exit Function
    RefreshData = False
    
    '清空汇总表格
    On Error GoTo errHandle
    With Msf待发汇总
        .Clear
        .rows = 2
        SetFormat (3)
    End With
    
    strID = ""
    StrBillNo = ""
    With Msf待发列表
    
        '获得NO号
        For intRow = 1 To .rows - 1
            If intRow = 1 Then
                StrBillNo = StrBillNo & .TextMatrix(intRow, 1)
            Else
                If Trim(.TextMatrix(intRow, 1)) <> "" Then
                    If intRow Mod 8 = 0 Then StrBillNo = StrBillNo & ";"
                    StrBillNo = StrBillNo & "," & .TextMatrix(intRow, 1)
                End If
            End If
        Next
        
        '组合ID
        For intRow = 1 To .rows - 1
            StrNoThis = .TextMatrix(intRow, 1)
            IntBillThis = .RowData(intRow)
            
            gstrSQL = " Select ID From 药品收发记录 Where No=[1] And 单据=[2] " & _
                " And Mod(记录状态,3)=1 And 审核人 Is Null And (库房ID+0=[3] Or 库房ID Is NULL)"
            Set RecTotal = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, StrNoThis, IntBillThis, lng药房ID)
            
            With RecTotal
                Do While Not .EOF
                    strID = strID & IIf(strID = "", "", ",") & !Id
                    .MoveNext
                Loop
            End With
        Next
    End With
    If strID = "" Then Exit Function
    
    '显示汇总数据
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("药房属性", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(lng药房ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(lng药房ID, gint门诊药房)
    Else
        strUnit = GetSpecUnit(lng药房ID, gint住院药房)
    End If
    Select Case strUnit
    Case "售价单位"
        str汇总单位串 = "C.计算单位 单位,B.零售价 单价,Sum(B.实际数量*Nvl(B.付数,1)) 数量"
    Case "门诊单位"
        str汇总单位串 = "D.门诊单位 单位,B.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装) 单价,Sum(B.实际数量/Decode(D.门诊包装,Null,1,0,1,D.门诊包装)*Nvl(B.付数,1)) 数量"
    Case "住院单位"
        str汇总单位串 = "D.住院单位 单位,B.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装) 单价,Sum(B.实际数量/Decode(D.住院包装,Null,1,0,1,D.住院包装)*Nvl(B.付数,1)) 数量"
    Case "药库单位"
        str汇总单位串 = "D.药库单位 单位,B.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装) 单价,Sum(B.实际数量/Decode(D.药库包装,Null,1,0,1,D.药库包装)*Nvl(B.付数,1)) 数量"
    End Select
    str汇总单位串 = str汇总单位串 & ",Sum(B.零售金额) 金额,Sum(Nvl(B.付数, 1) * B.实际数量 / (Nvl(B.费用付数, 1) * B.数次) * B.实收金额) As 实收金额  "
    
    gstrSQL = " Select A.No, A.药品id, A.批次, A.零售价, A.实际数量, A.付数, A.零售金额, B.付数 As 费用付数,B.数次, B.实收金额, A.产地 " & _
        " From 药品收发记录 A, 门诊费用记录 B , Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) C " & _
        " Where A.费用id = B.Id And A.Id =C.Column_Value "
    
    gstrSQL = "Select Distinct D.*,'['||D.编码||']'|| D.通用名称 As 品名,A.名称 As 商品名 " & _
             " From " & _
             "     (SELECT D.药品ID,C.编码,C.名称 通用名称,NVL(B.批次,0) 批次," & _
             "     DECODE(C.规格,NULL,B.产地,DECODE(B.产地,NULL,C.规格,C.规格||'|'||B.产地)) 规格," & str汇总单位串 & _
             "     FROM (" & gstrSQL & ") B," & _
             "           药品规格 D,收费项目目录 C " & _
             "     WHERE B.药品ID+0=D.药品ID AND D.药品ID=C.ID" & _
             "     GROUP BY D.药品ID,C.编码,C.名称,NVL(B.批次,0)," & _
             "     DECODE(C.规格,NULL,B.产地,DECODE(B.产地,NULL,C.规格,C.规格||'|'||B.产地)),"
    Select Case strUnit
    Case "售价单位"
        gstrSQL = gstrSQL & "C.计算单位,B.零售价"
    Case "门诊单位"
        gstrSQL = gstrSQL & "D.门诊单位,B.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装)"
    Case "住院单位"
        gstrSQL = gstrSQL & "D.住院单位,B.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装)"
    Case "药库单位"
        gstrSQL = gstrSQL & "D.药库单位,B.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装)"
    End Select
    gstrSQL = gstrSQL & ") D,收费项目别名 A" & _
            " Where D.药品ID=A.收费细目ID(+) AND A.性质(+)=3"
    gstrSQL = gstrSQL & " Order By D.编码"
    
    Set RecTotal = zldatabase.OpenSQLRecord(gstrSQL, "RefreshData", strID)
    
    Call WriteTotalDataToBill
    
    If err <> 0 Then
        MsgBox "显示汇总数据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnModify = False
    RefreshData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteTotalDataToBill() As Boolean
    Dim dbl应收金额 As Double
    Dim dbl实收金额 As Double
    Dim str金额显示 As String
    
    '将汇总数据装入
    On Error Resume Next
    err = 0
    
    WriteTotalDataToBill = False
    With Msf待发汇总
        .Clear
        .rows = 2
        Call SetFormat(3)
    End With
    
    '填充单据内容
    With RecTotal
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Msf待发汇总.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Msf待发汇总.TextMatrix(.AbsolutePosition, 1) = !品名
            Msf待发汇总.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!商品名), "", !商品名)
            Msf待发汇总.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!规格), "", !规格)
            Msf待发汇总.TextMatrix(.AbsolutePosition, 4) = IIf(IsNull(!单位), "", !单位)
            Msf待发汇总.TextMatrix(.AbsolutePosition, 5) = Format(!单价, "#####0.00000;-#####0.00000; ;")
            Msf待发汇总.TextMatrix(.AbsolutePosition, 6) = Format(!数量, "#####0.00000;-#####0.00000; ;")
            Msf待发汇总.TextMatrix(.AbsolutePosition, 7) = Format(!金额, "#####0.00;-#####0.00; ;")
            Msf待发汇总.TextMatrix(.AbsolutePosition, 8) = Format(!实收金额, "#####0.00;-#####0.00; ;")
            Msf待发汇总.MergeRow(.AbsolutePosition) = False
            dbl应收金额 = dbl应收金额 + !金额
            dbl实收金额 = dbl实收金额 + !实收金额
            
            If .AbsolutePosition >= Msf待发汇总.rows - 1 Then Msf待发汇总.rows = Msf待发汇总.rows + 1
            .MoveNext
        Loop
        
        '显示合计
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 0) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 1) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 2) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 3) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 4) = "合计"
        
        If mint金额显示 = 1 Then
            str金额显示 = "实收金额：" & Format(dbl实收金额, "#####0.00;-#####0.00; ;")
        ElseIf mint金额显示 = 2 Then
            str金额显示 = "应收金额：" & Format(dbl应收金额, "#####0.00;-#####0.00; ;") & "    实收金额：" & Format(dbl实收金额, "#####0.00;-#####0.00; ;")
        Else
            str金额显示 = "应收金额：" & Format(dbl应收金额, "#####0.00;-#####0.00; ;")
        End If
        
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 5) = str金额显示
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 6) = str金额显示
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 7) = str金额显示
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 8) = str金额显示
        Msf待发汇总.MergeCells = flexMergeFree
        Msf待发汇总.MergeRow(Msf待发汇总.rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "显示单据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    WriteTotalDataToBill = True
End Function

Private Function WriteDataToBill() As Boolean
    Dim dbl应收金额 As Double
    Dim dbl实收金额 As Double
    Dim str金额显示 As String
    
    '--显示指定处方的明细--
    On Error Resume Next
    err = 0
    
    WriteDataToBill = False
    With Msf待发明细
        .Clear
        .rows = 2
        Call SetFormat(2)
    End With
    
    '填充单据内容
    With RecBill
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Msf待发明细.MergeRow(.AbsolutePosition) = False
            Msf待发明细.TextMatrix(.AbsolutePosition, 0) = !品名
            Msf待发明细.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!商品名), "", !商品名)
            Msf待发明细.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!规格), "", !规格)
            Msf待发明细.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!单位), "", !单位)
            Msf待发明细.TextMatrix(.AbsolutePosition, 4) = Format(!单价, "#####0.00000;-#####0.00000; ;")
            Msf待发明细.TextMatrix(.AbsolutePosition, 5) = Format(!数量, "#####0.00000;-#####0.00000; ;")
            Msf待发明细.TextMatrix(.AbsolutePosition, 6) = Format(!金额, "#####0.00;-#####0.00; ;")
            Msf待发明细.TextMatrix(.AbsolutePosition, 7) = Format(!实收金额, "#####0.00;-#####0.00; ;")
            dbl应收金额 = dbl应收金额 + Val(!金额)
            dbl实收金额 = dbl实收金额 + Val(!实收金额)
            
            If .AbsolutePosition >= Msf待发明细.rows - 1 Then Msf待发明细.rows = Msf待发明细.rows + 1
            .MoveNext
        Loop
    End With
    With Msf待发明细
        .TextMatrix(.rows - 1, 0) = "合计"
        .TextMatrix(.rows - 1, 1) = "合计"
        .TextMatrix(.rows - 1, 2) = "合计"
        .TextMatrix(.rows - 1, 3) = "合计"
        
        If mint金额显示 = 1 Then
            str金额显示 = "实收金额：" & Format(dbl实收金额, "#####0.00;-#####0.00; ;")
        ElseIf mint金额显示 = 2 Then
            str金额显示 = "应收金额：" & Format(dbl应收金额, "#####0.00;-#####0.00; ;") & "    实收金额：" & Format(dbl实收金额, "#####0.00;-#####0.00; ;")
        Else
            str金额显示 = "应收金额：" & Format(dbl应收金额, "#####0.00;-#####0.00; ;")
        End If
        
        .TextMatrix(.rows - 1, 4) = str金额显示
        .TextMatrix(.rows - 1, 5) = str金额显示
        .TextMatrix(.rows - 1, 6) = str金额显示
        .TextMatrix(.rows - 1, 7) = str金额显示
        
        .MergeCells = flexMergeFree
        .MergeRow(.rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "显示单据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    WriteDataToBill = True
End Function

Private Function SetLocateBill(Optional ByVal strNo As String = "", _
    Optional ByVal BlnEnterCell As Boolean = True) As Boolean
    Dim intRow As Integer
    
    SetLocateBill = False
    With Msf待发列表
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 1) = strNo And .RowData(intRow) = 8 Then
                .Row = intRow
                .TopRow = intRow
                SetLocateBill = True
                Exit For
            End If
        Next
    End With
    
    If BlnEnterCell Then Msf待发列表_EnterCell
End Function

Private Function CheckStock() As Boolean
    Dim RecCheckStock As New ADODB.Recordset
    Dim dblStock As Double
    Dim strSubSql As String
    '检查库存
    If IntCheckStock = 0 Then CheckStock = True: Exit Function
    
    '将库存数量转换为对应单位的实际数量
    Dim intUnit As Integer
    On Error GoTo errHandle
    intUnit = Val(zldatabase.GetPara("药房属性", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(lng药房ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(lng药房ID, gint门诊药房)
    Else
        strUnit = GetSpecUnit(lng药房ID, gint住院药房)
    End If
    Select Case strUnit
    Case "售价单位"
        strSubSql = "/1"
    Case "门诊单位"
        strSubSql = "/Decode(B.门诊包装,Null,1,0,1,B.门诊包装)"
    Case "住院单位"
        strSubSql = "/Decode(B.住院包装,Null,1,0,1,B.住院包装)"
    Case "药库单位"
        strSubSql = "/Decode(B.药库包装,Null,1,0,1,B.药库包装)"
    End Select
    
    CheckStock = False
    With RecTotal
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
           gstrSQL = " Select nvl(实际数量,0)" & strSubSql & " AS 数量" & _
                 " From 药品库存 A,药品规格 B" & _
                 " Where B.药品ID=A.药品ID And A.性质=1 And A.库房ID=[1] And A.药品ID=[2] And Nvl(A.批次,0)=[3]"
           Set RecCheckStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID, CLng(RecTotal!药品id), CLng(RecTotal!批次))
           
           With RecCheckStock
                If .EOF Then
                    dblStock = 0
                Else
                    dblStock = !数量
                End If
                
                If dblStock < RecTotal!数量 Then
                    If RecTotal!批次 <> 0 Then
                        MsgBox RecTotal!品名 & "的批次库存数不够，不能继续发药！", vbInformation, gstrSysName: Exit Function
                    Else
                        Select Case IntCheckStock
                        Case 1
                            If MsgBox(RecTotal!品名 & "的库存数不够，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox RecTotal!品名 & "的库存数不够，不能继续发药！", vbInformation, gstrSysName: Exit Function
                        End Select
                    End If
                End If
            End With
            .MoveNext
        Loop
    End With
    
    CheckStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendBill() As Boolean
    Dim intRow As Integer
    Dim StrDate As String
    Dim rsSendRecipeByNo As ADODB.Recordset
    Dim int门诊 As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnInTrans As Boolean
    Dim str签名记录 As String
    Dim strReturn As String
    Dim strNo As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    
    arrSql = Array()

    SendBill = False
    
    Set rsSendRecipeByNo = New ADODB.Recordset
    With rsSendRecipeByNo
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 2, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 1, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    StrDate = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With Msf待发列表
        For intRow = 1 To .rows - 1
            If .RowData(intRow) <> 0 Then
                '零差价管理
                If CheckPriceAdjustByNO(Val(.RowData(intRow)), lng药房ID, .TextMatrix(intRow, 1)) = False Then
                    Exit Function
                End If
                
                With rsSendRecipeByNo
                    .AddNew
                    !NO = Msf待发列表.TextMatrix(intRow, 1)
                    !单据 = Msf待发列表.RowData(intRow)
                    !记录性质 = Val(Msf待发列表.TextMatrix(intRow, 9))
                    !门诊标志 = Val(Msf待发列表.TextMatrix(intRow, 10))
                    .Update
                End With
            End If
        Next
    End With
    
    '按处方号排序后批量发药
    rsSendRecipeByNo.Sort = "NO"
    rsSendRecipeByNo.MoveFirst
    For intRow = 1 To rsSendRecipeByNo.RecordCount
        '先检查或执行预调价
        Call AutoAdjustPrice_ByNO(rsSendRecipeByNo!单据, rsSendRecipeByNo!NO)
        
        If Val(rsSendRecipeByNo!记录性质) = 1 Or (Val(rsSendRecipeByNo!记录性质) = 2 And (Val(rsSendRecipeByNo!门诊标志) = 1 Or Val(rsSendRecipeByNo!门诊标志) = 4)) Then
            int门诊 = 1
        Else
            int门诊 = 2
        End If
        
        gstrSQL = "zl_药品收发记录_处方发药("
        '库房ID
        gstrSQL = gstrSQL & lng药房ID
        '单据
        gstrSQL = gstrSQL & "," & rsSendRecipeByNo!单据
        'NO
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
        '审核人
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '配药人
        gstrSQL = gstrSQL & ",'" & str配药人 & "'"
        '校验人
        gstrSQL = gstrSQL & ",NULL"
        '发药方式
        gstrSQL = gstrSQL & ",2"
        '发药时间
        gstrSQL = gstrSQL & ",to_date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')"
        '操作员编号
        gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
        '操作员姓名
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '金额保留位数
        gstrSQL = gstrSQL & "," & int金额保留位数
        '审核划价单
        gstrSQL = gstrSQL & "," & int审核划价单
        '是否门诊
        gstrSQL = gstrSQL & "," & int门诊
        gstrSQL = gstrSQL & ")"

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        
        strNo = strNo & rsSendRecipeByNo!单据 & "," & rsSendRecipeByNo!NO & "|"
        rsSendRecipeByNo.MoveNext
    Next
    
    '调用发药前的外挂接口
    err.Clear
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        If mobjPlugIn.DrugBeforeSendByRecipe(lng药房ID, strNo, strReserve) = False Then
            If err.Number <> 0 Then
                err.Clear: On Error GoTo 0
            Else
                Exit Function
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo ErrHand
    
    '先处理发药事务
    gcnOracle.BeginTrans
    blnInTrans = True
    
    '单独处理电子签名，放到业务处理前面，防止弹框造成数据死锁
    '如果已启用了电子签名，则需要对配药人进行电子签名处理
    If gblnESign处方发药 = True And gblnESignUserStoped = False Then
        rsSendRecipeByNo.MoveFirst
        For intRow = 1 To rsSendRecipeByNo.RecordCount
            str签名记录 = ""
            If GetSignatureRecored(EsignTache.send, rsSendRecipeByNo!单据, rsSendRecipeByNo!NO, lng药房ID, str签名记录, 0, CDate(StrDate), gstrUserName) = False Then
                gcnOracle.RollbackTrans
                blnInTrans = False
                Exit Function
            End If
            
            If str签名记录 <> "" Then
                gstrSQL = "Zl_药品签名记录_Insert(" & str签名记录 & ")"
               
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            Else
                gcnOracle.RollbackTrans
                blnInTrans = False
                MsgBox "对发药人电子签名失败！", vbInformation, gstrSysName
                Exit Function
            End If
            
            rsSendRecipeByNo.MoveNext
        Next
    End If
    
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_Abolish")
    Next
    
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        If mblnConPacker And strNo <> "" And mblnLoadDrug Then
            Call mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.用户编码, UserInfo.用户姓名, lng药房ID, Mid(strNo, 1, Len(strNo) - 1), strReturn)
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        If mblnConPacker Then
            If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
            mobjDrugMAC.Operation gstrDbUser, Val("22-开始发药"), "1|" & Replace(strNo, "|", ";"), strReturn
        End If
    End If
    
    If MsgBox("你需要打印汇总清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_2", "ZL8_BILL_1341_2"), Me, "库房=" & lng药房ID, "发药方式=批量发药|2", "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "发药时间=" & StrDate, 2)
    End If
    
    '调用发药后的外挂接口
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        mobjPlugIn.DrugSendByRecipe lng药房ID, strNo, CDate(StrDate), strReserve
        err.Clear: On Error GoTo 0
    End If
    
    SendBill = True
    Exit Function
ErrHand:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng单据 As Long, str序号 As String
    '检查处方是否已结帐、检查该病人是否已出院，并对权限进行检查
    With rs序号
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !单据标识
            lng单据 = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            str序号 = nvl(!序号)
            If Not IsReceiptBalance_Charge(0, strPrivs, lng单据, strNo, str序号, Val(!记录性质), Val(!门诊标志)) Then Exit Function
            If Not IsOutPatient(strPrivs, lng单据, strNo, Val(!记录性质), Val(!门诊标志)) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function

Private Sub InitRec()
    Set rs序号 = New ADODB.Recordset
    With rs序号
        If .State = 1 Then .Close
        .Fields.Append "单据标识", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "序号", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub txt医保号_GotFocus()
    GetFocus txt医保号
End Sub


Private Sub txt医保号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call GetRecipe(2, txt医保号)
End Sub


