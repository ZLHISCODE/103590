VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm特殊项目审批_贵阳 
   Caption         =   "特殊项目审批_贵阳"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frm特殊项目审批_贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11880
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   5985
      Width           =   11880
      Begin VB.CommandButton cmd全报 
         Caption         =   "全报(&A)"
         Height          =   350
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmd更新 
         Caption         =   "更新(&O)"
         Height          =   350
         Left            =   7320
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmd全自费 
         Caption         =   "全自费(&C)"
         Height          =   350
         Left            =   1200
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txt住院号 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   3360
         TabIndex        =   3
         Top             =   150
         Width           =   1455
      End
      Begin VB.CommandButton cmd日志 
         Caption         =   "日志(&L)"
         Height          =   345
         Left            =   6090
         TabIndex        =   2
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   555
         Width           =   4755
      End
      Begin VB.Label Label1 
         Caption         =   "!特别说明：打'√'表示报销,打'Х'表示自费,空表示不更新"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         TabIndex        =   8
         Top             =   615
         Width           =   7935
      End
      Begin VB.Label lbl住院号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   2715
         TabIndex        =   7
         Top             =   210
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfDetail 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   10081
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm特殊项目审批_贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mintInsure As Integer

Public Sub ShowSelect(ByVal intinsure As Integer)
    On Error Resume Next
    mintInsure = intinsure
    mlng病人ID = 0
    mlng主页ID = 0
    Me.Show 1
End Sub

'由于：需要每一笔明细都需要特殊药品审批
'所以：原收费细目ID保存收费ID

Private Sub Cmd更新_Click()
    Dim str医保编码 As String
    Dim lngRow As Long, lngRows As Long
    Dim rsTemp As New ADODB.Recordset
    Dim str费用类型 As String
    
    Dim strTableD       As String
    Dim strWhereD       As String
    Dim i               As Integer
    Dim sFileName       As String
    
    
    If mlng病人ID = 0 Then
        MsgBox "请先确定病人!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    
    '添加日志
    
    '记录修改前日志
    ' 表表名(用分号";"隔开)
    strTableD = "特殊药品收费"
    ' 表的条件(用分号";"隔开)
    strWhereD = "病人ID='" & mlng病人ID & "' And 主页ID = '" & mlng主页ID & "'"
    ' 记录修改前的数据
    sFileName = EditFormerWriteFileA(strTableD, strWhereD)
    
    lngRows = msfDetail.Rows - 1
    For lngRow = 1 To lngRows
        '插入特殊药品收费表
        If msfDetail.TextMatrix(lngRow, 0) <> "" Then
            If msfDetail.RowData(lngRow) <> 0 Then
                gstrSQL = "ZL_特殊药品收费_Update(" & mlng病人ID & "," & mlng主页ID & "," & msfDetail.RowData(lngRow) & "," & IIf(msfDetail.TextMatrix(lngRow, 0) = "√", 1, 0) & ",'" & gstrUserName & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
        End If
    Next
    
    '提取所有特殊药品收费,逐条更新
    gstrSQL = " Select A.ID,A.NO,A.记录性质,A.记录状态,A.序号,A.收费类别,C.项目编码 AS 医保编码,B.标志" & _
              " From 病人费用记录 A,特殊药品收费 B,保险支付项目 C" & _
              " Where Nvl(A.实收金额,0)<>0 And Nvl(A.数次,0)<>0 And Nvl(A.附加标志,0)<>9 And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 And Nvl(A.婴儿费,0)=0 " & _
              " And A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.ID=B.费用ID And A.收费细目ID+0=C.收费细目ID And C.险类=[3]" & _
              " And B.病人ID=[1] And B.主页ID=[2]" & _
              " Order by A.收费细目ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有特殊药品收费,逐条更新", mlng病人ID, mlng主页ID, mintInsure)
    Do While Not rsTemp.EOF
        str费用类型 = ""
        If rsTemp!标志 <> 0 Then    '自费 或 报销 状态未设置 【不进行处理】
            If rsTemp!标志 = 1 Then
                str医保编码 = rsTemp!医保编码
            Else
               '自费西药编码：810851900099,20110121周玉强修改，医保中心启用新编码
                '自费中成药编码：820851900099
                '自费中草药编码：829000900099
                If rsTemp!收费类别 = "5" Then
                str医保编码 = "810851900099"
                ElseIf rsTemp!收费类别 = "6" Then
                str医保编码 = "820851900099"
                Else
                str医保编码 = "829000900099"
                End If
                str费用类型 = "药品全自费"
            End If
            gstrSQL = "zl_病人费用记录_上传('" & rsTemp!NO & "'," & rsTemp!序号 & "," & rsTemp!记录性质 & "," & rsTemp!记录状态 & "," & _
                      "'" & str医保编码 & "'," & IIf(str费用类型 = "", "NULL", "'" & str费用类型 & "'") & ",0)"
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
        End If
        
        rsTemp.MoveNext
    Loop
    
     '记录修改后日志
    Call EditFormerWriteFileA(strTableD, strWhereD, sFileName)
    '保存修改日志
    AddLog "医保工具", "特殊药品收费", DBConnLTEdit, , sFileName, CStr(mlng病人ID), CStr(mlng主页ID), , "特殊药品收费", , True
    
    gcnOracle.CommitTrans
    
    '清除表格，等待继续办理病人
    mlng病人ID = 0
    mlng主页ID = 0
    Me.txt住院号.Text = ""
    Me.txt住院号.Tag = ""
    Me.lblNote.Caption = ""
    
    
     msfDetail.Rows = 2: msfDetail.Cols = 10
    msfDetail.TextMatrix(0, 0) = "打勾报销"
    msfDetail.TextMatrix(0, 1) = "特殊项目"
    msfDetail.TextMatrix(0, 2) = "规格"
    msfDetail.TextMatrix(0, 3) = "限病种信息"
    msfDetail.TextMatrix(0, 4) = "记帐日期"
    msfDetail.TextMatrix(0, 5) = "单据号"
    msfDetail.TextMatrix(0, 6) = "数量"
    msfDetail.TextMatrix(0, 7) = "单价"
    msfDetail.TextMatrix(0, 8) = "金额"
    msfDetail.TextMatrix(0, 9) = "记帐人"
    
    
    msfDetail.ColWidth(0) = 1000
    msfDetail.ColWidth(1) = 2400
    msfDetail.ColWidth(2) = 1500
    msfDetail.ColWidth(3) = 3000
    msfDetail.ColWidth(4) = 1000
    msfDetail.ColWidth(5) = 800
    msfDetail.ColWidth(6) = 800
    msfDetail.ColWidth(7) = 800
    msfDetail.ColWidth(8) = 1000
    msfDetail.ColWidth(9) = 800
  
    
    MsgBox "更新成功！", vbInformation, gstrSysName
   
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub cmd全报_Click()
    Dim lngRow As Long, lngRows As Long
    lngRows = msfDetail.Rows - 1
    For lngRow = 1 To lngRows
        msfDetail.TextMatrix(lngRow, 0) = "√"
    Next
End Sub

Private Sub cmd全自费_Click()
    Dim lngRow As Long, lngRows As Long
    lngRows = msfDetail.Rows - 1
    For lngRow = 1 To lngRows
        msfDetail.TextMatrix(lngRow, 0) = "Х"
    Next
End Sub

Private Sub RefreshData()
  ' On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    '检查该病人是否存在特殊用药
    
    gstrSQL = " Select a.Id,to_char(a.登记时间,'yyyy-mm-dd') As 记帐日期,a.No As 单据号,a.姓名, Substr(c.规格,1,Instr(规格, '┆'))  as 规格,a.数次 As 数量,a.标准单价 As 单价,nvl(a.实收金额,0) As 金额,a.操作员姓名 As 记帐人,a.收费细目id, c.名称 As 特殊项目,c.编码, b.项目编码, c.说明 As 限病种信息," & _
             " Nvl(d.标志, 1) As 报销" & _
              " From 病人费用记录 A,保险支付项目 B,收费细目 C,特殊药品收费 D" & _
              " Where Nvl(A.实收金额,0)<>0 And Nvl(A.数次,0)<>0 And Nvl(A.附加标志,0)<>9 And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 And Nvl(A.婴儿费,0)=0 " & _
              " And C.类别 IN ('5','6','7') And C.说明 Is Not NULL And A.收费细目ID=C.ID And C.ID=B.收费细目ID And B.险类=[3]" & _
              " And A.病人ID=D.病人ID(+) And A.主页ID=D.主页ID(+) And A.收费细目ID=c.ID(+)  and  a.Id=d.费用id(+) " & _
              " And A.病人ID=[1] And A.主页ID=[2]" & _
              " Order by C.编码,a.登记时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查该病人是否存在特殊用药", mlng病人ID, mlng主页ID, mintInsure)
    '周玉强20100521增加了列表框显示列：记帐日期,单据号,数量,单价,金额,记帐人,规格
    With rsTemp
        Do While Not .EOF
            msfDetail.TextMatrix(.AbsolutePosition, 0) = ""
            msfDetail.TextMatrix(.AbsolutePosition, 1) = !特殊项目
            msfDetail.TextMatrix(.AbsolutePosition, 2) = !规格
            msfDetail.TextMatrix(.AbsolutePosition, 3) = Nvl(!限病种信息)
            msfDetail.TextMatrix(.AbsolutePosition, 4) = !记帐日期
            msfDetail.TextMatrix(.AbsolutePosition, 5) = !单据号
            msfDetail.TextMatrix(.AbsolutePosition, 6) = !数量
            msfDetail.TextMatrix(.AbsolutePosition, 7) = !单价
            msfDetail.TextMatrix(.AbsolutePosition, 8) = !金额
            msfDetail.TextMatrix(.AbsolutePosition, 9) = !记帐人
            

         '   msfDetail.TextMatrix(.AbsolutePosition, 0) = IIf(!报销 = 1, "√", "Х"),周玉强修改默认为空
                       '保存明细ID
            msfDetail.RowData(.AbsolutePosition) = !ID
            msfDetail.Rows = msfDetail.Rows + 1
            .MoveNext
        Loop
    End With
    Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    
     
   msfDetail.Rows = 2: msfDetail.Cols = 10
    msfDetail.TextMatrix(0, 0) = "打勾报销"
    msfDetail.TextMatrix(0, 1) = "特殊项目"
    msfDetail.TextMatrix(0, 2) = "规格"
    msfDetail.TextMatrix(0, 3) = "限病种信息"
    msfDetail.TextMatrix(0, 4) = "记帐日期"
    msfDetail.TextMatrix(0, 5) = "单据号"
    msfDetail.TextMatrix(0, 6) = "数量"
    msfDetail.TextMatrix(0, 7) = "单价"
    msfDetail.TextMatrix(0, 8) = "金额"
    msfDetail.TextMatrix(0, 9) = "记帐人"
    
    
    msfDetail.ColWidth(0) = 1000
    msfDetail.ColWidth(1) = 2400
    msfDetail.ColWidth(2) = 1500
    msfDetail.ColWidth(3) = 3000
    msfDetail.ColWidth(4) = 1000
    msfDetail.ColWidth(5) = 800
    msfDetail.ColWidth(6) = 800
    msfDetail.ColWidth(7) = 800
    msfDetail.ColWidth(8) = 1000
    msfDetail.ColWidth(9) = 800
   
   
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    msfDetail.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Picture1.ScaleHeight
End Sub

Private Sub msfDetail_DblClick()
    If msfDetail.TextMatrix(msfDetail.Row, 0) = "√" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = "Х"
    ElseIf msfDetail.TextMatrix(msfDetail.Row, 0) = "Х" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = ""
    ElseIf msfDetail.TextMatrix(msfDetail.Row, 0) = "" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = "√"
    End If
End Sub

Private Sub msfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    If msfDetail.TextMatrix(msfDetail.Row, 0) = "√" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = "Х"
    ElseIf msfDetail.TextMatrix(msfDetail.Row, 0) = "Х" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = ""
    ElseIf msfDetail.TextMatrix(msfDetail.Row, 0) = "" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = "√"
    End If
End Sub

Private Sub txt住院号_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
      
    msfDetail.Rows = 2: msfDetail.Cols = 10
    msfDetail.TextMatrix(0, 0) = "打勾报销"
    msfDetail.TextMatrix(0, 1) = "特殊项目"
    msfDetail.TextMatrix(0, 2) = "规格"
    msfDetail.TextMatrix(0, 3) = "限病种信息"
    msfDetail.TextMatrix(0, 4) = "记帐日期"
    msfDetail.TextMatrix(0, 5) = "单据号"
    msfDetail.TextMatrix(0, 6) = "数量"
    msfDetail.TextMatrix(0, 7) = "单价"
    msfDetail.TextMatrix(0, 8) = "金额"
    msfDetail.TextMatrix(0, 9) = "记帐人"
    
    
    msfDetail.ColWidth(0) = 1000
    msfDetail.ColWidth(1) = 2400
    msfDetail.ColWidth(2) = 1500
    msfDetail.ColWidth(3) = 3000
    msfDetail.ColWidth(4) = 1000
    msfDetail.ColWidth(5) = 800
    msfDetail.ColWidth(6) = 800
    msfDetail.ColWidth(7) = 800
    msfDetail.ColWidth(8) = 1000
    msfDetail.ColWidth(9) = 800
    msfDetail.ColAlignment(0) = 3
    msfDetail.ColAlignment(3) = 1
    msfDetail.ColAlignment(2) = 1
    msfDetail.ColAlignment(6) = 3
    
    If Trim(txt住院号.Text) = "" Then
        txt住院号.Tag = ""
        Exit Sub
    End If
    

    
    gstrSQL = " Select A.病人ID,A.住院次数 AS 主页ID,A.姓名,A.性别,B.名称 AS 科室 " & _
              " From 病人信息 A,部门表 B" & _
              " Where A.当前科室ID=B.ID(+) And A.住院号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人信息", CStr(txt住院号.Text))
    If rsTemp.RecordCount = 0 Or IsNull(rsTemp!主页ID) = True Then
        MsgBox "没有找到该病人！", vbInformation, gstrSysName
        txt住院号.Tag = ""
        txt住院号.SetFocus
        
        Exit Sub
    End If
    
    Me.lblNote.Caption = rsTemp!科室 & " " & rsTemp!姓名 & " " & rsTemp!性别
    Me.txt住院号.Tag = rsTemp!病人ID & "|" & rsTemp!主页ID
    mlng病人ID = rsTemp!病人ID
    mlng主页ID = rsTemp!主页ID
    
    Call RefreshData
    
    '如果人员性质为护士的则只能查看护士所在病区的人员
    cmd更新.Enabled = True
    gstrSQL = "select Count(1) from 人员性质说明 where 人员性质 = '护士' and 人员ID = [1]"
    If Val(zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID).Fields(0)) = 1 Then
        '人员性质为护士的需要检测权限
               
        gstrSQL = "select Count(1) from 病人信息 a , 部门人员 b where a.当前病区id = b.部门id And 病人id = [1] " & _
        " And 当前病区id in(select 部门ID FROM 部门人员 WHERE  人员ID=[2] )"
        If Val(zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, UserInfo.ID).Fields(0)) = 0 Then
            cmd更新.Enabled = False
            MsgBox "此病人已出院或不属于本科室，请核对：如果已出院，请撤消出院后再审批，如转科，请联系医保科！", vbInformation, gstrSysName
        End If
    End If

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd日志_Click()
On Error GoTo ErrH
    If mlng病人ID = 0 Then Exit Sub
    With frm医保操作日志
        .str模块 = "医保工具"
        .str功能 = "特殊药品收费"
        .str主键1 = mlng病人ID
        .str主键2 = mlng主页ID
        .Show vbModal
    End With
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

