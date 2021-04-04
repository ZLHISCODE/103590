VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckNumber 
   Caption         =   "验收单导入"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15855
   Icon            =   "frmCheckNumber.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   15855
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd全清 
      Cancel          =   -1  'True
      Caption         =   "全清(&B)"
      Height          =   350
      Left            =   11520
      TabIndex        =   23
      Top             =   7335
      Width           =   1100
   End
   Begin VB.CommandButton cmd全选 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   10200
      TabIndex        =   22
      Top             =   7335
      Width           =   1100
   End
   Begin VB.PictureBox pic已导入过颜色 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   7440
      Width           =   260
   End
   Begin VB.Frame Frmline2 
      Height          =   120
      Left            =   120
      TabIndex        =   16
      Top             =   6960
      Width           =   15735
   End
   Begin VB.Frame fraFilter 
      Caption         =   " 提取数据条件"
      Height          =   5775
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2415
      Begin VB.TextBox txt验收NO 
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Width           =   2205
      End
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   2205
      End
      Begin VB.ComboBox cbo验收日期 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1920
         Width           =   2205
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "提取数据"
         Height          =   300
         Left            =   1110
         TabIndex        =   7
         Top             =   5280
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   216334339
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   216334339
         CurrentDate     =   36263
      End
      Begin VB.Label lbl提示 
         Caption         =   "(不填NO号则忽略NO号过滤)"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label lbl验收NO 
         Caption         =   "验收NO"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label lbl结束时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label lbl开始时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   2505
         Width           =   720
      End
      Begin VB.Label lbl验收时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验收时间"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label lblStore 
         AutoSize        =   -1  'True
         Caption         =   "验收库房"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   14640
      TabIndex        =   2
      Top             =   7335
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "生成入库单据(&O)"
      Height          =   350
      Left            =   12960
      TabIndex        =   1
      Top             =   7335
      Width           =   1575
   End
   Begin VB.Frame Frmline1 
      Height          =   120
      Left            =   120
      TabIndex        =   0
      Top             =   645
      Width           =   15735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   5655
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   13140
      _cx             =   23177
      _cy             =   9975
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
      Cols            =   25
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCheckNumber.frx":6852
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
      Editable        =   2
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
   Begin VB.Label lbl已导入过颜色 
      AutoSize        =   -1  'True
      Caption         =   "已产生过外购入库单"
      Height          =   180
      Left            =   5760
      TabIndex        =   18
      Top             =   7470
      Width           =   1620
   End
   Begin VB.Label lbl产生单据数 
      AutoSize        =   -1  'True
      Caption         =   "提示：共会产生0张入库单据"
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
      TabIndex        =   5
      Top             =   7410
      Width           =   2625
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   255
      Picture         =   "frmCheckNumber.frx":6C18
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInfor 
      Caption         =   " 说明：跟据提取条件，过滤出合格的验收单据，生成药品外购入库单。"
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   10695
   End
End
Attribute VB_Name = "frmCheckNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const m已导入过ColColor As Long = &H8080FF
Private Const m未导入过ColColor As Long = &H0&

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Private mintListIndex As Integer    '库房索引

Private mfrmMain As Form

Public Sub ShowCard(FrmMain As Form, ByVal intListIndex As Integer)

    mintListIndex = intListIndex
    Set mfrmMain = FrmMain
    
    Me.Show vbModal, FrmMain
End Sub

Private Sub cbo库房_Click()
    If Val(Cbo库房.ListIndex) <> Val(Cbo库房.Tag) And vsfList.Rows > 1 Then
        If MsgBox("如果改变库房，需要重新提取单据内容，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            vsfList.Rows = 1
            lbl产生单据数.Caption = "提示：界面中共可能会产生0张入库单据"
        Else
            Cbo库房.ListIndex = Val(Cbo库房.Tag)
        End If
    End If
    Cbo库房.Tag = Val(Cbo库房.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
'填充表格数据
    Dim rsTemp As ADODB.Recordset
    Dim lng库房ID As Long
    Dim str供应商id As String
    Dim int单据数量 As Integer
    
    On Error GoTo errHandle

    If Cbo库房.ListIndex = -1 Then
        MsgBox "请选择库房！", vbInformation, gstrSysName
        Exit Sub
    End If

    vsfList.Rows = 1
    '库房id
    lng库房ID = Val(Cbo库房.ItemData(Cbo库房.ListIndex))
    
    gstrSQL = "Select Distinct f.Id As 验收id, f.库房id,f.供药单位id,c.药品id,a.导入标记,g.名称 As 供药单位,f.No,f.验收人,f.验收日期," & vbNewLine & _
                    "                Nvl(f.是否合格,0) As 是否合格,b.名称,b.编码, b.规格, a.产地 As 生产商,a.批号,a.进药数量,b.计算单位," & vbNewLine & _
                    "                a.成本价, a.零售价, a.生产日期, a.效期, a.批准文号, a.验收结论,b.是否变价" & vbNewLine & _
                    "From 药品验收明细 A, 收费项目目录 B, 药品规格 C, 药品特性 D, 药品剂型 E, 药品验收记录 F, 供应商 G" & vbNewLine & _
                    "Where a.药品id = b.Id And b.Id = c.药品id And c.药名id = d.药名id And d.药品剂型 = e.名称(+) And f.Id = a.验收id And f.是否合格 = 0 And" & vbNewLine & _
                    "      f.供药单位id = g.Id(+) And f.库房id = [1] And f.验收日期 Between [2] And [3]" & IIf(Trim(txt验收NO.Text) <> "", " and f.No=[4]", "") & vbNewLine & _
                    "Order By a.导入标记 Desc, f.No"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "药品验收入库", lng库房ID, _
    CDate(Format(dtp开始时间.Value, "yyyy-mm-dd")), CDate(Format(dtp结束时间.Value, "yyyy-mm-dd") & " 23:59:59"), Trim(txt验收NO.Text))

   If rsTemp.RecordCount = 0 Then MsgBox "没有查询到合格的验收单据，请检查！", vbInformation, gstrSysName: Exit Sub
   
    With vsfList
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("验收id")) = rsTemp!验收id
            .TextMatrix(.Rows - 1, .ColIndex("库房id")) = rsTemp!库房id
            .TextMatrix(.Rows - 1, .ColIndex("供应商id")) = rsTemp!供药单位ID
            .TextMatrix(.Rows - 1, .ColIndex("药品id")) = rsTemp!药品id
            .TextMatrix(.Rows - 1, .ColIndex("供药单位")) = NVL(rsTemp!供药单位)
            .TextMatrix(.Rows - 1, .ColIndex("验收NO")) = rsTemp!NO
            .TextMatrix(.Rows - 1, .ColIndex("验收人")) = NVL(rsTemp!验收人)
            .TextMatrix(.Rows - 1, .ColIndex("验收日期")) = Format(NVL(rsTemp!验收日期), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("是否合格")) = IIf(NVL(rsTemp!是否合格, 0) = 0, "合格", "不合格")
            .TextMatrix(.Rows - 1, .ColIndex("药品")) = "[" & rsTemp!编码 & "]" & rsTemp!名称 & "(" & rsTemp!规格 & ")"
            .TextMatrix(.Rows - 1, .ColIndex("生产商")) = NVL(rsTemp!生产商)
            .TextMatrix(.Rows - 1, .ColIndex("进药数量")) = zlStr.FormatEx(NVL(rsTemp!进药数量, 0), mintNumberDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("单位")) = NVL(rsTemp!计算单位)
            .TextMatrix(.Rows - 1, .ColIndex("成本价")) = zlStr.FormatEx(NVL(rsTemp!成本价, 0), mintCostDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("成本金额")) = zlStr.FormatEx(NVL(rsTemp!成本价, 0) * NVL(rsTemp!进药数量, 0), mintMoneyDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("零售价")) = zlStr.FormatEx(NVL(rsTemp!零售价, 0), mintPriceDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("零售金额")) = zlStr.FormatEx(NVL(rsTemp!零售价, 0) * NVL(rsTemp!进药数量, 0), mintMoneyDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("生产日期")) = Format(NVL(rsTemp!生产日期), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("效期")) = Format(NVL(rsTemp!效期), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("批准文号")) = NVL(rsTemp!批准文号)
            .TextMatrix(.Rows - 1, .ColIndex("验收结论")) = NVL(rsTemp!验收结论)
            .TextMatrix(.Rows - 1, .ColIndex("批号")) = NVL(rsTemp!批号)
            .TextMatrix(.Rows - 1, .ColIndex("是否变价")) = NVL(rsTemp!是否变价, 0)
            
            If NVL(rsTemp!是否变价, 0) = 0 Then .TextMatrix(.Rows - 1, .ColIndex("定价零售价")) = NVL(rsTemp!零售价, 0)
            
            .Cell(flexcpForeColor, .Rows - 1, .ColIndex("供药单位"), .Rows - 1, .ColIndex("验收结论")) = IIf(NVL(rsTemp!导入标记, 0) = 0, m未导入过ColColor, m已导入过ColColor)
            
            If InStr(";" & str供应商id & ";", ";" & rsTemp!供药单位ID & ";") = 0 Then
                str供应商id = IIf(str供应商id = "", "", str供应商id & ";") & rsTemp!供药单位ID
                int单据数量 = int单据数量 + 1
            End If
            
            rsTemp.MoveNext
        Loop
    End With
    
    lbl产生单据数.Caption = "提示：界面中共可能会产生" & int单据数量 & "张入库单据"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd全清_Click()
    Dim i As Integer
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("选择")) = 0
        Next
    End With
End Sub

Private Sub cmd全选_Click()
    Dim i As Integer
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("选择")) = 1
        Next
    End With
End Sub

Private Sub Form_Load()
    vsfList.AllowSelection = False '不能多选
    vsfList.Rows = 1
    Call initComboBox
    Call SetMedicalWH
    Call GetDrugDigit(Cbo库房.ItemData(Cbo库房.ListIndex), "药品验收管理", 4, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    lbl产生单据数.Caption = "提示：界面中共可能会产生0张入库单据"
End Sub

Private Sub SetMedicalWH()
    Dim i As Integer

    With mfrmMain.cboStock
        Cbo库房.Clear
        For i = 0 To .ListCount - 1
            Cbo库房.AddItem .List(i)
            Cbo库房.ItemData(Cbo库房.NewIndex) = .ItemData(i)
        Next
        Cbo库房.ListIndex = .ListIndex
    End With
        
    Cbo库房.Tag = IIf(mintListIndex = -1, 0, mintListIndex)
    Cbo库房.ListIndex = IIf(mintListIndex = -1, 0, mintListIndex)
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Height < 7500 Then Me.Height = 7500
    If Me.Width < 15000 Then Me.Width = 15000
    
    Frmline1.Left = 0
    Frmline1.Top = Me.ScaleHeight / 12
    Frmline1.Width = Me.ScaleWidth

    Frmline2.Left = 0
    Frmline2.Top = Me.ScaleHeight * 43 / 48
    Frmline2.Width = Me.ScaleWidth

    fraFilter.Left = 50
    fraFilter.Top = Frmline1.Top + 200
    fraFilter.Height = Frmline2.Top - Frmline1.Top - 300
    
    cmdFilter.Left = 1200
    cmdFilter.Top = fraFilter.Height - cmdFilter.Height - 200
    
    vsfList.Left = fraFilter.Width + 100
    vsfList.Top = fraFilter.Top + 100
    vsfList.Width = Me.ScaleWidth - fraFilter.Left - fraFilter.Width - 150
    vsfList.Height = fraFilter.Height - 100
    
    lbl产生单据数.Left = 100
    lbl产生单据数.Top = Frmline2.Top + 350
    
    CmdCancel.Left = Me.Width - CmdCancel.Width - 200
    CmdCancel.Top = lbl产生单据数.Top - 50
    
    cmdOK.Left = CmdCancel.Left - cmdOK.Width - 50
    cmdOK.Top = lbl产生单据数.Top - 50
    
    cmd全清.Top = cmdOK.Top
    cmd全清.Left = cmdOK.Left - cmd全清.Width - 50
    
    cmd全选.Top = cmdOK.Top
    cmd全选.Left = cmd全清.Left - cmd全选.Width - 50
    
    lbl已导入过颜色.Top = lbl产生单据数.Top + 30
    pic已导入过颜色.Top = lbl产生单据数.Top

    
End Sub


Private Sub cmdOK_Click()
    Dim i As Integer
    Dim int单据数量 As Integer
    Dim strNo As String
    Dim str供应商id As String
    Dim strDate As String
    Dim int序号 As Integer
    Dim lng库房ID As Long
    Dim blnTrans As Boolean
    Dim arrSql As Variant
    Dim blnOK As Boolean
    Dim rsSort As New ADODB.Recordset   '按供应商排序
    Dim intRow As Integer
    
    If vsfList.Rows < 2 Then Exit Sub
    If Cbo库房.ItemData(Cbo库房.ListIndex) = -1 Then
        MsgBox "请选择库房！", vbInformation, gstrSysName
        Exit Sub
    End If

    On Error GoTo ErrHand

    If MsgBox("是否确定生成入库单？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    With rsSort
        .Fields.Append "行号", adDouble, 18, adFldIsNullable
        .Fields.Append "供应商id", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    With vsfList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("选择")) Like "*1" Then
                rsSort.AddNew
                rsSort!行号 = i
                rsSort!供应商id = Val(.TextMatrix(i, .ColIndex("供应商id")))
                
                rsSort.Update
            End If
        Next
    End With
    
    If rsSort.RecordCount = 0 Then
        MsgBox "没有选择要导入的单据，请至少勾选一个数据后再保存！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    rsSort.Sort = "供应商id,行号"
        
    arrSql = Array()
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    lng库房ID = Val(Cbo库房.ItemData(Cbo库房.ListIndex))
    
    With vsfList
        For i = 1 To rsSort.RecordCount
            intRow = rsSort!行号

            If InStr(";" & str供应商id & ";", ";" & .TextMatrix(intRow, .ColIndex("供应商id")) & ";") = 0 Then
                str供应商id = IIf(str供应商id = "", "", str供应商id & ";") & .TextMatrix(intRow, .ColIndex("供应商id"))
                strNo = zlDatabase.GetNextNo(21, lng库房ID)
                int单据数量 = int单据数量 + 1
                int序号 = 0
            End If

            gstrSQL = "Zl_药品验收明细_导入标记(" & Val(.TextMatrix(intRow, .ColIndex("验收id"))) & "," & Val(.TextMatrix(intRow, .ColIndex("药品id"))) & ")"
                
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            int序号 = int序号 + 1
            
            gstrSQL = "zl_药品外购_INSERT("
            'NO
            gstrSQL = gstrSQL & "'" & strNo & "'"
            '序号
            gstrSQL = gstrSQL & "," & int序号
            '库房ID
            gstrSQL = gstrSQL & "," & lng库房ID
            '对方部门ID
            gstrSQL = gstrSQL & ",NULL"
            '供药单位ID
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("供应商id")))
            '药品ID
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("药品ID")))
            '产地
            gstrSQL = gstrSQL & ",'" & .TextMatrix(intRow, .ColIndex("生产商")) & "'"
            '批号
            gstrSQL = gstrSQL & ",'" & .TextMatrix(intRow, .ColIndex("批号")) & "'"
            '效期
            gstrSQL = gstrSQL & "," & "to_date('" & .TextMatrix(intRow, .ColIndex("效期")) & "','yyyy-mm-dd')"
            '实际数量
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("进药数量")))
            '成本价
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("成本价")))
            '成本金额
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("成本金额")))
            '扣率
            gstrSQL = gstrSQL & "," & 100
            '零售价
            If .TextMatrix(intRow, .ColIndex("是否变价")) = 0 Then
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("定价零售价")))
            Else
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("零售价")))
            End If
            '零售金额
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("零售金额")))
            '差价
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("零售金额"))) - Val(.TextMatrix(intRow, .ColIndex("成本金额")))
            '摘要
            gstrSQL = gstrSQL & ",'药品验收管理导入'"
            '填制人
            gstrSQL = gstrSQL & ",'" & UserInfo.用户姓名 & "'"
            '发票号
            gstrSQL = gstrSQL & ",NULL"
            '发票日期
            gstrSQL = gstrSQL & ",NULL"
            '发票金额
            gstrSQL = gstrSQL & ",NULL"
            '填制日期
            gstrSQL = gstrSQL & "," & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS')"
            '外观
            gstrSQL = gstrSQL & ",NULL"
            '产品合格证
            gstrSQL = gstrSQL & ",NULL"
            '核查人
            gstrSQL = gstrSQL & ",NULL"
            '核查日期
            gstrSQL = gstrSQL & ",NULL"
            '批次
            gstrSQL = gstrSQL & "," & 0
            '是否退货
            gstrSQL = gstrSQL & "," & 1
            '生产日期
            gstrSQL = gstrSQL & "," & "to_date('" & .TextMatrix(intRow, .ColIndex("生产日期")) & "','yyyy-mm-dd')"
            '批准文号
            gstrSQL = gstrSQL & ",'" & .TextMatrix(intRow, .ColIndex("批准文号")) & "'"
            '随货单号
            gstrSQL = gstrSQL & ",NULL"
            '金额差
            gstrSQL = gstrSQL & "," & 0
            '加成率
            If Val(.TextMatrix(intRow, .ColIndex("成本价"))) = 0 Then
                gstrSQL = gstrSQL & "," & 0
            Else
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("零售价"))) / Val(.TextMatrix(intRow, .ColIndex("成本价"))) - 1
            End If
            '发票代码
            gstrSQL = gstrSQL & ",NULL"
            '计划id
            gstrSQL = gstrSQL & ",NULL"
            '财务审核
            gstrSQL = gstrSQL & "," & 0
            '原产地
            gstrSQL = gstrSQL & ",NULL"
            '随货日期
            gstrSQL = gstrSQL & ",NULL"
            '验收结论
            gstrSQL = gstrSQL & ",'" & .TextMatrix(intRow, .ColIndex("验收结论")) & "'"
            
            gstrSQL = gstrSQL & ")"
                
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            rsSort.MoveNext
        Next
    
    End With
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "保存外购入库单")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    MsgBox "本次共产生" & int单据数量 & "张外购入库单，请注意查看！", vbInformation, gstrSysName
    vsfList.Rows = 1
    
    lbl产生单据数.Caption = "提示：界面中共可能会产生0张入库单据"
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub initComboBox()
    With cbo验收日期
        .Clear
        .AddItem "今日"
        .AddItem "一星期内"
        .AddItem "一个月内"
        .AddItem "三个月内"
        .AddItem "自定义日期"
        .ListIndex = 0
    End With
End Sub
Private Sub cbo验收日期_Click()
    Dim dateCurrentDate As Date
    
    If cbo验收日期.Text = "自定义日期" Then
        dtp开始时间.Enabled = True
        dtp结束时间.Enabled = True
        
    Else
        dtp开始时间.Enabled = False
        dtp结束时间.Enabled = False
    End If
    
    '根据选择改变时间
    dateCurrentDate = sys.Currentdate
    Select Case cbo验收日期.ListIndex
        Case 0
            dtp开始时间.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
            dtp结束时间.Value = dateCurrentDate
        Case 1
            dtp开始时间.Value = CDate(Format(DateAdd("d", -7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp结束时间.Value = dateCurrentDate
        Case 2
            dtp开始时间.Value = CDate(Format(DateAdd("d", -30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp结束时间.Value = dateCurrentDate
        Case 3
            dtp开始时间.Value = CDate(Format(DateAdd("d", -90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp结束时间.Value = dateCurrentDate
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
End Sub

Private Sub txt验收NO_GotFocus()
    Me.txt验收NO.SelStart = 0: Me.txt验收NO.SelLength = 100
End Sub
Private Sub txt验收NO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub
Private Sub vsfList_EnterCell()
    If vsfList.Col = vsfList.ColIndex("选择") Then
        vsfList.Editable = flexEDKbdMouse
    Else
        vsfList.Editable = flexEDNone
    End If
End Sub
Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode <> vbKeyDelete Then Exit Sub
    
    With vsfList
        If .Rows < 2 Then Exit Sub
        If MsgBox("是否继续删除验收NO为 " & .TextMatrix(.Row, .ColIndex("验收NO")) & "，药品为 " & .TextMatrix(.Row, .ColIndex("药品")) & " 的数据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        .RemoveItem .Row
        
    End With
End Sub
