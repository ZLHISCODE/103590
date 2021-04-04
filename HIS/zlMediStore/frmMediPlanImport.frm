VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmMediPlanImport 
   Caption         =   "药品计划单批量导入"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   Icon            =   "frmMediPlanImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   12240
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6840
      ScaleHeight     =   255
      ScaleWidth      =   3855
      TabIndex        =   16
      Top             =   6120
      Width           =   3855
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "正常"
         Height          =   180
         Left            =   1680
         TabIndex        =   20
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "已停用"
         Height          =   180
         Left            =   2640
         TabIndex        =   19
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.PictureBox picCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   2640
      Width           =   4815
      Begin VB.ComboBox cbo指定 
         Height          =   300
         Left            =   2730
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lbl库房 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "指定导入库房"
         Height          =   180
         Left            =   1560
         TabIndex        =   22
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label lblDetail 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "计划内容"
         Height          =   180
         Left            =   100
         TabIndex        =   11
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   840
      ScaleHeight     =   1695
      ScaleWidth      =   5655
      TabIndex        =   10
      Top             =   600
      Width           =   5655
      Begin VB.CommandButton cmdGetData 
         Caption         =   "获取数据(&G)"
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox chkZeroInput 
         Caption         =   "允许0计划数量显示"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CommandButton cmdUnchoose 
         Caption         =   "全不选(&U)"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdChoose 
         Caption         =   "全选(&A)"
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   975
         Left            =   0
         TabIndex        =   0
         Top             =   120
         Width           =   1335
         _cx             =   2355
         _cy             =   1720
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   840
      ScaleHeight     =   2535
      ScaleWidth      =   7095
      TabIndex        =   8
      Top             =   3000
      Width           =   7095
      Begin VB.PictureBox picOperation 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   6135
         TabIndex        =   9
         Top             =   1920
         Width           =   6135
         Begin VB.CheckBox chk允许导入停用药品 
            Caption         =   "允许导入停用药品"
            Height          =   180
            Left            =   360
            TabIndex        =   14
            Top             =   147
            Width           =   1815
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "取消(&C)"
            Height          =   375
            Left            =   3480
            TabIndex        =   6
            Top             =   50
            Width           =   1095
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "导入(&I)"
            Height          =   375
            Left            =   2280
            TabIndex        =   5
            Top             =   50
            Width           =   1095
         End
         Begin VB.Label lblInfo 
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   495
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1215
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "执行数量和计划数量字体加粗表示执行数量大于了计划数量及该计划单已经产生了入库单并已审核"
         Top             =   120
         Width           =   2655
         _cx             =   4683
         _cy             =   2143
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   6435
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediPlanImport.frx":014A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16510
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpView 
      Left            =   8280
      Top             =   2520
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMediPlanImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INT_WIDTH = 50
Private mrsCBO_Provider As ADODB.Recordset
Private mlngID As Long
Private mlngSum As Long '记录停用药品数量
Private mstrMsg As String '不允许导入停用药品，有停用药品时的提示信息

Private mstrPrivs As String
Private mlng库房ID As Long

'从参数表中取药品价格、数量、金额小数位数（计算精度）
Public mintCostDigit As Integer        '成本价小数位数
Public mintPriceDigit As Integer       '售价小数位数
Public mintNumberDigit As Integer      '数量小数位数
Public mintMoneyDigit As Integer       '金额小数位数
Private mintUnit As Integer             '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mblnLoad As Boolean

Private Sub cbo指定_Click()
    Dim i As Integer
    Dim blncheck库房 As Boolean, bln记录 As Boolean

    '取对应库房的精度
    Call GetDrugDigit(Val(cbo指定.ItemData(cbo指定.ListIndex)), "", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
     '未加载过数据退出
    If Not mblnLoad Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    With vsfDetail
        
        For i = 1 To .rows - 1
            '改变导入库房的id
            If .TextMatrix(i, .ColIndex("whid")) <> "" Then .TextMatrix(i, .ColIndex("whid")) = Val(cbo指定.ItemData(cbo指定.ListIndex))
            
            '对应库房的精度
            .ColFormat(.ColIndex("planqty")) = "#0." & String(mintNumberDigit, "0")
            .ColFormat(.ColIndex("execqty")) = "#0." & String(mintNumberDigit, "0")
            .ColFormat(.ColIndex("costprice")) = "#0." & String(mintCostDigit, "0")
            .ColFormat(.ColIndex("cost")) = "#0." & String(mintMoneyDigit, "0")
            .ColFormat(.ColIndex("sale")) = "#0." & String(mintMoneyDigit, "0")
            .ColFormat(.ColIndex("saleprice")) = "#0." & String(mintPriceDigit, "0")
            
            Get药品分批属性 i
            
            '判断药品是否设置存储库房在该导入库房中
            blncheck库房 = Check库房(Val(.TextMatrix(i, .ColIndex("id"))), Val(cbo指定.ItemData(cbo指定.ListIndex)))
            If blncheck库房 = False Then
                If Trim(.TextMatrix(i, .ColIndex("id"))) <> "" Then '只改变有药品的行
                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &HFFC0C0
                    .TextMatrix(i, .ColIndex("choose")) = "0"
                    
                    bln记录 = True
                End If
            Else
                .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbBack
                .TextMatrix(i, .ColIndex("choose")) = "-1"
                
                vsfDetail_AfterEdit i, .Row
            End If
            
            If bln记录 Then
                lblInfo.Caption = "紫色背景是该药品在导入库房中未设置存储性质，将不产生相应外购单据！"
                lblInfo.ForeColor = vbRed
                lblInfo.Visible = True
            Else
                lblInfo.Visible = False
            End If
            
            '判断是否停用，停用显示未
            If 是否停用(Val(.TextMatrix(i, .ColIndex("id")))) Then
                .Cell(flexcpForeColor, i, .ColIndex("choose"), i, .Cols - 1) = &HFF00FF
            End If
            
        Next
    End With
    
    staThis.Panels(2).Text = CountBuilds
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkZeroInput_Click()
    Dim i As Integer
    
    With vsfMain
        vsfDetail.rows = 1

        For i = 1 To .rows - 1
            Call vsfMain_AfterEdit(i, .ColIndex("choose"))
        Next
    End With
End Sub

Private Sub chk允许导入停用药品_Click()
    If Not mblnLoad Then Exit Sub
    staThis.Panels(2).Text = CountBuilds
End Sub

Private Sub cmdCancel_Click()
    Unload frmMediPlanGetData
    Unload Me
End Sub

Private Sub cmdChoose_Click()
    If vsfMain.rows > 1 Then
        Dim i As Integer
        Dim blnCancel As Boolean
        Screen.MousePointer = vbHourglass
        vsfDetail.rows = 1
        For i = 1 To vsfMain.rows - 1
            vsfMain.TextMatrix(i, vsfMain.ColIndex("choose")) = "-1"
            vsfMain_BeforeEdit i, vsfMain.ColIndex("choose"), blnCancel
            If blnCancel = False Then
                vsfMain_AfterEdit i, vsfMain.ColIndex("choose")
            Else
                vsfMain.TextMatrix(i, vsfMain.ColIndex("choose")) = "0"
            End If
            DoEvents
        Next
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdGetData_Click()
'获取计划单数据
    Dim strsql As String, strWhere As String
    Dim rsTmp As ADODB.Recordset, rsTmp1 As ADODB.Recordset
    
    On Error GoTo errHandle
    frmMediPlanGetData.Show vbModal, Me
    If frmMediPlanGetData.SQLWhere = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    DoEvents
    strWhere = frmMediPlanGetData.SQLWhere
    
    
    strsql = "select a.ID,a.NO,a.期间,a.库房ID,b.名称 库房,a.编制说明,a.编制人,a.编制日期,a.审核人,a.审核日期 " _
           & "from 药品采购计划 a, 部门表 b " _
           & "where a.库房id=b.id(+) and a.审核日期 is not null "
               
    strsql = strsql & strWhere & " order by a.NO"
    
    vsfDetail.rows = 1
    vsfMain.rows = 1
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
    If Not rsTmp.EOF Then
        '装载计划单数据
        DataLoading 1, rsTmp
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbDefault
        'MsgBox "未获取到数据！", , gstrSysName
    End If
    
    vsfMain.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdImport_Click()
    If vsfDetail.rows <= 1 Then Exit Sub
    If MsgBox("请确定要批量导入'药品计划单'数据？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Dim strInsert As String, strNo As String
    Dim i As Integer, intSN As Integer
    Dim dateNO As Date
    Dim lngWHID As Long, lngPID As Long
    Dim rsTmp As New ADODB.Recordset
    Dim int摘要长度 As Integer
    Dim str摘要 As String
    Dim col摘要 As New Collection
    
    Screen.MousePointer = vbHourglass
    On Error GoTo errSoft
    
    int摘要长度 = Sys.FieldsLength("药品收发记录", "摘要")
    
    '数据检查,如果分批则必须录入产地
    With vsfDetail
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("choose")) = "-1" And .Cell(flexcpBackColor, i, 1, i, 2) <> &HFFC0C0 And Val(.TextMatrix(i, .ColIndex("id"))) <> 0 And Val(.TextMatrix(i, .ColIndex("batch"))) = 1 And Trim(.TextMatrix(i, .ColIndex("producer"))) = "" Then
                MsgBox "第" & i & "行药品在此库房是分批性质，请对此药品设置上次生产商！"
                
                .SetFocus
                .Row = i
                .MsfObj.TopRow = i
                .Col = .ColIndex("producer")
                Exit Sub
            End If
        Next
    End With
    With rsTmp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        '.Fields.Append "序号", adInteger, , adFldIsNullable
        .Fields.Append "库房ID", adBigInt, , adFldIsNullable
        .Fields.Append "供药单位ID", adBigInt, , adFldIsNullable
        .Fields.Append "药品ID", adBigInt, , adFldIsNullable
        .Fields.Append "生产商", adVarChar, 60, adFldIsNullable
        .Fields.Append "生产日期", adDBDate, , adFldIsNullable
        .Fields.Append "效期", adDBDate, , adFldIsNullable
        .Fields.Append "实际数量", adDouble, , adFldIsNullable
        .Fields.Append "成本价", adDouble, , adFldIsNullable
        .Fields.Append "成本金额", adDouble, , adFldIsNullable
        .Fields.Append "零售价", adDouble, , adFldIsNullable
        .Fields.Append "零售金额", adDouble, , adFldIsNullable
        .Fields.Append "差价", adDouble, , adFldIsNullable
        .Fields.Append "加成率", adDouble, , adFldIsNullable
        .Fields.Append "药库包装", adDouble, , adFldIsNullable
        .Fields.Append "计划ID", adBigInt, , adFldIsNullable
        .Fields.Append "批准文号", adVarChar, 40, adFldIsNullable
        .Fields.Append "摘要", adLongVarChar, int摘要长度, adFldIsNullable
        .Open
    End With
    
    With vsfDetail
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("choose")) = "-1" And .Cell(flexcpBackColor, i, 1, i, 2) <> &HFFC0C0 And 是否导入(i) Then
                rsTmp.AddNew
                rsTmp!库房id = .TextMatrix(i, .ColIndex("whid"))
                rsTmp!供药单位ID = GetProviderID(.TextMatrix(i, .ColIndex("provider")))
                rsTmp!药品id = .TextMatrix(i, .ColIndex("id"))
                rsTmp!生产商 = .TextMatrix(i, .ColIndex("producer"))
                rsTmp!生产日期 = IIf(.TextMatrix(i, .ColIndex("pdate")) = "", Null, .TextMatrix(i, .ColIndex("pdate")))
                If Not IsNull(rsTmp!生产日期) Then
                    If gtype_UserSysParms.P149_效期显示方式 = 1 Then
                        rsTmp!效期 = DateAdd("d", 1, DateAdd("m", .TextMatrix(i, .ColIndex("avail_day")), rsTmp!生产日期))
                    Else
                        rsTmp!效期 = DateAdd("m", .TextMatrix(i, .ColIndex("avail_day")), rsTmp!生产日期)
                    End If
                End If
                rsTmp!实际数量 = .TextMatrix(i, .ColIndex("planqty"))
                rsTmp!成本价 = .TextMatrix(i, .ColIndex("costprice"))
                rsTmp!成本金额 = .TextMatrix(i, .ColIndex("cost"))
                rsTmp!零售价 = .TextMatrix(i, .ColIndex("saleprice"))
                rsTmp!零售金额 = .TextMatrix(i, .ColIndex("sale"))
                rsTmp!差价 = IIf(IsNull(rsTmp!零售金额), 0, rsTmp!零售金额) - IIf(IsNull(rsTmp!成本金额), 0, rsTmp!成本金额)
                rsTmp!加成率 = Val(.TextMatrix(i, .ColIndex("add_rate")))
                rsTmp!药库包装 = .TextMatrix(i, .ColIndex("store_pak"))
                rsTmp!计划id = .TextMatrix(i, .ColIndex("planid"))
                rsTmp!批准文号 = .TextMatrix(i, .ColIndex("approval"))
                
                '合并摘要，同一个供应商的摘要如果不同则进行汇总（用;分隔）
                If Trim(.TextMatrix(i, .ColIndex("Summary"))) <> "" Then
                    If ExistsColObject(col摘要, "_" & Val(rsTmp!供药单位ID)) = False Then
                        '集合没找到元素则新增该元素
                        col摘要.Add Trim(.TextMatrix(i, .ColIndex("Summary"))), "_" & Val(rsTmp!供药单位ID)
                    Else
                        '集合找到元素，则在原来值的基础上进行汇总
                        str摘要 = col摘要("_" & Val(rsTmp!供药单位ID))
                        If str摘要 = "" Then
                            str摘要 = Trim(.TextMatrix(i, .ColIndex("Summary")))
                        ElseIf InStr(1, ";" & str摘要 & ";", ";" & Trim(.TextMatrix(i, .ColIndex("Summary"))) & ";") = 0 Then
                            If LenB(StrConv(str摘要 & ";" & Trim(.TextMatrix(i, .ColIndex("Summary"))), vbFromUnicode)) <= int摘要长度 Then
                                str摘要 = str摘要 & ";" & Trim(.TextMatrix(i, .ColIndex("Summary")))
                            End If
                        End If
            
                        col摘要.Remove "_" & Val(rsTmp!供药单位ID)
                        col摘要.Add str摘要, "_" & Val(rsTmp!供药单位ID)
                    End If
                End If
                
                rsTmp.Update
            End If
        Next
        
        '合并摘要
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If ExistsColObject(col摘要, "_" & Val(rsTmp!供药单位ID)) = True Then
                str摘要 = col摘要("_" & Val(rsTmp!供药单位ID))
                rsTmp!摘要 = str摘要
            Else
                rsTmp!摘要 = ""
            End If
            
            rsTmp.Update
            rsTmp.MoveNext
        Loop
        
        rsTmp.MoveFirst
        
        rsTmp.Sort = "库房ID,供药单位ID"
    End With
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
            
    With rsTmp
        dateNO = Sys.Currentdate
        .MoveFirst
        Do While Not .EOF
            If lngWHID <> rsTmp!库房id Or lngPID <> rsTmp!供药单位ID Then
                lngWHID = rsTmp!库房id
                lngPID = rsTmp!供药单位ID
                intSN = 1
                strNo = Sys.GetNextNo(21, rsTmp!库房id)
            End If
            '执行存储过程，提交数据
            strInsert = "zl_药品外购_INSERT("
            'NO
            strInsert = strInsert & "'" & strNo & "'"
            '序号
            strInsert = strInsert & "," & intSN
            '库房ID
            strInsert = strInsert & "," & rsTmp!库房id
            '对方部门ID
            strInsert = strInsert & ",null"
            strInsert = strInsert & "," & rsTmp!供药单位ID
            strInsert = strInsert & "," & rsTmp!药品id
            strInsert = strInsert & ",'" & rsTmp!生产商 & "'"
            '批号
            strInsert = strInsert & ",'1'"
            '效期
            strInsert = strInsert & "," & IIf(IsNull(rsTmp!效期), "null", "to_date('" & Format(rsTmp!效期, "yyyy-mm-dd") & "', 'yyyy-mm-dd')")
            '实际数量
            strInsert = strInsert & "," & Round(rsTmp!实际数量 * rsTmp!药库包装, mintNumberDigit)
            '成本价
            strInsert = strInsert & "," & Round(rsTmp!成本价 / rsTmp!药库包装, mintCostDigit)
            '成本金额
            strInsert = strInsert & "," & Round(rsTmp!成本金额, mintMoneyDigit)
            '扣率
            strInsert = strInsert & ",100"
            '零售价
            strInsert = strInsert & "," & Round(rsTmp!零售价 / rsTmp!药库包装, mintPriceDigit)
            strInsert = strInsert & "," & Round(rsTmp!零售金额, mintMoneyDigit)
            strInsert = strInsert & "," & Round(rsTmp!差价, mintMoneyDigit)
            '摘要
            strInsert = strInsert & ",'" & rsTmp!摘要 & "'"
            '填制人
            strInsert = strInsert & ",'" & UserInfo.用户姓名 & "'"
            '发票号
            strInsert = strInsert & ",null"
            '发票日期
            strInsert = strInsert & ",null"
            '发票金额
            strInsert = strInsert & ",Null"
            '填制日期
            strInsert = strInsert & ",to_date('" & dateNO & "','yyyy-mm-dd HH24:MI:SS')"
            '外观
            strInsert = strInsert & ",Null"
            '产品合格证
            strInsert = strInsert & ",Null"
            '核查人
            strInsert = strInsert & ",Null"
            '核查日期
            strInsert = strInsert & ",Null"
            '批次
            strInsert = strInsert & ",Null"
            '是否退货
            strInsert = strInsert & ",1"
            '生产日期
            strInsert = strInsert & "," & IIf(IsNull(rsTmp!生产日期), "null", "to_date('" & Format(rsTmp!生产日期, "yyyy-mm-dd") & "', 'yyyy-mm-dd')")
            '批准文号
            strInsert = strInsert & "," & IIf(IsNull(rsTmp!批准文号), "Null", "'" & rsTmp!批准文号 & "'")
            '随货单号
            strInsert = strInsert & ",Null"
            '金额差
            strInsert = strInsert & ",Null"
            '加成率
            strInsert = strInsert & "," & IIf(IsNull(rsTmp!加成率), "null", rsTmp!加成率)
            '发票代码
            strInsert = strInsert & ",Null"
            '计划id
            strInsert = strInsert & "," & rsTmp!计划id
            strInsert = strInsert & ")"
            
            zlDatabase.ExecuteProcedure strInsert, Me.Caption
            
            intSN = intSN + 1
            lngWHID = rsTmp!库房id
            lngPID = rsTmp!供药单位ID
            .MoveNext
        Loop
        gcnOracle.CommitTrans
    End With
    
    '提示信息
    If mlngSum > 0 Then
        MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "个药品已停用，这部分药品将不导入外购入库单中！", "等" & mlngSum & "个药品已停用，这部分药品将不导入外购入库单中！"), vbInformation, gstrSysName
        
        mlngSum = 0
        mstrMsg = ""
    End If
    
    '导入成功后，清理界面
    vsfMain.rows = 1
    vsfDetail.rows = 1
    rsTmp.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub
    
errSoft:
    Screen.MousePointer = vbDefault
    Call SaveErrLog
    Exit Sub

errHandle:
    gcnOracle.RollbackTrans
    'If ErrCenter() = 1 Then Resume
    Screen.MousePointer = vbDefault
    rsTmp.Close
    Call SaveErrLog
End Sub

'功能：判断药品是否停用，再根据复选框“允许导入停用药品”返回值
'当勾选时（允许导入停用药品），不用判断药品是否停用直接返回TRUE
'当不勾选时（不允许导入停用药品），判断药品是否停用，停用返回false
Private Function 是否导入(Row As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) = "0" Then Exit Function
    
    If chk允许导入停用药品.Value = 1 Then  '允许导入停用药品
        是否导入 = True
        Exit Function
    Else '不允许导入停用药品
    
        '判断药品是否停用
        gstrSQL = "select 名称,规格 from 收费项目目录 where ID = [1] and nvl(撤档时间,to_date('3000-01-01','YYYY-MM-DD')) <> to_date('3000-01-01','YYYY-MM-DD') "
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查药品是否停用", Val(vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("id"))))
        
        If rsTemp.RecordCount = 0 Then 'rsTemp.RecordCount = 0说明该药品未停用
            是否导入 = True
        Else
            是否导入 = False
            
            mlngSum = mlngSum + 1
            If mlngSum <= 3 Then '拼提示信息串
                mstrMsg = mstrMsg & "【" & rsTemp!名称 & "(" & rsTemp!规格 & ")】" & Chr(10)
            End If
            
        End If
    End If

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdUnchoose_Click()
    If vsfMain.rows > 1 Then
        Dim i As Integer
        vsfDetail.rows = 1
        For i = 1 To vsfMain.rows - 1
            vsfMain.TextMatrix(i, vsfMain.ColIndex("choose")) = "0"
        Next
        staThis.Panels(2).Text = ""
        lblInfo.Visible = False
    End If
End Sub

Private Sub dkpView_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
        Case 1
            Item.Handle = picMain.hWnd
    End Select
End Sub

Private Sub dkpView_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Left = INT_WIDTH
    Right = INT_WIDTH
    Bottom = staThis.Height
End Sub

Private Sub dkpView_Resize()
    On Error Resume Next
    
    Dim lngL As Long, lngT As Long, lngR As Long, lngB As Long

    dkpView.GetClientRect lngL, lngT, lngR, lngB
    Me.picCaption.Move lngL, lngT, lngR - lngL
    Me.picDetail.Move lngL, lngT + picCaption.Height, lngR - lngL, lngB - lngT - picCaption.Height

'    With Me.picDetail
'        .Left = Me.sstMain.Left + INT_WIDTH
'        .Top = Me.sstMain.TabHeight + INT_WIDTH
'        .Width = Me.sstMain.Width - INT_WIDTH * 3
'        .Height = Me.sstMain.Height - Me.sstMain.TabHeight - INT_WIDTH * 2
'    End With

    With Me.vsfMain
        .Left = 0: .Top = 0
        .Width = Me.picMain.Width
        .Height = Me.picMain.Height - Me.picOperation.Height
    End With
    chkZeroInput.Top = vsfMain.Height + INT_WIDTH * 2
    chkZeroInput.Left = 0
    Me.cmdUnchoose.Top = vsfMain.Height + INT_WIDTH * 2
    Me.cmdChoose.Top = Me.cmdUnchoose.Top
    Me.cmdGetData.Top = Me.cmdChoose.Top
    
    Me.cmdUnchoose.Left = Me.picMain.Width - Me.cmdUnchoose.Width - INT_WIDTH * 2
    Me.cmdChoose.Left = Me.cmdUnchoose.Left - Me.cmdChoose.Width - INT_WIDTH
    Me.cmdGetData.Left = Me.cmdChoose.Left - Me.cmdGetData.Width - INT_WIDTH
    With Me.vsfDetail
        .Left = 0: .Top = 0
        .Width = Me.picDetail.Width
        .Height = Me.picDetail.Height - Me.picOperation.Height
    End With
    
    With Me.picOperation
        .Left = 0: .Top = Me.vsfDetail.Height
        .Width = Me.vsfDetail.Width
    End With
    Me.cmdCancel.Left = Me.picOperation.Width - Me.cmdCancel.Width - INT_WIDTH * 2
    Me.cmdImport.Left = Me.cmdCancel.Left - Me.cmdImport.Width - INT_WIDTH
    
    Me.chk允许导入停用药品.Left = Me.cmdImport.Left - Me.chk允许导入停用药品.Width - INT_WIDTH
    lblInfo.Width = cmdGetData.Left - 50
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    
    mblnLoad = False
    
    mstrPrivs = gstrprivs
    
    SetWarehouse

    chk允许导入停用药品.Value = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9MediStore", "允许导入停用药品", 0)
    
    staThis.Panels(2).Picture = picColor
    
    '初始化
    InitVSF 1
    vsfMain.AllowSelection = False
    picMain.BackColor = &H8000000F
    InitVSF 2
    vsfDetail.ExplorerBar = flexExNone

    Call InitTabs
    
    mblnLoad = True
End Sub

Private Sub InitVSF(ByVal bytIndex As Byte)
'初始化VsfView
    Dim objVSF As VSFlex8Ctl.VSFlexGrid
    Dim strCols As String
    Dim arrCols As Variant
    Dim i As Integer
    
    If bytIndex = 1 Then
        Set objVSF = vsfMain
        strCols = "||选择,choose,440|H_计划ID,planid,1000|计划单号,no,880|H_库房ID,whid,1000|库房,wh,1000|期间,length,660|审核日期,verifydate,1900" _
                & "|审核人,verifyer,660|编制日期,builddate,1900|编制人,builder,660|编制说明,explain,3000"
    Else
        Set objVSF = vsfDetail
        strCols = "||选择,choose,440|H_计划id,planid,1000|H_库房id,whid,0|计划单号,planno,880|序号,sn,440|H_药品ID,id,1000|药品名称,name,1800|计划数量,planqty,1500|执行数量,execqty,1500" _
                & "|药库单位,unit,850|成本价,costprice,1500|成本金额,cost,1500|售价,saleprice,1500|售价金额,sale,1500|供应商,provider,2000" _
                & "|上次生产商,producer,2000|H_生产日期,pdate,0|H_最大效期,avail_day,0|H_加成率,add_rate,0|H_药库包装,store_pak,0|H_分批属性,batch,0|批准文号,approval,2000|摘要,summary,0"
    End If
    
    With objVSF
        .rows = 1
        .ColWidth(0) = 130 * 2                               '第一列宽
        .ColWidth(1) = 130
        .FixedCols = 2                                       '固定前二列
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeColumns               '运行时可调整Columns宽度
        .AllowSelection = True                               '多单元选择控制开关
        .SelectionMode = flexSelectionListBox                '多单元选择控制
        .ExplorerBar = flexExSortShow
        '.BackColorSel = &HC0E0FF
        .BackColorBkg = vbWhite
    End With
    
    arrCols = Split(strCols, "|")
    With objVSF
        .Cols = UBound(arrCols) + 1
        For i = LBound(arrCols) To UBound(arrCols)
            If arrCols(i) = "" Then
                .TextMatrix(0, i) = ""
            Else
                .TextMatrix(0, i) = Split(arrCols(i), ",")(0)
                .ColKey(i) = Split(arrCols(i), ",")(1)
                .ColWidth(i) = Split(arrCols(i), ",")(2)
                'H_为隐藏列
                If Mid(Split(arrCols(i), ",")(0), 1, 2) = "H_" Then
                    .colHidden(i) = True
                Else
                    .colHidden(i) = False
                End If
            End If
        Next
        If .ColIndex("choose") > 0 Then
            .ColDataType(.ColIndex("choose")) = flexDTBoolean    '设置为Check控件
        End If
    End With
    If bytIndex = 2 Then
        vsfDetail.ColComboList(vsfDetail.ColIndex("provider")) = GetComboVSF("select 名称 from 供应商 order by 名称")
        vsfDetail.ColComboList(vsfDetail.ColIndex("producer")) = GetComboVSF("Select 名称 From 药品生产商")
    End If
End Sub

Private Sub InitTabs()
'初始化Tabs
    Dim objPane1 As Pane

    Set objPane1 = dkpView.CreatePane(1, 0, Me.ScaleY(Me.Height * 0.5, vbTwips, vbPoints), DockTopOf)
    objPane1.Title = "计划单"
    objPane1.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable

    With dkpView
        .Options.ThemedFloatingFrames = True
        .Options.LunaColors = False
        .Options.AlphaDockingContext = True
        .VisualTheme = ThemeOffice2003
        '.Options.FloatingFrameCaption = "Panes"
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
    Dim blnData As Boolean
    If vsfMain.rows <= 1 Then Exit Sub
'    If mrsDetail.State = adStateClosed Then Exit Sub
'    If mrsDetail.RecordCount = 0 Then Exit Sub
'有勾选的数据提示
    For i = 1 To vsfDetail.rows - 1
        If vsfDetail.TextMatrix(i, vsfDetail.ColIndex("choose")) = "-1" Then
            blnData = True
            Exit For
        End If
    Next
    If blnData Then
        If MsgBox("您有数据未处理，确定要取消吗？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 7050 Then Me.Width = 7050
    If Me.Height < 6500 Then Me.Height = 6500
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - .Width - 300
    End With
End Sub

Public Sub DataLoading(ByVal bytIndex As Byte, ByRef rsVal As ADODB.Recordset)
    Dim i As Integer, j As Integer
    Dim strName As String, strSpec As String, strUnit As String, strProvider As String
    Dim blnGet As Boolean
    Dim vsfTmp As VSFlex8Ctl.VSFlexGrid
    Dim blncheck库房 As Boolean
'    Dim intCostDigit As Integer, intPriceDigit As Integer, intNumberDigit As Integer, intMoneydigit As Integer
'    '药库精度
'    intCostDigit = frmMediPlanGetData.mintCostDigit
'    intPriceDigit = frmMediPlanGetData.mintPriceDigit
'    intNumberDigit = frmMediPlanGetData.mintNumberDigit
'    intMoneydigit = frmMediPlanGetData.mintMoneyDigit

    If bytIndex = 1 Then
        Set vsfTmp = vsfMain
    Else
        Set vsfTmp = vsfDetail
        If chkZeroInput.Value = False Then
            rsVal.Filter = "计划数量<>0"
        End If
    End If
    
    With vsfTmp
        If rsVal.RecordCount = 0 Then Exit Sub
        If bytIndex = 1 Then
            j = 1
            .rows = rsVal.RecordCount + 1
        Else
            j = .rows
            .rows = .rows + rsVal.RecordCount
        End If
        rsVal.MoveFirst
        For i = j To j + rsVal.RecordCount - 1
            'strName = "": strSpec = "": strUnit = ""
            'blnGet = GetMedicalInfo(rsVal!药品id, strName, strSpec, strUnit)
            '序号
            .TextMatrix(i, 1) = i
            '计划单
            If bytIndex = 1 Then
                .TextMatrix(i, .ColIndex("choose")) = 0
                .TextMatrix(i, .ColIndex("planid")) = IIf(IsNull(rsVal!id), "", rsVal!id)
                .TextMatrix(i, .ColIndex("no")) = IIf(IsNull(rsVal!NO), "", rsVal!NO)
                .TextMatrix(i, .ColIndex("whid")) = IIf(IsNull(rsVal!库房id), 0, rsVal!库房id)
                .TextMatrix(i, .ColIndex("wh")) = IIf(IsNull(rsVal!库房), "全院", rsVal!库房)
                .TextMatrix(i, .ColIndex("length")) = IIf(IsNull(rsVal!期间), "", rsVal!期间)
                .TextMatrix(i, .ColIndex("verifydate")) = IIf(IsNull(rsVal!审核日期), "", rsVal!审核日期)
                .ColFormat(.ColIndex("verifydate")) = "yyyy-mm-dd hh:mm:ss"
                .TextMatrix(i, .ColIndex("verifyer")) = IIf(IsNull(rsVal!审核人), "", rsVal!审核人)
                .TextMatrix(i, .ColIndex("builddate")) = IIf(IsNull(rsVal!编制日期), "", rsVal!编制日期)
                .ColFormat(.ColIndex("builddate")) = "yyyy-mm-dd hh:mm:ss"
                .TextMatrix(i, .ColIndex("builder")) = IIf(IsNull(rsVal!编制人), "", rsVal!编制人)
                .TextMatrix(i, .ColIndex("explain")) = IIf(IsNull(rsVal!编制说明), "", rsVal!编制说明)
            '计划内容
            Else
                .TextMatrix(i, .ColIndex("approval")) = IIf(IsNull(rsVal!批准文号), "", rsVal!批准文号)
                .TextMatrix(i, .ColIndex("Summary")) = IIf(IsNull(rsVal!编制说明), "", rsVal!编制说明)
                .TextMatrix(i, .ColIndex("planid")) = IIf(IsNull(rsVal!计划id), "", rsVal!计划id)
                If IIf(IsNull(rsVal!上次供应商), "", rsVal!上次供应商) = "" Then
                    .TextMatrix(i, .ColIndex("choose")) = "0"
                Else
                    .TextMatrix(i, .ColIndex("choose")) = IIf(IsNull(rsVal!选择), "", rsVal!选择)
                End If
'                If mbln全院 Then
'                    .TextMatrix(i, .ColIndex("whid")) = mlngID
'                Else
'                    .TextMatrix(i, .ColIndex("whid")) = IIf(IsNull(rsVal!库房id), "", rsVal!库房id)
'                End If
                
                .TextMatrix(i, .ColIndex("whid")) = Val(cbo指定.ItemData(cbo指定.ListIndex))
                
                .TextMatrix(i, .ColIndex("planno")) = IIf(IsNull(rsVal!NO), "", rsVal!NO)
                .TextMatrix(i, .ColIndex("sn")) = IIf(IsNull(rsVal!序号), "", rsVal!序号)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rsVal!药品id), "", rsVal!药品id)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rsVal!名称), "", rsVal!名称)
                .TextMatrix(i, .ColIndex("planqty")) = IIf(IsNull(rsVal!计划数量), "", rsVal!计划数量)
                .TextMatrix(i, .ColIndex("execqty")) = IIf(IsNull(rsVal!执行数量), "", rsVal!执行数量)
                .ColFormat(.ColIndex("planqty")) = "#0." & String(mintNumberDigit, "0")
                .ColFormat(.ColIndex("execqty")) = "#0." & String(mintNumberDigit, "0")
                .TextMatrix(i, .ColIndex("unit")) = IIf(IsNull(rsVal!药库单位), "", rsVal!药库单位)
                .TextMatrix(i, .ColIndex("costprice")) = IIf(IsNull(rsVal!成本价), "", rsVal!成本价) * IIf(IsNull(rsVal!药库包装), 0, rsVal!药库包装)
                .ColFormat(.ColIndex("costprice")) = "#0." & String(mintCostDigit, "0")
                
                .TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(rsVal!计划数量), 0, rsVal!计划数量) * .TextMatrix(i, .ColIndex("costprice"))
                .ColFormat(.ColIndex("cost")) = "#0." & String(mintMoneyDigit, "0")
                '处理售价
                Dim dblTmp As Double
                dblTmp = 1 + IIf(IsNull(rsVal!加成率), 0, rsVal!加成率)
                If IIf(IsNull(rsVal!是否变价), 0, rsVal!是否变价) = 1 Then
                    '变价
                    dblTmp = dblTmp * IIf(IsNull(rsVal!成本价), 0, rsVal!成本价)
                    If dblTmp >= IIf(IsNull(rsVal!指导零售价), 0, rsVal!指导零售价) Then
                        dblTmp = IIf(IsNull(rsVal!指导零售价), 0, rsVal!指导零售价)
                    Else
                        dblTmp = dblTmp _
                               + (IIf(IsNull(rsVal!指导零售价), 0, rsVal!指导零售价) - dblTmp) _
                               * (1 - (IIf(IsNull(rsVal!差价让利比), 0, rsVal!差价让利比) / 100))
                    End If
                Else
                    '定价
                    'dblTmp = IIf(IsNull(rsVal!现价), 0, rsVal!现价)
                    'If dblTmp >= IIf(IsNull(rsVal!指导零售价), 0, rsVal!指导零售价) Then
                        dblTmp = Get售价(False, Val(.TextMatrix(i, .ColIndex("id"))), Val(.TextMatrix(i, .ColIndex("whid"))), 0) 'IIf(IsNull(rsVal!售价), 0, rsVal!售价)
                    'End If
                End If
                .TextMatrix(i, .ColIndex("saleprice")) = dblTmp * IIf(IsNull(rsVal!药库包装), 0, rsVal!药库包装)
                .ColFormat(.ColIndex("saleprice")) = "#0." & String(mintPriceDigit, "0")
                
                .TextMatrix(i, .ColIndex("sale")) = IIf(IsNull(rsVal!计划数量), 0, rsVal!计划数量) * .TextMatrix(i, .ColIndex("saleprice"))
                .ColFormat(.ColIndex("sale")) = "#0." & String(mintMoneyDigit, "0")
                
                If IsNull(rsVal!供应商id) Or rsVal!供应商id = "" Then
                    .TextMatrix(i, .ColIndex("provider")) = ""
                    .TextMatrix(i, .ColIndex("choose")) = "0"
                Else
                    .TextMatrix(i, .ColIndex("provider")) = IIf(IsNull(rsVal!上次供应商), "", rsVal!上次供应商)
                End If
                .TextMatrix(i, .ColIndex("producer")) = IIf(IsNull(rsVal!上次生产商), "", rsVal!上次生产商)
                '设置分批属性
                Call Get药品分批属性(i)
                
                .TextMatrix(i, .ColIndex("pdate")) = IIf(IsNull(rsVal!上次生产日期), "", rsVal!上次生产日期)
                .TextMatrix(i, .ColIndex("avail_day")) = IIf(IsNull(rsVal!最大效期), "", rsVal!最大效期)
                .TextMatrix(i, .ColIndex("add_rate")) = IIf(IsNull(rsVal!加成率), "", rsVal!加成率)
                .TextMatrix(i, .ColIndex("store_pak")) = IIf(IsNull(rsVal!药库包装), "", rsVal!药库包装)
                
                blncheck库房 = Check库房(Val(IIf(IsNull(rsVal!药品id), 0, rsVal!药品id)), Val(cbo指定.ItemData(cbo指定.ListIndex)))
                If blncheck库房 = False Then
                    If Trim(.TextMatrix(i, .ColIndex("id"))) <> "" Then '只改变有药品的行
                        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &HFFC0C0
                        .TextMatrix(i, .ColIndex("choose")) = "0"
                    End If
                End If
               
                '执行数量>计划数量 则对应行两列字体加粗
                If Val(.TextMatrix(i, .ColIndex("execqty"))) > Val(.TextMatrix(i, .ColIndex("planqty"))) Then
                    .Cell(flexcpFontBold, i, .ColIndex("planqty"), i, .ColIndex("execqty")) = True
                End If
                
                '判断是否停用，停用显示未
                If 是否停用(Val(rsVal!药品id)) Then
                    .Cell(flexcpForeColor, i, .ColIndex("choose"), i, .Cols - 1) = &HFF00FF
                End If

            End If
            rsVal.MoveNext
        Next
        
        '确定vsf序号的宽度
        .ColWidth(1) = IIf(.rows > 0, Len(Trim(Str(.rows))) * 130 + 70, 200)
        
        '为vsfDetail合并列
        If bytIndex = 2 Then
            With vsfDetail
                .rows = .rows + 1
                .Row = .rows - 1
                .TextMatrix(.Row, 1) = .Row
                .TextMatrix(.Row, .ColIndex("planno")) = .TextMatrix(.Row - 1, .ColIndex("planno"))
                .MergeCells = flexMergeFree
                '.ColDataType(.ColIndex("choose")) = flexDTSingle
                .MergeRow(.rows - 1) = True
                .Cell(flexcpText, .Row, .ColIndex("planno") + 1, .Row, .Cols - 1) = " "
                .Cell(flexcpForeColor, .Row, .ColIndex("choose"), .Row, .Cols - 1) = &H80000010
            End With
        End If
    End With
End Sub

Private Sub Get药品分批属性(ByVal intBillRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strsql As String
    Dim int分批属性 As Integer      '0-不分批;1-分批
    Dim int药库分批 As Integer      '0-不分批;1-分批
    Dim int药房分批 As Integer      '0-不分批;1-分批
    Dim bln是否具有药房性质 As Boolean  'True-具有药房性质;False-不具有药房性质
        
    On Error GoTo errHandle
    With vsfDetail
        If Val(vsfDetail.TextMatrix(intBillRow, .ColIndex("id"))) = 0 Then Exit Sub
        
        strsql = "SELECT NVL(药库分批, 0) 药库分批,NVL(药房分批, 0) 药房分批 " & _
                " From 药品规格 WHERE 药品ID = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, "取药品库房分批属性", Val(vsfDetail.TextMatrix(intBillRow, .ColIndex("id"))))
        
        If rsTemp.RecordCount > 0 Then
            int药库分批 = rsTemp!药库分批
            int药房分批 = rsTemp!药房分批
        End If
        
        If int药房分批 = 1 Then     '如果药房分批，则分批属性为1
            int分批属性 = 1
        Else
            If int药库分批 = 1 Then
                strsql = "SELECT 部门ID From 部门性质说明 " & _
                        " WHERE ((工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')) AND 部门ID = [1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(strsql, "取部门性质", Val(vsfDetail.TextMatrix(intBillRow, .ColIndex("whid"))))
                
                bln是否具有药房性质 = (rsTemp.RecordCount > 0)
                        
                If bln是否具有药房性质 Then
                    int分批属性 = 0
                Else
                    int分批属性 = 1
                End If
            End If
        End If
        
        vsfDetail.TextMatrix(intBillRow, .ColIndex("batch")) = int分批属性
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check库房(ByVal lng药品ID As Long, ByVal lng库房ID As Long) As Boolean
    Dim rsTemp As Recordset
    
    gstrSQL = "select 收费细目id from 收费执行科室 where 收费细目id=[1] and 执行科室id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查存储库房", lng药品ID, lng库房ID)
    If rsTemp.RecordCount > 0 Then
        Check库房 = True
    Else
        Check库房 = False
    End If
End Function

Private Sub txtPlanNO_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '保存注册表信息(是否显示停用药品)
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\zl9MediStore", "允许导入停用药品", chk允许导入停用药品.Value
End Sub

Private Sub vsfDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Trim(vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("provider"))) = "" Then
        vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) = "0"
    ElseIf Col = vsfDetail.ColIndex("provider") Then
        vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) = "-1"
    End If
    staThis.Panels(2).Text = CountBuilds
End Sub

Private Sub vsfDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDetail
        '库房id无，不能修改
        If .TextMatrix(Row, .ColIndex("whid")) = "" Then
            Cancel = True
            Exit Sub
        End If
        
        '选择项可修改
        If Col = .ColIndex("choose") Then
            If Trim(.TextMatrix(Row, .ColIndex("provider"))) = "" Or .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = &HFFC0C0 Then
                Cancel = True
            Else
                Cancel = False
            End If
        ElseIf Col = .ColIndex("provider") Then
            Cancel = False
        ElseIf Col = .ColIndex("producer") Then
            Cancel = False
        ElseIf Col = .ColIndex("approval") Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfDetail_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
'    If Col = vsfDetail.ColIndex("provider") Then
'        If vsfDetail.EditText <> "" And vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) <> "-1" Then
'            vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) = "-1"
'        End If
'    End If
End Sub
Private Sub vsfDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = vsfDetail.ColIndex("approval") Then
        If KeyAscii <> vbKeyBack Then
            If LenB(StrConv(vsfDetail.EditText, vbFromUnicode)) >= 40 Or InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub vsfDetail_GotFocus()
    picCaption.BackColor = &HFFFFE9
End Sub

Private Sub vsfDetail_LostFocus()
    picCaption.BackColor = &H8000000A
End Sub

Private Sub vsfMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strsql As String, strWhere As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim bln记录 As Boolean
    
    On Error GoTo errHandle
    If Col = vsfMain.ColIndex("choose") Then
        If vsfMain.TextMatrix(Row, Col) = -1 Then
            '装载药品计划内容
            strWhere = " and id=" & vsfMain.TextMatrix(Row, vsfMain.ColIndex("planid"))
            strsql = "select -1 选择,b.库房id, B.NO, A.计划id, A.药品id," _
                   & "  D.名称, A.序号, A.计划数量 / C.药库包装 as 计划数量,nvl(A.执行数量,0) / C.药库包装 as 执行数量, C.药库单位, C.药库包装," _
                   & "  Nvl(case When Nvl(A.单价, 0) = 0 then " _
                   & "        (Select 上次采购价 From 药品库存 Where Nvl(批次, 0) =" _
                   & "           (Select nvl(Max(批次),0) From 药品库存 Where B.库房id = 库房id And A.药品id = 药品id And 性质 = 1) " _
                   & " and b.库房id = 库房id And a.药品id = 药品id And 性质 = 1 and rownum=1 ) " _
                   & "      else  A.单价 End, C.成本价) 成本价,a.售价, " _
                   & "A.上次供应商, A.上次生产商, A.说明, F.是否变价, " _
                   & "c.加成率/100 as 加成率," _
                   & "G.id 供应商ID, (select max(上次生产日期) from 药品库存 where 药品id=c.药品id) 上次生产日期, F.最大效期, C.指导零售价, C.差价让利比, Nvl(a.批准文号, Nvl(c.上次批准文号, c.批准文号)) As 批准文号, b.编制说明 " _
                   & "From 药品计划内容 A," _
                   & "     (Select ID, NO, 库房id,编制说明 From 药品采购计划 Where 审核日期 is not null " & strWhere _
                   & ") B, 药品规格 C, 收费项目目录 D, 药品目录 F,供应商 G " _
                   & "Where A.计划id = B.ID And A.药品id = C.药品id And C.药品id = D.ID And C.药品id = F.药品id " _
                   & "   and A.上次供应商=G.名称(+) " _
                   & "order by a.计划id,a.序号 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
            '装载数据
            DataLoading 2, rsTmp
        Else
            '清除药品计划内容
            strWhere = vsfMain.TextMatrix(Row, vsfMain.ColIndex("no"))
            For i = vsfDetail.rows - 1 To 1 Step -1
                If strWhere = vsfDetail.TextMatrix(i, vsfDetail.ColIndex("planno")) Then
                    vsfDetail.RemoveItem i
                End If
            Next
            
            '刷新VSF的序号
            For i = 1 To vsfDetail.rows - 1
                vsfDetail.TextMatrix(i, 1) = i
            Next
            '确定vsf序号的宽度
            vsfDetail.ColWidth(1) = IIf(vsfDetail.rows > 0, Len(Trim(Str(vsfDetail.rows))) * 130 + 70, 200)
        End If
    End If
    
    With vsfDetail
        For i = 1 To .rows - 1
            If .Cell(flexcpBackColor, i, 1, i, 2) = &HFFC0C0 Then
                bln记录 = True
                Exit For
            End If
        Next
        If bln记录 = True Then
            lblInfo.Caption = "紫色背景是该药品在导入库房中未设置存储性质，将不产生相应外购单据！"
            lblInfo.ForeColor = vbRed
            lblInfo.Visible = True
        Else
            lblInfo.Visible = False
        End If
    End With
    
    staThis.Panels(2).Text = CountBuilds
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfDetail_RowColChange()
    '当前记录用箭头指示
    vsfDetail.Cell(flexcpText, 0, 0, vsfDetail.rows - 1, 0) = ""
    If vsfDetail.Row > 0 Then
        vsfDetail.Cell(flexcpFontName, , 0) = "Marlett"
        vsfDetail.TextMatrix(vsfDetail.Row, 0) = 4
    End If
End Sub

Private Sub vsfMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfMain.TextMatrix(Row, vsfMain.ColIndex("whid")) = "" Then
        Cancel = True
        Exit Sub
    End If
        
    If Col = vsfMain.ColIndex("choose") Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub vsfMain_RowColChange()
    '当前记录用箭头指示
    vsfMain.Cell(flexcpText, 0, 0, vsfMain.rows - 1, 0) = ""
    If vsfMain.Row > 0 Then
        vsfMain.Cell(flexcpFontName, , 0) = "Marlett"
        vsfMain.TextMatrix(vsfMain.Row, 0) = 4
    End If
End Sub

Private Function GetComboVSF(ByVal strsql As String)
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
    '格式: "#1;Full time|#23;Part time|#65;Contractor|#78;Intern|#0;Other"
    strTmp = " |"
    Do While Not rsTmp.EOF
        strTmp = strTmp & rsTmp.Fields(0) & "|"
        rsTmp.MoveNext
    Loop
    GetComboVSF = strTmp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CountBuilds()
    Dim i As Integer, j As Integer
'    Dim strOldProvider As String, intOldWHID As Integer
    Dim blnFind As Boolean
    Dim rsProvider As New ADODB.Recordset
    rsProvider.Fields.Append "provider", adVarChar, 1000, adFldIsNullable
    rsProvider.Fields.Append "whid", adInteger, 18, adFldIsNullable
    rsProvider.Open
    '计算多少张入库单据，及多少个供应商+库房ID
    With vsfDetail
'        If .Rows > 1 Then
'            strOldProvider = Trim(.TextMatrix(1, .ColIndex("provider")))
'            intOldWHID = .TextMatrix(1, .ColIndex("whid"))
'        End If
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("choose")) <> "-1" Then GoTo EndFor
            If Trim(.TextMatrix(i, .ColIndex("provider"))) = "" Then GoTo EndFor
            'If Trim(.TextMatrix(i, .ColIndex("provider"))) = strOldProvider Then GoTo EndFor
            blnFind = False
            If rsProvider.RecordCount > 0 Then rsProvider.MoveFirst
            Do While Not rsProvider.EOF
                If rsProvider!Provider = Trim(.TextMatrix(i, .ColIndex("provider"))) And rsProvider!whid = .TextMatrix(i, .ColIndex("whid")) Then
                    blnFind = True
                    Exit Do
                End If
                rsProvider.MoveNext
            Loop
            If blnFind = False Then
                If chk允许导入停用药品.Value = 0 Then '不允许导入停用
                    If Not 是否停用(Val(.TextMatrix(i, .ColIndex("id")))) Then
                        rsProvider.AddNew
                        rsProvider!Provider = Trim(.TextMatrix(i, .ColIndex("provider")))
                        rsProvider!whid = .TextMatrix(i, .ColIndex("whid"))
                        rsProvider.Update
                    End If
                Else
                    rsProvider.AddNew
                    rsProvider!Provider = Trim(.TextMatrix(i, .ColIndex("provider")))
                    rsProvider!whid = .TextMatrix(i, .ColIndex("whid"))
                    rsProvider.Update
                End If
            End If
            
'        strOldProvider = Trim(.TextMatrix(i, .ColIndex("provider")))
'        intOldWHID = .TextMatrix(i, .ColIndex("whid"))
            
EndFor:
        Next
        
        CountBuilds = "您选择的数据，将生成 " & rsProvider.RecordCount & " 张入库单据。"
        rsProvider.Close
    End With
End Function

Private Function GetProviderID(ByVal strProvider As String) As Integer
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    Set rsTmp = zlDatabase.OpenSQLRecord("select ID from 供应商 where rownum=1 and 名称=[1]", Me.Caption, strProvider)
    If Not rsTmp.EOF Then GetProviderID = rsTmp!id
    rsTmp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'功能：判断是否停用,true - 停用
Private Function 是否停用(ByVal lng药品ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If lng药品ID = 0 Then Exit Function

    
    '判断药品是否停用
    gstrSQL = "select 名称,规格 from 收费项目目录 where ID = [1] and nvl(撤档时间,to_date('3000-01-01','YYYY-MM-DD')) <> to_date('3000-01-01','YYYY-MM-DD') "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查药品是否停用", lng药品ID)
    
    是否停用 = rsTemp.RecordCount <> 0  '说明该药品未停用

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub SetWarehouse()
'设置库房、ID
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    Dim i As Integer
    Dim cboTmp As ComboBox
    
    On Error GoTo ErrHand
    
    Set cboTmp = cbo指定
    
    If InStr(1, mstrPrivs, "允许药房外购入库") = 0 Then
        strsql = "Select Distinct A.ID 库房ID, A.名称 库房 " _
               & "From 部门性质说明 C, 部门性质分类 B, 部门表 A " _
               & "Where C.工作性质 = B.名称 And B.编码 In ('H', 'I', 'J') And A.ID = C.部门id And To_Char(A.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    Else
        strsql = "Select Distinct A.ID 库房ID, A.名称 库房 " _
               & "From 部门性质说明 C, 部门性质分类 B, 部门表 A " _
               & "Where C.工作性质 = B.名称 And B.编码 In ('H', 'I', 'J','K', 'L', 'M','N') And A.ID = C.部门id And To_Char(A.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    End If
    
    cboTmp.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 0 To rsTmp.RecordCount - 1
            cboTmp.AddItem rsTmp!库房
            cboTmp.ItemData(i) = rsTmp!库房id
            
            If mlng库房ID = Val(rsTmp!库房id) Then cboTmp.ListIndex = i
            
            rsTmp.MoveNext
        Next
        
    End If
    
    If cboTmp.ListIndex = -1 Then cboTmp.ListIndex = 0
    rsTmp.Close
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub ShowCard(FrmMain As Form, ByVal lng库房ID As Long)

    mlng库房ID = lng库房ID

    Me.Show vbModal, FrmMain
End Sub
