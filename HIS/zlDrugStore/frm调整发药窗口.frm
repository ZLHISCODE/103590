VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm调整发药窗口 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "调整发药窗口"
   ClientHeight    =   7644
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8424
   Icon            =   "frm调整发药窗口.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7644
   ScaleWidth      =   8424
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraInput 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8175
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   600
         TabIndex        =   3
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         Caption         =   "录入NO，姓名，就诊卡号在列表中查找并定位"
         Height          =   180
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   3600
      End
      Begin VB.Label lbl查找 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFWindows 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8175
      _cx             =   14420
      _cy             =   11033
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm调整发药窗口.frx":1A72
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
   End
End
Attribute VB_Name = "frm调整发药窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng药房ID As Long
Private mrsList As Recordset
Private mstrWin As String
Private mdate开始日期 As Date
Private mdate结束日期 As Date
Private mstr窗口 As String
Private mstrSourceDep As String
Private mstrDeptNode As String
Private mstrCurrentWin As String           '当前窗口
Private mintShowBill收费 As Integer
Private mintShowBill记帐 As Integer
Private mblnOper As Boolean     '是否执行了窗口调整

'用户定义的处方颜色，从注册表取的字符串，用;分隔
Private mstrUserRecipeColor As String
Private Sub InitComman()
'--------------------------------------
'初始化CommandBars1控件

'--------------------------------------
    With CommandBarsGlobalSettings
        Set .App = App
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '设置中文语言资源文件
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '控件整体的颜色方案，根据系统自动识别
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

    With Me.cbsMain
        .VisualTheme = xtpThemeOffice2003 '设置控件显示风格
        .EnableCustomization False '是否允许自定义设置
        .Item(1).Delete
        .Icons = frmPublic.imgPublic.Icons
    End With
End Sub
Private Sub InitTool()
'-----------------------------------------------------
'设置工具栏
'----------------------------------------------------
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    Set objBar = cbsMain.Add("工具栏1", xtpBarTop)
    objBar.ContextMenuPresent = False '工具栏上点击鼠标右键时不弹出设置菜单
    objBar.ShowTextBelowIcons = False '工具栏中的按钮文字显示在图标右侧
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_Edit_Recipe_Guide, "批量设置")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        Set objControl = .Add(xtpControlButton, mconMenu_Edit_Recipe_Average, "平均分配")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        objControl.Visible = (mstrCurrentWin <> "")

        Set objControl = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        
        Set objControl = .Add(xtpControlButton, mconMenu_Edit_Recipe_OK, "确定")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        
        Set objControl = .Add(xtpControlButton, mconMenu_File_Exit, "退出")
        objControl.Style = xtpButtonIconAndCaption '同时显示图标和文字
        objControl.BeginGroup = True
        
    End With
End Sub


Private Sub Init窗口()
    Dim strsql As String
    Dim rsRecord As Recordset
    
    On Error GoTo errHandle
    
    strsql = "select 编码,名称 from 发药窗口 where 药房id=[1] and 上班否=1"
    If mstrCurrentWin <> "" Then strsql = strsql & " And 名称<>[2] "
    Set rsRecord = zldatabase.OpenSQLRecord(strsql, "Init窗口", mlng药房ID, mstrCurrentWin)
    
    If Not (rsRecord Is Nothing) Then
        Do While Not rsRecord.EOF
            mstrWin = mstrWin & rsRecord!名称 & "|"
            rsRecord.MoveNext
        Loop
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub SetAverageWindows()
    '平均分配发药窗口，适用于窗口下班时将本窗口处方平均分配到其他上班的窗口
    '思路：1）按单据号排序，并按病人ID统计本窗口有多少人
    '2）统计有几个已上班的发药窗口，并进行平均分配
    '3）同一个病人确保分配到同一个窗口
    Dim strPatis As String      '按单据排序，依次记录病人信息（病人姓名+病人id，如果病人同名且无病人id，则视为同一人）
    Dim arrPaits, arrWins, arrPatiWins
    Dim i As Integer, lngRows As Long
    Dim intWinCount As Integer  '已上班窗口数量
    Dim strWinPati As String    '窗口和病人对应关系
    
    mrsList.Filter = ""
    mrsList.Sort = "单据,NO"
    If mrsList.RecordCount = 0 Then Exit Sub
    
    '如果已经没有上班窗口了就退出
    If mstrWin = "" Then Exit Sub
    
    '按NO排序后记录病人ID，不重复
    With mrsList
        Do While Not .EOF
            If InStr(1, "|" & strPatis & "|", "|" & !姓名 & !病人ID & "|") = 0 Then
                strPatis = IIf(strPatis = "", "", strPatis & "|") & !姓名 & !病人ID
            End If
            
            .MoveNext
        Loop
    End With
    
    arrPaits = Array()
    arrPaits = Split(strPatis, "|")
    
    If Right(mstrWin, 1) = "|" Then mstrWin = Mid(mstrWin, 1, Len(mstrWin) - 1)
    arrWins = Array()
    arrWins = Split(mstrWin, "|")
    intWinCount = UBound(arrWins) + 1
    
    '平均分配窗口，记录窗口和病人ID关系
    For i = 0 To UBound(arrPaits)
        strWinPati = IIf(strWinPati = "", "", strWinPati & "|") & arrPaits(i) & "," & arrWins(IIf(((i + 1) Mod intWinCount) = 0, intWinCount - 1, (i + 1) Mod intWinCount - 1))
    Next
    
    arrPatiWins = Array()
    arrPatiWins = Split(strWinPati, "|")
    
    '设置新窗口
    With VSFWindows
        For lngRows = 1 To .rows - 1
            For i = 0 To UBound(arrPatiWins)
                If .TextMatrix(lngRows, .ColIndex("姓名")) & .TextMatrix(lngRows, .ColIndex("病人id")) = Split(arrPatiWins(i), ",")(0) Then
                    .TextMatrix(lngRows, .ColIndex("新窗口")) = Split(arrPatiWins(i), ",")(1)
                    Exit For
                End If
            Next
        Next
    End With
    
End Sub

Public Function showMe(ByVal lng药房ID As Long, ByVal FrmMain As Form, ByVal date开始日期 As Date, ByVal date结束日期 As Date, _
    ByVal strDeptNode As String, Optional ByVal strCurrentWin As String) As Boolean
    'lng药房ID：当前药房
    'FrmMain：主窗体
    'date开始日期,date结束日期：处方填制日期范围
    'strDeptNode：药房站点
    'strCurrentWin：当前窗口
    mlng药房ID = lng药房ID
    mdate开始日期 = date开始日期
    mdate结束日期 = date结束日期
    mstrDeptNode = strDeptNode
    mstrCurrentWin = strCurrentWin
    
    Call frm调整发药窗口.Show(1, FrmMain)
    
    showMe = mblnOper
End Function

Private Sub InitVSFGrid()
    With Me.VSFWindows
        .rows = 1
        
        .ColComboList(.ColIndex("新窗口")) = mstrWin
        .ComboList = mstrWin
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_View_Refresh  '执行刷新操作
            Me.VSFWindows.rows = 1
            LoadData
            InitGridData
        Case mconMenu_File_Exit      '执行退出操作
            Unload Me
        Case mconMenu_Edit_Recipe_Guide  '执行批量设置向导
            Edit_Recipe_Guide
        Case mconMenu_Edit_Recipe_OK     '执行确定操作
            Call Edit_Recipe_OK
            Me.VSFWindows.rows = 1
            LoadData
            InitGridData
        Case mconMenu_Edit_Recipe_Average   '平均分配
            Call SetAverageWindows
    End Select
End Sub
Private Sub Edit_Recipe_Guide()
    Dim strConWin As String
    Dim str处方 As String
    Dim str新窗口 As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    frm批量设置向导.showMe mlng药房ID, strConWin, str处方, str新窗口, mstrCurrentWin
    
    For i = -1 To UBound(Split(strConWin, ","))
        If i = UBound(Split(strConWin, ",")) - 1 Then Exit For
        
        For j = -1 To UBound(Split(str处方, ","))
            If j = UBound(Split(str处方, ",")) - 1 Then Exit For
            
            If UBound(Split(strConWin, ",")) <> -1 Or UBound(Split(str处方, ",")) <> -1 Then
                For k = 1 To Me.VSFWindows.rows - 1
                    If UBound(Split(strConWin, ",")) = -1 Then
                        If Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("类别")) = Split(str处方, ",")(j + 1) Then
                            Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("新窗口")) = str新窗口
                        End If
                    ElseIf UBound(Split(str处方, ",")) = -1 Then
                        If Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("现窗口")) = Split(strConWin, ",")(i + 1) Then
                            Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("新窗口")) = str新窗口
                        End If
                    Else
                        If Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("现窗口")) = Split(strConWin, ",")(i + 1) And Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("类别")) = Split(str处方, ",")(j + 1) Then
                             Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("新窗口")) = str新窗口
                        End If
                    End If
                Next
            End If
        Next
    Next
End Sub
Private Sub Edit_Recipe_OK()
    Dim i As Integer
    Dim arrSql As Variant
    
    arrSql = Array()
    
    With VSFWindows
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("现窗口")) <> .TextMatrix(i, .ColIndex("新窗口")) And .TextMatrix(i, .ColIndex("新窗口")) <> "" Then
                gstrSQL = "zl_未发药品记录_分配发药窗口("
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("NO")) & "'"
                gstrSQL = gstrSQL & "," & .TextMatrix(i, .ColIndex("单据"))
                gstrSQL = gstrSQL & "," & mlng药房ID
                gstrSQL = gstrSQL & ",'" & .TextMatrix(i, .ColIndex("新窗口")) & "')"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        Next
    End With
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "Edit_Recipe_OK")
    Next
    gcnOracle.CommitTrans
    
    mblnOper = True
    
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mblnOper = False
    
    mstrUserRecipeColor = zldatabase.GetPara("处方颜色", glngSys, 1341)
    If mstrUserRecipeColor = "" Then mstrUserRecipeColor = GetDefaultRecipeColor
    
    InitComman
    
    InitTool
    
    Init窗口
    
    InitVSFGrid
    
    LoadData
    
    InitGridData
End Sub

Private Sub LoadData()
    Dim strsql As String
    Dim str门诊 As String
    Dim strSub1 As String
    Dim strSub2 As String
    Dim str住院 As String
    
    On Error GoTo errHandle
    gstrSQL = "Select /*+ Rule*/ '' As 颜色, 处方类型 ,'' As 选择 ,'0' As 标志,类型,单据,已收费,配药人,NO,姓名,发药窗口," & _
            " 就诊卡号,门诊号,身份证号,IC卡号,病人ID,医保号,住院号," & _
            " 门诊标志,记录性质,Decode(A.收费类别, '7', '7', '0') As 收费类别,日期 " & _
            " From ("
            
    str门诊 = " Select A.发药窗口,A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.NO,A.姓名,C.零售金额,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.医保号,A.住院号,d.实收金额, Nvl(A.处方类型,Nvl(C.注册证号,0)) As 处方类型,D.门诊标志,D.记录性质,D.收费类别 " & _
            " From " & _
            " (Select distinct A.发药窗口,B.就诊卡号,B.门诊号,B.身份证号,B.IC卡号,B.医保号,B.住院号,A.优先级,A.填制日期,Decode(Nvl(A.已收费,0),1,'','(未)')||Decode(A.单据,8,'收费',9,'记帐') 类型,A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明,B.病人ID, A.处方类型,a.对方部门id " & _
            " From 未发药品记录 A,病人信息 B " & _
            " Where 1=1 "

    '主要条件
    str门诊 = str门诊 & " And (A.库房ID=[1] Or A.库房ID Is NULL) And A.填制日期 Between [2] And [3] "
    
    
    str门诊 = str门诊 & " And A.病人ID=B.病人ID(+)"
    
    
    str门诊 = str门诊 & " And A.单据 IN(8,9)"
    
    If mstrCurrentWin <> "" Then str门诊 = str门诊 & " And A.发药窗口=[5] "
        
    str门诊 = str门诊 & ") A,药品收发记录 C, 门诊费用记录 D, 部门表 B " & _
              " Where C.费用id = D.ID And nvl(c.发药方式,-999)<>-1 and A.单据=C.单据 And A.NO=C.NO And C.审核人 Is NULL " & _
              " And Nvl(D.费用状态,0)<>1 And (C.库房id=[1] Or C.库房id Is null)  And a.对方部门id = b.Id "
    
    If mstrDeptNode <> "" Then
        str门诊 = str门诊 & " And (b.站点 = [4] Or b.站点 Is Null) "
    End If
    
    str住院 = Replace(str门诊, "门诊费用记录", "住院费用记录")
    str住院 = Replace(str住院, "And Nvl(D.费用状态,0)<>1", "")
    
    '门诊划价及门诊记帐
    gstrSQL = gstrSQL & str门诊 & "Union All " & str住院
    
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.发药窗口,A.优先级,A.类型,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明,A.就诊卡号,A.门诊号,A.身份证号,A.IC卡号,A.病人ID,A.医保号,A.住院号,A.处方类型,A.门诊标志,A.记录性质,Decode(A.收费类别, '7', '7', '0'),A.日期 "
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.类型,A.单据,A.病人id,A.No"
    
    Set mrsList = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mlng药房ID, _
            mdate开始日期, _
            zldatabase.Currentdate, _
            mstrDeptNode, _
            mstrCurrentWin)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitGridData()
    Dim i As Integer
    
    With Me.VSFWindows
        .Redraw = flexRDNone
        
        .rows = .rows + mrsList.RecordCount
        For i = 1 To mrsList.RecordCount
            If mrsList!处方类型 = 1 Then
                .TextMatrix(i, .ColIndex("类型")) = "儿科"
            ElseIf mrsList!处方类型 = 2 Then
                .TextMatrix(i, .ColIndex("类型")) = "急诊"
            ElseIf mrsList!处方类型 = 3 Then
                .TextMatrix(i, .ColIndex("类型")) = "精二"
            ElseIf mrsList!处方类型 = 4 Then
                .TextMatrix(i, .ColIndex("类型")) = "精一"
            ElseIf mrsList!处方类型 = 5 Then
                .TextMatrix(i, .ColIndex("类型")) = "麻醉"
            Else
                .TextMatrix(i, .ColIndex("类型")) = "普通"
            End If
            
            .TextMatrix(i, .ColIndex("类别")) = IIf(IsNull(mrsList!类型), "", mrsList!类型)
            .TextMatrix(i, .ColIndex("NO")) = mrsList!NO
            .TextMatrix(i, .ColIndex("日期")) = Format(mrsList!日期, "YYYY-MM-DD HH:MM")
            .TextMatrix(i, .ColIndex("姓名")) = IIf(IsNull(mrsList!姓名), "", mrsList!姓名)
            .TextMatrix(i, .ColIndex("现窗口")) = zlStr.Nvl(mrsList!发药窗口)
            .TextMatrix(i, .ColIndex("单据")) = mrsList!单据
            .TextMatrix(i, .ColIndex("就诊卡号")) = zlStr.Nvl(mrsList!就诊卡号)
            .TextMatrix(i, .ColIndex("病人ID")) = zlStr.Nvl(mrsList!病人ID)
            
            .Cell(flexcpBackColor, i, .ColIndex("类型"), i, .ColIndex("类型")) = Val(Split(mstrUserRecipeColor, ";")(Val(mrsList!处方类型)))
            
            mrsList.MoveNext
        Next
        
        .Cell(flexcpFontBold, 0, .ColIndex("新窗口"), 0, .ColIndex("新窗口")) = True
        .Cell(flexcpForeColor, 0, .ColIndex("新窗口"), 0, .ColIndex("新窗口")) = vbBlue
        
        If .rows > 1 Then
            .Cell(flexcpBackColor, 1, .ColIndex("新窗口"), .rows - 1, .ColIndex("新窗口")) = &HFFE3C8    '&HFFEDDD     '&HFFC0C0
        End If
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.fraInput.Move 100, 490, Me.ScaleWidth - 200, Me.fraInput.Height
    Me.VSFWindows.Move 100, Me.fraInput.Top + Me.fraInput.Height, Me.ScaleWidth - 200, Me.ScaleHeight - (Me.fraInput.Top + Me.fraInput.Height) - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrWin = ""
End Sub

Private Sub txtMsg_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strText As String
    Dim i As Integer
    
    If KeyCode <> 13 Then Exit Sub
    strText = Trim(Me.txtMsg.Text)
    
    For i = 1 To Me.VSFWindows.rows - 1
        If InStr(1, Me.VSFWindows.TextMatrix(i, VSFWindows.ColIndex("NO")), strText) <> 0 Or InStr(1, Me.VSFWindows.TextMatrix(i, VSFWindows.ColIndex("姓名")), strText) Or InStr(1, Me.VSFWindows.TextMatrix(i, VSFWindows.ColIndex("就诊卡号")), strText) Then
            Me.VSFWindows.Row = i
            Exit Sub
        End If
    Next
End Sub

Private Sub VSFWindows_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> VSFWindows.ColIndex("新窗口") Then
        Cancel = True
    End If
End Sub


