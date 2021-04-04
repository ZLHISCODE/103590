VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPatientRptPrint 
   Caption         =   "病区检验报告打印"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientRptPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkPrinted 
      Caption         =   "包含已打印记录"
      Height          =   240
      Left            =   75
      TabIndex        =   8
      Top             =   5325
      Width           =   8565
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgData 
      Height          =   3855
      Left            =   510
      TabIndex        =   7
      Top             =   1950
      Width           =   8430
      _cx             =   14870
      _cy             =   6800
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   0
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
      AllowUserFreezing=   1
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame FraWhere 
      Height          =   1080
      Left            =   90
      TabIndex        =   0
      Top             =   555
      Width           =   8490
      Begin VB.ComboBox cboPatient 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   225
         Width           =   1560
      End
      Begin VB.ComboBox cboDir 
         Height          =   315
         Left            =   990
         TabIndex        =   13
         Text            =   "cboDir"
         Top             =   615
         Width           =   1560
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&P"
         Height          =   315
         Left            =   5850
         TabIndex        =   12
         Top             =   615
         Width           =   285
      End
      Begin VB.TextBox txtItem 
         Height          =   315
         Left            =   3585
         TabIndex        =   11
         Top             =   615
         Width           =   2265
      End
      Begin VB.TextBox txtPatiNo 
         Height          =   300
         Left            =   7065
         TabIndex        =   4
         Top             =   225
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker dtpS 
         Height          =   300
         Left            =   3585
         TabIndex        =   1
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   112001027
         CurrentDate     =   40954
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   300
         Left            =   4920
         TabIndex        =   2
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   112001027
         CurrentDate     =   40954
      End
      Begin VB.Label lblDir 
         AutoSize        =   -1  'True
         Caption         =   "住院医生"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "检验项目"
         Height          =   195
         Left            =   2700
         TabIndex        =   9
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申请科室↓"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   270
         Width           =   900
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         Caption         =   "住院号↓"
         Height          =   195
         Left            =   6315
         TabIndex        =   5
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         Caption         =   "申请日期↓"
         Height          =   195
         Left            =   2700
         TabIndex        =   3
         Top             =   270
         Width           =   900
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   180
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuPop 
      Caption         =   "选择"
      Visible         =   0   'False
      Begin VB.Menu mnuPatiNo 
         Caption         =   "住院号"
      End
      Begin VB.Menu mnuBedNo 
         Caption         =   "床号"
      End
   End
   Begin VB.Menu mnudate 
      Caption         =   "日期"
      Visible         =   0   'False
      Begin VB.Menu mnuAppdate 
         Caption         =   "申请日期"
      End
      Begin VB.Menu mnuAdjdate 
         Caption         =   "审核日期"
      End
   End
End
Attribute VB_Name = "frmPatientRptPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsDept As ADODB.Recordset  '操作员所在病区记录集
Private mlngOptDeptID As Long           '病区ID
Private mlngSys As Long                 '系统号
Private mcnOracle As ADODB.Connection
Private mstrPrivs As String         '模块权限

Private mintPrint As Integer    '是否在打印
Private Enum Col
    选择 = 0: 姓名: 性别: 年龄: 住院号: 床号: 检验项目: 申请科室: 申请人: 审核人: 审核时间: 医嘱id: 发送号: 病人ID: ID
End Enum
Private Function CbsSetting(ByRef cbsMain As CommandBars)
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.ActiveMenuBar.Closeable = False
      
End Function

Private Sub CbsButtonInit(ByRef cbsMain As CommandBars, Buttons As Collection, _
                         Optional blnLargeIcons As Boolean = False, _
                         Optional Position As XTPBarPosition)
    '创建工具栏菜单
    'cbsMain :工具栏对象
    'Buttons :菜单集合,每个元素的格式为 菜单id,标题,是否分组
    'blnLargeIcons :是否大图标
    'Position      :菜单位置
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim strButton As Variant
    Dim varButton As Variant

    Call CbsSetting(cbsMain)
    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.ActiveMenuBar
    cbsMain.Options.LargeIcons = blnLargeIcons  '小图标
    objBar.Position = Position   '工具栏在顶部

    For Each strButton In Buttons
        varButton = Split(strButton, ",")
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, Val(varButton(0)), varButton(1))     '固有
            objControl.Style = xtpButtonIconAndCaption
            If UCase(varButton(2)) = "TRUE" Then objControl.BeginGroup = True '固有
        End With
    Next
    cbsMain.RecalcLayout
End Sub

Public Sub ShowME(cnOracle As ADODB.Connection, ByVal lngSys As Long, Objfrm As Object, lngOpterDeptID As Long, MainPrivs As String)
    
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo hErr
    
    mlngOptDeptID = lngOpterDeptID
    mlngSys = lngSys
    mstrPrivs = MainPrivs
    Set mcnOracle = cnOracle
    
    strSQL = "Select a.Id, a.编码, a.名称 From 部门表 A, 病区科室对应 D Where a.Id = d.科室id And d.病区id = [1]"

    Set mrsDept = zlDatabase.OpenSQLRecord(strSQL, "取病区科室", mlngOptDeptID)
    If mrsDept.EOF Then
        MsgBox "操作员没有可操作的科室！", vbQuestion, Me.Caption
        Exit Sub
    End If

    Me.Show , Objfrm
    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadBaseData()
    '装入病区数据
    Dim lngloop As Long
    dtpS.Value = zlDatabase.Currentdate
    dtpE.Value = dtpS.Value
    txtPatiNo = ""
    cboPatient.Clear
    cboPatient.AddItem "<所有科室>"
    
    Do Until mrsDept.EOF
        cboPatient.AddItem Trim("" & mrsDept!编码 & "-" & mrsDept!名称)
        If mlngOptDeptID = Val("" & mrsDept!ID) Then
            cboPatient.ListIndex = cboPatient.NewIndex
        End If
        cboPatient.ItemData(cboPatient.NewIndex) = Val("" & mrsDept!ID)
        mrsDept.MoveNext
    Loop
    If cboPatient.ListIndex = -1 Then cboPatient.ListIndex = 0
    Call InitDoctors(cboPatient.ItemData(cboPatient.ListIndex))
    Call mnuPatiNo_Click
    Call mnuAppdate_Click
    Call vfgSetting(0, vfgData, "选择,800,1; 姓名,1200,1;性别,600,1;年龄,600,1;住院号,900,1;床号,900,1;检验项目,2000,1;申请科室,1200,1;申请人,900,1;审核人,900,1;审核时间,1200,1;医嘱id,0,1;发送号,0,1;病人ID,0,1;ID,0,1")
    vfgData.ColDataType(Col.选择) = flexDTBoolean
End Sub

Private Sub LoadDataToVfg()
    '根据界面上的条件，将数据填入表格控件
    '
    Dim strSQL As String, rsTmp As ADODB.Recordset, intIndex As Integer
    Dim dateS As Date, dateE As Date, lngPatiDeptID As Long, strNO As String, strDepts As String, strPatients As String
    
    On Error GoTo hErr
    
    txtPatiNo.SetFocus
    
    dateS = CDate(Format(dtpS.Value, "yyyy-MM-dd 00:00:00"))
    dateE = CDate(Format(dtpE.Value, "yyyy-MM-dd 23:59:59"))
    If dateE < dateS Then
        MsgBox "结束日期不能小于开始日期！", vbInformation, Me.Caption
        Exit Sub
    End If
    If DateDiff("d", dateS, dateE) > 31 Then
        If MsgBox("查询的日期范围超过了31天，会影响系统响应速度，是否继续？", vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
    End If
    
    Call vfgSetting(0, vfgData, "选择,800,1; 姓名,1200,1;性别,600,1;年龄,600,1;住院号,900,1;床号,900,1;检验项目,2000,1;申请科室,1200,1;申请人,900,1;审核人,900,1;审核时间,1200,1;医嘱id,0,1;发送号,0,1;病人ID,0,1;ID,0,1")
    vfgData.ColDataType(Col.选择) = flexDTBoolean
    
    
    lngPatiDeptID = Val(cboPatient.ItemData(cboPatient.ListIndex))
    
    strNO = Trim(DelInvalidChar(txtPatiNo))
    strDepts = ""
    If lngPatiDeptID <= 0 Then
        If cboPatient.ListCount > 0 Then
            For intIndex = 1 To cboPatient.ListCount - 1
                If Val(cboPatient.ItemData(intIndex)) <> 0 Then strDepts = strDepts & "," & Val(cboPatient.ItemData(intIndex))
            Next
        End If
    End If
    If strDepts <> "" Then
        strDepts = Mid(strDepts, 2)
        strSQL = "select /*+ RULE */ 0 as 选择,a.姓名, a.性别, a.年龄, a.住院号, a.床号, a.检验项目,f.名称 as 申请科室,a.申请人,a.审核人,a.审核时间,a.医嘱id,b.发送号,a.病人id,a.id " & vbNewLine & _
            "From 检验标本记录 A, 病人医嘱发送 B, 部门表 F,(Select * From Table(Cast(f_str2list([5]) As zltools.t_strlist))) G " & vbNewLine & _
            "where a.病人来源 = 2 and a.医嘱id=b.医嘱id and a.申请科室id=f.id and  a.审核时间 Is Not Null "
        strSQL = strSQL & " And A.申请科室id = G.Column_Value "
    Else
        strSQL = "select 0 as 选择,a.姓名, a.性别, a.年龄, a.住院号, a.床号, a.检验项目,f.名称 as 申请科室,a.申请人,a.审核人,a.审核时间,a.医嘱id,b.发送号,a.病人id,a.id " & vbNewLine & _
            "From 检验标本记录 A, 病人医嘱发送 B, 部门表 F" & vbNewLine & _
            "where a.病人来源 = 2 and a.医嘱id=b.医嘱id and a.申请科室id=f.id  And a.审核时间 Is Not Null "
    End If
    
    
    If chkPrinted.Value = 0 Then
        strSQL = strSQL & " And (a.打印次数 Is Null Or a.打印次数 <= 0)"
    End If
    
    
    If lngPatiDeptID <> 0 Then
        strSQL = strSQL & " and a.申请科室id = [3] "
    End If
    If strNO <> "" Then
        If lblNo.Caption = "住院号↓" Then
            strSQL = strSQL & " and a.住院号 = [4] "
        Else
            strSQL = strSQL & " and a.床号 = [4] "
        End If
    End If
    
    If lbldate.Caption = "申请日期↓" Then
        strSQL = strSQL & " and a.申请时间 Between [1] And [2] "
    Else
        strSQL = strSQL & " and a.审核时间 Between [1] And [2] "
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dateS, dateE, lngPatiDeptID, strNO, strDepts)
    With vfgData
        Do Until rsTmp.EOF
            If strPatients = "" Then
                strPatients = rsTmp("病人id")
            Else
                If InStr(strPatients, rsTmp("病人id")) = 0 Then
                    strPatients = strPatients & "," & rsTmp("病人id")
                End If
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    strSQL = "select distinct 0 as 选择,a.姓名, a.性别, a.年龄, a.住院号, a.床号, a.检验项目,f.名称 as 申请科室,a.申请人,a.审核人,a.审核时间,a.医嘱id,b.发送号,a.病人id,a.id " & vbNewLine & _
            "From 检验标本记录 A, 病人医嘱发送 B, 部门表 F,(Select * From Table(Cast(f_str2list([5]) As zltools.t_strlist))) G　" & vbNewLine & _
            "where a.病人来源 = 2 and a.医嘱id=b.医嘱id and a.申请科室id=f.id and  a.审核时间 Is Not Null and A.病人id = G.Column_Value "
    
    If cboDir.Text <> "" Then
        strSQL = strSQL & " and  a.申请人=[6]"
    End If
    
    If txtItem.Text <> "" Then
        strSQL = strSQL & " and a.检验项目=[7]"
    End If
    
    If chkPrinted.Value = 0 Then
        strSQL = strSQL & " And (a.打印次数 Is Null Or a.打印次数 <= 0)"
    End If
    
    If lbldate.Caption = "申请日期↓" Then
        strSQL = strSQL & " and a.申请时间 Between [1] And [2] "
    Else
        strSQL = strSQL & " and a.审核时间 Between [1] And [2] "
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dateS, dateE, lngPatiDeptID, strNO, strPatients, cboDir.Text, txtItem.Text)
    
    With vfgData
        Do Until rsTmp.EOF
            
            If Val(.TextMatrix(.Rows - 1, Col.ID)) <> 0 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, Col.ID) = Val("" & rsTmp!ID)
            .TextMatrix(.Rows - 1, Col.选择) = Val("" & rsTmp!选择)
            .TextMatrix(.Rows - 1, Col.姓名) = Trim("" & rsTmp!姓名)
            .TextMatrix(.Rows - 1, Col.性别) = Trim("" & rsTmp!性别)
            .TextMatrix(.Rows - 1, Col.年龄) = Trim("" & rsTmp!年龄)

            .TextMatrix(.Rows - 1, Col.住院号) = Trim("" & rsTmp!住院号)
            .TextMatrix(.Rows - 1, Col.床号) = Trim("" & rsTmp!床号)
            .TextMatrix(.Rows - 1, Col.检验项目) = Trim("" & rsTmp!检验项目)
            .TextMatrix(.Rows - 1, Col.申请科室) = Trim("" & rsTmp!申请科室)
            .TextMatrix(.Rows - 1, Col.申请人) = Trim("" & rsTmp!申请人)
            
            .TextMatrix(.Rows - 1, Col.审核人) = Trim("" & rsTmp!审核人)
            .TextMatrix(.Rows - 1, Col.审核时间) = Trim("" & rsTmp!审核时间)
            .TextMatrix(.Rows - 1, Col.医嘱id) = Trim("" & rsTmp!医嘱id)
            .TextMatrix(.Rows - 1, Col.发送号) = Trim("" & rsTmp!发送号)
            .TextMatrix(.Rows - 1, Col.病人ID) = Trim("" & rsTmp!病人ID)
            
            rsTmp.MoveNext
        Loop
        If rsTmp.RecordCount > 0 Then
            chkPrinted.Caption = "包含已打印记录    " & "共有" & .Rows - 1 & "条记录"
        Else
            chkPrinted.Caption = "包含已打印记录"
        End If
    End With
    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitDoctors(ByVal lngdptID As Long)
'功能：读取当前开单科室中包含的所有人员
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldDoctor As String
    
    strOldDoctor = Me.cboDir.Text
    
    Me.cboDir.Clear
    
    '科室医生或护士
    If lngdptID = 0 Then
        strSQL = _
            "Select Distinct A.ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
            " C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
            " And C.人员性质 IN('医生')  " & _
            " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "

    Else
        strSQL = _
            "Select Distinct A.ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
            " C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
            " From 人员表 A,部门人员 B,人员性质说明 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
            " And C.人员性质 IN('医生') And B.部门ID=[1] " & _
            " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
    End If
    strSQL = strSQL & " Order by 简码,人员性质 Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngdptID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboDir.AddItem rsTmp!姓名
            cboDir.ItemData(cboDir.ListCount - 1) = rsTmp!ID
            If rsTmp!姓名 = strOldDoctor Then
                cboDir.ListIndex = cboDir.NewIndex
            End If
            
            If rsTmp!ID = UserInfo.ID And cboDir.ListIndex = -1 Then cboDir.ListIndex = cboDir.NewIndex
            rsTmp.MoveNext
        Next
        
        If cboDir.ListCount = 1 And cboDir.ListIndex = -1 Then cboDir.ListIndex = 0
    End If
End Sub

Private Sub cboDir_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        If mbln微生物项目 = True Then
'            Call cbodir_Validate(False)
'            dtp(0).SetFocus
'        Else
'            zlCommFun.PressKey vbKeyTab
            Call cboDir_Validate(False)
'            SetFocusNextIndex Me.cboDir.TabIndex + 2
'            gintSelectFocus = 2
'        End If
    End If
End Sub

Private Sub cboDir_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
'    If cboDir.ListIndex <> -1 Then mstrReqDoctor = Me.cboDir.Text: Exit Sub '已选中
    If cboDir.Text = "" Then '无输入
        Exit Sub
    End If

    strInput = UCase(NeedName(cboDir.Text))
    '全院医生
    strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
        " And B.部门ID IN(" & strSQL & ")" & _
        " And (Upper(A.编号) Like [1] Or Upper(A.姓名) Like [2] Or Upper(A.简码) Like [2])" & _
        " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
        " Order by A.简码"

    On Error GoTo errH
    vRect = GetControlRect(cboDir.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱医生", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboDir.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        cboDir.Text = rsTmp!姓名
'        Me.dtp(0).SetFocus
'        SetFocusNextIndex Me.cbodir.TabIndex
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的医生。", vbInformation, gstrSysName
        End If
        Cancel = True:  Exit Sub
    End If
'    If Len(Trim(Me.cboDir.Text)) > 0 Then mstrReqDoctor = Me.cboDir.Text
'    gintSelectFocus = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgSelect(ByVal intSelect As Integer)
    Dim iRow As Integer
    With vfgData
        For iRow = .FixedRows To .Rows - 1
            If intSelect = 1 Then
                .TextMatrix(iRow, 0) = 1
            Else
                .TextMatrix(iRow, 0) = 0
            End If
        Next
    End With
End Sub
Private Sub PrintSelect()
    '打印选中的记录
    Dim iRow As Integer, intCount As Integer, intCurr As Integer
    Dim lngRedeID As Long, lngSendID As Long, lngPatiID As Long, lngSampleID As Long
    mintPrint = 1
    intCount = 0: intCurr = 0
    With vfgData
        For iRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(iRow, Col.选择)) <> 0 Then
                intCount = intCount + 1
            End If
            
        Next
        If intCount <= 0 Then
            MsgBox "请选择要打印的报告后再执行此操作！", vbInformation, Me.Caption
            mintPrint = 0
            Exit Sub
        End If
        If intCount > 300 Then
            If MsgBox("要打印的报告超过300份，打印时间会比较长，是否继续？" & vbNewLine & "点[是]开始打印，点[否]不打印。", vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                mintPrint = 0
                Exit Sub
            End If
        End If
        
        For iRow = .FixedRows To .Rows - 1
            DoEvents
            If Val(.TextMatrix(iRow, Col.选择)) <> 0 Then
                intCurr = intCurr + 1
                lngRedeID = Val(.TextMatrix(iRow, Col.医嘱id))
                lngSendID = Val(.TextMatrix(iRow, Col.发送号))
                lngPatiID = Val(.TextMatrix(iRow, Col.病人ID))
                lngSampleID = Val(.TextMatrix(iRow, Col.ID))
                Call ReportPrint(lngRedeID, lngSendID, lngPatiID, lngSampleID, intCurr, True)
                .TextMatrix(iRow, Col.选择) = 0
            End If
        Next
    End With
    mintPrint = 0
    
End Sub
Private Sub ReportPrint(ByVal lngRedeID As Long, ByVal lngSendID As Long, ByVal lngPatiID As Long, ByVal lngSampleID As Long, _
                        ByVal intCount As Integer, ByVal blnPrint As Boolean)
    '单个报告打印
    'lngRedeID :医嘱ID
    'lngSendID :发送号
    'lngPatiID :病人ID
    'lngSampleID :标本ID
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    Dim lngKey As Long
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "正在打印第" & intCount & "份报告...", Me
    
    '生成图形供自定义报表调用
    strSQL = "select id from 检验图像结果 where 标本id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lngSampleID)
    intLoop = 1
    Do Until rsTmp.EOF
        strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
        Call LoadImageData(App.path, rsTmp("ID"))
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    If GetReportCode(lngRedeID, lngSendID, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
'        Call ReportOpen(gcnOracle, mlngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lngRedeID, _
'                        "病人ID=" & lngPatiID, "标本ID=" & lngSampleID, "多个医嘱=" & lngRedeID, "多个标本=" & lngSampleID, _
'                        "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
'                        "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
'                        "图形9=" & strChart(9), IIf(blnPrint, 2, 1))

            Call ReportOpen(mcnOracle, mlngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lngRedeID, _
                            "病人ID=" & lngPatiID, "标本ID=" & lngSampleID, _
                            "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
                            "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
                            "图形9=" & strChart(9), IIf(blnPrint, 2, 1))
    End If
    
    
    On Error GoTo errH
    If blnPrint = True Then
        strSQL = "ZL_检验标本记录_标本质控(" & lngSampleID & ",'',1)"   '打印次数加1
        zlDatabase.ExecuteProcedure strSQL, gstrSysName
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

Public Sub vfgSetting(ByVal LngStyle As Long, ByRef objVfg As VSFlexGrid, Optional ByVal strTtile As String, Optional VsfImg As ImageList)
    'lngStyle＝0 默认设置，统一Vfg表格的外观
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'objVfg:    要初始化的控件
    'VsfImg:    ImageList图标集控件对象

    Dim arrHead As Variant, i As Long, strHead As String
    If strTtile = "" Then
        strHead = "第1列,900,1;第2列,900,1;第3列,900,1"
    Else
        strHead = strTtile
    End If
    arrHead = Split(strHead, ";")
    
    
    With objVfg
        '1.边框
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .GridLines = flexGridFlat
        .GridColorFixed = flexGridFlat
        
        '2.颜色
        .BackColor = vbWindowBackground '窗口背景
        .BackColorAlternate = vbWindowBackground
        .BackColorBkg = vbWindowBackground
        .BackColorFixed = vbButtonFace '按钮表面
        .BackColorFrozen = &H0&         '黑
        .FloodColor = &HC0&             '红
        .BackColorSel = &HFFEBD7        '浅绿
        .ForeColor = vbWindowText       '窗口文本
        .ForeColorFixed = vbButtonText  '按钮文本
        .ForeColorFrozen = &H0&         '黑
        .ForeColorSel = vbWindowText
        
        .GridColor = vbApplicationWorkspace '应用程序工作区
        .GridColorFixed = vbApplicationWorkspace
        .SheetBorder = vbWindowBackground
        .TreeColor = vbButtonShadow         '按钮阴影
        
        '3.初始化行列

        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值

            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        
        '固定行文字居中
        If .FixedRows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .RowHeight(0) = 300
        .RowHeightMin = 300
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
        
        
        '4.其他属性
        .SelectionMode = flexSelectionByRow     '整行选择
        .ExplorerBar = flexExNone               '点标题栏不响应（排序及移动列）操作
        .AllowUserResizing = flexResizeColumns  '可调整列宽
        .Editable = flexEDNone                  '只读
        
    End With
    
End Sub
'-------------------------------------------------------------------------------------------------

Private Sub cboPatient_Click()
    Call InitDoctors(cboPatient.ItemData(cboPatient.ListIndex))
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Manage_SelectAllImages       '全选
           If mintPrint = 0 Then Call vfgSelect(1)
        Case conMenu_Manage_UnSelectAllImages      '全清
            If mintPrint = 0 Then Call vfgSelect(0)
        Case conMenu_File_PrintSet        '打印设置
            If mintPrint = 0 Then Call zlPrintSet
        Case conMenu_File_Print           '打印
            If mintPrint = 0 Then Call PrintSelect
        Case conMenu_View_Find            '查询
            If mintPrint = 0 Then Call LoadDataToVfg
        Case ConMenu_pop_Dept
            Label3.Caption = "申请科室↓"
            InitDepts 0
        Case ConMenu_pop_DeptDistrict
            Label3.Caption = "申请病区↓"
            InitDepts 1
        Case conMenu_File_Exit            '退出
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    
    With FraWhere
        .Left = lngLeft + 15
        .Top = lngTop + 15
        .Width = lngRight - lngLeft - 30
    End With
    With vfgData
        .Left = lngLeft + 15
        .Top = FraWhere.Top + FraWhere.Height + 15
        .Width = lngRight - lngLeft - 30
        .Height = chkPrinted.Top - .Top - 30
    End With
    chkPrinted.Left = lngLeft + 15
    chkPrinted.Width = FraWhere.Width

End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID <> conMenu_File_Exit Then Control.Enabled = mintPrint = 0   '打印时，不能执行其他操作！
End Sub


Private Function ShowOpenTree()
    '-----------------------------------------------------------------------------------------
    '功能:打开树型+列表结构的诊疗项目数据
    '返回:出错返回2;成功返回1;取消返回0
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim objPoint As POINTAPI
    
    On Error GoTo ErrHand
    
    strLvw = "编码,1200,0,1;名称,2700,0,0;标本部位,900,0,0"

    ShowOpenTree = 2
    
    strSQL = "Select * " & vbNewLine & _
             "   From (Select Distinct ID, 上级id, 0 As 末级, 编码, 名称, Null + 0 As 标本部位," & vbNewLine & _
             "                          Decode(上级id, Null, ID * Power(10, 20), 上级id * Power(10, 20) + ID) As 排序" & vbNewLine & _
             "          From 诊疗分类目录" & vbNewLine & _
             "          Where 类型 = 5" & vbNewLine & _
             "          Start With ID In (Select Distinct 分类id From 诊疗项目目录 Where 类别 = 'C')" & vbNewLine & _
             "          Connect By Prior 上级id = ID" & vbNewLine & _
             "          Union All" & vbNewLine & _
             "          Select Distinct a.Id, a.分类id As 上级id, 1 As 末级, a.编码, a.名称, a.标本部位, 1 As 排序" & vbNewLine & _
             "          From 诊疗项目目录 A, 检验报告项目 B" & vbNewLine & _
             "          Where a.类别 = 'C' And (a.组合项目 = 1 Or a.单独应用 = 1) And a.Id = b.诊疗项目id(+) And" & vbNewLine & _
             "                (a.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or a.撤档时间 Is Null)) A" & vbNewLine & _
             "   Order By a.末级, a.排序, a.编码"
            
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rs.BOF Then Exit Function
    
    Call ClientToScreen(cmdFind.hWnd, objPoint)
    
    If frmSelectExplorer.ShowSelect(Me, _
                            rs, _
                            objPoint.X * 15 - 30, objPoint.Y * 15 + txtItem.Height - 30, 5400, 2400, _
                            txtItem.Height, _
                            "检验项目树型选择", _
                            strLvw, _
                            "请选择一个检验项目") Then
        
        txtItem.Text = zlCommFun.Nvl(rs("名称").Value) & IIf(rs("标本部位") = "", "", "(" & zlCommFun.Nvl(rs("标本部位").Value) & ")")
        txtItem.Tag = ""
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdFind_Click()
    ShowOpenTree
End Sub

Private Sub Form_Load()
    Dim Menus As New Collection
    '全选  全清 打印 退出
    Menus.Add conMenu_Manage_SelectAllImages & ",全选(&A),False"
    Menus.Add conMenu_Manage_UnSelectAllImages & ",全清(&U),False"
    Menus.Add conMenu_File_PrintSet & ",设置(&P),True"
    Menus.Add conMenu_File_Print & ",打印(&P),False"
    Menus.Add conMenu_View_Find & ",查询(&F),True"
    Menus.Add conMenu_File_Exit & ",退出(&Q),True"
    
    Call CbsButtonInit(cbsMain, Menus, True, xtpBarTop)
    Set Menus = Nothing

    Call LoadBaseData
    mintPrint = 0
    
End Sub

Private Sub Form_Resize()
    Me.chkPrinted.Top = Me.ScaleHeight - Me.chkPrinted.Height - 15
    Call cbsMain_Resize
End Sub



Private Sub Label3_Click()
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    Dim vPoint As POINTAPI
    On Error Resume Next

    Set objPopup = Me.cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_Dept, "申请科室")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_DeptDistrict, "申请病区")
    End With
    vPoint.X = Label3.Left / Screen.TwipsPerPixelX
    vPoint.Y = (Label3.Top + Label3.Height + 30) / Screen.TwipsPerPixelY
    ClientToScreen FraWhere.hWnd, vPoint

    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
End Sub

Private Function InitDepts(intDeptView As Integer, Optional strErr As String) As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strDeptIDs As String, lngPreDept As Long

    If cboPatient.ListIndex <> -1 Then
        lngPreDept = cboPatient.ItemData(cboPatient.ListIndex)
    End If

    On Error GoTo errH

    If intDeptView = 0 Then
        '按科室读取显示
        '包含门急诊观察室的病人还没有上床，不加只显床上有病人的科室的限制
        If InStr(mstrPrivs, "全院病人") > 0 Then

            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where B.部门ID=A.ID And B.工作性质='临床'" & _
                " And (B.服务对象 IN(2,3) Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编码"
        Else
            '求有权限的科室：本身所在科室+所属病区包含的科室
            strSQL = _
                " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
                " From 部门表 A,部门性质说明 B,部门人员 C" & _
                " Where B.部门ID=A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
                " And (B.服务对象 IN(2,3) Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And B.工作性质='临床'"
            strSQL = strSQL & " Union " & _
                " Select C.ID,C.编码,C.名称,Nvl(A.缺省,0) As 缺省" & _
                " From 部门人员 A,病区科室对应 B,部门表 C" & _
                " Where A.部门ID=B.病区ID And B.科室ID=C.ID And A.人员ID=[1]" & _
                " And Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=B.病区ID)" & _
                " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=B.病区ID)" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)"
            If InStr(mstrPrivs, "ICU病人") > 0 Then
                strSQL = strSQL & " Union " & _
                    " Select A.ID,A.编码,A.名称,0 As 缺省" & _
                    " From 部门表 A" & _
                    " Where Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='ICU')" & _
                    " And Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='临床')" & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
            End If
            strSQL = "Select ID,编码,名称,Max(缺省) As 缺省 From (" & strSQL & ") Group By ID,编码,名称 Order by 编码"
        End If
    Else
        '按病区读取显示
        If InStr(mstrPrivs, "全院病人") > 0 Then

            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称" & _
                " From 部门表 A,部门性质说明 B " & _
                " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
                " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编码"
        Else
            '求有权病区：直接所在病区+所在科室所属病区
            strSQL = _
                " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
                " From 部门表 A,部门性质说明 B,部门人员 C" & _
                " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
                " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
                " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
            strSQL = strSQL & " Union " & _
                " Select C.ID,C.编码,C.名称,Nvl(A.缺省,0) as 缺省" & _
                " From 部门人员 A,病区科室对应 B,部门表 C" & _
                " Where A.部门ID=B.科室ID And B.病区ID=C.ID And A.人员ID=[1]" & _
                " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=B.科室ID)" & _
                " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=B.科室ID)" & _
                " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)"
            If InStr(mstrPrivs, "ICU病人") > 0 Then
                strSQL = strSQL & " Union " & _
                    " Select A.ID,A.编码,A.名称,0 As 缺省" & _
                    " From 部门表 A" & _
                    " Where Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='ICU')" & _
                    " And Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='护理')" & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
            End If
            strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
        End If
    End If

    cboPatient.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    If intDeptView = 0 Then
        cboPatient.AddItem "<所有科室>"
    Else
        cboPatient.AddItem "<所有病区>"
    End If
    
    For i = 1 To rsTmp.RecordCount
        cboPatient.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboPatient.ItemData(cboPatient.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next
    If rsTmp.RecordCount > 0 Then
        cboPatient.ListIndex = 0
    End If
    InitDepts = True
    Exit Function
errH:
    strErr = "出错函数(GetSampleValCount),出错信息:" & Err.Number & " " & Err.Description
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Function


Private Sub lbldate_Click()
    PopupMenu mnudate
End Sub

Private Sub lblNo_Click()
    PopupMenu mnuPop
End Sub

Private Sub mnuAdjdate_Click()
    mnuAdjdate.Checked = Not mnuAdjdate.Checked
    If mnuAdjdate.Checked = True Then
        lbldate.Caption = "审核日期↓"
        mnuAppdate.Checked = False
    Else
        mnuAppdate.Checked = True
        lbldate.Caption = "申请日期↓"
    End If
End Sub

Private Sub mnuAppdate_Click()
    mnuAppdate.Checked = Not mnuAppdate.Checked
    If mnuAppdate.Checked = True Then
        lbldate.Caption = "申请日期↓"
        mnuAdjdate.Checked = False
    Else
        mnuAdjdate.Checked = True
        lbldate.Caption = "审核日期↓"
    End If
End Sub

Private Sub mnuBedNo_Click()
    mnuBedNo.Checked = Not mnuBedNo.Checked
    If mnuBedNo.Checked = True Then
        lblNo.Caption = "床号↓"
        mnuPatiNo.Checked = False
    Else
        mnuPatiNo.Checked = True
        lblNo.Caption = "住院号↓"
    End If
End Sub

Private Sub mnuPatiNo_Click()
    mnuPatiNo.Checked = Not mnuPatiNo.Checked
    If mnuPatiNo.Checked = True Then
        lblNo.Caption = "住院号↓"
        mnuBedNo.Checked = False
    Else
        mnuBedNo.Checked = True
        lblNo.Caption = "床号↓"
    End If
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OpenSelect (txtItem)
    End If
End Sub

Private Function OpenSelect(ByVal strText As String)
    '-----------------------------------------------------------------------------------------
    '功能:打开列表结构的诊疗项目数据
    '返回:出错返回2;成功返回1;取消返回0
    '-----------------------------------------------------------------------------------------
    Dim strInput As String
    Dim rs As New ADODB.Recordset
    Dim strLvw As String
    Dim objPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    strLvw = "编码,900,0,1;检验项目,3600,0,0;标本部位,900,0,0"
    
    strInput = "%" & UCase(strText) & "%"
    strSQL = " Select Distinct a.Id, a.编码, a.名称 as 检验项目, a.标本部位" & vbNewLine & _
             "    From 诊疗项目目录 A, 检验报告项目 B" & vbNewLine & _
             "    Where a.类别 = 'C' And (a.组合项目 = 1 Or a.单独应用 = 1) And a.Id = b.诊疗项目id(+) And " & vbNewLine & _
             "          (a.编码 Like [1] Or a.名称 Like [1] Or" & vbNewLine & _
             "          a.Id In (Select 诊疗项目id From 诊疗项目别名 Where (名称 Like [1] Or Upper(简码) Like Upper([1]))))" & vbNewLine & _
             "    Order By a.编码"
             
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput)
    If rs.BOF Then
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then GoTo Over
            
    Call ClientToScreen(txtItem.hWnd, objPoint)
    If frmSelectList.ShowSelect(Me, rs, strLvw, objPoint.X * 15 - 30, objPoint.Y * 15 + txtItem.Height - 30, 6000, 4200, Me.Name & "\检验项目选择", "请从下表中选择一个项目") Then
        GoTo Over
    End If
    Exit Function
Over:
    txtItem.Text = zlCommFun.Nvl(rs("检验项目").Value) & IIf(rs("标本部位") = "", "", "(" & zlCommFun.Nvl(rs("标本部位").Value) & ")")
    txtItem.Tag = ""
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub vfgData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iRow As Integer, iCol As Integer
    If Button = 1 Then
        With vfgData
            iCol = .MouseCol: iRow = .MouseRow
            
            If iCol = Col.选择 And iRow >= .FixedRows And iRow <= .Rows - 1 Then
                If Val(.TextMatrix(iRow, Col.选择)) = 0 Then
                    .TextMatrix(iRow, Col.选择) = 1
                Else
                    .TextMatrix(iRow, Col.选择) = 0
                End If
            End If
        End With
    End If
End Sub

