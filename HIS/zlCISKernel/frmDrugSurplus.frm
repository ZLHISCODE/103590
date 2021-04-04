VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugSurplus 
   AutoRedraw      =   -1  'True
   Caption         =   "药品留存登记"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10830
   Icon            =   "frmDrugSurplus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic条件 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   195
      ScaleHeight     =   2865
      ScaleWidth      =   2415
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2415
      Begin VB.ComboBox cbo比例 
         Height          =   300
         ItemData        =   "frmDrugSurplus.frx":058A
         Left            =   1380
         List            =   "frmDrugSurplus.frx":0597
         TabIndex        =   7
         Text            =   "50%"
         Top             =   1907
         Width           =   765
      End
      Begin VB.OptionButton optRule 
         Caption         =   "比例留存"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   6
         Top             =   1960
         Width           =   1380
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "提取待发药品(&L)"
         Height          =   350
         Left            =   870
         TabIndex        =   9
         Top             =   2535
         Width           =   1560
      End
      Begin VB.OptionButton optRule 
         Caption         =   "整体分零满足"
         Height          =   195
         Index           =   3
         Left            =   315
         TabIndex        =   8
         Top             =   2250
         Width           =   1380
      End
      Begin VB.OptionButton optRule 
         Caption         =   "全部留存"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   5
         Top             =   1670
         Width           =   1380
      End
      Begin VB.OptionButton optRule 
         Caption         =   "不处理"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   4
         Top             =   1380
         Value           =   -1  'True
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   345
         TabIndex        =   2
         Top             =   255
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   122617859
         CurrentDate     =   39610
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   345
         TabIndex        =   3
         Top             =   615
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   122617859
         CurrentDate     =   39610
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省留存计算："
         Height          =   180
         Left            =   105
         TabIndex        =   17
         Top             =   1095
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药品发送时间："
         Height          =   180
         Left            =   105
         TabIndex        =   16
         Top             =   15
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   105
         TabIndex        =   14
         Top             =   675
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从"
         Height          =   180
         Left            =   105
         TabIndex        =   13
         Top             =   315
         Width           =   180
      End
   End
   Begin VB.Frame fraLR 
      BorderStyle     =   0  'None
      Height          =   6345
      Left            =   3090
      MousePointer    =   9  'Size W E
      TabIndex        =   20
      Top             =   600
      Width           =   45
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   195
      ScaleHeight     =   990
      ScaleWidth      =   2415
      TabIndex        =   15
      Top             =   6150
      Width           =   2415
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   1335
         TabIndex        =   11
         ToolTipText     =   "F3"
         Top             =   660
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   105
         TabIndex        =   10
         Top             =   285
         Width           =   2310
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药品编码、简码、名称："
         Height          =   180
         Left            =   105
         TabIndex        =   18
         Top             =   60
         Width           =   1980
      End
   End
   Begin VB.ComboBox cbo药房 
      Height          =   300
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   870
      Width           =   2100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDrug 
      Height          =   6300
      Left            =   3195
      TabIndex        =   0
      Top             =   600
      Width           =   7560
      _cx             =   13335
      _cy             =   11112
      Appearance      =   2
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   11
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugSurplus.frx":05AA
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   115
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
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   6600
      Left            =   90
      TabIndex        =   19
      Top             =   645
      Width           =   3000
      _Version        =   589884
      _ExtentX        =   5292
      _ExtentY        =   11642
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   405
      Top             =   210
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDrugSurplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病区ID As Long

Private mrsDrug As ADODB.Recordset '待发药缓存
Private mrsApply As ADODB.Recordset '退药情况缓存
Private mrsAdvice As ADODB.Recordset '医保信息缓存

Private mblnReturn As Boolean
Private mlngPe药房 As Long
Private mstrLike As String
Private mint简码 As Integer
Private mblnChange As Boolean
Private Enum COL_DRUG
    col编码 = 0
    col药品 = 1
    col规格 = 2
    col产地 = 3
    col单位 = 4
    col应发数 = 5
    col申请退药数 = 6
    col审核退药数 = 7
    col留存数 = 8
    col类别 = 9
    col住院包装 = 10
End Enum
Private mstr给药IDs As String

Public Sub ShowMe(frmParent As Object, ByVal lng病区ID As Long)
    mlng病区ID = lng病区ID
    
    On Error Resume Next
    Me.Show , frmParent
End Sub

Private Sub cbo比例_GotFocus()
    Call zlControl.TxtSelAll(cbo比例)
End Sub

Private Sub cbo比例_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cbo比例_Validate(blnCancel)
        If Not blnCancel Then Call cmdLoad_Click
    Else
        If InStr("0123456789%" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub cbo比例_Validate(Cancel As Boolean)
    If Val(cbo比例.Text) < 0 Or Val(cbo比例.Text) > 100 Then
        Cancel = True
    Else
        cbo比例.Text = Val(cbo比例.Text) & "%"
    End If
End Sub

Private Sub cbo药房_Click()
    If cbo药房.ListIndex <> -1 Then
        If cbo药房.ListIndex = mlngPe药房 Then Exit Sub
        If mblnChange Then
            If MsgBox("当前的数据没有保存，确实要切换药房吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Call Cbo.SetIndex(cbo药房.Hwnd, mlngPe药房)
                Exit Sub
            End If
        End If
        If Not CheckDate Then
            Call Cbo.SetIndex(cbo药房.Hwnd, mlngPe药房)
            dtpBegin.SetFocus: Exit Sub
        End If
        
        mlngPe药房 = cbo药房.ListIndex
        Call ReleaseRecord(mrsDrug)
        Call ReleaseRecord(mrsAdvice)
        Call ReleaseRecord(mrsApply)
        
        Call LoadSurplus
        
        mblnChange = False
        If Me.Visible Then vsDrug.SetFocus
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_Edit_Save
        Call SaveData
    Case conMenu_File_Print
        Call OutputList(1)
    Case conMenu_File_Preview
        Call OutputList(2)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    Me.tkpMain.Left = lngLeft
    Me.tkpMain.Top = lngTop
    Me.tkpMain.Height = lngBottom - lngTop
    
    Me.fraLR.Left = lngLeft + tkpMain.Width
    Me.fraLR.Top = lngTop
    Me.fraLR.Height = lngBottom - lngTop

    Me.vsDrug.Left = fraLR.Left + fraLR.Width
    Me.vsDrug.Top = lngTop
    Me.vsDrug.Width = lngRight - lngLeft - fraLR.Width - tkpMain.Width
    Me.vsDrug.Height = lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Save
        Control.Enabled = mblnChange
    End Select
End Sub

Private Sub cmdFind_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strInput As String, strMatch As String
    Dim strSQL As String, i As Long
    Dim blnFirst As Boolean
    
    If vsDrug.Rows = vsDrug.FixedRows + 1 And vsDrug.RowData(vsDrug.Row) = 0 Then
        MsgBox "找不到匹配的药品。"
        txtFind.SetFocus: Exit Sub
    End If
    
    If txtFind.Tag = "" Then
        '不同的输入匹配方式
        strInput = UCase(txtFind.Text)
        strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.名称 Like [2] And C.码类=[3] Or C.简码 Like [2] And C.码类 IN([3],3))"
        If IsNumeric(strInput) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
            If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.简码 Like [2] And C.码类=3)"
        ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.输入全是字母时只匹配简码
            If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.简码 Like [2] And C.码类=[3]"
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strMatch = " And C.名称 Like [2] And C.码类=[3]"
        End If
        
        strSQL = _
            " Select Distinct A.ID" & _
            " From 收费项目目录 A,收费项目别名 C" & _
            " Where (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And A.服务对象 IN(2,3) And A.ID=C.收费细目ID And A.类别 IN('5','6','7')" & strMatch
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint简码 + 1)
        
        strSQL = "0"
        Do While Not rsTmp.EOF
            strSQL = strSQL & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        txtFind.Tag = strSQL
        
        blnFirst = True
    End If
    
    If txtFind.Tag = "0" Then
        MsgBox "找不到匹配的药品。"
        txtFind.SetFocus: Exit Sub
    End If
    
    With vsDrug
        For i = IIF(blnFirst, 1, .Row + 1) To .Rows - 1
            If .RowData(i) <> 0 And InStr("," & txtFind.Tag & ",", "," & .RowData(i) & ",") > 0 Then
                .Row = i: Call .ShowCell(i, .Col): .SetFocus: Exit For
            End If
        Next
        If i > .Rows - 1 Then
            MsgBox "找不到匹配的药品。"
            txtFind.SetFocus: Exit Sub
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveData
End Sub

Private Function CheckDate() As Boolean
    If dtpBegin.value > dtpEnd.value Then
        MsgBox "开始时间应该比结束时间小。", vbInformation, gstrSysName
        Exit Function
    End If
    If DateDiff("d", dtpBegin.value, dtpEnd.value) >= 7 And (dtpBegin.Tag <> "" Or dtpEnd.Tag <> "") Then
        If MsgBox("设置的时间范围太大，可能引起系统查询缓慢，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
        dtpBegin.Tag = "": dtpEnd.Tag = ""
    End If
    CheckDate = True
End Function

Private Sub cmdLoad_Click()
    Dim arrData As Variant, strData As String
    Dim lngRow As Long, i As Long
    Dim sng申请数 As Single, sng审核数 As Single
    
    If Not CheckDate Then dtpBegin.SetFocus: Exit Sub
    
    If Not (vsDrug.RowData(vsDrug.Row) = 0 And vsDrug.Row = vsDrug.Rows - 1 And vsDrug.Rows = vsDrug.FixedRows + 1) Then
        If MsgBox("确实要提取所有待发的药品吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    'If mrsDrug Is Nothing Then Call LoadDrugPut
    Call LoadDrugPut
    
    Screen.MousePointer = 11
    
    With vsDrug
        '记录原有留存数
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, col留存数)) > 0 Then
                strData = strData & ";" & .RowData(i) & "," & Val(.TextMatrix(i, col留存数))
            End If
        Next
        strData = Mid(strData, 2)
        
        '装入待发药品
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        mrsDrug.Filter = 0
        If Not mrsDrug.EOF Then
            .Rows = .FixedRows + mrsDrug.RecordCount
            For i = .FixedRows To .FixedRows + mrsDrug.RecordCount - 1
                .RowData(i) = Val(mrsDrug!药品ID)
                .TextMatrix(i, col编码) = Nvl(mrsDrug!编码)
                .TextMatrix(i, col药品) = Nvl(mrsDrug!名称)
                .TextMatrix(i, col规格) = Nvl(mrsDrug!规格)
                .TextMatrix(i, col产地) = Nvl(mrsDrug!产地)
                .TextMatrix(i, col单位) = Nvl(mrsDrug!单位)
                .TextMatrix(i, col类别) = Nvl(mrsDrug!类别)
                .TextMatrix(i, col住院包装) = Nvl(mrsDrug!住院包装, 0)
                
                .TextMatrix(i, col应发数) = FormatEx(Nvl(mrsDrug!数量, 0), 5)
                Call GetDrugApply(mrsDrug!药品ID, sng申请数, sng审核数)
                .TextMatrix(i, col申请退药数) = sng申请数
                .TextMatrix(i, col审核退药数) = sng审核数
                .TextMatrix(i, col留存数) = GetSurplus(i)
                
                mrsDrug.MoveNext
            Next
        Else
            .Rows = .FixedRows + 1
        End If
        
        '保留之前的留存数
        arrData = Split(strData, ";")
        For i = 0 To UBound(arrData)
            lngRow = .FindRow(Val(Split(arrData(i), ",")(0)))
            If lngRow <> -1 Then
                If Val(.TextMatrix(lngRow, col留存数)) = 0 Then
                    .TextMatrix(lngRow, col留存数) = Val(Split(arrData(i), ",")(1))
                    
                    '原输入的留存数比现在的应发数大，粗体显示
                    If Val(.TextMatrix(lngRow, col留存数)) > Val(.TextMatrix(lngRow, col应发数)) Then
                        .TextMatrix(lngRow, col留存数) = Val(.TextMatrix(lngRow, col应发数))
                        .Cell(flexcpFontBold, lngRow, col留存数) = True
                    End If
                End If
            End If
        Next
        
        .Row = .FixedRows
        .Col = IIF(.RowData(.Row) = 0, col药品, col留存数)
        Call vsDrug_AfterRowColChange(-1, -1, .Row, .Col)
        Call .ShowCell(.Row, .Col)
        
        .Redraw = flexRDDirect
    End With
    
    Screen.MousePointer = 0
    
    mblnChange = True
    
    vsDrug.SetFocus
End Sub

Private Sub dtpBegin_Change()
    dtpBegin.Tag = "1"
    Call ReleaseRecord(mrsDrug)
    Call ReleaseRecord(mrsAdvice)
    Call ReleaseRecord(mrsApply)
End Sub

Private Sub dtpEnd_Change()
    dtpEnd.Tag = "1"
    Call ReleaseRecord(mrsDrug)
    Call ReleaseRecord(mrsAdvice)
    Call ReleaseRecord(mrsApply)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If cmdFind.Enabled And cmdFind.Visible Then
            cmdFind_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objPane As Pane
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng药房ID As Long
    Dim i As Long
    
    For i = 0 To vsDrug.Cols - 1
        If vsDrug.ColHidden(i) Then
            vsDrug.ColWidth(i) = 0 '为支持PrintMode
        Else
            vsDrug.MergeCol(i) = True
        End If
    Next
    vsDrug.MergeRow(0) = True
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_Edit_Save
        .Add 0, vbKeyF1, conMenu_Help_Help
        .Add FALT, vbKeyX, conMenu_File_Exit
    End With
    
    '分组控件------------------------------------------
    Call tkpMain.SetMargins(8, 8, 8, 8, 8)
    
    Set objGroup = tkpMain.Groups.Add(0, "留存药房")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = cbo药房
    
    Set objGroup = tkpMain.Groups.Add(0, "待发药品")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = pic条件
    pic条件.BackColor = objItem.BackColor
    optRule(0).BackColor = objItem.BackColor
    optRule(1).BackColor = objItem.BackColor
    optRule(2).BackColor = objItem.BackColor
    optRule(3).BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(0, "查找")
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picFind
    picFind.BackColor = objItem.BackColor
    
    '数据初始-------------------------------------------
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    
    dtpEnd.value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.value = Format(dtpEnd.value, "yyyy-MM-dd 00:00:00")
    
    cbo比例.Text = Val(zlDatabase.GetPara("缺省留存比例", glngSys, p住院医嘱发送, "50", Array(cbo比例))) & "%"
    If Not cbo比例.Enabled Then
        cbo比例.Tag = "1" '标识固定不可用
    Else
        cbo比例.Enabled = False '缺省选项应是不可用
    End If
    
    optRule(Val(zlDatabase.GetPara("缺省留存计算", glngSys, p住院医嘱发送, "0", Array(optRule(0), optRule(1), optRule(2), optRule(3))))).value = True
    
    mstr给药IDs = zlDatabase.GetPara("留存登记给药途径限制", glngSys, p住院医嘱发送)
    
    '住院药房
    lng药房ID = Val(zlDatabase.GetPara("缺省留存药房", glngSys, p住院医嘱发送, , Array(cbo药房)))
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " AND B.部门ID=A.ID And B.服务对象 IN(2,3) and B.工作性质 in('中药房','西药房','成药房')" & _
        " Order by A.编码"
    On Error GoTo errH
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        cbo药房.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cbo药房.ItemData(cbo药房.NewIndex) = rsTmp!ID
        If rsTmp!ID = lng药房ID Then
            Call Cbo.SetIndex(cbo药房.Hwnd, cbo药房.NewIndex)
        End If
        rsTmp.MoveNext
    Loop
    If cbo药房.ListCount = 0 Then
        MsgBox "没有可用的住院药房，请先到部门管理中进行设置。", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    If cbo药房.ListIndex = -1 Then
        Call Cbo.SetIndex(cbo药房.Hwnd, 0)
    End If
    
    
    mlngPe药房 = -1
    mblnChange = False
    Call cbo药房_Click

    '-------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("当前的数据没有保存，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    
    If cbo药房.ListIndex <> -1 Then
        Call zlDatabase.SetPara("缺省留存药房", cbo药房.ItemData(cbo药房.ListIndex), glngSys, p住院医嘱发送)
    End If
    Call zlDatabase.SetPara("缺省留存计算", IIF(optRule(0).value, 0, IIF(optRule(1).value, 1, IIF(optRule(2).value, 2, 3))), glngSys, p住院医嘱发送)
    Call zlDatabase.SetPara("缺省留存比例", Val(cbo比例.Text), glngSys, p住院医嘱发送)
    
    Call SaveWinState(Me, App.ProductName)
    
    Call ReleaseRecord(mrsDrug)
    Call ReleaseRecord(mrsAdvice)
    Call ReleaseRecord(mrsApply)
End Sub

Private Sub ReleaseRecord(rsData As ADODB.Recordset)
    If Not rsData Is Nothing Then
        If rsData.State = 1 Then rsData.Close
        Set rsData = Nothing
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        If tkpMain.Width + x < 2700 Or vsDrug.Width - x < 3000 Then Exit Sub
        fraLR.Left = fraLR.Left + x
        tkpMain.Width = tkpMain.Width + x
        vsDrug.Left = vsDrug.Left + x
        vsDrug.Width = vsDrug.Width - x
        Me.Refresh
    End If
End Sub

Private Sub optRule_Click(Index As Integer)
    cbo比例.Enabled = optRule(2).value And cbo比例.Tag = ""
End Sub

Private Sub picFind_Resize()
    On Error Resume Next
    
    txtFind.Width = picFind.ScaleWidth - txtFind.Left
    cmdFind.Left = picFind.ScaleWidth - cmdFind.Width + 15
End Sub

Private Sub pic条件_Resize()
    On Error Resume Next
    
    dtpBegin.Width = pic条件.ScaleWidth - dtpBegin.Left
    dtpEnd.Width = pic条件.ScaleWidth - dtpEnd.Left
    cmdLoad.Left = pic条件.ScaleWidth - cmdLoad.Width + 15
End Sub

Private Sub txtFind_Change()
    txtFind.Tag = ""
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        zlControl.TxtSelAll txtFind
        Call cmdFind_Click
    End If
End Sub

Private Sub vsDrug_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        If vsDrug.Redraw <> flexRDNone Then
            If OldRow <> -1 And OldRow <= vsDrug.Rows - 1 Then
                vsDrug.Cell(flexcpForeColor, OldRow, 0, OldRow, vsDrug.Cols - 1) = vsDrug.ForeColor
            End If
            If NewRow <> -1 Then
                vsDrug.Cell(flexcpForeColor, NewRow, 0, NewRow, vsDrug.Cols - 1) = vbBlue
            End If
        End If
    End If
    
    If CellEditable(NewRow, NewCol) Then
        vsDrug.FocusRect = flexFocusSolid
        If NewCol = col药品 Then
            vsDrug.ComboList = "..."
        Else
            vsDrug.ComboList = ""
        End If
    Else
        vsDrug.ComboList = ""
        vsDrug.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vsDrug_AfterSort(ByVal Col As Long, Order As Integer)
    vsDrug.Cell(flexcpForeColor, vsDrug.Row, 0, vsDrug.Row, vsDrug.Cols - 1) = vbBlue
End Sub

Private Sub vsDrug_BeforeSort(ByVal Col As Long, Order As Integer)
    vsDrug.Cell(flexcpForeColor, vsDrug.Row, 0, vsDrug.Row, vsDrug.Cols - 1) = vsDrug.ForeColor
End Sub

Private Sub vsDrug_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnCancel As Boolean
    Dim str性质 As String, str类型 As String, str类别 As String

    If Col = col药品 Then
        str性质 = Get部门性质(cbo药房.ItemData(cbo药房.ListIndex))
        str类型 = " And 类型 IN(1,2,3)"
        str类别 = " And A.类别 IN('5','6','7')"
        If InStr(str性质, "西药房") = 0 Then
            str类型 = Replace(str类型, "1,", "")
            str类别 = Replace(str类别, "'5',", "")
        End If
        If InStr(str性质, "成药房") = 0 Then
            str类型 = Replace(str类型, "2,", "")
            str类别 = Replace(str类别, "'6',", "")
        End If
        If InStr(str性质, "中药房") = 0 Then
            str类型 = Replace(str类型, ",3", "")
            str类别 = Replace(str类别, ",'7'", "")
        End If
        
        strSQL = _
            " Select Distinct 0 as 末级,To_Number('999999999'||类型) as ID,-NULL as 上级ID," & _
            " CHR(13)||类型 as 编码,Decode(类型,1,'西成药',2,'中成药',3,'中草药') as 名称," & _
            " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别ID,-NULL as 系数ID" & _
            " From 诊疗分类目录 Where (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & str类型
        strSQL = strSQL & " Union ALL " & _
            " Select 0 as 末级,-1*ID as ID,Nvl(-1*上级ID,To_Number('999999999'||类型)) as 上级ID," & _
            " 编码,名称,NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别ID,-NULL as 系数ID" & _
            " From 诊疗分类目录 Where (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & str类型 & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID"
        strSQL = strSQL & " Union ALL " & _
            " Select Distinct 1 as 末级,A.ID,-1*E.分类ID as 上级ID,A.编码," & _
            " Nvl(F.名称,A.名称) as 名称,D.住院单位 as 单位,A.规格,A.产地,A.类别 as 类别ID,D.住院包装 as 系数ID" & _
            " From 收费项目目录 A,药品规格 D,诊疗项目目录 E,收费项目别名 F" & _
            " Where (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And A.服务对象 IN(2,3) And A.ID=D.药品ID And D.药名ID=E.ID" & _
            " And A.ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=[1]" & str类别
        strSQL = strSQL & " Order by 编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "药品", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, IIF(gbyt药品名称显示 = 0, 1, 3))
        If Not rsTmp Is Nothing Then
            If SetItemInput(Row, rsTmp) Then Call EnterNextCell(Row, Col)
        Else
            If Not blnCancel Then
                MsgBox "没有可用的药品，请先到药品目录管理中设置！", vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function SetItemInput(ByVal lngRow As Long, rsTmp As ADODB.Recordset) As Boolean
    Dim lngFind As Long
    Dim sng申请数 As Single, sng审核数 As Single
    
    With vsDrug
        lngFind = .FindRow(Val(rsTmp!ID))
        If lngFind <> -1 And lngFind <> lngRow Then
            MsgBox "药品""" & rsTmp!名称 & """已经录入。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '根据给药途径限制待发药品进行限制
        If mstr给药IDs <> "" Then
            If mrsDrug Is Nothing Then Call LoadDrugPut
            mrsDrug.Filter = "药品ID=" & rsTmp!ID
            If mrsDrug.EOF Then
                MsgBox "当前指定给药途径等条件下药品""" & rsTmp!名称 & """没有待发药记录。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        .RowData(lngRow) = Val(rsTmp!ID)
        .TextMatrix(lngRow, col编码) = Nvl(rsTmp!编码)
        .TextMatrix(lngRow, col药品) = Nvl(rsTmp!名称)
        .TextMatrix(lngRow, col规格) = Nvl(rsTmp!规格)
        .TextMatrix(lngRow, col产地) = Nvl(rsTmp!产地)
        .TextMatrix(lngRow, col单位) = Nvl(rsTmp!单位)
        .TextMatrix(lngRow, col类别) = Nvl(rsTmp!类别ID)
        .TextMatrix(lngRow, col住院包装) = Nvl(rsTmp!系数ID, 0)
        
        .TextMatrix(lngRow, col应发数) = GetDrugPut(rsTmp!ID)
        Call GetDrugApply(rsTmp!ID, sng申请数, sng审核数)
        .TextMatrix(lngRow, col申请退药数) = sng申请数
        .TextMatrix(lngRow, col审核退药数) = sng审核数
        .TextMatrix(lngRow, col留存数) = GetSurplus(lngRow)
        
        .Cell(flexcpFontBold, lngRow, col留存数) = False
        If Val(.TextMatrix(lngRow, col留存数)) > 0 Then mblnChange = True
    End With
    
    SetItemInput = True
End Function

Private Sub vsDrug_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsDrug
        '有留存的颜色标记
        If Col = col留存数 Then
            If Val(.TextMatrix(Row, Col)) > 0 Then
                .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = &HC0FFFF
            Else
                .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = .BackColor
            End If
        End If
    End With
End Sub

Private Sub vsDrug_DblClick()
    Call vsDrug_KeyPress(32)
End Sub

Private Sub vsDrug_GotFocus()
    '该窗体中从其他控件焦点切换过来才会激活
    If Not CheckDate Then dtpBegin.SetFocus
End Sub

Private Sub vsDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsDrug
        If KeyCode = vbKeyDelete Then
            If vsDrug.Col = col留存数 Then
                If vsDrug.TextMatrix(vsDrug.Row, vsDrug.Col) <> "" Then
                    vsDrug.TextMatrix(vsDrug.Row, vsDrug.Col) = ""
                    vsDrug.CellFontBold = False
                    mblnChange = True
                End If
            Else
                If .RowData(.Row) <> 0 Then
                    If MsgBox("确定要删除当前药品行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    If Val(.TextMatrix(.Row, col留存数)) <> 0 Then mblnChange = True
                End If
                
                .RemoveItem .Row
    
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = col药品
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDrug_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDrug_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call EnterNextCell(vsDrug.Row, vsDrug.Col)
    Else
        If vsDrug.Col = col药品 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsDrug_CellButtonClick(vsDrug.Row, vsDrug.Col)
            Else
                vsDrug.ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End If
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
    If lngCol < col药品 Then
        vsDrug.Col = col药品
    ElseIf lngCol < col留存数 Then
        If vsDrug.RowData(lngRow) <> 0 Then
            vsDrug.Col = col留存数
        End If
    ElseIf vsDrug.RowData(lngRow) <> 0 Then
        If lngRow = vsDrug.Rows - 1 Then vsDrug.AddItem ""
        vsDrug.Row = vsDrug.Row + 1
        vsDrug.Col = col药品
    End If
    vsDrug.ShowCell vsDrug.Row, vsDrug.Col
End Sub

Private Sub vsDrug_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
        If Col = col留存数 Then
            If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub vsDrug_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDrug.EditSelStart = 0
    vsDrug.EditSelLength = zlCommFun.ActualLen(vsDrug.EditText)
End Sub

Private Sub vsDrug_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not CellEditable(Row, Col) Then Cancel = True
    If Col = col留存数 Then
        vsDrug.EditMaxLength = 10
    Else
        vsDrug.EditMaxLength = 0
    End If
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    CellEditable = True
    If Not (lngCol = col药品 Or lngCol = col留存数) Then
        CellEditable = False
    ElseIf lngCol = col留存数 And vsDrug.RowData(lngRow) = 0 Then
        CellEditable = False
    End If
End Function

Private Function LoadSurplus() As Boolean
'功能：读取当前药房已填写的留存记录
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lng药房ID As Long
    Dim sng申请数 As Single, sng审核数 As Single
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    If cbo药房.ListIndex <> -1 Then
        lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
    End If
    strSQL = "Select A.药品ID,C.编码,Nvl(D.名称,C.名称) as 名称,C.规格,C.产地," & _
        " B.住院单位 as 单位,A.留存数量/Nvl(B.住院包装,1) as 留存数量,C.类别,B.住院包装" & _
        " From 药品留存计划 A,药品规格 B,收费项目目录 C,收费项目别名 D" & _
        " Where A.药品ID=B.药品ID And A.药品ID=C.ID And A.部门ID=[1] And A.库房ID=[2]" & _
        " And A.状态=0 And C.ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=[3]" & _
        " Order by C.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病区ID, lng药房ID, IIF(gbyt药品名称显示 = 0, 1, 3))
    
    With vsDrug
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = .FixedRows To .FixedRows + rsTmp.RecordCount - 1
                .RowData(i) = Val(rsTmp!药品ID)
                .TextMatrix(i, col编码) = Nvl(rsTmp!编码)
                .TextMatrix(i, col药品) = Nvl(rsTmp!名称)
                .TextMatrix(i, col规格) = Nvl(rsTmp!规格)
                .TextMatrix(i, col产地) = Nvl(rsTmp!产地)
                .TextMatrix(i, col单位) = Nvl(rsTmp!单位)
                .TextMatrix(i, col类别) = Nvl(rsTmp!类别)
                .TextMatrix(i, col住院包装) = Nvl(rsTmp!住院包装, 0)
                
                .TextMatrix(i, col应发数) = GetDrugPut(rsTmp!药品ID) '当前应发数
                Call GetDrugApply(rsTmp!药品ID, sng申请数, sng审核数)
                .TextMatrix(i, col申请退药数) = sng申请数
                .TextMatrix(i, col审核退药数) = sng审核数
                .TextMatrix(i, col留存数) = Nvl(rsTmp!留存数量)
                
                '原输入的留存数比现在的应发数大，粗体显示
                If Val(.TextMatrix(i, col留存数)) > Val(.TextMatrix(i, col应发数)) Then
                    .TextMatrix(i, col留存数) = Val(.TextMatrix(i, col应发数))
                    .Cell(flexcpFontBold, i, col留存数) = True
                End If
                
                rsTmp.MoveNext
            Next
        Else
            .Rows = .FixedRows + 1
        End If
        
        .Row = .FixedRows
        .Col = IIF(.RowData(.Row) = 0, col药品, col留存数)
        Call vsDrug_AfterRowColChange(-1, -1, .Row, .Col)
        Call .ShowCell(.Row, .Col)
    End With
    
    Screen.MousePointer = 0
    LoadSurplus = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    Dim arrSQL As Variant, i As Long
    
    arrSQL = Array()
    With vsDrug
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_药品留存计划_Delete(" & mlng病区ID & "," & cbo药房.ItemData(cbo药房.ListIndex) & ")"
        
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 And Val(.TextMatrix(i, col留存数)) > 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_药品留存计划_Insert(" & mlng病区ID & "," & cbo药房.ItemData(cbo药房.ListIndex) & "," & _
                    .RowData(i) & "," & Val(.TextMatrix(i, col留存数)) * Val(.TextMatrix(i, col住院包装)) & ",'" & UserInfo.姓名 & "')"
            End If
        Next
    End With
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    vsDrug.Cell(flexcpFontBold, vsDrug.FixedRows, col留存数, vsDrug.Rows - 1, col留存数) = False
    mblnChange = False
    SaveData = True
    
    Screen.MousePointer = 0
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadDrugApply() As Boolean
'功能：读取退药申请与审核信息
    Dim strSQL As String
    Dim intMouse As Integer
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.收费细目ID as 药品ID," & _
        " Sum(A.数量/Nvl(B.住院包装,1)) as 申请数," & _
        " Sum(Decode(A.状态,1,A.数量/Nvl(B.住院包装,1),0)) as 审核数" & _
        " From 病人费用销帐 A,药品规格 B" & _
        " Where A.收费细目ID=B.药品ID And A.申请时间 Between [1] And [2] And A.申请部门ID=[3] And A.审核部门ID=[4]" & _
        " Group by A.收费细目ID"
    Set mrsApply = New ADODB.Recordset
    Set mrsApply = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(dtpBegin.value), CDate(dtpEnd.value), mlng病区ID, cbo药房.ItemData(cbo药房.ListIndex))
    
    Screen.MousePointer = intMouse
    LoadDrugApply = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call ReleaseRecord(mrsApply)
End Function

Private Function LoadDrugPut() As Boolean
'功能：读取药品未发汇总
    Dim strSQL As String
    Dim str给药 As String
    Dim intMouse As Integer
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 11
        
    On Error GoTo errH
    
    str给药 = mstr给药IDs

    strSQL = _
        " Select /*+ Rule*/ A.药品ID,D.编码,Nvl(D.名称,E.名称) As 名称,C.住院单位 As 单位," & _
        " D.规格,D.产地,Sum(A.填写数量/Nvl(C.住院包装,1)) as 数量,D.类别,C.住院包装" & _
        " From 药品收发记录 A,住院费用记录 B,药品规格 C,收费项目目录 D,收费项目别名 E" & _
        " Where A.单据 = 9 And A.NO = B.NO And B.记录性质 = 2 And A.费用ID = B.ID And A.药品ID = C.药品ID" & _
        " And C.药品ID=D.ID And Mod(A.记录状态,3)=1 And A.审核人 is Null" & _
        " And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=[5]" & _
        IIF(str给药 <> "", " And Nvl(A.用法,'Null') Not IN(Select Column_Value From Table(Cast(f_Str2list([6]) As zlTools.t_Strlist)))", "") & _
        " And A.填制日期 Between [1] And [2] And A.库房ID=[3] And B.病人病区ID=[4]" & _
        " Group By A.药品ID,D.编码,Nvl(D.名称,E.名称),C.住院单位,D.规格,D.产地,D.类别,C.住院包装" & _
        " Order By 编码"
    Set mrsDrug = New ADODB.Recordset
    Set mrsDrug = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(dtpBegin.value), CDate(dtpEnd.value), cbo药房.ItemData(cbo药房.ListIndex), mlng病区ID, IIF(gbyt药品名称显示 = 0, 1, 3), str给药)
    
    Screen.MousePointer = intMouse
    LoadDrugPut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call ReleaseRecord(mrsDrug)
End Function

Private Sub GetDrugApply(ByVal lng药品ID As Long, sng申请数 As Single, sng审核数 As Single)
    sng申请数 = 0: sng审核数 = 0
    
    If mrsApply Is Nothing Then Call LoadDrugApply
    
    mrsApply.Filter = "药品ID=" & lng药品ID
    If Not mrsApply.EOF Then
        sng申请数 = FormatEx(Nvl(mrsApply!申请数, 0), 5)
        sng审核数 = FormatEx(Nvl(mrsApply!审核数, 0), 5)
    End If
End Sub

Private Function GetDrugPut(ByVal lng药品ID As Long) As Single
    If mrsDrug Is Nothing Then Call LoadDrugPut
    
    mrsDrug.Filter = "药品ID=" & lng药品ID
    If Not mrsDrug.EOF Then
        GetDrugPut = FormatEx(Nvl(mrsDrug!数量, 0), 5)
    End If
End Function

Private Sub vsDrug_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strInput As String, strMatch As String
    Dim vPoint As PointAPI, blnCancel As Boolean
    Dim str性质 As String, str类别 As String
    
    With vsDrug
        If Col = col药品 And .EditText <> "" Then
            str性质 = Get部门性质(cbo药房.ItemData(cbo药房.ListIndex))
            str类别 = " And A.类别 IN('5','6','7')"
            If InStr(str性质, "西药房") = 0 Then str类别 = Replace(str类别, "'5',", "")
            If InStr(str性质, "成药房") = 0 Then str类别 = Replace(str类别, "'6',", "")
            If InStr(str性质, "中药房") = 0 Then str类别 = Replace(str类别, ",'7'", "")
            
            '不同的输入匹配方式
            strInput = UCase(.EditText)
            strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.名称 Like [2] And C.码类=[3] Or C.简码 Like [2] And C.码类 IN([3],3))"
            If IsNumeric(strInput) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.简码 Like [2] And C.码类=3)"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.输入全是字母时只匹配简码
                If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.简码 Like [2] And C.码类=[3]"
            ElseIf zlCommFun.IsCharChinese(strInput) Then
                strMatch = " And C.名称 Like [2] And C.码类=[3]"
            End If
            
            strSQL = _
                " Select Distinct 1 as 末级,A.ID,A.编码,C.名称," & _
                " B.住院单位 as 单位,A.规格,A.产地,A.类别 as 类别ID,B.住院包装 as 系数ID" & _
                " From 收费项目目录 A,药品规格 B,收费项目别名 C" & _
                " Where (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And A.服务对象 IN(2,3) And A.ID=B.药品ID And A.ID=C.收费细目ID" & str类别 & strMatch & _
                " Order by 编码"
            vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药品", False, "", "", False, False, True, vPoint.x, vPoint.Y, .CellHeight, blnCancel, False, True, _
                strInput & "%", mstrLike & strInput & "%", mint简码 + 1)
            If Not rsTmp Is Nothing Then
                If Not SetItemInput(Row, rsTmp) Then
                    Cancel = True
                Else
                    .EditText = .Text
                    If mblnReturn Then
                        Call EnterNextCell(Row, Col)
                    End If
                End If
            Else
                If Not blnCancel Then
                    MsgBox "输入""" & .EditText & """没有找到可用的药品。", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
            mblnReturn = False
        ElseIf Col = col留存数 Then
            If Not IsNumeric(.EditText) And .EditText <> "" Or Val(.EditText) < 0 Or Val(.EditText) > LONG_MAX Then
                MsgBox "输入的留存数量""" & .EditText & """错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                Cancel = True
            ElseIf Val(.EditText) > Val(.TextMatrix(Row, col应发数)) Then
                MsgBox "输入的留存数量""" & .EditText & """不应大于应发数量""" & .TextMatrix(Row, col应发数) & """！", vbInformation, gstrSysName
                Cancel = True
            Else
                .CellFontBold = False
                If Val(.EditText) = 0 Then .EditText = ""
                mblnChange = True
                If mblnReturn Then
                    Call EnterNextCell(Row, Col)
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Function LoadDrugAdvice() As Boolean
'功能：读取药品医嘱信息，用于计算
'说明：
'   没有包含非药品医嘱的药品计价
'   处理了长嘱不按规格下达的情况
    Dim strSQL As String
    Dim intMouse As Integer
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 11
        
    On Error GoTo errH
    
    strSQL = _
        "Select M.入院日期,A.开始执行时间,A.医嘱期效,D.药品ID,D.剂量系数,D.住院包装,Nvl(A.可否分零,D.住院可否分零) as 可否分零," & _
        " B.首次时间,B.末次时间,A.天数,A.执行时间方案,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.单次用量,B.发送数次" & _
        " From 病人医嘱记录 A,病人医嘱发送 B,住院费用记录 C,药品规格 D,病案主页 M" & _
        " Where A.诊疗类别 IN('5','6') And A.ID=B.医嘱ID And A.病人ID=M.病人ID And A.主页ID=M.主页ID" & _
        " And B.NO=C.NO And B.记录性质=C.记录性质 And B.医嘱ID=C.医嘱序号 And C.记录状态 IN(0,1,3)" & _
        " And C.收费细目ID=D.药品ID And B.发送时间 Between [1] And [2] And B.执行部门ID=[3] And C.病人病区ID=[4]"
    Set mrsAdvice = New ADODB.Recordset
    Set mrsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(dtpBegin.value), CDate(dtpEnd.value), cbo药房.ItemData(cbo药房.ListIndex), mlng病区ID)
    
    Screen.MousePointer = intMouse
    LoadDrugAdvice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call ReleaseRecord(mrsAdvice)
End Function

Private Function GetSurplus(ByVal lngRow As Long) As String
'功能：计算指定行药品的缺省留存数量
    Dim dbl比例 As Double, dbl总单量 As Double
    Dim lng次数 As Long, strTime As String
    
    If optRule(0).value Then
        GetSurplus = ""
    ElseIf optRule(1).value Then
        GetSurplus = Val(vsDrug.TextMatrix(lngRow, col应发数))
    ElseIf optRule(2).value Then
        dbl比例 = Val(cbo比例.Text) / 100
        If dbl比例 > 1 Then dbl比例 = 1
        If dbl比例 < 0 Then dbl比例 = 0
        GetSurplus = IntEx(Val(vsDrug.TextMatrix(lngRow, col应发数)) * dbl比例)
    ElseIf optRule(3).value Then
        '整体分零满足
        If vsDrug.TextMatrix(lngRow, col类别) = "7" Then
            GetSurplus = "" '中药应不存在留存
        Else
            If mrsAdvice Is Nothing Then Call LoadDrugAdvice
            mrsAdvice.Filter = "药品ID=" & vsDrug.RowData(lngRow)
            If Not mrsAdvice.EOF Then
                Do While Not mrsAdvice.EOF
                    '参见药品医嘱发送：
                    '相同药品，不同医嘱可能期效、频率等不同
                    '总单量=发送用药次数*第次单量
                    If Nvl(mrsAdvice!医嘱期效, 0) = 0 Then
                        If Not IsNull(mrsAdvice!首次时间) And Not IsNull(mrsAdvice!末次时间) And Not IsNull(mrsAdvice!执行时间方案) Then
                            strTime = Calc段内分解时间(mrsAdvice!首次时间, mrsAdvice!末次时间, "", _
                                mrsAdvice!执行时间方案, mrsAdvice!频率次数, mrsAdvice!频率间隔, mrsAdvice!间隔单位, mrsAdvice!开始执行时间)
                            lng次数 = UBound(Split(strTime, ",")) + 1
                            dbl总单量 = dbl总单量 + lng次数 * mrsAdvice!单次用量
                        Else
                            dbl总单量 = dbl总单量 + mrsAdvice!发送数次 '异常情况，直接取发送总单量
                        End If
                    ElseIf Not IsNull(mrsAdvice!单次用量) Then
                        If Nvl(mrsAdvice!频率次数, 0) = 0 Or Nvl(mrsAdvice!频率间隔, 0) = 0 Then
                            lng次数 = 1 '设置为一次性的临嘱药品
                        ElseIf Nvl(mrsAdvice!天数, 0) <> 0 And Not IsNull(mrsAdvice!执行频次) Then
                            '用药天数内按频率周期的次数
                            If mrsAdvice!间隔单位 = "周" Then
                                lng次数 = IntEx(mrsAdvice!天数 * (mrsAdvice!频率次数 / 7))
                            ElseIf mrsAdvice!间隔单位 = "天" Then
                                lng次数 = IntEx(mrsAdvice!天数 * (mrsAdvice!频率次数 / mrsAdvice!频率间隔))
                            ElseIf mrsAdvice!间隔单位 = "小时" Then
                                lng次数 = IntEx(mrsAdvice!天数 * (mrsAdvice!频率次数 / mrsAdvice!频率间隔) * 24)
                            ElseIf mrsAdvice!间隔单位 = "分钟" Then
                                lng次数 = IntEx(mrsAdvice!天数 * (mrsAdvice!频率次数 / mrsAdvice!频率间隔) * (24 * 60))
                            End If
                        Else
                            '可分零药品时,按总量对单量的倍数计算给药途径的次数,否则按一个频率周期的次数计算
                            If Nvl(mrsAdvice!可否分零, Nvl(mrsAdvice!可否分零, 0)) = 0 And Nvl(mrsAdvice!单次用量, 0) <> 0 Then
                                lng次数 = IntEx(mrsAdvice!总给予量 * mrsAdvice!剂量系数 / mrsAdvice!单次用量)
                            Else
                                lng次数 = Nvl(mrsAdvice!频率次数, 0)
                            End If
                        End If
                        
                        dbl总单量 = dbl总单量 + lng次数 * mrsAdvice!单次用量
                    End If
                    
                    mrsAdvice.MoveNext
                Loop
                
                mrsAdvice.MoveFirst '取一些药品信息,转为住院单位
                dbl总单量 = IntEx(dbl总单量 / Nvl(mrsAdvice!剂量系数, 1) / Nvl(mrsAdvice!住院包装, 1))
                If dbl总单量 > Val(vsDrug.TextMatrix(lngRow, col应发数)) Then
                    dbl总单量 = Val(vsDrug.TextMatrix(lngRow, col应发数))
                End If
                '留存数量=应发数量-实际需量
                GetSurplus = Val(vsDrug.TextMatrix(lngRow, col应发数)) - dbl总单量
            End If
        End If
    End If
    
    If Val(GetSurplus) = 0 Then GetSurplus = ""
End Function

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    
    '表头
    objOut.Title.Text = Sys.RowValue("部门表", mlng病区ID, "名称") & "药品留存清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set objRow = New zlTabAppRow
    objRow.Add "药房：" & zlCommFun.GetNeedName(cbo药房.Text)
    objRow.Add "时间：" & Format(dtpBegin.value, "yyyy-MM-dd HH:mm") & " 至 " & Format(dtpEnd.value, "yyyy-MM-dd HH:mm")
    objOut.UnderAppRows.Add objRow
    
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = vsDrug
    
    '输出
    vsDrug.Redraw = flexRDNone
    lngRow = vsDrug.Row: lngCol = vsDrug.Col
    vsDrug.Cell(flexcpForeColor, vsDrug.Row, 0, vsDrug.Row, vsDrug.Cols - 1) = vsDrug.ForeColor
    
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    vsDrug.Row = lngRow: vsDrug.Col = lngCol
    vsDrug.Redraw = flexRDDirect
    
    Call vsDrug_AfterRowColChange(-1, -1, vsDrug.Row, vsDrug.Col)
End Sub

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub
