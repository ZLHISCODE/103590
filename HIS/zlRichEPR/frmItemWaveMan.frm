VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmItemWaveMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "波动项目设置"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5895
   Icon            =   "frmItemWaveMan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid VsfData 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5115
      _cx             =   9022
      _cy             =   6165
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   5000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmItemWaveMan.frx":020A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      AutoSizeMouse   =   0   'False
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
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmItemWaveMan.frx":026C
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmItemWaveMan.frx":0280
   End
End
Attribute VB_Name = "frmItemWaveMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnEdit As Boolean

Private Const conMenu_保存 = 2
Private Const conMenu_恢复 = 3
Private Const conMenu_帮助 = 4
Private Const conMenu_退出 = 5

Private Enum EnumCOl
    项目名称 = 0
    是否波动 = 1
End Enum


Private Function SaveData() As Boolean
'保存数据信息
    Dim rs As New ADODB.Recordset
    Dim strData As String
    Dim intRow As Integer
    Dim lngOrder As Long
    Dim lngID As String
    Dim arrSQL() As String
    Dim i As Integer
    Dim blnTrans As Boolean
    
    On Error GoTo errHand
    
    For intRow = 1 To VsfData.Rows - 1
        lngOrder = VsfData.RowData(intRow)
        If lngOrder <> 0 And VsfData.TextMatrix(intRow, 是否波动) = "√" Then
            lngID = lngID & "," & lngOrder
            strData = strData & "|" & lngOrder & ";" & VsfData.TextMatrix(intRow, 项目名称)
        End If
    Next intRow
    
    If Left(strData, 1) = "|" Then strData = Mid(strData, 2)
    If Left(lngID, 1) = "," Then lngID = Mid(lngID, 2)
    
    ReDim Preserve arrSQL(1 To 1)
    '修改波动项目记录频次(>2的情况)
    If Val(lngID) <> 0 Then
        gstrSQL = "Select  /*+ RULE*/ 项目序号,排列序号,记录名,记录法,记录符,记录色,最大值,最小值,单位值,单位,最高行,记录频次,刻度间隔,警示线 " & _
            "   From 体温记录项目 A,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) B" & _
            "  Where A.项目序号=B.Column_Value And nvl(A.记录频次,0)>2"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "护理记录项目", CStr(lngID))

        With rs
            Do While Not .EOF
                gstrSQL = "ZL_体温记录项目_INSERT(" & NVL(!项目序号, 0) & "," & _
                                                                Val(NVL(!排列序号)) & ",'" & _
                                                                Trim(NVL(!记录名)) & "'," & _
                                                                Val(NVL(!记录法)) & ",'" & _
                                                                NVL(!记录符) & "'," & _
                                                                Val(NVL(!记录色)) & "," & _
                                                                IIf(Trim(NVL(!最小值)) <> "", Val(NVL(!最小值)), "NULL") & "," & _
                                                                IIf(Trim(NVL(!最大值)) <> "", Val(NVL(!最大值)), "NULL") & "," & _
                                                                IIf(Trim(NVL(!单位值)) <> "", Val(NVL(!单位值)), "NULL") & ",'" & _
                                                                Trim(NVL(!单位)) & "'," & _
                                                                "NULL" & "," & _
                                                                2 & "," & IIf(Trim(NVL(!刻度间隔)) = "", "NULL", Val(NVL(!刻度间隔))) & "," & _
                                                                IIf(Trim(NVL(!警示线)) = "", "NULL", Val(NVL(!警示线))) & ")"
                arrSQL(ReDimArray(arrSQL)) = gstrSQL
            .MoveNext
            Loop
        End With
        
        '保存数据
        blnTrans = (rs.RecordCount > 0)
    End If
    gstrSQL = "zl_护理波动项目_Upate('" & strData & "')"
    arrSQL(ReDimArray(arrSQL)) = gstrSQL
    
    If blnTrans Then gcnOracle.BeginTrans
    For i = 1 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "护理波动项目")
    Next i
    
    If blnTrans Then gcnOracle.CommitTrans
    
    SaveData = True
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnEdit Then
        If MsgBox("修改的数据还未保存，你确定要退出吗？" & vbCrLf & "点“是”则放弃修改并退出，点“否”继续修改！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long, lngTop As Long, lngHeight As Long, lngWidth As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngWidth, lngHeight)
    With VsfData
        .Left = lngLeft
        .Top = lngTop
        .Height = lngHeight - lngTop
        .Width = lngWidth
    End With
End Sub

Private Sub VsfData_DblClick()
    Call VsfData_KeyDown(vbKeySpace, 0)
End Sub

Private Sub VsfData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngOrder As Long
    Dim intRow As Integer, intCount As Integer
    If VsfData.Col = 是否波动 And KeyCode = vbKeySpace Then
        
        '设置波动项目
        lngOrder = VsfData.RowData(VsfData.Row)
        If lngOrder = 0 Then Exit Sub
        VsfData.TextMatrix(VsfData.Row, 是否波动) = IIf(VsfData.TextMatrix(VsfData.Row, 是否波动) = "√", "", "√")
        intCount = VsfData.Rows - 1
        If lngOrder = 4 Or lngOrder = 5 Then
            For intRow = 1 To intCount
                If VsfData.RowData(intRow) = IIf(lngOrder = 4, 5, 4) Then
                    VsfData.TextMatrix(intRow, 是否波动) = VsfData.TextMatrix(VsfData.Row, 是否波动)
                    Exit For
                End If
            Next intRow
        End If
        
        mblnEdit = True
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_保存
        If Not SaveData Then Exit Sub
        mblnEdit = False
    Case conMenu_恢复
        Call LoadData
    Case conMenu_帮助
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_退出
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With VsfData
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight
        .Height = lngBottom
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_保存
        Control.Enabled = mblnEdit
    Case conMenu_恢复
        Control.Enabled = mblnEdit
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call MainDefCommandBar
    Call LoadData
End Sub

Private Sub LoadData()
'--初始化表格数据
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    
    mblnEdit = False
    With VsfData
        .Clear
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        .Cols = 2
        .TextMatrix(0, 项目名称) = "项目名称"
        .TextMatrix(0, 是否波动) = "是否波动"
        .ColWidth(项目名称) = 1400
        .ColWidth(是否波动) = 900
        
        .ColAlignment(项目名称) = flexAlignLeftCenter
        .ColAlignment(是否波动) = flexAlignCenterCenter
    End With
    
    '添加数据
    gstrSQL = " SELECT A.项目序号,A.项目名称,DECODE(NVL(C.项目序号,0),0,0,1) 波动项目" & vbNewLine & _
            "   FROM 护理记录项目 A,体温记录项目 B,护理波动项目 C" & vbNewLine & _
            "   WHERE A.项目序号=B.项目序号 AND A.项目序号=C.项目序号(+) AND A.项目类型=0" & vbNewLine & _
            "   AND A.项目表示=0 AND B.记录法=2 AND B.项目序号<>3" & vbNewLine & _
            "   ORDER BY A.项目序号"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取体温表格项目")
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition >= VsfData.Rows Then VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(.AbsolutePosition, 项目名称) = CStr(!项目名称)
            VsfData.TextMatrix(.AbsolutePosition, 是否波动) = IIf(NVL(!波动项目, 0) = 1, "√", "")
            VsfData.RowData(.AbsolutePosition) = CLng(!项目序号)
            .MoveNext
        Loop
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup, objFile As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim lngHandel As Long

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    'cbsMain
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
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
    cbsMain.Icons = imgPublic.Icons
    
    '工具栏定义
    '-----------------------------------------------------
    cbsMain.DeleteAll
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)      '固有
    objBar.EnableDocking xtpFlagStretched
    objBar.Closeable = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_保存, "保存"): objControl.STYLE = xtpButtonIconAndCaption: objControl.ToolTipText = "保存数据": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_恢复, "恢复"): objControl.STYLE = xtpButtonIconAndCaption: objControl.ToolTipText = "取消保存"
        Set objControl = .Add(xtpControlButton, conMenu_帮助, "帮助"): objControl.STYLE = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_退出, "退出"): objControl.STYLE = xtpButtonIconAndCaption
    End With
    
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_保存             '保存
    End With
End Sub

