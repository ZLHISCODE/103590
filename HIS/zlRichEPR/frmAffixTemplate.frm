VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmAffixTemplate 
   Caption         =   "附项模板设置"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12555
   Icon            =   "frmAffixTemplate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   12555
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vfgAffix 
      Height          =   6000
      Left            =   30
      TabIndex        =   0
      Top             =   840
      Width           =   3870
      _cx             =   6826
      _cy             =   10583
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12648447
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16769985
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6945
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAffixTemplate.frx":06EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17066
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vfgTemplate 
      Height          =   5970
      Left            =   3990
      TabIndex        =   2
      Top             =   855
      Width           =   8520
      _cx             =   15028
      _cy             =   10530
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorSel    =   16769985
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   1
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
      BackColorFrozen =   -2147483630
      ForeColorFrozen =   -2147483630
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   1110
      Top             =   225
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAffixTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'菜单类型枚举定义
Private Enum TMenuType
    mtFile = 1
    mtSave
    mtCancel
    mtImport
    mtExport
    mtQuit
    
    mtEdit
    mtNew
    mtDel
    mtClearCount
End Enum


'行状态
Private Enum TRowState
    rsNormal = 0
    rsNew
    rsDel
    rsModify
End Enum


'附项模板列表的列定义
Private Enum TAffixTemplateCol
    atcState = 0
    atcTitle = 1
    atcCount = 2
    atcContext = 3
End Enum

Private mrsAffixTemplate As ADODB.Recordset

Private mblnModifyState As Boolean      '修改状态
Private mlngRequestPageId As Long       '病历单据Id

Private mstrStartEditText As String

Public Sub ShowAffixConfig(ByVal lngRequestPageId As Long, objOwner As Object)
    mlngRequestPageId = lngRequestPageId
    Me.Show 1, objOwner
End Sub


Private Sub Menu_Help_Web_Mail_click()
On Error GoTo ErrHandle
    zlMailTo hWnd
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_About_click()
On Error GoTo ErrHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'功能：调用帮助主题
On Error GoTo ErrHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo ErrHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo ErrHandle
    zlHomePage hWnd
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_StatusBar_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Button_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Size_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).STYLE
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.STYLE = intStyle
        Next
    Next
    
    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub CancelModify()
'撤销修改
    vfgTemplate.Row = vfgTemplate.Row
    
    If MsgBox("是否撤销所做的修改？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
     
    Call LoadAffixTemplate(mlngRequestPageId, vfgAffix.Text)
     
    mblnModifyState = False
End Sub


Private Sub NewTemplate()
'新增模板
    vfgTemplate.Rows = vfgTemplate.Rows + 1
    vfgTemplate.TextMatrix(vfgTemplate.Rows - 1, atcCount) = 0
    
    vfgTemplate.Col = 1
    vfgTemplate.Row = vfgTemplate.Rows - 1
    
    '更新行数据状态
    Call SetRowDataState(vfgTemplate.Row, rsNew)
    
    vfgTemplate.EditCell
    
    mblnModifyState = True
End Sub

Private Sub DelTemplate()
'删除模板
    Dim lngNextRow As Long
    Dim i As Long
    
    If vfgTemplate.Rows <= 1 Then Exit Sub
    
    lngNextRow = -1
    
    If vfgTemplate.Row < vfgTemplate.Rows - 1 Then
        For i = vfgTemplate.Row + 1 To vfgTemplate.Rows - 1
            If Not vfgTemplate.RowHidden(i) Then
                lngNextRow = i
                Exit For
            End If
        Next i
    End If
    
    If lngNextRow = -1 Or vfgTemplate.Row = vfgTemplate.Rows - 1 Then
        For i = vfgTemplate.Rows - 1 To 1 Step -1
            If Not vfgTemplate.RowHidden(i) And i <> vfgTemplate.Row Then
                lngNextRow = i
                Exit For
            End If
        Next i
    End If
    


     vfgTemplate.RowHidden(vfgTemplate.Row) = True
     
     '更新行状态
     If vfgTemplate.Cell(flexcpData, vfgTemplate.Row, 1) <> "" Then
        Call SetRowDataState(vfgTemplate.Row, rsDel)
     Else
        Call SetRowDataState(vfgTemplate.Row, rsNormal)
     End If
     
     If lngNextRow > -1 Then vfgTemplate.Row = lngNextRow

    mblnModifyState = True
End Sub

Private Sub SetRowDataState(ByVal lngRow As Long, ByVal rsState As TRowState)
'设置数据行状态
    vfgTemplate.Cell(flexcpData, lngRow, atcState) = rsState
End Sub

Private Function GetRowDataState(ByVal lngRow As Long) As TRowState
'获取数据行状态
    GetRowDataState = Val(vfgTemplate.Cell(flexcpData, lngRow, atcState))
End Function


Private Function VerifyDataInputIsOk() As Boolean
'验证输入数据是否正确
    Dim i As Long
    
    VerifyDataInputIsOk = False
    
    For i = 1 To vfgTemplate.Rows - 1
        If Not vfgTemplate.RowHidden(i) Then
            If vfgTemplate.TextMatrix(i, atcTitle) = "" Then
                MsgBox "标题不能为空。", vbOKOnly, Me.Caption
                
                Call vfgTemplate.ShowCell(i, atcTitle)
                Call vfgTemplate.Select(i, atcTitle)
                Call vfgTemplate.EditCell
                
                Exit Function
            End If
            
            If Len(vfgTemplate.TextMatrix(i, atcCount)) > 8 Then
                MsgBox "数值位数不能超出8位。", vbOKOnly, Me.Caption
                
                Call vfgTemplate.ShowCell(i, atcCount)
                Call vfgTemplate.Select(i, atcCount)
                Call vfgTemplate.EditCell
                
                Exit Function
            End If
        End If
    Next i
    
    VerifyDataInputIsOk = True
End Function

Private Function SaveTemplate() As Boolean
'保存模板
    Dim i As Long
    Dim arySql() As String
    Dim rsRowState As TRowState
    
    vfgTemplate.Row = vfgTemplate.Row
    
    SaveTemplate = False
    
    If Not VerifyDataInputIsOk Then Exit Function
    
    ReDim Preserve arySql(1)
    
    arySql(0) = ""
    
    For i = 1 To vfgTemplate.Rows - 1
        rsRowState = GetRowDataState(i)
        
        Select Case rsRowState
            Case TRowState.rsNew
                If vfgTemplate.TextMatrix(i, atcTitle) <> "" Then
                    ReDim Preserve arySql(UBound(arySql) + 1)
                    arySql(UBound(arySql)) = "zl_病历附项模板_Insert(" & mlngRequestPageId & ",'" & _
                                                                        vfgAffix.Text & "','" & _
                                                                        vfgTemplate.TextMatrix(i, atcTitle) & "','" & _
                                                                        vfgTemplate.TextMatrix(i, atcContext) & "'," & _
                                                                        Val(vfgTemplate.TextMatrix(i, atcCount)) & ")"
                End If

            Case TRowState.rsDel
                ReDim Preserve arySql(UBound(arySql) + 1)
                arySql(UBound(arySql)) = "zl_病历附项模板_Del(" & vfgTemplate.Cell(flexcpData, i, atcTitle) & ")"
                
            Case TRowState.rsModify
                If vfgTemplate.TextMatrix(i, atcTitle) <> "" Then
                    ReDim Preserve arySql(UBound(arySql) + 1)
                    arySql(UBound(arySql)) = "zl_病历附项模板_Update(" & vfgTemplate.Cell(flexcpData, i, atcTitle) & ",'" & _
                                                                        vfgTemplate.TextMatrix(i, atcTitle) & "','" & _
                                                                        vfgTemplate.TextMatrix(i, atcContext) & "'," & _
                                                                        Val(vfgTemplate.TextMatrix(i, atcCount)) & ")"
                End If
                
        End Select
    Next i
    
    
On Error GoTo ErrHandle
    gcnOracle.BeginTrans
    
    '处理数据保存
    For i = LBound(arySql) To UBound(arySql)
        If arySql(i) <> "" Then
            Call zlDatabase.ExecuteProcedure(arySql(i), "保存附项模板")
        End If
    Next i
    
    gcnOracle.CommitTrans
    
    '重新载入附项模板数据
    Call LoadTemplateToDataSet(mlngRequestPageId)
    
    mblnModifyState = False
    SaveTemplate = True
Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function


Private Sub ClearUseCount()
'清除使用次数
    Dim i As Long
    
    For i = 1 To vfgTemplate.Rows - 1
        If Val(vfgTemplate.TextMatrix(i, atcCount)) <> 0 Then
            vfgTemplate.TextMatrix(i, atcCount) = 0
            Call SetRowDataState(i, rsModify)
            mblnModifyState = True
        End If
    Next i
End Sub


Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Select Case Control.ID
    
        Case TMenuType.mtCancel
            Call CancelModify       '撤销修改
            
        Case TMenuType.mtNew
            Call NewTemplate        '新增模板
            
        Case TMenuType.mtDel
            Call DelTemplate        '删除模板
                        
        Case TMenuType.mtSave
            Call SaveTemplate       '保存模板
            
        Case TMenuType.mtClearCount
            Call ClearUseCount      '清除使用次数
                            
        Case TMenuType.mtQuit
            Call Unload(Me)
            
'        Case TMenuType.mtImport
'            '导入方案......
'
'        Case TMenuType.mtExport
'            '导出模板......
            
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(Control)
        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(Control)
        Case conMenu_View_ToolBar_Size '大图标
            Call Menu_View_ToolBar_Size_click(Control)
        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(Control)
            
'--------------------------帮助-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
    End Select
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
'调整模板界面
    vfgAffix.Left = Left
    vfgAffix.Top = Top
    vfgAffix.Height = Bottom - IIf(stbThis.Visible, stbThis.Height, 0) - Top
    
    
    vfgTemplate.Left = vfgAffix.Width + 80
    vfgTemplate.Top = Top
    vfgTemplate.Width = ScaleWidth - vfgAffix.Width - 80
    vfgTemplate.Height = vfgAffix.Height
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case TMenuType.mtCancel, TMenuType.mtSave
            Control.Enabled = mblnModifyState
        Case TMenuType.mtNew, TMenuType.mtDel, TMenuType.mtClearCount
            Control.Enabled = vfgAffix.Rows > 1
    End Select
End Sub

Private Sub LoadRequestAffix(ByVal lngRequestPageId As Long)
'载入申请附项
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select 项目 from 病历单据附项 where 只读=0 and 文件Id=[1] order by 排列"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRequestPageId)
    
    vfgAffix.Rows = 1
    If rsData.RecordCount <= 0 Then Exit Sub
    
    
    While Not rsData.EOF
        vfgAffix.Rows = vfgAffix.Rows + 1
        vfgAffix.TextMatrix(vfgAffix.Rows - 1, 0) = NVL(rsData!项目)
        vfgAffix.Cell(flexcpAlignment, vfgAffix.Rows - 1) = flexAlignLeftCenter
    
        Call rsData.MoveNext
    Wend
    
    vfgAffix.Row = 1
End Sub

Private Sub LoadAffixTemplate(ByVal lngRequestPageId As Long, ByVal strProjectName As String)
'载入附项模板

    vfgTemplate.Rows = 1
    
    If mrsAffixTemplate Is Nothing Then Exit Sub
    
    mrsAffixTemplate.Filter = "病历文件Id=" & lngRequestPageId & " and 单据附项='" & strProjectName & "'"
    
    If mrsAffixTemplate.RecordCount <= 0 Then Exit Sub
    
    While Not mrsAffixTemplate.EOF
        vfgTemplate.Rows = vfgTemplate.Rows + 1
        
        vfgTemplate.Cell(flexcpText, vfgTemplate.Rows - 1, atcTitle) = NVL(mrsAffixTemplate!模板标题)
        vfgTemplate.Cell(flexcpData, vfgTemplate.Rows - 1, atcTitle) = NVL(mrsAffixTemplate!ID)
        
        vfgTemplate.Cell(flexcpText, vfgTemplate.Rows - 1, atcCount) = NVL(mrsAffixTemplate!使用次数)
        
        vfgTemplate.Cell(flexcpText, vfgTemplate.Rows - 1, atcContext) = NVL(mrsAffixTemplate!模板内容)
        
        Call SetRowDataState(vfgTemplate.Rows - 1, TRowState.rsNormal)
        
        Call mrsAffixTemplate.MoveNext
    Wend
    
    vfgTemplate.Row = 1
End Sub


Private Sub LoadTemplateToDataSet(ByVal lngRequestPageId As Long)
'载入附项模板到数据集
    Dim strSQL As String
    
    strSQL = "select Id,病历文件Id,单据附项,模板标题,模板内容,使用次数 from 病历附项模板 where 病历文件Id=[1] order by 单据附项,模板标题"
    
    Set mrsAffixTemplate = zlDatabase.OpenSQLRecord(strSQL, "查询附项模板", lngRequestPageId)
End Sub

Private Sub Form_Load()
'###############配置调试数据################
'    InitDebugObject 1290, Me, "zlhis", "HIS"
'    mlngRequestPageId = 118
'###########################################

    Call RestoreWinState(Me, App.ProductName)

    mblnModifyState = False

    
    Call InitFaceList
    Call InitCommandBars
    
    Call LoadTemplateToDataSet(mlngRequestPageId)
    Call LoadRequestAffix(mlngRequestPageId)
End Sub


Public Sub InitDebugObject(ByVal lngModuleNum As Long, ByVal frmMain As Object, ByVal strUser As String, ByVal strPwd As String)
'初始化调试状态下的所需对象
    Set gcnOracle = New ADODB.Connection
    
    Call OraDataOpen("", strUser, strPwd)
    
    glngSys = 100
    gstrPrivs = ";PACS报告打印;PACS报告删除;PACS报告书写;PACS报告他科报告;PACS报告修订;PACS他人报告;采集参数设置;参数设置;存储管理;关联病人;基本;检查报到;检查登记;检查完成;绿色通道;排队叫号;清除图像;取消报到;取消检查完成;删除临时影像;视频采集;随访;所有科室;图像关联;未缴费报到;文件发送;无报告完成;影像质控;档案分类设置;Excel输出;"
    glngModul = lngModuleNum
    
    
    
    Call InitCommon(gcnOracle)
    
    Call RegCheck
End Sub

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    gstrDBUser = UCase(strUserName)
    SetDbUser gstrDBUser
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function


Private Sub InitFaceList()
'初始化界面列表配置
    vfgAffix.TextMatrix(0, 0) = "申请附项"
    
    
    vfgTemplate.ColWidth(atcState) = 120
    
    vfgTemplate.TextMatrix(0, atcTitle) = "标题"
    vfgTemplate.ColWidth(atcTitle) = 1600
    vfgTemplate.ColAlignment(atcTitle) = flexAlignLeftCenter
    
    vfgTemplate.TextMatrix(0, atcCount) = "次数"
    vfgTemplate.ColWidth(2) = 520
    vfgTemplate.ColAlignment(atcCount) = flexAlignLeftCenter
    
    vfgTemplate.TextMatrix(0, atcContext) = "内容"
    vfgTemplate.ColAlignment(atcContext) = flexAlignLeftCenter
End Sub




Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '设置菜单栏和工具栏风格
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True                                '显示按钮提示
        .AlwaysShowFullMenus = False                            '不常用的菜单项先隐藏
        .UseFadedIcons = False                                  '图标显示为褪色效果
        .IconsWithShadow = True                                 '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True                                '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True                                      '工具栏显示为大图标
        .SetIconSize True, 24, 24                               '设置大图标的尺寸
        .SetIconSize False, 16, 16                              '设置小图标的尺寸
    End With
    
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '设置控件显示风格
        .EnableCustomization False                             '是否允许自定义设置
        Set .Icons = zlCommFun.GetPubIcons                     '设置关联的图标控件
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
'Begin------------------------编辑菜单--------------------------------------默认可见
    cbrMain.ActiveMenuBar.Title = "菜单"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "文件(&F)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "保存(&S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "取消(&C)"): cbrControl.IconId = 3565
'    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtImport, "导入(&I)"): cbrControl.IconId = 0: cbrControl.BeginGroup = True
'    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtExport, "导出(&E)"): cbrControl.IconId = 0
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "退出(&Q)"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "编辑(&E)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtNew, "新增(&N)"): cbrControl.IconId = 4010
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtDel, "删除(&D)"): cbrControl.IconId = 4008
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtClearCount, "清除使用次数(&F)"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True
        
    
    'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)
    
    'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(H)")
    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    
    '---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "保存", "保存方案"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "取消", "取消修改"): cbrControl.IconId = 3565
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtNew, "新增", "新增方案"): cbrControl.IconId = 4010: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtDel, "删除", "删除方案"): cbrControl.IconId = 4008
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "退出", "退出"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
End Sub




Public Sub CreateViewAndHelpMenu(objViewMenu As Object, objHelpMenu As Object, _
    Optional ByVal strMenuTag As String = "")
    
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    
    
    'Begin----------------------查看菜单--------------------------------------
    If Not (objViewMenu Is Nothing) Then
        Set cbrMenuBar = objViewMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(T)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
            
                With cbrControl.CommandBar '二级菜单
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(0)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(1)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                End With
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(S)")
                cbrControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
        End With
    End If

    'Begin----------------------帮助菜单--------------------------------------默认可见
    If Not (objHelpMenu Is Nothing) Then
        Set cbrMenuBar = objHelpMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "帮助主题(M)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 901
                
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB上的中联(W)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
                
                With cbrControl.CommandBar
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(0)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "中联主页(1)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(2)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 9022
                End With
                
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "关于…(A)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
        End With
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrHandle
    Dim lngResult As Long
    
    vfgTemplate.Row = vfgTemplate.Row
    
    If Not mblnModifyState Then Exit Sub
    
    lngResult = MsgBox("当前申请附项【" & vfgAffix.TextMatrix(vfgAffix.Row, vfgAffix.Col) & "】的模板数据已被修改，是否保存？", vbYesNoCancel, Me.Caption)
    
    Select Case lngResult
        Case vbNo
            mblnModifyState = False
            Exit Sub
        Case vbCancel
            Cancel = True
        Case vbYes
            If Not SaveTemplate Then Cancel = True
    End Select
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vfgAffix_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
On Error GoTo ErrHandle
    Dim lngResult As Long
    
    If Not mblnModifyState Then Exit Sub
    
    lngResult = MsgBox("当前申请附项【" & vfgAffix.TextMatrix(OldRowSel, OldColSel) & "】的模板数据已被修改，是否保存？", vbYesNoCancel, Me.Caption)
    
    Select Case lngResult
        Case vbNo
            mblnModifyState = False
            Exit Sub
        Case vbCancel
            Cancel = True
        Case vbYes
            If Not SaveTemplate Then Cancel = True
    End Select
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgAffix_SelChange()
On Error GoTo ErrHandle
    
    Call LoadAffixTemplate(mlngRequestPageId, vfgAffix.Text)
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub vfgTemplate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Dim rsRowState As TRowState
    
    '如果没有做任何改变，则不进入编辑状态
    If vfgTemplate.TextMatrix(Row, Col) = mstrStartEditText Then Exit Sub
    
    rsRowState = GetRowDataState(Row)
    If rsRowState <> rsNew And rsRowState <> rsDel Then
        Call SetRowDataState(Row, rsModify)
    End If

    mblnModifyState = True
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub vfgTemplate_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = atcCount Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
        If KeyAscii = vbKeyReturn Then Exit Sub
        If KeyAscii = vbKeyEscape Then Exit Sub
        If KeyAscii = vbKeyDelete Then Exit Sub
        If KeyAscii = vbKeyBack Then Exit Sub
            
        KeyAscii = 0
        
    End If
End Sub

Private Sub vfgTemplate_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        Call EditNextCell(vfgTemplate.Row)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub EditNextCell(ByVal lngRow As Long, Optional ByVal blnAutoNextRow As Boolean = True)
'编辑下一列
    Dim iRow As Long
    Dim iCol As Long
            
    If vfgTemplate.Editable = flexEDNone Then Exit Sub
    
    Do While vfgTemplate.Col + 1 < vfgTemplate.Cols
        If Not vfgTemplate.ColHidden(vfgTemplate.Col + 1) Then
            Exit Do
        Else
            Call vfgTemplate.Select(lngRow, vfgTemplate.Col + 1)
        End If
    Loop
    
nextCell:
    
    If vfgTemplate.Col + 1 >= vfgTemplate.Cols Then
 
        iRow = GetNextRowIndex(lngRow)
        
        If iRow > 0 Then
            For iCol = 1 To vfgTemplate.Cols - 1
                If Not vfgTemplate.ColHidden(iCol) Then Exit For
            Next iCol
            
            If iRow < vfgTemplate.Rows Then
                If iCol = vfgTemplate.Cols Then iCol = vfgTemplate.Cols - 1
                
                Call vfgTemplate.Select(iRow, iCol)
                Call vfgTemplate.ShowCell(iRow, iCol)
            End If
        End If
        
        Call vfgTemplate.EditCell
 
        Exit Sub
    End If
    
    
    Call vfgTemplate.Select(lngRow, vfgTemplate.Col + 1)
        
    Call vfgTemplate.EditCell
End Sub

Public Function GetNextRowIndex(ByVal lngRow As Long) As Long
'取得下一行的索引
    Dim i As Long
    
    GetNextRowIndex = -1
    
    For i = lngRow + 1 To vfgTemplate.Rows - 1
        If Not vfgTemplate.RowHidden(i) Then
            GetNextRowIndex = i
            Exit Function
        End If
    Next i
    
    If GetNextRowIndex = -1 Then
        i = lngRow - 1
        Do While i > 0
            If Not vfgTemplate.RowHidden(i) Then
                GetNextRowIndex = i
                Exit Function
            End If
            
            i = i - 1
        Loop
    End If
End Function

Private Sub vfgTemplate_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    mstrStartEditText = vfgTemplate.TextMatrix(Row, Col)
End Sub

