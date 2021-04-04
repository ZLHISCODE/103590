VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmKillProcessManage 
   Caption         =   "进程清单管理"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   Icon            =   "frmKillProcessManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11625
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraSplit 
      Height          =   30
      Left            =   -15
      TabIndex        =   10
      Top             =   930
      Width           =   11970
   End
   Begin VB.PictureBox picEdit 
      Height          =   5670
      Left            =   7860
      ScaleHeight     =   5610
      ScaleWidth      =   3600
      TabIndex        =   2
      Top             =   1005
      Width           =   3660
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1766
         TabIndex        =   17
         Top             =   5025
         Width           =   800
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "修改(&U)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   958
         TabIndex        =   16
         Top             =   5025
         Width           =   800
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增(&A)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   150
         TabIndex        =   15
         Top             =   5025
         Width           =   800
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   345
         Left            =   2475
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出(&D)"
         Height          =   345
         Left            =   2575
         TabIndex        =   13
         Top             =   5025
         Width           =   800
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   345
         Left            =   1575
         TabIndex        =   8
         Top             =   3240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   645
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   255
         Width           =   2745
      End
      Begin VB.TextBox txtDescription 
         Height          =   1600
         Left            =   645
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "该描述最多可输入100个汉字或200个字符"
         Top             =   1410
         Width           =   2745
      End
      Begin VB.ComboBox cboType 
         Enabled         =   0   'False
         Height          =   300
         Left            =   645
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   2745
      End
      Begin VB.Label lblAdd 
         Height          =   180
         Left            =   660
         TabIndex        =   12
         Top             =   3315
         Width           =   720
      End
      Begin VB.Label lblShowCheckName 
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   645
         TabIndex        =   11
         Top             =   570
         Width           =   2760
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "名称"
         Height          =   180
         Left            =   150
         TabIndex        =   9
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "类型"
         Height          =   180
         Left            =   150
         TabIndex        =   7
         Top             =   870
         Width           =   360
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "描述"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   1410
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfProcessList 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   1005
      Width           =   7740
      _cx             =   13652
      _cy             =   9975
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
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   260
      RowHeightMax    =   260
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmKillProcessManage.frx":6852
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
      ExplorerBar     =   1
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
   Begin VB.Image imgMain 
      Height          =   720
      Left            =   210
      Picture         =   "frmKillProcessManage.frx":68EB
      Top             =   105
      Width           =   720
   End
   Begin VB.Label lblCaption 
      Height          =   540
      Left            =   1170
      TabIndex        =   1
      Top             =   210
      Width           =   10395
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuAdd 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu mnuPopuModify 
         Caption         =   "修改(&U)"
      End
      Begin VB.Menu mnuPopuDelete 
         Caption         =   "删除(&D)"
      End
   End
End
Attribute VB_Name = "frmKillProcessManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnRightClick As Boolean   '标记是否为鼠标右键点击，主要用于屏蔽右键点击列表导致vsfProcessList_Click的隐式调用
Private mblnOk As Boolean  '标记数据是否保存成功
Private Enum ProcessList
    PL_序号 = 0
    PL_分类 = 1
    PL_名称 = 2
    PL_类型 = 3
    PL_描述 = 4
End Enum

Public Sub ShowMe(ByVal strModule As String)
'strModule：调用该窗体的模块号
    Select Case strModule
        Case "0102"   '系统升迁管理
            Me.Caption = "中断客户端连接的进程定义"
            lblCaption = "在系统升迁过程中，若存在进程对数据库进行操作或使用，会影响升级效率。" & vbNewLine & _
                        "另一方面，对临时表结构调整，必须杀掉所有杀掉所有使用临时表的会话（必须标识会话的进程，防止误杀）才能进行调整。" & vbNewLine & _
                        "基于产品数据结构的进程都应该加入到该清单中。使用的数据结构和产品数据结构存在外键关系的进程，也应该加入到该清单。"
        Case "0307"   '客户端升级管理
            Me.Caption = "客户端进程管理"
            lblCaption = "客户端自动升级时，如果正在运行某些应用程序，待升级的文件如果被占用就无法被替换。" & vbNewLine & _
                        "为了保障正常升级，以下列表中的应用程序将会被升级程序自动终止。"
    End Select
    Me.Show vbModal, frmMDIMain
End Sub

Private Sub cmdCancel_Click()
    Call FillSelectData(vsfProcessList.Row)
    Call SetEnable(False)
    mblnRightClick = False
    Call vsfProcessList_Click
    lblAdd.Caption = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    mblnOk = False
    '检查输入数据是否符合要求
    If CheckData = False Then Exit Sub
    
    txtName.Tag = txtName.Text
    cboType.Tag = cboType.Text
    txtDescription.Tag = txtDescription.Text
    If Val(mnuPopuAdd.Tag) = 1 Then
        '新增数据
        Call ExecuteProcedure("Zl_Zlkillprocess_Edit(1, Null, '" & UCase(Trim(txtName.Text)) & "', " & cboType.ListIndex & ", '" & txtDescription.Text & "')", Me.Caption)
        strSQL = "Select 序号 From Zlkillprocess Where 名称 = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, UCase(Trim(txtName.Text)))
        If rsTmp.RecordCount > 0 Then
            lblAdd.Caption = "添加成功"
            vsfProcessList.Rows = vsfProcessList.Rows + 1
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_序号) = rsTmp!序号
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_分类) = "自定义"
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_名称) = UCase(Trim(txtName.Text))
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_类型) = cboType.Text
            vsfProcessList.TextMatrix(vsfProcessList.Rows - 1, PL_描述) = txtDescription.Text
            vsfProcessList.Row = vsfProcessList.Rows - 1
            Call vsfProcessList.ShowCell(vsfProcessList.Row, PL_序号)
        End If
        mblnOk = True
        '添加成功后，自动进入新增状态，提高新增效率
        Call mnuPopuAdd_Click
    Else
        '修改数据
        Call ExecuteProcedure("Zl_Zlkillprocess_Edit(2, " & vsfProcessList.TextMatrix(vsfProcessList.Row, PL_序号) & ", '" & _
                                UCase(Trim(txtName.Text)) & "', " & cboType.ListIndex & ", '" & txtDescription.Text & "')", Me.Caption)
        lblAdd.Caption = "修改成功"
        vsfProcessList.TextMatrix(vsfProcessList.Row, PL_名称) = UCase(Trim(txtName.Text))
        vsfProcessList.TextMatrix(vsfProcessList.Row, PL_类型) = cboType.Text
        vsfProcessList.TextMatrix(vsfProcessList.Row, PL_描述) = txtDescription.Text
        Call SetEnable(False)
        mblnOk = True
        mblnRightClick = False
        Call vsfProcessList_Click
    End If
    Exit Sub
errH:
    If Val(mnuPopuAdd.Tag) = 1 Then
        MsgBox "添加失败！" & vbNewLine & err.Description, vbInformation, gstrSysName
    Else
        MsgBox "修改失败！" & vbNewLine & err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '禁止输入单引号
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    cboType.addItem "进程"
    cboType.addItem "服务"
    '填充进程数据
    Call FillProcessData
End Sub

Private Sub FillProcessData()
    Dim strSQL  As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select 序号, 名称, 类型, 描述, 是否固定 From Zltools.Zlkillprocess Order By 序号"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    With rsTmp
        vsfProcessList.Rows = .RecordCount + 1
        For i = 1 To .RecordCount
            vsfProcessList.TextMatrix(i, PL_序号) = !序号
            vsfProcessList.TextMatrix(i, PL_分类) = IIf(!是否固定 = 1, "固定", "自定义")
            vsfProcessList.TextMatrix(i, PL_名称) = !名称
            vsfProcessList.TextMatrix(i, PL_类型) = IIf(!类型 = 1, "服务", "进程")
            vsfProcessList.TextMatrix(i, PL_描述) = !描述 & ""
            .MoveNext
        Next
        vsfProcessList.ScrollTrack = True
        If .RecordCount > 0 Then
            vsfProcessList.Row = 1
            Call FillSelectData(1)
            Call vsfProcessList_Click
            If txtName.Locked = True Then
                txtName.ForeColor = &H80000011
                txtDescription.ForeColor = &H80000011
            End If
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '若当前正在进行新增或修改操作，且已经对内容进行了修改，则弹出提示，是否需要保存
    If UnloadMode = vbFormControlMenu Or UnloadMode = vbFormCode Then
        Select Case CheckChange
            Case 2    '进行了修改，但是选择了保存，且保存失败了的
                Cancel = 1
        End Select
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picEdit.Height = Me.ScaleHeight - picEdit.Top
    picEdit.Left = Me.ScaleWidth - picEdit.Width
    vsfProcessList.Width = Me.ScaleWidth - picEdit.Width
    vsfProcessList.Height = picEdit.Height
    cmdAdd.Top = picEdit.Height - cmdAdd.Height - 255
    cmdUpdate.Top = cmdAdd.Top
    cmdDel.Top = cmdAdd.Top
    cmdExit.Top = cmdAdd.Top
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnRightClick = False
    mblnOk = False
End Sub

Private Sub cmdAdd_Click()
    Call mnuPopuAdd_Click
End Sub

Private Sub cmdUpdate_Click()
    Call mnuPopuModify_Click
End Sub

Private Sub cmdDel_Click()
    Call mnuPopuDelete_Click
End Sub

'新增服务或进程信息
Private Sub mnuPopuAdd_Click()
    mnuPopuAdd.Tag = 1   '标记当前状态为新增
    Call FillSelectData(-1)
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDel.Enabled = False
    Call SetEnable(True)
    txtName.SetFocus
End Sub

'删除服务或进程信息
Private Sub mnuPopuDelete_Click()
    On Error GoTo errH
    If MsgBox("确定要将名称为“" & vsfProcessList.TextMatrix(vsfProcessList.Row, PL_名称) & "”的这个" & vsfProcessList.TextMatrix(vsfProcessList.Row, PL_类型) & "删除吗？", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
        Call ExecuteProcedure("Zl_Zlkillprocess_Edit(3, Null, '" & vsfProcessList.TextMatrix(vsfProcessList.Row, PL_名称) & "')", Me.Caption)
        vsfProcessList.RemoveItem (vsfProcessList.Row)
        Call FillSelectData(vsfProcessList.Row)
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'修改服务或进程信息
Private Sub mnuPopuModify_Click()
    mnuPopuAdd.Tag = 2  '标记当前状态为修改
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdDel.Enabled = False
    Call SetEnable(True)
    txtName.SetFocus
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    '控制文本长度，以及屏蔽换行符
    If (ActualLen(txtDescription.Text) >= 200 And KeyAscii <> 8) Or KeyAscii = 9 Or KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    '不能输入\/:*?"'<>|
    If InStr("\/:*?""<>|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_LostFocus()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset

    lblAdd.Caption = ""
    If txtName.Locked = True Then Exit Sub
    If Val(mnuPopuAdd.Tag) = 1 Then
        '新增状态下，检查该名称是否有重复的，如果有，在下方标签页上给出文字提示
        strSQL = "Select Count(1) 数量 From Zlkillprocess Where 名称 = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, UCase(Trim(txtName.Text)))
        If rsTmp!数量 = 1 Then
            lblShowCheckName.Caption = "名称已存在！"
        Else
            lblShowCheckName.Caption = ""
        End If
    Else
        '修改状态下，先检查名称有没有被修改，若被修改了，再检查该名称是否有重复，如果有，在下方标签页上给出文字提示
        If txtName.Text <> txtName.Tag Then
            strSQL = "Select Count(1) 数量 From Zlkillprocess Where 名称 = [1]"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, UCase(Trim(txtName.Text)))
            If rsTmp!数量 = 1 Then
                lblShowCheckName.Caption = "名称已存在！"
            Else
                lblShowCheckName.Caption = ""
            End If
        Else
            lblShowCheckName.Caption = ""
        End If
    End If
End Sub

Private Sub vsfProcessList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    '在换行之前，先进行数据检查，看是否对数据进行了修改
    Select Case CheckChange
        Case 1, 4  '进行了修改，但是选择了保存，且保存成功了的 或 未进行修改
            Call FillSelectData(NewRow)
            Call SetEnable(False)
        Case 2   '进行了修改，但是选择了保存，且保存失败了的
            Cancel = True
        Case 3   '进行了修改，但是选择了不保存
            Call FillSelectData(OldRow)
            Cancel = True
            Call SetEnable(False)
    End Select
End Sub

Private Sub vsfProcessList_Click()
    If mblnRightClick Then Exit Sub
    With vsfProcessList
        '还原弹出菜单标记及启停状态
        '如果不还原的话，若当前正在新增或修改模式下，但并没有修改内容，这时右键切换行，按钮启停状态就会有误
        If txtName.Locked = True Then
            mnuPopuAdd.Tag = ""
            mnuPopuAdd.Enabled = True
            mnuPopuDelete.Enabled = .TextMatrix(.Row, PL_分类) <> "固定"
            mnuPopuModify.Enabled = mnuPopuDelete.Enabled
            cmdAdd.Enabled = mnuPopuAdd.Enabled
            cmdUpdate.Enabled = mnuPopuModify.Enabled
            cmdDel.Enabled = mnuPopuDelete.Enabled
        End If
    End With
End Sub

Private Sub vsfProcessList_DblClick()
    With vsfProcessList
        If .MouseRow <> .Row Or vsfProcessList.TextMatrix(vsfProcessList.Row, PL_分类) = "固定" Then Exit Sub
        Call mnuPopuModify_Click
    End With
End Sub

Private Sub vsfProcessList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '若该进程为系统内置的进程，则不允许进行修改和删除操作
    mnuPopuDelete.Enabled = vsfProcessList.TextMatrix(vsfProcessList.Row, PL_分类) <> "固定"
    mnuPopuModify.Enabled = mnuPopuDelete.Enabled
    '判断其是否属于新增或修改状态，若是，则不允许再次进行新增、修改或删除操作
    If txtName.Locked = False Then
        mnuPopuModify.Enabled = False
        mnuPopuAdd.Enabled = False
        mnuPopuDelete.Enabled = False
    End If
    cmdAdd.Enabled = mnuPopuAdd.Enabled
    cmdUpdate.Enabled = mnuPopuModify.Enabled
    cmdDel.Enabled = mnuPopuDelete.Enabled
    
    With vsfProcessList
        '右键某一项
        If .MouseRow <> -1 And .MouseRow <> 0 And Button = 2 Then
            If .MouseRow <> .Row Then
                '若选择的是不同行，则将其选中
                .Row = .MouseRow
                mblnRightClick = False
                Call vsfProcessList_Click
            End If
        End If
    End With
End Sub

Private Sub vsfProcessList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnRightClick = False
    If Button = 1 Then Exit Sub
    With vsfProcessList
        If .MouseRow <> .Row Then Exit Sub
        mblnRightClick = True
        PopupMenu mnuPopu
    End With
End Sub

Private Sub SetEnable(ByVal blnEnable As Boolean)
    With vsfProcessList
        txtName.Locked = Not blnEnable
        cboType.Enabled = blnEnable
        txtDescription.Locked = Not blnEnable
        cmdOK.Visible = blnEnable
        cmdCancel.Visible = blnEnable
        '如果当前编辑界面是锁定状态，就将前景色置灰
        If txtName.Locked = True Then
            txtName.ForeColor = &H80000011
            txtDescription.ForeColor = &H80000011
        Else
            txtName.ForeColor = &H80000009
            txtDescription.ForeColor = &H80000009
        End If
    End With
    lblShowCheckName.Caption = ""
End Sub

Private Function CheckChange() As Long
    '判断当前行是否已经被修改，如果是，则弹出提示
    '若进行了修改，但是选择了保存，且保存成功了的，返回1
    '若进行了修改，但是选择了保存，且保存失败了的，返回2
    '若进行了修改，但是选择了不保存，返回3
    '若未进行修改，返回4
    If txtName.Text <> txtName.Tag Or cboType.Text <> cboType.Tag Or txtDescription.Text <> txtDescription.Tag Then
        If MsgBox("当前项已被修改，是否保存？", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbOK Then
            Call cmdOK_Click
            If mblnOk Then
                CheckChange = 1
            Else
                CheckChange = 2
            End If
        Else
            CheckChange = 3
        End If
    Else
        CheckChange = 4
    End If
End Function

Private Function CheckData() As Boolean
    '判断当前输入数据是否符合要求
    If Right(UCase(Trim(txtName.Text)), 4) <> ".EXE" Then
        MsgBox "名称不是一个可执行文件（*.EXE）,请进行调整！", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Function
    End If
    If InStr(txtName.Text, "'") > 0 Or InStr(txtName.Text, """") > 0 Or InStr(txtName.Text, "\") > 0 Or _
        InStr(txtName.Text, "/") > 0 Or InStr(txtName.Text, ":") > 0 Or InStr(txtName.Text, "*") > 0 Or _
        InStr(txtName.Text, "?") > 0 Or InStr(txtName.Text, "<") > 0 Or InStr(txtName.Text, ">") > 0 Or InStr(txtName.Text, "|") > 0 Then
        MsgBox "该名称中含有非法字符(\/:*?""'<>|)，请重新填写！", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Function
    End If
    If StrIsValid(txtDescription.Text, 200) = False Then
        txtDescription.SetFocus
        Exit Function
    End If
    '若描述中含有换行符，将其去掉
    If InStr(txtDescription.Text, vbNewLine) > 0 Then
        txtDescription.Text = Replace(txtDescription.Text, vbNewLine, "")
    End If
    CheckData = True
End Function

Private Sub FillSelectData(ByVal lngRow As Long)
'lngRow:vsfProcessList的行号
    If lngRow = -1 Then
        txtName.Text = ""
        cboType.ListIndex = 0
        txtDescription.Text = ""
        txtName.Tag = txtName.Text
        cboType.Tag = cboType.Text
        txtDescription.Tag = txtDescription.Text
    Else
        txtName.Text = vsfProcessList.TextMatrix(lngRow, PL_名称)
        cboType.Text = vsfProcessList.TextMatrix(lngRow, PL_类型)
        txtDescription.Text = vsfProcessList.TextMatrix(lngRow, PL_描述)
        txtName.Tag = txtName.Text
        cboType.Tag = cboType.Text
        txtDescription.Tag = txtDescription.Text
    End If
End Sub
