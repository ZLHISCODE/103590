VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLabMainSetQC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置质控"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmLabMainSetQC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk仪器维护更新 
      Caption         =   "刚进行仪器维护更新"
      Height          =   210
      Left            =   4845
      TabIndex        =   12
      Top             =   2865
      Width           =   1950
   End
   Begin VB.CheckBox chk新包装控制物 
      Caption         =   "使用了新包装控制物"
      Height          =   210
      Left            =   4845
      TabIndex        =   11
      Top             =   2580
      Width           =   1950
   End
   Begin VB.CheckBox chk新包装校准物 
      Caption         =   "使用了新包装校准物"
      Height          =   210
      Left            =   2415
      TabIndex        =   10
      Top             =   2865
      Width           =   1935
   End
   Begin VB.CheckBox chk新批号校准物 
      Caption         =   "使用了新批号校准物"
      Height          =   210
      Left            =   2415
      TabIndex        =   9
      Top             =   2580
      Width           =   1935
   End
   Begin VB.CheckBox chk新包装试剂 
      Caption         =   "使用了新包装试剂"
      Height          =   210
      Left            =   150
      TabIndex        =   8
      Top             =   2865
      Width           =   1935
   End
   Begin VB.CheckBox chk新批号试剂 
      Caption         =   "使用了新批号试剂"
      Height          =   210
      Left            =   150
      TabIndex        =   7
      Top             =   2580
      Width           =   1935
   End
   Begin VB.ComboBox cbo质控品 
      Height          =   300
      ItemData        =   "frmLabMainSetQC.frx":000C
      Left            =   3555
      List            =   "frmLabMainSetQC.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   510
      Width           =   3225
   End
   Begin VB.TextBox txt标本号 
      Enabled         =   0   'False
      Height          =   300
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   1155
   End
   Begin VB.TextBox txt检验人 
      Enabled         =   0   'False
      Height          =   300
      Left            =   720
      TabIndex        =   4
      Top             =   510
      Width           =   1155
   End
   Begin VB.TextBox txt仪器 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3555
      TabIndex        =   3
      Top             =   120
      Width           =   3225
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5490
      TabIndex        =   2
      Top             =   3420
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   1
      Top             =   3420
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   15
      Left            =   60
      TabIndex        =   0
      Top             =   3240
      Width           =   6795
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgRecord 
      Height          =   1530
      Left            =   150
      TabIndex        =   13
      Top             =   900
      Width           =   6630
      _cx             =   11695
      _cy             =   2699
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
   Begin VB.Label lbl检验仪器 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "仪器"
      Height          =   180
      Left            =   3120
      TabIndex        =   17
      Top             =   180
      Width           =   360
   End
   Begin VB.Label lbl标本号 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标本号"
      Height          =   180
      Left            =   150
      TabIndex        =   16
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lbl质控品 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "质控品"
      Height          =   180
      Left            =   2940
      TabIndex        =   15
      Top             =   570
      Width           =   540
   End
   Begin VB.Label lbl检验人 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检验人"
      Height          =   180
      Left            =   150
      TabIndex        =   14
      Top             =   570
      Width           =   540
   End
End
Attribute VB_Name = "frmLabMainSetQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long
Private mstrSelDate As String
Private mblnAllDev As Boolean
Private mMachineID As Long

Private Enum mCol
    中文名 = 0: 英文名: 单位: 结果值: ID
End Enum

'临时变量
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '功能：初始化设置参考值列表
    '参数： blnKeepData-是否保留数据，即只是重新设置格式
    With Me.vfgRecord
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 2: .FixedRows = 1: .Cols = 5: .FixedCols = 0
            .TextMatrix(0, mCol.中文名) = "中文名"
            .TextMatrix(0, mCol.英文名) = "英文名"
            .TextMatrix(0, mCol.单位) = "单位"
            .TextMatrix(0, mCol.结果值) = "结果值"
            .TextMatrix(0, mCol.ID) = "ID"
        End If
        .ColWidth(mCol.中文名) = 3000
        .ColWidth(mCol.英文名) = 1200
        .ColWidth(mCol.单位) = 1000
        .ColWidth(mCol.结果值) = 900
        .ColWidth(mCol.ID) = 0
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngID As Long) As Boolean
    '功能：根据id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset
    mlngID = lngID
    
    '清除此前项目的显示
    
    Me.txt标本号.Text = "": Me.txt标本号.Tag = "": Me.txt仪器.Text = ""
    Me.txt检验人.Text = "": Me.cbo质控品.Clear
    Me.chk新批号试剂.Value = vbUnchecked: Me.chk新包装试剂.Value = vbUnchecked
    Me.chk新批号校准物.Value = vbUnchecked: Me.chk新包装校准物.Value = vbUnchecked
    Me.chk新包装控制物.Value = vbUnchecked: Me.chk仪器维护更新.Value = vbUnchecked
    
    If lngID = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select L.标本序号, A.名称 As 仪器, L.质控品id, M.批号 || '-' || M.名称 As 质控品, L.检验人, L.新批号试剂," & vbNewLine & _
            "       L.新包装试剂, L.新批号校准物, L.新包装校准物, L.新包装控制物, L.仪器维护更新" & vbNewLine & _
            "From 检验质控记录 L, 检验仪器 A, 检验质控品 M" & vbNewLine & _
            "Where L.仪器id = A.ID And L.质控品id = M.ID And L.标本id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt标本号.Text = "" & !标本序号: Me.txt标本号.Tag = lngID: Me.txt仪器.Text = "" & !仪器
            Me.txt检验人.Text = "" & !检验人
            Me.cbo质控品.AddItem "" & !质控品
            Me.cbo质控品.ItemData(Me.cbo质控品.NewIndex) = Val("" & !质控品id)
            Me.cbo质控品.ListIndex = Me.cbo质控品.NewIndex
            Me.chk新批号试剂.Value = IIf(Val("" & !新批号试剂) = 0, vbUnchecked, vbChecked)
            Me.chk新包装试剂.Value = IIf(Val("" & !新包装试剂) = 0, vbUnchecked, vbChecked)
            Me.chk新批号校准物.Value = IIf(Val("" & !新批号校准物) = 0, vbUnchecked, vbChecked)
            Me.chk新包装校准物.Value = IIf(Val("" & !新包装校准物) = 0, vbUnchecked, vbChecked)
            Me.chk新包装控制物.Value = IIf(Val("" & !新包装控制物) = 0, vbUnchecked, vbChecked)
            Me.chk仪器维护更新.Value = IIf(Val("" & !仪器维护更新) = 0, vbUnchecked, vbChecked)
        End If
    End With
        
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function ZlEditStart(blnAdd As Boolean, lngID As Long, Optional strSelDate As String, Optional blnAllDev As Boolean) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngID-当前编辑的标本id，或者当前选中的标本id
    '       strDate-指定日期
    '       blnAllDev-是否有所有设备的权限，否则只能是本部门的仪器
    Dim rsTemp As New ADODB.Recordset
    
    If blnAdd Then
        Err = 0: On Error Resume Next
        mstrSelDate = Format(strSelDate, "yyyy-MM-dd")
        If Err <> 0 Or mstrSelDate = "" Then ZlEditStart = False: Exit Function
        Err = 0: On Error GoTo 0
        mblnAllDev = blnAllDev
    End If
    
    mlngID = lngID
    If blnAdd Then
        Me.txt标本号.Text = "": Me.txt标本号.Tag = "": Me.txt仪器.Text = ""
        Me.txt检验人.Text = "": Me.cbo质控品.Clear
        Me.chk新批号试剂.Value = vbUnchecked: Me.chk新包装试剂.Value = vbUnchecked
        Me.chk新批号校准物.Value = vbUnchecked: Me.chk新包装校准物.Value = vbUnchecked
        Me.chk新包装控制物.Value = vbUnchecked: Me.chk仪器维护更新.Value = vbUnchecked
        Call setListFormat(False)
    Else
        If Me.cbo质控品.ListIndex = -1 Then
            Me.cbo质控品.Tag = 0
        Else
            Me.cbo质控品.Tag = Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex)
        End If
        Me.cbo质控品.Clear
        gstrSql = "Select Distinct M.ID, M.批号 || '-' || M.名称 || LPad('^,' || M.标本号, 200, ' ') As 质控品" & vbNewLine & _
                "From 检验标本记录 L, 检验普通结果 R, 检验质控品 M, 检验质控品项目 I" & vbNewLine & _
                "Where L.ID = R.检验标本id And Nvl(L.报告结果, 0) = Nvl(R.记录类型, 0) And L.仪器id = M.仪器id And M.ID = I.质控品id And" & vbNewLine & _
                "      Nvl(R.弃用结果,0)=0 And I.项目id = R.检验项目id And (L.检验时间 + 0 Between M.开始日期 And M.结束日期 + 1 - 1 / 86400) And L.ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID)
        With rsTemp
            Do While Not .EOF
                Me.cbo质控品.AddItem "" & !质控品
                Me.cbo质控品.ItemData(Me.cbo质控品.NewIndex) = Val("" & !ID)
                If Val(Me.cbo质控品.Tag) = Val("" & !ID) Then Me.cbo质控品.ListIndex = Me.cbo质控品.NewIndex
                .MoveNext
            Loop
            If Me.cbo质控品.ListCount > 0 And Me.cbo质控品.ListIndex = -1 Then Me.cbo质控品.ListIndex = 0
        End With
    End If
    
    Me.Tag = IIf(blnAdd, "增加", "修改"): Call Form_Resize
    If blnAdd Then
'        Me.cmdSelect.SetFocus
    Else
        Me.cbo质控品.SetFocus
    End If
    ZlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ZlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim strInfo As String
    
    If Me.cbo质控品.ListIndex = -1 Then MsgBox "未选择质控品！", vbInformation, gstrSysName: zlEditSave = 0: Exit Function
    
    strInfo = Split(Me.cbo质控品.Text, "^,")(1)
    If Trim(strInfo) <> "" And InStr(1, "," & strInfo & ",", "," & Trim(Me.txt标本号.Text) & ",") = 0 Then
        strInfo = "当前标本号与质控品规定的样本号不符，请检查："
        strInfo = strInfo & vbCrLf & "   选择质控品的浓度水平是否符合？"
        strInfo = strInfo & vbCrLf & vbCrLf & "如果确认正确，选择“是”继续。"
        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then zlEditSave = 0: Exit Function
    End If

    gstrSql = "Zl_检验质控记录_Edit(" & Me.Tag
    gstrSql = gstrSql & "," & Val(Me.txt标本号.Tag) & "," & Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex)
    gstrSql = gstrSql & "," & IIf(Me.chk新批号试剂.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk新包装试剂.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk新批号校准物.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk新包装校准物.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk新包装控制物.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk仪器维护更新.Value = vbChecked, 1, 0) & ")"
    
    Err = 0: On Error GoTo ErrHand
    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = Val(Me.txt标本号.Tag)
    frmQCCompute.ShowMe Me, mMachineID, Me.vfgRecord.TextMatrix(1, mCol.ID), zlDatabase.Currentdate, Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex)
    Unload Me: Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cbo质控品_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lngID As Long, lngResId As Long
    
    lngID = Val(Me.txt标本号.Tag)
    If lngID = 0 Then Call setListFormat(False): Exit Sub
    
    If Me.cbo质控品.ListIndex = -1 Then Exit Sub
    lngResId = Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex)
    
    Err = 0: On Error GoTo ErrHand
    If Trim(Me.Tag) = "" Then
        gstrSql = "Select I.中文名, I.英文名, I.单位, R.检验结果 As 结果值,R.检验项目ID as ID" & vbNewLine & _
            "From 检验普通结果 R, 诊治所见项目 I" & vbNewLine & _
            "Where R.检验项目id = I.ID And R.是否检验 = 1 And Nvl(R.弃用结果,0)=0 And R.检验标本id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID)
    Else
        gstrSql = "Select I.中文名, I.英文名, I.单位, R.检验结果 As 结果值,R.检验项目ID as ID" & vbNewLine & _
                "From 检验普通结果 R, 诊治所见项目 I, (Select 项目id From 检验质控品项目 Where 质控品id = [2]) T" & vbNewLine & _
                "Where R.检验项目id = I.ID And R.检验项目id = T.项目id And Nvl(R.弃用结果,0)=0 And R.检验标本id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID, lngResId)
    End If
    Set Me.vfgRecord.DataSource = rsTemp: Call setListFormat(True)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo质控品_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk弃用_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk新包装控制物_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk新包装试剂_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk新包装校准物_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk新批号试剂_Click()
    If Me.chk新批号试剂.Value = vbChecked Then
        Me.chk新包装试剂.Value = vbChecked: Me.chk新包装试剂.Enabled = False
    Else
        Me.chk新包装试剂.Enabled = True
    End If
End Sub

Private Sub chk新批号试剂_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk新批号校准物_Click()
    If Me.chk新批号校准物.Value = vbChecked Then
        Me.chk新包装校准物.Value = vbChecked: Me.chk新包装校准物.Enabled = False
    Else
        Me.chk新包装校准物.Enabled = True
    End If
End Sub

Private Sub chk新批号校准物_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk仪器维护更新_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub



Private Sub cmdSelect_Click()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    zlEditSave
End Sub

Private Sub Form_Load()
    Call setListFormat(False)
End Sub

Private Sub Form_Resize()
    Select Case Trim(Me.Tag)
    Case "1"    '增加
        Me.Enabled = True: Me.BackColor = RGB(250, 250, 250)
        'Me.cmdSelect.Visible = True
    Case "2"    '修改
        Me.Enabled = True: Me.BackColor = RGB(250, 250, 250)
        'Me.cmdSelect.Visible = False
    Case Else   '删除
'        Me.Enabled = False: Me.BackColor = &H8000000F
        'Me.cmdSelect.Visible = False
    End Select
    Me.chk新批号试剂.BackColor = Me.BackColor: Me.chk新包装试剂.BackColor = Me.BackColor
    Me.chk新批号校准物.BackColor = Me.BackColor: Me.chk新包装校准物.BackColor = Me.BackColor
    Me.chk新包装控制物.BackColor = Me.BackColor: Me.chk仪器维护更新.BackColor = Me.BackColor
    
'    Me.chk新包装试剂.Top = Me.ScaleHeight - 300
'    Me.chk新包装校准物.Top = Me.ScaleHeight - 300
'    Me.chk仪器维护更新.Top = Me.ScaleHeight - 300
'    Me.chk新批号试剂.Top = Me.ScaleHeight - 300 * 2
'    Me.chk新批号校准物.Top = Me.ScaleHeight - 300 * 2
'    Me.chk新包装控制物.Top = Me.ScaleHeight - 300 * 2
'    Me.vfgRecord.Height = Me.chk新批号试剂.Top - Me.vfgRecord.Top - 75
    
End Sub

Public Sub ShowMe(Objfrm As Object, lngSampleID As Long, SampleNumber, MachineID, VerifyName As String, EditMode As Integer)
    '功能:              增加修改删除质控
    '参数:              objfrm 上级窗体传入窗体对象
    '                   lngSampleID 标本ID
    '                   SampleNumber    标本序号
    '                   MachineID       仪器ID
    '                   VerifyName      检验人
    '                   EditMode        编辑模式: 1=增加 2=修改 3=删除
    
    Dim rsTemp As New ADODB.Recordset
    
    Me.Tag = EditMode
    mMachineID = MachineID
    gstrSql = "select 名称 from 检验仪器 where id = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, MachineID)
    
    Me.txt标本号.Text = SampleNumber: Me.txt标本号.Tag = lngSampleID
    If MachineID > 0 Then Me.txt仪器.Text = rsTemp("名称"): Me.txt检验人.Text = VerifyName
    
        
    If Me.cbo质控品.ListIndex = -1 Then
        Me.cbo质控品.Tag = 0
    Else
        Me.cbo质控品.Tag = Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex)
    End If
    Me.cbo质控品.Clear
    
    gstrSql = "Select Distinct M.ID, M.批号 || '-' || M.名称 || LPad('^,' || M.标本号, 200, ' ') As 质控品" & vbNewLine & _
            "From 检验标本记录 L, 检验普通结果 R, 检验质控品 M, 检验质控品项目 I" & vbNewLine & _
            "Where L.ID = R.检验标本id And Nvl(L.报告结果, 0) = Nvl(R.记录类型, 0) And L.仪器id = M.仪器id And M.ID = I.质控品id And" & vbNewLine & _
            "     Nvl(R.弃用结果,0)=0 And I.项目id = R.检验项目id And (L.检验时间 + 0 Between M.开始日期 And M.结束日期 + 1 - 1 / 86400) And L.ID = [1]"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.txt标本号.Tag))
    With rsTemp
        Do While Not .EOF
            Me.cbo质控品.AddItem "" & !质控品
            Me.cbo质控品.ItemData(Me.cbo质控品.NewIndex) = Val("" & !ID)
            If Val(Me.cbo质控品.Tag) = Val("" & !ID) Then Me.cbo质控品.ListIndex = Me.cbo质控品.NewIndex
            .MoveNext
        Loop
        If Me.cbo质控品.ListCount > 0 And Me.cbo质控品.ListIndex = -1 Then Me.cbo质控品.ListIndex = 0
    End With
    
    Me.Show vbModal, Objfrm
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub






