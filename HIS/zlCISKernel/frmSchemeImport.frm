VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmSchemeImport 
   AutoRedraw      =   -1  'True
   Caption         =   "病人医嘱导入"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   Icon            =   "frmSchemeImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelNone 
      Caption         =   "全清(&R)"
      Height          =   350
      Left            =   1815
      TabIndex        =   8
      ToolTipText     =   "Ctrl+R"
      Top             =   5655
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelALL 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "Ctrl+A"
      Top             =   5655
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8160
      TabIndex        =   10
      Top             =   5655
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7065
      TabIndex        =   9
      Top             =   5655
      Width           =   1100
   End
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   135
      Width           =   3630
   End
   Begin VB.TextBox txtPati 
      Height          =   300
      Left            =   735
      TabIndex        =   1
      ToolTipText     =   "输入方法：刷卡，-病人ID，+住院号，*门诊号，.挂号单"
      Top             =   135
      Width           =   1275
   End
   Begin VB.ComboBox cboBaby 
      Height          =   300
      Left            =   7530
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   135
      Width           =   2130
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4965
      Left            =   60
      TabIndex        =   6
      Top             =   570
      Width           =   9585
      _cx             =   16907
      _cy             =   8758
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   23
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSchemeImport.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "就诊时间(&T)"
      Height          =   180
      Left            =   2115
      TabIndex        =   2
      Top             =   195
      Width           =   990
   End
   Begin VB.Label lblPati 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人(&P)"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lblBaby 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "婴儿(&B)"
      Height          =   180
      Left            =   6870
      TabIndex        =   4
      Top             =   195
      Width           =   630
   End
End
Attribute VB_Name = "frmSchemeImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint范围 As Integer '1-门诊,2-住院,3-门诊和住院
Private mlng病人ID As Long
Private mstrIDs As String
Private mblnOK As Boolean

Private Enum COL成套方案
    col选择 = 0
    col期效 = 1
    col内容 = 2
    col单量 = 3
    col单位 = 4
    col总量 = 5
    col总量单位 = 6
    col频次 = 7
    col用法 = 8
    col嘱托 = 9
    col执行时间 = 10
    col执行科室 = 11
    col执行性质 = 12
    col序号 = 13
    col相关 = 14
    col项目ID = 15
    col类别 = 16
    col收费细目ID = 17
    col标本部位 = 18
    col检查方法 = 19
    col频率次数 = 20
    col频率间隔 = 21
    col间隔单位 = 22
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal int范围 As Integer, lng病人ID As Long) As String
'返回：所选择的医嘱组ID
'      lng病人ID=导入医嘱的病人
    
    mint范围 = int范围
    lng病人ID = 0
    
    Me.Show 1, frmParent
    If mblnOK Then
        ShowMe = mstrIDs
        lng病人ID = mlng病人ID
    End If
End Function

Private Sub cboBaby_Click()
    Call LoadAdvice
End Sub

Private Sub cboTime_Click()
    Call LoadAdviceBaby
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strIDs As String, i As Long
    Dim lng医嘱ID As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col选择)) <> 0 And Val(.TextMatrix(i, col序号)) <> 0 Then
                lng医嘱ID = IIF(Val(.TextMatrix(i, col相关)) = 0, Val(.TextMatrix(i, col序号)), Val(.TextMatrix(i, col相关)))
                If InStr(strIDs & ",", "," & lng医嘱ID & ",") = 0 Then
                    strIDs = strIDs & "," & lng医嘱ID
                End If
            End If
        Next
        If strIDs = "" Then
            MsgBox "没有选择任何医嘱内容。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    mstrIDs = Mid(strIDs, 2)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSelALL_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col项目ID)) <> 0 Then
                '以前的检查医嘱不允许保存为成套方案
                If .TextMatrix(i, col类别) = "D" Then
                    If Val(.TextMatrix(i, col相关)) = 0 Then
                        If Not CheckIsOldAdvice(i) Then
                            .TextMatrix(i, col选择) = -1
                            Call RowSelectSame(i)
                        End If
                    Else
                        '主项行已处理
                    End If
                Else
                    .TextMatrix(i, col选择) = -1
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdSelNone_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, col选择) = 0
        Next
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdSelALL_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdSelNone_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    mstrIDs = ""
    mblnOK = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsAdvice.Width = Me.ScaleWidth - vsAdvice.Left * 2
    
    If vsAdvice.Left + vsAdvice.Width - cboBaby.Width > cboTime.Left + cboTime.Width + lblBaby.Width + 150 Then
        cboBaby.Left = vsAdvice.Left + vsAdvice.Width - cboBaby.Width
    Else
        cboBaby.Left = cboTime.Left + cboTime.Width + lblBaby.Width + 150
    End If
    lblBaby.Left = cboBaby.Left - lblBaby.Width - 30
    
    If Me.ScaleWidth - cmdCancel.Width - cmdSelALL.Left > cmdSelNone.Left + cmdSelNone.Width + cmdOK.Width Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdSelALL.Left
    Else
        cmdCancel.Left = cmdSelNone.Left + cmdSelNone.Width + cmdOK.Width
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
    
    vsAdvice.Height = Me.ScaleHeight - vsAdvice.Top - cmdSelNone.Height * 1.6
    cmdSelNone.Top = vsAdvice.Top + vsAdvice.Height + cmdSelNone.Height * 0.3
    cmdSelALL.Top = cmdSelNone.Top
    cmdOK.Top = cmdSelNone.Top
    cmdCancel.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    
    '输入号码后回车
    If KeyAscii = 13 And Trim(txtPati.Text) <> "" Then
        KeyAscii = 0
        
        '读取病人信息
        If Not GetPatient(Trim(txtPati.Text)) Then
            txtPati.PasswordChar = ""
            txtPati.Text = lblPati.Tag
            Call zlControl.TxtSelAll(txtPati)
        Else
            txtPati.PasswordChar = ""
            vsAdvice.SetFocus
        End If
    End If
End Sub

Private Function GetPatient(ByVal strCode As String) As Boolean
'功能：读取病人信息，并显示该病人存在的医嘱时间
'参数：
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strNO As String, str姓名 As String, lng病人ID As Long
    
    On Error GoTo errH
    
    If Left(strCode, 1) = "-" And IsNumeric(Mid(strCode, 2)) Then '病人ID
        strSQL = "Select 病人ID,姓名 From 病人信息 Where 病人ID=[3] "
    ElseIf Left(strCode, 1) = "+" And IsNumeric(Mid(strCode, 2)) Then '住院号
        strSQL = "Select 病人ID,姓名 From 病人信息 Where 住院号=[3] "
    ElseIf Left(strCode, 1) = "*" And IsNumeric(Mid(strCode, 2)) Then '门诊号
        strSQL = "Select 病人ID,姓名 From 病人信息 Where 门诊号=[3] "
    ElseIf Left(strCode, 1) = "." Then '挂号单
        strNO = GetFullNO(Mid(UCase(strCode), 2), 12)
        strSQL = "Select 病人ID,姓名 From 病人挂号记录 Where NO=[4] And 记录性质=1 And 记录状态=1"
    Else '当作姓名
        strSQL = "Select 病人ID,姓名 From 病人信息 Where 姓名=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode, UCase(strCode), Mid(strCode, 2), strNO)
    
    If rsTmp.EOF Then
        mlng病人ID = 0
        MsgBox "没有找到相关的病人信息。", vbInformation, gstrSysName
    Else
        str姓名 = rsTmp!姓名: lng病人ID = rsTmp!病人ID
        strSQL = _
            " Select B.ID as 挂号ID,A.挂号单,B.登记时间,D.名称 as 挂号科室," & _
            " A.主页ID,C.入院日期,E.名称 as 住院科室,Min(A.开嘱时间) as 顺序" & _
            " From 病人医嘱记录 A,病人挂号记录 B,病案主页 C,部门表 D,部门表 E" & _
            " Where A.病人ID=[1] And A.挂号单=B.NO(+) And B.记录性质(+)=1 And B.记录状态(+)=1" & _
            Decode(mint范围, 1, " And A.病人来源=1", 2, " And A.病人来源=2", 3, " And A.病人来源 Not IN(3,4)") & _
            " And A.病人ID=C.病人ID(+) And A.主页ID=C.主页ID(+)" & _
            " And B.执行部门ID=D.ID(+) And C.出院科室ID=E.ID(+)" & _
            " Group by B.ID,A.挂号单,B.登记时间,D.名称,A.主页ID,C.入院日期,E.名称" & _
            " Order by 顺序 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
        
        If rsTmp.EOF Then
            mlng病人ID = 0
            MsgBox "病人""" & str姓名 & """没有医嘱记录。", vbInformation, gstrSysName
        Else
            txtPati.Text = str姓名
            lblPati.Tag = str姓名
            mlng病人ID = lng病人ID
            cboTime.Clear
            For i = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!主页ID) And Not IsNull(rsTmp!挂号ID) Then
                    cboTime.AddItem rsTmp!挂号科室 & " " & Format(rsTmp!登记时间, "yyyy-MM-dd HH:mm") & " 门诊就诊"
                    cboTime.ItemData(cboTime.NewIndex) = rsTmp!挂号ID
                    If rsTmp!挂号单 = strNO Then cboTime.ListIndex = cboTime.NewIndex
                ElseIf Not IsNull(rsTmp!主页ID) And IsNull(rsTmp!挂号单) Then
                    cboTime.AddItem rsTmp!住院科室 & " " & Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm") & " 第" & rsTmp!主页ID & "次住院"
                    cboTime.ItemData(cboTime.NewIndex) = rsTmp!主页ID
                End If
                rsTmp.MoveNext
            Next
            If cboTime.ListIndex = -1 Then cboTime.ListIndex = 0
            GetPatient = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadAdviceBaby() As Boolean
'功能：读取当前病人指定时间医嘱的婴儿清单
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNO As String
    
    If Val(mlng病人ID) = 0 Then Exit Function
    If cboTime.ListIndex = -1 Then Exit Function
    
    On Error GoTo errH
    
    If InStr(cboTime.Text, "住院") = 0 Then
        strSQL = "Select Distinct Nvl(A.婴儿,0) as 婴儿,C.婴儿姓名" & _
            " From 病人医嘱记录 A,病人挂号记录 B,病人新生儿记录 C" & _
            " Where A.病人ID=[1] And A.挂号单=B.NO And B.ID=[2]" & _
            " And A.婴儿=C.序号(+) And C.病人ID(+)=[1] And C.主页ID(+)=[2]" & _
            " Order by 婴儿"
    Else
        strSQL = "Select Distinct Nvl(A.婴儿,0) as 婴儿,C.婴儿姓名" & _
            " From 病人医嘱记录 A,病人新生儿记录 C" & _
            " Where A.病人ID=[1] And A.主页ID=[2]" & _
            " And A.婴儿=C.序号(+) And C.病人ID(+)=[1] And C.主页ID(+)=[2]" & _
            " Order by 婴儿"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, cboTime.ItemData(cboTime.ListIndex))
    
    cboBaby.Clear
    Do While Not rsTmp.EOF
        If NVL(rsTmp!婴儿, 0) = 0 Then
            cboBaby.AddItem "病人医嘱"
        Else
            cboBaby.AddItem "婴儿 " & rsTmp!婴儿 & IIF(IsNull(rsTmp!婴儿姓名), " 医嘱", "：" & NVL(rsTmp!婴儿姓名))
        End If
        cboBaby.ItemData(cboBaby.NewIndex) = NVL(rsTmp!婴儿, 0)
        rsTmp.MoveNext
    Loop
    If cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0
    LoadAdviceBaby = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadAdvice() As Boolean
'功能：读取当前病人指定的医嘱
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    If mlng病人ID = 0 Then Exit Function
    If cboTime.ListIndex = -1 Then Exit Function
    If cboBaby.ListIndex = -1 Then Exit Function
    
    On Error GoTo errH
    
    If InStr(cboTime.Text, "住院") = 0 Then
        strSQL = "Select Distinct A.ID,A.序号,A.相关ID,A.医嘱期效,A.诊疗项目ID,A.医嘱内容," & _
            " A.单次用量,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.医生嘱托,A.执行性质,A.执行标记," & _
            " Nvl(C.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'-')) as 执行科室,A.执行科室id,A.执行时间方案," & _
            " A.执行科室ID,A.标本部位,A.检查方法,Nvl(B.类别,'*') as 类别,B.名称,B.计算单位," & _
            " A.总给予量 as 总量,D.计算单位 as 总量单位,D.id as 收费细目ID" & _
            " From 病人医嘱记录 A,诊疗项目目录 B,部门表 C,收费项目目录 D,病人挂号记录 R" & _
            " Where A.诊疗项目ID=B.ID(+) And A.执行科室ID=C.ID(+) And A.收费细目ID=D.ID(+)" & _
            " And A.病人ID=[1] And Nvl(A.婴儿,0)=[3] And A.挂号单=R.NO And R.ID=[2]" & _
            " And A.开始执行时间 is Not NULL And Nvl(A.医嘱状态,0)<>-1" & _
            " Order by A.序号"
    Else
        strSQL = "Select Distinct A.ID,A.序号,A.相关ID,A.医嘱期效,A.诊疗项目ID,A.医嘱内容," & _
            " A.单次用量,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.医生嘱托,A.执行性质,A.执行标记," & _
            " Nvl(C.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'-')) as 执行科室,A.执行科室id,A.执行时间方案," & _
            " A.执行科室ID,A.标本部位,A.检查方法,Nvl(B.类别,'*') as 类别,B.名称,B.计算单位," & _
            " A.总给予量 as 总量,D.计算单位 as 总量单位,D.id as 收费细目ID" & _
            " From 病人医嘱记录 A,诊疗项目目录 B,部门表 C,收费项目目录 D" & _
            " Where A.诊疗项目ID=B.ID(+) And A.执行科室ID=C.ID(+) And A.收费细目ID=D.ID(+)" & _
            " And A.病人ID=[1] And Nvl(A.婴儿,0)=[3] And A.主页ID=[2]" & _
            " And A.开始执行时间 is Not NULL And Nvl(A.医嘱状态,0)<>-1" & _
            " Order by A.序号"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, cboTime.ItemData(cboTime.ListIndex), cboBaby.ItemData(cboBaby.ListIndex))
    
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '清除表格内容
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                '.TextMatrix(i, col选择) = -1
                .TextMatrix(i, col序号) = rsTmp!ID
                .TextMatrix(i, col相关) = NVL(rsTmp!相关ID)
                .TextMatrix(i, col期效) = IIF(NVL(rsTmp!医嘱期效, 0) = 0, "长嘱", "临嘱")
                .TextMatrix(i, col内容) = rsTmp!医嘱内容
                .TextMatrix(i, col标本部位) = NVL(rsTmp!标本部位) '检验标本
                .TextMatrix(i, col检查方法) = NVL(rsTmp!检查方法)
                .TextMatrix(i, col单量) = FormatEx(NVL(rsTmp!单次用量), 4)
                If Not IsNull(rsTmp!单次用量) Then
                    If rsTmp!类别 = "4" Then
                        .TextMatrix(i, col单位) = NVL(rsTmp!总量单位)
                    Else
                        .TextMatrix(i, col单位) = NVL(rsTmp!计算单位)
                    End If
                End If
                If .TextMatrix(i, col期效) = "临嘱" Then
                    If Not IsNull(rsTmp!总量) Then
                        .TextMatrix(i, col总量) = FormatEx(NVL(rsTmp!总量), 4)
                        If Not IsNull(rsTmp!总量单位) Then
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!总量单位)
                        ElseIf InStr(",4,5,6,7,", rsTmp!类别) = 0 Then
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!计算单位)
                        End If
                    End If
                End If
                .TextMatrix(i, col频次) = NVL(rsTmp!执行频次)
                .TextMatrix(i, col频率次数) = NVL(rsTmp!频率次数)
                .TextMatrix(i, col频率间隔) = NVL(rsTmp!频率间隔)
                .TextMatrix(i, col间隔单位) = NVL(rsTmp!间隔单位)
                .TextMatrix(i, col嘱托) = NVL(rsTmp!医生嘱托)
                
                If InStr(NVL(rsTmp!执行时间方案), ",") > 0 Then
                    .TextMatrix(i, col执行时间) = Split(NVL(rsTmp!执行时间方案), ",")(1)
                Else
                    .TextMatrix(i, col执行时间) = NVL(rsTmp!执行时间方案)
                End If
                
                .TextMatrix(i, col执行科室) = NVL(rsTmp!执行科室)
                .Cell(flexcpData, i, col执行科室) = CLng(NVL(rsTmp!执行科室id, 0))
                .Cell(flexcpData, i, col执行性质) = Val(NVL(rsTmp!执行性质, 0))
                
                If Val(NVL(rsTmp!执行标记, 0)) = 1 Then
                    .TextMatrix(i, col执行性质) = "自取药"
                ElseIf Val(NVL(rsTmp!执行标记, 0)) = 2 Then
                    .TextMatrix(i, col执行性质) = "不取药"
                ElseIf Val(NVL(rsTmp!执行性质, 0)) = 5 And Val(NVL(rsTmp!执行标记, 0)) = 0 And Val(NVL(rsTmp!执行科室id, 0)) = 0 Then
                    .TextMatrix(i, col执行性质) = "自备药"
                Else
                    .TextMatrix(i, col执行性质) = "正常"
                End If
                .TextMatrix(i, col项目ID) = NVL(rsTmp!诊疗项目ID)
                .TextMatrix(i, col类别) = rsTmp!类别
                .TextMatrix(i, col收费细目ID) = zlCommFun.NVL(rsTmp!收费细目ID)
                
                '处理行隐藏及用法显示
                If InStr(",C,D,F,G,E,", rsTmp!类别) > 0 And Not IsNull(rsTmp!相关ID) Then
                    .RowHidden(i) = True
                    
                    '输血途径
                    If rsTmp!类别 = "E" And .TextMatrix(i - 1, col类别) = "K" And Val(.TextMatrix(i - 1, col序号)) = rsTmp!相关ID Then
                        .TextMatrix(i - 1, col用法) = NVL(rsTmp!名称)
                    End If
                ElseIf rsTmp!类别 = "7" Then
                    .RowHidden(i) = True
                ElseIf rsTmp!类别 = "E" And IsNull(rsTmp!相关ID) _
                    And Val(.TextMatrix(i - 1, col相关)) = rsTmp!ID _
                    And InStr(",5,6,", .TextMatrix(i - 1, col类别)) > 0 Then
                    '给药途径
                    .RowHidden(i) = True
                    '显示给药途径
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关)) = rsTmp!ID Then
                            .TextMatrix(j, col用法) = NVL(rsTmp!名称)
                            
                            '显示成药的执行性质
                            If Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                                .TextMatrix(j, col执行性质) = "离院带药"
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf rsTmp!类别 = "E" And IsNull(rsTmp!相关ID) _
                    And Val(.TextMatrix(i - 1, col相关)) = rsTmp!ID _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col类别)) > 0 Then
                    '中药用法或检验采集方法
                    .TextMatrix(i, col用法) = NVL(rsTmp!名称)
                    
                    '中药或检验的执行科室
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关)) = rsTmp!ID Then
                            If InStr(",7,C,", .TextMatrix(j, col类别)) > 0 Then
                                .TextMatrix(i, col执行科室) = .TextMatrix(j, col执行科室)
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    '中药付数
                    If .TextMatrix(i - 1, col类别) <> "C" Then
                        .TextMatrix(i, col总量单位) = "付"
                        
                        '显示中药配方执行性质:以药品为准判断
                        j = .FindRow(CStr(rsTmp!ID), , col相关)
                        If Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                            .TextMatrix(j, col执行性质) = "离院带药"
                        End If
                    End If
                End If
                rsTmp.MoveNext
            Next
            
            '以前方式的检查医嘱不选择
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) And .TextMatrix(i, col类别) = "D" Then
                    If CheckIsOldAdvice(i) Then
                        .TextMatrix(i, col选择) = 0
                        Call RowSelectSame(i)
                    End If
                End If
            Next
        End If
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col内容
        .Redraw = flexRDDirect
    End With
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col选择 Then Call RowSelectSame(Row)
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col内容 Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col选择 Then
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_DblClick()
    Call vsAdvice_KeyPress(32)
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsAdvice
        If KeyAscii = 32 Then
            If .Col <> col选择 Then
                KeyAscii = 0
                If Val(.TextMatrix(.Row, col项目ID)) <> 0 Then
                    .TextMatrix(.Row, col选择) = IIF(Val(.TextMatrix(.Row, col选择)) = 0, -1, 0)
                    Call RowSelectSame(.Row)
                End If
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col选择 Then
        Cancel = True
    Else
        '以前的检查医嘱不允许保存为成套方案
        If CheckIsOldAdvice(Row) Then
            MsgBox "该检查医嘱是系统升级以前下达的，与现有方式不兼容，不能保存为成套方案。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Function CheckIsOldAdvice(ByVal lngRow As Long) As Boolean
'功能：检查指定行的检查医嘱是否老方式
'参数：lngRow=检查医嘱可见行
    Dim lngIdx As Long

    With vsAdvice
        If .TextMatrix(lngRow, col类别) = "D" Then
            lngIdx = .FindRow(CStr(.TextMatrix(lngRow, col序号)), lngRow + 1, col相关)
            If lngIdx = -1 Then
                'CheckIsOldAdvice = True '以前的单部位检查
            ElseIf Val(.TextMatrix(lngIdx, col项目ID)) <> Val(.TextMatrix(lngRow, col项目ID)) Then
                CheckIsOldAdvice = True '以前的多部位项目检查
            End If
        End If
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '擦除一并给药相关行列的边线及内容
        lngLeft = col期效: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col频次: lngRight = col用法
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, col类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub RowSelectSame(ByVal lngRow As Long)
'功能：根据指定行(可能为任意行)的选择状态,将相关医嘱一并选择
    Dim i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, col相关)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) _
                    Or Val(.TextMatrix(i, col序号)) = Val(.TextMatrix(lngRow, col相关)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) _
                    Or Val(.TextMatrix(i, col序号)) = Val(.TextMatrix(lngRow, col相关)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col序号)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col序号)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub
