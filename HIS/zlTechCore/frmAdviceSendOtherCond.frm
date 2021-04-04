VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceSendOtherCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发送条件"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmAdviceSendOtherCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraDetail 
      Height          =   5445
      Index           =   0
      Left            =   135
      TabIndex        =   19
      Top             =   60
      Width           =   5460
      Begin VB.CommandButton cmd执行科室 
         Height          =   240
         Left            =   5010
         Picture         =   "frmAdviceSendOtherCond.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "选择执行科室(F4)"
         Top             =   930
         Width           =   270
      End
      Begin VB.CheckBox chk加班加价 
         Caption         =   "执行加班加价(&V)"
         Height          =   195
         Left            =   3525
         TabIndex        =   4
         Top             =   600
         Width           =   1650
      End
      Begin VB.ListBox lstClass 
         Columns         =   4
         Height          =   1110
         ItemData        =   "frmAdviceSendOtherCond.frx":0680
         Left            =   1215
         List            =   "frmAdviceSendOtherCond.frx":0682
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   4230
         Width           =   4095
      End
      Begin VB.OptionButton opt期效 
         Caption         =   "长嘱(&L)"
         Height          =   180
         Index           =   0
         Left            =   1215
         TabIndex        =   0
         Top             =   255
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton opt期效 
         Caption         =   "临嘱(&T)"
         Height          =   180
         Index           =   1
         Left            =   2190
         TabIndex        =   1
         Top             =   255
         Width           =   930
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "全选"
         Height          =   330
         Left            =   270
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   2955
         Width           =   870
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "全清"
         Height          =   330
         Left            =   270
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   3330
         Width           =   870
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1260
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1215
         TabIndex        =   3
         Top             =   540
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   24838147
         CurrentDate     =   37953
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   2070
         Left            =   1215
         TabIndex        =   11
         Top             =   1620
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3651
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "姓名"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "住院号"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "床号"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "剩余款"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "住院医师"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "费别"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "护理等级"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "科室"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "入院日期"
            Object.Width           =   2857
         EndProperty
      End
      Begin VB.TextBox txt执行科室 
         Height          =   300
         Left            =   1215
         TabIndex        =   6
         Top             =   900
         Width           =   4095
      End
      Begin MSComctlLib.Toolbar tbrAutoSel 
         Height          =   360
         Left            =   1215
         TabIndex        =   20
         Top             =   3750
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   635
         ButtonWidth     =   5318
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "按病区报警设置排开欠费病人   "
               Object.ToolTipText     =   "Ctrl + Q"
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label lbl执行科室 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室(&D)"
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   960
         Width           =   990
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   105
         X2              =   5360
         Y1              =   4170
         Y2              =   4170
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   5360
         Y1              =   4155
         Y2              =   4155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊疗类别(&T)"
         Height          =   180
         Left            =   180
         TabIndex        =   14
         Top             =   4275
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院病区(&U)"
         Height          =   180
         Left            =   180
         TabIndex        =   8
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间(&E)"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院病人(&P)"
         Height          =   180
         Left            =   180
         TabIndex        =   10
         Top             =   1695
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   495
      TabIndex        =   18
      Top             =   5610
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   17
      Top             =   5610
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   16
      Top             =   5610
      Width           =   1100
   End
End
Attribute VB_Name = "frmAdviceSendOtherCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String 'IN
Public mlng病区ID As Long 'IN/OUT
Public mlng病人ID As Long 'IN
Public mblnOK As Boolean 'OUT:是否确认
Public mstrEnd As String 'OUT:结束时间
Public mint期效 As Integer 'OUT:0-长嘱,1-临嘱
Public mlng执行科室ID As Long 'OUT-发送的执行科室
Public mstr病人IDs As String 'OUT:病人ID串
Public mstr类别s As String 'OUT:诊疗类别串

Private mrsWarn As ADODB.Recordset
Private mstrLike As String

Private Sub cboUnit_Click()
'功能：读取指定范围内的病人列表
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str病人IDs As String, lng病区ID As Long
        
    lvwPati.ListItems.Clear
    
    On Error GoTo errH
    
    strSQL = "Select 适用病人,报警方法,报警值 From 记帐报警线 Where 病区ID=[1]"
    Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
    
    str病人IDs = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药发送病人", "")
    If str病人IDs <> "" And InStr(str病人IDs, ":") > 0 Then
        lng病区ID = Val(Split(str病人IDs, ":")(0))
        str病人IDs = Split(str病人IDs, ":")(1)
    End If
        
    '在院病人:出院病人禁止下医嘱,发送医嘱
    strSQL = _
        "Select A.病人ID,A.姓名,A.住院号,B.出院病床 as 床号," & _
        " Nvl(E.预交余额,0)-Nvl(E.费用余额,0)+Decode(B.险类,Null,0,Nvl(F.金额,0)) as 剩余款," & _
        " A.担保额,Decode(X.编码,'1',1,Decode(B.险类,Null,0,1)) as 医保,B.险类," & _
        " B.住院医师,B.费别,D.名称 as 护理等级,C.名称 as 科室,B.入院日期" & _
        " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 D,病人余额 E,医疗付款方式 X," & _
        " (Select 病人ID,主页ID,Sum(金额) As 金额 From 保险模拟结算 Group By 病人ID,主页ID) F" & _
        " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.出院科室ID=C.ID" & _
        " And A.病人ID=E.病人ID(+) And E.性质(+)=1 And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+)" & _
        " And B.出院日期 is NULL And Nvl(B.状态,0)<>3 And B.护理等级ID=D.ID(+) And B.医疗付款方式=X.名称(+)" & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) > 0, " And B.当前病区ID=[1]", "") & _
        IIF(cboUnit.ItemData(cboUnit.ListIndex) = 0, " Order by A.住院号 Desc", " Order by LPAD(B.出院病床,10,' ')")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID, rsTmp!姓名)
        objItem.SubItems(1) = IIF(IsNull(rsTmp!住院号), "", rsTmp!住院号)
        objItem.SubItems(2) = IIF(IsNull(rsTmp!床号), "", rsTmp!床号)
        objItem.SubItems(3) = Format(Nvl(rsTmp!剩余款, 0), "0.00")
        objItem.SubItems(4) = IIF(IsNull(rsTmp!住院医师), "", rsTmp!住院医师)
        objItem.SubItems(5) = IIF(IsNull(rsTmp!费别), "", rsTmp!费别)
        objItem.SubItems(6) = IIF(IsNull(rsTmp!护理等级), "", rsTmp!护理等级)
        objItem.SubItems(7) = IIF(IsNull(rsTmp!科室), "", rsTmp!科室)
        objItem.SubItems(8) = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
        
        '附加信息
        objItem.ListSubItems(1).Tag = Nvl(rsTmp!医保, 0)
        objItem.ListSubItems(2).Tag = Nvl(rsTmp!担保额, 0)
        
        '保险病人用红色显示
        If Not IsNull(rsTmp!险类) Then
            objItem.ForeColor = vbRed
            For j = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(j).ForeColor = vbRed
            Next
        End If
        
        '上次是否选择
        If cboUnit.ItemData(cboUnit.ListIndex) = lng病区ID And str病人IDs <> "" Then
            If InStr("," & str病人IDs & ",", "," & rsTmp!病人ID & ",") > 0 Then
                objItem.Checked = True
                If k = 0 Then '为了看到有选择的
                    objItem.EnsureVisible
                    objItem.Selected = True
                    k = 1
                End If
            End If
        ElseIf rsTmp!病人ID = mlng病人ID Then
            objItem.Checked = True '缺省只选择当前病人
            objItem.EnsureVisible
            objItem.Selected = True
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(lvwPati, True)
    lvwPati.SetFocus
End Sub

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "请选择一个病区。", vbInformation, gstrSysName
        cboUnit.SetFocus: Exit Sub
    End If
    mlng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
    
    '时间和期效
    mint期效 = IIF(opt期效(1).Value, 1, 0)
    If opt期效(0).Value Then
        mstrEnd = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")
    Else
        mstrEnd = ""
    End If
    
    '执行科室
    mlng执行科室ID = Val(cmd执行科室.Tag)
    
    '住院病人
    mstr病人IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            mstr病人IDs = mstr病人IDs & "," & Mid(lvwPati.ListItems(i).Key, 2)
        End If
    Next
    mstr病人IDs = Mid(mstr病人IDs, 2)
    If mstr病人IDs = "" Then
        MsgBox "请至少选择一个需要发送医嘱病人。", vbInformation, gstrSysName
        lvwPati.SetFocus: Exit Sub
    End If
        
    '诊疗类别
    mstr类别s = ""
    For i = 0 To lstClass.ListCount - 1
        If lstClass.Selected(i) Then
            mstr类别s = mstr类别s & ",'" & Chr(lstClass.ItemData(i)) & "'"
        End If
    Next
    mstr类别s = Mid(mstr类别s, 2)
    If mstr类别s = "" Then
        MsgBox "请至少选择一种诊疗类别。", vbInformation, gstrSysName
        lstClass.SetFocus: Exit Sub
    End If
    If UBound(Split(mstr类别s, ",")) + 1 = lstClass.ListCount Then
        mstr类别s = ""
    End If
    
    gbln加班加价 = chk加班加价.Value = 1
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd执行科室_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    strSQL = _
        " Select 0 as ID,'-' as 编码,'所有科室' as 名称,NULL as 简码 From Dual" & _
        " Union ALL" & _
        " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " Order by 编码"
    vRect = GetControlRect(txt执行科室.Hwnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "执行科室", , , , , , True, vRect.Left, vRect.Top, txt执行科室.Height, blnCancel, , True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有可用的科室，请先到部门管理中设置。", vbInformation, gstrSysName
        End If
        txt执行科室.Text = txt执行科室.Tag
        Call zlControl.TxtSelAll(txt执行科室)
    Else
        txt执行科室.Text = rsTmp!名称
        txt执行科室.Tag = rsTmp!名称
        cmd执行科室.Tag = rsTmp!ID
    End If
    txt执行科室.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If ActiveControl Is lstClass Then
            j = lstClass.ListIndex
            For i = 0 To lstClass.ListCount - 1
                lstClass.Selected(i) = True
            Next
            lstClass.ListIndex = j
        Else
            Call cmdAllPati_Click
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If ActiveControl Is lstClass Then
            j = lstClass.ListIndex
            For i = 0 To lstClass.ListCount - 1
                lstClass.Selected(i) = False
            Next
            lstClass.ListIndex = j
        Else
            Call cmdNoPati_Click
        End If
    ElseIf KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        If tbrAutoSel.Visible Then
            Call tbrAutoSel_ButtonClick(tbrAutoSel.Buttons(1))
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Me.ActiveControl Is txt执行科室 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim strTmp As String, lngTmp As Long
    
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    
    mblnOK = False
    
    '输入匹配
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    
    '缺省医嘱期效
    lngTmp = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药医嘱期效", 0))
    opt期效(lngTmp).Value = True
    '至少有一个才可能进来
    If InStr(mstrPrivs, "发送其他临嘱") = 0 Then
        opt期效(0).Value = True
        opt期效(1).Enabled = False
    ElseIf InStr(mstrPrivs, "发送其他长嘱") = 0 Then
        opt期效(1).Value = True
        opt期效(0).Enabled = False
    End If
    
    '缺省结束时间
    curDate = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药结束时点", "23:59:59")
    lngTmp = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药时间间隔", 0))
    dtpEnd.Value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
        
    '缺省执行科室
    txt执行科室.Text = "所有科室"
    txt执行科室.Tag = txt执行科室.Text
    cmd执行科室.Tag = ""
        
    '病区/病人
    Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
                    
    '诊疗类别
    Call Load诊疗类别
End Sub

Private Function Load诊疗类别() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str类别s As String
    
    On Error GoTo errH
    
    str类别s = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药诊疗类别", "")
    
    strSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('5','6','7','8','9') Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lstClass.AddItem rsTmp!名称
        lstClass.ItemData(lstClass.NewIndex) = Asc(rsTmp!编码)
        If str类别s <> "" Then
            If InStr(str类别s, "'" & rsTmp!编码 & "'") > 0 Then
                lstClass.Selected(lstClass.NewIndex) = True
            End If
        Else
            lstClass.Selected(lstClass.NewIndex) = True
        End If
        rsTmp.MoveNext
    Next
    If lstClass.ListCount > 0 Then lstClass.ListIndex = 0
    Load诊疗类别 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '包含门诊观察室
    If InStr(mstrPrivs, "全院病人") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        If Not gbln病区科室独立 Then
            strSQL = strSQL & IIF(strSQL <> "", " Union ", "") & _
                " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
                " From 床位状况记录 A,部门人员 B,部门表 C" & _
                " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
                " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        End If
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If
    
    cboUnit.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng病区ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTmp As String
    
    '保存条件设置
    If mblnOK Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药结束时点", Format(dtpEnd.Value, "HH:mm:ss")
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药时间间隔", Int(CDate(Format(dtpEnd.Value, "yyyy-MM-dd")) - CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")))
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药医嘱期效", IIF(opt期效(1).Value, 1, 0)
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药诊疗类别", mstr类别s
        
        '病人：选择病人仅为当前病人时,不保存
        If UBound(Split(mstr病人IDs, ",")) = 0 And Val(mstr病人IDs) = mlng病人ID Then
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药发送病人", ""
        Else
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "非药发送病人", cboUnit.ItemData(cboUnit.ListIndex) & ":" & mstr病人IDs
        End If
    End If
    
    '释放私有及IN变量
    mstrPrivs = ""
    mlng病人ID = 0
    Set mrsWarn = Nothing
    
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
End Sub

Private Sub opt期效_Click(Index As Integer)
    dtpEnd.Enabled = opt期效(0).Value
End Sub

Private Sub txt执行科室_GotFocus()
    Call zlControl.TxtSelAll(txt执行科室)
End Sub

Private Sub txt执行科室_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call cmd执行科室_Click
End Sub

Private Sub txt执行科室_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt执行科室.Text = txt执行科室.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt执行科室.Text = "" Then
            txt执行科室.Text = "所有科室"
            txt执行科室.Tag = txt执行科室.Text
            cmd执行科室.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            strSQL = _
                " Select 0 as ID,'-' as 编码,'所有科室' as 名称,NULL as 简码 From Dual" & _
                " Union ALL" & _
                " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
            strSQL = "Select * From (" & strSQL & ")" & _
                " Where 编码 Like [1] Or Upper(名称) Like [2] Or Upper(简码) Like [2]" & _
                " Order by 编码"
            vRect = GetControlRect(txt执行科室.Hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "执行科室", False, "", "", False, False, True, _
                vRect.Left, vRect.Top, txt执行科室.Height, blnCancel, False, True, _
                UCase(txt执行科室.Text) & "%", mstrLike & UCase(txt执行科室.Text) & "%")
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配的科室。", vbInformation, gstrSysName
                End If
                txt执行科室.Text = txt执行科室.Tag
                Call zlControl.TxtSelAll(txt执行科室)
                txt执行科室.SetFocus
            Else
                txt执行科室.Text = rsTmp!名称
                txt执行科室.Tag = rsTmp!名称
                cmd执行科室.Tag = rsTmp!ID
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End If
End Sub

Private Sub txt执行科室_Validate(Cancel As Boolean)
    If txt执行科室.Text = "" Then
        txt执行科室.Text = "所有科室"
        txt执行科室.Tag = txt执行科室.Text
        cmd执行科室.Tag = ""
    ElseIf txt执行科室.Text <> txt执行科室.Tag Then
        txt执行科室.Text = txt执行科室.Tag '恢复人为的清除
    End If
End Sub

Private Sub tbrAutoSel_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
    
    If mrsWarn Is Nothing Then Exit Sub
    
    With lvwPati
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                '只根据累计报警方法进行处理
                mrsWarn.Filter = "报警方法=1 And 适用病人=" & Val(.ListItems(i).ListSubItems(1).Tag) + 1
                If Not mrsWarn.EOF Then
                    If Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) < Nvl(mrsWarn!报警值, 0) Then
                        .ListItems(i).Checked = False
                    End If
                End If
            End If
        Next
    End With
End Sub
