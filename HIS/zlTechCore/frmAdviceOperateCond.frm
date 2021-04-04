VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceOperateCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "校对条件"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmAdviceOperateCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraDetail 
      Height          =   5040
      Index           =   0
      Left            =   135
      TabIndex        =   14
      Top             =   60
      Width           =   5460
      Begin VB.CheckBox chkPauseLast 
         Caption         =   "默认从医嘱的上次执行时间之后开始暂停(&F)"
         Height          =   195
         Left            =   1215
         TabIndex        =   8
         Top             =   4260
         Width           =   3825
      End
      Begin MSComctlLib.Toolbar tbrAutoSel 
         Height          =   360
         Left            =   1215
         TabIndex        =   11
         Top             =   4575
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
               Caption         =   "按病区报警设置选择欠费病人   "
               Object.ToolTipText     =   "Ctrl + Q"
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.CheckBox chk期效 
         Caption         =   "临嘱(&T)"
         Height          =   195
         Index           =   1
         Left            =   2145
         TabIndex        =   1
         Top             =   330
         Width           =   930
      End
      Begin VB.CheckBox chk期效 
         Caption         =   "长嘱(&L)"
         Height          =   195
         Index           =   0
         Left            =   1215
         TabIndex        =   0
         Top             =   330
         Width           =   930
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "全选"
         Height          =   330
         Left            =   210
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   3450
         Width           =   870
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "全清"
         Height          =   330
         Left            =   210
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   3825
         Width           =   870
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   615
         Width           =   4095
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   3210
         Left            =   1215
         TabIndex        =   7
         Top             =   975
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5662
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
      Begin VB.CheckBox chk类别 
         Caption         =   "其他(&H)"
         Height          =   195
         Index           =   1
         Left            =   4425
         TabIndex        =   3
         Top             =   330
         Width           =   930
      End
      Begin VB.CheckBox chk类别 
         Caption         =   "药嘱(&D)"
         Height          =   195
         Index           =   0
         Left            =   3495
         TabIndex        =   2
         Top             =   330
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院病区(&U)"
         Height          =   180
         Left            =   150
         TabIndex        =   4
         Top             =   675
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院病人(&P)"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   1050
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   13
      Top             =   5235
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   12
      Top             =   5235
      Width           =   1100
   End
End
Attribute VB_Name = "frmAdviceOperateCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String 'IN
Public mint类型 As Integer 'IN:3-医嘱校对,5-暂停医嘱,6-启用医嘱
Public mlng病区ID As Long 'IN/OUT
Public mlng病人ID As Long 'IN
Public mstr病人IDs As String 'OUT:病人ID串(病人ID,主页ID;...)
Public mint期效 As Integer 'OUT:0-长嘱,1-临嘱,2-所有
Public mint类别 As Integer 'OUT:0-药嘱,1-其他,2-所有
Public mblnPauseLast As Boolean 'OUT:是否从上次执行时间开始暂停
Public mblnOK As Boolean 'OUT:是否确认

Private mrsWarn As ADODB.Recordset

Private Sub cboUnit_Click()
'功能：读取指定范围内的病人列表
    Dim rsTmp As New ADODB.Recordset
    Dim rsWarn As New ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str病人IDs As String, lng病区ID As Long
        
    lvwPati.ListItems.Clear
    
    On Error GoTo errH
    
    '读取病区报警设置
    If mint类型 = 5 Or mint类型 = 6 Then
        strSQL = "Select 适用病人,报警方法,报警值 From 记帐报警线 Where 病区ID=[1]"
        Set mrsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboUnit.ItemData(cboUnit.ListIndex))
    End If
    
    str病人IDs = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "医嘱操作病人" & mint类型, "")
    If str病人IDs <> "" And InStr(str病人IDs, ":") > 0 Then
        lng病区ID = Val(Split(str病人IDs, ":")(0))
        str病人IDs = Split(str病人IDs, ":")(1)
    End If
        
    '在院病人:出院病人禁止操作医嘱
    strSQL = _
        "Select A.病人ID,B.主页ID,A.姓名,A.住院号,B.出院病床 as 床号," & _
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
        objItem.Tag = rsTmp!主页ID
                
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

Private Sub chk类别_Click(Index As Integer)
    If chk类别(0).Value = 0 And chk类别(1).Value = 0 Then chk类别(Index).Value = 1
End Sub

Private Sub chk期效_Click(Index As Integer)
    If chk期效(0).Value = 0 And chk期效(1).Value = 0 Then chk期效(Index).Value = 1
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
    Dim strTmp As String, i As Long
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "请选择一个病区。", vbInformation, gstrSysName
        cboUnit.SetFocus: Exit Sub
    End If
    mlng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
    
    '住院病人
    mstr病人IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            strTmp = strTmp & "," & Mid(lvwPati.ListItems(i).Key, 2) '用于保存
            mstr病人IDs = mstr病人IDs & ";" & Mid(lvwPati.ListItems(i).Key, 2) & "," & lvwPati.ListItems(i).Tag
        End If
    Next
    strTmp = Mid(strTmp, 2)
    mstr病人IDs = Mid(mstr病人IDs, 2)
    If mstr病人IDs = "" Then
        MsgBox "请至少选择一个病人。", vbInformation, gstrSysName
        lvwPati.SetFocus: Exit Sub
    End If
        
    '医嘱期效
    mint期效 = IIF(chk期效(0).Value = 1 And chk期效(1).Value = 1, 0, IIF(chk期效(0).Value = 1, 1, 2))
        
    '医嘱类别
    mint类别 = IIF(chk类别(0).Value = 1 And chk类别(1).Value = 1, 0, IIF(chk类别(0).Value = 1, 1, 2))
    
    '默认暂停时间
    mblnPauseLast = chkPauseLast.Value = 1
    
    '保存条件设置
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "医嘱操作期效" & mint类型, IIF(chk期效(0).Value = 1 And chk期效(1).Value = 1, 0, IIF(chk期效(0).Value = 1, 1, 2))
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "医嘱操作类别" & mint类型, IIF(chk类别(0).Value = 1 And chk类别(1).Value = 1, 0, IIF(chk类别(0).Value = 1, 1, 2))
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "上次开始暂停", chkPauseLast.Value
    If UBound(Split(strTmp, ",")) = 0 And Val(strTmp) = mlng病人ID Then
        '病人：选择病人仅为当前病人时,不保存
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "医嘱操作病人" & mint类型, ""
    Else
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "医嘱操作病人" & mint类型, cboUnit.ItemData(cboUnit.ListIndex) & ":" & strTmp
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAllPati_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdNoPati_Click
    ElseIf KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        If tbrAutoSel.Visible Then
            Call tbrAutoSel_ButtonClick(tbrAutoSel.Buttons(1))
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim lngTmp As Long
    
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    
    mblnOK = False
    Me.Caption = Decode(mint类型, 3, "校对", 5, "暂停", 6, "启用") & "条件"
    If mint类型 <> 5 Then
        chkPauseLast.Visible = False
        
        If mint类型 = 6 Then
            tbrAutoSel.Buttons(1).Caption = "按病区报警设置排开欠费病人   "
            lvwPati.Height = chkPauseLast.Top + chkPauseLast.Height - lvwPati.Top
            cmdAllPati.Top = cmdAllPati.Top + chkPauseLast.Height
            cmdNoPati.Top = cmdNoPati.Top + chkPauseLast.Height
        Else
            tbrAutoSel.Visible = False
            lvwPati.Height = tbrAutoSel.Top + tbrAutoSel.Height - lvwPati.Top
            cmdAllPati.Top = cmdAllPati.Top + tbrAutoSel.Height + chkPauseLast.Height
            cmdNoPati.Top = cmdNoPati.Top + tbrAutoSel.Height + chkPauseLast.Height
        End If
    End If
    
    '缺省医嘱期效
    If mint类型 = 5 Or mint类型 = 6 Then
        chk期效(0).Enabled = False: chk期效(1).Enabled = False
        chk期效(0).Value = 1: chk期效(1).Value = 0
    Else
        chk期效(0).Enabled = True: chk期效(1).Enabled = True
        lngTmp = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "医嘱操作期效" & mint类型, 0))
        If lngTmp = 0 Then
            chk期效(0).Value = 1: chk期效(1).Value = 1
        Else
            chk期效(lngTmp - 1).Value = 1
        End If
    End If
    '缺省医嘱类别
    lngTmp = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "医嘱操作类别" & mint类型, 0))
    If lngTmp = 0 Then
        chk类别(0).Value = 1: chk类别(1).Value = 1
    Else
        chk类别(lngTmp - 1).Value = 1
    End If
    
    '默认暂停时间
    chkPauseLast.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "上次开始暂停", 0))
    
    '病区/病人
    Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
End Sub

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
    '释放私有及IN变量
    mstrPrivs = ""
    mint类型 = 0
    mlng病人ID = 0
    Set mrsWarn = Nothing
    
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
End Sub

Private Sub tbrAutoSel_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long, k As Long
    
    If mrsWarn Is Nothing Then Exit Sub
    
    With lvwPati
        If mint类型 = 5 Then
            k = 0
            For i = 1 To .ListItems.Count
                .ListItems(i).Checked = False
                '只根据累计报警方法进行处理
                mrsWarn.Filter = "报警方法=1 And 适用病人=" & Val(.ListItems(i).ListSubItems(1).Tag) + 1
                If Not mrsWarn.EOF Then
                    If Val(.ListItems(i).SubItems(3)) + Val(.ListItems(i).ListSubItems(2).Tag) < Nvl(mrsWarn!报警值, 0) Then
                        .ListItems(i).Checked = True
                        If k = 0 Then
                            .ListItems(i).Selected = True
                            .ListItems(i).EnsureVisible
                        End If
                        k = k + 1
                    End If
                End If
            Next
        ElseIf mint类型 = 6 Then
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
        End If
    End With
End Sub
