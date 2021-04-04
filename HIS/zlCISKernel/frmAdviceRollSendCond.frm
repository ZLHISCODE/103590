VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceRollSendCond 
   AutoRedraw      =   -1  'True
   Caption         =   "收回条件"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   6180
   Icon            =   "frmAdviceRollSendCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6705
   ScaleWidth      =   6180
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6180
      TabIndex        =   9
      Top             =   6210
      Width           =   6180
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4950
         TabIndex        =   11
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3720
         TabIndex        =   10
         Top             =   0
         Width           =   1100
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   5535
      Left            =   135
      TabIndex        =   5
      Top             =   555
      Width           =   5940
      Begin VB.CheckBox chkOut 
         Alignment       =   1  'Right Justify
         Caption         =   "显示最近出院的病人(&A)"
         Height          =   195
         Left            =   3600
         TabIndex        =   4
         Top             =   5235
         Width           =   2190
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "全选"
         Height          =   330
         Left            =   270
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   4410
         Width           =   870
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "全清"
         Height          =   330
         Left            =   270
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   4785
         Width           =   870
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   2655
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   4500
         Left            =   1215
         TabIndex        =   1
         Top             =   645
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   7938
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
            Object.Width           =   2117
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
            SubItemIndex    =   3
            Text            =   "住院医师"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "费别"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "护理等级"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "科室"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "入院日期"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "病人类型"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院病区(&U)"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   345
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院病人(&P)"
         Height          =   180
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   990
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "系统将收回病人超期发送所产生的多余费用及药品。请从病人清单中选择需要处理的病人。"
      Height          =   380
      Left            =   1155
      TabIndex        =   8
      Top             =   135
      Width           =   4140
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   390
      Picture         =   "frmAdviceRollSendCond.frx":058A
      Top             =   75
      Width           =   480
   End
End
Attribute VB_Name = "frmAdviceRollSendCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mMainPrivs As String 'IN
Public mlng病区ID As Long 'IN/OUT
Public mlng病人ID As Long 'IN

Public mblnOK As Boolean 'OUT:是否确认
Public mstr病人IDs As String 'OUT:病人ID串
Public mstr主页IDs As String 'OUT:病人ID对应主页ID串

Private Sub cboUnit_Click()
'功能：读取指定范围内的病人列表
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str病人IDs As String, lng病区ID As Long, lngUnitID As Long
    Dim lngColor As Long
        
    On Error GoTo errH
    lvwPati.ListItems.Clear
    lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    
    str病人IDs = zlDatabase.GetPara("发送病人", glngSys, p住院医嘱发送)
    If str病人IDs <> "" And InStr(str病人IDs, ":") > 0 Then
        lng病区ID = Val(Split(str病人IDs, ":")(0))
        str病人IDs = Split(str病人IDs, ":")(1)
    End If
            
    Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng病人ID, False, False, chkOut.value)
  
    For i = 1 To rsTmp.RecordCount
        If Val(rsTmp!审核标志 & "") < 1 Or gbyt病人审核方式 <> 1 Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID & "_" & rsTmp!主页ID, rsTmp!姓名)
            objItem.SubItems(1) = IIF(IsNull(rsTmp!住院号), "", rsTmp!住院号)
            objItem.SubItems(2) = IIF(IsNull(rsTmp!床号), "", rsTmp!床号)
            objItem.SubItems(3) = IIF(IsNull(rsTmp!住院医师), "", rsTmp!住院医师)
            objItem.SubItems(4) = IIF(IsNull(rsTmp!费别), "", rsTmp!费别)
            objItem.SubItems(5) = IIF(IsNull(rsTmp!护理等级), "", rsTmp!护理等级)
            objItem.SubItems(6) = IIF(IsNull(rsTmp!科室), "", rsTmp!科室)
            objItem.SubItems(7) = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
            objItem.SubItems(8) = Nvl(rsTmp!病人类型)
            
            '病人颜色
            lngColor = zlDatabase.GetPatiColor(Nvl(rsTmp!病人类型))
            objItem.ListSubItems(1).ForeColor = lngColor
            objItem.ListSubItems(8).ForeColor = lngColor
            
            '上次是否选择
            If lngUnitID = lng病区ID And str病人IDs <> "" Then
                If str病人IDs = "ALL" _
                    Or Left(str病人IDs, 1) <> "-" And InStr("," & str病人IDs & ",", "," & rsTmp!病人ID & ",") > 0 _
                    Or Left(str病人IDs, 1) = "-" And InStr("," & Mid(str病人IDs, 2) & ",", "," & rsTmp!病人ID & ",") = 0 Then
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
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkOut_Click()
    If Visible Then Call cboUnit_Click
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
    Dim strSel As String, strUnSel As String
    Dim i As Long
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "请选择一个病区。", vbInformation, gstrSysName
        cboUnit.SetFocus: Exit Sub
    End If
    mlng病区ID = cboUnit.ItemData(cboUnit.ListIndex)
    
    '住院病人
    mstr病人IDs = "": mstr主页IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            mstr病人IDs = mstr病人IDs & "," & Split(Mid(lvwPati.ListItems(i).Key, 2), "_")(0)
            mstr主页IDs = mstr主页IDs & "," & Split(Mid(lvwPati.ListItems(i).Key, 2), "_")(1)
            strSel = strSel & "," & Split(Mid(lvwPati.ListItems(i).Key, 2), "_")(0)
        Else
            strUnSel = strUnSel & "," & Split(Mid(lvwPati.ListItems(i).Key, 2), "_")(0)
        End If
    Next
    mstr病人IDs = Mid(mstr病人IDs, 2)
    mstr主页IDs = Mid(mstr主页IDs, 2)
    If mstr病人IDs = "" Then
        MsgBox "请至少选择一个需要发送医嘱病人。", vbInformation, gstrSysName
        lvwPati.SetFocus: Exit Sub
    End If
        
    '保存条件设置
    strSel = Mid(strSel, 2)
    strUnSel = Mid(strUnSel, 2)
    If strSel = "" Or (UBound(Split(strSel, ",")) = 0 And Val(strSel) = mlng病人ID) Then
        strSel = ""
    Else
        If strUnSel = "" Then
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":ALL"
        ElseIf UBound(Split(strSel, ",")) > UBound(Split(strUnSel, ",")) Then
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":-" & strUnSel
        Else
            strSel = cboUnit.ItemData(cboUnit.ListIndex) & ":" & strSel
        End If
    End If
    Call zlDatabase.SetPara("发送病人", strSel, glngSys, p住院医嘱发送)
        
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAllPati_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdNoPati_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
    Call RestoreWinState(Me, App.ProductName)
    '病区/病人
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    'Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '包含门诊观察室
    If InStr(mMainPrivs, "全院病人") > 0 Then
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
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
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

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    fraDetail.Width = Me.ScaleWidth - 240
    chkOut.Left = fraDetail.Width - chkOut.Width - 120
    lvwPati.Width = fraDetail.Width - lvwPati.Left - 120
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 60
        
    fraDetail.Height = Me.ScaleHeight - picBottom.Height - fraDetail.Top - 120
    
    chkOut.Top = fraDetail.Height - chkOut.Height - 60
    lvwPati.Height = chkOut.Top - lvwPati.Top - 60
    
    cmdNoPati.Top = lvwPati.Top + lvwPati.Height - 30 - cmdNoPati.Height
    cmdAllPati.Top = cmdNoPati.Top - cmdAllPati.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '释放私有及IN变量
    mMainPrivs = ""
    mlng病人ID = 0
    
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub
