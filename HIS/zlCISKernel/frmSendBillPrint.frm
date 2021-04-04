VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendBillPrint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊疗单据打印"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmSendBillPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPreview 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   3540
      TabIndex        =   4
      ToolTipText     =   "预览当前单据"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "设置(&S)"
      Height          =   350
      Left            =   2445
      TabIndex        =   3
      ToolTipText     =   "设置当前单据"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "打印所有选择的单据"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "返回(&X)"
      Height          =   350
      Left            =   5895
      TabIndex        =   2
      Top             =   4665
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwBill 
      Height          =   3795
      Left            =   75
      TabIndex        =   0
      Top             =   750
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   6694
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "单据号"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "诊疗单据"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "说明"
         Object.Width           =   6350
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "frmSendBillPrint.frx":058A
      Top             =   165
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSendBillPrint.frx":0E54
      Height          =   525
      Left            =   930
      TabIndex        =   5
      Top             =   120
      Width           =   6090
   End
End
Attribute VB_Name = "frmSendBillPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mstrBillPrint As String '当前打印的诊疗单据：报表编号、NO、记录性质

Private mlng发送号 As Long
Private mint场合 As Integer
Private mstr前提IDs As String
Private mint打印方式 As Integer
Private mblnItem As Boolean
Private mint申请单打印模式 As Integer  '1-发送时打印，2-新开时打印

Public Sub ShowMe(ByVal lng发送号 As Long, ByVal int场合 As Integer, frmParent As Object, Optional ByVal str前提IDs As String)
'参数：lng发送号=本次发送的发送号
'      int场合=1-门诊,2-住院(数据场合,不是调用场合)
'      str前提IDs医技站中在当前科室执行的所有医嘱
    mlng发送号 = lng发送号
    mint场合 = int场合
    mstr前提IDs = str前提IDs
    
    On Error Resume Next
    Me.Show 1, frmParent
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
'功能：对诊疗单据对应的自定义报表进行预览
    If lvwBill.SelectedItem Is Nothing Then Exit Sub
    With lvwBill.SelectedItem
        '输血医嘱打印申请单调用相关函数进行检查
        If InStr(1, ",ZL1_INSIDE_1254_17_1,ZL1_INSIDE_1254_17_2,", "," & .Tag & ",") <> 0 Then
            If BloodApplyPrintCheck(Val(.ListSubItems(2).Tag), mint场合, IIF(.Tag = "ZL1_INSIDE_1254_17_1", 1, 2), 0) = False Then Exit Sub
        End If
        mstrBillPrint = .Tag & "," & .Text & "," & .ListSubItems(1).Tag
        Call mobjReport.ReportOpen(gcnOracle, glngSys, .Tag, Me, "NO=" & .Text, "性质=" & Val(.ListSubItems(1).Tag), "医嘱ID=" & Val(.ListSubItems(2).Tag), 1)
        mstrBillPrint = ""
    End With
End Sub

Private Sub cmdPrint_Click()
'功能：对选择的诊疗单据进行打印
    Dim i As Long, j As Long
    Dim blnALL As Boolean
    
    If lvwBill.SelectedItem Is Nothing Then Exit Sub
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then j = j + 1
    Next
    If j = 0 Then
        MsgBox "请先选择需要打印的诊疗单据。", vbInformation, gstrSysName
        Exit Sub
    ElseIf j = lvwBill.ListItems.Count Then
        blnALL = True
    End If
    
    '输血医嘱打印申请单调用相关函数进行检查
    For i = 1 To lvwBill.ListItems.Count
        With lvwBill.ListItems(i)
            If .Checked Then
                If InStr(1, ",ZL1_INSIDE_1254_17_1,ZL1_INSIDE_1254_17_2,", "," & .Tag & ",") <> 0 Then
                    If BloodApplyPrintCheck(Val(.ListSubItems(2).Tag), mint场合, IIF(.Tag = "ZL1_INSIDE_1254_17_1", 1, 2), 1) = False Then
                        .Checked = False
                        If blnALL = True Then blnALL = False
                    End If
                End If
            End If
        End With
    Next
    
    cmdPrint.Enabled = False
    Screen.MousePointer = 11
    For i = 1 To lvwBill.ListItems.Count
        With lvwBill.ListItems(i)
            If .Checked Then
                .Selected = True: .EnsureVisible: Me.Refresh
                
                mstrBillPrint = .Tag & "," & .Text & "," & .ListSubItems(1).Tag
                Call mobjReport.ReportOpen(gcnOracle, glngSys, .Tag, Me, "NO=" & .Text, "性质=" & Val(.ListSubItems(1).Tag), "医嘱ID=" & Val(.ListSubItems(2).Tag), "PrintEmpty=0", 2)
                mstrBillPrint = ""
                
                '已打印的用颜色标识
                .Checked = False: .ForeColor = vbBlue
                For j = 1 To .ListSubItems.Count
                    .ListSubItems(j).ForeColor = vbBlue
                Next
            End If
        End With
    Next
    Screen.MousePointer = 0
    cmdPrint.Enabled = True
    
    '手工打印时，全部打印完毕后自动退出
    If mint打印方式 = 1 And blnALL Then
        Unload Me: Exit Sub
    ElseIf Visible Then
        cmdCancel.SetFocus
    End If
End Sub

Private Sub cmdSetup_Click()
'功能：对诊疗单据对应的自定义报表进行设置
    If lvwBill.SelectedItem Is Nothing Then Exit Sub
    Call mobjReport.ReportPrintSet(gcnOracle, glngSys, lvwBill.SelectedItem.Tag, Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    mblnItem = False
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To lvwBill.ListItems.Count
            lvwBill.ListItems(i).Checked = True
        Next
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        For i = 1 To lvwBill.ListItems.Count
            lvwBill.ListItems(i).Checked = False
        Next
    End If
End Sub

Private Sub Form_Load()
    '诊疗单据打印方式:0-不打印,1-手工打印,2-自动打印
    If mstr前提IDs = "" Then
        If mint场合 = 1 Then
            mint打印方式 = Val(zlDatabase.GetPara("门诊发送单据打印", glngSys, p门诊医嘱下达))
        Else
            mint打印方式 = Val(zlDatabase.GetPara("住院发送单据打印", glngSys, p住院医嘱发送))
        End If
    Else
        mint打印方式 = 1
    End If
    If mint打印方式 = 0 Then Unload Me: Exit Sub
    mint申请单打印模式 = Val(zlDatabase.GetPara("输血申请单打印模式", glngSys, p住院医嘱发送, "1"))
    
    Call RestoreListViewState(lvwBill, App.ProductName & "\" & Me.Name, "")
    If Not LoadBill Then Unload Me: Exit Sub
    If lvwBill.ListItems.Count = 0 Then Unload Me: Exit Sub
    mblnItem = False
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    
    '自动打印后退出
    If mint打印方式 = 2 Then
        Call cmdPrint_Click
        Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set mobjReport = Nothing   '自动缓存以便报表部件中的缓存能重复使用
    Call SaveListViewState(lvwBill, App.ProductName & "\" & Me.Name, "")
End Sub

Private Sub lvwBill_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwBill, ColumnHeader.Index)
End Sub

Private Function LoadBill() As Boolean
'功能：读取本次发送可以打印的诊疗单据列表
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim objItem As ListItem
    Dim strTmp As String
    
    lvwBill.ListItems.Clear
    
    On Error GoTo errH
    
    '如果是住院，则排除申请单下达的输血医嘱，通过特殊方式处理
    If mint场合 = 2 And mint申请单打印模式 = 1 Then
        If gbln血库系统 = True Then
            strTmp = " And (NVL(b.申请序号,0)=0 Or b.诊疗类别 <>'K')" & _
                " Union All " & _
                " Select 0, No, 记录性质, '-17', Decode(类别, 1, '输血申请单', '取血通知单') 名称, '对病人进行输血治疗的申请单据', 医嘱id, 类别" & vbNewLine & _
                " From (Select b.No, b.记录性质, b.医嘱id, Decode(C.操作类型, '8', Nvl(C.执行分类, 0), 0)+1 类别" & vbNewLine & _
                "       From 诊疗项目目录 c, 病人医嘱记录 d, 病人医嘱记录 a, 病人医嘱发送 b" & vbNewLine & _
                "       Where Instr(',8,9,', ',' || c.操作类型 || ',') > 0 And c.Id = d.诊疗项目id And d.诊疗类别 = 'E' And d.相关id = a.Id And" & vbNewLine & _
                "             a.Id = b.医嘱id And a.医嘱期效 = 1 And b.发送号 = [1] And Nvl(a.申请序号, 0) <> 0 And a.诊疗类别 = 'K' And a.医嘱状态 = 8)"
        Else
            strTmp = " And (NVL(b.申请序号,0)=0 Or b.诊疗类别 <>'K')" & _
                " Union All " & _
                " Select 0,B.NO,B.记录性质,'-17','输血申请单','对病人进行输血治疗的申请单据',B.医嘱ID,0 From 病人医嘱记录 A,病人医嘱发送 B Where A.ID=B.医嘱ID And A.医嘱期效=1 And B.发送号=[1] And NVL(A.申请序号,0)<>0 And A.诊疗类别 = 'K' And A.医嘱状态=8 "
        End If
    End If
    
    '包含申请诊疗单据,根据单据编号调用报表(相当于通知单)
    strSql = "Select Distinct D.ID,A.NO,A.记录性质,D.编号,D.名称,D.说明,0 AS 医嘱ID,0 类别" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,病历单据应用 C,病历文件列表 D" & _
        " Where A.发送号=[1] And A.医嘱ID=B.ID" & _
        " And B.诊疗项目ID=C.诊疗项目ID And C.应用场合=[2] and (not D.说明 like '%<新开时打印>%' Or NVL(D.格式,0)<>1)" & _
        " And C.病历文件ID=D.ID And D.种类=7" & _
        strTmp & _
        " Order by NO"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng发送号, mint场合)
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID & "_" & rsTmp!NO & "_" & rsTmp!记录性质, rsTmp!NO)
        objItem.SubItems(1) = Nvl(rsTmp!名称)
        objItem.SubItems(2) = Nvl(rsTmp!说明)
        '如果小于0表示使用病区固定报表
        If Val(rsTmp!编号 & "") < 0 And Val(rsTmp!ID & "") = 0 Then
            objItem.Tag = "ZL1_INSIDE_1254_" & Abs(Val(rsTmp!编号 & "")) & IIF(Val(rsTmp!类别 & "") = 0, "", "_" & Val(rsTmp!类别 & "")) '对应的自定义报表编号
        Else
            objItem.Tag = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
        End If
        objItem.ListSubItems(1).Tag = rsTmp!记录性质
        objItem.ListSubItems(2).Tag = rsTmp!医嘱ID
        objItem.Checked = True
        rsTmp.MoveNext
    Next
    LoadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lvwBill_DblClick()
    If mblnItem Then Call lvwBill_KeyPress(13)
End Sub

Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If mblnItem Then
        Item.Selected = True
        Item.EnsureVisible
    End If
End Sub

Private Sub lvwBill_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdSetup_Click
    End If
End Sub

Private Sub lvwBill_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    Dim strSql As String
    
    '申请单据打印之后的处理
    If mstrBillPrint <> "" Then
        If Split(mstrBillPrint, ",")(0) = ReportNum Then
            strSql = "Zl_诊疗单据打印_Insert('" & Split(mstrBillPrint, ",")(1) & "'," & Val(Split(mstrBillPrint, ",")(2)) & ",1,'" & UserInfo.姓名 & "')"
        End If
    End If
    
    On Error GoTo errH
    If strSql <> "" Then
        zlDatabase.ExecuteProcedure strSql, Me.Name
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
