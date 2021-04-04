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

Private mlng发送号 As Long
Private mint场合 As Integer
Private mlng前提ID As Long
Private mint打印方式 As Integer
Private mblnItem As Boolean

Public Sub ShowMe(ByVal lng发送号 As Long, ByVal int场合 As Integer, frmParent As Object, Optional ByVal lng前提ID As Long)
'参数：lng发送号=本次发送的发送号
'      int场合=1-门诊,2-住院(数据场合,不是调用场合)
    mlng发送号 = lng发送号
    mint场合 = int场合
    mlng前提ID = lng前提ID
    
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
        Call ReportOpen(gcnOracle, glngSys, .Tag, Me, "NO=" & .Text, "性质=" & Val(.ListSubItems(1).Tag), 1)
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
    
    cmdPrint.Enabled = False
    Screen.MousePointer = 11
    For i = 1 To lvwBill.ListItems.Count
        With lvwBill.ListItems(i)
            If .Checked Then
                .Selected = True: .EnsureVisible: Me.Refresh
                Call ReportOpen(gcnOracle, glngSys, .Tag, Me, "NO=" & .Text, "性质=" & Val(.ListSubItems(1).Tag), 2)
                
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
    Call ReportPrintSet(gcnOracle, glngSys, lvwBill.SelectedItem.Tag, Me)
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
    If mint场合 = 1 And mlng前提ID = 0 Then
        mint打印方式 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊发送单据打印", 1))
    Else
        mint打印方式 = 1
    End If
    If mint打印方式 = 0 Then Unload Me: Exit Sub
    
    Call RestoreListViewState(lvwBill, App.ProductName & "\" & Me.Name, "")
    If Not LoadBill Then Unload Me: Exit Sub
    If lvwBill.ListItems.Count = 0 Then Unload Me: Exit Sub
    mblnItem = False
    
    '自动打印后退出
    If mint打印方式 = 2 Then
        Call cmdPrint_Click
        Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveListViewState(lvwBill, App.ProductName & "\" & Me.Name, "")
End Sub

Private Sub lvwBill_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwBill, ColumnHeader.Index)
End Sub

Private Function LoadBill() As Boolean
'功能：读取本次发送可以打印的诊疗单据列表
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objItem As ListItem
    
    lvwBill.ListItems.Clear
    
    On Error GoTo errH
    
    '包含单病人单/多记录的诊疗单据,根据单据编号调用报表(相当于通知单)
    strSQL = "Select Distinct D.ID,A.NO,A.记录性质,D.编号,D.名称,D.说明" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,诊疗单据应用 C,病历文件目录 D" & _
        " Where A.发送号=[1] And A.医嘱ID=B.ID" & _
        " And B.诊疗项目ID=C.诊疗项目ID And C.应用场合=[2]" & _
        " And C.病历文件ID=D.ID And D.种类=5" & _
        " And D.前提 IN([2],3) And D.书写 IN(1,2)" & _
        " Order by A.NO"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng发送号, mint场合)
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID & "_" & rsTmp!NO & "_" & rsTmp!记录性质, rsTmp!NO)
        objItem.SubItems(1) = Nvl(rsTmp!名称)
        objItem.SubItems(2) = Nvl(rsTmp!说明)
        objItem.Tag = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
        objItem.ListSubItems(1).Tag = rsTmp!记录性质
        objItem.Checked = True
        rsTmp.MoveNext
    Next
    LoadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
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

Private Sub lvwBill_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnItem = False
End Sub
