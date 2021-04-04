VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceUse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结算方式应用设置"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmBalanceUse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3780
      TabIndex        =   1
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3780
      TabIndex        =   2
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3780
      TabIndex        =   3
      Top             =   3360
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw结算方式 
      Height          =   2925
      Left            =   90
      TabIndex        =   0
      Top             =   1230
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   5159
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_应用场合"
         Object.Tag             =   "应用场合"
         Text            =   "应用场合"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "_缺省标志"
         Object.Tag             =   "缺省标志"
         Text            =   "缺省标志"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.Label lbl提示 
      Caption         =   $"frmBalanceUse.frx":000C
      Height          =   915
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   3285
   End
End
Attribute VB_Name = "frmBalanceUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr场合 As String
Dim mblnItem As Boolean
Dim mintSuccess As Integer
Dim mblnChange As Boolean     '是否改变了

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, 5
End Sub

Private Sub cmdOK_Click()
    If Save结算() = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    mblnChange = False
    Unload Me
End Sub

Private Function Save结算() As Boolean
'功能:保存编辑的内容结算方式表中
'参数:
'返回值:成功返回True,否则为False
    Dim i As Integer
    Dim str场合 As String
    On Error GoTo ErrHandle
    '把所有选中的工作性质做成一个串
    '把所有选中的工作性质做成一个串
    For i = 1 To lvw结算方式.ListItems.Count
        If lvw结算方式.ListItems(i).Checked = True Then
            str场合 = str场合 & lvw结算方式.ListItems(i) & ":"
            str场合 = str场合 & IIF(lvw结算方式.ListItems(i).SubItems(1) = "", "0;", "1;")
        End If
    Next
    
    '修改
    gstrSQL = "zl_结算方式应用_update( '" & mstr场合 & "','" & str场合 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnChange = False
    Save结算 = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 编辑场合(ByVal str场合 As String) As Boolean
'功能:用来与调用的结算方式管理窗口进行通讯的程序
'参数:str场合     当前编辑的结算方式的编码
'返回值:编辑成功返回True,否则为False
    Dim rs结算方式 As New ADODB.Recordset
    Dim ObjItem As Object
    mblnChange = False
    mintSuccess = 0
        
    mstr场合 = str场合
    lbl提示.Caption = Replace(lbl提示.Caption, "本结算方式", "在『" & str场合 & "』结算场合中")
    
    '读出结算场合:代收款只应用于预交款
    '除了收费,结帐以外,不能应用于应付款:33722
    '75134:李南春,2014/7/14,排除性质为9的结算方式
    '82990:李南春,2015/3/9,医保不能用于补结算
    On Error GoTo ErrHandle
    gstrSQL = "Select A.名称,A.性质,B.结算方式,B.缺省标志,nvl(A.应付款,0) as 应付款" & _
        " From 结算方式 A,结算方式应用 B" & _
        " Where A.名称=B.结算方式(+) And B.应用场合(+)=[1] " & _
        IIF(str场合 <> "预交款", " And A.性质<>5", "") & _
        IIF(str场合 = "补结算" Or str场合 = "消费卡", " And A.性质 In(1,2,8)", "") & _
        IIF(InStr("收费,结帐", str场合) = 0, " And nvl(A.应付款,0)<>1 ", "") & _
        " And A.性质<>9 And b.付款方式(+) Is Null" & _
        " Order by A.编码"
        
    Set rs结算方式 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str场合)
        
    lvw结算方式.ListItems.Clear
    Do Until rs结算方式.EOF
       Set ObjItem = lvw结算方式.ListItems.Add(, "C" & rs结算方式!名称, rs结算方式!名称)
        ObjItem.Tag = Nvl(rs结算方式!性质, 1) & "," & Val(Nvl(rs结算方式!应付款))
        
        If Not IsNull(rs结算方式!结算方式) Then
            ObjItem.Checked = True
            If Nvl(rs结算方式!性质, 1) < 3 Then
                ObjItem.SubItems(1) = IIF(Nvl(rs结算方式!缺省标志, 0) = 1, "缺省", "")
            End If
        End If
        rs结算方式.MoveNext
    Loop
    
    frmBalanceUse.Show vbModal
    编辑场合 = mintSuccess > 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub lvw结算方式_DblClick()
    If mblnItem = False Then Exit Sub
    Call ChangeServer
    mblnItem = False
End Sub

Private Sub lvw结算方式_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("C") Or KeyAscii = Asc("c") Then Call ChangeServer
End Sub

Private Sub ChangeServer()
    Dim i As Integer, j As Integer
    Dim varData As Variant
    If lvw结算方式.SelectedItem Is Nothing Then Exit Sub
    
    With lvw结算方式.SelectedItem
        If .Checked = False Then Exit Sub
        varData = Split(.Tag & ",", ",")
        If InStr("1,2,7,8", Val(varData(0))) = 0 Then .SubItems(1) = "": Exit Sub '医保结算及代收款不能为缺省结算方式
        If Val(varData(1)) = 1 Then .SubItems(1) = "": Exit Sub '应付款不能设置成缺省方式
        
        If .SubItems(1) = "" Then
            .SubItems(1) = "缺省"
            mblnChange = True
        Else
            .SubItems(1) = ""
            mblnChange = True
        End If
        cmdOK.Enabled = True
    End With
    
    '保证唯一性
    If lvw结算方式.SelectedItem.SubItems(1) <> "" Then
        j = lvw结算方式.SelectedItem.Index
        For i = 1 To lvw结算方式.ListItems.Count
            If i <> j Then
                lvw结算方式.ListItems(i).SubItems(1) = ""
            End If
        Next
    End If
End Sub

Private Sub lvw结算方式_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    cmdOK.Enabled = True
    mblnChange = True
    If Item.Checked = False And Item.SubItems(1) = "缺省" Then Item.SubItems(1) = ""
End Sub

Private Sub lvw结算方式_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub
