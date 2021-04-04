VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结算方式编辑"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmBalanceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5010
      TabIndex        =   14
      Top             =   4560
      Width           =   1100
   End
   Begin VB.Frame fra场合 
      Caption         =   "应用场合"
      Height          =   3390
      Left            =   150
      TabIndex        =   9
      Top             =   1935
      Width           =   4680
      Begin MSComctlLib.ListView lvw场合 
         Height          =   2250
         Left            =   105
         TabIndex        =   11
         Top             =   1020
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   3969
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
         Caption         =   $"frmBalanceEdit.frx":000C
         Height          =   735
         Left            =   195
         TabIndex        =   10
         Top             =   300
         Width           =   4305
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5010
      TabIndex        =   13
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5010
      TabIndex        =   12
      Top             =   255
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本信息"
      Height          =   1725
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4635
      Begin VB.CheckBox chk应付款 
         Caption         =   "应付款"
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   1328
         Width           =   855
      End
      Begin VB.CheckBox chkDue 
         Caption         =   "应收款"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   1328
         Width           =   975
      End
      Begin VB.ComboBox cmb 
         Height          =   300
         ItemData        =   "frmBalanceEdit.frx":0095
         Left            =   840
         List            =   "frmBalanceEdit.frx":0097
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1305
         Width           =   1850
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   2
         TabIndex        =   2
         Top             =   210
         Width           =   405
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   840
         MaxLength       =   10
         TabIndex        =   4
         Top             =   570
         Width           =   3675
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   840
         MaxLength       =   4
         TabIndex        =   6
         Top             =   930
         Width           =   1850
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "性质(&Q)"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   7
         Top             =   1350
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编号(&U)"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   630
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   990
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmBalanceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr编码 As String        '原始的编码
Dim mstr名称 As String        '原始的名称
Dim mbln固定 As Boolean       '方式是否固定
Dim mblnItem As Boolean
Dim mblnChange As Boolean     '是否改变了
Dim mintSuccess As Integer
Dim mblnCancel As Boolean     '取消编辑

Private Function CheckUsedDue() As Boolean
'检查是否选择了结帐场合
    Dim i As Long
    If cmb.ListIndex <> -1 Then
        If cmb.ListIndex = 0 Or cmb.ListIndex = 1 Or cmb.ListIndex = 3 Then
            For i = 1 To lvw场合.ListItems.Count
                If lvw场合.ListItems(i).Checked = True Then
                    If lvw场合.ListItems(i).Text = "结帐" Then CheckUsedDue = True: Exit Function
                End If
            Next
        End If
    End If
End Function
Private Function IsCheckDueValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否使用了收费或结帐场合,并且未使用其他场合
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-04 11:01:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnUserOther As Boolean, blnUser As Boolean
    IsCheckDueValied = False
    If cmb.ListIndex <> -1 Then
        If cmb.ListIndex = 0 Or cmb.ListIndex = 1 Then
            With lvw场合
                For i = 1 To .ListItems.Count
                    If .ListItems(i).Checked = True Then
                        If InStr(1, ";收费;结帐;", ";" & .ListItems(i).Text & ";") > 0 Then
                            blnUser = True
                        Else
                            blnUserOther = True: Exit For
                        End If
                    End If
                Next
                IsCheckDueValied = Not blnUserOther And blnUser '检查是否使用了收费或结帐场合,并且未使用其他场合
            End With
        End If
    End If
End Function


Private Sub chk应付款_Click()
    Dim rsTmp As ADODB.Recordset, ObjItem As ListItem
    '应付款方式只能有一个:33722
    If chk应付款.value = 1 Then
        mblnItem = True
        '需要检查是否只有收费和结帐
        For Each ObjItem In Me.lvw场合.ListItems
            If InStr(1, ";收费;结帐;", ";" & ObjItem.Text & ";") > 0 Then
                ObjItem.Checked = True
                ObjItem.Selected = True
            Else
                ObjItem.Checked = False
            End If
            ObjItem.SubItems(1) = ""
        Next
        mblnItem = False
    End If
End Sub

Private Sub cmb_Click()
    Dim ObjItem As ListItem
    Dim rsTmp As New ADODB.Recordset
    
    If mblnCancel Then Exit Sub
    mblnChange = True
        
    On Error GoTo ErrHandle
    chkDue.Enabled = CheckUsedDue
    chk应付款.Enabled = IsCheckDueValied    '是否要使用应付款这个选项:33722
    
    If Not chk应付款.Enabled Then chk应付款.value = 0
    If Not chkDue.Enabled Then chkDue.value = 0
    
    '现金方式只能有一个
    If Trim(mstr编码) <> "" Then
        gstrSQL = "select 编码,名称,简码,nvl(性质,1) 性质,缺省标志 from 结算方式  where 编码<>[1] and nvl(性质,1)=1 "
    Else
        gstrSQL = "select 编码,名称,简码,nvl(性质,1) 性质,缺省标志 from 结算方式  where  nvl(性质,1)=1 "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr编码)
        
    If rsTmp.RecordCount > 0 Then
        If cmb.ListIndex + 1 = 1 Then
            mblnCancel = True
            cmb.ListIndex = 1
            mblnCancel = False
            Exit Sub
        End If
    End If
    
    '个人帐户只能有一个
    If Trim(mstr编码) <> "" Then
        gstrSQL = "select 编码,名称,简码,nvl(性质,1) 性质,缺省标志 from 结算方式  where 编码<>[1] and  nvl(性质,1)=3 "
    Else
        gstrSQL = "select 编码,名称,简码,nvl(性质,1) 性质,缺省标志 from 结算方式  where  nvl(性质,1)=3 "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr编码)
    
    If rsTmp.RecordCount > 0 Then
        If cmb.ListIndex + 1 = 3 Then
            mblnCancel = True
            cmb.ListIndex = 3
            For Each ObjItem In Me.lvw场合.ListItems
                ObjItem.SubItems(1) = ""
                If ObjItem.Text = "补结算" Or ObjItem.Text = "消费卡" Then
                    ObjItem.Checked = False
                End If
            Next
            mblnCancel = False
            Exit Sub
        End If
    End If
    
    If cmb.ListIndex = 4 Then
        '代收款项只能应用于预交款,且不能为缺省
        For Each ObjItem In Me.lvw场合.ListItems
            If ObjItem.Text = "预交款" Then
                ObjItem.Checked = True
                ObjItem.Selected = True
            Else
                ObjItem.Checked = False
            End If
            ObjItem.SubItems(1) = ""
        Next
    ElseIf cmb.ListIndex = 2 Or cmb.ListIndex = 3 Then
        '医保的结算方式不能为缺省结算方式
        For Each ObjItem In Me.lvw场合.ListItems
            ObjItem.SubItems(1) = ""
            If ObjItem.Text = "补结算" Or ObjItem.Text = "消费卡" Then
                ObjItem.Checked = False
            End If
        Next
    ElseIf cmb.ListIndex = 6 Then '一卡通不用于预交款和就诊卡，以及补结算和消费卡
        For Each ObjItem In Me.lvw场合.ListItems
            If ObjItem.Text = "就诊卡" Or ObjItem.Text = "预交款" Or ObjItem.Text = "补结算" Or ObjItem.Text = "消费卡" Then
                ObjItem.Checked = False
            End If
        Next
    ElseIf cmb.ListIndex = 5 Then
        For Each ObjItem In Me.lvw场合.ListItems
            If ObjItem.Text = "补结算" Or ObjItem.Text = "消费卡" Then
                ObjItem.Checked = False
            End If
        Next
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save结算() = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    If mstr编码 <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstr编码 = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(1).Text = zlDatabase.GetMax("结算方式", "编码", 2)
    Me.cmb.ListIndex = 0
    For i = 1 To lvw场合.ListItems.Count
        lvw场合.ListItems(i).Checked = False
        lvw场合.ListItems(i).SubItems(1) = ""
    Next
    mblnChange = False
    txtEdit(1).SetFocus
    frmBalanceManage.Fill结算方式
End Sub

Private Function IsValid() As Boolean
    '功能:分析输入有关结算方式的内容是否有效
    '参数:
    '返回值:有效返回True,否则为False
    Dim rsTmp As ADODB.Recordset
    
    Dim i As Integer
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    For i = 1 To 3
        txtEdit(i).Text = Trim(txtEdit(i).Text)
        strTemp = txtEdit(i).Text
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox "所输入内容不能超过" & Int(txtEdit(i).MaxLength / 2) & "个汉字" & "或" & txtEdit(i).MaxLength & "个字母。", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)
    
    If Len(txtEdit(1).Text) = 0 Then
        MsgBox "编码不能为空。", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        txtEdit(2).Text = ""
        txtEdit(2).SetFocus
        Exit Function
    End If
    If chk应付款.value = 1 Then
        If Trim(mstr编码) <> "" Then
            gstrSQL = "select 编码,名称,简码,nvl(性质,1) 性质,缺省标志 from 结算方式  where 编码<>[1] and nvl(应付款,0)=1 "
        Else
            gstrSQL = "select 编码,名称,简码,nvl(性质,1) 性质,缺省标志 from 结算方式  where  nvl(应付款,0)=1 "
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr编码)
        If Not rsTmp.EOF Then
             If MsgBox("注意:" & vbCrLf & _
                              "     由于应付款性质的结算方式只能一种,而在系统中存在" & vbCrLf & _
                              "     结算方式为『" & Nvl(rsTmp!名称) & "』的应付款,如果继续操作,将会清除" & vbCrLf & _
                              "     结算方式的应付款性质,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                 Exit Function
            End If
        End If
        If Trim(mstr编码) <> "" Then
            gstrSQL = "select A.名称 from 结算方式 A,结算方式应用 B   where (A.编码=[1]) and a.名称=b.结算方式 and nvl(b.缺省标志,0)=1 "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr编码)
            If rsTmp.EOF = False Then
                 If MsgBox("注意:" & vbCrLf & _
                                  "     由于应付款性质的结算方式不能设置成缺省,而你修改的" & vbCrLf & _
                                  "     结算方式为『" & Nvl(rsTmp!名称) & "』目前正处于缺省状态,如果继续操作,将会清除" & vbCrLf & _
                                  "     此结算方式的所有缺省标志,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                     Exit Function
                End If
            End If
        End If
    End If
    IsValid = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Save结算() As Boolean
'功能:保存编辑的内容结算方式表中
'参数:
'返回值:成功返回True,否则为False
    Dim i As Integer
    Dim str场合 As String
    On Error GoTo ErrHandle
    '把所有选中的工作性质做成一个串
    '把所有选中的工作性质做成一个串
    For i = 1 To lvw场合.ListItems.Count
        If lvw场合.ListItems(i).Checked = True Then
            str场合 = str场合 & lvw场合.ListItems(i) & ":"
            If chk应付款.value = 1 Then
                str场合 = str场合 & "0;"
            Else
                str场合 = str场合 & IIF(lvw场合.ListItems(i).SubItems(1) = "", "0;", "1;")
            End If
            
        End If
    Next
    
    If mstr编码 = "" Then       '新增一条记录
        gstrSQL = "zl_结算方式_insert( '" & _
            txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & txtEdit(3).Text & "'," & cmb.ListIndex + 1 & ",'" & str场合 & "'," & IIF(chkDue.Enabled, chkDue.value, 0) & "," & IIF(chk应付款.Enabled, chk应付款.value, 0) & ")"
    Else    '修改
        gstrSQL = "zl_结算方式_update( '" & mstr编码 & "','" & _
            txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & txtEdit(3).Text & "'," & cmb.ListIndex + 1 & ",'" & str场合 & "'," & IIF(chkDue.Enabled, chkDue.value, 0) & "," & IIF(chk应付款.Enabled, chk应付款.value, 0) & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Save结算 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 编辑结算方式(ByVal str编码 As String) As Boolean
'功能:用来与调用的结算方式管理窗口进行通讯的程序
'参数:str编码     当前编辑的结算方式的编码
'返回值:编辑成功返回True,否则为False
    
    Dim rs结算方式 As New ADODB.Recordset
    
    '得到允许的数据长度
    GetDefineSize
    
    mintSuccess = 0
    rs结算方式.CursorLocation = adUseClient
    rs结算方式.CursorType = adOpenKeyset
    rs结算方式.LockType = adLockReadOnly
    
    On Error GoTo ErrHandle
    mstr编码 = str编码
    mstr名称 = ""
    If str编码 <> "" Then
        gstrSQL = "select 编码,名称,简码,nvl(性质,1) 性质,是否固定,缺省标志,Nvl(应收款,0) 应收款,Nvl(应付款,0) 应付款 from 结算方式  where 编码=[1]"
        Set rs结算方式 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str编码)
                
        '75134:李南春,2014/7/14,如果性质为9则直接退出
        If rs结算方式!性质 = 9 Then Exit Function
        txtEdit(1).Text = mstr编码
        txtEdit(2).Text = rs结算方式("名称")
        mstr名称 = rs结算方式("名称")
        txtEdit(3).Text = IIF(IsNull(rs结算方式("简码")), "", rs结算方式("简码"))
        mblnCancel = True
        cmb.ListIndex = rs结算方式!性质 - 1
        mblnCancel = False
        chkDue.value = rs结算方式!应收款
        chk应付款.value = IIF(Val(Nvl(rs结算方式!应付款)) = 1, 1, 0)
        '75134:李南春,2014/7/14,方式是否固定
        mbln固定 = IIF(rs结算方式("是否固定") = 1, True, False)
        txtEdit(2).Enabled = Not mbln固定
        txtEdit(3).Enabled = Not mbln固定
        cmb.Enabled = Not mbln固定
    Else
        txtEdit(1).Text = zlDatabase.GetMax("结算方式", "编码", 2)
    End If
    '读出结算场合
    If rs结算方式.State = 1 Then rs结算方式.Close
    gstrSQL = "Select a.名称,b.结算方式,b.缺省标志 from 结算场合 A,结算方式应用 B" & vbNewLine & _
            " Where b.应用场合(+)=a.名称 and b.结算方式(+)= [1] And b.付款方式(+) Is Null" & vbNewLine & _
            " Order by A.编码"
    Set rs结算方式 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr名称)
         
    lvw场合.ListItems.Clear
    Do Until rs结算方式.EOF
        lvw场合.ListItems.Add , "C" & rs结算方式("名称"), rs结算方式("名称")
        If rs结算方式("结算方式") = mstr名称 Then
            lvw场合.ListItems("C" & rs结算方式("名称")).Checked = True
            lvw场合.ListItems("C" & rs结算方式("名称")).SubItems(1) = IIF(rs结算方式("缺省标志") = 1, "缺省", "")
        End If
        rs结算方式.MoveNext
    Loop
    chkDue.Enabled = CheckUsedDue
    If Not chkDue.Enabled Then chkDue.value = 0
    chk应付款.Enabled = IsCheckDueValied
    If Not chk应付款.Enabled Then chk应付款.value = 0
    mblnChange = False
    frmBalanceEdit.Show vbModal
    编辑结算方式 = mintSuccess > 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If Not mbln固定 Then txtEdit(2).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()

    Me.cmb.AddItem "1-现金结算方式"
    Me.cmb.AddItem "2-其他非医保结算"
    Me.cmb.AddItem "3-医保个人帐户"
    Me.cmb.AddItem "4-医保各类统筹"
    Me.cmb.AddItem "5-代收款项"
    Me.cmb.AddItem "6-费用折扣"
    Me.cmb.AddItem "7-一卡通结算"
    Me.cmb.AddItem "8-结算卡结算"
    mblnCancel = True
    If cmb.ListCount > 1 Then Me.cmb.ListIndex = 1
    chkDue.value = 0
    mblnCancel = False
    mblnChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub lvw场合_DblClick()
    If mblnItem = False Then Exit Sub
    Call ChangeServer
    mblnItem = False
End Sub

Private Sub lvw场合_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("C") Or KeyAscii = Asc("c") Then Call ChangeServer
End Sub

Private Sub ChangeServer()
    Dim ObjItem As ListItem
    
    If lvw场合.SelectedItem Is Nothing Then Exit Sub
    
    With lvw场合.SelectedItem
        If .Checked = False Then Exit Sub
        If InStr("0,1,6,7", cmb.ListIndex) = 0 Then cmb_Click: Exit Sub '医保结算及代收款不能作为缺省项
        '应付款不能设置为缺省的结算方式
        If chk应付款.Enabled And chk应付款.value = 1 Then cmb_Click: Exit Sub
        If .SubItems(1) = "" Then
            .SubItems(1) = "缺省"
            mblnChange = True
        Else
            .SubItems(1) = ""
            mblnChange = True
        End If
    End With
End Sub

Private Sub lvw场合_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    chkDue.Enabled = CheckUsedDue
    If Not chkDue.Enabled Then chkDue.value = 0
    
    chk应付款.Enabled = IsCheckDueValied
    If Not chk应付款.Enabled Then chk应付款.value = 0
    
    '82990:李南春,2015/3/9,医保不能用于补结算
    If Item.Text = "补结算" Or Item.Text = "消费卡" Then
        If cmb.ListIndex <> 0 And cmb.ListIndex <> 1 And cmb.ListIndex <> 7 Then Item.Checked = False
    '代收款项只能应用于预交款,且不能为缺省
    ElseIf cmb.ListIndex = 4 Then
        If Item.Text <> "预交款" Then
            Item.Checked = False
        End If
    ElseIf cmb.ListIndex = 6 Then   '一卡通不用于预交和就诊卡
        If Item.Text = "预交款" Or Item.Text = "就诊卡" Then Item.Checked = False
    ElseIf cmb.ListIndex = 7 Then
        '结算卡
        If InStr(",收费,结帐,预交款,补结算,挂号,就诊卡,消费卡,", "," & Item.Text & ",") = 0 Then
            Item.Checked = False
        End If
    Else
        mblnChange = True
    End If
    If Item.Checked = False And Item.SubItems(1) = "缺省" Then Item.SubItems(1) = ""
End Sub

Private Sub lvw场合_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(txtEdit(2).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = 2 Then zlCommFun.OpenIme True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("'}|,""/", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub GetDefineSize()
'功能：得到数据库的表字段的长度
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "SELECT 编码,名称,简码 FROM 结算方式 Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "结算方式编辑")
    
    txtEdit(1).MaxLength = rsTemp.Fields("编码").DefinedSize
    txtEdit(2).MaxLength = rsTemp.Fields("名称").DefinedSize
    txtEdit(3).MaxLength = rsTemp.Fields("简码").DefinedSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
