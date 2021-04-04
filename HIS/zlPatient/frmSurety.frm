VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSurety 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "担保信息管理"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "frmSurety.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7560
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   6120
      TabIndex        =   9
      ToolTipText     =   "(F9)退出"
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   2570
      TabIndex        =   8
      ToolTipText     =   "只允许删除最近一条担保记录"
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改(&M)"
      Height          =   350
      Left            =   1470
      TabIndex        =   7
      ToolTipText     =   "只允许修改最近一条担保记录"
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Cancel          =   -1  'True
      Caption         =   "增加(&A)"
      Height          =   350
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "仅当最近一条担保记录到期或没有限制期限时才允许增加"
      Top             =   1320
      Width           =   1100
   End
   Begin VB.Frame fraEdit 
      Caption         =   "信息输入"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtReason 
         Height          =   300
         Left            =   5040
         MaxLength       =   50
         TabIndex        =   5
         Top             =   720
         Width           =   2010
      End
      Begin VB.CheckBox chk临时担保 
         Caption         =   "临时担保"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1050
      End
      Begin VB.CheckBox chkUnlimit 
         Caption         =   "不限额度"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         ToolTipText     =   "不限担保额时必须设置担保时限"
         Top             =   720
         Width           =   1050
      End
      Begin VB.TextBox txt担保人 
         Height          =   300
         Left            =   840
         MaxLength       =   100
         TabIndex        =   0
         Top             =   360
         Width           =   1005
      End
      Begin VB.TextBox txt担保额 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   9
         TabIndex        =   1
         Top             =   360
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker dtp担保时限 
         Height          =   300
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   275709955
         CurrentDate     =   38915.6041666667
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保原因"
         Height          =   180
         Left            =   4140
         TabIndex        =   16
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lbl担保时限 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到期时间"
         Height          =   180
         Left            =   4140
         TabIndex        =   15
         ToolTipText     =   "在院病人才能使用时限担保"
         Top             =   450
         Width           =   720
      End
      Begin VB.Label lbl担保人 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   450
         Width           =   540
      End
      Begin VB.Label lbl担保额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   180
         Left            =   2160
         TabIndex        =   12
         Top             =   450
         Width           =   540
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   2265
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   3995
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483645
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   16777215
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   4080
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9472
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3775
            MinWidth        =   3775
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSurety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mlng病人ID As Long
Public mbln在院病人 As Boolean
Public mstrPrivs As String
Private mlng主页ID As Long      '门诊病人或出院病人为0,在院病人为当前住院登记的主页ID

Private Sub chkUnlimit_Click()
     '不限担保额不能是临时担保
    If chkUnlimit.Value = 1 And IsNull(dtp担保时限.Value) Then
        dtp担保时限.Value = DateAdd("d", 3, dtp担保时限.MinDate)
    End If
    chk临时担保.Enabled = Not (chkUnlimit.Value = 1)
    txt担保额.Enabled = Not (chkUnlimit.Value = 1)
    
    If chkUnlimit.Value = 1 Then
        txt担保额.Text = "999999999":  txt担保额.BackColor = vbInactiveCaptionText
    Else
        txt担保额.Text = "": txt担保额.BackColor = vbWhite
    End If
End Sub

Private Sub chk临时担保_Click()
    If chk临时担保.Value = 1 Then
        dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm"
        dtp担保时限.Value = Null
        chkUnlimit.Value = 0        '值改变时有隐式调用click事件
    End If
    chkUnlimit.Enabled = Not (chk临时担保.Value = 1) And mbln在院病人
    dtp担保时限.Enabled = Not (chk临时担保.Value = 1) And mbln在院病人
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    Dim str登记时间 As String
    Dim str删除标志 As String
    
    '问题21368 by lesfeng 2010-08-02
    str删除标志 = Trim(msh.TextMatrix(msh.Row, GetColNum("删除标志")))
    If str删除标志 = "删除" Then
        MsgBox "此条担保记录已经为删除标记，不能进行删除标记操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("确实要进行标记此条担保记录为删除标记吗?" & vbCrLf & vbCrLf & "注意,删除标记后，当前担保将会不能恢复!" _
        , vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errH
    
    If Trim(msh.TextMatrix(msh.Row, GetColNum("登记时间"))) = "" Then
        str登记时间 = "NULL"
    Else
        str登记时间 = zlStr.To_Date(Trim(msh.TextMatrix(msh.Row, GetColNum("登记时间"))))
    End If
    '问题21368 by lesfeng 2010-08-02
    strSQL = "zl_病人担保记录_delete(" & mlng病人ID & "," & mlng主页ID & ",NULL," & str登记时间 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    stbThis.Panels(1).Text = "删除操作成功!"
    Call LoadSurety
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdModify_Click()
    Dim strSQL As String, str担保人 As String, str到期时间 As String
    Dim str登记时间 As String
    Dim str删除标志 As String
    '只能修改当前选中并且有效的担保记录
    
    
    If cmdModify.Caption = "修改(&M)" Then
    '提取修改信息
        If msh.TextMatrix(msh.Row, GetColNum("担保人")) = "" Then
            stbThis.Panels(1).Text = "没有可以修改的担保信息!"
            Exit Sub
        End If
        '问题21368 by lesfeng 2010-08-02
        str删除标志 = Trim(msh.TextMatrix(msh.Row, GetColNum("删除标志")))
        If str删除标志 = "删除" Then
            MsgBox "此条担保记录已经为删除标记，不能进行修改操作！", vbInformation, gstrSysName
            Exit Sub
        End If
        cmdModify.Caption = "保存(&S)"
        cmdAdd.Enabled = False
        cmdDel.Enabled = False
        cmdExit.Caption = "取消(&C)"
        fraEdit.Enabled = True
        
        With msh
            txt担保人.Text = Trim(.TextMatrix(.Row, GetColNum("担保人")))
            If .TextMatrix(.Row, GetColNum("担保额")) = "不限" Then
                chkUnlimit.Value = 1    '值不同时隐式调用click事件
                txt担保额.Text = "999999999"
            Else
                chkUnlimit.Value = 0
                txt担保额.Text = Val(.TextMatrix(.Row, GetColNum("担保额")))
            End If
            
            If IsDate(.TextMatrix(.Row, GetColNum("到期时间"))) Then
                dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm"
                dtp担保时限.Value = CDate(.TextMatrix(.Row, GetColNum("到期时间")))
            Else
                dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm" '如果不可见，下面句执行会出错
                dtp担保时限.Value = Null
            End If
            
            chk临时担保.Value = IIf(.TextMatrix(.Row, GetColNum("临时担保")) = "√", 1, 0)
            If txt担保人.Enabled Then txt担保人.SetFocus
            txt担保人.Tag = Trim(.TextMatrix(msh.Row, GetColNum("登记时间")))
        End With
    Else
    '保存修改结果
        '1.数据检查
        If Not Check担保信息 Then Exit Sub
        
        
        '先恢复界面按钮状态
        cmdModify.Caption = "修改(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "退出(&X)"
        fraEdit.Enabled = True      'SetCanEdit会再次设置
        
        str担保人 = Replace(Trim(txt担保人.Text), "'", "''")
        str到期时间 = "null"
        If Not IsNull(dtp担保时限.Value) Then str到期时间 = zlStr.To_Date(dtp担保时限.Value)
        str登记时间 = zlStr.To_Date(txt担保人.Tag)
        
        '长度检查
        If Not CheckLen(txt担保人, 64) Then Exit Sub
        
        '2.数据保存
        On Error GoTo errH
        strSQL = "zl_病人担保记录_update(" & mlng病人ID & "," & mlng主页ID & ",'" & str担保人 & "'," & _
            Val(txt担保额.Text) & "," & chk临时担保.Value & ",'" & Trim(txtReason.Text) & "',NULL," & str到期时间 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & str登记时间 & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                
        '3.数据刷新
        stbThis.Panels(1).Text = "修改结果已保存!"
        Call LoadSurety
        Call Init担保信息
        If cmdExit.Enabled Then cmdExit.SetFocus
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Init担保信息()
    Dim datsys As Date

    txt担保人.Text = ""
    chkUnlimit.Enabled = mbln在院病人
    chkUnlimit.Value = 0            '如果值有变化,则隐式调用click事件
    txt担保额.Text = ""
    txtReason.Text = ""
    
    dtp担保时限.Enabled = mbln在院病人
    dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm" '设置checkbox可见性
    If dtp担保时限.Enabled Then
        datsys = zlDatabase.Currentdate
        dtp担保时限.MinDate = datsys
        dtp担保时限.Value = DateAdd("d", 3, datsys)
    End If
    dtp担保时限.Value = Null
    
    chk临时担保.Enabled = True
    chk临时担保.Value = 0
    chkUnlimit.TabStop = True
End Sub

Private Sub dtp担保时限_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    ElseIf KeyAscii = vbKeySpace Then
        If dtp担保时限.CheckBox Then
            KeyAscii = 0
            If IsNull(dtp担保时限.Value) Then
                dtp担保时限.Value = DateAdd("d", 3, zlDatabase.Currentdate)
            Else
                dtp担保时限.Value = Null
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
        
    Dim strSQL  As String
    Dim rsTmp As New ADODB.Recordset
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadSurety
    Call Init担保信息
    
    'Call GetSuretyBalance   '初始mlng主页id
    
    If InStr(mstrPrivs, "担保信息增加") <= 0 Then
        cmdAdd.Visible = False
    End If
    
    If InStr(mstrPrivs, "担保信息调整") <= 0 Then
        cmdModify.Visible = False
    End If
    
    If InStr(mstrPrivs, "担保信息删除") <= 0 Then
        cmdDel.Visible = False
    End If
    
    If InStr(mstrPrivs, "担保信息增加") <= 0 And InStr(mstrPrivs, "担保信息调整") And InStr(mstrPrivs, "担保信息删除") <= 0 Then
        fraEdit.Enabled = False
        Me.Caption = "担保信息查看(当前用户：" & UserInfo.姓名 & ")"
    End If
    
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
    GetColNum = -1
End Function

Private Sub SetHeader()
    Dim strHead As String, i As Long
    '问题21368 by lesfeng 2010-08-02
    strHead = ",4,300|类别,4,1000|担保人,4,800|担保额,7,1250|临时担保,4,850|担保原因,4,1800|登记时间,1,1800|到期时间,1,1800|删除标志,4,850|操作员姓名,4,1050|操作员编号,4,1050|删除操作员姓名,4,1050|删除操作员编号,4,1050|删除时间,1,1800"
    With msh
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(msh, App.ProductName & "\" & Me.Name)
        
        .ForeColor = &H80000003
        .RowHeight(0) = 320
        .Redraw = True
    End With
End Sub

Private Sub GetSuretyBalance()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = _
        " Select To_char(担保额,'99999999990.00') as 担保额,Decode(当前科室ID,null,0,主页ID) as 主页ID" & _
        " From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    If rsTmp.RecordCount > 0 Then
        stbThis.Panels(2).Text = "有效担保额:" & IIf(IsNull(rsTmp!担保额), "无", Val(Trim("" & rsTmp!担保额)))
        mlng主页ID = Val("" & rsTmp!主页ID)
    Else
        stbThis.Panels(2).Text = "有效担保额:无"
        mlng主页ID = 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSurety()
    Dim rsTmp As ADODB.Recordset, Curdate As Date
    Dim strSQL As String, i As Integer, lngRow As Integer, RowPageid As Integer
    Dim str删除标志 As String
    
    On Error GoTo errH
    Curdate = zlDatabase.Currentdate
    '问题21368 by lesfeng 2010-08-02
    '删除标志,4,850|操作员姓名,4,1050|操作员编号,4,1050|删除操作员姓名,4,1050|删除操作员编号,4,1050|删除时间,1,1800"
    strSQL = _
        "SELECT '',Decode(主页id, NULL, '门诊', '第' || 主页id || '次住院') 类别, 担保人," & vbNewLine & _
        "       Decode(担保额, 999999999, '不限', To_Char(担保额, '999999990.00')) AS 担保额," & vbNewLine & _
        "       Decode(担保性质, 1, '√', ' ') AS 临时担保, 担保原因, To_Char(登记时间, 'yyyy-mm-dd hh24:mi:ss') 登记时间," & vbNewLine & _
        "       To_Char(到期时间, 'yyyy-mm-dd hh24:mi:ss') 到期时间,decode(删除标志,1,'',-1,'删除','') as 删除标志," & vbNewLine & _
        "       操作员姓名,操作员编号,删除操作员姓名,删除操作员编号,删除时间" & vbNewLine & _
        "FROM 病人担保记录" & vbNewLine & _
        "WHERE 病人id = [1]" & vbNewLine & _
        "ORDER BY 登记时间 DESC"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    If rsTmp.RecordCount > 0 Then
        Set msh.DataSource = rsTmp
    Else
        msh.Clear
        msh.Rows = 2
    End If
    Call SetHeader
    GetSuretyBalance
    For lngRow = 1 To msh.Rows - 1
        If UBound(Split(Trim(msh.TextMatrix(lngRow, GetColNum("类别"))), "次住院")) > 0 Then '取出选中行主页ID
            RowPageid = Val(Split(Split(Trim(msh.TextMatrix(lngRow, GetColNum("类别"))), "次住院")(0), "第")(1))
        Else
            RowPageid = 0
        End If
        '问题21368 by lesfeng 2010-08-02
        str删除标志 = Trim(msh.TextMatrix(lngRow, GetColNum("删除标志")))
        
        If mlng主页ID = RowPageid And (Trim(msh.TextMatrix(lngRow, GetColNum("到期时间"))) = "" Or Trim(msh.TextMatrix(lngRow, GetColNum("到期时间"))) > Curdate) Then
            msh.Row = lngRow
            For i = 0 To msh.Cols - 1
                msh.Col = i
                '问题21368 by lesfeng 2010-08-02
                If str删除标志 = "" Then
                    msh.CellForeColor = &HC00000
                Else
                    msh.CellForeColor = &HFF&
                End If
            Next
        Else
             For i = 0 To msh.Cols - 1
                msh.Col = i
                '问题21368 by lesfeng 2010-08-02
                If str删除标志 = "" Then
                Else
                    msh.CellForeColor = &HFF&
                End If
            Next
        End If
        
    Next lngRow
    msh.Row = 1
    msh.Col = 0: msh.ColSel = msh.Cols - 1
    Call msh_EnterCell
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Check担保信息() As Boolean
    Check担保信息 = True
        
    If Trim(txt担保人.Text) = "" Then
        stbThis.Panels(1).Text = "请输入担保人姓名,担保人不能为空!"
        If txt担保人.Enabled Then txt担保人.SetFocus
        Check担保信息 = False
        Exit Function
    End If
    
    If Not IsNumeric(txt担保额.Text) Then
        stbThis.Panels(1).Text = "请输入正确的担保额,担保额要求是数值!"
        If txt担保额.Enabled Then txt担保额.SetFocus
        Check担保信息 = False
        Exit Function
    ElseIf Val(txt担保额.Text) = 0 Then
        stbThis.Panels(1).Text = "请输入担保额,担保额不能为零!"
        If txt担保额.Enabled Then txt担保额.SetFocus
        Check担保信息 = False
        Exit Function
    End If
    
    If chk临时担保.Value = 1 Then
        If Not IsNull(dtp担保时限.Value) Or chkUnlimit.Value = 1 Then
            stbThis.Panels(1).Text = "临时担保不允许设置担保时限或不限担保额!"
            If chk临时担保.Enabled Then chk临时担保.SetFocus
            Check担保信息 = False
            Exit Function
        End If
    End If
    
    If zlCommFun.ActualLen(Trim(txtReason.Text)) > 50 Then
        stbThis.Panels(1).Text = "担保原因过长，最多允许 25 个汉字或 50 个字符。"
        txtReason.SetFocus
        Check担保信息 = False
        Exit Function
    End If
    
End Function

Private Sub cmdAdd_Click()
    Dim str担保人 As String, str到期时间 As String
    Dim strSQL As String, i As Integer, Curdate As Date, bln未到期 As Boolean, bln临时 As Boolean, RowPageid As Integer
    Dim str删除标志 As String
    
    '1.数据检查
    If Not Check担保信息 Then Exit Sub
    
    Curdate = zlDatabase.Currentdate
    
    For i = 1 To msh.Rows - 1 '判断本次住院未到期的担保记录，加以提示
         If Trim(msh.TextMatrix(i, GetColNum("类别"))) <> "" Then
            If UBound(Split(Trim(msh.TextMatrix(i, GetColNum("类别"))), "次住院")) > 0 Then '取出选中行主页ID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(i, GetColNum("类别"))), "次住院")(0), "第")(1))
            Else
                RowPageid = 0
            End If
            If mlng主页ID = RowPageid Then
                '问题21368 by lesfeng 2010-08-02
                str删除标志 = Trim(msh.TextMatrix(i, GetColNum("删除标志")))
               If (Trim(Nvl(msh.TextMatrix(i, GetColNum("到期时间")))) = "" Or Nvl(msh.TextMatrix(i, GetColNum("到期时间"))) > Curdate) And str删除标志 = "" Then
                   bln临时 = Nvl(msh.TextMatrix(i, GetColNum("临时担保"))) = "√"
                   bln未到期 = True: Exit For
               End If
            End If
        End If
    Next
    
    If bln未到期 Then
        If MsgBox("尚有未到期的" & IIf(bln临时, "临时", "") & "担保记录，新增将会" & IIf(bln临时, "让之前的临时担保自动失效", "累计担保") & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
        
    str担保人 = Replace(Trim(txt担保人.Text), "'", "''")
    str到期时间 = "null"
    If Not IsNull(dtp担保时限.Value) Then str到期时间 = "To_Date('" & Format(dtp担保时限.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    '长度检查
    If Not CheckLen(txt担保人, 64) Then Exit Sub
    
    '2.数据保存
    On Error GoTo errH
    
    strSQL = "zl_病人担保记录_insert(" & mlng病人ID & "," & mlng主页ID & ",'" & str担保人 & "'," & _
        Val(txt担保额.Text) & "," & chk临时担保.Value & ",'" & Trim(txtReason.Text) & "',Null," & str到期时间 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '3.数据刷新
    stbThis.Panels(1).Text = "新增信息已保存!"
    Call LoadSurety
    Call Init担保信息
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    
    If cmdExit.Caption = "取消(&C)" Then
        cmdModify.Caption = "修改(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "退出(&X)"
        fraEdit.Enabled = True      'SetCanEdit会再次设置
       
        '刷新数据,考虑并发操作
        stbThis.Panels(1).Text = ""
        Call LoadSurety
        Call Init担保信息
    Else
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        Call cmdExit_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdModify.Caption = "保存(&S)" Then
        If MsgBox("当前修改的信息未保存,确实要退出吗?", vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Cancel = 1
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub
Private Sub msh_EnterCell()
    Dim str到期时间 As String
    Dim datsys As Date, RowPageid As Integer
    Dim str删除标志 As String
   '当前行主页与病人主页不同时不允许修改删除,已过期不允许修改删除
    datsys = zlDatabase.Currentdate
    
    '问题21368 by lesfeng 2010-08-02
    str删除标志 = Trim(msh.TextMatrix(msh.Row, GetColNum("删除标志")))
    
    If cmdModify.Caption = "修改(&M)" Then
        If mlng主页ID = 0 And Trim(msh.TextMatrix(msh.Row, GetColNum("类别"))) = "门诊" Then
            '问题21368 by lesfeng 2010-08-02
            If str删除标志 = "" Then
                cmdModify.Enabled = True
                cmdDel.Enabled = True
                stbThis.Panels(1).Text = "当前担保记录有效"
            Else
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "当前担保记录已经标记删除"
            End If
        Else
            If UBound(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("类别"))), "次住院")) > 0 Then '取出选中行主页ID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("类别"))), "次住院")(0), "第")(1))
            Else
                RowPageid = 0
            End If
            If mlng主页ID <> RowPageid Then
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "当前担保记录非本次住院担保。"
            Else
                str到期时间 = Trim(msh.TextMatrix(msh.Row, GetColNum("到期时间")))
            
                If str到期时间 <> "" Then
                    If CDate(str到期时间) < datsys Then
                         cmdModify.Enabled = False
                         cmdDel.Enabled = False
                        '问题21368 by lesfeng 2010-08-02
                         If str删除标志 = "" Then
                            stbThis.Panels(1).Text = "当前担保记录已过期"
                        Else
                            stbThis.Panels(1).Text = "当前担保记录已经标记删除"
                        End If
                    Else
                        '问题21368 by lesfeng 2010-08-02
                        If str删除标志 = "" Then
                            cmdModify.Enabled = True
                            cmdDel.Enabled = True
                            stbThis.Panels(1).Text = "当前担保记录有效"
                        Else
                            cmdModify.Enabled = False
                            cmdDel.Enabled = False
                            stbThis.Panels(1).Text = "当前担保记录已经标记删除"
                        End If
                    End If
                Else
                    '问题21368 by lesfeng 2010-08-02
                    If str删除标志 = "" Then
                        cmdModify.Enabled = True
                        cmdDel.Enabled = True
                        stbThis.Panels(1).Text = "当前担保记录有效"
                    Else
                        cmdModify.Enabled = False
                        cmdDel.Enabled = False
                        stbThis.Panels(1).Text = "当前担保记录已经标记删除"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtReason_GotFocus()
    zlControl.TxtSelAll txtReason
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    Else
        If InStr("'|?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txtReason, KeyAscii
    End If
End Sub

Private Sub txtReason_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt担保额_GotFocus()
    zlControl.TxtSelAll txt担保额
End Sub

Private Sub txt担保额_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        If KeyAscii = vbKeyReturn Then
            chkUnlimit.TabStop = (txt担保额.Text = "")
            SendKeys "{Tab}"
        Else
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc(".") And InStr(txt担保额.Text, ".") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt担保额_LostFocus()
    If IsNumeric(txt担保额.Text) Then
        If txt担保额.Text = "999999999" Then
            stbThis.Panels(1).Text = "不允许输入该值，该值表示无限担保．"
            If txt担保额.Enabled Then txt担保额.SetFocus
        Else
            txt担保额.Text = Format(txt担保额.Text, "0.00")
        End If
    Else
        txt担保额.Text = ""
    End If
    
    Call zlCommFun.OpenIme
End Sub

Private Sub txt担保人_GotFocus()
    zlControl.TxtSelAll txt担保人
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt担保人_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt担保人, KeyAscii
    End If
End Sub

Private Sub txt担保人_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub
