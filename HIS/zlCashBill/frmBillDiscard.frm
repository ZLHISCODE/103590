VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillDiscard 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "票据报损"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBillDiscard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraBack 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4275
      Left            =   120
      TabIndex        =   20
      Top             =   810
      Width           =   6375
      Begin VB.ComboBox cmb报损人 
         Height          =   360
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1635
         Width           =   1830
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   1
         Left            =   1740
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1095
         Width           =   1485
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   2
         Left            =   3990
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1095
         Width           =   1485
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1410
         TabIndex        =   2
         Top             =   90
         Width           =   1815
      End
      Begin VB.OptionButton opt范围 
         Caption         =   "单张报损(&S)"
         Height          =   240
         Index           =   0
         Left            =   1425
         TabIndex        =   3
         Top             =   630
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.OptionButton opt范围 
         Caption         =   "多张报损(&M)"
         Height          =   240
         Index           =   1
         Left            =   3180
         TabIndex        =   4
         Top             =   630
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   4755
         TabIndex        =   14
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   162332675
         CurrentDate     =   37007
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报损人(&G)"
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   11
         Top             =   1695
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报损时间(&D)"
         Height          =   240
         Index           =   3
         Left            =   3375
         TabIndex        =   13
         Top             =   1710
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据种类"
         Height          =   240
         Index           =   4
         Left            =   390
         TabIndex        =   1
         Top             =   150
         Width           =   960
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "号码范围(&B)"
         Height          =   240
         Index           =   6
         Left            =   30
         TabIndex        =   5
         Top             =   1155
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   240
         Index           =   5
         Left            =   3330
         TabIndex        =   8
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label lbl说明 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   2085
         Left            =   30
         TabIndex        =   15
         Top             =   2160
         Width           =   6300
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Index           =   1
         Left            =   1410
         TabIndex        =   6
         Top             =   1095
         Width           =   315
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Index           =   2
         Left            =   3660
         TabIndex        =   9
         Top             =   1095
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   3630
      TabIndex        =   17
      Top             =   5430
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   4980
      TabIndex        =   18
      Top             =   5430
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   -270
      TabIndex        =   16
      Top             =   5160
      Width           =   7065
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   420
      Left            =   270
      TabIndex        =   19
      Top             =   5430
      Width           =   1200
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "票据报损卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   2760
   End
End
Attribute VB_Name = "frmBillDiscard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint票种 As gBillType
Private mstrPrivs As String
Private mstrID As String

Private mblnOK As Boolean
Private mblnChange As Boolean     '为真时表示已改变了
Private mstr前缀 As String
Private mstr最小号码 As String
Private mstr最大号码 As String
Private mlng票据长度 As Long
Private mblnIsBIll As Boolean '当前票种是否为票据

Private Sub InitContext()
    Dim dtCurrnet As Date
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errHandle
    dtCurrnet = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    dtpDate.Value = dtCurrnet
    dtpDate.MaxDate = dtCurrnet
    
    If mblnIsBIll Then
        lblTitle.Caption = "票据报损卡"
        lbl(6).Caption = "号码范围(&B)"
    Else
        lblTitle.Caption = IIf(mint票种 = gBillType.就诊卡, "医疗卡报损卡", "消费卡报损卡")
        lbl(6).Caption = "卡号范围(&B)"
    End If
    
    txtEdit(0).Text = _
        Choose(mint票种, "收费收据", "预交收据", "结帐收据", "挂号收据", "就诊卡", "消费卡", "会员卡")
    
    mblnChange = True
    Select Case mint票种
        Case gBillType.收费收据
            strWhere = " And B.人员性质='门诊收费员'"
        Case gBillType.预交收据
            strWhere = " And B.人员性质 in ('预交收款员','入院登记员')"
        Case gBillType.结帐收据
            strWhere = " And B.人员性质='住院结帐员'"
        Case gBillType.挂号收据
            strWhere = " And B.人员性质='门诊挂号员'"
        Case gBillType.就诊卡, gBillType.消费卡
            strWhere = " And B.人员性质 in ('发卡登记人','入院登记员')"
        Case Else
            Exit Sub
    End Select
    strSQL = _
        "Select Distinct A.姓名" & vbNewLine & _
        "From 人员表 A,人员性质说明 B" & vbNewLine & _
        "Where A.ID=B.人员ID " & strWhere & _
        "      And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        "      And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & vbNewLine & _
        "Order By A.姓名"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cmb报损人.Clear
    Do Until rsTemp.EOF
        cmb报损人.AddItem rsTemp("姓名")
        rsTemp.MoveNext
    Loop
    If cmb报损人.ListCount > 0 Then cmb报损人.ListIndex = 0

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmb报损人_Click()
    mblnChange = True
End Sub

Private Sub cmb报损人_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub dtpDate_Change()
    mblnChange = True
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If ValidateContent() = False Then Exit Sub
    If MsgBox("一旦报损后，报损" & IIf(mblnIsBIll, "号码", "卡号") & "就不能再使用了。" & vbCrLf & _
        "是否确认要继续？", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If Save() = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Sub opt范围_Click(Index As Integer)
    mblnChange = True
    If opt范围(0).Value = True Then
        txtEdit(2).Enabled = False
        txtEdit(2).Text = txtEdit(1).Text
    Else
        txtEdit(2).Enabled = True
    End If
    Call ShowSum
End Sub

Private Sub opt范围_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 1 And opt范围(0).Value = True Then txtEdit(2).Text = txtEdit(1).Text
    Call ShowSum
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyDelete _
        And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
    If (Index = 1 Or Index = 2) And (KeyAscii >= vbKey0 Or KeyAscii <= vbKey9) _
        And txtEdit(Index).SelLength = 0 Then
        If Len(txtEdit(Index)) >= mlng票据长度 Then KeyAscii = 0
    End If
End Sub

Private Function ValidateContent() As Boolean
'功能:检查输入内容的是否有效
'返回:有效则返回True,否则返回False
    Dim lngCount As Long, i As Integer
    Dim strTemp As String, strName As String
    Dim strNOs As String, varPara() As Variant, strTable As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    strName = IIf(mblnIsBIll, "号码", "卡号")
    '字符串检查
    For lngCount = 1 To 2
        txtEdit(lngCount).Text = Trim(txtEdit(lngCount).Text)
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            txtEdit(lngCount).SetFocus
            zlControl.TxtSelAll txtEdit(lngCount)
            Exit Function
        End If
        For i = 1 To Len(txtEdit(lngCount).Text)
            strTemp = Mid(txtEdit(lngCount), i, 1)
            If InStr("0123456789", strTemp) = 0 Then
                MsgBox strName & "中含有非数字字符。", vbExclamation, gstrSysName
                txtEdit(lngCount).SetFocus
                zlControl.TxtSelAll txtEdit(lngCount)
                Exit Function
            End If
        Next
        If Len(txtEdit(lngCount).Text) <> Len(txtEdit(lngCount).Tag) - Len(mstr前缀) Then
            MsgBox strName & "的长度不对。", vbExclamation, gstrSysName
            txtEdit(lngCount).SetFocus
            zlControl.TxtSelAll txtEdit(lngCount)
            Exit Function
        End If
    Next
    
    If mstr前缀 & txtEdit(1).Text < txtEdit(1).Tag Then
        MsgBox "作废的开始" & strName & "必须大于领用的开始" & strName & "。", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        zlControl.TxtSelAll txtEdit(1)
        Exit Function
    End If
    If txtEdit(2).Enabled = True Then
        If mstr前缀 & txtEdit(2).Text > txtEdit(2).Tag Then
            MsgBox "作废的终止" & strName & "必须小于领用的终止" & strName & "。", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            zlControl.TxtSelAll txtEdit(2)
            Exit Function
        End If
    Else
        If mstr前缀 & txtEdit(1).Text > txtEdit(2).Tag Then
            MsgBox "作废的" & strName & "必须小于领用的终止" & strName & "。", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            zlControl.TxtSelAll txtEdit(1)
            Exit Function
        End If
    End If
        
    If txtEdit(1).Text > txtEdit(2).Text Then
        MsgBox "作废的开始" & strName & "必须小于作废的终止" & strName & "。", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        zlControl.TxtSelAll txtEdit(1)
        Exit Function
    End If
    If Val(txtEdit(2).Text) - Val(txtEdit(1).Text) + 1 > 10000 Then
        MsgBox "一次作废的总张数不能超过一万张。", vbExclamation, gstrSysName
        txtEdit(2).SetFocus
        zlControl.TxtSelAll txtEdit(2)
        Exit Function
    End If
    
    If mstr最小号码 <> "" Then
        '检查是否报损或使用
        If SplitCardNos(mstr前缀 & txtEdit(1).Text & "～" & mstr前缀 & txtEdit(2).Text, strNOs) = False Then Exit Function
        varPara = Array(mstrID)
        If FromStringListBulidSQL(0, strNOs, varPara, strTable, strName, 2) = False Then Exit Function
        If mint票种 = gBillType.消费卡 Then
            strSQL = _
                "Select Distinct a.卡号 As 号码" & vbNewLine & _
                "From 消费卡使用记录 A, (" & strTable & ") B" & vbNewLine & _
                "Where a.卡号 = b.卡号 And a.领用id = [1]"
        Else
            strSQL = _
                "Select Distinct a.号码" & vbNewLine & _
                "From 票据使用明细 A, (" & strTable & ") B" & vbNewLine & _
                "Where a.号码 = b.号码 And a.领用id = [1]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, varPara)
        strTemp = ""
        Do While Not rsTemp.EOF
            strTemp = strTemp & "," & Nvl(rsTemp!号码)
            rsTemp.MoveNext
        Loop
        If strTemp <> "" Then
            strTemp = Mid(strTemp, 2)
            ShowMsgbox "以下" & strName & "已经使用或已被作废，不能再作废：" & vbCrLf & strTemp
            zlControl.ControlSetFocus txtEdit(1)
            zlControl.TxtSelAll txtEdit(1)
            Exit Function
        End If
    End If
    If cmb报损人.Text = "" Then
        MsgBox "报损人不能为空。", vbExclamation, gstrSysName
        cmb报损人.SetFocus
        Exit Function
    End If
    
    ValidateContent = True
End Function

Private Function Save() As Boolean
'功能:保存编辑的内容
'参数:
'返回值:成功返回True,否则为False
    Dim strTemp As String
    Dim lngID As Long
    Dim strSQL As String
    
    On Error GoTo errHandle
    If mint票种 = gBillType.消费卡 Then
        'Zl_消费卡使用记录_Damage
        strSQL = "Zl_消费卡使用记录_Damage("
        '  领用id_In   In 消费卡使用记录.领用id%Type,
        strSQL = strSQL & "" & mstrID & ","
        '  前缀_In     In 票据领用记录.前缀文本%Type,
        strSQL = strSQL & "'" & mstr前缀 & "',"
        '  开始卡号_In In 消费卡使用记录.卡号%Type,
        strSQL = strSQL & "'" & txtEdit(1).Text & "',"
        '  结束卡号_In In 消费卡使用记录.卡号%Type,
        strSQL = strSQL & "'" & txtEdit(2).Text & "',"
        '  使用时间_In In 消费卡使用记录.使用时间%Type := Null,
        strSQL = strSQL & "To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
        '  使用人_In   In 消费卡使用记录.使用人%Type := Null
        strSQL = strSQL & "'" & cmb报损人.Text & "')"
    Else
        'Zl_票据使用明细_Damage
        strSQL = "Zl_票据使用明细_Damage("
        '  领用id_In   In 票据使用明细.领用id%Type,
        strSQL = strSQL & "" & mstrID & ","
        '  票种_In     In 票据使用明细.票种%Type,
        strSQL = strSQL & "" & mint票种 & ","
        '  前缀_In     In 票据领用记录.前缀文本%Type,
        strSQL = strSQL & "'" & mstr前缀 & "',"
        '  开始号码_In In 票据使用明细.号码%Type,
        strSQL = strSQL & "'" & txtEdit(1).Text & "',"
        '  结束号码_In In 票据使用明细.号码%Type,
        strSQL = strSQL & "'" & txtEdit(2).Text & "',"
        '  使用时间_In In 票据使用明细.使用时间%Type := Null,
        strSQL = strSQL & "To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
        '  使用人_In   In 票据使用明细.使用人%Type := Null
        strSQL = strSQL & "'" & cmb报损人.Text & "')"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If gblnBillPrint Then
        Call gobjBillPrint.zlDiscardBill(mstrID, Val(txtEdit(0).Tag), _
            mstr前缀, txtEdit(1).Text, txtEdit(2).Text, dtpDate.Value, cmb报损人.Text)
    End If
    
    Save = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowSum()
'功能:显示汇总信息
    Dim strTemp As String
    Dim strName1 As String, strName2 As String
    
    strName1 = IIf(mblnIsBIll, "号码", "卡号")
    strName2 = IIf(mblnIsBIll, "票据", "卡片")
    
    '作废的开始号码:
    '作废的结束号码:
    '作废的票据总张数:
    '
    '领用的开始号码:
    '领用的结束号码:
    '已经使用的最小号码:
    '已经使用的最大号码:
    
    strTemp = " 作废的开始" & strName1 & "：" & lbl(1).Caption & txtEdit(1).Text & vbCrLf
    strTemp = strTemp & "  作废的结束" & strName1 & "：" & lbl(2).Caption & txtEdit(2).Text & vbCrLf
    If txtEdit(1).Text = "" Or txtEdit(2).Text = "" Then
        strTemp = strTemp & "  作废的" & strName2 & "总张数：" & vbCrLf & vbCrLf
    Else
        strTemp = strTemp & "  作废的" & strName2 & "总张数：" & Val(txtEdit(2).Text) - Val(txtEdit(1).Text) + 1 & vbCrLf & vbCrLf
    End If
    strTemp = strTemp & "  领用的开始" & strName1 & "：" & Replace(txtEdit(1).Tag, "&", "&&") & vbCrLf
    strTemp = strTemp & "  领用的结束" & strName1 & "：" & Replace(txtEdit(2).Tag, "&", "&&") & vbCrLf
    If mstr最小号码 <> "" Then
        strTemp = strTemp & "  已经使用的最小" & strName1 & "：" & Replace(mstr最小号码, "&", "&&") & vbCrLf
        strTemp = strTemp & "  已经使用的最大" & strName1 & "：" & Replace(mstr最大号码, "&", "&&") & vbCrLf
    End If
    
    lbl说明.Caption = strTemp
End Sub

Public Function 编辑票据报损(frmParent As Object, ByVal strPrivs As String, _
    ByVal int票种 As gBillType, ByVal strID As String) As Boolean
    '功能:用来与调用的财务监控窗口进行通讯的程序,用来增加缴款记录
    '参数:
    '返回值:编辑成功返回True,否则为False
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
        
    mstrPrivs = strPrivs
    mint票种 = int票种: mstrID = strID
    
    mblnIsBIll = CurrentIsBill(int票种)
    Call InitContext
    
    If mint票种 = gBillType.消费卡 Then
        strSQL = _
            "Select 领用人,前缀文本,开始卡号 As 开始号码,终止卡号 As 终止号码,当前卡号 As 当前号码,使用方式" & vbNewLine & _
            "From 消费卡领用记录" & vbNewLine & _
            "Where ID=[1]"
    Else
        strSQL = _
            "Select 领用人,前缀文本,开始号码,终止号码,当前号码,使用方式" & vbNewLine & _
            "From 票据领用记录" & vbNewLine & _
            "Where ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrID)
    
    mstr前缀 = Nvl(rsTemp!前缀文本)
    lbl(1).Caption = Replace(mstr前缀, "&", "&&")
    lbl(2).Caption = lbl(1).Caption
    txtEdit(1).Tag = Nvl(rsTemp!开始号码)
    txtEdit(2).Text = Mid(Nvl(rsTemp!终止号码), Len(mstr前缀) + 1)
    mlng票据长度 = Len(Mid(Nvl(rsTemp!终止号码), Len(mstr前缀) + 1))
    txtEdit(2).Tag = Nvl(rsTemp!终止号码)
    If IsNull(rsTemp!当前号码) Then
        txtEdit(1).Text = Mid(Nvl(rsTemp!开始号码), Len(mstr前缀) + 1)
    Else
        '已经使用，就把最大值加一
        txtEdit(1).Text = Mid(zlStr.Increase(Nvl(rsTemp!当前号码)), Len(mstr前缀) + 1)
    End If
    
    On Error Resume Next
    If Val(rsTemp!使用方式) = 2 Then    '共享方式下,只能选择为本操作员:35846
        cmb报损人.Text = UserInfo.姓名
    Else
        cmb报损人.Text = Nvl(rsTemp!领用人)
    End If
    If Err <> 0 Then
        If Val(rsTemp!使用方式) = 2 Then
            cmb报损人.AddItem UserInfo.姓名
            cmb报损人.ListIndex = cmb报损人.NewIndex
        Else
            cmb报损人.AddItem Nvl(rsTemp!领用人)
            cmb报损人.ListIndex = cmb报损人.NewIndex
        End If
    End If
    If InStr(mstrPrivs, "所有操作员") = 0 Then cmb报损人.Enabled = False
    On Error GoTo errHandle
    
    If mint票种 = gBillType.消费卡 Then
        strSQL = _
            "Select Nvl(Min(卡号), ' ') As 最小号码, Nvl(Max(卡号), ' ') As 最大号码" & vbNewLine & _
            "From 消费卡使用记录" & vbNewLine & _
            "Where 领用id =[1]"
    Else
        strSQL = _
            "Select Nvl(Min(号码), ' ') As 最小号码, Nvl(Max(号码), ' ') As 最大号码" & vbNewLine & _
            "From 票据使用明细" & vbNewLine & _
            "Where 领用id =[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrID)
    
    mstr最小号码 = Trim(Nvl(rsTemp!最小号码))
    mstr最大号码 = Trim(Nvl(rsTemp!最大号码))
    Call opt范围_Click(0)
    
    mblnOK = False
    mblnChange = False
    frmBillDiscard.Show vbModal, frmParent
    编辑票据报损 = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SplitCardNos(ByVal strCardNoRange As String, ByRef strCardNos As String) As Boolean
    '功能:根据传入的卡号范围，分解成相关的卡号
    '入参:
    '   strCardNoRange-卡号范围
    '出参:
    '   strCardNos-返回卡号数(用逗号分离)
    '返回:分解成功返回True，否则返回False
    Dim varData As Variant, lngCount As Long
    Dim strCardStartNO As String, strCardEndNO As String, strCurNo As String
    Dim str数量 As String

    varData = Split(strCardNoRange & "～", "～")
    strCardStartNO = varData(0): strCardEndNO = varData(1)
    If strCardEndNO = "" Then
        strCardNos = strCardStartNO
        SplitCardNos = True
        Exit Function
    End If
    If strCardStartNO > strCardEndNO Then Exit Function
    
    str数量 = zlStr.ExpressValue(strCardEndNO & "-" & strCardStartNO & "+1")
    If InStr(UCase(str数量), "E") > 0 Or Len(str数量) > 4 Then '数量太大已经变成科学计算法
        ShowMsgbox "卡号范围不能大于10000，请分段作废！"
        Exit Function
    End If
    
    strCurNo = strCardStartNO
    strCardNos = strCardStartNO
    Do While True
        If strCurNo >= strCardEndNO Then Exit Do
        strCurNo = zlStr.Increase(strCurNo)
        strCardNos = strCardNos & "," & strCurNo
        
        lngCount = lngCount + 1
        If lngCount > 10000 Then
            ShowMsgbox "卡号范围不能大于10000，请分段作废！"
            Exit Function
        End If
    Loop
    SplitCardNos = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FromStringListBulidSQL(ByVal bytBulidType As Byte, ByVal strValues As String, _
    ByRef varPara As Variant, ByRef strBulitSQL As String, _
    ByVal strColumnAliaName As String, Optional intStartPara As Integer = 1) As Boolean
    '功能:将参数值(值列表组成的)超长的参数分解为含有多个参数的SQL,如:select ... From str2List Union ALL Selelct ..
    '入参:strValues-值,多个用逗号分离
    '     strColumnAliaName-列别名
    '     bytType-0-字符型;1-数字型;
    '     intStartPara-启动的参数序号
    '出参:varPara-返回的参数值数据组
    '     strBulitSQL-返回的构建的SQL串
    '返回:如果获取成功,返回true,否则返回False
    Dim varData As Variant, strTemp As String
    Dim i As Long, j As Long, strSQL As String
    Dim strTable As String, strColumnName As String
    
    On Error GoTo ErrHandler
    strColumnName = " a.Column_Value "
    If strColumnAliaName <> "" Then strColumnName = strColumnName & " As " & strColumnAliaName
    
    If bytBulidType = 0 Then
        strTable = "Table(f_str2list([0]))"
    Else
        strTable = "Table(f_Num2list([0]))"
    End If
    
    j = intStartPara
    ReDim Preserve varPara(0 To j - 1)
    
    varData = Split(strValues, ",")
    strTemp = ""
    For i = 0 To UBound(varData)
        If zlCommFun.ActualLen(strTemp & "," & varData(i)) > 4000 Then
            strSQL = strSQL & " Union ALL " & _
                " Select /*+cardinality(a,10) */" & strColumnName & _
                " From " & Replace(strTable, "[0]", "[" & j & "]") & " A"
            ReDim Preserve varPara(0 To j - 1)
            varPara(j - 1) = Mid(strTemp, 2)
            j = j + 1: strTemp = ""
        End If
        strTemp = strTemp & "," & varData(i)
    Next
    If strTemp <> "" Then
        strSQL = strSQL & " Union ALL " & _
            " Select /*+cardinality(a,10) */" & strColumnName & _
            " From " & Replace(strTable, "[0]", "[" & j & "]") & " A"
        ReDim Preserve varPara(0 To j - 1)
        varPara(j - 1) = Mid(strTemp, 2)
    End If
    
    If strSQL <> "" Then strSQL = Mid(strSQL, 11)
    strBulitSQL = strSQL
    FromStringListBulidSQL = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

