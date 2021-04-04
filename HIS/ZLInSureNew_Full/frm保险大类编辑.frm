VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frm保险大类编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险大类编辑"
   ClientHeight    =   6345
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   7020
   Icon            =   "frm保险大类编辑.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6735
      Left            =   5535
      TabIndex        =   28
      Top             =   -300
      Width           =   30
   End
   Begin VB.ComboBox cmb服务对象 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1335
      Width           =   1425
   End
   Begin VB.CheckBox chk医保 
      Caption         =   "医保项目(&I)"
      Height          =   225
      Left            =   3990
      TabIndex        =   8
      Top             =   1365
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5865
      TabIndex        =   27
      Top             =   5610
      Width           =   1100
   End
   Begin VB.Frame frmRule 
      Caption         =   "统筹支付计算规则"
      Height          =   3750
      Left            =   180
      TabIndex        =   13
      Top             =   2445
      Width           =   5130
      Begin ZL9BillEdit.BillEdit mshbill 
         Height          =   1695
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2990
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.OptionButton opt算法 
         Caption         =   "分类费用档次计算法(&T)"
         Height          =   240
         Index           =   3
         Left            =   225
         TabIndex        =   24
         Top             =   1560
         Width           =   2265
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   3705
         MaxLength       =   16
         TabIndex        =   23
         Top             =   1200
         Width           =   630
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   21
         Top             =   1200
         Width           =   630
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   19
         Top             =   870
         Width           =   630
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   3735
         MaxLength       =   16
         TabIndex        =   16
         Top             =   270
         Width           =   630
      End
      Begin VB.OptionButton opt算法 
         Caption         =   "住院日定额计算法(&Z)"
         Height          =   240
         Index           =   2
         Left            =   210
         TabIndex        =   17
         Top             =   630
         Width           =   2265
      End
      Begin VB.OptionButton opt算法 
         Caption         =   "总额比例计算法(&B)"
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "特准定额允许        天"
         Height          =   180
         Index           =   6
         Left            =   2595
         TabIndex        =   22
         Top             =   1260
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "每日特准定额        元"
         Height          =   180
         Index           =   5
         Left            =   465
         TabIndex        =   20
         Top             =   1260
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "每日基本定额        元"
         Height          =   180
         Index           =   4
         Left            =   465
         TabIndex        =   18
         Top             =   930
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "统筹支付比例        %"
         Height          =   180
         Index           =   3
         Left            =   2625
         TabIndex        =   15
         Top             =   330
         Width           =   1890
      End
   End
   Begin VB.Frame fraKind 
      Caption         =   "性质"
      Height          =   630
      Left            =   195
      TabIndex        =   9
      Top             =   1710
      Width           =   5160
      Begin VB.OptionButton opt性质 
         Caption         =   "服务(&W)"
         Height          =   180
         Index           =   3
         Left            =   4050
         TabIndex        =   12
         Top             =   270
         Width           =   945
      End
      Begin VB.OptionButton opt性质 
         Caption         =   "医疗(&D)"
         Height          =   180
         Index           =   2
         Left            =   2145
         TabIndex        =   11
         Top             =   285
         Width           =   945
      End
      Begin VB.OptionButton opt性质 
         Caption         =   "药品(&M)"
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5775
      TabIndex        =   25
      Top             =   225
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5775
      TabIndex        =   26
      Top             =   735
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   5
      Top             =   937
      Width           =   1425
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1170
      MaxLength       =   40
      TabIndex        =   3
      Top             =   536
      Width           =   4080
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   1
      Top             =   135
      Width           =   1425
   End
   Begin VB.Label lbl服务 
      AutoSize        =   -1  'True
      Caption         =   "服务对象(&F)"
      Height          =   180
      Left            =   135
      TabIndex        =   6
      Top             =   1398
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "编码(&U)"
      Height          =   180
      Index           =   0
      Left            =   495
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "简码(&S)"
      Height          =   180
      Index           =   2
      Left            =   495
      TabIndex        =   4
      Top             =   997
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Index           =   1
      Left            =   495
      TabIndex        =   2
      Top             =   596
      Width           =   630
   End
End
Attribute VB_Name = "frm保险大类编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum编辑
    text编码 = 0
    Text名称 = 1
    Text简码 = 2
    Text比例 = 3
    Text基本 = 4
    Text特准 = 5
    Text天数 = 6

    Check药品 = 1
    Check医疗 = 2
    Check服务 = 3
    
    Check比例 = 1
    Check住院日 = 2
    chk费用档次 = 3
End Enum
Private Enum mColHead
    档次 = 0
    上限
    下限
    比例
End Enum

Dim mlng险类 As Long
Dim mstrID As String         '当前编辑的医保大类ID
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Sub chk医保_Click()
    mblnChange = True
End Sub

Private Sub chk医保_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmb服务对象_Click()
    mblnChange = True
End Sub

Private Sub cmb服务对象_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    If IsValid() = False Then Exit Sub
    If Save项目() = False Then Exit Sub
    
    If mstrID = "" Then
        '连续新增
        txtEdit(text编码).Text = zlDatabase.GetMax("保险支付大类", "编码", 6, " where 险类=" & mlng险类)
        For lngIndex = Text名称 To Text天数
            txtEdit(lngIndex).Text = ""
        Next
        chk医保.Value = 1
        mblnChange = False
        txtEdit(text编码).SetFocus
    Else
        mblnChange = False
        Unload Me
    End If
End Sub

Private Function Save项目() As Boolean
    Dim lngID As Long, lng性质 As Long, lng算法 As Long
    Dim dbl统筹比额 As Double, dbl特准定额 As Double, dbl特准天数 As Double
    Dim lngIndex As Long, lst As ListItem
    
    On Error GoTo errHandle
    
    For lngIndex = 1 To 3
        If opt性质(lngIndex).Value = True Then
            lng性质 = lngIndex
            Exit For
        End If
    Next
    If opt算法(1).Value = True Then
        '按比例
        lng算法 = 1
        dbl统筹比额 = Val(txtEdit(Text比例).Text)
        
    Else
        If opt算法(3).Value = True Then
            lng算法 = 3
        Else
            '按住院日
            lng算法 = 2
            dbl统筹比额 = Val(txtEdit(Text基本).Text)
            dbl特准定额 = Val(txtEdit(Text特准).Text)
            dbl特准天数 = Val(txtEdit(Text天数).Text)
        End If
    End If
    
    
    If mstrID = "" Then
        '新增
        lngID = zlDatabase.GetNextID("保险支付大类")
        gstrSQL = "zl_保险支付大类_INSERT(" & lngID & "," & mlng险类 & ",'" & Trim(txtEdit(text编码).Text) & "','" & _
                Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng性质 & "," & lng算法 & "," & _
                 dbl统筹比额 & "," & dbl特准定额 & "," & dbl特准天数 & "," & GetTextFromCombo(cmb服务对象, False) & "," & chk医保.Value & ")"
    Else
        gstrSQL = "zl_保险支付大类_Update(" & mstrID & ",'" & Trim(txtEdit(text编码).Text) & "','" & _
                Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng性质 & "," & lng算法 & "," & _
                 dbl统筹比额 & "," & dbl特准定额 & "," & dbl特准天数 & "," & GetTextFromCombo(cmb服务对象, False) & "," & chk医保.Value & ")"
    End If
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If lng算法 = 3 Then
        If SaveGrdData(IIf(mstrID = "", lngID, Val(mstrID))) = False Then GoTo errHandle:
    End If
    gcnOracle.CommitTrans
    
    '更新主界面
    If mstrID = "" Then
        Set lst = frm保险大类.lvwItem.ListItems.Add(, "K" & lngID, txtEdit(text编码), "Class", "Class")
    Else
        Set lst = frm保险大类.lvwItem.SelectedItem
        lst.Text = Trim(txtEdit(text编码).Text)
    End If
    lst.SubItems(1) = Trim(txtEdit(Text名称).Text)
    lst.SubItems(2) = Trim(txtEdit(Text简码).Text)
    lst.SubItems(3) = Choose(lng性质, "药品", "医疗", "服务")
    lst.SubItems(4) = IIf(lng算法 = 1, "总额比例", "住院日核定")
    lst.SubItems(5) = Mid(cmb服务对象.Text, 3)
    lst.SubItems(6) = IIf(chk医保.Value = 1, "是", "否")
    lst.Tag = dbl统筹比额 & ";" & dbl特准定额 & ";" & dbl特准天数
    
    Save项目 = True
    mblnOK = True
    Exit Function

errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
'功能:分析输入有关医保类别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim lngIndex As Integer
    For lngIndex = text编码 To Text天数
        If txtEdit(lngIndex).Enabled = True Then
            If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
                txtEdit(lngIndex).SetFocus
                zlControl.TxtSelAll txtEdit(lngIndex)
                Exit Function
            End If
            
            If lngIndex = text编码 Or lngIndex = Text名称 Then
                If Len(Trim(txtEdit(lngIndex).Text)) = 0 Then
                    txtEdit(lngIndex).Text = ""
                    MsgBox "编码或名称都不能为空。", vbExclamation, gstrSysName
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
            End If
            
            If lngIndex >= Text比例 Then
                If IsNumeric(txtEdit(lngIndex).Text) = False Then
                    MsgBox "请输入合法的数值。", vbInformation, gstrSysName
                    zlControl.TxtSelAll txtEdit(lngIndex)
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
                        
                If Val(txtEdit(lngIndex).Text) < 0 Then
                    MsgBox "数值不能小于0。", vbInformation, gstrSysName
                    zlControl.TxtSelAll txtEdit(lngIndex)
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
                
                If lngIndex = Text比例 Then
                    If Val(txtEdit(Text比例).Text) > 100 Then
                        MsgBox "统筹支付比例不能超过100。", vbInformation, gstrSysName
                        zlControl.TxtSelAll txtEdit(Text比例)
                        txtEdit(lngIndex).SetFocus
                        Exit Function
                    End If
                Else
                    If Val(txtEdit(lngIndex).Text) > 10000 Then
                        MsgBox "数值不能超过10000。", vbInformation, gstrSysName
                        zlControl.TxtSelAll txtEdit(lngIndex)
                        txtEdit(lngIndex).SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    '对特准天数和特准金额的限制
    If opt算法(Check住院日).Value = True Then
        If Val(txtEdit(Text特准).Text) = 0 And Val(txtEdit(Text天数).Text) <> 0 Then
            MsgBox "特准定额为0，特准天数也须为0。", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text天数)
            txtEdit(Text天数).SetFocus
            Exit Function
        End If
        If Val(txtEdit(Text特准).Text) <> 0 And Val(txtEdit(Text天数).Text) = 0 Then
            MsgBox "特准天数为0，特准定额也须为0。", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text特准)
            txtEdit(Text特准).SetFocus
            Exit Function
        End If
        If Val(txtEdit(Text特准).Text) <> 0 And Val(txtEdit(Text基本).Text) > Val(txtEdit(Text特准).Text) Then
            MsgBox "基本定额不能大于特准定额。", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text基本)
            txtEdit(Text基本).SetFocus
            Exit Function
        End If
    End If
    Dim i As Long
    If opt算法(chk费用档次).Value = True Then
        With mshBill
            For i = 1 To .Rows - 1
                If i <> 1 Then
                    If Val(.TextMatrix(i - 1, mColHead.下限)) <> Val(.TextMatrix(i, mColHead.上限)) Then
                        MsgBox "在第" & i & "行的上限不等于在第" & i - 1 & "行的下限,请重输!", vbInformation + vbDefaultButton1, gstrSysName
                        mshBill.Row = i
                        mshBill.SetFocus
                        Exit Function
                    End If
               End If
                If (Val(.TextMatrix(i, mColHead.下限)) <> 0 Or Val(.TextMatrix(i, mColHead.上限)) <> 0) _
                    And Val(.TextMatrix(i, mColHead.比例)) = 0 Then
                    MsgBox "在第" & i & "行的比例等于零了,请重输!", vbInformation + vbDefaultButton1, gstrSysName
                    mshBill.Row = i
                    mshBill.SetFocus
                    Exit Function
                End If
                If Val(.TextMatrix(i, mColHead.下限)) = Val(.TextMatrix(i, mColHead.上限)) And Val(.TextMatrix(i, mColHead.上限)) <> 0 Then
                    MsgBox "在第" & i & "行的上限等于下限,请重输!", vbInformation + vbDefaultButton1, gstrSysName
                    mshBill.Row = i
                    mshBill.SetFocus
                    Exit Function
                End If
            Next
        End With
    End If
    If chk医保.Value = 0 Then
        If MsgBox("如果将大类设作非医保，会影响到属于它的所有医保项目。" & vbCrLf & "是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            chk医保.SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function


Private Sub MshBill_AfterAddRow(Row As Long)
        If Row = mshBill.Rows - 1 Then
            mshBill.TextMatrix(Row, mColHead.上限) = mshBill.TextMatrix(Row - 1, mColHead.下限)
        End If
End Sub

Private Sub MshBill_AfterDeleteRow()
   If mshBill.Row = mshBill.Rows - 1 Then Exit Sub
    Call ReSet上限(mshBill.Row - 1)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
''   If Row = MshBill.Rows - 1 Then Exit Sub
''    Call ReSet上限(Row)
End Sub

Private Sub opt算法_Click(Index As Integer)
    Dim bln比例 As Boolean
    
    mblnChange = True
    txtEdit(Text比例).Enabled = (opt算法(Check比例).Value = True)
    lblEdit(Text比例).Enabled = txtEdit(Text比例).Enabled
    
    txtEdit(Text基本).Enabled = (opt算法(Check住院日).Value = True)
    txtEdit(Text特准).Enabled = txtEdit(Text基本).Enabled
    txtEdit(Text天数).Enabled = txtEdit(Text基本).Enabled
    lblEdit(Text基本).Enabled = txtEdit(Text基本).Enabled
    lblEdit(Text特准).Enabled = txtEdit(Text基本).Enabled
    lblEdit(Text天数).Enabled = txtEdit(Text基本).Enabled
    
    mshBill.Active = (opt算法(chk费用档次).Value = True)
End Sub

Private Sub opt算法_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub opt性质_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt性质_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text名称 Then
        txtEdit(Text简码).Text = zlCommFun.SpellCode(txtEdit(Text名称).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text名称
          zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 '使之不响
        zlCommFun.PressKey (vbKeyTab)
    Else
        If Index = text编码 Then
            KeyAscii = asc(UCase(Chr(KeyAscii)))
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
    If Index >= Text比例 And Index <= Text特准 Then
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
    End If
End Sub

Public Function 编辑医保大类(ByVal lng险类 As Long, ByVal strID As String) As Boolean
'功能:用来与调用的医保类别管理窗口进行通讯的程序
'参数:str序号           当前编辑的医保类别的的序号
'返回值:编辑成功返回True,否则为False
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mblnOK = False
    mlng险类 = lng险类
    mstrID = strID
    
    cmb服务对象.AddItem "1.门诊病人"
    cmb服务对象.AddItem "2.住院病人"
    cmb服务对象.AddItem "3.所有病人"
    cmb服务对象.ListIndex = 2
    rsTemp.CursorLocation = adUseClient
    If mstrID <> "" Then
        '修改医保大类
        gstrSQL = "select 编码,名称,简码,nvl(性质,1) as 性质,nvl(算法,1) as 算法 " & _
                  ",统筹比额,特准定额,特准天数,是否医保,nvl(服务对象,3) as 服务对象 " & _
                  "from 保险支付大类 where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(mstrID))
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "该保险大类已经被删除，请刷新。", vbInformation, gstrSysName
            Exit Function
        End If
        txtEdit(text编码).Text = rsTemp("编码")
        txtEdit(Text名称).Text = rsTemp("名称")
        txtEdit(Text简码).Text = IIf(IsNull(rsTemp("简码")), "", rsTemp("简码"))
        Call SetComboByText(cmb服务对象, rsTemp("服务对象"), False)
        chk医保.Value = IIf(rsTemp("是否医保") = 1, 1, 0)
        opt性质(rsTemp("性质")).Value = True
        opt算法(rsTemp("算法")).Value = True
        Call opt算法_Click(rsTemp("算法"))
        If rsTemp("算法") = 1 Then
            '清除费用档次
            Call initGrd
            '1-比例计算项目
            txtEdit(Text比例).Text = Format(rsTemp("统筹比额"), "0.00")
            opt算法_Click (1)
        ElseIf Nvl(rsTemp!算法, 0) = 2 Then
            '清除费用档次
            Call initGrd
            '2-住院日核定项目
            txtEdit(Text基本).Text = Format(rsTemp("统筹比额"), "0.00")
            txtEdit(Text特准).Text = Format(rsTemp("特准定额"), "0.00")
            txtEdit(Text天数).Text = Format(rsTemp("特准天数"), "0")
            opt算法_Click (2)
        Else '3-费用档次计算法
            Call LoadGrd
            opt算法_Click (3)
        End If
    Else
        '新增医保大类
        txtEdit(text编码).Text = zlDatabase.GetMax("保险支付大类", "编码", 6, " where 险类=" & mlng险类)
        opt算法(1).Value = True
        Call opt算法_Click(1)
        '清除费用档次
        Call initGrd
    End If
    mblnChange = False
    frm保险大类编辑.Show vbModal, frm保险大类
    编辑医保大类 = mblnOK
End Function
Private Sub initGrd()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始设置档次的Grid特性
    '--入参数:
    '--出参数:
    '--返  回:
    '--编  制:刘兴宏:20040615
    '-----------------------------------------------------------------------------------------------------------
   With mshBill
        .Active = True
        .ClearBill
        .Cols = 4
        .Rows = 2
        .msfObj.FixedCols = 1
        .TextMatrix(0, mColHead.档次) = "档次"
        .TextMatrix(0, mColHead.上限) = "上限"
        .TextMatrix(0, mColHead.下限) = "下限"
        .TextMatrix(0, mColHead.比例) = "实收比例"
        
        .ColWidth(mColHead.档次) = 500
        .ColWidth(mColHead.上限) = 1400
        .ColWidth(mColHead.下限) = 1400
        .ColWidth(mColHead.比例) = 1400
                
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(mColHead.档次) = 5
        .ColData(mColHead.上限) = 5
        .ColData(mColHead.下限) = 4
        .ColData(mColHead.比例) = 4

        .ColAlignment(mColHead.档次) = flexAlignCenterCenter
        .ColAlignment(mColHead.上限) = flexAlignCenterCenter
        .ColAlignment(mColHead.下限) = flexAlignCenterCenter
        .ColAlignment(mColHead.比例) = flexAlignCenterCenter
        .PrimaryCol = mColHead.档次
        .LocateCol = mColHead.下限
    End With
End Sub

Private Sub LoadGrd()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    '初始表格
    Call initGrd
    
    If mstrID = "" Then Exit Sub
    '表示修改填充数据
    gstrSQL = "Select * From 大类档次比例 where 大类id=" & Val(mstrID) & " order by  档次"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then Exit Sub
    mshBill.Rows = rsTemp.RecordCount + 1
    lngRow = 1
    With mshBill
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, mColHead.档次) = lngRow
            .TextMatrix(lngRow, mColHead.上限) = Format(Nvl(rsTemp!上限, 0), "####0.00;####0.00; ;")
            .TextMatrix(lngRow, mColHead.下限) = Format(Nvl(rsTemp!下限, 0), "####0.00;####0.00; ;")
            .TextMatrix(lngRow, mColHead.比例) = Format(Nvl(rsTemp!比例, 0), "####0.00;####0.00; ;")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
End Sub
Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EnterCell(Row As Long, COL As Long)
    With mshBill
        Select Case .COL
            Case mColHead.上限
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mColHead.下限
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                If Trim(.TextMatrix(.Row, mColHead.下限)) = "" Then
                    .AllowAddRow = False
                Else
                    .AllowAddRow = True
                End If


            Case mColHead.比例
                .TxtCheck = True
                .MaxLength = 4
                .TextMask = ".1234567890"
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With mshBill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        
        Select Case .COL
            Case mColHead.上限
                If .TextMatrix(.Row, .COL) = "" And strKey = "" And .Row <> 1 And Val(.TextMatrix(.Row, mColHead.下限)) <> 0 Then
                    MsgBox "未输入上限,请重新输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "上限必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 1E+19 Then
                    MsgBox "上限只能在0~9000900090009范围内,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If strKey <> "" Then
                    .Text = Format(strKey, "####0.00;####0.00; ;")
                ElseIf Trim(.TextMatrix(.Row, mColHead.上限)) = "" Then
                    .Text = " "
                    .TextMatrix(.Row, mColHead.上限) = " "
                End If
                
            Case mColHead.下限
                If .TextMatrix(.Row, .COL) = "" And strKey = "" And .Row <> .Rows - 1 Then
                    MsgBox "未输入下限,请重新输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "下限必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 1E+19 Then
                    MsgBox "下限只能在0~9000900090009范围内,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Val(strKey) <= Val(.TextMatrix(.Row, mColHead.上限)) And (.Row <> .Rows - 1) Then
                    MsgBox "下限不能小于等于上限,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If

                If Val(strKey) <= Val(.TextMatrix(.Row, mColHead.上限)) And (.Row = .Rows - 1 And strKey <> "") Then
                    MsgBox "下限不能小于等于上限,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                                
                If strKey <> "" Then
                    .Text = Format(strKey, "####0.00;####0.00; ;")
                    .TextMatrix(.Row, mColHead.下限) = .Text
                    
                    If .Row <> .Rows - 1 Then
                        Call ReSet上限(.Row)
                    End If
                ElseIf Trim(.TextMatrix(.Row, mColHead.下限)) = "" Then
                    .Text = " "
                    .TextMatrix(.Row, mColHead.下限) = " "
                End If
                If strKey = "" Then
                    .AllowAddRow = False
                Else
                    .AllowAddRow = True
                End If

                
            Case mColHead.比例
                If .TextMatrix(.Row, .COL) = "" And strKey = "" And (Val(.TextMatrix(.Row, mColHead.上限)) <> 0 Or Val(.TextMatrix(.Row, mColHead.下限)) <> 0) Then
                    MsgBox "未输入比例,请重新输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "比例必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 100 Then
                    MsgBox "比例只能在0~100范围内,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If (Val(.TextMatrix(.Row, mColHead.上限)) = 0 And Val(.TextMatrix(.Row, mColHead.下限)) = 0) Or .AllowAddRow = False Then
                    .AllowAddRow = False
                Else
                    .AllowAddRow = True
                End If
                If strKey <> "" Then
                    .Text = Format(strKey, "####0.00;####0.00; ;")
                End If
        End Select
        If .TextMatrix(.Row, mColHead.档次) = "" Then
            '需重新确定档次
            Call Set档次
        End If
    End With
  
End Sub
Private Sub ReSet上限(ByVal lngRow As Long)
    Dim i As Long
    Dim dbl差额 As Double
    
    With mshBill
        For i = lngRow + 1 To .Rows - 1
            
            .TextMatrix(i, mColHead.档次) = i
            dbl差额 = Val(.TextMatrix(i, mColHead.下限)) - Val(.TextMatrix(i, mColHead.上限))
            .TextMatrix(i, mColHead.上限) = Format(Val(.TextMatrix(i - 1, mColHead.下限)), "####0.00;####0.00; ;")
            If dbl差额 < 0 Then
            Else
                .TextMatrix(i, mColHead.下限) = Format(Val(.TextMatrix(i, mColHead.上限)) + IIf(dbl差额 < 0, 0, dbl差额), "####0.00;####0.00; ;")
            End If
        Next
'        If .TextMatrix(.Row, .Col) <> .Text Then
'            .Text = .TextMatrix(.Row, .Col)
'        End If
    End With
    
End Sub
Private Sub Set档次()
    Dim lngRow As Long
    For lngRow = 1 To mshBill.Rows - 1
        mshBill.TextMatrix(lngRow, mColHead.档次) = lngRow
    Next
End Sub
Private Function SaveGrdData(ByVal lngID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存费用档次数据
    '--入参数:
    '--出参数:
    '--返  回:成功返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim dbl比例 As Double
    Dim dbl下限 As Double
    Dim dbl上限  As Double
    
    SaveGrdData = False
    Err = 0: On Error GoTo errHand:
    If mstrID <> "" Then
                gstrSQL = "zl_大类档次比例_Delete(" & _
                    lngID & ")"
                 Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    With mshBill
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, mColHead.上限)) <> "" Or Trim(.TextMatrix(lngRow, mColHead.下限)) <> "" Then
                '参数:
                '   大类ID_IN    大类档次比例.大类id%Type,
                '   档次_IN        大类档次比例.档次%Type,
                '   上限_IN        大类档次比例.上限%Type,
                '   下限_IN        大类档次比例.下限%Type,
                '   比例_IN
                dbl上限 = Val(.TextMatrix(lngRow, mColHead.上限))
                dbl下限 = Val(.TextMatrix(lngRow, mColHead.下限))
                dbl比例 = Val(.TextMatrix(lngRow, mColHead.比例))
                
                gstrSQL = "zl_大类档次比例_InSert(" & _
                    lngID & "," & _
                    lngRow & "," & _
                    dbl上限 & "," & _
                    dbl下限 & "," & _
                    dbl比例 & ")"
                
                 Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    SaveGrdData = True
    Exit Function
errHand:
    
End Function


