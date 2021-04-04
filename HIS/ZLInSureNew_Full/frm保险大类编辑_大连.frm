VERSION 5.00
Begin VB.Form frm保险大类编辑_大连 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险大类编辑"
   ClientHeight    =   4980
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   4530
   Icon            =   "frm保险大类编辑_大连.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cmb服务对象 
      Height          =   300
      ItemData        =   "frm保险大类编辑_大连.frx":000C
      Left            =   1170
      List            =   "frm保险大类编辑_大连.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1335
      Width           =   1425
   End
   Begin VB.CheckBox chk医保 
      Caption         =   "医保项目(&I)"
      Height          =   225
      Left            =   1170
      TabIndex        =   8
      Top             =   1770
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   315
      TabIndex        =   18
      Top             =   4470
      Width           =   1100
   End
   Begin VB.Frame frmRule 
      Caption         =   "统筹支付计算规则"
      Height          =   1500
      Left            =   285
      TabIndex        =   13
      Top             =   2820
      Width           =   4080
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   5
         Left            =   1860
         MaxLength       =   16
         TabIndex        =   22
         Top             =   990
         Width           =   1320
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   1845
         MaxLength       =   16
         TabIndex        =   19
         Top             =   630
         Width           =   1320
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   1845
         MaxLength       =   16
         TabIndex        =   15
         Top             =   285
         Width           =   1320
      End
      Begin VB.CheckBox chk定额 
         Caption         =   "定额报销                元"
         Height          =   330
         Left            =   810
         TabIndex        =   23
         Top             =   990
         Width           =   2775
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "定额报销               元"
         Height          =   180
         Left            =   1125
         TabIndex        =   21
         Top             =   1050
         Width           =   2250
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "住院统筹支付比例               %"
         Height          =   180
         Index           =   4
         Left            =   390
         TabIndex        =   20
         Top             =   690
         Width           =   2880
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "门诊统筹支付比例               %"
         Height          =   180
         Index           =   3
         Left            =   390
         TabIndex        =   14
         Top             =   345
         Width           =   2880
      End
   End
   Begin VB.Frame fraKind 
      Caption         =   "性质"
      Height          =   630
      Left            =   285
      TabIndex        =   9
      Top             =   2070
      Width           =   4095
      Begin VB.OptionButton opt性质 
         Caption         =   "服务(&F)"
         Height          =   180
         Index           =   3
         Left            =   2640
         TabIndex        =   12
         Top             =   315
         Width           =   945
      End
      Begin VB.OptionButton opt性质 
         Caption         =   "医疗(&D)"
         Height          =   180
         Index           =   2
         Left            =   1425
         TabIndex        =   11
         Top             =   315
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
      Left            =   2130
      TabIndex        =   16
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   17
      Top             =   4470
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
      Width           =   3195
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
Attribute VB_Name = "frm保险大类编辑_大连"
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
    Text住院 = 4
    Text定额 = 5
    
    Check药品 = 1
    Check医疗 = 2
    Check服务 = 3
    
    Check比例 = 1
    Check住院日 = 2
End Enum

Dim mlng险类 As Long
Dim mstrID As String         '当前编辑的医保大类ID
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Sub Check1_Click()

End Sub

Private Sub chk定额_Click()
    txtEdit(Text定额).Enabled = chk定额.Value = 1
    txtEdit(Text比例).Enabled = chk定额.Value <> 1
    txtEdit(Text住院).Enabled = chk定额.Value <> 1
End Sub

Private Sub chk医保_Click()
    mblnChange = True
End Sub

Private Sub chk医保_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmb服务对象_Click()
    mblnChange = True
    Select Case cmb服务对象.ListIndex
    Case 0  '门诊
        lblEdit(3).Enabled = True
        txtEdit(Text比例).Enabled = True
        lblEdit(4).Enabled = False
        txtEdit(Text住院).Enabled = False
    Case 1  '住院
        lblEdit(3).Enabled = False
        txtEdit(Text比例).Enabled = False
        lblEdit(4).Enabled = True
        txtEdit(Text住院).Enabled = True
    Case 2  '所有
        lblEdit(3).Enabled = True
        txtEdit(Text比例).Enabled = True
        lblEdit(4).Enabled = True
        txtEdit(Text住院).Enabled = True
    End Select
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
        For lngIndex = Text名称 To Text住院
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
    Dim dbl统筹比额 As Double, dbl特准定额 As Double, dbl特准天数 As Double, dbl住院比额 As Double
    Dim dbl门诊比额 As Double
    Dim lngIndex As Long, lst As ListItem
    
    On Error GoTo errHandle
    
    For lngIndex = 1 To 3
        If opt性质(lngIndex).Value = True Then
            lng性质 = lngIndex
            Exit For
        End If
    Next
    dbl统筹比额 = 0
    Select Case Me.cmb服务对象.ListIndex
    Case 0
        dbl统筹比额 = Val(txtEdit(Text比例).Text)
        dbl门诊比额 = Val(txtEdit(Text比例).Text)
        dbl住院比额 = 0
    Case 1
        dbl门诊比额 = 0
        dbl住院比额 = Val(txtEdit(Text住院).Text)
    Case Else
        dbl统筹比额 = Val(txtEdit(Text比例).Text)
        dbl门诊比额 = Val(txtEdit(Text比例).Text)
        dbl住院比额 = Val(txtEdit(Text住院).Text)
    End Select
    
    dbl特准定额 = Val(txtEdit(Text定额).Text)
    dbl特准天数 = 0
    If chk定额.Value = 1 Then
        lng算法 = 2
    Else
        lng算法 = 1
    End If
    
    'zl_保险支付大类_UPDATE (
    '   ID_IN,编码_IN,名称_IN,简码_IN,性质_IN,算法_IN,门诊比额_IN,住院比额_IN,特准定额_IN,特准天数_IN,
    '   服务对象_IN,是否医保_IN
    
    If mstrID = "" Then
        '新增
        lngID = zlDatabase.GetNextID("保险支付大类")
        gstrSQL = "zl_保险支付大类_INSERT(" & lngID & "," & mlng险类 & ",'" & Trim(txtEdit(text编码).Text) & "','" & _
                Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng性质 & "," & lng算法 & "," & _
                 dbl门诊比额 & "," & dbl住院比额 & "," & dbl特准定额 & "," & dbl特准天数 & "," & GetTextFromCombo(cmb服务对象, False) & "," & chk医保.Value & ")"
    Else
        gstrSQL = "zl_保险支付大类_Update(" & mstrID & ",'" & Trim(txtEdit(text编码).Text) & "','" & _
                Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng性质 & "," & lng算法 & "," & _
                  dbl门诊比额 & "," & dbl住院比额 & "," & dbl特准定额 & "," & dbl特准天数 & "," & GetTextFromCombo(cmb服务对象, False) & "," & chk医保.Value & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '更新主界面
    If mstrID = "" Then
        Set lst = frm保险大类_大连.lvwItem.ListItems.Add(, "K" & lngID, txtEdit(text编码), "Class", "Class")
    Else
        Set lst = frm保险大类_大连.lvwItem.SelectedItem
        lst.Text = Trim(txtEdit(text编码).Text)
    End If
    lst.SubItems(1) = Trim(txtEdit(Text名称).Text)
    lst.SubItems(2) = Trim(txtEdit(Text简码).Text)
    lst.SubItems(3) = Choose(lng性质, "药品", "医疗", "服务")
    lst.SubItems(4) = "总额比例"
    lst.SubItems(5) = Mid(cmb服务对象.Text, 3)
    lst.SubItems(6) = IIf(chk医保.Value = 1, "是", "否")
    lst.Tag = dbl统筹比额 & ";" & dbl住院比额
    
    Save项目 = True
    mblnOK = True
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
'功能:分析输入有关医保类别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim lngIndex As Integer
    For lngIndex = text编码 To Text住院
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
            
            If lngIndex >= Text比例 And chk定额.Value <> 1 Then
                
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
                
                If lngIndex = Text比例 Or lngIndex = Text住院 Then
                    If Val(txtEdit(lngIndex).Text) > 100 Then
                        MsgBox IIf(lngIndex = Text比例, "门诊统筹", "住院统筹") & "支付比例不能超过100。", vbInformation, gstrSysName
                        zlControl.TxtSelAll txtEdit(lngIndex)
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
        If lngIndex = Text定额 And chk定额.Value = 1 Then
            
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
            
            If Val(txtEdit(lngIndex).Text) > 100000 Then
                MsgBox "数值不能超过100000。", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
    
    Next
    
    '对特准天数和特准金额的限制
    If chk医保.Value = 0 Then
        If MsgBox("如果将大类设作非医保，会影响到属于它的所有医保项目。" & vbCrLf & "是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            chk医保.SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function


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
        ElseIf Index = Text比例 Or Index = Text住院 Or Index = Text定额 Then
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m金额式
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
    If Index >= Text比例 And Index <= Text住院 Then
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
                  ",统筹比额,住院比额,特准定额,特准天数,是否医保,nvl(服务对象,3) as 服务对象 " & _
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
        '周海全调试 2003-12-17
        '修改时重置性质内容
        opt性质(rsTemp("性质")).Value = True
        chk医保.Value = IIf(rsTemp("是否医保") = 1, 1, 0)
        '1-比例计算项目
        txtEdit(Text比例).Text = Format(rsTemp("统筹比额"), "0.00")
        txtEdit(Text住院).Text = Format(rsTemp("住院比额"), "0.00")
        '1-定准定额
        txtEdit(Text定额).Text = Format(rsTemp("特准定额"), "###0.00;-####0.00; ;")
        If Val(txtEdit(Text定额)) <> 0 Then
            chk定额.Value = 1
        Else
            chk定额.Value = 0
        End If
        chk定额_Click
    Else
        '新增医保大类
        txtEdit(text编码).Text = zlDatabase.GetMax("保险支付大类", "编码", 6, " where 险类=" & mlng险类)
    End If
    
    
    mblnChange = False
    frm保险大类编辑_大连.Show vbModal, frm保险大类
    编辑医保大类 = mblnOK
End Function



