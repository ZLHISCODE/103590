VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊疗项目分类"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "frmClinicClass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkGoOn 
      Caption         =   "连续增加分类"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   2085
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   300
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3030
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1875
      MaxLength       =   8
      TabIndex        =   12
      Tag             =   "编码"
      Text            =   "000000"
      Top             =   1350
      Width           =   570
   End
   Begin VB.TextBox txtSymbol 
      Height          =   300
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "名称"
      Top             =   2040
      Width           =   1425
   End
   Begin VB.CheckBox chkCodeLen 
      Caption         =   "允许更改编码长度，并按此调整各同级编码(&L)"
      Height          =   285
      Left            =   1020
      TabIndex        =   5
      Top             =   2565
      Width           =   4290
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   -15
      TabIndex        =   15
      Top             =   2865
      Width           =   5745
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   6
      Top             =   3045
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   7
      Top             =   3045
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1740
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "名称"
      Top             =   1680
      Width           =   3195
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&P"
      Height          =   300
      Left            =   4650
      TabIndex        =   13
      Top             =   690
      Width           =   285
   End
   Begin VB.TextBox txtParent 
      Height          =   300
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "(无)"
      Top             =   690
      Width           =   2895
   End
   Begin VB.TextBox txtUpCode 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "编码"
      Text            =   "0000"
      Top             =   1305
      Width           =   1335
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   60
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicClass.frx":000C
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicClass.frx":05A6
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMsg 
      Caption         =   "添加成功。"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      Caption         =   "(提示：按Del清除上级，设置初级分类)"
      Height          =   180
      Left            =   1740
      TabIndex        =   18
      Top             =   1035
      Width           =   3150
   End
   Begin VB.Label lblSymbol 
      AutoSize        =   -1  'True
      Caption         =   "简码(&S)"
      Height          =   180
      Left            =   1020
      TabIndex        =   2
      Top             =   2100
      Width           =   630
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "西成药"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   555
      Width           =   540
   End
   Begin VB.Label lblNote 
      Caption         =   "诊疗项目可根据临床与医技各科诊疗应用操作的特点进行统一分类设置。"
      Height          =   345
      Left            =   990
      TabIndex        =   16
      Top             =   210
      Width           =   4170
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmClinicClass.frx":0B40
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "编码(&D)"
      Height          =   180
      Left            =   1020
      TabIndex        =   11
      Top             =   1365
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Left            =   1020
      TabIndex        =   0
      Top             =   1725
      Width           =   630
   End
   Begin VB.Label lblParent 
      AutoSize        =   -1  'True
      Caption         =   "上级(&U)"
      Height          =   180
      Left            =   1020
      TabIndex        =   9
      Top             =   750
      Width           =   630
   End
End
Attribute VB_Name = "frmClinicClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim intMaxLen As Integer
Dim objNode As Node
Dim mblnCancel As Boolean
Dim mblnChanged As Boolean
Private mblnOK As Boolean
Private mstrID As String, mstr名称 As String, mstr编码 As String
Private mstr所有记录 As String
Private mbln药品目录 As Boolean '诊疗目录进入=true ,药品目录进入=false

Public Function ShowMe(ByVal lngModle As Long, ByVal objForm As Object, ByVal strText As String, ByVal strID As String, ByVal intType As Integer, Optional bln药品目录 As Boolean = False) As Boolean
    mbln药品目录 = bln药品目录
    txtParent.Text = strText
    txtParent.Tag = strID
'    If intType = 1 Then Call zlDefaultCode
    Me.Show lngModle, objForm
    ShowMe = mblnOK
End Function

Private Sub chkCodeLen_Click()
    If Me.chkCodeLen.Value = 1 Then
        Me.txtCode.MaxLength = intMaxLen - Len(Me.txtUpCode.Text)
    Else
        Me.txtCode.MaxLength = Me.txtCode.Tag
        Me.txtCode.Text = Mid(Me.txtCode.Text, 1, Me.txtCode.MaxLength)
    End If
'    Me.txtCode.SetFocus
    If mblnCancel = True Then Exit Sub
    If Me.Visible Then mblnChanged = True
    lblMsg.Visible = False
End Sub

Private Sub chkCodeLen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Dim strTemp As String
    
    strTemp = txtParent.Text & "|" & txtUpCode.Text & "|" & txtName.Text & "|" & txtSymbol.Text & "|" & chkGoOn.Value & "|" & chkCodeLen.Value
    
    If strTemp <> mstr所有记录 Then
        If MsgBox("已有内容被修改了，是否退出？", vbYesNo, gstrSysName) = vbYes Then
            gblnCancel = True
            Unload Me
        End If
    Else
        gblnCancel = True
        Unload Me
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim lngItemID As Long
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intTmp As Integer
    
    '重新检查名称，并去掉特殊字符
    strTmp = MoveSpecialChar(txtName.Text)
    If txtName.Text <> strTmp Then
        txtName.Text = strTmp
        Me.txtSymbol.Text = zlStr.GetCodeByORCL(Me.txtName.Text, True, 10)
    End If
    
    If Me.txtCode.MaxLength = 12 Then
        MsgBox "上级编码已经达到最大长度，不能设置下级！", vbExclamation, gstrSysName
        Me.cmdCancel.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.txtUpCode.Text)) + Len(Trim(Me.txtCode.Text)) > 8 Then
        MsgBox "编码超长，编码最大长度为8位字符！", vbExclamation, gstrSysName
        Me.cmdCancel.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtCode.Text) = "" Then
        MsgBox "编码必须输入", vbExclamation, gstrSysName
        Me.txtCode.SetFocus
        Exit Sub
    End If
    If Me.chkCodeLen.Value = 0 And Len(Trim(Me.txtCode.Text)) <> Me.txtCode.MaxLength Then
        MsgBox "编码长度必须为" & Me.txtCode.MaxLength & "位，除非你选择更改长度选项", vbExclamation, gstrSysName
        Me.txtCode.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "名称必须输入", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "名称超过" & Me.txtName.MaxLength & "的长度限制", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    '编码的子级是否大于8字符长
    If Me.chkCodeLen.Value = 1 And Me.lblKind.Tag <> "" Then
        If Me.txtParent.Tag = 0 Then
            gstrSql = "select nvl(max(length(编码)),0) LEN from 诊疗分类目录 where 类型=[2] start with 上级id is null connect by prior id=上级id"
        Else
            gstrSql = "select nvl(max(length(编码)),0) LEN from 诊疗分类目录  where 类型=[2] start with 上级id=[1] connect by prior id=上级id"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.txtParent.Tag, Val(Me.lblKind.Tag))
        If Not rsTmp.EOF Then
            intTmp = rsTmp!Len
        End If
        rsTmp.Close
        If Len(Trim(Me.txtCode.Text)) - Int(Me.txtCode.Tag) + intTmp > 8 Then
            MsgBox "编码的子级会超8位字符长！", vbExclamation, gstrSysName
            Me.txtCode.SetFocus
            Exit Sub
        End If
    End If
    
    err = 0: On Error GoTo ErrHand
    If Me.Tag = "增加" Then
        lngItemID = zlDatabase.GetNextId("诊疗分类目录")
        gstrSql = "zl_诊疗分类目录_Insert(" & _
            lngItemID & "," & _
            Me.txtParent.Tag & "," & _
            "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            "'" & Trim(Me.txtName.Text) & "'," & _
            "'" & Trim(Me.txtSymbol.Text) & "'," & _
            Me.lblKind.Tag & "," & _
            Me.chkCodeLen.Value & ")"
    Else
        lngItemID = Me.Tag
        gstrSql = "zl_诊疗分类目录_Update(" & _
            lngItemID & "," & _
            Me.txtParent.Tag & "," & _
            "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            "'" & Trim(Me.txtName.Text) & "'," & _
            "'" & Trim(Me.txtSymbol.Text) & "'," & _
            Me.chkCodeLen.Value & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    mblnOK = True
    If chkGoOn.Value Then
        txtName.Text = ""
        txtName.SetFocus
        txtSymbol.Text = ""
        Call zlDefaultCode
        lblMsg.Visible = True
    Else
        Unload Me
    End If
    mblnChanged = False
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    Dim blnRe As Boolean
    Dim lngCount As Long
    
    If Me.Tag = "增加" Then
        gstrSql = "select ID,上级ID,编码,名称,简码" & _
                " From 诊疗分类目录" & _
                " Where 类型 = " & Val(Me.lblKind.Tag) & _
                " start with 上级ID is null" & _
                " connect by prior ID=上级ID"
        chkGoOn.Visible = True
    Else
        gstrSql = "select ID,上级ID,编码,名称,简码" & _
                " From 诊疗分类目录" & _
                " Where 类型 = " & Val(Me.lblKind.Tag) & _
                " and  id not in (select id from 诊疗分类目录 start with ID = " & Val(Me.Tag) & "connect by prior id=上级id ) " & _
                " start with 上级ID is null" & _
                " connect by prior ID=上级ID"
    End If
    blnRe = frmTreeSel.ShowTree(gstrSql, mstrID, mstr名称, mstr编码, "", "药品分类", "所有分类", False)
    If blnRe Then
        txtParent.Text = "[" & mstr编码 & "]" & mstr名称
        txtParent.Tag = mstrID
        Me.txtParent.SetFocus
        Call zlDefaultCode
    End If
End Sub

Private Sub Form_Activate()
    gblnCancel = False
    With Me.lblKind
        Select Case Val(.Tag)
        Case 1
            .Caption = "西成药"
            Me.Caption = "西成药分类编辑"
            Me.lblNote.Caption = "西成药通常可按照药品药理作用和化学作用，结合临床各科用药区别等进行分类设置。"
        Case 2
            .Caption = "中成药"
            Me.Caption = "中成药分类编辑"
            Me.lblNote.Caption = "中成药通常可以根据其组成的性味与功用等进行分类，也可选择其他分类方法进行分类。"
        Case 3
            .Caption = "中草药"
            Me.Caption = "中草药分类编辑"
            Me.lblNote.Caption = "中草药通常可以根据其性味归经功用等进行分类，也可选择如自然系统属性的分类方法。"
        Case 4
            .Caption = "中药配方"
            Me.Caption = "中药配方分类编辑"
            Me.lblNote.Caption = "配方分类参照中成药的分类方法进行，根据配方多少可以更细致或粗略地进行。"
        Case 5
            .Caption = "诊疗项目"
            Me.Caption = "诊疗分类编辑"
            Me.lblNote.Caption = "诊疗项目可根据临床与医技各科诊疗应用操作的特点进行统一分类设置。"
        Case 6
            .Caption = "成套诊疗"
            Me.Caption = "成套诊疗分类编辑"
            Me.lblNote.Caption = "成套诊疗可根据医院内部诊疗操作规范和设置的成套诊疗方案的数目进行粗略的分类。"
        End Select
    End With
    
    err = 0: On Error GoTo ErrHand
    mblnCancel = True
    
    If Me.Tag = "增加" Then
        gstrSql = "select ID,上级ID,编码,名称,简码" & _
                " From 诊疗分类目录" & _
                " Where 类型 = [1] " & _
                " start with 上级ID is null" & _
                " connect by prior ID=上级ID"
        chkGoOn.Visible = True
    Else
        gstrSql = "select ID,上级ID,编码,名称,简码" & _
                " From 诊疗分类目录" & _
                " Where 类型 = [1] " & _
                " and  id not in (select id from 诊疗分类目录 start with ID = [2] connect by prior id=上级id ) " & _
                " start with 上级ID is null" & _
                " connect by prior ID=上级ID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblKind.Tag), Val(Me.Tag))
    intMaxLen = rsTemp.Fields("编码").DefinedSize
    Me.txtName.MaxLength = rsTemp.Fields("名称").DefinedSize
    Me.txtSymbol.MaxLength = rsTemp.Fields("简码").DefinedSize
    
    If Me.Tag = "增加" Then
        Call zlDefaultCode
    End If
    mblnCancel = False
    
    mstr所有记录 = ""
    mstr所有记录 = txtParent.Text & "|" & txtUpCode.Text & "|" & txtName.Text & "|" & txtSymbol.Text & "|" & chkGoOn.Value & "|" & chkCodeLen.Value
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub Form_Load()
    mblnChanged = False
    mblnOK = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnOK = False Then
        gblnCancel = True '也当是点击取消按钮，在返回主界面的时候不刷新数据
    Else
        gblnCancel = False
    End If
    mbln药品目录 = False
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If mblnChanged = True Then
'        If MsgBox("设置已经改变，您确认要退出吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
'            Cancel = 1
'        End If
'    End If
'End Sub

Private Sub lblCode_Click()
    Me.txtCode.SetFocus
End Sub

Private Sub lblName_Click()
    Me.txtName.SetFocus
End Sub

Private Sub lblParent_Click()
    Me.txtParent.SetFocus
End Sub

Private Sub lblSymbol_Click()
    Me.txtSymbol.SetFocus
End Sub

Private Sub txtCode_Change()
    If mblnCancel = True Then Exit Sub
    If Me.Visible Then mblnChanged = True
    lblMsg.Visible = False
End Sub

Private Sub txtcode_GotFocus()
    Me.txtCode.SelStart = 0: Me.txtCode.SelLength = 100
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If mbln药品目录 Then
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        End If
    End Select
    KeyAscii = 0
End Sub

Private Sub txtName_Change()
    If mblnCancel = True Then Exit Sub
    If Me.Visible Then mblnChanged = True
    lblMsg.Visible = False
End Sub

Private Sub txtName_GotFocus()
    Call zlCommFun.OpenIme(True)
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
'    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        txtName.Text = MoveSpecialChar(txtName.Text)
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim blnCodeType As Boolean
    blnCodeType = zlDatabase.GetPara("简码方式")
    If txtName.Text <> "" Then Me.txtSymbol.Text = zlStr.GetCodeByORCL(Me.txtName.Text, blnCodeType, 10)
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtParent_Change()
    If mblnCancel = True Then Exit Sub
    If Me.Visible Then mblnChanged = True
    lblMsg.Visible = False
End Sub

Private Sub txtParent_GotFocus()
    Me.txtParent.SelStart = 0: Me.txtParent.SelLength = 100
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Me.txtParent.Tag = 0
        Call zlDefaultCode
    End If
    Me.txtParent.SelStart = 0: Me.txtParent.SelLength = 100
End Sub

Private Sub txtParent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtSymbol_Change()
    If mblnCancel = True Then Exit Sub
    If Me.Visible Then mblnChanged = True
    lblMsg.Visible = False
End Sub

Private Sub txtSymbol_GotFocus()
    Me.txtSymbol.SelStart = 0: Me.txtSymbol.SelLength = 100
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtUpCode_Change()
    Me.txtCode.Width = txtUpCode.Width - TextWidth(txtUpCode.Text) - 120
    Me.txtCode.Left = txtUpCode.Left + TextWidth(txtUpCode.Text) + 60
    If mblnCancel = True Then Exit Sub
    If Me.Visible Then mblnChanged = True
    lblMsg.Visible = False
End Sub

Private Sub zlDefaultCode()
    Dim strSql As String
    Dim rs上级 As ADODB.Recordset
    '-----------------------------------------------------
    '功能：根据选择的上级ID(存放于Me.txtParent.Tag)，调整设置编码的缺省值
    '-----------------------------------------------------
    err = 0:
    
    On Error Resume Next
    
    Me.chkCodeLen.Value = 0
    Me.chkCodeLen.Enabled = True
   
    If Me.txtParent.Tag = 0 Then
        Me.txtParent.Text = "(无)"
        Me.txtUpCode.Text = ""
        gstrSql = "select max(编码) as 编码 From 诊疗分类目录 Where 类型=[1] and 上级ID is null "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblKind.Tag))
        intMaxLen = rsTemp.Fields("编码").DefinedSize
        With rsTemp
            If IIf(IsNull(!编码), "", !编码) = "" Then
                Me.txtCode.Text = "01"
                Me.txtCode.MaxLength = intMaxLen
                Me.txtCode.Tag = Me.txtCode.MaxLength
                Me.chkCodeLen.Value = 1
                Me.chkCodeLen.Enabled = False
            Else
                Me.txtCode.MaxLength = Len(Trim(!编码))
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If !编码 = String(Me.txtCode.MaxLength, "9") Or !编码 = String(Me.txtCode.MaxLength, "Z") Or Len(zlCommFun.IncStr(!编码)) > txtCode.MaxLength Then
                    If Me.txtCode.MaxLength >= intMaxLen Then
                        MsgBox "最大编码和编码长度已经达到最大限制，无法递增编码", vbExclamation, gstrSysName
                        Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                        Me.chkCodeLen.Value = 0
                        Me.chkCodeLen.Enabled = False
                    Else
                        MsgBox "最大编码已经达到本级限制，你可以扩充编码长度以满足需要", vbExclamation, gstrSysName
'                        Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                        Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                        Me.txtCode.Text = zlCommFun.IncStr(!编码)
                        Me.txtCode.Tag = Me.txtCode.MaxLength
                        Me.chkCodeLen.Value = 1
                    End If
                Else
                    Me.txtCode.Text = zlCommFun.IncStr(!编码)
'                    Me.txtCode.Text = Format(Mid(!编码, Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                End If
            End If
        End With
    Else
        Me.txtUpCode.Text = Mid(Split(txtParent.Text, "]")(0), 2)
        strSql = "select id,编码 from 诊疗分类目录 where 上级id=[1]"
        Set rs上级 = zlDatabase.OpenSQLRecord(strSql, "查询是否有下级", Val(txtParent.Tag))
        intMaxLen = rs上级.Fields("编码").DefinedSize
        If rs上级.RecordCount = 0 Then
            Me.txtCode.MaxLength = intMaxLen - Len(Me.txtUpCode.Text)
            Me.txtCode.Tag = Me.txtCode.MaxLength
            If Me.txtCode.MaxLength = 12 Then
                MsgBox "上级编码已经达到最大长度，不能设置下级！", vbExclamation, gstrSysName
                Me.cmdCancel.SetFocus
                Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                Me.chkCodeLen.Value = 0
                Me.chkCodeLen.Enabled = False
                Exit Sub
            End If
            If Me.txtCode.MaxLength > 1 Then
                Me.txtCode.Text = "01"
            Else
                Me.txtCode.Text = "1"
            End If
            Me.chkCodeLen.Value = 1
            Me.chkCodeLen.Enabled = False
        Else
            gstrSql = "select nvl(max(编码),'') as 编码  From 诊疗分类目录 Where 类型=[1] and 上级ID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblKind.Tag), Val(txtParent.Tag))
            With rsTemp
                Me.txtCode.MaxLength = Len(!编码) - Len(Me.txtUpCode.Text)
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Mid(!编码, Len(Me.txtUpCode.Text) + 1) = String(Me.txtCode.MaxLength, "9") Or Mid(!编码, Len(Me.txtUpCode.Text) + 1) = String(Me.txtCode.MaxLength, "Z") Or Len(zlCommFun.IncStr(Mid(!编码, Len(Me.txtUpCode.Text) + 1))) > txtCode.MaxLength Then
                    If Len(Me.txtUpCode.Text) + Me.txtCode.MaxLength >= intMaxLen Then
                        MsgBox "该分类下级最大编码和编码长度已经达到最大限制，无法递增编码", vbExclamation, gstrSysName
                        Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                        Me.chkCodeLen.Value = 0
                        Me.chkCodeLen.Enabled = False
                    Else
                        MsgBox "该分类下级最大编码已经达到本级限制，你可以扩充编码长度以满足需要", vbExclamation, gstrSysName
'                        Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                        Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                        Me.txtCode.Text = zlCommFun.IncStr(Mid(!编码, Len(Me.txtUpCode.Text) + 1))
                        Me.txtCode.Tag = Me.txtCode.MaxLength
                        Me.chkCodeLen.Value = 1
                    End If
                Else
'                    Me.txtCode.Text = Format(Mid(!编码, Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                    Me.txtCode.Text = zlCommFun.IncStr(Mid(!编码, Len(Me.txtUpCode.Text) + 1))
                End If
            End With
        End If
    End If
    Me.txtName.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
