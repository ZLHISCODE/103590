VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeWholeSetClassEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "成套收费项目分类编辑"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   Icon            =   "frmChargeWholeSetClassEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   375
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3135
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1950
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "编码"
      Text            =   "000000"
      Top             =   1545
      Width           =   1380
   End
   Begin VB.TextBox txtSymbol 
      Height          =   300
      Left            =   1815
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "名称"
      Top             =   2265
      Width           =   1425
   End
   Begin VB.CheckBox chkCodeLen 
      Caption         =   "允许更改编码长度，并按此调整各同级编码(&L)"
      Height          =   285
      Left            =   1095
      TabIndex        =   7
      Top             =   2685
      Width           =   4290
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   -30
      TabIndex        =   6
      Top             =   2970
      Width           =   6345
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3345
      TabIndex        =   5
      Top             =   3150
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4575
      TabIndex        =   4
      Top             =   3150
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1815
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "名称"
      Top             =   1905
      Width           =   3795
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&P"
      Height          =   300
      Left            =   5325
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   885
      Width           =   285
   End
   Begin VB.TextBox txtParent 
      Height          =   300
      Left            =   1815
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "(无)"
      ToolTipText     =   "按Del清除上级，设置初级分类"
      Top             =   885
      Width           =   3495
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   2310
      Left            =   -3720
      TabIndex        =   0
      Tag             =   "1000"
      Top             =   2100
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   4075
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   135
      Top             =   1275
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
            Picture         =   "frmChargeWholeSetClassEdit.frx":0442
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeWholeSetClassEdit.frx":09DC
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUpCode 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1815
      MaxLength       =   10
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "编码"
      Text            =   "0000"
      Top             =   1500
      Width           =   1620
   End
   Begin VB.Label lblSymbol 
      AutoSize        =   -1  'True
      Caption         =   "简码(&S)"
      Height          =   180
      Left            =   1095
      TabIndex        =   18
      Top             =   2325
      Width           =   630
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "成套分类"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   735
      Width           =   720
   End
   Begin VB.Label lblNote 
      Caption         =   "    成套收费项目可根据临床与医技各科应用操作的特点进行统一分类设置。"
      Height          =   435
      Left            =   1065
      TabIndex        =   16
      Top             =   255
      Width           =   4755
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmChargeWholeSetClassEdit.frx":0F76
      Top             =   225
      Width           =   480
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "编码(&D)"
      Height          =   180
      Left            =   1095
      TabIndex        =   15
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Left            =   1095
      TabIndex        =   14
      Top             =   1950
      Width           =   630
   End
   Begin VB.Label lblParent 
      AutoSize        =   -1  'True
      Caption         =   "上级(&U)"
      Height          =   180
      Left            =   1095
      TabIndex        =   13
      Top             =   945
      Width           =   630
   End
   Begin VB.Label lblHint 
      Caption         =   "(提示：按Del清除上级，设置初级分类)"
      Height          =   210
      Left            =   1785
      TabIndex        =   12
      Top             =   1245
      Width           =   3330
   End
End
Attribute VB_Name = "frmChargeWholeSetClassEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum EditWhileSetType
    Ed_增加 = 1
    Ed_修改 = 2
End Enum
Private mEditType As EditWhileSetType
Private mstrPrivs As String, mlngModule As Long
Private mstrID As String
Private mintMaxLen As Integer
Private mObjNode As Node
Private mblnChanged As Boolean
Private mintSucces As Integer
Private mlng上级ID As Long
Private mstrLike  As String
Private mblnFirst As Boolean
Public Function EditCard(ByVal frmMain As Form, ByVal EditType As EditWhileSetType, _
    ByVal strPrivs As String, ByVal lngModule As Long, ByVal lng上级ID As Long, ByVal strID As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进入分类编辑接口
    '入参:
    '出参:
    '返回:编辑成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2010-08-26 13:30:38
    '说明:27327
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mstrPrivs = strPrivs: mlngModule = lngModule: mstrID = strID: mintSucces = 0
    mlng上级ID = lng上级ID
    Me.Show 1, frmMain
    EditCard = mintSucces > 0
End Function
Private Sub chkCodeLen_Click()
    On Error GoTo ErrHandle
    If Me.chkCodeLen.value = 1 Then
        Me.txtCode.MaxLength = mintMaxLen - Len(Me.txtUpCode.Text)
    Else
        Me.txtCode.MaxLength = Me.txtCode.Tag
        Me.txtCode.Text = Mid(Me.txtCode.Text, 1, Me.txtCode.MaxLength)
    End If
    Me.txtCode.SetFocus
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkCodeLen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim lngItemID As Long
    
    If Me.txtCode.MaxLength = 0 Then
        MsgBox "上级编码已经达到最大长度，不能设置下级！", vbExclamation, gstrSysName
        Me.cmdCancel.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtCode.Text) = "" Then
        MsgBox "编码必须输入", vbExclamation, gstrSysName
        Me.txtCode.SetFocus
        Exit Sub
    End If
    If Me.chkCodeLen.value = 0 And Len(Trim(Me.txtCode.Text)) <> Me.txtCode.MaxLength Then
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
    
    Err = 0: On Error GoTo ErrHand
    If mEditType = Ed_增加 Then
        lngItemID = sys.NextId("成套项目分类")
        'Zl_成套项目分类_Insert
        gstrSQL = "ZL_成套项目分类_INSERT("
    Else
        lngItemID = Val(mstrID)
        'Zl_成套项目分类_Update
        gstrSQL = "ZL_成套项目分类_UPDATE("
    End If
    '  Id_In      成套项目分类.ID%Type,
    gstrSQL = gstrSQL & "" & lngItemID & ","
    '  上级id_In  成套项目分类.上级id%Type,
    gstrSQL = gstrSQL & "" & IIF(Val(Me.txtParent.Tag) = 0, "NULL", Val(Me.txtParent.Tag)) & ","
    '  编码_In    成套项目分类.编码%Type,
    gstrSQL = gstrSQL & "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "',"
    '  名称_In    成套项目分类.名称%Type,
    gstrSQL = gstrSQL & "'" & Trim(Me.txtName.Text) & "',"
    '  简码_In    成套项目分类.简码%Type,
    gstrSQL = gstrSQL & "'" & Trim(Me.txtSymbol.Text) & "',"
    '  v_Brethren Number
    '  --是否对同级编码进行长度处理,0-否,1-是
    gstrSQL = gstrSQL & "" & Me.chkCodeLen.value & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mintSucces = mintSucces + 1
    mblnChanged = False
    If mEditType = Ed_修改 Then Unload Me: Exit Sub
    txtName.Text = ""
    Call zlDefaultCode
    mblnChanged = False
    txtName.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    Call SearchPreLevel("")
End Sub
Private Function SearchPreLevel(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择上级分类
    '返回:
    '编制:刘兴洪
    '日期:2010-08-26 13:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey As String, strWhere As String
    Dim vRect As RECT, bytStyle As Byte
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHandle
    strKey = mstrLike & strInput & "%"
    If strInput <> "" Then
        If IsNumeric(strInput) Then
            strWhere = " 编码 Like [1]"
        ElseIf zlStr.IsCharAlpha(strInput) Then
            strWhere = " 简码 Like upper([1])"
        Else
            strWhere = " 编码 Like [1] or 简码 Like upper([1]) or 名称 like [1]"
        End If
        If mEditType = Ed_修改 Then
            strWhere = "( " & strWhere & ")  and ID not in (select id from 成套项目分类 start with ID = [2] connect by prior id=上级id )"
        End If
        gstrSQL = "" & _
        " Select ID,上级ID,编码,名称,简码" & _
        " From 成套项目分类" & _
        " Where " & strWhere
        bytStyle = 0
    Else
        If mEditType = Ed_修改 Then
            gstrSQL = "" & _
            " Select ID,上级ID,编码,名称,简码" & _
            " From 成套项目分类" & _
            " Where id not in (select id from 成套项目分类 start with ID = [2] connect by prior id=上级id ) " & _
            " Start with 上级ID is null  connect by prior ID=上级ID"
        Else
            gstrSQL = "" & _
            " Select ID,上级ID,编码,名称,简码" & _
            " From 成套项目分类" & _
            " Start with 上级ID is null Connect by prior ID=上级ID"
        End If
        bytStyle = 1
    End If
    
    vRect = zlControl.GetControlRect(txtParent.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, bytStyle, "成套收费项目分类", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txtParent.Height, blnCancel, False, True, strKey, mlng上级ID)
    
    If blnCancel = True Then
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "未找到匹配的分类信息,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "未找到匹配的分类信息,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtParent.Enabled And txtParent.Visible Then txtParent.SetFocus
        Exit Function
    End If
    txtParent.Text = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
    txtParent.Tag = Nvl(rsTemp!ID)
    Call zlDefaultCode
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    SearchPreLevel = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ReadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取分类信息
    '返回:读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-08-26 13:35:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    Select Case mEditType
    Case Ed_增加
        gstrSQL = "" & _
        " Select A.ID,A.上级ID,A.编码,A.名称,A.简码" & _
        " From 成套项目分类 A" & _
        " Where id=0"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        mintMaxLen = rsTemp.Fields("编码").DefinedSize
        Me.txtName.MaxLength = rsTemp.Fields("名称").DefinedSize
        Me.txtSymbol.MaxLength = rsTemp.Fields("简码").DefinedSize
        Me.txtParent.Tag = mlng上级ID
        Call zlDefaultCode
    Case Else
        gstrSQL = "" & _
        " Select A.ID,A.上级ID,A.编码,A.名称,A.简码,B.编码 as 上级编码,B.名称 as 上级名称" & _
        " From 成套项目分类 A,成套项目分类 B" & _
        " Where A.ID=[1] and A.上级id=b.id(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        If rsTemp.EOF Then
            MsgBox "该分类可能已经被他人删除,不能进行修改!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        mlng上级ID = Val(Nvl(rsTemp!上级id))
        txtParent.Text = IIF(mlng上级ID = 0, "无", Nvl(rsTemp!上级编码) & "-" & Nvl(rsTemp!上级名称))
        txtParent.Tag = mlng上级ID
        txtUpCode.Text = Nvl(rsTemp!上级编码)
        txtCode.Text = Mid(Nvl(rsTemp!编码), Len(txtUpCode.Text) + 1)
        txtCode.MaxLength = Len(txtCode.Text)
        txtName.Text = Nvl(rsTemp!名称)
        txtSymbol.Text = Nvl(rsTemp!简码)
        mintMaxLen = rsTemp.Fields("编码").DefinedSize
        Me.txtName.MaxLength = rsTemp.Fields("名称").DefinedSize
        Me.txtSymbol.MaxLength = rsTemp.Fields("简码").DefinedSize
    End Select
    ReadCard = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call ReadCard
    Me.txtCode.ZOrder
    mblnChanged = False
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
End Sub

Private Sub Form_Load()
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mblnChanged = False: mblnFirst = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChanged = True Then
        If MsgBox("设置已经改变，您确认要退出吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

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

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txtParent.SetFocus
    Call zlDefaultCode
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If cmdSelect Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtCode_Change()
    mblnChanged = True
End Sub

Private Sub txtCode_GotFocus()
    Me.txtCode.SelStart = 0: Me.txtCode.SelLength = 100
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtName_Change()
    mblnChanged = True
End Sub

Private Sub txtName_GotFocus()
    Call OS.OpenIme(True)
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=`;'"":/.,?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.txtSymbol.Text = Mid(zlStr.GetCodeByORCL(Me.txtName.Text), 1, 10)
End Sub

Private Sub txtName_LostFocus()
    Call OS.OpenIme(False)
End Sub

Private Sub txtParent_Change()
    mblnChanged = True
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
        Call OS.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtSymbol_Change()
    mblnChanged = True
End Sub

Private Sub txtSymbol_GotFocus()
    Me.txtSymbol.SelStart = 0: Me.txtSymbol.SelLength = 100
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtUpCode_Change()
    Me.txtCode.Width = txtUpCode.Width - TextWidth(txtUpCode.Text) - 120
    Me.txtCode.Left = txtUpCode.Left + TextWidth(txtUpCode.Text) + 60
    mblnChanged = True
End Sub

Private Sub zlDefaultCode()
    '-----------------------------------------------------
    '功能：根据选择的上级ID(存放于txtParent.Tag))，调整设置编码的缺省值
    '-----------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo ErrHand
    Me.chkCodeLen.value = 0
    Me.chkCodeLen.Enabled = True
NotPreID:
    If Val(txtParent.Tag) = 0 Then
        Me.txtParent.Text = "(无)"
        Me.txtUpCode.Text = ""
        gstrSQL = "select max(编码) as 编码 From 成套项目分类 Where 上级ID is null "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With rsTemp
            If IIF(IsNull(!编码), "", !编码) = "" Then
                Me.txtCode.Text = "01"
                Me.txtCode.MaxLength = mintMaxLen
                Me.txtCode.Tag = Me.txtCode.MaxLength
                Me.chkCodeLen.value = 1
                Me.chkCodeLen.Enabled = False
            Else
                Me.txtCode.MaxLength = Len(Trim(Nvl(!编码)))
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Nvl(!编码) = String(Me.txtCode.MaxLength, "9") Then
                    If Me.txtCode.MaxLength >= mintMaxLen Then
                        MsgBox "最大编码和编码长度已经达到最大限制，无法递增编码", vbExclamation, gstrSysName
                        Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                        Me.chkCodeLen.value = 0
                        Me.chkCodeLen.Enabled = False
                    Else
                        MsgBox "最大编码已经达到本级限制，你可以扩充编码长度以满足需要", vbExclamation, gstrSysName
                        Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                        Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                        Me.txtCode.Tag = Me.txtCode.MaxLength
                        Me.chkCodeLen.value = 1
                    End If
                Else
                    Me.txtCode.Text = Format(Mid(Nvl(!编码), Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                End If
            End If
        End With
    Else
        gstrSQL = "select 编码,名称 From 成套项目分类 Where ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(txtParent.Tag))
        If rsTemp.EOF Then
            MsgBox "上级编码可能被他人删除,请检查!", vbInformation + vbDefaultButton1, gstrSysName
            mlng上级ID = 0: GoTo NotPreID:
        End If
        Me.txtParent.Text = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
        Me.txtUpCode.Text = Nvl(rsTemp!编码)
        Me.txtCode.MaxLength = IIF(mintMaxLen - Len(Me.txtUpCode.Text) > 0, mintMaxLen - Len(Me.txtUpCode.Text), 1)
        Me.txtCode.Tag = Me.txtCode.MaxLength
        
        gstrSQL = "select nvl(编码,'') as 编码  From 成套项目分类 Where 上级ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(txtParent.Tag))
        If rsTemp.EOF Then
            '没有子项,
            If Me.txtCode.MaxLength > 1 Then
                Me.txtCode.Text = "01"
            Else
                Me.txtCode.Text = "1"
            End If
            Me.chkCodeLen.value = 1
            Me.chkCodeLen.Enabled = False
        Else
            With rsTemp
                Me.txtCode.MaxLength = IIF(Len(Nvl(!编码)) - Len(Me.txtUpCode.Text) > 0, Len(Nvl(!编码)) - Len(Me.txtUpCode.Text), 1)
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Mid(Nvl(!编码), Len(Me.txtUpCode.Text) + 1) = String(Me.txtCode.MaxLength, "9") Then
                    If Len(Me.txtUpCode.Text) + Me.txtCode.MaxLength >= mintMaxLen Then
                        MsgBox "该分类下级最大编码和编码长度已经达到最大限制，无法递增编码", vbExclamation, gstrSysName
                        Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                        Me.chkCodeLen.value = 0
                        Me.chkCodeLen.Enabled = False
                    Else
                        MsgBox "该分类下级最大编码已经达到本级限制，你可以扩充编码长度以满足需要", vbExclamation, gstrSysName
                        Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                        Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                        Me.txtCode.Tag = Me.txtCode.MaxLength
                        Me.chkCodeLen.value = 1
                    End If
                Else
                    If Len(Nvl(!编码)) >= Len(Me.txtUpCode.Text) + 1 Then
                        Me.txtCode.Text = Format(Mid(Nvl(!编码), Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                    End If
                End If
            End With
        End If
    End If
    Me.txtParent.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




