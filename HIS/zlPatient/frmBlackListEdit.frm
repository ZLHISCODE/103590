VERSION 5.00
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#3.4#0"; "zlIDKind.ocx"
Begin VB.Form frmBlackListEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "特殊病人"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frmBlackListEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt编号 
      ForeColor       =   &H00C00000&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4245
      MaxLength       =   10
      TabIndex        =   14
      Top             =   1050
      Width           =   885
   End
   Begin VB.TextBox txtNote 
      Height          =   870
      Left            =   105
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1395
      Width           =   5220
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3855
      TabIndex        =   11
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   10
      Top             =   2400
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Height          =   960
      Left            =   105
      TabIndex        =   12
      Top             =   15
      Width           =   5220
      Begin VB.CommandButton cmdPati 
         Height          =   300
         Left            =   2625
         Picture         =   "frmBlackListEdit.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "选择病人(F2)"
         Top             =   225
         Width           =   300
      End
      Begin VB.TextBox txtPatient 
         Height          =   300
         Left            =   1350
         TabIndex        =   1
         Top             =   225
         Width           =   1275
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   300
         Left            =   705
         TabIndex        =   15
         ToolTipText     =   "快捷键F4"
         Top             =   225
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   529
         Appearance      =   2
         IDKindStr       =   $"frmBlackListEdit.frx":0C74
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         BackColor       =   -2147483633
      End
      Begin VB.Label lbl床号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号："
         Height          =   180
         Left            =   3960
         TabIndex        =   7
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室："
         Height          =   180
         Left            =   2325
         TabIndex        =   6
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lbl标识号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标识号："
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Left            =   3960
         TabIndex        =   4
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Left            =   2985
         TabIndex        =   3
         Top             =   285
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         Height          =   180
         Left            =   330
         TabIndex        =   0
         Top             =   285
         Width           =   360
      End
   End
   Begin VB.Label lbl编号 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "登记编号"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3480
      TabIndex        =   13
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "该病人加入特殊病人名单的原因："
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   1155
      Width           =   2700
   End
End
Attribute VB_Name = "frmBlackListEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private mstrPrivs As String
Private mlng编号 As Long
Private mblnDelete As Boolean
Private mblnOK As Boolean
Private mblnNotClick As Boolean


Public Function ShowMe(frmParent As Object, ByVal strPrivs As String, Optional ByVal lng编号 As Long, Optional ByVal blnDelete As Boolean) As Boolean
    mlng编号 = lng编号
    mblnDelete = blnDelete
    mstrPrivs = strPrivs
    mblnNotClick = False
    Me.Show 1, frmParent
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim blnTrans As Boolean
    
    If Val(txtPatient.Tag) = 0 Then
        MsgBox "请确定要加入特殊病人名单的病人。", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Sub
    End If
    If Val(txt编号.Text) = 0 Then
        MsgBox "请确定要加入特殊病人的登记编号。", vbInformation, gstrSysName
        txt编号.SetFocus: Exit Sub
    End If
    If txtNote.Text = "" Then
        MsgBox "请输入原因。", vbInformation, gstrSysName
        txtNote.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txtNote.Text) > txtNote.MaxLength Then
        MsgBox "最多允许输入 " & txtNote.MaxLength & " 个字符或 " & txtNote.MaxLength \ 2 & " 个汉字。", vbInformation, gstrSysName
        txtNote.SetFocus: Exit Sub
    End If
    
    If mlng编号 = 0 Then
        strSQL = "ZL_特殊病人_Insert(" & Val(txt编号.Text) & "," & Val(txtPatient.Tag) & ",'" & txtNote.Text & "')"
    Else
        If mblnDelete Then
            strSQL = "ZL_特殊病人_Delete(" & mlng编号 & ",'" & txtNote.Text & "')"
        Else
            strSQL = "ZL_特殊病人_Update(" & mlng编号 & "," & Val(txt编号.Text) & ",'" & txtNote.Text & "')"
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
    
Private Sub cmdPati_Click()
    frmPatiSel.mstrPrivs = mstrPrivs
    frmPatiSel.Show 1, Me
    If frmPatiSel.mlng病人ID <> 0 Then
        txtPatient.Text = "-" & frmPatiSel.mlng病人ID
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("姓名")
        mblnNotClick = False
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intIndex As Integer
    If KeyCode = vbKeyF4 Then
        If Shift = vbCtrlMask And IDKind.Enabled Then
            intIndex = IDKind.GetKindIndex("IC卡号")
            If intIndex < 0 Then Exit Sub
            IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
        End If
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdPati_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
    Call CreateMobjCard
    Call CreateSquareCardObject(Me, 1101)
     '初始化
    Call IDKind.zlInit(Me, 100, 1101, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    
    If Not gobjSquare.objSquareCard Is Nothing Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    If mlng编号 <> 0 Then
        txt编号.Text = mlng编号
        If mblnDelete Then
            Me.Caption = "撤消特殊病人"
            txt编号.Enabled = False
        End If
        fraPati.Enabled = False
        txtPatient.Enabled = False
        cmdPati.Visible = False
        Call GetPatient(IDKind.GetCurCard, "编号" & mlng编号)
    Else
        txt编号.Text = GetMaxNum
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXml As String
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hwnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call txtPatient_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, 1101, lng卡类别ID, False, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    
    Set gobjSquare.objCurCard = objCard
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And mblnNotClick = False Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, False)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, False)
End Sub

Private Sub txtNote_GotFocus()
    Call zlControl.TxtSelAll(txtNote)
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjIDCard.SetEnabled (True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjICCard.SetEnabled (True)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Function FindPati(ByVal objCard As Card, Optional blnCard As Boolean = False) As Boolean
    If Not GetPatient(objCard, txtPatient.Text, blnCard) Then
        If IsNumeric(txtPatient.Text) Then
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.Text = ""
        End If
        Call zlControl.TxtSelAll(txtPatient)
        txtPatient.SetFocus: Exit Function
    Else
        txtPatient.PasswordChar = ""
        txtPatient.IMEMode = 0
        txtNote.SetFocus: Exit Function
    End If
    FindPati = True
End Function

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
'功能：读取病人信息
    Dim lng卡类别ID As Long, lng病人ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnDo As Boolean
    Dim blnCode As Boolean, blnHavePassWord As Boolean
    Dim strPassWord As String, strErrMsg As String
    Dim strCard As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    If strInput Like "编号*" Then
        blnCode = True
        strSQL = "Select A.病人ID,A.姓名,A.性别,A.年龄,A.病人类型,A.险类," & _
            " Decode(Nvl(A.住院次数,0),0,A.门诊号,A.住院号) as 标识号," & _
            " D.名称 as 科室,B.出院病床 as 床号,C.加入原因,C.撤消原因,C.撤消时间" & _
            " From 病人信息 A,病案主页 B,特殊病人 C,部门表 D" & _
            " Where A.病人ID=B.病人ID(+) And Nvl(A.主页ID,0)=B.主页ID(+)" & _
            " And B.出院科室ID=D.ID(+) And Nvl(B.主页ID(+),0)<>0" & _
            " And A.病人ID=C.病人ID And C.编号=[2]"
        strInput = Mid(strInput, 3)
    Else
        blnCode = False
        strSQL = "Select A.病人ID,A.姓名,A.性别,A.年龄,A.病人类型,A.险类," & _
            " Decode(Nvl(A.住院次数,0),0,A.门诊号,A.住院号) as 标识号," & _
            " D.名称 as 科室,B.出院病床 as 床号" & _
            " From 病人信息 A,病案主页 B,部门表 D" & _
            " Where A.病人ID=B.病人ID(+) And Nvl(A.主页ID,0)=B.主页ID(+)" & _
            " And B.出院科室ID=D.ID(+) And Nvl(B.主页ID(+),0)<>0 And A.停用时间 is NULL"
            
        If blnCard = True And objCard.名称 Like "姓名*" Then    '刷卡
            If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
                lng卡类别ID = IDKind.GetfaultCard.接口序号
            Else
                lng卡类别ID = "-1"
            End If
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
            If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
            If lng病人ID <= 0 Then GoTo NotFoundPati:
            strInput = "-" & lng病人ID
            strSQL = strSQL & " And A.病人ID=[1]"
            blnHavePassWord = True
        ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
            strSQL = strSQL & " And A.病人ID=[1]"
        ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
            strSQL = strSQL & " And A.住院号=[1]"
        ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
            strSQL = strSQL & " And A.门诊号=[1]"
        Else
            Select Case objCard.名称
                Case "姓名", "姓名或就诊卡"
                    If gblnShowCard = True Then
                        strCard = "A.就诊卡号 as 就诊卡,A.就诊卡号 as 就诊卡号,"
                    Else
                        strCard = "LPAD('*',Length(A.就诊卡号),'*') as 就诊卡,A.就诊卡号 as 就诊卡号,"
                    End If
                    '通过姓名模糊查找病人(允许输入病人标识时)
                    strPati = _
                        " Select A.病人ID ID,A.病人ID,A.门诊号,A.住院号," & strCard & "A.姓名,A.性别,A.年龄,A.费别 as 门诊费别," & _
                        "   B.名称 as 病区,C.名称 as 科室,A.当前床号 as 床号,To_Char(A.入院时间,'YYYY-MM-DD') as 入院时间," & _
                        "   To_Char(A.出院时间,'YYYY-MM-DD') as 出院时间,A.住院次数,To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期," & _
                        "   A.国籍,A.民族,A.区域,A.学历,A.职业,A.身份,A.身份证号,A.家庭地址,A.工作单位,To_Char(A.登记时间,'YYYY-MM-DD') as 登记时间," & _
                        "   Nvl(P.病人类型,Decode(P.险类,Null,'普通病人','医保病人')) 病人类型" & _
                        " From 病案主页 P,病人信息 A,部门表 B,部门表 C" & _
                        " Where A.当前病区ID=B.ID(+) And A.当前科室ID=C.ID(+) And A.病人ID=P.病人ID(+) And A.主页ID=P.主页ID(+)" & _
                        "   And Nvl(P.主页ID(+),0)<>0 And A.停用时间 is NULL And A.姓名 Like [1]" & _
                        " Order by A.姓名,A.登记时间 Desc"
                    
                    vRect = zlControl.GetControlRect(txtPatient.hwnd)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                                
                    '只有一行数据时,blncancel返回false,按取消返回也是一样
                    If Not rsTmp Is Nothing Then
                        strSQL = strSQL & " And A.病人ID=[1]"
                        lng病人ID = Val(Nvl(rsTmp!病人ID))
                        If lng病人ID <= 0 Then GoTo NotFoundPati:
                        strInput = "-" & lng病人ID
                    ElseIf blnCancel = True Then
                        strSQL = strSQL & " And A.病人ID=[1]"
                        lng病人ID = Val(txtPatient.Tag)
                        If lng病人ID <= 0 Then GoTo NotFoundPati:
                        strInput = "-" & lng病人ID
                    Else
                        GoTo NotFoundPati
                    End If
                Case "医保号"
                    strInput = UCase(strInput)
                    strSQL = strSQL & " And A.医保号=[2]"
                Case "门诊号"
                    If Not IsNumeric(strInput) Then strInput = "0"
                    strSQL = strSQL & " And A.门诊号=[2]"
                Case Else
                    '其他类别的,获取相关的病人ID
                    If Val(objCard.接口序号) > 0 Then
                        lng卡类别ID = Val(objCard.接口序号)
                        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                        If lng病人ID = 0 Then GoTo NotFoundPati:
                    Else
                        If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                            strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    End If
                    If lng病人ID <= 0 Then GoTo NotFoundPati:
                    strSQL = strSQL & " And A.病人ID=[1]"
                    strInput = "-" & lng病人ID
                    blnHavePassWord = True
            End Select
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    blnDo = Not rsTmp.EOF
    If Not blnCode Then
        If blnDo Then blnDo = PatiAllow(rsTmp!病人ID, rsTmp!姓名)
    End If
    If blnDo Then
        txtPatient.Tag = rsTmp!病人ID
        txtPatient.Text = rsTmp!姓名
        '74426:李南春,2014-7-9,病人姓名颜色处理
        Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), Me.ForeColor, vbRed))
        lblSex.Caption = "性别：" & Nvl(rsTmp!性别)
        lblAge.Caption = "年龄：" & Nvl(rsTmp!年龄)
        lbl标识号.Caption = "标识号：" & Nvl(rsTmp!标识号)
        lbl科室.Caption = "科室：" & Nvl(rsTmp!科室)
        lbl床号.Caption = "床号：" & Nvl(rsTmp!床号)
        
        '修改时才读取
        If blnCode Then
            If mblnDelete Then
                lblNote.Caption = "将病人从特殊名单中撤消的原因："
                txtNote.Text = ""
            ElseIf IsNull(rsTmp!撤消时间) Then
                lblNote.Caption = "该病人加入特殊病人名单的原因："
                txtNote.Text = Nvl(rsTmp!加入原因)
            Else
                lblNote.Caption = "将病人从特殊名单中撤消的原因："
                txtNote.Text = Nvl(rsTmp!撤消原因)
                txt编号.Enabled = False
            End If
        End If
            
        GetPatient = True
    Else
NotFoundPati:
        txtPatient.Tag = ""
        txtPatient.Text = ""
        txtPatient.ForeColor = Me.ForeColor
        lblSex.Caption = "性别："
        lblAge.Caption = "年龄："
        lbl标识号.Caption = "标识号："
        lbl科室.Caption = "科室："
        lbl床号.Caption = "床号："
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.ActiveFastKey
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim blnCard As Boolean
    
    If IDKind.GetCurCard.名称 = "姓名" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If
    
    '刷卡完毕或输入号码后回车
    If blnCard And Len(Me.txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '读取病人信息
        Call FindPati(IDKind.GetCurCard, blnCard)
    End If
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt编号_GotFocus()
    Call zlControl.TxtSelAll(txt编号)
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function GetMaxNum() As Long
'功能：获取特殊性病人中的当前最大可用编号(补缺)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 编号+1 as 编号 From 特殊病人 Minus Select 编号 From 特殊病人"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rsTmp.EOF Then
        GetMaxNum = 1
    Else
        GetMaxNum = rsTmp!编号
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function PatiAllow(ByVal lng病人ID As Long, ByVal str姓名 As String) As Boolean
'功能：判断指定病人是否可以加入特殊病人
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    'by lesfeng 2010-03-06 性能优化，绑定变量
    strSQL = "Select 编号,病人ID,加入原因,加入时间,登记人,撤消原因,撤消时间,撤消人 From 特殊病人 Where 撤消时间 is Null And 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    If Not rsTmp.EOF Then
        MsgBox str姓名 & "已经加入特殊病人,原因：" & vbCrLf & vbCrLf & vbTab & Nvl(rsTmp!加入原因, "<没有原因>"), vbInformation, gstrSysName
        Exit Function
    End If
    PatiAllow = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CreateMobjCard()
    '创建卡部件
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
End Sub
