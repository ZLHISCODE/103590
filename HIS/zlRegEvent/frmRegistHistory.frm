VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmRegistHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人历次挂号信息查询"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10200
   Icon            =   "frmRegistHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   120
      Width           =   680
      _ExtentX        =   1191
      _ExtentY        =   661
      Appearance      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "宋体"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   4
      Top             =   975
      Width           =   10785
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1275
      TabIndex        =   3
      Top             =   120
      Width           =   3090
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   10785
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8550
      TabIndex        =   1
      Top             =   5970
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7260
      TabIndex        =   0
      Top             =   5970
      Width           =   1230
   End
   Begin VSFlex8Ctl.VSFlexGrid vsRegist 
      Height          =   4470
      Left            =   90
      TabIndex        =   5
      Top             =   1170
      Width           =   9975
      _cx             =   17595
      _cy             =   7885
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRegistHistory.frx":0442
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lbl险类 
      AutoSize        =   -1  'True
      Caption         =   "医保"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   11
      Top             =   645
      Width           =   480
   End
   Begin VB.Label txt险类 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1275
      TabIndex        =   10
      Top             =   570
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5970
      TabIndex        =   9
      Top             =   645
      Width           =   480
   End
   Begin VB.Label txt性别 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6510
      TabIndex        =   8
      Top             =   585
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7965
      TabIndex        =   7
      Top             =   645
      Width           =   480
   End
   Begin VB.Label txt年龄 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   6
      Top             =   585
      Width           =   1275
   End
End
Attribute VB_Name = "frmRegistHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mblnOlnyBJYB As Boolean
Private mstr号别 As String
Private mlng病人ID As Long
Private mstrPrivs As String, mintIDKind As Integer
Private mblnOk As Boolean
Private mbln允许住院病人挂号 As Boolean
Private Const mlngModule = 1111
Private mblnNotClick As Boolean
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------

Public Function ShowRegist(ByVal frmMain As Form, ByVal strPrivs As String, _
     ByVal bln允许住院病人挂号 As Boolean, blnOlnyBjYb As Boolean, _
    ByRef lng病人ID As Long, ByRef str号别 As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：选择病人的历次挂号信息
    '入参：blnOlnyBjYb- 是否北京医保
    '出参：str号别-返回所选择的号别
    '         lng病人ID-返回的选择的病人ID
    '返回：成功返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-07-16 14:34:24
    '说明：28604
    '------------------------------------------------------------------------------------------------------------------------
    mblnOlnyBJYB = blnOlnyBjYb: mlng病人ID = lng病人ID: mstrPrivs = strPrivs: mblnOk = False
    mbln允许住院病人挂号 = bln允许住院病人挂号
    str号别 = ""
    Me.Show 1, frmMain
    str号别 = mstr号别: lng病人ID = mlng病人ID
    ShowRegist = mblnOk
End Function

Private Sub cmdCancel_Click()
    mblnOk = False: Unload Me
End Sub

Private Sub cmdOK_Click()
      With vsRegist
            If .Row < 0 Then Exit Sub
            If Trim(.TextMatrix(.Row, .ColIndex("号别"))) = "" Then Exit Sub
            If Val(.RowData(.Row)) = 0 Then Exit Sub
            
            If Val(.RowData(.Row)) <> mlng病人ID And mlng病人ID <> 0 Then
                If MsgBox("注意:" & vbCrLf & " 姓名为『 " & txtPatient.Text & "』的病人不是挂号确定的病人,是否继续?" & vbCrLf & _
                "选择『是』:表示以当前选择的病人作为挂号病人。" & vbCrLf & _
                "选择『否』:表示不以此病人为准，返回查询界面。", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            End If
            mlng病人ID = Val(.RowData(.Row))
            mstr号别 = Trim(.TextMatrix(.Row, .ColIndex("号别")))
            
            mblnOk = True
            Unload Me:
      End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    Select Case KeyCode
        Case vbKeyF4
            If Shift = vbCtrlMask Then
                If IDKind.Enabled Then IDKind.IDKind = IDKind.GetKindIndex("IC卡号"): Call IDKind_Click(IDKind.GetCurCard)
            ElseIf Me.ActiveControl Is txtPatient Then
                If IDKind.Enabled Then
                    If Shift = vbShiftMask Then
                        IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDkindStr, ";")), IDKind.IDKind - 1)
                    Else
                        IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), 0, IDKind.IDKind + 1)
                    End If
                End If
            End If
        Case vbKeyF11
            If txtPatient.Enabled And txtPatient.Visible And Not txtPatient.Locked Then
                If Me.ActiveControl Is txtPatient Then
                    IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), IDKind.GetKindIndex("姓名"), IDKind.IDKind + 1)
                Else
                    txtPatient.SetFocus
                End If
            End If
        Case vbKeyReturn
       
    End Select
End Sub
Private Sub Form_Load()
    Dim strTemp As String
    
    Call InitIDKind
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    
    Call mobjIDCard.SetParent(Me.Hwnd)
    Call mobjICCard.SetParent(Me.Hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
    Call GetRegInFor(g私有模块, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    If mlng病人ID <> 0 Then
        txtPatient.Text = "-" & mlng病人ID
        Call GetPatient(Trim(txtPatient.Text))
    End If
    Call InitVsGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
End Sub
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If IsCardType(IDKind, "IC卡号") Then
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call GetPatient(Trim(txtPatient))
            End If
        End If
        Exit Sub
    End If
    lng卡类别ID = IDKind.GetCurCard.接口序号
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
        Call GetPatient(Trim(txtPatient))
    End If
End Sub

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    
'    If IDKind.GetCardNoLen <> 0 Then
'        txtPatient.MaxLength = IDKind.GetCardNoLen
'    Else
'        txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
'    End If
    
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
    mblnNotClick = True

    intIndex = IDKind.GetKindIndex(objCard.名称)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex
    txtPatient.Text = objPatiInfor.卡号
    Call txtPatient_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = "": txtPatient.ForeColor = Me.ForeColor
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
     zlControl.TxtSelAll txtPatient
      If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If IDKind.GetCurCard.名称 Like "姓名*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, glngSys, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) Then
            KeyAscii = 0
            '刷新病人信息:"-病人ID"
            Call GetPatient(txtPatient.Tag, False)
            Exit Sub
        End If
        KeyAscii = 0
        If IDKind.IDKind = IDKind.GetKindIndex("IC卡号") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        Call GetPatient(txtPatient.Text, blnCard)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub GetPatient(ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取病人信息
    '入参：blnCard=是否就诊卡刷卡
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-07-16 14:24:14
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str非在院 As String
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    If mbln允许住院病人挂号 = False Then
        str非在院 = " And Not Exists(Select 1 From 病案主页 Where 病人ID=B.病人ID And 主页ID=B.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
    End If
    
    strSQL = ""
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        'If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[1] " & str非在院
        strInput = UCase(strInput)
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSQL = strSQL & " And B.门诊号=[2]" & str非在院
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSQL = strSQL & " And B.病人ID=[2]" & str非在院
    Else
        Select Case IDKind.GetCurCard.名称
            Case "姓名", "姓名或就诊卡"
                '姓名
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtPatient.Text = mrsInfo!姓名 Then blnSame = True
                End If
                If Not blnSame Then
                    If Not gblnSeekName Or gblnSeekName And Len(txtPatient.Text) < 2 Then
                        Set mrsInfo = Nothing: Exit Sub
                    Else
                        strPati = _
                            " Select 1 as 排序ID,B.病人ID as ID,B.病人ID,B.姓名,B.性别,B.年龄,B.门诊号,B.出生日期,B.身份证号,B.家庭地址,B.工作单位" & _
                            " From 病人信息 B" & _
                            " Where Rownum <101 And B.停用时间 is NULL And B.姓名 Like [1]" & str非在院 & _
                            IIf(gintNameDays = 0, "", " And Nvl(B.就诊时间,B.登记时间)>Trunc(Sysdate-[2])")
                     
                        strPati = strPati & " Order by 排序ID,姓名"
                            
                        vRect = zlControl.GetControlRect(txtPatient.Hwnd)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", gintNameDays)
                        If Not rsTmp Is Nothing Then
                            If rsTmp!ID = 0 Then '当作新病人
                                MsgBox "未找到满足条件的病人信息,请检查!", vbInformation + vbDefaultButton1, gstrSysName
                                 txtPatient.Text = ""
                                Call txtPatient_GotFocus
                                Set mrsInfo = Nothing: Exit Sub
                            Else '以病人ID读取
                                strInput = rsTmp!病人ID
                                strSQL = strSQL & " And A.病人ID=[1]"
                            End If
                        Else '取消选择
                           MsgBox "未找到满足条件的病人信息,请检查!", vbInformation + vbDefaultButton1, gstrSysName
                            txtPatient.Text = ""
                            Call txtPatient_GotFocus
                            Set mrsInfo = Nothing: Exit Sub
                        End If
                    End If
                Else
                    '修改问题：39164
                    strInput = mrsInfo!病人ID
                    strSQL = strSQL & " And A.病人ID=[1]"
                End If
            Case "医保号"
                strInput = UCase(strInput)
                If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                    '仅北京医保才有效:见问题:问题:26982
                    strSQL = strSQL & " And B.医保号 like [3] " & str非在院
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And B.医保号=[1]" & str非在院
                End If
            Case "身份证号", "二代身份证", "身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSQL = strSQL & " And A.病人ID=[2] " & str非在院
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                 If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSQL = strSQL & " And A.病人ID=[2] " & str非在院
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.门诊号=[1]" & str非在院
            Case "挂号单号"
                 strInput = GetFullNO(strInput, 12)
                 txtPatient.Text = strInput
                strSQL = strSQL & " And A.NO=[1]" & str非在院

         Case Else
            '其他类别的,获取相关的病人ID
            If IDKind.GetCurCard.接口序号 > 0 Then
                lng卡类别ID = IDKind.GetCurCard.接口序号
                If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                If lng病人ID = 0 Then lng病人ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, strInput, False, lng病人ID, _
                    strPassWord, strErrMsg) = False Then lng病人ID = 0
            End If
            If lng病人ID <= 0 Then lng病人ID = 0
            strSQL = strSQL & " And A.病人ID=[1]" & str非在院
            strInput = "-" & lng病人ID
            blnHavePassWord = True
        End Select
    End If
    
    strSQL = "" & _
            "   Select distinct A.NO,A.号别,A.执行部门id,C.名称 as  挂号科室, B.病人ID," & _
            "            to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') as 挂号时间  " & vbNewLine & _
            "   From 病人挂号记录 A, 病人信息 B,部门表 C" & vbNewLine & _
            "   Where  A.执行部门ID=C.ID (+) " & _
            "               And B.病人id =A.病人id(+) and a.记录性质=1 and　 a.记录状态=1  " & strSQL & _
            "    Order by 挂号时间 Desc"
                                                             
    '没有设置黑名单,保留以前的处理方式,否则只能取当天的预约单(如果失了约的,则以红色字体显示)
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, CStr(Mid(strInput, 2)))
    If rsTmp.RecordCount = 0 Then
        vsRegist.Clear 1: vsRegist.Rows = 2: vsRegist.Row = 1
        MsgBox "未找到满足条件的病人信息,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        txtPatient.Text = ""
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        Call txtPatient_GotFocus
        Exit Sub
    End If
    
    If Val(Nvl(rsTmp!病人ID)) <> 0 Then
        strSQL = "Select A.*,B.名称 险类名称 From 病人信息 A,保险类别 B Where A.险类 = B.序号(+) And A.停用时间 is NULL "
        strSQL = strSQL & " And A.病人id=[1]"
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp!病人ID)))
        If mrsInfo.EOF = False Then
            txtPatient.Text = Nvl(mrsInfo!姓名)
            txt险类.Caption = Nvl(mrsInfo!险类名称):
            txt性别 = Nvl(mrsInfo!性别)
            txt年龄 = Nvl(mrsInfo!年龄)
            txtPatient.PasswordChar = ""
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
            '74428：李南春，2014-7-8，病人姓名显示颜色处理
            Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(Trim(txt险类.Caption) = "", txtPatient.ForeColor, vbRed))
        Else
            txt险类.Caption = "": txt性别 = "": txt年龄 = ""
        End If
        
    Else
        Set mrsInfo = Nothing
        txt险类.Caption = "": txt性别 = "": txt年龄 = ""
    End If
    
    Dim lngRow As Long
    With vsRegist
        .Clear 1: .Rows = 2
        If rsTmp.RecordCount <> 0 Then .Rows = rsTmp.RecordCount + 1
        lngRow = 1
        Do While Not rsTmp.EOF
            .TextMatrix(lngRow, .ColIndex("标志")) = lngRow
            .TextMatrix(lngRow, .ColIndex("单据号")) = Nvl(rsTmp!NO)
            .TextMatrix(lngRow, .ColIndex("号别")) = Nvl(rsTmp!号别)
            .TextMatrix(lngRow, .ColIndex("科室")) = Nvl(rsTmp!挂号科室)
            .TextMatrix(lngRow, .ColIndex("挂号时间")) = Nvl(rsTmp!挂号时间)
            .RowData(lngRow) = Val(Nvl(rsTmp!病人ID))
            lngRow = lngRow + 1
            rsTmp.MoveNext
        Loop
        zl_vsGrid_Para_Restore mlngModule, vsRegist, Me.Caption, "挂号单列表", True
        .ColWidth(.ColIndex("标志")) = 285
    End With
    Call txtPatient_GotFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
  
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2009-09-09 15:45:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer
    With vsRegist
        .ColData(.ColIndex("标志")) = "1|1"
        .ColData(.ColIndex("预约单据号")) = "1|0"
    End With
End Sub

Private Sub vsRegist_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsRegist, Me.Caption, "挂号单列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsRegist_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRegist
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub

Private Sub vsRegist_DblClick()
        Call cmdOK_Click
End Sub

Private Sub vsRegist_GotFocus()
    vsRegist.BackColorSel = &H8000000D
End Sub

Private Sub vsRegist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub vsRegist_LostFocus()
    vsRegist.BackColorSel = GRD_LOSTFOCUS_COLORSEL
End Sub
Private Sub vsRegist_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsRegist, Me.Caption, "挂号单列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        txtPatient.Text = strID
        Call GetPatient(Trim(txtPatient.Text))
        IDKind.IDKind = lngPreIDKind
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then mobjICCard.SetEnabled (txtPatient.Text = "")
    End If
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("IC卡号")
        txtPatient.Text = strNO
        If txtPatient.Text <> "" Then
            Call GetPatient(Trim(txtPatient.Text))
        Else
            Call mobjICCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Me.ActiveControl Is txtPatient And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub

Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", Me.txtPatient)
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModule, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    IDKind.ShowPropertySet = InStr(";" & mstrPrivs & ";", "参数设置") > 0
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
End Function

Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
     End Select
End Function
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function
