VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#3.4#0"; "zlIDKind.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiInfoChangeLog 
   Caption         =   "病人基本信息变动日志"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13440
   Icon            =   "frmPatiInfoChangeLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   13440
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraPati 
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   345
      Width           =   13440
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1500
         TabIndex        =   2
         Top             =   270
         Width           =   2340
      End
      Begin VB.CommandButton cmdPati 
         Height          =   360
         Left            =   3840
         Picture         =   "frmPatiInfoChangeLog.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "选择病人(F2)"
         Top             =   270
         Width           =   360
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   855
         TabIndex        =   1
         ToolTipText     =   "快捷键F4"
         Top             =   270
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmPatiInfoChangeLog.frx":6DDC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   10.5
         FontName        =   "宋体"
         IDKind          =   -1
         BackColor       =   -2147483633
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   330
         TabIndex        =   7
         Top             =   345
         Width           =   420
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4320
         TabIndex        =   4
         Top             =   345
         Width           =   630
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6090
         TabIndex        =   5
         Top             =   345
         Width           =   630
      End
      Begin VB.Label lblBirthday 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8355
         TabIndex        =   6
         Top             =   345
         Width           =   1050
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VsgData 
      Height          =   5865
      Left            =   0
      TabIndex        =   8
      Top             =   1185
      Width           =   13425
      _cx             =   23680
      _cy             =   10345
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   5000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPatiInfoChangeLog.frx":6E63
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
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
      AutoSizeMouse   =   0   'False
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
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7605
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiInfoChangeLog.frx":6EC5
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20796
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatiInfoChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------
'65802:刘鹏飞,2013-11-14
'------------------------------------------

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private mlng病人ID As Long
Private mblnNotClick As Boolean
Private mstrPrivs As String

Private Enum VFGDATACOL
    序列 = 0
    病人ID = 1
    变动项目 = 2
    原信息 = 3
    新信息 = 4
    变动时间 = 5
    变动人 = 6
    变动模块 = 7
    变动说明 = 8
End Enum


Public Sub ShowMe(frmParent As Object, ByVal strPrivs As String, Optional ByVal lng病人ID As Long = 0)
'--------------------------------------------------------------------------------------------
'功能:查看病人基本信息变动日志
'参数:
'   frmParent:调用窗体对象
'   strPrivs:权限功能字符串
'   lng病人ID:病人ID<>0则直接提取病人
'--------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs
    mlng病人ID = lng病人ID
    mblnNotClick = False
    
    Me.Show 1, frmParent
End Sub
    
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    Dim objControl As CommandBarControl
    
    Select Case Control.ID
        Case conMenu_File_PrintSet '打印设置
            Call zlPrintSet
        Case conMenu_File_Preview  '预览
            Call OutputList(2)
        Case conMenu_File_Print   '打印
            Call OutputList(1)
        Case conMenu_File_Excel   '输出到Excel
            Call OutputList(3)
        Case conMenu_View_Refresh '刷新
            Call LoadPatiChangeInfo(Val(txtPatient.Tag))
        Case conMenu_View_ToolBar_Button '工具栏
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                cbsMain(i).Visible = Not cbsMain(i).Visible
            Next
            cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '按钮文字
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                For Each objControl In cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '大图标
            cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
            Control.Checked = Not Control.Checked
            cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '状态栏
            staThis.Visible = Not staThis.Visible
            Control.Checked = Not Control.Checked
            cbsMain.RecalcLayout
        Case conMenu_View_Refresh
            
        Case conMenu_Help_Web_Home 'Web上的中联
            Call zlHomePage(hwnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(hwnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call zlMailTo(hwnd)
        Case conMenu_Help_About '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, hwnd, Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '退出
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If staThis.Visible Then Bottom = staThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.fraPati
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
    End With
    
    With VsgData
        .Left = lngLeft: .Top = fraPati.Top + fraPati.Height
        .Width = lngRight - lngLeft
        .Height = lngBottom - .Top
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button '工具栏
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.staThis.Visible
    End Select
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
    Call CreateMobjCard
    Call CreateSquareCardObject(Me, 1101)
     '初始化
    Call IDKind.zlInit(Me, 100, 1101, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    
    If Not gobjSquare.objSquareCard Is Nothing Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    Call InitMainMunus
    
    RestoreWinState Me, App.ProductName
    
    If mlng病人ID <> 0 Then
        txtPatient.Text = "-" & mlng病人ID
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("姓名")
        mblnNotClick = False
        Call txtPatient_KeyPress(vbKeyReturn)
    Else
        txtPatient.Text = ""
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
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
    SaveFlexState VsgData, App.ProductName & "\" & Me.Name
    SaveWinState Me, App.ProductName
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
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call InitVsfDate(False)
    Else
        txtPatient.PasswordChar = ""
        txtPatient.IMEMode = 0
        Call LoadPatiChangeInfo(Val(txtPatient.Tag))
    End If
    FindPati = True
End Function

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
'功能：读取病人信息
    Dim lng卡类别ID As Long, lng病人ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnDo As Boolean
    Dim blnHavePassWord As Boolean
    Dim strPassWord As String, strErrMsg As String
    Dim strCard As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select A.病人ID,A.姓名,A.性别,A.年龄,A.出生日期,A.病人类型,A.险类" & _
        " From 病人信息 A" & _
        " Where A.停用时间 is NULL"
        
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
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    blnDo = Not rsTmp.EOF
    
    If blnDo Then
        txtPatient.Tag = rsTmp!病人ID
        txtPatient.Text = rsTmp!姓名
        '74426:李南春,2014-7-9,病人姓名显示颜色处理
        Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), Me.ForeColor, vbRed))
        lblSex.Caption = "性别：" & Nvl(rsTmp!性别)
        lblAge.Caption = "年龄：" & Nvl(rsTmp!年龄)
        lblBirthday.Caption = "出生日期：" & Format(Nvl(rsTmp!出生日期), "YYYY-MM-DD HH:mm")
        
        GetPatient = True
    Else
NotFoundPati:
        txtPatient.Tag = ""
        txtPatient.Text = ""
        txtPatient.ForeColor = Me.ForeColor
        lblSex.Caption = "性别："
        lblAge.Caption = "年龄："
        lblBirthday.Caption = "出生日期："
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

Private Sub CreateMobjCard()
    '创建卡部件
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
End Sub

Private Sub InitVsfDate(Optional blnSet As Boolean)
'功能:初始化病人信息变动日志表格

    Dim strHead As String
    Dim i As Integer

    strHead = "序列,4,500|病人ID,1,0|变动项目,4,1000|原信息,4,1500|新信息,4,1500|变动时间,1,2000|变动人,1,1000|变动模块,1,1200|变动说明,1,4000"
    
    With VsgData
        .Redraw = False
        .Clear
        .Rows = 2
        .Cols = UBound(Split(strHead, "|")) + 1
        
        .MergeCells = flexMergeRestrictColumns
        .MergeCellsFixed = flexMergeFree
        
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Or blnSet Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
        Next
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        If Not Visible Or blnSet Then Call RestoreFlexState(VsgData, App.ProductName & "\" & Me.Name)
        .FixedCols = 1
        .FixedRows = 1
        .ColHidden(病人ID) = True

        .RowHeight(0) = 320
        .RowHeight(1) = 300
        '恢复上次行
        .Row = 1
        .Col = 1:
        .Redraw = True
    End With
    staThis.Panels(2).Text = "请先确定病人"
End Sub

Private Sub LoadPatiChangeInfo(ByVal lng病人ID As Long)
'功能:提取并加载病人信息变动日期
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngRow As Long
    Dim strChangeDate As String
    Dim strTmp As String, lngNum As Long '基数
    
    On Error GoTo Errhand
    
    If lng病人ID = 0 Then
        Call InitVsfDate(True)
        Exit Sub
    End If
    
    strSQL = " Select 病人id, 变动项目, 原信息, 新信息, 变动时间, 变动人, 变动模块, 说明 变动说明" & vbNewLine & _
            " From 病人信息变动" & vbNewLine & _
            " Where 病人id = [1]" & vbNewLine & _
            " Order By 变动时间, 变动项目 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人信息变动", lng病人ID)
    With VsgData
        Call InitVsfDate(True)
        lngNum = 0
        strChangeDate = ""
        .Redraw = flexRDNone
        Do While Not rsTmp.EOF
            lngRow = rsTmp.AbsolutePosition
            If Format(strChangeDate, "YYYY-MM-DD HH:mm:ss") <> Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss") Then
                lngNum = lngNum + 1
                strChangeDate = Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss")
                If lngNum Mod 2 = 1 Then
                    strTmp = ""
                Else
                    strTmp = " "
                End If
            End If
            
            If lngRow > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(lngRow, 序列) = lngNum
            .TextMatrix(lngRow, 病人ID) = Nvl(rsTmp!病人ID)
            .TextMatrix(lngRow, 变动项目) = Nvl(rsTmp!变动项目)
            If Nvl(rsTmp!变动项目) = "出生日期" Then
                .TextMatrix(lngRow, 原信息) = Format(Nvl(rsTmp!原信息), "YYYY-MM-DD HH:mm")
                .TextMatrix(lngRow, 新信息) = Format(Nvl(rsTmp!新信息), "YYYY-MM-DD HH:mm")
            Else
                .TextMatrix(lngRow, 原信息) = Nvl(rsTmp!原信息)
                .TextMatrix(lngRow, 新信息) = Nvl(rsTmp!新信息)
            End If
            .TextMatrix(lngRow, 变动时间) = Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss") & strTmp
            .TextMatrix(lngRow, 变动人) = Nvl(rsTmp!变动人) & strTmp
            .TextMatrix(lngRow, 变动模块) = Nvl(rsTmp!变动模块) & strTmp
            .TextMatrix(lngRow, 变动说明) = Nvl(rsTmp!变动说明) & strTmp
        rsTmp.MoveNext
        Loop
        
        .WordWrap = False
        .MergeCol(序列) = False
        .MergeCol(变动时间) = False
        .MergeCol(变动人) = False
        .MergeCol(变动模块) = False
        .MergeCol(变动说明) = False
        
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        .RowHeight(0) = 320
        
        .MergeCol(序列) = True
        .MergeCol(变动时间) = True
        .MergeCol(变动人) = True
        .MergeCol(变动模块) = True
        .MergeCol(变动说明) = True
        
        For lngRow = .FixedRows To .Rows - 1
            If .RowHeight(lngRow) < 300 Then .RowHeight(lngRow) = 300
        Next lngRow
        .Row = 1
        .Redraw = flexRDDirect
    End With
    
    If rsTmp.RecordCount > 0 Then
        staThis.Panels(2).Text = "病人共发生了　" & lngNum & "　次基本信息变动"
    Else
        staThis.Panels(2).Text = "无基本信息变动记录"
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub InitMainMunus()
    Dim objBar As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
        
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    Set Me.cbsMain.Icons = zlCommFun.GetPubIcons
    With Me.cbsMain.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With
    
    '菜单定义
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        .Add xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…"
        .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&V)").BeginGroup = True
        .Add xtpControlButton, conMenu_File_Print, "打印(&P)"
        .Add xtpControlButton, conMenu_File_Excel, "输出到Excel(&L)"
        .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)").BeginGroup = True
    End With


    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
       Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With
    
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的中联")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False    '固有
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With

    
    '查找项特殊处理
    '-----------------------------------------------------
'    '主菜单右侧的查找
'    With cbsMain.ActiveMenuBar.Controls
'        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
'        objCustom.Handle = picFind.hwnd
'        objCustom.flags = xtpFlagRightAlign
'        IDKind.BackColor = picFind.BackColor
'    End With

    '工具栏定义
    '-----------------------------------------------------
    Set objBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
    End With
    
     For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '快键绑定
    With cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    cbsMain.RecalcLayout
End Sub


Private Sub OutputList(bytStyle As Byte)
'功能：输入出病人变动日志
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, lngRow As Long
    
    lngRow = VsgData.Row
    
    '表头
    objOut.Title.Text = "病人基本信息变动日志"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    objRow.Add "姓名：" & txtPatient.Text
    objRow.Add lblSex.Caption
    objRow.Add lblAge.Caption
    objRow.Add lblBirthday.Caption
    objOut.UnderAppRows.Add objRow
    
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    VsgData.Redraw = False
    Set objOut.Body = VsgData
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    VsgData.Row = lngRow
    VsgData.Redraw = True
End Sub

Private Sub VsgData_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long
    With VsgData
        .Redraw = flexRDNone
        .WordWrap = False
        .MergeCol(序列) = False
        .MergeCol(变动时间) = False
        .MergeCol(变动人) = False
        .MergeCol(变动模块) = False
        .MergeCol(变动说明) = False
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False

        .MergeCol(序列) = True
        .MergeCol(变动时间) = True
        .MergeCol(变动人) = True
        .MergeCol(变动模块) = True
        .MergeCol(变动说明) = True

        For lngRow = .FixedRows To .Rows - 1
            If .RowHeight(lngRow) < 300 Then .RowHeight(lngRow) = 300
        Next lngRow
        .Redraw = flexRDDirect
    End With
End Sub

