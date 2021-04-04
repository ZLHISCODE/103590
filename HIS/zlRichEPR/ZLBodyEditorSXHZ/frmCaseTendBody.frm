VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseTendBody 
   Caption         =   "体温作图"
   ClientHeight    =   7350
   ClientLeft      =   180
   ClientTop       =   450
   ClientWidth     =   10740
   Icon            =   "frmCaseTendBody.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picCustom 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2265
      ScaleHeight     =   300
      ScaleWidth      =   1965
      TabIndex        =   2
      Top             =   5220
      Width           =   1965
      Begin VB.CommandButton cmd 
         Height          =   300
         Left            =   1665
         Picture         =   "frmCaseTendBody.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   300
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1665
      End
   End
   Begin zl9BodyEditorSXHZ.usrBodyEditor BodyEdit 
      Height          =   4350
      Left            =   435
      TabIndex        =   0
      Top             =   615
      Width           =   6375
      _extentx        =   11245
      _extenty        =   7673
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6990
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBody.frx":6AD8
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16034
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmCaseTendBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'局部变量申明区域

Private mrsParam As New ADODB.Recordset
Private mstrSQL As String
Private mblnStartUp As Boolean
Private mblnChildForm As Boolean
Private mblnOK As Boolean
Private mfrmMain As Object
Private mblnChanged As Boolean
Private mcbr查看 As CommandBarControl
Private mstr体温部位 As String
Private mstr呼吸方式 As String
Private mstr脉搏 As String
Private mcbrMenuBar曲线 As CommandBarControl
Private mcbrMenuBar部位 As CommandBarControl
Private mcbrMenuBar编辑 As CommandBarControl
Private mcbrToolBar As CommandBar
Private mint就诊卡号码长度 As Integer
Private mstrSvr姓名 As String
Private mrsPatient As ADODB.Recordset
Private mlngRowNum As Long
Private mstrFindKey As String
Private mobjFindKey As CommandBarControl
Private mstrPrivs As String
Private mblnShowing As Boolean

Public Event AfterPrint()

'######################################################################################################################
'自定义函数、过程区域

Public Function ShowEdit(ByVal frmMain As Object, strParam As String, Optional ByVal bytMode As Byte = 1, Optional strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim blnShowing As Boolean
    
    mblnStartUp = True
    mblnChanged = False
    mstrPrivs = strPrivs
    mstr体温部位 = "腋温"
    mstr呼吸方式 = "自主呼吸"
    mstr脉搏 = ""
    
    blnShowing = mblnShowing
    
    mblnShowing = True
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    If blnShowing Then
        If Val(varParam(0)) = Val(mrsParam("病人id").Value) Or Val(varParam(1)) = Val(mrsParam("主页id").Value) And Val(mrsParam("科室id").Value) = Val(varParam(2)) Then
            Call ShowWindow(Me.hWnd, SW_RESTORE)
            Call BringWindowToTop(Me.hWnd)
            Exit Function
        End If
    End If
    
    Set mfrmMain = frmMain

    '参数格式：病人ID;主页ID;病区ID;出院;编辑;婴儿
    
    '初始化参数
    Set mrsParam = New ADODB.Recordset
    Call CreateParam(mrsParam, "病人id", adBigInt)
    Call CreateParam(mrsParam, "主页id", adBigInt)
    Call CreateParam(mrsParam, "病区id", adBigInt)
    Call CreateParam(mrsParam, "科室id", adBigInt)
    Call CreateParam(mrsParam, "出院", adTinyInt)
    Call CreateParam(mrsParam, "婴儿", adTinyInt)
    Call CreateParam(mrsParam, "编辑", adTinyInt)
    Call CreateParam(mrsParam, "护理等级", adTinyInt)
    Call CreateParam(mrsParam, "出院开始日期", adVarChar, 30)
    Call CreateParam(mrsParam, "出院结束日期", adVarChar, 30)
    Call CreateParam(mrsParam, "在院病人", adTinyInt)
    Call CreateParam(mrsParam, "出院病人", adTinyInt)
    Call CreateParam(mrsParam, "待入科病人", adTinyInt)
    Call CreateParam(mrsParam, "转出病人", adTinyInt)
    Call CreateParam(mrsParam, "转出天数", adTinyInt)

    mrsParam.Open
    mrsParam.AddNew
    
    mrsParam("婴儿").Value = 0
    mrsParam("出院").Value = 0
    mrsParam("编辑").Value = 0
    
    mrsParam("病人id").Value = Val(varParam(0))
    mrsParam("主页id").Value = Val(varParam(1))
    mrsParam("病区id").Value = Val(varParam(2))
    mrsParam("科室id").Value = Val(varParam(2))
    
    If UBound(varParam) >= 3 Then mrsParam("出院").Value = Val(varParam(3))
    If UBound(varParam) >= 4 Then mrsParam("编辑").Value = Val(varParam(4))
    If UBound(varParam) >= 5 Then mrsParam("婴儿").Value = Val(varParam(5))
    
    
    '出院开始日期;出院结束日期;在院病人;出院病人;转出病人;转出天数
    '------------------------------------------------------------------------------------------------------------------
    strPar = zlDatabase.GetPara("病人显示范围", glngSys, 1262, "10000")
    mrsParam("在院病人").Value = Val(Mid(strPar, 1, 1))
    mrsParam("出院病人").Value = Val(Mid(strPar, 2, 1))
    mrsParam("转出病人").Value = Val(Mid(strPar, 4, 1))
    On Error Resume Next
    mrsParam("待入科病人").Value = Val(Mid(strPar, 5, 1))
    On Error GoTo 0
    
    mrsParam("转出天数").Value = Val(zlDatabase.GetPara("最近转出天数", 7))
    
    Dim curDate As Date
    Dim intDay As Integer
    
    curDate = zlDatabase.Currentdate
    intDay = Val(zlDatabase.GetPara("出院病人结束间隔", glngSys, 1262, 7))
    mrsParam("出院结束日期").Value = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
    intDay = Val(zlDatabase.GetPara("出院病人开始间隔", glngSys, 1262, 30))
    mrsParam("出院开始日期").Value = Format(CDate(mrsParam("出院结束日期").Value) - intDay, "yyyy-MM-dd 00:00:00")
    
    If blnShowing = False Then Call InitMenuBar
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 出院科室ID from 病案主页 Where 病人id=[1] And 主页id=[2] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value))
    If rs.BOF = False Then
        mrsParam("科室id").Value = Val(zlCommFun.NVL(rs("出院科室ID").Value))
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 姓名 from 病人信息 Where 病人id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value))
    If rs.BOF = False Then
        txt.Text = zlCommFun.NVL(rs("姓名").Value)
        txt.Tag = ""
    End If
    
    '就诊卡长度
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 卡号长度 from 医疗卡类别 where 名称='就诊卡' and 是否固定=1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "就诊卡")
    If rs.BOF = False Then
        mint就诊卡号码长度 = Val(zlCommFun.NVL(rs("卡号长度").Value))
    Else
        mint就诊卡号码长度 = 7
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If ReadPatient = False Then
        mblnStartUp = False
        Unload Me
        Exit Function
    End If
    
    mrsPatient.Filter = ""
    mrsPatient.Filter = "病人id=" & Val(mrsParam("病人id").Value)
    If mrsPatient.RecordCount > 0 Then mlngRowNum = Val(mrsPatient("ID").Value)
    mrsPatient.Filter = ""
    
    Set BodyEdit.ParentForm = Me
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        mblnStartUp = False
        Unload Me
        Exit Function
    End If
    
    mblnStartUp = False
    
    If blnShowing = False Then
        Hook Me.hWnd
        
        If bytMode = 1 Then
            Me.Show , mfrmMain
        Else
            Me.Show 1, mfrmMain
        End If
        
        ShowEdit = mblnChanged
    End If
    
End Function

Public Function zlInit() As Boolean

    mblnChildForm = True

'    Call InitMenuBar

End Function

Public Function zlPrintBody(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDevice As String) As Long
    '入参:1-预览,2-打印
    '返回值:0-失败;1-成功;2-打印
    gblnPrinted = False
    
'    If bytMode = 1 Then
'        zlPrintBody = PrintData(2, strPrintDevice)
'    Else
'        zlPrintBody = PrintData(1, strPrintDevice)
'    End If
    
    Call PrintData(IIf(bytMode = 1, 2, 1), strPrintDevice)
    zlPrintBody = IIf(gblnPrinted, 2, 1)
End Function

Public Function zlRefresh(strParam As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    
    mblnChildForm = True
    stbThis.Visible = Not mblnChildForm
    picCustom.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.RecalcLayout
    
    mblnStartUp = True
    mblnChanged = False
'    mstrPrivs = strPrivs
    mstr体温部位 = "腋温"
    mstr呼吸方式 = "自主呼吸"
    mstr脉搏 = ""
    
'    Set mfrmMain = frmMain
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    '参数格式：病人ID;主页ID;病区ID;出院;编辑;婴儿
    
    '初始化参数
    Set mrsParam = New ADODB.Recordset
    Call CreateParam(mrsParam, "病人id", adBigInt)
    Call CreateParam(mrsParam, "主页id", adBigInt)
    Call CreateParam(mrsParam, "病区id", adBigInt)
    Call CreateParam(mrsParam, "科室id", adBigInt)
    Call CreateParam(mrsParam, "出院", adTinyInt)
    Call CreateParam(mrsParam, "婴儿", adTinyInt)
    Call CreateParam(mrsParam, "编辑", adTinyInt)
    Call CreateParam(mrsParam, "护理等级", adTinyInt)
    Call CreateParam(mrsParam, "出院开始日期", adVarChar, 30)
    Call CreateParam(mrsParam, "出院结束日期", adVarChar, 30)
    Call CreateParam(mrsParam, "在院病人", adTinyInt)
    Call CreateParam(mrsParam, "出院病人", adTinyInt)
    Call CreateParam(mrsParam, "待入科病人", adTinyInt)
    Call CreateParam(mrsParam, "转出病人", adTinyInt)
    Call CreateParam(mrsParam, "转出天数", adTinyInt)
    
    mrsParam.Open
    mrsParam.AddNew
    
    mrsParam("婴儿").Value = 0
    mrsParam("出院").Value = 0
    mrsParam("编辑").Value = 0
    
    mrsParam("病人id").Value = Val(varParam(0))
    mrsParam("主页id").Value = Val(varParam(1))
    mrsParam("病区id").Value = Val(varParam(2))
    mrsParam("科室id").Value = Val(varParam(2))
    
    If UBound(varParam) >= 3 Then mrsParam("出院").Value = Val(varParam(3))
    If UBound(varParam) >= 4 Then mrsParam("编辑").Value = Val(varParam(4))
    If UBound(varParam) >= 5 Then mrsParam("婴儿").Value = Val(varParam(5))
    
    
    '出院开始日期;出院结束日期;在院病人;出院病人;转出病人;转出天数
    '------------------------------------------------------------------------------------------------------------------
    strPar = zlDatabase.GetPara("病人显示范围", glngSys, 1262, "10000")
    mrsParam("在院病人").Value = Val(Mid(strPar, 1, 1))
    mrsParam("出院病人").Value = Val(Mid(strPar, 2, 1))
    mrsParam("转出病人").Value = Val(Mid(strPar, 4, 1))
    On Error Resume Next
    mrsParam("待入科病人").Value = Val(Mid(strPar, 5, 1))
    On Error GoTo 0
    
    mrsParam("转出天数").Value = Val(zlDatabase.GetPara("最近转出天数", 7))
    
    Dim curDate As Date
    Dim intDay As Integer
    
    curDate = zlDatabase.Currentdate
    intDay = Val(zlDatabase.GetPara("出院病人结束间隔", glngSys, 1262, 7))
    mrsParam("出院结束日期").Value = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
    intDay = Val(zlDatabase.GetPara("出院病人开始间隔", glngSys, 1262, 30))
    mrsParam("出院开始日期").Value = Format(CDate(mrsParam("出院结束日期").Value) - intDay, "yyyy-MM-dd 00:00:00")

'    Call InitMenuBar
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 出院科室ID from 病案主页 Where 病人id=[1] And 主页id=[2] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value))
    If rs.BOF = False Then
        mrsParam("科室id").Value = Val(zlCommFun.NVL(rs("出院科室ID").Value))
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 姓名 from 病人信息 Where 病人id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value))
    If rs.BOF = False Then
        txt.Text = zlCommFun.NVL(rs("姓名").Value)
        txt.Tag = ""
    End If
    
    '就诊卡长度
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 卡号长度 from 医疗卡类别 where 名称='就诊卡' and 是否固定=1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "就诊卡")
    If rs.BOF = False Then
        mint就诊卡号码长度 = Val(zlCommFun.NVL(rs("卡号长度").Value))
    Else
        mint就诊卡号码长度 = 7
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If ReadPatient = False Then
        mblnStartUp = False
'        Unload Me
        Exit Function
    End If
    
    mrsPatient.Filter = ""
    mrsPatient.Filter = "病人id=" & Val(mrsParam("病人id").Value)
    If mrsPatient.RecordCount > 0 Then mlngRowNum = Val(mrsPatient("ID").Value)
    mrsPatient.Filter = ""
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        mblnStartUp = False
'        Unload Me
        Exit Function
    End If
    
    mblnStartUp = False
    
'    Hook Me.hWnd
        
    zlRefresh = True
    
End Function

Private Function ShowTxtSelDialog(ByVal frmParent As Object, _
                                    ByVal objTXT As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rs As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False, _
                                    Optional strInitKey As String = "", _
                                    Optional ByVal WinStyle As Byte = 3, _
                                    Optional ByVal blnSort As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:打开树型+列表结构
    '返回:出错返回2;成功返回1;取消返回0
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim objPoint As POINTAPI
        
    
    On Error GoTo errHand
    
    If rs.BOF Then Exit Function
    
    Call ClientToScreen(objTXT.hWnd, objPoint)
                
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTXT.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
    
    If frmSelectDialog.ShowSelect(frmParent, WinStyle, rs, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTXT.Height, strInitKey, strSavePath, , False, blnMuliSel, , blnSort) Then
                            
        Set rsResult = rs
        ShowTxtSelDialog = True
        
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function OpenPatientMap() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strParam As String
    
    mstrSvr姓名 = txt.Text
    
    mrsParam("护理等级").Value = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If rs.BOF = False Then mrsParam("护理等级").Value = zlCommFun.NVL(rs("护理等级"), 3)
    
    '初始化曲线菜单
    If InitBodyLine = False Then Exit Function
    
    '参数：病人ID,主页ID,病区ID,科室ID,出院标志;编辑标志;婴儿
    strParam = Val(mrsParam("病人id")) & ";" & Val(mrsParam("主页id")) & ";" & Val(mrsParam("病区id")) & ";" & Val(mrsParam("出院")) & ";" & Val(mrsParam("编辑").Value) & ";" & Val(mrsParam("婴儿").Value)
    If Not BodyEdit.zlMenuClick("初始数据", strParam) Then Exit Function
'    If InitBody(Val(mrsParam("病人id")), Val(mrsParam("主页id")), Val(mrsParam("病区id"))) = False Then Exit Function
        
    OpenPatientMap = True
    
End Function

Private Function ReadPatient() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strParam As String
    
    '在院和出院病人:出院病人可能已有多次住院
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("在院病人").Value) <> 0 Or Val(mrsParam("出院病人").Value) <> 0 Or Val(mrsParam("待入科病人").Value) <> 0 Then
        gstrSQL = _
            "Select Decode(B.出院日期,NULL,Decode(B.状态,3,2,1),Decode(B.出院方式,'死亡',4,3)) as 排序," & _
            " Decode(B.出院日期,NULL,Decode(B.状态,3,'预出院病人','在院病人'),Decode(B.出院方式,'死亡','死亡病人','出院病人')) as 类型," & _
            " A.病人ID,B.主页ID,B.住院号,A.门诊号,A.姓名,A.性别,A.年龄,C.名称 as 科室,B.住院医师," & _
            " B.出院病床 as 床号,B.费别,B.入院日期,B.出院日期,B.状态,B.险类,A.就诊卡号" & _
            " From 病人信息 A,病案主页 B,部门表 C" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And ([6]=1 Or Nvl(B.状态,0)<>1) And B.出院科室ID=C.ID" & _
            " And B.当前病区ID=[1] And ([4]<>0 And B.出院日期 is NULL Or [5]<>0 And B.出院日期 Between [2] And [3]) "
    End If
    
    '转出病人:在院,医生和床号显示本科转出前的
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("转出病人").Value) <> 0 Then
        gstrSQL = gstrSQL & IIf(gstrSQL <> "", " Union All ", "") & _
            "Select Distinct 5 as 排序,'转出病人' as 类型," & _
            " A.病人ID,B.主页ID,B.住院号,A.门诊号,A.姓名,A.性别,A.年龄,D.名称 as 科室,C.经治医师 as 住院医师," & _
            " C.床号,B.费别,B.入院日期,B.出院日期,B.状态,B.险类,A.就诊卡号" & _
            " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And C.科室ID=D.ID" & _
            " And Nvl(B.状态,0)=0 And B.出院日期 is NULL And B.当前病区ID<>[1]" & _
            " And B.病人ID=C.病人ID And B.主页ID=C.主页ID And C.病区ID=[1]" & _
            " And C.终止原因=3 And C.终止时间 Between Sysdate-[7] And Sysdate "
    End If
    gstrSQL = gstrSQL & " Order by 排序,床号,主页ID Desc"
    gstrSQL = "Select RowNum As ID,1 As 末级,A.* From (" & gstrSQL & ") A"
    
    If Val(mrsParam("编辑").Value) = 1 Then
        Set mrsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
                                                                Val(mrsParam("病区id").Value), _
                                                                CDate(Format(mrsParam("出院开始日期").Value, "yyyy-MM-dd 00:00:00")), _
                                                                CDate(Format(mrsParam("出院结束日期").Value, "yyyy-MM-dd 23:59:59")), _
                                                                Val(mrsParam("在院病人").Value), _
                                                                0, _
                                                                Val(mrsParam("待入科病人").Value), _
                                                                Val(mrsParam("转出天数").Value))
    Else
        Set mrsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
                                                                Val(mrsParam("病区id").Value), _
                                                                CDate(Format(mrsParam("出院开始日期").Value, "yyyy-MM-dd 00:00:00")), _
                                                                CDate(Format(mrsParam("出院结束日期").Value, "yyyy-MM-dd 23:59:59")), _
                                                                Val(mrsParam("在院病人").Value), _
                                                                Val(mrsParam("出院病人").Value), _
                                                                Val(mrsParam("待入科病人").Value), _
                                                                Val(mrsParam("转出天数").Value))
    End If
    
    ReadPatient = True
    
End Function

Private Function PrintData(ByVal bytMode As Byte, Optional ByVal strPrintDevice As String = "") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnCur As Boolean
    Dim lngBeginY As Long
    Dim intBeginPage As Integer
    Dim intPrintRange As Integer
    
    '传入了打印机名称,说明是批量打印,自动从第1页开始打印,不进行任何询问
    '返回:0-取消,2-预览,1-打印
    
    frmCaseTendBodyPrintSet.cmdPrint.Visible = (bytMode = 1)
    frmCaseTendBodyPrintSet.cmdPreview.Visible = (bytMode = 2)
    
    If strPrintDevice = "" Then
        bytMode = frmCaseTendBodyPrintSet.PrintSet(Me, True, intPrintRange, lngBeginY, intBeginPage, mstrPrivs)
    Else
        bytMode = 2
        intPrintRange = 2
    End If
    If bytMode = 0 Then Exit Function
    If intBeginPage <= 0 Then intBeginPage = -1
            
    Select Case bytMode
    Case 2  '打印
        Call BodyEdit.PrintState(intPrintRange, True, lngBeginY, intBeginPage, strPrintDevice)
    Case 1  '预览
        Call BodyEdit.PrintState(intPrintRange, False, lngBeginY, intBeginPage, strPrintDevice)
    End Select

    
End Function

Private Function InitBodyLine() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    
    On Error GoTo errHand
    
    If mcbrMenuBar曲线 Is Nothing Then
        InitBodyLine = True
        Exit Function
    End If
    
    mstrSQL = "SELECT A.记录名,A.项目序号 FROM 体温记录项目 A,护理记录项目 B " & _
            "WHERE A.记录法 =1 And A.项目序号=B.项目序号 AND B.护理等级>=[1]  And Nvl(b.应用方式,0)=1 " & _
            "ORDER BY A.排列序号"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(mrsParam("护理等级").Value))
    If rsTmp.BOF Then
        ShowSimpleMsg "无体温单操作曲线项目，请在护理项目中设置！"
        Exit Function
    End If

    Do While Not rsTmp.EOF

        Set cbrItem = mcbrMenuBar曲线.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendOther, zlCommFun.NVL(rsTmp("记录名")), -1, False)
        cbrItem.Parameter = rsTmp.AbsolutePosition

        rsTmp.MoveNext
    Loop

    InitBodyLine = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'Private Function InitBody(ByVal lng病人id As Long, ByVal lng主页id As Long, ByVal lng病区id As Long) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim strSQL As String
'    Dim RS As New ADODB.Recordset
'    Dim rsTmp As New ADODB.Recordset
'    Dim cbrItem As CommandBarControl
'    Dim intCount As Integer
'    Dim strDateFrom As String
'    Dim strDateTo As String
'    Dim strEnterDate As String
'    Dim intCol As Integer
'    Dim strCaption As String
'    Dim strParameter As String
'    Dim strNow As String
'    Dim strCut As String
'    Dim lngLoop As Long
'    Dim strTmp As String
'    Dim lnglast科室id As Long
'
'    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
'    '删除操作页面菜单项
'
'    mcbrToolBar页面.Delete
'    mcbrMenuBar页面.Delete
'
'    Set mcbrToolBar页面 = mcbrToolBar.Controls.Add(xtpControlPopup, conMenu_Edit_NewItem, "页面", 5):  mcbrToolBar页面.BeginGroup = True
'    mcbrToolBar页面.IconId = conMenu_Edit_Modify
'    mcbrToolBar页面.Style = xtpButtonIconAndCaption
'
'    Set mcbrMenuBar页面 = mcbr查看.CommandBar.Controls.Add(xtpControlPopup, conMenu_Edit_NewParent, "体温页面(&P)", 3)
'    mcbrMenuBar页面.BeginGroup = True
'
'    '
'    '------------------------------------------------------------------------------------------------------------------
'    strSQL = "Select 入院时间, 出院时间, 1 + Round((b.出院时间 - b.入院时间) / 7) As 页数" & vbNewLine & _
'                "  from (Select Min(开始时间) as 入院时间," & vbNewLine & _
'                "               Max(Nvl(终止时间, Sysdate)) as 出院时间" & vbNewLine & _
'                "          From 病人变动记录" & vbNewLine & _
'                "         Where 开始时间 is Not Null And 病人ID = [1] And 主页ID = [2]) b"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人id, lng主页id)
'    If rsTmp.BOF Then
'        MsgBox "无病人本次住院记录！", vbExclamation, gstrSysName
'        Exit Function
'    End If
'
'    '
'    '------------------------------------------------------------------------------------------------------------------
'    strSQL = "Select 1 + Round((a.开始时间 - b.入院时间) / 7) As 开始页码,1 + Round((a.终止时间 - b.入院时间) / 7) As 结束页码,b.入院时间," & vbNewLine & _
'                "       病区id,c.名称," & vbNewLine & _
'                "       开始时间," & vbNewLine & _
'                "       终止时间" & vbNewLine & _
'                "  from (Select 病区id," & vbNewLine & _
'                "               Min(开始时间) as 开始时间," & vbNewLine & _
'                "               Max(Nvl(终止时间, Sysdate)) as 终止时间" & vbNewLine & _
'                "          From 病人变动记录" & vbNewLine & _
'                "         Where 开始时间 is Not Null And 病人ID = [1] And 主页ID = [2]" & vbNewLine & _
'                "         Group by 病区id) a," & vbNewLine & _
'                "       (Select Min(开始时间) as 入院时间" & vbNewLine & _
'                "          From 病人变动记录" & vbNewLine & _
'                "         Where 开始时间 is Not Null And 病人ID = [1] And 主页ID = [2]) b,部门表 c Where c.ID=a.病区id " & vbNewLine & _
'                " order by a.开始时间"
'    Set RS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人id, lng主页id)
'
'    strEnterDate = Format(rsTmp!入院时间, "yyyy-MM-dd HH:mm:ss")
'    For lngLoop = 0 To rsTmp("页数").Value - 1
'
'        strDateFrom = Format(rsTmp("入院时间").Value + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"
'        strDateTo = Format(rsTmp("入院时间").Value + 7 * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
'        If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
'            strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
'        End If
'
'        If strDateFrom < Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
'
'            If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
'            If strDateTo > Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss")
'
'            RS.Filter = ""
'            RS.Filter = "开始页码<=" & lngLoop + 1 & " And 结束页码>=" & lngLoop + 1
'            If RS.RecordCount > 0 Then RS.MoveFirst
'            For intCol = 1 To RS.RecordCount
'
'                If strDateFrom < Format(RS("开始时间").Value, "yyyy-MM-dd HH:mm:ss") Then
'                    strTmp = Format(RS("开始时间").Value, "yyyy-MM-dd HH:mm:ss")
'                Else
'                    strTmp = strDateFrom
'                End If
'
'                If strDateTo > Format(RS("终止时间").Value, "yyyy-MM-dd HH:mm:ss") Then
'                    strCaption = Format(RS("终止时间").Value, "yyyy-MM-dd HH:mm:ss")
'                Else
'                    strCaption = strDateTo
'                End If
'
'                strCaption = Format(strTmp, "yyyy-MM-dd") & "～" & Format(strCaption, "yyyy-MM-dd")
'                strCaption = "第" & lngLoop + 1 & "页：" & strCaption & "(" & RS("名称").Value & ")"
'
'                Set cbrItem = mcbrMenuBar页面.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
'
'                '入院时间;科室id;开始时间;结束时间;
'                cbrItem.Parameter = strEnterDate & ";" & RS!病区ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop
'
'                Set cbrItem = mcbrToolBar页面.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
'                cbrItem.Parameter = strEnterDate & ";" & RS!病区ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop
'
'                lnglast科室id = RS("病区ID").Value
'
'                RS.MoveNext
'
'                strParameter = cbrItem.Parameter
'            Next
'        End If
'
'    Next
'
'    If strParameter <> "" Then Call BodyEdit.zlMenuClick("装载数据", strParameter)
'
'    InitBody = True
'End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    
    cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&E)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
       
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存数据(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "恢复数据(&R)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    Set mcbrMenuBar编辑 = cbrMenuBar
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Notify, "设定开始日期(&B)")
        
        Set mcbrMenuBar曲线 = .Add(xtpControlPopup, conMenu_Edit_Modify, "操作曲线(&D)")
        mcbrMenuBar曲线.BeginGroup = True
        mcbrMenuBar曲线.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_SendOther, "无", -1, False
        
        Set cbrPop = .Add(xtpControlPopup, conMenu_Edit_Append, "特殊处理(&S)")
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 1, "失禁或假肛(&1)", -1, False): cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 2, "灌肠(&2)", -1, False):  cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 3, "灌肠后排泄(&3)", -1, False):  cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 4, "导尿(&4)", -1, False):   cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 5, "保留导尿(&5)", -1, False):   cbrControl.IconId = 1
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "添加项目(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "删除项目(&R)")
        
        '
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "设置记录(&E)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "清除记录(&U)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "设置手术/分娩(&W)"): cbrControl.IconId = 1
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "清除手术/分娩(&C)"): cbrControl.IconId = 1
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "复试合格(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "取消复试(&B)")
        
        Set cbrPop = .Add(xtpControlPopup, conMenu_View_ToolBar, "自动获取(&A)"): cbrPop.BeginGroup = True: cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Price, "获取饮入(&1)", -1, False): cbrControl.Parameter = "饮入": cbrControl.IconId = 1
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Price, "获取手术/分娩(&2)", -1, False): cbrControl.Parameter = "手术": cbrControl.IconId = 1
        
    End With

    Set mcbr查看 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    With mcbr查看.CommandBar.Controls
                
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
                
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
'        Set mcbrMenuBar页面 = .Add(xtpControlPopup, conMenu_Edit_NewParent, "体温页面(&P)")
'        mcbrMenuBar页面.BeginGroup = True
        
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."):
        cbrControl.BeginGroup = True
    End With
    
   
    '主菜单右侧的查找
    '------------------------------------------------------------------------------------------------------------------
    cbsThis.ActiveMenuBar.SetIconSize 16, 16
    
    mstrFindKey = Trim(zlDatabase.GetPara("查找方法", glngSys, 1255, "床  号"))
    If mstrFindKey = "" Then mstrFindKey = "床  号"
        
    Set mobjFindKey = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.ToolTipText = "快捷键:F4"
    mobjFindKey.Style = xtpButtonIconAndCaption
    mobjFindKey.flags = xtpFlagRightAlign
    
    Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.床  号"): cbrControl.Parameter = "床  号"
    Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.住院号"): cbrControl.Parameter = "住院号"
    Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&3.就诊卡"): cbrControl.Parameter = "就诊卡"

    Set cbrCustom = cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picCustom.hWnd
    txt.ToolTipText = "查找病人(F3)"
    cbrCustom.flags = xtpFlagRightAlign

    Set cbrControl = cbsThis.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一病人")
    cbrControl.ToolTipText = "前一病人(Ctrl+Left)"
    cbrControl.flags = xtpFlagRightAlign
    cbrControl.Style = xtpButtonIcon

    Set cbrControl = cbsThis.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一病人")
    cbrControl.ToolTipText = "后一病人(Ctrl+Right)"
    cbrControl.flags = xtpFlagRightAlign
    cbrControl.Style = xtpButtonIcon
    
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("标准", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "恢复")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    
    '定位工具栏
    '------------------------------------------------------------------------------------------------------------------
    
    For Each cbrControl In mcbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FALT, Asc("1"), conMenu_Edit_Append * 10 + 1
        .Add FALT, Asc("2"), conMenu_Edit_Append * 10 + 2
        .Add FALT, Asc("3"), conMenu_Edit_Append * 10 + 3
        .Add FALT, Asc("4"), conMenu_Edit_Append * 10 + 4
        .Add FALT, Asc("5"), conMenu_Edit_Append * 10 + 5
        .Add 0, VK_DELETE, conMenu_Edit_Untread
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add 0, VK_F1, conMenu_Help_Help
                
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
        
    End With
    
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Private Sub BodyEdit_PromptInfo(ByVal strInfo As String)
    stbThis.Panels(2).Text = strInfo
End Sub

Private Sub BodyEdit_RButton(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button <> 2 Then Exit Sub
    
    If cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub
    
    '组装右键菜单
    If mcbrMenuBar部位 Is Nothing Then Exit Sub
    If mcbrMenuBar部位.CommandBar.Controls.Count = 0 Then Exit Sub
    Set cbrMenuBar = mcbrMenuBar部位
    Set cbrPopupBar = cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.Id, cbrControl.Caption)
        cbrPopupItem.IconId = cbrControl.IconId
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
    
End Sub

Private Sub BodyEdit_SelectScale(ByVal intScale As Integer)
    Call AddActiveMenu
End Sub

'######################################################################################################################
'控件事件

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As CommandBarControl
    Dim lngKey As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.Id
        Case conMenu_Tool_Option
            
            If Control.Parameter = "" Then
                Control.Parameter = "1"
            Else
                Control.Parameter = ""
            End If
            
        Case conMenu_File_PrintSet
            
            On Error Resume Next
            frmPrintSet.mbytMode = 1
            frmPrintSet.mstrPrivs = mstrPrivs
            frmPrintSet.Show 1, Me
            
        Case conMenu_File_Preview
            
            Call PrintData(2)
            
        Case conMenu_File_Print
        
            Call PrintData(1)
        
        Case conMenu_View_ToolBar_Button
        
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_Notify    '设定体温单开始日期(已存在体温数据的不允许设定)
            Dim strParam As String
            
            If Not BodyEdit.zlMenuClick("设定开始日期") Then Exit Sub
            '参数：病人ID,主页ID,病区ID,科室ID,出院标志;编辑标志;婴儿
            strParam = Val(mrsParam("病人id")) & ";" & Val(mrsParam("主页id")) & ";" & Val(mrsParam("病区id")) & ";" & Val(mrsParam("出院")) & ";" & Val(mrsParam("编辑").Value) & ";" & Val(mrsParam("婴儿").Value)
            Call BodyEdit.zlMenuClick("初始数据", strParam)
        
        Case conMenu_Edit_Adjust
            
            If BodyEdit.CurPostion >= 0 Then Call BodyEdit.zlMenuClick("填写记录线")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread
            
            If BodyEdit.CurPostion >= 0 Then Call BodyEdit.zlMenuClick("清除记录线")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify
    
            Call BodyEdit.zlMenuClick("填写手术日")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
        
            Call BodyEdit.zlMenuClick("清除手术日")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Append            '添加项目
            Call BodyEdit.zlMenuClick("添加项目")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Stop              '删除项目
            Call BodyEdit.zlMenuClick("删除项目")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Compend * 10 + 1, conMenu_Edit_Compend * 10 + 2, conMenu_Edit_Compend * 10 + 3
            
            mstr体温部位 = Control.Parameter
            
            BodyEdit.体温部位 = mstr体温部位
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Compend * 10 + 5, conMenu_Edit_Compend * 10 + 6
            
            mstr呼吸方式 = Control.Parameter
            
            BodyEdit.呼吸方式 = mstr呼吸方式
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Compend * 10 + 8
            
            mstr脉搏 = IIf(Control.Checked = False, "起搏器", "")
            
            BodyEdit.脉搏方式 = mstr脉搏
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Reuse
            If BodyEdit.zlMenuClick("恢复数据") Then
    
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Audit
            
            Call BodyEdit.zlMenuClick("复试合格")
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Blankoff
            
            Call BodyEdit.zlMenuClick("取消复试")
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsThis.RecalcLayout
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Location
            
            Call LocationObj(txt)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Forward
            
            If mlngRowNum = 1 Then mlngRowNum = mrsPatient.RecordCount + 1
            
            mrsPatient.Filter = ""
            mrsPatient.Filter = "ID<" & mlngRowNum
            If mrsPatient.RecordCount > 0 Then
                mrsPatient.MoveLast
                mlngRowNum = Val(mrsPatient("ID").Value)
                txt.Text = zlCommFun.NVL(mrsPatient("姓名").Value)
                mrsParam("病人id").Value = Val(mrsPatient("病人id").Value)
                mrsParam("主页id").Value = Val(mrsPatient("主页id").Value)
                mrsParam("婴儿").Value = 0
                Select Case CStr(mrsPatient("类型").Value)
                Case "死亡", "死亡病人", "出院病人"
                    mrsParam("出院").Value = 1
                Case Else
                    mrsParam("出院").Value = 0
                End Select
                
                Call OpenPatientMap
                txt.Tag = ""
            End If
            mrsPatient.Filter = ""
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Backward
            
            If mlngRowNum = mrsPatient.RecordCount Then mlngRowNum = 0
            
            mrsPatient.Filter = ""
            mrsPatient.Filter = "ID>" & mlngRowNum
            If mrsPatient.RecordCount > 0 Then
                mrsPatient.MoveFirst
                mlngRowNum = Val(mrsPatient("ID").Value)
                txt.Text = zlCommFun.NVL(mrsPatient("姓名").Value)
                mrsParam("病人id").Value = Val(mrsPatient("病人id").Value)
                mrsParam("主页id").Value = Val(mrsPatient("主页id").Value)
                mrsParam("婴儿").Value = 0
                Select Case CStr(mrsPatient("类型").Value)
                Case "死亡", "死亡病人", "出院病人"
                    mrsParam("出院").Value = 1
                Case Else
                    mrsParam("出院").Value = 0
                End Select
                
                Call OpenPatientMap
                txt.Tag = ""
            End If
            mrsPatient.Filter = ""
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Save
            '保存数据
            
            If BodyEdit.zlMenuClick("保存数据") Then
                mblnChanged = True
            End If
            
            cbsThis.RecalcLayout
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Price
            
            '计算饮入
            Select Case Control.Parameter
            Case "饮入"
                mblnChanged = BodyEdit.zlMenuClick("计算饮入")
            Case "手术"
                mblnChanged = BodyEdit.zlMenuClick("获取手术日")
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Append * 10 + 1

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("假肛")

        Case conMenu_Edit_Append * 10 + 2

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("灌肠")

        Case conMenu_Edit_Append * 10 + 3

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("灌肠后排泄")

        Case conMenu_Edit_Append * 10 + 4

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("导尿")
            
        Case conMenu_Edit_Append * 10 + 5

'            Control.Checked = Not Control.Checked
            mblnChanged = BodyEdit.zlMenuClick("保留导尿")
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Jump
            
            Call BodyEdit.zlMenuClick("装载数据", Control.Parameter)
            cbsThis.RecalcLayout
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SendOther
            
            Call BodyEdit.zlMenuClick("操作曲线", Val(Control.Parameter))
            
            Call AddActiveMenu
            
            cbsThis.RecalcLayout
            
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.hWnd)
            
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hWnd)
            
        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)

    If stbThis.Visible Then Bottom = stbThis.Height
    
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With BodyEdit
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
    End With
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0
    On Error Resume Next
    
    Select Case Control.Id

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify, conMenu_Edit_Save, conMenu_Edit_Reuse, conMenu_Edit_Price
        
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit                 '复试合格
        
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1 And BodyEdit.AllowAudit)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Blankoff              '取消复试
        
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1 And BodyEdit.AllowUnAudit)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
    
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1 And Val(BodyEdit.GetUpObj.ColData(BodyEdit.GetUpObj.Col)) > 0)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append            '添加项目
        
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Stop              '删除项目
        
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Compend * 10 + 1, conMenu_Edit_Compend * 10 + 2, conMenu_Edit_Compend * 10 + 3    '口温/腋温/肛温
    
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1) And BodyEdit.体温项目
        Control.Checked = (Control.Parameter = mstr体温部位)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Compend * 10 + 5, conMenu_Edit_Compend * 10 + 6                                   '自主呼吸/呼吸机辅助
    
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1) And BodyEdit.呼吸项目
        Control.Checked = (Control.Parameter = mstr呼吸方式)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Compend * 10 + 8                                                                  '有无使用起搏器
    
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1) And BodyEdit.脉搏项目
        Control.Checked = (mstr脉搏 = "起搏器")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Adjust, conMenu_Edit_Untread
        
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1 And BodyEdit.CurPostion >= 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 1

'        Control.Checked = (BodyEdit.mbytSpecChar = 1)
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1 And BodyEdit.是否大便项目)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 2

'        Control.Checked = (BodyEdit.mbytSpecChar = 2)
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1 And BodyEdit.是否大便项目)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 3

'        Control.Checked = (BodyEdit.mbytSpecChar = 3)
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1 And BodyEdit.是否大便项目)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 4
    
'        Control.Checked = (BodyEdit.mbytSpecChar = 4)
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1 And BodyEdit.是否出液项目)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append * 10 + 5
    
'        Control.Checked = (BodyEdit.mbytSpecChar = 5)
        Control.Enabled = (Val(mrsParam("编辑").Value) = 1 And BodyEdit.是否出液项目)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump
        
        If Control.Parameter = "" Then
            Control.Checked = True
        Else
            Control.Checked = (Val(Split(Control.Parameter, ";")(4)) = BodyEdit.Page)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SendOther
        
        Control.Checked = (Val(Control.Parameter) = BodyEdit.LineType)
        
    Case conMenu_View_ToolBar_Button
    
        Control.Checked = Me.cbsThis(2).Visible
        
    Case conMenu_View_ToolBar_Text
    
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        
    Case conMenu_View_ToolBar_Size
    
        Control.Checked = Me.cbsThis.Options.LargeIcons
        
    Case conMenu_View_StatusBar
    
        Control.Checked = Me.stbThis.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        Control.Checked = (mstrFindKey = Control.Parameter)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
        
'        Control.Enabled = (mlngRowNum > 1)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward, conMenu_View_Backward
        
        Control.Enabled = (mrsPatient.RecordCount > 1)
        
    End Select
End Sub

Private Sub cmd_Click()
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset

    '------------------------------------------------------------------------------------------------------------------
    mrsPatient.Filter = ""
    If mrsPatient.RecordCount > 0 Then
        mrsPatient.MoveFirst
        If ShowTxtSelDialog(Me, txt, "床号,1200,0,0;姓名,1200,0,1;性别,600,0,0;科室,1800,0,0;住院号,1080,0,0", Me.Name & "\病人清单选择", "请从下面选择一个病人。", mrsPatient, rs, 5600, 4500, , CStr(mlngRowNum), 2, True) Then
            
            mlngRowNum = Val(mrsPatient("ID").Value)
            
            txt.Text = zlCommFun.NVL(rs("姓名").Value)
            
            
            mrsParam("病人id").Value = Val(rs("病人id").Value)
            mrsParam("主页id").Value = Val(rs("主页id").Value)
            mrsParam("婴儿").Value = 0
            Select Case CStr(rs("类型").Value)
            Case "死亡", "死亡病人", "出院病人"
                mrsParam("出院").Value = 1
            Case Else
                mrsParam("出院").Value = 0
            End Select
            
            Call OpenPatientMap
            
            txt.Tag = ""
        End If
    End If
    mrsPatient.Filter = ""
    
    Call LocationObj(txt)

    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
End Sub

Private Sub Form_Load()
        
    Call InitCommonControls
    
    If mblnChildForm Then
'        Call RestoreWinState(Me, App.ProductName, "ChildForm")
    Else
        Call RestoreWinState(Me, App.ProductName)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Call zlDatabase.SetPara("查找方法", mstrFindKey, glngSys, 1255)
    
    UnHook Me.hWnd
    
    If mblnChildForm Then
'        Call SaveWinState(Me, App.ProductName, "ChildForm")
    Else
        Call SaveWinState(Me, App.ProductName)
    End If
    
    
    
    Set mrsPatient = Nothing
    Set mobjFindKey = Nothing
    mblnShowing = False
End Sub

Private Sub BodyEdit_zlAfterPrint()
    gblnPrinted = True
    RaiseEvent AfterPrint
End Sub

Private Sub BodyEdit_DbClickCur()
    
    Call BodyEdit.zlMenuClick("填写记录线")
        
End Sub

Private Sub txt_Change()
    txt.Tag = "Changed"
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim bytMode As Byte
    Dim lng病人ID As Long
    Dim strInput As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If txt.Tag = "Changed" And txt.Text <> "" Then
            If InStr(txt.Text, "'") Then
                ShowSimpleMsg "输入的内容中有非法字符 ' ！"
                Exit Sub
            End If
            
            Select Case mstrFindKey
'            Case "病人id"
'                strInput = "病人id=" & Val(txt.Text)
'                bytMode = 2
'            Case "门诊号"
'                strInput = "门诊号=" & Val(txt.Text)
'                bytMode = 4
            Case "床  号"
                strInput = "床号='" & Trim(txt.Text) & "'"
                bytMode = 5
            Case "住院号"
                strInput = "住院号=" & Val(txt.Text)
                bytMode = 3
            Case "就诊卡"
                strInput = "就诊卡号='" & Trim(txt.Text) & "'"
                bytMode = 1
            End Select
                        
        End If

    ElseIf mstrFindKey = "就诊卡" And txt.Tag = "Changed" And txt.Text <> "" Then
        If Len(txt.Text) = mint就诊卡号码长度 - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txt.Text <> "" Then
            If KeyAscii <> 13 Then
                txt.Text = txt.Text & Chr(KeyAscii)
                txt.SelStart = Len(txt.Text)
                KeyAscii = 0
            End If

            strInput = "就诊卡号='" & Trim(txt.Text) & "'"
            bytMode = 1
        End If
    End If
    
    If strInput <> "" Then
        txt.Tag = ""
        mrsPatient.Filter = ""
        mrsPatient.Filter = strInput
        If mrsPatient.RecordCount > 0 Then
            mrsPatient.MoveFirst
            lng病人ID = Val(mrsPatient("病人id").Value)
            mlngRowNum = Val(mrsPatient("ID").Value)
            
            txt.Text = zlCommFun.NVL(mrsPatient("姓名").Value)
            txt.Tag = ""
            
            mrsParam("病人id").Value = Val(mrsPatient("病人id").Value)
            mrsParam("主页id").Value = Val(mrsPatient("主页id").Value)
            mrsParam("婴儿").Value = 0
            Select Case CStr(mrsPatient("类型").Value)
            Case "死亡", "死亡病人", "出院病人"
                mrsParam("出院").Value = 1
            Case Else
                mrsParam("出院").Value = 0
            End Select
            
            Call OpenPatientMap
        Else
            ShowSimpleMsg "没有找到符合条件的病人！"
            txt.Text = mstrSvr姓名
        End If
        mrsPatient.Filter = ""

        Call LocationObj(txt)
        
    End If

    Exit Sub

errHand:
End Sub

Private Sub AddActiveMenu()
    '------------------------------------------------------------
    '根据项目添加菜单(如果是体温则增加体温部位;如果是呼吸则增加呼吸方式)
    Dim varTmp As Variant
    Dim rs As New ADODB.Recordset
    Dim cbrControl As CommandBarControl
    
    If Not mcbrMenuBar部位 Is Nothing Then
        If mcbrMenuBar部位.CommandBar.Controls.Count <> 0 Then
            Call mcbrMenuBar部位.CommandBar.Controls.DeleteAll
            Call mcbrMenuBar编辑.CommandBar.Controls.Item(2).Delete
        End If
    End If
    If BodyEdit.体温项目 Then
        Set mcbrMenuBar部位 = mcbrMenuBar编辑.CommandBar.Controls.Add(xtpControlPopup, conMenu_Edit_Compend, "体温部位(&T)", 2)
        gstrSQL = "Select 记录符 From 体温记录项目 Where 项目序号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1)
        If rs.BOF = False Then
            varTmp = Split(zlCommFun.NVL(rs("记录符").Value, "·,×,○"), ",")
        Else
            varTmp = Split("·,×,○", ",")
        End If
        
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 1, "口温" & varTmp(0) & "(&1)", -1, False): cbrControl.Parameter = "口温": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 2, "腋温" & varTmp(1) & "(&2)", -1, False): cbrControl.Parameter = "腋温": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 3, "肛温" & varTmp(2) & "(&3)", -1, False): cbrControl.Parameter = "肛温": cbrControl.IconId = 1
    End If
    
    If BodyEdit.呼吸项目 Then
        Set mcbrMenuBar部位 = mcbrMenuBar编辑.CommandBar.Controls.Add(xtpControlPopup, conMenu_Edit_Compend, "呼吸方式(&T)", 2)
        gstrSQL = "Select 记录符 From 体温记录项目 Where 项目序号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 3)
        If rs.BOF = False Then
            varTmp = zlCommFun.NVL(rs("记录符").Value, "⊕")
        Else
            varTmp = "⊕"
        End If
        
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 5, "自主呼吸" & varTmp & "(&1)", -1, False): cbrControl.Parameter = "自主呼吸": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 6, "呼吸机 (&2)", -1, False): cbrControl.Parameter = "呼吸机": cbrControl.IconId = 1
    End If

    If BodyEdit.脉搏项目 Then
        Set mcbrMenuBar部位 = mcbrMenuBar编辑.CommandBar.Controls.Add(xtpControlPopup, conMenu_Edit_Compend, "脉博方式(&T)", 2)
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10 + 8, "使用起搏器" & "(&1)", -1, False): cbrControl.Parameter = "起搏器": cbrControl.IconId = 1
    End If
End Sub
