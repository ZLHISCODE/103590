VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.0#0"; "zlIDKind.ocx"
Begin VB.Form frmForceGet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "强制续诊"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   Icon            =   "frmForceGet.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin VB.Frame fraKind 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   280
      Left            =   960
      TabIndex        =   12
      Top             =   160
      Width           =   2190
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   270
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmForceGet.frx":058A
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowSortName    =   -1  'True
         DefaultCardType =   "就诊卡"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   Begin VB.CheckBox chk急诊 
      Caption         =   "标记为急诊"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   2303
      Width           =   1215
   End
   Begin VB.ComboBox cbo续诊科室 
      Height          =   300
      Left            =   885
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   2
      Left            =   -120
      TabIndex        =   8
      Top             =   585
      Width           =   7635
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   -120
      TabIndex        =   7
      Top             =   2130
      Width           =   7635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5205
      TabIndex        =   3
      Top             =   2250
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6300
      TabIndex        =   4
      Top             =   2250
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsRegist 
      Height          =   1080
      Left            =   135
      TabIndex        =   0
      Top             =   945
      Width           =   7200
      _cx             =   12700
      _cy             =   1905
      Appearance      =   1
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
      BackColorSel    =   16772055
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmForceGet.frx":0651
      ScrollTrack     =   -1  'True
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
   Begin VB.Image imgSentence 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3120
      Picture         =   "frmForceGet.frx":0730
      ToolTipText     =   "选择本科室最近的病人"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgStaKB 
      Height          =   330
      Left            =   3430
      Picture         =   "frmForceGet.frx":0E1A
      ToolTipText     =   "点击启用屏幕键盘"
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lbl续诊科室 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "续诊科室"
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   2340
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Top             =   195
      Width           =   90
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "挂号记录："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   705
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "续诊病人"
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   5
      Top             =   195
      Width           =   720
   End
End
Attribute VB_Name = "frmForceGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String 'IN
Private mstr挂号单 As String 'Out
Private mlng病人ID As Long
Private mbytSize As Byte
Private Enum COL_REGIST
    COL_NO = 0
    col_科室 = 1
    COL_项目 = 2
    COL_医生 = 3
    COL_诊室 = 4
    COL_时间 = 5
    COL_状态 = 6
    COL_急诊 = 7
End Enum
Private mblnStaKB As Boolean '是否自动启用屏幕键盘
Private mlng卡类别ID As Long
Private mobjSquare As Object     '刘兴洪 日期:2011-12-25 16:37:31
Private mblnCard As Boolean
Private mobjKeyBoard As Object '屏幕键盘对象动态创建
Private mbln刷卡回车 As Boolean
Private msinTime As Single
Private mlng接诊科室ID As Long

Public Function ShowMe(frmParent As Object, ByVal strPrivs As String, lng接诊科室ID As Long, _
    ByVal objSquare As Object) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:强制续诊入口
    '入参:objSquare-卡结算部件对象
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-12-25 16:14:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjSquare = objSquare
    mstrPrivs = strPrivs
    mlng接诊科室ID = lng接诊科室ID
    Me.Show 1, frmParent
    ShowMe = mstr挂号单
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim arrSQL As Variant
    Dim strTime As String
    Dim i As Integer
    Dim blnTrans As Boolean
    
    If mlng病人ID = 0 Then
        MsgBox "请输入需要续诊的病人。", vbInformation, gstrSysName
        PatiIdentify.SetFocus: Exit Sub
    End If
    If vsRegist.TextMatrix(1, 0) = "" Then
        MsgBox "该病人没有可用于续诊的挂号记录。", vbInformation, gstrSysName
        PatiIdentify.SetFocus: Exit Sub
    End If
    If cbo续诊科室.ListIndex = -1 Then
        MsgBox "请确定对病人进行续诊的科室。", vbInformation, gstrSysName
        cbo续诊科室.SetFocus: Exit Sub
    End If
    On Error GoTo errH
    With vsRegist
        If BillExpend(.TextMatrix(.Row, COL_NO)) Then
            MsgBox "该病人挂号已超过有效天数，不能再进行转诊。", vbInformation, gstrSysName
            Exit Sub
        End If
        arrSQL = Array()
        If Val(.RowData(.Row)) = 2 Then
            '刘兴洪 日期:2011-12-25 16:37:31
            '对预约进行强制续诊
            If Val(zlDatabase.GetPara("挂号模式", glngSys, 9000, 1)) <> 1 And Not mobjSquare Is Nothing Then
                If Not mobjSquare.zlRegisterIncept(Me, p门诊医生站, Trim(.TextMatrix(.Row, COL_NO)), cbo续诊科室.Text, IIf(mblnCard, mlng卡类别ID, 0), "") Then Exit Sub
            Else
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人预约挂号_接收('" & Trim(.TextMatrix(.Row, COL_NO)) & "','" & cbo续诊科室.Text & "')"
            End If
        End If
        '记录就诊变动记录
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_就诊变动记录_Insert('" & .TextMatrix(.Row, COL_NO) & "',3,'强制续诊','" & UserInfo.姓名 & "','" & UserInfo.编号 & "',NULL," & cbo续诊科室.ItemData(cbo续诊科室.ListIndex) & ",NULL," & UserInfo.ID & ",'" & UserInfo.姓名 & "')"
         
        '执行
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人接诊(" & mlng病人ID & ",'" & .TextMatrix(.Row, COL_NO) & "'," & cbo续诊科室.ItemData(cbo续诊科室.ListIndex) & ",'" & UserInfo.姓名 & "',Null," & IIf(chk急诊.Visible And chk急诊.Value = 1, "1", "0") & ")"
        '开启事务
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mstr挂号单 = .TextMatrix(.Row, COL_NO)
    End With
    '问题号：42196
    'Call zlDatabase.SetPara("本机门诊科室", cbo续诊科室.ItemData(cbo续诊科室.ListIndex), glngSys, p门诊医生站, InStr(mstrPrivs, "参数设置") > 0)
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If mbln刷卡回车 And KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        mbln刷卡回车 = False
    End If
End Sub

Private Sub Form_Resize()
    Frame1(1).Width = Me.Width + 100
    Frame1(2).Width = Frame1(1).Width
End Sub

Private Sub imgSentence_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim objCardData As Object
    Dim n As Long
    
    If gint急诊挂号天数 = 0 And gint普通挂号天数 = 0 Then
        n = 1
    Else
        If gint普通挂号天数 - gint急诊挂号天数 > 0 Then
            n = gint普通挂号天数
        Else
            n = gint急诊挂号天数
        End If
    End If
    
    vRect = zlControl.GetControlRect(fraKind.hwnd)
    
    blnCancel = True
    On Error GoTo errH
    strSQL = "Select A.病人ID as ID,A.门诊号,A.姓名 as 名称,A.性别,A.年龄,a.发生时间 as 挂号时间 From 病人挂号记录 A,病人信息 B" & _
    " Where A.病人ID=B.病人ID" & IIf(mlng接诊科室ID = 0, "", " And A.执行部门ID+0=[1]") & _
    " And A.记录性质 <> 2 And A.记录状态 = 1 " & _
    " And A.发生时间 Between Sysdate-" & n & " And trunc(Sysdate)+1-1/24/60/60 order by A.发生时间 desc"

    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "选择最近7天的病人", False, "", "", False, False, True, _
                vRect.Left, vRect.Top + 50, PatiIdentify.Height, blnCancel, False, True, mlng接诊科室ID)
    If blnCancel = True Then
        MsgBox "未查找到本科室近期的病人!", vbInformation, gstrSysName
        Exit Sub
    End If
    If (Not rsTmp Is Nothing) And blnCancel = False Then
        Set objCardData = New zlIDKind.PatiInfor
        mlng病人ID = Val(rsTmp!ID & "")
        PatiIdentify.Text = rsTmp!名称 & ""
        objCardData.病人ID = mlng病人ID
        objCardData.姓名 = rsTmp!名称 & ""
        objCardData.门诊号 = rsTmp!门诊号 & ""
        objCardData.性别 = rsTmp!性别 & ""
        objCardData.年龄 = rsTmp!年龄 & ""
        Call SetPati(objCardData)
    Else
        MsgBox "未查找到本科室近期的病人!", vbInformation, gstrSysName
        blnCancel = True
        Exit Sub
    End If
    Exit Sub
errH:
    MsgBox "未查找到本科室近期的病人!", vbInformation, gstrSysName
    blnCancel = True
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgStaKB_Click()
    On Error Resume Next
    If mobjKeyBoard Is Nothing Then Set mobjKeyBoard = CreateObject("zlScreenKeyboard.clsKeyBoard")
    Call mobjKeyBoard.StartUp
    Call mobjKeyBoard.SetPos
    err.Clear: On Error GoTo 0
End Sub
 
Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim lng本机科室ID As Long
    Dim strSQL As String
    
    mbln刷卡回车 = False
    
    '启用屏幕键盘
    mblnStaKB = Val(zlDatabase.GetPara("启用屏幕键盘", glngSys, p门诊医生站)) <> 0
    '字体设置
    mbytSize = zlDatabase.GetPara("字体", glngSys, p门诊医生站, "0")
    Call initCardSquareData
    '设置缺省查找方式
    On Error Resume Next
    If Val(zlDatabase.GetPara("使用个性化风格")) <> 0 Then
        PatiIdentify.objIDKind.IDKind = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "IDKind", 0))
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo errH
    mlng病人ID = 0
    mstr挂号单 = ""
    '问题号：42196
    'lng本机科室ID = Val(zlDatabase.GetPara("本机门诊科室", glngSys, p门诊医生站, , Array(lbl续诊科室, cbo续诊科室), InStr(mstrPrivs, "参数设置") > 0))
    'If lng本机科室ID = 0 Then
        '接诊范围：1=挂本人号的病人,2=本诊室病人,3=本科室病人   (本机参数，如果是其他医生用过该机器，则可能该科室不是当前医生的科室)
        '问题号：42196。原因：10.28由于分诊叫号的调整，已改为接诊科室必须填。所以，强制续诊时，缺省科室可不判断接诊范围，直接取本地参数的接诊科室。
        'If Val(zlDatabase.GetPara("接诊范围", glngSys, p门诊医生站, "2")) = 3 Then
            lng本机科室ID = Val(zlDatabase.GetPara("接诊科室", glngSys, p门诊医生站))
        'End If
    'End If
    
    '确定缺省的续诊科室
    On Error GoTo errH
    strSQL = "Select Distinct A.ID,A.名称,B.缺省" & _
        " From 部门表 A,部门人员 B,部门性质说明 C" & _
        " Where A.ID=B.部门ID And A.ID=C.部门ID And C.工作性质||''='临床' And C.服务对象 IN(1,3)" & _
        " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) And B.人员ID=[1]" & _
        " Order by A.名称"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    Do While Not rsTmp.EOF
        cbo续诊科室.AddItem rsTmp!名称
        cbo续诊科室.ItemData(cbo续诊科室.NewIndex) = rsTmp!ID
                
        If rsTmp!ID = lng本机科室ID Then
            cbo续诊科室.ListIndex = cbo续诊科室.NewIndex
        
        ElseIf Nvl(rsTmp!缺省, 0) = 1 And cbo续诊科室.ListIndex = -1 Then
            cbo续诊科室.ListIndex = cbo续诊科室.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cbo续诊科室.ListIndex = -1 And cbo续诊科室.ListCount > 0 Then cbo续诊科室.ListIndex = 0
    
    If mblnStaKB Then
        On Error Resume Next
        Set mobjKeyBoard = Nothing
        Set mobjKeyBoard = CreateObject("zlScreenKeyboard.clsKeyBoard")
        err.Clear: On Error GoTo 0
        If Not mobjKeyBoard Is Nothing Then
            imgStaKB.Visible = True
            Call mobjKeyBoard.StartUp
        Else
            MsgBox "屏幕键盘部件未能正确安装，不能使用！", vbInformation, gstrSysName
        End If
    End If
    Call SetFontSize(mbytSize)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Call PatiIdentify.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjKeyBoard = Nothing
    Set mobjSquare = Nothing
    If Val(zlDatabase.GetPara("使用个性化风格")) <> 0 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "IDKind", PatiIdentify.objIDKind.IDKind)
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    
    If objHisPati Is Nothing Then
        MsgBox "没有找到相关的病人信息", vbInformation, gstrSysName
        Exit Sub
    End If
    If objHisPati.病人ID = 0 Then
        MsgBox "没有找到相关的病人信息", vbInformation, gstrSysName
        Exit Sub
    End If
    Call SetPati(objHisPati)
End Sub

Private Sub SetPati(ByVal objHisPati As zlIDKind.PatiInfor)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNO As String
    Dim str科室IDs As String, i As Long
    Dim strMsg As String, blnDo As Boolean
    
    lblInfo(1).Caption = "门诊号:" & objHisPati.门诊号 & "  性别:" & objHisPati.性别 & "  年龄:" & objHisPati.年龄
    '为了明确提示，不通过SQL而通过程序检查条件
    strSQL = "Select A.NO,A.记录性质,D.ID as 科室ID,D.名称 as 科室," & _
        " C.ID as 项目ID,C.名称 as 项目,A.执行人,A.诊室,A.发生时间,A.执行状态,Decode(A.急诊,1,'是','否') as 急诊" & _
        " From 病人挂号记录 A,门诊费用记录 B,收费项目目录 C,部门表 D" & _
        " Where A.NO=B.NO And B.记录性质=4 And B.记录状态 in (1,0) And B.收费类别='1' And a.记录性质 in (1,2) And a.记录状态 =1" & _
        "           And B.价格父号 is Null And B.从属父号 is Null And B.收费细目ID=C.ID And A.执行部门ID=D.ID" & _
        "           And A.发生时间<=trunc(Sysdate)+1-1/24/60/60  And A.病人ID=[1]" & _
        IIf(Val(zlDatabase.GetPara(210, glngSys)) = 1, "", " And A.发生时间 Between Sysdate - Decode(A.急诊,1," & IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数) & "," & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数) & ") And trunc(Sysdate)+1-1/24/60/60") & _
        " Order by 发生时间 Desc "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objHisPati.病人ID)
    With vsRegist
        If Not rsTmp.EOF Then
            .Rows = .FixedRows
            str科室IDs = GetUser科室IDs
            mlng病人ID = objHisPati.病人ID
            Do While Not rsTmp.EOF
                blnDo = True
                If Nvl(rsTmp!执行人) = UserInfo.姓名 Then
                    strMsg = strMsg & vbCrLf & "挂号记录" & rsTmp!NO & "：医生是本人的" & Decode(Nvl(rsTmp!执行状态, 0), 0, "候诊", 1, "已诊", 2, "正在就诊的") & "号，无需使用续诊功能。"
                    blnDo = False '医生本人已诊或在诊的号，无需通过续诊功能。
                End If
                If blnDo Then
                    '科内续诊：病人在当前医生所属科室挂的号，就诊状态无限制
                    '全院续诊：病人在任何科室挂的号；他科不能是正在就诊
                    If InStr("," & str科室IDs & ",", "," & rsTmp!科室ID & ",") = 0 Then
                        If InStr(mstrPrivs, "全院病人续诊") = 0 Then
                            strMsg = strMsg & vbCrLf & "挂号记录" & rsTmp!NO & "：挂号科室为""" & rsTmp!科室 & """，不是本科挂号，没有权限进行续诊。"
                            blnDo = False
                        ElseIf Nvl(rsTmp!执行状态, 0) = 2 And InStr(GetInsidePrivs(p门诊医生站), ";允许强制续诊正在就诊的病人;") = 0 Then
                            strMsg = strMsg & vbCrLf & "挂号记录" & rsTmp!NO & "：""" & rsTmp!科室 & """的医生""" & rsTmp!执行人 & """正在就诊，不能进行续诊。"
                            blnDo = False
                        End If
                    End If
                End If
                If blnDo Then
                    .AddItem "": i = .Rows - 1
                    .TextMatrix(i, COL_NO) = rsTmp!NO
                    .RowData(i) = Val(Nvl(rsTmp!记录性质))
                    .TextMatrix(i, col_科室) = rsTmp!科室
                    .Cell(flexcpData, i, col_科室) = Val(rsTmp!科室ID)
                    .TextMatrix(i, COL_项目) = rsTmp!项目
                    .Cell(flexcpData, i, COL_项目) = Val(rsTmp!项目ID)
                    .TextMatrix(i, COL_医生) = Nvl(rsTmp!执行人)
                    .TextMatrix(i, COL_诊室) = Nvl(rsTmp!诊室)
                    .TextMatrix(i, COL_时间) = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm")
                    .TextMatrix(i, COL_状态) = Decode(Nvl(rsTmp!执行状态, 0), 0, "候诊", 1, "已诊", 2, "正在就诊")
                    .TextMatrix(i, COL_急诊) = rsTmp!急诊
                End If
                rsTmp.MoveNext
            Loop
            If .Rows = .FixedRows Then
                .Rows = .FixedRows + 1
                strMsg = "病人""" & objHisPati.姓名 & """没有可以续诊的挂号记录：" & vbCrLf & strMsg
            Else
                strMsg = ""
            End If
        Else
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
            strMsg = "病人""" & objHisPati.姓名 & """在挂号有效天数内没有挂号记录。"
        End If
        .Row = .FixedRows
    End With
    
    If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
    
    If blnDo Then
        mbln刷卡回车 = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mbln刷卡回车 Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mbln刷卡回车 = False
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNO As String
    Dim str科室IDs As String, i As Long
    Dim strMsg As String, blnDo As Boolean
    Dim lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String
    Dim vRect As RECT
    Dim blnLikeCode As Boolean '根据简码查找
    
    On Error GoTo errH
    If strShowText = "" Then blnCancel = True: Exit Sub
    If zlCommFun.IsCharAlpha(strShowText) Then
        blnLikeCode = True
        strWhere = " Instr(',' || Zlpinyincode(姓名), ',' || [1]) > 0"
        strWhere = strWhere & "And 记录性质 <> 2 And 记录状态 = 1 And 发生时间 Between Sysdate - Decode(急诊,1," & IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数) & "," & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数) & ") And trunc(Sysdate)+1-1/24/60/60"
        strShowText = UCase(strShowText)
    ElseIf Left(strShowText, 1) = "." Then '挂号单
        strNO = GetFullNO(Mid(UCase(strShowText), 2), 12)
        strSQL = "Select 病人ID,门诊号,姓名,性别,年龄 From 病人挂号记录 Where NO=[2] And 记录性质=1 "
    Else
        Select Case objCard.名称
            Case "姓名"
                strSQL = "Select A.病人ID,A.门诊号,A.姓名,A.性别,A.年龄 From 病人挂号记录 A,病人信息 B Where A.病人ID=B.病人ID And A.记录性质 <> 2 And A.记录状态 = 1 And b.姓名=[1]" & _
                            IIf(Val(zlDatabase.GetPara(210, glngSys)) = 1, "", " And a.发生时间 Between Sysdate - Decode(A.急诊,1," & IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数) & "," & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数) & ") And trunc(Sysdate)+1-1/24/60/60")
                strSQL = strSQL & " Order by A.发生时间 desc"
            Case "挂号单号"
                strNO = GetFullNO(UCase(strShowText), 12)
                strSQL = "Select 病人ID,门诊号,姓名,性别,年龄 From 病人挂号记录 Where NO=[2] And 记录性质=1  "
        End Select
    End If
    If strWhere <> "" Then
        strSQL = "Select /*+ RULE */ 病人ID,门诊号,姓名,性别,年龄 From 病人挂号记录 Where " & strWhere
    End If
    If strSQL = "" Then Exit Sub
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strShowText, strNO)

    If rsTmp.EOF = False Then
        If rsTmp.RecordCount > 1 Then
            vRect = zlControl.GetControlRect(fraKind.hwnd)
            If blnLikeCode Then
                strSQL = "Select /*+ RULE */ 病人ID as ID,门诊号,姓名 as 名称,性别,年龄,发生时间 as 挂号时间 From 病人挂号记录 Where " & strWhere & " order by 发生时间 desc"
            Else
                strSQL = "Select A.病人ID as ID,A.门诊号,A.姓名 as 名称,A.性别,A.年龄,a.发生时间 as 挂号时间 From 病人挂号记录 A,病人信息 B Where A.病人ID=B.病人ID And B.姓名=[1] And A.记录性质 <> 2 And A.记录状态 = 1" & IIf(Val(zlDatabase.GetPara(210, glngSys)) = 1, "", " And a.发生时间 Between Sysdate - Decode(A.急诊,1," & IIf(gint急诊挂号天数 = 0, 1, gint急诊挂号天数) & "," & IIf(gint普通挂号天数 = 0, 1, gint普通挂号天数) & ") And trunc(Sysdate)+1-1/24/60/60") & " order by a.发生时间 desc"
            End If
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "选择病人", False, "", "", False, False, True, _
                        vRect.Left, vRect.Top + 50, PatiIdentify.Height, blnCancel, False, True, strShowText, strNO)
            
            If Not rsTmp Is Nothing And blnCancel = False Then
                mlng病人ID = Val(rsTmp!ID & "")
                strShowText = rsTmp!名称 & ""
                Set objCardData = New zlIDKind.PatiInfor
                objCardData.病人ID = mlng病人ID
                objCardData.姓名 = rsTmp!名称 & ""
                objCardData.门诊号 = rsTmp!门诊号 & ""
                objCardData.性别 = rsTmp!性别 & ""
                objCardData.年龄 = rsTmp!年龄 & ""
                Call SetPati(objCardData)
                blnFindPatied = True
            Else
                blnCancel = True
                Exit Sub
            End If
        Else
            mlng病人ID = Val(rsTmp!病人ID & "")
            strShowText = rsTmp!姓名 & ""
            Set objCardData = New zlIDKind.PatiInfor
            objCardData.病人ID = mlng病人ID
            objCardData.姓名 = rsTmp!姓名 & ""
            objCardData.门诊号 = rsTmp!门诊号 & ""
            objCardData.性别 = rsTmp!性别 & ""
            objCardData.年龄 = rsTmp!年龄 & ""
            Call SetPati(objCardData)
            blnFindPatied = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    mlng卡类别ID = objCard.接口序号
End Sub


Private Sub vsRegist_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    chk急诊.Visible = (vsRegist.TextMatrix(NewRow, COL_急诊) = "否")
End Sub

Private Sub vsRegist_GotFocus()
    vsRegist.BackColorSel = &HFFEBD7
    vsRegist.ForeColorSel = &H0&
End Sub

Private Sub vsRegist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsRegist.TextMatrix(1, 0) <> "" And cbo续诊科室.ListCount = 1 Then
            Call cmdOK_Click
        End If
    End If
End Sub

Private Sub vsRegist_LostFocus()
    vsRegist.BackColorSel = &HC0C0C0
    vsRegist.ForeColorSel = &H0&
End Sub

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡对象的相关信息
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjSquare Is Nothing Then Exit Sub
    Call PatiIdentify.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquare, , "")
    PatiIdentify.objIDKind.AllowAutoICCard = True
    PatiIdentify.objIDKind.AllowAutoIDCard = True
End Sub

Private Sub SetFontSize(ByVal bytSize As Byte)
'功能：进行界面字体的统一设置
'参数：bytSize  0-9号字体，1-12号字体
    
    Me.Width = IIf(bytSize = 0, 7200, 9500)
    Me.Height = IIf(bytSize = 0, 3100, 4100)
    Call zlControl.SetPubFontSize(Me, bytSize)
    vsRegist.Height = 5 * vsRegist.RowHeight(0)
    Call SetCtlPos
    vsRegist.Width = Me.Width - 2 * vsRegist.Left
End Sub

Private Sub SetCtlPos()
'功能：设置界面控件位置
    Dim lngDis1 As Long, lngDis2 As Long
    lngDis1 = 30: lngDis2 = 120
    
    Call zlControl.SetPubCtrlPos(False, 0, lblInfo(2), lngDis1, fraKind, 0, imgSentence, lngDis1, imgStaKB)
    imgSentence.Top = imgSentence.Top - 10
    Call zlControl.SetPubCtrlPos(True, -1, lblInfo(2), 120 + Frame1(1).Height + 90, lblInfo(1), 30, vsRegist, 90 + Frame1(2).Height + 180, lbl续诊科室)
    Frame1(1).Top = lblInfo(2).Top + lblInfo(2).Height + 120
    Frame1(2).Top = vsRegist.Top + vsRegist.Height + 90
    Call zlControl.SetPubCtrlPos(False, 0, lbl续诊科室, lngDis1, cbo续诊科室, lngDis2, chk急诊, lngDis2, cmdOK, lngDis2, cmdCancel)
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - lngDis2
    cmdOK.Left = cmdCancel.Left - 60 - cmdOK.Width
End Sub
