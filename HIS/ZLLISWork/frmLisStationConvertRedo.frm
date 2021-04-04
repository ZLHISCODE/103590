VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisStationConvertRedo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "转为新标本的重做结果"
   ClientHeight    =   5460
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8205
   Icon            =   "frmLisStationConvertRedo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk 
      Caption         =   "替换申请标本结果(&R)"
      Height          =   255
      Index           =   1
      Left            =   2220
      TabIndex        =   19
      Top             =   4935
      Width           =   2115
   End
   Begin VB.CheckBox chk 
      Caption         =   "使用无主标本号(&U)"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   18
      Top             =   4935
      Value           =   1  'Checked
      Width           =   1830
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6930
      TabIndex        =   8
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5700
      TabIndex        =   7
      Top             =   4860
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "申请标本"
      Height          =   4560
      Left            =   4080
      TabIndex        =   3
      Top             =   75
      Width           =   4095
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&P"
         Height          =   300
         Left            =   3585
         TabIndex        =   15
         Top             =   900
         Width           =   300
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   555
         Width           =   2805
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   900
         Width           =   2475
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   3255
         Left            =   60
         TabIndex        =   4
         Top             =   1260
         Width           =   3960
         _cx             =   6985
         _cy             =   5741
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
         BackColorSel    =   16768667
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483639
         FocusRect       =   2
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   240
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   72613891
         CurrentDate     =   38229
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   2595
         TabIndex        =   16
         Top             =   180
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   72613891
         CurrentDate     =   38229
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   4
         Left            =   2355
         TabIndex        =   17
         Top             =   240
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.检验仪器"
         Height          =   180
         Index           =   3
         Left            =   90
         TabIndex        =   14
         Top             =   615
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.标本时间"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   13
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.标本号码"
         Height          =   180
         Index           =   5
         Left            =   90
         TabIndex        =   6
         Top             =   960
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "无主标本"
      Height          =   4560
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   4035
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   3285
         Left            =   60
         TabIndex        =   1
         Top             =   1230
         Width           =   3915
         _cx             =   6906
         _cy             =   5794
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
         BackColorSel    =   16768667
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483639
         FocusRect       =   2
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   240
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "标本号码:"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   930
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "检验仪器:"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   615
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "标本时间:"
         Height          =   180
         Index           =   7
         Left            =   90
         TabIndex        =   2
         Top             =   300
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmLisStationConvertRedo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mfrmMain As Form
Private mlngLoop As Long
Private mRs As New ADODB.Recordset
Private mstrSQL As String
Private mblnChangeEdit As Boolean
Private mlngKey As Long  '标本ID
Private mstrName As String

Private Function OpenSelect(ByVal strText As String) As Byte
    '-----------------------------------------------------------------------------------------
    '功能:打开列表结构的诊疗项目数据
    '返回:出错返回2;成功返回1;取消返回0
    '-----------------------------------------------------------------------------------------
    Dim strInput As String
    Dim rs As New ADODB.Recordset
    Dim strLvw As String
    Dim objPoint As POINTAPI
    Dim strStart As String
    Dim strEnd As String
    
    On Error GoTo ErrHand
    
    OpenSelect = 2
    
    strLvw = "标本时间,900,0,1;检验仪器,2400,0,0;标本序号,1200,0,0"
    
    strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "正检验标本范围", "今  天"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "正检验标本范围", "今  天"), 2)
    If strStart = "" Then strStart = GetDateTime("今  天", 1)
    If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    
    strInput = "'%" & strText & "%'"
    
    mstrSQL = "SELECT DISTINCT E.ID," & _
                  "E.标本序号," & _
                  "TO_CHAR(E.核收时间,'MM-DD HH24:MI') AS 标本时间," & _
                  "D.名称 AS 检验仪器 " & _
             "FROM 病人医嘱记录 A,检验仪器 D,检验标本记录 E,病人医嘱发送 F " & _
            "WHERE E.医嘱id=A.相关ID " & _
                  "AND A.医嘱状态=8 AND A.ID=F.医嘱id AND E.仪器id=D.ID(+) " & _
                  IIf(cbo(0).ListIndex = 0, "", " AND E.仪器id=[1] ") & _
                  "AND E.标本序号 LIKE [2] " & _
                  "AND F.执行状态=3 AND A.执行科室ID+0= [3] " & _
                  "AND A.开嘱时间 BETWEEN [4] and [5] " & _
                  "AND E.核收时间 BETWEEN [6] and [7] "
                  
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, cbo(0).ItemData(cbo(0).ListIndex), strInput, mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex), _
             CDate(Format(strStart, "yyyy-MM-dd hh:mm:ss")), CDate(Format(strEnd, "yyyy-MM-dd hh:mm:ss")), CDate(Format(dtp(0).Value & "00:00:00", "yyyy-MM-dd hh:mm:ss")), _
             CDate(Format(dtp(1).Value & "23:59:59", "yyyy-MM-dd hh:mm:ss")))
    
    If rs.BOF Then
        OpenSelect = 0
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then GoTo Over
            
    Call ClientToScreen(txt(1).Hwnd, objPoint)
    If frmSelectList.ShowSelect(Me, rs, strLvw, objPoint.x * 15 - 30, objPoint.y * 15 + txt(1).Height - 30, 6000, 4200, Me.Name & "\检验项目选择", "请从下表中选择一个项目") Then
        GoTo Over
    End If
    Exit Function
Over:
    txt(1).Text = zlCommFun.Nvl(rs("标本序号").Value)
    cmdOpen.Tag = zlCommFun.Nvl(rs("ID").Value)
    txt(1).Tag = ""
    
    OpenSelect = 1
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function RefreshData(ByVal lngKey As Long) As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    vsfDetail.Rows = 2
    vsfDetail.Cell(flexcpText, 1, 0, 1, vsfDetail.Cols - 1) = ""
    
    mstrSQL = "SELECT ROWNUM AS 序号,B.中文名 AS 检验项目," & _
                "A.检验结果," & _
                "DECODE(A.结果标志,3,'偏高',2,'偏低',1,'正常',4,'阳性',5,'阴性','') AS 结果标志 " & _
                "FROM 检验普通结果 A,诊治所见项目 B,检验标本记录 D " & _
                "WHERE A.检验项目id = B.ID " & _
                    "AND A.记录类型 =D.报告结果 " & _
                    "AND D.ID=A.检验标本ID " & _
                    "AND D.ID=" & lngKey
                    
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
    
    If rs.BOF = False Then
        vsfDetail.TextMatrix(0, 0) = "序号"
        Call FillGrid(vsfDetail, rs)
        vsfDetail.TextMatrix(0, 0) = ""
    End If
    
    RefreshData = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub AdjustEnableState()
    '-----------------------------------------------------------------------------------------
    '功能:根据修改状态设置按钮、菜单等的可用状态
    '-----------------------------------------------------------------------------------------
'    cmd(2).Enabled = True
'
'    If mblnChangeEdit = False Then cmd(2).Enabled = False
'
'    tbrThis.Buttons("审核").Enabled = cmd(2).Enabled
        
End Sub

Private Sub RefreshStatus()
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
'    If vsf.Rows = 2 And Trim(vsf.TextMatrix(1, 1)) = "" Then
'        stbThis.Panels(2).Text = "没有标本信息。"
'    Else
'        stbThis.Panels(2).Text = "共找到 " & vsf.Rows - 1 & " 个标本信息。"
'    End If
'
End Sub

Public Function ShowEdit(ByVal frmMain As Form, ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：显示本编辑窗体
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
            
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    If ReadData = False Then Exit Function
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
    lngKey = mlngKey
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHand
    
    vsf.Cols = 0
    Call NewColumn(vsf, "", 240, 4)
    Call NewColumn(vsf, "检验项目", 1500, 1)
    Call NewColumn(vsf, "检验结果", 900, 1)
    Call NewColumn(vsf, "结果标志", 810, 1)
    vsf.FixedCols = 1
    
    vsfDetail.Cols = 0
    Call NewColumn(vsfDetail, "", 240, 4)
    Call NewColumn(vsfDetail, "检验项目", 1500, 1)
    Call NewColumn(vsfDetail, "检验结果", 900, 1)
    Call NewColumn(vsfDetail, "结果标志", 810, 1)
    vsfDetail.FixedCols = 1
    
    InitData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strWhere As String
    
    On Error GoTo ErrHand
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""

    mstrSQL = "SELECT ROWNUM AS 序号,B.中文名 AS 检验项目,D.核收时间,D.标本序号,C.名称," & _
                "A.检验结果," & _
                "DECODE(A.结果标志,3,'偏高',2,'偏低',1,'正常',4,'阳性',5,'阴性','') AS 结果标志 " & _
                "FROM 检验普通结果 A,诊治所见项目 B,检验仪器 C,检验标本记录 D " & _
                "WHERE A.检验项目id = B.ID " & _
                    "AND A.记录类型 =D.报告结果 " & _
                    "AND D.仪器id =C.ID(+) " & _
                    "AND D.ID=A.检验标本ID " & _
                    "AND D.ID= [1] "
                    
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mlngKey)
    
    If rs.BOF = False Then
        
        lbl(7).Caption = "标本时间:" & Format(zlCommFun.Nvl(rs("核收时间")), "yyyy-mm-dd hh:mm")
        lbl(0).Caption = "检验仪器:" & zlCommFun.Nvl(rs("标本序号"))
        lbl(1).Caption = "标本号码:" & zlCommFun.Nvl(rs("名称"), "无")
        
        vsf.TextMatrix(0, 0) = "序号"
        Call FillGrid(vsf, rs)
        vsf.TextMatrix(0, 0) = ""
        
    End If
    
    ReadData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strError As String
    

        
    ValidData = True
    
    Exit Function
ErrHand:
    MsgBox strError, vbInformation, gstrSysName
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strNow As String
    
    Dim strsql() As String
    
    On Error GoTo ErrHand
    ReDim strsql(1 To 1)
    
    strsql(ReDimArray(strsql)) = "ZL_检验标本记录_转为重做(" & mlngKey & "," & Val(cmdOpen.Tag) & "," & chk(1).Value & "," & chk(0).Value & ")"
        
    blnTran = True
    
    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strsql)
        If strsql(mlngLoop) <> "" Then Call ExecuteProc(strsql(mlngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    If chk(0).Value = 0 Then mlngKey = Val(cmdOpen.Tag)
        
    
    SaveData = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function


Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_Click(Index As Integer)
    If Index = 0 Then
        chk(1).Value = 0
        chk(1).Enabled = (chk(0).Value = 0)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Val(cmdOpen.Tag) > 0 Then
        
        If ValidData = False Then Exit Sub
        
        If SaveData = False Then Exit Sub
        
        mblnOK = True
        
        cmdOpen.Tag = ""
        
        Unload Me
        
    End If
End Sub

Private Sub cmdOpen_Click()
    Select Case OpenSelect("")
    Case 0
        '没有匹配的项目
        MsgBox "没有找到相匹配的结果！", vbInformation, gstrSysName
    Case 1
        '选取了一个项目
        mstrName = txt(1).Text
        
        Call RefreshData(Val(cmdOpen.Tag))
        
    End Select
    txt(1).SetFocus
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False

    
    '检验仪器
    mstrSQL = "SELECT A.编码||'-'||A.名称,ID FROM 检验仪器 A ORDER BY A.编码||'-'||A.名称"
    Call OpenRecord(rs, mstrSQL, Me.Caption)
    cbo(0).AddItem "所有仪器"
    If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
    If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    
    dtp(0).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtp(1).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    cbo(0).ListIndex = 0
    txt(1).Text = ""
    
    dtp(0).SetFocus
    
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 1 Then txt(1).Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Dim strInput As String
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    If KeyAscii = vbKeyReturn Then
        If Index = 1 Then
            If txt(1).Tag <> "" Then
                txt(1).Tag = ""
                Select Case OpenSelect(txt(1).Text)
                Case 0
                    '没有匹配的项目
                    MsgBox "没有找到相匹配的结果！", vbInformation, gstrSysName
                    txt(1).Text = mstrName
                    
                Case 1
                    '选取了一个项目
                    mstrName = txt(1).Text
                    Call RefreshData(Val(cmdOpen.Tag))
                Case 2
                    '取消了本次选择
                    txt(1).Text = mstrName
                End Select
            Else
                zlCommFun.PressKey vbKeyTab
                zlCommFun.PressKey vbKeyTab
            End If
            txt(1).Tag = ""
        Else
            zlCommFun.PressKey vbKeyTab
            zlCommFun.PressKey vbKeyTab
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "ZXCVBNMASDFGHJKLQWERTYUIOP01234567890,-")
    End If
End Sub
Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Index = 1 Then
        If (txt(1).Tag = "Changed") Then txt(1).Text = mstrName
    End If
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChangeEdit = True
    Call AdjustEnableState
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    
    If NewRow + 1 > vsf.FixedRows And OldRow + 1 > vsf.FixedRows Then
        vsf.Cell(flexcpBackColor, OldRow, 1, OldRow, vsf.Cols - 1) = vsf.BackColor
        vsf.Cell(flexcpBackColor, NewRow, 1, NewRow, vsf.Cols - 1) = vsf.BackColorSel
    End If
'
'    If NewRow <> OldRow Then
'        Call RefreshData(vsf.RowData(NewRow))
'    End If
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf.RowData(Row)) = 0 Then Cancel = True
    If Col <> 0 Then Cancel = True
    
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    
    If NewRow + 1 > vsfDetail.FixedRows And OldRow + 1 > vsfDetail.FixedRows Then
        vsfDetail.Cell(flexcpBackColor, OldRow, 1, OldRow, vsfDetail.Cols - 1) = vsfDetail.BackColor
        vsfDetail.Cell(flexcpBackColor, NewRow, 1, NewRow, vsfDetail.Cols - 1) = vsfDetail.BackColorSel
    End If
End Sub

Private Sub vsfDetail_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
