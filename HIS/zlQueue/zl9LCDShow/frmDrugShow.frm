VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDrugShow 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicMsg 
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   7440
      Width           =   11655
      Begin VB.Timer TimerCall 
         Interval        =   1000
         Left            =   1680
         Top             =   240
      End
      Begin VB.Timer timerLCD 
         Interval        =   10000
         Left            =   10680
         Top             =   120
      End
      Begin VB.Timer timerPage 
         Interval        =   5000
         Left            =   9000
         Top             =   360
      End
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11655
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgCallingData 
      Height          =   7335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      _cx             =   20981
      _cy             =   12938
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   65280
      BackColorFixed  =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   0
      BackColorAlternate=   0
      GridColor       =   65280
      GridColorFixed  =   65280
      TreeColor       =   -2147483633
      FloodColor      =   0
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugShow.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
End
Attribute VB_Name = "frmDrugShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mintCols As Integer
Private mstrWins As String
Private mintRows As Integer
Private mrsData() As Recordset
Private mrsCallingData As Recordset
Private mrsPreparingData() As Recordset
Private mRowRec As Integer
Private mlng药房ID As Long
Private mIntCallRow As Integer
Private mIntPraRow As Integer
Private mbln配药 As Boolean
Private mbln配药确认 As Boolean
Private mstrSendNames() As String
Private mstrPraNames() As String
Private mIntSendPages() As Integer

Private Type Type_para
    bln单窗体显示模式 As Boolean             '窗体显示模式，单窗体：多窗体
    Str窗口 As String
    dblLeft As Double
    dblTop As Double
    dblWidth As Double
    dblHeight As Double
    
    lng呼叫中字体颜色 As Long
    
    bln显示待发药 As Boolean
    int待发药人数 As Integer
    int待发药行数 As Integer
    int待发药列数 As Integer
    lng待发药字体颜色 As Long
    
    bln显示待配药 As Boolean
    int待配药人数 As Integer
    int待配药行数 As Integer
    int待配药列数 As Integer
    lng待配药字体颜色 As Long
    
    bln显示窗口 As Boolean
    lng窗口字体颜色 As Long
    
    bln显示其他内容 As Boolean
    lng其他内容字体颜色 As Long
    
    
    intRowPeople  As Integer
    intPage As Integer
    intRefTime As Integer
    
    str显示内容 As String
End Type

Private mType_para As Type_para

Public Sub SetFacePostion()
'************************************************************************************
'
'设置界面的显示位置
'
'************************************************************************************
    Dim strReg As String
    
    On Error Resume Next
        
    '从注册表中，读取显示参数
    strReg = "公共模块\药房排队叫号\液晶电视"
    
    '设置显示参数
    Me.Left = GetSetting("ZLSOFT", strReg, "左", "1024") * Screen.TwipsPerPixelX
    Me.Top = GetSetting("ZLSOFT", strReg, "顶", "0") * Screen.TwipsPerPixelY
    Me.Width = GetSetting("ZLSOFT", strReg, "宽度", "1024") * Screen.TwipsPerPixelX
    Me.Height = GetSetting("ZLSOFT", strReg, "高度", "768") * Screen.TwipsPerPixelY
End Sub
Private Sub LoadPara()
    Dim strReg As String
    Dim i As Integer
    Dim strWin As String
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    
    With mType_para
        .bln单窗体显示模式 = (Val(GetSetting("ZLSOFT", strReg, "窗口模式", "0")) = 0)
        
        '加载窗口
        .Str窗口 = GetSetting("ZLSOFT", strReg, "窗口", "1,2,3")
        
        '加载屏幕信息
        .dblLeft = GetSetting("ZLSOFT", strReg, "左", "1024")
        .dblTop = GetSetting("ZLSOFT", strReg, "顶", "0")
        .dblWidth = GetSetting("ZLSOFT", strReg, "宽度", "1024")
        .dblHeight = GetSetting("ZLSOFT", strReg, "高度", "768")
        
        '呼叫中的字体颜色
        .lng呼叫中字体颜色 = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
        
        '待发药列表的设置
        .bln显示待发药 = (Val(GetSetting("ZLSOFT", strReg, "显示待发药", "1")) = 1)
        .int待发药人数 = Val(GetSetting("ZLSOFT", strReg, "待发药人数", "9"))
        .int待发药行数 = Val(GetSetting("ZLSOFT", strReg, "待发药行数", "3"))
        .int待发药列数 = Val(GetSetting("ZLSOFT", strReg, "待发药列数", "3"))
        .lng待发药字体颜色 = GetSetting("ZLSOFT", strReg, "待发药颜色", vbGreen)
        
        '待配药列表的设置
        .bln显示待配药 = (Val(GetSetting("ZLSOFT", strReg, "显示待配药", "1")) = 1)
        .int待配药人数 = Val(GetSetting("ZLSOFT", strReg, "待配药人数", "9"))
        .int待配药行数 = Val(GetSetting("ZLSOFT", strReg, "待配药行数", "9"))
        .int待配药列数 = Val(GetSetting("ZLSOFT", strReg, "待配药列数", "9"))
        .lng待配药字体颜色 = GetSetting("ZLSOFT", strReg, "待配药颜色", vbGreen)
        
        .intRowPeople = 3
        .intPage = GetSetting("ZLSOFT", strReg, "翻页时间", "5")
        .intRefTime = GetSetting("ZLSOFT", strReg, "刷新时间", "10")
        
        .bln显示窗口 = (Val(GetSetting("ZLSOFT", strReg, "显示窗口", "1")) = 1)
        .lng窗口字体颜色 = GetSetting("ZLSOFT", strReg, "窗口颜色", vbGreen)
        
        .bln显示其他内容 = (Val(GetSetting("ZLSOFT", strReg, "显示其他内容", "1")) = 1)
        .lng其他内容字体颜色 = GetSetting("ZLSOFT", strReg, "其他内容颜色", vbBlack)
        
        .str显示内容 = GetSetting("ZLSOFT", strReg, "显示内容", "")
    End With
End Sub


Private Sub InitData(ByVal intPage As Integer, ByVal blnRef As Boolean)
'***********************************************************************
'
'刷新数据：intPage=1为窗体加载时，加载数据；intpage=2为timer事件刷新数据
'
'************************************************************************
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strpeople As String
    Dim count As Integer
    Dim intcol As Integer
    Dim intSum As Integer
    Dim strTemp As String
    Dim intTemp As Integer
    Dim intCurPage As Integer
    Dim intPraPage As Integer
    Dim intCallPage As Integer
    
    '绘制待发药列表的边框
     
    If vfgCallingData.Cols = 0 Then Exit Sub
    If mType_para.bln显示待发药 Then
        vfgCallingData.Select mIntCallRow + 2, 0, mIntCallRow + 2, (mintCols) * mRowRec - 1
        vfgCallingData.CellBorder &HFF00&, -1, -1, -1, 1, 0, 1
    End If
    
    For k = 0 To mintCols - 1
        If intPage = 1 Or blnRef Then
            loadCalling (Split(mstrWins, ",")(k))
            For j = 1 To mRowRec
                intSum = intSum + 1
                Me.vfgCallingData.TextMatrix(0, intSum - 1) = Split(mstrWins, ",")(k)
                
                If k Mod 2 = 0 Then
                    strTemp = String(0, " ")
                Else
                    strTemp = String(1, " ")
                End If
                
                If Not mrsCallingData.EOF Then
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "请 " & mrsCallingData!姓名 & " 领药"
                Else
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "无呼叫人员"
                End If
            Next
        End If
        If blnRef = False Then
            '显示待发药信息
            If mType_para.bln显示待发药 Then
                loadData (Split(mstrWins, ",")(k)), k
                
                ShowSend k, intPage
            End If
        

            '显示待配药信息
            ShowPra k, intPage
        End If
        
        '画边框
        If k <> mintCols - 1 And vfgCallingData.Rows > 2 Then
            vfgCallingData.Select 2, (k + 1) * mRowRec - 1, mintRows - 1, (k + 1) * mRowRec
            vfgCallingData.CellBorder &HFF00&, -1, -1, -1, 0, 1, 0
            
            If mType_para.bln显示待发药 Then
                vfgCallingData.Select mIntCallRow + 2, (k + 1) * mRowRec - 1, mIntCallRow + 2, (k + 1) * mRowRec
                vfgCallingData.CellBorder &HFF00&, -1, -1, -1, 1, 1, 1
            End If
        End If
    Next
    
    '合并窗口和叫号信息
    vfgCallingData.MergeRow(0) = True
    vfgCallingData.MergeRow(1) = True
    vfgCallingData.Refresh
    
    vfgCallingData.Select 0, 0, 1, mintCols * mRowRec - 1
    vfgCallingData.CellBorder &HFF00&, 0, 0, 0, 1, 1, 1
End Sub


Private Sub Form_Load()
    '加载参数
    LoadPara
    
    SetFacePostion
    
    Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln显示其他内容, Round(Me.ScaleHeight * 0.9), Round(Me.ScaleHeight))
    
    '根据模式确定具体的显示窗口，单窗口是传参，多窗体是在参数设置时进行选择
    If mType_para.bln单窗体显示模式 = False Then
        mstrWins = mType_para.Str窗口
        If mstrWins = "" Then
            Exit Sub
        End If
    Else
        Me.TimerCall.Enabled = False
    End If
    
    mintCols = UBound(Split(mstrWins, ",")) + 1
'    If mintCols = 0 Then Exit Sub
    mRowRec = mType_para.intRowPeople
    
    '确认数据集数组的长度
    ReDim mrsData(mintCols)
    ReDim mrsPreparingData(mintCols)
    ReDim mstrSendNames(mintCols)
    ReDim mstrPraNames(mintCols)
    
    '初始化表格
    InitVSF
    
    InitData 1, False
    
    Me.timerPage.Interval = mType_para.intPage * 1000
    Me.timerLCD.Interval = mType_para.intRefTime * 1000
    
    Me.PicMsg.Visible = mType_para.bln显示其他内容
    Me.lblmsg.Caption = IIf(mType_para.str显示内容 = "", "祝您早日康复！  " & Format(zlDatabase.Currentdate, "yyyy-mm-dd  hh:mm"), mType_para.str显示内容)
End Sub

Private Sub loadData(ByVal strWin As String, ByVal Index As Integer)
'************************************************************************
'
'加载待发药列表的数据
'
'************************************************************************
    Dim strSql As String
    Dim date开始日期 As Date
    Dim date结束日期 As Date
        
    On Error GoTo errHandle
    date开始日期 = zlDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")
    
    date结束日期 = zlDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")

    strSql = "Select distinct A.姓名,B.配药日期,B.签到时间 From 未发药品记录 A,药品收发记录 B,门诊费用记录 C" & _
             " Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id and B.费用id=C.id and (A.单据=8 or A.单据=9 or A.单据=10) "
    
    If mbln配药 Then
        strSql = strSql & " and (A.排队状态=2 or A.排队状态=4) and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    ElseIf mbln配药确认 And mbln配药 = False Then
        strSql = strSql & " and (A.排队状态=1 or A.排队状态=2 or A.排队状态=4) and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    ElseIf mbln配药 = False And mbln配药确认 = False Then
        strSql = strSql & "  and (A.排队状态<>3 or A.排队状态 is null) and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    End If
    
    Set mrsData(Index) = zlDatabase.OpenSQLRecord(strSql, "", mlng药房ID, strWin, date开始日期, date结束日期)
    
    
    If Not mrsData(Index).EOF Then
        If Nvl(mrsData(Index)!配药日期) <> "" Then
            mrsData(Index).Sort = "配药日期"
        Else
            mrsData(Index).Sort = "签到时间"
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub loadCalling(ByVal strWin As String)
'************************************************************************
'
'加载当前呼叫的数据
'
'************************************************************************
    Dim strSql As String
    Dim date开始日期 As Date
    Dim date结束日期 As Date
    
    On Error GoTo errHandle
    date开始日期 = zlDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")
    
    date结束日期 = zlDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "select 姓名 from 未发药品记录 where 排队状态=3 and 库房id=[1] and 发药窗口=[2] and 填制日期 between [3] and [4]"
    Set mrsCallingData = zlDatabase.OpenSQLRecord(strSql, "", mlng药房ID, strWin, date开始日期, date结束日期)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub loadPreparing(ByVal strWin As String, ByVal intIndex As Integer)
'************************************************************************
'
'加载待配药列表的数据
'
'************************************************************************
    Dim strSql As String
    
    On Error GoTo errHandle
    Dim date开始日期 As Date
    Dim date结束日期 As Date
        
    On Error GoTo errHandle
    date开始日期 = zlDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")
    
    date结束日期 = zlDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "Select distinct A.姓名,B.配药日期,B.签到时间 From 未发药品记录 A,药品收发记录 B,门诊费用记录 C" & _
             " Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id and B.费用id=C.id and (A.单据=8 or A.单据=9 or A.单据=10) "
    If mbln配药确认 Then
        strSql = strSql & "and A.排队状态=1 and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    Else
        strSql = strSql & "and (A.排队状态=1 or A.排队状态=0 or A.排队状态 is null) and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    End If
    
    If mbln配药 = False Then strSql = strSql & " And 1=2"
    
    Set mrsPreparingData(intIndex) = zlDatabase.OpenSQLRecord(strSql, "", mlng药房ID, strWin, date开始日期, date结束日期)
    
    If Not mrsPreparingData(intIndex).EOF Then
        If Nvl(mrsPreparingData(intIndex)!配药日期) <> "" Then
             mrsPreparingData(intIndex).Sort = "配药日期"
        Else
             mrsPreparingData(intIndex).Sort = "签到时间"
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub InitVSF()
'************************************************************************
'
'初始化表格
'
'************************************************************************
    Dim intColWidth As Integer
    Dim intRowheight As Integer
    Dim i As Integer
    Dim strReg As String
    Dim dblHeight As Double
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    
    mintRows = 2
'    dblHeight = (20 * Val(GetSetting("ZLSOFT", strReg, "字号(0)", "14"))) * 1.5
'
'    If mType_para.bln显示窗口 Then
'        dblHeight = dblHeight + (20 * Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))) * 1.5
'        mintRows = 2
'    End If
'
'    If mType_para.bln显示待发药 = True Then
'        mIntCallRow = (Me.vfgCallingData.Height - dblHeight) * 0.6 \ ((20 * Val(GetSetting("ZLSOFT", strReg, "字号(2)", "14")) * 2))
'    End If
'
'    If mType_para.bln显示待配药 = True Then
'        mIntPraRow = (Me.vfgCallingData.Height - dblHeight) * 0.4 \ ((20 * Val(GetSetting("ZLSOFT", strReg, "字号(3)", "14")) * 2))
'    End If
'
'    If Val(GetSetting("ZLSOFT", strReg, "字号(2)", "14")) > Val(GetSetting("ZLSOFT", strReg, "字号(3)", "14")) Then
'        mRowRec = mType_para.int待发药人数 \ mIntCallRow + 1
'    Else
'        mRowRec = mType_para.int待配药人数 \ mIntPraRow + 1
'    End If
'
'    If mType_para.int待发药人数 Mod mRowRec = 0 Then
'        mIntCallRow = mType_para.int待发药人数 \ mRowRec
'    Else
'        mIntCallRow = mType_para.int待发药人数 \ mRowRec + 1
'    End If
'
'    If mType_para.int待配药人数 Mod mRowRec = 0 Then
'        mIntPraRow = mType_para.int待配药人数 \ mRowRec
'    Else
'        mIntPraRow = mType_para.int待配药人数 \ mRowRec + 1
'    End If
    
    mIntCallRow = IIf(mType_para.bln显示待发药, mType_para.int待发药行数, 0)
    mIntPraRow = IIf(mType_para.bln显示待配药, mType_para.int待配药行数, 0)
    mRowRec = mType_para.int待发药列数
    mintRows = mintRows + mIntCallRow + mIntPraRow + IIf(mType_para.bln显示待发药, 1, 0) + IIf(mType_para.bln显示待配药, 1, 0)
    With vfgCallingData
        .Rows = mintRows
        .Cols = mintCols * mRowRec
        
        If .Cols = 0 Then Exit Sub
        
'        If mintCols = 0 Then
'            Unload Me
'            Exit Sub
'        End If
        '设置表格为自由合并
        .MergeCells = flexMergeFree

         '设置字体和字体颜色大小
        SetFont
        
        intColWidth = Me.ScaleWidth / (mintCols * mRowRec)
        '设置内容居中显示
        For i = 0 To mintCols * mRowRec - 1
            .ColWidth(i) = intColWidth
            vfgCallingData.ColAlignment(i) = flexAlignCenterCenter
        Next
        
        If mType_para.bln显示待发药 = False And mType_para.bln显示待配药 = False Then
            .RowHeight(0) = (20 * Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))) * 1.5
            .RowHeight(1) = IIf(mType_para.bln显示窗口, .Height - .RowHeight(0), .Height)
        ElseIf mType_para.bln显示待发药 = False Or mType_para.bln显示待配药 = False Then
            .RowHeight(0) = (20 * Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))) * 1.5
            .RowHeight(1) = (20 * Val(GetSetting("ZLSOFT", strReg, "字号(0)", "14"))) * 2
        Else
            .RowHeight(0) = (20 * Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))) * 1.5
            .RowHeight(1) = (20 * Val(GetSetting("ZLSOFT", strReg, "字号(0)", "14"))) * 2
        End If
        
        If Not mType_para.bln显示窗口 Then
            .RowHeight(0) = 0
        End If
        If vfgCallingData.Rows > 2 Then
            intRowheight = (.Height - .RowHeight(0) - .RowHeight(1)) / (mintRows - 2)
            For i = 2 To mintRows - 1
                .RowHeight(i) = intRowheight
            Next
        End If
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln显示其他内容, Round(Me.ScaleHeight * 0.9), Round(Me.ScaleHeight))
    Me.PicMsg.Move 0, Me.vfgCallingData.Height, Me.vfgCallingData.Width, Round(Me.ScaleHeight * 0.1)
    Me.lblmsg.Move 0, Me.PicMsg.Height / 5, Me.PicMsg.Width, Me.PicMsg.Height
    
    InitVSF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    For i = 0 To mintCols - 1
        mstrSendNames(i) = ""
        mstrPraNames(i) = ""
        Set mrsData(i) = Nothing
        Set mrsData(i) = Nothing
    Next
End Sub

Private Sub TimerCall_Timer()
    InitData 2, True
End Sub

Private Sub timerPage_Timer()
'************************************************************************
'
'对待发药列表的数据进行翻页
'
'************************************************************************
    Dim i As Integer
    Dim intcol As Integer
    Dim k As Integer
    Dim count As Integer
    Dim intPage As Integer
    Dim strTemp As String
    Dim intCallPage As Integer
    Dim intPraPage As Integer

    If mType_para.bln显示待发药 = False And mType_para.bln显示待配药 = False Then Exit Sub

'    Me.timerLCD.Enabled = False

    For k = 0 To mintCols - 1
        intcol = k * mRowRec
        '计算各个窗口翻页之后的页数
'        For intcol = k * mRowRec To (k + 1) * mRowRec - 1
        If (intcol \ mRowRec) Mod 2 = 0 Then
            strTemp = String(0, " ")
        Else
            strTemp = String(1, " ")
        End If

        If mType_para.bln显示待发药 = True Then
            '待发药的总页数
            intCallPage = (mrsData(k).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(k).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))

            If intCallPage = 0 Then intCallPage = 1

            '当前页
            intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), "/"))) + 1

            If intPage > intCallPage Then intPage = 1

            Me.vfgCallingData.Cell(flexcpText, mIntCallRow + 2, intcol, mIntCallRow + 2, (k + 1) * mRowRec - 1) = "待发药   " & strTemp & intPage & "/" & intCallPage & " 共" & mrsData(k).RecordCount & "人"
        End If

    
        If mType_para.bln显示待发药 = True Then
            '当数据集有数据，但是已经移到最后一条数据了就移到第一条
            If mrsData(k).EOF And mrsData(k).RecordCount > 0 Then mrsData(k).MoveFirst
            '清空待发药列表的数据
            Me.vfgCallingData.Cell(flexcpText, 2, k * mRowRec, mIntCallRow + 1, (k + 1) * mRowRec - 1) = ""

            i = 2
            count = 0
            intcol = k * mRowRec
            If mstrSendNames(k) = "" Or intPage = 1 Then mstrSendNames(k) = ","
            Do While Not mrsData(k).EOF
                If InStr(1, mstrSendNames(k), "," & mrsData(k)!姓名 & ",") = 0 Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsData(k)!姓名)
                    
                    '到达每个窗口每行最后一列时，换到下一行显示
                    
                    intcol = intcol + 1
'                    If count = mRowRec Then
'                        count = 0
'                        intcol = k * mRowRec
'                        If Not mrsData(k).EOF Then i = i + 1
'                    End If
                End If

                mstrSendNames(k) = mstrSendNames(k) & mrsData(k)!姓名 & ","
                mrsData(k).MoveNext
                
                 '当数据的显示已经到达设置值时，退出循环
                If (i - 2) * mRowRec + count = mType_para.int待发药人数 Then
                    Exit Do
                End If
                '到达每个窗口每行最后一列时，换到下一行显示
                If count = mRowRec Then
                    count = 0
                    intcol = k * mRowRec
                    If Not mrsData(k).EOF Then i = i + 1
                End If
            Loop
        End If
        
        intcol = 0
        If mType_para.bln显示待配药 = True Then
             '当前的页数
            intPraPage = (mrsPreparingData(k).RecordCount \ (mIntPraRow * mRowRec) + IIf(mrsPreparingData(k).RecordCount Mod (mIntPraRow * mRowRec) = 0, 0, 1))

            If intPraPage = 0 Then intPraPage = 1

            intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), "/"))) + 1

            If intPage > intPraPage Then intPage = 1
            
            Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (k + 1) * mRowRec - 1) = "待配药   " & strTemp & intPage & "/" & intPraPage & " 共" & mrsPreparingData(k).RecordCount & "人"
        End If

        If mType_para.bln显示待配药 = True Then
            count = 0
            '当数据集有数据，但是已经移到最后一条数据了就移到第一条
            If mrsPreparingData(k).EOF And mrsPreparingData(k).RecordCount > 0 Then mrsPreparingData(k).MoveFirst
            '清空待发药列表的数据
            Me.vfgCallingData.Cell(flexcpText, mintRows - mIntPraRow - 1, k * mRowRec, mintRows - 2, (k + 1) * mRowRec - 1) = ""

            i = mintRows - mIntPraRow - 1
            intcol = k * mRowRec
            If mstrPraNames(k) = "" Or intPage = 1 Then mstrPraNames(k) = ","
            Do While Not mrsPreparingData(k).EOF
                If InStr(1, mstrPraNames(k), "," & mrsPreparingData(k)!姓名 & ",") = 0 Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = mrsPreparingData(k)!姓名

                    '到达每个窗口每行最后一列时，换到下一行显示
                    intcol = intcol + 1
                    
'                    If count = mRowRec Then
'                        count = 0
'                        intcol = k * mRowRec
'                        If Not mrsPreparingData(k).EOF Then i = i + 1
'                    End If
                    mstrPraNames(k) = mstrPraNames(k) & mrsPreparingData(k)!姓名 & ","
                    
                    If (i - (mintRows - mIntPraRow - 1)) * mRowRec + count = mType_para.int待配药人数 Then
                        Exit Do
                    End If
                    If count = mRowRec Then
                        count = 0
                        intcol = k * mRowRec
                        If Not mrsPreparingData(k).EOF Then i = i + 1
                    End If
                End If
                
                mrsPreparingData(k).MoveNext
                
                '当数据的显示已经到达设置值时，退出循环
                
            Loop
        End If

'        intcol = k * mRowRec
'        '计算各个窗口翻页之后的页数
''        For intcol = k * mRowRec To (k + 1) * mRowRec - 1
'        If (intcol \ mRowRec) Mod 2 = 0 Then
'            strTemp = String(0, " ")
'        Else
'            strTemp = String(1, " ")
'        End If
'
'        If mType_para.bln显示待发药 = True Then
'            '待发药的总页数
'            intCallPage = (mrsData(k).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(k).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))
'
'            If intCallPage = 0 Then intCallPage = 1
'
'            '当前页
'            intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), "/"))) + 1
'
'            If intPage > intCallPage Then intPage = 1
'
'            Me.vfgCallingData.Cell(flexcpText, mIntCallRow + 2, intcol, mIntCallRow + 2, (k + 1) * mRowRec - 1) = "待发药   " & strTemp & intPage & "/" & intCallPage & " 共" & mrsData(k).RecordCount & "人"
'        End If
'
'        If mType_para.bln显示待配药 = True Then
'             '当前的页数
'            intPraPage = (mrsPreparingData(k).RecordCount \ (mIntPraRow * mRowRec) + IIf(mrsPreparingData(k).RecordCount Mod (mIntPraRow * mRowRec) = 0, 0, 1))
'
'            If intPraPage = 0 Then intPraPage = 1
'
'            intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), "/"))) + 1
'
'            If intPage > intPraPage Then intPage = 1
'
'
'            Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (k + 1) * mRowRec - 1) = "待配药   " & strTemp & intPage & "/" & intPraPage & " 共" & mrsPreparingData(k).RecordCount & "人"
'        End If
'        Next
    Next
End Sub

Public Sub ShowMe(ByVal lng药房ID As Long, ByVal strWins As String, ByVal bln配药 As Boolean, ByVal bln配药确认 As Boolean)
'**************************************************************************
'打开窗体的接口，lng药房ID：当前的药房id；strWins：窗体连接串
'**************************************************************************
    mlng药房ID = lng药房ID
    mstrWins = strWins
    mbln配药 = bln配药
    mbln配药确认 = bln配药确认
    Dim strTemp As String
    Dim strReg As String
    Dim cls As New clsLCDShow
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    strTemp = GetSetting("ZLSOFT", strReg, "窗口", "1,2,3")
    If strTemp = "" And strWins = "" Then
        cls.zlClose
        Exit Sub
    End If
    
    Me.Show
End Sub

Private Sub timerLCD_Timer()
'************************************************************************
'
'刷新待发药列表的数据
'
'************************************************************************
    InitData 2, False
    
    Me.lblmsg.Caption = IIf(mType_para.str显示内容 = "", "祝您早日康复！  " & Format(zlDatabase.Currentdate, "yyyy-mm-dd  hh:mm"), mType_para.str显示内容)
End Sub

Public Sub ChangeCall(ByVal strWin As String, ByVal strName As String)
'****************************************************************************
'
'更新当前呼叫信息
'
'**************************************************************************

    InitData 2, True
End Sub

Private Sub ShowSend(ByVal Index As Integer, ByVal intPage As Integer)
'******************************************************************************
'
'将待发药的数据加载到待发区域
'
'******************************************************************************
    Dim count As Integer
    Dim i As Integer
    Dim intcol As Integer
    Dim intCallPage As Integer
    Dim intCurPage As Integer
    Dim strTemp As String
    Dim strNames As String
    
    '显示待发药信息
    If mType_para.bln显示待发药 Then
        '计算总的页数
        intCallPage = (mrsData(Index).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(Index).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))
        If intCallPage = 0 Then intCallPage = 1
        
        '判断是否为窗体加载
        If intPage <> 1 Then
            For i = 0 To Me.vfgCallingData.Cols - 1
                '计算当前页数
                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
                    intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, i), "/")))
                End If
            Next
            
            '将记录集的游标移向当前页显示的内容
            For i = 1 To mIntCallRow * mRowRec * (intCurPage - 1)
                If Not mrsData(Index).EOF Then
                    mrsData(Index).MoveNext
                End If
            Next
            
        Else
            intCurPage = 1
        End If
        
        count = 0
        i = 2
        intcol = Index * mRowRec
        
        '循环记录集，显示界面
        mstrSendNames(Index) = ","
        Do While Not mrsData(Index).EOF
            If InStr(1, mstrSendNames(Index), "," & mrsData(Index)!姓名 & ",") = 0 Then
                count = count + 1
                Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsData(Index)!姓名)
                
'                '每行显示了制定人数后，跳到下一行
                intcol = intcol + 1
'                If count = mRowRec Then
'                    count = 0
'                    intcol = Index * mRowRec
'                    If Not mrsData(Index).EOF Then i = i + 1
'                End If
            End If
            
            mstrSendNames(Index) = mstrSendNames(Index) & Nvl(mrsData(Index)!姓名) & ","
            '移向下一条记录
            mrsData(Index).MoveNext
            
            '界面显示了指定人数后，退出循环
            If (i - 2) * mRowRec + count = mType_para.int待发药人数 Then
                Exit Do
            End If
'            intcol = intcol + 1
            If count = mRowRec Then
                count = 0
                intcol = Index * mRowRec
                If Not mrsData(Index).EOF Then i = i + 1
            End If
            
        Loop
        
        '显示翻页信息
        intcol = Index * mRowRec
        For intcol = Index * mRowRec To (Index + 1) * mRowRec - 1
            If (intcol \ mRowRec) Mod 2 = 0 Then
                strTemp = String(0, " ")
            Else
                strTemp = String(1, " ")
            End If
            
            Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol) = "待发药   " & strTemp & intCurPage & "/" & intCallPage & " 共" & mrsData(Index).RecordCount & "人"
        Next
        '合并显示翻页信息
        vfgCallingData.MergeRow(mIntCallRow + 2) = True
    End If
End Sub

Private Sub ShowPra(ByVal Index As Integer, ByVal intPage As Integer)
    Dim count As Integer
    Dim i As Integer
    Dim intPraPage As Integer
    Dim intCurPage As Integer
    Dim intcol As Integer
    Dim strTemp As String
    
    If mType_para.bln显示待配药 Then
        '加载待配药信息
        loadPreparing (Split(mstrWins, ",")(Index)), Index
        
        '计算总页数
        intPraPage = (mrsPreparingData(Index).RecordCount \ (mIntPraRow * mRowRec) + IIf(mrsPreparingData(Index).RecordCount Mod (mIntPraRow * mRowRec) = 0, 0, 1))
        If intPraPage = 0 Then intPraPage = 1
        
        '判断是否是窗体加载
        If intPage <> 1 Then
            For i = 0 To Me.vfgCallingData.Cols - 1
                '得到当前页数
                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
                    intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), "/")))
                    Exit For
                End If
            Next
            
            '将记录集移到当前页数的位置
            For i = 1 To mIntPraRow * mRowRec * (intCurPage - 1)
                If Not mrsPreparingData(Index).EOF Then
                    mrsPreparingData(Index).MoveNext
                End If
            Next
            
        Else
            intCurPage = 1
        End If
        
        count = 0
        i = Me.vfgCallingData.Rows - mIntPraRow - 1
        intcol = Index * mRowRec
        mstrPraNames(Index) = ","
        '循环记录集，将数据显示到界面
        Do While Not mrsPreparingData(Index).EOF
            If InStr(1, mstrPraNames(Index), "," & Nvl(mrsPreparingData(Index)!姓名) & ",") = 0 Then
                count = count + 1
                Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsPreparingData(Index)!姓名)
                
                '一行数据加载完之后，跳到下一行
                intcol = intcol + 1
'                If count = mRowRec Then
'                    count = 0
'                    intcol = Index * mRowRec
'                End If
                mstrPraNames(Index) = mstrPraNames(Index) & Nvl(mrsPreparingData(Index)!姓名) & ","
            End If
            
            
            mrsPreparingData(Index).MoveNext
            
             '当界面的人数显示制定个数后，退出循环
            If (i - (Me.vfgCallingData.Rows - mIntPraRow - 1)) * mRowRec + count = mType_para.int待配药人数 Then
                Exit Do
            End If
            '一行数据加载完之后，跳到下一行
'            intcol = intcol + 1
            If count = mRowRec Then
                count = 0
                intcol = Index * mRowRec
                i = i + 1
            End If
        Loop
        
        '显示翻页信息
        intcol = Index * mRowRec
        For intcol = Index * mRowRec To (Index + 1) * mRowRec - 1
            If (intcol \ mRowRec) Mod 2 = 0 Then
                strTemp = String(0, " ")
            Else
                strTemp = String(1, " ")
            End If
            
            Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (Index + 1) * mRowRec - 1) = "待配药   " & strTemp & intCurPage & "/" & intPraPage & " 共" & mrsPreparingData(Index).RecordCount & "人"
        Next
        '合并显示翻页信息
        vfgCallingData.MergeRow(Me.vfgCallingData.Rows - 1) = True
    End If
End Sub

Public Sub SetFont()
    Dim strReg As String
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    With Me.vfgCallingData
        If .Cols = 0 Then Exit Sub
        '设置字体和字体颜色大小
        .Cell(flexcpFontSize, 1, 0, 1, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(0)", "14"))
        .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
        .Cell(flexcpFontName, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "字体(0)", "宋体")
        .Cell(flexcpFontBold, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "粗体(0)", "false")
        .Cell(flexcpFontItalic, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "斜体(0)", "false")
        If mType_para.bln显示窗口 Then
            .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "窗口颜色", vbGreen)
            .Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "字体(1)", "宋体")
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "粗体(1)", "false")
            .Cell(flexcpFontItalic, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "斜体(1)", "false")
        End If
        
        If mType_para.bln显示待发药 = True Then
            .Cell(flexcpFontSize, 2, 0, mIntCallRow + 2, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(2)", "14"))
            .Cell(flexcpForeColor, 2, 0, mIntCallRow + 2, .Cols - 1) = GetSetting("ZLSOFT", strReg, "待发药颜色", vbGreen)
            .Cell(flexcpFontName, 2, 0, mIntCallRow + 2, .Cols - 1) = GetSetting("ZLSOFT", strReg, "字体(2)", "宋体")
            .Cell(flexcpFontBold, 2, 0, mIntCallRow + 2, .Cols - 1) = GetSetting("ZLSOFT", strReg, "粗体(2)", "false")
            .Cell(flexcpFontItalic, 2, 0, mIntCallRow + 2, .Cols - 1) = GetSetting("ZLSOFT", strReg, "斜体(2)", "false")
        End If
        
        If mType_para.bln显示待配药 = True Then
            .Cell(flexcpFontSize, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(3)", "14"))
            .Cell(flexcpForeColor, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "待配药颜色", vbGreen)
            .Cell(flexcpFontName, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "字体(3)", "宋体")
            .Cell(flexcpFontBold, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "粗体(3)", "false")
            .Cell(flexcpFontItalic, mintRows - mIntPraRow - 1, 0, mintRows - 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "斜体(3)", "false")
        End If
    End With
    
    Me.lblmsg.ForeColor = GetSetting("ZLSOFT", strReg, "其他内容颜色", vbBlack)
    Me.lblmsg.FontSize = Val(GetSetting("ZLSOFT", strReg, "字号(4)", "14"))
    Me.lblmsg.FontName = GetSetting("ZLSOFT", strReg, "字体(4)", "宋体")
    Me.lblmsg.FontBold = GetSetting("ZLSOFT", strReg, "粗体(4)", "false")
    Me.lblmsg.FontItalic = GetSetting("ZLSOFT", strReg, "斜体(4)", "false")
End Sub

'Private Sub GetTotal(ByVal intType As Integer)
'    Dim i As Integer
'    Dim strTemp As String
'
'    If intType = 0 Then
'        For i = 0 To mintCols - 1
'            strTemp = ","
'            If Not mrsData(i) Is Nothing Then mrsData(i).MoveFirst
'            Do While mrsData(i).EOF
'
'                If InStr(1, strTemp, "," & mrsData(i)!姓名 & ",") Then
'                    mintSenpages(i) = mintSenpages(i) + 1
'                End If
'                strTemp = strTemp & mrsData(i)!姓名 & ","
'                mrsData(i).MoveNext
'            Loop
'            mrsData(i).MoveFirst
'        Next
'    Else
'
'    End If
'End Sub


