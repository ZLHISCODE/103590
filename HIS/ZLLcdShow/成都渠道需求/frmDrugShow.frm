VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDrugShow 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8736
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11640
   ForeColor       =   &H80000001&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8736
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timerPage 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7500
      Top             =   8160
   End
   Begin VB.Timer timerLCD 
      Interval        =   10000
      Left            =   9090
      Top             =   8190
   End
   Begin VB.Timer TimerCall 
      Interval        =   1000
      Left            =   90
      Top             =   8160
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgCallingData 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
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
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   7440
      Width           =   11655
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
Private mrsTimeoutData() As Recordset
Private mRowRec As Integer
Private mlng药房ID As Long
Private mIntCallRow As Integer
Private mIntPraRow As Integer
Private mIntCallCol As Integer
Private mIntPraCol As Integer
Private mIntTimeoutCol As Integer
Private mbln配药 As Boolean
Private mbln配药确认 As Boolean
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
    bln显示发药序号 As Boolean
    int待发药人数 As Integer
    int待发药行数 As Integer
    int待发药列数 As Integer
    lng待发药字体颜色 As Long
    
    bln显示待配药 As Boolean
    int待配药人数 As Integer
    int待配药行数 As Integer
    int待配药列数 As Integer
    lng待配药字体颜色 As Long
    
    bln显示已过号 As Boolean
    int已过号人数 As Integer
    int已过号行数 As Integer
    int已过号列数 As Integer
    lng已过号字体颜色 As Long
    lng当前过号页数 As Long
    
    bln显示窗口 As Boolean
    lng窗口字体颜色 As Long
    
    bln显示其他内容 As Boolean
    lng其他内容字体颜色 As Long
    
    
    intRowPeople  As Integer
    intPage As Integer
    intRefTime As Integer
    intTimeout As Integer
    
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
    
    On Error GoTo errHandle
        
    '从注册表中，读取显示参数
    strReg = "公共模块\药房排队叫号\液晶电视"
    
    '设置显示参数
    Me.Left = GetSetting("ZLSOFT", strReg, "左", "1024") * Screen.TwipsPerPixelX
    Me.Top = GetSetting("ZLSOFT", strReg, "顶", "0") * Screen.TwipsPerPixelY
    Me.Width = GetSetting("ZLSOFT", strReg, "宽度", "1024") * Screen.TwipsPerPixelX
    Me.Height = GetSetting("ZLSOFT", strReg, "高度", "768") * Screen.TwipsPerPixelY
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************SetFacePostion*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Sub LoadPara()
    Dim strReg As String
    Dim i As Integer
    Dim strWin As String
    Dim rsWin As New Recordset
    Dim strWins_temp As String, strSql As String
    On Error GoTo errHandle
    strReg = "公共模块\药房排队叫号\液晶电视"
    
    With mType_para
        .bln单窗体显示模式 = (Val(GetSetting("ZLSOFT", strReg, "窗口模式", "0")) = 0)
        
        '加载窗口
        .Str窗口 = GetSetting("ZLSOFT", strReg, "窗口", "1,2,3")
        strWins_temp = "'" & Replace(.Str窗口, ",", "','") & "'"
        '对窗口号进行重新排序
        strSql = "Select TO_CHAR(WMSYS.WM_CONCAT(名称)) 名称 From (select 名称 from 发药窗口 where 药房ID=[1] and 名称 in (" & strWins_temp & ") order by 编码)"
        Set rsWin = gobjDatabase.OpenSQLRecord(strSql, "", mlng药房ID)
        If rsWin.RecordCount > 0 Then
            .Str窗口 = Nvl(rsWin!名称)
        End If
        
        '加载屏幕信息
        .dblLeft = GetSetting("ZLSOFT", strReg, "左", "1024")
        .dblTop = GetSetting("ZLSOFT", strReg, "顶", "0")
        .dblWidth = GetSetting("ZLSOFT", strReg, "宽度", "1024")
        .dblHeight = GetSetting("ZLSOFT", strReg, "高度", "768")
        
        '呼叫中的字体颜色
        .lng呼叫中字体颜色 = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
        
        '待发药列表的设置
        .bln显示待发药 = (Val(GetSetting("ZLSOFT", strReg, "显示待发药", "1")) = 1)
        .bln显示发药序号 = (Val(GetSetting("ZLSOFT", strReg, "待发药序号", "0")) = 1)
        .int待发药人数 = Val(GetSetting("ZLSOFT", strReg, "待发药人数", "10"))
        .int待发药行数 = Val(GetSetting("ZLSOFT", strReg, "待发药行数", "5"))
        .int待发药列数 = Val(GetSetting("ZLSOFT", strReg, "待发药列数", "2"))
        .lng待发药字体颜色 = GetSetting("ZLSOFT", strReg, "待发药颜色", vbGreen)
        
        '待配药列表的设置
        .bln显示待配药 = (Val(GetSetting("ZLSOFT", strReg, "显示待配药", "1")) = 1)
        .int待配药人数 = Val(GetSetting("ZLSOFT", strReg, "待配药人数", "10"))
        .int待配药行数 = Val(GetSetting("ZLSOFT", strReg, "待配药行数", "5"))
        .int待配药列数 = Val(GetSetting("ZLSOFT", strReg, "待配药列数", "2"))
        .lng待配药字体颜色 = GetSetting("ZLSOFT", strReg, "待配药颜色", vbGreen)
        
        '待配药列表的设置
        .bln显示已过号 = (Val(GetSetting("ZLSOFT", strReg, "显示已过号", "1")) = 1)
        .int已过号人数 = Val(GetSetting("ZLSOFT", strReg, "已过号人数", "5"))
        .int已过号行数 = Val(GetSetting("ZLSOFT", strReg, "已过号行数", "5"))
        .int已过号列数 = Val(GetSetting("ZLSOFT", strReg, "已过号列数", "1"))
        .lng已过号字体颜色 = GetSetting("ZLSOFT", strReg, "已过号颜色", vbGreen)
        
        .intRowPeople = 5
        .intPage = GetSetting("ZLSOFT", strReg, "翻页时间", "5")
        .intRefTime = GetSetting("ZLSOFT", strReg, "刷新时间", "10")
'        .intTimeout = GetSetting("ZLSOFT", strReg, "过号时间", "10")
        
        .bln显示窗口 = (Val(GetSetting("ZLSOFT", strReg, "显示窗口", "1")) = 1)
        .lng窗口字体颜色 = GetSetting("ZLSOFT", strReg, "窗口颜色", vbGreen)
        
        .bln显示其他内容 = (Val(GetSetting("ZLSOFT", strReg, "显示其他内容", "1")) = 1)
        .lng其他内容字体颜色 = GetSetting("ZLSOFT", strReg, "其他内容颜色", vbBlack)
        
        .str显示内容 = GetSetting("ZLSOFT", strReg, "显示内容", "")
    End With
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadPara*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
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
    Dim rsTemp As Recordset
    Dim strSql As String
    '绘制待发药列表的边框
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "公共模块\药房排队叫号\液晶电视"
    If vfgCallingData.Cols = 0 Then Exit Sub
    If mType_para.bln显示待发药 Or mType_para.bln显示待配药 Or mType_para.bln显示已过号 Then
        vfgCallingData.Select mintRows - 1, 0, mintRows - 1, (mintCols) * mRowRec - 1
        'vfgCallingData.CellBorder &HFF00&, -1, -1, -1, 1, 0, 1
        vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), -1, -1, -1, 1, 0, 1
    End If
'    strsql = "Select 名称 From 部门表 Where ID=[1]"
'    Set rsTemp = gobjDatabase.OpenSQLRecord(strsql, "", mlng药房ID)
    For k = 0 To mintCols - 1
        'wwx timerCall 刷新
        If intPage = 2 Or blnRef Then
            loadCalling (Split(mstrWins, ",")(k))
            For j = 1 To mRowRec
                intSum = intSum + 1
                'Me.vfgCallingData.TextMatrix(0, intSum - 1) = rsTemp!名称 & "  " & Split(mstrWins, ",")(k)
                Me.vfgCallingData.Cell(flexcpFontSize, 0, intSum - 1, 0, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "窗口颜色", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "字体(1)", "宋体")
                Me.vfgCallingData.Cell(flexcpFontBold, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "粗体(1)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "斜体(1)", "false")
                Me.vfgCallingData.TextMatrix(0, intSum - 1) = Split(mstrWins, ",")(k)
                If k Mod 2 = 0 Then
                    strTemp = String(0, " ")
                Else
                    strTemp = String(1, " ")
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, 1, intSum - 1, 1, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(0)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "字体(0)", "宋体")
                Me.vfgCallingData.Cell(flexcpFontBold, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "粗体(0)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "斜体(0)", "false")
                If Not mrsCallingData.EOF Then
                    SaveDebug (mrsCallingData.RecordCount)
                    SaveDebug ("正在呼叫姓名：" & mrsCallingData!姓名)
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "请 " & mrsCallingData!姓名 & " 领药"
                Else
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "无呼叫人员"
                End If
            Next
        End If
        
        If intPage = 1 Or blnRef Then
            loadCalling (Split(mstrWins, ",")(k))
            For j = 1 To mRowRec
                intSum = intSum + 1
                'Me.vfgCallingData.TextMatrix(0, intSum - 1) = rsTemp!名称 & "  " & Split(mstrWins, ",")(k)
                Me.vfgCallingData.Cell(flexcpFontSize, 0, intSum - 1, 0, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "窗口颜色", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "字体(1)", "宋体")
                Me.vfgCallingData.Cell(flexcpFontBold, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "粗体(1)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "斜体(1)", "false")
                Me.vfgCallingData.TextMatrix(0, intSum - 1) = Split(mstrWins, ",")(k)
                If k Mod 2 = 0 Then
                    strTemp = String(0, " ")
                Else
                    strTemp = String(1, " ")
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, 1, intSum - 1, 1, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(0)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "字体(0)", "宋体")
                Me.vfgCallingData.Cell(flexcpFontBold, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "粗体(0)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "斜体(0)", "false")
                If Not mrsCallingData.EOF Then
                    SaveDebug (mrsCallingData.RecordCount)
                    SaveDebug ("正在呼叫姓名：" & mrsCallingData!姓名)
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "请 " & mrsCallingData!姓名 & " 领药"
                Else
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "无呼叫人员"
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, 2, intSum - 1, 2, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 2, intSum - 1, 2, intSum - 1) = GetSetting("ZLSOFT", strReg, "窗口颜色", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 2, intSum - 1, 2, intSum - 1) = GetSetting("ZLSOFT", strReg, "字体(1)", "宋体")
                Me.vfgCallingData.Cell(flexcpFontBold, 2, intSum - 1, 2, intSum - 1) = GetSetting("ZLSOFT", strReg, "粗体(1)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 2, intSum - 1, 2, intSum - 1) = GetSetting("ZLSOFT", strReg, "斜体(1)", "false")
                If mType_para.bln显示待配药 Then
                    If intSum - k * mRowRec <= mType_para.int待配药列数 Then
                        Me.vfgCallingData.TextMatrix(2, intSum - 1) = "待配药"
                    End If
                End If
                If mType_para.bln显示待发药 Then
                    If mType_para.bln显示待配药 Then
                        If intSum - k * mRowRec <= mType_para.int待发药列数 + mType_para.int待配药列数 And intSum - k * mRowRec > mType_para.int待配药列数 Then
                            Me.vfgCallingData.TextMatrix(2, intSum - 1) = "待发药"
                        End If
                    Else
                        If intSum - k * mRowRec <= mType_para.int待发药列数 Then
                            Me.vfgCallingData.TextMatrix(2, intSum - 1) = "待发药"
                        End If
                    End If
                End If
                If mType_para.bln显示已过号 Then
                    If mType_para.bln显示待配药 And mType_para.bln显示待发药 Then
                        If intSum - k * mRowRec <= mType_para.int待发药列数 + mType_para.int待配药列数 + mType_para.int已过号列数 And intSum - k * mRowRec > mType_para.int待配药列数 + mType_para.int待发药列数 Then
                            Me.vfgCallingData.TextMatrix(2, intSum - 1) = "过号"
                        End If
                    Else
                        If Not mType_para.bln显示待配药 And mType_para.bln显示待发药 Then
                            If intSum - k * mRowRec <= mType_para.int待发药列数 + mType_para.int已过号列数 And intSum - k * mRowRec > mType_para.int待发药列数 Then
                                Me.vfgCallingData.TextMatrix(2, intSum - 1) = "过号"
                            End If
                        Else
                            If mType_para.bln显示待配药 And Not mType_para.bln显示待发药 Then
                                If intSum - k * mRowRec <= mType_para.int待配药列数 + mType_para.int已过号列数 And intSum - k * mRowRec > mType_para.int待配药列数 Then
                                    Me.vfgCallingData.TextMatrix(2, intSum - 1) = "过号"
                                End If
                            Else
                                If intSum - k * mRowRec <= mType_para.int已过号列数 Then
                                    Me.vfgCallingData.TextMatrix(2, intSum - 1) = "过号"
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If blnRef = False Then
            '清空待发药列表的数据
            Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec, mIntCallRow + 2, (k + 1) * mRowRec - 1) = ""
            '显示待配药信息
            ShowPra k, intPage
            '显示待发药信息
            If mType_para.bln显示待发药 Then
                loadData (Split(mstrWins, ",")(k)), k
                ShowSend k, intPage
            End If
            '显示已过号信息
            ShowTimeout k, intPage
        End If
        If k <= mintCols - 1 And vfgCallingData.Rows > 2 Then
            '左,上,右,底部,竖直,水平线
            '画配药、发药、过号之间的分割线
            If mType_para.bln显示待配药 Then
                vfgCallingData.Select 3, mType_para.int待配药列数 - 1 + k * mRowRec, mintRows - 1, mType_para.int待配药列数 - 1 + k * mRowRec
                vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), 0, 0, 1, 0, 1, 0
                vfgCallingData.Select mintRows - 1, mType_para.int待配药列数 - 1 + k * mRowRec, mintRows - 1, mType_para.int待配药列数 - 1 + k * mRowRec
                vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), 0, 0, 1, 1, 1, 1
                If mType_para.bln显示待发药 And mType_para.bln显示已过号 Then
                    vfgCallingData.Select 3, mType_para.int待发药列数 + mType_para.int待配药列数 - 1 + k * mRowRec, mintRows - 1, mType_para.int待配药列数 + mType_para.int待发药列数 - 1 + k * mRowRec
                    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), 0, 0, 1, 0, 1, 0
                    vfgCallingData.Select mintRows - 1, mType_para.int待配药列数 + mType_para.int待发药列数 - 1 + k * mRowRec, mintRows - 1, mType_para.int待配药列数 + mType_para.int待发药列数 - 1 + k * mRowRec
                    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), 0, 0, 1, 1, 1, 1
                End If
            Else
                If mType_para.bln显示待发药 Then
                    vfgCallingData.Select 3, mType_para.int待发药列数 - 1 + k * mRowRec, mintRows - 1, mType_para.int待发药列数 - 1 + k * mRowRec
                    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), 0, 0, 1, 0, 1, 0
                    vfgCallingData.Select mintRows - 1, mType_para.int待发药列数 - 1 + k * mRowRec, mintRows - 1, mType_para.int待发药列数 - 1 + k * mRowRec
                    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), 0, 0, 1, 1, 1, 1
                End If
            End If
        End If
        '画边框
        If k <> mintCols - 1 And vfgCallingData.Rows > 2 Then
            vfgCallingData.Select 3, (k + 1) * mRowRec - 1, mintRows - 1, (k + 1) * mRowRec
            vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), -1, -1, -1, 0, 1, 0
           
            If mType_para.bln显示待发药 Or mType_para.bln显示待配药 Or mType_para.bln显示已过号 Then
                vfgCallingData.Select mIntCallRow + 2, (k + 1) * mRowRec - 1, mIntCallRow + 2, (k + 1) * mRowRec
                vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), -1, -1, -1, 1, 1, 1
            End If
        End If
    Next
    
    '合并窗口和叫号信息
    vfgCallingData.MergeRow(0) = True
    vfgCallingData.MergeRow(1) = True
    vfgCallingData.MergeRow(2) = True
    vfgCallingData.Refresh
    
    vfgCallingData.Select 0, 0, 2, mintCols * mRowRec - 1
    'vfgCallingData.CellBorder &HFF00&, 0, 0, 0, 1, 1, 1
    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "框体颜色", vbGreen), 0, 0, 0, 1, 1, 1
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************InitData*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub


Private Sub Form_Load()
    Dim strReg As String
    On Error GoTo errHandle
    '加载参数
    LoadPara
    
    SetFacePostion
    strReg = "公共模块\药房排队叫号\液晶电视"
'    If 25 * Val(GetSetting("ZLSOFT", strReg, "字号(4)", "14")) > Round(Me.ScaleHeight * 0.9) Then
    Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln显示其他内容, Round(Me.ScaleHeight - 30 * Val(GetSetting("ZLSOFT", strReg, "字号(4)", "14"))), Round(Me.ScaleHeight))
    Me.BackColor = vbBlack
    Me.lblmsg.Move 0, Me.vfgCallingData.Height + 100, Me.vfgCallingData.Width, Me.ScaleHeight

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
    ReDim mrsTimeoutData(mintCols)
    ReDim mstrSendNames(mintCols)
    ReDim mstrPraNames(mintCols)
    ReDim mstrTimeoutNames(mintCols)
    '初始化表格
    InitVSF
    
    InitData 1, False
    
    Me.timerPage.Interval = mType_para.intPage * 1000
    Me.timerLCD.Interval = mType_para.intRefTime * 1000

    Me.lblmsg.Visible = mType_para.bln显示其他内容
    Me.lblmsg.Caption = IIf(mType_para.str显示内容 = "", "祝您早日康复！", mType_para.str显示内容) & "   " & Format(gobjDatabase.Currentdate, "yyyy-mm-dd  hh:mm")
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************Load*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
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
    date开始日期 = gobjDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")

    date结束日期 = gobjDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")

    strSql = "Select A.病人ID,A.姓名,B.配药日期,B.签到时间,B.填制日期 " & _
             "From 未发药品记录 A,药品收发记录 B,部门性质说明 D " & _
             "Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id And (A.单据=8 or A.单据=9 or A.单据=10) and A.库房id=D.部门ID " & _
             "  And D.工作性质 in ('西药房','中药房')"
             '& IIf(mType_para.bln显示已过号, " And round((SYSDATE-Nvl(A.呼叫时间,SYSDATE))*24*60*60)<=" & mType_para.intTimeout, "")
    If mbln配药 Then
        strSql = strSql & " and (A.排队状态=2 " & IIf(mType_para.bln显示已过号, "", " or A.排队状态=4") & ") and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
        'strSQL = strSQL & " and (A.排队状态=2 or A.排队状态=4) and A.库房id=52 and A.发药窗口='窗口1' and A.填制日期 >sysdate-2 And (B.记录状态=1 Or Mod(B.记录状态,3)=0) "
    ElseIf mbln配药确认 And mbln配药 = False Then
        strSql = strSql & " and (A.排队状态=1 or A.排队状态=2 " & IIf(mType_para.bln显示已过号, "", " or A.排队状态=4") & ") and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
        'strSQL = strSQL & " and (A.排队状态=1 or A.排队状态=2 or A.排队状态=4) and A.库房id=52 and A.发药窗口='窗口1' and A.填制日期 >sysdate-2And (B.记录状态=1 Or Mod(B.记录状态,3)=0) "
    ElseIf mbln配药 = False And mbln配药确认 = False Then
        strSql = strSql & "  and (A.排队状态<3 or A.排队状态 is null " & IIf(mType_para.bln显示已过号, "", " or A.排队状态=4") & ") and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
        'strSQL = strSQL & "  and (A.排队状态<>3 or A.排队状态 is null) and A.库房id=52 and A.发药窗口='窗口1' and A.填制日期>sysdate-2 And (B.记录状态=1 Or Mod(B.记录状态,3)=0) "
    End If
    strSql = "Select Rownum 序号,姓名,日期 " & _
             "From ( " & _
                    "Select min(" & IIf(mbln配药, "配药日期", "Nvl(签到时间,填制日期)") & ") 日期,病人id,姓名 " & _
                    "From (" & strSql & ") " & _
                    "Where 病人ID Not In (Select distinct A.病人ID From 未发药品记录 A,药品收发记录 B,门诊费用记录 C " & _
                                         "Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id and B.费用id=C.id and (A.单据=8 or A.单据=9 or A.单据=10) " & _
                                         "  and A.排队状态=4 and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)) " & _
                    "Group By 姓名,病人id " & _
                    "Order by 日期 " & _
                    ")"
    'Call SaveErrLog(strSQL)
    Set mrsData(Index) = gobjDatabase.OpenSQLRecord(strSql, "", mlng药房ID, strWin, date开始日期, date结束日期)
    'Call SaveErrLog(mlng药房ID & "," & strWin & "," & date开始日期 & "," & date结束日期)
    'If mrsData(Index).State = 1 Then mrsData(Index).Close
    'mrsData(Index).Open strSQL, gcnOracle
'    If Not mrsData(Index).EOF Then
'        If Nvl(mrsData(Index)!配药日期) <> "" Then
'            mrsData(Index).Sort = "配药日期"
'        Else
'            mrsData(Index).Sort = "签到时间"
'        End If
'    End If
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadData*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
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
    date开始日期 = gobjDatabase.Currentdate
    'date开始日期 = Now - 1
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")
    
    date结束日期 = gobjDatabase.Currentdate
    'date结束日期 = Now
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "select 姓名 from 未发药品记录 where 排队状态=3 and 库房id=[1] and 发药窗口=[2] and 填制日期 between [3] and [4]"
    Set mrsCallingData = gobjDatabase.OpenSQLRecord(strSql, "", mlng药房ID, strWin, date开始日期, date结束日期)
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadCalling*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub


Private Sub loadPreparing(ByVal strWin As String, ByVal intIndex As Integer)
'************************************************************************
'
'加载待配药列表的数据
'
'************************************************************************
    Dim strSql As String
    
    Dim date开始日期 As Date
    Dim date结束日期 As Date
        
    On Error GoTo errHandle
    date开始日期 = gobjDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")

    date结束日期 = gobjDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "Select rownum 序号,A.姓名,B.填制日期,B.签到时间 From 未发药品记录 A,药品收发记录 B,门诊费用记录 C" & _
             " Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id and B.费用id=C.id and (A.单据=8 or A.单据=9 or A.单据=10) "
    If mbln配药确认 Then
        strSql = strSql & "and A.排队状态=1 and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    Else
        strSql = strSql & "and (A.排队状态=1 or A.排队状态=0 or A.排队状态 is null) and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    End If
    'strSQL = strSQL & "and (A.排队状态=1 or A.排队状态=0 or A.排队状态 is null) and A.库房id=52 and A.发药窗口='窗口1' and A.填制日期>sysdate-2 And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
    If mbln配药 = False Then strSql = strSql & " And 1=2"
    strSql = "Select 姓名,min(序号) 序号 From (" & strSql & " Order by Nvl(B.签到时间,A.填制日期)) group by 姓名 order by 序号"
    Set mrsPreparingData(intIndex) = gobjDatabase.OpenSQLRecord(strSql, "", mlng药房ID, strWin, date开始日期, date结束日期)
    
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadPreparing*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("strSQL:" & strSql)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Sub loadTimeout(ByVal strWin As String, ByVal intIndex As Integer)
'************************************************************************
'
'加载已过号（已配药超过多少秒）列表的数据
'
'************************************************************************
    Dim strSql As String
    
    Dim date开始日期 As Date
    Dim date结束日期 As Date
        
    On Error GoTo errHandle
    date开始日期 = gobjDatabase.Currentdate
    date开始日期 = CDate(Format(date开始日期, "yyyy-mm-dd") & " 00:00:00")

    date结束日期 = gobjDatabase.Currentdate
    date结束日期 = CDate(Format(date结束日期, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "Select distinct A.病人ID,A.姓名,A.呼叫时间 From 未发药品记录 A,药品收发记录 B,门诊费用记录 C" & _
             " Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id and B.费用id=C.id and (A.单据=8 or A.单据=9 or A.单据=10) " & _
             "   and A.排队状态=4 and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0) " & _
             " union all " & _
             " Select distinct A.病人ID,A.姓名,A.呼叫时间 From 未发药品记录 A,药品收发记录 B,住院费用记录 C" & _
             " Where A.单据=B.单据 And A.No=B.NO And A.库房id=B.库房id and B.费用id=C.id and (A.单据=8 or A.单据=9 or A.单据=10) " & _
             "   and A.排队状态=4 and A.库房id=[1] and A.发药窗口=[2] and A.填制日期 between [3] and [4] And (B.记录状态=1 Or Mod(B.记录状态,3)=0)"
'    If mbln配药 = False Then
'        strSQL = strSQL & " and round((sysdate-Nvl(B.签到时间,B.填制日期))*24*60)>" & mType_para.intTimeout
'        strSQL = strSQL & " Order by Nvl(B.签到时间,B.填制日期) DESC"
'    Else
'        strSQL = strSQL & " and round((sysdate-B.配药日期)*24*60)>" & mType_para.intTimeout
'        strSQL = strSQL & " Order by B.配药日期 DESC"
'    End If
'    strSQL = strSQL & " And A.呼叫时间 IS NOT NULL And round((SYSDATE-Nvl(A.呼叫时间,SYSDATE))*24*60*60)>" & mType_para.intTimeout
    strSql = "Select rownum 序号,姓名 " & _
             "From (Select 病人ID,姓名,min(呼叫时间) 呼叫时间 " & _
                   "From (" & strSql & ") " & _
                   "Group by 病人ID,姓名 " & _
                   "Order by 呼叫时间 asc " & _
                   ")"
    'Call SaveErrLog(strSQL)
    Set mrsTimeoutData(intIndex) = gobjDatabase.OpenSQLRecord(strSql, "", mlng药房ID, strWin, date开始日期, date结束日期)
    'Call SaveErrLog(mlng药房ID & "," & strWin & "," & date开始日期 & "," & date结束日期)
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadTimeout*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Sub InitVSF()
'************************************************************************
'
'初始化表格
'
'************************************************************************
    Dim intColWidth As Integer
    Dim lngRowheight As Long, lngRowheights As Long
    Dim i As Integer
    Dim strReg As String
    Dim dblHeight As Double
    On Error GoTo errHandle
    strReg = "公共模块\药房排队叫号\液晶电视"
    
    mintRows = 3

    
    mIntCallRow = IIf(mType_para.bln显示待发药, mType_para.int待发药行数, 0)
    mIntPraRow = IIf(mType_para.bln显示待配药, mType_para.int待配药行数, 0)
    
    mIntCallCol = IIf(mType_para.bln显示待发药, mType_para.int待发药列数, 0)
    mIntPraCol = IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0)
    mIntTimeoutCol = IIf(mType_para.bln显示已过号, mType_para.int已过号列数, 0)
    mRowRec = mIntCallCol + mIntPraCol + mIntTimeoutCol
    'mRowRec = mType_para.int待发药列数
    'mintRows = mintRows + mIntCallRow + mIntPraRow + IIf(mType_para.bln显示待发药, 1, 0) + IIf(mType_para.bln显示待配药, 1, 0)
    mintRows = mintRows + mIntCallRow
    
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
        
        If mType_para.bln显示待发药 Or mType_para.bln显示待配药 Or mType_para.bln显示已过号 Then
            .RowHeight(0) = 25 * Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))
            .RowHeight(1) = 30 * Val(GetSetting("ZLSOFT", strReg, "字号(0)", "14"))
            .RowHeight(2) = 25 * Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))
        Else
            .RowHeight(0) = 25 * Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))
            .RowHeight(1) = IIf(mType_para.bln显示窗口, .Height - .RowHeight(0), .Height)
            .RowHeight(2) = 0
        End If
        
        If Not mType_para.bln显示窗口 Then
            .RowHeight(0) = 0
        End If
        If vfgCallingData.Rows > 3 Then
            'mIntCallRow待发药行数,总的行数应该是mIntCallRow待发药行数+3
            lngRowheight = Round((.Height - .RowHeight(0) - .RowHeight(1) - .RowHeight(2)) / mIntCallRow)
            lngRowheights = 0
            For i = 3 To .Rows - 1
                .RowHeight(i) = lngRowheight
                lngRowheights = lngRowheights + lngRowheight
            Next
            vfgCallingData.Height = lngRowheights + .RowHeight(0) + .RowHeight(1) + .RowHeight(2) + 50
        End If
    End With
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************InitVSF*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub Form_Resize()
    Dim strReg As String
    On Error GoTo errHandle
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    'Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln显示其他内容, Round(Me.ScaleHeight * 0.9), Round(Me.ScaleHeight))
    'Me.PicMsg.Move 0, Me.vfgCallingData.Height, Me.vfgCallingData.Width, Round(Me.ScaleHeight * 0.1)
    Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln显示其他内容, Round(Me.ScaleHeight - 30 * Val(GetSetting("ZLSOFT", strReg, "字号(4)", "14"))), Round(Me.ScaleHeight))
    Me.lblmsg.Move 0, Me.vfgCallingData.Height + 100, Me.vfgCallingData.Width, Me.ScaleHeight
    'Me.lblmsg.Move 0, Me.PicMsg.Height / 20, Me.PicMsg.Width, Me.PicMsg.Height
    InitVSF
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************Form_Resize*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    On Error GoTo errHandle
    For i = 0 To mintCols - 1
'        mstrSendNames(i) = ""
'        mstrPraNames(i) = ""
        Set mrsData(i) = Nothing
        Set mrsPreparingData(i) = Nothing
        Set mrsTimeoutData(i) = Nothing
    Next
    Set mrsCallingData = Nothing
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************Form_Unload*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
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
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "公共模块\药房排队叫号\液晶电视"
    If mType_para.bln显示待发药 = False And mType_para.bln显示待配药 = False And mType_para.bln显示已过号 = False Then Exit Sub

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

'        If mType_para.bln显示待发药 = True Then
'            '待发药的总页数
'            'intCallPage = (mrsData(k).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(k).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))
'            intCallPage = 1
'            If intCallPage = 0 Then intCallPage = 1
'
'            '当前页
'            'intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), "/"))) + 1
'            intPage = 1
'            If intPage > intCallPage Then intPage = 1
'
'            Me.vfgCallingData.Cell(flexcpText, mIntCallRow + 2, intcol, mIntCallRow + 2, (k + 1) * mRowRec - 1) = "待发药   " & strTemp & intPage & "/" & intCallPage & " 共" & mrsData(k).RecordCount & "人"
'        End If

    
        If mType_para.bln显示待发药 = True Then
            '当数据集有数据，但是已经移到最后一条数据了就移到第一条
            If mrsData(k).EOF And mrsData(k).RecordCount > 0 Then mrsData(k).MoveFirst
            '清空待发药列表的数据
            If mType_para.bln显示待配药 = True Then
                Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec + mType_para.int待配药列数, mIntCallRow + 2, k * mRowRec + mType_para.int待配药列数 + mType_para.int待发药列数 - 1) = ""
                intcol = k * mRowRec + mType_para.int待配药列数
            Else
                Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec, mIntCallRow + 2, k * mRowRec + mType_para.int待发药列数 - 1) = ""
                intcol = k * mRowRec
            End If

            i = 3
            count = 0
            Do While Not mrsData(k).EOF
                If mrsData(k)!姓名 <> "" Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsData(k)!姓名)
                    
                    Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "字号(3)", "14"))
                    Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "待发药颜色", vbGreen)
                    Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "字体(3)", "宋体")
                    Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "粗体(3)", "false")
                    Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "斜体(3)", "false")
                    
    '                '每行显示了制定人数后，跳到下一行
                    If intcol < IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0) + mType_para.int待发药列数 + k * mRowRec - 1 Then
                        intcol = intcol + 1
                    Else
                        intcol = 0
                        i = i + 1
                    End If
                End If

                mrsData(k).MoveNext
                
                 '当数据的显示已经到达设置值时，退出循环
                If count >= mType_para.int待发药人数 Then
                    Exit Do
                End If
            Loop
        End If
        
'        intcol = 0
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
'            Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (k + 1) * mRowRec - 1) = "待配药   " & strTemp & intPage & "/" & intPraPage & " 共" & mrsPreparingData(k).RecordCount & "人"
'        End If

        If mType_para.bln显示待配药 = True Then
            count = 0
            '当数据集有数据，但是已经移到最后一条数据了就移到第一条
            If mrsPreparingData(k).EOF And mrsPreparingData(k).RecordCount > 0 Then mrsPreparingData(k).MoveFirst
            '清空待发药列表的数据
            Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec, mIntCallRow + 2, k * mRowRec + mType_para.int待配药列数 - 1) = ""

            i = 3
            intcol = k * mRowRec
            Do While Not mrsPreparingData(k).EOF
                If mrsPreparingData(k)!姓名 <> "" Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = mrsPreparingData(k)!姓名
                    Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "字号(2)", "14"))
                    Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "待配药颜色", vbGreen)
                    Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "字体(2)", "宋体")
                    Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "粗体(2)", "false")
                    Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "斜体(2)", "false")

                    '一行数据加载完之后，跳到下一行
                    If intcol < mType_para.int待配药列数 + k * mRowRec - 1 Then
                        intcol = intcol + 1
                    Else
                        intcol = 0
                        i = i + 1
                    End If

'                    If count = mRowRec Then
'                        count = 0
'                        intcol = k * mRowRec
'                        If Not mrsPreparingData(k).EOF Then i = i + 1
'                    End If
                End If
                
                mrsPreparingData(k).MoveNext
                '当数据的显示已经到达设置值时，退出循环
                If count >= mType_para.int待配药人数 Then
                    Exit Do
                End If
            Loop
        End If

        If mType_para.bln显示已过号 = True Then
            '当数据集有数据，但是已经移到最后一条数据了就移到第一条
            If mrsData(k).EOF And mrsData(k).RecordCount > 0 Then mrsData(k).MoveFirst
            '清空待发药列表的数据
            Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec + IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0) + IIf(mType_para.bln显示待发药, mType_para.int待发药列数, 0), mIntCallRow + 2, k * mRowRec + IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0) + IIf(mType_para.bln显示待发药, mType_para.int待发药列数, 0) + mType_para.int已过号列数 - 1) = ""
            intcol = k * mRowRec + IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0) + IIf(mType_para.bln显示待发药, mType_para.int待发药列数, 0)

            i = 3
            count = 0
            Do While Not mrsData(k).EOF
                If mrsData(k)!姓名 <> "" Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsData(k)!姓名)
                    
                    Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "字号(5)", "14"))
                    Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "已过号颜色", vbGreen)
                    Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "字体(5)", "宋体")
                    Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "粗体(5)", "false")
                    Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "斜体(5)", "false")
                    
    '                '每行显示了制定人数后，跳到下一行
                    If intcol < IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0) + IIf(mType_para.bln显示待发药, mType_para.int待发药列数, 0) + mType_para.int已过号列数 + k * mRowRec - 1 Then
                        intcol = intcol + 1
                    Else
                        intcol = 0
                        i = i + 1
                    End If
                End If

                mrsData(k).MoveNext
                
                 '当数据的显示已经到达设置值时，退出循环
                If count >= mType_para.int已过号人数 Then
                    Exit Do
                End If
            Loop
        End If
    Next
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************TimerPage_Timer*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Public Sub ShowMe(ByVal lng药房ID As Long, ByVal strWins As String, ByVal bln配药 As Boolean, ByVal bln配药确认 As Boolean)
'**************************************************************************
'打开窗体的接口，lng药房ID：当前的药房id；strWins：窗体连接串
'**************************************************************************
    Dim rsWin As ADODB.Recordset
    Dim strWins_temp As String, strSql As String
    Dim strTemp As String
    Dim strReg As String
    Dim cls As New clsLCDShow
    
    On Error GoTo errHandle
    mlng药房ID = lng药房ID
    strWins_temp = "'" & Replace(strWins, ",", "','") & "'"
    mstrWins = strWins
    '对窗口号进行重新排序
    strSql = "Select TO_CHAR(WMSYS.WM_CONCAT(名称)) 名称 From (select 名称 from 发药窗口 where 药房ID=[1] and 名称 in (" & strWins_temp & ") order by 编码)"
    Set rsWin = gobjDatabase.OpenSQLRecord(strSql, "", mlng药房ID)
    If rsWin.RecordCount > 0 Then
        mstrWins = Nvl(rsWin!名称)
    End If
    mbln配药 = bln配药
    mbln配药确认 = bln配药确认
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    strTemp = GetSetting("ZLSOFT", strReg, "窗口", "1,2,3")
    If strTemp = "" And strWins = "" Then
        Call SaveErrLog("************************************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("strTemp:" & strTemp)
        Call SaveErrLog("strWins:" & strWins)
        Call SaveErrLog("发药窗口为空,不显示排队屏幕")
        Call SaveErrLog("************************************")
        cls.zlClose
        Exit Sub
    End If
    Me.Show
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************ShoeMe*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub timerLCD_Timer()
'************************************************************************
'
'刷新待发药列表的数据
'
'************************************************************************
    InitData 2, False
'    Dim i As Long, j As Long, lngRowheight As Long, lngRowHeights As Long
'    j = 0
'    For i = 0 To vfgCallingData.Rows - 1
'        j = j + vfgCallingData.RowHeight(i)
'    Next
'    With vfgCallingData
'    lngRowheight = Round((.Height - .RowHeight(0) - .RowHeight(1) - .RowHeight(2)) / mIntCallRow)
'    lngRowHeights = 0
'    For i = 3 To .Rows - 1
'        .Cell(flexcpFontSize, i, 0, i, 14) = 7.5
'        .Cell(flexcpForeColor, i, 0, i, 14) = vbGreen
'        .Cell(flexcpFontName, i, 0, i, 14) = "宋体"
'        .Cell(flexcpFontBold, i, 0, i, 14) = "false"
'        .Cell(flexcpFontItalic, i, 0, i, 14) = "false"
'        .TextMatrix(i, 0) = .RowHeight(0)
'        .TextMatrix(i, 1) = .RowHeight(1)
'        .TextMatrix(i, 2) = .RowHeight(i)
'        .TextMatrix(i, 3) = .Height
'        .TextMatrix(i, 4) = lblmsg.Height
'        .TextMatrix(i, 5) = Me.ScaleHeight
'        .TextMatrix(i, 6) = Me.Height
'        .TextMatrix(i, 7) = lblmsg.Visible
'        lngRowHeights = lngRowHeights + .RowHeight(i)
'    Next
'        .Height = lngRowHeights + .RowHeight(0) + .RowHeight(1) + .RowHeight(2) + 100
'    End With
    Me.lblmsg.Caption = IIf(mType_para.str显示内容 = "", "祝您早日康复！", mType_para.str显示内容) & "   " & Format(gobjDatabase.Currentdate, "yyyy-mm-dd  hh:mm")
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
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "公共模块\药房排队叫号\液晶电视"
    '显示待发药信息
    If mType_para.bln显示待发药 Then
        '计算总的页数
        'intCallPage = (mrsData(Index).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(Index).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))
        intCallPage = 1
        If intCallPage = 0 Then intCallPage = 1
        
        '判断是否为窗体加载
        If intPage <> 1 Then
            For i = 0 To Me.vfgCallingData.Cols - 1
                '计算当前页数
                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
                    'intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, i), "/")))
                    intCurPage = 1
                End If
            Next
            '将记录集的游标移向当前页显示的内容
'            For i = 1 To mIntCallRow * mRowRec * (intCurPage - 1)
'                If Not mrsData(Index).EOF Then
'                    mrsData(Index).MoveNext
'                End If
'            Next
        Else
            intCurPage = 1
        End If
        
        count = 0
        i = 3
        intcol = Index * mRowRec + IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0)
        
        '循环记录集，显示界面
        Do While Not mrsData(Index).EOF
            If mrsData(Index)!姓名 <> "" Then
                count = count + 1
                '界面显示了指定人数后，退出循环
                If count > mType_para.int待发药人数 Then
                    Exit Do
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "字号(2)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "待发药颜色", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "字体(2)", "宋体")
                Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "粗体(2)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "斜体(2)", "false")
                Me.vfgCallingData.TextMatrix(i, intcol) = TranRowNum(Val(Nvl(mrsData(Index)!序号))) & Nvl(mrsData(Index)!姓名)
                
'                '每行显示了指定人数后，跳到下一行
                If intcol < IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0) + mType_para.int待发药列数 + Index * mRowRec - 1 Then
                    intcol = intcol + 1
                Else
                    intcol = Index * mRowRec + IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0)
                    i = i + 1
                End If
'                If count = mRowRec Then
'                    count = 0
'                    intcol = Index * mRowRec
'                    If Not mrsData(Index).EOF Then i = i + 1
'                End If
            End If
            
            'mstrSendNames(Index) = mstrSendNames(Index) & Nvl(mrsData(Index)!姓名) & ","
            '移向下一条记录
            mrsData(Index).MoveNext
'            intcol = intcol + 1
'            If count = mRowRec Then
'                count = 0
'                intcol = Index * mRowRec
'                If Not mrsData(Index).EOF Then i = i + 1
'            End If
            
        Loop
        
    End If
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************ShowSend*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub ShowPra(ByVal Index As Integer, ByVal intPage As Integer)
    Dim count As Integer
    Dim i As Integer
    Dim intPraPage As Integer
    Dim intCurPage As Integer
    Dim intcol As Integer
    Dim strTemp As String
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "公共模块\药房排队叫号\液晶电视"
    If mType_para.bln显示待配药 Then
        '加载待配药信息
        loadPreparing (Split(mstrWins, ",")(Index)), Index
        
        '计算总页数
        'intPraPage = (mrsPreparingData(Index).RecordCount \ (mIntPraRow * mRowRec) + IIf(mrsPreparingData(Index).RecordCount Mod (mIntPraRow * mRowRec) = 0, 0, 1))
        intPraPage = 1
        If intPraPage = 0 Then intPraPage = 1
    
        '判断是否是窗体加载
        If intPage <> 1 Then
            For i = 0 To Me.vfgCallingData.Cols - 1
                '得到当前页数
                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
                    'intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), "/")))
                    intCurPage = 1
                    Exit For
                End If
            Next

            '将记录集移到当前页数的位置
'            For i = 1 To mIntPraRow * mRowRec * (intCurPage - 1)
'                If Not mrsPreparingData(Index).EOF Then
'                    mrsPreparingData(Index).MoveNext
'                End If
'            Next

        Else
            intCurPage = 1
        End If
        
        count = 0
        i = 3
        intcol = mRowRec * Index
        '循环记录集，将数据显示到界面
        Do While Not mrsPreparingData(Index).EOF
            If Nvl(mrsPreparingData(Index)!姓名) <> "" Then
                count = count + 1
                '当界面的人数显示预定个数后，退出循环
                If count > mType_para.int待配药人数 Then
                    Exit Do
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "字号(3)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "待配药颜色", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "字体(3)", "宋体")
                Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "粗体(3)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "斜体(3)", "false")
                Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsPreparingData(Index)!姓名)
                '一行数据加载完之后，跳到下一行
                If intcol < mType_para.int待配药列数 + Index * mRowRec - 1 Then
                    intcol = intcol + 1
                Else
                    intcol = mRowRec * Index
                    i = i + 1
                End If
'                If count = mRowRec Then
'                    count = 0
'                    intcol = Index * mRowRec
'                End If
                'mstrPraNames(Index) = mstrPraNames(Index) & Nvl(mrsPreparingData(Index)!姓名) & ","
            End If
            mrsPreparingData(Index).MoveNext
            
            '一行数据加载完之后，跳到下一行
'            intcol = intcol + 1
'            If count = mRowRec Then
'                count = 0
'                intcol = Index * mRowRec
'                i = i + 1
'            End If
        Loop
        
        '显示翻页信息
'        intcol = Index * mRowRec
'        For intcol = Index * mRowRec To (Index + 1) * mRowRec - 1
'            If (intcol \ mRowRec) Mod 2 = 0 Then
'                strTemp = String(0, " ")
'            Else
'                strTemp = String(1, " ")
'            End If
'
'            'Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (Index + 1) * mRowRec - 1) = "待配药   " & strTemp & intCurPage & "/" & intPraPage & " 共" & mrsPreparingData(Index).RecordCount & "人"
'        Next
'        '合并显示翻页信息
'        vfgCallingData.MergeRow(Me.vfgCallingData.Rows - 1) = True
    End If
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************ShowPra*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Sub ShowTimeout(ByVal Index As Integer, ByVal intPage As Integer)
    Dim count As Integer
    Dim i As Integer
    Dim intPraPage As Integer
    Dim intCurPage As Integer
    Dim intcol As Integer
    Dim strTemp As String
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "公共模块\药房排队叫号\液晶电视"
    If mType_para.bln显示已过号 Then
        '加载已过号信息
        loadTimeout (Split(mstrWins, ",")(Index)), Index
        intPraPage = 1
        If intPraPage = 0 Then intPraPage = 1
    
        '判断是否是窗体加载
        If intPage <> 1 Then
            intCurPage = 1
'            For i = 0 To Me.vfgCallingData.Cols - 1
'                '得到当前页数
'                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
'                    intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), "/")))
'                    intCurPage = 1
'                    Exit For
'                End If
'            Next
            '将记录集移到当前页数的位置
'            For i = 1 To mType_para.int已过号人数 * (intCurPage - 1)
'                If Not mrsPreparingData(Index).EOF Then
'                    mrsPreparingData(Index).MoveNext
'                End If
'            Next
        Else
            intCurPage = 1
        End If
        count = 0
        i = 3
        intcol = mRowRec * Index + IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0) + IIf(mType_para.bln显示待发药, mType_para.int待发药列数, 0)
        intPraPage = -1 * Int(-1 * mrsTimeoutData(Index).RecordCount / mType_para.int已过号人数)
        If mType_para.lng当前过号页数 < intPraPage Then
            mType_para.lng当前过号页数 = mType_para.lng当前过号页数 + 1
        Else
            mType_para.lng当前过号页数 = 1
        End If
        '循环记录集，将数据显示到界面
        mrsTimeoutData(Index).MoveFirst
        Do While Not mrsTimeoutData(Index).EOF
            If Nvl(mrsTimeoutData(Index)!姓名) <> "" Then
                count = count + 1
                '当界面的人数显示预定个数后，退出循环
                If count > mType_para.int已过号人数 * mType_para.lng当前过号页数 Then
                    Exit Do
                End If
                If count <= mType_para.int已过号人数 * mType_para.lng当前过号页数 And count > mType_para.int已过号人数 * (mType_para.lng当前过号页数 - 1) Then
                    Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "字号(5)", "14"))
                    Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "已过号颜色", vbGreen)
                    Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "字体(5)", "宋体")
                    Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "粗体(5)", "false")
                    Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "斜体(5)", "false")
                    Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsTimeoutData(Index)!姓名)
                    '一行数据加载完之后，跳到下一行
                    If intcol < IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0) + IIf(mType_para.bln显示待发药, mType_para.int待发药列数, 0) + mType_para.int已过号列数 + Index * mRowRec - 1 Then
                        intcol = intcol + 1
                    Else
                        intcol = mRowRec * Index + IIf(mType_para.bln显示待配药, mType_para.int待配药列数, 0) + IIf(mType_para.bln显示待发药, mType_para.int待发药列数, 0)
                        i = i + 1
                    End If
                End If
            End If
            mrsTimeoutData(Index).MoveNext
        Loop
    End If
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then   '3021是BOF 或 EOF 有一个为真
        Call SaveErrLog("***************ShowTimeout*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Public Sub SetFont()
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "公共模块\药房排队叫号\液晶电视"
    With Me.vfgCallingData
        If .Cols = 0 Then Exit Sub
        '设置字体和字体颜色大小
        '叫号区域
        .Cell(flexcpFontSize, 1, 0, 1, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(0)", "14"))
        .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
        .Cell(flexcpFontName, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "字体(0)", "宋体")
        .Cell(flexcpFontBold, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "粗体(0)", "false")
        .Cell(flexcpFontItalic, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "斜体(0)", "false")
        '配药、发药、过号表头--与窗口号字体及颜色一样
        .Cell(flexcpFontSize, 2, 0, 0, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))
        .Cell(flexcpForeColor, 2, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "窗口颜色", vbGreen)
        .Cell(flexcpFontName, 2, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "字体(1)", "宋体")
        .Cell(flexcpFontBold, 2, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "粗体(1)", "false")
        .Cell(flexcpFontItalic, 2, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "斜体(1)", "false")
        If mType_para.bln显示窗口 Then
            .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "字号(1)", "14"))
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "窗口颜色", vbGreen)
            .Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "字体(1)", "宋体")
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "粗体(1)", "false")
            .Cell(flexcpFontItalic, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "斜体(1)", "false")
        End If
'        If mType_para.bln显示待配药 = True Then
'            .Cell(flexcpFontSize, 3, 0, mintRows - 1, 1) = Val(GetSetting("ZLSOFT", strReg, "字号(3)", "14"))
'            .Cell(flexcpForeColor, 3, 0, mintRows - 1, 1) = GetSetting("ZLSOFT", strReg, "待配药颜色", vbGreen)
'            .Cell(flexcpFontName, 3, 0, mintRows - 1, 1) = GetSetting("ZLSOFT", strReg, "字体(3)", "宋体")
'            .Cell(flexcpFontBold, 3, 0, mintRows - 1, 1) = GetSetting("ZLSOFT", strReg, "粗体(3)", "false")
'            .Cell(flexcpFontItalic, 3, 0, mintRows - 1, 1) = GetSetting("ZLSOFT", strReg, "斜体(3)", "false")
'        End If
'        If mType_para.bln显示待发药 = True Then
'            .Cell(flexcpFontSize, 3, 2, mintRows - 1, 3) = Val(GetSetting("ZLSOFT", strReg, "字号(2)", "14"))
'            .Cell(flexcpForeColor, 3, 2, mintRows - 1, 3) = GetSetting("ZLSOFT", strReg, "待发药颜色", vbGreen)
'            .Cell(flexcpFontName, 3, 2, mintRows - 1, 3) = GetSetting("ZLSOFT", strReg, "字体(2)", "宋体")
'            .Cell(flexcpFontBold, 3, 2, mintRows - 1, 3) = GetSetting("ZLSOFT", strReg, "粗体(2)", "false")
'            .Cell(flexcpFontItalic, 3, 2, mintRows - 1, 3) = GetSetting("ZLSOFT", strReg, "斜体(2)", "false")
'        End If
'        If mType_para.bln显示已过号 = True Then
'            .Cell(flexcpFontSize, 3, 4, mintRows - 1, 4) = Val(GetSetting("ZLSOFT", strReg, "字号(5)", "14"))
'            .Cell(flexcpForeColor, 3, 4, mintRows - 1, 4) = GetSetting("ZLSOFT", strReg, "已过号颜色", vbGreen)
'            .Cell(flexcpFontName, 3, 4, mintRows - 1, 4) = GetSetting("ZLSOFT", strReg, "字体(5)", "宋体")
'            .Cell(flexcpFontBold, 3, 4, mintRows - 1, 4) = GetSetting("ZLSOFT", strReg, "粗体(5)", "false")
'            .Cell(flexcpFontItalic, 3, 4, mintRows - 1, 4) = GetSetting("ZLSOFT", strReg, "斜体(5)", "false")
'        End If
    End With
    
    Me.lblmsg.ForeColor = GetSetting("ZLSOFT", strReg, "其他内容颜色", vbBlack)
    Me.lblmsg.FontSize = Val(GetSetting("ZLSOFT", strReg, "字号(4)", "14"))
    Me.lblmsg.FontName = GetSetting("ZLSOFT", strReg, "字体(4)", "宋体")
    Me.lblmsg.FontBold = GetSetting("ZLSOFT", strReg, "粗体(4)", "false")
    Me.lblmsg.FontItalic = GetSetting("ZLSOFT", strReg, "斜体(4)", "false")
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************SetFont*************")
        Call SaveErrLog("服务器时间:" & gobjDatabase.Currentdate)
        Call SaveErrLog("错误信息:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Function TranRowNum(ByVal lngRowNum As Long) As String
    '带圈圈的数字显示方式,最多只支持1-10
    '  1-9 TranRowNum = Chr(Asc(lngRowNum) - 23896)
    '  10  TranRowNum = Chr(Asc(lngRowNum) - 23887)
    TranRowNum = ""
    If mType_para.bln显示发药序号 Then
        If mType_para.int待发药人数 <= 10 Then
            If lngRowNum < 10 Then
                TranRowNum = Chr(Asc(lngRowNum) - 23896)
            Else
                TranRowNum = Chr(Asc(lngRowNum) - 23887)
            End If
        Else
            TranRowNum = lngRowNum & "."
        End If
    End If
End Function

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


