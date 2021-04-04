VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProcCollectCur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "检查变动过程"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9810
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmProcCollectCur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9810
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&Q)"
      Height          =   350
      Left            =   8520
      TabIndex        =   5
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "开始(&S)"
      Height          =   350
      Left            =   7200
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.PictureBox picFunCap 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "frmProcCollectCur.frx":6852
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfModule 
      Height          =   3855
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   9495
      _cx             =   16748
      _cy             =   6800
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483636
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin MSComctlLib.ProgressBar pgsBar 
      Height          =   135
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog cdgPub 
      Left            =   7800
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTargetCmd 
      AutoSize        =   -1  'True
      Caption         =   "更改…"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   3480
      TabIndex        =   11
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label lblTargetPath 
      AutoSize        =   -1  'True
      Caption         =   "C:\AppSoft"
      Height          =   180
      Left            =   2280
      TabIndex        =   10
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      Caption         =   "系统安装目录:"
      Height          =   180
      Index           =   1
      Left            =   990
      TabIndex        =   9
      Top             =   1200
      Width           =   1170
   End
   Begin VB.Label lblFunNote 
      AutoSize        =   -1  'True
      Caption         =   $"frmProcCollectCur.frx":771C
      Height          =   360
      Left            =   990
      TabIndex        =   8
      Top             =   600
      Width           =   4680
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      Caption         =   "正在收集系统:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblFunCap 
      AutoSize        =   -1  'True
      Caption         =   "检查变动过程"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   990
      TabIndex        =   2
      Top             =   150
      Width           =   1980
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   4
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   90
   End
End
Attribute VB_Name = "frmProcCollectCur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowMe() As Boolean
    Me.Show 1
    ShowMe = True
End Function

Private Sub LoadMoudle()
    '加载安装的模块
    Dim strSQL As String, rsSys As New ADODB.Recordset
    Dim i As Long
    Dim strTarget As String, strInitFile As String
    Dim blnStep As Boolean
    
    '首先获取系统编号等信息
    strSQL = "Select 编号 系统编号, 名称 系统名称, 版本号 系统版本号, 所有者 系统所有者, 正常安装 From Zlsystems where Upper(所有者)=[1] Order by Nvl(共享号,0),编号"
    Set rsSys = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "读取安装系统", gstrUserName)
    
    If rsSys.RecordCount = 0 Then Exit Sub
    
    With vsfModule
        .Redraw = flexRDNone
        
        .Rows = 1
        i = .FixedRows
        .Rows = .FixedRows
        .Rows = rsSys.RecordCount + .FixedRows
        Do While Not rsSys.EOF
            .TextMatrix(i, .ColIndex("编号")) = rsSys!系统编号 & ""
            .TextMatrix(i, .ColIndex("系统名称")) = rsSys!系统名称 & ""
            .TextMatrix(i, .ColIndex("当前版本号")) = rsSys!系统版本号 & ""
            rsSys.MoveNext
            i = i + 1
        Loop
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = flexAlignCenterCenter
        
        .AutoResize = True
        .AutoSize 1, .Cols - 1
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    Dim strMsg As String, strErr As String
    Dim i As Long, lngNum As Long
    Dim rsProc As ADODB.Recordset
    Dim strSysName As String, lngSysNum As Long
    Dim strCurVer As String, strInitPath As String
    
    strMsg = "为了确保变动过程检查结果无误，请确保本地安装脚本、升级脚本没有遗漏，且选择的配置文件目录没有错误。" & vbNewLine & _
                    "是否开始检查？"
    
    If MsgBox(strMsg, vbYesNo, "收集确认") = vbNo Then Exit Sub
    
    With vsfModule
        .Enabled = False
        
        '检查选中
        For i = 1 To .Rows - 1
             If .Cell(flexcpChecked, i, 0) = flexChecked Then
                lngNum = lngNum + 1
            End If
        Next
        
        If lngNum = 0 Then
            .Enabled = True
            MsgBox "没有选中系统，无法进行检查。", , "提示"
            Exit Sub
        End If
        
        '开始收集变动过程
        pgsBar.Max = lngNum
        pgsBar.value = 0
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                strSysName = .TextMatrix(i, .ColIndex("系统名称"))
                lngSysNum = .TextMatrix(i, .ColIndex("编号"))
                strCurVer = .TextMatrix(i, .ColIndex("当前版本号"))
                strInitPath = .TextMatrix(i, .ColIndex("配置文件"))
                
                '状态信息
                lblsta(0).Visible = True
                pgsBar.Visible = True
                
                .Select i, 0
                .TopRow = i
                pgsBar.value = pgsBar.value + 1
                lblsta(0).Caption = "正在检查系统：" & strSysName
                Me.Refresh
                
                Set rsProc = Nothing
                Set rsProc = GetCurProc(strSysName, lngSysNum, strCurVer, strInitPath, strErr)
                If strErr <> "" Then
                    MsgBox "检查过程发生错误。" & vbNewLine & strErr, , "错误信息"
                End If
                
                If Not GetChangedProc(rsProc, strErr) Then
                    If strErr <> "" Then
                        MsgBox "收集变动过程发生错误。" & vbNewLine & strErr, , "错误信息"
                        Exit Sub
                    End If
                End If
                
                If Not UpdateProc2DB(rsProc, 1, strErr) Then
                    If strErr <> "" Then
                        MsgBox "保存变动过程发生错误。" & vbNewLine & strErr, , gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        Next
    
    End With
    
    lblsta(0).Visible = False
    pgsBar.Visible = False
    MsgBox " 检查完成。", , "提示"
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strCol As String
    
    '表格初始化
    With vsfModule
        strCol = ",300,1;编号,2000,1;系统名称,2000,1;当前版本号,2000,1;配置文件,2000,1"
        Call InitTable(vsfModule, strCol)
        .Rows = 1
        .FixedCols = 1
        .ColDataType(0) = flexDTBoolean
        .Cell(flexcpChecked, 0, 0) = flexUnchecked
        .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = &H80000008
    End With

    
    Call LoadMoudle
    LoadInitFile lblTargetPath.Caption
End Sub

Private Sub lblTargetCmd_Click()
    Dim strPath As String
    
    strPath = OpenFolder(Me, "选择系统安装目录")
    If strPath = "" Then Exit Sub
    
    lblTargetPath.Caption = strPath
    lblTargetCmd.Left = lblTargetPath.Left + lblTargetPath.Width + 150
    
    LoadInitFile strPath
End Sub

Private Sub vsfModule_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    With vsfModule
        If NewCol = .ColIndex("配置文件") Then
            .ComboList() = "..."
            .FocusRect = flexFocusSolid
        Else
            .ComboList = ""
        End If
    End With

End Sub

Private Sub vsfModule_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strFile As String
    Dim strMainPath As String

    If Col = vsfModule.ColIndex("配置文件") Then
        With cdgPub
            .DialogTitle = "选择应用安装配置文件"
            If Trim(vsfModule.TextMatrix(Row, vsfModule.ColIndex("编号"))) = "" Then
                .Filter = "服务器工具脚本(zlServer.Sql)|zlServer.Sql"
            Else
                .Filter = "应用安装配置文件(zlSetup.ini)|zlSetup.ini"
                .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
            End If
            
            .ShowOpen
            If .Filename <> "" Then
                If CheckInitFile(vsfModule.TextMatrix(Row, vsfModule.ColIndex("编号")), .Filename) Then
                    vsfModule.TextMatrix(Row, Col) = .Filename
                End If
            End If
        End With
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub LoadInitFile(ByVal strPath As String)
    '根据路径加载配置文件
    Dim i As Long, strInitFile As String
    
    With vsfModule
        For i = 1 To .Rows - 1
            strInitFile = strPath & "\" & Decode(.TextMatrix(i, .ColIndex("编号")) \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                                6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                                23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                                25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\应用脚本\ZLSETUP.INI"
            If gobjFSO.FileExists(strInitFile) Then
                .TextMatrix(i, .ColIndex("配置文件")) = strInitFile
                .Cell(flexcpChecked, i, 0) = flexChecked    '有配置文件,就选中
            Else
                .TextMatrix(i, .ColIndex("配置文件")) = ""
                .Cell(flexcpChecked, i, 0) = flexUnchecked
            End If
        Next
    End With
End Sub

Private Sub vsfModule_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Or Col = vsfModule.ColIndex("配置文件") Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub vsfModule_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsfModule
        If .Redraw = flexRDNone Then Exit Sub
        If .Rows = 1 Then Exit Sub
            
        If Row = 0 Then
            If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("配置文件")) <> "" Then
                        .Cell(flexcpChecked, i, 0) = flexChecked
                    Else
                        .Cell(flexcpChecked, i, 0) = flexUnchecked
                    End If
                Next
            Else
                .Cell(flexcpChecked, 1, 0, .Rows - 1, 0) = flexUnchecked
            End If
        End If

    End With
End Sub


Private Function GetChangedProc(ByRef rsProcs As ADODB.Recordset, ByRef strErr As String) As Boolean
    '遍历传入的过程集合,判断是否是变动过程
    '有错误返回False
    Dim strProc As String
    
    On Error GoTo errH
    If rsProcs Is Nothing Then Exit Function
    If rsProcs.RecordCount = 0 Then Exit Function
    
    With rsProcs
        .Filter = 0
        Do While Not .EOF
            strProc = LoadBaseProcs(!P_Name)
            
            If ConvertStr(strProc) = "" Then
                .Delete adAffectCurrent
            ElseIf ConvertStr(!P_Define) = "" Then
                .Delete adAffectCurrent
            ElseIf ConvertStr(strProc) = ConvertStr(!P_Define) Then
                .Delete adAffectCurrent
            End If
            .MoveNext
        Loop
    End With
    
    GetChangedProc = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, , gstrSysName
End Function

Private Function GetCurProc(ByVal strSysName, ByVal lngSysNum As Long, ByVal strCurVer As String, strInitPath As String, Optional strErr As String) As ADODB.Recordset
    '传入系统名称\系统编号\系统版本\配置文件目录,返回当前版本的标准过程
    Dim rsProcedure As ADODB.Recordset
    Dim rsFiles As ADODB.Recordset
    Dim i As Long, strFile As String
    Dim strFileVer As String, strOwner As String
    
    On Error GoTo errH
    
    Set rsFiles = GetUpgradeFiles(rsFiles, lngSysNum, GetPrimaryVer(strCurVer), strInitPath, , , strCurVer, , , True, False)
    If rsFiles Is Nothing Then Exit Function

    '先收集安装文件 zlProgram.sql
    strFileVer = "ZLPROGRAM.SQL"
    strOwner = GetOwnerName(lngSysNum, gcnOracle)
    strFile = Left(strInitPath, InStrRev(UCase(strInitPath), "ZLSETUP.INI") - 1) & "ZLPROGRAM.SQL"
    GetProceduresByFile strFile, rsProcedure, strFileVer, lngSysNum, strSysName, strOwner
    
    With rsFiles
        If .RecordCount > 0 Then
            .MoveLast '因为是倒序的,所以从集合的最后一行开始循环
        End If
        
        Do While Not .BOF
            strFile = !FilePath
            strFileVer = Mid(!FilePath, InStrRev(!FilePath, "\") + 1)   '文件版本名称
            GetProceduresByFile strFile, rsProcedure, strFileVer, lngSysNum, strSysName, strOwner
            .MovePrevious
         Loop
    End With
    
    Set GetCurProc = rsProcedure
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbCritical, gstrSysName
End Function

