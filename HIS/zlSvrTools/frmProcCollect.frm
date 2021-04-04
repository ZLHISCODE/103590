VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProcCollect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "变动过程检查"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11145
   Icon            =   "frmProcCollect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11145
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pctBottom 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   11055
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   11055
      Begin MSComctlLib.ProgressBar pgsBar 
         Height          =   135
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         Caption         =   "正在检查本地文件"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   5400
         TabIndex        =   7
         Top             =   300
         Width           =   90
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         Caption         =   "所检查文件:"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   4320
         TabIndex        =   6
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   3240
         TabIndex        =   5
         Top             =   300
         Width           =   90
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         Caption         =   "当前进度:"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   2400
         TabIndex        =   4
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.PictureBox picFunCap 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "frmProcCollect.frx":6852
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFiles 
      Height          =   4935
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   10815
      _cx             =   19076
      _cy             =   8705
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
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
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   150
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
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      Caption         =   "将对以下系统进行变动过程检查,正在检查系统:"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   5
      Left            =   990
      TabIndex        =   10
      Top             =   600
      Width           =   3780
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   4
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Width           =   90
   End
   Begin VB.Label lblFunCap 
      AutoSize        =   -1  'True
      Caption         =   "变动过程检查"
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
      TabIndex        =   1
      Top             =   150
      Width           =   1980
   End
End
Attribute VB_Name = "frmProcCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSystem As String
Private mstrCurInitPath As String
Public Event ReturnChangedProc(ByVal rsTmp As ADODB.Recordset, ByVal intType As Integer) '用于向调用窗体返回记录集,intType: 1-变动过程记录集 2-升级后被修改过程记录集


Public Sub ShowMe(ByVal strSystem As String, Optional ByVal strCurInitPath As String)
    '窗体显示函数
    mstrSystem = strSystem
    mstrCurInitPath = strCurInitPath
    Me.Show 1
End Sub

Private Sub Form_Activate()
    GetChangedProc 1
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strCol As String
    
    strCol = ",500,1;系统编号,0,1;系统名称,1800,1;版本,1600,1;脚本类型,0,1;文件,2000,1"
    Call InitTable(vsfFiles, strCol)
    vsfFiles.Rows = 1
    vsfFiles.FixedCols = 1
    vsfFiles.Cell(flexcpForeColor, 0, 0, 0, vsfFiles.Cols - 1) = &H80000008
    
    LoadFiles mstrSystem, mstrCurInitPath
    pgsBar.Max = vsfFiles.Rows - 1
    pctBottom.Visible = True
    
End Sub

Private Sub LoadFiles(ByVal strSystem As String, Optional ByVal strCurInitPath As String)
    '根据传入的字符获取变动过程检查所需文件,加载至表格
    'strSystem - 需要收集系统 , 格式为 "系统号,系统名称,当前版本,目标版本,目录",多个系统之间用分号间隔
    'strCurInitPath - 跨大版本时升级时,当前版本系统的配置文件目录
    Dim arrSys() As String, i As Integer, j As Long
    Dim lngSysNum As Long, strCur As String, strSysName As String
    Dim strTarget As String, strTargetInitPath As String
    Dim rsTmpCur As New ADODB.Recordset, rsTmpTarget As New ADODB.Recordset
    Dim strInitPath As String
    
    arrSys = Split(strSystem, ";")
    strInitPath = strCurInitPath
    With vsfFiles
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .FixedCols = 1
        .ColAlignment(0) = flexAlignCenterCenter
        .MergeCells = flexMergeRestrictRows
        .MergeCol(.ColIndex("系统编号")) = True
        .MergeCol(.ColIndex("系统名称")) = True
        
        For i = 0 To UBound(arrSys)
            lngSysNum = Split(arrSys(i), ",")(0)
            strSysName = Split(arrSys(i), ",")(1)
            strCur = Split(arrSys(i), ",")(2)
            strTarget = Split(arrSys(i), ",")(3)
            strTargetInitPath = Split(arrSys(i), ",")(4)
            
            If strInitPath = "" Then
                strCurInitPath = strTargetInitPath
            Else
                strCurInitPath = strInitPath & "\" & Decode(lngSysNum \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                    6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                    23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                    25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\应用脚本\ZLSETUP.INI"
            End If
            
            '1首先加载当前版本的 安装脚本＆升级脚本
            '2.加载目标版本的升级脚本,并合并
            Set rsTmpCur = GetUpgradeFiles(rsTmpCur, lngSysNum, GetPrimaryVer(strCur), strCurInitPath, , , strCur, , , True, False)
            Set rsTmpTarget = GetUpgradeFiles(rsTmpTarget, lngSysNum, strCur, strTargetInitPath, , , strTarget, , , True)
            RecDataAppend rsTmpTarget, rsTmpCur
            
            If rsTmpTarget.RecordCount <> 0 Then
                '获取后加载信息至表格
                    rsTmpTarget.MoveLast '因为是倒序的,所以从集合的最后一行开始循环
                    
                    '安装脚本
                    .Rows = .Rows + 1
                    j = .Rows - 1
                    .TextMatrix(j, 0) = j
                    .TextMatrix(j, .ColIndex("系统编号")) = lngSysNum
                    .TextMatrix(j, .ColIndex("系统名称")) = strSysName
                    .TextMatrix(j, .ColIndex("版本")) = GetPrimaryVer(strCur)
                    .TextMatrix(j, .ColIndex("文件")) = Left(strCurInitPath, InStrRev(UCase(strCurInitPath), "ZLSETUP.INI") - 1) & "ZLPROGRAM.SQL"
                    .TextMatrix(j, .ColIndex("脚本类型")) = "当前"
                    
                    Do While Not rsTmpTarget.BOF
                        If InStr(1, UCase(rsTmpTarget!FilePath), "OPTIONAL.SQL") = 0 And InStr(1, UCase(rsTmpTarget!FilePath), "HISTORY.SQL") = 0 Then
                            'OPTIONAL,HISTORY脚本不需要进行检查
                            .Rows = .Rows + 1
                            j = .Rows - 1
                            .TextMatrix(j, 0) = j
                            .TextMatrix(j, .ColIndex("系统编号")) = lngSysNum
                            .TextMatrix(j, .ColIndex("系统名称")) = strSysName
                            .TextMatrix(j, .ColIndex("版本")) = rsTmpTarget!SPVer
                            .TextMatrix(j, .ColIndex("文件")) = rsTmpTarget!FilePath
                                
                        
                            If IsUpgradeFile(strCur, lngSysNum, rsTmpTarget!FilePath) Then
                                .TextMatrix(j, .ColIndex("脚本类型")) = "升级"
                            Else
                                .TextMatrix(j, .ColIndex("脚本类型")) = "当前"
                            End If
                        End If
                        rsTmpTarget.MovePrevious
                    Loop
                
            End If
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub GetChangedProc(ByVal lngRow As Long)
    '循环表格获取变动过程
    Dim i As Long, lngSysNum As Long, j As Long
    Dim blnCheck As Boolean, strErr As String
    Dim rsCurrent As ADODB.Recordset
    Dim rsUpdate As ADODB.Recordset
    Dim strProcTxt As String, strSysOwner As String
    
    '说明:表格中的数据是按照 系统、低版本->高版本进行排序的
    '遍历表格,获取同一系统低版本至高本版的脚本
    With vsfFiles
        If lngRow > .Rows - 1 Then Exit Sub
        For i = lngRow To .Rows - 1
            '状态信息
            lblsta(6).Caption = "正在检查本地文件"
            lblsta(4).Caption = .TextMatrix(i, .ColIndex("系统名称"))
            lblsta(3).Caption = .TextMatrix(i, .ColIndex("文件"))
            lblsta(1).Caption = "(" & i & "/" & .Rows - 1 & ")"
            pgsBar.Max = .Rows
            pgsBar.value = i
            .Select i, 0
            .TopRow = i
            Me.Refresh:
            
            '遍历表格,根据表格中的文件获取过程
            If i <= .Rows - 1 Then
                If lngSysNum = 0 Then
                    lngSysNum = .TextMatrix(lngRow, .ColIndex("系统编号"))
                End If
                
                If .TextMatrix(i, .ColIndex("脚本类型")) = "当前" Then
                    '获取同一系统下的当前版本脚本
                    GetProceduresByFile .TextMatrix(i, .ColIndex("文件")), rsCurrent
                Else
                    '升级脚本
                    GetProceduresByFile .TextMatrix(i, .ColIndex("文件")), rsUpdate
                End If
                
                If i = .Rows - 1 Then   '最后一行
                    blnCheck = True
                Else
                    If lngSysNum = .TextMatrix(i + 1, .ColIndex("系统编号")) Then
                        blnCheck = False
                    Else
                        lngSysNum = .TextMatrix(i + 1, .ColIndex("系统编号"))
                        blnCheck = True
                    End If
                End If
            End If
            
            '检查完一个系统的文件或检查到表格最后一行时,就对收集到的过程进行处理
            If blnCheck Or i = .Rows - 1 Then
                If rsCurrent.RecordCount > 0 Then
                    rsCurrent.MoveFirst
                End If
                
                lblsta(6).Caption = "正在检查数据库过程"
                lblsta(4).Caption = .TextMatrix(i, .ColIndex("系统名称"))
                lblsta(2).Caption = "所检查过程："
                
                strSysOwner = GetOwnerName(.TextMatrix(i, .ColIndex("系统编号")), gcnOracle)
                j = 1: pgsBar.Max = rsCurrent.RecordCount
                '1.和数据库比较
                Do While Not rsCurrent.EOF
                    '状态信息
                    lblsta(3).Caption = rsCurrent!P_Name
                    lblsta(1).Caption = "(" & j & "/" & pgsBar.Max & ")"
                    pgsBar.value = j
                    Me.Refresh
                    
                    '从数据库取出过程和记录集中的做对比
                    strProcTxt = LoadBaseProcs(rsCurrent!P_Name)
                    If strProcTxt = "" Then
                        rsCurrent.Delete adAffectCurrent
                    ElseIf ConvertStr(rsCurrent!P_Define) = ConvertStr(strProcTxt) Then '相同就删除
                        rsCurrent.Delete adAffectCurrent
                    Else
                        rsCurrent.Update Array("P_System", "P_Owner", "P_SysNum"), _
                                                 Array(.TextMatrix(i, .ColIndex("系统名称")), strSysOwner, .TextMatrix(i, .ColIndex("系统编号")))
                    End If
                    rsCurrent.MoveNext
                    j = j + 1
                Loop
                
                '比较后把记录集返回,并更新至数据库
                rsCurrent.Filter = 0
                If Not UpdateProc(rsCurrent, 1, strErr) Then
                    MsgBox "保存变动过程时发生错误，请重试。" & vbNewLine & strErr, , "错误"
                    Exit Sub
                End If
                RaiseEvent ReturnChangedProc(rsCurrent, 1)
                
                '2.对比升级脚本
                If Not rsUpdate Is Nothing Then
                    If rsCurrent.RecordCount > 0 And rsUpdate.RecordCount > 0 Then
                        lblsta(6).Caption = "正在检查升级脚本中的过程"
                        lblsta(2).Caption = "所检查过程："
                        j = 1: pgsBar.Max = rsCurrent.RecordCount
                        
                        rsCurrent.MoveFirst
                        Do While Not rsCurrent.EOF
                            '状态信息
                            lblsta(3).Caption = rsCurrent!P_Name
                            lblsta(1).Caption = "(" & j & "/" & pgsBar.Max & ")"
                            pgsBar.value = j
                            Me.Refresh
                        
                            '检查变动过程是否在升级脚本中涉及
                            rsUpdate.Filter = "P_Name = '" & rsCurrent!P_Name & "'"
                            If rsUpdate.RecordCount <> 0 Then
                                rsUpdate.Update Array("P_System", "P_Owner", "P_SysNum"), _
                                                           Array(.TextMatrix(i, .ColIndex("系统名称")), strSysOwner, .TextMatrix(i, .ColIndex("系统编号")))
                            End If
                            rsCurrent.MoveNext
                            j = j + 1
                        Loop
                        
                        '将本次升级不涉及的过程删除(即系统名称\所有者\系统编号为空的记录)
                        rsUpdate.Filter = "P_System ='' "
                        Do While Not rsUpdate.EOF
                            rsUpdate.Delete
                            rsUpdate.MoveNext
                        Loop
                        
                        rsUpdate.Filter = 0
                        RaiseEvent ReturnChangedProc(rsUpdate, 2)
                        If Not UpdateProc(rsUpdate, 2, strErr) Then
                            MsgBox "保存最新变动过程时发生错误，请重试。" & vbNewLine & strErr, , "错误"
                            Exit Sub
                        End If
                    End If
                End If
                
                '当前过程处理完毕,进入下一个过程的处理
                Set rsCurrent = Nothing
                Set rsUpdate = Nothing
                blnCheck = False
                
            End If
        Next
        
    End With
    
End Sub

Private Function IsUpgradeFile(ByVal strCurVer As String, ByVal lngSys As Long, ByVal strFileName As String) As Boolean
    '根据版本号和脚本文件名称判断是否是升级脚本,是升级脚本就返回Ture
    'strCurVer -版本号  strFileName - 脚本文件名称 lngSys-系统版本号
    Dim strFileVer As String
    Dim arrCurVer() As String, arrFileVer() As String

    If InStr(1, strFileName, "\") > 0 Then
        strFileName = Mid(strFileName, InStrRev(strFileName, "\") + 1)
    End If
    AnalysisFileName strFileName, lngSys, strFileVer
    
    
    arrCurVer = Split(strCurVer, ".")
    arrFileVer = Split(strFileVer, ".")
    
    If UBound(arrFileVer) < 2 Then Exit Function
    
    '前三位依次比较
    If Val(arrCurVer(0)) < Val(arrFileVer(0)) Then
        IsUpgradeFile = True
        Exit Function
    End If
    If Val(arrCurVer(1)) < Val(arrFileVer(1)) Then
        IsUpgradeFile = True
        Exit Function
    End If
    If Val(arrCurVer(2)) < Val(arrFileVer(2)) Then
        IsUpgradeFile = True
        Exit Function
    End If
    
    '比较第四位
    If UBound(arrCurVer) > 2 And UBound(arrFileVer) = 2 Then
        Exit Function
    ElseIf UBound(arrCurVer) = 2 And UBound(arrFileVer) > 2 Then
        IsUpgradeFile = True
        Exit Function
    ElseIf UBound(arrCurVer) > 2 And UBound(arrFileVer) > 2 Then
        If Val(arrCurVer(3)) < Val(arrFileVer(3)) Then
            IsUpgradeFile = True
            Exit Function
        End If
    End If

End Function

Private Sub vsfFiles_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfFiles
        If .Redraw = flexRDNone Then Exit Sub
        If .Rows = 1 Then Exit Sub
        
        .Cell(flexcpForeColor, OldRow, 0) = Color.深灰色
        .Cell(flexcpFontBold, OldRow, 0) = False
        .Cell(flexcpFontBold, NewRow, 0) = True
        .Cell(flexcpForeColor, NewRow, 0) = Color.兰色
    End With
    
End Sub

Private Function UpdateProc(ByVal rsProc As ADODB.Recordset, ByVal intType As Integer, Optional ByRef strErr As String) As Boolean
    '功能:将收集到的过程保存至数据库
    '参数:rsProc-过程集合 intType-过程类型(1-变动过程 2-升级后被修改的过程)
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim arrTxt() As String, i As Long, j As Long
    Dim lngID As Long
    Dim lngSysNum As Long, strIDs As String, arrIds As Variant
    
    On Error GoTo errH
    If rsProc Is Nothing Then
        UpdateProc = True
        Exit Function
    End If
    If rsProc.RecordCount = 0 Then
        UpdateProc = True
        Exit Function
    End If
    
    With rsProc
        .MoveFirst
        Do While Not rsProc.EOF
            '状态信息
            j = j + 1
            lblsta(6).Caption = IIf(intType = 1, "正在保存变动过程源定义", "正在保存变动过程最新定义")
            lblsta(1).Caption = "(" & j & "/" & .RecordCount & ")"
            lblsta(2).Caption = "正在保存过程："
            lblsta(3).Caption = !P_Name
            pgsBar.Max = .RecordCount
            pgsBar.value = j
            Me.Refresh
            
            lngID = GetProcIdByName(!P_Name)
            
            gcnOracle.BeginTrans
            '更新数据至zlProcedure
            If lngID = 0 Then
                If intType = 1 Then
                    strSQL = "Insert Into Zlprocedure (ID, 类型, 名称, 状态, 所有者, 系统编号, 升级前版本) Values" & vbNewLine & _
                                 "(Zlprocedure_Id.Nextval,1,'" & !P_Name & "',1,'" & !P_Owner & "'," & !P_SysNum & ",'" & !P_Ver & "')"
                Else
                    strSQL = "Insert Into Zlprocedure (ID, 类型, 名称, 状态, 所有者, 系统编号, 升级后版本) Values" & vbNewLine & _
                                 "(Zlprocedure_Id.Nextval,1,'" & !P_Name & "',1,'" & !P_Owner & "'," & !P_SysNum & ",'" & !P_Ver & "')"
                End If
            Else
                '删除已转出的内容
                gcnOracle.Execute "Delete from zlProcedureText where 性质=3 and 过程ID = (Select ID From zlProcedure where 状态 = 4 And ID = " & lngID & ")"
                gcnOracle.Execute "Update zlProcedure Set 状态 = 1 Where 状态 = 4 And ID = " & lngID    '只修改已转出过程的状态
                
                '更新数据
                If intType = 1 Then
                    strSQL = "Update zlProcedure Set 类型 = 1,所有者='" & !P_Owner & "',系统编号=" & !P_SysNum & ",升级前版本='" & !P_Ver & "'" & vbNewLine & _
                                 "Where Id = " & lngID
                Else
                    strSQL = "Update zlProcedure Set 类型 = 1,所有者='" & !P_Owner & "',系统编号=" & !P_SysNum & ",升级后版本='" & !P_Ver & "'" & vbNewLine & _
                                 "Where Id = " & lngID
                End If
            End If
            gcnOracle.Execute strSQL

            '删除zlProcedureText中的数据
            If lngID = 0 Then
                lngID = GetProcIdByName(!P_Name)
            End If
            
            If intType = 1 Then
                gcnOracle.Execute "Delete from zlProcedureText where 性质=1 and 过程ID = " & lngID
            Else
                gcnOracle.Execute "Delete from zlProcedureText where 性质=4 and 过程ID = " & lngID
            End If
            
            '插入过程定义到zlProcedureText
            arrTxt = Split(!P_Define, vbNewLine)
            strSQL = "Insert Into zlProcedureText(过程ID,性质,序号,内容) "
            For i = 0 To UBound(arrTxt)
                If i = UBound(arrTxt) Then
                    strSQL = strSQL & vbNewLine & "Select " & lngID & "," & IIf(intType = 1, "1", "4") & "," & (i + 1) & ",'" & Replace(arrTxt(i), "'", "''") & "' From Dual "
                Else
                    strSQL = strSQL & vbNewLine & "Select " & lngID & "," & IIf(intType = 1, "1", "4") & "," & (i + 1) & ",'" & Replace(arrTxt(i), "'", "''") & "' From Dual Union All "
                End If
            Next
            gcnOracle.Execute strSQL
            
            If strIDs = "" Then
                lngSysNum = !P_SysNum
                strIDs = lngID
            Else
                strIDs = strIDs & "," & lngID '拼接所有ID
            End If
            gcnOracle.CommitTrans
            rsProc.MoveNext
        Loop
    End With
    
    '删除非该系统的其他数据,因为有的库zlProcedureText表外键不是级联删除,因此要先删除子表
    If intType = 1 Then
         gcnOracle.BeginTrans
         arrIds = TranStr2Var(strIDs, ",", 2000) '防止字符超长
         For i = 0 To UBound(arrIds)
             strSQL = "Delete From zlProcedureText Where 过程ID In  " & vbNewLine & _
                         "(Select ID from Zlprocedure Where 类型 = 1 And 系统编号 = " & lngSysNum & " And  ID Not In (Select Column_Value From Table(f_Str2list('" & arrIds(i) & "', ','))))"
             gcnOracle.Execute strSQL
         
             strSQL = "Delete From zlProcedure Where 类型 = 1 And 系统编号 = " & lngSysNum & " And  ID Not In (Select Column_Value From Table(f_Str2list('" & arrIds(i) & "', ',')))"
             gcnOracle.Execute strSQL
        Next
        
         gcnOracle.CommitTrans
    End If
    UpdateProc = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    strErr = err.Description
End Function
