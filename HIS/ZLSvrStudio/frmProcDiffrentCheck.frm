VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcDiffrentCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过程差异检查"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   Icon            =   "frmProcDiffrentCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4545
      Index           =   0
      Left            =   240
      ScaleHeight     =   4545
      ScaleWidth      =   10080
      TabIndex        =   3
      Top             =   960
      Width           =   10080
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1755
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   405
         Width           =   1935
         _cx             =   3413
         _cy             =   3096
         Appearance      =   1
         BorderStyle     =   0
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
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
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
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9210
      TabIndex        =   2
      Top             =   5625
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   195
      Picture         =   "frmProcDiffrentCheck.frx":6852
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   75
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "开始(&O)"
      Height          =   350
      Left            =   7995
      TabIndex        =   0
      Top             =   5625
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pbr 
      Height          =   105
      Left            =   180
      TabIndex        =   5
      Top             =   6045
      Visible         =   0   'False
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6705
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   5805
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "读取最新脚本过程与当前自定过程对应的标准过程进行比较得出差异"
      Height          =   180
      Left            =   1215
      TabIndex        =   7
      Top             =   630
      Width           =   5400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "过程差异检查"
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
      Left            =   1185
      TabIndex        =   6
      Top             =   150
      Width           =   1980
   End
End
Attribute VB_Name = "frmProcDiffrentCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mobjMain As Object
Private mclsVsf As clsVsf
Private mblnReading As Boolean

Public Function ShowMe(ByVal objMain As Object)
    On Error GoTo errHand
    mblnOk = False
    Set mobjMain = objMain
    Me.Show 1, mobjMain
    
    ShowMe = mblnOk
    
    Exit Function
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim strSQL As String
    Dim objItem As Object
    Dim intRow As Integer
    Dim intFlag As Integer
    Dim strUpPath As String
    Dim strFlag As String
    
    On Error GoTo errHand
    mblnReading = True
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True)
            Call .ClearColumn
            Call .AppendColumn("选择", 500, flexAlignLeftCenter, flexDTBoolean, , "", False)
            Call .AppendColumn("版本号", 0, flexAlignLeftCenter, flexDTString, , "", False)
            Call .AppendColumn("系统名称", 2000, flexAlignLeftCenter, flexDTString, , "", False)
            Call .AppendColumn("安装脚本", 2800, flexAlignLeftCenter, flexDTString, , "", True)
            Call .AppendColumn("升级脚本", 0, flexAlignLeftCenter, flexDTString, , "", True)
            
            Call .InitializeEdit(True, False, False)
            Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
            Call .InitializeEditColumn(.ColIndex("安装脚本"), True, vbVsfEditCommand)
            Call .InitializeEditColumn(.ColIndex("升级脚本"), True, vbVsfEditCommand)

'            .AppendRows = True
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        
        With vsf(0)
            strSQL = "Select A.编号,A.版本号,A.名称 as 系统名称,B.文件名 From zlSystems A,zlSysFiles B Where A.编号 = B.系统 And B.操作=1"
            Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
            If rs.BOF = False Then
                For intRow = 0 To rs.RecordCount - 1
                    intFlag = intFlag + 1
                    If .Rows < intFlag + 1 Then .Rows = intFlag + 1
                    .TextMatrix(intRow + 1, .ColIndex("系统名称")) = rs("系统名称").value
                    .TextMatrix(intRow + 1, .ColIndex("安装脚本")) = rs("文件名").value
                    
                    strFlag = rs("版本号").value
                    .TextMatrix(intRow + 1, .ColIndex("版本号")) = strFlag
                    strFlag = Split(strFlag, ".")(0) & "." & Split(strFlag, ".")(1) & ".0"
                    '缺省升级脚本
                    strUpPath = Split(rs("文件名").value, "应用脚本")(0) & "升级脚本\" & strFlag & "\zlUpgrade.ini"
                    If gobjFile.FileExists(strUpPath) = True Then
                        .TextMatrix(intRow + 1, .ColIndex("升级脚本")) = strUpPath
                    End If
                    
                    .RowData(intRow + 1) = rs("编号").value
                    rs.MoveNext
                Next
            End If
        End With
    End Select
    ExecuteCommand = True
    GoTo errEnd
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
    Exit Function
errEnd:
    mblnReading = False
    Exit Function
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    Dim strTemp As String
    Dim str上次标准过程路径 As String
    Dim str最新标准过程路径 As String
    Dim str对比报告路径 As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim lngLoop As Long
    Dim i As Integer
    Dim strProcName As String
    Dim strIniPath As String
    Dim rsInit As ADODB.Recordset
    Dim intSysNumLast As Integer
    Dim strFlag As String
    Dim strCommand As String
    Dim lngTemp As Long
    Dim lngProcess As Long
    Dim rsSQL As ADODB.Recordset
    Dim objFolder As Folder
    Dim objFile As File
    Dim objFSO As TextStream
    Dim lngMaxLength As Long
    Dim str As String
    Dim strArr() As String
    Dim strIni1 As String
    Dim strIniSys As String
    Dim strIniApp As String
    Dim lngSys As Long
    
    Call gclsBase.SQLRecord(rsSQL)
    
    cmdOK.Enabled = False
    
    lblTitle = "正在初始化.."
    lblTitle.Visible = True
    
    str上次标准过程路径 = App.Path & "\Tmp1"
    str最新标准过程路径 = App.Path & "\NewProcedure"
    str对比报告路径 = App.Path & "\Reports"
        
        
    With vsf(0)
        strSQL = "Select 编号,名称,版本号 From zlSystems a"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
        If rsData.BOF = True Then
            MsgBox "当前数据库没有安装任何系统。", vbInformation + vbOKOnly, "中联软件"
            GoTo errEnd
        End If
        For i = 1 To .Rows - 1
            If IIf(Abs(Val(.TextMatrix(i, .ColIndex("选择")))) = 1, True, False) = True Then
                rsData.Filter = ""
                rsData.Filter = "编号=" & .RowData(i)

                If .TextMatrix(i, vsf(0).ColIndex("安装脚本")) = "" Then
                    MsgBox "请选择" & .TextMatrix(i, .ColIndex("系统名称")) & "安装脚本"
                    GoTo errEnd
                End If
                Set rsInit = ReadINIToRec(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")))
                rsInit.Filter = "项目='版本号'"
                strIniApp = rsInit("内容").value

                rsData.Filter = ""
                rsData.Filter = "编号=" & .RowData(i)
                strIniSys = Trim(rsData("版本号").value)
                
                If strIniSys <> strIniApp Then
                    MsgBox .TextMatrix(i, .ColIndex("系统名称")) & "数据库系统版本与配置文件版本不匹配。", vbInformation + vbOKOnly, "中联软件"
                    GoTo errEnd
                End If
            End If
        Next
    End With
    
    '生成两个临时文件夹
    If gobjFile.FolderExists(str上次标准过程路径) Then Call gobjFile.DeleteFolder(str上次标准过程路径)
    If gobjFile.FolderExists(str最新标准过程路径) Then Call gobjFile.DeleteFolder(str最新标准过程路径)
    DoEvents
    
    Call gobjFile.CreateFolder(str上次标准过程路径)
    Call gobjFile.CreateFolder(str最新标准过程路径)
    lblTitle.Visible = True
    
    
    '------------------------------------------------------------------------------------------------------------------
    '提取最新安装脚本及升级脚本中包含的变动过程，并放到临时文件夹1中
    
    For i = 1 To vsf(0).Rows - 1
        If Abs(Val(vsf(0).TextMatrix(i, vsf(0).ColIndex("选择")))) = 1 Then
            
            '提取安装脚本和升级脚本的过程再生成单个脚本文件
            '读取安装脚本
            
            If Not gobjFile.FileExists(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本"))) Then
                MsgBox "无法打开脚本文件" & vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")) & ",执行中断。", vbExclamation, gstrSysName
                Exit Sub
            Else
                strIniPath = Mid(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")), 1, Len(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本"))) - 11)
                strIniPath = strIniPath & "zlProgram.sql"
            End If
            
            lblTitle.Caption = "正在提取“" & vsf(0).TextMatrix(i, vsf(0).ColIndex("系统名称")) & "”安装脚本.."
            Call CheckProcedure(strIniPath, str最新标准过程路径)
            pbr.value = 0
            pbr.Visible = False
            
            '提取升级脚本
            strIniSys = vsf(0).TextMatrix(i, vsf(0).ColIndex("版本号"))
            If Split(strIniSys, ".")(2) = 0 Then
                GoTo errNext
            ElseIf Not gobjFile.FolderExists(Split(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")), "应用脚本")(0) & "升级脚本\" & Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & ".0") Then
                MsgBox "无法检测到升级脚本文件夹,执行中断。", vbExclamation, gstrSysName
                GoTo errEnd
            Else
                strIniPath = Split(vsf(0).TextMatrix(i, vsf(0).ColIndex("安装脚本")), "应用脚本")(0) & "升级脚本\" & Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & ".0" & "\"
            End If
        
'            If Not gobjFile.FileExists(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本"))) And vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本")) <> "" Then
'                MsgBox "无法打开脚本文件" & vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本")) & ",执行中断。", vbExclamation, gstrSysName
'                Exit Sub
'            ElseIf Trim(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本"))) = "" Then
'                GoTo errNext
'            Else
'                strIniPath = Mid(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本")), 1, Len(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本"))) - 13)
'            End If
            
'            Set rsInit = ReadINIToRec(vsf(0).TextMatrix(i, vsf(0).ColIndex("升级脚本")))
'            If Not CheckINIValid(rsInit, "系统号|目标版本") Then
'                MsgBox "升迁配置文件格式不正确。", vbExclamation, "中联软件"
'                Exit Sub
'            End If
            
            
'            lblTitle.Caption = "正在提取“" & vsf(0).TextMatrix(i, vsf(0).ColIndex("系统名称")) & "”升级脚本.."
'            rsInit.Filter = "项目='目标版本'"
             intSysNumLast = Split(strIniSys, ".")(2)
            For lngLoop = 10 To intSysNumLast Step 10
                strFlag = Split(strIniSys, ".")(0) & "." & Split(strIniSys, ".")(1) & "." & CStr(lngLoop)
                Call CheckProcedure(strIniPath & "ZL1_" & strFlag & ".sql", str最新标准过程路径)
            Next
        End If

errNext:

    Next
    
    '------------------------------------------------------------------------------------------------------------------
    '提取变动过程对应的上次标准过程
    lblTitle = "正在准备上次的标准过程.."
    lblTitle.Visible = True
    strSQL = "Select ID,名称,所有者 From zlprocedure Where 类型 In (1,2)"
    Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
    If rsData.BOF = False Then
        pbr.value = 0
        pbr.Visible = True
        pbr.Max = rsData.RecordCount
        For lngLoop = 0 To rsData.RecordCount - 1
            strProcName = Nvl(rsData("名称").value)
            
            If strProcName = "NEXTNO" Then
                strProcName = "NEXTNO"
            End If
            
            If gobjFile.FileExists(str最新标准过程路径 & "\" & strProcName & ".sql") Then
                strSQL = "Select A.ID,A.名称,Upper(B.内容) As 内容 From zlProcedure A,zlProcedureText B Where A.ID = B.过程ID And B.性质 = 4 And A.ID=[1] Order By B.序号"
                Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", Val(Nvl(rsData("ID").value)))
                If rs.BOF = False Then
                    strTemp = ""
                    Do While Not rs.EOF
                        strTemp = strTemp & UCase(Nvl(rs("内容").value))
                        rs.MoveNext
                    Loop
'                    rs.MoveFirst
                    Set objFSO = gobjFile.CreateTextFile(str上次标准过程路径 & "\" & CStr(strProcName) & ".sql")
                    Call objFSO.Write(strTemp)
                    Call objFSO.Close
                End If
            End If
            rsData.MoveNext
            pbr.value = pbr.value + 1
        Next
    Else
        lblTitle.Visible = False
        MsgBox "当前管理工具没有标准过程和空白过程！", vbInformation + vbOKOnly, "中联软件"
        Exit Sub
    End If
    
    '----------------------------调用第三方工具进行对比两个文件夹中的脚本并生成报告---------------------------------------------
    If gobjFile.FolderExists(str对比报告路径) Then
        Call gobjFile.DeleteFolder(str对比报告路径)
    End If
    Call gobjFile.CreateFolder(str对比报告路径)
    '将数据库中的过程与脚本进行比对，生成html报告
    lblTitle.Caption = "正在比对.."
    If Not CompareFolder(str上次标准过程路径, str最新标准过程路径, str对比报告路径) Then
        Exit Sub
    End If
    '--------------------------将有差异的过程自动将它的调整状态修改为"待调整"---------------------------------------------------
    Set objFolder = gobjFile.GetFolder(str对比报告路径)
    lblTitle.Caption = "正在调整过程状态.."
    '报告中存在的即为需要调整的过程
    rsData.MoveFirst
    For i = 0 To rsData.RecordCount - 1
        If gobjFile.FileExists(str最新标准过程路径 & "\" & Nvl(rsData("名称").value) & ".sql") Then
            If gobjFile.FileExists(str对比报告路径 & "\" & Nvl(rsData("名称").value) & ".sql.htm") Then
            '标准过程在升级前后有变化
                strProcName = Nvl(rsData("名称").value)
                Set rs = gclsBase.GetProInfo(strProcName)
                If rs.BOF = False Then
                    strSQL = "Zl_Zlprocedure_Update(" & rs("ID").value & "," & rs("类型").value & ",'" & strProcName & "'," & ProcState.待调整 & ",'','" & Nvl(rsData("所有者").value) & "')"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                End If
                strTemp = ""
                Set objFSO = gobjFile.OpenTextFile(str最新标准过程路径 & "\" & strProcName & ".sql")
                Do While Not objFSO.AtEndOfStream
                    If objFSO.Line = 1 Then
                        strTemp = strTemp & Replace(objFSO.ReadLine, "'", "''")
                    Else
                        strTemp = strTemp & vbCrLf & Replace(objFSO.ReadLine, "'", "''")
                    End If
                    DoEvents
                Loop
                Call objFSO.Close
                lngMaxLength = 3900
                If LenB(StrConv(strTemp, vbFromUnicode)) > lngMaxLength Then
                    strFlag = ""
                    str = ""
                    For lngLoop = 1 To Len(strTemp)
                        str = str & Mid(strTemp, lngLoop, 1)
                        If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngLoop = Len(strTemp)) And Mid(strTemp, lngLoop, 1) <> "'" Then
                            strFlag = strFlag & gstrSplite & str
                            str = ""
                        End If
                    Next
                    strFlag = Mid(strFlag, Len(gstrSplite) + 1)
                    strTemp = strFlag
                End If
                strArr = Split(strTemp, gstrSplite)
'                strSQL = "Zl_Zlproceduretext_Move(" & NVL(rsData("ID").value) & ",3,1,4,2)"
                For lngLoop = 0 To UBound(strArr)
'                    strSQL = "Zl_Zlproceduretext_Update(" & NVL(rsData("ID").value) & ",3," & (lngLoop + 1) & ",'" & strArr(lngLoop) & "')"
'                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                    strSQL = "Zl_Zlproceduretext_Update(" & Nvl(rsData("ID").value) & ",4," & (lngLoop + 1) & ",'" & TrimNull(strArr(lngLoop)) & "')"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                Next
            Else
                '标准过程在升级前后无变化
                strProcName = Nvl(rsData("名称").value)
                strSQL = "Select A.ID,A.类型,A.名称,Upper(B.内容) As 内容 From zlProcedure A,zlProcedureText B Where A.ID = B.过程ID And B.性质 = 3 And A.名称=[1]"
                Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", strProcName)
                If rs.BOF = False Then
                    strSQL = "Zl_Zlprocedure_Update(" & rs("ID").value & "," & rs("类型").value & ",'" & strProcName & "',3,'','" & Nvl(rsData("所有者").value) & "')"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                    strTemp = ""
                    For lngLoop = 0 To rs.RecordCount - 1
                        strTemp = Replace(Nvl(rs("内容").value), "'", "''")
                        
                        strSQL = "Zl_Zlproceduretext_Update(" & Nvl(rsData("ID").value) & ",3," & (lngLoop + 1) & ",'" & strTemp & "')"
                        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                        rs.MoveNext
                    Next
                End If
                Set objFSO = gobjFile.OpenTextFile(str最新标准过程路径 & "\" & strProcName & ".sql")
                strTemp = ""
                Do While Not objFSO.AtEndOfStream
                    If objFSO.Line = 1 Then
                        strTemp = strTemp & Replace(objFSO.ReadLine, "'", "''")
                    Else
                        strTemp = strTemp & vbCrLf & Replace(objFSO.ReadLine, "'", "''")
                    End If
                    DoEvents
                Loop
                
                Call objFSO.Close
                lngMaxLength = 3900
                If LenB(StrConv(strTemp, vbFromUnicode)) > lngMaxLength Then
                    strFlag = ""
                    str = ""
                    For lngLoop = 1 To Len(strTemp)
                        str = str & Mid(strTemp, lngLoop, 1)
                        If (LenB(StrConv(str, vbFromUnicode)) > lngMaxLength - 1 Or lngLoop = Len(strTemp)) And Mid(strTemp, lngLoop, 1) <> "'" Then
                            strFlag = strFlag & gstrSplite & str
                            str = ""
                        End If
                    Next
                    strFlag = Mid(strFlag, Len(gstrSplite) + 1)
                    strTemp = strFlag
                End If
                strArr = Split(strTemp, gstrSplite)
                
                For lngLoop = 0 To UBound(strArr)
                    strSQL = "Zl_Zlproceduretext_Update(" & Nvl(rsData("ID").value) & ",4," & (lngLoop + 1) & ",'" & TrimNull(strArr(lngLoop)) & "')"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                Next
                
                
            End If
        End If
        rsData.MoveNext
    Next
    
    On Error Resume Next
    
    
    objFSO.Close
    Set objFSO = Nothing
    On Error GoTo errHand
    Call SQLRecordExecute(rsSQL, "")
    If gobjFile.FolderExists(str最新标准过程路径) Then
        Call gobjFile.DeleteFolder(str最新标准过程路径)
    End If
    If gobjFile.FolderExists(str上次标准过程路径) Then
        Call gobjFile.DeleteFolder(str上次标准过程路径)
    End If
    If gobjFile.FolderExists(str对比报告路径) Then
        Call gobjFile.DeleteFolder(str对比报告路径)
    End If
    
    lblTitle.Visible = False
    
    MsgBox "差异检查已经完成！", vbInformation, Me.Caption
    cmdOK.Enabled = True
    mblnOk = True
    Exit Sub
errEnd:
    mblnOk = True
    cmdOK.Enabled = True
    Exit Sub
errHand:
    MsgBox "差异检查失败！" & vbCrLf & err.Description, vbCritical, Me.Caption
    cmdOK.Enabled = True
End Sub

Public Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'功能：将指定INI配置文件的内容读取到记录集中
'返回：Nothing或包含"项目,内容"的记录集,其中同一项目可能有多行内容
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As Scripting.TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "项目", adVarChar, 100
    rsTmp.Fields.Append "内容", adVarChar, 4000, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        If Left(Trim(strLine), 1) = "[" And InStr(strLine, "]") > InStr(strLine, "[") Then
            
            If strItem <> "" And strText = "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = Null
                rsTmp.Update
            End If
            
            strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
            strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))
            If strItem <> "" And strText <> "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = strText
                rsTmp.Update
            End If
        ElseIf Trim(strLine) <> "" And strItem <> "" Then
            strText = Trim(strLine)
            rsTmp.AddNew
            rsTmp!项目 = strItem
            rsTmp!内容 = strText
            rsTmp.Update
        End If
    Loop
    
    If strItem <> "" And strText = "" Then
        rsTmp.AddNew
        rsTmp!项目 = strItem
        rsTmp!内容 = Null
        rsTmp.Update
    End If
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function

Private Function CheckINIValid(rsINI As ADODB.Recordset, ByVal strItem As String) As Boolean
'功能：检查对应的配置文件格式是否正确
'参数：rsINI=存放配置文件内容的记录集，包含"项目,内容"字段
'      strItem=配置文件中必须要求有内容的项目串,如"项目1|项目2|..."
    Dim arrItem As Variant, i As Long
    
    arrItem = Split(strItem, "|")
    For i = 0 To UBound(arrItem)
        rsINI.Filter = "项目='" & arrItem(i) & "'"
        If rsINI.EOF Then Exit Function
        If IsNull(rsINI!内容) Then Exit Function
    Next
    CheckINIValid = True
End Function

Private Function CheckProcedure(ByVal strFile As String, Optional strFilePath As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim lngLine As Long
    Dim strLine As String
    Dim strTemp As String
    Dim strFMT As String
    Dim blnSQL As Boolean
    Dim blnBlock As Boolean
    Dim strFlag As String
    Dim strFileProName As String
    Dim lngFileLines As Long
    Dim objFileTemp As TextStream
    Dim objFile As TextStream
    Dim blnFlag As Boolean
    Dim objPercent As New clsPercent
    Dim lngMsg As Long
    
    On Error GoTo errHand
    
    pbr.value = 0
    pbr.Visible = True

    Set objFile = gobjFile.OpenTextFile(strFile, ForReading)
    If objFile.AtEndOfStream Then
        objFile.Close
        Exit Function
    End If
        
    Do While Not objFile.AtEndOfStream
        objFile.ReadLine
    Loop
    lngFileLines = objFile.Line
    
    Call objPercent.InitPercent(pbr, lngFileLines)
    
    objFile.Close
    
    Dim blnSpaceProc As Boolean
    
    Set objFile = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objFile.AtEndOfStream
        lngLine = objFile.Line '当前行号:未读取行之前,行指针未移到下一行
        strLine = objFile.ReadLine
        strFMT = UCase(TrimComment(TrimEx(strLine)))
        If strFMT Like "PROMPT *" Then GoTo NextLine
        
        
        If blnBlock Then
            If strFMT = "/" Then
                blnSQL = True
                blnBlock = False
                Do While Right(strTemp, 1) = Chr(10) Or Right(strTemp, 1) = Chr(13)
                   strTemp = Left(strTemp, Len(strTemp) - 1)
                Loop
                
                
                objFileTemp.Write "CREATE OR REPLACE " & strTemp
                DoEvents
                objFileTemp.Close
                strTemp = ""
                
                If blnSpaceProc = True Then
                    blnSpaceProc = False
                    
                    Set objFileTemp = gobjFile.OpenTextFile(strFilePath & "\" & strFileProName & ".sql")
                    strTemp = objFileTemp.ReadAll
                    objFileTemp.Close
                    strTemp = GetBlankProcedure(strTemp)
                    
                    DoEvents
                    Set objFileTemp = gobjFile.CreateTextFile(strFilePath & "\" & strFileProName & ".sql", True)
                    objFileTemp.Write strTemp
                    objFileTemp.Close
                    strTemp = ""
                End If
                
            Else
                strTemp = strTemp & vbCrLf & strLine
            End If
        ElseIf strFMT Like "CREATE OR REPLACE PROCEDURE *" Or strFMT Like "CREATE PROCEDURE *" _
            Or strFMT Like "CREATE OR REPLACE FUNCTION *" Or strFMT Like "CREATE FUNCTION *" _
            Or strFMT Like "CREATE OR REPLACE TRIGGER *" Or strFMT Like "CREATE TRIGGER *" _
            Or strFMT Like "CREATE OR REPLACE TYPE *" Or strFMT Like "CREATE TYPE *" _
            Or strFMT Like "CREATE OR REPLACE PACKAGE *" Or strFMT Like "CREATE PACKAGE *" Then
            
            blnBlock = True
            
            '创建单个过程脚本文件
            strFlag = Replace(strFMT, "CREATE OR REPLACE ", "")
            strFlag = Replace(strFlag, "CREATE ", "")
            
            If InStr(strFlag, "(") > 0 Then strFlag = Left(strFlag, InStr(strFlag, "(") - 1)
            If InStr(strFlag, ".") > 0 Then strFlag = Split(strFlag, ".")(1)
            strFileProName = Split(strFlag, " ")(1)
            If gobjFile.FileExists(strFilePath & "\" & strFileProName & ".sql") Then
                Call gobjFile.DeleteFile(strFilePath & "\" & strFileProName & ".sql")
            End If
            
            '检查是否为空白过程
            blnSpaceProc = False
            If IsSpaceProcedure("ZLHIS", strFileProName) = True Then
                blnSpaceProc = True
            End If
            
            Set objFileTemp = gobjFile.CreateTextFile(strFilePath & "\" & strFileProName & ".sql", True)
             
            strFlag = Replace(strFMT, "CREATE OR REPLACE ", "")
            strFlag = Replace(strFlag, "CREATE ", "")
            strTemp = strTemp & UCase(strFlag)
        End If
        
        Call objPercent.LoopPercent

NextLine:
    Loop
    objFile.Close
    pbr.Visible = False
    pbr.value = 0
'    MsgBox blnFlag
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    Do While InStr(strText, "  ") > 0
        strText = Replace(strText, "  ", " ")
    Loop
    TrimEx = strText
End Function

Public Function TrimComment(ByVal strSQL As String) As String
'功能：去掉写在单行strSQL语句后面的"--"注释
'说明：主要是RunSQLFile的子函数
    Dim blnStr As Boolean
    Dim i As Long, K As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                K = i: Exit For
            End If
        Next
        If K > 0 Then strSQL = RTrim(Left(strSQL, K - 1))
    End If
    TrimComment = strSQL
End Function

Private Sub Form_Load()
    Call ExecuteCommand("初始控件")
    Call ExecuteCommand("初始数据")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsf(0).Move 15, 15, picPane(0).ScaleWidth - 30, picPane(0).ScaleHeight - 30
'    mclsVsf.AppendRows = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If mblnReading = True Then Exit Sub
    Call mclsVsf.AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Call mclsVsf.AfterMoveColumn(Col, Position)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnReading = True Then Exit Sub
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnReading = True Then Exit Sub
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    With vsf(0)
        Select Case Col
        '--------------------------------------------------------------------------------------------------------------
        Case .ColIndex("安装脚本")
            With dlg
                .DialogTitle = "选择应用安装配置文件"
                .Filter = "(应用安装配置文件)|zlSetup.ini"
                .ShowOpen
                If .FileName = "" Then
                    Exit Sub
                Else
                    vsf(0).TextMatrix(vsf(0).Row, vsf(0).Col) = .FileName
                End If
            End With
        Case .ColIndex("升级脚本")
            With dlg
                .DialogTitle = "选择应用升迁配置文件"
                .Filter = "应用升迁配置文件(zlUpgrade.ini)|zlUpgrade.ini"
                .Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
                .ShowOpen
                On Error GoTo 0
                Me.Refresh
                If .FileName = "" Then
                    Exit Sub
                Else
                    vsf(0).TextMatrix(vsf(0).Row, vsf(0).Col) = .FileName
                End If
            End With
        End Select
        
        Call mclsVsf.SetFocus(, , True)
    End With
End Sub




