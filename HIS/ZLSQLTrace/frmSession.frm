VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSession 
   AutoRedraw      =   -1  'True
   Caption         =   "会话"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   8655
   Icon            =   "frmSession.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   8655
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSession 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   8655
      TabIndex        =   17
      Top             =   0
      Width           =   8655
      Begin VB.CheckBox chkLike 
         Caption         =   "模糊查找"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   83
         Width           =   1095
      End
      Begin VB.CheckBox chkTraceOnly 
         Caption         =   "只显示已跟踪的用户"
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   83
         Width           =   2055
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Left            =   1380
         TabIndex        =   20
         Top             =   60
         Width           =   855
      End
      Begin VSFlex8Ctl.VSFlexGrid vsSession 
         Height          =   2400
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   8655
         _cx             =   15266
         _cy             =   4233
         Appearance      =   2
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSession.frx":058A
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
         ExplorerBar     =   1
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
         Begin MSComDlg.CommonDialog cdgFile 
            Left            =   3090
            Top             =   1290
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList imgSession 
            Left            =   2310
            Top             =   1260
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSession.frx":06BD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSession.frx":0C57
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSession.frx":11F1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSession.frx":178B
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label lblPrompt 
         Caption         =   "提示"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5760
         TabIndex        =   23
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblLocate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户名查找(&S)"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1170
      End
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   30
      MousePointer    =   7  'Size N S
      TabIndex        =   11
      Top             =   3120
      Width           =   7110
   End
   Begin VB.Frame fraTrace 
      Caption         =   " Trace 文件解析 "
      Height          =   2000
      Left            =   30
      TabIndex        =   10
      Top             =   3240
      Width           =   8595
      Begin ZLSQLTrace.ccXPButton cmdOut 
         Height          =   240
         Left            =   8175
         TabIndex        =   3
         Top             =   645
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   423
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ZLSQLTrace.ccXPButton cmdIn 
         Height          =   240
         Left            =   6285
         TabIndex        =   1
         Top             =   285
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   423
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ZLSQLTrace.ccXPButton cmdTrace 
         Height          =   345
         Left            =   7305
         TabIndex        =   8
         Top             =   1005
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "解析(&A)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboCount 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3450
         TabIndex        =   5
         Text            =   "cboCount"
         Top             =   1035
         Width           =   975
      End
      Begin VB.CheckBox chkNoSys 
         Caption         =   "排开系统语句"
         Height          =   195
         Left            =   4515
         TabIndex        =   6
         Top             =   1088
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox chkOpen 
         Caption         =   "自动打开"
         Height          =   195
         Left            =   6240
         TabIndex        =   7
         Top             =   1088
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cboSort 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1035
         Width           =   1620
      End
      Begin VB.TextBox txtCmd 
         Height          =   465
         IMEMode         =   2  'OFF
         Left            =   990
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1425
         Width           =   7470
      End
      Begin VB.TextBox txtOut 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   990
         TabIndex        =   2
         Top             =   615
         Width           =   7470
      End
      Begin VB.TextBox txtIn 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   990
         TabIndex        =   0
         Top             =   255
         Width           =   5550
      End
      Begin ZLSQLTrace.ccXPButton cmdFile 
         Height          =   300
         Left            =   6600
         TabIndex        =   24
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Caption         =   "获取ZLTRACE文件"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "解析条数"
         Height          =   180
         Left            =   2685
         TabIndex        =   16
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "排序方式"
         Height          =   180
         Left            =   210
         TabIndex        =   15
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "命令文本"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   14
         Top             =   1470
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "目标文件"
         Height          =   180
         Left            =   210
         TabIndex        =   13
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "来源文件"
         Height          =   180
         Left            =   210
         TabIndex        =   12
         Top             =   315
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event PopSessionMenu() '子窗体特有事件
Public Event OpenNewFile(ByVal File As String)
Public Event UpdateStatus(ByVal strStatus As String)
Private mcolTrace As New Collection
Private mstrDest As String
Private mstrDBName As String
Private mblnEv As Boolean
Public mlngCount As Long

Private mcolCon As New Collection

Public Sub ShowMe(frmMain As Object)
    mlngCount = 0
    Me.Show
End Sub

Public Sub DoCommand(ByVal DoID As CommandBarIDCond, Optional ByVal blnIsZlTraceFile)
'功能：子窗体命令执行接口
    Dim intEv As Integer, strTmp As String
    Dim cnTmp As New adodb.Connection, strConnect As String
    Dim rstmp As adodb.Recordset, strSql As String, strNewInstance

    On Error GoTo errh
    
    With vsSession
        Select Case DoID
        Case conMenu_Edit_Trace_1, conMenu_Edit_Trace_4, conMenu_Edit_Trace_8, conMenu_Edit_Trace_12
            
            If gblnIsRac Then
                If .TextMatrix(.Row, .ColIndex("Inst_ID")) <> gintInstId Then
                    Set cnTmp = GetConnection(.TextMatrix(.Row, .ColIndex("Inst_ID")))
                    If cnTmp.ConnectionString <> "" Then
                         Set gcnOracle = cnTmp
                         gintInstId = Val(.TextMatrix(.Row, .ColIndex("Inst_ID")))
                    Else
                        strSql = "Select Inst_Name From Gv$active_Instances Where Inst_Id = [1]"
                        Set rstmp = OpenSQLRecord(strSql, "GETINSTANCESNAME", .TextMatrix(.Row, .ColIndex("Inst_ID")))
                        strNewInstance = rstmp!Inst_Name
                        Set cnTmp = gcnOracle
                        strConnect = gcnOracle.ConnectionString
reLogin:
                        MsgBox "当前登录实例与所跟踪实例(" & strNewInstance & ")不一致,请重新登录到实例(" & strNewInstance & ")后再执行跟踪。"
                        frmUserLogin.Show 1
                        
                        '如果点击取消,就不执行跟踪方法
                        If gcnOracle Is Nothing Then
                            Set gcnOracle = cnTmp
                            Exit Sub
                        End If
                        '如果再次登录同一个实例,就弹出窗体要求重新登录
                        If gcnOracle.ConnectionString = strConnect Then GoTo reLogin
                        mcolCon.Add gcnOracle, "_" & gintInstId   '缓存集合
                    End If
                End If
            End If
            
            'If MsgBox("确定要对 用户名 = " & .TextMatrix(.Row, vssession.colindex("用户名")) & ",SID = " & .TextMatrix(.Row, vssession.colindex("SID")) & ",Serial# = " & .TextMatrix(.Row, vssession.colindex("Serial#")) & " 的会话进行跟踪吗？", vbQuestion + vbYesNo + vbDefaultButton1, App.Title) = vbNo Then Exit Sub
            Me.Tag = "正在跟踪"
                                    
            intEv = Decode(DoID, conMenu_Edit_Trace_1, 1, conMenu_Edit_Trace_4, 4, conMenu_Edit_Trace_8, 8, conMenu_Edit_Trace_12, 12)
            If DoID = conMenu_Edit_Trace_1 Then
                '这种方式可能快些
                gcnOracle.Execute "SYS.DBMS_System.Set_SQL_Trace_In_Session(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",True)", , adCmdStoredProc
            Else
                gcnOracle.Execute "SYS.DBMS_System.Set_Bool_Param_In_Session(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",'Timed_Statistics',True)", , adCmdStoredProc
                gcnOracle.Execute "SYS.DBMS_System.Set_Ev(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",10046," & intEv & ",'')", , adCmdStoredProc
            End If
            
            mlngCount = mlngCount + 1
            lblPrompt.Caption = "跟踪用户" & .TextMatrix(.Row, vsSession.ColIndex("用户名")) & "(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ")，操作成功。"
            .RowData(.Row) = intEv
            Set .Cell(flexcpPicture, .Row, 0) = imgSession.ListImages(Decode(intEv, 1, 1, 4, 2, 8, 3, 12, 4)).Picture
            .TextMatrix(.Row, .ColIndex("跟踪状态")) = "已跟踪"
            .TextMatrix(.Row, .ColIndex("跟踪等级")) = intEv
            Err.Clear: On Error Resume Next
            mcolTrace.Add intEv, "_" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "_" & .TextMatrix(.Row, vsSession.ColIndex("Serial#"))
            Err.Clear: On Error GoTo errh
        Case conMenu_Edit_TraceOff
            'If MsgBox("确定要停止对 用户名 = " & .TextMatrix(.Row, vssession.colindex("用户名")) & ",SID = " & .TextMatrix(.Row, vssession.colindex("SID")) & ",Serial# = " & .TextMatrix(.Row, vssession.colindex("Serial#")) & " 的会话进行跟踪吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
            Me.Tag = ""
            
            gcnOracle.Execute "SYS.DBMS_System.Set_Bool_Param_In_Session(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",'Timed_Statistics',False)", , adCmdStoredProc
            gcnOracle.Execute "SYS.DBMS_System.Set_Ev(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",10046,0,'')", , adCmdStoredProc
            '强行把两种都停了,但Set_SQL_Trace_In_Session要放在后面
            gcnOracle.Execute "SYS.DBMS_System.Set_SQL_Trace_In_Session(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ",False)", , adCmdStoredProc
                        
            mlngCount = mlngCount - 1
            lblPrompt.Caption = "停止跟踪用户" & .TextMatrix(.Row, vsSession.ColIndex("用户名")) & "(" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#")) & ")，操作成功。"
            .RowData(.Row) = 0
            Set .Cell(flexcpPicture, .Row, 0) = Nothing
            .TextMatrix(.Row, .ColIndex("跟踪状态")) = ""
            .TextMatrix(.Row, .ColIndex("跟踪等级")) = ""
            Err.Clear: On Error Resume Next
            mcolTrace.Remove "_" & .TextMatrix(.Row, vsSession.ColIndex("SID")) & "_" & .TextMatrix(.Row, vsSession.ColIndex("Serial#"))
            Err.Clear: On Error GoTo errh
            
            If gstrFilePath = "" Then
                gstrFilePath = GetDirName
                If gstrFilePath = "" Then
                    lblPrompt.Caption = "未选择保存路径，无法保存。"
                    Exit Sub
                End If
                Call SaveSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "TraceFilePath", gstrFilePath)
            End If
            
            strTmp = GetTraceFile
            If strTmp <> "" Then
                txtIn.Text = strTmp
            End If
            
            
        
        Case conMenu_Edit_ChangeReg
            strTmp = GetDirName
            If strTmp <> "" Then
                gstrFilePath = strTmp
                Call SaveSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "TraceFilePath", gstrFilePath)
            End If
            
        Case conMenu_View_Refresh
            Call LoadSession
        End Select
    End With
    
    Exit Sub
errh:
    MsgBox Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Function GetCommand(ByVal DoID As CommandBarIDCond) As Boolean
'功能：子窗体命令状态接口
    Select Case DoID
    Case conMenu_Edit_Trace
        GetCommand = mblnEv
        
    Case conMenu_Edit_Trace_1, conMenu_Edit_Trace_4, conMenu_Edit_Trace_8, conMenu_Edit_Trace_12
        If vsSession.Rows > vsSession.FixedRows Then
            GetCommand = vsSession.RowData(vsSession.Row) = 0 And mblnEv
        End If
    Case conMenu_Edit_TraceOff
        GetCommand = Me.Tag = "正在跟踪" Or Me.Tag = "已跟踪"
    Case conMenu_View_Refresh
        GetCommand = gcnOracle.State = 1
    End Select
End Function

Private Sub cboCount_Change()
    Call MakeCmdLine
End Sub

Private Sub cboCount_Click()
    Call MakeCmdLine
End Sub

Private Sub cboCount_GotFocus()
    cboCount.SelStart = 0: cboCount.SelLength = Len(cboCount.Text)
End Sub

Private Sub cboCount_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cboSort_Click()
    Call MakeCmdLine
End Sub

Private Sub chkNoSys_Click()
    Call MakeCmdLine
End Sub

Private Sub chkTraceOnly_Click()
    Dim i As Long
    Dim lngCount As Long
    
    With vsSession
        For i = .FixedRows To .Rows - .FixedRows
            If chkTraceOnly.Value = 1 Then
                If .Cell(flexcpPicture, i, 0) Is Nothing Then
                    .RowHidden(i) = True
                Else
                    .RowHidden(i) = False
                    lngCount = lngCount + 1
                End If
            Else
                .RowHidden(i) = False
                lngCount = lngCount + 1
            End If
        Next
        lblPrompt.Caption = "共" & lngCount & "个用户会话。"
    End With
End Sub

Private Sub cmdFile_Click()
    Dim strTmp As String

    If gstrFilePath = "" Then
        gstrFilePath = GetDirName
        If gstrFilePath = "" Then
            lblPrompt.Caption = "未选择保存路径，无法保存。"
            Exit Sub
        End If
        Call SaveSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "TraceFilePath", gstrFilePath)
    End If
    
    strTmp = GetTraceFile(True)
    If strTmp <> "" Then
        txtIn.Text = strTmp
    End If
    
End Sub

Private Sub cmdIn_Click()
    With Me.cdgFile
        .DialogTitle = "选择要解析的 SQL Trace 文件"
        .Filter = "SQL Trace(*.trc)|*.trc"
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        
        If txtIn.Text <> "" Then
            .InitDir = gobjFile.GetParentFolderName(txtIn.Text)
            .FileName = gobjFile.GetFileName(txtIn.Text)
        Else
            .InitDir = GetSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "Input", mstrDest)
            .FileName = ""
        End If
        
        .CancelError = True
        On Error GoTo errh
        .ShowOpen
        
        SaveSetting "ZLSOFT\公共模块\ZLDBATools", "Setting", "Input", gobjFile.GetParentFolderName(.FileName)
        txtIn.Text = .FileName: txtIn.SetFocus
    End With
errh:
End Sub

Private Sub cmdOut_Click()
    With Me.cdgFile
        .DialogTitle = "确定要解析生成的报告文件"
        .Filter = "SQL Trace(*.log)|*.log"
        .flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
                
        If txtOut.Text <> "" Then
            .InitDir = gobjFile.GetParentFolderName(txtOut.Text)
            .FileName = gobjFile.GetFileName(txtOut.Text)
        Else
            .InitDir = GetSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "Output", mstrDest)
            If txtIn.Text <> "" Then
                .FileName = ChangeType(gobjFile.GetFileName(txtIn.Text), ".trc", ".log")
            Else
                .FileName = ""
            End If
        End If
        
        .CancelError = True
        On Error GoTo errh
        .ShowSave

        SaveSetting "ZLSOFT\公共模块\ZLDBATools", "Setting", "Output", gobjFile.GetParentFolderName(.FileName)
        txtOut.Text = .FileName: txtOut.SetFocus
    End With
errh:
End Sub

Private Sub cmdTrace_Click()
    Dim lngTemp As Long, lngProcess As Long
    
    If txtIn.Text = "" Then
        MsgBox "请确定要解析的 SQL Trace 文件。", vbInformation, App.Title
        txtIn.SetFocus: Exit Sub
    End If
    If Not gobjFile.FileExists(txtIn.Text) Then
        MsgBox "指定要解析的 SQL Trace 文件不存在。", vbInformation, App.Title
        txtIn.SetFocus: Exit Sub
    End If
    If txtOut.Text = "" Then
        MsgBox "请确定解析后生成的 SQL Trace 文件。", vbInformation, App.Title
        txtOut.SetFocus: Exit Sub
    End If
    If txtCmd.Text = "" Then
        MsgBox "无法进行解析。", vbInformation, App.Title
        txtIn.SetFocus: Exit Sub
    End If
        
    Screen.MousePointer = 11
    On Error GoTo errh
    lngTemp = Shell(txtCmd.Tag, vbHide)
    
    '为什么有的机器很慢没反应
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        GetExitCodeProcess lngProcess, lngTemp
    Loop While lngTemp = Still_Active
    CloseHandle lngProcess
    
    '检查是否解析成功
    Screen.MousePointer = 0
    If Dir(txtOut.Text) = "" Then
        MsgBox "来源文件名有误，文件名不能包含中文及特殊符号。" & vbNewLine & "请修改后重新解析。"
    Else
        If chkOpen.Value = 0 Then
            Screen.MousePointer = 0
            MsgBox "文件解析完成。", vbInformation, App.Title
        Else
            RaiseEvent OpenNewFile(txtOut.Text)
        End If
    End If
  
    Exit Sub
errh:
    Screen.MousePointer = 0
    MsgBox Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Activate()
    RaiseEvent UpdateStatus("共有 " & vsSession.Rows - vsSession.FixedRows & " 个会话" & "|当前用户:" & gstrDBUser)
End Sub

Private Sub Form_Load()
    Dim i As Long, strCol As String
    
    With cboSort
        .AddItem ""
        .AddItem "(解析)执行次数:prscnt  number of times parse was called"
        .AddItem "(解析)CPU时间 :prscpu  cpu time parsing"
        .AddItem "(解析)总时间  :prsela  elapsed time parsing"
        .AddItem "(解析)物理读  :prsdsk  number of disk reads during parse"
        .AddItem "(解析)一致读  :prsqry  number of buffers for consistent read during parse"
        .AddItem "(解析)当前读  :prscu   number of buffers for current read during parse"
        .AddItem "(解析)硬解析  :prsmis  number of misses in library cache during parse"
        .AddItem "(执行)执行次数:execnt  number of execute was called"
        .AddItem "(执行)CPU时间 :execpu  cpu time spent executing"
        .AddItem "(执行)总的时间:exeela  elapsed time executing"
        .AddItem "(执行)物理读  :exedsk  number of disk reads during execute"
        .AddItem "(执行)一致读  :exeqry  number of buffers for consistent read during execute"
        .AddItem "(执行)当前读  :execu   number of buffers for current read during execute"
        .AddItem "(执行)记录数  :exerow  number of rows processed during execute"
        .AddItem "(执行)硬解析  :exemis  number of library cache misses during execute"
        .AddItem "(提取)执行次数:fchcnt  number of times fetch was called"
        .AddItem "(提取)CPU时间 :fchcpu  cpu time spent fetching"
        
        .AddItem "(提取)总的时间:fchela  elapsed time fetching"
        
        .AddItem "(提取)物理读  :fchdsk  number of disk reads during fetch"
        .AddItem "(提取)一致读  :fchqry  number of buffers for consistent read during fetch"
        .AddItem "(提取)当前读  :fchcu   number of buffers for current read during fetch"
        .AddItem "(提取)记录数  :fchrow  number of rows fetched"
        .ListIndex = 18
        SendMessage .hWnd, CB_SETDROPPEDWIDTH, 7000 / Screen.TwipsPerPixelX, 0
        SetWindowPos .hWnd, 0, 0, 0, .Width / Screen.TwipsPerPixelX, 5000 / Screen.TwipsPerPixelY, &H2
        
        Set gcolSort = New Collection
        For i = 1 To .ListCount - 1
            gcolSort.Add Trim(Split(.List(i), ":")(0)), "_" & UCase(Split(Split(.List(i), ":")(1), " ")(0))
        Next
    End With
    
    With cboCount
        .AddItem ""
        .AddItem "前50条": .ItemData(.NewIndex) = 50
        .AddItem "前100条": .ItemData(.NewIndex) = 100
        .AddItem "前200条": .ItemData(.NewIndex) = 200
        .ListIndex = 0
    End With
    lblPrompt.Caption = ""
    
    Call InitData
    
    strCol = "  ,2000,1;" & IIf(gblnIsRac, "Inst_ID,2000,1;", "") & "用户名,2000,1;SID,1500,4;Serial#,1500,1;状态,1500,1;姓名,1500,1;部门,1500,1;工作站,500,1;系统用户,500,1;登录程序,500,1;登录时间,1500,1;跟踪状态,500,1;跟踪等级,500,1"
    Call InitTable(vsSession, strCol)
    Call LoadSession
    
    '从注册表读取路径
    gstrFilePath = GetSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "TraceFilePath")
    
    If gstrDBUser <> "" Then
        gblnZlhis = CheckZlhis
    End If
    
    cmdFile.Visible = gstrDBUser <> "" And gblnZlhis
    
    mcolCon.Add gcnOracle, "_" & gintInstId   '缓存集合
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    picSession.Left = 0
    picSession.Top = 0
    picSession.Width = Me.ScaleWidth
    picSession.Height = Me.ScaleHeight - fraTrace.Height - fraUD.Height - 45
    
    
    fraUD.Left = 0
    fraUD.Top = picSession.Top + picSession.Height
    fraUD.Width = Me.ScaleWidth
    
    fraTrace.Left = 60
    fraTrace.Top = fraUD.Top + fraUD.Height
    fraTrace.Width = Me.ScaleWidth - fraTrace.Left * 2
    
    If fraTrace.Width - txtIn.Left - 200 >= 7300 Then
        txtIn.Width = fraTrace.Width - txtIn.Left - IIf(cmdFile.Visible, cmdFile.Width + 245, 200)
        cmdFile.Left = txtIn.Left + txtIn.Width + 45
        cmdIn.Left = txtIn.Left + txtIn.Width - cmdIn.Width - 30
        txtOut.Width = fraTrace.Width - txtOut.Left - 200
        cmdOut.Left = txtOut.Left + txtOut.Width - cmdOut.Width - 30
        
        cmdTrace.Left = txtOut.Left + txtOut.Width - cmdTrace.Width - 30
        chkOpen.Left = cmdTrace.Left - chkOpen.Width - 100
        txtCmd.Width = txtOut.Width
    End If
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub
    If Button = 1 Then
        If picSession.Height + y < 1000 Or fraTrace.Height - y < 2000 Then Exit Sub
        fraUD.Top = fraUD.Top + y
        picSession.Height = picSession.Height + y
        fraTrace.Top = fraTrace.Top + y
        fraTrace.Height = fraTrace.Height - y
    End If
End Sub

Private Sub InitData()
    Dim rstmp As adodb.Recordset
    Dim cmdTmp As adodb.Command
    Dim strSql As String
    
    If gcnOracle.State = 0 Then
        mblnEv = False: Exit Sub
    End If
    
    Err.Clear: On Error Resume Next
    'ORA-00942: 表或视图不存在
    
    'Trace文件生成目录
    strSql = "Select Value as 目录 From v$Parameter Where Name='user_dump_dest'"
    Set rstmp = New adodb.Recordset
    rstmp.CursorLocation = adUseClient
    rstmp.Open strSql, gcnOracle, adOpenKeyset
    If Not rstmp.EOF Then mstrDest = Nvl(rstmp!目录)
    
    '当前数据库实例名
    strSql = "Select SYS_CONTEXT('USERENV','DB_NAME') as 名称 From Dual"
    Set rstmp = New adodb.Recordset
    rstmp.CursorLocation = adUseClient
    rstmp.Open strSql, gcnOracle, adOpenKeyset
    If Not rstmp.EOF Then mstrDBName = Nvl(rstmp!名称)
    
    '检查是否有使用SYS.DBMS_SYSTEM的权限
    mblnEv = True
    Err.Clear
    Set cmdTmp = New adodb.Command
    Set cmdTmp.ActiveConnection = gcnOracle
    cmdTmp.CommandType = adCmdStoredProc
    cmdTmp.CommandText = "SYS.DBMS_System.Read_Ev"
    cmdTmp.Parameters.Append cmdTmp.CreateParameter("IEv", adNumeric, adParamInput, , 10046)
    cmdTmp.Parameters.Append cmdTmp.CreateParameter("OEv", adNumeric, adParamOutput)
    cmdTmp.Execute
    If InStr(Err.Description, "PLS-00201") > 0 Then mblnEv = False
End Sub

Private Sub LoadSession()
    Dim rstmp As New adodb.Recordset
    Dim strSql As String, i As Long
    Dim strPre As String

    If gcnOracle.State = 0 Then
        vsSession.Rows = vsSession.FixedRows
        Exit Sub
    End If
    
    On Error GoTo errh
    
    Screen.MousePointer = 11
    If gblnZlhis Then
        strSql = "Select B.用户名,A.姓名,C.名称 as 部门 From 人员表 A,上机人员表 B,部门表 C,部门人员 D " & vbCrLf & _
                 "Where A.ID=B.人员ID And A.ID = D.人员ID And D.缺省 = 1 And D.部门ID = C.ID"
        strSql = "Select A.*,B.姓名,b.部门 From gv$Session A,(" & strSql & ") B Where A.UserName=B.用户名(+) And A.UserName Is Not Null  And A.AUDSID <> userenv('sessionid') Order By A.UserName"
    Else
        strSql = "Select A.*,Null as 姓名,Null as 部门 From gv$Session A Where A.UserName Is Not Null And A.AUDSID <> userenv('sessionid') Order By A.Logon_Time,A.SID"
    End If
    rstmp.Open strSql, gcnOracle, adOpenKeyset, adLockOptimistic
    
    With vsSession
        strPre = .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#"))
        
        .Redraw = flexRDNone
        .Rows = .FixedRows '清除数据
        .Rows = .FixedRows + rstmp.RecordCount
        For i = 1 To rstmp.RecordCount
            If gblnIsRac Then
                .TextMatrix(i, .ColIndex("Inst_ID")) = rstmp!INST_ID
            End If
            .TextMatrix(i, .ColIndex("用户名")) = rstmp!UserName
            .TextMatrix(i, .ColIndex("SID")) = rstmp!SID
            .TextMatrix(i, .ColIndex("Serial#")) = rstmp.Fields("Serial#").Value
            .TextMatrix(i, .ColIndex("状态")) = Decode(rstmp!Status, "ACTIVE", "当前", "INACTIVE", "正常", rstmp!Status)
            .TextMatrix(i, .ColIndex("姓名")) = Nvl(rstmp!姓名)
            .TextMatrix(i, .ColIndex("部门")) = Nvl(rstmp!部门)
            .TextMatrix(i, .ColIndex("工作站")) = Nvl(rstmp!Machine)
            .TextMatrix(i, .ColIndex("系统用户")) = Nvl(rstmp!OSUser)
            .TextMatrix(i, .ColIndex("登录程序")) = Nvl(rstmp!Program) & IIf(Not IsNull(rstmp!Action), ":" & rstmp!Action, "")
            .TextMatrix(i, .ColIndex("登录时间")) = rstmp!Logon_Time
            
            '检查是否有处于跟踪状态的会话
            .TextMatrix(i, .ColIndex("跟踪状态")) = IIf(Nvl(rstmp!SQL_TRACE) = "ENABLED", "已开启跟踪", "")
'            跟踪等级对应
'            等级    SQL_TRACE   SQL_TRACE_WAITS    SQL_TRACE_BINDS
'            1         ENABLED           FALSE                           FALSE
'            4         ENABLED           FALSE                           TRUE
'            8         ENABLED           TRUE                             FALSE
'            12       ENABLED           TRUE                            TRUE
            
            If Nvl(rstmp!SQL_TRACE) = "ENABLED" Then
                Select Case Nvl(rstmp!SQL_TRACE_WAITS)
                    Case "FALSE"
                        .TextMatrix(i, .ColIndex("跟踪等级")) = IIf(rstmp!SQL_TRACE_BINDS = "FALSE", 1, 4)
                    Case "TRUE"
                        .TextMatrix(i, .ColIndex("跟踪等级")) = IIf(rstmp!SQL_TRACE_BINDS = "FALSE", 8, 12)
                End Select
                Set .Cell(flexcpPicture, i, 0) = imgSession.ListImages(Decode(Val(.TextMatrix(i, .ColIndex("跟踪等级"))), 1, 1, 4, 2, 8, 3, 12, 4)).Picture
            End If
            
            Err.Clear: On Error Resume Next
            .RowData(i) = mcolTrace("_" & rstmp!SID & "_" & rstmp.Fields("Serial#").Value)
            Err.Clear: On Error GoTo errh
            If .RowData(i) <> 0 Then
                Set .Cell(flexcpPicture, i, 0) = imgSession.ListImages(Decode(.RowData(i), 1, 1, 4, 2, 8, 3, 12, 4)).Picture
            End If

            rstmp.MoveNext
        Next
        If .Rows > 2 Then
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = 4
            .AutoSize 1, 9
            .Row = 1: .Col = 1
            .ShowCell .Row, .Col
        End If
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDDirect
        
        lblPrompt.Caption = "共" & .Rows - .FixedRows & "个用户会话。"
    End With
    Screen.MousePointer = 0
    Exit Sub
errh:
    Screen.MousePointer = 0
    If InStr(Err.Description, "ORA-00942") > 0 Then
        MsgBox "当前登录用户没有读取数据库会话信息 v$Session 的权限。", vbCritical, App.Title
    Else
        MsgBox Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title
    End If
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub picSession_Resize()
    On Error Resume Next
    
    vsSession.Top = txtLocate.Top + txtLocate.Height + 60
    vsSession.Left = picSession.Left
    vsSession.Width = picSession.Width
    vsSession.Height = picSession.Height - txtLocate.Height - 60
    
    lblPrompt.Width = vsSession.Width - lblPrompt.Left
End Sub

Private Sub txtCmd_GotFocus()
    txtCmd.SelStart = 0: txtCmd.SelLength = Len(txtCmd.Text)
End Sub

Private Sub txtIn_Change()
    If txtIn.Text <> "" Then
        If gobjFile.FileExists(txtIn.Text) Then
            If txtOut.Text = "" Then
                txtOut.Text = ChangeType(txtIn.Text, ".trc", ".log")
            Else
                txtOut.Text = gobjFile.GetParentFolderName(txtOut.Text) & "\" & ChangeType(gobjFile.GetFileName(txtIn.Text), ".trc", ".log")
            End If
        End If
    End If

    Call MakeCmdLine
End Sub

Private Sub txtIn_GotFocus()
    txtIn.SelStart = 0: txtIn.SelLength = Len(txtIn.Text)
End Sub

Private Sub txtLocate_Change()
    txtLocate.Tag = ""
End Sub

Private Sub txtLocate_GotFocus()
    txtLocate.SelStart = 0
    txtLocate.SelLength = Len(txtLocate.Text)
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Len(Trim(txtLocate.Text)) > 0 Then
        Dim lngRow As Long
        Dim strText As String
        Dim lngStart As Long
        
        If txtLocate.Tag = "" Then
            lngStart = vsSession.FixedRows
        Else
            lngStart = Val(txtLocate.Tag)
        End If
        
        strText = UCase(Trim(txtLocate.Text))
        If chkLike.Value = 1 Then
            lngRow = vsSession.FindRow(strText, lngStart, vsSession.ColIndex("用户名"), , False)
        Else
            lngRow = vsSession.FindRow(strText, lngStart, vsSession.ColIndex("用户名"))
        End If
        If lngRow > 0 Then
            txtLocate.Tag = lngRow + 1
            vsSession.Row = lngRow
            vsSession.TopRow = lngRow
        Else
            If lngStart = vsSession.FixedRows Then
                lblPrompt.Caption = "没有找到匹配的用户名。"
            Else
                lblPrompt.Caption = "当前行是最后一个，后面没有了。"
                txtLocate.Tag = ""
            End If
        End If
        Call txtLocate_GotFocus
    End If
End Sub

Private Sub txtOut_Change()
    Call MakeCmdLine
End Sub

Private Sub txtOut_GotFocus()
    txtOut.SelStart = 0: txtOut.SelLength = Len(txtOut.Text)
End Sub

Private Sub vsSession_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow As Long
    Dim i As Long
    
    If chkTraceOnly.Tag <> "" Then
        With vsSession
            For i = .FixedRows To .Rows - .FixedRows
                If .TextMatrix(i, vsSession.ColIndex("SID")) & "," & .TextMatrix(i, vsSession.ColIndex("Serial#")) = chkTraceOnly.Tag Then
                    .Row = i
                    .TopRow = i
                End If
            Next
        End With
    End If
End Sub

Private Sub vsSession_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Me.Tag = "正在跟踪" Then
        Cancel = True
        Exit Sub
    End If
    
    '已有正在跟踪的会话,将tag标记为"已跟踪",用于设置按钮可用性
    With vsSession
        If .Redraw = flexRDNone Then Exit Sub
        If .TextMatrix(NewRow, .ColIndex("跟踪状态")) = "已开启跟踪" Then
            Me.Tag = "已跟踪"
        Else
            Me.Tag = ""
        End If
    End With
End Sub

Private Sub vsSession_BeforeSort(ByVal Col As Long, Order As Integer)
    With vsSession
        If .Row > .FixedRows Then
            chkTraceOnly.Tag = .TextMatrix(.Row, vsSession.ColIndex("SID")) & "," & .TextMatrix(.Row, vsSession.ColIndex("Serial#"))
        End If
    End With
End Sub

Private Sub vsSession_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsSession_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    
    With vsSession
        lngRow = .MouseRow
        If Button = 2 And .MouseRow >= .FixedRows Then
            .Row = lngRow
            RaiseEvent PopSessionMenu
        End If
    End With
End Sub

Private Sub MakeCmdLine()
    Dim strIn As String, strOut As String
    Dim strCmd As String
    
    If txtIn.Text = "" Or txtOut.Text = "" Then
        txtCmd.Text = "": txtCmd.Tag = ""
    Else
        strCmd = "tkprof " & txtIn.Text & " " & txtOut.Text
        If cboSort.Text <> "" Then
            strCmd = strCmd & " sort=" & Split(Split(cboSort.Text, ":")(1), " ")(0)
        End If
        If cboCount.Text <> "" Then
            If cboCount.ListIndex = -1 Then
                If Int(Val(cboCount.Text)) > 0 Then
                    strCmd = strCmd & " print=" & Int(Val(cboCount.Text))
                End If
            ElseIf cboCount.ItemData(cboCount.ListIndex) > 0 Then
                strCmd = strCmd & " print=" & cboCount.ItemData(cboCount.ListIndex)
            End If
        End If
        
        If chkNoSys.Value = 1 Then strCmd = strCmd & " sys=no"
        txtCmd.Text = strCmd
        
        '处理短文件路径
        strIn = GetShortName(gobjFile.GetParentFolderName(txtIn.Text)) & "\" & gobjFile.GetFileName(txtIn.Text)
        strOut = GetShortName(gobjFile.GetParentFolderName(txtOut.Text)) & "\" & gobjFile.GetFileName(txtOut.Text)
        txtCmd.Tag = Replace(Replace(txtCmd.Text, txtIn.Text, strIn), txtOut.Text, strOut)
    End If
End Sub

Private Function ChangeType(ByVal strFile As String, ByVal strS As String, ByVal strO As String) As String
    If UCase(strFile) Like "*" & UCase(strS) Then
        strFile = Left(strFile, Len(strFile) - Len(strS)) & strO
    End If
    ChangeType = strFile
End Function

Private Function GetTraceFile(Optional ByVal blnIsZlTraceFile As Boolean) As String
    '功能:获取服务器Trace文件，转储到本机
    '参数说明：  blnIsZlTraceFile -查找10046事件跟踪的日志文件
    
    Dim rstmp As adodb.Recordset, strSql As String
    Dim strFileName As String, strFilePath As String
    Dim objText As TextStream
    Dim intInstID As Integer
    
    On Error GoTo errh
    
    If gblnIsRac Then
        intInstID = Val(vsSession.TextMatrix(vsSession.Row, vsSession.ColIndex("Inst_ID")))
    Else
        intInstID = 1
    End If
    If blnIsZlTraceFile Then
    '如果blnIsZlTraceFile为True，说明需要从zlreginfo中取文件名
        strSql = "Select 内容 Filename, Value Filepath From zlRegInfo," & IIf(gblnIsRac, "G", "") & "V$parameter Where 项目 = 'TRACE文件' And Name = 'user_dump_dest'" & _
                IIf(gblnIsRac, " And Inst_ID = " & intInstID, "")
        Set rstmp = OpenSQLRecord(strSql, "DoCommand")
        
        If rstmp.RecordCount = 0 Then
            MsgBox "获取Trace文件失败，请先登录导航台获取Trace文件。": Exit Function
        End If
        
        If rstmp!FileName = "" Then
            MsgBox "获取Trace文件失败，请先登录导航台获取Trace文件。": Exit Function
        End If
        
    Else
    '根据SID、Serial#确定服务器Trace文件路径
        strSql = "Select a.Value FilePath, c.Instance_Name || '_ora_' || d.Spid || '.trc' FileName" & vbNewLine & _
                        "From (Select Value From " & IIf(gblnIsRac, "G", "") & "V$parameter Where Name = 'user_dump_dest'" & IIf(gblnIsRac, " And Inst_ID = " & intInstID, "") & ") A," & vbNewLine & _
                        "     (Select Instance_Name From " & IIf(gblnIsRac, "G", "") & "V$instance " & IIf(gblnIsRac, " Where Inst_ID = " & intInstID, "") & ") C," & vbNewLine & _
                        "     (Select Spid From " & IIf(gblnIsRac, "G", "") & "V$session S, " & IIf(gblnIsRac, "G", "") & "V$process P " & vbNewLine & _
                        "Where s.Paddr = p.Addr And s.Sid = [1] And s.Serial# = [2] " & IIf(gblnIsRac, " And S.Inst_ID = " & intInstID & " And P.Inst_ID = " & intInstID, "") & " ) D"
        Set rstmp = OpenSQLRecord(strSql, "DoCommand", vsSession.TextMatrix(vsSession.Row, vsSession.ColIndex("Sid")), vsSession.TextMatrix(vsSession.Row, vsSession.ColIndex("Serial#")))
    End If
                    
    strFileName = rstmp!FileName
    strFilePath = rstmp!FilePath
    

    
    '检查是否需要删除表
    strSql = "Select 1 From User_Tables Where table_name ='TRACEFILE'"
    Set rstmp = OpenSQLRecord(strSql, "DoCommand")
    If rstmp.RecordCount > 0 Then
        strSql = "Drop table tracefile Purge"
        gcnOracle.Execute strSql
        strSql = "Drop directory zltracefile"
        gcnOracle.Execute strSql
    End If
    
    '创建路径
    strSql = "create or replace directory zltracefile as '" & strFilePath & "'"
    gcnOracle.Execute strSql
    '创建外部表
    strSql = "create  table TRACEFILE" & vbNewLine & _
                    " (TEXT varchar2(4000))" & vbNewLine & _
                    " organization external (" & vbNewLine & _
                    " type oracle_loader default directory zltracefile" & vbNewLine & _
                    " access parameters (" & vbNewLine & _
                    " records delimited by newline nobadfile nodiscardfile  nologfile  FIELDS (text (1:4000) CHAR) )" & vbNewLine & _
                    " location('" & strFileName & "')" & vbNewLine & _
                    " ) reject limit Unlimited"
                    
    gcnOracle.Execute strSql

    '查询外部表
    strSql = "Select Text from TRACEFILE"
    Set rstmp = OpenSQLRecord(strSql, "DoCommand")
    
    '遍历外表表，将数据存至本地磁盘
    Set objText = gobjFile.CreateTextFile(gstrFilePath & "\" & strFileName, True)
    
    Do While Not rstmp.EOF
        objText.WriteLine "" & rstmp!Text
        rstmp.MoveNext
    Loop
    objText.Close

    GetTraceFile = gstrFilePath & "\" & strFileName
    Exit Function
    
errh:
    GetTraceFile = ""
    If InStr(Err.Description, "ORA-29400") > 0 Then
        MsgBox "服务器路径" & strFilePath & "下不存在Trace文件" & strFileName & "。" & vbNewLine & "注：如被跟踪会话没有执行SQL语句，将不会产生Trace文件。"
        Exit Function
    End If
    MsgBox Err.Description
    
    If 0 = 1 Then
        Resume
    End If
End Function

Public Sub InitTable(vsf As VSFlexGrid, strCol As String)
'功能: 初始化表头
    Dim arrHead As Variant
    Dim i As Long
    
    arrHead = Split(strCol, ";")
   
    With vsf
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        .Editable = flexEDNone
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Function GetConnection(intInstID As Integer) As adodb.Connection
    '功能: 根据实例号获取对应连接对象
    Dim cnResult As New adodb.Connection
    
    On Error Resume Next
    Set cnResult = mcolCon("_" & intInstID)
    
    Set GetConnection = cnResult
End Function
