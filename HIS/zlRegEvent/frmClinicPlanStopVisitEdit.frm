VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanStopVisitEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "停诊申请"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   Icon            =   "frmClinicPlanStopVisitEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5438.276
   ScaleMode       =   0  'User
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk部分停诊 
      Caption         =   "停诊部分号源"
      Height          =   180
      Left            =   3180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSignalSource 
      Height          =   1665
      Left            =   330
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5025
      _cx             =   8864
      _cy             =   2937
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   260
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicPlanStopVisitEdit.frx":000C
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
   Begin VB.CheckBox chkStopTime 
      Caption         =   "立即终止"
      Height          =   210
      Left            =   3270
      TabIndex        =   10
      Top             =   3915
      Value           =   1  'Checked
      Width           =   1065
   End
   Begin VB.Frame fraButton 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   60
      TabIndex        =   20
      Top             =   4320
      Width           =   5475
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   360
         Left            =   300
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   360
         Left            =   4140
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         Height          =   360
         Left            =   3090
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame fraSplit 
         Height          =   25
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   5745
      End
   End
   Begin VB.TextBox txtAuditTime 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3270
      TabIndex        =   9
      Top             =   3480
      Width           =   2085
   End
   Begin VB.TextBox txtAuditName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   900
      TabIndex        =   8
      Top             =   3480
      Width           =   1305
   End
   Begin VB.TextBox txtApplyTime 
      Enabled         =   0   'False
      Height          =   300
      Left            =   900
      TabIndex        =   7
      Top             =   3075
      Width           =   2085
   End
   Begin VB.ComboBox cboReason 
      Height          =   300
      Left            =   900
      TabIndex        =   6
      Top             =   2685
      Width           =   4455
   End
   Begin MSComCtl2.DTPicker dtpStartTime 
      Height          =   300
      Left            =   900
      TabIndex        =   4
      Top             =   2280
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   169869315
      CurrentDate     =   42362
   End
   Begin VB.ComboBox cboApplyName 
      Height          =   300
      Left            =   900
      TabIndex        =   1
      Top             =   120
      Width           =   2085
   End
   Begin MSComCtl2.DTPicker dtpEndTime 
      Height          =   300
      Left            =   3270
      TabIndex        =   5
      Top             =   2280
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   169869315
      CurrentDate     =   42367.9999884259
   End
   Begin MSComCtl2.DTPicker dtpStopTime 
      Height          =   300
      Left            =   900
      TabIndex        =   11
      Top             =   3870
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   169869315
      CurrentDate     =   42362
   End
   Begin VB.Label lblStopTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "终止时间"
      Height          =   180
      Left            =   150
      TabIndex        =   22
      Top             =   3930
      Width           =   720
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "停诊时间                        至"
      Height          =   180
      Left            =   150
      TabIndex        =   19
      Top             =   2340
      Width           =   3060
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "停诊原因"
      Height          =   180
      Left            =   150
      TabIndex        =   18
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label lblAuditName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "审批人"
      Height          =   180
      Left            =   330
      TabIndex        =   17
      Top             =   3540
      Width           =   540
   End
   Begin VB.Label lblAuditTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "审批时间"
      Height          =   180
      Left            =   2520
      TabIndex        =   16
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblApplyTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "申请时间"
      Height          =   180
      Left            =   150
      TabIndex        =   15
      Top             =   3135
      Width           =   720
   End
   Begin VB.Label lblApplyName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "申请人"
      Height          =   180
      Left            =   330
      TabIndex        =   0
      Top             =   180
      Width           =   540
   End
End
Attribute VB_Name = "frmClinicPlanStopVisitEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As m_BytFun '功能：1-申请，2-取消申请，3-审批，4-取消审批，4-终止安排
Private mlngID As Long '临床出诊停诊记录ID
Private mblnOk As Boolean
Private mlngModule As Long
Private mstrPrivs As String

Private Enum m_BytFun
    Fun_Applay = 1
    Fun_UnApplay = 2
    Fun_Audit = 3
    Fun_UnAudit = 4
    Fun_StopPlan = 5
End Enum
Private mrsDoctor As ADODB.Recordset
Private mbyt预约清单控制方式 As Byte
Private mbyt预约清单打印方式 As Byte
Private mstrDoctorName As String

Public Function ShowMe(frmParent As Form, ByVal lngModule As Long, _
    ByVal strPrivs As String, ByVal bytFun As Byte, _
    Optional ByVal lngID As Long, Optional ByRef strDoctorName As String) As Boolean
    '程序入口
    '入参：
    '   frmParent 父窗口
    '   bytFun 1-申请，2-取消申请，3-审批，4-取消审批
    '   lngID 临床出诊停诊记录ID
    mstrPrivs = strPrivs: mlngModule = lngModule
    mbytFun = bytFun: mlngID = lngID
    mstrDoctorName = ""
    
    Err = 0: On Error Resume Next
    If CheckDepend() = False Then Exit Function
    mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
    If mblnOk Then strDoctorName = mstrDoctorName
End Function

Private Sub cboApplyName_Click()
    Err = 0: On Error GoTo errHandle
    If cboApplyName.ListIndex = -1 Then Exit Sub
    Call LoadSignalSource(cboApplyName.ItemData(cboApplyName.ListIndex), vsfSignalSource.Tag)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboApplyName_GotFocus()
    zlControl.TxtSelAll cboApplyName
End Sub

Private Sub cboApplyName_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    
    Err = 0: On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    If cboApplyName.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If mrsDoctor Is Nothing Then Exit Sub
    If Trim(cboApplyName.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    If zlPersonSelect(Me, mlngModule, cboApplyName, mrsDoctor, Trim(cboApplyName.Text), True, "") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboApplyName_Validate(Cancel As Boolean)
    If cboApplyName.ListIndex < 0 Then cboApplyName.Text = ""
End Sub

Private Sub cboReason_GotFocus()
    zlControl.TxtSelAll cboReason
End Sub

Private Sub cboReason_KeyPress(KeyAscii As Integer)
    Dim strReason As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(cboReason.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    strReason = SearchStopVisitReason(Me, cboReason, Trim(cboReason.Text))
    If strReason = "" Then Exit Sub
    zlControl.CboLocate cboReason, strReason
    If cboReason.ListIndex = -1 Then cboReason.Text = strReason
End Sub

Private Sub chkStopTime_Click()
    dtpStopTime.Enabled = (chkStopTime.Value = vbUnchecked)
    If dtpStopTime.Visible And dtpStopTime.Enabled Then dtpStopTime.SetFocus
End Sub

Private Sub chkStopTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub chk部分停诊_Click()
    Err = 0: On Error GoTo errHandler
    With vsfSignalSource
        If chk部分停诊.Value = vbChecked Then
            .Editable = flexEDKbdMouse
            .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack
        Else
            .Editable = flexEDNone
            .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbGrayText
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str记录IDs As String, strTemp As String
    
    Err = 0: On Error GoTo errHandler
    If IsValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    mblnOk = True
    mstrDoctorName = NeedName(cboApplyName.Text)
    
    '检查是否要输出预约清单
    If mbytFun = Fun_Audit Then '审核
        strSQL = "Select a.ID as 记录ID" & vbNewLine & _
                " From 临床出诊记录 A, 临床出诊停诊记录 B, 病人挂号记录 C,临床出诊号源 D" & vbNewLine & _
                " Where ((a.替诊医生姓名 Is Null And a.医生id Is Not Null And a.医生姓名 = b.申请人)" & vbNewLine & _
                "       Or (a.替诊医生姓名 Is Not Null And a.替诊医生id Is Not Null And a.替诊医生姓名 = b.申请人))" & vbNewLine & _
                "       And b.Id = [1] And Not (a.开始时间 > b.终止时间 Or a.终止时间 < b.开始时间)" & vbNewLine & _
                "       And a.号源ID = d.ID And (b.停诊号码 Is Null Or Instr(','||b.停诊号码||',', ','||d.号码||',') > 0)" & vbNewLine & _
                "       And Exists (Select 1" & vbNewLine & _
                "                   From 临床出诊安排 C, 临床出诊表 D" & vbNewLine & _
                "                   Where c.出诊id = d.Id And c.Id = a.安排id And d.发布时间 Is Not Null)" & vbNewLine & _
                "        And a.Id = c.出诊记录id And c.记录状态 = 1" & vbNewLine & _
                "       And (c.记录性质 = 1 And c.发生时间 Between a.停诊开始时间 And a.停诊终止时间" & vbNewLine & _
                "           Or c.记录性质 = 2 And c.预约时间 Between a.停诊开始时间 And a.停诊终止时间)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        
        If rsTemp Is Nothing Then GoTo unloadForm:
        If rsTemp.EOF Then GoTo unloadForm:
    
        Do While Not rsTemp.EOF
            If InStr(strTemp & ",", "," & Nvl(rsTemp!记录ID) & ",") = 0 Then
                str记录IDs = str记录IDs & "," & Nvl(rsTemp!记录ID)
            End If
            rsTemp.MoveNext
        Loop
        If str记录IDs <> "" Then str记录IDs = Mid(str记录IDs, 2)
        
        If mbyt预约清单控制方式 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "出诊记录IDS=" & str记录IDs, 3)
        ElseIf mbyt预约清单控制方式 = 2 Then
            If MsgBox("当前医生停诊时间范围内存在预约或挂号病人，是否将预约清单输出到Excel中？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "出诊记录IDS=" & str记录IDs, 3)
            End If
        End If
        
        If mbyt预约清单打印方式 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "出诊记录IDS=" & str记录IDs, 2)
        ElseIf mbyt预约清单打印方式 = 2 Then
            If MsgBox("当前医生停诊时间范围内存在预约或挂号病人，你确定要打印预约清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "出诊记录IDS=" & str记录IDs, 2)
            End If
        End If
    End If
    
unloadForm:
    mblnOk = True
    Unload Me
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String
    Dim strNOs As String, i As Integer
    
    Err = 0: On Error GoTo errHandler
    '1-申请，2-取消申请，3-审批，4-取消审批
    If mbytFun = Fun_Applay Then
        If chk部分停诊.Value = vbChecked Then
            With vsfSignalSource
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, .ColIndex("选择")) = 1 Then
                        strNOs = strNOs & "," & .TextMatrix(i, .ColIndex("号码"))
                    End If
                Next
                If strNOs <> "" Then strNOs = Mid(strNOs, 2)
            End With
        End If
        
        'Zl_临床出诊停诊_Apply(
        strSQL = "Zl_临床出诊停诊_Apply("
        '操作类型_In Number,--0-申请，else-取消申请
        strSQL = strSQL & "" & 0 & ","
        'Id_In       临床出诊停诊记录.Id%Type,
        strSQL = strSQL & "" & mlngID & ","
        '停诊号码_In 临床出诊停诊记录.停诊号码%type := Null,
        strSQL = strSQL & "'" & strNOs & "',"
        '开始时间_In 临床出诊停诊记录.开始时间%Type := Null,
        strSQL = strSQL & "" & ZDate(dtpStartTime.Value) & ","
        '终止时间_In 临床出诊停诊记录.终止时间%Type := Null,
        strSQL = strSQL & "" & ZDate(dtpEndTime.Value) & ","
        '停诊原因_In 临床出诊停诊记录.停诊原因%Type := Null,
        strSQL = strSQL & "'" & NeedName(cboReason.Text) & "',"
        '申请人_In   临床出诊停诊记录.申请人%Type := Null,
        strSQL = strSQL & "'" & NeedName(cboApplyName.Text) & "',"
        '申请时间_In 临床出诊停诊记录.申请时间%Type := Null,
        strSQL = strSQL & "" & ZDate(txtApplyTime.Text) & ","
        '登记人_In   临床出诊停诊记录.登记人%Type := Null
        strSQL = strSQL & "'" & UserInfo.姓名 & "')"
    ElseIf mbytFun = Fun_UnApplay Then
        'Zl_临床出诊停诊_Apply(
        strSQL = "Zl_临床出诊停诊_Apply("
        '操作类型_In Number,--0-申请，else-取消申请
        strSQL = strSQL & "" & 1 & ","
        'Id_In       临床出诊停诊记录.Id%Type,
        strSQL = strSQL & "" & mlngID & ")"
        '停诊号码_In 临床出诊停诊记录.停诊号码%type := Null,
        '开始时间_In 临床出诊停诊记录.开始时间%Type := Null,
        '终止时间_In 临床出诊停诊记录.终止时间%Type := Null,
        '停诊原因_In 临床出诊停诊记录.停诊原因%Type := Null,
        '申请人_In   临床出诊停诊记录.申请人%Type := Null,
        '申请时间_In 临床出诊停诊记录.申请时间%Type := Null,
        '登记人_In   临床出诊停诊记录.登记人%Type := Null
    ElseIf mbytFun = Fun_Audit Then
        'Zl_临床出诊停诊_Audit(
        strSQL = "Zl_临床出诊停诊_Audit("
        '操作类型_In Number,--1-审核，2-取消审核
        strSQL = strSQL & "" & 1 & ","
        'Id_In       临床出诊停诊记录.Id%Type,
        strSQL = strSQL & "" & mlngID & ","
        '审批人_In   临床出诊停诊记录.审批人%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '审批时间_In 临床出诊停诊记录.审批时间%Type := Null
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    ElseIf mbytFun = Fun_UnAudit Then
        'Zl_临床出诊停诊_Audit(
        strSQL = "Zl_临床出诊停诊_Audit("
        '操作类型_In Number,--1-审核，2-取消审核
        strSQL = strSQL & "" & 2 & ","
        'Id_In       临床出诊停诊记录.Id%Type,
        strSQL = strSQL & "" & mlngID & ")"
        '审批人_In   临床出诊停诊记录.审批人%Type := Null,
        '审批时间_In 临床出诊停诊记录.审批时间%Type := Null
    ElseIf mbytFun = Fun_StopPlan Then
        'Zl_临床出诊停诊_Stop
        strSQL = "Zl_临床出诊停诊_Stop("
        '  Id_In       临床出诊停诊记录.Id%Type,
        strSQL = strSQL & "" & mlngID & ","
        '  终止人_In   临床出诊停诊记录.取消人%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  终止时间_In 临床出诊停诊记录.失效时间%Type := Null--Null-立即终止，其它-具体的终止时间
        If chkStopTime.Value = vbChecked Then
            strSQL = strSQL & "" & "NULL" & ")"
        Else
            strSQL = strSQL & "" & ZDate(dtpStopTime.Value) & ")"
        End If
    End If
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub dtpEndTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub dtpStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub dtpStartTime_LostFocus()
    dtpEndTime.Value = Format(dtpStartTime.Value, "yyyy-mm-dd 23:59:59")
End Sub

Private Sub dtpStopTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    Me.Caption = Choose(mbytFun, "停诊申请", "取消申请", "停诊审批", "取消审批", "终止安排")
    
    If mbytFun = Fun_Audit Then
        mbyt预约清单控制方式 = Val(zlDatabase.GetPara("预约清单控制方式", glngSys, mlngModule, "0"))
        mbyt预约清单打印方式 = Val(zlDatabase.GetPara("预约清单打印方式", glngSys, mlngModule, "0"))
    End If
    
    If mbytFun = Fun_Applay Then
        If Not zlStr.IsHavePrivs(mstrPrivs, "允许代他人停诊申请") Then
            cboApplyName.Enabled = False
        End If
    Else
        cboApplyName.Enabled = False
        dtpStartTime.Enabled = False
        dtpEndTime.Enabled = False
        cboReason.Enabled = False
    End If
    
    lblAuditTime.Visible = (mbytFun = Fun_UnAudit Or mbytFun = Fun_StopPlan)
    txtAuditTime.Visible = (mbytFun = Fun_UnAudit Or mbytFun = Fun_StopPlan)
    lblAuditName.Visible = (mbytFun = Fun_UnAudit Or mbytFun = Fun_StopPlan)
    txtAuditName.Visible = (mbytFun = Fun_UnAudit Or mbytFun = Fun_StopPlan)
    
    lblStopTime.Visible = (mbytFun = Fun_StopPlan)
    dtpStopTime.Visible = (mbytFun = Fun_StopPlan)
    chkStopTime.Visible = (mbytFun = Fun_StopPlan)
    
    If mbytFun = Fun_UnAudit Then
        fraButton.Top = txtAuditTime.Top + txtAuditTime.Height + 150
    ElseIf mbytFun = Fun_StopPlan Then
        fraButton.Top = dtpStopTime.Top + dtpStopTime.Height + 150
    Else
        fraButton.Top = txtApplyTime.Top + txtApplyTime.Height + 150
    End If
    Me.Height = fraButton.Top + fraButton.Height + 280
    
    Call SetEnabledBackColor(Me.Controls)
    If InitData() = False Then Unload Me: Exit Sub
    If LoadData(mbytFun) = False Then Unload Me: Exit Sub
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strPersons As String
    
    Err = 0: On Error GoTo errHandler
    If zlStr.IsHavePrivs(mstrPrivs, "允许代他人停诊申请") Then
        Set mrsDoctor = GetDoctor(, "编号")
        cboApplyName.Clear
        Do While Not mrsDoctor.EOF
            If InStr("," & strPersons & ",", "," & Nvl(mrsDoctor!ID) & ",") = 0 Then
                strPersons = strPersons & "," & Nvl(mrsDoctor!ID)
                cboApplyName.AddItem Nvl(mrsDoctor!简码) & "-" & Nvl(mrsDoctor!姓名)
                cboApplyName.ItemData(cboApplyName.NewIndex) = Nvl(mrsDoctor!ID)
            End If
            mrsDoctor.MoveNext
        Loop
    Else
        cboApplyName.Clear
        cboApplyName.AddItem UserInfo.简码 & "-" & UserInfo.姓名
        cboApplyName.ItemData(cboApplyName.NewIndex) = UserInfo.ID
    End If
    
    strSQL = "Select 编码, 名称, 简码, Nvl(缺省标志, 0) As 缺省标志" & vbNewLine & _
            " From 常用停诊原因"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboReason.Clear
    Do While Not rsTemp.EOF
        cboReason.AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
        If Val(Nvl(rsTemp!缺省标志)) = 1 Then cboReason.ListIndex = cboReason.NewIndex
        rsTemp.MoveNext
    Loop
    InitData = True
    Exit Function
errHandler:
    vsfSignalSource.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadSignalSource(ByVal lng医生ID As Long, _
    Optional ByVal str停诊号码 As String) As String
    '获取医生的所有有效号源
    '入参：
    '   str停诊号码 已申请的号码，多个用逗号分隔
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Long, strWhere As String
    
    Err = 0: On Error GoTo errHandler
    If str停诊号码 = "" Then
        strWhere = _
            " And Nvl(a.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate" & vbNewLine & _
            " And Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate"
    Else
        strWhere = " And a.号码 In(Select /*+cardinality(j,10)*/Column_Value From Table(f_Str2list([2])) J)"
    End If
    strSQL = _
        " Select a.号码, b.名称 As 科室, c.名称 As 收费项目" & vbNewLine & _
        " From 临床出诊号源 A, 部门表 B, 收费项目目录 C" & vbNewLine & _
        " Where a.科室id = b.Id And a.项目id = c.Id And a.医生id = [1]" & vbNewLine & _
                strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医生ID, str停诊号码)
    
    With vsfSignalSource
        .Redraw = flexRDNone
        .Clear 1
        .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = 1 '标记为选择
            .TextMatrix(lngRow, .ColIndex("号码")) = Nvl(rsTemp!号码)
            .TextMatrix(lngRow, .ColIndex("科室")) = Nvl(rsTemp!科室)
            .TextMatrix(lngRow, .ColIndex("收费项目")) = Nvl(rsTemp!收费项目)
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = IIf(chk部分停诊.Value = vbChecked, vbBlack, vbGrayText)
        .Redraw = flexRDBuffered
    End With
    chk部分停诊.Visible = mbytFun = Fun_Applay And rsTemp.RecordCount > 0
    LoadSignalSource = True
    Exit Function
errHandler:
    vsfSignalSource.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadData(ByVal bytFun As Byte) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dtNow As Date
    
    Err = 0: On Error GoTo errHandler
    dtNow = zlDatabase.Currentdate
    If bytFun = Fun_Applay Then
        zlControl.CboSetText cboApplyName, UserInfo.姓名
        dtpStartTime.MinDate = Format(dtNow, "yyyy-mm-dd hh:mm:ss")
        dtpStartTime.Value = Format(dtNow + 1, "yyyy-mm-dd 00:00:00")
        dtpEndTime.MinDate = Format(dtNow, "yyyy-mm-dd hh:mm:ss")
        dtpEndTime.Value = Format(dtNow + 1, "yyyy-mm-dd 23:59:59")
        txtApplyTime.Text = Format(dtNow, "yyyy-mm-dd hh:mm:ss")
        LoadData = True: Exit Function
    End If
    
    strSQL = "Select 停诊原因, 开始时间, 终止时间, 申请人, 申请时间, 审批人, 审批时间, 登记人, 停诊号码" & vbNewLine & _
            " From 临床出诊停诊记录" & vbNewLine & _
            " Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If rsTemp.EOF Then
        MsgBox "记录不存在，可能已被他人" & IIf(bytFun = Fun_UnAudit, "取消审批", "取消申请或审批") & "！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    vsfSignalSource.Tag = Nvl(rsTemp!停诊号码)
    
    zlControl.CboSetText cboApplyName, Nvl(rsTemp!申请人)
    If cboApplyName.ListIndex = -1 Then cboApplyName.Text = Nvl(rsTemp!申请人)
    cboApplyName.Tag = Nvl(rsTemp!登记人) '存储登记人，用于检查
    
    dtpStartTime.Value = Format(Nvl(rsTemp!开始时间), "yyyy-mm-dd hh:mm:ss")
    dtpEndTime.Value = Format(Nvl(rsTemp!终止时间), "yyyy-mm-dd hh:mm:ss")
    
    zlControl.CboSetText cboReason, Nvl(rsTemp!停诊原因)
    If cboReason.ListIndex = -1 Then cboReason.Text = Nvl(rsTemp!停诊原因)
    
    txtApplyTime.Text = Format(Nvl(rsTemp!申请时间), "yyyy-mm-dd hh:mm:ss")
    If bytFun = Fun_UnApplay Then LoadData = True: Exit Function
    
    txtAuditName.Text = Nvl(rsTemp!审批人)
    txtAuditTime.Text = Format(Nvl(rsTemp!审批时间), "yyyy-mm-dd hh:mm:ss")
    
    dtpStopTime.Value = Format(dtNow, "yyyy-mm-dd hh:mm:ss")
    dtpStopTime.MaxDate = Format(Nvl(rsTemp!终止时间), "yyyy-mm-dd hh:mm:ss")
    LoadData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mrsDoctor = Nothing
End Sub

Private Function IsValied() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngCount As Long, i As Integer, strStopNOs As String
    Dim str已停号码 As String, varData As Variant
    
    Err = 0: On Error GoTo errHandle
    If mbytFun = Fun_Applay Then
        If zlControl.FormCheckInput(Me) = False Then Exit Function
        If cboApplyName.ListIndex < 0 Or cboApplyName.Text = "" Then
            MsgBox "请选择申请人！", vbInformation, gstrSysName
            If cboApplyName.Visible And cboApplyName.Enabled Then cboApplyName.SetFocus
            Exit Function
        End If
        
        If chk部分停诊.Value = vbChecked Then
            With vsfSignalSource
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, .ColIndex("选择")) = 1 Then
                        lngCount = lngCount + 1
                        strStopNOs = strStopNOs & "," & .TextMatrix(i, .ColIndex("号码"))
                    End If
                Next
                If strStopNOs <> "" Then strStopNOs = Mid(strStopNOs, 2)
                If .Rows > 1 And lngCount = 0 Then
                    MsgBox "请选择停诊号码！", vbInformation, gstrSysName
                    If vsfSignalSource.Visible And vsfSignalSource.Enabled Then vsfSignalSource.SetFocus
                    Exit Function
                End If
                If lngCount > 100 Then
                    MsgBox "每一次申请的停诊号码不能超过100个，请分多次申请！", vbInformation, gstrSysName
                    If vsfSignalSource.Visible And vsfSignalSource.Enabled Then vsfSignalSource.SetFocus
                    Exit Function
                End If
            End With
        End If
        
        If DateDiff("s", dtpStartTime.Value, zlDatabase.Currentdate) >= 0 Then
            MsgBox "停诊时间的开始时间必须大于当前时间！", vbInformation, gstrSysName
            If dtpEndTime.Visible And dtpEndTime.Enabled Then dtpEndTime.SetFocus
            Exit Function
        End If
        
        If DateDiff("s", dtpStartTime.Value, dtpEndTime.Value) <= 0 Then
            MsgBox "停诊时间的结束时间必须大于开始时间！", vbInformation, gstrSysName
            If dtpEndTime.Visible And dtpEndTime.Enabled Then dtpEndTime.SetFocus
            Exit Function
        End If
        
        If zlControl.TxtCheckInput(cboReason, "停诊原因", 50, False) = False Then Exit Function
        
        strSQL = "Select 1 From 临床出诊停诊记录" & vbNewLine & _
                " Where 记录id Is Null And Not (开始时间 > [2] Or Nvl(失效时间, 终止时间) < [1])" & vbNewLine & _
                "       And 申请人 = [3] And 停诊号码 Is Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtpStartTime.Value, dtpEndTime.Value, NeedName(cboApplyName.Text))
        If Not rsTemp.EOF Then
            MsgBox "当前停诊时间与已申请停诊时间范围存在重叠，请检查！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 停诊号码 From 临床出诊停诊记录" & vbNewLine & _
                " Where 记录id Is Null And Not (开始时间 > [2] Or Nvl(失效时间, 终止时间) < [1])" & vbNewLine & _
                "       And 申请人 = [3] And 停诊号码 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtpStartTime.Value, dtpEndTime.Value, NeedName(cboApplyName.Text))
        Do While Not rsTemp.EOF
            If strStopNOs = "" Then
                str已停号码 = str已停号码 & "," & Nvl(rsTemp!停诊号码)
            Else
                varData = Split(strStopNOs, ",")
                For i = 0 To UBound(varData)
                    If InStr("," & Nvl(rsTemp!停诊号码) & ",", "," & varData(i) & ",") > 0 Then
                        str已停号码 = str已停号码 & "," & varData(i)
                    End If
                Next
            End If
            rsTemp.MoveNext
        Loop
        If str已停号码 <> "" Then
            str已停号码 = Mid(str已停号码, 2)
            MsgBox "号码(" & str已停号码 & ")当前停诊时间与已申请停诊时间范围存在重叠，请检查！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mbytFun = Fun_UnApplay Then
        If Not zlStr.IsHavePrivs(mstrPrivs, "允许代他人停诊申请") Then
            If Not (NeedName(cboApplyName.Text) = UserInfo.姓名 Or cboApplyName.Tag = UserInfo.姓名) Then
                MsgBox "你只能删除自己的申请，不能删除他人的申请！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    ElseIf mbytFun = Fun_StopPlan Then
        If chkStopTime.Value = vbUnchecked Then
            If DateDiff("s", dtpStopTime.Value, zlDatabase.Currentdate) >= 0 Then
                MsgBox "终止时间必须大于当前时间！", vbInformation, gstrSysName
                If dtpStopTime.Visible And dtpStopTime.Enabled Then dtpStopTime.SetFocus
                Exit Function
            End If
        End If
    End If
    If CheckDepend() = False Then Exit Function
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function CheckDepend() As Boolean
    '功能:检查数据
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    If mbytFun = Fun_UnApplay Then
        strSQL = "Select 1 From 临床出诊停诊记录 Where ID = [1] And 审批人 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "该申请已被审批，不能取消申请。", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mbytFun = Fun_Audit Then
        strSQL = "Select 1 From 临床出诊停诊记录 Where ID = [1] And 审批人 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "该申请已被审批，不能再次审批！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mbytFun = Fun_UnAudit Then
        strSQL = "Select 1 From 临床出诊停诊记录 Where ID = [1] And 终止时间 < Sysdate"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "该停诊安排已失效，不能取消审批！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From 临床出诊停诊记录 Where ID = [1] And 失效时间 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "该停诊安排已被终止，不能取消审批！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From 临床出诊记录 A, 临床出诊停诊记录 B, 病人服务信息记录 C" & vbNewLine & _
                " Where Nvl(a.替诊医生姓名, a.医生姓名) = b.申请人 And Nvl(a.替诊医生id, a.医生id) Is Not Null" & vbNewLine & _
                "       And a.Id = c.记录id And (a.开始时间 Between b.开始时间 And b.终止时间 Or a.终止时间 Between b.开始时间 And b.终止时间)" & vbNewLine & _
                "       And c.处理人 Is Not Null And b.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "该停诊安排的部分停诊信息已被处理，不能取消审批！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mbytFun = Fun_StopPlan Then
        strSQL = "Select 1 From 临床出诊停诊记录 Where ID = [1] And 终止时间 < Sysdate"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "该停诊安排已失效，不能终止！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From 临床出诊停诊记录 Where ID = [1] And 失效时间 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If Not rsTemp.EOF Then
            MsgBox "该停诊安排已被终止，不能再终止！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From 临床出诊停诊记录 Where ID = [1] And 审批人 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
        If rsTemp.EOF Then
            MsgBox "该停诊安排还未审批，不能终止！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDepend = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsfSignalSource_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfSignalSource.ColIndex("选择") Then Exit Sub
    Cancel = True
End Sub
