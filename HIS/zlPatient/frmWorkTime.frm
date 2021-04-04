VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印缴款书"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmWorkTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboTimes 
      Height          =   300
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   990
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "缴款书(&P)"
      Height          =   350
      Left            =   405
      TabIndex        =   5
      Top             =   2880
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   2610
      Width           =   6555
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   3960
      TabIndex        =   6
      Top             =   2895
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   2865
      TabIndex        =   4
      Top             =   2895
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Index           =   0
      Left            =   -270
      TabIndex        =   11
      Top             =   1440
      Width           =   6555
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   1155
      TabIndex        =   3
      Top             =   2175
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
      Format          =   134086659
      CurrentDate     =   38175
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1155
      TabIndex        =   2
      Top             =   1710
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
      Format          =   134086659
      CurrentDate     =   38175
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Left            =   1155
      TabIndex        =   0
      Top             =   990
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   134086659
      CurrentDate     =   2
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3720
      TabIndex        =   13
      Top             =   2280
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "终止时间"
      Height          =   180
      Left            =   375
      TabIndex        =   10
      Top             =   2235
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间"
      Height          =   180
      Left            =   375
      TabIndex        =   9
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上班日期"
      Height          =   180
      Left            =   375
      TabIndex        =   8
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "在打印缴款书之前，请先指定上班日期,根据实际工作时间填写并保存该天的上班开始时间和结束时间(夜班可能跨天)。"
      Height          =   540
      Left            =   975
      TabIndex        =   7
      Top             =   195
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frmWorkTime.frx":000C
      Top             =   390
      Width           =   480
   End
End
Attribute VB_Name = "frmWorkTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte '1-预交款,2-结帐,3-收费,4-挂号,5-就诊卡,6-消费卡缴款
Private mrsTimes As ADODB.Recordset '当前缴款日期的缴款次数

Public Sub ShowMe(frmParent As Object, bytType As Byte)
    mbytType = bytType
    
    On Error Resume Next
    Me.Show 1, frmParent
End Sub


Private Sub cboTimes_Click()
    
    If cboTimes.Visible Then
        Call SetTimeRange
        cboTimes.Tag = "Click"
    End If
End Sub

Private Sub cboTimes_Validate(Cancel As Boolean)
    If cboTimes.Tag = "Click" Then
        cboTimes.Tag = ""
    Else
        Call SetTimeRange
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim strReport As String
    
    If cmdSave.Enabled Then
        MsgBox "在打印缴款书之前，请先保存上班开始终止时间。", vbInformation, gstrSysName
        cmdSave.SetFocus: Exit Sub
    End If
    
    Select Case mbytType
        Case 1 '预交款缴款书
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1103_1"
        Case 2 '结帐缴款书
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1137_1"
        Case 3 '收费缴款书
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1121_1"
        Case 4 '挂号缴款书
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1111_1"
        Case 5 '就诊卡缴款书
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1102_1"
        Case 6 '消费卡缴款书
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1503_2"
    End Select
    
    Call ReportOpen(gcnOracle, glngSys, strReport, Me, _
        "开始时间=" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss"), _
        "结束时间=" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss"), _
        "操作员=" & UserInfo.姓名)
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String, strOldDate As String
    
    If cboTimes.ListIndex < 0 Then MsgBox "请选择缴款次数!", vbInformation, App.ProductName
    If dtpBegin.Value >= dtpEnd.Value Then
        MsgBox "开始时间应该小于终止时间。", vbInformation, gstrSysName
        If dtpBegin.Enabled Then dtpBegin.SetFocus
        Exit Sub
    End If
    
    If InStr(";" & gstrPrivs & ";", ";修改上班时间;") = 0 Then
        If MsgBox("保存当前设置的上班时间后将不允许修改,你确定要继续吗?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    If cboTimes.ItemData(cboTimes.ListIndex) <> 0 Then '修改缴款书
        mrsTimes.Filter = "次数=" & cboTimes.ItemData(cboTimes.ListIndex)
        strOldDate = "To_Date('" & Format(mrsTimes!开始时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    Else
        strOldDate = "Null"
    End If
    
    
    strSQL = "ZL_收费清点记录_Insert(To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','YYYY-MM-DD')," & mbytType & "," & _
        "To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
        "To_Date('" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & strOldDate & ")"
    On Error GoTo errH
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    '刷新数据，以便没有权限修改的，重读固定的上班时间
    Call dtpDate_Change
    
    cmdSave.Enabled = False
    cmdPrint.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpBegin_Change()
    cmdSave.Enabled = True
    lblMessage.Caption = ""
    If dtpEnd.Value < dtpBegin.Value Then
        lblMessage.Top = Label3.Top
        lblMessage.Caption = "不能比结束时间大!"
        cmdSave.Enabled = False
    End If
End Sub

Private Sub SetTimeRange()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, curDate As Date, DatBegin As Date, strBDate As String, strEDate As String
    Dim blnHave As Boolean, dat上次终止 As Date, dat下次开始 As Date
    
    On Error GoTo errH
    '后面设置了禁用,重选日期时要恢复
    dtpBegin.Enabled = True
    dtpEnd.Enabled = True
    
    '一天有多次缴款
    If cboTimes.ListCount > 1 Then
        If cboTimes.ItemData(cboTimes.ListIndex) = 0 Then  '新缴款
            mrsTimes.Filter = "次数=" & cboTimes.ItemData(cboTimes.ListCount - 1)
            DatBegin = mrsTimes!开始时间
            
            strBDate = " And 日期=[1] And 开始时间=[4]"
            strEDate = " And 日期>[1]"
        Else
            mrsTimes.Filter = "次数=" & cboTimes.ItemData(cboTimes.ListIndex)
            DatBegin = mrsTimes!开始时间
            
            If cboTimes.ItemData(cboTimes.ListIndex) = 1 Then  '当天第1次缴款
                strBDate = " And 日期<[1]"
            Else
                strBDate = " And 日期=[1] And 开始时间<[4]"
            End If
            
            If cboTimes.ListIndex = cboTimes.ListCount - 1 Then '当天最后一次缴款
                strEDate = " And 日期>[1]"
            Else
                strEDate = " And 日期=[1] And 开始时间>[4]"
            End If
        End If
    Else
        strBDate = " And 日期<[1]"
        strEDate = " And 日期>[1]"
    End If
    
    
    '设置该天可以修改的时间范围
    '------------------------------------------------------------------------------------------
    curDate = zldatabase.Currentdate
    dtpBegin.MinDate = "1601-01-01": dtpBegin.MaxDate = "9999-12-31"
    dtpEnd.MinDate = "1601-01-01": dtpEnd.MaxDate = "9999-12-31"
    
    '开始时间：应大于上次终止时间,或前一天内
    strSQL = "Select Max(终止时间) as 上次终止 From 收费清点记录 Where 收款员=[2] And 性质=[3]" & strBDate
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.姓名, mbytType, DatBegin)
    
    blnHave = False
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!上次终止) Then blnHave = True
    End If
    If blnHave Then
        dat上次终止 = rsTmp!上次终止
        dtpBegin.MinDate = DateAdd("s", 1, rsTmp!上次终止)
    Else
        dtpBegin.MinDate = Int(dtpDate.Value - 1)
    End If
    
    '终止时间：应小于下次开始时间,或后一天内(不超过当前时间)
    strSQL = "Select Min(开始时间) as 下次开始 From 收费清点记录 Where 收款员=[2] And 性质=[3]" & strEDate
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.姓名, mbytType, DatBegin)
    
    blnHave = False
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!下次开始) Then blnHave = True
    End If
    If blnHave Then
        dat下次开始 = rsTmp!下次开始
        dtpEnd.MaxDate = DateAdd("s", -1, rsTmp!下次开始)
    Else
        dtpEnd.MaxDate = curDate
    End If
    dtpBegin.MaxDate = dtpEnd.MaxDate
    dtpEnd.MinDate = dtpBegin.MinDate
    
    
        
    '设置缺省上班时间范围
    '------------------------------------------------------------------------------------------
    '缴款重打或修改
    If cboTimes.ItemData(cboTimes.ListIndex) > 0 Then
        If cboTimes.ListCount = 1 Then
            strSQL = "Select 开始时间,终止时间 From 收费清点记录 Where 收款员=[2] And 性质=[3] And 日期=[1] Order by 终止时间 Desc"
        Else
            strSQL = "Select 开始时间,终止时间 From 收费清点记录 Where 收款员=[2] And 性质=[3] And 日期=[1] And 开始时间=[4] Order by 终止时间 Desc"
        End If
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.姓名, mbytType, DatBegin)
    
        cmdSave.Enabled = False
        
        If rsTmp.RecordCount = 1 Then
            dtpBegin.Value = rsTmp!开始时间
            dtpEnd.Value = rsTmp!终止时间
        Else
            '如果一天缴款多次,则缺省继续连续缴款
            dtpBegin.Value = DateAdd("s", 1, rsTmp!终止时间)
            '终止时间缺省为下次开始时间-1s,或为该天最后时间
            If dat下次开始 <> CDate(0) Then
                dtpEnd.Value = DateAdd("s", -1, dat下次开始)
            Else
                If Format(dtpDate.Value, "yyyy-MM-dd 23:59:59") <= curDate And dtpEnd.MinDate <= dtpDate.Value Then
                    dtpEnd.Value = Format(dtpDate.Value, "yyyy-MM-dd 23:59:59")
                Else
                    dtpEnd.Value = curDate
                End If
            End If
        End If
        
        '修改上班时间权限,指是否有修改上班日期那一天的开始时间和结束时间的权限,而不是上班时间本身，因为允许一次打印多天的缴款书
        If InStr(";" & gstrPrivs & ";", ";修改上班时间;") = 0 Then
            dtpBegin.Enabled = False
            dtpEnd.Enabled = False
        End If
    Else
        '新的缴款
        cmdSave.Enabled = True
        
        '开始时间缺省上次终止时间+1s
        If dat上次终止 <> CDate(0) Then
            dtpBegin.Value = DateAdd("s", 1, dat上次终止)
            If InStr(";" & gstrPrivs & ";", ";修改上班时间;") = 0 Then dtpBegin.Enabled = False
        Else
            dtpBegin.Value = Int(dtpDate.Value) '允许修改至前一天,缺省为当天
        End If
        
        '终止时间缺省为下次开始时间-1s,或为该天最后时间
        If dat下次开始 <> CDate(0) Then
            dtpEnd.Value = DateAdd("s", -1, dat下次开始)
        Else
            If Format(dtpDate.Value, "yyyy-MM-dd 23:59:59") <= curDate And dtpEnd.MinDate <= dtpDate.Value Then
                dtpEnd.Value = Format(dtpDate.Value, "yyyy-MM-dd 23:59:59")
            Else
                dtpEnd.Value = curDate
            End If
        End If
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTimes(datThis As Date)
'加载缴款次数
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    
    '注意:顺序影响相关取数功能
    strSQL = "Select Rownum 次数, 开始时间 From (Select 开始时间 From 收费清点记录 Where 日期 = [1] And 收款员=[2] And 性质=[3] Order By 开始时间)"
    On Error GoTo errH
    Set mrsTimes = zldatabase.OpenSQLRecord(strSQL, Me.Caption, datThis, UserInfo.姓名, mbytType)
        
    '当天的最大终止时间与当天之后最近一次缴款日期的最小开始时间之间无间隔时，不允许再新增缴款
    strSQL = "Select 1" & vbNewLine & _
            "From (Select Min(开始时间) 开始时间" & vbNewLine & _
            "       From 收费清点记录" & vbNewLine & _
            "       Where 日期 > [1] And 收款员 = [2] And 性质 = [3]) A," & vbNewLine & _
            "     (Select Max(终止时间) 终止时间" & vbNewLine & _
            "       From 收费清点记录" & vbNewLine & _
            "       Where 日期 = [1] And 收款员 = [2] And 性质 = [3]) B" & vbNewLine & _
            "Where A.开始时间 = B.终止时间 + 1 / 60 / 60 / 24"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, datThis, UserInfo.姓名, mbytType)
        
    With cboTimes
        .Clear
        If rsTmp.RecordCount = 0 Then .AddItem "新增缴款": .ItemData(.NewIndex) = 0
        For i = 1 To mrsTimes.RecordCount
            .AddItem "第" & i & "次缴款": .ItemData(.NewIndex) = i
        Next
    End With
    Call zlControl.CboSetIndex(cboTimes.hWnd, 0)
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpDate_Change()
    If dtpDate.Tag = "SelfChange" Then Exit Sub
    
    Call ValidateDate
    
    Call LoadTimes(dtpDate.Value)
    Call SetTimeRange
End Sub

Private Sub ValidateDate()
    Dim rsTmp As ADODB.Recordset, blnDo As Boolean
    Dim strSQL As String
        
    On Error GoTo errH
    
    '检查输入的上班时间是否已包含在已存在的缴款时间段内
    '如果是一天多次缴款,保存时再检查
    '例如:假设已存在以下[收费清点记录],此时输入2006-12-11到2006-12-13之间的日期都是不允许的,自动改为2006-12-10或2006-12-14
    '    日期    收款员  性质    开始时间    终止时间
    '1   2006-12-10  曹丽华  3   2006-12-10  2006-12-13 11:59:59
    '2   2006-12-14  曹丽华  3   2006-12-13 12:00:01 2006-12-14 08:00:00
    strSQL = "Select 日期 From 收费清点记录 Where 日期<>[1] And [1] Between 开始时间 And 终止时间 And 收款员=[2] And 性质=[3]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.姓名, mbytType)
    blnDo = rsTmp.RecordCount > 0
    If Not blnDo Then
        '当天之前的最大终止时间与当天之后最近一次缴款日期的最小开始时间之间无间隔时，不允许选择
        strSQL = "Select 1" & vbNewLine & _
                "From (Select Min(开始时间) 开始时间" & vbNewLine & _
                "       From 收费清点记录" & vbNewLine & _
                "       Where 日期 > [1] And 收款员 = [2] And 性质 = [3]) A," & vbNewLine & _
                "     (Select Max(终止时间) 终止时间" & vbNewLine & _
                "       From 收费清点记录" & vbNewLine & _
                "       Where 日期 < [1] And 收款员 = [2] And 性质 = [3]) B" & vbNewLine & _
                "Where A.开始时间 = B.终止时间 + 1 / 60 / 60 / 24"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.姓名, mbytType)
        blnDo = rsTmp.RecordCount > 0
    End If
    
    If blnDo Then
        MsgBox "上班时间:" & Format(dtpDate.Value, "YYYY-MM-DD") & "已包含在" & Format(rsTmp!日期, "YYYY-MM-DD") & "的缴款时间范围内!", vbInformation, gstrSysName
        dtpDate.Tag = "SelfChange"
        dtpDate.Value = rsTmp!日期
        dtpDate.Tag = ""
        If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
        '问题:38829
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpEnd_Change()
    cmdSave.Enabled = True
    lblMessage.Caption = ""
    If dtpEnd.Value < dtpBegin.Value Then
        lblMessage.Top = Label4.Top
        lblMessage.Caption = "不能比开始时间小!"
        cmdSave.Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    dtpDate.Value = Int(zldatabase.Currentdate)
    dtpDate.MaxDate = Int(dtpDate.Value)
    Call dtpDate_Change
End Sub

