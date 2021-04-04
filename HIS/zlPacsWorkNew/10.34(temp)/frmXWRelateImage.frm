VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXWRelateImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关联影像"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   Icon            =   "frmXWRelateImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdRepair 
      Caption         =   "修复关联状态(&R)"
      Height          =   350
      Left            =   7275
      TabIndex        =   20
      Top             =   6120
      Width           =   1845
   End
   Begin VB.Frame frmFilter 
      Caption         =   "过滤条件"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   11655
      Begin VB.ComboBox cboModality 
         Height          =   300
         ItemData        =   "frmXWRelateImage.frx":038A
         Left            =   960
         List            =   "frmXWRelateImage.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1057
         Width           =   1600
      End
      Begin VB.TextBox txtStudyNo 
         Height          =   300
         Left            =   5280
         TabIndex        =   14
         Top             =   1057
         Width           =   1600
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   9840
         TabIndex        =   13
         Top             =   1057
         Width           =   1600
      End
      Begin VB.Frame frmTime 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11295
         Begin MSComCtl2.DTPicker dtpStart 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   7560
            TabIndex        =   12
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   112328707
            CurrentDate     =   40833
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   9720
            TabIndex        =   11
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   112328707
            CurrentDate     =   40833
         End
         Begin VB.OptionButton optDays 
            Caption         =   "                   到"
            Height          =   180
            Index           =   6
            Left            =   7200
            TabIndex        =   19
            Top             =   240
            Width           =   2295
         End
         Begin VB.OptionButton optDays 
            Caption         =   "1天"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "3天"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "5天"
            Height          =   180
            Index           =   3
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "7天"
            Height          =   180
            Index           =   4
            Left            =   4800
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "半月"
            Height          =   180
            Index           =   5
            Left            =   6000
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optDays 
            Caption         =   "2天"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         Caption         =   "影像类别："
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "病人ID："
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "姓  名："
         Height          =   255
         Left            =   8880
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确 定(&O)"
      Height          =   350
      Left            =   9465
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   350
      Left            =   10665
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwUnMatched 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7223
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmXWRelateImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pstrStudyDate As String

Private mlngStudyID As Long         '新网的PACS关键字
Private mlngOrderID As Long         '中联的医嘱ID
Private mblnMatch As Boolean        '关联或者取消关联，True--关联；False--取消关联
Private mblnOpenDB As Boolean       '本窗口打开的数据库连接，关闭窗口时要关闭
Private mstrModality As String      '默认关联的影像类别

Private mrsUnMatchData As ADODB.Recordset

Private Sub cboModality_Click()
    
    If mblnMatch = True Then '关联图像
        If cboModality.ListIndex < 0 Then Exit Sub
        
        Call subFillUnMatched
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mblnMatch And Not lvwUnMatched.SelectedItem Is Nothing Then
        mlngStudyID = Val(Mid(lvwUnMatched.SelectedItem.Key, 2))
        pstrStudyDate = lvwUnMatched.SelectedItem.SubItems(5)
    ElseIf mblnMatch = False And Not lvwUnMatched.SelectedItem Is Nothing Then
        mlngStudyID = lvwUnMatched.SelectedItem.SubItems(1)
        pstrStudyDate = lvwUnMatched.SelectedItem.SubItems(4)
    End If
    Unload Me
End Sub

Public Function zlShowMe(frmParent As Form, lngOrderID As Long, blnMatch As Boolean, strModality As String) As Long
''--------------------------------------------
''功能： 显示未匹配的图像记录
''参数：frmParent --父窗体；
''      lngOrderID -- 医嘱ID ；
''      blnMatch --关联或者取消关联，True--关联；False--取消关联
''      strModality -- 需要关联图像的影像类别
''返回：需要匹配的医嘱ID
''--------------------------------------------
    On Error GoTo err
    
    mblnMatch = blnMatch
    mlngOrderID = lngOrderID
    mstrModality = strModality
    
    '判断数据库是否已经连接，如果没有连接，则打开连接
    If gcnXWDBServer.State <> adStateOpen Then
        If XWDBServerOpen = 0 Then
            mblnOpenDB = True
        End If
    End If
    
    mlngStudyID = 0
    
    If mblnMatch Then
        optDays(3).value = True
        Call subQueryUnmatched
        Call subFillUnMatched
        Call FillModality
    Else
        Call subFillMatched
    End If
    
    Me.Show 1, frmParent
    
    zlShowMe = mlngStudyID
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub subQueryUnmatched()
''--------------------------------------------
''功能： 查询未匹配的图像记录
''参数：无
''返回：无
''--------------------------------------------
    Dim strSql As String
    Dim dtNow As Date
    Dim i As Integer
    
    On Error GoTo err
    
    With lvwUnMatched
        If .ColumnHeaders.Count <= 0 Then
            With .ColumnHeaders
                .Clear
                .Add , , "姓名", 1000
                .Add , , "性别", 600
                .Add , , "出生日期", 1200
                .Add , , "年龄", 600
                .Add , , "病人ID", 1000
                .Add , , "检查日期", 1200
                .Add , , "检查时间", 1000
                .Add , , "检查描述", 1000
                .Add , , "影像类别", 1000
                .Add , , "检查项目", 2200
                .Add , , "图像数量", 800
            End With
            .ListItems.Add , , "Temp"
        End If
    End With
    
    On Error GoTo err
    
    dtNow = zlDatabase.Currentdate
    For i = 0 To 5
        If optDays(i).value = True Then
            Select Case i
                Case 0
                    dtpStart.value = dtNow
                    dtpEnd.value = dtNow
                Case 1
                    dtpStart.value = DateAdd("d", -1, dtNow)
                    dtpEnd.value = dtNow
                Case 2
                    dtpStart.value = DateAdd("d", -2, dtNow)
                    dtpEnd.value = dtNow
                Case 3
                    dtpStart.value = DateAdd("d", -4, dtNow)
                    dtpEnd.value = dtNow
                Case 4
                    dtpStart.value = DateAdd("d", -6, dtNow)
                    dtpEnd.value = dtNow
                Case 5
                    dtpStart.value = DateAdd("d", -14, dtNow)
                    dtpEnd.value = dtNow
            End Select
        End If
    Next i
    
    strSql = "select F_PAT_NAME as 姓名,F_PAT_NO as 病人ID,F_SEX as 性别,F_STU_BIRTH as 出生日期,F_STU_ID as 检查主键, " _
            & "F_STU_NO as 医嘱ID,F_STU_UID as 检查UID,F_AGE as 年龄,F_STU_DATE as 检查日期,F_STU_TIME as 检查时间, " _
            & " F_STU_SUSPICION as 检查描述,F_MODALITY as 影像类别,F_STU_PLACE as 检查项目,F_COUNT_IMG as 图像数量 from V_OEM_STUDY_UNMATCHED " _
            & " where F_MATCHED_FLAG = 0 and F_STU_DATE between '" & Format(dtpStart, "yyyy.mm.dd 00:00") & "' and '" & Format(dtpEnd, "yyyy.mm.dd 23:59") & "'"
    Set mrsUnMatchData = gcnXWDBServer.Execute(strSql)
Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subFillUnMatched()
''--------------------------------------------
''功能： 填充未匹配的图像记录
''参数：无
''返回：无
''--------------------------------------------
    Dim strFilter As String
    Dim tmpItem As ListItem
    
    On Error GoTo err
    
    '设置过滤条件
    strFilter = ""
    If cboModality.ListIndex >= 0 Then
        strFilter = IIf(strFilter = "", "", strFilter & " and ") & "影像类别 = '" & Split(cboModality.Text, "-")(0) & "'"
    End If
    
    If txtName.Text <> "" Then
        strFilter = IIf(strFilter = "", "", strFilter & " and ") & "姓名 = '" & txtName.Text & "'"
    End If
    
    If txtStudyNo.Text <> "" Then
        strFilter = IIf(strFilter = "", "", strFilter & " and ") & "病人ID = '" & txtStudyNo.Text & "'"
    End If
    
    mrsUnMatchData.Filter = strFilter
    
    lvwUnMatched.ListItems.Clear
    
    If Not mrsUnMatchData.EOF Then
        Do While Not mrsUnMatchData.EOF
            Set tmpItem = lvwUnMatched.ListItems.Add(, "_" & mrsUnMatchData("检查主键"), Nvl(mrsUnMatchData("姓名")))
            With tmpItem
                .SubItems(1) = Nvl(mrsUnMatchData("性别"))
                .SubItems(2) = Nvl(mrsUnMatchData("出生日期"))
                .SubItems(3) = Nvl(mrsUnMatchData("年龄"))
                .SubItems(4) = Nvl(mrsUnMatchData("病人ID"))
                .SubItems(5) = Nvl(mrsUnMatchData("检查日期"))
                .SubItems(6) = Nvl(mrsUnMatchData("检查时间"))
                .SubItems(7) = Nvl(mrsUnMatchData("检查描述"))
                .SubItems(8) = Nvl(mrsUnMatchData("影像类别"))
                .SubItems(9) = Nvl(mrsUnMatchData("检查项目"))
                .SubItems(10) = Nvl(mrsUnMatchData("图像数量"))

            End With
            mrsUnMatchData.MoveNext
        Loop
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub subFillMatched()
''--------------------------------------------
''功能： 填充已匹配的图像记录
''参数：无
''返回：无
''--------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim tmpItem As ListItem
    
    On Error GoTo err
    
    With lvwUnMatched
        If .ColumnHeaders.Count <= 0 Then
            With .ColumnHeaders
                .Clear
                .Add , , "影像类别", 2000
                .Add , , "检查号", 1000
                .Add , , "序列号", 4000
                .Add , , "说明", 4000
                .Add , , "采集时间", 2000
            End With
            .ListItems.Add , , "Temp"
        End If
        .ListItems.Clear
    End With
    
    strSql = "select F_SER_ID as SERIES主键,F_STU_ID as Study主键,F_SER_UID as 序列UID,F_SER_DATE as 序列日期,F_SER_TIME as 序列时间, " _
                & " F_SER_CONTEXT as 序列描述,F_MODALITY as 影像类型,F_STU_NO as 医嘱ID from V_OEM_SERIES where F_STU_NO ='" & mlngOrderID _
                & "' order by F_STU_ID ,F_SER_ID"
    Set rsTemp = gcnXWDBServer.Execute(strSql)
    
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
             Set tmpItem = lvwUnMatched.ListItems.Add(, "_" & rsTemp!SERIES主键, rsTemp!影像类型)
            With tmpItem
                .SubItems(1) = Nvl(rsTemp("Study主键"))
                .SubItems(2) = Nvl(rsTemp("序列UID"))
                .SubItems(3) = Nvl(rsTemp("序列描述"))
                .SubItems(4) = Nvl(rsTemp("序列日期"), date)
            End With
            rsTemp.MoveNext
        Loop
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdRepair_Click()
On Error GoTo errHandle
    Dim strSql As String
    Dim rsData As ADODB.Recordset

    '根据医嘱ID查询xwpacs中已经存在的图像检查数据
    strSql = "select F_STU_ID as 检查主键, F_STU_NO as 医嘱ID, F_STU_UID as 检查UID, F_STU_DATE as 检查日期, F_STU_TIME as 检查时间 " _
            & " from V_OEM_STUDY_UNMATCHED " _
            & " where F_STU_NO = '" & mlngOrderID & "'"
    
    Set rsData = gcnXWDBServer.Execute(strSql)
    
    If rsData.RecordCount <= 0 Then
        MsgBoxD Me, "图像关联状态修复失败，在影像服务器中未匹配到该检查信息。", vbOKOnly, "提示"
        Exit Sub
    End If


    '调用中联存储过程"b_XINWANGInterface.PacsStatusChange"，关联图像
    strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsStatusChange(1," & mlngOrderID & ",null,null,to_date('" _
                & Now & "','YYYY.MM.DD'),null,null)"
    zlDatabase.ExecuteProcedure strSql, "关联图像"
    
    mlngStudyID = -1
    
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpEnd_Change()
    If dtpStart.value > dtpEnd.value Then
        dtpEnd.value = dtpStart.value
    End If
    Call optDays_Click(6)
End Sub

Private Sub dtpEnd_GotFocus()
    optDays(6).value = True
End Sub

Private Sub dtpStart_Change()
    If dtpStart.value > dtpEnd.value Then
        dtpStart.value = dtpEnd.value
    End If
    Call optDays_Click(6)
End Sub

Private Sub dtpStart_GotFocus()
    optDays(6).value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '如果是在过程中打开的数据库连接，则退出时关闭连接
    If mblnOpenDB = True Then
        Call XWDBServerClose
    End If
End Sub

Private Sub FillModality()
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select 编码,名称 from 影像检查类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "影像检查类别")
    
    cboModality.Clear
    Do Until rsTemp.EOF
        cboModality.AddItem rsTemp!编码 & "-" & rsTemp!名称
        If rsTemp!编码 = mstrModality Then cboModality.ListIndex = cboModality.ListCount - 1
        rsTemp.MoveNext
    Loop
    
    If cboModality.ListIndex = -1 Then
        If cboModality.ListCount >= 1 Then
            cboModality.ListIndex = 1
        End If
    End If
End Sub

Private Sub optDays_Click(Index As Integer)
    Call subQueryUnmatched
    Call subFillUnMatched
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    '是回车，则查询
    Call subFillUnMatched
End Sub

Private Sub txtStudyNo_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    '是回车，则查询
    Call subFillUnMatched
End Sub
