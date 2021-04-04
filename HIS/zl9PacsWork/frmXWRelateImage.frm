VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXWRelateImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关联影像"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   Icon            =   "frmXWRelateImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdRepair 
      Caption         =   "修复关联状态(&R)"
      Height          =   350
      Left            =   7275
      TabIndex        =   20
      Top             =   6600
      Width           =   1845
   End
   Begin VB.Frame frmFilter 
      Caption         =   "过滤条件"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   11655
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查 询(&Q)"
         Height          =   350
         Left            =   10440
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cboModality 
         Height          =   300
         ItemData        =   "frmXWRelateImage.frx":038A
         Left            =   960
         List            =   "frmXWRelateImage.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox txtStudyNo 
         Height          =   300
         Left            =   3240
         TabIndex        =   14
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   5640
         TabIndex        =   13
         Top             =   960
         Width           =   1600
      End
      Begin VB.Frame frmTime 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11415
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
            Left            =   8160
            TabIndex        =   12
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   155779075
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
            Left            =   9960
            TabIndex        =   11
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   155779075
            CurrentDate     =   40833
         End
         Begin VB.OptionButton optDays 
            Caption         =   "                 到"
            Height          =   180
            Index           =   6
            Left            =   7800
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton optDays 
            Caption         =   "1天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "3天"
            Height          =   180
            Index           =   2
            Left            =   2040
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "5天"
            Height          =   180
            Index           =   3
            Left            =   3000
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "7天"
            Height          =   180
            Index           =   4
            Left            =   3960
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "半月"
            Height          =   180
            Index           =   5
            Left            =   4920
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optDays 
            Caption         =   "2天"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         Caption         =   "影像类别"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "病人ID"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   1000
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "姓  名："
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   1005
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确 定(&O)"
      Height          =   350
      Left            =   9465
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   350
      Left            =   10665
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwUnMatched 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
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
   Begin MSComctlLib.ListView lvwSeries 
      Height          =   1695
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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


Private mlngOrderID As Long         '中联的医嘱ID
Private mblnMatch As Boolean        '关联或者取消关联，True--关联；False--取消关联
Private mblnOpenDB As Boolean       '本窗口打开的数据库连接，关闭窗口时要关闭
Private mstrModality As String      '默认关联的影像类别
Private mlngReleationState As Long  '-1修复关联，0成功关联，1未关联，2关联失败

Private mrsUnMatchData As ADODB.Recordset

Private Sub cboModality_Click()
On Error GoTo errHandle
    If mblnMatch = True Then '关联图像
        If cboModality.ListIndex < 0 Then Exit Sub
        
        Call subFillUnMatched
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    mlngReleationState = 1
    Unload Me
End Sub

Private Function IsCheckSeries() As Boolean
'获取序列是否进行了勾选
    Dim i As Long
    
    IsCheckSeries = False
    
    For i = 1 To lvwSeries.ListItems.Count
        If lvwSeries.ListItems(i).Checked = True Then
            IsCheckSeries = True
            Exit Function
        End If
    Next i
End Function

Private Sub CmdOK_Click()
On Error GoTo errHandle
    Dim blnOpenDb As Boolean
    
    '如果数据库仍未打开成功，则退出处理
    If gcnXWDBServer.State <> adStateOpen Then
        MsgBox "PACS数据库服务不能正常连接，该操作将不能继续。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    If IsCheckSeries = False Then
        MsgBoxD Me, "请勾选需要处理的序列信息。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    If mblnMatch = True Then
        '关联影像
        mlngReleationState = IIf(ReleationImages(mlngOrderID) = True, 0, 2)
        
    ElseIf mblnMatch = False Then
        '取消关联
        mlngReleationState = IIf(CancelReleation(mlngOrderID) = True, 0, 2)
    End If
    
    
    Unload Me
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function ReleationImages(ByVal lngOrderID As Long) As Boolean
'关联图像
    Dim lngStudyId As Long
    Dim strSeriesIds As String
    Dim strUnCheckIds As String
    Dim lngSelectState As Long
    Dim lngSourceStudyId As Long
    Dim strSql As String
    
    Dim rsTemp As ADODB.Recordset
    Dim rsOrderInfo As ADODB.Recordset
    Dim strStudyDate As String
    Dim strStudyUID As String
    
    lngSelectState = GetSelectData(lngStudyId, strSeriesIds, strUnCheckIds, strStudyUID)
    
    If lngSelectState = 0 Then
        '没有选择任何数据，则直接退出
        MsgBoxD Me, "没有选择需要撤销的关联数据。", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    lngSourceStudyId = 0
    
    '查询医嘱ID是否有对应的图像关联信息
    strSql = "select f_stu_id from v_oem_study_unmatched where f_stu_no='" & lngOrderID & "'"
    
    Set rsTemp = gcnXWDBServer.Execute(strSql)
    If rsTemp.RecordCount > 0 Then lngSourceStudyId = Val(Nvl(rsTemp!f_stu_id))
    
    
    strStudyDate = Trim(lvwUnMatched.SelectedItem.SubItems(5)) & " " & Trim(lvwUnMatched.SelectedItem.SubItems(6))
    
    If lngSelectState = 2 Then
        '关联整个检查
        ReleationImages = ReleationStudy(lngOrderID, lngStudyId, strStudyDate, strStudyUID, IIf(lngSourceStudyId = 0, True, False))
    Else
        '关联部分序列
        
        '先对没有进行选择的序列进行检查拆分
        If XWUnmatchSeries(lngStudyId, strUnCheckIds) <> 0 Then
            Exit Function
        End If
        
        '对选中的序列进行关联操作
        ReleationImages = ReleationStudy(lngOrderID, lngStudyId, strStudyDate, strStudyUID, IIf(lngSourceStudyId = 0, True, False))
    End If
    
End Function

Private Function ReleationStudy(ByVal lngOrderID As Long, ByVal lngStudyId As Long, _
    ByVal strStudyDate As String, ByVal strStudyUID As String, Optional ByVal blnUpdateRis As Boolean = True) As Boolean
'--------------------------------------------
'功能： 关联检查
'参数： lngOrderID -- 医嘱ID
'       lngStudyId -- PACS中的检查主键
'       strStudyDate -- PACS中的检查日期
'       strStudyUID -- PACS中的检查UID
'       blnUpdateRis -- 可选参数，是否更新RIS数据
'返回：
'--------------------------------------------
    Dim strSql As String
    Dim rsOrderInfo As ADODB.Recordset
    
    ReleationStudy = False
    
    '查询关联检查所需的医嘱信息
    strSql = "Select b.病人ID,b.门诊号,b.住院号,b.健康号 as 体检号,b.姓名,b.性别,b.年龄,To_char(b.出生日期,'yyyymmdd') As 出生日期, " _
                & " c.英文名 as 拼音名,c.影像类别,c.检查号,a.病人来源,a.执行科室ID,d.名称 As 执行科室,a.开嘱时间,a.开始执行时间 " _
                & " From 病人医嘱记录 a,病人信息 b,影像检查记录 c,部门表 d  " _
                & " Where a.病人Id = b.病人ID And a.Id = c.医嘱ID And a.执行科室ID =d.Id  and a.Id = [1]"
    Set rsOrderInfo = zlDatabase.OpenSQLRecord(strSql, "查询检查信息", lngOrderID)
    
    If rsOrderInfo.RecordCount > 0 Then
        '调用新网存储过程“P_OEM_MATCHING_RIS”，关联图像
                
        strSql = "P_OEM_MATCHING_RIS(" & lngStudyId & ",'" & lngOrderID & "','" & rsOrderInfo!病人ID & "','" & Nvl(rsOrderInfo!门诊号, 0) _
                & "','" & Nvl(rsOrderInfo!住院号, 0) & "','" & Nvl(rsOrderInfo!体检号, 0) & "','" & Nvl(rsOrderInfo!姓名) & "','" _
                & Nvl(rsOrderInfo!性别) & "','" & Nvl(rsOrderInfo!年龄, 0) & "','" & Nvl(rsOrderInfo!出生日期) & "','" & Nvl(rsOrderInfo!拼音名) _
                & "','" & Nvl(rsOrderInfo!影像类别) & "','" & rsOrderInfo!检查号 & "'," & Nvl(rsOrderInfo!病人来源, 3) & "," & Nvl(rsOrderInfo!执行科室ID) _
                & ",'" & Nvl(rsOrderInfo!执行科室) & "','','')"
                
        gcnXWDBServer.Execute strSql
        
        If blnUpdateRis = True Then
            '调用中联存储过程"b_XINWANGInterface.PacsStatusChange"，关联图像
            strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") _
                & "b_XINWANGInterface.PacsStatusChange(1," & lngOrderID & ",'" & Nvl(rsOrderInfo!影像类别) _
                & "','" & rsOrderInfo!检查号 & "',to_date('" & Trim(strStudyDate) _
                & "','yyyy-mm-dd hh24:mi:ss'),null,null,'" & strStudyUID & "')"
                        
            zlDatabase.ExecuteProcedure strSql, "关联图像"
        End If
    End If
        
    ReleationStudy = True
End Function


Private Function GetSelectData(ByRef lngStudyId As Long, ByRef strSeriesIds As String, _
    ByRef strUnCheckIds As String, ByRef strStudyUID As String) As Long
'--------------------------------------------
'功能： 获取界面的数据选择情况,从界面树形结构中提取PACS数据的信息
'参数： lngStudyId -- V_OEM_STUDY_UNMATCHED中的F_STU_ID=检查主键
'       strSeriesIds --已选中的序列ID，V_OEM_SERIES的F_SER_ID=SERIES主键，的组合
'       strUnCheckIds --未选中的序列ID，V_OEM_SERIES的F_SER_ID=SERIES主键，的组合
'       strStudyUID -- V_OEM_STUDY_UNMATCHED中的F_STU_UID=检查UID
'返回：0-未选择,1-选择部分序列,2-选择整个检查
'--------------------------------------------
    Dim i As Long
    Dim lngSelectState As Long
    
    lngSelectState = 0
    
    lngStudyId = 0
    strSeriesIds = ""
    strUnCheckIds = ""
    
    For i = 1 To lvwSeries.ListItems.Count
        If lvwSeries.ListItems(i).Checked Then
            If strSeriesIds <> "" Then strSeriesIds = strSeriesIds & ","
            strSeriesIds = strSeriesIds & Val(Mid(lvwSeries.ListItems(i).Key, 2))
        Else
            If strUnCheckIds <> "" Then strUnCheckIds = strUnCheckIds & ","
            strUnCheckIds = strUnCheckIds & Val(Mid(lvwSeries.ListItems(i).Key, 2))
            
            lngSelectState = 1
        End If
    Next i
    
    '判断序列是否全选
    If lngSelectState <> 1 And strSeriesIds <> "" Then lngSelectState = 2
    
    lngStudyId = Val(Mid(lvwUnMatched.SelectedItem.Key, 2))
    strStudyUID = lvwUnMatched.SelectedItem.SubItems(11)
    
    GetSelectData = lngSelectState

End Function

Private Function CancelReleation(ByVal lngOrderID As Long) As Boolean
'取消关联处理
    Dim lngStudyId As Long
    Dim strSeriesIds As String
    Dim lngSelectState As Long
    Dim strUnCheckIds As String
    Dim strStudyUID As String
    
    '判断是否选中了所有序列，如果选中了所有序列，则直接与检查级别进行关联撤销
    CancelReleation = False
    
    lngSelectState = GetSelectData(lngStudyId, strSeriesIds, strUnCheckIds, strStudyUID)
    
    If lngSelectState = 0 Then
        '没有选择任何数据，则直接退出
        MsgBoxD Me, "没有选择需要撤销的关联数据。", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    If lngSelectState = 2 Then
        '处理检查级别的关联撤销
        CancelReleation = IIf(XWUnmatchImage(lngOrderID, lngStudyId) = 0, True, False)
    Else
        '处理序列级别的关联撤销
        CancelReleation = IIf(XWUnmatchSeries(lngStudyId, strSeriesIds) = 0, True, False)
    End If
    
End Function



Public Function zlShowMe(frmParent As Form, lngOrderID As Long, blnMatch As Boolean, strModality As String) As Long
''--------------------------------------------
''功能： 显示未匹配的图像记录
''参数：frmParent --父窗体；
''      lngOrderID -- 医嘱ID ；
''      blnMatch --关联或者取消关联，True--关联；False--取消关联
''      strModality -- 需要关联图像的影像类别
''返回：-1修复关联，0成功，1失败
''--------------------------------------------
    On Error GoTo err
    
    mblnMatch = blnMatch
    mlngOrderID = lngOrderID
    mstrModality = strModality
    
    mlngReleationState = 2
    
    '判断数据库是否已经连接，如果没有连接，则打开连接
    If gcnXWDBServer.State <> adStateOpen Then
        If XWDBServerOpen = 0 Then
            mblnOpenDB = True
        End If
    End If
    
    
    InitSeriesList
    InitStudyList
    
    frmFilter.Visible = blnMatch
    
    
    If mblnMatch Then
        '关联影像
        optDays(3).value = True
        
        Call subQueryUnmatched
        Call subFillUnMatched
        
        Call FillModality
                
        lvwSeries.Height = 1695
        Me.Caption = "关联影像"
    Else
        '取消关联
        'Call subFillMatched
        Call subQueryCurStudy(mlngOrderID)
        Call FillStudyData(mrsUnMatchData)
        
        lvwSeries.Height = 3375
        Me.Caption = "关联取消"
    End If
    
    Me.Show 1, frmParent
    
    zlShowMe = mlngReleationState
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub subQuerySeries(ByVal strStudyId As String)
'查询检查对应的序列信息
    Dim strSql As String
    Dim rsSeries As ADODB.Recordset
        
    strSql = "select F_SER_ID as 序列ID,F_STU_ID as 检查ID,F_SER_DATE as 序列日期, F_SER_TIME as 序列时间," & _
            " F_SER_NO as 序列号,  F_SER_CONTEXT as 序列描述,F_MODALITY as 设备类型," & _
            " F_SER_PLACE as 序列部位,F_COUNT_IMG as 图像数量  from v_oem_series where F_STU_ID=" & strStudyId & " order by 序列号 "
    Set rsSeries = gcnXWDBServer.Execute(strSql)
    
    Call FillSeriesData(rsSeries)
End Sub

Private Sub InitSeriesList()
'初始化序列列表
    Dim tmpItem As ListItem
    
    With lvwSeries
        .ListItems.Clear
        
        '如果未初始化列，则进行初始化
        If .ColumnHeaders.Count <= 0 Then
            With .ColumnHeaders
                .Clear
                .Add , , "序列ID", 1000
                .Add , , "序列号", 1000
                .Add , , "生成日期", 1200
                .Add , , "生成时间", 1200
                .Add , , "序列描述", 1600
                .Add , , "设备类型", 1000
                .Add , , "序列部位", 1200
                .Add , , "图像数量", 1000
            End With
        End If
    End With
End Sub

Private Sub FillSeriesData(rsSeries As ADODB.Recordset)
'填充检查数据
    Dim tmpItem As ListItem
    
    lvwSeries.ListItems.Clear
    
    If Not rsSeries.EOF Then
        Do While Not rsSeries.EOF
            Set tmpItem = lvwSeries.ListItems.Add(, "_" & rsSeries("序列ID"), Nvl(rsSeries("序列ID")))
            With tmpItem
                .SubItems(1) = Nvl(rsSeries("序列号"))
                .SubItems(2) = Nvl(rsSeries("序列日期"))
                .SubItems(3) = Nvl(rsSeries("序列时间"))
                .SubItems(4) = Nvl(rsSeries("序列描述"))
                .SubItems(5) = Nvl(rsSeries("设备类型"))
                .SubItems(6) = Nvl(rsSeries("序列部位"))
                .SubItems(7) = Nvl(rsSeries("图像数量"))
                '.Checked = True
            End With
            rsSeries.MoveNext
        Loop
    End If
End Sub

Private Sub InitStudyList()
'初始化检查列表
    Dim tmpItem As ListItem
    
    With lvwUnMatched
        .ListItems.Clear
        '如果未初始化列，则进行初始化
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
                .Add , , "检查UID", 1000
            End With
        End If
    End With
End Sub


Private Sub FillStudyData(rsData As ADODB.Recordset)
'填充检查数据
    Dim tmpItem As ListItem
    
    lvwUnMatched.ListItems.Clear
    lvwSeries.ListItems.Clear
    
    If Not rsData.EOF Then
        Do While Not rsData.EOF
            Set tmpItem = lvwUnMatched.ListItems.Add(, "_" & rsData("检查主键"), Nvl(rsData("姓名")))
            With tmpItem
                .SubItems(1) = Nvl(rsData("性别"))
                .SubItems(2) = Nvl(rsData("出生日期"))
                .SubItems(3) = Nvl(rsData("年龄"))
                .SubItems(4) = Nvl(rsData("病人ID"))
                .SubItems(5) = Nvl(rsData("检查日期"))
                .SubItems(6) = Nvl(rsData("检查时间"))
                .SubItems(7) = Nvl(rsData("检查描述"))
                .SubItems(8) = Nvl(rsData("影像类别"))
                .SubItems(9) = Nvl(rsData("检查项目"))
                .SubItems(10) = Nvl(rsData("图像数量"))
                .SubItems(11) = Nvl(rsData("检查UID"))

            End With
            rsData.MoveNext
        Loop
    End If
    
    If lvwUnMatched.ListItems.Count > 0 Then
        lvwUnMatched.ListItems(1).Selected = True
        
        Call lvwUnMatched_Click
    End If
End Sub

Private Sub subQueryCurStudy(ByVal lngOrderID As Long)
'查询当前检查信息
    Dim strSql As String
    Dim tmpItem As ListItem
    
    
        
    strSql = "select F_PAT_NAME as 姓名,F_PAT_NO as 病人ID,F_SEX as 性别,F_STU_BIRTH as 出生日期,F_STU_ID as 检查主键, " _
            & "F_STU_NO as 医嘱ID,F_STU_UID as 检查UID,F_AGE as 年龄,F_STU_DATE as 检查日期,F_STU_TIME as 检查时间, " _
            & " F_STU_SUSPICION as 检查描述,F_MODALITY as 影像类别,F_STU_PLACE as 检查项目,F_COUNT_IMG as 图像数量 from V_OEM_STUDY_UNMATCHED " _
            & " where F_Stu_No='" & lngOrderID & "'"
    
    Set mrsUnMatchData = gcnXWDBServer.Execute(strSql)
End Sub



Private Sub subQueryUnmatched()
''--------------------------------------------
''功能： 查询未匹配的图像记录
''参数：无
''返回：无
''--------------------------------------------
    Dim strSql As String
    Dim dtNow As Date
    Dim i As Integer
    Dim tmpItem As ListItem
    
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
            & " where F_MATCHED_FLAG = 0 and F_STU_DATE between '" & Format(dtpStart, "yyyy.mm.dd 00:00") & "' and '" & Format(dtpEnd, "yyyy.mm.dd 23:59") & "' order by F_PAT_NAME, F_STU_NO"
    
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
    
    Call FillStudyData(mrsUnMatchData)
    
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
    Dim strStudyUID As String

    '根据医嘱ID查询xwpacs中已经存在的图像检查数据
    strSql = "select F_STU_ID as 检查主键, F_STU_NO as 医嘱ID, F_STU_UID as 检查UID, F_STU_DATE as 检查日期, F_STU_TIME as 检查时间 " _
            & " from V_OEM_STUDY_UNMATCHED " _
            & " where F_STU_NO = '" & mlngOrderID & "'"
    
    Set rsData = gcnXWDBServer.Execute(strSql)
    
    If rsData.RecordCount <= 0 Then
        MsgBoxD Me, "图像关联状态修复失败，在影像服务器中未匹配到该检查信息。", vbOKOnly, "提示"
        Exit Sub
    End If

    strStudyUID = Nvl(rsData!检查UID)
    '调用中联存储过程"b_XINWANGInterface.PacsStatusChange"，关联图像
    strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsStatusChange(1," _
        & mlngOrderID & ",null,null,to_date('" & Now & "','YYYY.MM.DD'),null,null,'" & strStudyUID & "')"
    zlDatabase.ExecuteProcedure strSql, "关联图像"
    
    mlngReleationState = -1
    
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Command1_Click()
On Error GoTo errHandle
    '是回车，则查询
    Call subQueryUnmatched
    Call subFillUnMatched
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpEnd_Change()
On Error GoTo errHandle
    If dtpStart.value > dtpEnd.value Then
        dtpEnd.value = dtpStart.value
    End If
    
    Call optDays_Click(6)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpEnd_GotFocus()
    optDays(6).value = True
End Sub

Private Sub dtpStart_Change()
On Error GoTo errHandle
    If dtpStart.value > dtpEnd.value Then
        dtpStart.value = dtpEnd.value
    End If
    Call optDays_Click(6)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
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

Private Sub lvwUnMatched_Click()
'查询检查对应的序列信息
On Error GoTo errHandle
    Dim strStudyId As String
    
    If lvwUnMatched.SelectedItem Is Nothing Then Exit Sub
    
    strStudyId = Mid(lvwUnMatched.SelectedItem.Key, 2)
    
    Call subQuerySeries(strStudyId)

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optDays_Click(Index As Integer)
On Error GoTo errHandle
    Call subQueryUnmatched
    Call subFillUnMatched
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    
    '是回车，则查询
    Call subFillUnMatched
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtStudyNo_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    
    '是回车，则查询
    Call subFillUnMatched
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
