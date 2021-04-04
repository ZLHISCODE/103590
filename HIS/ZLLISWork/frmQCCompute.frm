VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQCCompute 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "失控计算"
   ClientHeight    =   8415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7710
   Icon            =   "frmQCCompute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo项目 
      Height          =   300
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   795
      Width           =   5000
   End
   Begin VB.Frame fraRule 
      Height          =   5190
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   7300
      Begin VB.CheckBox chk多规则 
         Caption         =   "多规则"
         Height          =   225
         Left            =   375
         TabIndex        =   18
         Top             =   225
         Visible         =   0   'False
         Width           =   6600
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "复制规则"
         Height          =   350
         Left            =   5940
         TabIndex        =   14
         Top             =   4650
         Width           =   1100
      End
      Begin VB.Frame fra2 
         Caption         =   "计算控制界限规则"
         Height          =   1545
         Left            =   210
         TabIndex        =   13
         Top             =   2130
         Width           =   6900
         Begin VB.CheckBox chk控界 
            Height          =   210
            Index           =   0
            Left            =   135
            TabIndex        =   20
            Top             =   255
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lbl提示 
            Caption         =   "当仪器每批水平数>1时才能选择计算控制界限规则"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   1065
            TabIndex        =   22
            Top             =   630
            Visible         =   0   'False
            Width           =   4050
         End
      End
      Begin VB.Frame fra1 
         Caption         =   "常用质控规则"
         Height          =   1545
         Left            =   225
         TabIndex        =   12
         Top             =   525
         Width           =   6900
         Begin VB.CheckBox chk常用 
            Height          =   210
            Index           =   0
            Left            =   135
            TabIndex        =   19
            Top             =   255
            Visible         =   0   'False
            Width           =   1200
         End
      End
      Begin VB.Frame fra3 
         Caption         =   "累积和规则"
         Height          =   795
         Left            =   210
         TabIndex        =   11
         Top             =   3765
         Width           =   6900
         Begin VB.CheckBox chk累积 
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   300
            Visible         =   0   'False
            Width           =   1600
         End
      End
   End
   Begin VB.CheckBox chkALL 
      Caption         =   "计算本仪器所有项目"
      Height          =   195
      Left            =   4290
      TabIndex        =   9
      Top             =   1485
      Width           =   1995
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -45
      TabIndex        =   8
      Top             =   345
      Width           =   11865
   End
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "返回(&X)"
      Height          =   350
      Left            =   6210
      TabIndex        =   1
      Top             =   495
      Width           =   1100
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "计算(&E)"
      Height          =   350
      Left            =   6210
      TabIndex        =   0
      Top             =   930
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1275
      Left            =   480
      TabIndex        =   6
      Top             =   1725
      Width           =   6900
      _cx             =   12171
      _cy             =   2249
      Appearance      =   2
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComCtl2.DTPicker dtp日期 
      Height          =   300
      Index           =   0
      Left            =   990
      TabIndex        =   15
      Top             =   1140
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   118161411
      CurrentDate     =   39110
   End
   Begin MSComCtl2.DTPicker dtp日期 
      Height          =   300
      Index           =   1
      Left            =   4295
      TabIndex        =   16
      Top             =   1140
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   118161411
      CurrentDate     =   39110
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmQCCompute.frx":058A
      Top             =   75
      Width           =   240
   End
   Begin VB.Label lbl质控品 
      AutoSize        =   -1  'True
      Caption         =   "请选择2个不同浓度水平的质控品:"
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   1485
      Width           =   2700
   End
   Begin VB.Label lbl项目 
      AutoSize        =   -1  'True
      Caption         =   "项目: ####"
      Height          =   180
      Left            =   450
      TabIndex        =   5
      Top             =   825
      Width           =   900
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "根据仪器设置的质控规则，自动进行失控计算，标记失控状态。"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   4
      Top             =   90
      Width           =   5040
   End
   Begin VB.Label lbl日期 
      AutoSize        =   -1  'True
      Caption         =   "日期: ####"
      Height          =   180
      Left            =   450
      TabIndex        =   3
      Top             =   1155
      Width           =   900
   End
   Begin VB.Label lbl仪器 
      AutoSize        =   -1  'True
      Caption         =   "仪器: ####"
      Height          =   180
      Left            =   450
      TabIndex        =   2
      Top             =   495
      Width           =   900
   End
End
Attribute VB_Name = "frmQCCompute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum mCol
    ID = 0: 选择: 质控品: 水平: 必选
End Enum

Private mlngDevID As Long       '仪器id
Private mlngItemID As Long      '项目id
Private mdtBeging As Date        '日期
Private mdtEnd As Date          '结束日期
Private mintLevel As Integer    '由仪器决定的质控水平数
Private mblnModify As Boolean   '是否执行

Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function ShowMe(frmParent As Form, lngDevId As Long, lngItemID As Long, dtBegin As Date, Optional lngResId As Long, Optional dtEnd As Date) As Boolean
    '功能：根据指定仪器、项目、日期，并显示计算对话框
    Dim rsTemp As New adodb.Recordset
    
    mlngDevID = lngDevId
    mlngItemID = lngItemID
    mdtBeging = dtBegin
    If dtEnd = CDate(0) Then
        mdtEnd = dtBegin
    Else
        mdtEnd = dtEnd
    End If
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select 编码 || ', ' || 名称 As 仪器名, 质控水平数 From 检验仪器 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe", mlngDevID)
    With rsTemp
        If .RecordCount <= 0 Then MsgBox "指定仪器不存在！", vbInformation, gstrSysName: Exit Function
        mintLevel = !质控水平数
        Me.lbl仪器.Caption = "仪器: " & !仪器名
    End With
    

    
    gstrSql = "Select 编码 || ', ' || 中文名 || ', ' || 英文名 As 项目名 From 诊治所见项目 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
    With rsTemp
        If .RecordCount <= 0 Then MsgBox "指定项目不存在！", vbInformation, gstrSysName: Exit Function
        Me.lbl项目.Caption = "项目: " & !项目名
    End With
    Me.lbl日期.Caption = "日期: " & Format(dtBegin, "yyyy年mm月dd日")
    Me.lbl质控品.Caption = "请选择" & mintLevel & "个不同浓度水平的质控品:"
    
    Me.dtp日期(0).Value = mdtBeging: Me.dtp日期(0).MinDate = mdtBeging: Me.dtp日期(0).MaxDate = mdtEnd
    Me.dtp日期(1).Value = mdtEnd: Me.dtp日期(1).MinDate = mdtBeging: Me.dtp日期(1).MaxDate = mdtEnd
    
    gstrSql = "Select M.ID, '' As 选择, M.批号 || ', ' || M.名称 || ', 水平:' || M.水平 As 质控品, M.水平, 0 As 必选" & vbNewLine & _
            "From 检验质控品 M, 检验质控品项目 I" & vbNewLine & _
            "Where M.ID = I.质控品id And M.仪器id=[1] And I.项目id = [2] And " & vbNewLine & _
            "      To_Date([3],'yyyy-MM-dd') Between M.开始日期 And M.结束日期" & vbNewLine & _
            "Order By M.开始日期, M.水平"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDevId, lngItemID, Format(mdtBeging, "yyyy-MM-dd"))
    With Me.vfgList
        Set .DataSource = rsTemp
        .ColWidth(mCol.选择) = 280
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.水平) = 0: .ColWidth(mCol.必选) = 0
        .ColHidden(mCol.ID) = True: .ColHidden(mCol.水平) = True: .ColHidden(mCol.必选) = True
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mCol.ID) = lngResId Or lngCount < mintLevel Then
                .Cell(flexcpChecked, lngCount, mCol.选择) = flexChecked
            Else
                .Cell(flexcpChecked, lngCount, mCol.选择) = flexUnchecked
            End If
        Next
    End With
    
    mblnModify = False
    Me.Show vbModal, frmParent
    ShowMe = mblnModify
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initRuleCtr()
    '初始化质控规则选择控件
    Dim rsTmp As adodb.Recordset
    Dim strSQL As String
    Dim lngLeft As Long, lngTop As Long
    
    On Error GoTo ErrHandle
    '多水平才能使用控制界限规则
    lbl提示.Visible = False
    If mintLevel > 1 Then
        strSQL = "Select B.ID, B.名称, B.种类, B.说明, B.形式, B.多水平 From 检验质控规则 B Order By 种类,形式, B.编码 "
    Else
        strSQL = "Select B.ID, B.名称, B.种类, B.说明, B.形式, B.多水平 From 检验质控规则 B Where 种类 In (1, 3)  Order By 种类,形式, B.编码 "
        lbl提示.Visible = True
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do Until rsTmp.EOF
        Select Case "" & rsTmp!种类
        Case 1 '常用质控规则
            
            If Trim(chk常用(chk常用.UBound).Tag) <> "" Then
                Load chk常用(chk常用.UBound + 1)
                lngLeft = chk常用(chk常用.UBound - 1).Left + chk常用(chk常用.UBound - 1).Width + 45
                lngTop = chk常用(chk常用.UBound - 1).Top
            Else
                lngLeft = chk常用(chk常用.UBound).Left
                lngTop = chk常用(chk常用.UBound).Top
            End If
            
            chk常用(chk常用.UBound).Caption = rsTmp!名称
            chk常用(chk常用.UBound).Value = 0
            chk常用(chk常用.UBound).Tag = rsTmp!ID
            chk常用(chk常用.UBound).Visible = False
            If lngLeft + chk常用(chk常用.UBound).Width > Me.fra1.Left + Me.fra1.Width - 155 Then
                lngLeft = chk常用(0).Left
                lngTop = chk常用(chk常用.UBound - 1).Top + chk常用(chk常用.UBound - 1).Height + 45
            End If
            chk常用(chk常用.UBound).Left = lngLeft
            chk常用(chk常用.UBound).Top = lngTop
            If Trim(chk常用(chk常用.UBound).Caption) <> "" Then chk常用(chk常用.UBound).Visible = True

        Case 2 '控制界限规则
            
            If Trim(chk控界(chk控界.UBound).Tag) <> "" Then
                Load chk控界(chk控界.UBound + 1)
                lngLeft = chk控界(chk控界.UBound - 1).Left + chk控界(chk控界.UBound - 1).Width + 45
                lngTop = chk控界(chk控界.UBound - 1).Top
            Else
                lngLeft = chk控界(chk控界.UBound).Left
                lngTop = chk控界(chk控界.UBound).Top
            End If
            
            chk控界(chk控界.UBound).Caption = rsTmp!名称
            chk控界(chk控界.UBound).Value = 0
            chk控界(chk控界.UBound).Tag = rsTmp!ID
            chk控界(chk控界.UBound).Visible = False
            If lngLeft + chk控界(chk控界.UBound).Width > Me.fra1.Left + Me.fra1.Width - 155 Then
                lngLeft = chk控界(0).Left
                lngTop = chk控界(chk控界.UBound - 1).Top + chk控界(chk控界.UBound - 1).Height + 45
            End If
            chk控界(chk控界.UBound).Left = lngLeft
            chk控界(chk控界.UBound).Top = lngTop
            If Trim(chk控界(chk控界.UBound).Caption) <> "" Then chk控界(chk控界.UBound).Visible = True
        
        Case Else   '累积和规则
            
            If Trim(chk累积(chk累积.UBound).Tag) <> "" Then
                Load chk累积(chk累积.UBound + 1)
                lngLeft = chk累积(chk累积.UBound - 1).Left + chk累积(chk累积.UBound - 1).Width + 45
                lngTop = chk累积(chk累积.UBound - 1).Top
            Else
                lngLeft = chk累积(chk累积.UBound).Left
                lngTop = chk累积(chk累积.UBound).Top
            End If
            
            chk累积(chk累积.UBound).Caption = rsTmp!名称
            chk累积(chk累积.UBound).Value = 0
            chk累积(chk累积.UBound).Tag = rsTmp!ID
            chk累积(chk累积.UBound).Visible = False
            If lngLeft + chk累积(chk累积.UBound).Width > Me.fra1.Left + Me.fra1.Width - 155 Then
                lngLeft = chk累积(0).Left
                lngTop = chk累积(chk累积.UBound - 1).Top + chk累积(chk累积.UBound - 1).Height + 45
            End If
            chk累积(chk累积.UBound).Left = lngLeft
            chk累积(chk累积.UBound).Top = lngTop
            If Trim(chk累积(chk累积.UBound).Caption) <> "" Then chk累积(chk累积.UBound).Visible = True
        End Select
        
        rsTmp.MoveNext
    Loop
    
    '其他控件
    strSQL = "Select Distinct I.ID, I.编码, I.中文名, I.英文名" & vbNewLine & _
            "From 检验质控品 M, 检验质控品项目 Q, 诊治所见项目 I" & vbNewLine & _
            "Where M.ID = Q.质控品id And Q.项目id = I.ID And M.仪器id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDevID)
    With rsTmp
        Me.cbo项目.Clear
        Do Until .EOF
            Me.cbo项目.AddItem !编码 & ", " & !中文名 & "/" & !英文名
            Me.cbo项目.ItemData(Me.cbo项目.NewIndex) = !ID
            If !ID = mlngItemID Then
                Me.cbo项目.ListIndex = Me.cbo项目.NewIndex
            End If
            .MoveNext
        Loop
        If Me.cbo项目.ListCount = 0 Then MsgBox "尚未完成仪器质控品设置！", vbInformation, gstrSysName: Unload Me: Exit Sub
        If cbo项目.ListIndex = -1 Then
            Me.cbo项目.ListIndex = 0
        End If

    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub refRuleStat()
    '显示当前项目的规则状态
    Dim rsTmp As adodb.Recordset
    Dim strSQL As String
    Dim intCount As Integer
    
    On Error GoTo ErrHandle
    
    '---- 置为初始状态
    chk多规则.Value = 0
    For intCount = chk常用.LBound To chk常用.UBound
        chk常用(intCount).Value = 0
    Next
    For intCount = chk控界.LBound To chk控界.UBound
        chk控界(intCount).Value = 0
    Next
    For intCount = chk累积.LBound To chk累积.UBound
        chk累积(intCount).Value = 0
    Next
    '-----
    mlngItemID = Me.cbo项目.ItemData(Me.cbo项目.ListIndex)
    
    strSQL = "Select A.ID, A.仪器id, A.规则id, A.性质, B.名称, B.种类, B.说明, B.形式, B.多水平, A.是否使用" & vbNewLine & _
            "From 检验仪器规则 A, 检验质控规则 B" & vbNewLine & _
            "Where A.规则id = B.ID And A.上级id Is Null And A.仪器id = [1] And A.项目id = [2] " & vbNewLine & _
            "Order By A.性质, B.种类, A.规则id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDevID, mlngItemID)
    chk多规则.Visible = False
    Do Until rsTmp.EOF
        If Val("" & rsTmp!性质) = 0 Then
            
            chk多规则.Value = Val("" & rsTmp!是否使用)
            chk多规则.Visible = True
        ElseIf Val("" & rsTmp!性质) = 1 Then
            If Val("" & rsTmp!是否使用) = 1 Then
                Select Case Val("" & rsTmp!种类)
                Case 1
                    For intCount = chk常用.LBound To chk常用.UBound
                        If Val(chk常用(intCount).Tag) = Val("" & rsTmp!规则id) Then
                            chk常用(intCount).Value = 1
                            Exit For
                        End If
                    Next
                Case 2
                    For intCount = chk控界.LBound To chk控界.UBound
                        If Val(chk控界(intCount).Tag) = Val("" & rsTmp!规则id) Then
                            chk控界(intCount).Value = 1
                            Exit For
                        End If
                    Next
                Case Else
                    For intCount = chk累积.LBound To chk累积.UBound
                        If Val(chk累积(intCount).Tag) = Val("" & rsTmp!规则id) Then
                            chk累积(intCount).Value = 1
                            Exit For
                        End If
                    Next
                End Select
            End If
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CheckRule(ByVal lng规则ID As Long, ByVal int是否使用 As Integer)
    '核对检验仪器规则的记录，该增加则增加，该修改则修改。
    '注：只能对附加规则进行处理
    Dim strSQL As String
    Dim rsTmp As adodb.Recordset
    Dim blnHave  As Boolean
    On Error GoTo ErrHandle

    strSQL = "ZL_检验仪器规则_SetUsed(" & mlngDevID & "," & mlngItemID & "," & lng规则ID & "," & int是否使用 & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub cbo项目_Click()
    Call refRuleStat
End Sub


Private Sub chk常用_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lng规则ID As Long
    If Button = 1 Then
        lng规则ID = Val(Me.chk常用(Index).Tag)
        If Me.chk常用(Index).Value = 0 Then
            Call CheckRule(lng规则ID, 1)
        Else
            Call CheckRule(lng规则ID, 0)
        End If
    End If
End Sub

Private Sub chk控界_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lng规则ID As Long
    If Button = 1 Then
        lng规则ID = Val(Me.chk控界(Index).Tag)
        If Me.chk控界(Index).Value = 0 Then
            Call CheckRule(lng规则ID, 1)
        Else
            Call CheckRule(lng规则ID, 0)
        End If
    End If
End Sub

Private Sub chk累积_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lng规则ID As Long
    If Button = 1 Then
        lng规则ID = Val(Me.chk累积(Index).Tag)
        If Me.chk累积(Index).Value = 0 Then
            Call CheckRule(lng规则ID, 1)
        Else
            Call CheckRule(lng规则ID, 0)
        End If
    End If
End Sub

Private Sub cmdApply_Click()
    Call frmAppRuleCopy.ShowMe(mlngDevID, mlngItemID, Me)
End Sub

Private Sub cmdExecute_Click()
    Dim strResList As String, strLevels As String
    Dim rsTemp As New adodb.Recordset, rsTmp As New adodb.Recordset
    Dim strReturn As String
    Dim lngLoop As Long, lngDate As Date, lngCount As Long, strInfo As String
    
    Dim dtBeging As Date, dtEnd As Date
    
    strResList = "": strLevels = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.选择) = flexChecked Then
                If InStr(1, strLevels, Trim(.TextMatrix(lngCount, mCol.水平))) > 0 Then
                    MsgBox "只能允许一个水平" & Trim(.TextMatrix(lngCount, mCol.水平)) & "的质控品！", vbInformation, gstrSysName
                    Exit Sub
                End If
                strLevels = strLevels & "," & Trim(.TextMatrix(lngCount, mCol.水平))
                strResList = strResList & "," & Trim(.TextMatrix(lngCount, mCol.ID))
            End If
        Next
    End With
    If strResList <> "" Then strResList = Mid(strResList, 2)
'    2009-06-03 塘厦：多水平计算时，不需要每个水平的测试个数一致，只需要总个数相符就可以计算
'    If UBound(Split(strResList, ",")) <> mintLevel - 1 Then
'        MsgBox "请按仪器质控要求选择" & mintLevel & "个不同水平的质控品！", vbInformation, gstrSysName
'        Exit Sub
'    End If
    
    Err = 0: On Error GoTo ErrHand
    dtBeging = dtp日期(0).Value: dtEnd = dtp日期(1).Value
    
    If Me.chkALL.Value = 1 Then
        
        gstrSql = "Select Distinct B.项目id, C.编码, C.中文名, C.英文名" & vbNewLine & _
                    " From 检验质控品 A, 检验质控品项目 B, 诊治所见项目 C" & vbNewLine & _
                    " Where A.ID = B.质控品id And B.项目id = C.ID And A.仪器id = [1] "
            
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevID)
        Do Until rsTmp.EOF
            '计算一段时间
            lngCount = DateDiff("d", dtBeging, dtEnd)
            For lngLoop = 0 To lngCount
                gstrSql = "Select Zl_检验质控记录_Compute(" & mlngDevID & ", " & rsTmp("项目ID") & ", To_Date('" & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & "','yyyy-mm-dd'), '" & strResList & "') From Dual"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)

                If rsTemp.RecordCount <= 0 Then strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")  计算过程调用错误！" & vbCrLf
                If InStr(rsTemp.Fields(0).Value, "出现失控！") > 0 Then
                    strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")" & rsTemp.Fields(0).Value & vbCrLf
                    
                    ' 2009-06-03 塘厦：当前计算点失控时，后续点不再计算,
                    Exit For
                ElseIf InStr(rsTemp.Fields(0).Value, "计算完成！") <= 0 Then
                    If InStr(rsTemp.Fields(0).Value, "按规则未发现警告和失控！") <= 0 Then
                    strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")" & rsTemp.Fields(0).Value & vbCrLf
                    End If
                End If
            Next
            rsTmp.MoveNext
        Loop
    Else
        lngCount = DateDiff("d", dtBeging, dtEnd)
        
        For lngLoop = 0 To lngCount
            gstrSql = "Select Zl_检验质控记录_Compute(" & mlngDevID & ", " & mlngItemID & ", To_Date('" & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & "','yyyy-mm-dd'), '" & strResList & "') From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            If rsTemp.RecordCount <= 0 Then strReturn = strReturn & Format(DateAdd("d", lngCount, dtBeging), "yyyy-mm-dd") & " " & Nvl(rsTmp("中文名")) & "(" & Nvl(rsTmp("英文名")) & ")  计算过程调用错误！" & vbCrLf
            If InStr(rsTemp.Fields(0).Value, "出现失控！") > 0 Then
                strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & rsTemp.Fields(0).Value & vbCrLf
                ' 2009-06-03 塘厦：当前计算点失控时，后续点不再计算,
                Exit For
            ElseIf InStr(rsTemp.Fields(0).Value, "计算完成！") <= 0 Then
                If InStr(rsTemp.Fields(0).Value, "按规则未发现警告和失控！") <= 0 Then
                strReturn = strReturn & Format(DateAdd("d", lngLoop, dtBeging), "yyyy-mm-dd") & " " & rsTemp.Fields(0).Value & vbCrLf
                End If
            End If
       Next
    End If
    If Trim(strReturn) = "" Then
        strReturn = "计算完成，按规则未发现警告和失控！"
        MsgBox strReturn, vbInformation, gstrSysName
    Else
        Call frmQCShowInfo.ShowMe(Me.Caption, strReturn, Me)
    End If
    mblnModify = True
    If Left(strReturn, 4) = "计算完成" Then Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
    
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    chkALL.Value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "计算所有项目", 1)
    Call initRuleCtr
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "计算所有项目", chkALL.Value)

    For lngCount = 0 To chk常用.Count - 1
        If lngCount > 0 Then Unload chk常用(lngCount)
    Next
    For lngCount = 0 To chk控界.Count - 1
        If lngCount > 0 Then Unload chk控界(lngCount)
    Next
    
    For lngCount = 0 To chk累积.Count - 1
        If lngCount > 0 Then Unload chk累积(lngCount)
    Next
End Sub

Private Sub vfgList_DblClick()
    With Me.vfgList
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        If .Cell(flexcpChecked, .Row, mCol.选择) = flexUnchecked Or Val(.TextMatrix(.Row, mCol.必选)) = 1 Then
            .Cell(flexcpChecked, .Row, mCol.选择) = flexChecked
        Else
            .Cell(flexcpChecked, .Row, mCol.选择) = flexUnchecked
        End If
    End With
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call vfgList_DblClick
End Sub
