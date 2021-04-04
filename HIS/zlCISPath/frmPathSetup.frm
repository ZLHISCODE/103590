VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmPathSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "路径选项"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8295
   Icon            =   "frmPathSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPath 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   0
      ScaleHeight     =   1920
      ScaleWidth      =   8295
      TabIndex        =   10
      Top             =   5700
      Width           =   8295
      Begin VB.Frame fraRule 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         TabIndex        =   28
         Top             =   1560
         Width           =   2415
         Begin VB.OptionButton optPrtRule 
            Caption         =   "按天打印"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optPrtRule 
            Caption         =   "按阶段打印"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   29
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkIsForwardSend 
         Caption         =   "允许提前生成明天的路径项目"
         Height          =   180
         Left            =   4080
         TabIndex        =   27
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CheckBox chkIsEvaluate 
         Caption         =   "允许前一天不评估就生成今天的路径项目"
         Height          =   180
         Left            =   4080
         TabIndex        =   26
         Top             =   840
         Width           =   3615
      End
      Begin VB.OptionButton optPrintDay 
         Caption         =   "3天"
         Height          =   180
         Index           =   1
         Left            =   2760
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton optPrintDay 
         Caption         =   "2天"
         Height          =   180
         Index           =   0
         Left            =   2160
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txt中药味数 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "30"
         Top             =   840
         Width           =   300
      End
      Begin VB.CheckBox chk未评估时允许添加医嘱到昨天 
         Caption         =   "未评估时允许添加医嘱到昨天"
         Height          =   180
         Left            =   4080
         TabIndex        =   18
         Top             =   480
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   3000
         TabIndex        =   17
         Top             =   1740
         Width           =   300
      End
      Begin VB.TextBox txtSendAdvice 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   16
         Text            =   "1"
         Top             =   1560
         Width           =   300
      End
      Begin VB.CheckBox chkOutTable 
         Caption         =   "医技科室下达医嘱不在路径表上记录"
         Height          =   180
         Left            =   4080
         TabIndex        =   15
         Top             =   120
         Width           =   3375
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   2805
         TabIndex        =   14
         Top             =   1030
         Width           =   420
      End
      Begin VB.Frame fraPathExe 
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3135
         Begin VB.CheckBox chkInvocation 
            Caption         =   "启用路径执行环节"
            Height          =   180
            Left            =   240
            TabIndex        =   25
            Top             =   0
            Width           =   1815
         End
         Begin VB.CheckBox chkDoctor 
            Caption         =   "医生"
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox chkNurse 
            Caption         =   "护士"
            Height          =   255
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Label lblPrtRule 
         Caption         =   "路径表单打印规则："
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblPrintDays 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径表单每页打印的天数"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1980
      End
      Begin VB.Label lblSendAdvice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生成路径时，医嘱允许超过当前时间    天"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   3420
      End
      Begin VB.Label lbl中药味数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中药配方允许修改的中药味数上限    %"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   3150
      End
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   9
      Top             =   5520
      Width           =   6975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7080
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8295
      TabIndex        =   6
      Top             =   0
      Width           =   8295
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医嘱类路径项目生成顺序"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1095
         TabIndex        =   8
         Top             =   120
         Width           =   2145
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "    路径项目生成医嘱时缺省按路径表中该阶段定义的分类及项目顺序列出，但优先按下表定义的类别顺序排列，每次生成医嘱时也可以调整顺序。"
         Height          =   360
         Left            =   1095
         TabIndex        =   7
         Top             =   360
         Width           =   6165
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathSetup.frx":038A
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   1
      Left            =   7680
      Picture         =   "frmPathSetup.frx":6504
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   0
      Left            =   7080
      Picture         =   "frmPathSetup.frx":69B5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7080
      TabIndex        =   2
      Top             =   1440
      Width           =   1100
   End
   Begin VB.PictureBox picAddRow 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   220
      Left            =   6000
      Picture         =   "frmPathSetup.frx":6E6E
      ScaleHeight     =   225
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   360
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6900
      _cx             =   12171
      _cy             =   8017
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathSetup.frx":71F8
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmPathSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mbytFun  As Byte     '0=临床路径模块调用,1=医生站调用
Private Enum CNAME
    c顺序 = 0
    c期效 = 1
    c诊疗类别 = 2
    c操作类型 = 3
    c给药分类 = 4
    c操作 = 5
End Enum

Private Sub chkDoctor_Click()
    If chkNurse.Value = Unchecked And chkDoctor.Value = Unchecked Then
        chkInvocation.Value = Unchecked
    End If
End Sub

Private Sub chkInvocation_Click()
    
    If chkInvocation.Value = Checked Then
        chkDoctor.Enabled = True
        chkNurse.Enabled = True
        chkDoctor.Value = Checked
        chkNurse.Value = Checked
    Else
        chkDoctor.Enabled = False
        chkNurse.Enabled = False
        chkDoctor.Value = Unchecked
        chkNurse.Value = Unchecked
    End If
End Sub

Private Sub chkIsEvaluate_Click()
    If chkIsEvaluate.Value = Unchecked Then
        chkIsForwardSend.Value = Unchecked
        chkIsForwardSend.Enabled = False
    Else
        chkIsForwardSend.Enabled = True
    End If
End Sub

Private Sub chkNurse_Click()
    If chkNurse.Value = Unchecked And chkDoctor.Value = Unchecked Then
        chkInvocation.Value = Unchecked
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click(Index As Integer)
    With vsItem
            If Index = 0 And .Row > .FixedRows Then
                .RowPosition(.Row) = .Row - 1
                .TextMatrix(.Row, c顺序) = .TextMatrix(.Row, c顺序) + 1
                .TextMatrix(.Row - 1, c顺序) = .TextMatrix(.Row - 1, c顺序) - 1
                .Row = .Row - 1
            ElseIf Index = 1 And .Row < .Rows - 1 Then
                .RowPosition(.Row) = .Row + 1
                .TextMatrix(.Row, c顺序) = .TextMatrix(.Row, c顺序) - 1
                .TextMatrix(.Row + 1, c顺序) = .TextMatrix(.Row + 1, c顺序) + 1
                .Row = .Row + 1
            End If
    End With
End Sub

Private Sub cmdOK_Click()
    If Not (vsItem.Rows = 2 And vsItem.TextMatrix(1, CNAME.c期效) = "" And vsItem.TextMatrix(1, CNAME.c诊疗类别) = "" _
            And vsItem.TextMatrix(1, CNAME.c操作类型) = "" And vsItem.TextMatrix(1, CNAME.c给药分类) = "") Then
        If CheckData = False Then Exit Sub
    End If
    
    Call SaveData
    Unload Me
End Sub

Private Function CheckData() As Boolean
    Dim i As Long, str操作类型 As String, str给药分类 As String
    Dim rsSQL As ADODB.Recordset, strKey As String
    
    
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "行号", adBigInt
    rsSQL.Fields.Append "值", adVarChar, 200, adFldIsNullable
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    With vsItem
        .Cell(flexcpBackColor, .FixedRows, c顺序, .Rows - 1, .Cols - 1) = vbWhite
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, c期效) = "" Then
                MsgBox "请选择医嘱期效。", vbInformation, gstrSysName
                .Select i, c期效
                Exit Function
            ElseIf .TextMatrix(i, c诊疗类别) = "" Then
                MsgBox "请选择诊疗项目类别。", vbInformation, gstrSysName
                .Select i, c诊疗类别
                Exit Function
            ElseIf .TextMatrix(i, c操作类型) = "" Then
                If .Cell(flexcpData, i, c诊疗类别) = "H" Or .Cell(flexcpData, i, c诊疗类别) = "E" Then
                    MsgBox "请选择操作类型。", vbInformation, gstrSysName
                    .Select i, c操作类型
                    Exit Function
                End If
            ElseIf .TextMatrix(i, c给药分类) = "" Then
                If .TextMatrix(i, c诊疗类别) = "西药中成药" Then
                    MsgBox "请选择给药分类。", vbInformation, gstrSysName
                    .Select i, c给药分类
                    Exit Function
                End If
            End If
        
    
            '检查重复值
            If .TextMatrix(i, c操作类型) = "" Then
                str操作类型 = "Null"
            Else
                str操作类型 = .Cell(flexcpData, i, c操作类型)
            End If
            
            If .TextMatrix(i, c给药分类) = "" Then
                str给药分类 = "Null"
            Else
                str给药分类 = .Cell(flexcpData, i, c给药分类)
            End If
            strKey = .Cell(flexcpData, i, c期效) & "," & .Cell(flexcpData, i, c诊疗类别) & "," & str操作类型 & "," & str给药分类
            
            rsSQL.Filter = "值='" & strKey & "'"
            If rsSQL.RecordCount > 0 Then
                MsgBox "第" & i & "行与第" & rsSQL!行号 & "行的数据重复。", vbInformation, gstrSysName
                .Cell(flexcpBackColor, Val(rsSQL!行号), c顺序, Val(rsSQL!行号), .Cols - 1) = &H80C0FF
                .Select i, c期效
                Exit Function
            Else
                rsSQL.AddNew
                rsSQL!行号 = i
                rsSQL!值 = strKey
                rsSQL.Update
            End If
        Next
    End With
    CheckData = True
End Function

Private Sub SaveData()
    Dim strSql As String
    Dim i As Long, str操作类型 As String, str给药分类 As String
    Dim colSQL As New Collection, blnTrans As Boolean, blnSetup As Boolean
    Dim intOnlyDel As Integer
    Dim strTmp As String
    
    On Error GoTo errH
    With vsItem
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, c操作类型) = "" Then
                str操作类型 = "Null"
            Else
                str操作类型 = "'" & .Cell(flexcpData, i, c操作类型) & "'"
            End If
            
            If .TextMatrix(i, c给药分类) = "" Then
                str给药分类 = "Null"
            Else
                str给药分类 = .Cell(flexcpData, i, c给药分类)
            End If

            If vsItem.Rows = 2 And vsItem.TextMatrix(1, CNAME.c期效) = "" And vsItem.TextMatrix(1, CNAME.c诊疗类别) = "" _
                    And vsItem.TextMatrix(1, CNAME.c操作类型) = "" And vsItem.TextMatrix(1, CNAME.c给药分类) = "" Then
                intOnlyDel = 1
            Else
                intOnlyDel = 0
            End If
            strSql = "Zl_路径项目顺序_Insert(" & .TextMatrix(i, c顺序) & "," & _
                IIf(.Cell(flexcpData, i, c期效) = "", "null", .Cell(flexcpData, i, c期效)) & _
                ",'" & .Cell(flexcpData, i, c诊疗类别) & "'," & str操作类型 & "," & str给药分类 & "," & _
                intOnlyDel & ")"
            colSQL.Add strSql, "C" & colSQL.count + 1
        Next
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    
    '保存参数设置
    If mbytFun = 0 Then
        blnSetup = InStr(GetInsidePrivs(p临床路径应用), ";参数设置;") > 0
        Call zlDatabase.SetPara("是否启用路径执行环节", chkInvocation.Value, glngSys, p临床路径应用, blnSetup)
        Call zlDatabase.SetPara("路径执行环节启用场合", chkDoctor.Value & chkNurse.Value, glngSys, p临床路径应用, blnSetup)
      
        Call zlDatabase.SetPara("医技医嘱在路径表外", chkOutTable.Value, glngSys, p临床路径应用, blnSetup)

        Call zlDatabase.SetPara("路径医嘱生成超前天数", txtSendAdvice.Text, glngSys, p临床路径应用, blnSetup)
        Call zlDatabase.SetPara("中药配方允许修改的中药味数上限", txt中药味数.Text, glngSys, p临床路径应用, blnSetup)
        i = IIf(optPrintDay(0).Value, Val(optPrintDay(0).Caption), Val(optPrintDay(1).Caption))
        Call zlDatabase.SetPara("路径表单每页打印的天数", i & "", glngSys, p临床路径应用, blnSetup)
        Call zlDatabase.SetPara("未评估时允许添加医嘱到昨天", chk未评估时允许添加医嘱到昨天.Value, glngSys, p临床路径应用, blnSetup)
        Call zlDatabase.SetPara("允许前一天不评估就生成今天的路径项目", chkIsEvaluate.Value, glngSys, p临床路径应用, blnSetup)
        Call zlDatabase.SetPara("允许提前生成明天的路径项目", chkIsForwardSend.Value, glngSys, p临床路径应用, blnSetup)
        i = IIf(optPrtRule(0).Value, 0, 1)
        Call zlDatabase.SetPara("路径表单打印规则", i & "", glngSys, p临床路径应用, blnSetup)
    End If
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim blnParSet As Boolean
    Dim lngDays As Long
    Dim strTmp As String
    
    If mbytFun = 0 Then
        Me.Caption = "路径选项"
        lblInfo.Caption = "医嘱类路径项目生成顺序"
        lblNote.Caption = "    路径项目生成医嘱时缺省按路径表中该阶段定义的分类及项目顺序列出，但优先按下表定义的类别顺序排列，每次生成医嘱时也可以调整顺序。"
    Else
        Me.Caption = "医嘱排序设置"
        lblInfo.Caption = "医嘱下达后自动排序"
        lblNote.Caption = "    医嘱保存前，对本次新开的医嘱，自动按下表定义的类别顺序排列，保存后也可以使用医嘱顺序调整功能重新排列顺序。"
        Frame1.Visible = False
        picPath.Visible = False
        vsItem.Height = vsItem.Height + 1200
    End If
    
    picAddRow.Visible = False
    Call InitData
    Call LoadData
    
    If mbytFun = 0 Then
        blnParSet = InStr(GetInsidePrivs(p临床路径应用), ";参数设置;")

        chkInvocation.Value = zlDatabase.GetPara("是否启用路径执行环节", glngSys, p临床路径应用, "1", Array(chkInvocation), blnParSet)
        If chkInvocation.Value = Checked Then
            strTmp = zlDatabase.GetPara("路径执行环节启用场合", glngSys, p临床路径应用, "11", Array(chkDoctor, chkNurse), blnParSet)
            chkDoctor.Value = Val(Mid(strTmp, 1, 1))
            chkNurse.Value = Val(Mid(strTmp, 2, 1))
        Else
            chkDoctor.Enabled = False
            chkNurse.Enabled = False
        End If
        chkOutTable.Value = zlDatabase.GetPara("医技医嘱在路径表外", glngSys, p临床路径应用, "0", Array(chkOutTable), blnParSet)
        lngDays = Val(zlDatabase.GetPara("路径表单每页打印的天数", glngSys, p临床路径应用, "2", Array(lblPrintDays, optPrintDay(0), optPrintDay(1)), blnParSet))
        If lngDays = 2 Then
            optPrintDay(0).Value = True
        Else
            optPrintDay(1).Value = True
        End If
        strTmp = zlDatabase.GetPara("路径表单打印规则", glngSys, p临床路径应用, "0", Array(lblPrtRule, optPrtRule(0), optPrtRule(1)), blnParSet)
        optPrtRule(Val(strTmp)).Value = True
        txtSendAdvice.Text = zlDatabase.GetPara("路径医嘱生成超前天数", glngSys, p临床路径应用, "1", Array(lblSendAdvice, txtSendAdvice), blnParSet)
        chkIsEvaluate.Value = zlDatabase.GetPara("允许前一天不评估就生成今天的路径项目", glngSys, p临床路径应用, "1", Array(chkIsEvaluate), blnParSet)
        chkIsForwardSend.Value = zlDatabase.GetPara("允许提前生成明天的路径项目", glngSys, p临床路径应用, "0", Array(chkIsForwardSend), blnParSet)
        chk未评估时允许添加医嘱到昨天.Value = zlDatabase.GetPara("未评估时允许添加医嘱到昨天", glngSys, p临床路径应用, "1", Array(chk未评估时允许添加医嘱到昨天), blnParSet)
        txt中药味数.Text = zlDatabase.GetPara("中药配方允许修改的中药味数上限", glngSys, p临床路径应用, "30", Array(lbl中药味数, txt中药味数), blnParSet)
        txt中药味数.Tag = txt中药味数.Text
        If chkIsEvaluate.Value = Unchecked Then
            chkIsForwardSend.Value = Unchecked
            chkIsForwardSend.Enabled = False
        End If
    End If
    '将焦点移到第一行
    If vsItem.Rows > 0 Then vsItem.Row = 1
End Sub

Private Sub LoadData()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    
    On Error GoTo errH
    strSql = "Select a.顺序,a.医嘱期效,a.诊疗类别 as 类别编码,a.执行分类,a.操作类型,b.名称 as 类别名称 From 路径项目顺序 a,诊疗项目类别 b " & _
        "Where a.诊疗类别 = b.编码 Order by a.顺序"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "路径项目顺序")
    
    If rsTmp.RecordCount = 0 Then vsItem.TextMatrix(1, c顺序) = 1: Exit Sub
    
    With vsItem
        .Redraw = False
        .Rows = .FixedRows + rsTmp.RecordCount
        i = .FixedRows
        While Not rsTmp.EOF
            .TextMatrix(i, c顺序) = i
            .TextMatrix(i, c期效) = IIf(rsTmp!医嘱期效 = 0, "长嘱", "临嘱")
            .Cell(flexcpData, i, c期效) = Val(rsTmp!医嘱期效)
            
            .Cell(flexcpData, i, c诊疗类别) = CStr(rsTmp!类别编码)  '药品还是存为治疗类的编码E
            
            If rsTmp!类别编码 = "E" And Val("" & rsTmp!操作类型) = 2 Then
                .TextMatrix(i, c诊疗类别) = "西药中成药"
                
            ElseIf rsTmp!类别编码 = "E" And Val("" & rsTmp!操作类型) = 4 Then
                .TextMatrix(i, c诊疗类别) = "中草药"
                
            Else
                .TextMatrix(i, c诊疗类别) = rsTmp!类别名称
            End If
            
            '简化只支持：治疗类：0-普通;1-过敏试验;2-给药方法(西药);3-中药煎法;4-中药用(服)法;5-特殊治疗;6-采集方法(检验);7-配血方法(血库);8-输血途径；
            '            护理类：0-护理常规；1-护理等级；
            If Not IsNull(rsTmp!操作类型) And (rsTmp!类别编码 = "H" Or rsTmp!类别编码 = "E") Then
                If rsTmp!类别编码 = "H" Then
                    .TextMatrix(i, c操作类型) = IIf(rsTmp!操作类型 = 0, "护理常规", "护理等级")
                Else
                     .TextMatrix(i, c操作类型) = Choose(Val(rsTmp!操作类型) + 1, "普通", "过敏试验", "给药方法", "中药煎法", "中药用法", "特殊治疗", "采集方法", "配血方法", "输血途径")
                End If
                .Cell(flexcpData, i, c操作类型) = Val(rsTmp!操作类型)
            End If
            
            If Not IsNull(rsTmp!执行分类) Then
                .TextMatrix(i, c给药分类) = Choose(rsTmp!执行分类 + 1, "其他", "输液", "注射", "皮试", "口服")
                .Cell(flexcpData, i, c给药分类) = Val("" & rsTmp!执行分类)
            End If
            
            i = i + 1
            rsTmp.MoveNext
        Wend
        .Redraw = True
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitData()
    Dim rsTmp As ADODB.Recordset, strTmp As String
    
    Set rsTmp = Get诊疗类别
    strTmp = "#E;西药中成药|#E;中草药"  '固定名称，按治疗类存储
    While Not rsTmp.EOF
        strTmp = strTmp & "|#" & rsTmp!编码 & ";" & rsTmp!名称
        rsTmp.MoveNext
    Wend
    
    With vsItem
        
        .ColComboList(c期效) = "#1;临嘱|#0;长嘱"
        .ColComboList(c诊疗类别) = strTmp
        .ColComboList(c给药分类) = "#0;其他|#1;输液|#2;注射|#3;皮试|#4;口服"
        .Rows = .FixedRows
        .Rows = .FixedRows + 1 '初始一空行
    End With
End Sub

Private Function Get诊疗类别() As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not In('5','6','7')"
    Set Get诊疗类别 = zlDatabase.OpenSQLRecord(strSql, "诊疗类别")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub picAddRow_Click()
    Dim i As Long
    
    If vsItem.Row = vsItem.Rows - 1 Then
        vsItem.Rows = vsItem.Rows + 1
        vsItem.TextMatrix(vsItem.Rows - 1, c顺序) = vsItem.Rows - 1
        vsItem.Select vsItem.Rows - 1, c期效
    Else
        i = vsItem.Row
        vsItem.AddItem "", i
        Call Reset序号
        vsItem.Select i, c期效
    End If
    
End Sub

Private Sub Reset序号()
    Dim i As Long
    
    For i = vsItem.FixedRows To vsItem.Rows - 1
        vsItem.TextMatrix(i, c顺序) = i
    Next
End Sub

Private Sub txtSendAdvice_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txt中药味数_GotFocus()
    zlControl.TxtSelAll txt中药味数
End Sub

Private Sub txt中药味数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txt中药味数_Validate(Cancel As Boolean)
    If Val(txt中药味数.Text) > 100 Then
        MsgBox "中药配方允许修改的中药味数上限比例只能是0-100之间的数字。", vbInformation, Me.Caption
        Cancel = True
        txt中药味数.Text = txt中药味数.Tag
        zlControl.TxtSelAll txt中药味数
    Else
        txt中药味数.Tag = txt中药味数.Text
    End If
End Sub

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsItem.ComboData = "" Then   '未选择时离开焦点
        vsItem.TextMatrix(Row, Col) = CStr(vsItem.Tag)
        Exit Sub
    End If
    
    With vsItem
        If .Tag <> "" Then
            If .Tag = CStr(.ComboItem) Then Exit Sub
        End If
        .TextMatrix(Row, Col) = .ComboItem
        .Cell(flexcpData, Row, Col) = .ComboData
        
        If Col = c诊疗类别 Then
            If .TextMatrix(Row, c诊疗类别) = "西药中成药" Then
                .TextMatrix(Row, c操作类型) = "给药方法"
                .Cell(flexcpData, Row, c操作类型) = 2
            
            ElseIf .TextMatrix(Row, c诊疗类别) = "中草药" Then
                .TextMatrix(Row, c操作类型) = "中药用法"
                .Cell(flexcpData, Row, c操作类型) = 4
                
            Else
                .TextMatrix(Row, c操作类型) = ""
                .Cell(flexcpData, Row, c操作类型) = ""
            End If
            
            .TextMatrix(Row, c给药分类) = ""
            .Cell(flexcpData, Row, c给药分类) = ""
            
        ElseIf Col = c操作类型 Then
            .TextMatrix(Row, c给药分类) = ""
            .Cell(flexcpData, Row, c给药分类) = ""
        End If
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If (OldRow <> NewRow Or OldRow = NewRow And OldRow = 1) And NewRow > vsItem.FixedRows - 1 Then
        If Me.Visible Then
            If picAddRow.Visible = False Then picAddRow.Visible = True
        End If
        picAddRow.Top = vsItem.Top + vsItem.Cell(flexcpTop, NewRow, c操作) + 30
        picAddRow.Left = vsItem.Left + vsItem.Cell(flexcpLeft, NewRow, c操作) + 120
    End If
End Sub

Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        '简化只支持：治疗类：0-普通;1-过敏试验;2-给药方法(西药);3-中药煎法;4-中药用(服)法;5-特殊治疗;6-采集方法(检验);7-配血方法(血库);8-输血途径；
            '            护理类：0-护理常规；1-护理等级；
    With vsItem
        .Tag = .TextMatrix(Row, Col)  '用于AfterEdit中判断是否改变了值
        If Col = c操作 Then
            Cancel = True
        ElseIf Col = c操作类型 Then '治疗和护理才允许
            If .Cell(flexcpData, Row, c诊疗类别) = "H" Then
                .ComboList = "#0;护理常规|#1;护理等级"
                
            ElseIf .TextMatrix(Row, c诊疗类别) = "西药中成药" Or .TextMatrix(Row, c诊疗类别) = "中草药" Then
                .ComboList = ""
                Cancel = True
            
            ElseIf .Cell(flexcpData, Row, c诊疗类别) = "E" Then
                .ComboList = "#0;普通|#1;过敏试验|#5;特殊治疗"
                
            Else
                .ComboList = ""
                Cancel = True
            End If
        ElseIf Col = c给药分类 Then '药品才允许
            If Not (.TextMatrix(Row, c诊疗类别) = "西药中成药") Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub vsItem_ChangeEdit()
    Call vsItem_AfterEdit(vsItem.Row, vsItem.Col)
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
        If vsItem.Row = 0 Then Exit Sub
        If MsgBox("要删除当前行吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
        With vsItem
        
            If .Rows > 2 Then
                vsItem.RemoveItem vsItem.Row
                Call Reset序号
            ElseIf .Rows = 2 Then
                If .TextMatrix(1, CNAME.c期效) = "" And .TextMatrix(1, CNAME.c诊疗类别) = "" _
                        And .TextMatrix(1, CNAME.c操作类型) = "" And .TextMatrix(1, CNAME.c给药分类) = "" Then
                    MsgBox "没有可删除的行了。", vbInformation, gstrSysName
                Else
                    .TextMatrix(1, CNAME.c期效) = ""
                    .TextMatrix(1, CNAME.c诊疗类别) = ""
                    .TextMatrix(1, CNAME.c操作类型) = ""
                    .TextMatrix(1, CNAME.c给药分类) = ""
                End If
            End If
        
        End With
       
    End If
End Sub

Private Sub EnterNextCell()
   
    With vsItem
        If .Col = .Cols - 1 And .Row < .Rows - 1 Then
            .Select .Row + 1, c期效
        ElseIf .Col < .Cols - 1 Then
            .Col = .Col + 1
        End If
    End With
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call EnterNextCell
    End If
End Sub

Private Sub vsItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsItem.ComboIndex <> -1 Then
            Call vsItem_KeyPress(13)
        End If
    End If
End Sub
