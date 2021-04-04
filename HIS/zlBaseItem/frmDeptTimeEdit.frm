VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDeptTimeEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "部门上班安排"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   Icon            =   "frmDeptTimeEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8415
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5760
      TabIndex        =   37
      Top             =   3840
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   -210
      TabIndex        =   34
      Top             =   675
      Width           =   9060
   End
   Begin VB.OptionButton OptUnit 
      Caption         =   "10分钟"
      Height          =   180
      Index           =   1
      Left            =   7305
      TabIndex        =   33
      Top             =   450
      Width           =   915
   End
   Begin VB.OptionButton OptUnit 
      Caption         =   "半小时"
      Height          =   180
      Index           =   0
      Left            =   6330
      TabIndex        =   32
      Top             =   450
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   7065
      TabIndex        =   29
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   2490
      TabIndex        =   28
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "休息(&B)"
      Height          =   350
      Left            =   1320
      TabIndex        =   27
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdPlan 
      Caption         =   "安排(&P)"
      Height          =   350
      Left            =   150
      TabIndex        =   26
      Top             =   3840
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgdPlan 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1260
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   4260
      _Version        =   393216
      ForeColor       =   -2147483630
      Rows            =   7
      FixedRows       =   0
      FocusRect       =   0
      ScrollBars      =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblSelColor 
      BackColor       =   &H8000000D&
      Height          =   180
      Left            =   915
      TabIndex        =   36
      Top             =   480
      Width           =   345
   End
   Begin VB.Label lblSelNote 
      AutoSize        =   -1  'True
      Caption         =   "当前选择时间范围："
      Height          =   180
      Left            =   1320
      TabIndex        =   35
      Top             =   495
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   165
      Picture         =   "frmDeptTimeEdit.frx":030A
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "基本刻度单位："
      Height          =   180
      Left            =   6135
      TabIndex        =   31
      Top             =   165
      Width           =   1260
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "部门上班安排"
      Height          =   180
      Left            =   795
      TabIndex        =   30
      Top             =   165
      Width           =   1080
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   24
      X1              =   5925
      X2              =   5925
      Y1              =   825
      Y2              =   1210
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   23
      X1              =   5385
      X2              =   5385
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   22
      X1              =   5175
      X2              =   5175
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   21
      X1              =   4965
      X2              =   4965
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   20
      X1              =   4740
      X2              =   4740
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   19
      X1              =   4530
      X2              =   4530
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   18
      X1              =   4320
      X2              =   4320
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   17
      X1              =   4110
      X2              =   4110
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   16
      X1              =   3690
      X2              =   3690
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   15
      X1              =   3480
      X2              =   3480
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   14
      X1              =   3900
      X2              =   3900
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   13
      X1              =   3270
      X2              =   3270
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   12
      X1              =   2925
      X2              =   2925
      Y1              =   825
      Y2              =   1210
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   11
      X1              =   2715
      X2              =   2715
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   10
      X1              =   2505
      X2              =   2505
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   9
      X1              =   2295
      X2              =   2295
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   8
      X1              =   2085
      X2              =   2085
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   7
      X1              =   1860
      X2              =   1860
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   6
      X1              =   1650
      X2              =   1650
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   5
      X1              =   1440
      X2              =   1440
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   4
      X1              =   1020
      X2              =   1020
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   3
      X1              =   810
      X2              =   810
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   2
      X1              =   1230
      X2              =   1230
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   1
      X1              =   600
      X2              =   600
      Y1              =   1035
      Y2              =   1230
   End
   Begin VB.Line linTime 
      BorderColor     =   &H000080FF&
      Index           =   0
      X1              =   390
      X2              =   390
      Y1              =   825
      Y2              =   1225
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "次日"
      Height          =   180
      Index           =   24
      Left            =   5955
      TabIndex        =   25
      Top             =   1065
      Width           =   360
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "11"
      Height          =   180
      Index           =   23
      Left            =   4890
      TabIndex        =   24
      Top             =   1065
      Width           =   180
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "10"
      Height          =   180
      Index           =   22
      Left            =   4620
      TabIndex        =   23
      Top             =   1065
      Width           =   180
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "9"
      Height          =   180
      Index           =   21
      Left            =   4440
      TabIndex        =   22
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "8"
      Height          =   180
      Index           =   20
      Left            =   4275
      TabIndex        =   21
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "7"
      Height          =   180
      Index           =   19
      Left            =   4095
      TabIndex        =   20
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "6"
      Height          =   180
      Index           =   18
      Left            =   3930
      TabIndex        =   19
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "5"
      Height          =   180
      Index           =   17
      Left            =   3750
      TabIndex        =   18
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   180
      Index           =   16
      Left            =   3570
      TabIndex        =   17
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   180
      Index           =   15
      Left            =   3405
      TabIndex        =   16
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Index           =   14
      Left            =   3225
      TabIndex        =   15
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   180
      Index           =   13
      Left            =   3060
      TabIndex        =   14
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "下午"
      Height          =   180
      Index           =   12
      Left            =   2715
      TabIndex        =   13
      Top             =   1065
      Width           =   360
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "11"
      Height          =   180
      Index           =   11
      Left            =   2415
      TabIndex        =   12
      Top             =   1065
      Width           =   180
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "10"
      Height          =   180
      Index           =   10
      Left            =   2145
      TabIndex        =   11
      Top             =   1065
      Width           =   180
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "9"
      Height          =   180
      Index           =   9
      Left            =   1965
      TabIndex        =   10
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "8"
      Height          =   180
      Index           =   8
      Left            =   1800
      TabIndex        =   9
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "7"
      Height          =   180
      Index           =   7
      Left            =   1620
      TabIndex        =   8
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "6"
      Height          =   180
      Index           =   6
      Left            =   1455
      TabIndex        =   7
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "5"
      Height          =   180
      Index           =   5
      Left            =   1275
      TabIndex        =   6
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   180
      Index           =   4
      Left            =   1095
      TabIndex        =   5
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   180
      Index           =   3
      Left            =   930
      TabIndex        =   4
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Index           =   2
      Left            =   750
      TabIndex        =   3
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   585
      TabIndex        =   2
      Top             =   1065
      Width           =   90
   End
   Begin VB.Label lblAM 
      AutoSize        =   -1  'True
      Caption         =   "上午"
      Height          =   180
      Index           =   0
      Left            =   405
      TabIndex        =   1
      Top             =   1065
      Width           =   360
   End
End
Attribute VB_Name = "frmDeptTimeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intTax As Integer
Dim rsPlan As New ADODB.Recordset
Dim blnAskSave As Boolean

Dim intCnt As Long
Dim intRow As Integer, intCol As Integer
Dim dblTime As Double
Dim strTime As String

Private Sub cmdClose_Click()
    If blnAskSave Then
        If MsgBox("需要保存当前进行的工作安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then cmdSave_Click
    End If
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdPlan_Click()
    UpdatePlan 1
    blnAskSave = True
End Sub

Private Sub cmdSave_Click()
    Dim col安排 As New Collection
    Dim varTemp As Variant
    Dim strFlag As String
    Dim lngCount As Long
    Dim strTimeBegin As String
    
    With Me.fgdPlan
        .redraw = False
        For intRow = 0 To .Rows - 1
            strFlag = "结束"
            For intCol = 1 To .Cols - 1
                .Row = intRow
                .Col = intCol
                dblTime = (intCol - 1) * (24 / (.Cols - 1))
                strTime = Format(Int(dblTime), "00") & ":" & Format(60 * (dblTime - Int(dblTime)), "00") & ":00"
                If strFlag = "结束" And .CellBackColor <> &H80000005 Then
                    strFlag = "开始"
                    strTimeBegin = "to_date('" & strTime & "','HH24:MI:SS')" '开始时间
                End If
                If strFlag = "开始" And .CellBackColor = &H80000005 Then
                    '中间可能有停顿
                    lngCount = lngCount + 1 '安排次数
                    col安排.Add Array(intRow, strTimeBegin, "to_date('" & strTime & "','HH24:MI:SS')"), "C" & lngCount
                    strFlag = "结束"
                End If
            Next
            If strFlag = "开始" Then
                '一直安排到当天结束
                lngCount = lngCount + 1 '安排次数
                col安排.Add Array(intRow, strTimeBegin, "to_date('23:59:59','HH24:MI:SS')"), "C" & lngCount
                strFlag = "结束"
            End If
        Next
        .redraw = True
    End With
    
    On Error GoTo errSave
    SQLTest App.ProductName, Me.Name
    gcnOracle.BeginTrans
        gstrSQL = "zl_部门安排_delete(" & lblDept.Tag & ")"
        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        For intCnt = 1 To lngCount '安排次数
            varTemp = col安排("C" & intCnt)
            gstrSQL = "zl_部门安排_insert(" & lblDept.Tag & "," & varTemp(0) & "," & varTemp(1) & "," & varTemp(2) & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
    gcnOracle.CommitTrans
    Call SQLTest
    blnAskSave = False
    Exit Sub
errSave:
    gcnOracle.RollbackTrans
    If ERRCENTER() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdStop_Click()
    UpdatePlan 0
    blnAskSave = True
End Sub

Private Sub Form_Load()
    Me.lblDept.Tag = Mid(frmDeptTime.tvw.SelectedItem.Key, 2)
    Me.lblDept.Caption = frmDeptTime.tvw.SelectedItem.Text & " 上班时间安排"
    
    On Error GoTo errHandle
    blnAskSave = False
'    If rsPlan.State = adStateOpen Then rsPlan.Close
    gstrSQL = "select 部门id, 星期, 开始时间, 终止时间 from 部门安排 where 部门id=[1] and (to_char(开始时间,'MI') not in ('00','30','59') or to_char(终止时间,'MI') not in ('00','30','59'))"
    Set rsPlan = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.lblDept.Tag))
        
    If rsPlan.EOF Then
        intTax = 2
        Me.OptUnit(0).Value = True
        Me.OptUnit(1).Value = False
    Else
        intTax = 6
        Me.OptUnit(0).Value = False
        Me.OptUnit(1).Value = True
    End If
    RefPlan
    Exit Sub
errHandle:
    If ERRCENTER() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub UpdatePlan(bytPlan As Byte)
    Dim intSRow As Integer, intSCol As Integer
    Dim intERow As Integer, intECol As Integer
    
    With Me.fgdPlan
        .redraw = False
        If .Row <= .RowSel Then
            intSRow = .Row
            intERow = .RowSel
        Else
            intSRow = .RowSel
            intERow = .Row
        End If
        If .Col < .ColSel Then
            intSCol = .Col
            intECol = .ColSel
        Else
            intSCol = .ColSel
            intECol = .Col
        End If
        For intRow = intSRow To intERow
            For intCol = intSCol To intECol
                .Row = intRow
                .Col = intCol
                If bytPlan = 1 Then
                    .CellBackColor = &H80FF&
                Else
                    .CellBackColor = &H80000005
                End If
            Next
        Next
        .redraw = True
    End With
End Sub

Public Sub RefPlan()
    On Error GoTo errHandle
    With Me.fgdPlan
        .redraw = False
        .Clear
        For intCnt = 0 To .Rows - 1
            .RowHeight(intCnt) = (.Height - 45) / .Rows
        Next
        .Width = Me.ScaleWidth - 400
        .Left = Me.ScaleLeft + 100
        .ColAlignmentFixed(0) = 4
        .ColWidth(0) = 600
        .Cols = 24 * intTax + 1
        For intCnt = 1 To .Cols - 1
            .ColWidth(intCnt) = Int((.Width - .ColWidth(0)) / (.Cols - 1))
        Next
        .ColWidth(0) = .Width - .ColWidth(1) * (.Cols - 1) - 45
        
        .TextMatrix(0, 0) = IIF(.ColWidth(0) > 600, "星期", "") & "日"
        .TextMatrix(1, 0) = IIF(.ColWidth(0) > 600, "星期", "") & "一"
        .TextMatrix(2, 0) = IIF(.ColWidth(0) > 600, "星期", "") & "二"
        .TextMatrix(3, 0) = IIF(.ColWidth(0) > 600, "星期", "") & "三"
        .TextMatrix(4, 0) = IIF(.ColWidth(0) > 600, "星期", "") & "四"
        .TextMatrix(5, 0) = IIF(.ColWidth(0) > 600, "星期", "") & "五"
        .TextMatrix(6, 0) = IIF(.ColWidth(0) > 600, "星期", "") & "六"
        
        Me.lblAM(0).Left = .Left + .ColWidth(0) + 15
        Me.linTime(0).X1 = Me.lblAM(0).Left - 15
        Me.linTime(0).X2 = Me.linTime(0).X1
            
        For intCnt = 1 To 24
            Me.lblAM(intCnt).Left = Me.lblAM(intCnt - 1).Left + .ColWidth(1) * (.Cols - 1) / 24
            Me.linTime(intCnt).X1 = Me.lblAM(intCnt).Left - 15
            Me.linTime(intCnt).X2 = Me.linTime(intCnt).X1
        Next
        For intCnt = 0 To 24
            Me.linTime(intCnt).Y2 = .Top
            Me.linTime(intCnt).Y1 = .Top - (Me.linTime(intCnt).Y2 - Me.linTime(intCnt).Y1)
            Me.lblAM(intCnt).Top = Me.linTime(intCnt).Y1
        Next
        
        gstrSQL = "select 部门id, 星期, 开始时间, 终止时间 from 部门安排 where 部门id=[1] order by 星期,开始时间"
        Set rsPlan = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.lblDept.Tag))
                
        Do While Not rsPlan.EOF
            dblTime = Hour(rsPlan!开始时间) + Minute(rsPlan!开始时间) / 60
            .Row = rsPlan!星期
            .Col = dblTime / (24 / (.Cols - 1)) + 1

            .RowSel = .RowSel
            If Format(rsPlan!终止时间, "HH:MM:SS") = "23:59:59" Then
                .ColSel = .Cols - 1
            Else
                dblTime = Hour(rsPlan!终止时间) + Minute(rsPlan!终止时间) / 60
                .ColSel = dblTime / (24 / (.Cols - 1))
            End If
            UpdatePlan 1
            .redraw = False
            rsPlan.MoveNext
        Loop
        .redraw = True
    End With
    Exit Sub
errHandle:
    If ERRCENTER() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub fgdPlan_SelChange()
    Dim intSCol As Integer
    Dim intECol As Integer
    
    With Me.fgdPlan
        If .Col < .ColSel Then
            intSCol = .Col
            intECol = .ColSel
        Else
            intSCol = .ColSel
            intECol = .Col
        End If
        dblTime = (intSCol - 1) * (24 / (.Cols - 1))
        strTime = Format(Int(dblTime), "00") & ":" & Format(60 * (dblTime - Int(dblTime)), "00") & ":00"
        Me.lblSelNote = "当前选择时间范围： " & strTime
        
        dblTime = intECol * (24 / (.Cols - 1))
        strTime = Format(Int(dblTime), "00") & ":" & Format(60 * (dblTime - Int(dblTime)), "00") & ":00"
        Me.lblSelNote = Me.lblSelNote & "—" & strTime
    End With
End Sub

Private Sub OptUnit_Click(Index As Integer)
    If blnAskSave Then
        If MsgBox("需要保存当前进行的工作安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then cmdSave_Click
    End If
    If OptUnit(0).Value Then
        intTax = 2
    Else
        intTax = 6
    End If
    RefPlan
End Sub
