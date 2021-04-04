VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLeaveSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医嘱药品选择"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   Icon            =   "frmLeaveSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdAll 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Ctrl+A"
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清(&R)"
      Height          =   350
      Left            =   1350
      TabIndex        =   3
      ToolTipText     =   "Ctrl+R"
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7920
      TabIndex        =   2
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9030
      TabIndex        =   1
      Top             =   5520
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   5325
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   10440
      _cx             =   18415
      _cy             =   9393
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLeaveSelect.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
Attribute VB_Name = "frmLeaveSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As frmLeaveMediMana
Private mLeaveRecord As New ADODB.Recordset

Public Function LeaveSelect(ByVal frmMain As frmLeaveMediMana, ByVal strSQL As String)
    '
    Dim i As Integer
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    Set mLeaveRecord = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not mLeaveRecord.EOF Then
        Call FillVsList
        Me.Show vbModal, frmMain
    Else
        MsgBox "没有可供选择的内容！", vbInformation, gstrSysName
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdAll_Click()
    Dim lngRow As Long
    With vsList
        For lngRow = .FixedRows To .Rows - 1
            If Not (Val(.TextMatrix(.Rows - 1, .ColIndex("数量"))) = 0 Or Trim(.TextMatrix(.Rows - 1, .ColIndex("药品名称与编码"))) = "") Then
                vsList.Cell(flexcpChecked, lngRow, 0) = flexChecked
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    vsList.Cell(flexcpChecked, 1, 0, vsList.Rows - 1, 0) = flexUnchecked
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer, strItem As String
    Dim blnAdd As Boolean
    Dim lngRow As Long
    
    lngRow = mfrmMain.vsList.Row
    For i = 1 To vsList.Rows - 1
        If vsList.Cell(flexcpChecked, i, 0) = flexChecked Then
        
               strItem = "医嘱" & vbTab & vsList.TextMatrix(i, vsList.ColIndex("药品名称与编码")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("规格")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("用途")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("数量")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("计算单位")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("单价")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("金额")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("药品ID")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("医嘱ID")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("发送号")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("剂量单位")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("剂量系数")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("门诊单位")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("门诊包装")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("容量")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("可存数量"))
                mfrmMain.vsList.AddItem strItem, mfrmMain.vsList.Rows - 1
                mfrmMain.vsList.Select mfrmMain.vsList.Row + 1, 0
            blnAdd = True
        End If
    Next
    
    If blnAdd = True Then
        Unload Me
    End If
End Sub

Private Sub FillVsList()
    Dim strHead As String
    Dim lngLast相关ID As Long, cur已用数量 As Currency, cur已执行次  As Currency, date日期 As Date
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo hErr
    strHead = ",300,1;发送时间,1500,1;NO,900,1;药品名称与编码,2500,1;规格,1600,1;用途,550,1;数量,750,7;计算单位,450,4;单价,750,7;金额,1000,7;" & _
              "药品ID,0,1;医嘱ID,0,1;发送号,0,1;剂量单位,0,1;剂量系数,0,1;门诊单位,0,1;门诊包装,0,1;容量,0,1;可存数量,0,1"
    Call SetVsFlexGridHead(strHead, vsList)
    With vsList
        '合并单元格
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(.ColIndex("发送时间")) = True
        .MergeCol(.ColIndex("NO")) = True
        .AutoSize 1, .ColIndex("金额")
    End With
    
    Do Until mLeaveRecord.EOF
        With vsList
            
            
            '求可用数量
            cur已用数量 = 0
            date日期 = zlDatabase.Currentdate
            strSQL = "Select Min(登记时间) as 日期 From 暂存药品记录 Where 医嘱ID=[1] And 发送号=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & mLeaveRecord.Fields("医嘱ID")), Val("" & mLeaveRecord.Fields("发送号")))
            Do Until rsTmp.EOF
                date日期 = IIf(IsNull(rsTmp!日期), zlDatabase.Currentdate, rsTmp!日期)
                rsTmp.MoveNext
            Loop
            
            strSQL = "Select Sum(Nvl(A.本次数次, 0)) As 已用数次" & vbNewLine & _
                    "From  病人医嘱执行 A" & vbNewLine & _
                    "Where A.医嘱id = [1] And A.发送号 = [2] And A.执行时间  < [3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & mLeaveRecord.Fields("相关ID")), Val("" & mLeaveRecord.Fields("发送号")), date日期)
            Do Until rsTmp.EOF
                cur已执行次 = Val("" & rsTmp!已用数次)
                rsTmp.MoveNext
            Loop
'            cur已用数量 = (cur已执行次 * Val(mLeaveRecord.Fields("单次用量"))) / Val("" & mLeaveRecord.Fields("剂量系数"))
            If zlCommFun.NVL(mLeaveRecord.Fields("门诊可否分零"), 0) = 0 Then
                cur已用数量 = (cur已执行次 * Val(mLeaveRecord.Fields("单次用量"))) / Val("" & mLeaveRecord.Fields("剂量系数"))
            Else
                '门诊不可分零，Abs函数是向上取整
                cur已用数量 = cur已执行次 * Abs(Int(0 - Val(mLeaveRecord.Fields("单次用量")) / Val("" & mLeaveRecord.Fields("剂量系数"))))
            End If
            
            If Val("" & mLeaveRecord.Fields("可存数量")) - cur已用数量 - Val(mLeaveRecord.Fields("已存数量")) > 0 Then
                If lngLast相关ID <> 0 And lngLast相关ID <> Val(mLeaveRecord.Fields("相关ID")) Then
                    .AddItem ""
                    .RowHidden(.Rows - 2) = True
                End If
                lngLast相关ID = Val(mLeaveRecord.Fields("相关ID"))
                .TextMatrix(.Rows - 1, .ColIndex("发送时间")) = Format(mLeaveRecord.Fields("发送时间"), "yy-MM-dd hh:mm")
                .TextMatrix(.Rows - 1, .ColIndex("NO")) = mLeaveRecord.Fields("NO")
                .TextMatrix(.Rows - 1, .ColIndex("药品名称与编码")) = "[" & mLeaveRecord.Fields("编码") & "]" & mLeaveRecord.Fields("名称")
                .TextMatrix(.Rows - 1, .ColIndex("规格")) = mLeaveRecord.Fields("规格")
                Select Case mLeaveRecord.Fields("用途")
                    Case 1
                        .TextMatrix(.Rows - 1, .ColIndex("用途")) = "输液"
                    Case 2
                        .TextMatrix(.Rows - 1, .ColIndex("用途")) = "注射"
                    Case 3
                        .TextMatrix(.Rows - 1, .ColIndex("用途")) = "皮试"
                    Case Else
                        .TextMatrix(.Rows - 1, .ColIndex("用途")) = "治疗"
                End Select

                
                .TextMatrix(.Rows - 1, .ColIndex("数量")) = Val("" & mLeaveRecord.Fields("可存数量")) - cur已用数量 - Val(mLeaveRecord.Fields("已存数量"))
                .TextMatrix(.Rows - 1, .ColIndex("计算单位")) = "" & mLeaveRecord.Fields("计算单位")
                .TextMatrix(.Rows - 1, .ColIndex("单价")) = Format(Val("" & mLeaveRecord.Fields("现价")), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format((Val("" & mLeaveRecord.Fields("可存数量")) - cur已用数量 - Val("" & mLeaveRecord.Fields("已存数量"))) * Val("" & mLeaveRecord.Fields("现价")), "0.00")
                
                .TextMatrix(.Rows - 1, .ColIndex("药品ID")) = "" & mLeaveRecord.Fields("收费细目ID")
                .TextMatrix(.Rows - 1, .ColIndex("医嘱ID")) = "" & mLeaveRecord.Fields("医嘱ID")
                .TextMatrix(.Rows - 1, .ColIndex("发送号")) = "" & mLeaveRecord.Fields("发送号")
                .TextMatrix(.Rows - 1, .ColIndex("剂量单位")) = "" & mLeaveRecord.Fields("剂量单位")
                .TextMatrix(.Rows - 1, .ColIndex("剂量系数")) = "" & mLeaveRecord.Fields("剂量系数")
                .TextMatrix(.Rows - 1, .ColIndex("门诊单位")) = "" & mLeaveRecord.Fields("门诊单位")
                .TextMatrix(.Rows - 1, .ColIndex("门诊包装")) = "" & mLeaveRecord.Fields("门诊包装")
                .TextMatrix(.Rows - 1, .ColIndex("容量")) = Val("" & mLeaveRecord.Fields("容量") + 0)
                .TextMatrix(.Rows - 1, .ColIndex("可存数量")) = Val("" & mLeaveRecord.Fields("可存数量")) - cur已用数量 - Val("" & mLeaveRecord.Fields("已存数量"))
                .AddItem ""
            End If
        End With
        mLeaveRecord.MoveNext
    Loop
    If vsList.Rows > 2 Then
        vsList.RemoveItem (vsList.Rows - 1)
    End If
    vsList.Cell(flexcpChecked, 1, 0, vsList.Rows - 1, 0) = flexUnchecked '全部设为未选
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_Click()
    If vsList.MouseCol = 0 Then
        With vsList
            If Val(.TextMatrix(.Rows - 1, .ColIndex("数量"))) = 0 Or Trim(.TextMatrix(.Rows - 1, .ColIndex("药品名称与编码"))) = "" Then Exit Sub
        End With
        vsList.Cell(flexcpChecked, vsList.Row, 0) = IIf(vsList.Cell(flexcpChecked, vsList.Row, 0) = flexUnchecked, flexChecked, flexUnchecked)
    End If
End Sub



