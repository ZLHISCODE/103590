VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPunctureDeskSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "穿刺台设置"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPunctureDeskSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6645
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar sbarSub 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4650
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11668
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgDesk 
      Height          =   2925
      Left            =   135
      TabIndex        =   6
      Top             =   1500
      Width           =   5040
      _cx             =   8890
      _cy             =   5159
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VB.Frame fraBase 
      Height          =   855
      Left            =   15
      TabIndex        =   0
      Top             =   525
      Width           =   6600
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "是否使用"
         Height          =   195
         Left            =   5400
         TabIndex        =   5
         Top             =   330
         Width           =   1020
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   270
         Width           =   3090
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   0
         Left            =   555
         TabIndex        =   1
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "呼叫器编号"
         Height          =   195
         Index           =   1
         Left            =   1305
         TabIndex        =   3
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "序号"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   315
         Width           =   360
      End
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   225
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPunctureDeskSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintEdit As Integer  '0-查看 1-新增,2-修改,3-删除
Private mlngDeptID As Long  '科室ID
Private mblnOk As Boolean
Private mstrPrivs As String

Public Function ShowMe(ByVal lngDeptID As Long) As Boolean
    mlngDeptID = lngDeptID
    mblnOk = False
    mstrPrivs = gstrPrivs
    Show vbModal
    ShowMe = mblnOk
End Function

Private Sub RefData()
    '刷新数据
    Dim strSQL As String, strErr As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "select ID,科室ID,序号,有效,呼叫器编号 from 门诊穿刺台 Where 科室ID=[1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID)
    Call vfgSetting(0, Me.vfgDesk)
    
    If vfgLoadFromRecord(Me.vfgDesk, rsTmp, strErr) Then
        With vfgDesk
            .ColWidth(.ColIndex("序号")) = 900: .ColHidden(.ColIndex("序号")) = False
            .ColWidth(.ColIndex("呼叫器编号")) = 2800: .ColHidden(.ColIndex("呼叫器编号")) = False
            .ColWidth(.ColIndex("有效")) = 800: .ColHidden(.ColIndex("有效")) = False
            .ColDataType(.ColIndex("有效")) = flexDTBoolean
            
            .ExtendLastCol = True
        End With
    Else
        MsgBox strErr, vbQuestion, Me.Caption
        Exit Sub
    End If
    Call vfgDesk_RowColChange
End Sub

Private Sub LockEdit()
    
    Dim objCtrl As Control
    fraBase.Enabled = Not (mintEdit = 0)
    
    txt(0).Locked = True    '序号不能改
    If mintEdit = 1 Then
        For Each objCtrl In Me.Controls
            If TypeName(objCtrl) = "TextBox" Then
                objCtrl = ""
            End If
        Next
        txt(0) = GetMaxNoAddOne("序号", "门诊穿刺台 Where 科室ID=" & mlngDeptID)
        chk.Value = 1
        txt(0).Locked = False
    End If
    
End Sub
Private Function SaveData() As Boolean
    '保存仪器
    Dim strSeqNo As String, strBeepCode As String, lngID As Long
    Dim intStat As Integer, strSQL As String, rsTmp As ADODB.Recordset
    Dim strErr As String
    
    lngID = Val(txt(0).Tag)
    If mintEdit = 1 Then lngID = Val(GetMaxNoAddOne("ID", "门诊穿刺台"))
    
    strSeqNo = Trim(DelInvalidChar(txt(0)))
    If strSeqNo = "" Then
        strErr = "序号不能为空！"
        MsgBox strErr, vbInformation, Me.Caption
        Exit Function
    End If
    
    If mintEdit = 3 Then
        If lngID <> 0 Then
            If MsgBox("是否删除" & strSeqNo & "号穿刺台？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                
                strSQL = "Select 病人id From 排队记录 Where 穿刺台 = [1] And 科室id = [2] And 日期 > Sysdate - 35 And Rownum < 2"
                Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, strSeqNo, mlngDeptID)
                If rsTmp.EOF Then
                    strSQL = "ZL_门诊穿刺台_Edit(2," & lngID & ")"
                    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
                Else
                    MsgBox strSeqNo & "号穿刺台已使用，不能删除！"
                End If
            End If
        Else
            strErr = "请选择一个项目！"
            MsgBox strErr, vbInformation, Me.Caption
        End If
        Exit Function
    End If
    '检查待保存的数据
    

    strBeepCode = Trim(DelInvalidChar(txt(1)))
    intStat = 0
    If chk.Value = 1 Then intStat = 1
    '


    If mintEdit = 2 And lngID <= 0 Then
        strErr = "请选择一个项目！"
        MsgBox strErr, vbInformation, Me.Caption
        Exit Function
    ElseIf mintEdit = 1 Then
        '新增的序号，检查是否已存在
        strSQL = "Select ID From 门诊穿刺台 Where 序号 = [1] And 科室id = [2]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, strSeqNo, mlngDeptID)
        If Not rsTmp.EOF Then
            MsgBox strSeqNo & "号穿刺台已经存在！", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    
    strSQL = "ZL_门诊穿刺台_Edit(1," & lngID & "," & mlngDeptID & ",'" & strSeqNo & "','" & strBeepCode & "'," & intStat & ")"
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
End Function
Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Edit_Seat_Add      '增加
            mintEdit = 1
            txt(0).Tag = ""
            Call LockEdit
        Case conMenu_Edit_Seat_Delete   '删除
            mintEdit = 3
            Call SaveData
            mintEdit = 0
            Call LockEdit
            RefData
        Case conMenu_Edit_Seat_Modify   '修改
            mintEdit = 2
            Call LockEdit
        Case conMenu_Edit_Transf_Cancle '取消
            mintEdit = 0
            Call LockEdit
        Case conMenu_Edit_Transf_Save   '保存
            If SaveData Then
                mintEdit = 0
                Call LockEdit
                RefData
            End If
        Case conMenu_File_Exit          '退出
            Unload Me
    End Select
End Sub

Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    fraBase.Left = lngLeft + 15
    fraBase.Top = lngTop
    fraBase.Width = lngRight - lngLeft
    
    vfgDesk.Left = lngLeft
    vfgDesk.Top = fraBase.Top + fraBase.Height
    vfgDesk.Width = lngRight - lngLeft
    vfgDesk.Height = lngBottom - lngTop - sbarSub.Height - fraBase.Height
    
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Edit_Seat_Add
            Control.Enabled = mintEdit = 0
        Case conMenu_Edit_Seat_Delete
            Control.Enabled = mintEdit = 0
        Case conMenu_Edit_Seat_Modify
            Control.Enabled = mintEdit = 0
        Case conMenu_Edit_Transf_Cancle
            Control.Enabled = mintEdit <> 0
        Case conMenu_Edit_Transf_Save
            Control.Enabled = mintEdit <> 0
        Case conMenu_File_Exit
            Control.Enabled = mintEdit = 0
    End Select
End Sub

Private Sub Form_Load()
    Dim Menus As New Collection
    Dim strName As String
    
    Menus.Add conMenu_Edit_Seat_Add & ",增加(&A),False"
    Menus.Add conMenu_Edit_Seat_Modify & ",修改(&E),False"
    Menus.Add conMenu_Edit_Seat_Delete & ",删除(&D),False"
    Menus.Add conMenu_Edit_Transf_Cancle & ",取消(&U),True"
    Menus.Add conMenu_Edit_Transf_Save & ",保存(&S),False"
    Menus.Add conMenu_File_Exit & ",退出(&Q),True"
    Call CbsButtonInit(cbsSub, Menus, False, xtpBarTop)
    
    Set Menus = Nothing
    
    Call vfgSetting(0, Me.vfgDesk, "")
    Call LockEdit
    Call RefData
    
    strName = GetDeptName(mlngDeptID)
    If strName <> "" Then
        Caption = Caption & "(" & strName & ")"
    End If
    
End Sub

Private Sub Form_Resize()
    Call cbsSub_Resize
End Sub

Private Sub vfgDesk_RowColChange()
    
    With vfgDesk
        
        If Not (.Row >= .FixedRows And .Row <= .Rows - 1) Then Exit Sub
        If .Cols < 5 Then Exit Sub
        If Val("" & .TextMatrix(.Row, .ColIndex("id"))) <= 0 Then Exit Sub
        
        txt(0) = "" & .TextMatrix(.Row, .ColIndex("序号"))
        txt(0).Tag = Val("" & .TextMatrix(.Row, .ColIndex("id")))
        
        txt(1) = "" & .TextMatrix(.Row, .ColIndex("呼叫器编号"))
        
        If Val("" & .TextMatrix(.Row, .ColIndex("有效"))) = 0 Then
            chk.Value = 0
        Else
            chk.Value = 1
        End If

    End With
End Sub

Private Function GetDeptName(ByVal lngID As Long)
'功能：获取部门名称和编码
'参数：
'  lngID：部门ID

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select 编码 || '-' || 名称 名称 From 部门表 Where ID = [1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "获取部门名称", lngID)
    If rsTmp.EOF = False Then
        GetDeptName = zlcommfun.NVL(rsTmp!名称)
    End If
    rsTmp.Close
    
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function
