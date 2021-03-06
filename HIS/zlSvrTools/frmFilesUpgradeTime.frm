VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmFilesUpgradeTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "预升级时间设置"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7800
   Icon            =   "frmFilesUpgradeTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraBounds 
      Height          =   30
      Index           =   1
      Left            =   -315
      TabIndex        =   11
      Top             =   1005
      Width           =   8355
   End
   Begin VB.Frame fraBounds 
      Height          =   30
      Index           =   0
      Left            =   -540
      TabIndex        =   10
      Top             =   3765
      Width           =   8625
   End
   Begin VB.CommandButton Cmd添加 
      Caption         =   "添加(&A)"
      Height          =   300
      Left            =   3315
      TabIndex        =   2
      Top             =   3930
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   7800
      TabIndex        =   6
      Top             =   0
      Width           =   7800
      Begin VB.Image imgCaption 
         Height          =   720
         Left            =   405
         Picture         =   "frmFilesUpgradeTime.frx":6852
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "应用：保存应用时间点，对客户端自动分配预升级时间点"
         Height          =   225
         Index           =   0
         Left            =   1485
         TabIndex        =   9
         Top             =   405
         Width           =   5070
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "添加：添加一个新的预升级时间点"
         Height          =   180
         Index           =   1
         Left            =   1485
         TabIndex        =   8
         Top             =   135
         Width           =   2700
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间点：整点的一个小时内客户端预升级，如12:30时间点实际是12:00-12:59"
         Height          =   180
         Index           =   2
         Left            =   1485
         TabIndex        =   7
         Top             =   690
         Width           =   6120
      End
   End
   Begin VB.CheckBox chkParameter 
      Caption         =   "对未设置时间点的客户端生效"
      Height          =   285
      Left            =   75
      TabIndex        =   5
      Top             =   3945
      Width           =   3105
   End
   Begin VB.CommandButton Cmd删除 
      Caption         =   "删除(&D)"
      Height          =   300
      Left            =   4455
      TabIndex        =   3
      Top             =   3930
      Width           =   1000
   End
   Begin VB.CommandButton cmd保存 
      Caption         =   "应用(&S)"
      Height          =   300
      Left            =   5595
      TabIndex        =   1
      Top             =   3930
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   6735
      TabIndex        =   0
      Top             =   3930
      Width           =   1000
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfShift 
      Height          =   2610
      Left            =   15
      TabIndex        =   4
      Top             =   1095
      Width           =   7750
      _cx             =   13670
      _cy             =   4604
      Appearance      =   1
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483626
      GridColorFixed  =   -2147483632
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
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
End
Attribute VB_Name = "frmFilesUpgradeTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOk As Boolean
Private WithEvents mclsVsfShift As clsVsf
Attribute mclsVsfShift.VB_VarHelpID = -1
Private mstrOldTimes As String

Private Sub chkParameter_Click()
    If chkParameter.value = 1 Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "chk参数", "1"
    ElseIf chkParameter.value = 0 Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "chk参数", "0"
    End If
End Sub

'关闭
Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub Cmd保存_Click()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngLoop As Long
    Dim strTemp As String
    On Error GoTo errHand
    
    For lngLoop = 1 To vsfShift.Rows - 1
        If Len(strTemp) = 0 Then
            If vsfShift.TextMatrix(lngLoop, 1) <> "" Then
                strTemp = vsfShift.TextMatrix(lngLoop, 1)
            End If
        Else
            If vsfShift.TextMatrix(lngLoop, 1) <> "" Then
                strTemp = strTemp & "," & Format(vsfShift.TextMatrix(lngLoop, 1), "HH:mm")
            End If
        End If
    Next
    
    If strTemp <> mstrOldTimes Then
        Set rsTmp = New ADODB.Recordset
        gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 = '客户端预升级时间点'"
        Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        
        If rsTmp.EOF = False Then
            strSQL = "Update zlRegInfo Set 内容='" & strTemp & "' Where 项目='客户端预升级时间点'"
            gcnOracle.Execute strSQL
        Else
            strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('客户端预升级时间点',Null,'" & strTemp & "')"
            gcnOracle.Execute strSQL
        End If

        If chkParameter.value = 1 Then
            If MsgBox("是否只对未设置预升级时间的客户端设置预升级时间点？", vbOKCancel + vbQuestion, gstrSysName) = vbOK Then
                Call ExecuteProcedure("Zl_Zlclients_SetPRETime(1,0)", Me.Caption)
            Else
                Exit Sub
            End If
        Else
            If MsgBox("是否对所有客户端设置预升级时间点？该操作会覆盖已设置过预升级时间点的客户端！", vbOKCancel + vbQuestion, gstrSysName) = vbOK Then
                Call ExecuteProcedure("Zl_Zlclients_SetPRETime(1,1)", Me.Caption)
            Else
                Exit Sub
            End If
        End If

        If MsgBox("是否重新对所有客户端进行预升级?" & vbNewLine & "是:所有客户端预升级的完成状态将被清除。" & vbNewLine & "否:不清除预升级完成状态的标志。", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
'            strSQL = "Zl_Zlclients_Control(3,Null,Null,Null,Null,Null,0)"
            strSQL = "Zl_Zlclients_SetPRETime(3)"
            Call ExecuteProcedure(strSQL, Me.Caption)
        End If
        
        mblnOk = True
    Else
        mblnOk = False
    End If
    
    Unload Me
  Exit Sub
errHand:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub Cmd删除_Click()
    If vsfShift.Rows > 1 Then
        If vsfShift.Row > 1 Then
            Call mclsVsfShift.DeleteRow(vsfShift.Row)
        Else
            Call mclsVsfShift.DeleteRow(vsfShift.Rows - 1)
        End If
    End If
End Sub

Private Sub Cmd添加_Click()
    If vsfShift.Rows = 2 And vsfShift.TextMatrix(1, 1) = "" Then
        vsfShift.TextMatrix(1, 0) = "1"
        vsfShift.TextMatrix(1, 1) = Format("12:00", "HH:mm")
    Else
        Call mclsVsfShift.AutoAddRow(vsfShift.MouseRow, vsfShift.MouseCol)
    End If
    vsfShift.ComboList = ""
End Sub

Private Sub Form_Load()
    Call InitVSF
    Call LoadVSF
    chkParameter.value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "chk参数", ""))
End Sub

Private Sub Form_Resize()
    With vsfShift
        .Top = 1000
        .Left = 0
        .Width = Me.ScaleWidth + 30
        .Height = 2600
    End With
    mclsVsfShift.AppendRows = True
    
'    With Cmd添加
'        .Top = ScaleHeight - .Height - 30
'        .Left = 15
'    End With
'
'    With Cmd删除
'        .Top = Cmd添加.Top
'        .Left = Cmd添加.Left + Cmd添加.Width + 30
'    End With
    
'    With cmdCancel
'        .Top = Cmd添加.Top
'        .Left = ScaleWidth - .Width - 30
'    End With
'
'    With cmd保存
'        .Top = Cmd添加.Top
'        .Left = cmdCancel.Left - .Width - 30
'    End With
End Sub

Private Sub InitVSF()
     
     Set mclsVsfShift = New clsVsf
     
    With mclsVsfShift
        Call .Initialize(Me.Controls, vsfShift, True, False)
        Call .ClearColumn
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "", False)
        Call .AppendColumn("预设客户端后台升级时间", 1670, flexAlignCenterCenter, flexDTString, "HH:mm", , True)
        
        .AppendRows = True
        .IndicatorMode = 2
        Call .InitializeEdit(True, True, True)
        Call .InitializeEditColumn(.ColIndex("预设客户端后台升级时间"), True, vbVsfEditDate, , , , "99:99")
    End With
End Sub

Private Sub LoadVSF()
    Dim rsTmp As ADODB.Recordset
    Dim strTemp() As String
    Dim i As Integer
    Set rsTmp = New ADODB.Recordset
    
    mstrOldTimes = ""
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 like '客户端预升级时间点'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    With vsfShift
'        .Redraw = flexRDNone
        If rsTmp.RecordCount = 1 Then
            If Nvl(rsTmp!内容) <> "" Then
                strTemp = Split(Nvl(rsTmp!内容), ",")
                .Rows = UBound(strTemp) + 2
                For i = 0 To UBound(strTemp)
                    
                    .TextMatrix(i + 1, 0) = i + 1
                    .TextMatrix(i + 1, 1) = strTemp(i)
                Next
            Else
                For i = 0 To .Cols - 1
                    .TextMatrix(1, i) = ""
                Next
'                .Redraw = flexRDBuffered
                Exit Sub
            End If
        Else
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
'            .Redraw = flexRDBuffered
            Exit Sub
        End If
        For i = 1 To vsfShift.Rows - 1
            If Len(mstrOldTimes) = 0 Then
                If vsfShift.TextMatrix(i, 1) <> "" Then
                    mstrOldTimes = vsfShift.TextMatrix(i, 1)
                End If
            Else
                If vsfShift.TextMatrix(i, 1) <> "" Then
                    mstrOldTimes = mstrOldTimes & "," & Format(vsfShift.TextMatrix(i, 1), "HH:mm")
                End If
            End If
        Next
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsfShift = Nothing
End Sub


Private Sub mclsVsfShift_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    mclsVsfShift.AppendRows = True
End Sub

Private Sub mclsVsfShift_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    mclsVsfShift.AppendRows = True
End Sub

Private Sub mclsVsfShift_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsfShift
            Cancel = (.TextMatrix(Row, .ColIndex("预设客户端后台升级时间")) = "" Or .TextMatrix(Row, .ColIndex("预设客户端后台升级时间")) = "    -  -     :  " Or .TextMatrix(Row, .ColIndex("预设客户端后台升级时间")) = "__:__")
    End With
    mclsVsfShift.AppendRows = True
End Sub

Private Sub vsfShift_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsfShift.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    Dim lngCol As Long
    lngCol = vsfShift.ColIndex("预设客户端后台升级时间")
   
    With vsfShift
        Select Case NewCol
        Case lngCol
            If mclsVsfShift.AllowColEdit(NewCol) = False Or mclsVsfShift.AllowEdit = False Then Exit Sub
            If IsDate(.TextMatrix(NewRow, NewCol)) = False Then
                
                If NewRow > 1 Then
                    If IsDate(.TextMatrix(NewRow - 1, NewCol)) Then
                        .TextMatrix(NewRow, NewCol) = GetUpgradeTime(.TextMatrix(NewRow - 1, NewCol)) & ":00"
                    Else
                        .TextMatrix(NewRow, NewCol) = Format(CurrentDate, "HH:mm")
                    End If
                Else
                    .TextMatrix(NewRow, NewCol) = Format(CurrentDate, "HH:mm")
                End If
            End If
        End Select
    End With
    mclsVsfShift.AppendRows = True
    vsfShift.ComboList = ""
End Sub

Private Sub vsfShift_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsfShift.AppendRows = True
End Sub

Private Sub vsfShift_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsfShift.AppendRows = True
End Sub

Private Sub vsfShift_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsfShift
    
        If mclsVsfShift.CellButtonClick(Row, Col) Then
            Call mclsVsfShift.SetFocus(, , True)
        End If
    End With
End Sub

Private Sub vsfShift_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsfShift.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfShift_KeyPress(KeyAscii As Integer)
    '编辑处理
    Call mclsVsfShift.KeyPress(KeyAscii)
End Sub

Private Sub vsfShift_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Call mclsVsfShift.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfShift_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
    Case 1
        Call mclsVsfShift.AutoAddRow(vsfShift.MouseRow, vsfShift.MouseCol)
    End Select
End Sub

Private Sub vsfShift_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Call mclsVsfShift.EditSelAll
End Sub

Private Sub vsfShift_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsfShift.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsfShift_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsfShift.ValidateEdit(Col, Cancel)
End Sub

Private Sub vsfShift_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case vsfShift.ColIndex("预设客户端后台升级时间")
        '只有服务器列和升级列才能更改
    Case Else
        '其他列不能更改
        Cancel = True
    End Select
End Sub

Private Function GetUpgradeTime(ByVal strTemp As String) As String
    Dim i As Integer
    Dim strTime As String
    If strTemp = "" Then
        GetUpgradeTime = Format(CurrentDate, "HH:mm")
        Exit Function
    End If
    
    i = InStrRev(strTemp, ":")
    If i > 0 Then
        strTime = Left(strTemp, i)
        strTime = Val(strTime) + 1
        If Val(strTime) >= 24 Then
            strTime = "00"
        End If
        
        GetUpgradeTime = strTime
        Exit Function
    End If
    GetUpgradeTime = Format(CurrentDate, "HH:mm")
End Function
