VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOpsStationArrange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "手术安排"
   ClientHeight    =   4515
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6600
   Icon            =   "frmOpsStationArrange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5355
      TabIndex        =   2
      Top             =   1350
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5355
      TabIndex        =   0
      Top             =   60
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5355
      TabIndex        =   1
      Top             =   570
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   4500
      Left            =   45
      TabIndex        =   3
      Top             =   -45
      Width           =   5115
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   525
         Width           =   3510
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   1
         Left            =   4635
         Picture         =   "frmOpsStationArrange.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "多选，快捷键：F3"
         Top             =   495
         Width           =   345
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   0
         Left            =   3945
         TabIndex        =   4
         Text            =   "1"
         Top             =   165
         Width           =   450
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Left            =   1125
         TabIndex        =   5
         Top             =   165
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   82182147
         CurrentDate     =   38083
      End
      Begin MSComCtl2.UpDown udp 
         Height          =   300
         Left            =   4395
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   165
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         OrigLeft        =   2040
         OrigTop         =   1155
         OrigRight       =   2280
         OrigBottom      =   1455
         Max             =   12
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   3510
         Left            =   1125
         TabIndex        =   12
         Top             =   885
         Width           =   3870
         _cx             =   6826
         _cy             =   6191
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "开始时间(&T)"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "时长(&H)"
         Height          =   180
         Index           =   1
         Left            =   3285
         TabIndex        =   10
         Top             =   225
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手 术 间(&M)"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "小时"
         Height          =   180
         Left            =   4650
         TabIndex        =   8
         Top             =   210
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手术人员(&R)"
         Height          =   180
         Index           =   3
         Left            =   90
         TabIndex        =   7
         Top             =   930
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmOpsStationArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'（１）窗体级变量定义

Private mblnReading As Boolean
Private mblnDataChanged As Boolean
Private mblnOK As Boolean
Private mlngKey As Long
Private mlngDeptKey As Long
Private mfrmMain As Form
'Private mlngDept As Long
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Form, Optional lngKey As Long = 0, Optional lngDeptKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '功能：打开编辑窗体进行数据的新增、修改操作
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngKey = lngKey
    mlngDeptKey = lngDeptKey
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("初始数据") = False Then Exit Function
    
    Call ExecuteCommand("读取数据")
    
    DataChanged = False
    
    Me.Show 1, mfrmMain
    
    ShowEdit = mblnOK
    
End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：对新增、修改的数据进行合法性校验
    '返回：校验合法返回True，否则返回False
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset

        
    If Val(txt(0).Text) < 1 Or Val(txt(0).Text) > 12 Then
        ShowSimpleMsg "手术时长必须大于1小时而小于12小时！"
        
        zlControl.TxtSelAll txt(0)
        txt(0).SetFocus
        Exit Function
    End If
    
    gstrSQL = "SELECT 1 FROM 医技执行房间 WHERE 执行间=[1] AND 科室id=[2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt(1).Text, mlngDeptKey)
    
    If rs.BOF Then
        ShowSimpleMsg "安排手术间了一个不存在的手术间！"
        zlControl.TxtSelAll txt(1)
        txt(1).SetFocus
        Exit Function
    End If
    
    '检查手术时间与申请时间的关系
    gstrSQL = "SELECT 开嘱时间 FROM 病人医嘱记录 WHERE ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    If rs.BOF = False Then
        If Format(dtp.Value, "yyyy-MM-dd HH:mm") < Format(rs("开嘱时间").Value, "yyyy-MM-dd HH:mm") Then
            
            If MsgBox("手术开始时间(" & Format(dtp.Value, "yyyy-MM-dd HH:mm") & ")早于申请时间(" & Format(rs("开嘱时间").Value, "yyyy-MM-dd HH:mm") & ")" & vbCrLf & "是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                dtp.SetFocus
                Exit Function
            End If
            
        End If
    End If
    
    '检查一个病人是否在同一时间段内做两种手术
    gstrSQL = "SELECT 1 FROM 病人手术记录 B " & _
                "WHERE B.手术状态 In (2,3) AND  " & _
                       "B.医嘱id <> [3] AND  " & _
                       "(B.病人id, NVL(B.主页id,0)) IN (SELECT 病人id, NVL(主页id,0) FROM 病人医嘱记录 WHERE ID = [3]) AND  " & _
                       "((B.手术开始时间 BETWEEN [1] AND [2]) OR  " & _
                       "(B.手术结束时间 BETWEEN [1] AND [2]))"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(DateAdd("h", Val(txt(0).Text), dtp.Value), "YYYY-MM-DD HH:MM:SS")), mlngKey)
    If rs.BOF = False Then
        ShowSimpleMsg "当前病人不能同时进行二场手术。"
        dtp.SetFocus
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Function SaveData() As Boolean
    '******************************************************************************************************************
    '功能：对新增、修改后的数据进行保存/更新处理
    '参数：返回参lngKey，表示更新记录的关键字
    '返回：保存成功返回True，否则返回False
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim rsSQL As ADODB.Recordset
    Dim strTmp As String
    
    Call SQLRecord(rsSQL)
    
    With vsf
        For lngLoop = 1 To .Rows - 1
            If Val(.RowData(lngLoop)) > 0 Then
                strTmp = strTmp & ";" & Val(.RowData(lngLoop)) & "," & .TextMatrix(lngLoop, .ColIndex("岗位")) & "," & .TextMatrix(lngLoop, .ColIndex("姓名")) & "," & .TextMatrix(lngLoop, .ColIndex("编号"))
            End If
        Next
    End With
    
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)

    gstrSQL = "zl_病人手术记录_Arrange(" & mlngKey & ",To_Date('" & Format(dtp.Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'),TO_DATE('" & Format(DateAdd("h", Val(txt(0).Text), dtp.Value), "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'),'" & txt(1).Text & "'," & mlngDeptKey & ",'" & strTmp & "',2)"
    Call SQLRecordAdd(rsSQL, gstrSQL)
    
    SaveData = SQLRecordExecute(rsSQL, Me.Caption)
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf, True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)

            Call .AppendColumn("岗位", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("姓名", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("编号", 900, flexAlignLeftCenter, flexDTString, "", , True)

            Call .InitializeEdit(True, True, True)
            

            Call .InitializeEditColumn(.ColIndex("岗位"), True, vbVsfEditCombox)
            Call .InitializeEditColumn(.ColIndex("姓名"), True, vbVsfEditCommand)
            
            .IndicatorCol = 0
            Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            
            .AppendRows = True
        End With
        txt(1).BackColor = COLOR.锁色
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        gstrSQL = "SELECT 名称 FROM 手术岗位 Order by 编码"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
        Call mclsVsf.InitializeEditColumn(mclsVsf.ColIndex("岗位"), True, vbVsfEditCombox, vsf.BuildComboList(rs, "名称", "名称"))

        
        dtp.Value = Format(zlDatabase.Currentdate + 1, dtp.CustomFormat)
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
        
        mblnReading = True
        
        
        mblnReading = False
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
        
        
    End Select

    ExecuteCommand = True

    Exit Function
    
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 1      '手术执行间
        
        gstrSQL = "Select RowNum As ID,执行间,Decode(b.手术间,Null,'空闲',Decode(b.手术状态,2,'预订',3,'在用')) As 状态" & vbNewLine & _
                    "From 医技执行房间 a," & vbNewLine & _
                    "     (" & vbNewLine & _
                    "      Select 手术间,手术状态" & vbNewLine & _
                    "      From 病人手术记录" & vbNewLine & _
                    "      Where Not (手术结束时间<[2] OR 手术开始时间>[3]) AND 手术室id=[1] AND 手术状态 In (2,3)" & vbNewLine & _
                    "     ) b" & vbNewLine & _
                    "Where a.科室id=[1]" & vbNewLine & _
                    "      And a.执行间=b.手术间(+)"
                        
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngDeptKey, CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(DateAdd("h", Val(txt(0).Text), dtp.Value))))
 
        If ShowPubSelect(Me, txt(1), 2, "执行间,2100,0,;状态,900,0,", Me.Name & "\手术执行间选择", "请从下表中选择一个手术执行间", rsData, rs, 3600, 4200) = 1 Then
            txt(1).Text = zlCommFun.NVL(rs("执行间").Value)
            DataChanged = True
        End If

    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    If ValidData = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    DataChanged = False
    
    Unload Me
End Sub

Private Sub txt_Change(Index As Integer)
    If mblnReading Then Exit Sub
    
    DataChanged = True

End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        zlCommFun.PressKey vbKeyTab

    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub
    
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
    DataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    Dim strTmp As String
    
    With vsf
        If Col = .ColIndex("姓名") Then

            gstrSQL = GetPublicSQL(SQL.人员安排选择)
            
            strTmp = "医生"
            If InStr(.TextMatrix(.Row, .ColIndex("岗位")), "护士") > 0 Then strTmp = "护士"
                
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strTmp, mlngDeptKey, mlngKey)
            bytRet = ShowPubSelect(Me, vsf, 2, "编号,1200,0,;姓名,1200,0,;简码,900,0,;科室,1200,0,;状态,900,0,", Me.Name & "\人员安排选择", "请从下表中选择一个手术人员", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                        
            If bytRet = 1 Then
            
'                If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
'                    ShowSimpleMsg "选择的人员“" & zlCommFun.NVL(rs("姓名").Value) & "”已被选择！"
'                    Exit Sub
'                End If
                       
                .EditText = zlCommFun.NVL(rs("姓名").Value)
                .TextMatrix(Row, .ColIndex("姓名")) = zlCommFun.NVL(rs("姓名").Value)
                .TextMatrix(Row, .ColIndex("编号")) = zlCommFun.NVL(rs("编号").Value)
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                DataChanged = True
    
            End If
            
        End If
    End With
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    Dim bytRet As Byte
    Dim strDoctor As String
    
    With vsf
        If KeyCode = vbKeyReturn Then
            If Col = .ColIndex("姓名") Then
            
                If InStr(.EditText, "'") > 0 Then
                    KeyCode = 0
                    .EditText = ""
                    Exit Sub
                End If

                strText = UCase(.EditText)
                bytMode = GetApplyMode(strText)
                    
                strText = strText & "%"
                strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, strText, "%" & strText)
    
                gstrSQL = GetPublicSQL(SQL.人员安排过滤, bytMode)
                
                strDoctor = "医生"
                If InStr(.TextMatrix(.Row, .ColIndex("岗位")), "护士") > 0 Then strDoctor = "护士"
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strDoctor, mlngDeptKey, mlngKey, strText, strTmp)
    
                If ShowPubSelect(Me, vsf, 2, "编号,1200,0,;姓名,1200,0,;简码,900,0,;科室,1200,0,;状态,900,0,", Me.Name & "\人员安排过滤", "请从下表中选择一个人员", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

'                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
'                        ShowSimpleMsg "选择的人员“" & zlCommFun.NVL(rs("姓名").Value) & "”已被选择！"
'                        Exit Sub
'                    End If
                           
                    .EditText = zlCommFun.NVL(rs("姓名").Value)
                    .TextMatrix(Row, .ColIndex("姓名")) = zlCommFun.NVL(rs("姓名").Value)
                    .TextMatrix(Row, .ColIndex("编号")) = zlCommFun.NVL(rs("编号").Value)
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    DataChanged = True
                Else
                    KeyCode = 0

                    .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                    .EditText = .Cell(flexcpData, Row, Col)
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                End If

            End If
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf.MouseRow, vsf.MouseCol)
    End Select
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub


