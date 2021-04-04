VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmChildSchemeOps 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2670
      Index           =   0
      Left            =   585
      ScaleHeight     =   2670
      ScaleWidth      =   6090
      TabIndex        =   0
      Top             =   630
      Width           =   6090
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   1
         Left            =   5730
         Picture         =   "frmChildSchemeOps.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "多选，快捷键：F3"
         Top             =   30
         Width           =   345
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   555
         TabIndex        =   2
         Top             =   900
         Width           =   1860
         _cx             =   3281
         _cy             =   2143
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
         GridColor       =   -2147483626
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
   End
End
Attribute VB_Name = "frmChildSchemeOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mlngKey As Long
Private mlngReferKey As Long
Private mfrmMain As Object
Private mbytMode As Byte
Private mblnAllowModify As Boolean
Private mblnDataChanged As Boolean
Private mblnReading As Boolean

Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Event AfterDataChanged()


Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    Set mfrmMain = frmMain

    If ExecuteCommand("初始数据") = False Then Exit Function
    Call ExecuteCommand("控件状态")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    mlngKey = lngKey
    mbytMode = 2
    
    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("控件状态")
    
    If mlngKey > 0 Then
        If ExecuteCommand("读取数据") = False Then Exit Function
    End If

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function NewData(Optional ByVal lngReferKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 1
    
    mlngReferKey = lngReferKey
    If mlngReferKey > 0 Then
        mlngKey = mlngReferKey
        Call ExecuteCommand("读取数据")
        mlngKey = 0
    Else
        Call ExecuteCommand("清空数据")
    End If
    
    Call ExecuteCommand("控件状态")
    
    DataChanged = True
    
    Call LocationGrid(vsf(0))
        
    NewData = True
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************

    
    ValidData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset, ByRef lngKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lngLoop As Long
    Dim strTmp As String

    On Error GoTo errHand

    strSQL = "ZL_方案适用手术_DELETE(" & lngKey & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    
    For lngLoop = 1 To vsf(0).Rows - 1
        If Val(vsf(0).RowData(lngLoop)) > 0 Then
            strSQL = "ZL_方案适用手术_INSERT(" & lngKey & "," & Val(vsf(0).RowData(lngLoop)) & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
        End If
    Next
        
    SaveData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            If mblnAllowModify Then
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
            Else
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            End If
            Call .AppendColumn("手术名称", 2700, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("编码", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            
            If mblnAllowModify Then
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("手术名称"), True, vbVsfEditCommand)
                
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            End If
            
            .AppendRows = True
        End With
        cmd(1).Enabled = mblnAllowModify
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 And mbytMode = 2 Then blnAllowModify = False
        mclsVsf.AllowEdit = blnAllowModify
        cmd(1).Enabled = blnAllowModify
    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
        mclsVsf.ClearGrid
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
        
        mclsVsf.ClearGrid
        
        gstrSQL = "SELECT B.ID,B.名称 As 手术名称,b.编码 FROM 方案适用手术 A,诊疗项目目录 B WHERE A.手术项目ID=B.ID AND A.方案ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)
        
        
    End Select

    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim lngLoop As Long
    
    Select Case Index
    Case 1

        gstrSQL = GetPublicSQL(SQL.手术项目选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)

        If ShowPubSelect(Me, cmd(1), 3, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目多选", "请从下表中选择一个或多个手术项目", rsData, rs, 8790, 4800, True) = 1 Then
            With vsf(0)
                For lngLoop = 1 To rs.RecordCount
                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) = False Then

                        If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1

                        .TextMatrix(.Rows - 1, mclsVsf.ColIndex("手术名称")) = zlCommFun.NVL(rs("名称").Value)
                        .TextMatrix(.Rows - 1, mclsVsf.ColIndex("编码")) = zlCommFun.NVL(rs("编码").Value)
                        .RowData(.Rows - 1) = zlCommFun.NVL(rs("ID").Value, 0)

                        DataChanged = True
                    End If

                    rs.MoveNext
                Next
            End With
        End If

    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        KeyCode = 0
        If cmd(0).Enabled And cmd(0).Visible Then
            Call cmd_Click(0)
        End If
    End If
End Sub

Private Sub Form_Load()
'    picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub mclsVsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width - cmd(1).Width - 30, picPane(Index).Height
        cmd(1).Move vsf(0).Left + vsf(0).Width + 15
        mclsVsf.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    '编辑处理
    Call mclsVsf.AfterEdit(Row, Col)
    
    DataChanged = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf(0).RowData(Row)) <= 0)
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If Col = mclsVsf.ColIndex("手术名称") Then

        gstrSQL = GetPublicSQL(SQL.手术项目选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
'
        If ShowPubSelect(Me, vsf(0), 3, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目选择", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(vsf(0).RowData(Row))) = 1 Then
            If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                Exit Sub
            End If

            vsf(0).EditText = zlCommFun.NVL(rs("名称").Value)
            vsf(0).TextMatrix(Row, mclsVsf.ColIndex("手术名称")) = zlCommFun.NVL(rs("名称").Value)
            vsf(0).TextMatrix(Row, mclsVsf.ColIndex("编码")) = zlCommFun.NVL(rs("编码").Value)
            vsf(0).RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

            DataChanged = True

        End If
        
    End If
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    
    With mclsVsf
        If KeyCode = vbKeyReturn Then
            If Col = .ColIndex("手术名称") Then
                
                If InStr(vsf(0).EditText, "'") > 0 Then
                    KeyCode = 0
                    vsf(0).EditText = ""
                    Exit Sub
                End If

                strText = UCase(vsf(0).EditText)
                bytMode = GetApplyMode(strText)
'
                gstrSQL = GetPublicSQL(SQL.手术项目过滤, bytMode)

                strText = strText & "%"
                If ParamInfo.项目输入匹配方式 = 1 Then
                    strTmp = strText
                Else
                    strTmp = "%" & strText
                End If
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                '
                If ShowPubSelect(Me, vsf(0), 2, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目过滤", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(vsf(0).RowData(Row))) = 1 Then

                    If .CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                        Exit Sub
                    End If

                    vsf(0).EditText = zlCommFun.NVL(rs("名称").Value)
                    .TextMatrix(Row, .ColIndex("手术名称")) = zlCommFun.NVL(rs("名称").Value)
                    .TextMatrix(Row, .ColIndex("编码")) = zlCommFun.NVL(rs("编码").Value)

                    vsf(0).RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

                    DataChanged = True

                Else
                    KeyCode = 0

                    vsf(0).Cell(flexcpData, Row, Col) = vsf(0).Cell(flexcpData, Row, Col)
                    vsf(0).EditText = vsf(0).Cell(flexcpData, Row, Col)
                    .TextMatrix(Row, Col) = vsf(0).Cell(flexcpData, Row, Col)

                End If
            End If
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    '编辑处理
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub



