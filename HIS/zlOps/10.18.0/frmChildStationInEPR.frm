VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{099B2A6C-9CCE-43CF-AEF0-C526C98F4B7F}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmChildStationInEPR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2670
      Index           =   1
      Left            =   615
      ScaleHeight     =   2670
      ScaleWidth      =   6090
      TabIndex        =   2
      Top             =   3585
      Width           =   6090
      Begin zlRichEditor.Editor edt 
         Height          =   1245
         Left            =   3675
         TabIndex        =   3
         Top             =   765
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   2196
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1785
      Index           =   0
      Left            =   945
      ScaleHeight     =   1785
      ScaleWidth      =   6090
      TabIndex        =   0
      Top             =   -75
      Width           =   6090
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   555
         TabIndex        =   1
         Top             =   270
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChildStationInEPR"
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
Private mlng病人id As Long
Private mlng主页id As Long
Private mlng医嘱id As Long
Private mlng科室ID As Long
    
Private WithEvents mobjDoc As zlRichEPR.cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Event AfterDataChanged()

'######################################################################################################################

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

Public Sub zlDefCommandBars(ByVal cbsMain As CommandBars)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "病历(&E)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Edit_NewItem, "新增(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除(&D)")

    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain(2)
    
    For Each objControl In objBar.Controls  '先求出前面的最后一个Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
        
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "保存", True, , , , objControl.Index + 1)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消", , , , , objControl.Index + 1)
    
    '命令的快键绑定:公共部份主界面已处理
    '------------------------------------------------------------------------------------------------------------------
'    With cbsMain.KeyBindings
'        .Add 0, vbKeyF2, conMenu_Edit_NewItem              '保存
'    End With

End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim rs As New ADODB.Recordset
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem * 2
        
        If Val(Split(Control.Parameter, ";")(1)) = -1 Then
        
            If Control.Caption = "麻醉记录" Then
            
                '特别的处理方式
'                Call mobjDoc.ShowCaseNarcosis(mlng病人id, mlng主页id, 0, mlng科室ID, 1, mfrmMain, True)
                
            End If
            
        ElseIf Val(Split(Control.Parameter, ";")(0)) > 0 Then
        
            Call mobjDoc.InitEPRDoc(cprEM_新增, cprET_单病历编辑, Val(Split(Control.Parameter, ";")(0)), cprPF_住院, mlng病人id, mlng主页id, 0, mlng科室ID, mlng医嘱id)
            mobjDoc.ShowEPREditor mfrmMain
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
    
        With vsf(0)
            
            If Val(.TextMatrix(.Row, .ColIndex("本次"))) = 0 Then Exit Sub
            
            gstrSQL = "Select a.名称,a.保留 From 病历文件列表 a,电子病历记录 b Where b.ID=[1] And b.文件id=a.ID"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(.Row)))
            If rs("名称").Value = "麻醉记录" And zlCommFun.NVL(rs("保留").Value, 0) = -1 Then

                '特别的处理方式
'                Call mobjDoc.ShowCaseNarcosis(mlng病人id, mlng主页id, Val(.RowData(.Row)), mlng科室ID, 2, mfrmMain, True)
                
            Else
            
                If Val(.RowData(.Row)) > 0 Then
                    Call mobjDoc.InitEPRDoc(cprEM_修改, cprET_单病历编辑, Val(.RowData(.Row)), cprPF_住院, mlng病人id, mlng主页id, 0, mlng科室ID, mlng医嘱id)
                    mobjDoc.ShowEPREditor mfrmMain
                End If
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        If ExecuteCommand("删除病历") Then
            Call ExecuteCommand("读取数据")
        End If
        
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Enabled = mblnAllowModify And mlngKey > 0
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        With vsf(0)
            Control.Enabled = mblnAllowModify And mlngKey > 0 And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("本次"))) = 1
        End With
    End Select
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim objControl As CommandBarControl

    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_NewItem
        With CommandBar.Controls
            .DeleteAll
            
            strSQL = "Select a.ID,a.名称,a.保留 From 病历文件列表 a,病历时限要求 b Where a.Id=b.文件id And b.事件='手术' And 种类=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, 2)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 2, rs("名称").Value)
                    objControl.Parameter = rs("ID").Value & ";" & zlCommFun.NVL(rs("保留").Value, 0)
                    rs.MoveNext
                Loop
            End If
        End With
    End Select
        
End Sub

Public Function RefreshData(ByVal lngKey As Long, ByVal lng病人id As Long, ByVal lng主页id As Long, ByVal lng科室id As Long, ByVal lng医嘱id As Long, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngKey = lngKey
    mlng病人id = lng病人id
    mlng主页id = lng主页id
    mlng医嘱id = lng医嘱id
    mlng科室ID = lng科室id
    
    mblnAllowModify = blnAllowModify
    
    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("控件状态")
    
    If mlng病人id > 0 Then
        If ExecuteCommand("读取数据") = False Then Exit Function
        
    End If
    Call ExecuteCommand("显示病历")
    
    DataChanged = False
    
    RefreshData = True
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
'    Dim objPane As Pane
'
'    Set objPane = dkpMain.CreatePane(10, 100, 100, DockTopOf, Nothing)
'    objPane.Title = "病历列表"
'    objPane.Options = PaneNoCaption
'
'    Set objPane = dkpMain.CreatePane(11, 100, 100, DockBottomOf, objPane)
'    objPane.Title = "病历内容"
'    objPane.Options = PaneNoCaption
'
''    dkpMain.SetCommandBars cbsMain
'
'    dkpMain.Options.ThemedFloatingFrames = True
'    dkpMain.Options.UseSplitterTracker = False '实时拖动
'    dkpMain.Options.AlphaDockingContext = True
'    dkpMain.Options.HideClient = True
End Sub

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
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("病历名称", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("创建人", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("创建时间", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("保存人", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("当前版本", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("科室名", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("本次", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("当前情况", 900, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        Call InitDockPannel
        
        Set mobjDoc = New zlRichEPR.cEPRDocument
        
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 And mbytMode = 2 Then blnAllowModify = False

    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
    
        mclsVsf.ClearGrid
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
        
        mclsVsf.ClearGrid
        
        gstrSQL = "Select r.Id, r.病历名称, r.创建人 As 创建人,Decode(s.记录id,Null,0,[3],1,0) As 本次," & _
                "        r.创建时间, r.保存人," & _
                "        r.完成时间, r.最后版本 As 当前版本," & _
                "        Decode(r.最后版本, 1, '书写：', '修订：') || r.保存人 || '在' || To_Char(r.保存时间, 'yyyy-mm-dd hh24:mi') ||" & _
                "         Decode(Nvl(r.签名级别, 0), 0, '保存(未完成)', 1, '完成', '审签') As 当前情况, r.归档人, r.归档日期," & _
                "        d.名称 As 科室名, f.保留, r.处理状态, p.病人状态" & _
                " From 电子病历记录 r, 部门表 d,病人手术病历 s," & _
                "      (Select Decode(出院日期, Null, Decode(状态, 3, '预出院', '在院'), '出院') As 病人状态" & vbNewLine & _
                "        From 病案主页" & vbNewLine & _
                "        Where 病人id = [1] And 主页id = [2]) p," & _
                "      (Select d.Id As 文件id, f.种类, f.编号, f.名称 As 页面, d.保留" & _
                "        From 病历文件列表 d, 病历页面格式 f" & _
                "        Where d.种类 In (2, 5, 6) And d.种类 = f.种类 And d.页面 = f.编号) f" & _
                " Where r.文件id = f.文件id(+) And r.病人来源 = 2 And r.病历种类 In (2, 5, 6) And r.科室id = d.Id And r.病人id = [1] And r.主页id = [2] And s.文件id(+)=r.ID" & _
                " Order By r.病历种类, f.编号, r.序号, r.Id"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlng病人id, mlng主页id, mlngKey)
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)

    '------------------------------------------------------------------------------------------------------------------
    Case "显示病历"
        
        With vsf(0)
            Call ShowDocument(edt, Val(.RowData(.Row)))
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "删除病历"
    
        With vsf(0)
            
            If Val(.TextMatrix(.Row, .ColIndex("本次"))) = 0 Then Exit Function
            
            If MsgBox("是否真的要删除“" & .TextMatrix(.Row, .ColIndex("病历名称")) & "”病历内容？", vbYesNo + vbDefaultButton2 + vbQuestion, ParamInfo.系统名称) = vbNo Then Exit Function
            
                        
            gstrSQL = "zl_病人手术病历_Delete(" & mlngKey & "," & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)
            
            ExecuteCommand = SQLRecordExecute(rsSQL, mfrmMain.Caption)
            
            Exit Function
        End With
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

'Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
'    Select Case Item.ID
'    Case 10
'        Item.Handle = picPane(0).hWnd
'    Case 11
'        Item.Handle = picPane(1).hWnd
'    End Select
'End Sub

Private Sub Form_Resize()
    On Error Resume Next
'
'    Call SetPaneRange(dkpMain, 10, 100, 100, Me.ScaleWidth, 250)
'
'    dkpMain.RecalcLayout
    
    picPane(0).Move 0, 0, Me.ScaleWidth
    picPane(1).Move 0, picPane(0).Top + picPane(0).Height + 30, Me.ScaleWidth, Me.ScaleHeight - (picPane(0).Top + picPane(0).Height + 30)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Set mclsVsf = Nothing
    Set mobjDoc = Nothing

End Sub

Private Sub mobjDoc_AfterSaved(lngRecordId As Long)

    '刷新界面
    Call ExecuteCommand("读取数据")
    Call ExecuteCommand("显示病历")
End Sub

Private Sub mobjDoc_BeforeSaved(lngRecordId As Long, Cancel As Boolean)
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    
    If mlngKey > 0 And lngRecordId > 0 Then
        
        Call SQLRecord(rsSQL)
        
        strSQL = "zl_病人手术病历_Update(" & mlngKey & "," & lngRecordId & ")"
        Call SQLRecordAdd(rsSQL, strSQL)
        
        Cancel = Not SQLRecordExecute(rsSQL, mfrmMain.Caption, False)
        
    End If
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf.AppendRows = True
    Case 1
        edt.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow Then Call ExecuteCommand("显示病历")
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        Set cbrPopupBar = CopyMenu(mfrmMain.cbsMain, 3)
        cbrPopupBar.ShowPopup
    End Select
End Sub
