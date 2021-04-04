VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmbloodReactionPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印报表"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6480
   Icon            =   "frmbloodReactionPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VSFlex8Ctl.VSFlexGrid VSFPrint 
      Height          =   2295
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      _cx             =   7435
      _cy             =   4048
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
      AllowUserResizing=   1
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Bindings        =   "frmbloodReactionPrint.frx":000C
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmbloodReactionPrint.frx":0020
   End
End
Attribute VB_Name = "frmbloodReactionPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlng病人ID As Long
Private mlng病人来源 As Long          '1-门诊  2-住院
Private mlng主页id As String
Private mRsFY As ADODB.Recordset      '病人信息记录集
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mstr收发id As String
Private mstr挂号单 As String
Private mlng阶段 As Long              '0-门诊  1-住院  2-输血科
Private marrFilter                    '分解mstrFilter后得到的过滤条件数组
Private mstrFilter As String          '过滤条件字符串
Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：初始化Commandbar
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始化处理
    
    Call CommandBarInit(cbsMain)
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    Set cbsMain.Icons = ImageManager.Icons
    cbsMain.Options.LargeIcons = False
    
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap Or xtpFlagAlignBottom
        
        Set objControl = NewToolBar(objBar, xtpControlButton, 1, "全选", True, , xtpButtonIconAndCaption)
        Set objControl = NewToolBar(objBar, xtpControlButton, 2, "全清", True, , xtpButtonIconAndCaption)
        Set objControl = NewToolBar(objBar, xtpControlButton, 4, "确定", True, , xtpButtonIconAndCaption)
        mobjStateInfo.Flags = xtpFlagRightAlign

    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings

        .Add FCONTROL, vbKeyA, 1            '全选
        .Add FSHIFT, vbKeyDelete, 2         '全清
        
    End With
    
    InitCommandBar = True
    Exit Function
ErrHand:
    
End Function

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim lngi As Long
    Dim lngj As Long
    Dim rsSAD As New ADODB.Recordset
    Dim StrSqlSAD As String
    Dim strOPT As String
    Dim CanbeTransfer As Boolean
    Dim blnHD As Boolean
    blnHD = True
    On Error GoTo Error
    
    Call SQLRecord(rsSAD)
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        
            Case "初始表格"
                '初始vsf表格
                Set mclsVsf = New clsVsf
                With mclsVsf
                    Call .Initialize(Me.Controls, VSFPrint, True, True)
                    Call .ClearColumn
                    Call .AppendColumn("收发id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("", 400, flexAlignLeftCenter, flexDTBoolean, "", , True)
                    Call .AppendColumn("血袋编号", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("输血项目", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("输入量", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("状态", 1000, flexAlignLeftCenter, flexDTString, , "", True)
                    Call .AppendColumn("输血史", 0, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("输血次数", 0, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("反应时间", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("记录人", 900, flexAlignLeftCenter, flexDTString, , "", True)
                    Call .AppendColumn("记录时间", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("确认人", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("确认时间", 1200, flexAlignLeftCenter, flexDTString, "", , True)

                    .AppendRows = False
                    .SysHidden(.ColIndex("收发id")) = True
                    .SysHidden(.ColIndex("输血史")) = True
                    .SysHidden(.ColIndex("输血次数")) = True
                    Call .InitializeEdit(True, True, True)
                    Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
                    
                End With
                
            Case "获取反应记录"
                Dim lngFilter
                Dim ArrTime
                Dim str是否本人 As String
                Dim lng部门id As Long
                Dim str开始时间 As String
                Dim str结束时间 As String
                
                If mstrFilter <> "" Then
                    lngFilter = marrFilter(3) '提交状态
                    str是否本人 = marrFilter(2) '记录人
                    lng部门id = marrFilter(0) '部门id
                    str开始时间 = Split(marrFilter(1), "'")(0)
                    str结束时间 = Split(marrFilter(1), "'")(1)
                Else
                    lngFilter = 0
                    str是否本人 = ""
                    lng部门id = 0
                    str开始时间 = Now
                    str结束时间 = Now
                End If
                '获取病人的输血反应数据，主要是从输血反应记录中获取
                strSqlFY = " select distinct d.收发id, c.血袋编号, Decode(d.输血史, 0, '无', '有') As 输血史, d.输血次数, d.输血项目, d.输入量, " & _
                           " Decode(d.状态, 0, '未提交', 1, '医生已提交', '输血科已提交') As 状态, to_char(d.反应时间,'yyyy-mm-dd HH24:mi:ss') as 反应时间, d.记录人, d.记录时间, d.确认人,to_char(d.确认时间,'yyyy-mm-dd HH24:mi:ss') as 确认时间 " & _
                           " from 病人医嘱记录 a,血液配血记录 b,血液收发记录 c,输血反应记录 d " & _
                           " where d.收发id=c.id and c.配发id=b.id and mod(c.记录状态,3)=1 and c.核对人 is not null and b.申请id=a.id " & _
                           " and a.病人id=[1] "
                
                If mlng病人来源 = 2 Then '住院病人
                    If lngFilter = 0 And mlng阶段 = 1 Then '全部数据,医生阶段
                        strSqlFY = strSqlFY & "and a.主页id=[2] "
                    ElseIf lngFilter = 0 And mlng阶段 = 2 Then '全部数据,输血科阶段
                        strSqlFY = strSqlFY & "and a.主页id=[2] and d.状态 <>0 "
                    ElseIf lngFilter = 1 And mlng阶段 = 1 Then '未提交数据,医生
                        strSqlFY = strSqlFY & "and a.主页id=[2] and d.状态=0 "
                    ElseIf lngFilter = 1 And mlng阶段 = 2 Then '未提交数据,输血科
                        strSqlFY = strSqlFY & "and a.主页id=[2] and d.状态=1 "
                    ElseIf lngFilter = 2 And mlng阶段 = 1 Then '已提交数据，医生
                        strSqlFY = strSqlFY & "and a.主页id=[2] and d.状态 <>0 "
                    ElseIf lngFilter = 2 And mlng阶段 = 2 Then '已提交数据，输血科
                        strSqlFY = strSqlFY & "and a.主页id=[2] and d.状态=2 "
                    End If
                Else
                    If lngFilter = 0 And mlng阶段 = 1 Then '全部数据,医生阶段
                        strSqlFY = strSqlFY & "and a.挂号单=[7] "
                    ElseIf lngFilter = 0 And mlng阶段 = 2 Then '全部数据,输血科阶段
                        strSqlFY = strSqlFY & "and a.挂号单=[7] and d.状态 <>0 "
                    ElseIf lngFilter = 1 And mlng阶段 = 1 Then '未提交数据,医生
                        strSqlFY = strSqlFY & "and a.挂号单=[7] and d.状态=0 "
                    ElseIf lngFilter = 1 And mlng阶段 = 2 Then '未提交数据,输血科
                        strSqlFY = strSqlFY & "and a.挂号单=[7] and d.状态=1 "
                    ElseIf lngFilter = 2 And mlng阶段 = 1 Then '已提交数据，医生
                        strSqlFY = strSqlFY & "and a.挂号单=[7] and d.状态 <>0 "
                    ElseIf lngFilter = 2 And mlng阶段 = 2 Then '已提交数据，输血科
                        strSqlFY = strSqlFY & "and a.挂号单=[7] and d.状态=2 "
                    End If
                End If
                
                If marrFilter(2) <> "" Then
                    strSqlFY = strSqlFY & " and d.记录人=[3] "
                End If
    
                If Val(marrFilter(0)) <> -1 And mstrFilter <> "" Then '如果没有过滤条件时Val(mArrFilter(0))=0
                    strSqlFY = strSqlFY & " and c.对方部门id =[4] "
                End If
                
                If mlng阶段 = 2 And mstrFilter <> "" Then  '输血科要按时间过滤反应记录
                    strSqlFY = strSqlFY & " and d.反应时间 Between [5] and [6] order by 反应时间"
                Else
                    strSqlFY = strSqlFY & " order by 反应时间 "
                End If
                
                Set mRsFY = gobjDatabase.OpenSQLRecord(strSqlFY, "病人输血反应记录", mlng病人ID, mlng主页id, str是否本人, lng部门id, CDate(str开始时间), CDate(str结束时间), mstr挂号单)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End Select
    Next
    ExecuteCommand = True
    Exit Function
Error:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    ExecuteCommand = False
End Function

Public Function BloodPrintList(lng病人ID As Long, lng病人来源 As Long, lng主页id As Long, strFilter As String, lng阶段 As Long, Optional strSelBloodid As String = "") As String
    '功能：初始化打印界面
    '参数：lng病人ID-病人id
    '      lng病人来源-1-门诊 2-住院
    '      lng主页id-主页id
    '      strFilter-过滤条件字符串
    '      lng阶段-0-门诊 1-住院 2-输血科
    Dim strSQL As String
    Dim rsSql As ADODB.Recordset
    Dim lngi As Long
    mstrFilter = strFilter
    If strFilter <> "" Then
        marrFilter = Split(strFilter, "|")
    Else
        ReDim marrFilter(0 To 3)
    End If
    
    mlng病人ID = lng病人ID
    mlng病人来源 = lng病人来源
    mlng主页id = lng主页id
    mstr收发id = ""
'    mlngFilter = lngFilter
    mlng阶段 = lng阶段
    strSQL = " select no from 病人挂号记录 where id=[1]"
    Set rsSql = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", mlng主页id)
    If rsSql.RecordCount > 0 Then
        mstr挂号单 = rsSql.Fields("no")
    End If
    InitCommandBar
    Call ExecuteCommand("初始表格")
    Call ExecuteCommand("获取反应记录")
    Call mclsVsf.LoadGrid(mRsFY)
    
    If strSelBloodid <> "" Then '定位到选中的输血反应记录
        For lngi = 1 To VSFPrint.Rows - 1
            If mclsVsf.TextMatrix(lngi, mclsVsf.ColIndex("收发id")) = Val(strSelBloodid) Then
                VSFPrint.TextMatrix(lngi, 1) = -1
                Exit For
            End If
        Next
    End If
    Me.Show (1)
    BloodPrintList = mstr收发id
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngi As Long
    Dim blnOk As Boolean
    blnOk = False
    Select Case Control.id
        Case 1
            For lngi = 1 To VSFPrint.Rows - 1
                VSFPrint.TextMatrix(lngi, 1) = -1
            Next
        Case 2
            For lngi = 1 To VSFPrint.Rows - 1
                VSFPrint.TextMatrix(lngi, 1) = 0
            Next
        Case 4
            mstr收发id = ""
            For lngi = 1 To VSFPrint.Rows - 1
                If Val(VSFPrint.TextMatrix(lngi, 1)) = -1 Then
                    mstr收发id = mstr收发id & VSFPrint.TextMatrix(lngi, VSFPrint.ColIndex("收发id")) & ";"
                    blnOk = True
                End If
            Next
            If mstr收发id <> "" Then mstr收发id = Left(mstr收发id, Len(mstr收发id) - 1)
            If blnOk = True Then
                Unload Me
            Else
                MsgBox "未选择反应记录！", vbInformation, gstrSysName
            End If
            
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long
    On Error GoTo Errorhand
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    '窗体其它控件Resize处理
    VSFPrint.Move lngLeft, lngTop + 50, lngRight - lngLeft, lngBottom - lngTop
Errorhand:
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mRsFY = Nothing
End Sub

Private Sub VSFPrint_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

