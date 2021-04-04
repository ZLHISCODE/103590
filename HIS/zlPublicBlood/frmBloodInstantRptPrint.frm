VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmBloodInstantRptPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输血执行单打印"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7830
   Icon            =   "frmBloodInstantRptPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid VSFPrint 
      Height          =   1560
      Left            =   1755
      TabIndex        =   0
      Top             =   390
      Width           =   2790
      _cx             =   4921
      _cy             =   2752
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
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   0
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBloodInstantRptPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean
Private mlngAdviceId As Long
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Function ShowMe(ByVal objfrm As Object, ByVal lngAdviceid As Long) As Boolean
    
    mblnOk = False
    mlngAdviceId = lngAdviceid
    Call InitCommandBar
    If LoadVsfPrint = False Then Exit Function
    If Not objfrm Is Nothing Then
        Me.Show 1, objfrm
    Else
        Me.Show 1
    End If
    ShowMe = mblnOk
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Select Case Control.id
        Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll
            For i = VSFPrint.FixedRows To VSFPrint.Rows - 1
                VSFPrint.TextMatrix(i, VSFPrint.ColIndex("选择")) = IIf(Control.id = conMenu_Edit_SelAll, 1, 0)
            Next
        Case conMenu_File_PrintSet
            Call Rptprint(0)
        Case conMenu_File_Preview
            Call Rptprint(1)
        Case conMenu_File_Print
            Call Rptprint(2)
        Case conMenu_File_Exit
            mblnOk = False
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim Rmain As RECT
    
    On Error Resume Next
    Call cbsMain.GetClientRect(Rmain.Left, Rmain.Top, Rmain.Right, Rmain.Bottom)
    With VSFPrint
        .Left = Rmain.Left
        .Top = Rmain.Top
        .Width = Rmain.Right - Rmain.Left
        .Height = Rmain.Bottom - Rmain.Top
    End With
End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：初始化Commandbar
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始化处理
    
    Call CommandBarInit(cbsMain)
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    Set cbsMain.Icons = gobjCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap Or xtpFlagAlignBottom
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "全选", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "全清", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_PrintSet, "打印设置", True, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出", True, , xtpButtonIconAndCaption)

    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll            '全选
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_ClsAll         '全清
    End With
    
    InitCommandBar = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadVsfPrint() As Boolean
'功能：初始化表格，并加载表格数据
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, VSFPrint, True, True)
        Call .ClearColumn
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTBoolean, "", "选择", False)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)  '收发ID
        Call .AppendColumn("状态", 810, flexAlignLeftCenter, flexDTString, , "血液状态") '接收执行状态
        Call .AppendColumn("血袋编号", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("血液名称", 1800, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("规格", 810, flexAlignLeftCenter, flexDTString, , "血液规格")
        Call .AppendColumn("ABO", 810, flexAlignLeftCenter, flexDTString, , "ABO", True)
        Call .AppendColumn("Rh(D)", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
        
        '执行信息
        Call .AppendColumn("执行人", 1200, flexAlignLeftCenter, flexDTString, , "开始执行人")
        Call .AppendColumn("开始时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("结束人", 1200, flexAlignLeftCenter, flexDTString, , "结束执行人")
        Call .AppendColumn("结束时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
       
        .AppendRows = False
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
    End With
    
    strSQL = _
        "Select Id, Max(血袋编号) 血袋编号, Max(Abo) Abo, Max(Rh) Rh, Max(血液名称) 血液名称, Max(血液规格) 血液规格, Max(血液状态) 血液状态, Max(开始时间) 开始时间," & vbNewLine & _
        "       Max(开始执行人) 开始执行人, Max(结束时间) 结束时间, Max(结束执行人) 结束执行人" & vbNewLine & _
        "From (Select a.Id, a.血袋编号, a.Abo, a.Rh, e.名称 As 血液名称, e.规格 血液规格," & vbNewLine & _
        "              Decode(Nvl(f.执行状态, 0), 1, '正在执行', 2, '完成执行', 3, '停止执行') 血液状态, Decode(g.记录性质, 1, g.执行时间) 开始时间," & vbNewLine & _
        "              Decode(g.记录性质, 1, g.执行人) 开始执行人, Decode(g.记录性质, 3, g.执行时间) 结束时间, Decode(g.记录性质, 3, g.执行人) 结束执行人" & vbNewLine & _
        "       From 收费项目目录 e, 血液收发记录 a, 血液执行记录 g, 血液发送记录 f, 血液配血记录 b" & vbNewLine & _
        "       Where e.Id = a.血液id And Nvl(a.填写数量, 0) <> 0 And Mod(a.记录状态, 3) = 1 And a.Id = f.收发id And g.收发id = f.收发id And" & vbNewLine & _
        "             f.配发id = b.Id And b.申请id = [1])" & vbNewLine & _
        "Group By Id" & vbNewLine & _
        "Order By 开始时间"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "已发血液信息提取", mlngAdviceId)
    If rsTemp.EOF Then
        MsgBox "该医嘱还未进行输血执行情况登记，请登记后再进行此操作！", vbInformation, gstrSysName
        Exit Function
    End If
    Call mclsVsf.LoadGrid(rsTemp, "", True)
    LoadVsfPrint = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Rptprint(ByVal bytMode As Byte)
    Dim i As Integer
    Dim strIDs As String
    Dim strRptName As String
    
    strRptName = "ZL22_BILL_9005_1" 'ZL22_BILL_1938
    Select Case bytMode
        Case 0  '打印设置
            Call ReportPrintSet(gcnOracle, 2200, strRptName, Me)
        Case 1, 2 '预览  打印
            With VSFPrint
                For i = .FixedRows To .Rows - 1
                    If Abs(Val(.TextMatrix(i, .ColIndex("选择")))) = 1 Then
                        strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
                    End If
                Next
            End With
            strIDs = Mid(strIDs, 2)
            ReportOpen gcnOracle, 2200, strRptName, Me, "医嘱id=" & mlngAdviceId, "收发ID=" & strIDs, bytMode
            mblnOk = True
    End Select
End Sub

Private Sub VSFPrint_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub VSFPrint_DblClick()
    With VSFPrint
        If .Row >= .FixedRows And .Col >= .FixedCols Then
            If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 Then
                .TextMatrix(.Row, .ColIndex("选择")) = IIf(Abs(Val(.TextMatrix(.Row, .ColIndex("选择")))) = 1, 0, 1)
            End If
        End If
    End With
End Sub

Private Sub VSFPrint_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not Col = VSFPrint.ColIndex("选择") Then
        Cancel = True
    End If
End Sub
