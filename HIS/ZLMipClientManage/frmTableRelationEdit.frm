VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmTableRelationEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   4290
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7650
   Icon            =   "frmTableRelationEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3900
      Index           =   0
      Left            =   -15
      ScaleHeight     =   3900
      ScaleWidth      =   7755
      TabIndex        =   11
      Top             =   465
      Width           =   7755
      Begin VB.Frame fra 
         Height          =   3135
         Left            =   30
         TabIndex        =   12
         Top             =   -75
         Width           =   7620
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   2
            Left            =   7215
            Picture         =   "frmTableRelationEdit.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   195
            Width           =   300
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   210
            Width           =   1170
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   4905
            TabIndex        =   5
            Top             =   195
            Width           =   2325
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   930
            TabIndex        =   1
            Top             =   210
            Width           =   1785
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   2460
            Index           =   0
            Left            =   915
            TabIndex        =   8
            Top             =   585
            Width           =   6630
            _cx             =   11695
            _cy             =   4339
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "参数赋值"
            Height          =   180
            Index           =   4
            Left            =   150
            TabIndex        =   7
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "目标"
            Height          =   180
            Index           =   2
            Left            =   4470
            TabIndex        =   4
            Top             =   270
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "关系命名"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   0
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "类型"
            Height          =   180
            Index           =   0
            Left            =   2790
            TabIndex        =   2
            Top             =   255
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(C)"
         Height          =   350
         Left            =   6450
         TabIndex        =   10
         Top             =   3240
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   5115
         TabIndex        =   9
         Top             =   3240
         Width           =   1100
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmTableRelationEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################

Private Type Items
    目标信息 As String
End Type
Private usrSaveItem As Items

Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private mstrSourceDataKey As String
Private mstrSourceField As String
Private mrsCondition As ADODB.Recordset
Private mrsSourceTabField As ADODB.Recordset
Private mblnContiune As Boolean
Private mstrBusiness As String

Private WithEvents mclsVsf As zlVSFlexGrid.clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Event AfterNewData(ByVal DataKey As String)
Public Event AfterModifyData(ByVal DataKey As String)
Public Event AfterDeleteData(ByVal DataKey As String)
Public Event Forward(ByRef DataKey As String, ByRef Cancel As Boolean)
Public Event Backward(ByRef DataKey As String, ByRef Cancel As Boolean)

'######################################################################################################################

Public Function InitDialog(ByVal frmParent As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    
    InitDialog = True
    
End Function

Public Sub NewData(ByVal strBusiness As String, ByVal strSourceDataKey As String, ByVal strSourceField As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 1
    mstrBusiness = strBusiness
    Me.Caption = "新增关系"
    mstrDataKey = ""
    mstrSourceDataKey = strSourceDataKey
    mstrSourceField = strSourceField
    
    Call InitData
    Call InitCommandBar
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub ModifyData(ByVal strBusiness As String, ByVal strSourceDataKey As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 2
    mstrBusiness = strBusiness
    mstrSourceDataKey = strSourceDataKey
    
    mstrDataKey = strDataKey
    
    Me.Caption = "修改关系"
        
    Call InitData
    Call InitCommandBar
    
    Call ReadData(mstrDataKey)
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub DeleteData(ByVal strBusiness As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 3
    mstrBusiness = strBusiness
    If strDataKey = "" Then Exit Sub
    mstrDataKey = strDataKey
    
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "ID", mstrDataKey)
        
    If gclsBusiness.TableRelationEdit("Delete", mrsPara) Then
        RaiseEvent AfterDeleteData(mstrDataKey)
    End If
End Sub

'######################################################################################################################
Private Function InitData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    mblnContiune = False
    
    Set rsTmp = gclsBusiness.TableRelationStruct()
    If Not (rsTmp Is Nothing) Then
        txt(1).MaxLength = rsTmp("ext_title").Precision
    End If
    
    With cbo(0)
        .AddItem "1-引用"
        .ItemData(.NewIndex) = 1
        .AddItem "2-组合"
        .ItemData(.NewIndex) = 2
        
        .ListIndex = 0
    End With
    
    Set mclsVsf = New zlVSFlexGrid.clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf(0), True, True, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[序号]", False)
        Call .AppendColumn("参数名称", 2100, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("参数类型", 0, flexAlignLeftCenter, flexDTString, , "", , , , True)
        Call .AppendColumn("赋值关系", 900, flexAlignCenterCenter, flexDTString, , "", True)
        Call .AppendColumn("参数赋值", 2100, flexAlignLeftCenter, flexDTString, , "", True)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("参数赋值"), True, vbVsfEditCombox, "|")
                           
        .AppendRows = True
        
    End With
    
    '------------------------------------------------------------------------------------------------------------------
        
    Set mrsCondition = zlCommFun.CreateCondition
    
    Call zlCommFun.SetCondition(mrsCondition, "业务信息id", mstrSourceDataKey)
    Call zlCommFun.SetCondition(mrsCondition, "字段前缀", "原表")
    Set mrsSourceTabField = gclsBusiness.TableFieldRead("字段", mrsCondition)
    
    InitData = True
End Function

Private Function ReadRelationParameter(ByVal strDataKey As String, ByVal strTargetDataKey As String) As Boolean
    
    Dim rsTmp As ADODB.Recordset

    Set mrsCondition = zlCommFun.CreateCondition
    
    mclsVsf.ClearGrid
    
    Call zlCommFun.SetCondition(mrsCondition, "id", strDataKey)
    Call zlCommFun.SetCondition(mrsCondition, "目标信息表id", strTargetDataKey)
    Set rsTmp = gclsBusiness.TableRelationParameterRead("", mrsCondition)
    If rsTmp.BOF = False Then
        Call mclsVsf.LoadGrid(rsTmp)
    End If
                    
    ReadRelationParameter = True
    
End Function

Private Sub FillRowSourceField(ByVal intRow As Integer, ByVal strDefaultField As String)
    Dim strTemp As String
    
    With vsf(0)
        mrsSourceTabField.Filter = ""
        mrsSourceTabField.Filter = "字段类型='" & .TextMatrix(intRow, .ColIndex("参数类型")) & "'"
        If mrsSourceTabField.RecordCount > 0 Then
            mrsSourceTabField.MoveFirst
            strTemp = vsf(0).BuildComboList(mrsSourceTabField, "字段名称")
            mclsVsf.DropColData(mclsVsf.ColIndex("参数赋值")) = strTemp
            
            If InStr(UCase(strTemp), "原表." & UCase(strDefaultField)) > 0 And mbytMode = 1 Then
                .TextMatrix(intRow, .ColIndex("参数赋值")) = "原表." & UCase(strDefaultField)
            End If
        End If
    
    End With

End Sub

Private Function ReadData(ByVal strDataKey As String) As Boolean

    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)
    
    mblnReading = True
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsTmp = gclsBusiness.TableRelationRead("id", rsCondition)
    If rsTmp.BOF = False Then
        zlControl.CboLocate cbo(0), zlCommFun.NVL(rsTmp("类型").Value)
        txt(1).Text = zlCommFun.NVL(rsTmp("命名").Value)
        txt(2).Text = zlCommFun.NVL(rsTmp("目标").Value)
        cmd(2).Tag = zlCommFun.NVL(rsTmp("目标id").Value)
    End If
    
    Call ReadRelationParameter(strDataKey, cmd(2).Tag)
    
    Call FillRowSourceField(vsf(0).Row, "")
    
    mblnReading = False
    mblnDataChanged = False
    
    ReadData = True
    
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeWhidbey
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Select Case mbytMode
    Case 1
        Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, "确定后继续新增", True)
    Case 2
        Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, "确定后继续修改", True)
    End Select
    objControl.IconId = conMenu_View_UnCheck
    
'    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Forward, "上一条", True)
'    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Backward, "下一条")
    

    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
        
    If Len(txt(1).Text) = 0 Then
        ShowSimpleMsg "关系的命名不能为空！"
        Call LocationObj(txt(1))
        Exit Function
    End If
    
    If cmd(2).Tag = "" Then
        ShowSimpleMsg "目标信息不能为空！"
        Call LocationObj(txt(2))
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Function SaveData(ByRef strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    Dim strLine As String
    Dim strTemp As String
    Dim lngCount As Long
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
    
    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
    Call zlCommFun.SetParameter(rsPara, "关系命名", Trim(txt(1).Text))
    Call zlCommFun.SetParameter(rsPara, "关系类型", cbo(0).ItemData(cbo(0).ListIndex))
    Call zlCommFun.SetParameter(rsPara, "来源信息", mstrSourceDataKey)
    Call zlCommFun.SetParameter(rsPara, "目标信息", Trim(cmd(2).Tag))
    
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
    lngCount = 0
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If Trim(.TextMatrix(lngLoop, .ColIndex("参数名称"))) <> "" And Trim(.TextMatrix(lngLoop, .ColIndex("参数赋值"))) <> "" Then
                
                strLine = lngLoop
                strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("参数名称")))
                strLine = strLine & "," & Mid(Trim(.TextMatrix(lngLoop, .ColIndex("参数赋值"))), 4)
                
                If LenB(strTemp & ";" & strLine) > 3500 Then
                    If strTemp <> "" Then
                        lngCount = lngCount + 1
                        strTemp = Mid(strTemp, 2)
                        Call zlCommFun.SetParameter(rsPara, "赋值关系_" & lngCount, strTemp)
                        strTemp = ""
                    End If
                End If
                strTemp = strTemp & ";" & strLine
            
            End If
        Next
    End With
    If strTemp <> "" Then
        lngCount = lngCount + 1
        strTemp = Mid(strTemp, 2)
        Call zlCommFun.SetParameter(rsPara, "赋值关系_" & lngCount, strTemp)
    End If
    Call zlCommFun.SetParameter(rsPara, "赋值关系段数", lngCount)
            
            
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '新增
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        
        SaveData = gclsBusiness.TableRelationEdit("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '修改
        SaveData = gclsBusiness.TableRelationEdit("UPDATE", rsPara)
    End Select
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Click(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Dim blnCancel As Boolean
    Dim strDataKey As String
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward               '上一条
        
        strDataKey = mstrDataKey
        
        RaiseEvent Forward(strDataKey, blnCancel)
        If blnCancel = False Then
        
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
    
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward               '下一条
        
        strDataKey = mstrDataKey
        
        RaiseEvent Backward(strDataKey, blnCancel)
        If blnCancel = False Then
            
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
            
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        mblnContiune = Not mblnContiune
    End Select
    
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '窗体其它控件Resize处理
    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Filter, conMenu_View_LocationItem, conMenu_View_Backward, conMenu_View_Forward, 0
        Control.Visible = (mbytMode = 2)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnContiune
        Control.IconId = IIf(mblnContiune = True, conMenu_View_Check, conMenu_View_UnCheck)
    End Select
    
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        
        Set rsCondition = zlCommFun.CreateCondition
        Call zlCommFun.SetCondition(rsCondition, "data_code", mstrBusiness)
        Set rsData = gclsBusiness.TableRead("SelectData", rsCondition)
        
        If zlCommFun.ShowPubSelect(Me, txt(Index), 2, "编码,900,0,1;名称,2400,0,0;说明,1200,0,0", mfrmParent.Name & "\入口信息选择", "请从下表中选择一个入口信息", rsData, rs, , , , Trim(cmd(Index).Tag), , True) = 1 Then
            
            If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
                
                txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
                txt(Index).Tag = ""
                usrSaveItem.目标信息 = txt(Index).Text
                cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
                
                mblnDataChanged = True
                                                
                Call ReadRelationParameter(mstrDataKey, cmd(Index).Tag)
                
                With vsf(0)
                    Call FillRowSourceField(.Row, mstrSourceField)
                End With
            End If
            Call LocationObj(txt(Index), True)
        Else
            Call LocationObj(txt(Index), True)
            Exit Sub
        End If
        
    End Select
    
End Sub

Private Sub cmdCancel_Click()
    '
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strOldDataKey As String
    Dim rsTmp As ADODB.Recordset
    
    If mblnDataChanged = True Then
        If ValidData = True Then
    
            If SaveData(mstrDataKey) = True Then
                
                If strOldDataKey <> "" Then
                    RaiseEvent AfterModifyData(strOldDataKey)
                End If
                
                Select Case mbytMode
                Case 1
                    RaiseEvent AfterNewData(mstrDataKey)
                Case 2
                    RaiseEvent AfterModifyData(mstrDataKey)
                End Select
                
                If mblnContiune = False Then
                    mblnDataChanged = False
                    Unload Me
                Else
                    '重置环境，进入下一次新增状态
                    If mbytMode = 1 Then
                        mstrDataKey = ""
                        txt(0).Text = ""
                        txt(2).Text = ""
                        cmd(2).Tag = ""
                        mclsVsf.ClearGrid
                    End If
                    Call LocationObj(txt(0))
                    mblnDataChanged = False
                End If
                
            End If
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Set mobjFindKey = Nothing
    Set mrsPara = Nothing
    Set mclsVsf = Nothing
    
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    Select Case Index
    Case 2
        txt(Index).Tag = "Changed"
    End Select
    
    mblnDataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 4
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case Index
    Case 2
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
            txt(Index).Text = ""
            cmd(Index).Tag = ""
            txt(Index).Tag = ""
            usrSaveItem.目标信息 = ""
        End If
    End Select

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim bytMode As Byte
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
            
        '--------------------------------------------------------------------------------------------------------------
        Case 2
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""

                Set rsCondition = zlCommFun.CreateCondition
                Call zlCommFun.SetCondition(rsCondition, "FilterText", txt(Index).Text)
                
                Set rsData = gclsBusiness.TableRead("FilterData", rsCondition)
                If zlCommFun.ShowPubSelect(Me, txt(Index), 2, "编码,900,0,1;名称,2400,0,0;说明,1200,0,0", mfrmParent.Name & "\入口信息过滤", "请从下表中选择一个入口信", rsData, rs, , , , Trim(cmd(Index).Tag), , True) = 1 Then
                    
                    If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
                        txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
                        txt(Index).Tag = ""
                        usrSaveItem.目标信息 = txt(Index).Text
                        cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
                        mblnDataChanged = True
                                                
                        Call ReadRelationParameter(mstrDataKey, cmd(Index).Tag)
                        Call FillRowSourceField(vsf(0).Row, mstrSourceField)
                    End If
                Else
                    txt(Index).Text = usrSaveItem.目标信息
                    txt(Index).Tag = ""
                    Exit Sub
                End If
            End If
        End Select
        
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 4
        zlCommFun.OpenIme False
    End Select

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
    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Cancel Then Exit Sub

    Select Case Index
    Case 2
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.目标信息
            txt(Index).Tag = ""
        End If
    End Select
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    '编辑处理
    Call mclsVsf.AfterEdit(Row, Col)
    mblnDataChanged = True
        
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    
    If OldRow <> NewRow Then
        Call FillRowSourceField(NewRow, mstrSourceField)
    End If
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    Call mclsVsf.DbClick
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsf(Index)
        Select Case Button
        Case 1
            Call mclsVsf.AutoAddRow(.MouseRow, .MouseCol)
        End Select
    End With
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub

