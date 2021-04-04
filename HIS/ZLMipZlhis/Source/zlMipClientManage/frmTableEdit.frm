VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmTableEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   8325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12165
   Icon            =   "frmTableEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   7860
      Index           =   0
      Left            =   30
      ScaleHeight     =   7860
      ScaleWidth      =   12195
      TabIndex        =   15
      Top             =   540
      Width           =   12195
      Begin VB.Frame fra 
         Height          =   7125
         Left            =   30
         TabIndex        =   16
         Top             =   -90
         Width           =   12105
         Begin VB.TextBox txt 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   1
            Left            =   1845
            TabIndex        =   1
            Text            =   "001"
            Top             =   285
            Width           =   720
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            ForeColor       =   &H80000010&
            Height          =   300
            Left            =   780
            TabIndex        =   18
            Text            =   "ZLHIS_USER_"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Frame Frame1 
            Height          =   4545
            Left            =   780
            TabIndex        =   17
            Top             =   555
            Width           =   11145
            Begin RichTextLib.RichTextBox txtSQL 
               Height          =   3930
               Left            =   90
               TabIndex        =   7
               Top             =   135
               Width           =   10905
               _ExtentX        =   19235
               _ExtentY        =   6932
               _Version        =   393217
               BorderStyle     =   0
               ScrollBars      =   1
               TextRTF         =   $"frmTableEdit.frx":000C
            End
            Begin VB.CommandButton cmdVerfiy 
               Caption         =   "校验(&V)"
               Height          =   350
               Left            =   1290
               TabIndex        =   9
               Top             =   4140
               Width           =   1100
            End
            Begin VB.CommandButton cmdPara 
               Caption         =   "参数(&P)"
               Height          =   350
               Left            =   75
               TabIndex        =   8
               Top             =   4125
               Width           =   1100
            End
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   5580
            TabIndex        =   5
            Top             =   225
            Width           =   6345
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   3090
            TabIndex        =   3
            Top             =   240
            Width           =   1980
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1875
            Index           =   0
            Left            =   765
            TabIndex        =   11
            Top             =   5145
            Width           =   11160
            _cx             =   19685
            _cy             =   3307
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
            Caption         =   "SQL参数"
            Height          =   180
            Index           =   5
            Left            =   90
            TabIndex        =   10
            Top             =   5160
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "SQL语句"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   6
            Top             =   675
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "说明"
            Height          =   180
            Index           =   2
            Left            =   5145
            TabIndex        =   4
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "名称"
            Height          =   180
            Index           =   1
            Left            =   2685
            TabIndex        =   2
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "编码"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   0
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(C)"
         Height          =   350
         Left            =   10950
         TabIndex        =   13
         Top             =   7260
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9675
         TabIndex        =   12
         Top             =   7260
         Width           =   1100
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1470
      TabIndex        =   14
      Top             =   75
      Width           =   1575
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
Attribute VB_Name = "frmTableEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private mintParsCount As Long
Private mblnContiune As Boolean
Private mlngModualCode As Long
Private mstrBusiness As String

Private WithEvents mclsVsf As zlVSFlexGrid.clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Event AfterNewData(ByVal DataKey As String)
Public Event AfterModifyData(ByVal DataKey As String)
Public Event AfterDeleteData(ByVal DataKey As String)
Public Event Forward(ByRef DataKey As String, ByRef Cancel As Boolean)
Public Event Backward(ByRef DataKey As String, ByRef Cancel As Boolean)

'######################################################################################################################

Public Function InitDialog(ByVal frmParent As Object, ByVal lngModualCode As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mlngModualCode = lngModualCode
    
    InitDialog = True
    
End Function

Public Sub NewData(ByVal strBusiness As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 1
    mstrDataKey = ""
    mstrBusiness = strBusiness
    
    Me.Caption = "新增业务信息"
        
    Call InitData
    Call InitGrid
    Call InitCommandBar
    
    txt(1).Text = gclsBusiness.GetMaxUserTableCode("ZLHIS_USER_")
    
    mblnDataChanged = False
    
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub ModifyData(ByVal strBusiness As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mbytMode = 2
    mstrDataKey = strDataKey
    mstrBusiness = strBusiness
    
    Call InitData
    Call InitGrid
    Call InitCommandBar
    
    txt(1).Text = gclsBusiness.GetMaxCode("zlmip_table", "tab_code")
    
    Me.Caption = "修改业务信息"
        
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
        
    If gclsBusiness.TableEdit("Delete", mrsPara) Then
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
    
    Set rsTmp = gclsBusiness.TableStruct()
    If Not (rsTmp Is Nothing) Then
        txt(0).MaxLength = rsTmp("tab_title").DefinedSize
        txt(1).MaxLength = rsTmp("tab_code").Precision - Len(txtCode.Text)
        txt(2).MaxLength = rsTmp("tab_note").DefinedSize
        txtSQL.MaxLength = rsTmp("tab_sqltext").DefinedSize
    End If
    
    InitData = True
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '初始网格控件
    Set mclsVsf = New zlVSFlexGrid.clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf(0), True, False, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 270, flexAlignLeftCenter, flexDTString, , "[序号]", False, False, False)
        
        Call .AppendColumn("参数命名", 1500, flexAlignLeftCenter, flexDTString, , "para_title", True)
        Call .AppendColumn("参数类型", 1500, flexAlignLeftCenter, flexDTString, , "para_type", True)
        Call .AppendColumn("参数缺省", 0, flexAlignLeftCenter, flexDTString, , "para_default", True, , , True)
        Call .AppendColumn("参数说明", 3000, flexAlignLeftCenter, flexDTString, , "para_note", True)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        .UpdateSerial
        .AppendRows = True
        
        Call .InitializeEdit(True, False, False)
        
        Call .InitializeEditColumn(.ColIndex("参数命名"), True, vbVsfEditText)
        Call .InitializeEditColumn(.ColIndex("参数类型"), True, vbVsfEditCombox, "字符|数值|日期")
        Call .InitializeEditColumn(.ColIndex("参数缺省"), True, vbVsfEditText)
        Call .InitializeEditColumn(.ColIndex("参数说明"), True, vbVsfEditText)
    End With
                
    InitGrid = True
    
End Function

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
    Set rsTmp = gclsBusiness.TableRead("id", rsCondition)
    If rsTmp.BOF = False Then
        txt(0).Text = zlCommFun.NVL(rsTmp("tab_title").Value)
        txt(1).Text = Replace(zlCommFun.NVL(rsTmp("tab_code").Value), "ZLHIS_USER_", "")
        txt(2).Text = zlCommFun.NVL(rsTmp("tab_note").Value)
        txtSQL.Text = zlCommFun.NVL(rsTmp("tab_sqltext").Value)
    End If
    
    Set rsTmp = gclsBusiness.TableParameterRead("tab_id", rsCondition)
    If rsTmp.BOF = False Then Call mclsVsf.LoadGrid(rsTmp)
    
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
        
    mstrFindKey = zlDataBase.GetPara("定位依据", ParamInfo.系统号, mlngModualCode, "名称")
    If mstrFindKey = "" Then mstrFindKey = "名称"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, True, , , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.名称"): objControl.Parameter = "名称"
    objControl.IconId = 1
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.编码"): objControl.Parameter = "编码"
    objControl.IconId = 1

    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "搜索")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Forward, "上一条")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Backward, "下一条")

    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, IIf(mbytMode = 1, "确定之继续新增", "确定之继续修改"), False)
    objControl.IconId = conMenu_View_UnCheck
    If mbytMode <> 1 Then objControl.flags = xtpFlagRightAlign

    
    txtLocation.Visible = (mbytMode = 2)
    
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
    Dim lngLoop As Long
    
    If Len(txt(0).Text) = 0 Then
        ShowSimpleMsg "业务信息的名称不能为空！"
        Call LocationObj(txt(0))
        Exit Function
    End If
    
    If Len(txt(1).Text) = 0 Then
        ShowSimpleMsg "业务信息的编码不能为空！"
        Call LocationObj(txt(1))
        Exit Function
    End If
    
    '检查编码是否为数字字符
    If zlCommFun.CheckStrType(Trim(txt(1).Text), 99, "0123456789") = False Then
        ShowSimpleMsg "编码必须为数字字符！"
        LocationObj txt(1)
        Exit Function
    End If
    
    If Len(txtSQL.Text) = 0 Then
        ShowSimpleMsg "业务信息的SQL语句不能为空！"
        Call LocationObj(txtSQL)
        Exit Function
    End If
        
    If VerfiySQL = False Then
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
    Dim intType As Integer
    Dim strLine As String
    Dim strTemp As String
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
    
    Call zlCommFun.SetParameter(rsPara, "data_code", mstrBusiness)
    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
    Call zlCommFun.SetParameter(rsPara, "tab_code", txtCode.Text & Trim(txt(1).Text))
    Call zlCommFun.SetParameter(rsPara, "tab_title", Trim(txt(0).Text))
    Call zlCommFun.SetParameter(rsPara, "tab_sqltext", Replace(txtSQL.Text, "'", "''"))
    Call zlCommFun.SetParameter(rsPara, "tab_note", Trim(txt(2).Text))
        
    '------------------------------------------------------------------------------------------------------------------
    With vsf(0)
        lngCount = 0
        strTemp = ""
        For lngLoop = 1 To .Rows - 1
            
            Select Case .TextMatrix(lngLoop, .ColIndex("参数类型"))
            Case "数值"
                intType = 1
            Case "字符"
                intType = 2
            Case "日期"
                intType = 3
            End Select
            
            strLine = lngLoop
            
            strLine = strLine & "," & .TextMatrix(lngLoop, .ColIndex("参数命名"))
            strLine = strLine & "," & intType
            strLine = strLine & "," & .TextMatrix(lngLoop, .ColIndex("参数缺省"))
            strLine = strLine & "," & .TextMatrix(lngLoop, .ColIndex("参数说明"))
                        
            If LenB(strTemp & ";" & strLine) > 3500 Then
                If strTemp <> "" Then
                    lngCount = lngCount + 1
                    strTemp = Mid(strTemp, 2)
                    Call zlCommFun.SetParameter(rsPara, "SQL参数_" & lngCount, strTemp)
                    strTemp = ""
                End If
            End If
            strTemp = strTemp & ";" & strLine
        Next
    End With
    If strTemp <> "" Then
        lngCount = lngCount + 1
        strTemp = Mid(strTemp, 2)
        Call zlCommFun.SetParameter(rsPara, "SQL参数_" & lngCount, strTemp)
    End If
    Call zlCommFun.SetParameter(rsPara, "SQL参数个数", lngCount)
    
    
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
    strLine = ""
    lngCount = 0
    Set rsTmp = gclsBusiness.GetSQLField(Trim(txtSQL.Text))
    If Not (rsTmp Is Nothing) Then
        If rsTmp.BOF = False Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                strLine = rsTmp("序号").Value
                strLine = strLine & "," & rsTmp("名称").Value
                strLine = strLine & "," & rsTmp("类型").Value
                
                If LenB(strTemp & ";" & strLine) > 3500 Then
                    If strTemp <> "" Then
                        lngCount = lngCount + 1
                        strTemp = Mid(strTemp, 2)
                        Call zlCommFun.SetParameter(rsPara, "SQL字段_" & lngCount, strTemp)
                        strTemp = ""
                    End If
                End If
                strTemp = strTemp & ";" & strLine
            
                rsTmp.MoveNext
            Loop
        End If
    Else
        GoTo errHand
    End If
    
    If strTemp <> "" Then
        lngCount = lngCount + 1
        strTemp = Mid(strTemp, 2)
        Call zlCommFun.SetParameter(rsPara, "SQL字段_" & lngCount, strTemp)
    End If
    Call zlCommFun.SetParameter(rsPara, "SQL字段个数", lngCount)
    
    
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '新增
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        
        SaveData = gclsBusiness.TableEdit("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '修改
        SaveData = gclsBusiness.TableEdit("UPDATE", rsPara)
    End Select
    
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

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
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        
        Dim strText As String
        Dim rsCondition As ADODB.Recordset
        Dim rsData As ADODB.Recordset
        Dim rs As ADODB.Recordset
        
        If txtLocation.Text <> "" Then
            
            txtLocation.Tag = ""
                        
            Set rsCondition = zlCommFun.CreateCondition
            
            Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
            Call zlCommFun.SetCondition(rsCondition, "FilterText", txtLocation.Text)
            Set rsData = gclsBusiness.TableRead("FilterData", rsCondition)
                        
            If zlCommFun.ShowPubSelect(Me, txtLocation, 2, "编码,900,0,1;名称,1800,0,0;说明,2400,0,0", Me.Name & "\业务信息表过滤", "请从下表中选择一个业务信息表", rsData, rs, , , , , , True) = 1 Then
                mstrDataKey = rs("id").Value
                Call ReadData(mstrDataKey)
                txtLocation.Tag = ""
            Else
                txtLocation.Tag = ""
                Call LocationObj(txtLocation, True)
                Exit Sub
            End If
                        
            Call LocationObj(txtLocation, True)
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
                        txt(1).Text = gclsBusiness.GetMaxUserTableCode("ZLHIS_USER_")
                        txt(2).Text = ""
                        txtSQL.Text = ""
                    End If
                    Call LocationObj(txt(0))
                    mblnDataChanged = False
                End If
            End If
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub cmdPara_Click()
    Dim strSQL As String
    Dim intLoop As Integer
    Dim rsSQLPara As ADODB.Recordset
    
    strSQL = TrimChar(RemoveNote(txtSQL.Text))
        
    If strSQL <> "" Then
            
        '检查参数定义
        '------------------------------------------------------------------------------------------------------------------
        If gclsBusiness.CheckSQLPara(strSQL) = False Then
            MsgBox "参数定义不正确！请检查参数括号是否配对,参数号是否都为数值,且连续编号！", vbInformation + vbOKOnly, ParamInfo.系统名称
            txtSQL.SetFocus
            Exit Sub
        End If
                
        Set rsSQLPara = gclsBusiness.GetSQLPara(strSQL)
                
        If rsSQLPara Is Nothing Then
            mclsVsf.ClearGrid
        Else
            If rsSQLPara.RecordCount = 0 Then
                mclsVsf.ClearGrid
            Else
                
                With vsf(0)
                    .Rows = rsSQLPara.RecordCount + 1
                    For intLoop = 1 To .Rows - 1
                        
                        If .TextMatrix(intLoop, .ColIndex("参数命名")) = "" Then
                            .TextMatrix(intLoop, .ColIndex("参数命名")) = "参数" & intLoop
                            .TextMatrix(intLoop, .ColIndex("参数类型")) = "字符"
                        End If
                        
                    Next
                    mclsVsf.UpdateSerial
                    mclsVsf.AppendRows = True
                End With
        
            End If
        End If
    End If
        
End Sub

Private Sub cmdVerfiy_Click()
    
    If VerfiySQL Then
        MsgBox "当前SQL是合法的！", vbInformation, Me.Caption
    End If
    
End Sub

Private Function VerfiySQL() As Boolean
    Dim intLoop As Integer
    Dim rsSQLPara As ADODB.Recordset
            
    '------------------------------------------------------------------------------------------------------------------
    Set rsSQLPara = New ADODB.Recordset
    With rsSQLPara
        .Fields.Append "序号", adTinyInt
        .Fields.Append "名称", adVarChar, 60
        .Fields.Append "类型", adVarChar, 10
        .Open
    End With
    With vsf(0)
        For intLoop = 1 To .Rows - 1
            If .TextMatrix(intLoop, .ColIndex("参数命名")) <> "" Then
                rsSQLPara.AddNew
                rsSQLPara("序号").Value = intLoop - 1
                rsSQLPara("名称").Value = .TextMatrix(intLoop, .ColIndex("参数命名"))
                rsSQLPara("类型").Value = .TextMatrix(intLoop, .ColIndex("参数类型"))
            End If
        Next
    End With
    
    VerfiySQL = gclsBusiness.CheckSQL(TrimChar(RemoveNote(txtSQL.Text)), rsSQLPara)
    
End Function

Private Function RemoveNote(ByVal strSQL As String) As String
    '功能：移除SQL语句中的注释
    '说明：只支持移除整行的注释
    Dim strTmp As String
    Dim i As Integer
    Dim arrLine() As String
    
    strSQL = Replace(strSQL, vbTab, " ")
    strSQL = Replace(strSQL, vbLf, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr, vbCrLf)
    
'    strSQL = Replace(strSQL, "'", "''")
    
    arrLine = Split(strSQL, vbCrLf)
    
    For i = 0 To UBound(arrLine)
        If Not Trim(arrLine(i)) Like "--*" Then
            RemoveNote = RemoveNote & vbCrLf & arrLine(i)
        End If
    Next
    RemoveNote = Mid(RemoveNote, 3)
End Function

Private Function TrimChar(ByVal strSQL As String) As String
'功能:去除字符串中连续的空格和回车(含两头的空格,回车),不去除TAB字符,哪怕是连续的
    Dim strTmp As String
    Dim i As Long
    Dim j As Long
    
    If Trim(strSQL) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(strSQL)
    
    strTmp = Replace(strTmp, "  ", " ")
    strTmp = Replace(strTmp, "  ", " ")
    
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)

    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Set mclsVsf = Nothing
    
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = vsf(0).Rows > mintParsCount
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    mblnDataChanged = True
        
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    Select Case Index
    Case 4
        
    Case Else
        zlControl.TxtSelAll txt(Index)
    End Select
    
    Select Case Index
    Case 0, 2, 4
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
        '
        
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim bytMode As Byte
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0

        Select Case Index
        Case 1
            If zlCommFun.FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        End Select
        
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
End Sub

Private Sub txtLocation_Change()
    txtLocation.Tag = "Changed"
End Sub

Private Sub txtLocation_GotFocus()
    zlControl.TxtSelAll txtLocation
End Sub

Private Sub txtLocation_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        KeyCode = 0
        txtLocation.Text = ""
        txtLocation.Tag = ""
    End If

End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        If txtLocation.Text <> "" Then
            txtLocation.Tag = ""
            
            Dim obj As CommandBarControl
            
            Set obj = cbsMain.FindControl(, conMenu_View_Filter, True)
            If obj.Enabled = True Then
                Call cbsMain_Execute(obj)
            End If

        End If
        txtLocation.Tag = ""
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txtLocation_Validate(Cancel As Boolean)
    If (txtLocation.Tag = "Changed") Then
        txtLocation.Tag = ""
    End If
End Sub

Private Sub txtSQL_Change()
    If mblnReading Then Exit Sub
    mblnDataChanged = True
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
    If mblnReading Then Exit Sub
    mblnDataChanged = True
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

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
    End Select
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub
