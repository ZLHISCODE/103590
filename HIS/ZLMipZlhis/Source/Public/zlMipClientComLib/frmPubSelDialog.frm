VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmPubSelDialog 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8535
   Icon            =   "frmPubSelDialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2130
      Index           =   4
      Left            =   1545
      ScaleHeight     =   2130
      ScaleWidth      =   4905
      TabIndex        =   6
      Top             =   3030
      Width           =   4905
      Begin VSFlex8Ctl.VSFlexGrid vsfPro 
         Height          =   1215
         Left            =   555
         TabIndex        =   7
         Top             =   435
         Width           =   4380
         _cx             =   7726
         _cy             =   2143
         Appearance      =   0
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
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   2025
      ScaleHeight     =   240
      ScaleWidth      =   1695
      TabIndex        =   4
      Top             =   150
      Width           =   1720
      Begin VB.TextBox txtLocation 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   30
         TabIndex        =   5
         Top             =   15
         Width           =   1700
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1530
      Index           =   1
      Left            =   4110
      ScaleHeight     =   1530
      ScaleWidth      =   2820
      TabIndex        =   1
      Top             =   1110
      Width           =   2820
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Left            =   555
         TabIndex        =   3
         Top             =   435
         Width           =   4380
         _cx             =   7726
         _cy             =   2143
         Appearance      =   0
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
         Rows            =   5
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
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1530
      Index           =   0
      Left            =   465
      ScaleHeight     =   1530
      ScaleWidth      =   2820
      TabIndex        =   0
      Top             =   1635
      Width           =   2820
      Begin MSComctlLib.TreeView tvw 
         Height          =   1395
         Left            =   105
         TabIndex        =   2
         Top             =   270
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   2461
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2145
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":69AC
            Key             =   "Cur"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":6C42
            Key             =   "All"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":6ED8
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":7472
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":780C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":E06E
            Key             =   "SelAll"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":148D0
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":1B132
            Key             =   "ClsAll"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":21994
            Key             =   "Icon"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSelDialog.frx":21BCE
            Key             =   "Flag"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   3030
      Top             =   510
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPubSelDialog.frx":21F68
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPubSelDialog.frx":2496A
      Left            =   360
      Top             =   15
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin VB.Image imgX 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1860
      Left            =   2325
      MousePointer    =   9  'Size W E
      Top             =   585
      Width           =   30
   End
End
Attribute VB_Name = "frmPubSelDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'（１）窗体级变量定义
'######################################################################################################################

Private Enum STYLE
    vbTreeView = 1
    vbListView = 2
    vbTreeListView = 3
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Enum mCol
    可选 = 0
    体检项目
    编码
    类别
    采集方式
    标本部位
    基本价格
    体检价格
    折扣
    计费明细
    采集方式id
    操作类型
    可选状态
End Enum

Private mblnStartUp As Boolean
Private mstrStatePath As String
Private mlngX As Long
Private mlngY As Long
Private mstrSvrKey As String
Private mlngSortColumn As Long
Private msglTxtH As Single
Private mstrSvrTag As String
Private mstrTitle As String
Private mblnLeftSelect As Boolean
Public mrsData As ADODB.Recordset
Private mbytOK As Byte
Public mbytWinStyle As Byte
Private mInitSelectKey As String
Private mblnAllowResize As Boolean
Private mblnSearch As Boolean
Private mstrLvw As String
Private mblnMuliSelect As Boolean
Private mrsSelData As New ADODB.Recordset
Private gstrDBUser As String
Private gstrUserName As String
Private mblnShowAll As Boolean
Private mfrmMain As Object
Private mstrFilterControl As String
Private mobjStateInfo As CommandBarControl
Private WithEvents mclsVsfPeis As clsVsf
Attribute mclsVsfPeis.VB_VarHelpID = -1
Private mstrSelectInfo As String
Private Declare Function GetWindowRect& Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Event Search(ByVal strValue As String)


'######################################################################################################################
'
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
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain

        cbsMain.VisualTheme = xtpThemeOffice2003
        
        With cbsMain.Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True
            .UseDisabledIcons = True
            .LargeIcons = True
            .SetIconSize True, 24, 24
            .SetIconSize False, 16, 16
            
        End With
        cbsMain.EnableCustomization False
        
        Set cbsMain.Icons = ImageManager.Icons
        cbsMain.Options.LargeIcons = False
    

        .ActiveMenuBar.Title = "菜单"
        .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
        .ActiveMenuBar.Visible = False
        

        Set objBar = .Add("工具栏", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap Or xtpFlagAlignBottom
        
        picPane(3).Visible = mblnSearch
        If mblnSearch Then
            Set objControl = NewToolBar(objBar, xtpControlLabel, 6, "")
            
            Set cbrCustom = NewToolBar(objBar, xtpControlCustom, 747, "")
            cbrCustom.Handle = picPane(3).hWnd
            cbrCustom.STYLE = xtpButtonIconAndCaption
            
            Set objControl = NewToolBar(objBar, xtpControlButton, 7, "搜索")
        End If
        Set objControl = NewToolBar(objBar, xtpControlButton, 1, "全选", , , xtpButtonIcon)
        Set objControl = NewToolBar(objBar, xtpControlButton, 2, "全清", , , xtpButtonIcon)
        Set objControl = NewToolBar(objBar, xtpControlButton, 3, "所有", , , xtpButtonIconAndCaption)
        Set objControl = NewToolBar(objBar, xtpControlButton, 4, "确定", True, , xtpButtonIconAndCaption)
        
        
        Set mobjStateInfo = NewToolBar(objBar, xtpControlLabel, 0, " 状态信息")
        
        mobjStateInfo.flags = xtpFlagRightAlign
        
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings

        .Add 0, vbKeyEscape, 99              '退出
        .Add FCONTROL, vbKeyA, 1            '全选
        .Add FSHIFT, vbKeyDelete, 2         '全清

    End With
    
    Exit Function
    
errHand:
    MsgBox Err.Description
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "分类"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, objPane)
    objPane.Title = "条件"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(3, 300, 100, DockBottomOf, objPane)
    objPane.Title = "项目"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(4, 300, 200, DockBottomOf, Nothing)
    objPane.Title = "体检项目"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)
    
    
End Sub

Private Sub InitVsfPro()
    
            '初始化VsflexGrid
        Set mclsVsfPeis = New clsVsf
        With mclsVsfPeis
            
            Call .Initialize(Me.Controls, vsfPro, True, False)      'frmPubResource.GetImageList(16)
            Call .ClearColumn
            Call .AppendColumn("可选", 600, flexAlignLeftCenter, flexDTBoolean, "", , True)
            Call .AppendColumn("体检项目", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("编码", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("类别", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("采集方式", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("标本部位", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("基本价格", 1080, flexAlignRightCenter, flexDTString, "0.00", , True)
            Call .AppendColumn("体检价格", 1080, flexAlignRightCenter, flexDTString, "0.00", , True)
            Call .AppendColumn("折扣", 1080, flexAlignRightCenter, flexDTString, "0.00%", , True)
            Call .AppendColumn("计费明细", 0, flexAlignLeftCenter, flexDTString, "", , True, False, False, True)
            Call .AppendColumn("采集方式id", 0, flexAlignLeftCenter, flexDTString, "", , True, False, False, True)
            Call .AppendColumn("操作类型", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("可选状态", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
           
            Call .InitializeEdit(True, True, True)
            Call .InitializeEditColumn(0, True, vbVsfEditCheck)
            .IndicatorCol = 0

            .AppendRows = True
        End With
    
End Sub

Private Function GetTrayHeight() As Long
    '******************************************************************************************************************
    '功能:获取任务栏的高度
    '******************************************************************************************************************
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function

Private Function CreateRec(ByVal rsFrom As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim lngCol As Long
    
    For lngCol = 0 To rsFrom.Fields.count - 1
        
        Select Case rsFrom.Fields(lngCol).type
        Case 131, 139

            rs.Fields.Append rsFrom.Fields(lngCol).Name, adBigInt, , adFldIsNullable
            
        Case Else

            rs.Fields.Append rsFrom.Fields(lngCol).Name, adVarChar, rsFrom.Fields(lngCol).DefinedSize + 20, adFldIsNullable
            
        End Select
    Next
    
    If mstrFilterControl = "usrPeisItem" Then
        rs.Fields.Append "可选", adVarChar, 500, adFldIsNullable
    End If
    
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    
    Set CreateRec = rs
End Function

Private Function CopyRec(ByVal rsFrom As ADODB.Recordset, ByVal rsTo As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngCol As Long
    
    If rsFrom.RecordCount > 0 Then
        rsFrom.MoveFirst
        Do While Not rsFrom.EOF
            rsTo.AddNew
            For lngCol = 0 To rsFrom.Fields.count - 1
                rsTo.Fields(lngCol) = rsFrom.Fields(lngCol).value
            Next
            rsFrom.MoveNext
        Loop
        rsTo.Update
    End If
End Function

Private Sub SaveFormState()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath <> "" Then
    
        SaveSetting "ZLSOFT", mstrStatePath, "宽度", Me.Width
        SaveSetting "ZLSOFT", mstrStatePath, "高度", Me.Height
        SaveSetting "ZLSOFT", mstrStatePath, "分隔条", imgX.Left
        SaveSetting "ZLSOFT", mstrStatePath, "所有下级", IIf(mblnShowAll, 1, 0)
        
        For lngLoop = 0 To vsf.Cols - 1
            strTmp = strTmp & ";" & vsf.ColWidth(lngLoop)
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        
        SaveSetting "ZLSOFT", mstrStatePath, "列宽", strTmp
    
    End If
End Sub

Private Sub RestoreFormState()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim strTmp As String
    
    On Error Resume Next
        
    If mstrStatePath <> "" Then
        
        If mblnAllowResize Then
            Me.Width = GetSetting("ZLSOFT", mstrStatePath, "宽度", Me.Width)
            Me.Height = GetSetting("ZLSOFT", mstrStatePath, "高度", Me.Height)
        End If
        
        imgX.Left = GetSetting("ZLSOFT", mstrStatePath, "分隔条", imgX.Left)
        mblnShowAll = (Val(GetSetting("ZLSOFT", mstrStatePath, "所有下级", 0)) = 1)
        
        For lngLoop = 0 To vsf.Cols - 1
            strTmp = strTmp & ";" & vsf.ColWidth(lngLoop)
        Next
        
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        
        strTmp = GetSetting("ZLSOFT", mstrStatePath, "列宽", strTmp)
        
        For lngLoop = 0 To vsf.Cols - 1
            vsf.ColWidth(lngLoop) = Val(Split(strTmp, ";")(lngLoop))
        Next
        
    End If
    
    '检查是否超过屏幕高和宽度
    Dim lngTrayH As Long
    Dim lngH0 As Long
    Dim lngH1 As Long
    
    lngTrayH = GetTrayHeight
    
    If Me.Left + Me.Width > Screen.Width Then
        If (Screen.Width - Me.Width) >= 0 Then
            Me.Left = Screen.Width - Me.Width
        Else
            Me.Left = 0
            Me.Width = Screen.Width
        End If
    End If
    
    If Me.Top + Me.Height > (Screen.Height - lngTrayH) Then
        
        If (Me.Top - Me.Height - msglTxtH) >= 0 Then
            '放在输入框的上面
            Me.Top = Me.Top - Me.Height - msglTxtH
        Else
            
            '分别计算放置上面和放置下面的高度,取最大高度
            lngH0 = Me.Top - msglTxtH
            lngH1 = Screen.Height - lngTrayH - Me.Top
            
            If lngH0 > lngH1 Then
            
                '上面高
                Me.Top = 0
                Me.Height = lngH0
            Else
                Me.Height = Screen.Height - lngTrayH - Me.Top
            End If
        End If
    End If
    
End Sub

Private Sub LoadData(ByVal strKey As String)
    '******************************************************************************************************************
    '功能：装载数据
    '参数：要装载数据的分类关键字
    '返回：
    '******************************************************************************************************************
    Dim strKeys As String
    Dim objItem As ListItem
    Dim lngLoop As Long
    Dim strFilter As String
    
    On Error GoTo errHand
    
    mrsData.Filter = ""

    '显示所有下级
    If Not (tvw.SelectedItem Is Nothing) And mblnShowAll Then
        strKeys = GetDownKey(tvw, tvw.SelectedItem)
        strFilter = ""
        If strKeys <> "" Then
            For lngLoop = 0 To UBound(Split(strKeys, ","))
                strFilter = strFilter & " Or (末级=1 And 上级id=" & Split(strKeys, ",")(lngLoop) & ")"
            Next
            strFilter = Mid(strFilter, 5)
        End If
        
        mrsData.Filter = strFilter
    Else
        mrsData.Filter = "末级=1 AND 上级id=" & Mid(strKey, 2)
    End If
    
    If mrsData.RecordCount > 0 Then mrsData.MoveFirst
    
    vsf.Rows = 2
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    Set vsf.Cell(flexcpPicture, 1, 0, 1, vsf.Cols - 1) = Nothing
    vsf.RowData(1) = ""
    
    Do While Not mrsData.EOF
        
        If vsf.RowData(vsf.Rows - 1) <> "" Then vsf.Rows = vsf.Rows + 1
        
        vsf.RowData(vsf.Rows - 1) = mrsData("ID").value
        Set vsf.Cell(flexcpPicture, vsf.Rows - 1, 0) = ils16.ListImages(1).Picture
        For lngLoop = 1 To vsf.Cols - 1
            If vsf.TextMatrix(0, lngLoop) <> "" Then
                vsf.TextMatrix(vsf.Rows - 1, lngLoop) = NVL(mrsData(vsf.TextMatrix(0, lngLoop)).value, "")
            End If
        Next
        
        mrsData.MoveNext
    Loop
    
    vsf.Cell(flexcpPictureAlignment, 0, 0, vsf.Rows - 1, 0) = 4
    
    Call AppendRows(Me, vsf)
    
'    If mstrFilterControl = "usrPeisItem" Then
'        If Val(vsf.RowData(vsf.Row)) > 0 Then
'            Call ReadPeisProject(Val(vsf.RowData(vsf.Row)))
'        End If
'        mstrSelectInfo = ""
'    End If
    
    Exit Sub
    
errHand:
        
End Sub

Public Function RefreshData()
    
    Call ExecuteCommand("刷新数据")
    Call ExecuteCommand("刷新状态")
'    Call zlControl.TxtSelAll(txtLocation)
End Function

Private Function AppendRows(frmParent As Form, ByVal objVsf As Object) As Boolean
    '******************************************************************************************************************
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '******************************************************************************************************************
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long

    Dim strName As String
    Dim strName2 As String

    Dim objTmp As Object
    Dim objAryX() As Object
    Dim objAryY() As Object

    On Error GoTo errHand

    ReDim objAryX(0)
    ReDim objAryY(0)

    If objVsf.Rows = 0 Then Exit Function

    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next

    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)

    On Error Resume Next
    Err = 0
    strName = objVsf.Name & objVsf.Index
    If Err <> 0 Then
        strName = objVsf.Name
    End If
    Err = 0
    On Error GoTo errHand

    '1.隐藏所有的线
    For Each objTmp In frmParent.Controls

        If TypeName(objTmp) = "Line" And Left(objTmp.Name, Len("ln" & strName & "_")) = "ln" & strName & "_" Then

            objTmp.Visible = False
            lngIndex = Val(Mid(objTmp.Name, Len("ln" & strName & "_") + 2))

            Select Case Mid(objTmp.Name, Len("ln" & strName & "_") + 1, 1)
            Case "X"
                ReDim Preserve objAryX(lngIndex)
                Set objAryX(lngIndex) = objTmp
            Case "Y"
                ReDim Preserve objAryY(lngIndex)
                Set objAryY(lngIndex) = objTmp
            End Select
        End If
    Next

    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If UBound(objAryY) < lngLoop Then ReDim Preserve objAryY(lngLoop)

        If objAryY(lngLoop) Is Nothing Then
            Set objTmp = frmParent.Controls.Add("VB.Line", "ln" & strName & "_Y" & lngLoop)
            Set objTmp.Container = objVsf
            Set objAryY(lngLoop) = objTmp
        End If

        With objAryY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With
    Next

   '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeightMin) < objVsf.Height

        lngIndex = lngIndex + 1

        If UBound(objAryX) < lngIndex Then ReDim Preserve objAryX(lngIndex)

        If objAryX(lngIndex) Is Nothing Then

            Set objTmp = frmParent.Controls.Add("VB.Line", "ln" & strName & "_X" & lngIndex)
            Set objTmp.Container = objVsf

            Set objAryX(lngIndex) = objTmp
        End If

        With objAryX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeightMin + IIf(lngIndex = 1, 30, 0)
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop

    AppendRows = True

    Exit Function

errHand:
'    MsgBox Err.Description
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Private Function GetDownKey(ByVal objTvw As TreeView, ByVal objNode As Node) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    Dim objNodeChild As Node
                
    strTmp = ""
    Set objNodeChild = objNode.Child
    strTmp = strTmp & "," & Trim(Mid(objNode.Key, 2))
    
    Do While Not (objNodeChild Is Nothing)
        strTmp = strTmp & "," & GetDownKey(objTvw, objNodeChild)
        Set objNodeChild = objNodeChild.Next
    Loop
    
    GetDownKey = IIf(strTmp <> "", Mid(strTmp, 2), "")
    
End Function

Private Sub ReadTreeData()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objItem As Node
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    mrsData.Filter = ""
    mrsData.Filter = "末级<>1"
    If mrsData.RecordCount > 0 Then mrsData.MoveFirst
    tvw.Nodes.Clear
    
    Do While Not mrsData.EOF
        'If IIf(IsNull(mrsData("上级id").value), 0, mrsData("上级id").value) <> 0 Then
        If IIf(IsNull(mrsData("上级id").value), "", mrsData("上级id").value) <> "" Then
            Set objItem = tvw.Nodes.Add("K" & mrsData("上级ID").value, tvwChild, "K" & mrsData("ID").value, mrsData("名称").value, "Close", "Open")
        Else
            Set objItem = tvw.Nodes.Add(, , "K" & mrsData("ID").value, mrsData("名称").value, "Close", "Open")
        End If
        mrsData.MoveNext
    Loop
     
    If tvw.Nodes.count > 0 Then
        tvw.Nodes(1).Selected = True
        tvw.Nodes(1).EnsureVisible
        tvw.Nodes(1).Expanded = True
    End If
    If Trim(txtLocation.Text) <> "" Then
        tvw.Nodes(tvw.Nodes.count).Selected = True
    End If
    Exit Sub
    
errHand:
    MsgBox Err.Description
'    If ErrCenter = 1 Then
        Resume
'    End If
End Sub

Private Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    '******************************************************************************************************************
    '功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    '参数：
    '返回：
    '******************************************************************************************************************
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Private Sub ReadListData()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objItem As ListItem
    
    '装载数据
    
    vsf.Rows = 2
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    vsf.RowData(1) = ""
    
    Do While Not mrsData.EOF
        
        If vsf.RowData(vsf.Rows - 1) <> "" Then vsf.Rows = vsf.Rows + 1
        
        vsf.RowData(vsf.Rows - 1) = mrsData("ID").value
        Set vsf.Cell(flexcpPicture, vsf.Rows - 1, 0) = ils16.ListImages(1).Picture
        For lngLoop = 1 To vsf.Cols - 1
            If mblnMuliSelect And lngLoop = 1 Then
                
                On Error Resume Next
                
                vsf.TextMatrix(vsf.Rows - 1, lngLoop) = NVL(mrsData("选择").value, "")
                
                If Val(NVL(mrsData("选择").value, "")) = 1 Then
                    vsf.Cell(flexcpForeColor, vsf.Rows - 1, 2, vsf.Rows - 1, vsf.Cols - 1) = &HFF0000
                Else
                    vsf.Cell(flexcpForeColor, vsf.Rows - 1, 2, vsf.Rows - 1, vsf.Cols - 1) = 0
                End If
    
                On Error GoTo 0
                
            Else
                If vsf.TextMatrix(0, lngLoop) <> "" Then
                    vsf.TextMatrix(vsf.Rows - 1, lngLoop) = NVL(mrsData(vsf.TextMatrix(0, lngLoop)).value, "")
                End If
            End If
            

        Next
        
        mrsData.MoveNext
    Loop
    
End Sub

Public Function ShowDialog(ByVal frmMain As Object, _
                            ByVal WinStyle As Byte, _
                            ByRef rsData As ADODB.Recordset, _
                            ByVal lvwHead As String, _
                            ByVal Title As String, _
                            ByVal X As Single, _
                            ByVal Y As Single, _
                            Optional ByVal cx As Single = 7200, _
                            Optional ByVal cy As Single = 4500, _
                            Optional ByVal CtlHeight As Single = 300, _
                            Optional ByVal InitKey As String = "", _
                            Optional ByVal RegPath As String = "", _
                            Optional ByVal LeftSelect As Boolean = False, _
                            Optional ByVal AllowResize As Boolean = True, _
                            Optional ByVal MuliSelect As Boolean = False, _
                            Optional ByVal strFilterControl As String, _
                            Optional ByVal strValue As String = "", Optional ByVal blnSearch As Boolean = False) As Byte
    '******************************************************************************************************************
    '功能：显示查询选择器
    '参数：
    '返回：0:取消选择;1:选择;2:无数据返回
    '******************************************************************************************************************
    On Error GoTo errHand
    
    mblnStartUp = True
    mbytOK = 0
    
    If rsData.BOF Then
        ShowDialog = 2
        Exit Function
    End If
    
    mstrFilterControl = strFilterControl
    Set mrsData = rsData
    
    mrsData.MoveFirst
    
    mblnLeftSelect = LeftSelect
    mstrSvrKey = ""
    mstrSvrTag = ""
    mlngSortColumn = 1
    msglTxtH = CtlHeight
    mstrTitle = Title
    mbytWinStyle = WinStyle              '1-TreeView;2-ListView;3-TreeView+ListView
    mInitSelectKey = InitKey
    mblnAllowResize = AllowResize
    mstrLvw = lvwHead
    mstrStatePath = RegPath
    
    mblnMuliSelect = MuliSelect
    mblnSearch = blnSearch
        
    Me.Left = X + 90
    Me.Top = Y + 90
    Me.Width = cx
    Me.Height = cy
    
    If InitData = False Then Exit Function
    If Trim(strValue) <> "" Then
        txtLocation = strValue
'        Call LocationObj(txtLocation, True)
    End If
    
    Me.Show 1, frmMain
    
    Set rsData = mrsSelData
    If Not (rsData Is Nothing) Then
        If rsData.State = adStateOpen Then
            If rsData.RecordCount > 0 Then rsData.MoveFirst
        End If
    End If
    ShowDialog = mbytOK
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Private Function InitData() As Boolean
    '******************************************************************************************************************
    '功能：显示查询选择器
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngUpKey As Long
    Dim objNode As Node
    Dim objItem As ListItem
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    Call InitCommandBar
    Call InitDockPannel
    
    If mstrFilterControl = "" Then
        dkpMain.Panes(2).Close
        dkpMain.Panes(4).Close
    End If

    Select Case mbytWinStyle
    Case 1
        
        dkpMain.Panes(3).Close
        
        mblnLeftSelect = True
        
    Case 2
    
        dkpMain.Panes(1).Close
        
    Case 4
        
        dkpMain.Panes(1).Close
        
    End Select
    
    Me.Caption = mstrTitle
    
    vsf.Rows = 2
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    Set vsf.Cell(flexcpPicture, 1, 0, 1, vsf.Cols - 1) = Nothing
    vsf.RowData(1) = ""
    
    tvw.Nodes.Clear
    
    '
    vsf.Cols = 0
    
    '增加图标列
    vsf.Cols = vsf.Cols + 1
    vsf.TextMatrix(1, vsf.Cols - 1) = ""
    Set vsf.Cell(flexcpPicture, 0, 0) = ils16.ListImages("Icon").Picture
    vsf.ColWidth(vsf.Cols - 1) = 255
    
    Dim lngStart As Long
    
    lngStart = 1
    
    If mblnMuliSelect Then
        '增加选择列
        vsf.Cols = vsf.Cols + 1
        vsf.TextMatrix(1, vsf.Cols - 1) = ""
        vsf.ColWidth(vsf.Cols - 1) = 255
        Set vsf.Cell(flexcpPicture, 0, 1) = ils16.ListImages("Flag").Picture
        vsf.ForeColorSel = 0
        vsf.ColDataType(1) = flexDTBoolean
        vsf.Editable = flexEDKbdMouse
        lngStart = 2
    Else
        vsf.ForeColorSel = &HFF0000
    End If
    
    vsf.Cols = vsf.Cols + UBound(Split(mstrLvw, ";")) + 1
    
    vsf.ExplorerBar = flexExSortShowAndMove
    
    For lngLoop = lngStart To vsf.Cols - 1
    
        vsf.TextMatrix(0, lngLoop) = Trim(Split((Split(mstrLvw, ";")(lngLoop - lngStart)), ",")(0))
        vsf.ColWidth(lngLoop) = Val(Split((Split(mstrLvw, ";")(lngLoop - lngStart)), ",")(1))
        vsf.ColAlignment(lngLoop) = Val(Split((Split(mstrLvw, ";")(lngLoop - lngStart)), ",")(2))
        vsf.ColFormat(lngLoop) = Trim(Split((Split(mstrLvw, ";")(lngLoop - lngStart)), ",")(3))
    Next
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    
'    If mblnLeftSelect Then tvw.LineStyle = tvwTreeLines
    
    Call RestoreFormState
    
    If mblnMuliSelect Then vsf.ColWidth(1) = 255
    
    tvw.Checkboxes = IIf(mbytWinStyle = 1, mblnMuliSelect, False)
    
    
    Call ExecuteCommand("刷新数据")
    
    '定位初始项InitSelectKey
    If mInitSelectKey <> "" Then
        
        On Error Resume Next
        
        If mbytWinStyle = 1 Or (mbytWinStyle = 3 And mblnLeftSelect) Then
            'tvw
            
            Set objNode = tvw.Nodes("K" & mInitSelectKey)
            
            If Not (objNode Is Nothing) Then
                objNode.Selected = True
                objNode.EnsureVisible
                
                Call tvw_NodeClick(objNode)
            End If
            
        Else
            'lvw
            If mbytWinStyle = 3 Then
                mrsData.Filter = ""
                mrsData.Filter = "ID=" & mInitSelectKey
                If mrsData.RecordCount > 0 Then
                    
                    lngUpKey = NVL(mrsData("上级id"), 0)
                    
                    Set objNode = tvw.Nodes("K" & lngUpKey)
                    
                    If Not (objNode Is Nothing) Then
                        objNode.Selected = True
                        objNode.EnsureVisible
                        
                        Call tvw_NodeClick(objNode)
                    End If
                    
                End If
                
                mrsData.Filter = ""
            End If
        End If
    End If
    
    Call ExecuteCommand("刷新状态")
    
    InitData = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
'    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String

    
    On Error GoTo errHand

'    Call SQLRecord(rsSQL)

    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"

        
    '--------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        Select Case mbytWinStyle
        Case 1
            
            Call ReadTreeData
            
        Case 2
            Call ReadListData
            Call AppendRows(Me, vsf)
        Case 3
            
            Call ReadTreeData
            If Not (tvw.SelectedItem Is Nothing) Then Call tvw_NodeClick(tvw.SelectedItem)
            
        End Select
                            
    '--------------------------------------------------------------------------------------------------------------
    Case "刷新状态"
        
        If tvw.SelectedItem Is Nothing Then
            mobjStateInfo.Caption = ""
        Else
            mobjStateInfo.Caption = tvw.SelectedItem.Text & "下有 " & vsf.Rows - 1 & " 条信息"
        End If
        
        strTmp = "没有搜索到结果"
        If dkpMain.Panes(2).Closed = False Then
        
            If dkpMain.Panes(1).Closed = False Then
                If tvw.SelectedItem Is Nothing Then
                    strTmp = "没有搜索到结果"
                Else
                    strTmp = "在" & tvw.SelectedItem.Text & "下搜索到 " & vsf.Rows - 1 & " 条结果"
                End If
            Else
                strTmp = "搜索到 " & vsf.Rows - 1 & " 条结果"
            End If
            
        ElseIf dkpMain.Panes(1).Closed = False Then
            If tvw.SelectedItem Is Nothing Then
                strTmp = "没有搜索到结果"
            Else
                strTmp = "搜索到 " & tvw.Nodes.count & " 条结果"
            End If
        End If

        mobjStateInfo.Caption = strTmp

        cbsMain.RecalcLayout
    
    End Select


    ExecuteCommand = True

    Exit Function
errHand:

'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call SaveErrLog

End Function

'######################################################################################################################

Private Sub ReturnSelect()
    Dim lngLoop As Long
    Dim lngCol As Long
    Dim strTemp() As String
    Dim i As Long
    Dim j As Long
    Dim strSelID As String
    
    Set mrsSelData = CreateRec(mrsData)
    
    If mblnMuliSelect Then
        For lngLoop = 1 To vsf.Rows - 1
            If vsf.RowData(lngLoop) <> "" Then
                If Abs(Val(vsf.TextMatrix(lngLoop, 1))) = 1 Then
                    mrsData.Filter = ""
                    If mbytWinStyle = 3 Then
                        mrsData.Filter = "末级=1 And ID=" & vsf.RowData(lngLoop)
                    Else
                        mrsData.Filter = "ID=" & vsf.RowData(lngLoop)
                    End If
                    
                    strSelID = ""
                    mrsSelData.AddNew
                    For lngCol = 0 To mrsData.Fields.count - 1
                        mrsSelData.Fields(lngCol).value = mrsData.Fields(lngCol).value
                    Next
                    
                     If mstrFilterControl = "usrPeisItem" And mstrSelectInfo <> "" Then
                        strTemp = Split(mstrSelectInfo, "|")
                        For i = 0 To UBound(strTemp) - 1
                            j = InStrRev(strTemp(i), ";")
                            If j > 0 Then
                                If vsf.RowData(lngLoop) = Val(Left(strTemp(i), j - 1)) Then
                                    strSelID = strSelID & strTemp(i) & "|"
                                End If
                            End If
                        Next
                        
                        mrsSelData.Fields(mrsData.Fields.count).value = strSelID
                        
                        
                    End If
                
                End If

            End If
        Next
    Else
        If mblnLeftSelect = False Then
            If vsf.RowData(vsf.Row) <> "" Then
                
                mrsData.Filter = ""
                If mbytWinStyle = 3 Then
                    mrsData.Filter = "末级=1 AND ID='" & vsf.RowData(vsf.Row) & "'"
                Else
                    mrsData.Filter = "ID='" & vsf.RowData(vsf.Row) & "'"
                End If
                
                Call CopyRec(mrsData, mrsSelData)
            End If
        Else
            If Not (tvw.SelectedItem Is Nothing) Then
                mrsData.Filter = ""
                If mbytWinStyle = 3 Then
                    mrsData.Filter = "末级=0 And ID=" & Val(Mid(tvw.SelectedItem.Key, 2))
                Else
                    mrsData.Filter = "ID=" & Val(Mid(tvw.SelectedItem.Key, 2))
                End If
                
                Call CopyRec(mrsData, mrsSelData)
            End If
        End If
    End If
    
    mbytOK = 1
    
    Unload Me
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long
    Dim bytMode As Byte
    Dim strText As String
    Dim intCol As Integer
    
    Select Case Control.id
    Case 1
        If mblnMuliSelect Then
            vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, 1) = 1
            vsf.Cell(flexcpForeColor, 1, 2, vsf.Rows - 1, vsf.Cols - 1) = &HFF0000
            vsf.ForeColorSel = vsf.Cell(flexcpForeColor, vsf.Row, 2)
        End If
    Case 2
        If mblnMuliSelect Then
            vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, 1) = 0
            vsf.Cell(flexcpForeColor, 1, 2, vsf.Rows - 1, vsf.Cols - 1) = 0
            vsf.ForeColorSel = vsf.Cell(flexcpForeColor, vsf.Row, 2)
        End If
    Case 3
    
        mblnShowAll = Not mblnShowAll
        
        If tvw.SelectedItem Is Nothing Then Exit Sub
        
        mstrSvrKey = ""
        Call tvw_NodeClick(tvw.SelectedItem)
        
        On Error Resume Next
        vsf.SetFocus
        On Error GoTo 0
        
    Case 4
        Call ReturnSelect
    Case 7 '搜索
        If Right(Me.Caption, 4) = "危害因素" Then  '危害因素特殊处理
            
            lngRow = -1
            strText = UCase(txtLocation.Text)
            bytMode = GetApplyModeCheck(strText)
            
            intCol = bytMode
    
            
            lngRow = FindRow(UCase(txtLocation.Text), intCol, bytMode, vsf.Row + 1)
            If lngRow = -1 Then
                lngRow = FindRow(UCase(txtLocation.Text), intCol, bytMode)
            End If
            If lngRow > 0 And vsf.Row <> lngRow Then
                vsf.Row = lngRow
                vsf.ShowCell vsf.Row, vsf.Col
            End If

'            Call LocationObj(txtLocation)
        ElseIf Right(Me.Caption, 4) = "受检人员" Then '受检人员2次过滤特殊处理
        
            lngRow = -1
            strText = UCase(txtLocation.Text)
            bytMode = GetApplyModeCheck(strText)
            
            intCol = bytMode
    
            lngRow = FindRow(UCase(txtLocation.Text), intCol, bytMode, vsf.Row + 1)
            If lngRow = -1 Then
                lngRow = FindRow(UCase(txtLocation.Text), intCol, bytMode)
            End If
            If lngRow > 0 And vsf.Row <> lngRow Then
                vsf.Row = lngRow
                vsf.ShowCell vsf.Row, vsf.Col
            End If

'            Call LocationObj(txtLocation)
            
        Else
            If InStr(txtLocation.Text, "'") > 0 Then
                txtLocation.Text = ""
            End If
            RaiseEvent Search(UCase(txtLocation.Text))
        End If
    Case 99
        Unload Me
    End Select
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case 1, 2
        Control.Visible = mblnMuliSelect
    Case 3
        Control.Visible = (mbytWinStyle = 3)
        Control.IconId = IIf(mblnShowAll, 103, 3)
    Case 4
        
    End Select
    
End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 3
        Item.Handle = picPane(1).hWnd
    Case 4
        Item.Handle = picPane(4).hWnd
    End Select
End Sub

Private Sub Form_Activate()
         
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
     '置缺焦点
    If vsf.Visible Then
        If mblnLeftSelect Then
            tvw.SetFocus
        Else
            vsf.SetFocus
        End If
    Else
        tvw.SetFocus
    End If
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'    If Shift = 4 And KeyCode = vbKeyA Then
'        If tbr.Buttons("全选").Visible Then
'            Call tbr_ButtonClick(tbr.Buttons("全选"))
'        End If
'    End If
'
'    If Shift = 4 And KeyCode = vbKeyD Then
'        If tbr.Buttons("全清").Visible Then
'            Call tbr_ButtonClick(tbr.Buttons("全清"))
'        End If
'    End If
'
'    If Shift = 4 And KeyCode = vbKeyO Then
'        If tbr.Buttons("确定").Visible Then
'            Call tbr_ButtonClick(tbr.Buttons("确定"))
'        End If
'    End If
'
'    If Shift = 4 And KeyCode = vbKeyS Then
'        If tbr.Buttons("所有").Visible Then
'            Call tbr_ButtonClick(tbr.Buttons("所有"))
'        End If
'    End If
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 2, 15, 75, Me.ScaleWidth, 75)
    
    dkpMain.RecalcLayout
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ObjPtr(mclsVsfPeis) > 0 Then Set mclsVsfPeis = Nothing
    Call SaveFormState
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgX.Left = imgX.Left + X
    
    If imgX.Left < 1500 Then imgX.Left = 1500
    If Me.Width - imgX.Left - imgX.Width < 1000 Then imgX.Left = Me.Width - imgX.Width - 1000
    
    Form_Resize
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        tvw.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 1
        vsf.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        Call AppendRows(Me, vsf)
    Case 4
        vsfPro.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub txtLocation_GotFocus()
'    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    Dim strText As String
    Dim bytMode As Byte
    Dim intCol As Integer
    
    If KeyCode = vbKeyReturn Then
        If InStr(txtLocation.Text, "'") > 0 Then
            KeyCode = 0
            txtLocation.Text = ""
            Exit Sub
        End If
        
        If Right(Me.Caption, 4) = "危害因素" Or Right(Me.Caption, 4) = "受检人员" Then '受检人员2次过滤特殊处理
        
            lngRow = -1
            strText = UCase(txtLocation.Text)
            bytMode = GetApplyModeCheck(strText)
            
            intCol = bytMode
    
            lngRow = FindRow(UCase(txtLocation.Text), intCol, bytMode, vsf.Row + 1)
            If lngRow = -1 Then
                lngRow = FindRow(UCase(txtLocation.Text), intCol, bytMode)
            End If
            If lngRow > 0 And vsf.Row <> lngRow Then
                vsf.Row = lngRow
                vsf.ShowCell vsf.Row, vsf.Col
            End If

'            Call LocationObj(txtLocation)
        Else
            RaiseEvent Search(UCase(txtLocation.Text))
        End If
    End If
End Sub

Private Sub usrPeisFilterCompanyPerson_AfterFilterData(rs As ADODB.Recordset)
    Set mrsData = rs
    Call ExecuteCommand("刷新数据")
    Call ExecuteCommand("刷新状态")
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Abs(Val(vsf.TextMatrix(Row, 1))) = 1 Then
        vsf.Cell(flexcpForeColor, Row, 2, Row, vsf.Cols - 1) = &HFF0000
    Else
        vsf.Cell(flexcpForeColor, Row, 2, Row, vsf.Cols - 1) = 0
    End If
    
    vsf.ForeColorSel = vsf.Cell(flexcpForeColor, Row, 2)
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    If mstrFilterControl = "usrPeisItem" Then
'        Call ReadPeisProject(Val(vsf.RowData(NewRow)))
'    End If
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(Me, vsf)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(Me, vsf)
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnMuliSelect Then
        vsf.ForeColorSel = vsf.Cell(flexcpForeColor, NewRow, 2)
    End If
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnMuliSelect Then
        Cancel = (Col < 2)
    Else
        Cancel = (Col < 1)
    End If
End Sub

Private Sub vsf_DblClick()
    
    If vsf.RowData(vsf.Row) = "" Then Exit Sub
    
    If mblnLeftSelect Then Exit Sub
    
    Call ReturnSelect

End Sub


Private Sub vsf_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
    Case vbKeyReturn
    
        Call vsf_DblClick
        
    Case vbKeySpace
        With vsf
            If mblnMuliSelect And .Col <> 1 Then
    
                If Abs(Val(.TextMatrix(.Row, 1))) = 0 Then
                    .TextMatrix(.Row, 1) = 1
                    .Cell(flexcpForeColor, .Row, 2, .Row, .Cols - 1) = &HFF0000
                Else
                    .TextMatrix(.Row, 1) = 0
                    .Cell(flexcpForeColor, .Row, 2, .Row, .Cols - 1) = 0
                End If
                .ForeColorSel = .Cell(flexcpForeColor, .Row, 2)
            End If
        End With
    
    End Select
    
    
End Sub

Private Sub tvw_DblClick()

    If mblnLeftSelect = False And mbytWinStyle = 3 Then Exit Sub
    
    If tvw.SelectedItem Is Nothing Then Exit Sub
    Call ReturnSelect
    
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    '如果重复点击同一节点，则不再刷新数据
       
    If Node.Key <> mstrSvrKey Or Node.Key = "K0" Then
        mstrSvrKey = Node.Key
        
        '先清除数据再装载新数据
        Call LoadData(Node.Key)
        
        Call ExecuteCommand("刷新状态")
    End If

End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> 1)
End Sub

Private Function FindRow(ByVal strKey As String, _
                        Optional ByVal intCol As Integer = -1, _
                        Optional ByVal bytMatch As Byte = 0, _
                        Optional ByVal lngStartRow As Long = -1) As Long
    '******************************************************************************************************************
    '功能:
    '参数:bytMatch=0，完全匹配；=1，左向匹配；=2；包含匹配；=3，右向匹配
    '返回:
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngLoop As Long
    
    If lngStartRow = -1 Then lngStartRow = vsf.FixedRows
    lngRow = -1
    
    If intCol = -1 Then
        lngRow = vsf.FindRow(strKey)
    Else
        Select Case bytMatch
        Case 0
            lngRow = vsf.FindRow(strKey, lngStartRow, intCol, False, True)
        Case 1
            lngRow = vsf.FindRow(strKey, lngStartRow, intCol, False, False)
        Case 2
            For lngLoop = lngStartRow To vsf.Rows - 1
                If InStr(UCase(vsf.TextMatrix(lngLoop, intCol)), UCase(strKey)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        Case 3
            For lngLoop = lngStartRow To vsf.Rows - 1
                If Right(UCase(vsf.TextMatrix(lngLoop, intCol)), Len(strKey)) = UCase(strKey) Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End Select
    End If
        
    FindRow = lngRow
    
End Function

Private Function GetApplyModeCheck(ByVal strText As String) As Byte
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
'    If IsNumeric(strText) Then
'        '是全数字，按编码查找
'
'        GetApplyModeCheck = 1
'
'    ElseIf CheckStrType(strText, 2) Then
'        '是全字母，按简码查找
'        GetApplyModeCheck = 3
'    Else
'        GetApplyModeCheck = 2
'    End If
End Function

Private Sub vsfPeis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsfPeis.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    
    If mblnStartUp Or NewRow = OldRow Then Exit Sub
    
    Call ExecuteCommand("显示收费项目")
    
End Sub

Private Sub vsfPro_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strTmp As String
    Dim i As Long
    If mstrFilterControl = "usrPeisItem" Then
        If Col = 0 Then
            
            If vsfPro.TextMatrix(Row, Col) = 0 Then
                strTmp = Val(vsf.RowData(vsf.Row)) & ";" & Val(vsfPro.RowData(Row)) & "|"
                mstrSelectInfo = Replace(mstrSelectInfo, strTmp, "")
                
                
                If InStrRev(mstrSelectInfo, Val(vsf.RowData(vsf.Row)) & ";") = 0 Then
                    vsf.TextMatrix(vsf.Row, 1) = 0
                End If
            Else
                strTmp = Val(vsf.RowData(vsf.Row)) & ";" & Val(vsfPro.RowData(Row)) & "|"
                mstrSelectInfo = Replace(mstrSelectInfo, strTmp, "")
                
                strTmp = Val(vsf.RowData(vsf.Row)) & ";" & Val(vsfPro.RowData(Row))
                mstrSelectInfo = mstrSelectInfo & strTmp & "|"
                
                '设置套餐可选
                
                vsf.TextMatrix(vsf.Row, 1) = 1
                
            End If
        End If
    End If
End Sub

Private Sub vsfPro_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsfPeis.AppendRows = True
End Sub

Private Sub vsfPro_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsfPeis.AppendRows = True
End Sub

Private Sub vsfPro_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mstrFilterControl = "usrPeisItem" Then
        If Val(vsfPro.TextMatrix(Row, mCol.可选状态)) = 0 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsfPro_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfPeis.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsfPro_KeyPress(KeyAscii As Integer)
    Call mclsVsfPeis.KeyPress(KeyAscii)
End Sub

Private Sub vsfPro_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsfPeis.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfPro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsfPeis.AutoAddRow(vsfPro.MouseRow, vsfPro.MouseCol)
    End Select
End Sub

Private Sub vsfPro_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsfPeis.EditSelAll
End Sub

Private Sub vsfPro_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfPeis.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub mclsVsfPeis_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsfPro.RowData(Row)) = 0)
End Sub

Private Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.id = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.type = xtpControlButton Or objControl.type = xtpControlPopup Then
            objControl.STYLE = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Private Function DockPannelInit(ByRef dkpMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

