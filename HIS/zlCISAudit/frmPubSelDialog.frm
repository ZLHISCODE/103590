VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin zl9Peis.usrPeisFilterCompanyPerson usrPeisFilterCompanyPerson 
      Height          =   420
      Left            =   330
      TabIndex        =   4
      Top             =   4575
      Visible         =   0   'False
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   741
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000011&
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
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      Height          =   1530
      Index           =   0
      Left            =   480
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
      Bindings        =   "frmPubSelDialog.frx":23B92
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
Private mrsData As ADODB.Recordset
Private mbytOK As Byte
Private mbytWinStyle As Byte
Private mInitSelectKey As String
Private mblnAllowResize As Boolean
Private mstrLvw As String
Private mblnMuliSelect As Boolean
Private mrsSelData As New ADODB.Recordset
Private gstrDBUser As String
Private gstrUserName As String
Private mblnShowAll As Boolean
Private mfrmMain As Object
Private mstrFilterControl As String
Private mobjStateInfo As CommandBarControl
Private Declare Function GetWindowRect& Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


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

        Set objControl = NewToolBar(objBar, xtpControlButton, 1, "全选", , , xtpButtonIcon)
        Set objControl = NewToolBar(objBar, xtpControlButton, 2, "全清", , , xtpButtonIcon)
        Set objControl = NewToolBar(objBar, xtpControlButton, 3, "所有", , , xtpButtonIconAndCaption)
        Set objControl = NewToolBar(objBar, xtpControlButton, 4, "确定", True, , xtpButtonIconAndCaption)
        

        If mstrFilterControl <> "" Then

            Set cbrCustom = NewToolBar(objBar, xtpControlCustom, 5, "过滤", True, , xtpButtonIconAndCaption)
            
            Select Case UCase(mstrFilterControl)
            Case UCase("usrPeisFilterCompanyPerson")
                cbrCustom.Handle = usrPeisFilterCompanyPerson.hWnd
            End Select
            
            
        End If
        
        Set mobjStateInfo = NewToolBar(objBar, xtpControlLabel, 0, " 状态信息")
        mobjStateInfo.Flags = xtpFlagRightAlign
        
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings

        .Add 0, vbKeyEscape, 99              '刷新
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

    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, Nothing)
    objPane.Title = "项目"
    objPane.Options = PaneNoCaption

    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)
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
    
    For lngCol = 0 To rsFrom.Fields.Count - 1
        
        Select Case rsFrom.Fields(lngCol).Type
        Case 131, 139

            rs.Fields.Append rsFrom.Fields(lngCol).Name, adBigInt, , adFldIsNullable
            
        Case Else

            rs.Fields.Append rsFrom.Fields(lngCol).Name, adVarChar, rsFrom.Fields(lngCol).DefinedSize, adFldIsNullable
            
        End Select
    Next
    
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
            For lngCol = 0 To rsFrom.Fields.Count - 1
                rsTo.Fields(lngCol) = rsFrom.Fields(lngCol).Value
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
        
        vsf.RowData(vsf.Rows - 1) = mrsData("ID").Value
        Set vsf.Cell(flexcpPicture, vsf.Rows - 1, 0) = ils16.ListImages(1).Picture
        For lngLoop = 1 To vsf.Cols - 1
            If vsf.TextMatrix(0, lngLoop) <> "" Then
                vsf.TextMatrix(vsf.Rows - 1, lngLoop) = NVL(mrsData(vsf.TextMatrix(0, lngLoop)).Value, "")
            End If
        Next
        
        mrsData.MoveNext
    Loop
    
    vsf.Cell(flexcpPictureAlignment, 0, 0, vsf.Rows - 1, 0) = 4
    
    Call AppendRows(Me, vsf)
    Exit Sub
    
errHand:
        
End Sub

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
    If ErrCenter = 1 Then
        Resume
    End If
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
    strTmp = strTmp & "," & Val(Mid(objNode.Key, 2))
    
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
    
    Do While Not mrsData.EOF
        If IIf(IsNull(mrsData("上级id").Value), 0, mrsData("上级id").Value) <> 0 Then
            Set objItem = tvw.Nodes.Add("K" & mrsData("上级ID").Value, tvwChild, "K" & mrsData("ID").Value, mrsData("名称").Value, "Close", "Open")
        Else
            Set objItem = tvw.Nodes.Add(, , "K" & mrsData("ID").Value, mrsData("名称").Value, "Close", "Open")
        End If
        mrsData.MoveNext
    Loop
     
    If tvw.Nodes.Count > 0 Then
        tvw.Nodes(1).Selected = True
        tvw.Nodes(1).EnsureVisible
        tvw.Nodes(1).Expanded = True
    End If
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
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
        
        vsf.RowData(vsf.Rows - 1) = mrsData("ID").Value
        Set vsf.Cell(flexcpPicture, vsf.Rows - 1, 0) = ils16.ListImages(1).Picture
        For lngLoop = 1 To vsf.Cols - 1
            If mblnMuliSelect And lngLoop = 1 Then
                
                On Error Resume Next
                
                vsf.TextMatrix(vsf.Rows - 1, lngLoop) = NVL(mrsData("选择").Value, "")
                
                If Val(NVL(mrsData("选择").Value, "")) = 1 Then
                    vsf.Cell(flexcpForeColor, vsf.Rows - 1, 2, vsf.Rows - 1, vsf.Cols - 1) = &HFF0000
                Else
                    vsf.Cell(flexcpForeColor, vsf.Rows - 1, 2, vsf.Rows - 1, vsf.Cols - 1) = 0
                End If
    
                On Error GoTo 0
                
            Else
                If vsf.TextMatrix(0, lngLoop) <> "" Then
                    vsf.TextMatrix(vsf.Rows - 1, lngLoop) = NVL(mrsData(vsf.TextMatrix(0, lngLoop)).Value, "")
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
                            Optional ByVal strFilterControl As String) As Byte
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
        
    Me.Left = X + 90
    Me.Top = Y + 90
    Me.Width = cx
    Me.Height = cy
    
    If InitData = False Then Exit Function
        
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
    If ErrCenter = 1 Then
        Resume
    End If
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
    
    Select Case mbytWinStyle
    Case 1
        
        dkpMain.Panes(2).Close
        
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
    
    If mblnLeftSelect Then tvw.LineStyle = tvwTreeLines
    
    Call RestoreFormState
    
    If mblnMuliSelect Then vsf.ColWidth(1) = 255
    
    tvw.Checkboxes = IIf(mbytWinStyle = 1, mblnMuliSelect, False)
    
    Select Case mstrFilterControl
    Case "usrPeisFilterCompanyPerson"
        Call usrPeisFilterCompanyPerson.InitData(mrsData)
    End Select
    
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
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String

    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

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
                strTmp = "搜索到 " & tvw.Nodes.Count & " 条结果"
            End If
        End If

        mobjStateInfo.Caption = strTmp

        cbsMain.RecalcLayout
    
    End Select


    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

'######################################################################################################################

Private Sub ReturnSelect()
    Dim lngLoop As Long
    Dim lngCol As Long
    
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
                    
                    mrsSelData.AddNew
                    For lngCol = 0 To mrsData.Fields.Count - 1
                        mrsSelData.Fields(lngCol).Value = mrsData.Fields(lngCol).Value
                    Next
                End If

            End If
        Next
    Else
        If mblnLeftSelect = False Then
            If vsf.RowData(vsf.Row) <> "" Then
                
                mrsData.Filter = ""
                If mbytWinStyle = 3 Then
                    mrsData.Filter = "末级=1 AND ID=" & vsf.RowData(vsf.Row)
                Else
                    mrsData.Filter = "ID=" & vsf.RowData(vsf.Row)
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
    
    Select Case Control.ID
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
    Case 99
        Unload Me
    End Select
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 1, 2
        Control.Visible = mblnMuliSelect
    Case 3
        Control.Visible = (mbytWinStyle = 3)
        Control.IconId = IIf(mblnShowAll, 103, 3)
    Case 4
        
    End Select
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
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
    
'    With picTitle
'        .Left = -15
'        .Top = -30
'        .Width = Me.ScaleWidth + 30
'    End With
    
'    Select Case mbytWinStyle
'    Case 1
'
'        With tvw
'            .Left = -15
'            .Top = 0
'            .Height = Me.ScaleHeight - stb.Height - .Top
'            .Width = Me.ScaleWidth - .Left
'        End With
'
'    Case 2, 4
'
'        With vsf
'            .Left = -15
'            .Top = 0
'            .Height = Me.ScaleHeight - stb.Height - .Top
'            .Width = Me.ScaleWidth - .Left
'        End With
'
'    Case 3
'        With tvw
'            .Left = -15
'            .Top = 0
'            .Height = Me.ScaleHeight - stb.Height - .Top
'            .Width = imgX.Left
'        End With
'
'        With vsf
'            .Left = imgX.Left + imgX.Width
'            .Top = tvw.Top
'            .Width = Me.ScaleWidth - .Left
'            .Height = tvw.Height
'        End With
'
'    End Select
'
'    With imgX
'        .Top = vsf.Top - 30
'        .Height = vsf.Height + 60
'    End With
'
'    With picDrag
'        .Left = Me.ScaleWidth - .Width - 30
'        .Top = Me.ScaleHeight - .Height - 30
'    End With
'
'    With tbr
'        .Left = stb.Left
'        .Top = stb.Top + 45
'    End With
'    Call AppendRows(Me, vsf)
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
        tvw.Move 0, 0, picPane(Index).Width - 15, picPane(Index).Height - 15
    Case 1
        vsf.Move 15, 0, picPane(Index).Width - 15, picPane(Index).Height - 15
        Call AppendRows(Me, vsf)
    End Select
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
    
    If mblnMuliSelect And vsf.Col <> 1 Then
        If Abs(Val(vsf.TextMatrix(vsf.Row, 1))) = 0 Then
            vsf.TextMatrix(vsf.Row, 1) = 1
            vsf.Cell(flexcpForeColor, vsf.Row, 2, vsf.Row, vsf.Cols - 1) = &HFF0000
        Else
            vsf.TextMatrix(vsf.Row, 1) = 0
            vsf.Cell(flexcpForeColor, vsf.Row, 2, vsf.Row, vsf.Cols - 1) = 0
        End If
        vsf.ForeColorSel = vsf.Cell(flexcpForeColor, vsf.Row, 2)
    Else
        Call ReturnSelect
    End If
End Sub


Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vsf_DblClick
    If mblnMuliSelect And KeyAscii = vbKeySpace And vsf.Col <> 1 Then Call vsf_DblClick
End Sub

Private Sub tvw_DblClick()

    If mblnLeftSelect = False And mbytWinStyle = 3 Then Exit Sub
    
    If tvw.SelectedItem Is Nothing Then Exit Sub
    Call ReturnSelect
    
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    '如果重复点击同一节点，则不再刷新数据
       
    If Node.Key <> mstrSvrKey Then
        mstrSvrKey = Node.Key
        
        '先清除数据再装载新数据
        Call LoadData(Node.Key)
        
        Call ExecuteCommand("刷新状态")
    End If

End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> 1)
End Sub
