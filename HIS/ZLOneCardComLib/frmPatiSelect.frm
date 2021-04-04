VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiSelect 
   Caption         =   "病人选择"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "frmPatiSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10710
   Begin VB.PictureBox picCmd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   10710
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6495
      Width           =   10710
      Begin VB.CommandButton cmdPatiMerge 
         Caption         =   "病人合并(&M)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   225
         TabIndex        =   6
         Top             =   120
         Width           =   1650
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   8505
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   9600
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   10710
      TabIndex        =   1
      Top             =   0
      Width           =   10710
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择一个项目,然后点击确定"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   120
         Width           =   2430
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsSelect 
      Height          =   5655
      Left            =   1665
      TabIndex        =   0
      Top             =   630
      Width           =   6825
      _cx             =   1968320263
      _cy             =   1968318199
      Appearance      =   2
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
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
      ExplorerBar     =   7
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
   Begin MSComctlLib.ImageList img16 
      Left            =   9450
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":6852
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpLine 
      Height          =   1815
      Left            =   255
      Top             =   1125
      Width           =   675
   End
End
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSys As Long, mlngModule As Long
Private mfrmMain As Form, mobjControl As Object
Private mrsBindings As ADODB.Recordset
Private mblnShowHead As Boolean, mstr参数名 As String, mstrHideCols As String '列1,列2,...
Private mblnOk As Boolean
Private mstrWinTittle As String, mstrNotes As String, mblnShowPatiMerge As Boolean, mblnNotShowWin As Boolean
Private mblnUserCancel As Boolean
Private mrsSelData As ADODB.Recordset
Private mrsOutSel  As ADODB.Recordset

'-------------------------------------------------------------------------------------------------------------------
'控件定位
Private Type ty_ctlObject_Locale
    '控件的位置
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    '下拉列表的最小高度和宽度
    minWidth As Single
    minHeight As Single
    
    '下接列表的实际位置
    DownLeft As Single
    DownTop As Single
    DownWidth As Single
    DownHeight As Single
    '屏模相关
    ScreenWidth As Single
    ScreenHeight As Single
End Type
Private mTyCtl_Locale As ty_ctlObject_Locale
'-------------------------------------------------------------------------------------------------------------------
'--API声明
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private mcnOracle As ADODB.Connection
Private mobjDataBase As clsDataBase

Private Sub SetWindowsProperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体特性
    '编制:刘兴洪
    '日期:2012-08-21 11:33:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cmdPatiMerge.Visible = mblnShowPatiMerge
    
    If mstrWinTittle <> "" Then Me.Caption = mstrWinTittle & IIf(InStr(mstrWinTittle, "选择") > 0, "", "选择")
    picInfo.Visible = False
    If mstrNotes <> "" Then
        lblInfo.Caption = mstrNotes
        picInfo.Visible = True
    End If
    If mblnNotShowWin Then
        Call FormSetCaption(Me, False, False)
    End If
    '存在病人合并或存在窗体时,显示下面的功能按钮
   picCmd.Visible = mblnNotShowWin = False Or mblnShowPatiMerge
   If mblnShowPatiMerge Then
        vsSelect.AllowSelection = True
        vsSelect.AllowBigSelection = True
        vsSelect.SelectionMode = flexSelectionListBox
   End If
End Sub

Private Function GetBoundData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取绑定数据
    '返回: 成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-21 14:54:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsBindings As ADODB.Recordset, strLog As String
  
    On Error GoTo errH
    
    Screen.MousePointer = 11
      
    mblnOk = False
    
    If mrsBindings Is Nothing Then
        Screen.MousePointer = 0
        Set mrsBindings = Nothing: Set mrsOutSel = Nothing
        mblnOk = False: Unload Me: Exit Function
        Exit Function
    End If
    
    '没有数据则返回
    If mrsBindings.EOF Then
        Screen.MousePointer = 0
        Set mrsBindings = Nothing: Set mrsOutSel = Nothing
        mblnOk = False: Unload Me: Exit Function
    End If
    If mrsBindings.RecordCount = 1 Then
        Set mrsOutSel = mrsBindings
        Screen.MousePointer = 0
        mblnOk = True: Unload Me: Exit Function
    End If
    Screen.MousePointer = 0
    GetBoundData = True
    Exit Function
errH:
    If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
    Call mobjDataBase.SaveErrLog
End Function

Public Function ShowSelect(ByVal frmMain As Form, ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, _
    ByVal lngModule As Long, ByVal objControl As Object, _
    ByVal rsSelData As ADODB.Recordset, _
    ByVal strWinTittle As String, _
    ByVal strNotes As String, _
    ByVal blnShowHead As Boolean, _
    ByVal blnShowPatiMerge As Boolean, _
    ByVal blnNotShowWin As Boolean, _
    ByVal str参数名 As String, _
    ByVal strHideCols As String, _
    ByRef rsOutSel As ADODB.Recordset, _
    ByRef blnUserCancel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择器入口
    '入参:frmMain-调用的主窗口
    '     lngSys-系统号
    '     lngModule-模块号
    '     objControl-控件对象(目前只支:textBox,Combox)
    '     rsSel-不能为空,主要字段,ID,......
    '     str参数-个性化保存的参数名.
    '     blnShowHead-是否显示现列头
    '出参:rsOutSel-选择后的记录集
    '返回:选中返回True, 否则返回False(可以按Esc进行返回)
    '编制:刘兴洪
    '日期:2009-01-01 15:35:30
    '问题:52913
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
   Set mrsBindings = rsSelData:   Set rsOutSel = Nothing
   If mrsBindings Is Nothing Then Exit Function
   If mrsBindings.RecordCount = 0 Then Exit Function
   
   mstrWinTittle = strWinTittle: mstrNotes = strNotes: mblnShowHead = blnShowHead
   mblnShowPatiMerge = blnShowPatiMerge: mblnNotShowWin = blnNotShowWin
    
    Set mfrmMain = frmMain: mlngSys = lngSys: mlngModule = lngModule
    mblnOk = False: Set mobjControl = objControl
    mblnShowHead = blnShowHead: mstr参数名 = str参数名: mstrHideCols = strHideCols
    
    On Error Resume Next
    If Not frmMain Is Nothing Then
        Me.Show 1, frmMain
    Else
        Me.Show 1
    End If
    On Error GoTo 0
    Set rsOutSel = mrsOutSel
    ShowSelect = mblnOk
    blnUserCancel = mblnUserCancel
End Function

Private Function SelectedItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择指定的子项
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-20 12:21:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    With vsSelect
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
    'If lngID = 0 Then Exit Function
    
    Set mrsOutSel = mrsBindings
    mrsOutSel.Filter = "ID=" & lngID
    If mrsOutSel.RecordCount = 0 Then Exit Function
    mblnUserCancel = False
    mblnOk = True
    Unload Me
    SelectedItem = True
End Function


Private Function zlBindingData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绑定数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-20 11:37:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngCol As Long
    
    If GetBoundData = False Then Exit Function
    
    '初始化网格
    With vsSelect
        .Redraw = flexRDNone
        Err = 0: On Error Resume Next
        Set vsSelect.Font = mobjControl.Font
        Set Me.Font = mobjControl.Font
        Err = 0: On Error GoTo 0
        Set .DataSource = mrsBindings
        If mrsBindings.EOF Then .Rows = 2: .Clear 1
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If UCase(.ColKey(i)) Like "*ID" Then .ColHidden(i) = True
            If mblnNotShowWin And UCase(.ColKey(i)) Like "*病人ID" Then .ColHidden(i) = False
            If InStr(1, "," & mstrHideCols & ",", "," & UCase(.ColKey(i)) & ",") > 0 Then .ColHidden(i) = True
            If .ColKey(i) Like "*余额" Then .ColAlignment(i) = flexAlignRightCenter
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
        Next
        .RowHidden(0) = False
        If mblnShowHead = False Then .RowHidden(0) = True
        
        '恢复网格控件顺序
        If mstr参数名 <> "" Then Call zl_vsGrid_Para_Restore(mlngSys, mlngModule, vsSelect, mstr参数名)
        
        '增加行号列
        If mblnNotShowWin Then
            .Cols = .Cols + 1
            lngCol = .Cols - 1
            .ColKey(lngCol) = "行号"
            .TextMatrix(0, lngCol) = "行"
            .ColAlignment(lngCol) = flexAlignCenterCenter
            For i = 1 To .Rows - 1
                .TextMatrix(i, lngCol) = i
            Next
            .AutoSize lngCol
            .ColPosition(lngCol) = 0
            
            '病人ID列加图标
            If .ColIndex("病人ID") > 0 Then
                Set .Cell(flexcpPicture, 1, .ColIndex("病人ID"), .Rows - 1) = img16.ListImages(1).Picture
                .AutoSize .ColIndex("病人ID")
            End If
        End If
        
        .Redraw = flexRDBuffered
    End With
    zlBindingData = True
End Function
Private Sub InitCtrlLocal()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定控件初始化控件位置
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-20 10:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngH As Long, lngW As Long, vRect As RECT, sngX As Single, sngY As Single
    
    '无对应窗体控件
    If mobjControl Is Nothing Then
        With mTyCtl_Locale
            .Top = 0
            .Left = 0
            .Width = Me.Width
            .Height = Me.Height
            .minHeight = vsSelect.RowHeight(0) * 5  '一般四行
            .minHeight = .minHeight + IIf(mblnNotShowWin = False Or mblnShowPatiMerge, picCmd.Height, 0)
            .minHeight = .minHeight + IIf(mstrNotes <> "", picInfo.Height, 0)
            .minWidth = .Width
            .ScreenHeight = GetSystemMetrics(SM_CYFULLSCREEN) * 15   '屏幕可用高度
            .ScreenWidth = Screen.Width  ' GetSystemMetrics(SM_CXVSCROLL) * 15 + 75  '屏幕可用宽度
        End With
        Exit Sub
    End If
   
   '通过Api计算出控件的相关坐标信息
    Select Case UCase(TypeName(mobjControl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, mobjControl)
        lngH = mobjControl.CellHeight
        lngW = mobjControl.CellWidth
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, mobjControl.MsfObj)
        lngH = mobjControl.MsfObj.CellHeight
        lngW = mobjControl.MsfObj.CellWidth
    Case Else
        vRect = GetControlRect(mobjControl.hWnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = mobjControl.Height
        lngW = mobjControl.Width
    End Select

    With mTyCtl_Locale
        .Top = sngY
        .Left = sngX
        .Width = lngW
        .Height = lngH
        .minHeight = vsSelect.RowHeight(0) * 5 '一般四行
        .minWidth = .Width
        .ScreenHeight = GetSystemMetrics(SM_CYFULLSCREEN) * 15   '屏幕可用高度
        .ScreenWidth = Screen.Width  ' GetSystemMetrics(SM_CXVSCROLL) * 15 + 75  '屏幕可用宽度
    End With
End Sub

Public Sub ReSetWindowsFormLocal()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置窗口的大小和位置
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-20 10:30:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblColsWidth As Double, dblRowsHeight As Double
    Dim dblTemp As Double, lngPicHeight As Long
    Dim i As Long
    
    '定位
    With mTyCtl_Locale
        .DownTop = .Top + .Height
        .DownLeft = .Left
        .DownWidth = .Width
    End With
    lngPicHeight = IIf(mstrNotes <> "", picInfo.Height, 0)
    lngPicHeight = lngPicHeight + IIf(mblnNotShowWin = False Or mblnShowPatiMerge, picCmd.Height, 0)
    
    
    '计算总列数的宽度
    dblColsWidth = 0
    For i = 0 To vsSelect.Cols - 1
        If Not vsSelect.ColHidden(i) Then
            dblColsWidth = dblColsWidth + vsSelect.ColWidth(i) + Screen.TwipsPerPixelX
        End If
    Next
    dblColsWidth = dblColsWidth + 300
    
    '计算总行数的高度
    dblRowsHeight = vsSelect.Cell(flexcpHeight, 0, 0, 0, 0)
    
    dblRowsHeight = (dblRowsHeight) * (vsSelect.Rows) + 100 + lngPicHeight
    If dblRowsHeight < mTyCtl_Locale.minHeight Then dblRowsHeight = mTyCtl_Locale.minHeight
    
    dblColsWidth = IIf(dblColsWidth < mTyCtl_Locale.minWidth, mTyCtl_Locale.minWidth, dblColsWidth)
    If vsSelect.Width > dblColsWidth - 300 Then
        vsSelect.ExtendLastCol = True
    Else
        vsSelect.ExtendLastCol = False
    End If
    
    If mblnNotShowWin = False Then
        If (mTyCtl_Locale.ScreenWidth - Me.Width) \ 2 > 0 Then
            Me.Left = (mTyCtl_Locale.ScreenWidth - Me.Width) \ 2
        End If
        If (mTyCtl_Locale.ScreenHeight - Me.Width) \ 2 > 0 Then
            Me.Top = (mTyCtl_Locale.ScreenHeight - Me.Width) \ 2
        End If
        Exit Sub
    End If
    
    With mTyCtl_Locale
        '计算窗体的Y坐标和下接高度
        If .ScreenHeight - (.Top + .Height + dblRowsHeight) < 0 Then
            '证明控件的行数要高度比控件以下的位置要大
            If dblRowsHeight < .Top Then
                '证明上部分能装下数据,因此控件下拉,放在上部分
               .DownHeight = dblRowsHeight
               .DownTop = .Top - dblRowsHeight
            ElseIf .Top > .ScreenHeight - (.Top + .Height) Then
                '检查上屏大还是下屏大,此分支表示上屏大
                .DownTop = 0
                .DownHeight = .Top
            Else '此分支表示下屏大
                .DownHeight = .ScreenHeight - (.Top + .Height)
            End If
        Else
            '证明下拉列表可以在控件下放显示
            .DownHeight = dblRowsHeight
        End If
        
        '计算窗体的Y坐标和下拉宽度
        If .ScreenWidth - .Left >= dblColsWidth Then
            '右屏能装入所有列宽
            .DownWidth = dblColsWidth
            .DownLeft = .Left
        Else
           If .Left + .Width >= dblColsWidth Then
                '右屏能装入所有列宽
                .DownLeft = .Left + .Width - dblColsWidth
                .DownWidth = dblColsWidth
           ElseIf .Left + .Width > .ScreenWidth - .Left Then
                '证明左边大于右边
                .DownWidth = .Left + .Width
                .DownLeft = 0
           Else
                '证明右边大于左边
                .DownWidth = .ScreenWidth - .Left
                .DownLeft = .Left
           End If
        End If
        
        '可以进行定位了
        Me.Left = .DownLeft
        Me.Top = .DownTop
        Me.Width = .DownWidth
        Me.Height = .DownHeight
    End With
    
End Sub

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15
        Y = objPoint.Y * 15 + objBill.Height
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Function zl_vsGrid_Para_Save(ByVal lngSys As Long, ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strKey As String, _
    Optional bln强制保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:保存vsFlex的宽度到注册表
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '返回:保存成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    
    Dim objDatabase As clsDataBase
    
    If zlGetOneDataBase(mcnOracle, objDatabase) = False Then Exit Function
    
    If bln强制保存 = False Then
        zl_vsGrid_Para_Save = True
        If Val(objDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    End If
    
    
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    objDatabase.SetPara strKey, strCol, lngSys, lngModule
    zl_vsGrid_Para_Save = True
End Function

Public Function zl_vsGrid_Para_Restore(ByVal lngSys As Long, ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strKey As String, _
    Optional bln强制恢复保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:从数据库中恢复网格的宽度等信息
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '     bln强制恢复保存-决定是否将保存的参数值,进行强制恢复
    '返回:恢复成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, arrtemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    Dim objDatabase As clsDataBase
    
    If zlGetOneDataBase(mcnOracle, objDatabase) = False Then Exit Function
    

    
    If bln强制恢复保存 = False Then
        zl_vsGrid_Para_Restore = True
        If Val(objDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    End If
    strParaValue = objDatabase.GetPara(strKey, lngSys, lngModule)
    
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    Err = 0: On Error GoTo errHand:
    arrReg = Split(strParaValue, "|")
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrtemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrtemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(arrtemp(1))
                If Val(arrtemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
errHand:
End Function

Private Sub cmdCancel_Click()
    mblnUserCancel = True
    Set mrsOutSel = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call vsSelect_DblClick
End Sub
Private Sub cmdPatiMerge_Click()
    Dim lng病人ID As Long
    Dim lng病人ID1 As Long
    Dim i As Long
    
    With vsSelect
        For i = 1 To .Rows - 1
            If .IsSelected(i) Then
                If lng病人ID <> 0 Then
                    If lng病人ID1 <> 0 Then Exit For
                    lng病人ID1 = Val(.TextMatrix(i, .ColIndex("病人ID")))
                    If lng病人ID1 <> 0 Then Exit For
                Else
                    lng病人ID = Val(.TextMatrix(i, .ColIndex("病人ID")))
                End If
            End If
        Next
    End With
    If frmMergePatient.zlShowPatiMerge(Me, mcnOracle, lng病人ID1, lng病人ID) = False Then Exit Sub
    
    '绑定数据
    If zlBindingData = False Then Exit Sub
    '初始化控件位置
    Call InitCtrlLocal
    '调整窗体位置
    Call ReSetWindowsFormLocal
End Sub

Private Sub Form_Activate()
    mblnUserCancel = True
    If vsSelect.Visible And vsSelect.Enabled Then vsSelect.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        Call SelectedItem
    Case vbKeyEscape
        Unload Me: Exit Sub
    End Select
End Sub

Private Sub Form_Load()
    Call SetWindowsProperty
    '绑定数据
    If zlBindingData = False Then Exit Sub
    '初始化控件位置
    Call InitCtrlLocal
    '调整窗体位置
    Call ReSetWindowsFormLocal
End Sub

Private Sub Form_Resize()
    Dim lngLineWidth As Long
    Err = 0: On Error Resume Next
    lngLineWidth = IIf(shpLine.Visible, 20, 0)
    With shpLine
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    With picInfo
        .Left = ScaleLeft + lngLineWidth
        .Top = ScaleTop + lngLineWidth
        .Width = ScaleWidth - lngLineWidth * 2
    End With
    With picCmd
        .Left = ScaleLeft + lngLineWidth
        .Top = ScaleHeight - .Height - lngLineWidth
        .Width = ScaleWidth - lngLineWidth * 2
    End With
    
    With vsSelect
        .Left = ScaleLeft + lngLineWidth
        .Width = ScaleWidth - lngLineWidth * 2
        .Top = ScaleTop + IIf(picInfo.Visible, picInfo.Top + picInfo.Height, 0)
        .Height = IIf(picCmd.Visible, picCmd.Top, ScaleHeight - lngLineWidth) - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If mstr参数名 <> "" Then zl_vsGrid_Para_Save mlngSys, mlngModule, vsSelect, mstr参数名
    Set mrsBindings = Nothing
    If Not mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
End Sub

Private Sub picCmd_Resize()
    Err = 0: On Error Resume Next
    With picCmd
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub

Private Sub vsSelect_AfterUserResize(ByVal Row As Long, ByVal col As Long)
    Call ReSetWindowsFormLocal
End Sub
Private Sub vsSelect_DblClick()
    Call SelectedItem
End Sub

Private Sub vsSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call SelectedItem
End Sub

Private Sub vsSelect_SelChange()
    '------------------------------------------------------------------------------
    '功能:设置合并按钮是否可用
    '编制:王吉
    '日期:20012/09/28
    '问题号:54215
    '------------------------------------------------------------------------------
     If cmdPatiMerge.Visible = True Then cmdPatiMerge.Enabled = vsSelect.SelectedRows >= 2
End Sub
