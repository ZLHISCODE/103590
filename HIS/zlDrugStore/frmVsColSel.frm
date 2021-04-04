VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVsColSel 
   BackColor       =   &H00FFEDDD&
   BorderStyle     =   0  'None
   Caption         =   "列设置"
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picModel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   380
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   2745
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
      Begin VB.OptionButton optModel 
         BackColor       =   &H80000005&
         Caption         =   "完整"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   80
         Width           =   700
      End
      Begin VB.OptionButton optModel 
         BackColor       =   &H80000005&
         Caption         =   "简洁"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   80
         Width           =   700
      End
      Begin VB.OptionButton optModel 
         BackColor       =   &H80000005&
         Caption         =   "自定义"
         Height          =   180
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   80
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.PictureBox picClose 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEDDD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Left            =   2520
      Picture         =   "frmVsColSel.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   50
      Width           =   250
   End
   Begin VB.PictureBox picUpDownCols 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEDDD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   0
      Left            =   2040
      Picture         =   "frmVsColSel.frx":0342
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   50
      Width           =   250
   End
   Begin VB.PictureBox picUpDownCols 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEDDD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Index           =   1
      Left            =   2280
      Picture         =   "frmVsColSel.frx":0684
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   50
      Width           =   250
   End
   Begin VSFlex8Ctl.VSFlexGrid vsColSet 
      Height          =   2970
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2820
      _cx             =   4974
      _cy             =   5239
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmVsColSel.frx":09C6
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   -1  'True
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEDDD&
      Caption         =   "列设置(按ESC退出)"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   75
      Width           =   1530
   End
End
Attribute VB_Name = "frmVsColSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type WinLocate
    Left As Double
    Top As Double
    lngTxtH As Long
End Type
Private mWindowPosition As WinLocate           '窗体位置
Private mVsGrid As VSFlexGrid
Private Const MFRM_MIN_WIDTH = 2775
Private Const MFRM_MIN_HEIGHT = 3255
Private Const SM_CXVSCROLL = 2
Private Const SPI_GETWORKAREA = 48

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Function ShowColSet(ByVal FrmMain As Form, ByVal strTittle As String, vsGrid As VSFlexGrid, _
                    Optional ByVal WinLeft As Double = 0, Optional ByVal WinTop As Double = 0, _
                    Optional ByVal lngTxtHeight As Long = 0) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:列设置接口
    '参数:
    '返回:列设置成功,返回true,否则返回False
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    
    err = 0: On Error Resume Next
    Set mVsGrid = vsGrid
    With mWindowPosition
        .Left = WinLeft
        .Top = WinTop
        .lngTxtH = lngTxtHeight
    End With
    Call LoadFulltoColSel
    Call ReSetWindowsFormLocal
    picModel.Visible = False
    With Me
        .Show 1, FrmMain
    End With
End Function

Public Sub ReSetWindowsFormLocal()
    '功能:重新设置窗口的大小和位置
    Dim dblColsWidth As Double, dblMinRowheight As Double, lngScrW As Long
    Dim lngTaskHeight As Long
    Dim dblRowsHeight As Double
    Dim dblTemp As Double
    Dim i As Long
    
    '定位
    With mWindowPosition
        Me.Left = .Left + 15
        Me.Top = .Top
    End With
    '隐藏表格头
    vsColSet.RowHidden(0) = True
    
    dblColsWidth = 0
    For i = 0 To vsColSet.Cols - 1
        If Not vsColSet.ColHidden(i) Then
            dblColsWidth = dblColsWidth + vsColSet.ColWidth(i) + 15
        End If
    Next
    dblMinRowheight = vsColSet.RowHeightMin
    lngTaskHeight = GetTaskbarHeight
    dblColsWidth = dblColsWidth + 300
    lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 75
    dblRowsHeight = dblMinRowheight * vsColSet.rows + 30
    
    dblColsWidth = IIf(dblColsWidth < MFRM_MIN_WIDTH, MFRM_MIN_WIDTH, dblColsWidth)
    
    '窗体高度
    If dblRowsHeight <= mWindowPosition.lngTxtH Then
        Me.Height = dblRowsHeight
    Else
        Me.Height = mWindowPosition.lngTxtH
    End If
    
    '窗体宽度定位
    '如果列宽总数小于等于当前窗体的宽度,则以列总数为准
    If dblColsWidth + Me.Left < Screen.Width Then
        '总列的宽度完全能显示
        Me.Width = dblColsWidth
    Else
        '检查是否左边屏幕大还是右边屏幕大
        If Screen.Width - Me.Left >= Me.Left Then
            '右边屏幕大
            Me.Width = Screen.Width - Me.Left
        Else
            Me.Left = Me.Left
            '左边屏幕大
            If dblColsWidth < Me.Left Then
                Me.Width = dblColsWidth
            Else
                Me.Width = Me.Left
            End If
            Me.Left = Me.Left - Me.Width
        End If
    End If
 
End Sub

Private Function LoadFulltoColSel() As Boolean
'--------------------------------------
'功能:加载列设置
'返回:成功,返回true,否则返回False
'--------------------------------------
    Dim i As Long, lngRow As Long
    Dim sngFrmHeight As Single, sngSelSumHeight As Single

    vsColSet.Clear 1
    vsColSet.rows = 2
    With mVsGrid
        lngRow = 1
        For i = 0 To .Cols - 1
            '（0：内部显示，可移动；1：内部隐藏，不可移动，不可显示；2：用户隐藏；3：用户显示(默认值)）
            If Trim(.ColKey(i)) <> "" And (Val(.ColData(i)) = 0 Or Val(.ColData(i)) = 2 Or Val(.ColData(i)) = 3) Then
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("列名")) = .TextMatrix(0, i)
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("选择")) = IIf(.ColWidth(i) = 0 Or .ColHidden(i), False, True)
                vsColSet.RowData(lngRow) = Val(.ColData(i))
                If Val(.ColData(i)) = 0 Then
                    vsColSet.Cell(flexcpForeColor, lngRow, 0, lngRow, vsColSet.Cols - 1) = vbBlue
                End If
                vsColSet.rows = vsColSet.rows + 1
                lngRow = lngRow + 1
            End If
        Next
    End With
    If vsColSet.rows > 2 Then vsColSet.rows = vsColSet.rows - 1
    sngFrmHeight = Me.ScaleHeight
    With vsColSet
        sngSelSumHeight = (.RowHeight(0) + 60) * (.rows) + 60
        '.Cell(flexcpBackColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000004
        '.Cell(flexcpForeColor, 0, 0, 0, vsColSet.Cols - 1) = vbBlack
        .AllowSelection = False
        .BackColorSel = &H8000000D
        .Row = 1
        .Visible = True
        .Editable = flexEDKbdMouse
        .ZOrder 0
        .Left = mVsGrid.Left + .Cell(flexcpWidth, 0, 0, 0, 0) + 30
        .Top = mVsGrid.Top + mVsGrid.RowHeight(0) + 15
        sngFrmHeight = sngFrmHeight - .Top
        If sngFrmHeight > sngSelSumHeight Then
            .Height = sngSelSumHeight
        Else
            .Height = IIf(sngFrmHeight < 0, 0, sngFrmHeight)
        End If
        .SetFocus
    End With
End Function

Private Function SetVsGridCol(ByVal strColKey As String, ByVal blnShow As Boolean, ByVal blnBatch As Boolean) As Boolean
'---------------------------------------------------------
'功能:设置显示列
'返回:成功,返回true,否则返回False
'---------------------------------------------------------
    Dim i As Long, lngRow As Long
    
    If InStr(1, strColKey, "效期") > 0 Then
        strColKey = "有效期"
    End If
    With mVsGrid
        .ColHidden(.ColIndex(strColKey)) = Not blnShow
        If .ColWidth(.ColIndex(strColKey)) = 0 Then .ColWidth(.ColIndex(strColKey)) = 1000
    End With
End Function

Private Sub cmdClose_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
 
Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With vsColSet
        .Left = ScaleLeft
        .Top = ScaleTop + 300
        .Height = ScaleHeight - 300 - IIf(picModel.Visible, picModel.Height, 0)
        .Width = ScaleWidth
    End With
    picClose.Left = ScaleWidth - picClose.Width - 10
    picUpDownCols(1).Left = picClose.Left - picUpDownCols(1).Width - 20
    picUpDownCols(0).Left = picUpDownCols(1).Left - picUpDownCols(0).Width - 20
    
    'picModel.Top = vsColSet.Top + vsColSet.Height
    'picModel.Width = vsColSet.Width
End Sub

Private Sub picClose_Click()
    Unload Me
End Sub

Private Sub picClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picClose.BorderStyle = 1
End Sub

Private Sub picClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picClose.BorderStyle = 0
End Sub

Private Sub picUpDownCols_Click(index As Integer)
    Dim intNewRow As Integer, intNewCol As Integer
    
    If vsColSet.RowSel < 1 Then Exit Sub
    '先行上下移动
    With vsColSet
        If index = 0 Then
            If .RowSel <= 1 Then
                intNewRow = 1
            Else
                intNewRow = .RowSel - 1
                '调整目标VSF列顺序
                DoEvents
                ChangeVSFColIndex .TextMatrix(.Row, .ColIndex("列名")), .TextMatrix(intNewRow, .ColIndex("列名"))
                DoEvents
            End If
        Else
            If .RowSel >= .rows - 1 Then
                intNewRow = .rows - 1
            Else
                intNewRow = .RowSel + 1
                '调整目标VSF列顺序
                DoEvents
                ChangeVSFColIndex .TextMatrix(.Row, .ColIndex("列名")), .TextMatrix(intNewRow, .ColIndex("列名"))
                DoEvents
            End If
        End If
        DoEvents
        .RowPosition(.RowSel) = intNewRow
        DoEvents
        .Row = intNewRow
    End With
    
End Sub

Private Sub ChangeVSFColIndex(ByVal strColKey As String, ByVal strColKeyTo As String)
'再调整目标VSF列顺序
    
    With mVsGrid
        .ColPosition(.ColIndex(strColKey)) = .ColIndex(strColKeyTo)
    End With
    
End Sub

Private Sub picUpDownCols_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    picUpDownCols(index).BorderStyle = 1
End Sub

Private Sub picUpDownCols_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    picUpDownCols(index).BorderStyle = 0
End Sub

Private Sub vsColSet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '修改后
    Dim strColKey As String, blnShow As Boolean
    With vsColSet
        Select Case Col
        Case .ColIndex("选择")
            blnShow = GetVsGridBoolColVal(vsColSet, Row, .ColIndex("选择"))
            Call SetVsGridCol(.TextMatrix(Row, .ColIndex("列名")), blnShow, IIf(.Tag = "Head", False, True))
        Case Else
        End Select
    End With
End Sub

Private Sub vsColSet_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        If Val(vsColSet.RowData(NewRow)) = 0 Then
            '以前景色作选择的前景色
            'vscolset.forecolorsel = vsColSet.Cell(flexcpForeColor, NewRow, 0, NewRow, vsColSet.Cols - 1)
            vsColSet.ForeColorSel = vbRed
        Else
            vsColSet.ForeColorSel = vbWhite
        End If
    End If
End Sub

Private Sub vsColSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsColSet
        Select Case Col
        Case .ColIndex("选择")
            If Val(.RowData(Row)) = 0 Then
                Cancel = True
            Else
                Cancel = False
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '返回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT
    err = 0: On Error GoTo ErrHand:
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
ErrHand:
End Function

Private Function GetVsGridBoolColVal(ByVal vsGrid As VSFlexGrid, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '功能:获取bool列的值
    '返回:是该单元格为true,返回true,否则返回False
    '------------------------------------------------------------------------------
    Dim strTemp As String
    err = 0: On Error GoTo ErrHand:
    With vsGrid
        strTemp = .TextMatrix(lngRow, lngCol)
    End With
    If UCase(strTemp) = UCase("True") Then
        GetVsGridBoolColVal = True: Exit Function
    End If
    GetVsGridBoolColVal = Val(strTemp) <> 0
    Exit Function
ErrHand:
End Function
