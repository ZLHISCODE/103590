VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVsColSel 
   BorderStyle     =   0  'None
   Caption         =   "列设置"
   ClientHeight    =   3252
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2772
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3252
   ScaleWidth      =   2772
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsColSet 
      Height          =   3210
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2700
      _cx             =   4762
      _cy             =   5662
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      FormatString    =   $"frmVsColSel.frx":0000
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
      Begin VB.CommandButton cmdClose 
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2415
         TabIndex        =   1
         Top             =   30
         Width           =   240
      End
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

Public Function ShowColSet(ByVal frmMain As Form, ByVal strTittle As String, vsGrid As VSFlexGrid, _
                    Optional ByVal WinLeft As Double = 0, Optional ByVal WinTop As Double = 0, _
                    Optional ByVal lngTxtHeight As Long = 0) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:列设置接口
    '参数:
    '返回:列设置成功,返回true,否则返回False
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Err = 0: On Error Resume Next
    Set mVsGrid = vsGrid
    With mWindowPosition
        .Left = WinLeft
        .Top = WinTop
        .lngTxtH = lngTxtHeight
    End With
    Call LoadFulltoColSel
    Call ReSetWindowsFormLocal
    With Me
        .Show 1, frmMain
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
    dblRowsHeight = dblMinRowheight * vsColSet.Rows + 30
    
    dblColsWidth = IIf(dblColsWidth < MFRM_MIN_WIDTH, MFRM_MIN_WIDTH, dblColsWidth)
    
    If Me.Top + dblRowsHeight <= Screen.Height Then
        '窗体顶部+总行高度+小于等于屏幕高度。
        '看是否比最小高度还小,如果还小,就以最小度高为准
        If dblRowsHeight < MFRM_MIN_HEIGHT Then
            Me.Height = MFRM_MIN_HEIGHT
        Else
            Me.Height = dblRowsHeight
        End If
    Else
        '窗体顶部+总行数高度+批号的总高度大于屏幕高度,需要进一下检查
        '1.看上半屏幕高度是否比下半屏高度要高，如果，以上半屏的高度为准，否则以下半屏为准.
        If Screen.Height - Me.Top > Me.Top - mWindowPosition.lngTxtH - 15 Then
            '下半屏要大
            Me.Height = Screen.Height - Me.Top - lngTaskHeight
            '不能完全装下,只能根据情况来分配规格列表和批次列表的高度
         Else
            dblTemp = Me.Top - mWindowPosition.lngTxtH - 15
            Me.Top = Me.Top - mWindowPosition.lngTxtH - 15
            '上半屏要大
            If dblTemp - dblRowsHeight > 0 Then
                '上半屏能完全能装下
                Me.Height = dblRowsHeight
                If Me.Height < MFRM_MIN_HEIGHT Then Me.Height = MFRM_MIN_HEIGHT
            Else
                Me.Height = dblTemp
            End If
            Me.Top = Me.Top - Me.Height
        End If
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
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载列设置
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-09 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long, arrSplit As Variant
    Dim sngFrmHeight As Single, sngSelSumHeight As Single
    

    vsColSet.Clear 1
    vsColSet.Rows = 2
    With mVsGrid
        lngRow = 1
        For i = 0 To .Cols - 1
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            arrSplit = Split(.ColData(i) & "||", "||")
            
            If Trim(.ColKey(i)) <> "" And (Val(arrSplit(0)) = 1 Or Val(arrSplit(0)) = 0) Then
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("列名")) = .ColKey(i)
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("选择")) = IIf(.ColWidth(i) = 0 Or .ColHidden(i), False, True)
                vsColSet.RowData(lngRow) = Val(arrSplit(0))
                If Val(arrSplit(0)) = 1 Then
                    vsColSet.Cell(flexcpForeColor, lngRow, 0, lngRow, vsColSet.Cols - 1) = vbBlue
                End If
                vsColSet.Rows = vsColSet.Rows + 1
                lngRow = lngRow + 1
            End If
        Next
    End With
    If vsColSet.Rows > 2 Then vsColSet.Rows = vsColSet.Rows - 1
    sngFrmHeight = Me.ScaleHeight
    With vsColSet
        sngSelSumHeight = (.RowHeight(0) + 60) * (.Rows) + 60
        .Cell(flexcpBackColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000001
        .Cell(flexcpForeColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000005
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
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置显示列
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-09 17:31:22
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long
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
    Err = 0: On Error Resume Next
    With vsColSet
        .Left = ScaleLeft
        .Top = ScaleTop
        .Height = ScaleHeight
        .Width = ScaleWidth
        cmdClose.Left = .Left + .Width - cmdClose.Width - 10
    End With
    
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

Private Sub vsColSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsColSet
        Select Case Col
        Case .ColIndex("选择")
            'rowdata(i):1-固定,-1-不能选,0-可选
            If Val(.RowData(Row)) = 1 Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub
'Private Sub vsColSet_LostFocus()
'    vsColSet.Visible = False
'End Sub


