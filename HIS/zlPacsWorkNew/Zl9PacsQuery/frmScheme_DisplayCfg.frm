VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmScheme_DisplayCfg 
   BorderStyle     =   0  'None
   Caption         =   "数据显示设置"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdReset 
      Caption         =   "全部重置"
      Height          =   350
      Left            =   9360
      TabIndex        =   3
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清空设置"
      Height          =   350
      Left            =   9360
      TabIndex        =   2
      Top             =   1080
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDisSet 
      Height          =   5415
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   7455
      _cx             =   13150
      _cy             =   9551
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   350
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
   Begin VB.Image imgCheck 
      Height          =   255
      Left            =   840
      Picture         =   "frmScheme_DisplayCfg.frx":0000
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNoCheck 
      Height          =   255
      Left            =   840
      Picture         =   "frmScheme_DisplayCfg.frx":0372
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTag 
      AutoSize        =   -1  'True
      Caption         =   "列显示设置"
      Height          =   180
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frmScheme_DisplayCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnIsEdit As Boolean

Private mblnState As Boolean
Private WithEvents mobjSetRelated As frmSetRelated
Attribute mobjSetRelated.VB_VarHelpID = -1
Private mrsRecord As Recordset
Private Const M_STR_COLNAME = "列名称|列图标|隐藏列|不显示标题|隐藏数据显示|排序对照|按数字排序|启用定位|数据转换|列统计|行关联显示配置"
Private Const M_STR_CROOK = "√"
Private mobjSqlScheme As New clsSqlScheme
Private mstrQuerySql As String
Private mlngBackColor As Long   '已设置背景色的列的位置
Private mlngTimeOut As Long     '已设置闪烁超时的列的位置
Private mblnAddItem As Boolean

Private Enum ColTitle
    ct列名称 = 0
    ct列图标 = 1
    ct隐藏列 = 2
    ct不显示标题 = 3
    ct隐藏数据显示 = 4
    ct排序对照 = 5
    ct按数字排序 = 6
    ct启用定位 = 7
    ct数据转换 = 8
    ct列统计 = 9
    ct行关联显示配置 = 10
End Enum

Private Sub cmdClear_Click()
    Dim i As Long
    
    On Error GoTo errHandle
    
    For i = 1 To vsfDisSet.Rows - 1
        vsfDisSet.Cell(flexcpBackColor, i, 1, i, vsfDisSet.Cols - 1) = &H80000005
        Call DisDataDefault(i)
    Next
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdReset_Click()
    On Error GoTo errHandle
    
    Call ShowDisplaySet(mobjSqlScheme)
    Call RefreshDisplaySet(mstrQuerySql)
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    mblnAddItem = False
    Call GridInit(M_STR_COLNAME, vsfDisSet)
    Call GridShow
    Call RefreshWindowState(False)
    Call SetFontSize(gbytFontSize)
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lblTag.Move Me.ScaleLeft + 100, Me.ScaleTop + 100
    fraLine.Move lblTag.Left + lblTag.Width, lblTag.Top + lblTag.Height / 2, Me.ScaleWidth - fraLine.Left
    vsfDisSet.Move lblTag.Left, fraLine.Top + 200, Me.ScaleWidth - cmdClear.Width - 300, Me.ScaleHeight - vsfDisSet.Top - 350
    cmdClear.Move vsfDisSet.Left + vsfDisSet.Width + 100, vsfDisSet.Top
    cmdReset.Move cmdClear.Left, cmdClear.Top + cmdClear.Height + 200
End Sub

Private Sub GridShow()
    With vsfDisSet
        .ColComboList(ColTitle.ct列图标) = "..."
        .ColComboList(ColTitle.ct数据转换) = "..."
        .ColComboList(ColTitle.ct行关联显示配置) = "..."
'        .ColWidth(ColTitle.ct列名称) = 1200
'        .ColWidth(ColTitle.ct数据转换) = 2000
'        .ColWidth(ColTitle.ct隐藏数据显示) = 1200
'        .ColWidth(ColTitle.ct列统计) = 700
'        .ColWidth(ColTitle.ct隐藏列) = 700
'        .ColWidth(ColTitle.ct列图标) = 700
'        .ColWidth(ColTitle.ct不显示标题) = 1000
'        .ColWidth(ColTitle.ct排序对照) = 1200
        .FixedCols = 1
    End With
End Sub


Private Sub SetColWithd(ByVal bytSize As Long)
    With vsfDisSet
        Select Case bytSize
            Case 0
                .ColWidth(ColTitle.ct列名称) = 1200
                .ColWidth(ColTitle.ct数据转换) = 2000
                .ColWidth(ColTitle.ct隐藏数据显示) = 1200
                .ColWidth(ColTitle.ct列统计) = 700
                .ColWidth(ColTitle.ct隐藏列) = 700
                .ColWidth(ColTitle.ct列图标) = 700
                .ColWidth(ColTitle.ct不显示标题) = 1000
                .ColWidth(ColTitle.ct排序对照) = 1200
                .ColWidth(ColTitle.ct按数字排序) = 1000
            Case 1
                .ColWidth(ColTitle.ct列名称) = 1400
                .ColWidth(ColTitle.ct数据转换) = 2000
                .ColWidth(ColTitle.ct隐藏数据显示) = 1600
                .ColWidth(ColTitle.ct列统计) = 900
                .ColWidth(ColTitle.ct隐藏列) = 900
                .ColWidth(ColTitle.ct列图标) = 900
                .ColWidth(ColTitle.ct不显示标题) = 1350
                .ColWidth(ColTitle.ct排序对照) = 1450
                .ColWidth(ColTitle.ct按数字排序) = 1350
            Case 2
                .ColWidth(ColTitle.ct列名称) = 1600
                .ColWidth(ColTitle.ct数据转换) = 2000
                .ColWidth(ColTitle.ct隐藏数据显示) = 2000
                .ColWidth(ColTitle.ct列统计) = 1100
                .ColWidth(ColTitle.ct隐藏列) = 1100
                .ColWidth(ColTitle.ct列图标) = 1100
                .ColWidth(ColTitle.ct不显示标题) = 1700
                .ColWidth(ColTitle.ct排序对照) = 1700
                .ColWidth(ColTitle.ct按数字排序) = 1700
        End Select
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mobjSetRelated Is Nothing Then Unload mobjSetRelated
    
    Set mrsRecord = Nothing
    Set mobjSetRelated = Nothing
    Set mobjSqlScheme = Nothing
End Sub

Private Sub mobjSetRelated_ClearItemSet(ByVal lngItem As Long, ByVal lngRow As Long)
    Call ClearItemSet(lngItem, lngRow)
End Sub

Private Sub mobjSetRelated_IsItemSetted(ByVal lngItem As Long, lngRow As Long, strRowName As String)
    Call IsItemSetted(lngItem, lngRow, strRowName)
End Sub

Private Sub vsfDisSet_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    Dim objIconManage As frmIconManage
    Dim objCustomQueryForm As frmSetDataFrom
    Dim ObjScShowCfg As New clsScShowCfg
    Dim strPerformCol As String
    Dim strIconName As String
    Dim objIcon As Object
    Dim blnEdit As Boolean
    Dim i As Long
    
    On Error GoTo errHandle
    
    If Col = ColTitle.ct列图标 Then
        strIconName = vsfDisSet.Cell(flexcpData, Row, Col)
        
        Set objIconManage = New frmIconManage
        Set objIcon = objIconManage.ShowIconWindow(strIconName, Me, 1)
        If Not objIcon Is Nothing Then
            vsfDisSet.Cell(flexcpPicture, Row, Col) = objIcon
            vsfDisSet.Cell(flexcpPictureAlignment, Row, Col) = flexPicAlignCenterCenter
            mblnIsEdit = True
        End If

        vsfDisSet.Cell(flexcpData, Row, Col) = strIconName
        
        If Not objIconManage Is Nothing Then Unload objIconManage
        Set objIconManage = Nothing
    End If
    If Col = ColTitle.ct行关联显示配置 Then
        For i = 1 To vsfDisSet.Rows - 1
            If i <> Row Then
                strPerformCol = strPerformCol & "|" & vsfDisSet.TextMatrix(i, ColTitle.ct列名称)
            End If
        Next
        strPerformCol = "|    " & "|" & Mid(strPerformCol, 2)
        
        If Not IsObject(vsfDisSet.RowData(Row)) Then
            If mobjSetRelated Is Nothing Then
                Set mobjSetRelated = New frmSetRelated
            End If
            Set ObjScShowCfg = mobjSetRelated.ShowScRowRelation(Nothing, vsfDisSet.TextMatrix(Row, ColTitle.ct数据转换), strPerformCol, mblnState, blnEdit, Me)
        Else
            If mobjSetRelated Is Nothing Then
                Set mobjSetRelated = New frmSetRelated
            End If
            Set ObjScShowCfg = vsfDisSet.RowData(Row)
            Set ObjScShowCfg = mobjSetRelated.ShowScRowRelation(ObjScShowCfg, vsfDisSet.TextMatrix(Row, ColTitle.ct数据转换), strPerformCol, mblnState, blnEdit, Me)

        End If
        
        mblnIsEdit = blnEdit
        
        If Not mobjSetRelated Is Nothing Then Unload mobjSetRelated
        Set mobjSetRelated = Nothing
        
        If Not (ObjScShowCfg Is Nothing) And mblnState Then
            vsfDisSet.RowData(Row) = ObjScShowCfg
        End If
        If ObjScShowCfg.RowRelationCount > 0 Then
            vsfDisSet.TextMatrix(Row, Col) = "..."
        Else
            vsfDisSet.TextMatrix(Row, Col) = ""
        End If
        
    End If
    
    If Col = ColTitle.ct数据转换 Then
        Set objCustomQueryForm = New frmSetDataFrom
        strValue = vsfDisSet.TextMatrix(Row, Col)
        strValue = objCustomQueryForm.ShowSqlFromWindow(strValue, "", 3, mblnState, gbytFontSize, Me)
        vsfDisSet.TextMatrix(Row, Col) = strValue
        
        If Not objCustomQueryForm Is Nothing Then Unload objCustomQueryForm
        Set objCustomQueryForm = Nothing
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption

    If Not objIconManage Is Nothing Then Unload objIconManage
    Set objIconManage = Nothing
    
    If Not mobjSetRelated Is Nothing Then Unload mobjSetRelated
    Set mobjSetRelated = Nothing
    
    If Not objCustomQueryForm Is Nothing Then Unload objCustomQueryForm
    Set objCustomQueryForm = Nothing
End Sub

Private Sub vsfDisSet_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errHandle
    
    If mblnState And Not mblnAddItem Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub



Private Sub vsfDisSet_Click()
    Dim lngRow As Long
    Dim lngCol As Long

    On Error GoTo errHandle
    
    lngRow = vsfDisSet.Row
    lngCol = vsfDisSet.Col
    If mblnState Then
        If lngRow <= 0 Then Exit Sub
        
        If lngCol <> ColTitle.ct隐藏列 Then
            If vsfDisSet.Cell(flexcpData, lngRow, ColTitle.ct隐藏列) = 1 Then
                Exit Sub
            End If
        End If

        If lngCol = ColTitle.ct隐藏列 Or lngCol = ColTitle.ct按数字排序 Or lngCol = ColTitle.ct不显示标题 Or lngCol = ColTitle.ct隐藏数据显示 Or lngCol = ColTitle.ct启用定位 Or lngCol = ColTitle.ct列统计 Then
            If lngCol = ColTitle.ct列统计 Then
                If ColCount >= 2 Then
                    MsgBox "【列统计】不能超过2列，请检查。", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
            mblnIsEdit = True
            If vsfDisSet.Cell(flexcpData, lngRow, lngCol) = 1 Then
                vsfDisSet.Cell(flexcpData, lngRow, lngCol) = 0
                vsfDisSet.Cell(flexcpPicture, lngRow, lngCol) = imgNoCheck.Picture
                If lngCol = ColTitle.ct隐藏列 Then
                    vsfDisSet.Cell(flexcpBackColor, lngRow, 1, lngRow, vsfDisSet.Cols - 1) = &H80000005
                End If
            Else
                vsfDisSet.Cell(flexcpPicture, lngRow, lngCol) = imgCheck.Picture
                vsfDisSet.Cell(flexcpData, lngRow, lngCol) = 1
                If lngCol = ColTitle.ct隐藏列 Then
                    vsfDisSet.Cell(flexcpBackColor, lngRow, 1, lngRow, vsfDisSet.Cols - 1) = &HC0FFFF
                End If
            End If
        End If
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Function ColCount() As Long
    Dim i As Long
    Dim lngCount As Long
    
    For i = 1 To vsfDisSet.Rows - 1
        If vsfDisSet.Cell(flexcpData, i, ct列统计) = 1 And Not vsfDisSet.RowHidden(i) And i <> vsfDisSet.Row Then
            lngCount = lngCount + 1
            If lngCount = 2 Then
                ColCount = lngCount
                Exit Function
            End If
        End If
    Next
    ColCount = lngCount
End Function

Private Sub vsfDisSet_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo errHandle
    
    If Not mblnState Then Exit Sub
    If KeyAscii <> 8 Then Exit Sub
    lngRow = vsfDisSet.Row
    lngCol = vsfDisSet.Col
    
    If lngRow <= 0 Then Exit Sub
    
    Select Case lngCol
        Case ct列图标
            vsfDisSet.Cell(flexcpPicture, lngRow, lngCol) = Nothing
            vsfDisSet.Cell(flexcpData, lngRow, lngCol) = ""
    End Select
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub vsfDisSet_RowColChange()
    Dim strSortContrasCol As String
    Dim i As Long
    Dim lngCol As Long


    On Error GoTo errHandle
    If vsfDisSet.Row < 1 Then Exit Sub
    lngCol = vsfDisSet.Col
    vsfDisSet.Editable = flexEDKbdMouse

    If mblnState Then
        If lngCol = ColTitle.ct列名称 Or lngCol = ColTitle.ct按数字排序 Or lngCol = ColTitle.ct隐藏列 Or lngCol = ColTitle.ct不显示标题 Or lngCol = ColTitle.ct隐藏数据显示 Or lngCol = ColTitle.ct启用定位 Or lngCol = ColTitle.ct列统计 Then
            vsfDisSet.Editable = flexEDNone
        End If

        If lngCol = ColTitle.ct排序对照 Then
            For i = 0 To mrsRecord.Fields.Count - 1
                If UCase(Trim(vsfDisSet.TextMatrix(vsfDisSet.Row, ColTitle.ct列名称))) <> UCase(Trim(mrsRecord.Fields(i).Name)) Then
                    strSortContrasCol = strSortContrasCol & "|" & mrsRecord.Fields(i).Name
                End If
            Next
            strSortContrasCol = "|    " & "|" & Mid(strSortContrasCol, 2)
            vsfDisSet.ColComboList(ColTitle.ct排序对照) = strSortContrasCol
            If Len(strSortContrasCol) = 0 Then
                vsfDisSet.Editable = flexEDNone
            End If
        End If

        If Val(vsfDisSet.Cell(flexcpData, vsfDisSet.Row, ColTitle.ct隐藏列)) = 1 Then
            vsfDisSet.Editable = flexEDNone
        End If
    Else
        If Not (lngCol = ColTitle.ct数据转换 Or lngCol = ColTitle.ct行关联显示配置) Then
            vsfDisSet.Editable = flexEDNone
        End If
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Public Sub SetShowCfg(objSqlScheme As clsSqlScheme)
'写入显示配置
    Dim ObjScShowCfg As clsScShowCfg
    Dim objScRowRelation As clsScRowRelation
    Dim objShow As clsScShowCfg
    Dim i As Long
    Dim j  As Long

    For i = 1 To vsfDisSet.Rows - 1
        Set objShow = New clsScShowCfg

        
        If vsfDisSet.RowHidden(i) = False And Len(vsfDisSet.TextMatrix(i, ColTitle.ct列名称)) > 0 And Not IsNoneSetRow(i) Then
            Set ObjScShowCfg = New clsScShowCfg
            
            If IsObject(vsfDisSet.RowData(i)) Then
                Set objShow = vsfDisSet.RowData(i)
            End If
            
            With ObjScShowCfg
                '显示设置
                .Name = vsfDisSet.TextMatrix(i, ColTitle.ct列名称)
                .Icon = vsfDisSet.Cell(flexcpData, i, ColTitle.ct列图标)
                .HiddenCol = IIf(Val(vsfDisSet.Cell(flexcpData, i, ColTitle.ct隐藏列)) = 1, True, False)
                .HiddenTitle = IIf(Val(vsfDisSet.Cell(flexcpData, i, ColTitle.ct不显示标题)) = 1, True, False)
                .HiddenData = IIf(Val(vsfDisSet.Cell(flexcpData, i, ColTitle.ct隐藏数据显示)) = 1, True, False)
                .SortContrastCol = Trim(vsfDisSet.TextMatrix(i, ColTitle.ct排序对照))
                .UseListLocate = IIf(Val(vsfDisSet.Cell(flexcpData, i, ColTitle.ct启用定位)) = 1, True, False)
                .DataConvert = vsfDisSet.TextMatrix(i, ColTitle.ct数据转换)
                .IsTotal = IIf(Val(vsfDisSet.Cell(flexcpData, i, ColTitle.ct列统计)) = 1, True, False)
                .IsNumerSort = IIf(Val(vsfDisSet.Cell(flexcpData, i, ColTitle.ct按数字排序)) = 1, True, False)
                
                '行关联设置
                For j = 1 To objShow.RowRelationCount
                    Set objScRowRelation = New clsScRowRelation
                    objScRowRelation.TiggerData = objShow.RowRelation(j).TiggerData
                    objScRowRelation.Icon = objShow.RowRelation(j).Icon
                    objScRowRelation.IconPerformCol = objShow.RowRelation(j).IconPerformCol
                    objScRowRelation.IsStateIcon = objShow.RowRelation(j).IsStateIcon
                    objScRowRelation.RowFontColor = objShow.RowRelation(j).RowFontColor
                    objScRowRelation.RowBackColor = objShow.RowRelation(j).RowBackColor
                    objScRowRelation.CellFontColor = objShow.RowRelation(j).CellFontColor
                    objScRowRelation.CellBackColor = objShow.RowRelation(j).CellBackColor
                    objScRowRelation.ColorPerformCol = objShow.RowRelation(j).ColorPerformCol
                    objScRowRelation.FlickerTimeOut = objShow.RowRelation(j).FlickerTimeOut
                    objScRowRelation.TimeOutReferCol = objShow.RowRelation(j).TimeOutReferCol
                    .AddRowRelation objScRowRelation
                Next
            End With
            objSqlScheme.AddShowCfg ObjScShowCfg
        End If
    Next
End Sub

Private Function IsNoneSetRow(lngRow As Long) As Boolean
    IsNoneSetRow = False
    With vsfDisSet
        If Len(.Cell(flexcpData, lngRow, ColTitle.ct列图标)) > 0 Then
            Exit Function
        ElseIf Val(.Cell(flexcpData, lngRow, ColTitle.ct隐藏列)) = 1 Then
            Exit Function
        ElseIf Val(vsfDisSet.Cell(flexcpData, lngRow, ColTitle.ct不显示标题)) = 1 Then
            Exit Function
        ElseIf Val(vsfDisSet.Cell(flexcpData, lngRow, ColTitle.ct隐藏数据显示)) = 1 Then
            Exit Function
        ElseIf Len(Trim(vsfDisSet.TextMatrix(lngRow, ColTitle.ct排序对照))) > 0 Then
            Exit Function
        ElseIf Val(vsfDisSet.Cell(flexcpData, lngRow, ColTitle.ct按数字排序)) = 1 Then
            Exit Function
        ElseIf Val(vsfDisSet.Cell(flexcpData, lngRow, ColTitle.ct启用定位)) = 1 Then
            Exit Function
        ElseIf Len(.TextMatrix(lngRow, ColTitle.ct数据转换)) > 0 Then
            Exit Function
        ElseIf Val(vsfDisSet.Cell(flexcpData, lngRow, ColTitle.ct列统计)) = 1 Then
            Exit Function
        ElseIf Len(.TextMatrix(lngRow, ColTitle.ct行关联显示配置)) > 0 Then
            Exit Function
        End If
    End With
    IsNoneSetRow = True
End Function

Public Sub RefreshWindowState(blnState As Boolean)
    mblnState = blnState
    cmdClear.Enabled = blnState
    cmdReset.Enabled = blnState
    If blnState Then
        vsfDisSet.BackColor = &H80000005
    Else
        vsfDisSet.BackColor = &H8000000F
    End If
End Sub

Public Sub ShowDisplaySet(objSqlScheme As clsSqlScheme)
'显示
    Dim ObjScShowCfg As New clsScShowCfg
    Dim i As Long
    Dim strDataChange As String
    Dim strFile As String
    
    Set mobjSqlScheme = objSqlScheme
    vsfDisSet.Rows = 1
    For i = 1 To objSqlScheme.ShowCfgCount
        Set ObjScShowCfg = objSqlScheme.ShowCfg(i)
        With vsfDisSet
            If Len(Trim(ObjScShowCfg.Name)) > 0 Then
                .Rows = .Rows + 1
                .TextMatrix(i, ColTitle.ct列名称) = ObjScShowCfg.Name
                .Cell(flexcpData, i, ColTitle.ct列图标) = ObjScShowCfg.Icon
                
                If Len(ObjScShowCfg.Icon) > 0 Then
                    strFile = zlBlobRead(ObjScShowCfg.Icon)
                    If Len(strFile) > 0 Then
                        If Len(Dir(strFile)) > 0 Then
                            vsfDisSet.Cell(flexcpPicture, i, ColTitle.ct列图标) = LoadPicture(strFile)
                            vsfDisSet.Cell(flexcpPictureAlignment, i, ColTitle.ct列图标) = flexPicAlignCenterCenter
                            Kill strFile
                        End If
                    End If
                End If
                
                If NVL(ObjScShowCfg.HiddenCol, False) Then
                    .Cell(flexcpData, i, ColTitle.ct隐藏列) = 1
                    .Cell(flexcpPicture, i, ColTitle.ct隐藏列) = imgCheck.Picture
                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &HC0FFFF
                Else
                    .Cell(flexcpData, i, ColTitle.ct隐藏列) = 0
                    .Cell(flexcpPicture, i, ColTitle.ct隐藏列) = imgNoCheck.Picture
                End If
                .Cell(flexcpPictureAlignment, i, ColTitle.ct隐藏列) = flexPicAlignCenterCenter
                
                If NVL(ObjScShowCfg.HiddenTitle, False) Then
                    .Cell(flexcpData, i, ColTitle.ct不显示标题) = 1
                    .Cell(flexcpPicture, i, ColTitle.ct不显示标题) = imgCheck.Picture
                Else
                    .Cell(flexcpData, i, ColTitle.ct不显示标题) = 0
                    .Cell(flexcpPicture, i, ColTitle.ct不显示标题) = imgNoCheck.Picture
                End If
                .Cell(flexcpPictureAlignment, i, ColTitle.ct不显示标题) = flexPicAlignCenterCenter
                
                If NVL(ObjScShowCfg.HiddenData, False) Then
                    .Cell(flexcpData, i, ColTitle.ct隐藏数据显示) = 1
                    .Cell(flexcpPicture, i, ColTitle.ct隐藏数据显示) = imgCheck.Picture
                Else
                    .Cell(flexcpData, i, ColTitle.ct隐藏数据显示) = 0
                    .Cell(flexcpPicture, i, ColTitle.ct隐藏数据显示) = imgNoCheck.Picture
                End If
                .Cell(flexcpPictureAlignment, i, ColTitle.ct隐藏数据显示) = flexPicAlignCenterCenter
                
                .TextMatrix(i, ColTitle.ct排序对照) = ObjScShowCfg.SortContrastCol
                
                If NVL(ObjScShowCfg.UseListLocate, False) Then
                    .Cell(flexcpData, i, ColTitle.ct启用定位) = 1
                    .Cell(flexcpPicture, i, ColTitle.ct启用定位) = imgCheck.Picture
                Else
                    .Cell(flexcpData, i, ColTitle.ct启用定位) = 0
                    .Cell(flexcpPicture, i, ColTitle.ct启用定位) = imgNoCheck.Picture
                End If
                .Cell(flexcpPictureAlignment, i, ColTitle.ct启用定位) = flexPicAlignCenterCenter
                
                If NVL(ObjScShowCfg.IsNumerSort, False) Then
                    .Cell(flexcpData, i, ColTitle.ct按数字排序) = 1
                    .Cell(flexcpPicture, i, ColTitle.ct按数字排序) = imgCheck.Picture
                Else
                    .Cell(flexcpData, i, ColTitle.ct按数字排序) = 0
                    .Cell(flexcpPicture, i, ColTitle.ct按数字排序) = imgNoCheck.Picture
                End If
                .Cell(flexcpPictureAlignment, i, ColTitle.ct按数字排序) = flexPicAlignCenterCenter
                
                .TextMatrix(i, ColTitle.ct数据转换) = ObjScShowCfg.DataConvert
                
                If NVL(ObjScShowCfg.IsTotal, False) Then
                    .Cell(flexcpData, i, ColTitle.ct列统计) = 1
                    .Cell(flexcpPicture, i, ColTitle.ct列统计) = imgCheck.Picture
                Else
                    .Cell(flexcpData, i, ColTitle.ct列统计) = 0
                    .Cell(flexcpPicture, i, ColTitle.ct列统计) = imgNoCheck.Picture
                End If
                .Cell(flexcpPictureAlignment, i, ColTitle.ct列统计) = flexPicAlignCenterCenter
                
                .TextMatrix(i, ColTitle.ct行关联显示配置) = IIf(ObjScShowCfg.RowRelationCount > 0, "...", "")
                .RowData(i) = ObjScShowCfg
            End If
        End With
    Next
End Sub


Public Sub RefreshDisplaySet(strQuerySql As String)
'刷新显示
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim ObjScShowCfg As New clsScShowCfg
    Dim strQueryItem As String
    Dim strCurItem As String
    Dim arrItem() As String
    Dim i As Long
    
    mstrQuerySql = strQuerySql
    objSqlParse.init strQuerySql
    Set mrsRecord = objQuery.GetQueryField(objSqlParse)
    If mrsRecord Is Nothing Then
        Exit Sub
    End If
    
    For i = 1 To vsfDisSet.Rows - 1
        If Not vsfDisSet.RowHidden(i) Then
            strCurItem = strCurItem & ",[" & vsfDisSet.TextMatrix(i, ColTitle.ct列名称) & "]"
            If Not HasSelectItem(vsfDisSet.TextMatrix(i, ColTitle.ct列名称), mrsRecord) Then
                vsfDisSet.RowHidden(i) = True
            End If
        End If
    Next
    
    strCurItem = Mid(strCurItem, 2)

    For i = 0 To mrsRecord.Fields.Count - 1
        If Len(Trim(mrsRecord.Fields(i).Name)) > 0 Then
            If InStr(strCurItem, "[" & mrsRecord.Fields(i).Name & "]") = 0 Then
                mblnAddItem = True
                vsfDisSet.AddItem mrsRecord.Fields(i).Name, vsfDisSet.Rows
                mblnAddItem = False
                Call DisDataDefault(vsfDisSet.Rows - 1)
            End If
        End If
    Next
End Sub

Private Function HasSelectItem(strItem As String, mrsRecord As Recordset) As Boolean
    Dim i As Long
    
    HasSelectItem = False

    For i = 0 To mrsRecord.Fields.Count - 1
        If strItem = mrsRecord.Fields(i).Name Then
            HasSelectItem = True
            Exit Function
        End If
    Next
End Function

Private Sub DisDataDefault(lngRow As Long)
    Dim i As Long
    
    With vsfDisSet
        For i = 1 To 9
            .TextMatrix(lngRow, i) = ""
            If i = ColTitle.ct隐藏列 Or i = ColTitle.ct隐藏数据显示 Or i = ColTitle.ct不显示标题 Or i = ColTitle.ct启用定位 Or i = ColTitle.ct列统计 Or i = ColTitle.ct按数字排序 Then
                .Cell(flexcpData, lngRow, i) = 0
                .Cell(flexcpPicture, lngRow, i) = imgNoCheck.Picture
                .Cell(flexcpPictureAlignment, lngRow, i) = flexPicAlignCenterCenter
            ElseIf i = ColTitle.ct列图标 Then
                .Cell(flexcpData, lngRow, i) = ""
                .Cell(flexcpPicture, lngRow, i) = Nothing
            ElseIf i = ColTitle.ct行关联显示配置 Then
                .TextMatrix(lngRow, i) = ""
                .RowData(lngRow) = ""
            End If
        Next
    End With
End Sub

Public Sub UnloadMe()
    Unload Me
End Sub

'判断该行关联设置是否已设置过
Private Sub IsItemSetted(lngItem As Long, ByRef lngRow As Long, ByRef strRowName As String)
    Dim ObjScShowCfg As New clsScShowCfg
    Dim blnSetted As Boolean
    Dim i As Long
    Dim j As Long
    
    For i = 1 To vsfDisSet.Rows - 1
        Set ObjScShowCfg = Nothing
        If vsfDisSet.RowHidden(i) = False Then
            If IsObject(vsfDisSet.RowData(i)) Then
                Set ObjScShowCfg = vsfDisSet.RowData(i)
            End If
            
            If ObjScShowCfg.RowRelationCount > 0 Then
                If i <> vsfDisSet.Row Then
                    Select Case lngItem
                        Case 0  '行背景色
                            For j = 1 To ObjScShowCfg.RowRelationCount
                                If ObjScShowCfg.RowRelation(j).RowBackColor > 0 Then
                                    blnSetted = True
                                    Exit For
                                End If
                            Next
                        Case 1  '闪烁超时
                            For j = 1 To ObjScShowCfg.RowRelationCount
                                If Val(ObjScShowCfg.RowRelation(j).FlickerTimeOut) > 0 Then
                                    blnSetted = True
                                    Exit For
                                End If
                            Next
                        Case 2  '行前景色
                            For j = 1 To ObjScShowCfg.RowRelationCount
                                If ObjScShowCfg.RowRelation(j).RowFontColor > 0 Then
                                    blnSetted = True
                                    Exit For
                                End If
                            Next
                    End Select
                    
                    If blnSetted Then
                        lngRow = i
                        strRowName = vsfDisSet.TextMatrix(i, ColTitle.ct列名称)
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
End Sub

Private Sub ClearItemSet(ByVal lngItem As Long, ByVal lngRow As Long)
    Dim ObjScShowCfg As New clsScShowCfg
    Dim i As Long
    
    If lngRow > 0 Then
        Set ObjScShowCfg = vsfDisSet.RowData(lngRow)
    End If
    
    Select Case lngItem
        Case 0  '行背景色
            For i = 1 To ObjScShowCfg.RowRelationCount
                ObjScShowCfg.RowRelation(i).RowBackColor = 0
            Next
        Case 1  '闪烁超时
            For i = 1 To ObjScShowCfg.RowRelationCount
                ObjScShowCfg.RowRelation(i).FlickerTimeOut = 0
            Next
        Case 2  '行前景色
            For i = 1 To ObjScShowCfg.RowRelationCount
                ObjScShowCfg.RowRelation(i).RowFontColor = 0
            Next
    End Select
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim lngCmdHeight As Long
    Dim lngCmdWithd As Long
    
    If bytFontSize = 9 Then
        lngCmdHeight = 350
        lngCmdWithd = 1100
        vsfDisSet.Width = 7455
        Call SetColWithd(0)
    ElseIf bytFontSize = 12 Then
        lngCmdHeight = 385
        lngCmdWithd = 1300
        vsfDisSet.Width = 7255
        Call SetColWithd(1)
    ElseIf bytFontSize = 15 Then
        lngCmdHeight = 420
        lngCmdWithd = 1500
        vsfDisSet.Width = 7055
        Call SetColWithd(2)
    End If
    
    lblTag.FontSize = bytFontSize
    vsfDisSet.FontSize = bytFontSize
    
    cmdClear.FontSize = bytFontSize
    cmdClear.Height = lngCmdHeight
    cmdClear.Width = lngCmdWithd
    cmdReset.FontSize = bytFontSize
    cmdReset.Height = lngCmdHeight
    cmdReset.Width = lngCmdWithd
    
    Call Form_Resize
    
    If Not mobjSetRelated Is Nothing Then
        mobjSetRelated.SetFontSize bytFontSize
    End If
    
    
End Sub

