VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicPlanTemp 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList imgList16 
      Left            =   390
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanTemp.frx":0000
            Key             =   "FixedItem"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanTemp.frx":059A
            Key             =   "InvalidFixedItem"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanTemp.frx":0B34
            Key             =   "MonthItem"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanTemp.frx":10CE
            Key             =   "InvalidMonthItem"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanTemp.frx":1668
            Key             =   "WeekItem"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanTemp.frx":1C02
            Key             =   "InvalidWeekItem"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanTemp.frx":219C
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanTemp.frx":26D6
            Key             =   "ASC"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanTemp.frx":2C70
            Key             =   "DESC"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPlanColor 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   330
      ScaleHeight     =   1170
      ScaleWidth      =   2580
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   2610
      Begin VSFlex8Ctl.VSFlexGrid vsfPlanColor 
         Height          =   945
         Left            =   -15
         TabIndex        =   3
         Top             =   60
         Width           =   2625
         _cx             =   4630
         _cy             =   1667
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmClinicPlanTemp.frx":320A
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
   Begin VB.PictureBox picImage 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2610
      ScaleHeight     =   2175
      ScaleWidth      =   4725
      TabIndex        =   1
      Top             =   2265
      Visible         =   0   'False
      Width           =   4725
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfGridPrint 
      Height          =   1395
      Left            =   3630
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   3165
      _cx             =   5583
      _cy             =   2461
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
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfGridCopy 
      Height          =   1395
      Left            =   3960
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   3165
      _cx             =   5583
      _cy             =   2461
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
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmClinicPlanTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINT
     X As Long
     Y As Long
End Type
Private MP As POINTAPI
Private blnLoading As Boolean

Public Function GetPlanItemImage(ByVal strKey As String) As IPictureDisp
    '获取安排号源类型图像
    '入参：
    '   strKey 图像索引
    On Error GoTo errHandler
    Select Case UCase(strKey)
    Case UCase("InvalidFixedItem") '无效按固定排班号源"
        Set GetPlanItemImage = imgList16.ListImages("InvalidFixedItem").Picture
    Case UCase("FixedItem") '正常按固定排班号源"
        Set GetPlanItemImage = imgList16.ListImages("FixedItem").Picture
    Case UCase("InvalidMonthItem") '无效按月排班号源
        Set GetPlanItemImage = imgList16.ListImages("InvalidMonthItem").Picture
    Case UCase("MonthItem") '正常按月排班号源"
        Set GetPlanItemImage = imgList16.ListImages("MonthItem").Picture
    Case UCase("InvalidWeekItem") '无效按周排班号源"
        Set GetPlanItemImage = imgList16.ListImages("InvalidWeekItem").Picture
    Case UCase("WeekItem") '正常按周排班号源
        Set GetPlanItemImage = imgList16.ListImages("WeekItem").Picture
    End Select
    Exit Function
errHandler:
    Err.Clear
End Function

Public Function GetSortIcon(ByVal strKey As String) As IPictureDisp
    '获取排序图标
    '入参：
    '   strKey 图像索引
    On Error GoTo errHandler
    Select Case UCase(strKey)
    Case "ASC" '升序
        Set GetSortIcon = imgList16.ListImages("ASC").Picture
    Case "DESC" '降序
        Set GetSortIcon = imgList16.ListImages("DESC").Picture
    End Select
    Exit Function
errHandler:
    Err.Clear
End Function

Public Function GetLockPicture() As IPictureDisp
    '获取锁号图像
    On Error GoTo errHandler
    Set GetLockPicture = imgList16.ListImages("Lock").Picture
    Exit Function
errHandler:
    Err.Clear
End Function

Public Function GetTempPicture(ByVal strTxt As String, _
    Optional ByVal dblWidth As Double, Optional ByVal dblHight As Double, _
    Optional ByVal lngBackColor As OLE_COLOR = vbButtonFace, _
    Optional ByVal lngForeColor As OLE_COLOR = vbBlue, _
    Optional ByVal objFont As StdFont, _
    Optional ByVal intAlignment As PictureTextAlignmentSettings = pictxtAlignCenterCenter, _
    Optional ByVal strSubTxt As String, _
    Optional ByVal lngSubForeColor As OLE_COLOR = vbBlack, _
    Optional ByVal objSubFont As StdFont, _
    Optional ByVal intSubAlignment As PictureTextAlignmentSettings = pictxtAlignCenterCenter) As IPictureDisp
    '功能：根据参数生成图片
    '入参：
    '   strTxt - 主要显示文本
    '   dblWidth,dblHight - 图片大小，缺省为文本打印出来后的宽度和高度
    '   lngBackColor - 图片背景色，缺省为按钮表面颜色
    '   lngForeColor - 主要文本的前景色
    '   objFont - 主要文本的字体
    '   intAlignment - 主要文本的相对位置
    '   strSubTxt - 附加文本
    '   lngSubForeColor - 附加文本的前景色
    '   objSubFont - 附加文本的字体
    '   intSubAlignment - 附加文本的相对位置
    '返回：图片对象
    Dim p As POINT
    Dim objPic As IPictureDisp
    
    With picImage
        .AutoRedraw = True
        .Cls
        .BackColor = lngBackColor
        
        If objFont Is Nothing Then Set objFont = New StdFont
        Set .Font = objFont
        
        '确定图片大小
        If dblWidth = 0 Then dblWidth = .TextWidth(strTxt)
        If dblHight = 0 Then dblHight = .TextHeight(strTxt)
        .Width = dblWidth: .Height = dblHight
        
        '打印主要文本
        .ForeColor = lngForeColor
        p = GetTxtPostion(.Width, .Height, .TextWidth(strTxt), .TextHeight(strTxt), intAlignment)
        .CurrentX = p.X: .CurrentY = p.Y
        picImage.Print strTxt
        
        '打印附加文本
        If Trim(strSubTxt) <> "" Then
            If objSubFont Is Nothing Then Set objSubFont = New StdFont
            Set .Font = objSubFont
            .ForeColor = lngSubForeColor
            p = GetTxtPostion(.Width, .Height, .TextWidth(strSubTxt), .TextHeight(strSubTxt), intSubAlignment)
            .CurrentX = p.X: .CurrentY = p.Y
            picImage.Print strSubTxt
        End If
    End With
    
    '裁剪一下
    Set objPic = picImage.Image
    picImage.Cls
    picImage.PaintPicture objPic, 0, 0, dblWidth, dblHight, 0, 0, dblWidth, dblHight
    Set GetTempPicture = picImage.Image
End Function

Private Function GetTxtPostion(ByVal dblWidth As Double, ByVal dblHight As Double, _
    ByVal dblTxtWidth As Double, ByVal dblTxtHight As Double, _
    ByVal intAlignment As PictureTextAlignmentSettings) As POINT
    '确定文本打印位置
    Dim p As POINT
    
    Select Case intAlignment
    Case pictxtAlignLeftTop '左上
        p.X = 0
        p.Y = 0
    Case pictxtAlignLeftCenter
        p.X = 0
        p.Y = (dblHight - dblTxtHight) / 2
    Case pictxtAlignLeftBottom
        p.X = 0
        p.Y = dblHight - dblTxtHight
    Case pictxtAlignCenterTop
        p.X = (dblWidth - dblTxtWidth) / 2
        p.Y = 0
    Case pictxtAlignCenterCenter
        p.X = (dblWidth - dblTxtWidth) / 2
        p.Y = (dblHight - dblTxtHight) / 2
    Case pictxtAlignCenterBottom
        p.X = (dblWidth - dblTxtWidth) / 2
        p.Y = dblHight - dblTxtHight
    Case pictxtAlignRightTop
        p.X = dblWidth - dblTxtWidth
        p.Y = 0
    Case pictxtAlignRightCenter
        p.X = dblWidth - dblTxtWidth
        p.Y = (dblHight - dblTxtHight) / 2
    Case pictxtAlignRightBottom
        p.X = dblWidth - dblTxtWidth
        p.Y = dblHight - dblTxtHight
    End Select
    GetTxtPostion = p
End Function

Public Function GetVsfGrid(rptData As ReportControl, _
    Optional ByVal strHiddenCols As String) As VSFlexGrid
    '功能:将ReportControl转换为VSFlexGrid
    '入参:
    '   strHiddenCols 隐藏列索引(索引从0开始)，格式：列1,列2,列3,...
    Dim i As Long, j As Long, lngRowIndex As Long
    Dim varData As Variant
    
    Err = 0: On Error GoTo errHandler
    With vsfGridPrint
        .Clear
        .Cols = rptData.Columns.Count
        .Rows = rptData.Records.Count + 1
        .FixedAlignment(-1) = flexAlignCenterCenter
        
        '标题行
        For i = 0 To rptData.Columns.Count - 1
            .TextMatrix(0, i) = rptData.Columns(i).Caption
            .ColWidth(i) = rptData.Columns(i).Width * 16
            .ColAlignment(i) = Choose(rptData.Columns(i).Alignment + 1, 1, 4, 7)
        Next
        '隐藏列
        If strHiddenCols <> "" Then
            varData = Split(strHiddenCols, ",")
            For i = 0 To UBound(varData)
                .ColWidth(Val(varData(i))) = 0
            Next
        End If
        
        '数据行
        lngRowIndex = 1
        For i = 0 To rptData.Rows.Count - 1
            If rptData.Rows(i).GroupRow = False Then
                For j = 0 To rptData.Columns.Count - 1
                    .TextMatrix(lngRowIndex, j) = rptData.Rows(i).Record(j).Value
                Next
                lngRowIndex = lngRowIndex + 1
            End If
        Next
    End With
    Set GetVsfGrid = vsfGridPrint
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function VSFlexGridCopyTo(ByVal vsfSource As VSFlexGrid, ByRef vsfNew As VSFlexGrid, _
    Optional ByVal bytMode As Byte) As Boolean
    '功能: 将vsfSource的数据复制到vsfNew中，包括显示格式，用于打印\预览
    '参数:
    '     vsfNew-复制后的对象
    '     vsfSource-被复制的对象
    '     bytMode=1 打印;2 预览;3 输出到EXCEL
    '返回：复制成功，返回True；否则，返回False
    Dim i As Long, j As Long
    
    On Error GoTo errHandler
    Set vsfNew = vsfGridCopy
    With vsfNew
        .Redraw = flexRDNone
        .Clear
        '1.复制数据
        .LoadArray vsfSource
        
        '2.复制格式
        .FixedRows = vsfSource.FixedRows
        .FixedCols = vsfSource.FixedCols
        
        '合并属性
        .MergeCells = vsfSource.MergeCells
        .MergeCellsFixed = vsfSource.MergeCellsFixed
        .MergeCompare = vsfSource.MergeCompare
        
        For i = 0 To .Rows - 1
            .RowHeight(i) = vsfSource.RowHeight(i)
            .RowHidden(i) = vsfSource.RowHidden(i)
            If .RowHidden(i) Then .RowHeight(i) = 0
            If .RowHeight(i) = 0 Then .RowHidden(i) = True
            
            '合并属性
            .MergeRow(i) = vsfSource.MergeRow(i)
        Next
        
        '特殊处理：移除隐藏行，每间隔一个号源在文本后加一个空格(防止不同号源间行合并)，用以解决输出到Excel会显示隐藏行的问题
        '在加载数据时为了防止不同日期间列合并，是在文本两边各加了一个空格：strSpace & .TextMatrix(i, j) & strSpace
        If bytMode = 3 Then
            For i = 0 To .Rows - 1
                If i > .Rows - 1 Then Exit For
            
                Dim blnAddSpace As Boolean, strPrevNo As String
                If .RowHidden(i) Then
                    .RemoveItem i
                    i = i - 1
                Else
                    If strPrevNo <> .TextMatrix(i, COL_号码) Then
                        strPrevNo = .TextMatrix(i, COL_号码)
                        blnAddSpace = Not blnAddSpace
                    End If
                    If blnAddSpace Then
                        For j = 0 To .Cols - 1
                            .TextMatrix(i, j) = .TextMatrix(i, j) & " "
                        Next
                    End If
                End If
            Next
        End If

        For j = 0 To .Cols - 1
            .FixedAlignment(j) = vsfSource.FixedAlignment(j)
            .ColAlignment(j) = vsfSource.ColAlignment(j)
            .ColHidden(j) = vsfSource.ColHidden(j)
            .ColWidth(j) = vsfSource.ColWidth(j)
            If .ColHidden(j) Then .ColWidth(j) = 0
            
            '合并属性
            .MergeCol(j) = vsfSource.MergeCol(j)
        Next
        
        For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                .Cell(flexcpBackColor, i, j) = vsfSource.Cell(flexcpBackColor, i, j)
                .Cell(flexcpFont, i, j) = vsfSource.Cell(flexcpFont, i, j)
                .Cell(flexcpForeColor, i, j) = vsfSource.Cell(flexcpForeColor, i, j)
            Next
        Next
        
        .BackColor = vsfSource.BackColor
        .BackColorAlternate = vsfSource.BackColorAlternate
        .BackColorBkg = vsfSource.BackColorBkg
        .BackColorFixed = vsfSource.BackColorFixed
        .Redraw = flexRDBuffered
    End With
    VSFlexGridCopyTo = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ShowPlanColor(ByVal frmParent As Object)
    '功能:显示安排颜色标识
    On Error Resume Next
    With vsfPlanColor
        .Rows = 4
        .Cell(flexcpText, 0, 0, 0, 1) = "操作类型" & vbTab & "显示方式"
        .Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
        '1.停诊号用红色背景显示
        '2.替诊号用蓝色字体显示并显示替诊医生
        '3.临时出诊号用蓝色字体显示
        .Cell(flexcpText, 1, 0, 1, 1) = "停诊(含部分)" & vbTab & "上午"
        .Cell(flexcpBackColor, 1, 1) = vbRed
        
        .Cell(flexcpText, 2, 0, 2, 1) = "替诊(含部分)" & vbTab & "上午(张三)"
        .Cell(flexcpForeColor, 2, 1) = vbBlue
        
        .Cell(flexcpText, 3, 0, 3, 1) = "临时出诊" & vbTab & "上午"
        .Cell(flexcpForeColor, 3, 1) = vbBlue
        
        .Cell(flexcpAlignment, 1, 1, 3, 1) = flexAlignCenterCenter
        
        .Left = -10
        .Top = -10
        .Height = 300 * .Rows
        
        picPlanColor.Left = 0
        picPlanColor.Top = 0
        picPlanColor.Height = .Height
    End With
    
    With Me
        If Me.Top < 0 Or Me.Left < 0 Then
            Me.Top = 0: Me.Left = 0
        End If
        .Width = picPlanColor.Width
        .Height = picPlanColor.Height
        
        Dim objBar As Object, objPoint As RECT
        For Each objBar In frmParent
            If UCase(TypeName(objBar)) = "STATUSBAR" Then Exit For
        Next
        Call GetWindowRect(objBar.Hwnd, objPoint)
        
        Me.Top = objPoint.Top * Screen.TwipsPerPixelY - Me.Height
        Me.Left = objPoint.Left * Screen.TwipsPerPixelX + objBar.Panels("PlanColor").Left - (Me.Width - objBar.Panels("PlanColor").Width) / 2
        
        Me.Show 0, frmParent
        vsfPlanColor.SetFocus
    End With
End Function

Public Function ShowDoctorsTitle(ByVal frmParent As Object, ByVal rsDoctorsTitle As ADODB.Recordset) As Boolean
    '功能:显示医生的专业技术职务
    Dim i As Long

    On Error Resume Next
    With vsfPlanColor
        .Clear
        If rsDoctorsTitle.RecordCount = 0 Then Exit Function
        .Rows = rsDoctorsTitle.RecordCount + 1
        
        .TextMatrix(0, 0) = "职称"
        .TextMatrix(0, 1) = "标识符"
        .Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
        rsDoctorsTitle.MoveFirst
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = Nvl(rsDoctorsTitle!名称)
            .TextMatrix(i, 1) = Nvl(rsDoctorsTitle!标识符)
            .Cell(flexcpAlignment, i, 0) = flexAlignLeftCenter
            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
            rsDoctorsTitle.MoveNext
        Next
        
        .Left = -10
        .Top = -10
        .Height = 300 * .Rows
        
        picPlanColor.Left = 0
        picPlanColor.Top = 0
        picPlanColor.Height = .Height
    End With
    With Me
        If Me.Top < 0 Or Me.Left < 0 Then
            Me.Top = 0: Me.Left = 0
        End If
        .Width = picPlanColor.Width
        .Height = picPlanColor.Height

        Dim objBar As Object, objPoint As RECT
        For Each objBar In frmParent
            If UCase(TypeName(objBar)) = "STATUSBAR" Then Exit For
        Next
        Call GetWindowRect(objBar.Hwnd, objPoint)
        Me.Top = objPoint.Top * Screen.TwipsPerPixelY - Me.Height
        Me.Left = objPoint.Left * Screen.TwipsPerPixelX + objBar.Panels("DoctorsTitle").Left - (Me.Width - objBar.Panels("DoctorsTitle").Width) / 2

        Me.Show 0, frmParent
        vsfPlanColor.SetFocus
    End With
    ShowDoctorsTitle = True
End Function

Private Sub vsfPlanColor_LostFocus()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
