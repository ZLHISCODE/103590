VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicDefaultModus 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vfg方法 
      Height          =   2490
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4350
      _cx             =   7673
      _cy             =   4392
      Appearance      =   0
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      Begin MSComctlLib.ImageList imgList 
         Left            =   2985
         Top             =   1860
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicDefaultModus.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicDefaultModus.frx":059A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicDefaultModus.frx":0B34
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicDefaultModus.frx":10CE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmClinicDefaultModus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    方法 = 0
    造影 = 2
    默认 = 3
End Enum

Private mblnOK As Boolean
Private mstrModus As String
Private mstrDefault As String
'Private mlngRow As Long
'Private mlngCol As Long

Public Sub ShowModus(ByVal strModus As String, ByRef strDefault As String)
    mblnOK = False: mstrModus = strModus: mstrDefault = strDefault
    'mlngRow = lngRow: mlngCol = lngCol
    Me.Show vbModal
    If mblnOK And mstrDefault <> strDefault Then strDefault = mstrDefault
End Sub


Private Sub Form_Activate()
    Call FormatList(mstrModus)
End Sub

Private Sub Form_Deactivate()
    Dim strReturn As String, lngRow As Long
    Dim strTemp As String
    
    With vfg方法
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngRow, mCol.默认) = flexChecked Then
                    If .RowData(lngRow) = 3 Then
                        strReturn = strReturn & strTemp & "〈□" & .Cell(flexcpText, lngRow, mCol.方法 + 1) & ","
                    Else
                        strTemp = .Cell(flexcpText, lngRow, mCol.方法 + 1)
                        strReturn = strReturn & .Cell(flexcpText, lngRow, mCol.方法 + 1) & ","
                    End If
            End If
        Next
    End With
    If InStr(strReturn, ",") > 0 Then strReturn = Mid(strReturn, 1, Len(strReturn) - 1)
    If mstrDefault <> strReturn Then mstrDefault = strReturn
    mblnOK = True
    Me.Tag = ""
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim strReturn As String, lngRow As Long
    Dim strTemp As String
    ' 按ESC ,清空后返回
    If KeyAscii = 27 Then Unload Me
    ' 按回车,返回
    If KeyAscii = 13 Then
        With vfg方法
            For lngRow = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, lngRow, mCol.默认) = flexChecked Then
                        If .RowData(lngRow) = 3 Then
                            strReturn = strReturn & strTemp & "〈□" & .Cell(flexcpText, lngRow, mCol.方法 + 1) & ","
                        Else
                            strTemp = .Cell(flexcpText, lngRow, mCol.方法 + 1)
                            strReturn = strReturn & .Cell(flexcpText, lngRow, mCol.方法 + 1) & ","
                        End If
                End If
            Next
        End With
        If InStr(strReturn, ",") > 0 Then strReturn = Mid(strReturn, 1, Len(strReturn) - 1)
        If mstrDefault <> strReturn Then mstrDefault = strReturn
        mblnOK = True
        Unload Me
    End If
End Sub


Private Sub Form_Resize()
    vfg方法.Move 15, 15, ScaleWidth - 30, ScaleHeight - 30
End Sub

Private Sub FormatList(Optional strMode As String)
    '功能：初始化设置参考值列表
    '参数：strMode-方法串
    Dim aryItem() As String, strItems As String, strTemp As String
    Dim aryChild() As String, lngChild As Long, lngCount As Long
    
    With Me.vfg方法
        .Redraw = flexRDNone
        .Clear
        .Rows = 1: .FixedRows = 1: .Cols = 4: .FixedCols = 0
        .TextMatrix(0, mCol.方法) = "检查方法": .ColWidth(mCol.方法) = 280: .FixedAlignment(mCol.方法) = flexAlignCenterCenter
        .TextMatrix(0, mCol.方法 + 1) = "检查方法": .ColWidth(mCol.方法 + 1) = 2500
        .TextMatrix(0, mCol.造影) = "造影"
        .TextMatrix(0, mCol.默认) = "默认": .ColWidth(mCol.默认) = 0
        .MergeCells = flexMergeRestrictRows: .MergeRow(0) = True
        If strMode = "" Then .Redraw = flexRDDirect: Exit Sub
        
        '.Editable = flexEDKbdMouse
        
        strMode = Replace(strMode, vbTab, ";" & vbTab)
        aryItem() = Split(strMode, ";")
        mstrDefault = "," & mstrDefault & ","
        For lngCount = 0 To UBound(aryItem)
            strTemp = aryItem(lngCount)
            If strTemp <> "" Then
                If InStr(1, aryItem(lngCount), ",") > 0 Then strTemp = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), ",") - 1)
                .Rows = .Rows + 1 ': .MergeRow(.Rows - 1) = True
                If InStr(1, strTemp, vbTab) > 0 Then
                    .RowData(.Rows - 1) = 2
                Else
                    .RowData(.Rows - 1) = 1
                End If
                Set .Cell(flexcpPicture, .Rows - 1, mCol.方法) = Me.imgList.ListImages(.RowData(.Rows - 1)).Picture
                .TextMatrix(.Rows - 1, mCol.方法) = IIf(InStr(1, strTemp, vbTab) > 0, Mid(strTemp, 3), Mid(strTemp, 2))
                .TextMatrix(.Rows - 1, mCol.方法 + 1) = .TextMatrix(.Rows - 1, mCol.方法)
                If Val(Left(strTemp, 1)) = 1 Then
                    .Cell(flexcpText, .Rows - 1, mCol.造影) = "√"
                Else
                    .Cell(flexcpText, .Rows - 1, mCol.造影) = ""
                End If
                
                If InStr(1, strTemp, vbTab) > 0 Then
                    If InStr(mstrDefault & ",", "," & Mid(strTemp, 3) & ",") > 0 Then
                        .Cell(flexcpChecked, .Rows - 1, mCol.默认) = flexChecked
                        Set .Cell(flexcpPicture, .Rows - 1, mCol.方法) = Me.imgList.ListImages(.RowData(.Rows - 1)).Picture
                    Else
                        .Cell(flexcpChecked, .Rows - 1, mCol.默认) = flexUnchecked
                        Set .Cell(flexcpPicture, .Rows - 1, mCol.方法) = Me.imgList.ListImages(.RowData(.Rows - 1) + 2).Picture
                    End If
                Else
                    If InStr(mstrDefault & ",", "," & Mid(strTemp, 2) & ",") > 0 Then
                        .Cell(flexcpChecked, .Rows - 1, mCol.默认) = flexChecked
                        Set .Cell(flexcpPicture, .Rows - 1, mCol.方法) = Me.imgList.ListImages(.RowData(.Rows - 1)).Picture
                    Else
                        .Cell(flexcpChecked, .Rows - 1, mCol.默认) = flexUnchecked
                        Set .Cell(flexcpPicture, .Rows - 1, mCol.方法) = Me.imgList.ListImages(.RowData(.Rows - 1) + 2).Picture
                    End If
                End If
                
                If InStr(1, aryItem(lngCount), ",") > 0 Then
                    strTemp = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), ",") + 1)
                    aryChild = Split(strTemp, ",")
                    For lngChild = 0 To UBound(aryChild)
                        strTemp = aryChild(lngChild)
                        .Rows = .Rows + 1 ': .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 3
                        Set .Cell(flexcpPicture, .Rows - 1, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(.Rows - 1) - 1).Picture
                        .TextMatrix(.Rows - 1, mCol.方法 + 1) = Mid(strTemp, 2)
                        If Val(Left(strTemp, 1)) = 1 Then
                            .Cell(flexcpText, .Rows - 1, mCol.造影) = "√"
                        Else
                            .Cell(flexcpText, .Rows - 1, mCol.造影) = ""
                        End If
                        
                        strItems = Replace(Mid(aryItem(lngCount), 1, InStr(aryItem(lngCount), ",") - 1), vbTab, "")
                        strItems = Mid(strItems, 2)
                        If InStr("," & mstrDefault & ",", "," & strItems & "〈□" & Mid(strTemp, 2) & ",") > 0 Then
                            .Cell(flexcpChecked, .Rows - 1, mCol.默认) = flexChecked
                            Set .Cell(flexcpPicture, .Rows - 1, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(.Rows - 1) - 1).Picture

                        Else
                            .Cell(flexcpChecked, .Rows - 1, mCol.默认) = flexUnchecked
                            Set .Cell(flexcpPicture, .Rows - 1, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(.Rows - 1) + 1).Picture
                        End If
                    Next
                End If
            End If
        Next
        If .Rows > .FixedRows Then .Row = .FixedRows
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vfg方法_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.默认 Then
        Cancel = True
    End If
End Sub

Private Sub vfg方法_DblClick()
    Call zlCommFun.PressKey(13)
End Sub

Private Sub vfg方法_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeySpace Then Exit Sub
    
    Call EditDefault(vfg方法.Row, vfg方法.Col)
End Sub

Private Sub EditDefault(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngTmp As Long, blnNoChecked As Boolean

    With vfg方法
        .Cell(flexcpChecked, .Row, mCol.默认) = IIf(.Cell(flexcpChecked, .Row, mCol.默认) = flexChecked, flexUnchecked, flexChecked)
        If .RowData(Row) <> 3 And .Cell(flexcpChecked, .Row, mCol.默认) = flexChecked Then
            '当修改为选中状态时
            If .Cell(flexcpPicture, .Row, mCol.方法) Is Nothing Then
                .Cell(flexcpPicture, .Row, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(.Row)).Picture
            Else
                .Cell(flexcpPicture, .Row, mCol.方法) = Me.imgList.ListImages(.RowData(.Row)).Picture
            End If
            For lngRow = Row To .FixedRows Step -1
                '去掉当前行之上的其他排斥项的选中状态
                If .RowData(Row) = 1 And lngRow <> Row And .RowData(lngRow) = 1 Then
                    .Cell(flexcpChecked, lngRow, mCol.默认) = flexUnchecked
                    
                    If .Cell(flexcpPicture, lngRow, mCol.方法) Is Nothing Then
                        .Cell(flexcpPicture, lngRow, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(lngRow) + 2).Picture
                    Else
                        .Cell(flexcpPicture, lngRow, mCol.方法) = Me.imgList.ListImages(.RowData(lngRow) + 2).Picture
                    End If
                    
                    For lngTmp = lngRow To Row
                    '取掉排斥项下的附加项的选中状态
                        If lngTmp <> Row And .Cell(flexcpPicture, lngTmp, mCol.方法) Is Nothing Then
                            .Cell(flexcpChecked, lngTmp, mCol.默认) = flexUnchecked
                            .Cell(flexcpPicture, lngTmp, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(lngTmp) + 1).Picture
                        ElseIf lngTmp <> lngRow And .RowData(lngTmp) <> 0 Then
                            Exit For
                        End If
                    Next
                End If
            Next
            
            For lngRow = Row To .Rows - 1
                '去掉当前行之下的其他排斥项的选中状态
                If lngRow <> Row And .RowData(lngRow) = 1 And .RowData(Row) = 1 Then
                    .Cell(flexcpChecked, lngRow, mCol.默认) = flexUnchecked
                    
                    If .Cell(flexcpPicture, lngRow, mCol.方法) Is Nothing Then
                        .Cell(flexcpPicture, lngRow, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(lngRow) + 2).Picture
                    Else
                        .Cell(flexcpPicture, lngRow, mCol.方法) = Me.imgList.ListImages(.RowData(lngRow) + 2).Picture
                    End If
                    For lngTmp = lngRow To .Rows - 1
                        '取掉排斥项下的附加项的选中状态
                        If lngTmp <> lngRow And .Cell(flexcpPicture, lngTmp, mCol.方法) Is Nothing Then
                            .Cell(flexcpChecked, lngTmp, mCol.默认) = flexUnchecked
                            .Cell(flexcpPicture, lngTmp, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(lngTmp) + 1).Picture
                        ElseIf lngTmp <> lngRow And .RowData(lngTmp) <> 0 Then
                            Exit For
                        End If
                    Next
                End If
            Next
            
             
       ElseIf .RowData(Row) <> 3 And .Cell(flexcpChecked, Row, mCol.默认) = flexUnchecked Then
'                '当修改为不选中状态时
'
            If .Cell(flexcpPicture, .Row, mCol.方法) Is Nothing Then
                .Cell(flexcpPicture, .Row, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(.Row) + 2).Picture
            Else
                .Cell(flexcpPicture, .Row, mCol.方法) = Me.imgList.ListImages(.RowData(.Row) + 2).Picture
            End If
            
            For lngRow = Row To .Rows - 1
                '去掉当前行之下的其他排斥项的选中状态
                If (.RowData(lngRow) = 1 Or .RowData(lngRow) = 2) And lngRow <> Row Then
                    Exit For
                End If
                If lngRow <> Row And .RowData(lngRow) = 3 Then
                    For lngTmp = lngRow To .Rows - 1
                        '取掉排斥项下的附加项的选中状态
                        If lngTmp >= lngRow And .Cell(flexcpPicture, lngTmp, mCol.方法) Is Nothing Then
                            .Cell(flexcpChecked, lngTmp, mCol.默认) = flexUnchecked
                            
                            .Cell(flexcpPicture, lngTmp, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(lngTmp) + 1).Picture

                        ElseIf lngTmp <> lngRow And .RowData(lngTmp) <> 0 Then
                            Exit For
                        End If
                    Next
                End If
            Next
        Else
            If .Cell(flexcpChecked, Row, mCol.默认) = flexUnchecked Then
                If .Cell(flexcpPicture, Row, mCol.方法) Is Nothing Then
                    .Cell(flexcpPicture, Row, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(Row) + 1).Picture
                Else
                    .Cell(flexcpPicture, Row, mCol.方法) = Me.imgList.ListImages(.RowData(Row) + 1).Picture
                End If
            Else
            
                If .Cell(flexcpPicture, Row, mCol.方法) Is Nothing Then
                    '检查父项是事选中
                    For lngRow = Row To .FixedRows Step -1
                        If lngRow <> Row And .RowData(lngRow) <> 3 Then
                            Exit For
                        End If
                    Next
                    If lngRow >= .FixedRows And lngRow < Row Then
                        If .Cell(flexcpChecked, lngRow, mCol.默认) = flexUnchecked Then
                            .Cell(flexcpPicture, Row, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(Row) + 1).Picture
                            .Cell(flexcpChecked, Row, mCol.默认) = flexUnchecked
                        Else
                            .Cell(flexcpPicture, Row, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(Row) - 1).Picture
                        End If
                    End If
                    
                Else
                    .Cell(flexcpPicture, Row, mCol.方法) = Me.imgList.ListImages(.RowData(Row)).Picture
                End If
                
            End If
            
        End If
    End With


End Sub

Private Sub vfg方法_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 1 And x > vfg方法.CellLeft And x < vfg方法.CellLeft + 250 Then
        Call EditDefault(vfg方法.Row, vfg方法.Col)
    End If
    
End Sub
