VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCompendWord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关联词句示范"
   ClientHeight    =   6540
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6165
   Icon            =   "frmCompendWord.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkChildren 
      Caption         =   "同步清除下级分类(&2)"
      Height          =   195
      Index           =   1
      Left            =   2505
      TabIndex        =   5
      Top             =   5730
      Width           =   2055
   End
   Begin VB.CheckBox chkChildren 
      Caption         =   "同步选中下级分类(&1)"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   5730
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox chkOnlySel 
      Caption         =   "仅显示已选择分类(&S)"
      Height          =   195
      Left            =   165
      TabIndex        =   6
      Top             =   6135
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3750
      TabIndex        =   2
      Top             =   6060
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4890
      TabIndex        =   3
      Top             =   6060
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   5265
      Left            =   150
      TabIndex        =   1
      Top             =   390
      Width           =   5835
      _cx             =   10292
      _cy             =   9287
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
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
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   3
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
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmCompendWord.frx":000C
      Top             =   75
      Width           =   240
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "设置在实际编辑过程中，当前提纲关联可选的词句示范分类。"
      Height          =   180
      Left            =   435
      TabIndex        =   0
      Top             =   105
      Width           =   4860
   End
End
Attribute VB_Name = "frmCompendWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    选择 = 0: ID: 上级id: 编码: 名称: 说明
End Enum

Private mlngCompendId As Long   '当前提纲ID
Private mblnOK As Boolean       '是否确认

Dim lngCount As Long

'-----------------------------------------------------
'以下为外部公共程序
'-----------------------------------------------------
Public Function ShowMe(ByVal frmParent As Form, ByVal lngCompendID As Long, bytFileType As Byte) As Boolean
    '功能：显示本编辑窗体
    '参数： frmParent-父窗体
    '       lngCompendId-提纲ID
    '       bytFileType-文件类型
    Dim rsTemp As New ADODB.Recordset
    mlngCompendId = lngCompendID
    
    '装入可选分类数据
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select Decode(U.词句分类id, Null, 0, 1) As 选择, C.ID, C.上级id, C.编码, C.名称, C.说明" & vbNewLine & _
            "From 病历词句分类 C, (Select 词句分类id From 病历提纲词句 Where 提纲id = [1]) U" & vbNewLine & _
            "Where C.ID = U.词句分类id(+) And Substr(范围, [2], 1) = '1'" & vbNewLine & _
            "Order By C.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngCompendID, bytFileType)
    With Me.vfgList
        .Redraw = flexRDNone
        Set .DataSource = rsTemp
        .ColWidth(mCol.选择) = 280
        .ColWidth(mCol.ID) = 0: .ColHidden(mCol.ID) = True
        .ColWidth(mCol.上级id) = 0: .ColHidden(mCol.上级id) = True
        For lngCount = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngCount, mCol.选择)) = 1 Then
                .Cell(flexcpChecked, lngCount, mCol.选择) = flexChecked
            Else
                .Cell(flexcpChecked, lngCount, mCol.选择) = flexUnchecked
            End If
            .TextMatrix(lngCount, mCol.选择) = ""
        Next
        If .Rows > .FixedRows Then .Row = .FixedRows
        .Col = mCol.选择
        .Redraw = flexRDDirect
    End With
    
    Me.Show vbModal, frmParent
    ShowMe = mblnOK
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub chkChildren_Click(Index As Integer)
    With Me.vfgList
        If .Visible And .Enabled Then .SetFocus
    End With
End Sub

Private Sub chkOnlySel_Click()
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If Me.chkOnlySel.Value = vbChecked Then
                If .Cell(flexcpChecked, lngCount, mCol.选择) = flexUnchecked Then
                    .RowHidden(lngCount) = True
                End If
            Else
                .RowHidden(lngCount) = False
            End If
        Next
        If .Visible And .Enabled Then .SetFocus
    End With
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strClass As String
    
    Err = 0: On Error GoTo errHand
    strClass = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.选择) = flexChecked Then
                strClass = strClass & "," & .TextMatrix(lngCount, mCol.ID)
            End If
        Next
    End With
    If strClass <> "" Then strClass = Mid(strClass, 2)
    
    gstrSQL = "Zl_病历提纲词句_Update(" & mlngCompendId & ",'" & strClass & "')"
    zlDatabase.ExecuteProcedure gstrSQL, "保存词句关联"
    mblnOK = True: Unload Me
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgList_DblClick()
    With Me.vfgList
        If .Row < .FixedRows Then Exit Sub
        If .Cell(flexcpChecked, .Row, mCol.选择) = flexUnchecked Then
            .Cell(flexcpChecked, .Row, mCol.选择) = flexChecked
            If Me.chkChildren(0).Value = vbChecked Then
                For lngCount = .Row To .Rows - 1
                    If Val(.TextMatrix(lngCount, mCol.上级id)) = Val(.TextMatrix(.Row, mCol.ID)) Then
                        .Cell(flexcpChecked, lngCount, mCol.选择) = flexChecked
                    End If
                Next
            End If
        Else
            .Cell(flexcpChecked, .Row, mCol.选择) = flexUnchecked
            If Me.chkChildren(1).Value = vbChecked Then
                For lngCount = .Row To .Rows - 1
                    If Val(.TextMatrix(lngCount, mCol.上级id)) = Val(.TextMatrix(.Row, mCol.ID)) Then
                        .Cell(flexcpChecked, lngCount, mCol.选择) = flexUnchecked
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii = vbKeySpace Then Call vfgList_DblClick
End Sub
