VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmOpsStationList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2145
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   3990
      _cx             =   7038
      _cy             =   3784
      Appearance      =   3
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
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
Attribute VB_Name = "frmOpsStationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum mCol
    图标
    紧急
    状态
    来源
    姓名
    标识号
    病人科室
    申请时间
    床号
    病人id
    医嘱id
End Enum

Public Property Get VsfGrid() As VSFlexGrid
    Set VsfGrid = vsf
End Property

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '==================================================================================================================
    '功能：
    '参数：
    '返回：
    '==================================================================================================================
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strStart As String
    Dim strEnd As String
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "初始数据"
    
            strTmp = ",255,4,1,1,[图标];,255,4,1,1,[紧急];,255,4,1,1,[状态];,255,4,1,1,[来源];手术名称,1500,1,1,1,;姓名,810,1,1,1,;标识号,1080,1,1,1,;病人科室,1200,1,1,1,;申请时间,1670,1,1,1,;床号,600,1,1,1,;病人id,0,1,1,0,;医嘱id,0,1,1,0,"
            
            Call CreateVsf(vsf, strTmp)
            With vsf
                .Cols = .Cols + 1
                .ColWidth(.Cols - 1) = 15
                .ExtendLastCol = True
                Set .Cell(flexcpPicture, 0, 0) = frmPubIcons.ils13.ListImages("图标").Picture
                Set .Cell(flexcpPicture, 0, 1) = frmPubIcons.ils13.ListImages("状态").Picture
                Set .Cell(flexcpPicture, 0, 2) = frmPubIcons.ils13.ListImages("状态").Picture
                Set .Cell(flexcpPicture, 0, 3) = frmPubIcons.ils13.ListImages("状态").Picture
            End With
        
        End Select
    Next
    
    ExecuteCommand = True
    
    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Sub Form_Load()
    Call ExecuteCommand("初始数据")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsf.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Call AppendRows(Me, vsf)
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(Me, vsf)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(Me, vsf)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col < 4)
End Sub

