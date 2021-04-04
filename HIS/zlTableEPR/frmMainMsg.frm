VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMainMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "内容监控提醒"
   ClientHeight    =   4305
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7440
   Icon            =   "frmMainMsg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdLocation 
      Caption         =   "定位(&L)"
      Height          =   350
      Left            =   6000
      TabIndex        =   3
      Top             =   1545
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdTerm 
      Cancel          =   -1  'True
      Caption         =   "终止(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   2
      Top             =   870
      Width           =   1100
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "继续(&O)"
      Height          =   350
      Left            =   6000
      TabIndex        =   0
      Top             =   450
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   3780
      Left            =   180
      TabIndex        =   1
      Top             =   435
      Width           =   5565
      _cx             =   9816
      _cy             =   6667
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
      AutoSizeMode    =   1
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "注意：本文件中存在如下内容为空的诊治要素，是否继续保存操作"
      Height          =   180
      Left            =   165
      TabIndex        =   4
      Top             =   150
      Width           =   5220
   End
End
Attribute VB_Name = "frmMainMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mfrmMain As Object

Public Event Location(ByVal strRange As String, ByVal elKey As Long)

Public Function ShowNotice(ByVal frmMain As Object, Optional CheckItemMust As Boolean = True) As Boolean
'CheckItemMust：强制检查必填要素
    Dim intCount As Integer, intLoop As Integer, elKey As String, le As Long, lKey As String
    mblnOK = False
    Set mfrmMain = frmMain

    With vsf
        .Rows = 2
        .Cols = 4
        .FixedCols = 0
        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 1) = "要素名称"
        .TextMatrix(0, 2) = "必填要素"
        .TextMatrix(0, 3) = "单元Key"
        .ColWidth(0) = 800
        .ColWidth(1) = 3600
        .ColWidth(2) = 800
        .ColWidth(3) = 0
        
        .FixedAlignment(0) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .FixedAlignment(2) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ExtendLastCol = True
        
        Dim lRow As Long, lCol As Long
        For intLoop = 1 To mfrmMain.Document.Cells.Count
            elKey = mfrmMain.Document.Cells(intLoop).ElementKey
            If elKey <> "" Then
                For le = 0 To UBound(Split(elKey, "|"))
                    lKey = Split(elKey, "|")(le)
                    If lKey <> "" Then
                        lRow = Split(mfrmMain.Document.Elements("K" & lKey).区域, "|")(0): lCol = Split(mfrmMain.Document.Elements("K" & lKey).区域, "|")(1)
                        If Trim(mfrmMain.Document.Elements("K" & lKey).内容文本) = "" Or _
                            (mfrmMain.Document.Elements("K" & lKey).输入形态 = 1 And mfrmMain.Document.Elements("K" & lKey).要素表示 = 2 And InStr(mfrmMain.Document.Elements("K" & lKey).内容文本, "●") = 0) Or _
                            (mfrmMain.Document.Elements("K" & lKey).输入形态 = 1 And mfrmMain.Document.Elements("K" & lKey).要素表示 = 3 And InStr(mfrmMain.Document.Elements("K" & lKey).内容文本, "■") = 0) Then
                            Select Case mfrmMain.Document.Elements("K" & lKey).要素名称
                            Case "经治医师签名", "主治医师签名", "主任医师签名"
                            Case Else
                                intCount = intCount + 1
                                If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                                .RowData(.Rows - 1) = mfrmMain.Document.Elements("K" & lKey).Key
                                .TextMatrix(.Rows - 1, 0) = intCount
                                .TextMatrix(.Rows - 1, 1) = mfrmMain.Document.Elements("K" & lKey).要素名称
                                .TextMatrix(.Rows - 1, 2) = IIf(mfrmMain.Document.Elements("K" & lKey).必填 = 0, "否", "是")
                                .TextMatrix(.Rows - 1, 3) = mfrmMain.Document.Elements("K" & lKey).区域
                                
                                If mfrmMain.Document.Elements("K" & lKey).必填 = 1 And CheckItemMust Then cmdContinue.Enabled = False '只要有必填项目未填则不允许通过
                            End Select
                        End If
                    End If
                Next
            End If
        Next
        
        If intCount = 0 Then
            ShowNotice = True
            Exit Function
        End If
    End With
    
    Me.Show 1, frmMain
    
    ShowNotice = mblnOK
    
End Function

Private Sub cmdContinue_Click()
    Unload Me
    mblnOK = True
End Sub

Private Sub cmdLocation_Click()
    RaiseEvent Location(vsf.TextMatrix(vsf.Row, 3), Val(vsf.RowData(vsf.Row)))
End Sub

Private Sub cmdTerm_Click()
    Unload Me
    mblnOK = False
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vsf_DblClick()
    Call cmdLocation_Click
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdLocation_Click
End Sub
