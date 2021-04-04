VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmReasonSelector 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "理由选择器"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   5520
      TabIndex        =   5
      Top             =   2760
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&O)"
      Height          =   360
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   990
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   360
      Left            =   3360
      TabIndex        =   3
      Top             =   2760
      Width           =   990
   End
   Begin VB.CommandButton cmdUnselect 
      Caption         =   "全取(&U)"
      Height          =   360
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   990
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全选(&A)"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   990
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfReason 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _cx             =   11245
      _cy             =   4048
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
Attribute VB_Name = "frmReasonSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_VSF As String = "选择,,3,600|内容,,3,2000"

Private mrsSelect As New ADODB.Recordset
Private mblnOK As Boolean

Public Function ShowMe(ByVal frmOwner As Form, ByVal strSQL As String, ByRef blnCancel As Boolean, ByVal arrInput As Variant) As ADODB.Recordset
    Dim l As Long
    
    mblnOK = True
    Set mrsSelect = Nothing
    
    Call InitVSF
    Call mdlDefine.SetVSFHead(vsfReason, MSTR_VSF)
    vsfReason.ColDataType(vsfReason.ColIndex("选择")) = flexDTBoolean
    
    If ShowReason(strSQL, arrInput) = False Then
        blnCancel = True
        Exit Function
    End If
    
    '自动行高
    vsfReason.AutoSize 0, vsfReason.Cols - 1
    
    RestoreWinState Me, App.ProductName
    
    Show vbModal, frmOwner
    blnCancel = mblnOK
    Set ShowMe = mrsSelect
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If vsfReason.Rows <= 1 Then Exit Sub
    
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = zlStr.FormatString("Zl_处方审查常用理由_Update(0, '[1]', '[2]')", _
                    UserInfo.用户名, _
                    vsfReason.TextMatrix(vsfReason.Row, vsfReason.ColIndex("内容")))
    Call zlDatabase.ExecuteProcedure(strSQL, "删除处方审查常用理由")
    
    vsfReason.RemoveItem vsfReason.Row
    
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Sub cmdOK_Click()
    mblnOK = Not SetRecordset()
    Unload Me
End Sub

Private Function SetRecordset() As Boolean
    Dim i As Integer
    Dim l As Long
    
    On Error GoTo errHandle
    
    With mrsSelect
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockPessimistic
        
        For i = 0 To vsfReason.Cols - 1
            .Fields.Append vsfReason.ColKey(i), adVariant
        Next
        .Open
        
        For l = 1 To vsfReason.Rows - 1
            If Val(vsfReason.TextMatrix(l, vsfReason.ColIndex("选择"))) = -1 Then
                .AddNew
                For i = 0 To vsfReason.Cols - 1
                    .Fields(i).Value = vsfReason.TextMatrix(l, i)
                Next
                .Update
            End If
        Next
    End With
    
    SetRecordset = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Function ShowReason(ByVal strSQL As String, ByVal arrInput As Variant) As Boolean
    Dim rsReason As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set rsReason = zlDatabase.OpenSQLRecord(strSQL, "查询审方常用理由", arrInput)
    Call mdlDefine.FillVSFData(vsfReason, rsReason)
    If vsfReason.Rows > 1 Then vsfReason.Row = 1
    
    ShowReason = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub InitVSF()
'功能：初始化窗体的VSFlexGrid控件的风格
'参数：
'  vsfVar：要初始化的VSFlexGrid控件

    With vsfReason
        .Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .SheetBorder = .BackColor
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoResize = True
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub cmdSelect_Click()
    Dim l As Long
    
    With vsfReason
        .Redraw = False
        For l = 1 To .Rows - 1
            .TextMatrix(l, .ColIndex("选择")) = "-1"
        Next
        .Redraw = True
    End With

End Sub

Private Sub cmdUnselect_Click()
    Dim l As Long
    
    With vsfReason
        .Redraw = False
        For l = 1 To .Rows - 1
            .TextMatrix(l, .ColIndex("选择")) = ""
        Next
        .Redraw = True
    End With
End Sub

Private Sub Form_Resize()
    Const INT_BOUND As Integer = 60

    On Error Resume Next
    
    With vsfReason
        .Left = INT_BOUND
        .Top = INT_BOUND
        .Width = Me.ScaleWidth - INT_BOUND * 2
        .Height = Me.ScaleHeight - INT_BOUND * 3 - cmdClose.Height
    End With
    
    With cmdSelect
        .Top = vsfReason.Top + vsfReason.Height + INT_BOUND
        .Left = INT_BOUND
    End With
    
    With cmdUnselect
        .Top = cmdSelect.Top
        .Left = cmdSelect.Left + cmdSelect.Width + INT_BOUND
    End With
    
    With cmdClose
        .Top = cmdSelect.Top
        .Left = Me.ScaleWidth - INT_BOUND - .Width
    End With
    
    With cmdOK
        .Top = cmdSelect.Top
        .Left = cmdClose.Left - INT_BOUND - .Width
    End With
    
    With cmdDelete
        .Top = cmdSelect.Top
        .Left = cmdOK.Left - INT_BOUND - .Width
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub vsfReason_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfReason.ColIndex("选择") Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

