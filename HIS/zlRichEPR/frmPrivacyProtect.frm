VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPrivacyProtect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人隐私保护项目设置"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   3855
   Icon            =   "frmPrivacyProtect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   4110
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   3840
      _cx             =   6773
      _cy             =   7250
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   260
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   3855
      TabIndex        =   3
      Top             =   0
      Width           =   3855
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请勾选需要保护的病人隐私项目。"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   165
         TabIndex        =   4
         Top             =   75
         Width           =   2700
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1485
      TabIndex        =   1
      Top             =   4545
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2655
      TabIndex        =   2
      Top             =   4545
      Width           =   1095
   End
End
Attribute VB_Name = "frmPrivacyProtect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    ID = 0: 保护: 编码: 名称: 单位
End Enum
Private mblnOk As Boolean, mstrPrivs As String

Public Function ShowMe(ByRef frmParent As Object, Optional lngModul As Long = 1074) As Boolean
    mstrPrivs = GetPrivFunc(glngSys, lngModul)
    If InStr(1, mstrPrivs, "隐私设置") = 0 Then
        MsgBox "对不起，您没有隐私设置权限！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    Call InitGrid
    Me.Show vbModal, frmParent
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim blnTran As Boolean
    '保存设置
    On Error GoTo LL
    gcnOracle.BeginTrans
    blnTran = True
    Dim i As Long
    gstrSQL = "Zl_隐私保护项目_Clear"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    For i = 1 To vfgThis.Rows - 1
        If vfgThis.Cell(flexcpChecked, i, mCol.保护) = flexChecked Then
            gstrSQL = "Zl_隐私保护项目_Insert(" & Val(vfgThis.TextMatrix(i, mCol.ID)) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    mblnOk = True
    Unload Me
    Exit Sub
LL:
    If blnTran Then gcnOracle.RollbackTrans
    
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnOk = False
End Sub

Private Sub InitGrid()
    Dim rsTemp As New ADODB.Recordset, i As Long
    '执行监测
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select A.ID, Decode(C.项目id, Null, 0, 1) As 保护, A.编码, A.中文名 As 名称, A.单位, B.名称 As 分类 " & _
        " From 诊治所见项目 A, 诊治所见分类 B, 隐私保护项目 C " & _
        " Where A.分类id = B.ID And (B.名称 = '病人基本信息' Or B.名称 = '病人附加信息') And A.ID = C.项目id(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    With Me.vfgThis
        .Clear
        .Editable = flexEDKbdMouse
        Set .DataSource = rsTemp
        .ColWidth(mCol.ID) = 0
        .ColWidth(mCol.保护) = 300
        For i = 0 To .Rows - 1
            .Cell(flexcpChecked, i, mCol.保护) = IIf(Val(.Cell(flexcpText, i, mCol.保护)) = 1, flexChecked, flexUnchecked)
            .Cell(flexcpText, i, mCol.保护) = ""
        Next
    End With
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgThis_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    If Row = 0 And Col = mCol.保护 Then
        For i = 1 To vfgThis.Rows - 1
            vfgThis.Cell(flexcpChecked, i, mCol.保护) = vfgThis.Cell(flexcpChecked, 0, mCol.保护)
        Next
    Else
        vfgThis.Cell(flexcpChecked, 0, mCol.保护) = flexUnchecked
    End If
End Sub

Private Sub vfgThis_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 1 Then Cancel = True
End Sub

Private Sub vfgThis_RowColChange()
    vfgThis.Col = mCol.保护
End Sub
