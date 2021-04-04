VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmOpsStationRoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "手术执行间"
   ClientHeight    =   3555
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4920
   Icon            =   "frmOpsStationRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   3570
      TabIndex        =   2
      Top             =   -1380
      Width           =   30
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3690
      TabIndex        =   1
      Top             =   495
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3690
      TabIndex        =   0
      Top             =   75
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   3345
      Left            =   75
      TabIndex        =   4
      Top             =   60
      Width           =   3405
      _cx             =   6006
      _cy             =   5900
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483638
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   765
      TabIndex        =   3
      Top             =   135
      Width           =   90
   End
End
Attribute VB_Name = "frmOpsStationRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'（１）窗体级变量定义

Private mblnOK As Boolean
Private mfrmParent As Form
Private mlngDept As Long
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Form, ByVal lngDept As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    mblnOK = False
    mlngDept = lngDept
    
    Set mfrmParent = frmMain
    
    Call ExecuteCommand("初始数据")
    Call ExecuteCommand("读取手术间")
    
    DataChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf, True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
            Call .AppendColumn("原执行间", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("执行间", 900, flexAlignLeftCenter, flexDTString, "", , True)
            

            Call .InitializeEdit(True, True, True)
            Call .InitializeEditColumn(.ColIndex("执行间"), True, vbVsfEditText)
            
            .IndicatorCol = 0
            Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            
            .AppendRows = True
        End With

        
        gstrSQL = "SELECT 名称 FROM 部门表 where ID=[1] "
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDept)
        If rs.BOF = False Then Me.Caption = "手术间设置：" & zlCommFun.NVL(rs("名称").Value)
    
    '--------------------------------------------------------------------------------------------------------------
    Case "读取手术间"
        
        mblnReading = True
        
        Call mclsVsf.ClearGrid
        
        gstrSQL = "SELECT 执行间,执行间 As 原执行间 FROM 医技执行房间 where 科室id=[1] "
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDept)
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)
        
        mblnReading = False
        DataChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case "校验数据"
        With vsf
            For intLoop = 1 To .Rows - 1
                For intRow = 1 To intLoop - 1
                    If Trim(.TextMatrix(intLoop, .ColIndex("执行间"))) = Trim(.TextMatrix(intRow, .ColIndex("执行间"))) Then
                        ShowSimpleMsg "“" & Trim(.TextMatrix(intLoop, .ColIndex("执行间"))) & "”已经存在！"
                        Exit Function
                    End If
                Next
            Next
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "保存数据"
        
        With vsf
            For intLoop = 1 To .Rows - 1
            
                If Trim(.TextMatrix(intLoop, .ColIndex("原执行间"))) <> "" Then
                    If Trim(.TextMatrix(intLoop, .ColIndex("执行间"))) = "" Then
                        strSQL = "zl_医技执行房间_Delete(" & mlngDept & ",'" & Trim(.TextMatrix(intLoop, .ColIndex("原执行间"))) & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                    ElseIf Trim(.TextMatrix(intLoop, .ColIndex("执行间"))) <> Trim(.TextMatrix(intLoop, .ColIndex("原执行间"))) Then
                        strSQL = "zl_医技执行房间_Update(" & mlngDept & ",'" & Trim(.TextMatrix(intLoop, .ColIndex("原执行间"))) & "','" & Trim(.TextMatrix(intLoop, .ColIndex("执行间"))) & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                    End If
                ElseIf Trim(.TextMatrix(intLoop, .ColIndex("执行间"))) <> "" Then
                    strSQL = "zl_医技执行房间_Insert(" & mlngDept & ",'" & Trim(.TextMatrix(intLoop, .ColIndex("执行间"))) & "')"
                    Call SQLRecordAdd(rsSQL, strSQL)
                End If
            Next
        End With
        
        ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        
        Exit Function
    End Select

    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    If ExecuteCommand("校验数据") = False Then Exit Sub
    If ExecuteCommand("保存数据") = False Then Exit Sub
    
    mblnOK = True
    
    DataChanged = False

    Unload Me

End Sub

Private Sub mclsVsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsSQL As ADODB.Recordset
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    If Trim(vsf.TextMatrix(Row, vsf.ColIndex("执行间"))) <> "" Then
        gstrSQL = "zl_医技执行房间_Delete(" & mlngDept & "," & Trim(vsf.TextMatrix(Row, vsf.ColIndex("执行间"))) & ")"
        Call SQLRecordAdd(rsSQL, gstrSQL)
        
    End If
    
    Cancel = Not SQLRecordExecute(rsSQL, Me.Caption, False)
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Cancel = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Trim(vsf.TextMatrix(Row, vsf.ColIndex("执行间"))) = "")
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
    DataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub


Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf.MouseRow, vsf.MouseCol)
    End Select
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub

