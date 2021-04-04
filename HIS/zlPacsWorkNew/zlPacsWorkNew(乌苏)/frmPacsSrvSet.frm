VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPacsSrvSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   Icon            =   "frmPacsSrvSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Caption         =   "退出(&C)"
      Height          =   350
      Left            =   8640
      TabIndex        =   5
      Top             =   8550
      Width           =   1125
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6360
      TabIndex        =   4
      Top             =   8550
      Width           =   1125
   End
   Begin VB.Frame FraList 
      Caption         =   "服务列表"
      Height          =   1860
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "列表来自于<影像设备目录>中的设置"
      Top             =   0
      Width           =   11505
      Begin VB.CommandButton CmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   10200
         TabIndex        =   3
         Top             =   645
         Width           =   1100
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   10200
         TabIndex        =   2
         Top             =   210
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   1560
         Left            =   120
         TabIndex        =   6
         Top             =   210
         Width           =   9990
         _cx             =   17621
         _cy             =   2752
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
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
         Rows            =   3
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
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
   Begin XtremeSuiteControls.TabControl TabList 
      Height          =   7080
      Left            =   0
      TabIndex        =   0
      Top             =   1905
      Width           =   11535
      _Version        =   589884
      _ExtentX        =   20346
      _ExtentY        =   12488
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmPacsSrvSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum ColList
        ColID = 0
        Col服务名
        Col服务功能
        ColPACS角色
        Col服务IP
        Col服务AE
        Col服务端口
        Col设备IP
        Col设备AE
        Col设备端口
End Enum

Private mFrmImgSrv As New frmImgSrv
Private mfrmWorkList As New frmWorklist
Private mFrmQRSrv As New frmQrSrv
Private mDevNo As String, mDevIP As String
Private mblnNeedSaveSrv As Boolean, mblnInitOk As Boolean
Public Sub ShowMe(ByVal DevNo As String, ByVal DevName As String, ByVal DevIP As String, ByVal frmobj As Object)
    mDevNo = DevNo
    mDevIP = DevIP
    mblnNeedSaveSrv = False
    Me.Caption = "设备(" & DevName & ")" & "参数设置"
    Me.Show , frmobj
End Sub
Private Sub InitFaceScheme()
'初始界面布局
    FraList.Top = 0
    FraList.Left = 0
    With TabList
        .Top = FraList.Top + FraList.Height
        .Left = FraList.Left
        .Height = Me.ScaleHeight - FraList.Height
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 1, "图像接收服务", mFrmImgSrv.hWnd, 0
        .InsertItem 2, "WorkList服务", mfrmWorkList.hWnd, 0
        .InsertItem 3, "Q/R 查询服务", mFrmQRSrv.hWnd, 0
        
        .Item(0).Enabled = False
        .Item(1).Enabled = False
        .Item(2).Enabled = False
    End With
    cmdCancel.Top = TabList.Top + TabList.Height - cmdCancel.Height - 50
    cmdSave.Top = TabList.Top + TabList.Height - cmdSave.Height - 50
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdDel_Click()
    If vfgList.TextMatrix(vfgList.Row, Col服务名) = "" Then Exit Sub
    If MsgBoxD(Me, "确实要删除服务(" & vfgList.TextMatrix(vfgList.Row, Col服务名) & ")吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If vfgList.TextMatrix(vfgList.Row, ColID) <> "" Then
        gstrSQL = "Zl_影像DICOM服务对_DELETE(" & vfgList.TextMatrix(vfgList.Row, ColID) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "删除服务")
    End If
    Call InitSrvList
End Sub
Private Function ValidData() As Boolean
Dim i As Long, j As Integer

    With vfgList
        For i = 1 To .Rows - 1
            If i <> .Rows - 1 Then
                If .TextMatrix(i, Col服务名) = "" Then MsgBoxD Me, "第" & i & "行 服务名 不能为空", vbInformation, gstrSysName: Exit Function
                If .TextMatrix(i, Col服务功能) = "" Then MsgBoxD Me, "第" & i & "行 服务功能 不能为空", vbInformation, gstrSysName: Exit Function
                If UBound(Split(Trim(.TextMatrix(i, Col服务IP)), ".")) <> 3 Then
                    MsgBoxD Me, "第" & i & "行 网关IP格式不正确，请检查！", vbInformation, gstrSysName: Exit Function
                Else
                    For j = 0 To 3
                        If Not IsNumeric(Split(Trim(.TextMatrix(i, Col服务IP)), ".")(j)) Then
                            MsgBoxD Me, "第" & i & "行 网关IP格式不正确，请检查！", vbInformation, gstrSysName: Exit Function
                        Else
                            If Split(Trim(.TextMatrix(i, Col服务IP)), ".")(j) < 0 Or Split(Trim(.TextMatrix(i, Col服务IP)), ".")(j) >= 256 Then
                                MsgBoxD Me, "第" & i & "行 网关IP格式不正确，请检查！", vbInformation, gstrSysName: Exit Function
                            End If
                        End If
                    Next
                End If
                If .TextMatrix(i, Col服务AE) = "" Then MsgBoxD Me, "第" & i & "行 网关AE 不能为空", vbInformation, gstrSysName: Exit Function
                If .TextMatrix(i, Col服务端口) = "" Then MsgBoxD Me, "第" & i & "行 网关端口 不能为空", vbInformation, gstrSysName: Exit Function
                If Not IsNumeric(.TextMatrix(i, Col服务端口)) Then MsgBoxD Me, "第" & i & "行 网关端口 必须为数值", vbInformation, gstrSysName: Exit Function
                If .TextMatrix(i, Col设备AE) = "" Then MsgBoxD Me, "第" & i & "行 设备AE 不能为空", vbInformation, gstrSysName: Exit Function
            Else
                If .TextMatrix(i, Col服务名) <> "" Then
                    If .TextMatrix(i, Col服务功能) = "" Then MsgBoxD Me, "第" & i & "行 服务功能 不能为空", vbInformation, gstrSysName: Exit Function
                    If UBound(Split(Trim(.TextMatrix(i, Col服务IP)), ".")) <> 3 Then
                        MsgBoxD Me, "第" & i & "行 网关IP格式不正确，请检查！", vbInformation, gstrSysName: Exit Function
                    Else
                        For j = 0 To 3
                            If Not IsNumeric(Split(Trim(.TextMatrix(i, Col服务IP)), ".")(j)) Then
                                MsgBoxD Me, "第" & i & "行 网关IP格式不正确，请检查！", vbInformation, gstrSysName: Exit Function
                            Else
                                If Split(Trim(.TextMatrix(i, Col服务IP)), ".")(j) < 0 Or Split(Trim(.TextMatrix(i, Col服务IP)), ".")(j) >= 256 Then
                                    MsgBoxD Me, "第" & i & "行 网关IP格式不正确，请检查！", vbInformation, gstrSysName: Exit Function
                                End If
                            End If
                        Next
                    End If
                    If .TextMatrix(i, Col服务AE) = "" Then MsgBoxD Me, "第" & i & "行 网关AE 不能为空", vbInformation, gstrSysName: Exit Function
                    If .TextMatrix(i, Col服务端口) = "" Then MsgBoxD Me, "第" & i & "行 网关端口 不能为空", vbInformation, gstrSysName: Exit Function
                    If Not IsNumeric(.TextMatrix(i, Col服务端口)) Then MsgBoxD Me, "第" & i & "行 网关端口 必须为数值", vbInformation, gstrSysName: Exit Function
                    If .TextMatrix(i, Col设备AE) = "" Then MsgBoxD Me, "第" & i & "行 设备AE 不能为空", vbInformation, gstrSysName: Exit Function
                ElseIf i = 1 Then
                    Exit Function
                End If
            End If
        Next
    End With
    ValidData = True
End Function
Private Sub cmdOK_Click()
Dim i As Long, Count As Long
    If Not ValidData Then Exit Sub
    On Error GoTo errHandle
    If Trim(vfgList.TextMatrix(vfgList.Rows - 1, Col服务名)) = "" Then
        Count = vfgList.Rows - 2
    Else
        Count = vfgList.Rows - 1
    End If
    
    For i = 1 To Count
        With vfgList
            Select Case .TextMatrix(i, ColID)
                Case ""
                    gstrSQL = "Zl_影像DICOM服务对_INSERT('" & Trim(.TextMatrix(i, Col服务名)) & "','" & mDevNo & "','" & _
                                            .TextMatrix(i, Col服务功能) & "','" & .TextMatrix(i, ColPACS角色) & "','" & _
                                            Trim(.TextMatrix(i, Col服务IP)) & "','" & Trim(.TextMatrix(i, Col服务AE)) & "','" & _
                                            Trim(.TextMatrix(i, Col服务端口)) & "','" & mDevIP & "','" & _
                                            Trim(.TextMatrix(i, Col设备AE)) & "','" & Trim(.TextMatrix(i, Col服务端口)) & "')"
                Case Else
                    gstrSQL = "Zl_影像DICOM服务对_UPDATE(" & .TextMatrix(i, ColID) & ",'" & Trim(.TextMatrix(i, Col服务名)) & "','" & mDevNo & "','" & _
                                            .TextMatrix(i, Col服务功能) & "','" & .TextMatrix(i, ColPACS角色) & "','" & _
                                            Trim(.TextMatrix(i, Col服务IP)) & "','" & Trim(.TextMatrix(i, Col服务AE)) & "','" & _
                                            Trim(.TextMatrix(i, Col服务端口)) & "','" & mDevIP & "','" & _
                                            Trim(.TextMatrix(i, Col设备AE)) & "','" & Trim(.TextMatrix(i, Col服务端口)) & "')"
            End Select
            Call zlDatabase.ExecuteProcedure(gstrSQL, "保存服务")
        End With
    Next
    Call InitSrvList(Trim(vfgList.TextMatrix(vfgList.Row, Col服务名)))
    mblnNeedSaveSrv = False
    MsgBoxD Me, "服务保存成功，请为该服务设定参数！", vbInformation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSave_Click()
    If mblnNeedSaveSrv Then MsgBoxD Me, "上方服务列表中有变动且未保存，请先保存服务列表变动", vbInformation, gstrSysName: Exit Sub
    Select Case TabList.Selected.Caption
        Case "图像接收服务"
            Call mFrmImgSrv.SavePara
        Case "WorkList服务"
            Call mfrmWorkList.SavePara
        Case "Q/R 查询服务"
            Call mFrmQRSrv.SavePara
    End Select
    MsgBoxD Me, "参数保存成功", vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    InitFaceScheme '初始化界面
    InitSrvList '初始化服务列表
End Sub
Private Sub InitSrvList(Optional ByVal strSrvName As String)
    mblnInitOk = False
    With vfgList
        .Clear
        .Rows = 2
        .Cols = 10
        .ColWidth(ColID) = 0 'ID
        .ColWidth(Col服务名) = 1000 '服务名
        .ColWidth(Col服务功能) = 1200 '服务功能
        .ColWidth(ColPACS角色) = 500  'PACS角色 SUC/SUP
        .ColWidth(Col服务IP) = 1400 '服务IP
        .ColWidth(Col服务AE) = 1000  '服务AE
        .ColWidth(Col服务端口) = 500 '服务端口
        .ColWidth(Col设备IP) = 0
        .ColWidth(Col设备AE) = 1000
        .ColWidth(Col设备端口) = 0
        .TextMatrix(0, ColID) = "ID"
        .TextMatrix(0, Col服务名) = "服务名"
        .TextMatrix(0, Col服务功能) = "服务功能"
        .TextMatrix(0, ColPACS角色) = "角色"
        .TextMatrix(0, Col服务IP) = "网关IP"
        .TextMatrix(0, Col服务AE) = "网关AE"
        .TextMatrix(0, Col服务端口) = "端口"
        .TextMatrix(0, Col设备IP) = "设备IP"
        .TextMatrix(0, Col设备AE) = "设备AE"
        .TextMatrix(0, Col设备端口) = "设备端口"
        

        .FixedAlignment(ColID) = flexAlignCenterCenter
        .FixedAlignment(Col服务名) = flexAlignCenterCenter
        .FixedAlignment(Col服务功能) = flexAlignCenterCenter
        .FixedAlignment(ColPACS角色) = flexAlignCenterCenter
        .FixedAlignment(Col服务IP) = flexAlignCenterCenter
        .FixedAlignment(Col服务AE) = flexAlignCenterCenter
        .FixedAlignment(Col服务端口) = flexAlignCenterCenter
        .FixedAlignment(Col设备IP) = flexAlignCenterCenter
        .FixedAlignment(Col设备AE) = flexAlignCenterCenter
        .FixedAlignment(Col设备端口) = flexAlignCenterCenter
        
        .ColAlignment(ColID) = flexAlignLeftCenter
        .ColAlignment(Col服务名) = flexAlignLeftCenter
        .ColAlignment(Col服务功能) = flexAlignLeftCenter
        .ColAlignment(ColPACS角色) = flexAlignLeftCenter
        .ColAlignment(Col服务IP) = flexAlignLeftCenter
        .ColAlignment(Col服务AE) = flexAlignLeftCenter
        .ColAlignment(Col服务端口) = flexAlignLeftCenter
        .ColAlignment(Col设备IP) = flexAlignLeftCenter
        .ColAlignment(Col设备AE) = flexAlignLeftCenter
        .ColAlignment(Col设备端口) = flexAlignLeftCenter

        .Editable = flexEDKbdMouse
        .ColComboList(Col服务功能) = "图像接收|Worklist|Q/R服务|胶片接收"
    End With
    Call FillBILL(strSrvName)
    mblnInitOk = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload mFrmImgSrv
    Unload mfrmWorkList
    Unload mFrmQRSrv
End Sub
Private Sub FillBILL(ByVal strSrvName As String)
Dim rsTemp As ADODB.Recordset, i As Integer
    On Error GoTo errHandle
    With vfgList
        gstrSQL = "select B.设备IP地址,B.设备AE名称,B.设备端口,B.服务ID,B.服务名,B.服务功能,B.PACS角色,B.PACSIP地址,B.PACSAE名称,B.PACS端口" & _
                    " from 影像设备目录 A,影像DICOM服务对 B where A.设备号=[1] and A.设备号=B.设备号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取服务信息", CStr(mDevNo))
        Do Until rsTemp.EOF
            .TextMatrix(rsTemp.AbsolutePosition, ColID) = rsTemp!服务ID
            .TextMatrix(rsTemp.AbsolutePosition, Col服务名) = Nvl(rsTemp!服务名)
            .TextMatrix(rsTemp.AbsolutePosition, Col服务功能) = Nvl(rsTemp!服务功能)
            .TextMatrix(rsTemp.AbsolutePosition, ColPACS角色) = Nvl(rsTemp!PACS角色)
            .TextMatrix(rsTemp.AbsolutePosition, Col服务IP) = Nvl(rsTemp!PACSIP地址)
            .TextMatrix(rsTemp.AbsolutePosition, Col服务AE) = Nvl(rsTemp!PACSAE名称)
            .TextMatrix(rsTemp.AbsolutePosition, Col服务端口) = Nvl(rsTemp!PACS端口)
            .TextMatrix(rsTemp.AbsolutePosition, Col设备IP) = Nvl(rsTemp!设备IP地址)
            .TextMatrix(rsTemp.AbsolutePosition, Col设备AE) = Nvl(rsTemp!设备AE名称)
            .TextMatrix(rsTemp.AbsolutePosition, Col设备端口) = Nvl(rsTemp!设备端口)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If strSrvName <> "" Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, Col服务名) = strSrvName Then .Row = i: .RowSel = i: .Col = 0
            Next
        Else
            .Row = 1
            vfgList_EnterCell
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If mblnInitOk = True Then '初始化完成才触发,初始化期间触发无效
        mblnNeedSaveSrv = True
        If Col = Col服务功能 Then
            Select Case vfgList.TextMatrix(Row, Col)
                Case "图像接收"
                    vfgList.TextMatrix(vfgList.Row, ColPACS角色) = "SCP"
                Case "Worklist"
                    vfgList.TextMatrix(vfgList.Row, ColPACS角色) = "SCP"
                Case "Q/R服务"
                    vfgList.TextMatrix(vfgList.Row, ColPACS角色) = "SCU"
                Case "胶片接收"
                    vfgList.TextMatrix(vfgList.Row, ColPACS角色) = "SCP"
            End Select
        End If
    End If
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .Col = ColPACS角色 Then
            mblnNeedSaveSrv = True
            .Editable = flexEDNone
            If (.TextMatrix(.Row, Col服务功能) <> "胶片接收") Then
                If .TextMatrix(.Row, .Col) = "SCP" Then
                    .TextMatrix(.Row, .Col) = "SCU"
                Else
                    .TextMatrix(.Row, .Col) = "SCP"
                End If
            End If
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vfgList_EnterCell()
Dim i As Long

    If Trim(vfgList.TextMatrix(vfgList.Row, Col服务名) = "") And vfgList.Rows <= 2 Then
        TabList.Enabled = False
    Else
        TabList.Enabled = True
    End If
    TabList.Visible = True
    Select Case vfgList.TextMatrix(vfgList.Row, Col服务功能)
        Case "图像接收", "胶片接收"
            TabList.Item(0).Enabled = True
            TabList.Item(0).Visible = True
            TabList.Item(0).Selected = True
            TabList.Item(1).Enabled = False
            TabList.Item(2).Enabled = False
            TabList.Item(1).Visible = False
            TabList.Item(2).Visible = False
            cmdSave.Enabled = True
            Call mFrmImgSrv.ShowRefresh(IIf(vfgList.TextMatrix(vfgList.Row, ColID) = "", 0, vfgList.TextMatrix(vfgList.Row, ColID)))
        Case "Worklist"
            TabList.Item(0).Enabled = False
            TabList.Item(0).Visible = False
            TabList.Item(1).Enabled = True
            TabList.Item(1).Visible = True
            TabList.Item(1).Selected = True
            TabList.Item(2).Enabled = False
            TabList.Item(2).Visible = False
            cmdSave.Enabled = True
            Call mfrmWorkList.ShowRefresh(IIf(vfgList.TextMatrix(vfgList.Row, ColID) = "", 0, vfgList.TextMatrix(vfgList.Row, ColID)))
        Case "Q/R服务"
            TabList.Item(0).Enabled = False
            TabList.Item(0).Visible = False
            TabList.Item(1).Enabled = False
            TabList.Item(1).Visible = False
            TabList.Item(2).Selected = True
            TabList.Item(2).Enabled = True
            TabList.Item(2).Visible = True
            cmdSave.Enabled = True
            Call mFrmQRSrv.ShowRefresh(IIf(vfgList.TextMatrix(vfgList.Row, ColID) = "", 0, vfgList.TextMatrix(vfgList.Row, ColID)))
        Case Else
            TabList.Enabled = False
            TabList.Visible = False
            cmdSave.Enabled = False
    End Select
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = Col服务IP Or Col = Col服务端口 Then
        If InStr("0123456789." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    ElseIf Col = Col服务功能 Or Col = ColPACS角色 Then
        If KeyAscii <> 13 Then KeyAscii = 0
    End If
End Sub
