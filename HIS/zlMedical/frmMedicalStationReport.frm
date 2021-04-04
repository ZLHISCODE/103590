VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicalStationReport 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "体检报告"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   3135
      Left            =   135
      ScaleHeight     =   3075
      ScaleWidth      =   5400
      TabIndex        =   1
      Top             =   1860
      Width           =   5460
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1530
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   5430
      _cx             =   9578
      _cy             =   2699
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
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
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
      Begin VB.Line lnY 
         Index           =   1
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
      Begin VB.Line lnX 
         Index           =   1
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   6345
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":0000
            Key             =   "公共"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":039A
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":0734
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":0ACE
            Key             =   "单据"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":0E68
            Key             =   "附加"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":1202
            Key             =   "up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationReport.frx":13C4
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgX 
      Height          =   135
      Left            =   1080
      MousePointer    =   7  'Size N S
      Top             =   1695
      Width           =   5115
   End
End
Attribute VB_Name = "frmMedicalStationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean
Private mfrmReport As Object
Private mclsCore As New clsCISCore
Private mlngKey As Long
Private mfrmMain As Object
Private mvarParam As Variant
Private mblnNoAllowChange As Boolean
Private mblnDataMoved As Boolean
Private mblnShow As Boolean         '是否显示内容

Private Enum mCol
    公共
    状态
    项目
    执行科室
    执行状态
    报告人
    时间
    报告id
    单据id
    No
    结算途径
End Enum

Public Function zlMenuClick(ByVal frmMain As Object, ByVal strMenuItem As String, Optional ByVal strParam As String = "") As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '参数：lngKey 档案ID
    '--------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    Dim strNO As String
    Dim lng单据id As Long
    Dim lng报告id As Long
    Dim lng记录性质 As Long
    
    On Error GoTo errHand
    
    mvarParam = Split(strParam, "'")
    
    mlngKey = Val(mvarParam(0))
    
    Set mfrmMain = frmMain
    
    Select Case strMenuItem
    Case "刷新"
        
        lngSvrKey = Val(vsf.RowData(vsf.Row))
        Call zlClearData
        Call RefreshData(strMenuItem)
        Call RestoreRow(vsf, lngSvrKey)
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
        
    Case "填写报告", "查看报告"
        
        If Val(vsf.RowData(vsf.Row)) <= 0 Then Exit Function
        
        strNO = vsf.TextMatrix(vsf.Row, GetCol(vsf, "No"))
        lng单据id = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "单据id")))
        lng报告id = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "报告id")))
        lng记录性质 = IIf(Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "结算途径"))) = 1, 2, 1)
        
        If strNO = "" Then Exit Function
        If lng单据id = 0 And lng报告id = 0 Then Exit Function
                
        Call EditReport(frmMain, strNO, lng记录性质, lng单据id, lng报告id, "", IIf(strMenuItem = "填写报告", False, True), True, , , , False, , mblnDataMoved, "001")
                            
        '退出后进行刷新
        mblnNoAllowChange = True
        
        lngSvrKey = Val(vsf.RowData(vsf.Row))
        Call zlClearData
        Call RefreshData("刷新")
        Call RestoreRow(vsf, lngSvrKey)
        
        mblnNoAllowChange = False
        
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)

    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    'If ErrCenter = 1 Then Resume
End Function

Public Sub zlClearData(Optional ByVal strPart As String = "所有")
    '--------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '--------------------------------------------------------------------------------------------------
    Dim blnSvr As Boolean
    
    blnSvr = mblnNoAllowChange
    
    mblnNoAllowChange = True
    
    Call ResetVsf(vsf)
    Call AppendSapceRows(vsf, lnX, lnY)
        
    On Error Resume Next
    If Not (mfrmReport Is Nothing) Then mfrmReport.zlClearData
    
    mblnNoAllowChange = blnSvr
End Sub

Public Property Get Body(ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property

Public Property Let ShowResult(ByVal v_Data As Boolean)
    
    mblnShow = v_Data

    picContainer.Visible = mblnShow
    imgX.Visible = mblnShow
    
    Call Form_Resize
        
    If mblnShow Then
        
        If mfrmReport Is Nothing Then Set mfrmReport = mclsCore.ShowFileObject(Me, Me.picContainer, 0, 0, gcnOracle, "", glngSys, "", "")
                
        Call RefreshData("报告")
        
    Else

        On Error Resume Next
        
        If Not (mfrmReport Is Nothing) Then mfrmReport.zlClearData
        
    End If
    
End Property

Private Function RefreshData(ByVal strMenu As String) As Boolean
    
    Dim rs As New ADODB.Recordset
    
    Select Case strMenu
    Case "刷新"
        If Val(mvarParam(1)) = 0 Then
            '未开始之前的查询
            mblnDataMoved = False
            gstrSQL = GetPublicSQL(SQL.病人所有项目)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mvarParam(0)))
        Else
            
            gstrSQL = "Select X. *, " & _
                               "Y.名称 As 执行科室, " & _
                               "Z.名称 As 项目, " & _
                               "Decode(X.报告id, Null, Decode(D.病历文件id, Null, '', '单据'), Decode(H.书写人, Null, '单据', '报告')) As 状态, " & _
                               "D.病历文件id As 单据id, " & _
                               "H.书写人 As 报告人, " & _
                               "To_Char(H.书写日期, 'yyyy-mm-dd hh24:mi') As 时间,Decode(x.复查清单id,0,0,Null,0,255) As 前景色 " & _
                        "From ( Select E.ID, " & _
                                      "B.执行科室id,b.复查清单id, " & _
                                      "A.诊疗项目id, " & _
                                      "A.结算途径, " & _
                                      "Decode(G.执行状态, 1, '完全执行', 2, '取消执行', 3, '正在执行', '') As 执行状态, G.报告id, G.NO, " & _
                                      "Decode(A.病人id, Null, '', '附加') As 公共 " & _
                               "From 体检项目医嘱 B, 体检项目清单 A, 体检人员档案 C, 体检登记记录 D, 病人医嘱记录 E, 病人医嘱发送 G " & _
                               "Where A.ID = B.清单id " & _
                                     "And B.病人id = C.病人id " & _
                                     "And C.登记id = A.登记id " & _
                                     "AND D.ID=C.登记id And d.体检号=E.挂号单 And e.病人id=c.病人id " & _
                                     "AND E.病人来源=4 " & _
                                     "AND E.医嘱状态<>4 " & _
                                     "And E.诊疗类别 In ('C', 'D') " & _
                                     "And G.医嘱id = E.ID And b.医嘱id In (e.ID,e.相关id) "
            gstrSQL = gstrSQL & _
                                     "And C.ID = [1] " & _
                               ") X, 部门表 Y, 诊疗项目目录 Z, 诊疗单据应用 D, 病人病历记录 H " & _
                        "Where x.执行科室id = y.ID " & _
                              "And Z.ID = X.诊疗项目id " & _
                              "And X.报告id = H.ID(+) " & _
                              "And D.应用场合(+) = 4 " & _
                              "And X.诊疗项目id = D.诊疗项目id(+) " & _
                        "Order By Y.名称"
            
            '数据转储处理
            '----------------------------------------------------------------------------------------------------------
            mblnDataMoved = DataMove(mlngKey, 2)
            If mblnDataMoved Then
                gstrSQL = Replace(gstrSQL, "体检项目医嘱", "H体检项目医嘱")
                gstrSQL = Replace(gstrSQL, "体检项目清单", "H体检项目清单")
                gstrSQL = Replace(gstrSQL, "体检人员档案", "H体检人员档案")
                gstrSQL = Replace(gstrSQL, "体检登记记录", "H体检登记记录")
                gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
                gstrSQL = Replace(gstrSQL, "病人医嘱发送", "H病人医嘱发送")
                gstrSQL = Replace(gstrSQL, "病人病历记录", "H病人病历记录")
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        End If
        
        If rs.BOF = False Then
            
            Call LoadGrid(vsf, rs, , , ils13)
            Call AppendSapceRows(vsf, lnX, lnY)
            
        End If
    
    Case "报告"
    
        If Not (mfrmReport Is Nothing) Then Call mfrmReport.zlMenuClick(Me, Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "报告id"))), "刷新")
        
    End Select
    
End Function

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
            
    picContainer.Visible = False
    imgX.Visible = False
    
    Set mfrmReport = Nothing
    
    strVsf = ",255,4,1,1,[公共];,255,4,1,1,[状态];项目,2400,1,1,1,;执行科室,1080,1,1,1,;执行状态,900,1,1,1,;报告人,900,1,1,1,;时间,1670,1,1,1,;报告id,0,1,1,1,;单据id,0,1,1,1,;No,0,1,1,1,;结算途径,0,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    
    Set vsf.Cell(flexcpPicture, 0, 0) = ils13.ListImages("公共").Picture
    Set vsf.Cell(flexcpPicture, 0, 1) = ils13.ListImages("状态").Picture
    
    Call InitCISCore(gcnOracle)
    
    Call AppendSapceRows(vsf, lnX, lnY)
        
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
        
    Call InitLoad
       
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    If imgX.Top > Me.ScaleHeight - 1000 Then imgX.Top = Me.ScaleHeight - 1000
    
    With vsf
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = IIf(mblnShow, imgX.Top, Me.ScaleHeight)
    End With
    
    If mblnShow Then
        With imgX
            .Left = vsf.Left
            .Width = vsf.Width
            .Height = 45
            .BorderStyle = 0
        End With
    
        With picContainer
            .Left = 0
            .Top = imgX.Top + imgX.Height
            .Width = vsf.Width
            .Height = Me.ScaleHeight - .Top
        End With
    End If
    
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmReport = Nothing
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX.Top = imgX.Top + Y
    
    If imgX.Top < 1500 Then imgX.Top = 1500
    If Me.Height - imgX.Top - imgX.Height < 1000 Then imgX.Top = Me.Height - imgX.Height - 1000
    
            
    Form_Resize
End Sub

Private Sub picContainer_Resize()
    On Error Resume Next
    
    If Not (mfrmReport Is Nothing) Then
        mfrmReport.Width = picContainer.Width
        mfrmReport.Height = picContainer.Height
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNoAllowChange Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    
    Call SelectRow(vsf, OldRow, NewRow)
    
    If mblnShow Then Call RefreshData("报告")
    
    On Error GoTo errHand
    Call mfrmMain.ActiveFormEnabled
    
errHand:
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col < 2)
End Sub

Private Sub vsf_DblClick()
    '
    Dim strNO As String
    Dim lng单据id As Long
    Dim lng报告id As Long
    Dim lng记录性质 As Long
    
    If Val(vsf.RowData(vsf.Row)) <= 0 Then Exit Sub
        
    strNO = vsf.TextMatrix(vsf.Row, GetCol(vsf, "No"))
    lng单据id = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "单据id")))
    lng报告id = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "报告id")))
    lng记录性质 = IIf(Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "结算途径"))) = 1, 2, 1)
    
    If strNO = "" Or lng报告id = 0 Then Exit Sub
                
    Call EditReport(mfrmMain, strNO, lng记录性质, lng单据id, lng报告id, "", True, True, , , , False, , , "001")
    
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.焦点
    Call SelectRow(vsf, 1, vsf.Row)
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.非焦点
    Call SelectRow(vsf, 1, vsf.Row)
End Sub

