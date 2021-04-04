VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicalStationCase 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3075
      ScaleWidth      =   5400
      TabIndex        =   0
      Top             =   1950
      Width           =   5460
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1530
      Left            =   240
      TabIndex        =   1
      Top             =   210
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
      Begin VB.Line lnX 
         Index           =   1
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY 
         Index           =   1
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   6195
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationCase.frx":0000
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationCase.frx":039A
            Key             =   "图标"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgX 
      Height          =   135
      Left            =   2505
      MousePointer    =   7  'Size N S
      Top             =   1635
      Width           =   5115
   End
End
Attribute VB_Name = "frmMedicalStationCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean
Private mfrmCaseFile As Object
Private mclsCore As New clsCISCore
Private mlngKey As Long
Private mfrmMain As Object

Public Function zlMenuClick(ByVal frmMain As Object, ByVal lngKey As Long, ByVal strMenuItem As String) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '参数：lngKey 档案ID
    '--------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    On Error GoTo errHand
    
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    Select Case strMenuItem
    Case "刷新"
        
        lngSvrKey = Val(vsf.RowData(vsf.Row))
        Call zlClearData
        Call RefreshData(strMenuItem)
        Call RestoreRow(vsf, lngSvrKey)
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
                   
    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub zlClearData(Optional ByVal strPart As String = "所有")
    '--------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '--------------------------------------------------------------------------------------------------
    
    Call ResetVsf(vsf)
    Call AppendSapceRows(vsf, lnX, lnY)
    
End Sub

Public Property Get Body(ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property


Private Function RefreshData(ByVal strMenu As String) As Boolean
    
    Dim rs As New ADODB.Recordset
    Dim strStart As String
    Dim strEnd As String
        
    Select Case strMenu
    Case "刷新"

    
        strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "历史病历范围", "今  天"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "历史病历范围", "今  天"), 2)
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    

        gstrSQL = "SELECT A.ID," & _
                        "TO_CHAR(A.书写日期,'yyyy-mm-dd hh24:mi') AS 时间," & _
                        "A.病历名称," & _
                        "A.书写人 AS 医生 " & _
                    "FROM 病人病历记录 A,体检人员档案 B,病历文件目录 C " & _
                    "WHERE A.病人id=B.病人id " & _
                        "AND A.病历种类 IN (1,2,-2) " & _
                        "AND A.文件id=C.ID  AND 作废日期 IS NULL " & _
                        "AND A.书写日期 BETWEEN [1] AND [2] " & _
                        "AND B.ID=[3] ORDER BY A.书写日期 DESC "
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(strStart), CDate(strEnd), mlngKey)
        
        If rs.BOF = False Then
                                    
            Call LoadGrid(vsf, rs, , , ils13)
            Call AppendSapceRows(vsf, lnX, lnY)
            
        End If
    
    Case "病历"
        Call mfrmCaseFile.zlMenuClick(Me, Val(vsf.RowData(vsf.Row)), "刷新")
    End Select
    
End Function

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    Dim strVsf As String
    
    strVsf = "时间,1670,1,1,1,;病历名称,2400,1,1,1,;医生,1200,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
        
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
    
    Set mfrmCaseFile = mclsCore.ShowFileObject(Me, Me.picContainer, 0, 0, gcnOracle, "", 100, "", "")
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With vsf
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = imgX.Top
    End With
    
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
    
    Call AppendSapceRows(vsf, lnX, lnY)
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
    
    If Not (mfrmCaseFile Is Nothing) Then
        mfrmCaseFile.Width = picContainer.Width
        mfrmCaseFile.Height = picContainer.Height
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    
    Call SelectRow(vsf, OldRow, NewRow)
    
    Call RefreshData("病历")
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.焦点
    Call SelectRow(vsf, 1, vsf.Row)
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.非焦点
    Call SelectRow(vsf, 1, vsf.Row)
End Sub


