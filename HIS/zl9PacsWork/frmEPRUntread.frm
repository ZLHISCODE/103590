VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRUntread 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "版本回退"
   ClientHeight    =   3555
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5700
   Icon            =   "frmEPRUntread.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdUntread 
      Caption         =   "回退(&U)"
      Height          =   375
      Left            =   2790
      TabIndex        =   3
      Top             =   2955
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   4095
      TabIndex        =   2
      Top             =   2955
      Width           =   1230
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2085
      Left            =   285
      TabIndex        =   1
      Top             =   720
      Width           =   5055
      _cx             =   8916
      _cy             =   3678
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
      Rows            =   2
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   255
      Picture         =   "frmEPRUntread.frx":058A
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "该病历审阅修订情况如下，可以逐步回退以撤消对病历的修订和签名。"
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   195
      Width           =   4500
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEPRUntread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mblnOK As Boolean
Private mSignLevel As EPRSignLevelEnum

'临时变量

Dim lngCount As Long
Dim edtType As EditTypeEnum

Public Function ShowMe(ByVal lngID As Long, _
    ByVal EditType As EditTypeEnum, _
    ByRef lngVersion As Long, _
    ByRef lngSignKey As Long, _
    ByRef fParent As Object) As Boolean
    
    '功能：显示病历的版本修订变化情况，让用户决定执行回退
    '参数：lngId-电子病历记录id
    '      lngVersion-需要回退的版本号
    '      lngSignKey-需要回退的签名Key值
    '返回：成功与否
    
    '----------------------
    '读取版本修订变化
    
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    edtType = EditType
    err = 0: On Error GoTo errHand
    strSql = "Select 0 As Id, -null As 对象标记, 1 As 版本, '新增编辑' As 操作, l.创建人 As 人员," & _
        "        To_Char(l.创建时间, 'yyyy-mm-dd hh24:mi:ss') As 时间,-1 as 排序 " & _
        " From 电子病历记录 l" & _
        " Where L.ID = [1]" & _
        " Union All" & _
        " Select c.Id, c.对象标记, c.开始版 As 版本," & _
        "        Decode(l.病历种类, 4, Decode(c.要素表示, 3, '护士长', '护士')," & _
        "                Decode(c.要素表示, 3, '主任医师', 2, '主治医师', '经治医师')) || Decode(c.开始版, 1, '签名', '修订') As 操作," & _
        "        c.内容文本 As 人员, Rtrim(Substr(c.对象属性, Instr(c.对象属性, ';', 1, 4) + 1)) As 时间,对象标记 as 排序 " & _
        " From 电子病历记录 l, 电子病历内容 c" & _
        " Where L.ID = c.文件ID And L.ID = [1] And c.对象类型 = 8" & _
        " Union All" & _
        " Select c.Id,  -null as 对象标记, l.最后版本 As 版本, '正在修订…' As 操作, l.保存人 As 人员," & _
        "        To_Char(l.保存时间, 'yyyy-mm-dd hh24:mi:ss') As 时间,c.对象标记 as 排序 " & _
        " From 电子病历记录 l," & _
        "      (Select Max(c.开始版) As 开始版, Max(Id + 1) As Id,Max(对象标记+1) as 对象标记 From 电子病历内容 c Where c.文件id = [1] And c.对象类型 = 8) c" & _
        " Where L.ID = [1] And L.最后版本 > c.开始版" & _
        " Order By 排序 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    With Me.vfgThis
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 0: .ColHidden(0) = True
        .ColWidth(1) = 0: .ColHidden(1) = True
        .ColWidth(6) = 0: .ColHidden(6) = True
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        For lngCount = .FixedRows To .Rows - 1
            If InStr(1, .TextMatrix(lngCount, 5), ";") > 0 Then .TextMatrix(lngCount, 5) = Left(.TextMatrix(lngCount, 5), 19)
        Next
        If EditType = cprET_单病历编辑 Then
            If .Rows <= .FixedRows + 1 Then Me.cmdUntread.Enabled = False
        Else
            If .Rows <= .FixedRows + 2 Then Me.cmdUntread.Enabled = False
        End If
    End With
    
    mSignLevel = GetUserSignLevel(UserInfo.ID)
    If mSignLevel <= cprSL_空白 Then Me.cmdUntread.Enabled = False
    
    Me.Show vbModal, fParent
    If mblnOK = False Then ShowMe = False: Unload Me: Exit Function
    
    '----------------------
    '返回
    lngVersion = Val(vfgThis.TextMatrix(1, 2))
    lngSignKey = Val(vfgThis.TextMatrix(1, 1))
    
    ShowMe = True: Unload Me: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdUntread_Click()
    mblnOK = True: Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
    '窗口显示在最前面
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
err.Clear
End Sub

Private Sub vfgThis_RowColChange()
    Dim blnEnable As Boolean
    If mSignLevel <= cprSL_空白 Then Me.cmdUntread.Enabled = False: Exit Sub
    blnEnable = True
    If edtType = cprET_单病历编辑 Then
        If vfgThis.Rows <= vfgThis.FixedRows + 1 Then blnEnable = False
    Else
        If vfgThis.Rows <= vfgThis.FixedRows + 2 Then blnEnable = False
    End If
    cmdUntread.Enabled = blnEnable And (vfgThis.Row = 1)
End Sub
