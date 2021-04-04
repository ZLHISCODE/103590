VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmUserQueryReleation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户常用查询配置"
   ClientHeight    =   5844
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   9144
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.8
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserQueryReleation.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5844
   ScaleWidth      =   9144
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
      Height          =   3612
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   8652
      _cx             =   15261
      _cy             =   6371
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.8
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
      BackColorSel    =   13082765
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
      Editable        =   2
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
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   5280
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取 消(&Q)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7728
      TabIndex        =   4
      Top             =   5280
      Width           =   1185
   End
   Begin VB.ComboBox cbxUser 
      Appearance      =   0  'Flat
      Height          =   312
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   220
      Width           =   3012
   End
   Begin VB.ComboBox cbxDepart 
      Appearance      =   0  'Flat
      Height          =   312
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   216
      Width           =   2892
   End
   Begin VB.Image imgNoCheck 
      Height          =   252
      Left            =   240
      Picture         =   "frmUserQueryReleation.frx":000C
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   252
      Left            =   240
      Picture         =   "frmUserQueryReleation.frx":037E
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label labNote 
      BackColor       =   &H00DDF8FB&
      Caption         =   "方案说明："
      Height          =   732
      Left            =   240
      TabIndex        =   7
      Top             =   4300
      Width           =   8652
   End
   Begin VB.Label Label2 
      Caption         =   "科室名称："
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "用户名称:"
      Height          =   252
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "frmUserQueryReleation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TColDef
    cdName = 0          '方案名称
    cdDefault = 1       '是否默认
    cdCommonUse = 2     '是否常用
    cdStationName = 3   '站点名称
    cdSchemeDescript = 4 '方案描述
End Enum

Private mlngModuleNo As Long
Private mlngUserId As Long
Private mblnIsLoading As Boolean

Private mblnIsOK As Boolean


Public Function ShowUserScheme(owner As Object, ByVal lngModuleNo As Long, Optional ByVal lngUserId As Long = 0) As Boolean
    mblnIsOK = False
    
    ShowUserScheme = False
    mlngModuleNo = lngModuleNo
    mlngUserId = lngUserId
    
    Me.Show 1, owner
    
    ShowUserScheme = mblnIsOK
End Function

Private Sub LoadDepartInfo()
'载入科室信息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    cbxDepart.Clear
    
    If mlngUserId <> 0 Then
        cbxDepart.BackColor = &H8000000F
        cbxDepart.Enabled = False
        Exit Sub
    Else
        cbxDepart.BackColor = vbWhite
        cbxDepart.Enabled = True
    End If
    
    strSQL = "Select ID,名称 From 部门表 A, 部门性质说明 B where A.ID=B.部门ID And B.工作性质='检查' Order By 名称"
    Set rsData = ExecuteSql(strSQL, "查询部门信息")
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    While Not rsData.EOF
        
        cbxDepart.AddItem NVL(rsData!名称)
        cbxDepart.ItemData(cbxDepart.ListCount - 1) = Val(NVL(rsData!Id))
        
        Call rsData.MoveNext
    Wend
    
    cbxDepart.AddItem ""
    
    cbxDepart.ListIndex = 0
End Sub

Private Sub LoadUserInfo()
'载入用户信息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngUserId As Long
    Dim lngIndex As Long
    Dim blnIsQueryCurUser As Boolean
    
    cbxUser.Clear
    
    If mlngUserId <= 0 Then
        cbxUser.BackColor = vbWhite
        cbxUser.Enabled = True
        
        If cbxDepart.Text = "" Then Exit Sub
        
        
        strSQL = "Select ID, 姓名,用户名 From 人员表 A, 部门人员 B, 上机人员表 C Where A.ID=B.人员ID  And A.ID=C.人员ID And B.部门ID=[1] Order By 姓名"
        Set rsData = ExecuteSql(strSQL, "查询人员信息", cbxDepart.ItemData(cbxDepart.ListIndex))
        
        If rsData.RecordCount <= 0 Then Exit Sub
    Else
        cbxUser.BackColor = &H8000000F
        cbxUser.Enabled = False
        
        strSQL = "Select Id, 姓名,'当前用户' as 用户名 From 人员表 Where ID=[1]"
        Set rsData = ExecuteSql(strSQL, "查询当前人员信息", mlngUserId)
        
        If rsData Is Nothing Then Exit Sub
        If rsData.RecordCount <= 0 Then Exit Sub
    End If
        
    While Not rsData.EOF
        lngUserId = Val(NVL(rsData!Id))
        
        cbxUser.AddItem NVL(rsData!用户名) & "-" & NVL(rsData!姓名)
        cbxUser.ItemData(cbxUser.ListCount - 1) = lngUserId
        
        If lngUserId = mlngUserId Then
            lngIndex = cbxUser.ListCount - 1
        End If
        
        Call rsData.MoveNext
    Wend
    
    cbxUser.ListIndex = lngIndex
End Sub

Public Sub LoadSchemeConfig()
'载入用户方案配置
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim i As Long
    Dim blnIsUser As Boolean
    
    vsfGrid.Rows = 1
    
    If cbxUser.Text = "" Then Exit Sub
    
    strSQL = "Select A.ID, A.方案名称, B.用户ID, " & vbCrLf & _
            "   case when B.用户ID Is Null then A.是否默认 else decode(B.是否默认, null,B.是否默认,B.是否默认+1) End As 是否默认, " & vbCrLf & _
            "   case when B.用户ID Is Null then A.是否常用 else decode(B.是否常用, null,B.是否常用,B.是否常用+1) End As 是否常用, " & vbCrLf & _
            "   B.所属站点 , A.方案说明 " & vbCrLf & _
            " From 影像查询方案 A, 影像查询关联 B " & vbCrLf & _
            " Where A.ID=B.查询方案ID(+) And A.使用状态=1 And A.所属模块=[1] And B.用户ID(+)=[2] order by 方案序号"
              

    Set rsData = ExecuteSql(strSQL, "载入所有方案", mlngModuleNo, Val(cbxUser.ItemData(cbxUser.ListIndex)))
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Filter = "用户ID=" & Val(cbxUser.ItemData(cbxUser.ListIndex))
    
    '判断是否根据用户进行了配置
    blnIsUser = IIf(rsData.RecordCount <= 0, False, True)
 
    rsData.Filter = ""

    
    vsfGrid.Rows = rsData.RecordCount + 1
    
    i = 1
    While Not rsData.EOF
        vsfGrid.RowData(i) = NVL(rsData!Id)
        
        vsfGrid.Cell(flexcpText, i, cdName) = NVL(rsData!方案名称)
        
        If Val(NVL(rsData!是否默认)) > IIf(blnIsUser, 1, 0) Then
            vsfGrid.Cell(flexcpData, i, cdDefault) = 1
            vsfGrid.Cell(flexcpPicture, i, cdDefault) = imgCheck.Picture
        Else
            vsfGrid.Cell(flexcpData, i, cdDefault) = 0
            vsfGrid.Cell(flexcpPicture, i, cdDefault) = imgNoCheck.Picture
        End If
        
        If Val(NVL(rsData!是否常用)) > IIf(blnIsUser, 1, 0) Then
            vsfGrid.Cell(flexcpData, i, cdCommonUse) = 1
            vsfGrid.Cell(flexcpPicture, i, cdCommonUse) = imgCheck.Picture
        Else
            vsfGrid.Cell(flexcpData, i, cdCommonUse) = 0
            vsfGrid.Cell(flexcpPicture, i, cdCommonUse) = imgNoCheck.Picture
        End If
                
        vsfGrid.Cell(flexcpText, i, cdStationName) = NVL(rsData!所属站点)
        vsfGrid.Cell(flexcpText, i, cdSchemeDescript) = NVL(rsData!方案说明)
        
        i = i + 1
        
        Call rsData.MoveNext
    Wend
    
    vsfGrid.Cell(flexcpBackColor, 1, cdName, i - 1, cdName) = &HDDF8FB
    vsfGrid.Cell(flexcpPictureAlignment, 1, cdDefault, i - 1, cdDefault) = flexPicAlignCenterCenter
    vsfGrid.Cell(flexcpPictureAlignment, 1, cdCommonUse, i - 1, cdCommonUse) = flexPicAlignCenterCenter

End Sub

Private Sub cbxDepart_Change()
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call LoadUserInfo
    Call LoadStationInfos
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxDepart_Click()
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call LoadUserInfo
    Call LoadStationInfos
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxUser_Change()
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call LoadSchemeConfig
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxUser_Click()
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call LoadSchemeConfig
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub SaveConfig(ByVal lngUserId As Long)
    Dim i As Long
    Dim blnIsDefault As Boolean
    Dim blnIsCommonUse As Boolean
    Dim strStation As String
    Dim strSQL As String
    Dim blnIsStartTrans As Boolean
    
    strSQL = "zl_影像查询_清除关联(" & lngUserId & ")"
    Call ExecuteCmd(strSQL, "清除用户查询关联")
    
    On Error GoTo errHandle:
    
    blnIsStartTrans = False
    For i = 1 To vsfGrid.Rows - 1
        blnIsDefault = IIf(vsfGrid.Cell(flexcpData, i, cdDefault) = 1, True, False)
        blnIsCommonUse = IIf(vsfGrid.Cell(flexcpData, i, cdCommonUse) = 1, True, False)
        strStation = vsfGrid.Cell(flexcpText, i, cdStationName)
        
        If blnIsDefault Or blnIsCommonUse Or Trim(strStation) <> "" Then
            If blnIsStartTrans = False Then
                gcnOracle.BeginTrans
                blnIsStartTrans = True
            End If
            
            strSQL = "zl_影像查询_更新关联(" & lngUserId & "," & Val(vsfGrid.RowData(i)) & "," & _
                                            IIf(blnIsDefault, 1, 0) & "," & IIf(blnIsCommonUse, 1, 0) & ",'" & _
                                            strStation & "')"
            Call ExecuteCmd(strSQL, "用户查询关联")
        End If
    Next i
    
    If blnIsStartTrans Then gcnOracle.CommitTrans
Exit Sub
errHandle:
    If blnIsStartTrans Then gcnOracle.RollbackTrans
    Debug.Print "SaveConfig Err:" & Err.Description
    Err.Raise -1, "SaveConfig", "[SaveConfig]保存用户关联配置错误>>" & Err.Description
    Resume
End Sub


Private Sub cmdSure_Click()
'确认处理
On Error GoTo errHandle
    Call SaveConfig(Val(cbxUser.ItemData(cbxUser.ListIndex)))
    mblnIsOK = True
    
    MsgBox "方案设置成功,配置将在下次加载时生效。", vbOKOnly, Me.Caption
    
    Unload Me
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub Form_Load()
    mblnIsLoading = True
    
    Call InitList
    
    Call LoadDepartInfo
    Call LoadUserInfo
    Call LoadStationInfos
    
    Call LoadSchemeConfig
    
    mblnIsLoading = False
End Sub

Private Function GetStationCfgString(ByVal strDepartName As String) As String
    Dim strResult As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strCurStationName As String
    
    strCurStationName = UCase(StationName)
    
    strResult = " |" & strCurStationName
    GetStationCfgString = strResult
    
    strSQL = "Select 工作站 From ZlClients Where 部门=[1] or 部门 Is Null Order By 工作站"
    
    Set rsData = ExecuteSql(strSQL, "查询站点", strDepartName)
    
    If rsData Is Nothing Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    While Not rsData.EOF
        If NVL(rsData!工作站) <> strCurStationName Then
            If strResult <> "" Then strResult = strResult & "|"
            strResult = strResult & "|" & NVL(rsData!工作站)
        End If
        
        Call rsData.MoveNext
    Wend
    
    GetStationCfgString = strResult
End Function

Private Sub LoadStationInfos()
    vsfGrid.ColComboList(3) = GetStationCfgString(cbxDepart.Text)
End Sub

Private Sub InitList()
    vsfGrid.Cell(flexcpText, 0, cdName) = "方案名称"
    vsfGrid.Cell(flexcpText, 0, cdDefault) = "是否默认"
    vsfGrid.Cell(flexcpText, 0, cdCommonUse) = "是否常用"
    vsfGrid.Cell(flexcpText, 0, cdStationName) = "所属站点"
    vsfGrid.Cell(flexcpText, 0, cdSchemeDescript) = "方案说明"
    
    
    vsfGrid.ColHidden(4) = True
    
    
    
    vsfGrid.ColWidth(0) = 4000
End Sub

Private Sub vsfGrid_Click()
On Error GoTo errHandle
    Dim i As Long
    
    If vsfGrid.RowSel < 1 Then Exit Sub
    
    Select Case vsfGrid.ColSel
        Case 1  '是否默认列处理
            If vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdDefault) = 1 Then
                vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdDefault) = 0
                vsfGrid.Cell(flexcpPicture, vsfGrid.RowSel, cdDefault) = imgNoCheck.Picture
            Else
                For i = 1 To vsfGrid.Rows - 1
                    vsfGrid.Cell(flexcpData, i, cdDefault) = 0
                    vsfGrid.Cell(flexcpPicture, i, cdDefault) = imgNoCheck.Picture
                Next i
                
                vsfGrid.Cell(flexcpPicture, vsfGrid.RowSel, cdDefault) = imgCheck.Picture
                vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdDefault) = 1
            End If
        Case 2  '是否常用列处理
            If vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdCommonUse) = 1 Then
                vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdCommonUse) = 0
                vsfGrid.Cell(flexcpPicture, vsfGrid.RowSel, cdCommonUse) = imgNoCheck.Picture
            Else
                vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdCommonUse) = 1
                vsfGrid.Cell(flexcpPicture, vsfGrid.RowSel, cdCommonUse) = imgCheck.Picture
            End If
    End Select
    
    
Exit Sub
errHandle:
    Debug.Print "vsfGrid_DblClick Err:" & Err.Description
End Sub


 
Private Sub vsfGrid_SelChange()
On Error GoTo errHandle
    labNote.Caption = "方案说明：" & vsfGrid.Cell(flexcpText, vsfGrid.RowSel, 4)
Exit Sub
errHandle:
    Debug.Print "" & Err.Description
End Sub

Private Sub vsfGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 3 Then Cancel = True
End Sub
