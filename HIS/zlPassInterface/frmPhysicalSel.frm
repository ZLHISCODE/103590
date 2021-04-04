VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPhysicalSel 
   BorderStyle     =   0  'None
   Caption         =   "1"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2475
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraMK 
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         ItemData        =   "frmPhysicalSel.frx":0000
         Left            =   1200
         List            =   "frmPhysicalSel.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         ItemData        =   "frmPhysicalSel.frx":001A
         Left            =   1200
         List            =   "frmPhysicalSel.frx":0024
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtDate 
         Height          =   300
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "2015-02-01"
         Top             =   795
         Width           =   1455
      End
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         ItemData        =   "frmPhysicalSel.frx":0034
         Left            =   3960
         List            =   "frmPhysicalSel.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "妊娠日期"
         Height          =   180
         Index           =   4
         Left            =   3000
         TabIndex        =   17
         ToolTipText     =   "日期格式:YYYY-MM-DD"
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "哺乳\妊娠"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   840
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "肝功"
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "肾功"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   14
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   3
         Left            =   3960
         TabIndex        =   13
         ToolTipText     =   "妊娠开始日期"
         Top             =   1080
         Width           =   90
      End
   End
   Begin VB.Timer tmrThis 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2880
      Top             =   1920
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   2490
      Index           =   1
      Left            =   9120
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   0
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   2490
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   8775
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   7
      Top             =   0
      Width           =   8655
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   8760
      TabIndex        =   6
      Top             =   1875
      Width           =   8760
      Begin VB.Timer tmrAir 
         Interval        =   1000
         Left            =   2040
         Top             =   120
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "退出(&Q)"
         Height          =   350
         Left            =   7080
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.ListBox lstItem 
      Height          =   1740
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VSFlex8Ctl.VSFlexGrid vsSymptom 
      Height          =   1740
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Width           =   6615
      _cx             =   11668
      _cy             =   3069
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   11250603
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   400
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPhysicalSel.frx":004E
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
      WordWrap        =   -1  'True
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
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   2160
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line lin 
      Index           =   7
      X1              =   5760
      X2              =   6435
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line lin 
      Index           =   6
      X1              =   5880
      X2              =   6555
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line lin 
      Index           =   5
      X1              =   5880
      X2              =   6555
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line lin 
      Index           =   4
      X1              =   5880
      X2              =   6555
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line lin 
      Index           =   3
      X1              =   5880
      X2              =   6555
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lin 
      Index           =   2
      X1              =   5880
      X2              =   6555
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line lin 
      Index           =   1
      X1              =   5760
      X2              =   6435
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   5880
      X2              =   6555
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "frmPhysicalSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mstrPhysical As String
'Private mobjAir As clsAirBubble zl9ComLib.clsAirBubble
Private mobjAir As Object
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mbytSex As Integer    '性别 0-男,1-女
Private mbyt场合 As Byte      '1-门诊编辑，2-住院编辑
Private mstrDelOrder As String   '记录删除症状记录序号:序号1，序号2，...
Private mlngNum As Long  '记录时间修改位数
Private mIntWaitTime As Integer   '记录气泡延迟时间，由于调用气泡时传人的第一个参数是Picture,导致气泡不能自动延迟
'症状列号
Private Enum COL症状列号
    COL_序号 = 0
    COL_状态 = 1
    col_症状 = 2
    col_开始日期 = 3
    col_结束日期 = 4
    COL_医生 = 5
    COL_操作 = 6
End Enum

Private Enum CBO_IX
    Ix_肝功 = 0
    IX_肾功 = 1
    IX_妊娠 = 2
    IX_日期 = 3
End Enum

Public Sub ShowMe(ByRef objMap As clsPassMap)
'功能:
'参数:
'      byt场合-1-门诊编辑，2-住院编辑
    Dim objPoint As RECT
    
    mlng病人ID = objMap.PassPati.lng病人ID
    mlng主页ID = IIf(objMap.lngModel = PM_门诊编辑, objMap.PassPati.lng挂号ID, objMap.PassPati.lng主页ID)
    mbytSex = IIf(objMap.PassPati.str性别 = "女", 1, 0)
    mbyt场合 = IIf(objMap.lngModel = PM_门诊编辑, 1, 2)
    
    Call GetWindowRect(objMap.objCmdBar.hWnd, objPoint)
    If gbytPass = DT Or gbytPass = YWS Or gbytPass = ZL Then
        Me.Width = 2040
        Me.Top = objPoint.Top * Screen.TwipsPerPixelY + objMap.objCmdBar.Height
        Me.Left = objPoint.Left * Screen.TwipsPerPixelX - Me.Width + objMap.objCmdBar.Width
    ElseIf gbytPass = TYT Then
        Me.Width = 8760
        Me.Top = objPoint.Top * Screen.TwipsPerPixelY + objMap.objCmdBar.Height
        Me.Left = objPoint.Left * Screen.TwipsPerPixelX - Me.Width + objMap.objCmdBar.Width
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        Me.Width = 6150: Me.Height = 2100
        Me.Top = objPoint.Top * Screen.TwipsPerPixelY + objMap.objCmdBar.Height
        Me.Left = objPoint.Left * Screen.TwipsPerPixelX - Me.Width + objMap.objCmdBar.Width
        
        lstItem.Visible = False
        vsSymptom.Visible = False
        fraMK.Visible = True
    End If
    Me.Show 1, objMap.frmMain
End Sub

Private Sub LoadDict()
'功能:加载病生理情况字典数据
'参数:bytSex:0-男,1-女
    Dim strSQL As String, i As Long
    Dim strList As String
    Dim rsDict As ADODB.Recordset
    If gbytPass <> ZL Then
        strSQL = "Select 名称 From 病生理情况 Order by 编码"
    Else
        strList = "妊娠,哺乳,肝功能不全,严重肝功能不全,肾功能不全,严重肾功能不全"
        strSQL = "Select /*+cardinality(A,10)*/ a.Column_Value As 名称 From Table(f_Str2list([1])) A"
    End If
    On Error GoTo errH
    Set rsDict = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strList)
    lstItem.Clear
    With rsDict
        For i = 1 To .RecordCount
            If !名称 = "孕妇" Or !名称 = "哺乳期" Or !名称 = "妊娠" Or !名称 = "哺乳" Then
                If mbytSex = 1 Then lstItem.AddItem !名称
            Else
                lstItem.AddItem !名称
            End If
            .MoveNext
        Next
    End With
    Exit Sub

errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadLists()
'功能:加载病人的病生理情况
'参数:bytSex:0-男,1-女
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    Dim lngTmp As Long
    Dim strTmp As String

    If gbytPass = YWS Or gbytPass = DT Or gbytPass = TYT Or gbytPass = ZL Then
       Call LoadDict
    
       If mbyt场合 = 1 Then
           lngTmp = Val(zlDatabase.GetPara(21, glngSys))
           strSQL = "Select 病生理情况" & vbNewLine & _
                   "From 病人挂号记录" & vbNewLine & _
                   "Where 病人id = [1] And 登记时间 > Trunc(Sysdate-[2]) And 病生理情况 Is Not Null And Rownum = 1"
       Else
           strSQL = "Select 信息值 As 病生理情况" & vbNewLine & _
                   "From 病案主页从表 Where 病人id = [1] And 主页id = [2] And 信息名 = '病生理情况'"
       End If
       
       On Error GoTo errH
       Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, IIf(mbyt场合 = 1, lngTmp, mlng主页ID))
       If rsTmp.RecordCount > 0 Then
           For i = 0 To lstItem.ListCount - 1
               lstItem.Selected(i) = InStr(1, "," & rsTmp!病生理情况 & ",", "," & lstItem.List(i) & ",") > 0
           Next
       End If
       
       mstrPhysical = GetLists
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        '加载数据
        txtDate.Text = ""
        cboStatus(Ix_肝功).Clear
        cboStatus(IX_肾功).Clear
        cboStatus(IX_妊娠).Clear
        '-1-不确定（默认）,0-无肝损害,1-肝功能不全,2-轻度肝损害,3-中度肝损害,4-重度肝损害"
        strTmp = ",无肝损害,肝功能不全,轻度肝损害,中度肝损害,重度肝损害"
        For i = LBound(Split(strTmp, ",")) To UBound(Split(strTmp, ","))
            cboStatus(Ix_肝功).AddItem Split(strTmp, ",")(i)
        Next
        cboStatus(Ix_肝功).ListIndex = -1 '缺省
        
        '-1-不确定（默认）；0-无肾损害；1-肾功能不全；2-轻度肾损害；3-中度肾损害；4-重度肾损害
        strTmp = ",无肾损害,肾功能不全,轻度肾损害,中度肾损害,重度肾损害"
        For i = LBound(Split(strTmp, ",")) To UBound(Split(strTmp, ","))
            cboStatus(IX_肾功).AddItem Split(strTmp, ",")(i)
        Next
        cboStatus(IX_肾功).ListIndex = -1
        
        If mbytSex = 1 Then '女性病人
            '哺乳\妊娠
            strTmp = ",哺乳,妊娠"
            For i = LBound(Split(strTmp, ",")) To UBound(Split(strTmp, ","))
                cboStatus(IX_妊娠).AddItem Split(strTmp, ",")(i)
            Next
            cboStatus(IX_妊娠).ListIndex = -1
            
            cboStatus(IX_妊娠).Enabled = True
            cboStatus(IX_妊娠).BackColor = &H80000005
        Else
            cboStatus(IX_妊娠).Enabled = False
            cboStatus(IX_妊娠).BackColor = &H8000000C
        End If
        '日期缺省设置
        txtDate.Enabled = False
        txtDate.BackColor = &H8000000C
        
        strTmp = Get病人病生理情况(mlng病人ID, IIf(mbyt场合 = 1, 0, mlng主页ID))
        mstrPhysical = strTmp
        If strTmp <> "" Then
            For i = LBound(Split(strTmp, ",")) To UBound(Split(strTmp, ","))
                If InStr(",无肝损害,肝功能不全,轻度肝损害,中度肝损害,重度肝损害,", "," & Split(strTmp, ",")(i) & ",") > 0 Then
                    zlControl.CboLocate cboStatus(Ix_肝功), Split(strTmp, ",")(i)
                ElseIf InStr(",无肾损害,肾功能不全,轻度肾损害,中度肾损害,重度肾损害,", "," & Split(strTmp, ",")(i) & ",") > 0 Then
                    zlControl.CboLocate cboStatus(IX_肾功), Split(strTmp, ",")(i)
                ElseIf InStr(",哺乳,妊娠,", "," & Split(strTmp, ",")(i) & ",") > 0 Then
                    zlControl.CboLocate cboStatus(IX_妊娠), Split(strTmp, ",")(i)
                ElseIf InStr("," & Split(strTmp, ",")(i), ",妊娠日期|") > 0 Then '妊娠日期|2015-02-01
                    strTmp = Split(strTmp, ",")(i)
                    txtDate.Text = Split(strTmp, "|")(1)
                End If
            Next
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetLists() As String
'功能：获取选择的病生理情况字符串，以逗号分隔
    Dim i As Long, strRetu As String
    
    For i = 0 To lstItem.ListCount - 1
        If lstItem.Selected(i) Then strRetu = strRetu & "," & lstItem.List(i)
    Next
    
    If strRetu <> "" Then GetLists = Mid(strRetu, 2)
End Function

Private Sub cmdQuit_Click()
    '检查
    If CheckCell Then Exit Sub
    '保存数据
    Call SaveData
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyQ = KeyCode And Shift = vbCtrlMask Then
        Call cmdQuit_Click
    End If
End Sub

Private Sub Form_Load()
    
    Call LoadLists
    If gbytPass = TYT Then
        '初始化症状列
        Call InitSymptom
        '加载数据
        Call LoadSyptoms
        'Set mobjAir = New gobjComLib.clsAirBubble
        Set mobjAir = CreateObject("zl9ComLib.clsAirBubble")
    End If
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    '窗体边框设置
    Call InitFormBorder
    If gbytPass = DT Or gbytPass = YWS Then
        lstItem.Top = fraBorder(0).Height + 80
        lstItem.Left = fraBorder(3).Width + 80
        vsSymptom.Visible = False
    ElseIf gbytPass = TYT Then
        lstItem.Top = fraBorder(0).Height + 80
        lstItem.Left = fraBorder(3).Width + 80
        vsSymptom.Top = fraBorder(0).Height + 80
        vsSymptom.Left = fraBorder(3).Width + 80 + lstItem.Width + 80
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        fraMK.Top = fraBorder(0).Height + 80
        fraMK.Left = fraBorder(3).Width + 80
    End If
    cmdQuit.Left = picBottom.Width - cmdQuit.Width - 200
    
End Sub

Private Sub SaveData()
    Dim strTmp As String
    Dim bytFunc As Byte
    Dim arrSQL As Variant
    Dim curDate As Date
    Dim i As Long
    arrSQL = Array()
    
    If gbytPass = MK And gstrVersion = "4.0" Then
        strTmp = IIf(cboStatus(Ix_肝功).Tag <> "", "," & cboStatus(Ix_肝功).Tag, "")
        strTmp = strTmp & IIf(cboStatus(IX_肾功).Tag <> "", "," & cboStatus(IX_肾功).Tag, "")
        strTmp = strTmp & IIf(cboStatus(IX_妊娠).Tag <> "", "," & cboStatus(IX_妊娠).Tag, "")
        strTmp = strTmp & IIf(txtDate.Tag <> "", "," & txtDate.Tag, "")
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    Else
        strTmp = GetLists
    End If
    
    If strTmp <> mstrPhysical Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        If mbyt场合 = 1 Then    '1-门诊编辑
            arrSQL(UBound(arrSQL)) = "Zl_病人病生理情况_Insert(" & mlng病人ID & ",0," & mlng主页ID & ",'" & strTmp & "')"   '此时mlng主页ID是门诊挂号ID
        Else    '2-住院编辑
            arrSQL(UBound(arrSQL)) = "Zl_病人病生理情况_Insert(" & mlng病人ID & "," & mlng主页ID & ",Null,'" & strTmp & "')"
        End If
    End If

    If gbytPass = TYT Then
        '组装删除sql
        If mstrDelOrder <> "" Then
            For i = 0 To UBound(Split(mstrDelOrder, ",")) - 1    '最后一个不取
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人症状记录_Update(3," & mlng病人ID & "," & mlng主页ID & "," & Split(mstrDelOrder, ",")(i) & ")"
            Next
        End If
        curDate = zlDatabase.Currentdate
        With vsSymptom
            For i = .FixedRows To .Rows - 2  '最后一行空白
                bytFunc = Val(.TextMatrix(i, COL_状态))
                If bytFunc = 2 Then  '新增 序号在过程中取最大值插入
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人症状记录_Update(1," & mlng病人ID & "," & mlng主页ID & "," & Val(.TextMatrix(i, COL_序号)) & " ,'" & _
                                             .RowData(i) & "','" & .TextMatrix(i, col_症状) & "',To_Date('" & .TextMatrix(i, col_开始日期) & "','YYYY-MM-DD HH24:MI:SS')," & _
                                             "To_date('" & .TextMatrix(i, col_结束日期) & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.姓名 & _
                                             "',To_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'))"
                ElseIf bytFunc = 3 Then   '修改
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人症状记录_Update(2," & mlng病人ID & "," & mlng主页ID & "," & Val(.TextMatrix(i, COL_序号)) & " ,'" & _
                                             .RowData(i) & "','" & .TextMatrix(i, col_症状) & "',To_Date('" & .TextMatrix(i, col_开始日期) & "','YYYY-MM-DD HH24:MI:SS')," & _
                                             "To_date('" & .TextMatrix(i, col_结束日期) & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.姓名 & _
                                             "',To_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'))"
                End If
            Next

        End With
    End If

    On Error GoTo errH
    '首先执行删除，再修改，其次才新增 否则因为乱序而出错。
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "病生状态")
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitSymptom()
'功能: 初始化病人症状记录表头
    Dim strCol As String, arrHead As Variant
    Dim i As Long
    '状态: 0-未标记,1-原始，2-新增，3-修改
    strCol = "序号;状态;症状,2000,4;开始日期,1300,4;结束日期,1300,4;医生,1000,4;操作,50,1"
    arrHead = Split(strCol, ";")
    With vsSymptom
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows + 1    '缺省显示一行空白

        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeColumns

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)

            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrDelOrder = ""
    mstrPhysical = ""
    If Not mobjAir Is Nothing Then
        mobjAir.CloseAirBubble
        Set mobjAir = Nothing
    End If
End Sub

Private Sub tmrAir_Timer()
'功能:弹出气泡的时候，设置mIntWaitTime=3
    If mIntWaitTime = 0 Then Exit Sub
    mIntWaitTime = mIntWaitTime - 1
    If mIntWaitTime = 1 Then
        If Not mobjAir Is Nothing Then
            mobjAir.CloseAirBubble
        End If
        mIntWaitTime = 0
    End If
End Sub

Private Sub tmrThis_Timer()
    Dim lngTmp As Long
    With vsSymptom
        If .Col = col_开始日期 Or .Col = col_结束日期 Then
            lngTmp = .EditSelStart
            If .EditSelText = "-" Then
                Call Vs_EditSelChange(lngTmp - 1)    '选不中"-"
            ElseIf lngTmp = 0 Or lngTmp = 5 Or lngTmp = 8 Then
                mlngNum = 0
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strDate As String

    With vsSymptom
        If .TextMatrix(Row, col_症状) <> "" Then
            '设置默认值
            If .TextMatrix(Row, col_开始日期) = "" And .TextMatrix(Row, col_结束日期) = "" Then
                strDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
                .TextMatrix(Row, col_开始日期) = strDate
                .TextMatrix(Row, col_结束日期) = strDate
            End If
            If .TextMatrix(Row, col_开始日期) <> "" And .TextMatrix(Row, col_结束日期) <> "" _
               And (.TextMatrix(Row, COL_医生) = "" Or .Cell(flexcpData, Row, COL_医生) <> UserInfo.姓名) Then
                .TextMatrix(Row, COL_医生) = UserInfo.姓名
                .Cell(flexcpAlignment, Row, COL_医生) = flexAlignLeftCenter
            End If
        End If
        '状态更新
        If .TextMatrix(Row, COL_状态) = "1" Then
            If .TextMatrix(Row, Col) <> .Cell(flexcpData, Row, Col) Then
                .TextMatrix(Row, COL_状态) = "3"   '3-修改
            End If
        ElseIf .TextMatrix(Row, COL_状态) = "" And .TextMatrix(Row, COL_医生) <> "" Then  '医生录入代表一行数据录入完毕
            .TextMatrix(Row, COL_状态) = "2"   '2-新增
        End If

    End With
End Sub

Private Sub vsSymptom_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsSymptom

        If col_症状 = NewCol Then
            .ColComboList(col_症状) = "..."
            .FocusRect = flexFocusLight
        Else
            .ColComboList(col_症状) = ""
            .FocusRect = flexFocusLight
        End If

        If .TextMatrix(.Row, col_症状) <> "" And .TextMatrix(.Rows - 1, col_症状) <> "" Then
            .Rows = .Rows + 1
        End If
    End With
End Sub

Private Sub vsSymptom_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With vsSymptom
        '光标靠左
        If Row > .FixedRows Then
            .Cell(flexcpAlignment, Row, Col) = flexAlignLeftCenter
        End If
        '医生、操作列不可编辑
        If COL_医生 = Col Or COL_操作 = Col Then
            Cancel = True
        End If
        
        If col_开始日期 = Col Or Col = col_结束日期 Then
            If .TextMatrix(Row, col_症状) = "" Then
                Cancel = True '必须症状不为空的情况才允许编辑
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSymptom As String

    If col_症状 = Col Then
        On Error Resume Next
        strSymptom = gobjPass.inputDiagside
        If Err.Number <> 0 Then
            MsgBox "太元通接口调用失败，请检查是否配置有误。", vbInformation + vbOKOnly, Me.Caption
        End If
        Err.Clear: On Error GoTo 0
        If strSymptom <> "" Then
            vsSymptom.RowData(Row) = Val(Split(strSymptom, ";")(0))
            vsSymptom.TextMatrix(Row, Col) = Split(strSymptom, ";")(1)
            Call vsSymptom_AfterEdit(Row, Col)
        End If

    End If
End Sub

Private Sub vsSymptom_Click()
    Dim i As Long

    With vsSymptom
        If .Col = COL_操作 And Not .Cell(flexcpPicture, .Row, .Col) Is Nothing Then
            If .Rows - 1 = .FixedRows Then
                .Cell(flexcpText, .Row, col_症状, .Row, COL_操作) = ""
            Else
                If MsgBox("确定要删除症状【" & .TextMatrix(.Row, col_症状) & "】？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                '保存删除的SQL
                If Val(.TextMatrix(.Row, COL_状态)) = 1 Or Val(.TextMatrix(.Row, COL_状态)) = 3 Then
                    mstrDelOrder = mstrDelOrder & Val(.TextMatrix(.Row, COL_序号)) & ","
                End If
                '删掉显示行
                .RemoveItem (.Row)
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsSymptom
        If KeyCode = vbKeyF1 Then
            If .Col = col_症状 Then
            .ColComboList(vsSymptom.Col) = ""
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col_症状) <> "" Then
                If MsgBox("确定要删除症状【" & .TextMatrix(.Row, col_症状) & "】？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                '保存删除的SQL
                If Val(.TextMatrix(.Row, COL_状态)) = 1 Or Val(.TextMatrix(.Row, COL_状态)) = 3 Then
                    mstrDelOrder = mstrDelOrder & Val(.TextMatrix(.Row, COL_序号)) & ","
                End If
                '删掉显示行
                .RemoveItem (.Row)
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Then 'delete键与del键保持一致
        Call vsSymptom_KeyPressEdit(Row, Col, vbKeyDelete)
    End If
End Sub

Private Sub vsSymptom_KeyPress(KeyAscii As Integer)

    With vsSymptom
        If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            KeyAscii = 0
            If .Col <> COL_医生 Then
                .TextMatrix(.Row, .Col) = ""
            End If
        ElseIf KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            Call EnterNextCell
            If .Row = .Rows - 1 And .TextMatrix(.Row, col_症状) = "" And .Col >= col_结束日期 Then
                cmdQuit.SetFocus
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strChr As String
    Dim lngTmp As Long

    With vsSymptom

        If KeyAscii = vbKeyBack Then
            If col_症状 = Col And .ColComboList(col_症状) = "" Then
                .EditText = ""
            End If

            If Col = col_开始日期 Or Col = col_结束日期 Then
                If .EditText <> "" Then
                    If Len(.EditText) = .EditSelStart Then    '光标在最后
                        .EditText = Left(.EditText, .EditSelStart - 1)
                    ElseIf Len(.EditText) > .EditSelStart And .EditSelLength = 0 Then    '光标在中间
                        lngTmp = .EditSelStart
                        If lngTmp <> 0 Then
                            .EditText = Mid(.EditText, 1, lngTmp - 1) & Mid(.EditText, lngTmp)
                            .EditSelStart = lngTmp
                        End If
                        Exit Sub
                    ElseIf Len(.EditText) = .EditSelLength Then    '全选中
                        .EditText = ""
                    ElseIf .EditSelText <> "-" And .EditSelLength <> 0 Then
                        If .EditSelLength = 4 Then
                            .EditText = "2000" & Mid(.EditText, 5)
                            lngTmp = 4
                        ElseIf .EditSelStart <= 7 Then
                            .EditText = Left(.EditText, 5) & "01" & Right(.EditText, 3)
                            lngTmp = 7
                        Else
                            .EditText = Left(.EditText, 8) & "01"
                            lngTmp = 8
                        End If
                        Call Vs_EditSelChange(lngTmp)
                    End If
                End If
            End If
        ElseIf KeyAscii = vbKeyReturn Then
            KeyAscii = 0

            Call EnterNextCell: Exit Sub

        ElseIf KeyAscii = vbKeyDelete Then
            KeyAscii = 0
            .EditText = Mid(.EditText, 1, .EditSelStart)
            .EditSelStart = Len(.EditText)
            Exit Sub
        End If

        If Col = col_开始日期 Or Col = col_结束日期 Then
            '只能输入数字
            strChr = Chr(KeyAscii)

            If InStr("0123456789-", strChr) = 0 Then KeyAscii = 0: Exit Sub
            If .EditSelStart < 10 And Len(.EditText) = .EditSelStart Then
                '新增
                '年份
                If Len(.EditText) = 0 And InStr("123", strChr) = 0 Then KeyAscii = 0: Exit Sub

                '月份
                If (.EditSelStart = 4 Or .EditSelStart = 5) And InStr("01", strChr) = 0 Then KeyAscii = 0: Exit Sub
                If .EditSelStart = 6 Then
                    If (Val(Right(.EditText, 1)) = 0 And Val(strChr) = 0) Or (Val(Right(.EditText, 1)) = 1 And Val(strChr) > 2) Then
                        KeyAscii = 0: Exit Sub
                    End If
                End If
                '日期
                If (.EditSelStart = 7 Or .EditSelStart = 8) And InStr("0123", strChr) = 0 Then KeyAscii = 0: Exit Sub
                If .EditSelStart = 9 Then
                    If (Val(Right(.EditText, 1)) = 0 And Val(strChr) = 0) Or (Val(Right(.EditText, 1)) = 3 And Val(strChr) > 1) Then
                        KeyAscii = 0: Exit Sub
                    End If
                End If
                '自动添加分隔符
                If .EditSelStart = 4 Or .EditSelStart = 7 Then
                    .EditText = .EditText & "-"
                End If
            ElseIf .EditSelStart < Len(.EditText) And .EditSelLength = 0 And Len(.EditText) < 10 Then    '中间插入
                KeyAscii = 0
                lngTmp = .EditSelStart

                If lngTmp = 4 Or lngTmp = 7 Then
                    .EditText = Mid(.EditText, 1, lngTmp) & "-" & strChr & Mid(.EditText, lngTmp + 1)
                    .EditSelStart = lngTmp + 2
                Else
                    .EditText = Mid(.EditText, 1, lngTmp) & strChr & Mid(.EditText, lngTmp + 1)
                    .EditSelStart = lngTmp + 1
                End If
                Exit Sub
            ElseIf Len(.EditText) >= 10 Or .EditSelStart < Len(.EditText) Then
                '修改
                strChr = Chr(KeyAscii)
                KeyAscii = 0

                If .EditSelStart <= 4 Then
                    '年份
                    mlngNum = mlngNum + 1
                    If mlngNum = 1 And InStr("123", strChr) = 0 Then mlngNum = mlngNum - 1: Exit Sub
                    .EditText = Left(.EditText, mlngNum - 1) & strChr & Mid(.EditText, mlngNum + 1)
                    .EditSelStart = mlngNum
                    .EditSelLength = 4 - mlngNum
                    If mlngNum = 4 Then Call Vs_EditSelChange(5)
                ElseIf .EditSelStart >= 5 And .EditSelStart <= 7 Then
                    '月份
                    mlngNum = mlngNum + 1
                    If mlngNum = 1 And InStr("01", strChr) = 0 Then mlngNum = mlngNum - 1: Exit Sub
                    If mlngNum = 2 Then
                        If Val(Mid(.EditText, 6, 1)) = 0 And Val(strChr) = 0 Then
                            mlngNum = mlngNum - 1: Exit Sub  '月份最小：01
                        ElseIf Val(Mid(.EditText, 6, 1)) = 1 And Val(strChr) > 2 Then
                            mlngNum = mlngNum - 1: Exit Sub     '月份最大：12
                        End If
                    End If
                    .EditText = Left(.EditText, mlngNum + 4) & strChr & Mid(.EditText, mlngNum + 6)
                    .EditSelStart = 5 + mlngNum
                    .EditSelLength = 2 - mlngNum
                    If mlngNum = 2 Then Call Vs_EditSelChange(8)
                Else
                    '日期
                    mlngNum = mlngNum + 1
                    If mlngNum = 1 And InStr("0123", strChr) = 0 Then mlngNum = mlngNum - 1: Exit Sub
                    If mlngNum = 2 Then
                        If Val(Mid(.EditText, 9, 1)) = 0 And Val(strChr) = 0 Then
                            mlngNum = mlngNum - 1: Exit Sub  '日期最小：01
                        ElseIf Val(Mid(.EditText, 9, 1)) = 3 And Val(strChr) > 1 Then
                            mlngNum = mlngNum - 1: Exit Sub     '日期最大：31
                        End If
                    End If
                    .EditText = Left(.EditText, mlngNum + 7) & strChr & Right(.EditText, 2 - mlngNum)
                    .EditSelStart = 8 + mlngNum
                    .EditSelLength = 2 - mlngNum
                    If mlngNum = 2 Then Call Vs_EditSelChange(4)
                End If

            End If
        End If
    End With
End Sub

Private Sub EnterNextCell()
    With vsSymptom
        If .Col >= col_结束日期 Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = col_症状
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            .Col = .Col + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Private Sub vsSymptom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsSymptom
        If .Row >= .FixedRows And .Row <= .Rows - 2 Then
            '清除删除按钮
            Set .Cell(flexcpPicture, .FixedRows, COL_操作, .Rows - 1, COL_操作) = Nothing
            '显示删除按钮
            Set .Cell(flexcpPicture, .Row, COL_操作) = imgButtonDel.Picture
        End If
        If .Col = col_症状 Then
            .ToolTipText = "按F1可自由录入症状"
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub vsSymptom_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsSymptom
        If Col = col_症状 Then
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
        ElseIf Col = col_开始日期 Or Col = col_结束日期 Then
            tmrThis.Enabled = True
            .EditSelStart = 0
            .EditSelLength = 4
        End If
    End With
End Sub

Private Function ValidateDate(ByRef Row As Long, ByRef Col As Long) As Boolean
    Dim curDate As Date

    With vsSymptom    '日期格式检查
        If Col = col_开始日期 Or Col = col_结束日期 Then
            If Not IsDate(.TextMatrix(Row, Col)) Then '非日期提示
                Call mobjAir.OpenTransparentAirBubble(picBottom, "输入的日期格式不正确或日期错误", 1, 1, 80, &H99CCFF, vbRed, , 1, , , 咳嗽, True)
                mIntWaitTime = 3: ValidateDate = True
                Exit Function
            Else  '日期提示
                If .TextMatrix(Row, Col) <> "" Then
                    curDate = zlDatabase.Currentdate
                    curDate = Format(curDate, "yyyy-mm-dd")
                    If CDate(.TextMatrix(Row, Col)) > curDate Then
                        Call mobjAir.OpenTransparentAirBubble(picBottom, "您输入的日期不能大于当前日期。当前日期：" & curDate & "。", 1, 1, 80, &H99CCFF, vbRed, , 1, , , 咳嗽, True)
                        mIntWaitTime = 3: ValidateDate = True
                        Exit Function
                    End If
                    '开始日期<结束日期
                    If Col = col_结束日期 Then
                        If CDate(.TextMatrix(Row, col_开始日期)) > CDate(.TextMatrix(Row, Col)) Then
                            Call mobjAir.OpenTransparentAirBubble(picBottom, "开始日期不能大于结束日期", 1, 1, 80, &H99CCFF, vbRed, , 1, , , 咳嗽, True)
                            mIntWaitTime = 3: ValidateDate = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Sub LoadSyptoms()
'功能:加载病人症状
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim lngRow As Long

    strSQL = "Select 编码,序号,名称,开始日期,结束日期,记录人 From 病人症状记录 Where 病人id = [1] And 主页id = [2]"
    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    With vsSymptom
        .Redraw = flexRDNone
        .Rows = 2  '默认显示一行
        For i = 1 To rsTmp.RecordCount
            lngRow = .Rows - 1
            .RowData(lngRow) = rsTmp!编码 & ""
            .TextMatrix(lngRow, col_症状) = rsTmp!名称 & ""
            .Cell(flexcpData, lngRow, col_症状) = rsTmp!名称 & ""
            .TextMatrix(lngRow, COL_序号) = rsTmp!序号 & ""
            .TextMatrix(lngRow, col_开始日期) = Format(rsTmp!开始日期 & "", "yyyy-mm-dd")
            .Cell(flexcpData, lngRow, col_开始日期) = Format(rsTmp!开始日期 & "", "yyyy-mm-dd")
            .TextMatrix(lngRow, col_结束日期) = Format(rsTmp!结束日期 & "", "yyyy-mm-dd")
            .Cell(flexcpData, lngRow, col_结束日期) = Format(rsTmp!结束日期 & "", "yyyy-mm-dd")
            .TextMatrix(lngRow, COL_医生) = rsTmp!记录人 & ""
            .Cell(flexcpData, lngRow, COL_医生) = rsTmp!记录人 & ""
            .TextMatrix(lngRow, COL_状态) = "1"    '1-原始

            .Rows = lngRow + 2
            rsTmp.MoveNext
        Next
        .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter  '单元格内容左中对齐
        .Redraw = flexRDDirect

    End With

    Exit Sub
errH:
    If ErrCenter() Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckCell() As Boolean
'功能:检查必填内容不能为空,检查日期单元的正确性。
    Dim i As Long, j As Long
    If gbytPass <> TYT Then Exit Function
    With vsSymptom
        For i = .FixedRows To .Rows - 2
            For j = col_症状 To COL_医生
                If .TextMatrix(i, j) = "" Then
                    MsgBox "症状数据填写不完整，请填写完整后再退出", vbInformation + vbOKOnly, gstrSysName
                    '定位单元格
                    .Row = i: .Col = j
                    .EditCell
                    CheckCell = True
                    Exit Function
                End If
                If j = col_开始日期 Or j = col_结束日期 Then
                    If ValidateDate(i, j) Then
                        .Row = i: .Col = j
                        .EditCell
                        CheckCell = True
                        Exit Function
                    End If
                End If
            Next
        Next
    End With
End Function

Private Sub InitFormBorder()
    Dim i As Long
    
    fraBorder(0).Left = 0
    fraBorder(0).Top = 0
    fraBorder(0).Width = Me.ScaleWidth
    fraBorder(1).Top = fraBorder(0).Height
    fraBorder(1).Left = Me.ScaleWidth - fraBorder(1).Width
    fraBorder(1).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    fraBorder(2).Left = 0
    fraBorder(2).Top = Me.ScaleHeight - fraBorder(2).Height
    fraBorder(2).Width = Me.ScaleWidth
    fraBorder(3).Top = fraBorder(0).Height
    fraBorder(3).Left = 0
    fraBorder(3).Height = Me.ScaleHeight - fraBorder(0).Height * 2

    '边框设置
    For i = 0 To fraBorder.UBound
        fraBorder(i).BackColor = vbButtonFace
    Next
    Set lin(0).Container = fraBorder(0): Set lin(1).Container = fraBorder(0)
    Set lin(2).Container = fraBorder(1): Set lin(3).Container = fraBorder(1)
    Set lin(4).Container = fraBorder(2): Set lin(5).Container = fraBorder(2)
    Set lin(6).Container = fraBorder(3): Set lin(7).Container = fraBorder(3)
    lin(0).X1 = 0: lin(0).Y1 = 0: lin(0).X2 = Screen.Width: lin(0).Y2 = lin(0).Y1: lin(0).BorderColor = &H8000000F
    lin(1).X1 = 0: lin(1).Y1 = Screen.TwipsPerPixelY: lin(1).X2 = Screen.Width: lin(1).Y2 = lin(1).Y1: lin(1).BorderColor = &H8000000E
    lin(2).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).Y1 = 0: lin(2).X2 = lin(2).X1: lin(2).Y2 = Screen.Height: lin(2).BorderColor = &H80000011
    lin(3).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).Y1 = 0: lin(3).X2 = lin(3).X1: lin(3).Y2 = Screen.Height: lin(3).BorderColor = &H80000010
    lin(4).X1 = 0: lin(4).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).X2 = Screen.Width: lin(4).Y2 = lin(4).Y1: lin(4).BorderColor = &H80000011
    lin(5).X1 = 0: lin(5).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).X2 = Screen.Width: lin(5).Y2 = lin(5).Y1: lin(5).BorderColor = &H80000010
    lin(6).X1 = 0: lin(6).Y1 = 0: lin(6).X2 = lin(6).X1: lin(6).Y2 = Screen.Height: lin(6).BorderColor = &H8000000F
    lin(7).X1 = Screen.TwipsPerPixelX: lin(7).Y1 = 0: lin(7).X2 = lin(7).X1: lin(7).Y2 = Screen.Height: lin(7).BorderColor = &H8000000E
End Sub

Private Sub Vs_EditSelChange(ByVal lngSelNum As Long)
'当用户切换光标的时候触发
    With vsSymptom
        If lngSelNum <= 4 Then
            .EditSelStart = 0
            .EditSelLength = 4
        ElseIf lngSelNum <= 7 Then
            .EditSelStart = 5
            .EditSelLength = 2
        ElseIf lngSelNum <= 10 Then
            .EditSelStart = 8
            .EditSelLength = 2
        End If
        mlngNum = 0
    End With
End Sub

Private Sub cboStatus_Click(Index As Integer)
    If Index = IX_妊娠 Then
        If cboStatus(Index).Text = "妊娠" Then
            txtDate.Enabled = True
            txtDate.BackColor = &H80000005
            If Me.Visible Then txtDate.SetFocus
        Else
            txtDate.Enabled = False
            txtDate.BackColor = &H8000000C
            txtDate.Text = "": txtDate.Tag = ""
        End If
    End If
    cboStatus(Index).Tag = cboStatus(Index).Text
End Sub

Private Sub cboStatus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub txtDate_GotFocus()
    Call zlControl.TxtSelAll(txtDate)
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter Then
        cmdQuit.SetFocus
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtDate.Text <> "" Then
            txtDate.Text = zlStr.FullDate(txtDate.Text)
        End If
        zlCommFun.PressKey vbKeyTab
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    Dim datCurr As Date
    
    If Trim(txtDate.Text) <> "" Then
        If Not IsDate(txtDate) Then
            MsgBox "输入内容:" & txtDate.Text & "不是有效的日期格式:YYYY-MM-dd！", vbInformation + vbOKOnly, gstrSysName
            Cancel = True
        Else
            datCurr = zlDatabase.Currentdate
            If DateDiff("D", CDate(txtDate), datCurr) < 0 Then
                MsgBox "妊娠日期超过当前时间！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
            Else
                txtDate.Text = Format(txtDate.Text, "YYYY-MM-DD")
                txtDate.Tag = "妊娠日期|" & txtDate.Text
            End If
        End If
    Else
        MsgBox "妊娠日期不能为空！", vbInformation + vbOKOnly, gstrSysName
        Cancel = True
    End If
    If Cancel Then
        Call txtDate_GotFocus
    End If
End Sub
