VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceCopy 
   AutoRedraw      =   -1  'True
   Caption         =   "复制医嘱"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmAdviceCopy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   465
      TabIndex        =   8
      ToolTipText     =   "F1"
      Top             =   5730
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   5010
      Left            =   15
      TabIndex        =   0
      Top             =   555
      Width           =   9435
      _cx             =   16642
      _cy             =   8837
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceCopy.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
      Editable        =   2
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
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   4
      Top             =   5730
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6090
      TabIndex        =   3
      Top             =   5730
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清(&R)"
      Height          =   350
      Left            =   2745
      TabIndex        =   2
      ToolTipText     =   "Ctrl+R"
      Top             =   5730
      Width           =   1100
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   1635
      TabIndex        =   1
      ToolTipText     =   "Ctrl+A"
      Top             =   5730
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Height          =   615
      Left            =   15
      TabIndex        =   9
      Top             =   -75
      Width           =   9420
      Begin VB.CommandButton cmdPati 
         Height          =   240
         Left            =   1740
         Picture         =   "frmAdviceCopy.frx":077D
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "选择病人(F4)"
         Top             =   225
         Width           =   255
      End
      Begin VB.TextBox txtPati 
         Height          =   300
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   195
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人(&P)"
         Height          =   180
         Left            =   135
         TabIndex        =   11
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "###病人信息###"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2130
         TabIndex        =   10
         Top             =   255
         Width           =   1260
      End
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   3975
      Left            =   795
      TabIndex        =   7
      Top             =   450
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   7011
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "病人"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "住院号"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "床号"
         Object.Width           =   1111
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "住院医师"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "性别"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "年龄"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "费别"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "护理等级"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   $"frmAdviceCopy.frx":0873
         Object.Width           =   2081
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "出院日期"
         Object.Width           =   2081
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "付款方式"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2295
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceCopy.frx":0880
            Key             =   "Pati"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdviceCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As Object
Private mstrPrivs As String
Private mbln护士站 As Boolean
Private mlng前提ID As Long
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr挂号单 As String
Private mblnMoved As Boolean
Private mblnItem As Boolean
Private mstrIDs As String
Private mstrAlter As String

Private Enum COL成套方案
    col选择 = 0
    col期效 = 1
    col时间 = 2
    col内容 = 3
    col总量 = 4
    col总量单位 = 5
    col单量 = 6
    col单量单位 = 7
    col频次 = 8
    col用法 = 9
    col嘱托 = 10
    col执行时间 = 11
    col执行科室 = 12
    colID = 13
    col相关ID = 14
    col诊疗类别 = 15
    col诊疗项目ID = 16
    col收费细目ID = 17
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal strPrivs As String, _
    lng病人ID As Long, varTime As Variant, blnMoved As Boolean, _
    Optional ByVal bln护士站 As Boolean, Optional ByVal lng前提ID As Long, _
    Optional strAlter As String) As String
'返回：lng病人ID,varTime=要复制医嘱的病人ID，主页ID(挂号单NO)
'      blnMoved=要复制病人的医嘱是否转出
'      strAlter=本次复制的医嘱中要切换期效的医嘱ID(组ID):123,456,...
'      ShowMe=要复制的医嘱的组ID串
    Set mfrmParent = frmParent
    mstrPrivs = strPrivs
    mbln护士站 = bln护士站
    mlng前提ID = lng前提ID
    mlng病人ID = lng病人ID
    If TypeName(varTime) = "String" Then
        mstr挂号单 = varTime
        mlng主页ID = 0
    Else
        mlng主页ID = varTime
        mstr挂号单 = ""
    End If
    mblnMoved = blnMoved
    strAlter = "": mstrAlter = strAlter
    
    Me.Show 1, frmParent
    
    lng病人ID = mlng病人ID
    If TypeName(varTime) = "String" Then
        varTime = mstr挂号单
    Else
        varTime = mlng主页ID
    End If
    blnMoved = mblnMoved
    strAlter = mstrAlter
    ShowMe = mstrIDs
End Function

Private Function LoadPatients() As Boolean
'功能：读取与调用界面相同范围的病人列表
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer
    Dim lng部门ID As Long, intBedLen As Long
        
    On Error GoTo errH
    
    If mlng前提ID <> 0 Then
        cmdPati.Visible = False
        If mstr挂号单 <> "" Then
            strSQL = "Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,A.险类" & _
                " From 病人信息 A,病人挂号记录 B Where A.病人ID=B.病人ID And A.病人ID=[1] And B.NO=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
        Else
            strSQL = _
                "Select A.病人ID,B.主页ID,A.住院号,A.姓名,A.性别,A.年龄," & _
                " B.入院日期,B.出院日期,B.住院医师,B.出院病床 as 床号,B.费别," & _
                " B.险类,B.出院科室ID as 科室ID,B.当前病区ID as 病区ID,C.名称 as 护理等级," & _
                " B.状态,B.数据转出,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式" & _
                " From 病人信息 A,病案主页 B,收费项目目录 C" & _
                " Where A.病人ID=B.病人ID And B.护理等级ID=C.ID(+)" & _
                " And A.病人ID=[1] And B.主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        End If
    Else
        If mstr挂号单 <> "" Then
            '提供当前医生正在就诊的病人清单供选择:因此这点暂不涉及判断和读取"H病人挂号记录"
            strSQL = "Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,A.险类" & _
                " From 病人信息 A,病人挂号记录 B Where A.病人ID=B.病人ID And A.病人ID=[1] And B.NO=[2]"
            strSQL = strSQL & " Union " & _
                " Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,A.险类" & _
                " From 病人信息 A,病人挂号记录 B" & _
                " Where A.病人ID=B.病人ID And B.执行状态=2 And B.执行人||''=[3]" & _
                " Order By NO"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单, UserInfo.姓名)
        Else
            strSQL = "Select 出院科室ID as 科室ID,当前病区ID as 病区ID From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
                
            '提供当前科室/病区的在院病人清单供选择
            lng部门ID = IIF(mbln护士站, Nvl(rsTmp!病区ID, 0), Nvl(rsTmp!科室ID, 0))
            intBedLen = GetMaxBedLen(lng部门ID, Not mbln护士站)
            strSQL = _
                "Select A.病人ID,B.主页ID,A.住院号,A.姓名,A.性别,A.年龄,B.入院日期,B.出院日期," & _
                " B.住院医师,LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,B.险类," & _
                " B.出院科室ID as 科室ID,C.名称 as 护理等级,B.状态,B.数据转出," & _
                " Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式" & _
                " From 病人信息 A,病案主页 B,收费项目目录 C" & _
                " Where A.病人ID=B.病人ID And B.护理等级ID=C.ID(+) And A.病人ID=[1] And B.主页ID=[2]"
            strSQL = strSQL & " Union " & _
                "Select A.病人ID,B.主页ID,A.住院号,A.姓名,A.性别,A.年龄,B.入院日期,B.出院日期," & _
                " B.住院医师,LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,B.险类," & _
                " B.出院科室ID as 科室ID,C.名称 as 护理等级,B.状态,B.数据转出," & _
                " Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式" & _
                " From 病人信息 A,病案主页 B,收费项目目录 C" & _
                " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.护理等级ID=C.ID(+)" & _
                " And (B.出院日期>=Sysdate-30 Or Nvl(B.状态,0)=3 Or B.出院日期 is NULL And Nvl(B.状态,0)<>3)" & _
                IIF(mbln护士站, " And B.当前病区ID=[3]", " And B.出院科室ID=[3]") & _
                IIF(Not mbln护士站 And InStr(mstrPrivs, "本科病人") = 0, " And B.住院医师=[4]", "") & _
                " Order by 床号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, lng部门ID, UserInfo.姓名)
        End If
    End If
    
    lvwPati.ListItems.Clear
    For i = 1 To rsTmp.RecordCount
        If mstr挂号单 <> "" Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID & "_" & rsTmp!NO, rsTmp!姓名, , "Pati")
            objItem.SubItems(1) = Nvl(rsTmp!门诊号)
            objItem.SubItems(2) = Nvl(rsTmp!NO)
            objItem.SubItems(3) = Nvl(rsTmp!性别)
            objItem.SubItems(4) = Nvl(rsTmp!年龄)
            
            '保险病人用红色显示
            If Not IsNull(rsTmp!险类) Then
                Call SetItemColor(objItem, vbRed)
            End If
            
            '显示初始病人的信息
            If rsTmp!病人ID = mlng病人ID And rsTmp!NO = mstr挂号单 Then
                With objItem
                    txtPati.ForeColor = .ForeColor
                    txtPati.Text = .Text
                    lblPati.Caption = "门诊号:" & .SubItems(1) & "　挂号单:" & .SubItems(2) & _
                        "　性别:" & .SubItems(3) & "　年龄:" & .SubItems(4)
                    .Selected = True '一定要选中当前病人
                End With
            End If
        Else
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID & "_" & rsTmp!主页ID, rsTmp!姓名, , "Pati")
            objItem.SubItems(1) = Nvl(rsTmp!住院号)
            objItem.SubItems(2) = Nvl(rsTmp!床号)
            objItem.SubItems(3) = Nvl(rsTmp!住院医师)
            objItem.SubItems(4) = Nvl(rsTmp!性别)
            objItem.SubItems(5) = Nvl(rsTmp!年龄)
            objItem.SubItems(6) = Nvl(rsTmp!费别)
            objItem.SubItems(7) = Nvl(rsTmp!护理等级)
            objItem.SubItems(8) = Format(rsTmp!入院日期, "MM-dd HH:mm")
            objItem.SubItems(9) = Format(Nvl(rsTmp!出院日期), "MM-dd HH:mm")
            objItem.SubItems(10) = Nvl(rsTmp!医疗付款方式)
            objItem.Tag = Nvl(rsTmp!数据转出, 0)
            
            '保险病人用红色显示
            If Not IsNull(rsTmp!险类) Then
                Call SetItemColor(objItem, vbRed)
            End If
            
            '显示初始病人的信息
            If rsTmp!病人ID = mlng病人ID And rsTmp!主页ID = mlng主页ID Then
                With objItem
                    txtPati.ForeColor = .ForeColor
                    txtPati.Text = .Text
                    lblPati.Caption = "住院号:" & .SubItems(1) & "　床号:" & .SubItems(2) & _
                        "　性别:" & .SubItems(4) & "　年龄:" & .SubItems(5) & _
                        "　费别:" & .SubItems(6) & "　付款方式:" & .SubItems(10)
                    .Selected = True '一定要选中当前病人
                End With
            End If
        End If
        rsTmp.MoveNext
    Next
    
    LoadPatients = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetItemColor(ByVal objItem As ListItem, ByVal lngColor As Long)
    Dim i As Long
    
    objItem.ForeColor = lngColor
    For i = 1 To objItem.ListSubItems.Count
        objItem.ListSubItems(i).ForeColor = lngColor
    Next
End Sub

Private Sub cmdAll_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, colID)) <> 0 And RowCanSelect(i) = 0 Then
                .TextMatrix(i, col选择) = -1
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, col选择) = 0
        Next
    End With
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim lng组ID As Long, i As Long
    Dim strIDs As String, strAlter As String
    
    With vsAdvice
        '取一组医嘱的ID
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col选择)) <> 0 Then
                lng组ID = Val(.TextMatrix(i, col相关ID))
                If lng组ID = 0 Then lng组ID = Val(.TextMatrix(i, colID))
                
                '选择复制部份
                If InStr(strIDs & ",", "," & lng组ID & ",") = 0 Then
                    strIDs = strIDs & "," & lng组ID
                End If
                
                '切换期效部份
                If .TextMatrix(i, col期效) <> .Cell(flexcpData, i, col期效) Then
                    If InStr(strAlter & ",", "," & lng组ID & ",") = 0 Then
                        strAlter = strAlter & "," & lng组ID
                    End If
                End If
            End If
        Next
        strAlter = Mid(strAlter, 2)
        strIDs = Mid(strIDs, 2)
        If strIDs = "" Then
            MsgBox "请选择要复制的医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    mstrAlter = strAlter
    mstrIDs = strIDs
    Unload Me
End Sub

Private Sub cmdPati_Click()
    If mstr挂号单 <> "" Then
        lvwPati.ListItems("_" & mlng病人ID & "_" & mstr挂号单).Selected = True
    Else
        lvwPati.ListItems("_" & mlng病人ID & "_" & mlng主页ID).Selected = True
    End If
    lvwPati.SelectedItem.EnsureVisible
    lvwPati.Left = txtPati.Left + fraPati.Left
    lvwPati.Top = txtPati.Top + txtPati.Height + fraPati.Top
    lvwPati.Height = vsAdvice.Height - 300
    lvwPati.ZOrder
    lvwPati.Visible = True
    lvwPati.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    ElseIf KeyCode = vbKeyEscape Then
        If lvwPati.Visible Then
            lvwPati.Visible = False
        Else
            Unload Me
        End If
    ElseIf KeyCode = vbKeyF4 Or KeyCode = vbKeyDown Then
        If Not (KeyCode = vbKeyDown And Shift <> vbAltMask) Then
            If Me.ActiveControl Is txtPati Then
                If cmdPati.Visible And cmdPati.Enabled Then cmdPati_Click
            End If
        End If
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Dim strLvw As String
    If mstr挂号单 <> "" Then
        strLvw = "病人,1000,0,1;门诊号,1000,0,1;挂号单,1000,0,1;性别,600,0,1;年龄,600,0,1"
    Else
        strLvw = "病人,1000,0,1;住院号,1000,0,1;床号,630,0,1;住院医师,1000,0,1;性别,600,0,1;年龄,600,0,1;费别,850,0,1;护理等级,1150,0,1;入院日期,1180,0,1;出院日期,1180,0,1;付款方式,1500,0,1"
    End If
    Call zlControl.LvwSelectColumns(lvwPati, strLvw, True)
    Call RestoreWinState(Me, App.ProductName, IIF(mstr挂号单 <> "", 1, 2))
    If mlng主页ID <> 0 Then
        vsAdvice.FrozenCols = col期效 + 1
    Else
        vsAdvice.FrozenCols = col选择 + 1
    End If
    
    Call LoadPatients
    Call LoadAdvice
    
    mstrIDs = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraPati.Top = -75
    fraPati.Left = 0
    fraPati.Width = Me.ScaleWidth
    
    vsAdvice.Left = 0
    vsAdvice.Top = fraPati.Top + fraPati.Height
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - vsAdvice.Top - cmdOK.Height * 1.6
        
    cmdHelp.Top = Me.ScaleHeight - cmdAll.Height * 1.3
    cmdAll.Top = cmdHelp.Top
    cmdClear.Top = cmdAll.Top
    cmdOK.Top = cmdAll.Top
    cmdCancel.Top = cmdAll.Top
    
    If Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3) < 5000 Then
        cmdCancel.Left = 5000
    Else
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3)
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, IIF(mstr挂号单 <> "", 1, 2))
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwPati_DblClick()
    If mblnItem Then Call lvwPati_KeyPress(13)
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Dim lng病人ID As Long, lng主页ID As Long, strNO As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not lvwPati.SelectedItem Is Nothing Then
            lng病人ID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(0))
            If mstr挂号单 <> "" Then
                strNO = Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1)
                If lng病人ID = mlng病人ID And strNO = mstr挂号单 Then
                    lvwPati.Visible = False
                    vsAdvice.SetFocus: Exit Sub
                End If
                With lvwPati.SelectedItem
                    mlng病人ID = lng病人ID
                    mstr挂号单 = strNO
                    mblnMoved = MovedByNO(strNO, "病人挂号记录")
                    
                    txtPati.Text = .Text
                    txtPati.ForeColor = .ForeColor
                    lblPati.Caption = "门诊号:" & .SubItems(1) & "　挂号单:" & .SubItems(2) & _
                        "　性别:" & .SubItems(3) & "　年龄:" & .SubItems(4)
                End With
            Else
                lng主页ID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1))
                If lng病人ID = mlng病人ID And lng主页ID = mlng主页ID Then
                    lvwPati.Visible = False
                    vsAdvice.SetFocus: Exit Sub
                End If
                With lvwPati.SelectedItem
                    mlng病人ID = lng病人ID
                    mlng主页ID = lng主页ID
                    mblnMoved = Val(.Tag) = 1
                    
                    txtPati.Text = .Text
                    txtPati.ForeColor = .ForeColor
                    lblPati.Caption = "住院号:" & .SubItems(1) & "　床号:" & .SubItems(2) & _
                        "　性别:" & .SubItems(4) & "　年龄:" & .SubItems(5) & _
                        "　费别:" & .SubItems(6) & "  付款方式:" & .SubItems(10)
                End With
            End If
            lvwPati.Visible = False
            
            '读取并显示病人医嘱
            Call LoadAdvice
            
            vsAdvice.SetFocus
        End If
    End If
End Sub

Private Sub lvwPati_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnItem = False
End Sub

Private Sub lvwPati_Validate(Cancel As Boolean)
    lvwPati.Visible = False
End Sub

Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub RowSelectSame(ByVal lngRow As Long)
'功能：根据指定行(可能为任意行)的选择状态,将相关医嘱一并选择
    Dim i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, col相关ID)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) _
                    Or Val(.TextMatrix(i, colID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                    .TextMatrix(i, col期效) = .TextMatrix(lngRow, col期效)
                    .Cell(flexcpFontBold, i, col期效) = .Cell(flexcpFontBold, lngRow, col期效)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) _
                    Or Val(.TextMatrix(i, colID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                    .TextMatrix(i, col期效) = .TextMatrix(lngRow, col期效)
                    .Cell(flexcpFontBold, i, col期效) = .Cell(flexcpFontBold, lngRow, col期效)
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, colID)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                    .TextMatrix(i, col期效) = .TextMatrix(lngRow, col期效)
                    .Cell(flexcpFontBold, i, col期效) = .Cell(flexcpFontBold, lngRow, col期效)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, colID)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                    .TextMatrix(i, col期效) = .TextMatrix(lngRow, col期效)
                    .Cell(flexcpFontBold, i, col期效) = .Cell(flexcpFontBold, lngRow, col期效)
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Function LoadAdvice() As Boolean
'功能：读取当前病人指定的医嘱
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    On Error GoTo errH
    
    '排开撤档和不服务于的内容
    strSQL = "Select Distinct A.ID,A.序号,A.相关ID,A.医嘱期效,A.开始执行时间,A.诊疗项目ID," & _
        " A.医嘱内容,A.单次用量,A.执行频次,A.医生嘱托,C.名称 as 执行科室,A.执行时间方案,A.收费细目ID," & _
        " A.标本部位,B.类别,B.名称,B.计算单位,A.总给予量 as 总量,E.门诊包装,E.门诊单位,E.住院包装,E.住院单位," & _
        " B.撤档时间,B.服务对象,D.撤档时间 as 收费撤档,D.服务对象 as 收费服务" & _
        " From 病人医嘱记录 A,诊疗项目目录 B,部门表 C,收费项目目录 D,药品规格 E" & _
        " Where A.诊疗项目ID=B.ID(+) And A.执行科室ID=C.ID(+) And A.收费细目ID=D.ID(+) And A.收费细目ID=E.药品ID(+)" & _
        " And A.医嘱状态 Not IN(2,4) And A.开始执行时间 is Not Null And A.病人来源<>3" & _
        IIF(mstr挂号单 <> "", " And A.病人ID+0=[1] And A.挂号单=[2]", " And A.病人ID=[1] And A.主页ID=[2]") & _
        " Order by A.序号"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, IIF(mstr挂号单 <> "", mstr挂号单, mlng主页ID))
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '清除表格内容
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col选择) = 0
                .TextMatrix(i, colID) = rsTmp!ID
                .TextMatrix(i, col相关ID) = Nvl(rsTmp!相关ID)
                .TextMatrix(i, col诊疗类别) = Nvl(rsTmp!类别, "*")
                .TextMatrix(i, col诊疗项目ID) = Nvl(rsTmp!诊疗项目ID)
                .TextMatrix(i, col收费细目ID) = Nvl(rsTmp!收费细目ID)
                .TextMatrix(i, col期效) = IIF(Nvl(rsTmp!医嘱期效, 0) = 0, "长嘱", "临嘱")
                .Cell(flexcpData, i, col期效) = .TextMatrix(i, col期效)
                .TextMatrix(i, col时间) = Format(rsTmp!开始执行时间, "MM-dd HH:mm")
                .TextMatrix(i, col内容) = rsTmp!医嘱内容
                .Cell(flexcpData, i, col内容) = Nvl(rsTmp!标本部位) '检验标本
                .TextMatrix(i, col单量) = FormatEx(Nvl(rsTmp!单次用量), 4)
                If Not IsNull(rsTmp!单次用量) Then
                    .TextMatrix(i, col单量单位) = Nvl(rsTmp!计算单位)
                End If
                If InStr(",5,6,", Nvl(rsTmp!类别, "*")) > 0 Then
                    If mstr挂号单 <> "" Then
                        If Not IsNull(rsTmp!总量) And Not IsNull(rsTmp!门诊包装) Then
                            .TextMatrix(i, col总量) = FormatEx(rsTmp!总量 / rsTmp!门诊包装, 5)
                        End If
                        If Nvl(rsTmp!医嘱期效, 0) = 1 Then
                            .TextMatrix(i, col总量单位) = Nvl(rsTmp!门诊单位)
                        End If
                    Else
                        If Not IsNull(rsTmp!总量) And Not IsNull(rsTmp!住院包装) Then
                            .TextMatrix(i, col总量) = FormatEx(rsTmp!总量 / rsTmp!住院包装, 5)
                        End If
                        If Nvl(rsTmp!医嘱期效, 0) = 1 Then
                            .TextMatrix(i, col总量单位) = Nvl(rsTmp!住院单位)
                        End If
                    End If
                Else
                    If Not IsNull(rsTmp!总量) Then
                        .TextMatrix(i, col总量) = rsTmp!总量
                    End If
                    If Nvl(rsTmp!医嘱期效, 0) = 1 Then
                        .TextMatrix(i, col总量单位) = Nvl(rsTmp!计算单位)
                    End If
                End If
                
                .TextMatrix(i, col频次) = Nvl(rsTmp!执行频次)
                .TextMatrix(i, col嘱托) = Nvl(rsTmp!医生嘱托)
                .TextMatrix(i, col执行时间) = Nvl(rsTmp!执行时间方案)
                .TextMatrix(i, col执行科室) = Nvl(rsTmp!执行科室)
                
                '处理行隐藏及用法显示
                If InStr(",C,D,F,G,E,", Nvl(rsTmp!类别, "*")) > 0 And Not IsNull(rsTmp!相关ID) Then
                    .RowHidden(i) = True
                ElseIf Nvl(rsTmp!类别) = "7" Then
                    .RowHidden(i) = True
                ElseIf Nvl(rsTmp!类别) = "E" And IsNull(rsTmp!相关ID) _
                    And Val(.TextMatrix(i - 1, col相关ID)) = rsTmp!ID _
                    And InStr(",5,6,", .TextMatrix(i - 1, col诊疗类别)) > 0 Then
                    '给药途径
                    .RowHidden(i) = True
                    '显示给药途径
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关ID)) = rsTmp!ID Then
                            .TextMatrix(j, col用法) = rsTmp!名称
                        Else
                            Exit For
                        End If
                    Next
                ElseIf Nvl(rsTmp!类别) = "E" And IsNull(rsTmp!相关ID) _
                    And Val(.TextMatrix(i - 1, col相关ID)) = rsTmp!ID _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col诊疗类别)) > 0 Then
                    '中药用法或检验采集方法
                    .TextMatrix(i, col用法) = rsTmp!名称
                    
                    '中药或检验的执行科室
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关ID)) = rsTmp!ID Then
                            If InStr(",7,C,", .TextMatrix(j, col诊疗类别)) > 0 Then
                                .TextMatrix(i, col执行科室) = .TextMatrix(j, col执行科室)
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    '中药付数
                    If .TextMatrix(i - 1, col诊疗类别) <> "C" Then
                        .TextMatrix(i, col总量单位) = "付"
                    End If
                End If
                
                '标记包含得有撤档或不服务的项目
                If Not IsNull(rsTmp!诊疗项目ID) Then
                    If Not (IsNull(rsTmp!撤档时间) Or Format(Nvl(rsTmp!撤档时间), "yyyy-MM-dd") = "3000-01-01") Then
                        .RowData(i) = 1
                    ElseIf Not (Nvl(rsTmp!服务对象, 0) = 3 Or Nvl(rsTmp!服务对象, 0) = IIF(mstr挂号单 <> "", 1, 2)) Then
                        .RowData(i) = 1
                    ElseIf Not IsNull(rsTmp!收费细目ID) Then
                        '对药品,同时要判断到收费项目目录
                        If Not (IsNull(rsTmp!收费撤档) Or Format(Nvl(rsTmp!收费撤档), "yyyy-MM-dd") = "3000-01-01") Then
                            .RowData(i) = 1
                        ElseIf Not (Nvl(rsTmp!收费服务, 0) = 3 Or Nvl(rsTmp!收费服务, 0) = IIF(mstr挂号单 <> "", 1, 2)) Then
                            .RowData(i) = 1
                        End If
                    End If
                End If

                rsTmp.MoveNext
            Next
        End If
        If mlng主页ID <> 0 Then
            .Cell(flexcpBackColor, .FixedRows, col选择, .Rows - 1, col期效) = &HC0FFC0
        End If
        
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col内容
        .ColHidden(col期效) = mstr挂号单 <> ""
        .Redraw = flexRDDirect
    End With
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function RowCanSelect(ByVal lngRow As Long) As Long
'功能：判断指定行的(相关)医嘱可否选择
'返回：如果可以选择，返回0,否则返回行号
    Dim i As Long
    
    With vsAdvice
        If .RowData(lngRow) = 1 Then RowCanSelect = lngRow: Exit Function
        
        If Val(.TextMatrix(lngRow, col相关ID)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) _
                    Or Val(.TextMatrix(i, colID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) _
                    Or Val(.TextMatrix(i, colID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, colID)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, colID)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Function

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col选择 Then Call RowSelectSame(Row)
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col内容 Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col选择 Then Cancel = True
End Sub

Private Sub vsAdvice_DblClick()
    Call vsAdvice_KeyPress(32)
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '擦除一并给药相关行列的边线及内容
        lngLeft = col期效: lngRight = col时间
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col频次: lngRight = col用法
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            If Col = lngLeft And lngLeft = col期效 Then
                SetBkColor hDC, SysColor2RGB(.Cell(flexcpBackColor, Row, lngLeft))
            Else
                SetBkColor hDC, SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsAdvice
        If KeyAscii = 32 Then
            If .Col <> col选择 Then
                KeyAscii = 0
                If Val(.TextMatrix(.Row, colID)) = 0 Then Exit Sub
                
                '检查是否可以被选择
                i = RowCanSelect(.Row)
                If i > 0 And Val(.TextMatrix(.Row, col选择)) = 0 Then
                    MsgBox "因为医嘱""" & .TextMatrix(i, col内容) & """对应的项目已撤档或服务对象不匹配，该医嘱不能被选择。", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                If .Col = col期效 And mlng主页ID <> 0 Then
                    If CanAlterType(.Row) Then
                        .TextMatrix(.Row, .Col) = IIF(.TextMatrix(.Row, .Col) = "长嘱", "临嘱", "长嘱")
                        .Cell(flexcpFontBold, .Row, .Col) = .TextMatrix(.Row, .Col) <> .Cell(flexcpData, .Row, .Col)
                        If .Cell(flexcpFontBold, .Row, .Col) Then
                            .TextMatrix(.Row, col选择) = -1
                        End If
                        Call RowSelectSame(.Row)
                    End If
                Else
                    .TextMatrix(.Row, col选择) = IIF(Val(.TextMatrix(.Row, col选择)) = 0, -1, 0)
                    Call RowSelectSame(.Row)
                End If
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    If Col <> col选择 Then
        Cancel = True
    ElseIf Val(vsAdvice.TextMatrix(vsAdvice.Row, colID)) = 0 Then
        Cancel = True
    Else
        i = RowCanSelect(Row)
        If i > 0 Then
            Cancel = True
            MsgBox "因为医嘱""" & vsAdvice.TextMatrix(i, col内容) & """对应的项目已撤档或服务对象不匹配，该医嘱不能被选择。", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Function CanAlterType(ByVal lngRow As Long) As Boolean
'功能：判断指定的医嘱是否可以切换期效
'参数：lngRow=可见的医嘱行
'说明：允许切换期效的条件：
'   1.成长嘱：执行频率=0(可选频率),2(持续性)
'   2.成临嘱：执行频率=0(可选频率),1(一次性);药品必须指定了规格
    Dim rsMore As New ADODB.Recordset
    Dim strSQL As String, strType As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, colID)) = 0 Then
            CanAlterType = True: Exit Function
        ElseIf Val(.TextMatrix(lngRow, col诊疗项目ID)) = 0 Then
            '自由输入的可以切换
            CanAlterType = True: Exit Function
        ElseIf RowIn配方行(lngRow) Then
            '中药配方固定可以切换
            CanAlterType = True: Exit Function
        ElseIf RowIn检验行(lngRow) Then
            '检验以检验行为准判断
            lngRow = .FindRow(.TextMatrix(lngRow, colID), , col相关ID)
            If lngRow = -1 Then Exit Function
        End If
    
        strType = IIF(.TextMatrix(lngRow, col期效) = "长嘱", "临嘱", "长嘱")
        
        '以原始频率为准判断:因为可选择频率的可能已缺成一次性
        strSQL = "Select 执行频率 From 诊疗项目目录 Where ID=[1]"
        Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, col诊疗项目ID)))
        
        If strType = "长嘱" Then
            If InStr(",0,2,", Nvl(rsMore!执行频率, 0)) = 0 Then Exit Function
        Else
            If InStr(",0,1,", Nvl(rsMore!执行频率, 0)) = 0 Then Exit Function
            If InStr(",5,6,", .TextMatrix(lngRow, col诊疗类别)) > 0 Then
                Call GetRowScope(lngRow, lngBegin, lngEnd)
                For i = lngBegin To lngEnd
                    If InStr(",5,6,", .TextMatrix(i, col诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, col收费细目ID)) = 0 Then Exit Function
                    End If
                Next
            End If
        End If
    End With
    CanAlterType = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
'功能：获取组ID相同的一组医嘱行号范围(注意考虑一并给药中的空行)
    Dim lngS组ID As Long, lngO组ID As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS组ID = IIF(Val(.TextMatrix(lngRow, col相关ID)) = 0, Val(.TextMatrix(lngRow, colID)), Val(.TextMatrix(lngRow, col相关ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            lngO组ID = IIF(Val(.TextMatrix(i, col相关ID)) = 0, Val(.TextMatrix(i, colID)), Val(.TextMatrix(i, col相关ID)))
            If Not (Val(.TextMatrix(i, colID)) = 0 And i >= .FixedRows) Then '跳过空行
                If lngO组ID = lngS组ID Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            lngO组ID = IIF(Val(.TextMatrix(i, col相关ID)) = 0, Val(.TextMatrix(i, colID)), Val(.TextMatrix(i, col相关ID)))
            If Not (Val(.TextMatrix(i, colID)) = 0 And i >= .FixedRows) Then '跳过空行
                If lngO组ID = lngS组ID Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
End Sub

Private Function RowIn检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于检验组合中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.TextMatrix(lngRow, colID) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "E" And Val(.TextMatrix(lngRow, col相关ID)) = 0 Then
            '采集方法行
            If .TextMatrix(lngRow - 1, col诊疗类别) = "C" _
                And Val(.TextMatrix(lngRow - 1, col相关ID)) = .TextMatrix(lngRow, colID) Then
                RowIn检验行 = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, col诊疗类别) = "C" And Val(.TextMatrix(lngRow, col相关ID)) <> 0 Then
            '检验项目行
            RowIn检验行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于中药配方中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.TextMatrix(lngRow, colID) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "E" Then
            If Val(.TextMatrix(lngRow, col相关ID)) = 0 Then
                '用法行
                If Val(.TextMatrix(lngRow - 1, col相关ID)) = .TextMatrix(lngRow, colID) _
                    And .TextMatrix(lngRow - 1, col诊疗类别) = "E" Then
                    RowIn配方行 = True: Exit Function
                End If
            Else
                '煎法行
                If .TextMatrix(lngRow - 1, col诊疗类别) = "7" _
                    And Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    RowIn配方行 = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, col诊疗类别) = "7" And Val(.TextMatrix(lngRow, col相关ID)) <> 0 Then
            '中药行
            RowIn配方行 = True: Exit Function
        End If
    End With
End Function
