VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiffPriceAdjustCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自动调差设置"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4920
      TabIndex        =   16
      Top             =   4200
      Width           =   1100
   End
   Begin VB.Frame fraRangeSelect 
      Caption         =   " 内容 "
      Height          =   5190
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4680
      Begin VB.CheckBox Chk剂型 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1080
         Width           =   675
      End
      Begin MSComCtl2.UpDown updRate 
         Height          =   300
         Left            =   3225
         TabIndex        =   11
         Top             =   3825
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtRate"
         BuddyDispid     =   196613
         OrigLeft        =   3720
         OrigTop         =   4200
         OrigRight       =   3960
         OrigBottom      =   4575
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtRate 
         Height          =   300
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "10"
         Top             =   3825
         Width           =   1935
      End
      Begin VB.CommandButton Cmd用途 
         Caption         =   "…"
         Height          =   300
         Left            =   3885
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   675
         Width           =   285
      End
      Begin VB.TextBox Txt用途 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   4
         Top             =   675
         Width           =   2775
      End
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   3045
      End
      Begin MSComctlLib.ListView Lvw剂型 
         Height          =   2430
         Left            =   345
         TabIndex        =   8
         Top             =   1290
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   4286
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Lbl剂型 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "剂型(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   405
         TabIndex        =   6
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "  说明：实际差价与实际金额之比大于或小于指导差价率的百分点为调差波动率的那些药品才出来。"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   330
         TabIndex        =   13
         Top             =   4245
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3570
         TabIndex        =   12
         Top             =   3885
         Width           =   255
      End
      Begin VB.Label Lbl盘点方式 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "调差波动率"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   345
         TabIndex        =   9
         Top             =   3885
         Width           =   900
      End
      Begin VB.Label Lbl用途分类 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用途分类"
         Height          =   180
         Left            =   315
         TabIndex        =   3
         Top             =   735
         Width           =   720
      End
      Begin VB.Label lbl库房 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房"
         Height          =   180
         Left            =   675
         TabIndex        =   1
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4920
      TabIndex        =   15
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4920
      TabIndex        =   14
      Top             =   285
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5280
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCondition.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCondition.frx":0E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCondition.frx":2B5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tvw用途分类 
      Height          =   2700
      Left            =   510
      TabIndex        =   17
      Top             =   1095
      Visible         =   0   'False
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   4763
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmDiffPriceAdjustCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mblnBootUp As Boolean
Private mblnFirstUp As Boolean

Private mstr用途ID As String
Private mstr剂型 As String
Private mlng库房ID As Long
Private mintRate As Integer
Private mbln中药库房 As Boolean
Private mfrmMain As Form

Public Function GetCondition(FrmMain As Form, ByRef str用途ID, ByRef str剂型 As String, _
    ByRef lng库房ID As Long, ByRef int波动率 As Integer) As Boolean
    
    mstr用途ID = ""
    mstr剂型 = ""
    mlng库房ID = 0
    mintRate = int波动率
    mblnSelect = False
    
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    str用途ID = mstr用途ID
    str剂型 = mstr剂型
    lng库房ID = mlng库房ID
    int波动率 = mintRate
End Function

Private Sub cbo库房_Click()
    Dim rsTemp As New ADODB.Recordset
    '提取该库房现有剂型，供用户选择
    mbln中药库房 = False
    On Error GoTo errHandle
    
    gstrSQL = "Select 1 From 部门性质说明 " & _
             " Where 工作性质 Like '中药%' And 部门ID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查部门性质]", Me.cbo库房.ItemData(cbo库房.ListIndex))
             
    If Not rsTemp.EOF Then mbln中药库房 = True
    
    gstrSQL = "Select Distinct J.编码,J.名称 " & _
             " From 诊疗执行科室 A,药品特性 B,药品剂型 J " & _
             " Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称 " & _
             "     And A.执行科室ID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", Me.cbo库房.ItemData(cbo库房.ListIndex))
    
    lvw剂型.ListItems.Clear
    
    With rsTemp
        Do While Not .EOF
            lvw剂型.ListItems.Add , "K" & !编码, !名称, , 1
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Chk材质_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnSelect = False
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intItem As Integer, intItems As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    If Tvw用途分类.SelectedItem.Key <> "R" Then
        Select Case Tvw用途分类.SelectedItem.Key
            Case "R_中成药", "R_中草药", "R_西成药"
                mstr用途ID = "'" & Tvw用途分类.SelectedItem & "'"
            Case Else
                strsql = "Select ID From 诊疗分类目录 " & _
                         "Start With ID=[1] " & _
                         "Connect by Prior ID=上级ID "
                Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption & "-诊疗分类ID", Mid(Tvw用途分类.SelectedItem.Key, 3))
                With rsTmp
                    mstr用途ID = ""
                    Do While Not .EOF
                        mstr用途ID = mstr用途ID & !id & ","
                        .MoveNext
                    Loop
                End With
                mstr用途ID = Mid(mstr用途ID, 1, Len(mstr用途ID) - 1)
        End Select
    End If
    
    mstr剂型 = ""
    intItems = Me.lvw剂型.ListItems.count
    For intItem = 1 To intItems
        If lvw剂型.ListItems(intItem).Checked Then
            mstr剂型 = mstr剂型 & "," & lvw剂型.ListItems(intItem).Text
        End If
    Next
    If mbln中药库房 Then mstr剂型 = mstr剂型 & "," & "方剂"
    If mstr剂型 <> "" Then mstr剂型 = Mid(mstr剂型, 2)
    
    mlng库房ID = cbo库房.ItemData(cbo库房.ListIndex)
    mintRate = Val(txtRate.Text)
    
    mblnSelect = True
    frmDiffPriceAdjustCard.txtStock.Caption = cbo库房.Text
    frmDiffPriceAdjustCard.txtStock.Tag = mlng库房ID
    
    frmDiffPriceAdjustCard.CmdSave.Enabled = False
    frmDiffPriceAdjustCard.cmdCancel.Enabled = False
    
    Hide
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd用途_Click()
    '把药品用途分类装入TREEVIEW
    Tvw用途分类.Visible = Tvw用途分类.Visible Xor True
    If Tvw用途分类.Visible Then
        Tvw用途分类.Top = Txt用途.Top + Txt用途.Height + fraRangeSelect.Top
        Tvw用途分类.Left = Txt用途.Left + fraRangeSelect.Left
        Tvw用途分类.ZOrder 0
        Tvw用途分类.SetFocus
    End If
End Sub


Private Sub Command1_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Click()
    If Tvw用途分类.Visible = True Then
        Tvw用途分类.Visible = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rs用途分类 As New Recordset
    Dim rs剂型 As New Recordset
    Dim rs材质分类 As New Recordset
    Dim Str材质 As String
    
    Dim blnSelectStock As String
    On Error GoTo errHandle
    blnSelectStock = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & mfrmMain.Caption, "库房", "0")
    
    mblnBootUp = False
    mblnFirstUp = True
    
    With mfrmMain.cboStock
        cbo库房.Clear
        For i = 0 To .ListCount - 1
            cbo库房.AddItem .List(i)
            cbo库房.ItemData(cbo库房.NewIndex) = .ItemData(i)
        Next
        cbo库房.ListIndex = .ListIndex
    End With
        
    If zlStr.IsHavePrivs(gstrprivs, "所有库房") Then
        If blnSelectStock = "0" Then
            cbo库房.Enabled = False
        Else
            cbo库房.Enabled = True
        End If
    Else
        cbo库房.Enabled = False
    End If
        
    '药品材质权限控制
    Str材质 = ""
    If UserInfo.strMaterial <> "" Then      '为空表示有所有库房权限
        If InStr(1, UserInfo.strMaterial, "中成药") <> 0 Then Str材质 = Str材质 & IIf(Str材质 = "", "", ",") & "2"
        If InStr(1, UserInfo.strMaterial, "西成药") <> 0 Then Str材质 = Str材质 & IIf(Str材质 = "", "", ",") & "1"
        If InStr(1, UserInfo.strMaterial, "中草药") <> 0 Then Str材质 = Str材质 & IIf(Str材质 = "", "", ",") & "3"
        If Str材质 = "" Then
            MsgBox "对不起，你一个材质分类权限都没有，请与系统管理员联系！", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        Str材质 = "1,2,3"
    End If
    
    gstrSQL = " SELECT a.ID,a.上级ID,a.名称,1 AS 末级,DECODE(a.类型,1,'西成药',2,'中成药','中草药') AS 材质  " & _
              " FROM 诊疗分类目录 a " & _
              IIf(Str材质 = "", "", " WHERE a.类型 in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList))) ") & _
              " START WITH a.上级ID IS NULL CONNECT BY PRIOR a.ID =a.上级ID ORDER BY LEVEL,a.ID "
    Set rs用途分类 = zlDataBase.OpenSQLRecord(gstrSQL, "读取用途分类", Str材质)
    
    With rs用途分类
        If .EOF Then
            MsgBox "药品用途分类不完整！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        With rs用途分类
            Tvw用途分类.Nodes.Clear
            Tvw用途分类.Nodes.Add , , "R", "所有用途分类", 1, 1
            Txt用途.Text = "所有用途分类"
            
            
            gstrSQL = "Select 名称 From 诊疗项目类别 Where 编码 IN ('5','6','7')"
            Set rs材质分类 = zlDataBase.OpenSQLRecord(gstrSQL, "Form_Load")
            
            With rs材质分类
                Do While Not .EOF
                    Tvw用途分类.Nodes.Add "R", tvwChild, "R_" & !名称, !名称, 2, 2
                    .MoveNext
                Loop
                .Close
            End With
            
            .MoveFirst
            Do While Not .EOF
                If IsNull(!上级ID) Then
                    If !末级 = 1 Then
                        Tvw用途分类.Nodes.Add "R_" & !材质, tvwChild, "K_" & !id, !名称, 3, 3
                    Else
                        Tvw用途分类.Nodes.Add "R_" & !材质, tvwChild, "K_" & !id, !名称, 2, 2
                    End If
                Else
                    If !末级 = 1 Then
                        Tvw用途分类.Nodes.Add "K_" & !上级ID, tvwChild, "K_" & !id, !名称, 3, 3
                    Else
                        Tvw用途分类.Nodes.Add "K_" & !上级ID, tvwChild, "K_" & !id, !名称, 2, 2
                    End If
                End If
                Tvw用途分类.Nodes("K_" & !id).Tag = !末级
                .MoveNext
            Loop
        End With
    
        Tvw用途分类.Nodes("R").Selected = True
        Tvw用途分类.Nodes("R").Expanded = True
    End With
    
    mblnBootUp = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Select Case UnloadMode
        Case vbFormControlMenu, vbAppWindows, vbAppTaskManager, vbFormOwner
            Me.Hide
        Case vbFormCode
            If Tvw用途分类.Visible Then
                Tvw用途分类.Visible = False
                Cmd用途.SetFocus
                Cancel = 1
                Exit Sub
            End If
    End Select
End Sub

Private Sub fraRangeSelect_Click()
    If Tvw用途分类.Visible = True Then
        Tvw用途分类.Visible = False
    End If
End Sub

Private Sub lst剂型_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    
End Sub

Private Sub Tvw用途分类_DblClick()
    Me.Txt用途.Text = Tvw用途分类.SelectedItem.Text
    Tvw用途分类.Visible = False
    lvw剂型.SetFocus
End Sub

Private Sub Tvw用途分类_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Tvw用途分类_DblClick
    End If
End Sub

Private Sub Tvw用途分类_LostFocus()
    Tvw用途分类.Visible = False
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyAdd
            If Val(txtRate.Text) < 100 Then
                txtRate.Text = Val(txtRate.Text) + 1
            End If
        Case vbKeySubtract
            If Val(txtRate.Text) > 1 Then
                txtRate.Text = Val(txtRate.Text) - 1
            End If
    End Select
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
        
        Case 48 To 57
            If IsNumeric(txtRate.Text) Then
                If txtRate.SelLength <> Len(txtRate.Text) Then
                    If Val(txtRate.Text & Chr(KeyAscii)) > 100 Then
                        KeyAscii = 0
                    End If
                End If
            End If
        Case 8          '退格键
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtRate_Validate(Cancel As Boolean)
    If Trim(txtRate.Text) = "" Or Trim(txtRate.Text) = "0" Then
        Cancel = True
    End If
End Sub

Private Sub Chk剂型_Click()
    If Chk剂型.Value = 2 Then Exit Sub
    Call SetSelect(lvw剂型, Chk剂型.Value)
End Sub

Private Sub Lvw剂型_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ItemCheck(lvw剂型, Item)
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem)
    Dim lngCheck As Long, blnCheck As Boolean, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.count Then
            Chk剂型.Value = 1
        ElseIf intCount > 0 Then
            Chk剂型.Value = 2
        Else
            Chk剂型.Value = 0
        End If
    End With
End Sub
