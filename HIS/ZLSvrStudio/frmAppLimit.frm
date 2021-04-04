VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppLimit 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   16545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAppLimit.frx":0000
   ScaleHeight     =   10590
   ScaleWidth      =   16545
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   0
      ScaleHeight     =   9615
      ScaleWidth      =   15735
      TabIndex        =   10
      Top             =   600
      Width           =   15735
      Begin MSComctlLib.ImageList img16 
         Left            =   5280
         Top             =   0
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
               Picture         =   "frmAppLimit.frx":803A
               Key             =   "unCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppLimit.frx":85D4
               Key             =   "Check"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pctOpt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3495
         ScaleWidth      =   14175
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   6000
         Width           =   14175
         Begin VB.TextBox txtTip 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2775
            Left            =   1680
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "frmAppLimit.frx":8B6E
            Top             =   480
            Width           =   5535
         End
         Begin VB.PictureBox pctOption 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   8160
            ScaleHeight     =   2385
            ScaleWidth      =   5505
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   360
            Width           =   5535
            Begin VB.CommandButton cmdCancel 
               Caption         =   "取消(&C)"
               Height          =   350
               Left            =   3900
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   1920
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.ComboBox txtApp 
               Height          =   300
               Left            =   1020
               TabIndex        =   2
               Top             =   630
               Width           =   2775
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   2640
               MaxLength       =   3
               TabIndex        =   6
               Tag             =   "IP地址"
               Top             =   1125
               Width           =   315
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   0
               Left            =   1080
               MaxLength       =   3
               TabIndex        =   3
               Tag             =   "IP地址"
               Top             =   1125
               Width           =   315
            End
            Begin VB.TextBox txtIP 
               Height          =   300
               Index           =   2
               Left            =   3420
               MaxLength       =   3
               TabIndex        =   7
               Tag             =   "IP"
               Top             =   1095
               Width           =   390
            End
            Begin VB.TextBox txtUser 
               Height          =   350
               Left            =   1020
               TabIndex        =   1
               Top             =   120
               Width           =   2415
            End
            Begin VB.CommandButton cmdStop 
               Caption         =   "停用(&S)"
               Height          =   350
               Left            =   3900
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   1560
               Width           =   1455
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "修改限制(&M)"
               Height          =   350
               Left            =   3900
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   600
               Width           =   1455
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "添加限制(&A)"
               Height          =   350
               Left            =   3900
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   120
               Width           =   1455
            End
            Begin VB.CommandButton cmdDel 
               Caption         =   "删除限制(&D)"
               Height          =   350
               Left            =   3900
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1455
            End
            Begin VB.CommandButton cmdMore 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   6.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   3420
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   2160
               MaxLength       =   3
               TabIndex        =   5
               Tag             =   "IP地址"
               Top             =   1125
               Width           =   315
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   4
               Tag             =   "IP地址"
               Top             =   1125
               Width           =   315
            End
            Begin VB.TextBox txtDesc 
               Height          =   495
               Left            =   1020
               MaxLength       =   99
               MultiLine       =   -1  'True
               TabIndex        =   8
               Top             =   1560
               Width           =   2775
            End
            Begin VB.TextBox txtIP 
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   1020
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   18
               TabStop         =   0   'False
               Tag             =   "IP"
               Text            =   "    ．    ．    ．"
               Top             =   1080
               Width           =   1965
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "确定(&O)"
               Height          =   350
               Left            =   3900
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   1560
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label lblIP 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "IP段"
               Height          =   180
               Left            =   480
               TabIndex        =   27
               Top             =   1170
               Width           =   360
            End
            Begin VB.Label lblUser 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "用户名"
               Height          =   180
               Left            =   300
               TabIndex        =   26
               Top             =   210
               Width           =   540
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "-"
               Height          =   180
               Index           =   11
               Left            =   3180
               TabIndex        =   25
               Top             =   1155
               Width           =   90
            End
            Begin VB.Label lblDesc 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "说明"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   480
               TabIndex        =   24
               Top             =   1560
               Width           =   360
            End
            Begin VB.Label lblApp 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "应用程序"
               Height          =   180
               Left            =   120
               TabIndex        =   23
               Top             =   690
               Width           =   720
            End
         End
         Begin VB.Image imgIcon 
            Appearance      =   0  'Flat
            Height          =   1155
            Left            =   480
            Picture         =   "frmAppLimit.frx":8C24
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1125
         End
      End
      Begin VB.PictureBox pctPer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   0
         ScaleHeight     =   4695
         ScaleWidth      =   15135
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   15135
         Begin VB.TextBox txtFind 
            ForeColor       =   &H80000010&
            Height          =   350
            Left            =   960
            TabIndex        =   9
            Text            =   "输入用户名、姓名或应用程序后按回车定位"
            Top             =   80
            Width           =   3855
         End
         Begin VB.TextBox txtStop 
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   120
            Width           =   90
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfPer 
            Height          =   3255
            Left            =   120
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   600
            Width           =   7215
            _cx             =   12726
            _cy             =   5741
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
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
            ExplorerBar     =   1
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
         Begin VB.Label lblFind 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "查找"
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   165
            Width           =   360
         End
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用程序授权"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmAppLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsUsers As ADODB.Recordset  '保存用户的记录集,用于检测用户名可用性和加载查找窗体
Private mrsApps As ADODB.Recordset  '保存用户的记录集,用于检测用户名可用性和加载查找窗体
Private mstrApp As String   '保存选中的数据
Private mblnKeypress As Boolean

Private Enum Color
    tipColor = &H80000010
    txtColor = &H80000012
End Enum
Private Const conCol = "选择,250,1;应用程序,1200,1;用户名,1200,1;姓名,1200,1;开始IP,1200,1;结束IP,1200,1;状态,500,1;说明,1200,1"

Private Sub ChangeCmdVisiable(ByVal blnIsAdd)
    '修改按钮可见性
    cmdAdd.Visible = Not blnIsAdd
    cmdDel.Visible = Not blnIsAdd
    cmdEdit.Visible = Not blnIsAdd
    cmdStop.Visible = Not blnIsAdd
    cmdSave.Visible = blnIsAdd
    cmdCancel.Visible = blnIsAdd
    
    '选项框
    If blnIsAdd Then
        txtUser.Text = ""
        txtApp.Text = ""
        txtbeforeIp(0).Text = ""
        txtbeforeIp(1).Text = ""
        txtbeforeIp(2).Text = ""
        txtbeforeIp(3).Text = ""
        txtIP(2).Text = ""
        txtDesc.Text = ""
    Else
        With vsfPer
            vsfPer_AfterRowColChange 0, 0, .Row, .Col
        End With
    End If
    
    '输入框可用性
    cmdMore.Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtUser.Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtApp.Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(0).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(1).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(2).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(3).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtIP(2).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtDesc.Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
End Sub

Private Sub cmdAdd_Click()
    ChangeCmdVisiable True
End Sub

Private Sub cmdCancel_Click()
    ChangeCmdVisiable False
End Sub
Private Sub cmdSave_Click()
    Dim strTmp As String, i As Integer
    Dim strStartIP As String, strEndIp As String
    Dim strUser As String, strApp As String, strDesc As String
    Dim varUsers As Variant
    
    On Error GoTo errh
    '校验输入
    Call GetDataFromCard(strUser, strApp, strStartIP, strEndIp, strDesc)

    If mrsUsers Is Nothing Then
        Set mrsUsers = LoadUsers
    End If
    
    strTmp = CheckExist("用户名", strUser, mrsUsers)
    If strTmp <> "" Then
        MsgBox "以下用户名:" & strTmp & "不存在,请检查后重新输入。", , "提示"
        Exit Sub
    End If
    
    '检验唯一性
    strTmp = CheckPerOnly(strApp, strUser)
    If strTmp <> "" Then
        MsgBox "以下用户对于所选应用程序已有相关规则,请修改后再操作" & vbNewLine & strTmp, , "提示"
        Exit Sub
    End If
    
    strTmp = ValidateTxt
    If strTmp <> "" Then
        frmMDIMain.stbThis.Panels(2).Text = strTmp
        Exit Sub
    End If
    
    '提交数据
    gcnOracle.BeginTrans
    Screen.MousePointer = vbArrowHourglass
    If Len(strUser) < 2000 Then
        Call ExecuteProcedure("zltools.Zl_Zlapppermission_Edit(1,'" & strApp & "','" & strUser & "','" & strStartIP & "','" & strEndIp & "',1,'" & strDesc & "','','')", Me.Caption)
    Else
        varUsers = TranStr2Var(strUser, ",", 2000)
        For i = 0 To UBound(varUsers)
            Call ExecuteProcedure("zltools.Zl_Zlapppermission_Edit(1,'" & strApp & "','" & varUsers(i) & "','" & strStartIP & "','" & strEndIp & "',1,'" & strDesc & "','','')", Me.Caption)
        Next
    End If
    gcnOracle.CommitTrans
    Screen.MousePointer = vbDefault
    
    With vsfPer
        .Redraw = flexRDNone
        Call LoadAppPermission
    End With
    frmMDIMain.stbThis.Panels(2).Text = "添加规则成功。"
    Exit Sub
errh:
    Screen.MousePointer = vbDefault
    frmMDIMain.stbThis.Panels(2).Text = ""
    
    If InStr(1, UCase(err.Description), "ORA") Then '数据库错误,字符串较长,弹窗提示,同时回退事务
        MsgBox "添加规则失败。原因：" & vbNewLine & err.Description
        gcnOracle.RollbackTrans
    Else
        frmMDIMain.stbThis.Panels(2).Text = "添加规则失败。原因：" & vbNewLine & err.Description
    End If
End Sub

Private Sub cmdDel_Click()
    Dim varApps As Variant
    Dim i As Integer, intSRow As Integer
    
    mstrApp = GetSelectData
    '字符串长度小于2000的,直接进行删除,超过2000的,进行拆分后分批删除
    Screen.MousePointer = vbArrowHourglass
    gcnOracle.BeginTrans
    If Len(mstrApp) < 2000 Then
        Call ExecuteProcedure("Zl_ZlApppermission_Delete('" & mstrApp & "')", Me.Caption)
    Else
        varApps = TranStr2Var(mstrApp, ",", 2000)
        For i = 0 To UBound(varApps)
            Call ExecuteProcedure("Zl_ZlApppermission_Delete('" & varApps(i) & "')", Me.Caption)
        Next
    End If
    gcnOracle.CommitTrans
    Screen.MousePointer = vbDefault
    
    With vsfPer
        intSRow = .Row
        .Redraw = flexRDNone
        
        '判断删除的数据
        For i = .FixedRows To .Rows - .FixedRows
            If i > .Rows - .FixedRows Or .Rows = .FixedRows Then
                Exit For
            End If
            If InstrEx(mstrApp, .TextMatrix(i, .ColIndex("应用程序")) & ":" & .TextMatrix(i, .ColIndex("用户名"))) Then
                .RemoveItem (i)
                i = i - 1
            End If
        Next
        .Redraw = flexRDDirect
        
        '选中当前行
        If intSRow > .Rows - .FixedRows Then
            .Select .Rows - .FixedRows, 0
        Else
            .Select intSRow, 0
        End If
        .TopRow = .Row
    End With
    mstrApp = GetSelectData
    frmMDIMain.stbThis.Panels(2).Text = "删除规则成功。"
    Exit Sub
    
errh:
    Screen.MousePointer = vbDefault
    
    If InStr(1, UCase(err.Description), "ORA") Then '数据库错误,字符串较长,弹窗提示,同时回退事务
        MsgBox "删除规则失败。原因：" & vbNewLine & err.Description
        gcnOracle.RollbackTrans
    Else
        frmMDIMain.stbThis.Panels(2).Text = "删除规则失败。原因：" & vbNewLine & err.Description
    End If
End Sub

Private Sub cmdEdit_Click()
    EditPermission
End Sub

Private Sub CmdStop_Click()
    EditPermission ("停用")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 And Shift = 2 Then
        cmdMore_Click
    End If
End Sub

Private Sub cmdMore_Click()
    Dim strUsers As String
    Dim p As PointAPI
    Dim rstmp As ADODB.Recordset
    Dim strTmp() As String, i As Integer
    
    p.X = (pctOption.Left + cmdMore.Left + cmdMore.Width - FindUserWidth) / Screen.TwipsPerPixelX
    p.Y = (pctOpt.Top + pctContainer.Top - FindUserHeight) / Screen.TwipsPerPixelY
    ClientToScreen Me.hwnd, p
    
    If mrsUsers Is Nothing Then
        Set mrsUsers = LoadUsers
    End If
    
    strUsers = frmFindUser.ShowMe(Me, mrsUsers, Trim(txtUser.Text), p.X * Screen.TwipsPerPixelX, p.Y * Screen.TwipsPerPixelY)
    txtUser.Text = strUsers
    
End Sub

Private Sub Form_Load()
    Call InitTable(vsfPer, conCol)
    Call LoadAppPermission
    Call ChangeCmdVisiable(False)
    '初始化表格选择框
    With vsfPer
        .ColSort(-1) = flexSortCustom
        .ColSort(0) = flexSortNone
        .ColDataType(0) = flexDTBoolean
        .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        .Cell(flexcpText, 0, 0) = ""
        .Cell(flexcpPictureAlignment, 0, 0) = flexPicAlignCenterCenter
        .Editable = flexEDKbdMouse
    End With
    
    FindApp
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pctContainer.Width = Me.ScaleWidth
    pctContainer.Height = Me.ScaleHeight - pctContainer.Top
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set mrsApps = Nothing
    Set mrsUsers = Nothing
End Sub

Private Sub pctContainer_Resize()
    On Error Resume Next
    
    pctPer.Width = pctContainer.Width
    pctPer.Height = pctContainer.Height - pctOpt.Height
    
    pctOpt.Width = pctContainer.Width
    pctOpt.Top = pctPer.Top + pctPer.Height
End Sub

Private Sub pctOpt_Resize()
    On Error Resume Next
    
    pctOption.Left = pctOpt.Width - pctOption.Width - 120
End Sub

Private Sub pctPer_Resize()
    On Error Resume Next
    
    vsfPer.Width = pctPer.ScaleWidth - 240
    vsfPer.Height = pctPer.ScaleHeight - vsfPer.Top - 30
    
    lblFind.Left = vsfPer.Left
    txtFind.Left = lblFind.Left + lblFind.Width + 45
End Sub

Private Sub txtApp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = "输入用户名、姓名或应用程序后按回车定位" Then
        txtFind.Text = ""
        txtFind.ForeColor = txtColor
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intRow As Integer
    
    If KeyAscii = 13 Then
        If Trim(txtFind.Text) = "" Then
            '不输入数据按下回车,就刷新
            LoadAppPermission
        Else
            Call GetRowPos(vsfPer, txtFind.Text, "用户名,姓名,应用程序")
        End If
    End If
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "输入用户名、姓名或应用程序后按回车定位"
        txtFind.ForeColor = tipColor
    End If
End Sub

Private Sub txtIp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtUser_LostFocus()
    Call txtUser_KeyPress(13)
End Sub

Private Sub txtUser_Validate(Cancel As Boolean)
     If mblnKeypress Then
        mblnKeypress = False
    Else
        Call txtUser_KeyPress(13)
    End If
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
    Dim strTmp As String, intPos As Integer
    
    If KeyAscii = 13 Then    '按下回车
        strTmp = Trim(txtUser.Text)
        intPos = InStrRev(strTmp, ",")
        strTmp = UCase(Mid(strTmp, intPos + 1))
        If strTmp = "" Then Exit Sub
        strTmp = Left(Trim(txtUser.Text), intPos) & FindUser(strTmp)
        
        txtUser.Text = strTmp
        txtUser.SelStart = Len(strTmp)
    End If
    
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub LoadAppPermission()
'功能:加载应用程序授权信息
    Dim strSQL As String, rstmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errh
    
    strSQL = "Select a.应用程序, a.用户名, c.姓名, a.开始ip, a.结束ip, decode(a.状态,1,'已启用','未启用') 状态, a.说明" & vbNewLine & _
                    "From Zlapppermission A, 上机人员表 B, 人员表 C" & vbNewLine & _
                    "Where a.用户名 = b.用户名(+) And b.人员id = c.Id(+)"
                    
    Set rstmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "LoadAppLimit")
    Set mrsApps = rstmp
                            
    With vsfPer
        If rstmp.RecordCount = 0 Then
             .Rows = .FixedRows
            Exit Sub
        End If

        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = rstmp.RecordCount + .FixedRows
        
        i = .FixedRows
        Do While Not rstmp.EOF
            .RowData(i) = "" & rstmp!应用程序 & ":" & rstmp!用户名
            .TextMatrix(i, 0) = "0"
            .TextMatrix(i, .ColIndex("应用程序")) = rstmp!应用程序 & ""
            .TextMatrix(i, .ColIndex("用户名")) = rstmp!用户名 & ""
            .TextMatrix(i, .ColIndex("姓名")) = rstmp!姓名 & ""
            .TextMatrix(i, .ColIndex("开始ip")) = rstmp!开始IP & ""
            .TextMatrix(i, .ColIndex("结束ip")) = rstmp!结束IP & ""
            .TextMatrix(i, .ColIndex("状态")) = rstmp!状态 & ""
            .TextMatrix(i, .ColIndex("说明")) = rstmp!说明 & ""
            i = i + 1: rstmp.MoveNext
        Loop
        
        .AutoResize = True: .AutoSize 0, .Cols - 1
        
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Select .FixedRows, 0
        End If
    End With
    
    Exit Sub
errh:
    MsgBox err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub vsfPer_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        Dim strTmp() As String
        
        With vsfPer
            If .Redraw = flexRDNone Then Exit Sub
            If .Row = 0 Then Exit Sub
            
            txtUser.Text = .TextMatrix(NewRow, .ColIndex("用户名"))
            txtApp.Text = .TextMatrix(NewRow, .ColIndex("应用程序"))
            txtDesc.Text = .TextMatrix(NewRow, .ColIndex("说明"))
            cmdStop.Caption = IIf(.TextMatrix(NewRow, .ColIndex("状态")) = "已启用", "停用", "启用")
            
            If .TextMatrix(NewRow, .ColIndex("开始IP")) <> "" Then
                strTmp = Split(.TextMatrix(NewRow, .ColIndex("开始IP")), ".")
                txtbeforeIp(0).Text = strTmp(0)
                txtbeforeIp(1).Text = strTmp(1)
                txtbeforeIp(2).Text = strTmp(2)
                txtbeforeIp(3).Text = strTmp(3)
                txtIP(2).Text = Split(.TextMatrix(NewRow, .ColIndex("结束IP")), ".")(3)
            Else
                txtbeforeIp(0).Text = ""
                txtbeforeIp(1).Text = ""
                txtbeforeIp(2).Text = ""
                txtbeforeIp(3).Text = ""
                txtIP(2).Text = ""
            End If
        End With
End Sub

Private Sub vsfper_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsfper_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer
    
    With vsfPer
        If .Rows = .FixedRows Then Exit Sub
        If Col = 0 Then
            If .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture Then
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "-1"
                Next
            Else
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "0"
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfper_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer, blnAllSelectd As Boolean
    
    blnAllSelectd = True
    With vsfPer
        If .Redraw = flexRDNone Then Exit Sub
        
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 0) = "0" Then
                blnAllSelectd = False
                Exit For
            End If
        Next

        
        If blnAllSelectd Then
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
        Else
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        End If
    End With
End Sub


Private Function ValidateTxt() As String
'功能:校验输入是否合法,如果输入不合法,返回错误信息.
    Dim strStartIP As String, strEndIp As String
    Dim strErr As String
    
    If txtUser.Text = "" Or txtApp.Text = "" Then
        ValidateTxt = "用户名和应用程序不能为空。"
        Exit Function
    End If
    
    strStartIP = txtbeforeIp(0) & "." & txtbeforeIp(1) & "." & txtbeforeIp(2) & "." & txtbeforeIp(3)
    strEndIp = txtbeforeIp(0) & "." & txtbeforeIp(1) & "." & txtbeforeIp(2) & "." & txtIP(2)
    Call CheckIpValidate(strStartIP, strEndIp, strErr)
    If strErr <> "" Then
        ValidateTxt = strErr
        Exit Function
    End If
  
End Function


Private Sub GetDataFromCard(ByRef strUser As String, ByRef strApp As String, ByRef strStartIP As String, ByRef strEndIp As String, ByRef strDesc As String)
'功能:从卡片表单中获取数据
    
    strUser = Trim(txtUser.Text)
    strApp = Trim(txtApp.Text)
    
    If txtbeforeIp(0).Text = "" Then
        strStartIP = "": strEndIp = ""
    Else
        strStartIP = txtbeforeIp(0).Text & "." & txtbeforeIp(1).Text & "." & txtbeforeIp(2).Text & "." & txtbeforeIp(3).Text
        strEndIp = txtbeforeIp(0).Text & "." & txtbeforeIp(1).Text & "." & txtbeforeIp(2).Text & "." & IIf(txtIP(2).Text = "", txtbeforeIp(3).Text, txtIP(2).Text)
    End If
    
    strDesc = txtDesc.Text
    
End Sub


Private Sub EditPermission(Optional ByVal strStop As String)
'功能:修改授权
    Dim strTmp As String, i As Integer
    Dim strStartIP As String, strEndIp As String
    Dim strDesc As String, strUser As String, strApp As String
    Dim strNewUser As String, strNewApp As String
    
    On Error GoTo errh
    
    With vsfPer
        strApp = .TextMatrix(.Row, .ColIndex("应用程序"))
        strUser = .TextMatrix(.Row, .ColIndex("用户名"))
    End With
    '校验输入
    Call GetDataFromCard(strNewUser, strNewApp, strStartIP, strEndIp, strDesc)
    
    strTmp = ValidateTxt
    If strTmp <> "" Then
        frmMDIMain.stbThis.Panels(2).Text = strTmp
        Exit Sub
    End If
    
    '提交数据
    If strStop = "" Then
        '说明没有传入停用参数,不做更改,只对其他数据进行修改
        strStop = IIf(vsfPer.TextMatrix(vsfPer.Row, vsfPer.ColIndex("状态")) = "已启用", 1, 0)
    Else
        strStop = IIf(vsfPer.TextMatrix(vsfPer.Row, vsfPer.ColIndex("状态")) = "已启用", 0, 1)
    End If

    Screen.MousePointer = vbArrowHourglass
    Call ExecuteProcedure("zltools.Zl_Zlapppermission_Edit(2,'" & strApp & "','" & strUser & "','" & strStartIP & "','" & strEndIp & "'," & strStop & ",'" & strDesc & "','" & strNewApp & "','" & strNewUser & "' )", Me.Caption)
    Screen.MousePointer = vbDefault
    
    cmdStop.Caption = IIf(strStop = 0, "启用", "停用")
    With vsfPer
        .Redraw = flexRDNone
        Call LoadAppPermission
    End With
    frmMDIMain.stbThis.Panels(2).Text = "修改规则成功。"
    Exit Sub
errh:
    Screen.MousePointer = vbDefault
    frmMDIMain.stbThis.Panels(2).Text = ""
    
    If InStr(1, UCase(err.Description), "ORA") Then '数据库错误,字符串较长,弹窗提示,同时回退事务
        MsgBox "修改规则失败。原因：" & vbNewLine & err.Description
    Else
        frmMDIMain.stbThis.Panels(2).Text = "修改规则失败。原因：" & vbNewLine & err.Description
    End If
End Sub

Private Function GetSelectData() As String
'功能:获取勾选了选项框的数据,返回对应ID,否则返回空值
    Dim i As Integer, strTmp As String
    
    With vsfPer
        If .Rows = .FixedRows Then Exit Function
        
        '检查是否有选中数据
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 0) = "-1" Then
                If strTmp = "" Then
                    strTmp = .TextMatrix(i, .ColIndex("应用程序")) & ":" & .TextMatrix(i, .ColIndex("用户名"))
                Else
                    strTmp = strTmp & "," & .TextMatrix(i, .ColIndex("应用程序")) & ":" & .TextMatrix(i, .ColIndex("用户名"))
                End If
            End If
        Next
        
        If strTmp = "" Then
            '若为空值,获取当前选中行数据
            GetSelectData = .TextMatrix(.Row, .ColIndex("应用程序")) & ":" & .TextMatrix(.Row, .ColIndex("用户名"))
        Else
            GetSelectData = strTmp
        End If
    End With
End Function

Private Sub txtbeforeIp_Change(Index As Integer)
    Dim lngLineNo As Long '行号
    Dim lngColNo  As Long '列号
    err = 0
    On Error Resume Next
    If Trim(txtbeforeIp(0).Text) <> "" And Trim(txtbeforeIp(1).Text) <> "" And Trim(txtbeforeIp(2).Text) <> "" And Trim(txtbeforeIp(3).Text) <> "" And Trim(txtIP(2).Text) <> "" Then
        cmdAdd.Enabled = True
    End If
    Call GetCursorPos(Me.txtbeforeIp(Index).hwnd, lngLineNo, lngColNo)
    If lngColNo > 3 Then
        If Index < 3 Then
            If txtbeforeIp(Index + 1).Enabled Then txtbeforeIp(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtbeforeIp_GotFocus(Index As Integer)
    txtbeforeIp(Index).SelStart = 0
    txtbeforeIp(Index).SelLength = Len(txtbeforeIp(Index).Text)
End Sub

Private Sub txtbeforeIp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lngLineNo As Long '行号
    Dim lngColNo  As Long '列号
    err = 0
    Call GetCursorPos(Me.txtbeforeIp(Index).hwnd, lngLineNo, lngColNo)
    
    Select Case KeyCode
    Case 37    '<-
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtbeforeIp(Index - 1).Enabled Then
                txtbeforeIp(Index - 1).SelStart = Len(txtbeforeIp(Index - 1))
                txtbeforeIp(Index - 1).SetFocus
            End If
        End If
    Case 39    '->
        If Index < 3 Then
            If lngColNo <= Len(txtbeforeIp(Index)) Then Exit Sub
            If txtbeforeIp(Index + 1).Enabled Then
                txtbeforeIp(Index + 1).SelStart = 0
                txtbeforeIp(Index + 1).SetFocus
            End If
        End If
    Case 8     'BACKSPACE
        If Index > 0 Then
            If lngColNo > 1 Then Exit Sub
            If txtbeforeIp(Index - 1).Enabled Then
                txtbeforeIp(Index - 1).SelStart = Len(txtbeforeIp(Index - 1))
                txtbeforeIp(Index - 1).SetFocus
            End If
        End If
    End Select
    
    If InStr(1, "1234567890", Chr(KeyCode)) = 0 Then
        KeyCode = 0
    End If
    
End Sub

Private Sub txtbeforeIp_KeyPress(Index As Integer, KeyAscii As Integer)
    err = 0
    On Error Resume Next
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> 13 Then
            If KeyAscii <> 8 Then
                If KeyAscii = Asc(".") Then
                    If Index < 3 And Index >= 0 And Trim(txtbeforeIp(Index)) <> "" Then
                        If txtbeforeIp(Index + 1).Enabled Then txtbeforeIp(Index + 1).SetFocus
                    End If
                End If
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Function CheckPerOnly(ByVal strApp As String, ByVal strUser As String) As String
'功能:检查程序名和用户的唯一性,唯一返回True,不唯一返回False
    
    Dim strSQL As String, rstmp As ADODB.Recordset
    Dim arrUser() As Variant, strResult As String, i As Integer
    
    On Error GoTo errh
    
    strSQL = "Select 应用程序,用户名" & vbNewLine & _
                "From Zlapppermission" & vbNewLine & _
                "Where 用户名 In (Select /*+ cardinality(A,10) */" & vbNewLine & _
                "               Column_Value" & vbNewLine & _
                "              From Table(f_Str2list([1])) A) And 应用程序 = [2]"
    '有可能因为传入的用户名超出Oracle支持的长度,所以需要进行拆分
    If Len(strUser) > 2000 Then
        arrUser = TranStr2Var(strUser, ",", 2000)
        For i = 0 To UBound(arrUser)
            Set rstmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "CheckPerOnly", arrUser(i), strApp)
            
            Do While Not rstmp.EOF
                If strResult = "" Then
                    strResult = rstmp!用户名
                Else
                    strResult = strResult & "," & rstmp!用户名
                End If
                rstmp.MoveNext
            Loop
        Next
    Else
        Set rstmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "CheckPerOnly", strUser, strApp)
        Do While Not rstmp.EOF
            If strResult = "" Then
                strResult = rstmp!用户名
            Else
                strResult = strResult & "," & rstmp!用户名
            End If
            rstmp.MoveNext
        Loop
    End If
    
    CheckPerOnly = strResult
    Exit Function
errh:
    MsgBox err.Description
    If 0 = 1 Then
        Resume
    End If
End Function


Private Sub GetCursorPos(ByVal hwnd5 As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long
    Dim lParam As Long, wParam As Long
    Dim K As Long
    
    i = SendMessage(hwnd5, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16 '取得目前光标所在位置前有多少个Byte
    LineNo = SendMessage(hwnd5, EM_LINEFROMCHAR, j, 0) '取得光标前面有多少行
    LineNo = LineNo + 1
    K = SendMessage(hwnd5, EM_LINEINDEX, -1, 0)
    '取得目前光标所在行前面有多少个Byte
    ColNo = j - K + 1
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
End Function


Private Sub FindApp()
'功能:获取当前的应用程序至下拉框
    
    Dim strSQL As String, rstmp As ADODB.Recordset
    Dim strResult As String
    
    On Error GoTo errh
    strSQL = "Select Distinct Program From V$session where Program not like 'ORACLE.EXE%' order by Program"
    
    Set rstmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "FindApp")
    
    Do While Not rstmp.EOF
        txtApp.addItem rstmp!Program
        rstmp.MoveNext
    Loop
    Exit Sub
errh:
    MsgBox err.Description
End Sub

