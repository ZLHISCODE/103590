VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.OCX"
Begin VB.Form frm批量付款条件设置 
   Caption         =   "付款单批量生成条件设置"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   Icon            =   "frm批量付款条件设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7665
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPopu 
      Height          =   810
      Left            =   2850
      ScaleHeight     =   750
      ScaleWidth      =   2040
      TabIndex        =   32
      Top             =   5340
      Visible         =   0   'False
      Width           =   2100
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "所有未付供应商(A)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   31
         Top             =   420
         Width           =   1725
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "指定供应商(D)"
         ForeColor       =   &H8000000E&
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   30
         Top             =   135
         Width           =   1725
      End
      Begin VB.Label lblBackColor 
         BackColor       =   &H8000000D&
         Height          =   285
         Left            =   75
         TabIndex        =   33
         Top             =   90
         Width           =   1890
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBuilded 
      Height          =   3645
      Left            =   945
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   6429
      _Version        =   393216
      Rows            =   20
      Cols            =   3
      FixedCols       =   0
      HighLight       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.TextBox txt供应商 
      Height          =   300
      Left            =   4845
      TabIndex        =   17
      Top             =   795
      Width           =   1755
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   300
      Left            =   6945
      Picture         =   "frm批量付款条件设置.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   795
      Width           =   315
   End
   Begin VB.CommandButton cmdDele 
      Enabled         =   0   'False
      Height          =   300
      Left            =   7260
      Picture         =   "frm批量付款条件设置.frx":06D4
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   795
      Width           =   315
   End
   Begin MSComctlLib.ListView lvw供应商 
      Height          =   3240
      Left            =   4005
      TabIndex        =   20
      Top             =   1140
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5715
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "供应商"
         Object.Tag             =   "供应商"
         Text            =   "供应商"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "开户银行"
         Text            =   "开户银行"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "帐号"
         Text            =   "帐号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "联系人"
         Object.Tag             =   "联系人"
         Text            =   "联系人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "类型"
         Object.Tag             =   "类型"
         Text            =   "类型"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fra 
      Caption         =   "付款单设置"
      Height          =   1395
      Left            =   60
      TabIndex        =   11
      Top             =   2940
      Width           =   3870
      Begin VB.ComboBox cbo结算方式 
         Height          =   300
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   330
         Width           =   2640
      End
      Begin VB.TextBox txt说明 
         Height          =   300
         Left            =   165
         MaxLength       =   50
         TabIndex        =   15
         Top             =   915
         Width           =   3570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "付款说明(&F)"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   14
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "结算方式(&J)"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   12
         Top             =   390
         Width           =   990
      End
   End
   Begin VB.Frame fra类型 
      Caption         =   "付款条件"
      Height          =   1845
      Left            =   60
      TabIndex        =   1
      Top             =   945
      Width           =   3840
      Begin VB.CheckBox chkType 
         Caption         =   "其他(&Q)"
         Height          =   225
         Index           =   4
         Left            =   2730
         TabIndex        =   10
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chkType 
         Caption         =   "卫生材料(&L)"
         Height          =   225
         Index           =   3
         Left            =   225
         TabIndex        =   9
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkType 
         Caption         =   "设备(&S)"
         Height          =   225
         Index           =   2
         Left            =   2730
         TabIndex        =   8
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chkType 
         Caption         =   "物资(&W)"
         Height          =   225
         Index           =   1
         Left            =   1545
         TabIndex        =   7
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chkType 
         Caption         =   "药品(&Y)"
         Height          =   225
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Index           =   0
         Left            =   1050
         TabIndex        =   3
         Top             =   240
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   73531395
         CurrentDate     =   38936.9668171296
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Index           =   1
         Left            =   1035
         TabIndex        =   5
         Top             =   660
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   73531395
         CurrentDate     =   38936.4668171296
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "结束时间(&E)"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "开始时间(&K)"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   315
         Width           =   990
      End
   End
   Begin VB.Frame fraTemp 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   24
      Top             =   705
      Width           =   7995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6510
      TabIndex        =   22
      Top             =   4620
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5340
      TabIndex        =   21
      Top             =   4620
      Width           =   1100
   End
   Begin VB.Frame fraTemp 
      Height          =   30
      Index           =   0
      Left            =   -30
      TabIndex        =   23
      Top             =   4425
      Width           =   7905
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   -75
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm批量付款条件设置.frx":0C5E
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   645
      Top             =   4560
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
            Picture         =   "frm批量付款条件设置.frx":10B6
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPross 
      Height          =   615
      Left            =   -15
      ScaleHeight     =   555
      ScaleWidth      =   7620
      TabIndex        =   25
      Top             =   4455
      Visible         =   0   'False
      Width           =   7680
      Begin MSComctlLib.ProgressBar prgb 
         Height          =   285
         Left            =   45
         TabIndex        =   26
         Top             =   255
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblPer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   180
         Left            =   7215
         TabIndex        =   28
         Top             =   45
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lbl供应商 
         Caption         =   "正在生成："
         Height          =   195
         Left            =   75
         TabIndex        =   27
         Top             =   30
         Width           =   2040
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "供应商(&G)"
      Height          =   180
      Index           =   2
      Left            =   4035
      TabIndex        =   16
      Top             =   870
      Width           =   810
   End
   Begin VB.Label lblInfor 
      Caption         =   "设置产生各供应商的批量付款条件，其中供应商可以通过简码等方式进行录入。"
      Height          =   285
      Left            =   735
      TabIndex        =   0
      Top             =   375
      Width           =   6360
   End
   Begin VB.Image img晋升 
      Height          =   480
      Left            =   120
      Picture         =   "frm批量付款条件设置.frx":150E
      Top             =   180
      Width           =   480
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuLocal 
         Caption         =   "指定供应商(&L)"
      End
      Begin VB.Menu mnuPopuAll 
         Caption         =   "所有未付供应商(&A)"
      End
   End
End
Attribute VB_Name = "frm批量付款条件设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnCancel As Boolean
Private mstr供应商权限 As String
Private mstrPrivs As String
Private mfrmMain As Form
Private mblnFirst As Boolean
Private mintColumn As Integer
Public Function ShowCard(ByVal FrmMain As Form, ByVal strPrivs As String) As Boolean
    '------------------------------------------------------------------------------------------------
    '功能:显示批量增加的条件设置
    '参数:frmMain-主窗体
    '     strPrivs-权限串
    '返回:生成成功，返回True,否则False
    '------------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs
    mblnCancel = True
    mstr供应商权限 = ""
    Set mfrmMain = FrmMain
    Me.Show 1, mfrmMain
    ShowCard = Not mblnCancel
End Function

Private Sub cbo结算方式_Click()
    Call SetCtrlEn
End Sub

Private Sub cbo结算方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkType_Click(Index As Integer)
    Call SetCtrlEn
End Sub

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 

Private Sub cmdAdd_Click()
    picPopu.Left = cmdAdd.Left + cmdAdd.Width - picPopu.Width
    picPopu.Top = cmdAdd.Top + cmdAdd.Height
    picPopu.Visible = True
    picPopu.SetFocus
    RaisEffect picPopu, 2
End Sub

Private Sub cmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'PopupMenu Me.mnuPopu, vbPopupMenuRightAlign
End Sub
Private Sub cmdCancel_Click()
'    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdDele_Click()
    Call Dele供应商
    Call SetCtrlEn
End Sub

Private Sub cmdOK_Click()
    Dim lvwItem  As ListItem
    Dim intByte As Integer
    mblnCancel = False
    If zlCommFun.ActualLen(txt说明.Text) > 50 Then
        ShowMsgbox "说明不能大于25个汉字或50个字符!"
        If txt说明.Enabled Then txt说明.SetFocus
        Exit Sub
    End If
    If InStr(1, txt说明.Text, "'") > 0 Then
        ShowMsgbox "说明中不能包含特殊字符(单引号)!"
        If txt说明.Enabled Then txt说明.SetFocus
        Exit Sub
    End If
    If lvw供应商.ListItems.Count = 0 Then Exit Sub
    Screen.MousePointer = 11
    Me.Enabled = False
    picPross.Visible = True
    picPross.ZOrder 1
    prgb.Max = lvw供应商.ListItems.Count
    prgb.Min = 0
    prgb.Value = 0
    cmdOk.Visible = False
    cmdCancel.Visible = False
    
    Call initGrid
    
    For Each lvwItem In Me.lvw供应商.ListItems
        
        Call BuildingData(Val(Mid(lvwItem.Key, 2)), lvwItem.Text)
        prgb.Value = prgb.Value + 1
        lblPer.Caption = Round(prgb.Value / prgb.Max * 100, 2) & "/100"
        DoEvents
    Next
    picPross.Visible = False
    Screen.MousePointer = 0
    'cmdOk.Visible = True
    cmdCancel.Visible = True
    cmdCancel.Caption = "退出(&C)"
    Call ShowBuildedGrid
    
    Me.Enabled = True
    
End Sub


Private Sub dtpDate_Change(Index As Integer)
    Err = 0: On Error Resume Next
    If Index = 0 Then
        If dtpDate(0).Value >= dtpDate(1).Value Then
            dtpDate(1).Value = Format(dtpDate(0).Value, "yyyy-mm-dd") & " 23:59:59"
        End If
    Else
        If dtpDate(0).Value > dtpDate(1).Value Then
            dtpDate(0).Value = Format(dtpDate(1).Value, "yyyy-mm-dd") & " 00:00:00"
        End If
    End If
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If Load结算方式 = False Then Unload Me: Exit Sub
    Call 权限控制
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mstr供应商权限 = " (末级=1 and " & Get分类权限(mstrPrivs) & ") "
    dtpDate(1).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59"
    dtpDate(1).MaxDate = dtpDate(1).Value
    dtpDate(0).Value = Format(DateAdd("d", -7, dtpDate(1).Value), "yyyy-mm-dd") & " 00:00:00"
    dtpDate(0).MaxDate = dtpDate(1).Value
    '恢复相关参数
    RestoreWinState Me, App.ProductName
End Sub

 

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 7785 Then Me.Width = 7785
    If Me.Height < 5475 Then Me.Height = 5475
    
    With picPross
        .Top = ScaleHeight - .Height
        .Width = ScaleWidth - .Left
    End With
    With cmdCancel
        .Top = picPross.Top + (picPross.Height - .Height) / 2
        .Left = Me.ScaleWidth - .Width - 100
        cmdOk.Top = .Top
        cmdOk.Left = .Left - cmdOk.Width - 50
    End With
    With fraTemp(0)
        .Top = picPross.Top
        .Width = Me.ScaleWidth + 100
    End With
    With lvw供应商
        .Width = ScaleWidth - .Left - 50
        .Height = fraTemp(0).Top - .Top - 50
    End With
    With cmdDele
        .Left = lvw供应商.Left + lvw供应商.Width - .Width
        cmdAdd.Left = .Left - cmdAdd.Width
    End With
    fraTemp(1).Width = Me.ScaleWidth + 100
    With mshBuilded
        .Left = 50
        .Top = Me.txt供应商.Top
        .Height = fraTemp(0).Top - .Top - 10
        .Width = Me.ScaleWidth - .Left
    End With
    With txt供应商
        .Width = cmdAdd.Left - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lblMenu_Click(Index As Integer)
        picPopu.Visible = False
        Select Case Index
        Case 0
            Call mnuPopuLocal_Click
        Case 1
            Call mnuPopuAll_Click
        End Select
    
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    
    If Index = 0 Then
        lblMenu(Index).ForeColor = vbWhite
        lblMenu(1).ForeColor = &H80000012
    Else
        lblMenu(Index).ForeColor = vbWhite
        lblMenu(0).ForeColor = &H80000012
    End If
    With lblBackColor
        .Top = lblMenu(Index).Top - 25
    End With

End Sub

Private Sub lvw供应商_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvw供应商.Sorted = True
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvw供应商.SortOrder = IIf(lvw供应商.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvw供应商.SortKey = mintColumn
        lvw供应商.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw供应商_ItemClick(ByVal Item As MSComctlLib.ListItem)
        Call SetCtrlEn
End Sub

Private Sub lvw供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Function Get系统类型() As String
    '-------------------------------------------------------------------------------------------
    '功能:获取系统类型
    '返回:以1,2,3的形式返回
    '-------------------------------------------------------------------------------------------
    Dim str类型  As String
    str类型 = ""
    '1——药品应付款   2——物资应付款   3——设备应付款   4——其他,5--卫生材料
    str类型 = IIf(chkType(0).Value = 1, ",1", "")
    str类型 = str类型 & IIf(chkType(1).Value = 1, ",2", "")
    str类型 = str类型 & IIf(chkType(2).Value = 1, ",3", "")
    str类型 = str类型 & IIf(chkType(3).Value = 1, ",5", "")
    str类型 = str类型 & IIf(chkType(4).Value = 1, ",4", "")
    If str类型 <> "" Then
        str类型 = Mid(str类型, 2)
    End If
    Get系统类型 = str类型
End Function
Private Sub mnuPopuAll_Click()
        '全选所有数据的供应商
    Dim rsData As New ADODB.Recordset
    Dim str类型 As String
    Dim dtStartdate As Date
    Dim dtEndDate As Date
    Dim lvwItem As ListItem
    
    Err = 0: On Error GoTo errHand:
    
    str类型 = Get系统类型
    If str类型 = "" Then
        ShowMsgbox "未选择本次生成的系统类型!"
        Exit Sub
    End If
    
    str类型 = " And A.系统标识 in (" & str类型 & ")"
    
    If Me.lvw供应商.ListItems.Count <> 0 Then
        If MsgBox("已经选择了供应商，是否先清除所选择的供应商？", vbQuestion + vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
            Me.lvw供应商.ListItems.Clear
        End If
    End If
    zlCommFun.ShowFlash "正在获取供应商,请稍后...."
    
    gstrSQL = "Select distinct b.id, b.编码,b.名称 ,b.开户银行,b.帐号,b.联系人,b.类型" & _
             "   FROM 应付记录 A,供应商 b" & _
             "   WHERE  a.单位id+0=b.id and a.计划日期 IS NULL AND a.记录性质 <> -1  AND a.付款序号 IS NULL AND " & _
             "         a.审核日期 BETWEEN [1] AND [2] " & str类型 & IIf(mstr供应商权限 <> "", " and " & mstr供应商权限, "") & _
             "   order by b.编码"

    dtStartdate = Format(Me.dtpDate(0).Value, "yyyy-MM-DD hh:mm:ss")
    dtEndDate = Format(Me.dtpDate(1).Value, "yyyy-MM-DD hh:mm:ss")
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, dtStartdate, dtEndDate)
    
    If rsData.EOF Then
        zlCommFun.StopFlash
        ShowMsgbox "当前范围内没有相应数据发生的供应商!"
        Exit Sub
    End If
    With rsData
        Do While Not .EOF
            Err = 0: On Error Resume Next
            Set lvwItem = Me.lvw供应商.ListItems.Add(, "K" & !ID, !编码 & "-" & !名称, 1, 1)
            If Err = 0 Then
                lvwItem.SubItems(1) = Nvl(!开户银行)
                lvwItem.SubItems(2) = Nvl(!帐号)
                lvwItem.SubItems(3) = Nvl(!联系人)
                lvwItem.SubItems(4) = Nvl(!类型)
            Else
                Err = 0: On Error GoTo 0:
            End If
            .MoveNext
        Loop
    End With
    zlCommFun.StopFlash
    Call SetCtrlEn
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuPopuLocal_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim blnCancel As Boolean
    
    gstrSQL = "" & _
        "   Select id,上级ID, 编码,名称,末级,简码,许可证号,许可证效期,执照号,执照效期,税务登记号,地址,开户银行,帐号,联系人,建档时间,类型,信用期" & _
        "   From 供应商 " & _
        "   where (撤档时间 is null or 撤档时间>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & zl_获取站点限制 & " " & _
            IIf(mstr供应商权限 <> "", " and (末级<>1 or " & mstr供应商权限 & ")", "") & _
        "   start with 上级id is null connect by prior id=上级id"
        
    
    'ShowSelect:
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 2, "供应商选择", False, , "选择供应商", False, True, False, , , , blnCancel, , True)
    If blnCancel Or rsTemp Is Nothing Then Exit Sub
    If Add供应商(rsTemp) = False Then Exit Sub
    Call SetCtrlEn
End Sub

Private Sub picPopu_KeyDown(KeyCode As Integer, Shift As Integer)
      
        If KeyCode = vbKeyReturn Then
            If lblBackColor.Top < lblMenu(0).Top Then
               Call lblMenu_Click(0)
            Else
               Call lblMenu_Click(1)
            End If
        End If
        If KeyCode = vbKeyDown Then
            If lblBackColor.Top < lblMenu(0).Top Then
                lblMenu_MouseMove 1, 0, 0, 0, 0
            Else
                lblMenu_MouseMove 0, 0, 0, 0, 0
            End If
        ElseIf KeyCode = vbKeyUp Then
            If lblBackColor.Top < lblMenu(0).Top Then
                lblMenu_MouseMove 1, 0, 0, 0, 0
            Else
                lblMenu_MouseMove 0, 0, 0, 0, 0
            End If
        End If
        If KeyCode = vbKeyA And Shift = 4 Then
               Call lblMenu_Click(1)
        End If
        If KeyCode = vbKeyD And Shift = 4 Then
               Call lblMenu_Click(0)
        End If
End Sub

Private Sub picPross_Resize()
    prgb.Width = picPross.ScaleWidth - prgb.Left - 50
    lblPer.Left = prgb.Width + prgb.Left - lblPer.Width - 20
End Sub

Private Sub picPopu_LostFocus()
    picPopu.Visible = False
End Sub

Private Sub picPopu_Paint()
    RaisEffect picPopu, 2
          
End Sub

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txt供应商.Tag <> "" Then Exit Sub
    strKey = GetMatchingSting(UCase(txt供应商.Text))
    
    gstrSQL = "" & _
        "   Select id, 编码,名称,末级,简码,许可证号,许可证效期,执照号,执照效期,税务登记号,地址,开户银行,帐号,联系人,建档时间,类型,信用期" & _
        "   From 供应商 " & _
        "   where 末级=1 " & zl_获取站点限制 & " and  (撤档时间 is null or 撤档时间>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & IIf(mstr供应商权限 <> "", " and " & mstr供应商权限, "") & _
        "          and (编码 like [1] or 名称 like [1] or 简码 like [1])"
    'ShowSelect:
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Dim lngX As Long, lngY As Long, lngH As Long
    lngX = Me.Left + txt供应商.Left + Screen.TwipsPerPixelX
    lngY = Me.Top + Me.Height - Me.ScaleHeight + txt供应商.Top
    lngH = txt供应商.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "供应商选择", False, "", "选择供应商", False, True, True, lngX, lngY, lngH, blnCancel, False, True, strKey)
    
    If blnCancel Or rsTemp Is Nothing Then Exit Sub
    If rsTemp.State <> 1 Then Exit Sub
    If Add供应商(rsTemp) = False Then Exit Sub
    Call SetCtrlEn
End Sub
Private Function Add供应商(ByVal rsTemp As ADODB.Recordset) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能:增加供应商
    '------------------------------------------------------------------------------------------------------------------------
    Dim lvwItem As ListItem
    If rsTemp.EOF Then Exit Function
    
    
    Err = 0: On Error Resume Next:
    Set lvwItem = Me.lvw供应商.ListItems.Add(, "K" & rsTemp!ID, rsTemp!编码 & "-" & rsTemp!名称, 1, 1)
    If Err <> 0 Then
        MsgBox "你选择的供应商已经存在,不能再选择！", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    Err = 0: On Error GoTo errHand:
    lvwItem.SubItems(1) = Nvl(rsTemp!开户银行)
    lvwItem.SubItems(2) = Nvl(rsTemp!帐号)
    lvwItem.SubItems(3) = Nvl(rsTemp!联系人)
    lvwItem.SubItems(4) = Nvl(rsTemp!类型)
    Add供应商 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function Dele供应商() As Boolean
    Dim intIndex As Integer
    Err = 0: On Error Resume Next
    With lvw供应商
        '再删除ListView中对应节点
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
End Function
Private Sub SetCtrlEn()
    '功能:设置控件属性
    Dim blnData As Boolean
    Dim blnSel As Boolean
    Dim blnCheck As Boolean
    blnData = Me.lvw供应商.ListItems.Count <> 0
    blnSel = Not Me.lvw供应商.SelectedItem Is Nothing
    blnCheck = Me.chkType(0).Value = 1 Or Me.chkType(1).Value = 1 Or Me.chkType(2).Value = 1 Or Me.chkType(3).Value = 1 Or Me.chkType(4).Value = 1
    Me.cmdDele.Enabled = blnSel And blnData
    Me.cmdOk.Enabled = blnData And blnCheck And cbo结算方式.Text <> ""
    
End Sub
Private Function Load结算方式() As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:加载结算方式
    '-----------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo errHand:
    gstrSQL = "Select 结算方式,缺省标志 From 结算方式应用 Where 应用场合='付货款' Order by 缺省标志 desc"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.EOF Then
        ShowMsgbox "未设置结算方式或应用块合，请到［结算方式管理］中设置!"
        Exit Function
    End If
    With rsTemp
        Me.cbo结算方式.Clear
        Do While Not .EOF
            Me.cbo结算方式.AddItem Nvl(!结算方式)
            If Val(Nvl(!缺省标志)) = 1 Then
                Me.cbo结算方式.ListIndex = Me.cbo结算方式.NewIndex
            End If
            .MoveNext
        Loop
        If Me.cbo结算方式.ListCount <> 0 And Me.cbo结算方式.ListIndex < 0 Then
            Me.cbo结算方式.ListIndex = 0
        End If
    End With
    Load结算方式 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Sub txt说明_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub

Private Function BuildingData(ByVal lng供应商ID As Long, ByVal str供应商名称 As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------
    '功能:生成供应商的付款单
    '------------------------------------------------------------------------------------------------------------------------------
    Dim rsData As New ADODB.Recordset
    Dim str类型 As String
    
    Dim dtStartdate As Date
    Dim dtEndDate As Date
    Err = 0: On Error GoTo errHand:
    
    str类型 = Get系统类型
    If str类型 = "" Then
        ShowMsgbox "未选择本次生成的系统类型!"
        Exit Function
    End If
    str类型 = " And 系统标识 in (" & str类型 & ")"
    
    lbl供应商.Caption = "正在获取" & str供应商名称 & " 的数据...."
    
    gstrSQL = "Select  MAX(ID) ID, MAX(记录状态) 记录状态, '' 计划日期, 发票号, 入库单据号, " & _
             "          SUM(Nvl(数量, 0)) AS 数量, " & _
             "          SUM(Nvl(发票金额, 0)) AS 发票金额 " & _
             "   FROM 应付记录 " & _
             "   WHERE 计划日期 IS NULL AND 记录性质 <> -1 AND 单位id+0 = [3] AND 付款序号 IS NULL AND " & _
             "         审核日期 BETWEEN [1] AND [2] " & str类型 & _
             "   GROUP BY 记录性质, NO, 项目id,序号,付款序号,发票号,入库单据号 " & _
             "   HAVING SUM(Nvl(发票金额, 0)) <> 0 " & _
             "   ORDER BY 发票号"
    dtStartdate = Format(Me.dtpDate(0).Value, "yyyy-MM-DD hh:mm:ss")
    dtEndDate = Format(Me.dtpDate(1).Value, "yyyy-MM-DD hh:mm:ss")
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, dtStartdate, dtEndDate, lng供应商ID)
    '获取单据号及付款序号
    Dim lng付款序号 As Long, strNO As String
    Dim dbl付款金额 As Double, lngCount As Long
        
    Err = 0: On Error GoTo ErrRoll:
    gcnOracle.BeginTrans
    If rsData.RecordCount = 0 Then
        '无数据
        Call SetGridNewRowValue(str供应商名称, "", 0, 0)
    Else
        lng付款序号 = zlDatabase.GetNextId("付款记录")
        strNO = zlDatabase.GetNextNo(31)
        dbl付款金额 = 0
        With rsData
            Do While Not .EOF
                '过程参数
                '       ID_IN ,计划序号_IN(以0,1,2,3方式传入),
                '        付款序号_IN,预付款_IN
            
                gstrSQL = "zl_付款序号_UPDATE(" & _
                     "'" & Nvl(!ID) & "'," & _
                    "NULL," & _
                    lng付款序号 & "," & _
                    "0)"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                dbl付款金额 = dbl付款金额 + Val(Nvl(!发票金额))
                lngCount = lngCount + 1
                lbl供应商.Caption = "正在获取" & str供应商名称 & " 的数据 " & lngCount & "...."
                .MoveNext
            Loop
        End With
        dbl付款金额 = Round(dbl付款金额, 2)
        '保存单据
        gstrSQL = "" & _
        "   zl_付款管理_INSERT('" & _
            strNO & "'," & _
            1 & "," & _
            0 & "," & _
            lng供应商ID & "," & _
            dbl付款金额 & ",'" & _
            cbo结算方式.Text & "'," & _
            "NULL,'" & _
            gstrUserName & "',to_date('" & _
            Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')," & _
            lng付款序号 & ",'" & _
            txt说明.Text & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        Call SetGridNewRowValue(str供应商名称, strNO, lngCount, dbl付款金额)
    End If
    gcnOracle.CommitTrans
    BuildingData = True
    Exit Function
errHand:
    '无数据
    Call SetGridNewRowValue(str供应商名称, "", 0, 0)
    If ErrCenter = 1 Then Resume
    Exit Function
ErrRoll:
    
    Call ErrCenter
    gcnOracle.RollbackTrans
End Function
Private Sub initGrid()
    '------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化列头
    '------------------------------------------------------------------------------------------------------------------------------
    With mshBuilded
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "供应商"
        .TextMatrix(0, 1) = "付款单号"
        .TextMatrix(0, 2) = "明细总数"
        .TextMatrix(0, 3) = "总发票额"
        .TextMatrix(0, 4) = "说明"
        
        .ColWidth(0) = 2000
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1400
        .ColWidth(4) = 1500
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 1
    End With
End Sub
Private Sub SetGridNewRowValue(ByVal str供应商 As String, ByVal strNO As String, ByVal lngCount As Long, ByVal dbl发票总额 As Double)
    '------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置grid的新行值
    '参数; str供应商-供应商
    '      strNo-单据号
    '      lngCount-明细总数
    '      dbl发票总额-发票总额
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With mshBuilded
        .Rows = .Rows + 1
        .row = .Rows - 2
        .TextMatrix(.row, 0) = str供应商
        .TextMatrix(.row, 1) = strNO
        .TextMatrix(.row, 2) = lngCount
        .TextMatrix(.row, 3) = Format(dbl发票总额, "###0.00;-###0.00;0;0")
        If strNO = "" Then
            .TextMatrix(.row, 4) = "无数据发生,未生成付款单据"
            For i = 0 To .Cols - 1
                .Col = i
                .CellForeColor = vbRed
            Next
            .Col = 0
        End If
    End With
End Sub
Private Sub ShowBuildedGrid()
    '--------------------------------------------------------------------------
    '功能:显示生成的数据
    '--------------------------------------------------------------------------
    mshBuilded.Visible = True
    Me.lblInfor.Caption = "以下网格数据是本次生成付款单情况!"
    Me.Caption = "批量生成结果预览"
End Sub



Private Sub 权限控制()
    '权限控制
    Dim bln药品 As Boolean
    Dim bln物资 As Boolean
    Dim bln设备 As Boolean
    Dim bln其他 As Boolean
    Dim bln卫材 As Boolean
    
    bln药品 = InStr(1, mstrPrivs, ";药品;") <> 0
    bln物资 = InStr(1, mstrPrivs, ";物资;") <> 0
    bln设备 = InStr(1, mstrPrivs, ";设备;") <> 0
    bln其他 = InStr(1, mstrPrivs, ";其他;") <> 0
    bln卫材 = InStr(1, mstrPrivs, ";卫材;") <> 0
    
    chkType(0).Enabled = bln药品
    chkType(0).Value = IIf(bln药品, 1, 0)
    chkType(1).Enabled = bln物资
    chkType(1).Value = IIf(bln物资, 1, 0)
    chkType(2).Enabled = bln设备
    chkType(2).Value = IIf(bln设备, 1, 0)
    chkType(3).Enabled = bln卫材
    chkType(3).Value = IIf(bln卫材, 1, 0)
    chkType(4).Enabled = bln其他
    chkType(4).Value = IIf(bln其他, 1, 0)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt说明, KeyAscii, m文本式
End Sub
