VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRunLog 
   BackColor       =   &H80000005&
   Caption         =   "运行日志管理"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmRunLog.frx":0000
   ScaleHeight     =   6090
   ScaleWidth      =   8130
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   300
      ScaleHeight     =   3480
      ScaleWidth      =   3225
      TabIndex        =   12
      Top             =   1155
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Frame Fra查找 
         BackColor       =   &H80000005&
         Height          =   3720
         Left            =   -15
         TabIndex        =   13
         Top             =   -195
         Width           =   3255
         Begin VB.TextBox txt工作内容 
            Height          =   300
            Left            =   960
            TabIndex        =   23
            Top             =   1590
            Width           =   2145
         End
         Begin VB.TextBox txt部件名 
            Height          =   300
            Left            =   960
            TabIndex        =   22
            Top             =   1230
            Width           =   2145
         End
         Begin VB.CommandButton cmdReset 
            Cancel          =   -1  'True
            Caption         =   "重设条件"
            Height          =   350
            Left            =   90
            TabIndex        =   21
            Top             =   3120
            Width           =   915
         End
         Begin VB.Frame FraHead 
            BackColor       =   &H80000005&
            Height          =   405
            Left            =   30
            TabIndex        =   14
            Top             =   60
            Width           =   3195
            Begin VB.PictureBox PicClose 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   200
               Left            =   2925
               Picture         =   "FrmRunLog.frx":04F9
               ScaleHeight     =   195
               ScaleWidth      =   210
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   150
               Width           =   215
            End
            Begin VB.Label LblHead 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "条件设置"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   90
               TabIndex        =   15
               Top             =   160
               Width           =   720
            End
         End
         Begin VB.ComboBox Cbo工作站 
            Height          =   300
            Left            =   960
            TabIndex        =   4
            Top             =   510
            Width           =   2145
         End
         Begin VB.ComboBox Cbo用户名 
            Height          =   300
            Left            =   960
            TabIndex        =   5
            Top             =   870
            Width           =   2145
         End
         Begin VB.CommandButton Cmd确定 
            Caption         =   "确定(&O)"
            Height          =   350
            Left            =   1230
            TabIndex        =   7
            Top             =   3120
            Width           =   915
         End
         Begin VB.CommandButton Cmd取消 
            Caption         =   "取消(&C)"
            Height          =   350
            Left            =   2160
            TabIndex        =   8
            Top             =   3120
            Width           =   915
         End
         Begin MSComCtl2.DTPicker dtpDateStart 
            Height          =   315
            Left            =   960
            TabIndex        =   6
            Top             =   1950
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   106299395
            CurrentDate     =   37029
         End
         Begin MSComCtl2.DTPicker dtpDateEnd 
            Height          =   315
            Left            =   960
            TabIndex        =   24
            Top             =   2625
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   106299395
            CurrentDate     =   37029
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "至"
            Height          =   180
            Left            =   960
            TabIndex        =   25
            Top             =   2340
            Width           =   180
         End
         Begin VB.Label Lbl工作站 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "工作站"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   375
            TabIndex        =   20
            Top             =   570
            Width           =   540
         End
         Begin VB.Label Lbl部件名 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "部件名"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   375
            TabIndex        =   19
            Top             =   1290
            Width           =   540
         End
         Begin VB.Label Lbl用户名 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "用户名"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   375
            TabIndex        =   18
            Top             =   930
            Width           =   540
         End
         Begin VB.Label Lbl模块名 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "模块名"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   375
            TabIndex        =   17
            Top             =   1650
            Width           =   540
         End
         Begin VB.Label Lbl进入日期 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "进入时间"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   16
            Top             =   2010
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.ImageList ImgLvw 
      Left            =   255
      Top             =   1185
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
            Picture         =   "FrmRunLog.frx":0A47
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRunLog.frx":0BA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRunLog.frx":19F3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   1020
      TabIndex        =   3
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "查看(&V)"
      Height          =   350
      Left            =   4380
      TabIndex        =   1
      Top             =   630
      Width           =   1100
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   255
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   570
      Width           =   495
      Begin VB.Image imgMain 
         Height          =   480
         Left            =   15
         Picture         =   "FrmRunLog.frx":2845
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComctlLib.ListView LvwList 
      Height          =   4155
      Left            =   285
      TabIndex        =   0
      Top             =   1155
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImgLvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "类型"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "工作站"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "用户名"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "部件名"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "模块名"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "进入时间"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "退出时间"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   5670
      TabIndex        =   2
      Top             =   630
      Width           =   1100
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行日志管理"
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
      Left            =   195
      TabIndex        =   10
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "FrmRunLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'用于调整listview行高
Private Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Const LVM_FIRST = &H1000
Private Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Private Const LVSIL_SMALL = 1
Private Const LVM_UPDATE = (LVM_FIRST + 42)
Private hImageList As Long

Private RecLog As New ADODB.Recordset                       '日志记录集
Private strSQL As String                                    'SQL语句
Private StrDefaultSQL As String                             '缺省查找串
Private StrFindSQL As String                                '查找串

Private Type MousePoint
    X As Single
    Y As Single
End Type
Private Type WindowRect
    Left As Single
    Top As Single
End Type
Private CurMousePoint As MousePoint
Private CurWindowRect As WindowRect

Private Sub CmdDelete_Click()
    Dim ItemThis As ListItem
    '显示或屏蔽"删除选择菜单"
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub

    For Each ItemThis In LvwList.ListItems
        If ItemThis.Selected Then Exit For
    Next

    If ItemThis.Selected = False Then Exit Sub
    PopupMenu frmRegMenus.TrackMenu, 2, CmdDelete.Left, CmdDelete.Top + CmdDelete.Height
End Sub

Private Sub cmdReset_Click()
    Cbo工作站.Text = ""
    Cbo用户名.Text = ""
    txt部件名.Text = ""
    txt工作内容.Text = ""
    
    dtpDateStart.value = date
    dtpDateEnd.value = date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload frmRegMenus
    SetListViewRowHeight_Destroy
End Sub

Private Sub FraHead_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicFind_MouseDown Button, Shift, X, Y
End Sub

Private Sub FraHead_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicFind_MouseMove Button, Shift, X, Y
End Sub

Private Sub Fra查找_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicFind_MouseDown Button, Shift, X, Y
End Sub

Private Sub Fra查找_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicFind_MouseMove Button, Shift, X, Y
End Sub

Private Sub LvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With LvwList
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = 0, 1, 0)
        .Sorted = True
    End With
End Sub

Private Sub LvwList_DblClick()
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub
    
    CmdView_Click
End Sub

Private Sub LvwList_KeyDown(KeyCode As Integer, Shift As Integer)
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = vbKeyDelete Then Call DeleteCurLog(Me, True): Exit Sub
    If KeyCode = vbKeyReturn Then CmdView_Click
End Sub

Private Sub LvwList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ItemThis As ListItem
    '显示或屏蔽"删除选择菜单"
    
    If Button <> 2 Then Exit Sub
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub
    
    For Each ItemThis In LvwList.ListItems
        If ItemThis.Selected Then Exit For
    Next
    
    If ItemThis.Selected = False Then Exit Sub
    PopupMenu frmRegMenus.TrackMenu, 2
End Sub

Private Sub CmdView_Click()
    Dim ItemThis As ListItem
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub
    
    Set ItemThis = LvwList.SelectedItem
    With FrmRunLogProperty
        .Txt会话号 = ItemThis.Tag
        .Txt工作站 = ItemThis.SubItems(1)
        .Txt用户名 = ItemThis.SubItems(2)
        .txt部件名 = ItemThis.SubItems(3)
        .Txt模块名 = ItemThis.SubItems(4)
        .Txt进入时间 = ItemThis.SubItems(5)
        .Txt退出时间 = ItemThis.SubItems(6)
        .Txt退出原因 = ItemThis
        .Show 1
    End With
End Sub

Private Sub PicClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then RaisEffect PicClose, -2
End Sub

Private Sub PicClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then RaisEffect PicClose, 2
    
    If X > 0 And X < PicClose.Width And Y > 0 And Y < PicClose.Height Then Cmd取消_Click
End Sub

Private Sub PicFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With CurMousePoint
            .X = X
            .Y = Y
        End With
    End If
End Sub

Private Sub PicFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With CurWindowRect
            .Left = PicFind.Left + X - CurMousePoint.X
            .Top = PicFind.Top + Y - CurMousePoint.Y
            
            If .Left < ScaleLeft Then .Left = ScaleLeft
            If .Left + PicFind.Width > ScaleWidth Then .Left = ScaleWidth - PicFind.Width
            If .Top < ScaleTop Then .Top = ScaleTop
            If .Top + PicFind.Height > ScaleHeight Then .Top = ScaleHeight - PicFind.Height
        End With
        
        With PicFind
            .Move CurWindowRect.Left, CurWindowRect.Top
        End With
    End If
End Sub

Private Sub cmdFind_Click()
    With PicFind
        .Visible = True
        
        CmdFind.Enabled = .Visible Xor True
        CmdDelete.Enabled = CmdFind.Enabled
        CmdView.Enabled = CmdFind.Enabled
        LvwList.Enabled = CmdFind.Enabled
        
        Cbo工作站.SetFocus
    End With
End Sub

Private Sub Cmd取消_Click()
    CmdFind.Enabled = True
    CmdDelete.Enabled = (LvwList.ListItems.Count <> 0)
    CmdView.Enabled = (LvwList.ListItems.Count <> 0)
    LvwList.Enabled = CmdFind.Enabled
    LvwList.SetFocus
    PicFind.Visible = False
End Sub

Private Sub Cmd确定_Click()
    If GetFindSQL = False Then Exit Sub
    
    CmdDelete.Enabled = True
    CmdView.Enabled = True
    LvwList.Enabled = True
    LvwList.SetFocus
    PicFind.Visible = False
    frmMDIMain.stbThis.Panels(2).Text = "正在查找！"
    Call RefreshData
    
    CmdFind.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim StrDate As String
    
    With frmRegMenus
        .Bln日志 = True
        Set .FrmObj = Me
    End With
    
    RaisEffect PicClose, 2
    
    '获取供用户选择的内容
    Call InitCons
    
    '设置缺省查找串(查找当天的运行日志)
    StrDate = Format(CurrentDate(), "yyyy-MM-dd")
    StrDefaultSQL = " 进入时间 Between To_Date('" & StrDate & " 00:00:00','yyyy-MM-dd hh24:mi:ss') And To_date('" & StrDate & " 23:59:59','yyyy-MM-dd hh24:mi:ss')"
    
    Call RefreshData
    SetListViewRowHeight LvwList.hwnd, 15
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With LvwList
        .Width = ScaleWidth - .Left
        .Height = ScaleHeight - .Top
    End With
    With CmdDelete
        .Left = LvwList.Width - 300 - .Width
    End With
    With CmdView
        .Left = CmdDelete.Left - 150 - .Width
    End With
    With CmdFind
        .Left = PicMain.Left + PicMain.Width + 150
    End With
End Sub

Private Function GetFindSQL() As Boolean
    Dim strDateStart As String, strDateEnd As String
    
    '--根据输入产生对应的查找串--
    GetFindSQL = False
    StrFindSQL = ""
    'Substr(工作站, Instr(工作站, '\') + 1):过滤工作站添加这个条件是为了向上兼容，因为原来的版本记录的工作站信息格式为"工作组\工作站"，现在为"工作站"
    If Cbo工作站.Text <> "" Then StrFindSQL = StrFindSQL & IIf(StrFindSQL = "", " ", " And ") & " Substr(工作站, Instr(工作站, '\') + 1) = '" & Cbo工作站.Text & "'"
    If Cbo用户名.Text <> "" Then StrFindSQL = StrFindSQL & IIf(StrFindSQL = "", " ", " And ") & " 用户名 = '" & Cbo用户名.Text & "'"
    If txt部件名.Text <> "" Then StrFindSQL = StrFindSQL & IIf(StrFindSQL = "", " ", " And ") & " Upper(部件名) Like '" & UCase(txt部件名.Text) & "'"
    If txt工作内容.Text <> "" Then StrFindSQL = StrFindSQL & IIf(StrFindSQL = "", " ", " And ") & " Upper(工作内容) Like '" & UCase(txt工作内容.Text) & "'"
    strDateStart = Format(dtpDateStart, "yyyy-MM-dd")
    strDateEnd = Format(dtpDateEnd, "yyyy-MM-dd")
    StrFindSQL = StrFindSQL & IIf(StrFindSQL = "", " ", " And ") & " 进入时间 Between To_Date('" & strDateStart & " 00:00:00','yyyy-MM-dd hh24:mi:ss') And To_date('" & strDateEnd & " 23:59:59','yyyy-MM-dd hh24:mi:ss')"
    
    GetFindSQL = True
End Function

Private Function InitCons()
    Call ReadInitData(Cbo工作站, Right(Cbo工作站.Name, 3))
    Call ReadInitData(Cbo用户名, Right(Cbo用户名.Name, 3))
    dtpDateStart.value = CurrentDate()
    dtpDateEnd.value = CurrentDate()
End Function

Private Function ReadInitData(ByVal ConObj As Object, ByVal StrColumnName As String)
    Dim RecInit As New ADODB.Recordset
    '--获取初始值--
    
    With ConObj
        .Clear
    End With
    
    If StrColumnName = "工作站" Then
        strSQL = "Select Distinct " & StrColumnName & " As ColumnName From Zlclients"
    Else
        strSQL = "Select Distinct " & StrColumnName & " As ColumnName From 上机人员表"
    End If
    Set RecInit = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    With RecInit
        Do While Not .EOF
            If Not IsNull(!ColumnName) Then
                ConObj.addItem !ColumnName
            End If
            .MoveNext
        Loop
    End With
End Function

Private Function RefreshData()
    '--根据查找串,重新获取数据--
    Set RecLog = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Log", "运行日志", IIf(StrFindSQL = "", StrDefaultSQL, StrFindSQL))
    Call LoadData
End Function

Private Function LoadData()
    Dim lngCount As Long
    Dim ItemThis As ListItem
    '--装数--
    
    LvwList.ListItems.Clear
    With RecLog
        Do While Not .EOF
            If IsNull(!退出时间) Then
                Set ItemThis = LvwList.ListItems.Add(, "K_" & .AbsolutePosition, "正在运行", , 1)
            Else
                If !退出原因 = "正常退出" Then
                    Set ItemThis = LvwList.ListItems.Add(, "K_" & .AbsolutePosition, !退出原因, , 2)
                Else
                    Set ItemThis = LvwList.ListItems.Add(, "K_" & .AbsolutePosition, !退出原因, , 3)
                End If
            End If
            With ItemThis
                .SubItems(1) = IIf(IsNull(RecLog!工作站), "", Mid(RecLog!工作站, InStr(RecLog!工作站, "\") + 1))
                .SubItems(2) = IIf(IsNull(RecLog!用户名), "", RecLog!用户名)
                .SubItems(3) = IIf(IsNull(RecLog!部件名), "", RecLog!部件名)
                .SubItems(4) = IIf(IsNull(RecLog!工作内容), "", RecLog!工作内容)
                .SubItems(5) = IIf(IsNull(RecLog!进入时间), "", RecLog!进入时间)
                .SubItems(6) = IIf(IsNull(RecLog!退出时间), "", RecLog!退出时间)
                .Tag = RecLog!会话号
            End With
            .MoveNext
        Loop
    End With
    With LvwList
        If .ListItems.Count <> 0 Then
            .ListItems(1).Selected = True
            .SelectedItem.Selected = True
        End If
        CmdView.Enabled = (.ListItems.Count <> 0)
        CmdDelete.Enabled = (.ListItems.Count <> 0)
    End With
    If CmdFind.Enabled = False Then
        frmMDIMain.stbThis.Panels(2).Text = "查找完毕，共查找到" & RecLog.RecordCount & "条数据！"
    End If
End Function

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As zlPrintLvw
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "运行日志"
    Set objPrint.Body.objData = LvwList
    objPrint.BelowAppItems.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

'调整listview行高
Private Sub SetListViewRowHeight(ByVal listViewHwnd As Long, ByVal rowHeight As Long)
    Call SetListViewRowHeight_Destroy
    hImageList = ImageList_Create(1, rowHeight, 1, 0, 0)
    SendMessage listViewHwnd, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal hImageList
    SendMessage listViewHwnd, LVM_UPDATE, 0, ByVal 0
End Sub

Private Sub SetListViewRowHeight_Destroy()
    If hImageList <> 0 Then ImageList_Destroy hImageList
End Sub
