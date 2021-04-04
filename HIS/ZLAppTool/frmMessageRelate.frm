VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMessageRelate 
   Caption         =   "相关消息"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   Icon            =   "frmMessageRelate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5310
      TabIndex        =   5
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "打开(&O)"
      Height          =   350
      Left            =   5280
      TabIndex        =   4
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5280
      TabIndex        =   3
      Top             =   1290
      Width           =   1100
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   990
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2850
      Width           =   3000
   End
   Begin RichTextLib.RichTextBox rtfContent 
      Height          =   2265
      Left            =   450
      TabIndex        =   0
      Top             =   2970
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   3995
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMessageRelate.frx":0442
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2205
      Left            =   420
      TabIndex        =   2
      Top             =   330
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_主题"
         Object.Tag             =   "主题"
         Text            =   "主题"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "重要性"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "_发件人"
         Object.Tag             =   "发件人"
         Text            =   "发件人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "_收件人"
         Object.Tag             =   "收件人"
         Text            =   "收件人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "时间"
         Object.Tag             =   "时间"
         Text            =   "时间"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   4320
      Top             =   2490
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
            Picture         =   "frmMessageRelate.frx":04DF
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageRelate.frx":0639
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageRelate.frx":0793
            Key             =   "NewReply"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageRelate.frx":08ED
            Key             =   "ReadReply"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageRelate.frx":0A47
            Key             =   "High"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageRelate.frx":0BA1
            Key             =   "Low"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageRelate.frx":0CFB
            Key             =   "Script"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMessageRelate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sngStartY As Single   '移动前鼠标的位置
Dim mblnItem As Boolean   '为真表示单击到ListView某一项上

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
End Sub

Private Sub cmdOpen_Click()
    frmMessageEdit.OpenWindow Mid(lvwMain.SelectedItem.Key, 3), "", lvwMain.SelectedItem.Tag
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True Then cmdOpen_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal item As MSComctlLib.ListItem)
    mblnItem = True
    Call FillText
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartY = Y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    On Error Resume Next

    If Button = 1 Then
        sngTemp = picSplitH.Top + Y - sngStartY
        If sngTemp - lvwMain.Top > 2500 And ScaleHeight - (sngTemp + picSplitH.Height) > 1200 Then
            picSplitH.Top = sngTemp
            lvwMain.Height = picSplitH.Top - lvwMain.Top
            rtfContent.Top = picSplitH.Top + picSplitH.Height
            rtfContent.Height = ScaleHeight - rtfContent.Top - 60
        End If
        lvwMain.SetFocus
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    
    lvwMain.Left = ScaleLeft
    lvwMain.Top = 60
    If Me.ScaleWidth - 1500 - lvwMain.Left > 0 Then lvwMain.Width = Me.ScaleWidth - 1500 - lvwMain.Left
    
    cmdClose.Left = ScaleWidth - cmdClose.Width - 200
    cmdOpen.Left = cmdClose.Left
    cmdHelp.Left = cmdClose.Left
    
    picSplitH.Left = ScaleLeft
    picSplitH.Top = lvwMain.Top + lvwMain.Height
    picSplitH.Width = ScaleWidth
    
    rtfContent.Left = lvwMain.Left
    rtfContent.Top = picSplitH.Top + picSplitH.Height
    rtfContent.Height = ScaleHeight - rtfContent.Top - 60
    rtfContent.Width = ScaleWidth
End Sub

Public Sub FillList(ByVal strID As String)
'功能:装入相关消息到lvwMain

    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strTemp As String
    Dim strICO As String
    
    On Error GoTo ErrH
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select M.ID,M.会话ID,M.发件人,M.收件人,M.主题,to_char(M.时间,'YYYY-MM-DD HH24:MI:SS') as 时间,S.类型,S.状态" & _
        " from zlMessages M,zlMsgState S where M.ID=S.消息ID and S.删除<>2 and S.用户=[1] and M.会话ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrDbUser, Val(strID))
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "没找到相关消息。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    lvwMain.ListItems.Clear
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("状态")), "0000", rsTemp("状态"))
        
        If rsTemp("类型") = 0 Then
            strICO = "Script"
        Else
            strICO = IIf(Mid(strTemp, 1, 1) = "1", "Read", "New") & IIf(Mid(strTemp, 2, 2) <> "00", "Reply", "")   '已读+已处理
        End If
        Set lst = lvwMain.ListItems.Add(, "C" & rsTemp("类型") & rsTemp("ID"), IIf(IsNull(rsTemp("主题")), "", rsTemp("主题")), strICO, strICO)
        If Mid(strTemp, 4, 1) <> "0" Then
            lst.SubItems(1) = IIf(Mid(strTemp, 4, 1) = 1, "高", "低")
            lst.ListSubItems(1).ReportIcon = IIf(Mid(strTemp, 4, 1) = 1, "High", "Low")
        End If
        lst.SubItems(2) = IIf(IsNull(rsTemp("发件人")), "", rsTemp("发件人"))
        lst.SubItems(3) = IIf(IsNull(rsTemp("收件人")), "", rsTemp("收件人"))
        lst.SubItems(4) = IIf(IsNull(rsTemp("时间")), "", rsTemp("时间"))
        lst.Tag = rsTemp("类型")
        rsTemp.MoveNext
    Loop
    If lvwMain.ListItems.Count > 0 Then
        lvwMain.ListItems(1).Selected = True
    End If
    '统一调用显示文本
    Call FillText
    frmMessageRelate.Show , frmMessageManager
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub FillText()
'功能:把消息的内容装入到RichText中

    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrH
    If lvwMain.SelectedItem Is Nothing Then
        '保留原有键值
        rtfContent.Text = ""
        Exit Sub
    End If
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select 内容,背景色 from zlMessages where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(lvwMain.SelectedItem.Key, 3)))
    
    rtfContent.BackColor = IIf(IsNull(rsTemp("背景色")), RGB(255, 255, 255), rsTemp("背景色"))
    rtfContent.TextRTF = IIf(IsNull(rsTemp("内容")), "", rsTemp("内容"))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

