VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDiagRefer 
   BackColor       =   &H8000000C&
   Caption         =   "诊断参考编辑"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8790
   Icon            =   "frmDiagRefer.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   5970
      MultiLine       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5070
      Visible         =   0   'False
      Width           =   2445
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdRefer 
      Height          =   4935
      Left            =   390
      TabIndex        =   2
      Top             =   1125
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   8705
      _Version        =   393216
      BackColor       =   -2147483628
      Rows            =   10
      Cols            =   7
      FixedCols       =   4
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483628
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8790
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiagRefer.frx":0442
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7858
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "编者：专家姓名"
            TextSave        =   "编者：专家姓名"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8790
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   9705
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "Save"
               Description     =   "预览"
               Object.ToolTipText     =   "保存参考内容"
               Object.Tag             =   "保存"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "恢复"
               Key             =   "Restore"
               Object.ToolTipText     =   "恢复上次保存时内容"
               Object.Tag             =   "恢复"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览参考内容"
               Object.Tag             =   "预览"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印参考内容"
               Object.Tag             =   "打印"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "添加"
               Key             =   "Insert"
               Object.ToolTipText     =   "在当前内容后加入一段"
               Object.Tag             =   "添加"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除本段落"
               Object.Tag             =   "删除"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "措施"
               Key             =   "Method"
               Description     =   "措施"
               Object.ToolTipText     =   "修改本段对应诊疗措施"
               Object.Tag             =   "措施"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Object.ToolTipText     =   "查找单词"
               Object.Tag             =   "查找"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7680
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":0CD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":0EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":1108
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":1322
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":153C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":1C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":2330
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":2A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":2C44
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":2E64
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6915
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":3084
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":329E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":34B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":36D2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":38F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":3FEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":46E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":4DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":4FFA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagRefer.frx":521A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   1965
      Top             =   6495
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "比例尺寸"
      Height          =   180
      Left            =   7245
      TabIndex        =   5
      Top             =   6855
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSaveRefer 
         Caption         =   "保存参考(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "恢复(&R)"
      End
      Begin VB.Menu mnuFileSaveTitle 
         Caption         =   "保存标题(&C)"
      End
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintset 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditRowInsert 
         Caption         =   "添加段落(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditRowDelete 
         Caption         =   "删除段落(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditRowMethod 
         Caption         =   "对应措施(&M)..."
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "查找(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "替换(&R)..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditString 
         Caption         =   "特殊符号(&S)..."
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditTitleInsert 
         Caption         =   "添加标题(&I)..."
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuEditTitleUpdate 
         Caption         =   "修改标题(&U)..."
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditTitleDelete 
         Caption         =   "删除标题(&E)"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuEditSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditProofInsert 
         Caption         =   "添加证候(&B)..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditProofUpdate 
         Caption         =   "修改证候(&G)..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditProofDelete 
         Caption         =   "删除证候(&Y)"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuToolBar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStates 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "字体(&F)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
         WindowList      =   -1  'True
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)..."
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmDiagRefer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mlngBarSize As Long

Dim rsTemp As New ADODB.Recordset
Dim rsMethod As New ADODB.Recordset
Dim strTemp As String
Dim intCount As Integer, lngRow As Integer, lngCol As Integer
Dim blnActive As Boolean

Const conRowHeight As Long = 255
Const conCol项目 As Integer = 0
Const conCol证候 As Integer = 1
Const conCol措施 As Integer = 2
Const conCol标记 As Integer = 3
Const conCol内容 As Integer = 4

Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    If blnActive Then Exit Sub
    If Me.Tag = "西医" Then
        Me.mnuEditSpt3.Visible = False
        Me.mnuEditProofInsert.Visible = False
        Me.mnuEditProofUpdate.Visible = False
        Me.mnuEditProofDelete.Visible = False
    Else
        Me.mnuEditSpt3.Visible = True
        Me.mnuEditProofInsert.Visible = True
        Me.mnuEditProofUpdate.Visible = True
        Me.mnuEditProofDelete.Visible = True
    End If
    Err = 0: On Error GoTo ErrHand

    gstrSql = "select ID,名称,编者" & _
            " from 疾病诊断目录" & _
            " where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdRefer.Tag))
    
    With rsTemp
        Me.Caption = !名称 & "．诊断参考"
        Me.stbThis.Tag = IIf(IsNull(!编者), "", !编者)
        Me.stbThis.Panels(3).Text = "编者：" & Me.stbThis.Tag
    End With
    Call zlGetContent
    Call hgdRefer_RowColChange
    blnActive = True
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    blnActive = False
    Call RestoreWinState(Me, App.ProductName)
    With Me.hgdRefer
        .ColAlignmentFixed(conCol标记) = 3
        .ColAlignment(conCol内容 + 0) = 0
        .ColAlignment(conCol内容 + 1) = 0
        .ColAlignment(conCol内容 + 2) = 0
        .RowHeight(0) = 0
        .ColWidth(conCol项目) = 0
        .ColWidth(conCol证候) = 0
        .ColWidth(conCol措施) = 0
        .ColWidth(conCol标记) = 240
    End With
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If blnActive Then Me.hgdRefer.SetFocus
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    On Error Resume Next
    With Me.hgdRefer
        .Redraw = False
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - lngStatus - .Top
        .ColWidth(conCol内容 + 0) = Me.TextWidth("三个字") + 90
        .ColWidth(conCol内容 + 1) = Me.TextWidth("三个字") + 90
        .ColWidth(conCol内容 + 2) = .Width - .ColWidth(conCol标记) - .ColWidth(conCol内容) - .ColWidth(conCol内容 + 1) - mlngBarSize - 75
        Call zlGrdRowHeight
        .Redraw = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub hgdRefer_DblClick()
    Call hgdRefer_KeyPress(vbKeySpace)
End Sub

Private Sub hgdRefer_KeyPress(KeyAscii As Integer)
    With Me.hgdRefer
        Select Case KeyAscii
        Case vbKeyReturn, vbKeyTab
            .Row = .Row + IIf(.Row = .Rows - 1, 0, 1): Call hgdRefer_RowColChange: Exit Sub
        Case vbKeyDelete
            If .RowData(.Row) = 0 Then Exit Sub
            If .Col - conCol内容 < .RowData(.Row) Then Exit Sub
            If .RowData(.Row) = 1 Then
                .TextMatrix(.Row, conCol内容 + 1) = " "
                .TextMatrix(.Row, conCol内容 + 2) = " "
            Else '.RowData(.Row) = 2 Then
                .TextMatrix(.Row, conCol内容 + 2) = " "
            End If
        Case Else
            If .RowData(.Row) = 0 Then Exit Sub
            If .Col - conCol内容 < .RowData(.Row) Then Exit Sub
            Me.txtInput.Top = .Top + .CellTop
            Me.txtInput.Height = .CellHeight
            If .RowData(.Row) = 1 Then
                Me.txtInput.Left = .Left + .ColWidth(conCol标记) + .ColWidth(conCol内容) + 45
                Me.txtInput.Width = .ColWidth(conCol内容 + 1) + .ColWidth(conCol内容 + 2) - 15
            Else '.RowData(.Row) = 2 Then
                Me.txtInput.Left = .Left + .ColWidth(conCol标记) + .ColWidth(conCol内容) + .ColWidth(conCol内容 + 1) + 45
                Me.txtInput.Width = .ColWidth(conCol内容 + 2) - 15
            End If
            If KeyAscii < 0 _
                Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9") _
                Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") _
                Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
                Me.txtInput.Text = Chr(KeyAscii)
                Me.txtInput.SelStart = Len(Me.txtInput.Text)
            Else
                Me.txtInput.Text = .Text
                Me.txtInput.SelStart = 0
                Me.txtInput.SelLength = 30000
            End If
            Me.txtInput.Visible = True
            Me.mnuEditString.Visible = True
            Me.txtInput.SetFocus
        End Select
    End With
End Sub

Private Sub hgdRefer_KeyUp(KeyCode As Integer, Shift As Integer)
    With Me.hgdRefer
        Select Case KeyCode
        Case vbKeyDelete
            If .RowData(.Row) = 0 Then Exit Sub
            If .Col - conCol内容 < .RowData(.Row) Then Exit Sub
            If .RowData(.Row) = 1 Then
                .TextMatrix(.Row, conCol内容 + 1) = " "
                .TextMatrix(.Row, conCol内容 + 2) = " "
            Else '.RowData(.Row) = 2 Then
                .TextMatrix(.Row, conCol内容 + 2) = " "
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub hgdRefer_RowColChange()
    With Me.hgdRefer
        '内容操作控制
        If .TextMatrix(.Row, conCol证候) <> "" Then
            '辨证子项目不能直接增删段落
            Me.mnuEditRowInsert.Enabled = False
            Me.tlbThis.Buttons("Insert").Enabled = False
            Me.mnuEditRowDelete.Enabled = False
            Me.tlbThis.Buttons("Delete").Enabled = False
        Else
            Me.mnuEditRowInsert.Enabled = True
            Me.tlbThis.Buttons("Insert").Enabled = True
            If .RowData(.Row) = 0 Then
                '项目行不能直接删除
                Me.mnuEditRowDelete.Enabled = False
                Me.tlbThis.Buttons("Delete").Enabled = False
            Else
                Me.mnuEditRowDelete.Enabled = True
                Me.tlbThis.Buttons("Delete").Enabled = True
            End If
        End If
        If .TextMatrix(.Row, conCol标记) <> "" Then
            Me.mnuEditRowMethod.Enabled = True
            Me.tlbThis.Buttons("Method").Enabled = True
        Else
            Me.mnuEditRowMethod.Enabled = False
            Me.tlbThis.Buttons("Method").Enabled = False
        End If
        
        '标题操作控制
        If .TextMatrix(.Row, conCol项目) = "" Then
            '处于辨证行时，不能进行标题操作
            Me.mnuEditTitleInsert.Enabled = False
            Me.mnuEditTitleUpdate.Enabled = False
            Me.mnuEditTitleDelete.Enabled = False
        Else
            Me.mnuEditTitleInsert.Enabled = True
            Me.mnuEditTitleUpdate.Enabled = True
            Me.mnuEditTitleDelete.Enabled = True
        End If
        
        '证候操作控制
        If .TextMatrix(.Row, conCol证候) = "" Then
            '处于辨证行时，不能进行标题操作
            Me.mnuEditProofInsert.Enabled = False
            Me.mnuEditProofUpdate.Enabled = False
            Me.mnuEditProofDelete.Enabled = False
        Else
            Me.mnuEditProofInsert.Enabled = True
            Me.mnuEditProofUpdate.Enabled = True
            Me.mnuEditProofDelete.Enabled = True
        End If
       
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        If .RowData(.Row) = 0 Then Exit Sub
        If .RowData(.Row) = 1 Then
            .Col = conCol内容 + 1
        Else
            .Col = conCol内容 + 2
        End If
    End With
End Sub

Private Sub mnuEditFind_Click()
    With frmDiagRefFind
        Set .frmParent = Me
        .Tag = "查找"
        .Show , Me
    End With
End Sub

Private Sub mnuEditProofDelete_Click()
    Dim strProof As String   '当前项目证候
    Dim lngCurRow As Long
    
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        strProof = .TextMatrix(.Row, conCol证候)
        '删除检测
        If strProof = Mid(.TextMatrix(0, conCol证候), 2) Then
            MsgBox "要求参考至少保留一个证候段，不能删除", vbExclamation, gstrSysName
            Exit Sub
        End If
        If MsgBox("真的删除“" & Split(strProof, ",")(2) & "”证候段吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '删除操作
        For lngCurRow = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(lngCurRow, conCol证候) = strProof Then
                .Row = lngCurRow
                Call mnuEditRowDelete_Click
            End If
        Next
        .TextMatrix(0, conCol证候) = Split(.TextMatrix(0, conCol证候), ";" & strProof)(0) & Split(.TextMatrix(0, conCol证候), ";" & strProof)(1)
    End With
End Sub

Private Sub mnuEditProofInsert_Click()
    Dim strLefts As String   '已经存在的前面的证候
    Dim strRights As String  '已经存在的后面的证候
    Dim strProof As String   '当前项目证候
    Dim aryRows() As String
    
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        strProof = .TextMatrix(.Row, conCol证候)
        strLefts = Split(.TextMatrix(0, conCol证候), ";" & strProof)(0) & ";" & strProof
        strRights = Split(.TextMatrix(0, conCol证候), ";" & strProof)(1)
    End With
    
    '---------------------------------------------
    '调用证候设置窗体，获得证候
    With frmDiagProof
        .Tag = 0  '证候序号
        .txtName.Tag = ""   '证候ID
        .strLefts = strLefts
        .strRights = strRights
        .Show 1, Me
        strProof = .strProof
        Unload frmDiagProof
    End With
    '取消增加，退出处理
    If strProof = "" Then Exit Sub
    
    '---------------------------------------------
    '分情况进行增加处理
    Dim strOldProof As String       '依据增加的证候
    
    With Me.hgdRefer
        strOldProof = .TextMatrix(.Row, conCol证候)
        .TextMatrix(0, conCol证候) = strLefts & ";" & strProof & strRights
        
        '找到该项目第末行，增加一段，将段更改为证候
        For lngRow = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(lngRow, conCol证候) = strOldProof Then Exit For
        Next
        .Row = lngRow
        Call mnuEditRowInsert_Click
        
        .Row = .Row + 1
        .TextMatrix(.Row, conCol项目) = ""
        .TextMatrix(.Row, conCol证候) = strProof
        .TextMatrix(.Row, conCol措施) = ""
        .TextMatrix(.Row, conCol标记) = ""
        .TextMatrix(.Row, conCol内容 + 0) = ""
        .TextMatrix(.Row, conCol内容 + 1) = "［" & Split(strProof, ",")(2) & "］"
        .TextMatrix(.Row, conCol内容 + 2) = "［" & Split(strProof, ",")(2) & "］"
        .MergeRow(.Row) = True
        .RowData(.Row) = 0
        
        '查找二级辨证项，每个增加一行，并填写标题
        aryRows = Split(Mid(.TextMatrix(0, conCol项目), 2), ";")
        For intCount = LBound(aryRows) To UBound(aryRows)
            If Split(aryRows(intCount), ",")(3) = 1 And Split(aryRows(intCount), ",")(4) = 2 Then
                Call mnuEditRowInsert_Click
                .Row = .Row + 1
                .TextMatrix(.Row, conCol项目) = aryRows(intCount)
                .TextMatrix(.Row, conCol证候) = strProof
                .TextMatrix(.Row, conCol措施) = ""
                If Split(aryRows(intCount), ",")(5) = 1 Then
                    .TextMatrix(.Row, conCol标记) = "○"
                Else
                    .TextMatrix(.Row, conCol标记) = ""
                End If
                .TextMatrix(.Row, conCol内容 + 0) = ""
                .TextMatrix(.Row, conCol内容 + 1) = Split(aryRows(intCount), ",")(2) & "："
                .TextMatrix(.Row, conCol内容 + 2) = " "
                .MergeRow(.Row) = False
                .RowData(.Row) = 2
            End If
        Next
    End With

End Sub

Private Sub mnuEditProofUpdate_Click()
    Dim strLefts As String   '已经存在的前面的证候
    Dim strRights As String  '已经存在的后面的证候
    Dim strProof As String   '当前项目证候
    Dim aryRows() As String
    
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        strProof = .TextMatrix(.Row, conCol证候)
        strLefts = Split(.TextMatrix(0, conCol证候), ";" & strProof)(0)
        strRights = Split(.TextMatrix(0, conCol证候), ";" & strProof)(1)
    End With
    
    '---------------------------------------------
    '调用证候设置窗体，获得证候
    With frmDiagProof
        .Tag = Val(Split(strProof, ",")(1))  '证候序号
        .txtName.Tag = Split(strProof, ",")(0)   '证候ID
        .txtName.Text = Split(strProof, ",")(2)
        .strLefts = strLefts
        .strRights = strRights
        .Show 1, Me
        strProof = .strProof
        Unload frmDiagProof
    End With
    '取消增加，退出处理
    If strProof = "" Then Exit Sub
    
    '---------------------------------------------
    '分情况进行增加处理
    Dim strOldProof As String       '依据增加的证候
    With Me.hgdRefer
        strOldProof = .TextMatrix(.Row, conCol证候)
        .TextMatrix(0, conCol证候) = strLefts & ";" & strProof & strRights
        For lngRow = .FixedRows To .Rows - 1
            If .TextMatrix(lngRow, conCol证候) = strOldProof Then
                .TextMatrix(lngRow, conCol证候) = strProof
                If .TextMatrix(lngRow, conCol项目) = "" Then
                    .TextMatrix(lngRow, conCol内容 + 0) = ""
                    .TextMatrix(lngRow, conCol内容 + 1) = "［" & Split(strProof, ",")(2) & "］"
                    .TextMatrix(lngRow, conCol内容 + 2) = "［" & Split(strProof, ",")(2) & "］"
                End If
            End If
        Next
    End With
End Sub

Private Sub mnuEditReplace_Click()
    With frmDiagRefFind
        Set .frmParent = Me
        .Tag = "替换"
        .Show , Me
    End With
End Sub

Private Sub mnuEditRowDelete_Click()
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        .Redraw = False
        For lngRow = .Row To .Rows - 2
            For lngCol = 0 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow + 1, lngCol)
            Next
            .MergeRow(lngRow) = .MergeRow(lngRow + 1)
            .RowHeight(lngRow) = .RowHeight(lngRow + 1)
            .RowData(lngRow) = .RowData(lngRow + 1)
        Next
        .RowData(.Rows - 1) = 0
        .Rows = .Rows - 1
        .Redraw = True
    End With
    Call hgdRefer_RowColChange
End Sub

Private Sub mnuEditRowInsert_Click()
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        .Redraw = False
        .Rows = .Rows + 1
        For lngRow = .Rows - 1 To .Row + 1 Step -1
            For lngCol = 0 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow - 1, lngCol)
            Next
            .MergeRow(lngRow) = .MergeRow(lngRow - 1)
            .RowHeight(lngRow) = .RowHeight(lngRow - 1)
            .RowData(lngRow) = .RowData(lngRow - 1)
        Next
        .TextMatrix(.Row + 1, conCol项目) = .TextMatrix(.Row, conCol项目)
        .TextMatrix(.Row + 1, conCol证候) = .TextMatrix(.Row, conCol证候)
        If .TextMatrix(.Row, conCol项目) <> "" Then
            If Split(.TextMatrix(.Row, conCol项目), ",")(5) = 1 Then
                .TextMatrix(.Row + 1, conCol标记) = "○"
            Else
                .TextMatrix(.Row + 1, conCol标记) = ""
            End If
            .TextMatrix(.Row + 1, conCol内容 + 0) = ""
            If Split(.TextMatrix(.Row, conCol项目), ",")(4) = 1 Then
                .TextMatrix(.Row + 1, conCol内容 + 1) = " "
                .TextMatrix(.Row + 1, conCol内容 + 2) = " "
                .MergeRow(.Row + 1) = True
                .RowData(.Row + 1) = 1
            Else
                .TextMatrix(.Row + 1, conCol内容 + 1) = ""
                .TextMatrix(.Row + 1, conCol内容 + 2) = " "
                .MergeRow(.Row + 1) = False
                .RowData(.Row + 1) = 2
            End If
        Else
            .TextMatrix(.Row + 1, conCol标记) = ""
            .TextMatrix(.Row + 1, conCol内容 + 0) = ""
            .TextMatrix(.Row + 1, conCol内容 + 1) = ""
            .TextMatrix(.Row + 1, conCol内容 + 2) = " "
            .MergeRow(.Row + 1) = False
            .RowData(.Row + 1) = 2
        End If
        .RowHeight(.Row + 1) = conRowHeight
        .Redraw = True
    End With
    Call hgdRefer_RowColChange
End Sub

Private Sub mnuEditRowMethod_Click()
    Dim strMethod As String
    With Me.hgdRefer
        If Trim(.TextMatrix(.Row, conCol标记)) = "" Then Exit Sub
        strMethod = .TextMatrix(.Row, conCol措施)
    End With
    With frmDiagMethod
        .strMethod = strMethod
        .Show 1, Me
        strMethod = .strMethod
        Unload frmDiagMethod
    End With
    With Me.hgdRefer
        .TextMatrix(.Row, conCol措施) = strMethod
        If strMethod = "" Then
            .TextMatrix(.Row, conCol标记) = "○"
        Else
            .TextMatrix(.Row, conCol标记) = "●"
        End If
    End With
End Sub

Private Sub mnuEditString_Click()
    If Me.txtInput.Visible = False Then Exit Sub
    strTemp = ShowSpecChar(Me)
    With Me.txtInput
        intCount = .SelStart
        .Text = Left(.Text, .SelStart) & strTemp & Mid(.Text, .SelStart + .SelLength + 1)
        .SelStart = intCount + Len(strTemp)
        DoEvents
        .Visible = True
        .SetFocus
        Me.mnuEditString.Visible = True
    End With
End Sub

Private Sub mnuEditTitleDelete_Click()
    Dim strTitle As String   '当前项目标题
    Dim lngCurRow As Long
    
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        strTitle = .TextMatrix(.Row, conCol项目)
        '删除检测
        If strTitle = Mid(.TextMatrix(0, conCol项目), 2) Then
            MsgBox "要求参考至少保留一个标题段，不能删除", vbExclamation, gstrSysName
            Exit Sub
        End If
        If Split(Mid(strTitle, 2), ",")(3) = 1 And Split(Mid(strTitle, 2), ",")(4) = 1 Then
            '如果删除1级辨证项，首先检查是否有2级辨证项
            For lngRow = .FixedRows To .Rows - 1
                If .TextMatrix(lngRow, conCol项目) <> "" Then
                    If Split(Mid(.TextMatrix(lngRow, conCol项目), 2), ",")(3) = 1 And Split(Mid(.TextMatrix(lngRow, conCol项目), 2), ",")(4) = 2 Then
                        MsgBox "该标题段还存在2级辨证段(如：" & Split(Mid(.TextMatrix(lngRow, conCol项目), 2), ",")(2) & ")，不能删除！", vbExclamation, gstrSysName
                        Exit Sub
                    End If
                End If
            Next
        End If
        If MsgBox("真的删除“" & Split(strTitle, ",")(2) & "”标题段吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '删除操作
        For lngCurRow = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(lngCurRow, conCol项目) = strTitle Then
                .Row = lngCurRow
                Call mnuEditRowDelete_Click
            End If
        Next
        .TextMatrix(0, conCol项目) = Split(.TextMatrix(0, conCol项目), ";" & strTitle)(0) & Split(.TextMatrix(0, conCol项目), ";" & strTitle)(1)
        
        If Split(Mid(strTitle, 2), ",")(3) = 1 And Split(Mid(strTitle, 2), ",")(4) = 2 Then
            '如果删除2级辨证项，检查是否还有2级辨证项，如果没有则删除所有辨证
            For lngRow = .FixedRows To .Rows - 1
                If .TextMatrix(lngRow, conCol项目) <> "" Then
                    If Split(Mid(.TextMatrix(lngRow, conCol项目), 2), ",")(3) = 1 And Split(Mid(.TextMatrix(lngRow, conCol项目), 2), ",")(4) = 2 Then
                        Exit Sub
                    End If
                End If
            Next
            For lngCurRow = .Rows - 1 To .FixedRows Step -1
                If .TextMatrix(lngCurRow, conCol证候) <> "" Then
                    .Row = lngCurRow
                    Call mnuEditRowDelete_Click
                End If
            Next
            .TextMatrix(0, conCol证候) = ""
        End If
    
    End With

End Sub

Private Sub mnuEditTitleInsert_Click()
    Dim strLefts As String   '已经存在的前面的标题
    Dim strRights As String  '已经存在的后面的标题
    Dim strTitle As String   '当前项目标题
    Dim aryRows() As String, blnProof As Boolean
    
    Me.hgdRefer.SetFocus
    
    '查找项目中是否已经包含辨证项，决定是否允许插入辨证项
    With Me.hgdRefer
        blnProof = False
        aryRows = Split(Mid(.TextMatrix(0, conCol项目), 2), ";")
        For intCount = LBound(aryRows) To UBound(aryRows)
            If Split(aryRows(intCount), ",")(3) = 1 Then blnProof = True: Exit For
        Next
        strTitle = .TextMatrix(.Row, conCol项目)
        strLefts = Split(.TextMatrix(0, conCol项目), ";" & strTitle)(0) & ";" & strTitle
        strRights = Split(.TextMatrix(0, conCol项目), ";" & strTitle)(1)
    End With
    
    '---------------------------------------------
    '调用标题设置窗体，获得标题
    With frmDiagTitle
        .Tag = "0,0,"  '新项目的主序号、子序号为0
        .lblKind.Caption = Me.Tag
        If Split(strTitle, ",")(4) = 1 Then
            .optTier(0).Value = True
            .optTier(1).Value = False
        Else
            .optTier(0).Value = False
            .optTier(1).Value = True
        End If
        If .lblKind.Caption = "西医" Then
            .chkProof.Value = 0
            .chkProof.Enabled = False
        Else
            If Split(strTitle, ",")(3) = 1 Then
                '如果当前项目标题为辨证项，则在其后插入的必然为辨证项，且层次为2
                .chkProof.Value = 1
                .chkProof.Enabled = False
                .optTier(1).Value = True
                .optTier(0).Enabled = False
                .optTier(1).Enabled = False
            ElseIf blnProof Then
                '如果已经存在辨证项，则不能再增加辨证项
                .chkProof.Value = 0
                .chkProof.Enabled = False
            Else
                .chkProof.Enabled = True
            End If
        End If
        .chkMethod.Value = Split(strTitle, ",")(5)
        .strLefts = strLefts
        .strRights = strRights
        .Show 1, Me
        strTitle = .strTitle
        Unload frmDiagTitle
    End With
    '取消增加，退出处理
    If strTitle = "" Then Exit Sub
    
    '---------------------------------------------
    '分情况进行增加处理
    Dim strFromItem As String       '依据增加的项目
    
    With Me.hgdRefer
        strFromItem = .TextMatrix(.Row, conCol项目)
        .TextMatrix(0, conCol项目) = strLefts & ";" & strTitle & strRights
        
        If Split(strTitle, ",")(3) <> 1 Or Split(strTitle, ",")(3) = 1 And Split(strTitle, ",")(4) = 1 Then
            '独立项，或辨证项但层次为1，找到该项目第末行，增加一段，将段更改为项目
            For lngRow = .Rows - 1 To .FixedRows Step -1
                If .TextMatrix(lngRow, conCol项目) = strFromItem Then Exit For
            Next
            .Row = lngRow
            Call mnuEditRowInsert_Click
            
            .Row = .Row + 1
            .TextMatrix(.Row, conCol项目) = strTitle
            .TextMatrix(.Row, conCol措施) = ""
            .TextMatrix(.Row, conCol标记) = ""
            If Split(strTitle, ",")(4) = 1 Then
                .TextMatrix(.Row, conCol内容 + 0) = "【" & Split(strTitle, ",")(2) & "】"
                .TextMatrix(.Row, conCol内容 + 1) = "【" & Split(strTitle, ",")(2) & "】"
                .TextMatrix(.Row, conCol内容 + 2) = "【" & Split(strTitle, ",")(2) & "】"
            Else
                .TextMatrix(.Row, conCol内容 + 0) = ""
                .TextMatrix(.Row, conCol内容 + 1) = "〖" & Split(strTitle, ",")(2) & "〗"
                .TextMatrix(.Row, conCol内容 + 2) = "〖" & Split(strTitle, ",")(2) & "〗"
            End If
            .MergeRow(.Row) = True
            .RowData(.Row) = 0
        Else
            '辨证项，层次为2，需要分如下情况处理：
            If .TextMatrix(0, conCol证候) = "" Then
                '1、如果无证候记录，说明为第一个2级辨证项，找到该项目第末行；增加一段，填写∴∵证候；再加一段，将段更改为项目；
                For lngRow = .Rows - 1 To .FixedRows Step -1
                    If .TextMatrix(lngRow, conCol项目) = strFromItem Then Exit For
                Next
                .Row = lngRow
                Call mnuEditRowInsert_Click
                
                .Row = .Row + 1
                .TextMatrix(.Row, conCol证候) = "0,1,∴∵证候"
                .TextMatrix(0, conCol证候) = .TextMatrix(0, conCol证候) & ";" & .TextMatrix(.Row, conCol证候)
                .TextMatrix(.Row, conCol项目) = ""
                .TextMatrix(.Row, conCol内容 + 0) = ""
                .TextMatrix(.Row, conCol内容 + 1) = "［∴∵证候］"
                .TextMatrix(.Row, conCol内容 + 2) = "［∴∵证候］"
                .MergeRow(.Row) = True
                .RowData(.Row) = 0
                
                Call mnuEditRowInsert_Click
                .Row = .Row + 1
                .TextMatrix(.Row, conCol项目) = strTitle
                .TextMatrix(.Row, conCol措施) = ""
                If Split(strTitle, ",")(5) = 1 Then
                    .TextMatrix(.Row, conCol标记) = "○"
                Else
                    .TextMatrix(.Row, conCol标记) = ""
                End If
                .TextMatrix(.Row, conCol内容 + 1) = Split(strTitle, ",")(2) & "："
                .MergeRow(.Row) = False
                .RowData(.Row) = 2
            Else
                '2、如果有证候记录，且根据1级还是2级辨证项增加，逐个证候查找对比增加
                aryRows = Split(Mid(.TextMatrix(0, conCol证候), 2), ";")
                For intCount = LBound(aryRows) To UBound(aryRows)
                    For lngRow = .FixedRows To .Rows - 1
                        If Split(strFromItem, ",")(4) = 1 Then
                            If .TextMatrix(lngRow, conCol项目) = "" And .TextMatrix(lngRow, conCol证候) = aryRows(intCount) Then Exit For
                        Else
                            If .TextMatrix(lngRow, conCol项目) = strFromItem And .TextMatrix(lngRow, conCol证候) = aryRows(intCount) Then Exit For
                        End If
                    Next
                    .Row = lngRow
                    Call mnuEditRowInsert_Click
                    .Row = .Row + 1
                    .TextMatrix(.Row, conCol项目) = strTitle
                    .TextMatrix(.Row, conCol措施) = ""
                    If Split(strTitle, ",")(5) = 1 Then
                        .TextMatrix(.Row, conCol标记) = "○"
                    Else
                        .TextMatrix(.Row, conCol标记) = ""
                    End If
                    .TextMatrix(.Row, conCol内容 + 1) = Split(strTitle, ",")(2) & "："
                    .MergeRow(.Row) = False
                    .RowData(.Row) = 2
                Next
            End If
            
        End If
    End With
End Sub

Private Sub mnuEditTitleUpdate_Click()
    Dim strLefts As String   '已经存在的前面的标题
    Dim strRights As String  '已经存在的后面的标题
    Dim strTitle As String   '当前项目标题
    Dim aryRows() As String, blnProof As Boolean
    
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        strTitle = .TextMatrix(.Row, conCol项目)
        strLefts = Split(.TextMatrix(0, conCol项目), ";" & strTitle)(0)
        strRights = Split(.TextMatrix(0, conCol项目), ";" & strTitle)(1)
    End With
    
    '---------------------------------------------
    '调用标题设置窗体，获得标题
    With frmDiagTitle
        .lblKind.Caption = Me.Tag
        .Tag = Val(Split(strTitle, ",")(0)) & "," & Val(Split(strTitle, ",")(1)) & ","
        .txtName.Text = Split(strTitle, ",")(2)
        If Split(strTitle, ",")(4) = 1 Then
            .optTier(0).Value = True
            .optTier(1).Value = False
        Else
            .optTier(0).Value = False
            .optTier(1).Value = True
        End If
        .optTier(0).Enabled = False
        .optTier(1).Enabled = False
        .chkProof.Value = Split(strTitle, ",")(3)
        .chkProof.Enabled = False
        .chkMethod.Value = Split(strTitle, ",")(5)
        .strLefts = strLefts
        .strRights = strRights
        .Show 1, Me
        strTitle = .strTitle
        Unload frmDiagTitle
    End With
    '取消修改，退出处理
    If strTitle = "" Then Exit Sub
    
    '---------------------------------------------
    '分情况进行修改处理
    Dim strFromItem As String       '被修改的项目
    With Me.hgdRefer
        strFromItem = .TextMatrix(.Row, conCol项目)
        .TextMatrix(0, conCol项目) = strLefts & ";" & strTitle & strRights
        For lngRow = .FixedRows To .Rows - 1
            If .TextMatrix(lngRow, conCol项目) = strFromItem Then
                .TextMatrix(lngRow, conCol项目) = strTitle
                If Split(strTitle, ",")(5) = 1 And .RowData(lngRow) <> 0 Then
                    .TextMatrix(lngRow, conCol标记) = "○"
                Else
                    .TextMatrix(lngRow, conCol标记) = ""
                    .TextMatrix(lngRow, conCol措施) = ""
                End If
                If .TextMatrix(lngRow, conCol证候) = "" Then
                    If .RowData(lngRow) = 0 Then
                        If Split(strTitle, ",")(4) = 1 Then
                            .TextMatrix(lngRow, conCol内容 + 0) = "【" & Split(strTitle, ",")(2) & "】"
                            .TextMatrix(lngRow, conCol内容 + 1) = "【" & Split(strTitle, ",")(2) & "】"
                            .TextMatrix(lngRow, conCol内容 + 2) = "【" & Split(strTitle, ",")(2) & "】"
                        Else
                            .TextMatrix(lngRow, conCol内容 + 0) = ""
                            .TextMatrix(lngRow, conCol内容 + 1) = "〖" & Split(strTitle, ",")(2) & "〗"
                            .TextMatrix(lngRow, conCol内容 + 2) = "〖" & Split(strTitle, ",")(2) & "〗"
                        End If
                    End If
                Else
                    .TextMatrix(lngRow, conCol内容 + 1) = Split(strTitle, ",")(2) & "："
                End If
            End If
        Next
    End With
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call zlRptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileRestore_Click()
    If MsgBox("如果恢复上次保存的内容，刚才进行修改将无效" & vbCrLf & "真的恢复吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call zlGetContent
    Call hgdRefer_RowColChange
End Sub

Private Sub mnuFileSaveRefer_Click()
    Dim intOdd As Integer, intShowChars As Integer
    Dim strUpItem As String, strUpProof As String, strContent As String
    
    Me.hgdRefer.SetFocus
    
    '项目编排整理
    Call zlGrdSortItems
    
    '证候编排整理
    If Me.hgdRefer.TextMatrix(0, conCol证候) <> "" Then
        Call zlGrdSortProofs
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    gcnOracle.BeginTrans
    With Me.hgdRefer
        gstrSql = "zl_疾病诊断参考_Delete(" & .Tag & ")"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        intShowChars = Int(Me.stbThis.Panels(2).Width / Me.TextWidth("■"))
        intOdd = 0: strUpItem = "-": strUpProof = "-"
        For lngRow = .FixedRows To .Rows - 1
            Me.stbThis.Panels(2).Text = String(intShowChars * lngRow / .Rows, "■")
            If .TextMatrix(lngRow, conCol证候) <> strUpProof Then
                intOdd = 0
            ElseIf .TextMatrix(lngRow, conCol证候) = "" And .TextMatrix(lngRow, conCol项目) <> strUpItem Then
                intOdd = 0
            End If
            If .TextMatrix(lngRow, conCol项目) <> "" Then
                gstrSql = "zl_疾病诊断参考_Insert(" & .Tag & "," & _
                        Split(.TextMatrix(lngRow, conCol项目), ",")(0) & "," & _
                        "'" & Split(.TextMatrix(lngRow, conCol项目), ",")(2) & "'," & _
                        Split(.TextMatrix(lngRow, conCol项目), ",")(3) & "," & _
                        Split(.TextMatrix(lngRow, conCol项目), ",")(4) & "," & _
                        Split(.TextMatrix(lngRow, conCol项目), ",")(5) & ","
                If .TextMatrix(lngRow, conCol证候) = "" Then
                    gstrSql = gstrSql & "null,null,null,"
                ElseIf Val(Split(.TextMatrix(lngRow, conCol证候), ",")(0)) = 0 Then
                    If InStr(1, Split(.TextMatrix(lngRow, conCol证候), ",")(2), "∴∵") > 0 Then
                        gcnOracle.RollbackTrans
                        MsgBox "参考中存在不明确的证候：" & Split(.TextMatrix(lngRow, conCol证候), ",")(2) & "，请修改后保存。", vbExclamation, gstrSysName
                        Me.stbThis.Panels(2).Text = ""
                        Exit Sub
                    End If
                    gstrSql = gstrSql & "null," & _
                        Split(.TextMatrix(lngRow, conCol证候), ",")(1) & "," & _
                        "'" & Split(.TextMatrix(lngRow, conCol证候), ",")(2) & "',"
                Else
                    gstrSql = gstrSql & _
                        Split(.TextMatrix(lngRow, conCol证候), ",")(0) & "," & _
                        Split(.TextMatrix(lngRow, conCol证候), ",")(1) & "," & _
                        "'" & Split(.TextMatrix(lngRow, conCol证候), ",")(2) & "',"
                End If
                If .RowData(lngRow) = 0 Then
                    gstrSql = gstrSql & "0,null,"
                Else
                    strContent = Trim(.TextMatrix(lngRow, conCol内容 + 2))
                    strContent = Replace(strContent, vbCrLf, "")
                    strContent = Replace(strContent, vbCr, "")
                    strContent = Replace(strContent, vbLf, "")
                    strContent = Replace(strContent, "'", StrConv("'", vbWide))
                    strContent = Replace(strContent, "&", StrConv("&", vbWide))
                    gstrSql = gstrSql & intOdd & ",'" & strContent & "',"
                End If
                If .TextMatrix(lngRow, conCol措施) = "" Then
                    gstrSql = gstrSql & "null,'" & Trim(Me.stbThis.Tag) & "')"
                Else
                    gstrSql = gstrSql & "'" & .TextMatrix(lngRow, conCol措施) & "','" & Trim(Me.stbThis.Tag) & "')"
                End If
                
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
                
            End If
            intOdd = intOdd + 1
            strUpItem = .TextMatrix(lngRow, conCol项目)
            strUpProof = .TextMatrix(lngRow, conCol证候)
        Next
    End With
    
    gcnOracle.CommitTrans
    Me.stbThis.Panels(2).Text = ""
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuFileSaveTitle_Click()
    If MsgBox("真的保存本文标题作为" & Me.Tag & "缺省参考标题吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    
    Me.hgdRefer.SetFocus
    Call zlGrdSortItems     '项目编排整理
    On Error GoTo ErrHand
    gstrSql = "zl_疾病参考项目_Save(" & IIf(Me.Tag = "西医", 1, 2) & ",'" & Mid(Me.hgdRefer.TextMatrix(0, conCol项目), 2) & "')"
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuViewFont_Click()
    Me.hgdRefer.SetFocus
    With comDlg
        .FontName = Me.Font.Name
        .FontSize = Me.Font.Size
        .FontBold = Me.Font.Bold
        .FontItalic = Me.Font.Italic
        .Flags = cdlCFANSIOnly _
            + cdlCFApply _
            + cdlCFPrinterFonts
        .ShowFont
        Me.Font.Name = .FontName
        Me.Font.Size = .FontSize
        Me.Font.Bold = .FontBold
        Me.Font.Italic = .FontItalic
    End With
    Set Me.txtInput.Font = Me.Font
    Set Me.hgdRefer.Font = Me.Font
    Call Form_Resize
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarStand_Click()
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.clbThis.Visible = Me.mnuViewToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolBarText_Click()
    Dim i As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub stbThis_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 3 Then
        strTemp = InputBox("本参考著述的编者姓名" & vbCrLf & "  (通常应选择权威专家的著述作为参考)", "编者", Me.stbThis.Tag, Me.Left + Me.Width / 2 - 2500, Me.Top + Me.Height / 2)
        If Trim(strTemp) <> "" Then
            Me.stbThis.Tag = Left(strTemp, 10)
            Panel.Text = "编者：" & Left(strTemp, 10)
        End If
    End If
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Save"
        Call mnuFileSaveRefer_Click
    Case "Restore"
        Call mnuFileRestore_Click
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Insert"
        Call mnuEditRowInsert_Click
    Case "Delete"
        Call mnuEditRowDelete_Click
    Case "Method"
        Call mnuEditRowMethod_Click
    Case "Find"
        Call mnuEditFind_Click
    Case "Help"
        Call mnuHelpHelp_Click
    Case "Exit"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuToolBar, 2
End Sub

Private Sub txtInput_Change()
    Dim lngColWidth As Long, lngTxtWidth As Long, intAskRows As Integer
    With Me.hgdRefer
        .Redraw = False
        If .RowData(.Row) = 2 Then
            lngColWidth = .ColWidth(conCol内容 + 2)
            If Trim(Me.txtInput.Text) = "" Then
                .TextMatrix(.Row, conCol内容 + 2) = " "
            Else
                .TextMatrix(.Row, conCol内容 + 2) = Me.txtInput.Text
            End If
        ElseIf .RowData(.Row) = 1 Then
            lngColWidth = .ColWidth(conCol内容 + 1) + .ColWidth(conCol内容 + 2)
            If Trim(Me.txtInput.Text) = "" Then
                .TextMatrix(.Row, conCol内容 + 1) = " "
                .TextMatrix(.Row, conCol内容 + 2) = " "
            Else
                .TextMatrix(.Row, conCol内容 + 1) = Me.txtInput.Text
                .TextMatrix(.Row, conCol内容 + 2) = Me.txtInput.Text
            End If
        End If
        Me.lblScale.Width = lngColWidth - 90
        Me.lblScale.Caption = .TextMatrix(.Row, conCol内容 + 2)
        .RowHeight(.Row) = Me.lblScale.Height + 75
        Me.txtInput.Height = .RowHeight(.Row)
        .Redraw = True
    End With
End Sub

Private Sub txtInput_GotFocus()
    Me.mnuEditString.Visible = True
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Me.hgdRefer.SetFocus
        Call zlCommFun.PressKey(vbKeyReturn)
    End If
End Sub

Private Sub txtInput_LostFocus()
    Me.txtInput.Visible = False
    Me.mnuEditString.Visible = False
End Sub

Private Sub zlGetContent()
    '---------------------------------------------
    '提取参考内容
    '---------------------------------------------
    Err = 0: On Error GoTo ErrHand
    '--------------------------------------------------------
    Me.hgdRefer.Redraw = False
    Me.hgdRefer.Clear
    Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + 1
    
    '如果已经保存有参考内容，则提取显示；
    gstrSql = "select 项目序号,参考项目,nvl(项目格式,0) as 项目格式,项目层次,nvl(内容性质,0) as 内容性质," & _
            "       证候ID,nvl(证候序号,0) as 证候序号,证候名称,nvl(内容行号,0) as 内容行号,nvl(内容文本,'') as 内容文本" & _
            " from 疾病诊断参考" & _
            " where 诊断id=[1] " & _
            " order by 项目序号,证候序号,内容行号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdRefer.Tag))
    
    With rsTemp
        lngRow = 0
        Do While Not .EOF
            lngRow = lngRow + 1
            If lngRow > Me.hgdRefer.Rows - Me.hgdRefer.FixedRows Then Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + lngRow
            
            If !证候序号 = 0 Then
                Me.hgdRefer.TextMatrix(lngRow, conCol项目) = !项目序号 & ",0," & !参考项目 & "," & !项目格式 & "," & !项目层次 & "," & !内容性质
                If InStr(1, Me.hgdRefer.TextMatrix(0, conCol项目), ";" & Me.hgdRefer.TextMatrix(lngRow, conCol项目)) = 0 Then
                    Me.hgdRefer.TextMatrix(0, conCol项目) = Me.hgdRefer.TextMatrix(0, conCol项目) & ";" & Me.hgdRefer.TextMatrix(lngRow, conCol项目)
                    If !项目层次 = 1 Then
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 0) = "【" & !参考项目 & "】"
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = "【" & !参考项目 & "】"
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = "【" & !参考项目 & "】"
                    Else
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 0) = ""
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = "〖" & !参考项目 & "〗"
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = "〖" & !参考项目 & "〗"
                    End If
                    Me.hgdRefer.MergeRow(lngRow) = True
                    Me.hgdRefer.RowData(lngRow) = 0
                    If Trim(!内容文本) <> "" Then
                        lngRow = lngRow + 1
                        If lngRow > Me.hgdRefer.Rows - Me.hgdRefer.FixedRows Then Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + lngRow
                        Me.hgdRefer.TextMatrix(lngRow, conCol项目) = !项目序号 & ",0," & !参考项目 & "," & !项目格式 & "," & !项目层次 & "," & !内容性质
                        If !项目层次 = 1 Then
                            Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = IIf(IsNull(!内容文本), " ", !内容文本)
                            Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = IIf(IsNull(!内容文本), " ", !内容文本)
                            Me.hgdRefer.MergeRow(lngRow) = True
                            Me.hgdRefer.RowData(lngRow) = 1
                        Else
                            Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = ""
                            Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = IIf(IsNull(!内容文本), " ", !内容文本)
                            Me.hgdRefer.MergeRow(lngRow) = False
                            Me.hgdRefer.RowData(lngRow) = 2
                        End If
                    End If
                Else
                    Me.hgdRefer.TextMatrix(lngRow, conCol项目) = !项目序号 & ",0," & !参考项目 & "," & !项目格式 & "," & !项目层次 & "," & !内容性质
                    If !项目层次 = 1 Then
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = IIf(IsNull(!内容文本), " ", !内容文本)
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = IIf(IsNull(!内容文本), " ", !内容文本)
                        Me.hgdRefer.MergeRow(lngRow) = True
                        Me.hgdRefer.RowData(lngRow) = 1
                    Else
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = ""
                        Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = IIf(IsNull(!内容文本), " ", !内容文本)
                        Me.hgdRefer.MergeRow(lngRow) = False
                        Me.hgdRefer.RowData(lngRow) = 2
                    End If
                End If
            Else
                Me.hgdRefer.TextMatrix(lngRow, conCol证候) = IIf(IsNull(!证候ID), "", !证候ID) & "," & !证候序号 & "," & !证候名称
                If InStr(1, Me.hgdRefer.TextMatrix(0, conCol证候), ";" & Me.hgdRefer.TextMatrix(lngRow, conCol证候)) = 0 Then
                    Me.hgdRefer.TextMatrix(0, conCol证候) = Me.hgdRefer.TextMatrix(0, conCol证候) & ";" & Me.hgdRefer.TextMatrix(lngRow, conCol证候)
                    Me.hgdRefer.TextMatrix(lngRow, conCol项目) = ""
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 0) = ""
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = "［" & !证候名称 & "］"
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = "［" & !证候名称 & "］"
                    Me.hgdRefer.MergeRow(lngRow) = True
                    Me.hgdRefer.RowData(lngRow) = 0
                    lngRow = lngRow + 1
                    If lngRow > Me.hgdRefer.Rows - Me.hgdRefer.FixedRows Then Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + lngRow
                End If
                Me.hgdRefer.TextMatrix(lngRow, conCol证候) = IIf(IsNull(!证候ID), "", !证候ID) & "," & !证候序号 & "," & !证候名称
                Me.hgdRefer.TextMatrix(lngRow, conCol项目) = !项目序号 & "," & !内容行号 & "," & !参考项目 & "," & !项目格式 & "," & !项目层次 & "," & !内容性质
                If InStr(1, Me.hgdRefer.TextMatrix(0, conCol项目), ";" & Me.hgdRefer.TextMatrix(lngRow, conCol项目)) = 0 Then
                    Me.hgdRefer.TextMatrix(0, conCol项目) = Me.hgdRefer.TextMatrix(0, conCol项目) & ";" & Me.hgdRefer.TextMatrix(lngRow, conCol项目)
                End If
                Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = !参考项目 & "："
                Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = IIf(IsNull(!内容文本), " ", !内容文本)
                Me.hgdRefer.MergeRow(lngRow) = False
                Me.hgdRefer.RowData(lngRow) = 2
            End If
            If !内容性质 = 1 And Me.hgdRefer.RowData(lngRow) <> 0 Then
                Me.hgdRefer.TextMatrix(lngRow, conCol标记) = "○"
                gstrSql = "select 诊疗项目id " & _
                        " from 疾病诊疗措施" & _
                        " where 诊断id=[1] " & _
                        "       and 参考项目=[2] " & _
                        "       and nvl(内容行号,0)=[3] "
                If IsNull(!证候名称) Then
                    gstrSql = gstrSql & "       and 证候名称 is null "
                Else
                    gstrSql = gstrSql & "       and 证候名称=[4] "
                End If
                strTemp = ""
                Set rsMethod = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdRefer.Tag), CStr("" & rsTemp!参考项目), CLng(Val("" & rsTemp!内容行号)), CStr("" & rsTemp!证候名称))
                    
                With rsMethod
                    Do While Not .EOF
                        strTemp = strTemp & "," & !诊疗项目id
                        .MoveNext
                    Loop
                End With
                If strTemp <> "" Then
                    Me.hgdRefer.TextMatrix(lngRow, conCol措施) = Mid(strTemp, 2)
                    Me.hgdRefer.TextMatrix(lngRow, conCol标记) = "●"
                End If
            Else
                Me.hgdRefer.TextMatrix(lngRow, conCol标记) = ""
            End If
            .MoveNext
        Loop
    End With
    If rsTemp.RecordCount > 0 Then
        Call zlGrdRowHeight
        Me.hgdRefer.Redraw = True
        Exit Sub
    End If
    
    '如果没有编辑过参考，则按照缺省的项目组织参考格式；
    gstrSql = "select 主序号,nvl(子序号,0) as 子序号,名称,nvl(格式,0) as 格式,nvl(层次,1) as 层次,nvl(性质,0) as 性质" & _
            " from 疾病参考项目" & _
            " where 类别=[1] " & _
            " order by 主序号,子序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "西医", 1, 2))
    
    With rsTemp
        lngRow = 0
        Do While Not .EOF
            lngRow = lngRow + 1
            If lngRow > Me.hgdRefer.Rows - Me.hgdRefer.FixedRows Then Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + lngRow
            
            If !子序号 = 0 Then
                Me.hgdRefer.TextMatrix(lngRow, conCol项目) = !主序号 & ",0," & !名称 & "," & !格式 & "," & !层次 & "," & !性质
                Me.hgdRefer.TextMatrix(0, conCol项目) = Me.hgdRefer.TextMatrix(0, conCol项目) & ";" & Me.hgdRefer.TextMatrix(lngRow, conCol项目)
                If !层次 = 1 Then
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 0) = "【" & !名称 & "】"
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = "【" & !名称 & "】"
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = "【" & !名称 & "】"
                Else
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 0) = ""
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = "〖" & !名称 & "〗"
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = "〖" & !名称 & "〗"
                End If
                Me.hgdRefer.MergeRow(lngRow) = True
                Me.hgdRefer.RowData(lngRow) = 0
            Else
                Me.hgdRefer.TextMatrix(lngRow, conCol证候) = "0,1,∴∵证候"
                If InStr(1, Me.hgdRefer.TextMatrix(0, conCol证候), ";" & Me.hgdRefer.TextMatrix(lngRow, conCol证候)) = 0 Then
                    Me.hgdRefer.TextMatrix(0, conCol证候) = Me.hgdRefer.TextMatrix(0, conCol证候) & ";" & Me.hgdRefer.TextMatrix(lngRow, conCol证候)
                    Me.hgdRefer.TextMatrix(lngRow, conCol项目) = ""
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 0) = ""
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = "［∴∵证候］"
                    Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 2) = "［∴∵证候］"
                    Me.hgdRefer.MergeRow(lngRow) = True
                    Me.hgdRefer.RowData(lngRow) = 0
                    lngRow = lngRow + 1
                    If lngRow > Me.hgdRefer.Rows - Me.hgdRefer.FixedRows Then Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + lngRow
                End If
                Me.hgdRefer.TextMatrix(lngRow, conCol证候) = "0,1,∴∵证候"
                Me.hgdRefer.TextMatrix(lngRow, conCol项目) = !主序号 & "," & !子序号 & "," & !名称 & "," & !格式 & "," & !层次 & "," & !性质
                If InStr(1, Me.hgdRefer.TextMatrix(0, conCol项目), ";" & Me.hgdRefer.TextMatrix(lngRow, conCol项目)) = 0 Then
                    Me.hgdRefer.TextMatrix(0, conCol项目) = Me.hgdRefer.TextMatrix(0, conCol项目) & ";" & Me.hgdRefer.TextMatrix(lngRow, conCol项目)
                End If
                Me.hgdRefer.TextMatrix(lngRow, conCol内容 + 1) = !名称 & "："
                Me.hgdRefer.MergeRow(lngRow) = False
                Me.hgdRefer.RowData(lngRow) = 2
            End If
            .MoveNext
        Loop
    End With
    
    Call zlGrdRowHeight
    Me.hgdRefer.Redraw = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog

End Sub

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '根据调整内容调整内容网格的行高度，以保证内容的正常显示
    '---------------------------------------------
    Dim lngColWidth As Long
    With Me.hgdRefer
        For lngRow = .FixedRows To .Rows - 1
            Select Case .RowData(lngRow)
            Case 0
                lngColWidth = .ColWidth(conCol内容 + 2)
                If .TextMatrix(lngRow, conCol内容 + 2) = .TextMatrix(lngRow, conCol内容 + 1) Then
                    lngColWidth = lngColWidth + .ColWidth(conCol内容 + 1)
                    If .TextMatrix(lngRow, conCol内容 + 1) = .TextMatrix(lngRow, conCol内容) Then
                        lngColWidth = lngColWidth + .ColWidth(conCol内容)
                    End If
                End If
            Case 1
                lngColWidth = .ColWidth(conCol内容 + 1) + .ColWidth(conCol内容 + 2)
            Case 2
                lngColWidth = .ColWidth(conCol内容 + 2)
            End Select
            Me.lblScale.Width = lngColWidth - 90
            Me.lblScale.Caption = .TextMatrix(lngRow, conCol内容 + 2)
            .RowHeight(lngRow) = Me.lblScale.Height + 75
        Next
    End With
End Sub

Private Sub zlGrdSortItems()
    '---------------------------------------------
    '编排整理设置的标题项目，以便保存
    '---------------------------------------------
    Dim aryRows() As String, aryFlds() As String, strNewRows As String
    Dim intPNO As Integer, intSNO As String
    Dim bytFormat As Byte   '上一项目是否为辨证项
    
    aryRows = Split(Mid(Me.hgdRefer.TextMatrix(0, conCol项目), 2), ";")
    intPNO = 0: intSNO = 0: bytFormat = 0
    For intCount = LBound(aryRows) To UBound(aryRows)
        aryFlds = Split(aryRows(intCount), ",")
        If bytFormat = 1 And aryFlds(4) = 2 Then
            intSNO = intSNO + 1
        Else
            intPNO = intPNO + 1: intSNO = 0
        End If
        bytFormat = aryFlds(3)
        aryFlds(0) = intPNO: aryFlds(1) = intSNO
        strNewRows = Join(aryFlds, ",")
        
        '按照新的项目修改表中项目单元的内容
        With Me.hgdRefer
            For lngRow = .FixedRows To .Rows - 1
                If .TextMatrix(lngRow, conCol项目) = aryRows(intCount) Then
                    .TextMatrix(lngRow, conCol项目) = strNewRows
                End If
            Next
        End With
        aryRows(intCount) = strNewRows
    Next
    Me.hgdRefer.TextMatrix(0, conCol项目) = ";" & Join(aryRows, ";")
End Sub

Private Sub zlGrdSortProofs()
    '---------------------------------------------
    '编排整理设置的证候，以便保存
    '---------------------------------------------
    Dim aryRows() As String, aryFlds() As String, strNewRows As String
    
    If Me.hgdRefer.TextMatrix(0, conCol证候) = "" Then Exit Sub
    aryRows = Split(Mid(Me.hgdRefer.TextMatrix(0, conCol证候), 2), ";")
    For intCount = LBound(aryRows) To UBound(aryRows)
        aryFlds = Split(aryRows(intCount), ",")
        aryFlds(1) = intCount + 1
        strNewRows = Join(aryFlds, ",")
        
        '按照新的证候修改表中证候单元的内容
        With Me.hgdRefer
            For lngRow = .FixedRows To .Rows - 1
                If .TextMatrix(lngRow, conCol证候) = aryRows(intCount) Then
                    .TextMatrix(lngRow, conCol证候) = strNewRows
                End If
            Next
        End With
        aryRows(intCount) = strNewRows
    Next
    Me.hgdRefer.TextMatrix(0, conCol证候) = ";" & Join(aryRows, ";")
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    Set objPrint.Body = Me.hgdRefer
    With objPrint.Title
        .Text = Me.Caption
        .Font.Size = Me.Font.Size + 2
    End With
    objRow.Add ""
    objRow.Add "(" & Me.stbThis.Tag & ")"
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode = 0 Then Exit Sub
    End If
    Call zlPrintOrView1Grd(objPrint, bytMode)
End Sub

Public Function zlWordSelect(lngCurRow As Long, strWord As String) As Long
    '-------------------------------------------------
    '功能:在指定行内容中选中指定的文本
    '入参:  lngCurRow-指定行；strWord-指定选中的文本
    '返回:  未查找到，返回0；查找到则返回该文本的下一个位置
    '-------------------------------------------------
    Me.txtInput.Visible = False
    Me.hgdRefer.Row = lngCurRow
    Me.hgdRefer.Col = conCol内容 + 2
    With Me.hgdRefer
        If .RowData(.Row) = 0 Then zlWordSelect = 0: Exit Function
        Me.txtInput.Top = .Top + .CellTop
        Me.txtInput.Height = .CellHeight
        If .RowData(.Row) = 1 Then
            Me.txtInput.Left = .Left + .ColWidth(conCol标记) + .ColWidth(conCol内容) + 45
            Me.txtInput.Width = .ColWidth(conCol内容 + 1) + .ColWidth(conCol内容 + 2) - 15
        Else '.RowData(.Row) = 2 Then
            Me.txtInput.Left = .Left + .ColWidth(conCol标记) + .ColWidth(conCol内容) + .ColWidth(conCol内容 + 1) + 45
            Me.txtInput.Width = .ColWidth(conCol内容 + 2) - 15
        End If
        Me.txtInput.Text = .Text
        zlWordSelect = InStr(1, Me.txtInput.Text, strWord)
        If zlWordSelect <> 0 Then
            Me.txtInput.SelStart = zlWordSelect - 1
            Me.txtInput.SelLength = Len(strWord)
        End If
        Me.txtInput.Visible = True
        Me.mnuEditString.Visible = True
        Me.txtInput.SetFocus
    End With
    DoEvents

End Function

Public Sub zlWordReplace(lngCurRow As Long, strSource As String, strObject As String)
    '-------------------------------------------------
    '功能:替换指定行的文本内容
    '入参:  lngCurRow-指定行；strSource-指定被替换的文本；strObject-替换为文本
    '-------------------------------------------------
    Me.txtInput.Visible = False
    Me.hgdRefer.Row = lngCurRow
    Me.hgdRefer.Col = conCol内容 + 2
    With Me.hgdRefer
        If .RowData(.Row) = 0 Then Exit Sub
        Me.txtInput.Text = .Text
        Me.txtInput.Text = Replace(Me.txtInput.Text, strSource, strObject)
        If .RowData(.Row) = 1 Then
            .TextMatrix(.Row, conCol内容 + 1) = Me.txtInput.Text
            .TextMatrix(.Row, conCol内容 + 2) = Me.txtInput.Text
        Else '.RowData(.Row) = 2 Then
            .TextMatrix(.Row, conCol内容 + 2) = Me.txtInput.Text
        End If
    End With
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

