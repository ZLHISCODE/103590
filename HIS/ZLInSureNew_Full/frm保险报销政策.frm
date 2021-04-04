VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm保险报销政策 
   Caption         =   "年度结算规则"
   ClientHeight    =   7215
   ClientLeft      =   825
   ClientTop       =   2505
   ClientWidth     =   7320
   Icon            =   "frm保险报销政策.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ils32 
      Left            =   1470
      Top             =   2730
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":08CA
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":0BE4
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":0EFE
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":1218
            Key             =   "CommonD"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   2130
      ScaleHeight     =   4905
      ScaleWidth      =   4125
      TabIndex        =   8
      Top             =   1860
      Width           =   4125
      Begin VB.PictureBox picSplitH 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   30
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   4275
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1755
         Width           =   4275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh支付限额 
         Height          =   2430
         Left            =   90
         TabIndex        =   9
         Top             =   375
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   4286
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483643
         GridColor       =   4210752
         GridColorFixed  =   4210752
         FocusRect       =   2
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh支付比例 
         Height          =   1410
         Left            =   120
         TabIndex        =   10
         Top             =   3300
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   2487
         _Version        =   393216
         Rows            =   3
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483643
         GridColor       =   4210752
         GridColorFixed  =   4210752
         FocusRect       =   2
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lbl支付限额 
         AutoSize        =   -1  'True
         Caption         =   "起付线与封顶线"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   90
         Width           =   1260
      End
      Begin VB.Label lbl支付比例 
         AutoSize        =   -1  'True
         Caption         =   "统筹支付比例"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   90
         TabIndex        =   12
         Top             =   3000
         Width           =   1080
      End
   End
   Begin VB.ComboBox cmb中心 
      Height          =   300
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   900
      Width           =   2385
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   7320
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "打印预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "New"
               Object.ToolTipText     =   "增加新年度规则"
               Object.Tag             =   "增加"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除末年度规则"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   5205
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":1532
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":174C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":1966
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":1B80
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":1D9A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":1FB4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":21CE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   4485
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":23E8
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":2602
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":281C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":2A36
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":2C50
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":2E6A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险报销政策.frx":3084
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwKind_S 
      Height          =   5325
      Left            =   105
      TabIndex        =   0
      Top             =   870
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   9393
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6855
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   635
      SimpleText      =   $"frm保险报销政策.frx":329E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm保险报销政策.frx":32E5
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6540
      Left            =   1890
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6540
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   780
      Width           =   45
   End
   Begin MSComctlLib.TabStrip tab年度 
      Height          =   5400
      Left            =   2055
      TabIndex        =   5
      Top             =   1395
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   9525
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "2002年"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label lbl中心 
      AutoSize        =   -1  'True
      Caption         =   "医保中心(&N)"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   2010
      TabIndex        =   7
      Top             =   930
      Width           =   990
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加新年度规则(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除末年度规则(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditLimit 
         Caption         =   "封顶与起付线(&L)"
      End
      Begin VB.Menu mnuEditProportion 
         Caption         =   "统筹支付比例(&P)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolspilt1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
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
Attribute VB_Name = "frm保险报销政策"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim msngStartX As Single, msngStartY As Single    '移动前鼠标的位置
Dim mstrKey As String
Dim mlng中心 As Long
Dim mlng当前年 As Long
Dim mblnLoad As Boolean

Private Sub Form_Activate()
    If mblnLoad = True Then
        '显示当前项
        lvwKind_S.SelectedItem.EnsureVisible
        lvwKind_S_ItemClick lvwKind_S.SelectedItem
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    
    mblnLoad = True
    Call 权限控制
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    lvwKind_S.Top = sngTop
    lvwKind_S.Height = IIf(sngBottom - lvwKind_S.Top > 0, sngBottom - lvwKind_S.Top, 0)
    lvwKind_S.Left = ScaleLeft
    
    picSplitV.Top = sngTop
    picSplitV.Height = IIf(sngBottom - picSplitV.Top > 0, sngBottom - picSplitV.Top, 0)
    picSplitV.Left = lvwKind_S.Left + lvwKind_S.Width
    
    With cmb中心
        '设置控件的左边距与宽度
        lbl中心.Left = picSplitV.Left + picSplitV.Width
        .Left = lbl中心.Left + lbl中心.Width + 30
        .Width = IIf(ScaleWidth - cmb中心.Left > 0, ScaleWidth - cmb中心.Left, 0)
    
        tab年度.Left = lbl中心.Left
        tab年度.Width = IIf(ScaleWidth - tab年度.Left > 0, ScaleWidth - tab年度.Left, 0)
    End With
    
    With tab年度
        If cmb中心.Visible = True Then
            cmb中心.Top = sngTop
            lbl中心.Top = sngTop + 60
            .Top = cmb中心.Top + cmb中心.Height + 90
        Else
            .Top = sngTop
        End If
        
        .Height = IIf(sngBottom - .Top > 0, sngBottom - .Top, 0)
        picContainer.Left = .ClientLeft
        picContainer.Width = .ClientWidth
        picContainer.Top = .ClientTop
        picContainer.Height = .ClientHeight
    End With
    Me.Refresh
End Sub

Private Sub picContainer_Resize()
    
    With lbl支付限额
        .Top = 60
        .Left = 60
    
        msh支付限额.Top = .Top + .Height + 30
        msh支付限额.Left = .Left
       
    End With
    With msh支付限额
        .Width = IIf(picContainer.ScaleWidth - 120 > 0, picContainer.ScaleWidth - 120, 0)
    
        picSplitH.Top = msh支付限额.Top + msh支付限额.Height + 90
        picSplitH.Left = .Left
        picSplitH.Width = .Width
        
        lbl支付比例.Top = picSplitH.Top + picSplitH.Height
        lbl支付比例.Left = .Left
        
        msh支付比例.Top = lbl支付比例.Top + lbl支付比例.Height + 30
        msh支付比例.Height = IIf(picContainer.ScaleHeight - msh支付比例.Top > 0, picContainer.ScaleHeight - msh支付比例.Top, 0)
        msh支付比例.Left = .Left
        msh支付比例.Width = .Width
    End With
    
End Sub

Private Sub msh支付限额_DblClick()
    If mnuEditLimit.Visible = True And mnuEditLimit.Enabled = True Then
        Call mnuEditLimit_Click
    End If
End Sub

Private Sub msh支付比例_DblClick()
    If mnuEditProportion.Visible = True And mnuEditProportion.Enabled = True Then
        Call mnuEditProportion_Click
    End If
End Sub

Private Sub lvwKind_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrKey = Item.Key Then Exit Sub
    
    Dim rsTemp As New ADODB.Recordset
    
    cmb中心.Clear
    cmb中心.Visible = (Item.Tag = "1")
    lbl中心.Visible = cmb中心.Visible
    Call Form_Resize
    
    On Error GoTo errHandle
    If cmb中心.Visible = False Then
        '该医保只能有一个中心
        cmb中心.AddItem "1." & Item.Text
    Else
        gstrSQL = "select 序号,编码,名称 from 保险中心目录 where 险类=" & Mid(Item.Key, 2) & " order by 序号"
        Call OpenRecordset(rsTemp, Me.Caption)
        
        Do Until rsTemp.EOF
            cmb中心.AddItem rsTemp("编码") & "." & rsTemp("名称")
            cmb中心.ItemData(cmb中心.NewIndex) = rsTemp("序号")
            rsTemp.MoveNext
        Loop
    End If
    
    If cmb中心.ListCount > 0 Then
        '避免产生Click事件
        zlControl.CboSetIndex cmb中心.hwnd, 0
    End If
    Call Fill年度
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmb中心_Click()
    If cmb中心.ItemData(cmb中心.ListIndex) = mlng中心 Then Exit Sub
    
    Call Fill年度
End Sub

Private Sub Fill年度()
'功能：根据当前险类与中心得到年度信息
    Dim lng险类 As Long
    Dim rsTemp As New ADODB.Recordset
    
    If lvwKind_S.SelectedItem Is Nothing Then
        mstrKey = ""
        lng险类 = 0
    Else
        mstrKey = lvwKind_S.SelectedItem.Key
        lng险类 = Mid(mstrKey, 2)
    End If
    If cmb中心.ListIndex < 0 Then
        mlng中心 = -1
    Else
        mlng中心 = cmb中心.ItemData(cmb中心.ListIndex)
    End If
    
    
    '读出该医保中心使用过医保年
    gstrSQL = "select distinct 年度 from 保险报销政策 where 险类=" & lng险类 & " and 中心=" & mlng中心
    Call OpenRecordset(rsTemp, Me.Caption)
    
    tab年度.Tabs.Clear
    If rsTemp.RecordCount = 0 Then
        tab年度.Tabs.Add , "K0", "无年度信息"
    Else
        Do Until rsTemp.EOF
            tab年度.Tabs.Add , "K" & rsTemp("年度"), rsTemp("年度") & "年"
            If rsTemp("年度") = mlng当前年 Then
                tab年度.Tabs("K" & mlng当前年).Selected = True
            End If
            rsTemp.MoveNext
        Loop
    End If
    If tab年度.SelectedItem Is Nothing Then
        tab年度.Tabs(1).Selected = True
    End If
    
    Call tab年度_Click
    
End Sub

Private Sub tab年度_Click()
'功能：显示当前年度的结算规则。如果没有，则显示空表
    Dim lng险类 As Long
    Dim lng年度 As Long
    
    lng年度 = Mid(tab年度.SelectedItem.Key, 2)
    If lng年度 = 0 Then
        Call InitTable
    Else
        lng险类 = Mid(lvwKind_S.SelectedItem.Key, 2)
        
        Call Fill支付限额(lng险类, lng年度)
        Call Fill支付比例(lng险类, lng年度)
    End If
    
    Call SetMenu
End Sub

Private Sub Fill支付限额(ByVal lng险类 As Long, ByVal lng年度 As Long)
'功能：显示当前年度的支付限额
'参数：传递进来险类、年度，中心由全局变量得到
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lngCol As Long
    Dim int人群 As Integer
    
    '得到人员性质
    gstrSQL = "select 序号,名称 from 保险人群 where 险类=" & lng险类 & " order by 序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    '初始化表格
    With msh支付限额
        If rsTemp.RecordCount <> 0 Then
            .Rows = 3
            .FixedRows = 2
            .Cols = rsTemp.RecordCount * 2 + 1
            .MergeCells = flexMergeRestrictRows
            
            lngRow = 0
            lngCol = 1
            
            .TextMatrix(0, 0) = "住院次数"
            .TextMatrix(1, 0) = "住院次数"
            
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, lngCol) = rsTemp!名称
                .ColData(lngCol) = rsTemp!序号
                lngCol = lngCol + 1
                .TextMatrix(lngRow, lngCol) = rsTemp!名称
                .ColData(lngCol) = rsTemp!序号
                lngCol = lngCol + 1
                rsTemp.MoveNext
            Loop
            
            lngRow = lngRow + 1
            lngCol = 1
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, lngCol) = "本院"
                lngCol = lngCol + 1
                .TextMatrix(lngRow, lngCol) = "转院"
                lngCol = lngCol + 1
                rsTemp.MoveNext
            Loop
        End If
    End With
    
    gstrSQL = " Select 人群,本院,类别,金额 From 保险报销政策 " & _
              " Where 险类=" & lng险类 & " And 中心=" & mlng中心 & " And 年度=" & lng年度 & _
              " And 性质=2 And 本院=1 And 人群=1 Order by 类别"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    With msh支付限额
        '首先清除数据
        For lngRow = 2 To .Rows - 1
            For lngCol = 1 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = ""
            Next
        Next
        If rsTemp.RecordCount = 0 Then
            .Rows = 3
            .TextMatrix(.Rows - 1, 0) = 1
        End If
        
        Do Until rsTemp.EOF
            If Val(rsTemp!类别) <> 0 Then
                lngRow = 1 + rsTemp!类别
            Else
                lngRow = .Rows - 1
            End If
            lngCol = 0
            If rsTemp("类别") = "A" Then
                .TextMatrix(lngRow, 0) = "封顶线"
                .RowData(lngRow) = -1
            Else
                .TextMatrix(lngRow, 0) = "第" & rsTemp("类别") & "次住院"
                .RowData(lngRow) = rsTemp!类别
            End If
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If .TextMatrix(.Rows - 1, 0) = "" Then .Rows = .Rows - 1
        
        gstrSQL = " Select 人群,本院,类别,金额 From 保险报销政策 " & _
                  " Where 险类=" & lng险类 & " And 中心=" & mlng中心 & " And 年度=" & lng年度 & _
                  " And 性质=2 " & _
                  " Order by 人群,本院,类别"
        Call OpenRecordset(rsTemp, Me.Caption)
        Do Until rsTemp.EOF
            If Val(rsTemp!类别) <> 0 Then
                lngRow = 1 + rsTemp!类别
            Else
                lngRow = .Rows - 1
            End If
            lngCol = rsTemp!人群 * 2 - IIf(rsTemp!本院 = 1, 1, 0)
            .TextMatrix(lngRow, lngCol) = Format(rsTemp("金额"), "########0;-########0; ; ")
            rsTemp.MoveNext
        Loop
        If .TextMatrix(.Rows - 1, 0) = "" And .Rows - 1 > 3 Then .Rows = .Rows - 1
    End With
    
    '设置对齐
    With msh支付限额
        For lngCol = 0 To .Cols - 1
            .ColAlignmentFixed(lngCol) = 4
            .ColWidth(lngCol) = 1200
        Next
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
    End With
End Sub

Private Sub Fill支付比例(ByVal lng险类 As Long, ByVal lng年度 As Long)
'功能：显示当前年度的支付比例
'参数：传递进来险类、年度，中心由全局变量得到
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lngCol As Long, lngRows As Long
    Dim col起始行 As New Collection      '每种人员性质的起始行
    
    '得到人员性质
    gstrSQL = "select 序号,名称 from 保险人群 where 险类=" & lng险类 & " order by 序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    '初始化表格
    With msh支付比例
        If rsTemp.RecordCount <> 0 Then
            .Rows = 3
            .FixedRows = 2
            .Cols = rsTemp.RecordCount * 2 + 1
            
            lngRow = 0
            lngCol = 1
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, lngCol) = rsTemp!名称
                .ColData(lngCol) = rsTemp!序号
                lngCol = lngCol + 1
                .TextMatrix(lngRow, lngCol) = rsTemp!名称
                .ColData(lngCol) = rsTemp!序号
                lngCol = lngCol + 1
                rsTemp.MoveNext
            Loop
            
            lngRow = lngRow + 1
            lngCol = 1
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, lngCol) = "本院"
                lngCol = lngCol + 1
                .TextMatrix(lngRow, lngCol) = "转院"
                lngCol = lngCol + 1
                rsTemp.MoveNext
            Loop
        End If
    End With
    
    '再得到费用档
    gstrSQL = "select 档次,名称 from 保险费用档 where 险类=" & lng险类 & " and 中心=" & mlng中心 & " order by 档次"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    With msh支付比例
        .ColWidth(0) = 1200
        .ColAlignment(0) = 7
        .TextMatrix(0, 0) = "费用档"
        .TextMatrix(1, 0) = "费用档"
        
        If rsTemp.RecordCount = 0 Then
            .RowData(2) = 0
        Else
            .Rows = .Rows + rsTemp.RecordCount - 1
            lngRow = 2
            Do Until rsTemp.EOF
                .TextMatrix(lngRow, 0) = rsTemp("名称")
                .RowData(lngRow) = rsTemp("档次")
                
                lngRow = lngRow + 1
                rsTemp.MoveNext
            Loop
        End If
        
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = 2: .Col = 1
    End With
    
    '最后得到支付比例
    gstrSQL = " Select 人群,本院,档次,比例 From 保险报销政策 " & _
              " Where 年度=" & lng年度 & " And 险类=" & lng险类 & " And 中心=" & mlng中心 & " And 性质=1 " & _
              " Order by 人群,本院,档次"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    With msh支付比例
        '首先清除数据
        For lngRow = 2 To .Rows - 1
            For lngCol = 1 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = ""
            Next
        Next
        
        Do Until rsTemp.EOF
            lngRow = 1 + rsTemp!档次 + 1
            lngCol = rsTemp!人群 * 2 - IIf(rsTemp!本院 = 1, 1, 0)
            
            .TextMatrix(lngRow, lngCol) = Format(rsTemp("比例"), "0.00")
            rsTemp.MoveNext
        Loop
    End With
    
    '设置对齐
    With msh支付比例
        For lngCol = 0 To .Cols - 1
            .ColAlignmentFixed(lngCol) = 4
            .ColWidth(lngCol) = 1200
        Next
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
    End With
End Sub

Private Sub mnuEditAdd_Click()
    Dim lng险类 As Long
    Dim lng年度 As Long, lng末年 As Long
    Dim blnReturn As VbMsgBoxResult
    
    lng险类 = Mid(lvwKind_S.SelectedItem.Key, 2)
    lng年度 = Val(Mid(tab年度.SelectedItem.Key, 2))
    If lng年度 = 0 Then
        '完全新增
        blnReturn = MsgBox("你是否要增加" & mlng当前年 & "年度的结算规则？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
        gstrSQL = "ZL_保险报销政策_NEW(" & lng险类 & "," & mlng中心 & "," & mlng当前年 & ")"
    Else
        lng末年 = Val(Mid(tab年度.Tabs(tab年度.Tabs.Count).Key, 2))
        blnReturn = MsgBox("你是否要将" & lng末年 & "年度的结算规则复制产生成为" & (lng末年 + 1) & "的？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
        gstrSQL = "ZL_保险报销政策_Copy(" & lng险类 & "," & mlng中心 & ")"
        
    End If
    
    If blnReturn = vbNo Then Exit Sub
    
    On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '刷新
    Call Fill年度
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDelete_Click()
    Dim lng险类 As Long
    Dim lng年度 As Long
    
    
    lng险类 = Mid(lvwKind_S.SelectedItem.Key, 2)
    lng年度 = Val(Mid(tab年度.SelectedItem.Key, 2))
    
    If MsgBox("真的需要删除“" & lng年度 & "年度结算规则”吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    gstrSQL = "ZL_保险报销政策_DELETE(1," & lng险类 & "," & mlng中心 & "," & lng年度 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "ZL_保险报销政策_DELETE(2," & lng险类 & "," & mlng中心 & "," & lng年度 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    
    Call Fill年度
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub mnuEditLimit_Click()
    Dim lng险类 As Long, lng年度 As Long
    
    lng险类 = Mid(lvwKind_S.SelectedItem.Key, 2)
    lng年度 = Mid(tab年度.SelectedItem.Key, 2)
    
    If frm报销支付限额.编辑支付限额(lng险类, mlng中心, lng年度) = True Then
        Call Fill支付限额(lng险类, lng年度)
    End If
End Sub

Private Sub mnuEditProportion_Click()
    Dim lng险类 As Long, lng年度 As Long
    
    lng险类 = Mid(lvwKind_S.SelectedItem.Key, 2)
    lng年度 = Mid(tab年度.SelectedItem.Key, 2)
    
    If frm报销支付比例.编辑支付比例(lng险类, mlng中心, lng年度) = True Then
        Call Fill支付比例(lng险类, lng年度)
    End If
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    
    Dim objPrint As New zlPrintGrds
    Dim objRow As New zlTabAppRow
    
    Set objPrint.Grds = New Collection
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "保险类别设置"
        
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add lbl支付限额
    objPrint.UnderAppRows.Add objRow
    
'    Set objRow = New zlTabAppRow
'    objRow.Add lblTax
'    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.姓名    '& "   打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    objPrint.Grds.Add msh支付限额
    objPrint.Grds.Add msh支付比例
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewGrds objPrint, 1
          Case 2
              zlPrintOrViewGrds objPrint, 2
          Case 3
              zlPrintOrViewGrds objPrint, 3
      End Select
    Else
        zlPrintOrViewGrds objPrint, bytMode
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewRefresh_Click()
'    Call RefList
End Sub

Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    stbThis.Visible = Me.mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    Me.mnuViewToolButton.Checked = Not Me.mnuViewToolButton.Checked
    Me.mnuViewToolText.Enabled = Me.mnuViewToolButton.Checked
    Me.cbrThis.Visible = Me.mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer, intRow As Integer, intCol As Integer
    
    Me.mnuViewToolText.Checked = Not Me.mnuViewToolText.Checked
    If Me.mnuViewToolText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Call Form_Resize
End Sub

Private Sub picSplitV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    
    If Button = 1 Then
        sngTemp = picSplitV.Left + x - msngStartX
        If sngTemp > 1000 And Me.ScaleWidth - (sngTemp + picSplitV.Width) > 1500 Then
            picSplitV.Left = sngTemp
            lvwKind_S.Width = picSplitV.Left - lvwKind_S.Left
            
            Call Form_Resize
        End If
        lvwKind_S.SetFocus
    End If
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartY = y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    
    If Button = 1 Then
        sngTemp = picSplitH.Top + y - msngStartY
        If sngTemp - msh支付限额.Top > 1500 And (msh支付比例.Top + msh支付比例.Height) - (sngTemp + picSplitV.Width) > 1500 Then
            picSplitH.Top = sngTemp
            msh支付限额.Height = picSplitH.Top - 90 - msh支付限额.Top
            
            Call picContainer_Resize
        End If
        msh支付限额.SetFocus
    End If
End Sub


Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnuFilePreview_Click
    Case "Print"
        mnuFilePrint_Click
    Case "New"
        mnuEditAdd_Click
    Case "Delete"
        mnuEditDelete_Click
    Case "Help"
        mnuHelpTitle_Click
    Case "Exit"
        mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub 权限控制()
    If InStr(gstrPrivs, "增删改") = 0 Then
        tbrThis.Buttons("New").Visible = False
        tbrThis.Buttons("Delete").Visible = False
        tbrThis.Buttons("Split1").Visible = False
        
        mnuEditAdd.Visible = False
        mnuEditDelete.Visible = False
        mnuEditSplit0.Visible = False
    End If
    
    If InStr(gstrPrivs, "封顶与起付线") = 0 Then
       mnuEditLimit.Visible = False
    End If
    
    If InStr(gstrPrivs, "统筹支付比例") = 0 Then
        If mnuEditAdd.Visible = True Or mnuEditLimit.Visible = True Then
            mnuEditProportion.Visible = False
        Else
            mnuEditSplit0.Visible = True
            mnuEditProportion.Visible = False
            mnuEdit.Visible = False
        End If
    End If
End Sub

Private Sub SetMenu()
'功能：根据当前的显示内容设置菜单可用性
    Dim blnEnable As Boolean
    
    stbThis.Panels(2).Text = "当前选择的医保类别是：" & lvwKind_S.SelectedItem.Text & " 年度为：" & tab年度.SelectedItem.Caption
    
    blnEnable = True 'Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_自贡市 And Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_泸州市
    
    mnuEditDelete.Enabled = (Mid(tab年度.SelectedItem.Key, 2) > 0 And Mid(tab年度.SelectedItem.Key, 2) <> mlng当前年) And blnEnable
    If tab年度.SelectedItem.Index < tab年度.Tabs.Count Then
        '只能从最后一年开始删除
        mnuEditDelete.Enabled = False
    End If
    
    tbrThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled
    
    mnuEditAdd.Enabled = cmb中心.ListIndex >= 0 And blnEnable
    tbrThis.Buttons("New").Enabled = mnuEditAdd.Enabled
    '只要中心一建立，缺省就有支付限额了
    mnuEditLimit.Enabled = (Mid(tab年度.SelectedItem.Key, 2) > 0) And blnEnable
    '有年龄档与费用档
    mnuEditProportion.Enabled = (msh支付比例.Rows - 1 >= 2) And blnEnable
End Sub

Private Sub InitTable()
    With msh支付限额
        .Cols = 2: .Rows = 2
        .Clear
        .ColAlignmentFixed(0) = 1
        .ColAlignment(1) = 7
        .ColWidth(0) = 1200
        .ColWidth(1) = 1200
        .TextMatrix(0, 0) = "住院次数"
        .TextMatrix(0, 1) = "人群"
        
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = 1: .Col = 1
    End With
    
    With msh支付比例
        .Cols = 2: .Rows = 2
        .Clear
        .ColWidth(0) = 1200
        .ColWidth(1) = 1200
        .ColAlignmentFixed(0) = 1
        .ColAlignment(1) = 7
        
        .TextMatrix(0, 0) = "费用档"
        .TextMatrix(0, 1) = "人群"
        .RowData(1) = 0: .ColData(1) = 0
        
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = 1: .Col = 1
    End With
End Sub

Public Sub ShowForm(frmParent As Form)
'功能：装入医保类别
'说明：使用本功能的主要原因是在出错退出时窗体不会闪
    Dim rsTemp As New ADODB.Recordset
    Dim strIcon As String, lst As ListItem
    
    gstrSQL = "select 序号,名称,是否固定,具有中心 from 保险类别 where nvl(是否禁止,0)<>1  order by 序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        '如果是在窗体初始化时调用，就不用处理其它内容了
        MsgBox "没有可用保险类别，不能使用本功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frm保险报销政策.Visible = True Then
        frm保险报销政策.Show
        Exit Sub
    End If
    
    '现在才能开始使用控件
    Call InitTable
    mlng当前年 = Format(zlDatabase.Currentdate, "yyyy")
    
    mstrKey = ""
    lvwKind_S.ListItems.Clear
    Do Until rsTemp.EOF
        strIcon = IIf(rsTemp("是否固定") = 1, "Fix", "Common")
        If rsTemp("序号") = gintInsure Then strIcon = strIcon & "D"
        
        Set lst = lvwKind_S.ListItems.Add(, "K" & rsTemp("序号"), rsTemp("名称"), strIcon, strIcon)
        If rsTemp("序号") = gintInsure Then
            lst.Selected = True
        End If
        
        lst.Tag = IIf(rsTemp("具有中心") = 1, 1, 0)
        
        rsTemp.MoveNext
    Loop
    If lvwKind_S.SelectedItem Is Nothing Then
        lvwKind_S.ListItems(1).Selected = True
    End If
    frm保险报销政策.Show , frmParent
End Sub


