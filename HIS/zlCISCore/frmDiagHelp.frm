VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDiagHelp 
   BackColor       =   &H8000000C&
   Caption         =   "疾病参考查阅"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   Icon            =   "frmDiagHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picVBar 
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
      Height          =   6660
      Left            =   4155
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6660
      ScaleWidth      =   30
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   30
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   7890
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiagHelp.frx":038A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12806
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
   Begin VB.PictureBox picItem 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   4605
      Picture         =   "frmDiagHelp.frx":0C1C
      ScaleHeight     =   675
      ScaleWidth      =   7245
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   645
      Width           =   7245
      Begin VB.Label lblAlias 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "别名"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3495
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "原发性高血压"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   60
         Width           =   5940
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3210
      Top             =   7380
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
            Picture         =   "frmDiagHelp.frx":1238
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":17D2
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":1D6C
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   8070
      Top             =   1155
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
            Picture         =   "frmDiagHelp.frx":2306
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":2520
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":273A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":2954
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":2B6E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":2D88
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":2FA8
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   7380
      Top             =   1155
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
            Picture         =   "frmDiagHelp.frx":31C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":33E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":35FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":3816
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":3A30
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":3C50
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagHelp.frx":3E70
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   9330
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdRefer 
      Height          =   5295
      Left            =   4665
      TabIndex        =   15
      Top             =   1860
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   9340
      _Version        =   393216
      BackColor       =   -2147483628
      Rows            =   5
      Cols            =   4
      FixedRows       =   0
      BackColorFixed  =   -2147483628
      ForeColorFixed  =   32768
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483628
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   900
      BandCount       =   1
      _CBWidth        =   10125
      _CBHeight       =   510
      _Version        =   "6.0.8169"
      Child1          =   "tlbThis"
      MinWidth1       =   24000
      MinHeight1      =   450
      Width1          =   8730
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   450
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   24000
         _ExtentX        =   42333
         _ExtentY        =   794
         ButtonWidth     =   1561
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "隐藏"
               Key             =   "隐藏"
               Description     =   "隐藏"
               Object.ToolTipText     =   "隐藏目录区"
               Object.Tag             =   "隐藏"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "显示"
               Key             =   "显示"
               Description     =   "显示"
               Object.ToolTipText     =   "显示目录区域"
               Object.Tag             =   "显示"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "后退"
               Key             =   "后退"
               Description     =   "后退"
               Object.ToolTipText     =   "后退一个主题"
               Object.Tag             =   "后退"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "前进"
               Key             =   "前进"
               Description     =   "前进"
               Object.ToolTipText     =   "前进一个主题"
               Object.Tag             =   "前进"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Description     =   "打印"
               Object.ToolTipText     =   "打印当前表"
               Object.Tag             =   "打印"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab stbCatalog 
      Height          =   6750
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   11906
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "目录(&C)"
      TabPicture(0)   =   "frmDiagHelp.frx":4090
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "搜索(&S)"
      TabPicture(1)   =   "frmDiagHelp.frx":40AC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdShow(0)"
      Tab(1).Control(1)=   "cmdFind"
      Tab(1).Control(2)=   "cboFind"
      Tab(1).Control(3)=   "lvwFind"
      Tab(1).Control(4)=   "lblNote"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "书签(&I)"
      TabPicture(2)   =   "frmDiagHelp.frx":40C8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwMark"
      Tab(2).Control(1)=   "cmdMark"
      Tab(2).Control(2)=   "cmdDel"
      Tab(2).Control(3)=   "cmdShow(1)"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdShow 
         Caption         =   "显示(&D)"
         Height          =   350
         Index           =   1
         Left            =   -72660
         TabIndex        =   11
         Top             =   5745
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&E)"
         Height          =   350
         Left            =   -72765
         TabIndex        =   9
         Top             =   510
         Width           =   1110
      End
      Begin VB.CommandButton cmdMark 
         Caption         =   "当前项标记为书签(&M)"
         Height          =   350
         Left            =   -74820
         TabIndex        =   8
         Top             =   495
         Width           =   2010
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "显示(&D)"
         Enabled         =   0   'False
         Height          =   350
         Index           =   0
         Left            =   -72690
         TabIndex        =   7
         Top             =   6075
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "列出搜索项目(&L)"
         Height          =   350
         Left            =   -73290
         TabIndex        =   5
         Top             =   1350
         Width           =   1530
      End
      Begin VB.ComboBox cboFind 
         Height          =   300
         Left            =   -74910
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   855
         Width           =   3165
      End
      Begin VB.PictureBox picList 
         Height          =   5880
         Left            =   90
         ScaleHeight     =   5820
         ScaleWidth      =   3000
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "100"
         Top             =   390
         Width           =   3060
         Begin VB.CommandButton cmdKind 
            Caption         =   "中医疾病(&2)"
            Height          =   350
            Index           =   1
            Left            =   0
            TabIndex        =   20
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   345
            Width           =   2295
         End
         Begin VB.CommandButton cmdKind 
            Caption         =   "西医疾病(&1)"
            Height          =   350
            Index           =   0
            Left            =   0
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   15
            Width           =   2295
         End
         Begin MSComctlLib.TreeView tvwList 
            Height          =   4005
            Left            =   45
            TabIndex        =   2
            Tag             =   "1000"
            Top             =   1020
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   7064
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imgList"
            Appearance      =   0
         End
      End
      Begin MSComctlLib.ListView lvwFind 
         Height          =   4065
         Left            =   -74925
         TabIndex        =   6
         Top             =   1860
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   7170
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwMark 
         Height          =   4605
         Left            =   -74835
         TabIndex        =   10
         Top             =   930
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   8123
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "键入要查找的编码、名称、别名或简码(&W):"
         Height          =   180
         Left            =   -74895
         TabIndex        =   3
         Top             =   495
         Width           =   3495
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "比例尺寸"
      Height          =   180
      Left            =   7530
      TabIndex        =   22
      Top             =   7350
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDiagHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrSteps As String    '后退和前进项目记录步骤

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String, aryTemp() As String

Private WithEvents objParentForm As Form
Attribute objParentForm.VB_VarHelpID = -1

Public Sub ShowMe(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemId As Long)
    '---------------------------------------------
    '功能：根据上级程序要求，以模态或非模态显示诊疗参考
    '入参：frmParent-父窗体；
    '      blnModal-是否模态显示（通常和上级窗体一致）；
    '      lngItemId-要显示的诊断ID，不为0时，缺省不显示目录区；
    '---------------------------------------------
    If lngItemId = 0 Then
        Me.tlbThis.Buttons("隐藏").Visible = True: Me.tlbThis.Buttons("显示").Visible = False
        Call cmdKind_Click(0)
        Call stbCatalog_Click(0)
    Else
        Me.tlbThis.Buttons("隐藏").Visible = False: Me.tlbThis.Buttons("显示").Visible = True
        Call zlShowRef(lngItemId)
    End If
    
    On Error Resume Next
    Set objParentForm = frmParent
    Me.Show bytModal, frmParent
End Sub

Private Sub cboFind_Click()
    If Trim(Me.cboFind.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cboFind_GotFocus()
    Me.cboFind.SelStart = 0: Me.cboFind.SelLength = 100
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFind_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Me.cboFind.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cmdDel_Click()
    If Me.lvwMark.SelectedItem Is Nothing Then Exit Sub
    Me.lvwMark.ListItems.Remove (Me.lvwMark.SelectedItem.Key)
    If Me.lvwMark.ListItems.Count > 0 Then
        Me.cmdDel.Enabled = True: Me.cmdShow(1).Enabled = True
    Else
        Me.cmdDel.Enabled = False: Me.cmdShow(1).Enabled = False
    End If
    
    strTemp = ""
    For Each objItem In Me.lvwMark.ListItems
        strTemp = strTemp & "," & Mid(objItem.Key, 2)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\疾病诊断书签\", UserInfo.用户名, strTemp)
End Sub

Private Sub cmdFind_Click()
    If Trim(Me.cboFind.Text) = "" Then
        MsgBox "请输入查找的内容", vbExclamation, gstrSysName
        Me.cboFind.SetFocus: Exit Sub
    End If
    strTemp = ""
    For intCount = 0 To Me.cboFind.ListCount
        strTemp = strTemp & ";" & Me.cboFind.List(intCount)
    Next
    If InStr(1, strTemp, ";" & Trim(Me.cboFind.Text)) = 0 Then
        Me.cboFind.AddItem Trim(Me.cboFind.Text), 0
    End If
    
    Me.lvwFind.ListItems.Clear
    gstrSql = "select distinct I.ID,I.编码,I.名称,decode(I.类别,1,'西医','中医') as 类别" & _
            " from 疾病诊断目录 I,疾病诊断别名 N" & _
            " where I.ID=N.诊断ID" & _
            "       and (I.编码 like '" & Trim(Me.cboFind.Text) & "%'" & _
            "           or N.名称 like '" & gstrMatch & Trim(Me.cboFind.Text) & "%'" & _
            "           or N.简码 like '" & gstrMatch & Trim(Me.cboFind.Text) & "%')"
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            Set objItem = Me.lvwFind.ListItems.Add(, "_" & !ID, !名称, "item", "item")
            objItem.SubItems(Me.lvwFind.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwFind.ColumnHeaders("类别").Index - 1) = !类别
            .MoveNext
        Loop
    End With
    If Me.lvwFind.ListItems.Count > 0 Then
        Me.lvwFind.ListItems(1).Selected = True
        Me.lvwFind.SelectedItem.EnsureVisible
        Me.cmdShow(0).Enabled = True
    Else
        Me.cmdShow(0).Enabled = False
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdKind_Click(Index As Integer)
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = 0
        Else
            Me.cmdKind(intCount).Tag = 1
        End If
    Next
    If Me.stbCatalog.Visible And Me.tvwList.Visible Then Me.tvwList.SetFocus
    If Val(Me.tvwList.Tag) <> Index Then
        Call picList_Resize
        Me.tvwList.Tag = Index
        Call zlRefList
    End If
End Sub

Private Sub cmdMark_Click()
    If Val(Me.lblItem.Tag) = 0 Then Exit Sub
    gstrSql = "select I.ID,I.编码,I.名称,decode(I.类别,1,'西医','中医') as 类别" & _
            " from 疾病诊断目录 I" & _
            " where I.ID=" & Me.lblItem.Tag
    With rsTemp
        Err = 0: On Error GoTo ErrHand
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        If Not .EOF Then
            Err = 0: On Error Resume Next
            Set objItem = Me.lvwMark.ListItems.Add(, "_" & !ID, !名称, "item", "item")
            objItem.SubItems(Me.lvwMark.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwMark.ColumnHeaders("类别").Index - 1) = !类别
            Me.lvwMark.ListItems("_" & Me.lblItem.Tag).Selected = True
            Me.lvwMark.SelectedItem.EnsureVisible
        End If
    End With
    If Me.lvwMark.ListItems.Count > 0 Then
        Me.cmdDel.Enabled = True: Me.cmdShow(1).Enabled = True
    Else
        Me.cmdDel.Enabled = False: Me.cmdShow(1).Enabled = False
    End If
    
    strTemp = ""
    For Each objItem In Me.lvwMark.ListItems
        strTemp = strTemp & "," & Mid(objItem.Key, 2)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\疾病诊断书签\", UserInfo.用户名, strTemp)
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdShow_Click(Index As Integer)
    If Index = 0 Then
        If Me.lvwFind.SelectedItem Is Nothing Then Exit Sub
        Call zlShowRef(Mid(Me.lvwFind.SelectedItem.Key, 2))
    Else
        If Me.lvwMark.SelectedItem Is Nothing Then Exit Sub
        Call zlShowRef(Mid(Me.lvwMark.SelectedItem.Key, 2))
    End If
End Sub

Private Sub Form_Activate()
    If Me.tlbThis.Buttons("隐藏").Visible = True Then
        Me.stbCatalog.Visible = True: Me.picVBar.Visible = True
    Else
        Me.stbCatalog.Visible = False: Me.picVBar.Visible = False
    End If
End Sub

Private Sub Form_Load()
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    
    '截面元素形态设置
    Me.lvwFind.ListItems.Clear
    With Me.lvwFind.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2200
        .Add , "编码", "编码", 1000
        .Add , "类别", "类别", 1000
    End With
    With Me.lvwFind
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1: .SortOrder = lvwAscending
    End With
    Me.lvwMark.ListItems.Clear
    With Me.lvwMark.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2200
        .Add , "编码", "编码", 1000
        .Add , "类别", "类别", 1000
    End With
    With Me.lvwMark
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1: .SortOrder = lvwAscending
    End With
    
    With Me.hgdRefer
        .ColWidth(0) = 250
        .ColWidth(1) = Me.TextWidth("空格")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
    End With
        
    '提取书签显示
    strTemp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\疾病诊断书签\", UserInfo.用户名, "")
    If strTemp = "" Then Exit Sub
        
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "select I.ID,I.编码,I.名称,decode(I.类别,1,'西医','中医') as 类别" & _
                " from 疾病诊断目录 I" & _
                " where I.ID in (" & strTemp & ")"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            Set objItem = Me.lvwMark.ListItems.Add(, "_" & !ID, !名称, "item", "item")
            objItem.SubItems(Me.lvwMark.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwMark.ColumnHeaders("类别").Index - 1) = !类别
            .MoveNext
        Loop
    End With
    If Me.lvwMark.ListItems.Count > 0 Then
        Me.lvwMark.ListItems(1).Selected = True
        Me.lvwMark.SelectedItem.EnsureVisible
        Me.cmdDel.Enabled = True: Me.cmdShow(1).Enabled = True
    Else
        Me.cmdDel.Enabled = False: Me.cmdShow(1).Enabled = False
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    
    If Me.tlbThis.Buttons("隐藏").Visible = True Then
        With Me.picVBar
            .Top = lngTools
            .Height = Me.ScaleHeight - picList.Top - lngStatus
            If .Left < 2000 Then .Left = 2000
            If .Left > Me.ScaleWidth - 4000 Then .Left = Me.ScaleWidth - 4000
        End With
        
        With Me.stbCatalog
            .Left = Me.ScaleLeft
            .Top = lngTools
            .Height = Me.ScaleHeight - .Top - lngStatus + 15
            .Width = Me.picVBar.Left - .Left + 15
        End With
        With Me.picList
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = 390: .Height = Me.stbCatalog.Height - .Top - 90
        End With
        With Me.lblNote
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = 450
        End With
        With Me.cboFind
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = Me.lblNote.Top + Me.lblNote.Height + 45
        End With
        With Me.cmdFind
            .Left = Me.stbCatalog.Width - .Width - 90
            .Top = Me.cboFind.Top + Me.cboFind.Height + 90
        End With
        With Me.cmdShow(0)
            .Left = Me.stbCatalog.Width - .Width - 90
            .Top = Me.stbCatalog.Height - 180 - .Height
        End With
        With Me.lvwFind
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = Me.cmdFind.Top + Me.cmdFind.Height + 90
            .Height = Me.cmdShow(0).Top - .Top - 90
        End With
        
        With Me.cmdMark
            .Left = 90: .Top = 450
        End With
        With Me.cmdDel
            .Left = Me.cmdMark.Left + Me.cmdMark.Width + 45: .Top = 450
        End With
        With Me.cmdShow(1)
            .Left = Me.stbCatalog.Width - .Width - 90
            .Top = Me.stbCatalog.Height - 180 - .Height
        End With
        With Me.lvwMark
            .Left = 90: .Width = Me.stbCatalog.Width - .Left - 90
            .Top = Me.cmdMark.Top + Me.cmdMark.Height + 90
            .Height = Me.cmdShow(1).Top - .Top - 90
        End With
        
        With Me.picItem
            .Left = Me.picVBar.Left + Me.picVBar.Width: .Width = Me.ScaleWidth - .Left
            .Top = lngTools
        End With
    Else
        With Me.picItem
            .Left = Me.ScaleLeft: .Width = Me.ScaleWidth - .Left
            .Top = lngTools
        End With
    End If
    
    With Me.lblItem
        .Left = 45: .Width = Me.picItem.ScaleWidth
    End With
    With Me.lblAlias
        .Left = 45: .Width = Me.picItem.ScaleWidth - 90
    End With
    Me.picItem.Height = Me.lblAlias.Top + Me.lblAlias.Height + 90
    
    With Me.hgdRefer
        .Redraw = False
        .Left = Me.picItem.Left: .Width = Me.ScaleWidth - .Left
        .Top = Me.picItem.Top + Me.picItem.Height: .Height = Me.ScaleHeight - .Top - lngStatus
        .ColWidth(0) = 250
        .ColWidth(1) = Me.TextWidth("空格")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
         Call zlGrdRowHeight
        .Redraw = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub hgdRefer_DblClick()
    With Me.hgdRefer
        If .TextMatrix(.Row, 0) <> "Δ" And .TextMatrix(.Row, 0) <> "▼" Then Exit Sub
        .Redraw = False
        For intCount = .Row + 1 To .Rows - 1
            If .TextMatrix(intCount, 0) = "Δ" Or .TextMatrix(intCount, 0) = "▼" Then Exit For
            If .TextMatrix(.Row, 0) = "▼" Then
                .RowHeight(intCount) = 255
            Else
                .RowHeight(intCount) = 0
            End If
        Next
        If .TextMatrix(.Row, 0) = "▼" Then
            .TextMatrix(.Row, 0) = "Δ"
        Else
            .TextMatrix(.Row, 0) = "▼"
        End If
        Call zlGrdRowHeight
        .Redraw = True
    End With
End Sub

Private Sub hgdRefer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call hgdRefer_DblClick
End Sub

Private Sub lvwFind_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwFind.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwFind.SortOrder = IIf(Me.lvwFind.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwFind.SortKey = ColumnHeader.Index - 1
        Me.lvwFind.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwFind_DblClick()
    If Me.lvwFind.SelectedItem Is Nothing Then Exit Sub
    Call zlShowRef(Mid(Me.lvwFind.SelectedItem.Key, 2))
End Sub

Private Sub lvwMark_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwMark.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwMark.SortOrder = IIf(Me.lvwMark.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwMark.SortKey = ColumnHeader.Index - 1
        Me.lvwMark.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMark_DblClick()
    If Me.lvwMark.SelectedItem Is Nothing Then Exit Sub
    Call zlShowRef(Mid(Me.lvwMark.SelectedItem.Key, 2))
End Sub

Private Sub objParentForm_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub picList_Resize()
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picList.ScaleLeft + 0
        Me.cmdKind(intCount).Width = Me.picList.ScaleWidth + 15
        Me.cmdKind(intCount).Height = 300
        If Val(Me.cmdKind(intCount).Tag) = 0 Then
            Me.cmdKind(intCount).Top = Me.picList.ScaleTop + 285 * intCount
            Me.tvwList.Top = Me.picList.ScaleTop + 285 * (intCount + 1)
        Else
            Me.cmdKind(intCount).Top = Me.picList.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
        End If
    Next
    Me.tvwList.Left = Me.picList.ScaleLeft + 15
    Me.tvwList.Width = Me.picList.ScaleWidth
    Me.tvwList.Height = Me.picList.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Me.picVBar.Left = Me.picVBar.Left + X
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub stbCatalog_Click(PreviousTab As Integer)
    Select Case Me.stbCatalog.Tab
    Case 0
        Me.picList.Visible = True
        Me.lblNote.Visible = False
        Me.cboFind.Visible = False
        Me.cmdFind.Visible = False
        Me.lvwFind.Visible = False
        Me.cmdShow(0).Visible = False
        Me.cmdMark.Visible = False
        Me.cmdDel.Visible = False
        Me.lvwMark.Visible = False
        Me.cmdShow(1).Visible = False
    Case 1
        Me.picList.Visible = False
        Me.lblNote.Visible = True
        Me.cboFind.Visible = True
        Me.cmdFind.Visible = True
        Me.lvwFind.Visible = True
        Me.cmdShow(0).Visible = True
        Me.cmdMark.Visible = False
        Me.cmdDel.Visible = False
        Me.lvwMark.Visible = False
        Me.cmdShow(1).Visible = False
    Case 2
        Me.picList.Visible = False
        Me.lblNote.Visible = False
        Me.cboFind.Visible = False
        Me.cmdFind.Visible = False
        Me.lvwFind.Visible = False
        Me.cmdShow(0).Visible = False
        Me.cmdMark.Visible = True
        Me.cmdDel.Visible = True
        Me.lvwMark.Visible = True
        Me.cmdShow(1).Visible = True
    End Select
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "隐藏"
        Me.tlbThis.Buttons("隐藏").Visible = False: Me.tlbThis.Buttons("显示").Visible = True
        If Me.WindowState <> 2 Then
            Me.Left = Me.Left + (Me.stbCatalog.Width + Me.picVBar.Width)
            Me.Width = Me.Width - (Me.stbCatalog.Width + Me.picVBar.Width)
        End If
        Me.stbCatalog.Visible = False: Me.picVBar.Visible = False
        Call Form_Resize
    Case "显示"
        Me.tlbThis.Buttons("隐藏").Visible = True: Me.tlbThis.Buttons("显示").Visible = False
        If Me.WindowState <> 2 Then
            If Me.Left - (Me.stbCatalog.Width + Me.picVBar.Width) >= 0 Then
                Me.Left = Me.Left - (Me.stbCatalog.Width + Me.picVBar.Width)
            Else
                Me.Left = 0
            End If
            Me.Width = Me.Width + (Me.stbCatalog.Width + Me.picVBar.Width)
        End If
        Me.stbCatalog.Visible = True: Me.picVBar.Visible = True
        
        '定位到当前显示项目所在的目录
        If Val(Me.picItem.Tag) = 1 Then
            Call cmdKind_Click(0)
        Else
            Call cmdKind_Click(1)
        End If
        Call stbCatalog_Click(0)
        Call Form_Resize
    Case "后退"
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        Call zlShowRef(Val(aryTemp(intCount - 1)))
    Case "前进"
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        Call zlShowRef(Val(aryTemp(intCount + 1)))
    Case "打印"
        Dim objPrint As New zlPrint1Grd
        Dim objRow As New zlTabAppRow
        Dim bytMode As Byte
        If Me.hgdRefer.TextMatrix(Me.hgdRefer.FixedRows, 2) = "" Then Exit Sub
        Set objPrint.Body = Me.hgdRefer
        With objPrint.Title
            .Text = Me.lblItem.Caption & ".诊断参考"
            .Font.Size = 11
        End With
        objRow.Add Me.lblAlias.Caption
        objPrint.UnderAppRows.Add objRow
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Case "帮助"
        ShowHelp App.ProductName, Me.Hwnd, Me.Name ', Int((glngSys) / 100)
    Case "退出"
        Unload Me
    End Select
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If InStr(1, Node.Key, "_") = 0 Then Exit Sub
    Call zlShowRef(Mid(Node.Key, InStr(1, Node.Key, "_") + 1))
End Sub

Private Sub zlRefList()
    '---------------------------------------------
    '填写诊疗项目分类与目录
    '---------------------------------------------
    Me.tvwList.Visible = False
    Me.tvwList.Nodes.Clear
    '填写分类
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "select ID,上级ID,编码,名称,简码" & _
                " From 疾病诊断分类" & _
                " Where 类别=" & IIf(Val(Me.tvwList.Tag) = 0, 1, 2) & _
                " start with 上级ID is null" & _
                " connect by prior ID=上级ID"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwList.Nodes.Add(, , "C" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwList.Nodes.Add("C" & !上级ID, tvwChild, "C" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        
        gstrSql = "select R.分类ID,I.ID,I.编码,I.名称" & _
                " from 疾病诊断目录 I,疾病诊断属类 R" & _
                " where I.ID=R.诊断ID and I.类别=" & IIf(Val(Me.tvwList.Tag) = 0, 1, 2) & _
                " order by I.编码"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            Set objNode = Me.tvwList.Nodes.Add("C" & !分类ID, tvwChild, "C" & !分类ID & "_" & !ID, "[" & !编码 & "]" & !名称, "item")
            .MoveNext
        Loop
        
        If Me.tvwList.SelectedItem Is Nothing And .RecordCount <> 0 Then
            .MoveFirst
            Me.tvwList.Nodes("C" & !分类ID & "_" & !ID).Selected = True
        End If
        If Not (Me.tvwList.SelectedItem Is Nothing) Then
            Me.tvwList.SelectedItem.EnsureVisible
            Call tvwList_NodeClick(Me.tvwList.SelectedItem)
        End If
    End With
    
    Me.tvwList.Visible = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlShowRef(lngItemId As Long)
    '------------------------------------------------
    '功能：显示指定项目的诊疗参考内容
    '------------------------------------------------
    '清理详细信息显示区
    With Me.hgdRefer
        .Rows = 1: .ColAlignment(0) = 1: .ColAlignment(1) = 1: .ColAlignment(2) = 1
        For intCount = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCount) = ""
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand
    '------------------------------------------------
    Me.hgdRefer.Redraw = False
    With rsTemp
        Me.lblAlias.Caption = ""
        gstrSql = "select distinct I.ID,I.编码,N.性质,N.名称||decode(N.性质,1,'',2,'(英文名)','(别名)') as 名称,I.类别" & _
                " from 疾病诊断目录 I,疾病诊断别名 N" & _
                " where I.ID=N.诊断ID and I.ID=" & lngItemId
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            If !性质 = 1 Then
                Me.picItem.Tag = !类别: Me.lblItem.Tag = !ID: Me.lblItem.Caption = IIf(IsNull(!名称), "", !名称)
            Else
                Me.lblAlias.Caption = Me.lblAlias.Caption & "、" & IIf(IsNull(!名称), "", !名称)
            End If
            .MoveNext
        Loop
        If Me.lblAlias.Caption <> "" Then Me.lblAlias.Caption = Mid(Me.lblAlias.Caption, 2)
        
        gstrSql = "select 项目层次,项目序号,nvl(证候序号,0) as 证候序号,0 as 内容行号,decode(nvl(证候名称,''),'',参考项目,证候名称) as 内容" & _
                " from 疾病诊断参考 " & _
                " where 诊断id=" & lngItemId & _
                " union" & _
                " select 项目层次,项目序号,nvl(证候序号,0) as 证候序号,内容行号,decode(nvl(证候名称,''),'',内容文本,参考项目||'：'||内容文本) as 内容" & _
                " from 疾病诊断参考" & _
                " where 诊断id=" & lngItemId & " and length(ltrim(nvl(内容文本,'')))<>0" & _
                " order by 项目序号,证候序号,内容行号"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        If .EOF Or .BOF Then
            Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + 1
        Else
            Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            Me.hgdRefer.RowHeight(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = 255
            Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 0) = ""
            If !内容行号 = 0 Then
                If !项目层次 = 1 Then
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 0) = "Δ"
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = "【" & !内容 & "】"
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = "【" & !内容 & "】"
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = True
                Else
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = ""
                    If !证候序号 = 0 Then
                        Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = "［" & !内容 & "］"
                    Else
                        Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = !证候序号 & "." & !内容
                    End If
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = False
                End If
            Else
                If !项目层次 = 1 Then
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = Space(4) & !内容
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = Space(4) & !内容
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = True
                Else
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = ""
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = Space(4) & !内容
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = False
                End If
            End If
            .MoveNext
        Loop
    End With
    Call Form_Resize
    Me.hgdRefer.Redraw = True
    
    If mstrSteps = "" Then
        mstrSteps = Val(Me.lblItem.Tag)
    Else
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        If intCount > UBound(aryTemp) Then mstrSteps = mstrSteps & ";" & Val(Me.lblItem.Tag)
    End If
    
    aryTemp = Split(mstrSteps, ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
    Next
    If intCount > LBound(aryTemp) Then
        Me.tlbThis.Buttons("后退").Enabled = True
    Else
        Me.tlbThis.Buttons("后退").Enabled = False
    End If
    If intCount < UBound(aryTemp) Then
        Me.tlbThis.Buttons("前进").Enabled = True
    Else
        Me.tlbThis.Buttons("前进").Enabled = False
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '根据调整内容调整内容网格的行高度，以保证内容的正常显示
    '---------------------------------------------
    Dim lngColWidth As Long
    With Me.hgdRefer
        For intRow = .FixedRows To .Rows - 1
            If .RowHeight(intRow) <> 0 Then
                If .TextMatrix(intRow, 1) = "" Then
                    lngColWidth = .ColWidth(2)
                Else
                    lngColWidth = .ColWidth(1) + .ColWidth(2)
                End If
                Me.lblScale.Width = lngColWidth - 90
                Me.lblScale.Caption = .TextMatrix(intRow, 2)
                .RowHeight(intRow) = Me.lblScale.Height + 75
            End If
        Next
    End With
End Sub
