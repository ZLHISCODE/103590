VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDiagHelp 
   BackColor       =   &H8000000C&
   Caption         =   "�����ο�����"
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
         Name            =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3495
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ���Ը�Ѫѹ"
         BeginProperty Font 
            Name            =   "����"
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
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����Ŀ¼��"
               Object.Tag             =   "����"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ʾ"
               Key             =   "��ʾ"
               Description     =   "��ʾ"
               Object.ToolTipText     =   "��ʾĿ¼����"
               Object.Tag             =   "��ʾ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����һ������"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "ǰ��"
               Key             =   "ǰ��"
               Description     =   "ǰ��"
               Object.ToolTipText     =   "ǰ��һ������"
               Object.Tag             =   "ǰ��"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ��ǰ��"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
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
      TabCaption(0)   =   "Ŀ¼(&C)"
      TabPicture(0)   =   "frmDiagHelp.frx":4090
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "����(&S)"
      TabPicture(1)   =   "frmDiagHelp.frx":40AC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdShow(0)"
      Tab(1).Control(1)=   "cmdFind"
      Tab(1).Control(2)=   "cboFind"
      Tab(1).Control(3)=   "lvwFind"
      Tab(1).Control(4)=   "lblNote"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "��ǩ(&I)"
      TabPicture(2)   =   "frmDiagHelp.frx":40C8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwMark"
      Tab(2).Control(1)=   "cmdMark"
      Tab(2).Control(2)=   "cmdDel"
      Tab(2).Control(3)=   "cmdShow(1)"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdShow 
         Caption         =   "��ʾ(&D)"
         Height          =   350
         Index           =   1
         Left            =   -72660
         TabIndex        =   11
         Top             =   5745
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&E)"
         Height          =   350
         Left            =   -72765
         TabIndex        =   9
         Top             =   510
         Width           =   1110
      End
      Begin VB.CommandButton cmdMark 
         Caption         =   "��ǰ����Ϊ��ǩ(&M)"
         Height          =   350
         Left            =   -74820
         TabIndex        =   8
         Top             =   495
         Width           =   2010
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "��ʾ(&D)"
         Enabled         =   0   'False
         Height          =   350
         Index           =   0
         Left            =   -72690
         TabIndex        =   7
         Top             =   6075
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "�г�������Ŀ(&L)"
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
            Caption         =   "��ҽ����(&2)"
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
            Caption         =   "��ҽ����(&1)"
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
         Caption         =   "����Ҫ���ҵı��롢���ơ����������(&W):"
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
      Caption         =   "�����ߴ�"
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
Dim mstrSteps As String    '���˺�ǰ����Ŀ��¼����

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String, aryTemp() As String

Private WithEvents objParentForm As Form
Attribute objParentForm.VB_VarHelpID = -1

Public Sub ShowMe(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemId As Long)
    '---------------------------------------------
    '���ܣ������ϼ�����Ҫ����ģ̬���ģ̬��ʾ���Ʋο�
    '��Σ�frmParent-�����壻
    '      blnModal-�Ƿ�ģ̬��ʾ��ͨ�����ϼ�����һ�£���
    '      lngItemId-Ҫ��ʾ�����ID����Ϊ0ʱ��ȱʡ����ʾĿ¼����
    '---------------------------------------------
    If lngItemId = 0 Then
        Me.tlbThis.Buttons("����").Visible = True: Me.tlbThis.Buttons("��ʾ").Visible = False
        Call cmdKind_Click(0)
        Call stbCatalog_Click(0)
    Else
        Me.tlbThis.Buttons("����").Visible = False: Me.tlbThis.Buttons("��ʾ").Visible = True
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
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���������ǩ\", UserInfo.�û���, strTemp)
End Sub

Private Sub cmdFind_Click()
    If Trim(Me.cboFind.Text) = "" Then
        MsgBox "��������ҵ�����", vbExclamation, gstrSysName
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
    gstrSql = "select distinct I.ID,I.����,I.����,decode(I.���,1,'��ҽ','��ҽ') as ���" & _
            " from �������Ŀ¼ I,������ϱ��� N" & _
            " where I.ID=N.���ID" & _
            "       and (I.���� like '" & Trim(Me.cboFind.Text) & "%'" & _
            "           or N.���� like '" & gstrMatch & Trim(Me.cboFind.Text) & "%'" & _
            "           or N.���� like '" & gstrMatch & Trim(Me.cboFind.Text) & "%')"
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            Set objItem = Me.lvwFind.ListItems.Add(, "_" & !ID, !����, "item", "item")
            objItem.SubItems(Me.lvwFind.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwFind.ColumnHeaders("���").Index - 1) = !���
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
    gstrSql = "select I.ID,I.����,I.����,decode(I.���,1,'��ҽ','��ҽ') as ���" & _
            " from �������Ŀ¼ I" & _
            " where I.ID=" & Me.lblItem.Tag
    With rsTemp
        Err = 0: On Error GoTo ErrHand
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        If Not .EOF Then
            Err = 0: On Error Resume Next
            Set objItem = Me.lvwMark.ListItems.Add(, "_" & !ID, !����, "item", "item")
            objItem.SubItems(Me.lvwMark.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwMark.ColumnHeaders("���").Index - 1) = !���
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
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���������ǩ\", UserInfo.�û���, strTemp)
    
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
    If Me.tlbThis.Buttons("����").Visible = True Then
        Me.stbCatalog.Visible = True: Me.picVBar.Visible = True
    Else
        Me.stbCatalog.Visible = False: Me.picVBar.Visible = False
    End If
End Sub

Private Sub Form_Load()
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    
    '����Ԫ����̬����
    Me.lvwFind.ListItems.Clear
    With Me.lvwFind.ColumnHeaders
        .Clear
        .Add , "����", "����", 2200
        .Add , "����", "����", 1000
        .Add , "���", "���", 1000
    End With
    With Me.lvwFind
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1: .SortOrder = lvwAscending
    End With
    Me.lvwMark.ListItems.Clear
    With Me.lvwMark.ColumnHeaders
        .Clear
        .Add , "����", "����", 2200
        .Add , "����", "����", 1000
        .Add , "���", "���", 1000
    End With
    With Me.lvwMark
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1: .SortOrder = lvwAscending
    End With
    
    With Me.hgdRefer
        .ColWidth(0) = 250
        .ColWidth(1) = Me.TextWidth("�ո�")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
    End With
        
    '��ȡ��ǩ��ʾ
    strTemp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���������ǩ\", UserInfo.�û���, "")
    If strTemp = "" Then Exit Sub
        
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "select I.ID,I.����,I.����,decode(I.���,1,'��ҽ','��ҽ') as ���" & _
                " from �������Ŀ¼ I" & _
                " where I.ID in (" & strTemp & ")"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            Set objItem = Me.lvwMark.ListItems.Add(, "_" & !ID, !����, "item", "item")
            objItem.SubItems(Me.lvwMark.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwMark.ColumnHeaders("���").Index - 1) = !���
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
    
    If Me.tlbThis.Buttons("����").Visible = True Then
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
        .ColWidth(1) = Me.TextWidth("�ո�")
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
        If .TextMatrix(.Row, 0) <> "��" And .TextMatrix(.Row, 0) <> "��" Then Exit Sub
        .Redraw = False
        For intCount = .Row + 1 To .Rows - 1
            If .TextMatrix(intCount, 0) = "��" Or .TextMatrix(intCount, 0) = "��" Then Exit For
            If .TextMatrix(.Row, 0) = "��" Then
                .RowHeight(intCount) = 255
            Else
                .RowHeight(intCount) = 0
            End If
        Next
        If .TextMatrix(.Row, 0) = "��" Then
            .TextMatrix(.Row, 0) = "��"
        Else
            .TextMatrix(.Row, 0) = "��"
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
    Case "����"
        Me.tlbThis.Buttons("����").Visible = False: Me.tlbThis.Buttons("��ʾ").Visible = True
        If Me.WindowState <> 2 Then
            Me.Left = Me.Left + (Me.stbCatalog.Width + Me.picVBar.Width)
            Me.Width = Me.Width - (Me.stbCatalog.Width + Me.picVBar.Width)
        End If
        Me.stbCatalog.Visible = False: Me.picVBar.Visible = False
        Call Form_Resize
    Case "��ʾ"
        Me.tlbThis.Buttons("����").Visible = True: Me.tlbThis.Buttons("��ʾ").Visible = False
        If Me.WindowState <> 2 Then
            If Me.Left - (Me.stbCatalog.Width + Me.picVBar.Width) >= 0 Then
                Me.Left = Me.Left - (Me.stbCatalog.Width + Me.picVBar.Width)
            Else
                Me.Left = 0
            End If
            Me.Width = Me.Width + (Me.stbCatalog.Width + Me.picVBar.Width)
        End If
        Me.stbCatalog.Visible = True: Me.picVBar.Visible = True
        
        '��λ����ǰ��ʾ��Ŀ���ڵ�Ŀ¼
        If Val(Me.picItem.Tag) = 1 Then
            Call cmdKind_Click(0)
        Else
            Call cmdKind_Click(1)
        End If
        Call stbCatalog_Click(0)
        Call Form_Resize
    Case "����"
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        Call zlShowRef(Val(aryTemp(intCount - 1)))
    Case "ǰ��"
        aryTemp = Split(mstrSteps, ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            If Val(aryTemp(intCount)) = Val(Me.lblItem.Tag) Then Exit For
        Next
        Call zlShowRef(Val(aryTemp(intCount + 1)))
    Case "��ӡ"
        Dim objPrint As New zlPrint1Grd
        Dim objRow As New zlTabAppRow
        Dim bytMode As Byte
        If Me.hgdRefer.TextMatrix(Me.hgdRefer.FixedRows, 2) = "" Then Exit Sub
        Set objPrint.Body = Me.hgdRefer
        With objPrint.Title
            .Text = Me.lblItem.Caption & ".��ϲο�"
            .Font.Size = 11
        End With
        objRow.Add Me.lblAlias.Caption
        objPrint.UnderAppRows.Add objRow
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Case "����"
        ShowHelp App.ProductName, Me.Hwnd, Me.Name ', Int((glngSys) / 100)
    Case "�˳�"
        Unload Me
    End Select
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If InStr(1, Node.Key, "_") = 0 Then Exit Sub
    Call zlShowRef(Mid(Node.Key, InStr(1, Node.Key, "_") + 1))
End Sub

Private Sub zlRefList()
    '---------------------------------------------
    '��д������Ŀ������Ŀ¼
    '---------------------------------------------
    Me.tvwList.Visible = False
    Me.tvwList.Nodes.Clear
    '��д����
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "select ID,�ϼ�ID,����,����,����" & _
                " From ������Ϸ���" & _
                " Where ���=" & IIf(Val(Me.tvwList.Tag) = 0, 1, 2) & _
                " start with �ϼ�ID is null" & _
                " connect by prior ID=�ϼ�ID"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwList.Nodes.Add(, , "C" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwList.Nodes.Add("C" & !�ϼ�ID, tvwChild, "C" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        
        gstrSql = "select R.����ID,I.ID,I.����,I.����" & _
                " from �������Ŀ¼ I,����������� R" & _
                " where I.ID=R.���ID and I.���=" & IIf(Val(Me.tvwList.Tag) = 0, 1, 2) & _
                " order by I.����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            Set objNode = Me.tvwList.Nodes.Add("C" & !����ID, tvwChild, "C" & !����ID & "_" & !ID, "[" & !���� & "]" & !����, "item")
            .MoveNext
        Loop
        
        If Me.tvwList.SelectedItem Is Nothing And .RecordCount <> 0 Then
            .MoveFirst
            Me.tvwList.Nodes("C" & !����ID & "_" & !ID).Selected = True
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
    '���ܣ���ʾָ����Ŀ�����Ʋο�����
    '------------------------------------------------
    '������ϸ��Ϣ��ʾ��
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
        gstrSql = "select distinct I.ID,I.����,N.����,N.����||decode(N.����,1,'',2,'(Ӣ����)','(����)') as ����,I.���" & _
                " from �������Ŀ¼ I,������ϱ��� N" & _
                " where I.ID=N.���ID and I.ID=" & lngItemId
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Do While Not .EOF
            If !���� = 1 Then
                Me.picItem.Tag = !���: Me.lblItem.Tag = !ID: Me.lblItem.Caption = IIf(IsNull(!����), "", !����)
            Else
                Me.lblAlias.Caption = Me.lblAlias.Caption & "��" & IIf(IsNull(!����), "", !����)
            End If
            .MoveNext
        Loop
        If Me.lblAlias.Caption <> "" Then Me.lblAlias.Caption = Mid(Me.lblAlias.Caption, 2)
        
        gstrSql = "select ��Ŀ���,��Ŀ���,nvl(֤�����,0) as ֤�����,0 as �����к�,decode(nvl(֤������,''),'',�ο���Ŀ,֤������) as ����" & _
                " from ������ϲο� " & _
                " where ���id=" & lngItemId & _
                " union" & _
                " select ��Ŀ���,��Ŀ���,nvl(֤�����,0) as ֤�����,�����к�,decode(nvl(֤������,''),'',�����ı�,�ο���Ŀ||'��'||�����ı�) as ����" & _
                " from ������ϲο�" & _
                " where ���id=" & lngItemId & " and length(ltrim(nvl(�����ı�,'')))<>0" & _
                " order by ��Ŀ���,֤�����,�����к�"
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
            If !�����к� = 0 Then
                If !��Ŀ��� = 1 Then
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 0) = "��"
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = "��" & !���� & "��"
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = "��" & !���� & "��"
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = True
                Else
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = ""
                    If !֤����� = 0 Then
                        Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = "��" & !���� & "��"
                    Else
                        Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = !֤����� & "." & !����
                    End If
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = False
                End If
            Else
                If !��Ŀ��� = 1 Then
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = Space(4) & !����
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = Space(4) & !����
                    Me.hgdRefer.MergeRow(.AbsolutePosition + Me.hgdRefer.FixedRows - 1) = True
                Else
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 1) = ""
                    Me.hgdRefer.TextMatrix(.AbsolutePosition + Me.hgdRefer.FixedRows - 1, 2) = Space(4) & !����
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
        Me.tlbThis.Buttons("����").Enabled = True
    Else
        Me.tlbThis.Buttons("����").Enabled = False
    End If
    If intCount < UBound(aryTemp) Then
        Me.tlbThis.Buttons("ǰ��").Enabled = True
    Else
        Me.tlbThis.Buttons("ǰ��").Enabled = False
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '���ݵ������ݵ�������������и߶ȣ��Ա�֤���ݵ�������ʾ
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
