VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMedInOutClass 
   AutoRedraw      =   -1  'True
   Caption         =   "�������"
   ClientHeight    =   7170
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11610
   FillColor       =   &H00FF0000&
   Icon            =   "frmMedInOutClass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmMedInOutClass.frx":1CFA
   ScaleHeight     =   7170
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Tag             =   "15"
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedInOutClass.frx":2004
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15399
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
   Begin MSComctlLib.ImageList ImgLvw����Small 
      Left            =   4890
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgLvwBig 
      Left            =   4350
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgLvwSmall 
      Left            =   4890
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarBlack 
      Left            =   6750
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarColor 
      Left            =   6180
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView Lvw�������� 
      Height          =   1725
      Left            =   2040
      TabIndex        =   4
      Top             =   1980
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   3043
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��������"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "˵��"
         Object.Width           =   9701
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw������ 
      Height          =   1185
      Left            =   2010
      TabIndex        =   3
      Top             =   750
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2090
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "���"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComCtl3.CoolBar Cbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   1164
      BandCount       =   1
      FixedOrder      =   -1  'True
      _CBWidth        =   11610
      _CBHeight       =   660
      _Version        =   "6.7.8988"
      Child1          =   "Tbar"
      MinHeight1      =   600
      Width1          =   4995
      Key1            =   "Common"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   600
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   1058
         ButtonWidth     =   820
         ButtonHeight    =   1058
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "�鿴"
               Object.Tag             =   "�鿴"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Big"
                     Text            =   "��ͼ��(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Text            =   "Сͼ��(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Text            =   "�б�(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Text            =   "��ϸ����(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
            EndProperty
         EndProperty
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   10080
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "����"
            Top             =   210
            Width           =   1425
         End
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   9480
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   5
            Top             =   210
            Width           =   495
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   74
               Width           =   495
            End
         End
      End
   End
   Begin VB.Image ImgUpDown_S 
      Height          =   45
      Left            =   2040
      MousePointer    =   7  'Size N S
      Top             =   1920
      Width           =   3345
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintset 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolS 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolT 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu mnuView2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewBill 
         Caption         =   "��������ʾҩƷ������(&B)"
      End
      Begin VB.Menu mnuView3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "WEB�ϵ�����(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedInOutClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BlnStartUp As Boolean
Private RecClient As New ADODB.Recordset
Private strSQL As String                    '��дSQL���
Private BlnEditReturn As Boolean            '�޸ĳɹ����
Private mlngMode As Long
Private mstrPrivs As String                              'Ȩ�޴�
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mrsFind As ADODB.Recordset
Private mstrFindValue As String             '��¼��ѯ��ֵ

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
    Form_Resize
End Sub

Private Sub Form_Load()
    BlnStartUp = False
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    '--�ָ����弰�ؼ����״̬--
    RestoreWinState Me, App.ProductName
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '--������ز˵�--
    Lvw��������.View = lvwReport
    mnuViewToolT.Enabled = mnuViewToolS.Checked
    ClearViewState Lvw������.View + 1
    
    If LoadInIcon = False Then Exit Sub
    Call LoadInLvw
    Call Ȩ�޿���
    
    BlnStartUp = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    SetParent txtFind.hwnd, Tbar.hwnd
    SetParent picFind.hwnd, Tbar.hwnd
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    If Me.WindowState = 1 Then Exit Sub
    
    With Cbar
        .Bands(1).MinHeight = Tbar.Height
        Set .Bands(1).Child = Tbar
    End With
    
    With ImgUpDown_S
        .Left = 0
        .Top = (Me.ScaleHeight - stbThis.Height - Cbar.Top + Cbar.Height) * 0.5
        .Width = Me.ScaleWidth - .Left
    End With
    
    With Lvw������
        .Left = 0
        .Top = IIF(Cbar.Visible, Cbar.Height, 0)
        .Width = ImgUpDown_S.Width
        .Height = ImgUpDown_S.Top - .Top
    End With
    
    With Lvw��������
        .Left = 0
        .Top = ImgUpDown_S.Top + ImgUpDown_S.Height
        .Width = ImgUpDown_S.Width
        .Height = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Function LoadInIcon() As Boolean
    '--Ϊ���ؼ�װ��ͼ��--
    On Error Resume Next
    Err = 0
    LoadInIcon = False
    
    '--������Tbar--
    With ImgTbarBlack
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add 1, , LoadResPicture("BPREVIEW", vbResIcon)
        .ListImages.Add 2, , LoadResPicture("BPRINT", vbResIcon)
        .ListImages.Add 3, , LoadResPicture("BADD", vbResIcon)
        .ListImages.Add 4, , LoadResPicture("BMODIFY", vbResIcon)
        .ListImages.Add 5, , LoadResPicture("BDELETE", vbResIcon)
        .ListImages.Add 6, , LoadResPicture("BVIEW", vbResIcon)
        .ListImages.Add 7, , LoadResPicture("BHELP", vbResIcon)
        .ListImages.Add 8, , LoadResPicture("BEXIT", vbResIcon)
    End With
    With ImgTbarColor
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add 1, , LoadResPicture("CPREVIEW", vbResIcon)
        .ListImages.Add 2, , LoadResPicture("CPRINT", vbResIcon)
        .ListImages.Add 3, , LoadResPicture("CADD", vbResIcon)
        .ListImages.Add 4, , LoadResPicture("CMODIFY", vbResIcon)
        .ListImages.Add 5, , LoadResPicture("CDELETE", vbResIcon)
        .ListImages.Add 6, , LoadResPicture("CVIEW", vbResIcon)
        .ListImages.Add 7, , LoadResPicture("CHELP", vbResIcon)
        .ListImages.Add 8, , LoadResPicture("CEXIT", vbResIcon)
    End With
    With Tbar
        Set .ImageList = ImgTbarBlack
        Set .HotImageList = ImgTbarColor
        .Buttons("Preview").Image = 1
        .Buttons("Print").Image = 2
        .Buttons("Add").Image = 3
        .Buttons("Modify").Image = 4
        .Buttons("Delete").Image = 5
        .Buttons("View").Image = 6
        .Buttons("Help").Image = 7
        .Buttons("Exit").Image = 8
    End With
    Cbar.Bands("Common").MinHeight = Tbar.Height
    
    '--�б�Lvw������--
    With ImgLvwBig
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("CLASS", vbResIcon)
    End With
    With ImgLvwSmall
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add , , LoadResPicture("CLASS", vbResIcon)
    End With
    With Lvw������
        Set .SmallIcons = ImgLvwSmall
        Set .Icons = ImgLvwBig
    End With
    
    With ImgLvw����Small
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add , , LoadResPicture("BILL1", vbResIcon)
    End With
    With Lvw��������
        Set .SmallIcons = ImgLvw����Small
    End With
    
    If Err <> 0 Then
        MsgBox "�����Դ�ļ���ʧ�����������������ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function

Private Sub ImgUpDown_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgUpDown_S
        If .Top + Y < 2000 Then Exit Sub
        If .Top + Y > Me.ScaleHeight - 2500 Then Exit Sub
        
        .Move .Left, .Top + Y
    End With
    
    With Lvw������
        .Left = 0
        .Top = IIF(Cbar.Visible, Cbar.Height, 0)
        .Width = ImgUpDown_S.Width
        .Height = ImgUpDown_S.Top - .Top
    End With
    
    With Lvw��������
        .Left = 0
        .Top = ImgUpDown_S.Top + ImgUpDown_S.Height
        .Width = ImgUpDown_S.Width
        .Height = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
End Sub

Private Sub Lvw������_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw������
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw������_DblClick()
    If Lvw������.ListItems.Count = 0 Then Exit Sub
    
    If InStr(1, mstrPrivs, "��ɾ��") <> 0 Then mnuEditModify_Click
End Sub

Private Sub Lvw������_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '--������������������Щ����--
    
    On Error GoTo ErrHandle
    strSQL = "Select ����,����,˵�� From ҩƷ���ݷ��� Where ���� In" & _
             " (Select ���� From ҩƷ�������� Where ���ID=[1]) Order by ����"
    Set RecClient = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(Item.Key, 3)))

    LoadInBill
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Lvw������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Lvw������_DblClick
End Sub

Private Sub Lvw������_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Lvw������.ListItems.Count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim ItemThis As ListItem
    
    On Error Resume Next
    Err = 0
    
    With Lvw������
        Set ItemThis = .HitTest(X, Y)
        If Err <> 0 Then Exit Sub
        
        ItemThis.Selected = True
        .SelectedItem.Selected = True
    End With
    Lvw������_ItemClick Lvw������.SelectedItem
    If InStr(1, mstrPrivs, "��ɾ��") <> 0 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Lvw��������_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw��������
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub mnuEditAdd_Click()
    With frmEditInOutClass
        .EditState = 1
        .ϵ�� = 1
        .Show 1, Me
    End With
    If BlnEditReturn Then mnuViewRefresh_Click
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo ErrHand
    If MsgBox("��ȷ��Ҫɾ������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "zl_ҩƷ������_delete (" & Mid(Lvw������.SelectedItem.Key, 3) & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ɾ��ҩƷ������")
    
    mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
    With frmEditInOutClass
        .EditState = 2
        .���ID = Mid(Lvw������.SelectedItem.Key, 3)
        .���� = Lvw������.SelectedItem
        .���� = Lvw������.SelectedItem.SubItems(1)
        .ϵ�� = IIF(Lvw������.SelectedItem.SubItems(2) = "���", 1, -1)
        .Show 1, Me
    End With
    If BlnEditReturn Then mnuViewRefresh_Click
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub mnufilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintset_Click()
    zlPrintSet
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
End Sub

Private Sub mnuViewBill_Click()
    With frmByBillShow
        .Show 1, Me
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    LoadInLvw
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuhelpTitle_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    ClearViewState Index
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = mnuViewStatus.Checked Xor True
    stbThis.Visible = mnuViewStatus.Checked
    
    Form_Resize
End Sub

Private Sub MnuViewToolS_Click()
    mnuViewToolS.Checked = mnuViewToolS.Checked Xor True
    Cbar.Visible = mnuViewToolS.Checked
    mnuViewToolT.Enabled = mnuViewToolS.Checked
    
    Form_Resize
End Sub

Private Sub MnuViewToolT_Click()
    mnuViewToolT.Checked = mnuViewToolT.Checked Xor True
    If mnuViewToolT.Checked Then
        With Tbar
            .Buttons("Preview").Caption = .Buttons("Preview").Tag
            .Buttons("Print").Caption = .Buttons("Print").Tag
            .Buttons("Add").Caption = .Buttons("Add").Tag
            .Buttons("Modify").Caption = .Buttons("Modify").Tag
            .Buttons("Delete").Caption = .Buttons("Delete").Tag
            .Buttons("View").Caption = .Buttons("View").Tag
            .Buttons("Help").Caption = .Buttons("Help").Tag
            .Buttons("Exit").Caption = .Buttons("Exit").Tag
        End With
    Else
        With Tbar
            .Buttons("Preview").Caption = ""
            .Buttons("Print").Caption = ""
            .Buttons("Add").Caption = ""
            .Buttons("Modify").Caption = ""
            .Buttons("Delete").Caption = ""
            .Buttons("View").Caption = ""
            .Buttons("Help").Caption = ""
            .Buttons("Exit").Caption = ""
        End With
    End If
    Cbar.Bands(1).MinHeight = Tbar.Height
    
    Form_Resize
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnufilePreview_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Add"
        mnuEditAdd_Click
    Case "Modify"
        mnuEditModify_Click
    Case "Delete"
        mnuEditDelete_Click
    Case "View"
        ClearViewState IIF(Lvw������.View < lvwReport, Lvw������.View + 2, 1)
    Case "Help"
        mnuhelpTitle_Click
    Case "Exit"
        mnuFileExit_Click
    End Select
End Sub

Private Sub Tbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    ClearViewState ButtonMenu.Index
End Sub

Private Sub Tbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Function LoadInLvw() As Boolean
    '--װ��ListView--
    Dim ItemThis As ListItem
    
    Lvw������.ListItems.Clear
    stbThis.Panels(2) = ""
    
    On Error GoTo ErrHandle
    
'        If .State = 1 Then .Close
    strSQL = "Select ID,����,����,Decode(ϵ��,1,'���','����') ϵ�� From ҩƷ������ Order by ϵ��"
        
'        Call SQLTest(App.Title, Me.Caption, strSQL)
    Set RecClient = zldatabase.OpenSQLRecord(strSQL, "LoadInLvw")
'        Call SQLTest
    With RecClient
        Do While Not .EOF
            Set ItemThis = Lvw������.ListItems.Add(, "K_" & !ID, !����, 1, 1)
            ItemThis.SubItems(1) = !����
            ItemThis.SubItems(2) = !ϵ��
            .MoveNext
        Loop
        
    End With
    
    With Lvw������
        If .ListItems.Count <> 0 Then
            .ListItems(1).Selected = True
            .SelectedItem.Selected = True
            Lvw������_ItemClick Lvw������.SelectedItem
        End If
    End With
    
    SetMenuAndButton
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadInBill() As Boolean
    '--װ��ListView--
    Dim ItemThis As ListItem
    
    Lvw��������.ListItems.Clear
    With RecClient
        If .EOF Then Exit Function
        
        Do While Not .EOF
            Set ItemThis = Lvw��������.ListItems.Add(, "K_" & !����, !����, , 1)
            ItemThis.SubItems(1) = !˵��
            .MoveNext
        Loop
        
    End With
    
    With Lvw��������
        .ListItems(1).Selected = True
        .SelectedItem.Selected = True
    End With
End Function

Private Function ClearViewState(ByVal Index As Integer)
    '--������ʾ״̬--
    Dim intIndex As Integer
    For intIndex = 1 To 4
        mnuViewIcon(intIndex).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    
    Lvw������.View = Index - 1
End Function

Public Function EditReturn(ByVal EditValue As Boolean)
    BlnEditReturn = EditValue
End Function
    
Public Function subPrint(ByVal bytMode As Byte)
    Dim objPrint As New zlPrintLvw
    
    objPrint.Title.Text = "ҩƷ������"
    Set objPrint.Body.objData = Lvw������
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zldatabase.Currentdate, "yyyy��MM��dd��")

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
End Function

Private Sub SetMenuAndButton()
    
    '���ð�ť�빤����
    mnuFilePrint.Enabled = (Lvw������.ListItems.Count <> 0)
    mnuFilePreview.Enabled = mnuFilePrint.Enabled
    mnuFileExcel.Enabled = mnuFilePrint.Enabled
    Tbar.Buttons("Preview").Enabled = mnuFilePrint.Enabled
    Tbar.Buttons("Print").Enabled = mnuFilePrint.Enabled
    
    mnuEditModify.Enabled = (Lvw������.ListItems.Count <> 0)
    mnuEditDelete.Enabled = (Lvw������.ListItems.Count <> 0)
    Tbar.Buttons("Modify").Enabled = mnuEditModify.Enabled
    Tbar.Buttons("Delete").Enabled = mnuEditDelete.Enabled
End Sub

Private Sub Ȩ�޿���()
    If InStr(1, mstrPrivs, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
        Tbar.Buttons("Add").Visible = False
        Tbar.Buttons("Modify").Visible = False
        Tbar.Buttons("Delete").Visible = False
        Tbar.Buttons("Split1").Visible = False
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            
            gstrSQL = "select * from ҩƷ������ where ���� like [1] or ���� like [1]"
            Set mrsFind = zldatabase.OpenSQLRecord(gstrSQL, "ҩƷ�������ѯ", txtFind.Text & "%")
            Call LocateItem
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                Call LocateItem
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call LocateItem
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LocateItem()
    Dim strTemp As String
    
    txtFind.SetFocus
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    If mrsFind.RecordCount = 0 Then
        MsgBox " û���ҵ�������������Ϣ��", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        MsgBox " �Ѿ���λ�������ҵ�����Ϣ������������������", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    
    Lvw������.ListItems("K_" & mrsFind("ID")).Selected = True
    Lvw������.SelectedItem.EnsureVisible
End Sub
