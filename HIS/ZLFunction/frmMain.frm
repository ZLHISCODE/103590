VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "��������"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvw 
      Height          =   3780
      Left            =   45
      TabIndex        =   0
      Top             =   1200
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   6668
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "������"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "������"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "������"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "˵��"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "״̬"
         Object.Width           =   970
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   1125
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   1984
      BandCount       =   2
      _CBWidth        =   8670
      _CBHeight       =   1125
      _Version        =   "6.7.8988"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Caption2        =   "Ӧ��ϵͳ"
      Child2          =   "cboSys"
      MinHeight2      =   300
      Width2          =   1125
      UseCoolbarColors2=   0   'False
      NewRow2         =   -1  'True
      Begin VB.ComboBox cboSys 
         Height          =   300
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   780
         Width           =   7635
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   8415
         _ExtentX        =   14843
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
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Test"
               Description     =   "����"
               Object.ToolTipText     =   "���Ժ���"
               Object.Tag             =   "����"
               ImageKey        =   "Test"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Test_"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸ĺ���"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Del"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ������"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit_"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Description     =   "�鿴"
               Object.ToolTipText     =   "�б�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Icon"
                     Object.Tag             =   "��ͼ��(&I)"
                     Text            =   "��ͼ��(&I)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "Сͼ��(&S)"
                     Text            =   "Сͼ��(&S)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "�б�(&L)"
                     Text            =   "�б�(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "��ϸ����(&D)"
                     Text            =   "��ϸ����(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5025
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":014A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10213
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
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
   Begin MSComctlLib.ImageList img16 
      Left            =   2745
      Top             =   735
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
            Picture         =   "frmMain.frx":09DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3315
      Top             =   735
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
            Picture         =   "frmMain.frx":0B38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   0
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E52
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":106C
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1286
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14A0
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16BA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18D4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AEE
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D08
            Key             =   "Test"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   585
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F22
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":213C
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2356
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2570
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":278A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29A4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BBE
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DD8
            Key             =   "Test"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileTest 
         Caption         =   "���Ժ���(&T)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFile_Test_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "��������(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModi 
         Caption         =   "�޸ĺ���(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ������(&R)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&B)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&L)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuView_View 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_reFlash 
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
         Caption         =   "&WEB�ϵ�����"
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
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnRegModi As Boolean
Private mstrOwner As String
Private mintIndex As Integer

Private Sub cboSys_Click()
    If cboSys.ListIndex = mintIndex Then Exit Sub
    mintIndex = cboSys.ListIndex
    
    mstrOwner = Split(cboSys.Text, "(")(UBound(Split(cboSys.Text, "(")))
    mstrOwner = Left(mstrOwner, Len(mstrOwner) - 1)
    
    Call ReadFunc
End Sub

Private Sub Form_Load()
    lvw.ColumnHeaders(2).Position = 1
    RestoreWinState Me, App.ProductName
    
    mintIndex = -1
    mstrOwner = ""
    
    'mblnRegModi = (zlRegReport(GetUnitInfo("ע����")) And 2) = 2
    mblnRegModi = True
    
    Call ReadSystem
    
    Call SetEditable
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long
    Dim staH As Long
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(sta.Visible, sta.Height, 0)
    
    With lvw
        .Left = Me.ScaleLeft
        .Top = cbrH + Me.ScaleTop
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - cbrH - staH
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub lvw_GotFocus()
    Call SetEditable
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call SetEditable
End Sub

Private Sub lvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetEditable
    If lvw.HitTest(X, Y) Is Nothing Then
        If Button = 2 Then
            PopupMenu mnuView, 2
        ElseIf Button = 1 Then
            sta.Panels(2) = "�� " & lvw.ListItems.Count & " ������"
        End If
    Else
        If Button = 2 And mnuEdit.Visible Then PopupMenu mnuEdit, 2
    End If
End Sub

Private Sub mnuEditAdd_Click()
    frmSQLEdit.mblnModi = False
    frmSQLEdit.mstrOwner = mstrOwner
    frmSQLEdit.mlngSys = cboSys.ItemData(cboSys.ListIndex)
    frmSQLEdit.Show 1, Me
    
    If gblnOK Then
        'ˢ�²���Ȩ��
        Me.Refresh
        Screen.MousePointer = 11
        grsObject.Requery
        Screen.MousePointer = 0
        If Not lvw.SelectedItem Is Nothing Then Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub mnuEditDel_Click()
    Dim intIdx As Integer
    
    If MsgBox("ȷʵҪɾ���ú�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    On Error GoTo errH
    gcnOracle.Execute "Delete From zlFunctions Where ϵͳ=" & cboSys.ItemData(cboSys.ListIndex) & " And ������=" & Mid(lvw.SelectedItem.Key, 2)
    On Error Resume Next
    gcnOracle.Execute "Drop Function " & mstrOwner & "." & lvw.SelectedItem.Text
    On Error GoTo 0
    
    'ˢ�²���Ȩ��
    Me.Refresh
    Screen.MousePointer = 11
    grsObject.Requery
    Screen.MousePointer = 0
    
    intIdx = lvw.SelectedItem.Index
    lvw.ListItems.Remove intIdx
    If lvw.ListItems.Count > 0 Then
        If intIdx <= lvw.ListItems.Count Then
            lvw.ListItems(intIdx).Selected = True
        Else
            lvw.ListItems(lvw.ListItems.Count).Selected = True
        End If
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    Else
        sta.Panels(2) = "û�ж��庯��"
        Call SetEditable
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditModi_Click()
    Dim objPars As FuncPars, strObj As String, strText As String
    
    '��麯���Ƿ����ExecuteȨ��(�����߻�DBAһ������)
    grsObject.Filter = "OWNER='" & UCase(mstrOwner) & "' And OBJECT_TYPE='FUNCTION' And OBJECT_NAME='" & UCase(lvw.SelectedItem.Text) & "'"
    If grsObject.EOF Then
        MsgBox "��ǰ�û�û��Ȩ��ִ�иú�����", vbInformation, App.Title
        Exit Sub
    End If
    
    '��麯�������Ƿ������ȷ���Ƿ���Ȩ��ִ��
    strText = GetFunSource(mstrOwner, lvw.SelectedItem.Text)
    If strText = "" Then
        MsgBox "���ܶ�ȡ��������,�����û��Ȩ��ִ�иú�����", vbInformation, App.Title
        Exit Sub
    End If
    
    '��麯������ѡ���������Ƿ����SelectȨ��
    strObj = CheckParPrivs(cboSys.ItemData(cboSys.ListIndex), mstrOwner, Val(Mid(lvw.SelectedItem.Key, 2)))
    If strObj <> "" Then
        MsgBox "��ǰ�û�û��Ȩ�޷��ʺ�������ѡ������" & vbCrLf & _
                "��һЩ�������Щ���󲻴���,��������", vbInformation, App.Title
        Exit Sub
    End If
    
    frmSQLEdit.mblnModi = True
    frmSQLEdit.mstrOwner = mstrOwner
    frmSQLEdit.mlngSys = cboSys.ItemData(cboSys.ListIndex)
    frmSQLEdit.Show 1, Me
    If gblnOK And Not lvw.SelectedItem Is Nothing Then Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub mnuFileTest_Click()
    Dim objPars As FuncPars, strObj As String
    Dim strText As String, strPars As String
    Dim i As Integer
    
    '��麯���Ƿ����ExecuteȨ��(�����߻�DBAһ������)
    grsObject.Filter = "OWNER='" & UCase(mstrOwner) & "' And OBJECT_TYPE='FUNCTION' And OBJECT_NAME='" & UCase(lvw.SelectedItem.Text) & "'"
    If grsObject.EOF Then
        MsgBox "��ǰ�û�û��Ȩ��ִ�иú�����", vbInformation, App.Title
        Exit Sub
    End If
    
    '��麯�������Ƿ������ȷ���Ƿ���Ȩ��ִ��
    strText = GetFunSource(mstrOwner, lvw.SelectedItem.Text)
    If strText = "" Then
        MsgBox "���ܶ�ȡ��������,�����û��Ȩ��ִ�иú�����", vbInformation, App.Title
        Exit Sub
    End If
    
    '��麯������ѡ���������Ƿ����SelectȨ��
    strObj = CheckParPrivs(cboSys.ItemData(cboSys.ListIndex), mstrOwner, Val(Mid(lvw.SelectedItem.Key, 2)))
    If strObj <> "" Then
        MsgBox "��ǰ�û�û��Ȩ�޷��ʺ�������ѡ����" & vbCrLf & _
                "�е�һЩ�������Щ���󲻴��ڣ�", vbInformation, App.Title
        Exit Sub
    End If
    
    '������û�ж�����������
    strPars = GetFuncPars(strText)
    If strPars = "" Then
        Set objPars = New FuncPars
        strObj = ExeFunction(mstrOwner, lvw.SelectedItem.Text, strPars, objPars)
        If strObj Like "ERROR*" Then
            MsgBox "ִ��ʧ�ܣ�" & vbCrLf & vbCrLf & Mid(strObj, 6), vbInformation, App.Title
        Else
            MsgBox "ִ�гɹ���" & vbCrLf & vbCrLf & lvw.SelectedItem.Text & " = " & strObj & "    ", vbInformation, App.Title
        End If
    Else
        frmParInput.mlngSys = cboSys.ItemData(cboSys.ListIndex)
        frmParInput.mstrOwner = mstrOwner
        frmParInput.mintNum = Val(Mid(lvw.SelectedItem.Key, 2))
        frmParInput.mstrFun = lvw.SelectedItem.Text
        frmParInput.Show 1, Me
    End If
End Sub

Private Sub mnuView_reFlash_Click()
    mintIndex = -1
    Call cboSys_Click
End Sub

Private Sub mnuViewStatus_Click()
    Sub�鿴�˵� mnuViewStatus.Caption
End Sub

Private Sub mnuViewToolButton_Click()
    Sub�鿴�˵� mnuViewToolButton.Caption
End Sub

Private Sub mnuViewToolText_Click()
    Sub�鿴�˵� mnuViewToolText.Caption
End Sub

Private Sub Sub�鿴�˵�(ByVal mnuLable As String)
    Dim i As Integer
    Select Case mnuLable
        Case "��׼��ť(&B)"
            mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
            mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
            cbr.Visible = Not cbr.Visible
            Form_Resize
        Case "�ı���ǩ(&L)"
            mnuViewToolText.Checked = Not mnuViewToolText.Checked
            For i = 1 To tbr.Buttons.Count
                If mnuViewToolText.Checked Then
                    tbr.Buttons(i).Caption = tbr.Buttons(i).Tag
                Else
                    tbr.Buttons(i).Caption = ""
                End If
            Next
            cbr.Bands(1).MinHeight = tbr.ButtonHeight
            Form_Resize
        Case "״̬��(&S)"
            mnuViewStatus.Checked = Not mnuViewStatus.Checked
            sta.Visible = Not sta.Visible
            Form_Resize
    End Select
End Sub

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Call Form_Resize
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFileQuit_Click
        Case "View"
            Call SetView((lvw.View + 1) Mod 4)
        Case "Add"
            mnuEditAdd_Click
        Case "Modi"
            mnuEditModi_Click
        Case "Del"
            mnuEditDel_Click
        Case "Test"
            mnuFileTest_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub mnuView_View_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub SetView(bytStyle As Byte)
'���ܣ������б���ʾ��ʽ
'������bytstyle=0-��ͼ��,1-Сͼ��,2-�б�,3-��ϸ����
    mnuView_View(0).Checked = False
    mnuView_View(1).Checked = False
    mnuView_View(2).Checked = False
    mnuView_View(3).Checked = False
    mnuView_View(bytStyle).Checked = True
    
    On Error Resume Next
    lvw.View = bytStyle
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Icon"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelpFunc(Me.hwnd, "main", 0)
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub ReadFunc()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, objItem As ListItem
    Dim strSQL As String, j As Integer
    
    On Error GoTo errH
    
    lvw.ListItems.Clear
    If cboSys.ListIndex = -1 Then Exit Sub
    
    strSQL = "Select A.*,B.STATUS From zlFunctions A,All_Objects B" & _
        " Where A.ϵͳ=" & cboSys.ItemData(cboSys.ListIndex) & _
        " And B.Owner='" & mstrOwner & "' And B.Object_Type='FUNCTION'" & _
        " And Upper(A.������)=B.Object_Name" & _
        " Order by A.������"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "ReadFunc")
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvw.ListItems.Add(, "_" & rsTmp!������, rsTmp!������, 1, 1)
        objItem.SubItems(1) = rsTmp!������
        objItem.SubItems(2) = IIf(IsNull(rsTmp!������), "", rsTmp!������)
        objItem.SubItems(3) = IIf(IsNull(rsTmp!˵��), "", rsTmp!˵��)
        If rsTmp!Status <> "VALID" Then
            objItem.SubItems(4) = "��"
            objItem.ForeColor = &H808080
            For j = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(j).ForeColor = &H808080
            Next
        End If
        
        rsTmp.MoveNext
    Next
    Call SetEditable
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetEditable()
    If cboSys.ListIndex = -1 Then
        mnuEdit.Visible = False
        tbr.Buttons("Add").Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Modi").Visible = False
        tbr.Buttons("Edit_").Visible = False
        
        mnuFileTest.Visible = False
        mnuFile_Test_.Visible = False
        tbr.Buttons("Test").Visible = False
        tbr.Buttons("Test_").Visible = False
        
        Exit Sub
    Else
        mnuFileTest.Visible = True
        mnuFile_Test_.Visible = True
        tbr.Buttons("Test").Visible = True
        tbr.Buttons("Test_").Visible = True
    End If
    
    If UCase(mstrOwner) = UCase(gstrDBUser) Then 'Or gblnDBA Then
        mnuEdit.Visible = mblnRegModi
        tbr.Buttons("Add").Visible = mblnRegModi
        tbr.Buttons("Del").Visible = mblnRegModi
        tbr.Buttons("Modi").Visible = mblnRegModi
        tbr.Buttons("Edit_").Visible = mblnRegModi
        
        mnuEditModi.Enabled = Not lvw.SelectedItem Is Nothing
        mnuEditDel.Enabled = Not lvw.SelectedItem Is Nothing
        mnuFileTest.Enabled = Not lvw.SelectedItem Is Nothing
        
        tbr.Buttons("Modi").Enabled = Not lvw.SelectedItem Is Nothing
        tbr.Buttons("Del").Enabled = Not lvw.SelectedItem Is Nothing
        tbr.Buttons("Test").Enabled = Not lvw.SelectedItem Is Nothing
    Else
        mnuEdit.Visible = False
        tbr.Buttons("Add").Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Modi").Visible = False
        tbr.Buttons("Edit_").Visible = False
        
        mnuFileTest.Enabled = Not lvw.SelectedItem Is Nothing
        tbr.Buttons("Test").Enabled = Not lvw.SelectedItem Is Nothing
    End If
End Sub

Public Sub ReadSystem()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strSQL As String
    
    On Error GoTo errH
    
    cboSys.Clear
    strSQL = "Select * From zlSystems Order by ���"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "ReadSystem")
    
    For i = 1 To rsTmp.RecordCount
        cboSys.AddItem rsTmp!���� & "-" & Right(Format(rsTmp!���, "00000"), 2) & "(" & rsTmp!������ & ")"
        cboSys.ItemData(cboSys.NewIndex) = rsTmp!���
        If rsTmp!������ = gstrDBUser And cboSys.ListIndex = -1 Then cboSys.ListIndex = cboSys.NewIndex
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

