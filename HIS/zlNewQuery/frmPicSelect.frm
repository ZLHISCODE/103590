VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPicSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ͼƬ"
   ClientHeight    =   4530
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7740
   Icon            =   "frmPicSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgTmp 
      Left            =   5460
      Top             =   3750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4920
      Top             =   3795
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboRang 
      Height          =   300
      ItemData        =   "frmPicSelect.frx":000C
      Left            =   1020
      List            =   "frmPicSelect.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   15
      Width           =   2010
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   6780
      Top             =   3660
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
            Picture         =   "frmPicSelect.frx":007E
            Key             =   "pic"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   5910
      Top             =   3615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicSelect.frx":0ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicSelect.frx":126A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicSelect.frx":1604
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicSelect.frx":199E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPicSelect.frx":1D38
            Key             =   "pic"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Height          =   300
      Left            =   1095
      TabIndex        =   4
      Top             =   3600
      Width           =   2550
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Left            =   1095
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4065
      Width           =   2565
   End
   Begin MSComctlLib.Toolbar tbrThis 
      Height          =   345
      Left            =   3030
      TabIndex        =   9
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   609
      ButtonWidth     =   2064
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ilsMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "ɾ��(&D)"
            Key             =   "ɾ��"
            Object.ToolTipText     =   "ɾ��"
            Object.Tag             =   "ɾ��"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "����(&N)"
            Key             =   "����"
            Object.ToolTipText     =   "����"
            Object.Tag             =   "����"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�б�(&L)"
            Key             =   "�б�"
            Object.ToolTipText     =   "�б�"
            Object.Tag             =   "�б�"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List0"
                  Text            =   "��ͼ��"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List1"
                  Text            =   "Сͼ��"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List2"
                  Text            =   "�б�"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List3"
                  Text            =   "��ϸ����"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&H)"
            Key             =   "����"
            Object.ToolTipText     =   "����"
            Object.Tag             =   "����"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3750
      TabIndex        =   7
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3750
      TabIndex        =   8
      Top             =   4035
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3165
      Left            =   15
      TabIndex        =   2
      Top             =   360
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   5583
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ilsMenu"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ͼƬ����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���(����)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�߶�(����)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�޸�����"
         Object.Width           =   3175
      EndProperty
   End
   Begin zl9NewQuery.ctlPicture picBack 
      Height          =   3165
      Left            =   5280
      TabIndex        =   10
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5583
   End
   Begin VB.Label Label4 
      Caption         =   "ͼƬ��Χ(&R)"
      Height          =   210
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "ͼƬ����(&T)"
      Height          =   210
      Left            =   30
      TabIndex        =   5
      Top             =   4110
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "ͼƬ����(&N)"
      Height          =   210
      Left            =   30
      TabIndex        =   3
      Top             =   3660
      Width           =   1245
   End
End
Attribute VB_Name = "frmPicSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnLoad As Boolean
Private mblnFirst As Boolean
Private mOK As Boolean
Private mKey As Long
Private mType As String

Private mvarDefaultRange As String
Private mvarDefaultType As String

Private frmParent As Object

Public Function OpenPictureBox(frmMain As Object, ByVal Title As String, ByVal strType As String, Key As Long, DefaultRange As String, DefaultType As String) As Boolean
    'strType="0-ҽԺ����ͼ��;1-ҽԺ����ͼƬ;2-��淢��ͼƬ;3-��Ŀ����ͼ��;4-����ͼƬ;9-����ͼƬ"
        
    mType = ";" & strType & ";"
    frmPicSelect.Caption = Title
    mvarDefaultRange = DefaultRange
    mvarDefaultType = DefaultType
    Set frmParent = frmMain
    
    frmPicSelect.Show 1, frmMain
    Key = mKey
    DefaultRange = mvarDefaultRange
    DefaultType = mvarDefaultType
    OpenPictureBox = mOK
End Function

Private Sub cbo_Click()
    If mblnLoad = False Then
        Call LoadPictureList(Val(Mid(cboRang.Text, 1, 1)), Val(Mid(cbo.Text, 1, 1)))
    End If
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboRang_Click()
    If mblnFirst Then Exit Sub
    
    mblnLoad = True
    cbo.Clear
    If cboRang.ListIndex >= 0 Then
    Select Case cboRang.ItemData(cboRang.ListIndex)
    Case 0, 1, 2, 9
        cbo.AddItem "0-����ͼƬ"
        cbo.ItemData(cbo.NewIndex) = 0
        cbo.AddItem "2-FLASH����"
        cbo.ItemData(cbo.NewIndex) = 2
    Case 3
        cbo.AddItem "1-ͼ��"
        cbo.ItemData(cbo.NewIndex) = 1
    Case 4
        cbo.AddItem "0-����ͼƬ"
        cbo.ItemData(cbo.NewIndex) = 0
    End Select
    End If
    
    If cbo.ListCount > 0 Then cbo.ListIndex = 0

    mblnLoad = False
    
    Call LoadPictureList(Val(Mid(cboRang.Text, 1, 1)), Val(Mid(cbo.Text, 1, 1)))
    
End Sub

Private Sub cboRang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdCancel_Click()
    mvarDefaultType = cbo.Text
    mvarDefaultRange = cboRang.Text
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If frmParent Is frmAdvice Then
        Call frmAdvice.AddPicture(Val(Mid(lvw.SelectedItem.Key, 2)))
    Else
        mKey = Val(Mid(lvw.SelectedItem.Key, 2))
        mOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    'DoEvents
    '���潫��ʼ��һЩ����
    cboRang.Clear
    If InStr(mType, ";0;") > 0 Then
        cboRang.AddItem "0-ҽԺ��־ͼ��"
        cboRang.ItemData(cboRang.NewIndex) = 0
    End If
    If InStr(mType, ";1;") > 0 Then
        cboRang.AddItem "1-ҽԺ����ͼƬ"
        cboRang.ItemData(cboRang.NewIndex) = 1
    End If
    If InStr(mType, ";2;") > 0 Then
        cboRang.AddItem "2-��淢��ͼƬ"
        cboRang.ItemData(cboRang.NewIndex) = 2
    End If
    If InStr(mType, ";3;") > 0 Then
        cboRang.AddItem "3-��Ŀ����ͼ��"
        cboRang.ItemData(cboRang.NewIndex) = 3
    End If
    If InStr(mType, ";4;") > 0 Then
        cboRang.AddItem "4-ҳ�汳��ͼƬ"
        cboRang.ItemData(cboRang.NewIndex) = 4
    End If
    If InStr(mType, ";9;") > 0 Then
        cboRang.AddItem "9-����ͼƬ"
        cboRang.ItemData(cboRang.NewIndex) = 9
    End If
    If cboRang.ListCount > 0 Then
        If mvarDefaultRange <> "" Then
            On Error Resume Next
            cboRang.Text = mvarDefaultRange
            On Error GoTo 0
        Else
            cboRang.ListIndex = FindCboIndex(cboRang, Val(Split(mType, ";")(1)))
            If cboRang.ListIndex = -1 Then cboRang.ListIndex = 0
        End If
    End If
    mblnFirst = False
    Call cboRang_Click
    On Error Resume Next
    If mvarDefaultType <> "" Then cbo.Text = mvarDefaultType
End Sub

Private Sub Form_Load()
    mblnFirst = True
    RestoreWinState Me, App.ProductName
    mOK = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvw_DblClick()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    Call cmdOK_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    picBack.Tag = Mid(Item.Key, 2)
    txt.Text = Item.Text
    Call ShowPicture
End Sub

Private Sub ShowPicture()
    gstrSQL = "select ���,���,�߶�,���� from ��ѯͼƬԪ�� where ���=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(picBack.Tag))
    If gRs.BOF = False Then
        Call picBack.ShowPictureByFieldNew(gRs!���, gRs!��� * Screen.TwipsPerPixelX, gRs!�߶� * Screen.TwipsPerPixelY, IIf(IsNull(gRs!����), 0, gRs!����))
    End If
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
   
    Select Case Button.Key
    Case "ɾ��"
        Call DeletePicture
        If Not (lvw.SelectedItem Is Nothing) Then Call lvw_ItemClick(lvw.SelectedItem)
    Case "����"
        
        Select Case cbo.Text
        Case "0-����ͼƬ"
            dlg.DialogTitle = "��ѡ��Ҫ��ӵ�ͼƬ"
            dlg.Filter = "����ͼƬ�ļ�|*.bmp;*.jpg;*.gif"
        Case "1-ͼ��"
            dlg.DialogTitle = "��ѡ��Ҫ��ӵ�ͼ��"
            dlg.Filter = "����ͼ���ļ�|*.ico"
        Case "2-FLASH����"
            dlg.DialogTitle = "��ѡ��Ҫ��ӵ�FLASH"
            dlg.Filter = "FLASH�ļ�|*.swf"
        End Select
                        
        On Error Resume Next
        dlg.flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        dlg.FileName = ""
        dlg.MaxFileSize = 32767
        dlg.CancelError = True
        dlg.ShowOpen
        If Err.Number = 0 Then
            On Error GoTo 0
            If cbo.Text = "2-FLASH����" Then
                If SaveFlash(dlg.FileName, cboRang.ItemData(cboRang.ListIndex), cbo.ItemData(cbo.ListIndex)) Then Call LoadPictureList(Val(Mid(cboRang.Text, 1, 1)), Val(Mid(cbo.Text, 1, 1)))
            Else
                If SavePicture(dlg.FileName, imgTmp, cboRang.ItemData(cboRang.ListIndex), cbo.ItemData(cbo.ListIndex)) Then Call LoadPictureList(Val(Mid(cboRang.Text, 1, 1)), Val(Mid(cbo.Text, 1, 1)))
            End If
        Else
            Err.Clear
        End If
    Case "�б�"
        If lvw.View = 3 Then
            lvw.View = 0
        Else
            lvw.View = lvw.View + 1
        End If
    Case "����"
        ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
    End Select
End Sub

Private Sub LoadPictureList(ByVal Range As Byte, ByVal Filter As Byte)
    Dim Itmx As ListItem
            
    picBack.Tag = ""
    picBack.Cls
    
    lvw.ListItems.Clear
    txt.Text = ""
        
    gstrSQL = "select ���,����,���,�߶�,�޸����� from ��ѯͼƬԪ�� where ����=[1] and ����=[2] "
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Range, Filter)
    If gRs.BOF = False Then
        While Not gRs.EOF
            Set Itmx = lvw.ListItems.Add(, "K" & gRs!���, IIf(IsNull(gRs!����), "", gRs!����), "pic", "pic")
            Itmx.SubItems(1) = IIf(IsNull(gRs!���), "", gRs!���)
            Itmx.SubItems(2) = IIf(IsNull(gRs!�߶�), "", gRs!�߶�)
            Itmx.SubItems(3) = IIf(IsNull(gRs!�޸�����), "", gRs!�޸�����)
            gRs.MoveNext
        Wend
    End If
    tbrThis.Buttons("ɾ��").Enabled = True
    If lvw.SelectedItem Is Nothing Then
        tbrThis.Buttons("ɾ��").Enabled = False
        Exit Sub
    End If
    lvw.ListItems(1).Selected = True
    If Not (lvw.SelectedItem Is Nothing) Then Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub DeletePicture()
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo errHand
    
    gstrSQL = "zl_��ѯͼƬԪ��_delete(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    lvw.ListItems.Remove lvw.SelectedItem.Index
        
    Exit Sub
errHand:
    
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "List0"
        lvw.View = 0
    Case "List1"
        lvw.View = 1
    Case "List2"
        lvw.View = 2
    Case "List3"
        lvw.View = 3
    End Select
End Sub

Private Sub txt_GotFocus()
    SelAll txt
    zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim intLen As Long
    
    
    For i = 1 To lvw.ListItems.Count
        intLen = Len(txt.Text)
        If Mid(lvw.ListItems(i).Text, 1, intLen) = txt.Text Then
            lvw.ListItems(i).Selected = True
            Exit Sub
        End If
    Next
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme
End Sub
