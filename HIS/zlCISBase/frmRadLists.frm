VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmRadLists 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ӱ������Ŀ"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   8010
   Icon            =   "frmRadLists.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Visible         =   0   'False
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8010
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinWidth1       =   24000
      MinHeight1      =   720
      Width1          =   8730
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   24000
         _ExtentX        =   42333
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
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ����ǰ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ��ǰ��"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "���ļ�"
               Object.Tag             =   "����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Mod"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸��ļ�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Del"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ���ļ�"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6990
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRadLists.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9049
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSComctlLib.ImageList imgKind 
      Left            =   2220
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":115C
            Key             =   "kind"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":16F6
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLine 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   2040
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5895
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   960
      Width           =   30
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7080
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":1C90
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":1EAA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":20C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":22DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":24F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":2712
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":292C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":2B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":2D60
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":2F7A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":319A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6315
      Top             =   435
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":33BA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":35DA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":37FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":3A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":3C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":3E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":4062
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":427C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":4496
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":46B0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadLists.frx":48D0
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwKind 
      Height          =   5625
      Left            =   15
      TabIndex        =   0
      Top             =   945
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   9922
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgKind"
      SmallIcons      =   "imgKind"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   5385
      Left            =   2130
      TabIndex        =   1
      Top             =   930
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   9499
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgKind"
      SmallIcons      =   "imgKind"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&U)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "Ԥ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine0 
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
      Begin VB.Menu mnuEditMod 
         Caption         =   "�޸�(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&Q)"
      Begin VB.Menu mnuViewTools 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolsButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolsText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&E)..."
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
Attribute VB_Name = "frmRadLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim intCount As Integer       '�������ɼ�����

Private Sub Form_Activate()
    If Me.lvwKind.ListItems.count = 0 Then
        MsgBox "Ӱ����������ݶ�ʧ��(��ϵ����Ա)", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    '����ָ�
    Me.lvwItem.ColumnHeaders.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_����", "����", 2500
        .Add , "_����", "����", 1000
        .Add , "_��λ", "��λ", 900
        .Add , "_��λ", "��λ", 600
        .Add , "_���в���", "���в���", 1000
        .Add , "_�ɷ���Ƭ", "�ɷ���Ƭ", 1000
        .Add , "_����ͼ��", "����ͼ��", 900
        .Add , "_���׼��", "���׼��", 2000
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_����").Index - 1: .SortOrder = lvwAscending
    End With
    
    Call RestoreWinState(Me, App.ProductName)
    Me.lvwKind.View = lvwReport
    Me.lvwItem.ColumnHeaders("_����").Position = 1
    
    'Ȩ�޿���
    If InStr(1, gstrPrivs, "��ɾ��") = 0 Then
        Me.mnuEdit.Enabled = False
        Me.mnuEditAdd.Enabled = False
        Me.mnuEditMod.Enabled = False
        Me.mnuEditDel.Enabled = False
        Me.tlbThis.Buttons("Add").Enabled = False
        Me.tlbThis.Buttons("Mod").Enabled = False
        Me.tlbThis.Buttons("Del").Enabled = False
    End If
    
    'װ������
    gstrSql = "Select * From Ӱ������� Order By ����"
    Err = 0: On Error GoTo ErrHand
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    With rsTemp
        Me.lvwKind.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwKind.ListItems.Add(, "_" & !����, !����, "kind", "kind")
            objItem.SubItems(1) = !����
            .MoveNext
        Loop
    End With
    Err = 0: On Error GoTo 0
    If Me.lvwKind.ListItems.count > 0 Then
        Me.lvwKind.ListItems(1).Selected = True
        Me.lvwKind.SelectedItem.EnsureVisible
        Call zlRefItems
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    '-------------------------------------------------
    '���ݴ���仯����������������λ��
    '-------------------------------------------------
    Dim lngHeightTools As Long, lngHeightState As Long
    lngHeightTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngHeightState = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    Me.picLine.Top = 0
    Me.picLine.Height = Me.ScaleHeight
    On Error Resume Next
    If Me.picLine.Left < 1000 Then Me.picLine.Left = 1000
    If Me.picLine.Left > Me.ScaleWidth - 2600 Then Me.picLine.Left = Me.ScaleWidth - 2600
    
    With Me.lvwKind
        .Left = Me.ScaleLeft
        .Width = Me.picLine.Left - .Left
        .Top = Me.ScaleTop + lngHeightTools
        .Height = Me.ScaleHeight - .Top - lngHeightState
    End With
    
    With Me.lvwItem
        .Left = Me.picLine.Left + Me.picLine.Width
        .Width = Me.ScaleWidth - .Left
        .Top = Me.ScaleTop + lngHeightTools
        .Height = Me.ScaleHeight - .Top - lngHeightState
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwItem
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwItem_DblClick()
    If Me.mnuEditMod.Enabled Then Call mnuEditMod_Click
End Sub

Private Sub lvwItem_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.mnuEditMod.Enabled Then Call mnuEditMod_Click
End Sub

Private Sub lvwItem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Me.mnuEdit.Enabled Then PopupMenu Me.mnuEdit, 2
End Sub

Private Sub lvwKind_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call zlRefItems
End Sub

Private Sub mnuEditAdd_Click()
    frmRadNew.Show 1, Me
End Sub

Private Sub mnuEditDel_Click()
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("��Ľ���" & Me.lvwItem.SelectedItem.Text & "����Ӱ������Ŀ��ɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSql = "zl_Ӱ������Ŀ_Delete(" & Mid(Me.lvwItem.SelectedItem.Key, 2) & ")"
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Call Me.lvwItem.ListItems.Remove(Me.lvwItem.SelectedItem.Key)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub mnuEditMod_Click()
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    With frmRadMod
        .lblBaseInfo.Tag = Mid(Me.lvwItem.SelectedItem.Key, 2)
        .Show 1, Me
    End With
    Call zlRefItems(Mid(Me.lvwItem.SelectedItem.Key, 2))
End Sub

Private Sub mnuFileExcel_Click()
    Call RptPrint(3)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    Call RptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call RptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.lvwItem.SelectedItem Is Nothing Then
        Call zlRefItems
    Else
        Call zlRefItems(Mid(Me.lvwItem.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolsButton_Click()
    Me.mnuViewToolsButton.Checked = Not Me.mnuViewToolsButton.Checked
    Me.clbThis.Visible = Me.mnuViewToolsButton.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolsText_Click()
    Dim i As Integer
    Me.mnuViewToolsText.Checked = Not Me.mnuViewToolsText.Checked
    If Me.mnuViewToolsText.Checked Then
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

Private Sub picLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picLine.Left = Me.picLine.Left + x
    End If
End Sub

Private Sub picLine_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_Resize
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case UCase("Preview")
        Call mnuFilePreview_Click
    Case UCase("Print")
        Call mnuFilePrint_Click
    Case UCase("Add")
        Call mnuEditAdd_Click
    Case UCase("Mod")
        Call mnuEditMod_Click
    Case UCase("Del")
        Call mnuEditDel_Click
    Case UCase("Help")
        Call mnuHelpHelp_Click
    Case UCase("Exit")
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub RptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    Dim bytR As Byte
    On Error Resume Next
    
    Set objPrint.Body.objData = Me.lvwItem
    objPrint.Title.Text = Me.lvwKind.SelectedItem.Text & "�����Ŀ"
    objPrint.UnderAppItems.Add ""
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Now
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objPrint)
        If bytR <> 0 Then zlPrintOrViewLvw objPrint, bytR
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub zlRefItems(Optional lngItemId As Long)
    '-------------------------------------------------
    '����:ˢ�µ�ǰ����Ŀ�б�
    '-------------------------------------------------
    If Me.lvwKind.SelectedItem Is Nothing Then Exit Sub
    
    gstrSql = "Select I.ID,I.����, I.����,I.�걾��λ, I.���㵥λ,R.���в���,R.�ɷ���Ƭ,R.����ͼ��,R.���׼��" & _
            "  From ������ĿĿ¼ I, Ӱ������Ŀ R" & _
            " Where I.ID = R.������Ŀid And R.Ӱ�����=[1] "
    
    Err = 0: On Error GoTo ErrHand
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(Me.lvwKind.SelectedItem.Key, 2))
        
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����, "item", "item")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_��λ").Index - 1) = IIf(IsNull(!�걾��λ), "", !�걾��λ)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_��λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            Select Case !���в���
            Case 1
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_���в���").Index - 1) = "1-����"
            Case 2
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_���в���").Index - 1) = "2-ѡ�����"
            Case Else
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_���в���").Index - 1) = "0-������"
            End Select
            Select Case !�ɷ���Ƭ
            Case 1
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_�ɷ���Ƭ").Index - 1) = "1-����"
            Case 2
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_�ɷ���Ƭ").Index - 1) = "2-ѡ�񷢷�"
            Case Else
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_�ɷ���Ƭ").Index - 1) = "0-������"
            End Select
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����ͼ��").Index - 1) = IIf(IsNull(!����ͼ��), "", !����ͼ��)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_���׼��").Index - 1) = IIf(IsNull(!���׼��), "", !���׼��)
            .MoveNext
        Loop
    End With
    If Me.lvwItem.ListItems.count > 0 Then
        Err = 0: On Error Resume Next
        Me.lvwItem.ListItems("_" & lngItemId).Selected = True
        If Me.lvwItem.SelectedItem Is Nothing Then Me.lvwItem.ListItems(1).Selected = True
        Me.lvwItem.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "�����" & Me.lvwItem.ListItems.count & "����Ŀ"
    Else
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

