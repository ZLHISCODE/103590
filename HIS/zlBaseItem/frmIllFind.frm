VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIllFind 
   Caption         =   "��������"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   Icon            =   "frmIllFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   5730
      TabIndex        =   14
      ToolTipText     =   "��ݼ�: F3"
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "��λ(&L)"
      Height          =   350
      Left            =   5730
      TabIndex        =   15
      Top             =   570
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   5730
      TabIndex        =   23
      Top             =   2395
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "��������"
      Height          =   2745
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   5445
      Begin VB.ComboBox cbo���� 
         Height          =   300
         ItemData        =   "frmIllFind.frx":020A
         Left            =   1140
         List            =   "frmIllFind.frx":020C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1215
         Width           =   4035
      End
      Begin VB.ComboBox cmb�Ա� 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2025
         Width           =   1455
      End
      Begin VB.ComboBox cmb��Ч 
         Height          =   300
         Left            =   3870
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2025
         Width           =   1305
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   3
         Top             =   765
         Width           =   1425
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   1
         Top             =   270
         Width           =   1425
      End
      Begin VB.OptionButton optMode 
         Caption         =   "������������(&C)"
         Height          =   195
         Index           =   2
         Left            =   3540
         TabIndex        =   20
         Top             =   930
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optMode 
         Caption         =   "���������ݿ�ͷ(&B)"
         Height          =   180
         Index           =   1
         Left            =   3540
         TabIndex        =   19
         Top             =   600
         Width           =   1845
      End
      Begin VB.OptionButton optMode 
         Caption         =   "��ȫ��ͬ(&A)"
         Height          =   180
         Index           =   0
         Left            =   3540
         TabIndex        =   18
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1140
         TabIndex        =   7
         Top             =   1620
         Width           =   4035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "������Ϣ(&M)"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   2460
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������Ч(&G)"
         Height          =   180
         Index           =   7
         Left            =   2820
         TabIndex        =   10
         Top             =   2085
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�����(&S)"
         Height          =   180
         Index           =   6
         Left            =   150
         TabIndex        =   8
         Top             =   2085
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&E)"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   810
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&D)"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&O)"
         Height          =   180
         Index           =   3
         Left            =   480
         TabIndex        =   6
         Top             =   1680
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   5730
      TabIndex        =   16
      Top             =   1020
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   6210
      Top             =   2340
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
            Picture         =   "frmIllFind.frx":020E
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   5520
      Top             =   2280
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
            Picture         =   "frmIllFind.frx":0662
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2865
      Left            =   60
      TabIndex        =   21
      Top             =   2850
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�Ա�����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "������Ч"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "������Ϣ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "˵��"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   5925
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   635
      SimpleText      =   "CoolBar1"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmIllFind.frx":0AB6
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7355
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
   Begin VB.Menu mnuShort 
      Caption         =   "��ݲ˵�(&S)"
      Visible         =   0   'False
      Begin VB.Menu mnuShortLocate 
         Caption         =   "��λ(&P)"
      End
      Begin VB.Menu mnuShortSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ͼ��(&B)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmIllFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr������� As String
Dim mblnChange As Boolean
Dim mintColumn As Long
Dim mblnShowStop As Boolean         '�Ƿ���ʾͣ����Ŀ
Private Sub cbo����_Click()
    If cbo����.Tag <> cbo����.Text Then
        mblnChange = True
        cbo����.Tag = cbo����.Text
    End If
End Sub

Private Sub cmb�Ա�_Click()
    mblnChange = True
End Sub

Private Sub cmb��Ч_Click()
    mblnChange = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    Dim i As Long
    
    If mblnChange = False Then
        Exit Sub
    End If
    mblnChange = False
    
    On Error GoTo errHandle
    gstrSQL = ""
    If txtEdit(0).Text <> "" Then
        strTemp = Replace(Trim(txtEdit(0).Text), "'", "''")
        gstrSQL = IIF(optMode(0).value = True, "  ����=[1] and ", _
                IIF(optMode(1).value = True, " ���� like [2] and ", " ���� like [3] and "))
    End If
    If txtEdit(1).Text <> "" Then
        strTemp = Replace(Trim(txtEdit(1).Text), "'", "''")
        gstrSQL = gstrSQL & IIF(optMode(0).value = True, " ����=[1] and ", _
                IIF(optMode(1).value = True, " ���� like [2] and ", " ���� like [3] and "))
    End If
    If txtEdit(2).Text <> "" Then
        strTemp = Replace(Trim(txtEdit(2).Text), "'", "''")
        gstrSQL = gstrSQL & IIF(optMode(0).value = True, " ����=[1] and ", _
                IIF(optMode(1).value = True, " ���� like [2] and ", " ���� like [3] and "))
    End If
    If txtEdit(3).Text <> "" Then
        strTemp = UCase(Replace(Trim(txtEdit(3).Text), "'", "''"))
        gstrSQL = gstrSQL & IIF(optMode(0).value = True, " (����=[1] or �����=[1]) and ", _
                IIF(optMode(1).value = True, " (���� like [2] or ����� like [2]) and ", " (���� like [3] or ����� like [3]) and "))
    End If
    If cmb��Ч.Text <> "" Then
        gstrSQL = gstrSQL & " ��Ч����=[4] and "
    End If
    If cmb�Ա�.Text <> "" Then
        gstrSQL = gstrSQL & " �Ա�����=[5] and "
    End If
    
    If cbo����.Text <> "" Then
        gstrSQL = gstrSQL & IIF(cbo����.ListIndex = 1, "����='1'", "(���� Is Null or ����='0')") & " AND"
    End If
    
    If gstrSQL = "" Then
        MsgBox "�������ѯ������", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrSQL = "select ID,����,����,����,����,˵��,�Ա�����,��Ч���� as ������Ч,decode(����,1,'¼��') ������Ϣ,����ID " & _
              " From ��������Ŀ¼  Where " & gstrSQL & " ���=[6] " & _
              IIF(mblnShowStop = False, " and (����ʱ�� is null or ����ʱ��>=to_date('3000-01-01','yyyy-mm-dd'))", "")
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp, strTemp & "%", "%" & strTemp & "%", cmb��Ч.Text, cmb�Ա�.Text, mstr�������)
        
    With lvwMain.ListItems
        .Clear
        Do Until rsTemp.EOF
            '�ó���ȷ��ͼ��
            '��ӽڵ�
            Set lst = .Add(, "K" & rsTemp("id"), rsTemp("����"), "Item", "Item")
            
            Dim varValue As Variant
            '����ListView�����������ݿ�ȡ��
            For i = 2 To lvwMain.ColumnHeaders.Count
                varValue = rsTemp(lvwMain.ColumnHeaders(i).Text).value
                lst.SubItems(i - 1) = IIF(IsNull(varValue), "", varValue)
            Next
            lst.Tag = rsTemp("����ID")
            rsTemp.MoveNext
        Loop
    End With
    If rsTemp.RecordCount = 0 Then
        stbThis.Panels(2).Text = "�Բ���û�ҵ�����Ҫ�ļ�����������������ԡ�"
        txtEdit(0).SetFocus
    Else
        lvwMain.ListItems(1).Selected = True
        stbThis.Panels(2).Text = "���ҵ�" & rsTemp.RecordCount & "��������¼��"
        lvwMain.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, 3
End Sub

Private Sub cmdLocate_Click()
    Dim nod As Node
    Dim lst As ListItem
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
            
    If frmIllManage.tvwMain_S.Visible = True Then
        '���ȶ�λ����
        Set nod = frmIllManage.tvwMain_S.Nodes("K" & lvwMain.SelectedItem.Tag)
        If Not nod Is frmIllManage.tvwMain_S.SelectedItem Then
            nod.Selected = True
            nod.EnsureVisible
            Call frmIllManage.FillList
        End If
    End If
    Set lst = frmIllManage.lvwMain.ListItems(lvwMain.SelectedItem.Key)
    lst.Selected = True
    lst.EnsureVisible
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then Call cmdFind_Click
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
    If Not lvwMain.SelectedItem Is Nothing Then
        lvwMain.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim sngWidth As Single, sngHeight As Single
    
    If WindowState = 1 Then Exit Sub
    
    sngWidth = IIF(ScaleWidth < 7000, 7000, ScaleWidth)
    sngHeight = IIF(ScaleHeight < 3525, 3525, ScaleHeight)
    On Error Resume Next
    
    cmdFind.Left = sngWidth - cmdFind.Width - 200
    cmdLocate.Left = cmdFind.Left
    cmdClose.Left = cmdFind.Left
    cmdHelp.Left = cmdFind.Left
    
    lvwMain.Width = sngWidth - lvwMain.Left - 90
    lvwMain.Height = sngHeight - stbThis.Height - lvwMain.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Public Sub ShowFind(ByVal str������� As String, ByVal blnShowStop As Boolean)
    mblnChange = False
    mstr������� = str�������
    mblnShowStop = blnShowStop
        
    cmb�Ա�.AddItem ""
    cmb�Ա�.AddItem "��"
    cmb�Ա�.AddItem "Ů"
    
    cmb��Ч.AddItem ""
    cmb��Ч.AddItem "����"
    cmb��Ч.AddItem "��ת"
    cmb��Ч.AddItem "����"
    
    cbo����.AddItem ""
    cbo����.AddItem "¼�������Ϣ"
    cbo����.AddItem "��¼�������Ϣ"
    
    frmIllFind.Show vbModal, frmIllManage
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub lvwMain_DblClick()
    Call cmdLocate_Click
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdLocate_Click
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    If Button = 2 Then
        For i = 0 To 3
            mnuShortIcon(i).Checked = False
        Next
        mnuShortIcon(lvwMain.View).Checked = True
        
        PopupMenu mnuShort
    End If
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    lvwMain.View = Index
End Sub

Private Sub mnuShortLocate_Click()
    Call cmdLocate_Click
End Sub

Private Sub optMode_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 2 Then
        zlcommfun.OpenIme True
    Else
        zlcommfun.OpenIme False
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
'���ڶ����ı�����ò�Ҫ�ӿո�
    If Index = 0 Or Index = 1 Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
        'ֻ��ȡ��Щ��ĸ
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/*-+. " & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub



