VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.OCX"
Begin VB.Form frm���������������� 
   Caption         =   "�������������������"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   Icon            =   "frm����������������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7665
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPopu 
      Height          =   810
      Left            =   2850
      ScaleHeight     =   750
      ScaleWidth      =   2040
      TabIndex        =   32
      Top             =   5340
      Visible         =   0   'False
      Width           =   2100
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "����δ����Ӧ��(A)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   31
         Top             =   420
         Width           =   1725
      End
      Begin VB.Label lblMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "ָ����Ӧ��(D)"
         ForeColor       =   &H8000000E&
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   30
         Top             =   135
         Width           =   1725
      End
      Begin VB.Label lblBackColor 
         BackColor       =   &H8000000D&
         Height          =   285
         Left            =   75
         TabIndex        =   33
         Top             =   90
         Width           =   1890
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBuilded 
      Height          =   3645
      Left            =   945
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   6429
      _Version        =   393216
      Rows            =   20
      Cols            =   3
      FixedCols       =   0
      HighLight       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.TextBox txt��Ӧ�� 
      Height          =   300
      Left            =   4845
      TabIndex        =   17
      Top             =   795
      Width           =   1755
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   300
      Left            =   6945
      Picture         =   "frm����������������.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   795
      Width           =   315
   End
   Begin VB.CommandButton cmdDele 
      Enabled         =   0   'False
      Height          =   300
      Left            =   7260
      Picture         =   "frm����������������.frx":06D4
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   795
      Width           =   315
   End
   Begin MSComctlLib.ListView lvw��Ӧ�� 
      Height          =   3240
      Left            =   4005
      TabIndex        =   20
      Top             =   1140
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5715
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "��Ӧ��"
         Object.Tag             =   "��Ӧ��"
         Text            =   "��Ӧ��"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "��������"
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "�ʺ�"
         Text            =   "�ʺ�"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "��ϵ��"
         Object.Tag             =   "��ϵ��"
         Text            =   "��ϵ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fra 
      Caption         =   "�������"
      Height          =   1395
      Left            =   60
      TabIndex        =   11
      Top             =   2940
      Width           =   3870
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   300
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   330
         Width           =   2640
      End
      Begin VB.TextBox txt˵�� 
         Height          =   300
         Left            =   165
         MaxLength       =   50
         TabIndex        =   15
         Top             =   915
         Width           =   3570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����˵��(&F)"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   14
         Top             =   690
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���㷽ʽ(&J)"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   12
         Top             =   390
         Width           =   990
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "��������"
      Height          =   1845
      Left            =   60
      TabIndex        =   1
      Top             =   945
      Width           =   3840
      Begin VB.CheckBox chkType 
         Caption         =   "����(&Q)"
         Height          =   225
         Index           =   4
         Left            =   2730
         TabIndex        =   10
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chkType 
         Caption         =   "��������(&L)"
         Height          =   225
         Index           =   3
         Left            =   225
         TabIndex        =   9
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkType 
         Caption         =   "�豸(&S)"
         Height          =   225
         Index           =   2
         Left            =   2730
         TabIndex        =   8
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chkType 
         Caption         =   "����(&W)"
         Height          =   225
         Index           =   1
         Left            =   1545
         TabIndex        =   7
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chkType 
         Caption         =   "ҩƷ(&Y)"
         Height          =   225
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Index           =   0
         Left            =   1050
         TabIndex        =   3
         Top             =   240
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   73531395
         CurrentDate     =   38936.9668171296
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Index           =   1
         Left            =   1035
         TabIndex        =   5
         Top             =   660
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   73531395
         CurrentDate     =   38936.4668171296
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��(&E)"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��(&K)"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   315
         Width           =   990
      End
   End
   Begin VB.Frame fraTemp 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   24
      Top             =   705
      Width           =   7995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6510
      TabIndex        =   22
      Top             =   4620
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5340
      TabIndex        =   21
      Top             =   4620
      Width           =   1100
   End
   Begin VB.Frame fraTemp 
      Height          =   30
      Index           =   0
      Left            =   -30
      TabIndex        =   23
      Top             =   4425
      Width           =   7905
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   -75
      Top             =   4560
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
            Picture         =   "frm����������������.frx":0C5E
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   645
      Top             =   4560
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
            Picture         =   "frm����������������.frx":10B6
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPross 
      Height          =   615
      Left            =   -15
      ScaleHeight     =   555
      ScaleWidth      =   7620
      TabIndex        =   25
      Top             =   4455
      Visible         =   0   'False
      Width           =   7680
      Begin MSComctlLib.ProgressBar prgb 
         Height          =   285
         Left            =   45
         TabIndex        =   26
         Top             =   255
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblPer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   180
         Left            =   7215
         TabIndex        =   28
         Top             =   45
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lbl��Ӧ�� 
         Caption         =   "�������ɣ�"
         Height          =   195
         Left            =   75
         TabIndex        =   27
         Top             =   30
         Width           =   2040
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ��(&G)"
      Height          =   180
      Index           =   2
      Left            =   4035
      TabIndex        =   16
      Top             =   870
      Width           =   810
   End
   Begin VB.Label lblInfor 
      Caption         =   "���ò�������Ӧ�̵������������������й�Ӧ�̿���ͨ������ȷ�ʽ����¼�롣"
      Height          =   285
      Left            =   735
      TabIndex        =   0
      Top             =   375
      Width           =   6360
   End
   Begin VB.Image img���� 
      Height          =   480
      Left            =   120
      Picture         =   "frm����������������.frx":150E
      Top             =   180
      Width           =   480
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuLocal 
         Caption         =   "ָ����Ӧ��(&L)"
      End
      Begin VB.Menu mnuPopuAll 
         Caption         =   "����δ����Ӧ��(&A)"
      End
   End
End
Attribute VB_Name = "frm����������������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnCancel As Boolean
Private mstr��Ӧ��Ȩ�� As String
Private mstrPrivs As String
Private mfrmMain As Form
Private mblnFirst As Boolean
Private mintColumn As Integer
Public Function ShowCard(ByVal FrmMain As Form, ByVal strPrivs As String) As Boolean
    '------------------------------------------------------------------------------------------------
    '����:��ʾ�������ӵ���������
    '����:frmMain-������
    '     strPrivs-Ȩ�޴�
    '����:���ɳɹ�������True,����False
    '------------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs
    mblnCancel = True
    mstr��Ӧ��Ȩ�� = ""
    Set mfrmMain = FrmMain
    Me.Show 1, mfrmMain
    ShowCard = Not mblnCancel
End Function

Private Sub cbo���㷽ʽ_Click()
    Call SetCtrlEn
End Sub

Private Sub cbo���㷽ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkType_Click(Index As Integer)
    Call SetCtrlEn
End Sub

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 

Private Sub cmdAdd_Click()
    picPopu.Left = cmdAdd.Left + cmdAdd.Width - picPopu.Width
    picPopu.Top = cmdAdd.Top + cmdAdd.Height
    picPopu.Visible = True
    picPopu.SetFocus
    RaisEffect picPopu, 2
End Sub

Private Sub cmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'PopupMenu Me.mnuPopu, vbPopupMenuRightAlign
End Sub
Private Sub cmdCancel_Click()
'    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdDele_Click()
    Call Dele��Ӧ��
    Call SetCtrlEn
End Sub

Private Sub cmdOK_Click()
    Dim lvwItem  As ListItem
    Dim intByte As Integer
    mblnCancel = False
    If zlCommFun.ActualLen(txt˵��.Text) > 50 Then
        ShowMsgbox "˵�����ܴ���25�����ֻ�50���ַ�!"
        If txt˵��.Enabled Then txt˵��.SetFocus
        Exit Sub
    End If
    If InStr(1, txt˵��.Text, "'") > 0 Then
        ShowMsgbox "˵���в��ܰ��������ַ�(������)!"
        If txt˵��.Enabled Then txt˵��.SetFocus
        Exit Sub
    End If
    If lvw��Ӧ��.ListItems.Count = 0 Then Exit Sub
    Screen.MousePointer = 11
    Me.Enabled = False
    picPross.Visible = True
    picPross.ZOrder 1
    prgb.Max = lvw��Ӧ��.ListItems.Count
    prgb.Min = 0
    prgb.Value = 0
    cmdOk.Visible = False
    cmdCancel.Visible = False
    
    Call initGrid
    
    For Each lvwItem In Me.lvw��Ӧ��.ListItems
        
        Call BuildingData(Val(Mid(lvwItem.Key, 2)), lvwItem.Text)
        prgb.Value = prgb.Value + 1
        lblPer.Caption = Round(prgb.Value / prgb.Max * 100, 2) & "/100"
        DoEvents
    Next
    picPross.Visible = False
    Screen.MousePointer = 0
    'cmdOk.Visible = True
    cmdCancel.Visible = True
    cmdCancel.Caption = "�˳�(&C)"
    Call ShowBuildedGrid
    
    Me.Enabled = True
    
End Sub


Private Sub dtpDate_Change(Index As Integer)
    Err = 0: On Error Resume Next
    If Index = 0 Then
        If dtpDate(0).Value >= dtpDate(1).Value Then
            dtpDate(1).Value = Format(dtpDate(0).Value, "yyyy-mm-dd") & " 23:59:59"
        End If
    Else
        If dtpDate(0).Value > dtpDate(1).Value Then
            dtpDate(0).Value = Format(dtpDate(1).Value, "yyyy-mm-dd") & " 00:00:00"
        End If
    End If
End Sub

Private Sub dtpDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If Load���㷽ʽ = False Then Unload Me: Exit Sub
    Call Ȩ�޿���
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mstr��Ӧ��Ȩ�� = " (ĩ��=1 and " & Get����Ȩ��(mstrPrivs) & ") "
    dtpDate(1).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59"
    dtpDate(1).MaxDate = dtpDate(1).Value
    dtpDate(0).Value = Format(DateAdd("d", -7, dtpDate(1).Value), "yyyy-mm-dd") & " 00:00:00"
    dtpDate(0).MaxDate = dtpDate(1).Value
    '�ָ���ز���
    RestoreWinState Me, App.ProductName
End Sub

 

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 7785 Then Me.Width = 7785
    If Me.Height < 5475 Then Me.Height = 5475
    
    With picPross
        .Top = ScaleHeight - .Height
        .Width = ScaleWidth - .Left
    End With
    With cmdCancel
        .Top = picPross.Top + (picPross.Height - .Height) / 2
        .Left = Me.ScaleWidth - .Width - 100
        cmdOk.Top = .Top
        cmdOk.Left = .Left - cmdOk.Width - 50
    End With
    With fraTemp(0)
        .Top = picPross.Top
        .Width = Me.ScaleWidth + 100
    End With
    With lvw��Ӧ��
        .Width = ScaleWidth - .Left - 50
        .Height = fraTemp(0).Top - .Top - 50
    End With
    With cmdDele
        .Left = lvw��Ӧ��.Left + lvw��Ӧ��.Width - .Width
        cmdAdd.Left = .Left - cmdAdd.Width
    End With
    fraTemp(1).Width = Me.ScaleWidth + 100
    With mshBuilded
        .Left = 50
        .Top = Me.txt��Ӧ��.Top
        .Height = fraTemp(0).Top - .Top - 10
        .Width = Me.ScaleWidth - .Left
    End With
    With txt��Ӧ��
        .Width = cmdAdd.Left - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lblMenu_Click(Index As Integer)
        picPopu.Visible = False
        Select Case Index
        Case 0
            Call mnuPopuLocal_Click
        Case 1
            Call mnuPopuAll_Click
        End Select
    
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    
    If Index = 0 Then
        lblMenu(Index).ForeColor = vbWhite
        lblMenu(1).ForeColor = &H80000012
    Else
        lblMenu(Index).ForeColor = vbWhite
        lblMenu(0).ForeColor = &H80000012
    End If
    With lblBackColor
        .Top = lblMenu(Index).Top - 25
    End With

End Sub

Private Sub lvw��Ӧ��_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvw��Ӧ��.Sorted = True
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvw��Ӧ��.SortOrder = IIf(lvw��Ӧ��.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvw��Ӧ��.SortKey = mintColumn
        lvw��Ӧ��.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw��Ӧ��_ItemClick(ByVal Item As MSComctlLib.ListItem)
        Call SetCtrlEn
End Sub

Private Sub lvw��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Function Getϵͳ����() As String
    '-------------------------------------------------------------------------------------------
    '����:��ȡϵͳ����
    '����:��1,2,3����ʽ����
    '-------------------------------------------------------------------------------------------
    Dim str����  As String
    str���� = ""
    '1����ҩƷӦ����   2��������Ӧ����   3�����豸Ӧ����   4��������,5--��������
    str���� = IIf(chkType(0).Value = 1, ",1", "")
    str���� = str���� & IIf(chkType(1).Value = 1, ",2", "")
    str���� = str���� & IIf(chkType(2).Value = 1, ",3", "")
    str���� = str���� & IIf(chkType(3).Value = 1, ",5", "")
    str���� = str���� & IIf(chkType(4).Value = 1, ",4", "")
    If str���� <> "" Then
        str���� = Mid(str����, 2)
    End If
    Getϵͳ���� = str����
End Function
Private Sub mnuPopuAll_Click()
        'ȫѡ�������ݵĹ�Ӧ��
    Dim rsData As New ADODB.Recordset
    Dim str���� As String
    Dim dtStartdate As Date
    Dim dtEndDate As Date
    Dim lvwItem As ListItem
    
    Err = 0: On Error GoTo errHand:
    
    str���� = Getϵͳ����
    If str���� = "" Then
        ShowMsgbox "δѡ�񱾴����ɵ�ϵͳ����!"
        Exit Sub
    End If
    
    str���� = " And A.ϵͳ��ʶ in (" & str���� & ")"
    
    If Me.lvw��Ӧ��.ListItems.Count <> 0 Then
        If MsgBox("�Ѿ�ѡ���˹�Ӧ�̣��Ƿ��������ѡ��Ĺ�Ӧ�̣�", vbQuestion + vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
            Me.lvw��Ӧ��.ListItems.Clear
        End If
    End If
    zlCommFun.ShowFlash "���ڻ�ȡ��Ӧ��,���Ժ�...."
    
    gstrSQL = "Select distinct b.id, b.����,b.���� ,b.��������,b.�ʺ�,b.��ϵ��,b.����" & _
             "   FROM Ӧ����¼ A,��Ӧ�� b" & _
             "   WHERE  a.��λid+0=b.id and a.�ƻ����� IS NULL AND a.��¼���� <> -1  AND a.������� IS NULL AND " & _
             "         a.������� BETWEEN [1] AND [2] " & str���� & IIf(mstr��Ӧ��Ȩ�� <> "", " and " & mstr��Ӧ��Ȩ��, "") & _
             "   order by b.����"

    dtStartdate = Format(Me.dtpDate(0).Value, "yyyy-MM-DD hh:mm:ss")
    dtEndDate = Format(Me.dtpDate(1).Value, "yyyy-MM-DD hh:mm:ss")
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, dtStartdate, dtEndDate)
    
    If rsData.EOF Then
        zlCommFun.StopFlash
        ShowMsgbox "��ǰ��Χ��û����Ӧ���ݷ����Ĺ�Ӧ��!"
        Exit Sub
    End If
    With rsData
        Do While Not .EOF
            Err = 0: On Error Resume Next
            Set lvwItem = Me.lvw��Ӧ��.ListItems.Add(, "K" & !ID, !���� & "-" & !����, 1, 1)
            If Err = 0 Then
                lvwItem.SubItems(1) = Nvl(!��������)
                lvwItem.SubItems(2) = Nvl(!�ʺ�)
                lvwItem.SubItems(3) = Nvl(!��ϵ��)
                lvwItem.SubItems(4) = Nvl(!����)
            Else
                Err = 0: On Error GoTo 0:
            End If
            .MoveNext
        Loop
    End With
    zlCommFun.StopFlash
    Call SetCtrlEn
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuPopuLocal_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim blnCancel As Boolean
    
    gstrSQL = "" & _
        "   Select id,�ϼ�ID, ����,����,ĩ��,����,���֤��,���֤Ч��,ִ�պ�,ִ��Ч��,˰��ǼǺ�,��ַ,��������,�ʺ�,��ϵ��,����ʱ��,����,������" & _
        "   From ��Ӧ�� " & _
        "   where (����ʱ�� is null or ����ʱ��>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & zl_��ȡվ������ & " " & _
            IIf(mstr��Ӧ��Ȩ�� <> "", " and (ĩ��<>1 or " & mstr��Ӧ��Ȩ�� & ")", "") & _
        "   start with �ϼ�id is null connect by prior id=�ϼ�id"
        
    
    'ShowSelect:
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 2, "��Ӧ��ѡ��", False, , "ѡ��Ӧ��", False, True, False, , , , blnCancel, , True)
    If blnCancel Or rsTemp Is Nothing Then Exit Sub
    If Add��Ӧ��(rsTemp) = False Then Exit Sub
    Call SetCtrlEn
End Sub

Private Sub picPopu_KeyDown(KeyCode As Integer, Shift As Integer)
      
        If KeyCode = vbKeyReturn Then
            If lblBackColor.Top < lblMenu(0).Top Then
               Call lblMenu_Click(0)
            Else
               Call lblMenu_Click(1)
            End If
        End If
        If KeyCode = vbKeyDown Then
            If lblBackColor.Top < lblMenu(0).Top Then
                lblMenu_MouseMove 1, 0, 0, 0, 0
            Else
                lblMenu_MouseMove 0, 0, 0, 0, 0
            End If
        ElseIf KeyCode = vbKeyUp Then
            If lblBackColor.Top < lblMenu(0).Top Then
                lblMenu_MouseMove 1, 0, 0, 0, 0
            Else
                lblMenu_MouseMove 0, 0, 0, 0, 0
            End If
        End If
        If KeyCode = vbKeyA And Shift = 4 Then
               Call lblMenu_Click(1)
        End If
        If KeyCode = vbKeyD And Shift = 4 Then
               Call lblMenu_Click(0)
        End If
End Sub

Private Sub picPross_Resize()
    prgb.Width = picPross.ScaleWidth - prgb.Left - 50
    lblPer.Left = prgb.Width + prgb.Left - lblPer.Width - 20
End Sub

Private Sub picPopu_LostFocus()
    picPopu.Visible = False
End Sub

Private Sub picPopu_Paint()
    RaisEffect picPopu, 2
          
End Sub

Private Sub txt��Ӧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txt��Ӧ��.Tag <> "" Then Exit Sub
    strKey = GetMatchingSting(UCase(txt��Ӧ��.Text))
    
    gstrSQL = "" & _
        "   Select id, ����,����,ĩ��,����,���֤��,���֤Ч��,ִ�պ�,ִ��Ч��,˰��ǼǺ�,��ַ,��������,�ʺ�,��ϵ��,����ʱ��,����,������" & _
        "   From ��Ӧ�� " & _
        "   where ĩ��=1 " & zl_��ȡվ������ & " and  (����ʱ�� is null or ����ʱ��>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & IIf(mstr��Ӧ��Ȩ�� <> "", " and " & mstr��Ӧ��Ȩ��, "") & _
        "          and (���� like [1] or ���� like [1] or ���� like [1])"
    'ShowSelect:
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    Dim lngX As Long, lngY As Long, lngH As Long
    lngX = Me.Left + txt��Ӧ��.Left + Screen.TwipsPerPixelX
    lngY = Me.Top + Me.Height - Me.ScaleHeight + txt��Ӧ��.Top
    lngH = txt��Ӧ��.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "��Ӧ��ѡ��", False, "", "ѡ��Ӧ��", False, True, True, lngX, lngY, lngH, blnCancel, False, True, strKey)
    
    If blnCancel Or rsTemp Is Nothing Then Exit Sub
    If rsTemp.State <> 1 Then Exit Sub
    If Add��Ӧ��(rsTemp) = False Then Exit Sub
    Call SetCtrlEn
End Sub
Private Function Add��Ӧ��(ByVal rsTemp As ADODB.Recordset) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '����:���ӹ�Ӧ��
    '------------------------------------------------------------------------------------------------------------------------
    Dim lvwItem As ListItem
    If rsTemp.EOF Then Exit Function
    
    
    Err = 0: On Error Resume Next:
    Set lvwItem = Me.lvw��Ӧ��.ListItems.Add(, "K" & rsTemp!ID, rsTemp!���� & "-" & rsTemp!����, 1, 1)
    If Err <> 0 Then
        MsgBox "��ѡ��Ĺ�Ӧ���Ѿ�����,������ѡ��", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    Err = 0: On Error GoTo errHand:
    lvwItem.SubItems(1) = Nvl(rsTemp!��������)
    lvwItem.SubItems(2) = Nvl(rsTemp!�ʺ�)
    lvwItem.SubItems(3) = Nvl(rsTemp!��ϵ��)
    lvwItem.SubItems(4) = Nvl(rsTemp!����)
    Add��Ӧ�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function Dele��Ӧ��() As Boolean
    Dim intIndex As Integer
    Err = 0: On Error Resume Next
    With lvw��Ӧ��
        '��ɾ��ListView�ж�Ӧ�ڵ�
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
End Function
Private Sub SetCtrlEn()
    '����:���ÿؼ�����
    Dim blnData As Boolean
    Dim blnSel As Boolean
    Dim blnCheck As Boolean
    blnData = Me.lvw��Ӧ��.ListItems.Count <> 0
    blnSel = Not Me.lvw��Ӧ��.SelectedItem Is Nothing
    blnCheck = Me.chkType(0).Value = 1 Or Me.chkType(1).Value = 1 Or Me.chkType(2).Value = 1 Or Me.chkType(3).Value = 1 Or Me.chkType(4).Value = 1
    Me.cmdDele.Enabled = blnSel And blnData
    Me.cmdOk.Enabled = blnData And blnCheck And cbo���㷽ʽ.Text <> ""
    
End Sub
Private Function Load���㷽ʽ() As Boolean
    '-----------------------------------------------------------------------------------------
    '����:���ؽ��㷽ʽ
    '-----------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo errHand:
    gstrSQL = "Select ���㷽ʽ,ȱʡ��־ From ���㷽ʽӦ�� Where Ӧ�ó���='������' Order by ȱʡ��־ desc"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.EOF Then
        ShowMsgbox "δ���ý��㷽ʽ��Ӧ�ÿ�ϣ��뵽�۽��㷽ʽ�����������!"
        Exit Function
    End If
    With rsTemp
        Me.cbo���㷽ʽ.Clear
        Do While Not .EOF
            Me.cbo���㷽ʽ.AddItem Nvl(!���㷽ʽ)
            If Val(Nvl(!ȱʡ��־)) = 1 Then
                Me.cbo���㷽ʽ.ListIndex = Me.cbo���㷽ʽ.NewIndex
            End If
            .MoveNext
        Loop
        If Me.cbo���㷽ʽ.ListCount <> 0 And Me.cbo���㷽ʽ.ListIndex < 0 Then
            Me.cbo���㷽ʽ.ListIndex = 0
        End If
    End With
    Load���㷽ʽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Sub txt˵��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub

Private Function BuildingData(ByVal lng��Ӧ��ID As Long, ByVal str��Ӧ������ As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------
    '����:���ɹ�Ӧ�̵ĸ��
    '------------------------------------------------------------------------------------------------------------------------------
    Dim rsData As New ADODB.Recordset
    Dim str���� As String
    
    Dim dtStartdate As Date
    Dim dtEndDate As Date
    Err = 0: On Error GoTo errHand:
    
    str���� = Getϵͳ����
    If str���� = "" Then
        ShowMsgbox "δѡ�񱾴����ɵ�ϵͳ����!"
        Exit Function
    End If
    str���� = " And ϵͳ��ʶ in (" & str���� & ")"
    
    lbl��Ӧ��.Caption = "���ڻ�ȡ" & str��Ӧ������ & " ������...."
    
    gstrSQL = "Select  MAX(ID) ID, MAX(��¼״̬) ��¼״̬, '' �ƻ�����, ��Ʊ��, ��ⵥ�ݺ�, " & _
             "          SUM(Nvl(����, 0)) AS ����, " & _
             "          SUM(Nvl(��Ʊ���, 0)) AS ��Ʊ��� " & _
             "   FROM Ӧ����¼ " & _
             "   WHERE �ƻ����� IS NULL AND ��¼���� <> -1 AND ��λid+0 = [3] AND ������� IS NULL AND " & _
             "         ������� BETWEEN [1] AND [2] " & str���� & _
             "   GROUP BY ��¼����, NO, ��Ŀid,���,�������,��Ʊ��,��ⵥ�ݺ� " & _
             "   HAVING SUM(Nvl(��Ʊ���, 0)) <> 0 " & _
             "   ORDER BY ��Ʊ��"
    dtStartdate = Format(Me.dtpDate(0).Value, "yyyy-MM-DD hh:mm:ss")
    dtEndDate = Format(Me.dtpDate(1).Value, "yyyy-MM-DD hh:mm:ss")
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, dtStartdate, dtEndDate, lng��Ӧ��ID)
    '��ȡ���ݺż��������
    Dim lng������� As Long, strNO As String
    Dim dbl������ As Double, lngCount As Long
        
    Err = 0: On Error GoTo ErrRoll:
    gcnOracle.BeginTrans
    If rsData.RecordCount = 0 Then
        '������
        Call SetGridNewRowValue(str��Ӧ������, "", 0, 0)
    Else
        lng������� = zlDatabase.GetNextId("�����¼")
        strNO = zlDatabase.GetNextNo(31)
        dbl������ = 0
        With rsData
            Do While Not .EOF
                '���̲���
                '       ID_IN ,�ƻ����_IN(��0,1,2,3��ʽ����),
                '        �������_IN,Ԥ����_IN
            
                gstrSQL = "zl_�������_UPDATE(" & _
                     "'" & Nvl(!ID) & "'," & _
                    "NULL," & _
                    lng������� & "," & _
                    "0)"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                dbl������ = dbl������ + Val(Nvl(!��Ʊ���))
                lngCount = lngCount + 1
                lbl��Ӧ��.Caption = "���ڻ�ȡ" & str��Ӧ������ & " ������ " & lngCount & "...."
                .MoveNext
            Loop
        End With
        dbl������ = Round(dbl������, 2)
        '���浥��
        gstrSQL = "" & _
        "   zl_�������_INSERT('" & _
            strNO & "'," & _
            1 & "," & _
            0 & "," & _
            lng��Ӧ��ID & "," & _
            dbl������ & ",'" & _
            cbo���㷽ʽ.Text & "'," & _
            "NULL,'" & _
            gstrUserName & "',to_date('" & _
            Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')," & _
            lng������� & ",'" & _
            txt˵��.Text & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        Call SetGridNewRowValue(str��Ӧ������, strNO, lngCount, dbl������)
    End If
    gcnOracle.CommitTrans
    BuildingData = True
    Exit Function
errHand:
    '������
    Call SetGridNewRowValue(str��Ӧ������, "", 0, 0)
    If ErrCenter = 1 Then Resume
    Exit Function
ErrRoll:
    
    Call ErrCenter
    gcnOracle.RollbackTrans
End Function
Private Sub initGrid()
    '------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ͷ
    '------------------------------------------------------------------------------------------------------------------------------
    With mshBuilded
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "��Ӧ��"
        .TextMatrix(0, 1) = "�����"
        .TextMatrix(0, 2) = "��ϸ����"
        .TextMatrix(0, 3) = "�ܷ�Ʊ��"
        .TextMatrix(0, 4) = "˵��"
        
        .ColWidth(0) = 2000
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1400
        .ColWidth(4) = 1500
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColAlignment(4) = 1
    End With
End Sub
Private Sub SetGridNewRowValue(ByVal str��Ӧ�� As String, ByVal strNO As String, ByVal lngCount As Long, ByVal dbl��Ʊ�ܶ� As Double)
    '------------------------------------------------------------------------------------------------------------------------------------
    '����:����grid������ֵ
    '����; str��Ӧ��-��Ӧ��
    '      strNo-���ݺ�
    '      lngCount-��ϸ����
    '      dbl��Ʊ�ܶ�-��Ʊ�ܶ�
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With mshBuilded
        .Rows = .Rows + 1
        .row = .Rows - 2
        .TextMatrix(.row, 0) = str��Ӧ��
        .TextMatrix(.row, 1) = strNO
        .TextMatrix(.row, 2) = lngCount
        .TextMatrix(.row, 3) = Format(dbl��Ʊ�ܶ�, "###0.00;-###0.00;0;0")
        If strNO = "" Then
            .TextMatrix(.row, 4) = "�����ݷ���,δ���ɸ����"
            For i = 0 To .Cols - 1
                .Col = i
                .CellForeColor = vbRed
            Next
            .Col = 0
        End If
    End With
End Sub
Private Sub ShowBuildedGrid()
    '--------------------------------------------------------------------------
    '����:��ʾ���ɵ�����
    '--------------------------------------------------------------------------
    mshBuilded.Visible = True
    Me.lblInfor.Caption = "�������������Ǳ������ɸ�����!"
    Me.Caption = "�������ɽ��Ԥ��"
End Sub



Private Sub Ȩ�޿���()
    'Ȩ�޿���
    Dim blnҩƷ As Boolean
    Dim bln���� As Boolean
    Dim bln�豸 As Boolean
    Dim bln���� As Boolean
    Dim bln���� As Boolean
    
    blnҩƷ = InStr(1, mstrPrivs, ";ҩƷ;") <> 0
    bln���� = InStr(1, mstrPrivs, ";����;") <> 0
    bln�豸 = InStr(1, mstrPrivs, ";�豸;") <> 0
    bln���� = InStr(1, mstrPrivs, ";����;") <> 0
    bln���� = InStr(1, mstrPrivs, ";����;") <> 0
    
    chkType(0).Enabled = blnҩƷ
    chkType(0).Value = IIf(blnҩƷ, 1, 0)
    chkType(1).Enabled = bln����
    chkType(1).Value = IIf(bln����, 1, 0)
    chkType(2).Enabled = bln�豸
    chkType(2).Value = IIf(bln�豸, 1, 0)
    chkType(3).Enabled = bln����
    chkType(3).Value = IIf(bln����, 1, 0)
    chkType(4).Enabled = bln����
    chkType(4).Value = IIf(bln����, 1, 0)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt˵��, KeyAscii, m�ı�ʽ
End Sub
