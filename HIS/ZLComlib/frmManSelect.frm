VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmManSelect 
   Caption         =   "��Աѡ��"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6645
   Icon            =   "frmManSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3660
      TabIndex        =   5
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5130
      TabIndex        =   4
      Top             =   4470
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -600
      TabIndex        =   3
      Top             =   4170
      Width           =   7770
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3030
      Top             =   2220
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
            Picture         =   "frmManSelect.frx":06EA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSelect.frx":0D36
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSelect.frx":1052
            Key             =   "WriteNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSelect.frx":1372
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManSelect.frx":190C
            Key             =   "Man"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   3120
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2595
      ScaleWidth      =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1620
      Width           =   30
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2955
      Left            =   3390
      TabIndex        =   1
      Top             =   570
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   3345
      Left            =   330
      TabIndex        =   0
      Top             =   330
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   5900
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmManSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrReturn As String

Dim msngStartX As Single    '�ƶ�ǰ����λ��
Dim mblnItem As Boolean  'Ϊ���ʾ������ListViewĳһ����
Dim mintColumn As Integer
Dim mblnLoad As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not lvwMain.SelectedItem Is Nothing Then
        With lvwMain.SelectedItem
            mstrReturn = Mid(.Key, 2) & ";" & .SubItems(1) & ";" & .Text
        End With
    Else
        mstrReturn = ""
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        FillTree
        tvwMain_NodeClick tvwMain.Nodes("Root")
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mstrReturn = ""
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    cmdCancel.Top = ScaleHeight - cmdCancel.Height - 200
    cmdOK.Top = cmdCancel.Top
    Frame1.Top = cmdOK.Top - 300
    
    
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 300
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    Frame1.Width = ScaleWidth + 800
    
    sngTop = Me.ScaleTop
    sngBottom = Frame1.Top
    
    
    tvwMain.Top = sngTop
    tvwMain.Height = IIf(sngBottom - tvwMain.Top > 0, sngBottom - tvwMain.Top, 0)
    tvwMain.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = tvwMain.Left + tvwMain.Width
    
    lvwMain.Left = picSplit.Left + picSplit.Width
    lvwMain.Top = sngTop
    lvwMain.Height = IIf(sngBottom - lvwMain.Top > 0, sngBottom - lvwMain.Top, 0)
    If Me.ScaleWidth - lvwMain.Left > 0 Then lvwMain.Width = Me.ScaleWidth - lvwMain.Left
    Me.Refresh
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub


Private Sub lvwMain_DblClick()
    If mblnItem = True Then cmdOK_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - msngStartX
        If sngTemp > 500 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 500 Then
            picSplit.Left = sngTemp
            tvwMain.Width = picSplit.Left - tvwMain.Left
            lvwMain.Left = picSplit.Left + picSplit.Width
            lvwMain.Width = Me.ScaleWidth - lvwMain.Left
        End If
        tvwMain.SetFocus
    End If
End Sub

Private Sub tvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
    FillList Node.Key
End Sub


Private Sub FillTree()
'����:װ�����в��ŵ�tvwMain
'����:
    Dim strTemp As String
    Dim rs���� As New ADODB.Recordset
    
    rs����.CursorLocation = adUseClient
    rs����.CursorType = adOpenKeyset
    rs����.LockType = adLockReadOnly
    
    strTemp = " where (����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or ����ʱ�� is null ) "
    rs����.Open "select id,�ϼ�id,���� ,����,����ʱ��  from ���ű� " & strTemp & " start with �ϼ�id is null connect by prior id =�ϼ�id ", gcnOracle
    tvwMain.Nodes.Clear
    tvwMain.Nodes.Add , , "Root", "���в���", "Root", "Root"
    tvwMain.Nodes("Root").Sorted = True
    Do Until rs����.EOF
        If CDate(IIf(IsNull(rs����("����ʱ��")), CDate("3000/1/1"), rs����("����ʱ��"))) = CDate("3000/1/1") Then
            strTemp = "Write"
        Else
            strTemp = "WriteNo"
        End If
        If IsNull(rs����("�ϼ�id")) Then
            tvwMain.Nodes.Add "Root", tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), strTemp, strTemp
        Else
            tvwMain.Nodes.Add "C" & rs����("�ϼ�id"), tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), strTemp, strTemp
        End If
        tvwMain.Nodes("C" & rs����("id")).Sorted = True
        rs����.MoveNext
    Loop
    tvwMain.Nodes("Root").Selected = True
    tvwMain.Nodes("Root").Expanded = True
End Sub

Public Sub FillList(ByVal str����ID As String)
'����:װ���Ӧ���ŵ���Ա��lvwMain
'����:str����ID ���ŵı�ʶ

    Dim rs��Ա As New ADODB.Recordset
    Dim fld As Field, lst As ListItem
    Dim strSQL As String
    
    Dim strIcon As String
    Set rs��Ա = New ADODB.Recordset
    rs��Ա.CursorLocation = adUseClient
    rs��Ա.CursorType = adOpenKeyset
    rs��Ա.LockType = adLockReadOnly
    
    If str����ID = "Root" Then
        strSQL = "" & _
            "   Select distinct a.ID,C.����ID,a.����,a.���,a.����,A.�Ա� ,b.���� as ���� " & _
            "   from ��Ա�� a,���ű� b,������Ա C " & _
            "   where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
            "        And  A.id=c.��ԱID and C.����id = b.id And nvl(C.ȱʡ,0)=1"
        rs��Ա.Open strSQL, gcnOracle
    Else
        strSQL = "" & _
            "   Select A.ID,A.����,A.���,A.����,A.�Ա�  " & _
            "   From ��Ա�� A,������Ա C " & _
            "   Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null )  " & _
            "        and C.��Աid=A.id   And C.����ID = " & Mid(str����ID, 2)
        
        rs��Ա.Open strSQL, gcnOracle
    End If
    '����������ͷ
    lvwMain.ColumnHeaders.Clear
    For Each fld In rs��Ա.Fields
       If InStr(fld.Name, "ID") = 0 Then lvwMain.ColumnHeaders.Add , fld.Name, fld.Name
    Next
    
    'ȡ������
    lvwMain.ListItems.Clear
    Do Until rs��Ա.EOF
        strIcon = IIf(rs��Ա("�Ա�") = "Ů", "Woman", "Man")
        Set lst = lvwMain.ListItems.Add(, "C" & rs��Ա("ID"), rs��Ա("����"), strIcon, strIcon)
        For Each fld In rs��Ա.Fields
            If InStr(fld.Name, "ID") = 0 And fld.Name <> "����" Then
                lst.SubItems(lvwMain.ColumnHeaders(fld.Name).Index - 1) = IIf(IsNull(fld.Value), "", fld.Value)
            End If
        Next
        If str����ID = "Root" Then
            lst.ListSubItems(1).Tag = rs��Ա("����ID")
        End If
        rs��Ա.MoveNext
    Loop
End Sub
