VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelElement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ԫ��ѡ��"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4425
   Icon            =   "frmSelElement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdTab 
      Caption         =   "ר��ֽ"
      Height          =   300
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   3930
      Width           =   1335
   End
   Begin VB.CommandButton cmdTab 
      Caption         =   "������"
      Height          =   300
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   3645
      Width           =   1335
   End
   Begin VB.CommandButton cmdTab 
      Caption         =   "���ӱ�"
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdTab 
      Caption         =   "���ͼ"
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   1
      Top             =   3075
      Width           =   1335
   End
   Begin VB.CommandButton cmdTab 
      Caption         =   "�ı���"
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   2790
      Width           =   1335
   End
   Begin MSComctlLib.ImageList iLsTree32 
      Left            =   840
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelElement.frx":000C
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelElement.frx":08E6
            Key             =   "Attr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelElement.frx":0C00
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2295
      Index           =   3
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      SmallIcons      =   "iLsTree"
      ColHdrIcons     =   "iLsTree"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "˵��"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�ֺ�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ת�ı�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "����"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList iLsTree 
      Left            =   0
      Top             =   5000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelElement.frx":0D5A
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelElement.frx":0EB4
            Key             =   "Attr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelElement.frx":11CE
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2295
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      SmallIcons      =   "iLsTree"
      ColHdrIcons     =   "iLsTree"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "˵��"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�ֺ�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ת�ı�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "����"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2295
      Index           =   2
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      SmallIcons      =   "iLsTree"
      ColHdrIcons     =   "iLsTree"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "˵��"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�ֺ�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ת�ı�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "����"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2295
      Index           =   0
      Left            =   960
      TabIndex        =   8
      Top             =   840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      SmallIcons      =   "iLsTree"
      ColHdrIcons     =   "iLsTree"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "˵��"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�ֺ�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ת�ı�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "����"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2295
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      SmallIcons      =   "iLsTree"
      ColHdrIcons     =   "iLsTree"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "˵��"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�ֺ�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ת�ı�"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "����"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmSelElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pElementID As String, pElementName As String, pElementType As String
Public pDepartID As String, pFileType As Integer

Private clsDatabase As New zl9ComLib.clsDatabase
Private iCurrTab As Integer

Private Sub cmdTab_Click(Index As Integer)
    iCurrTab = Index
    ListItem
    
    Form_Resize
    On Error Resume Next
    Set lvwItem(iCurrTab).SelectedItem = lvwItem(iCurrTab).ListItems(1)
    lvwItem(iCurrTab).SetFocus
End Sub

Private Sub Form_Activate()
    pElementID = "": pElementName = "" ': pElementType = ""
    
    lvwItem(iCurrTab).SetFocus
End Sub

Private Sub Form_Load()
    Dim i As Integer
    iCurrTab = IIf(Len(pElementType) = 0, 0, pElementType)
    If Len(pElementType) > 0 Then
        For i = 0 To cmdTab.Count - 1
            cmdTab(i).Visible = IIf(i = iCurrTab, True, False)
        Next
    End If
    ListItem
    On Error Resume Next
    Set lvwItem(iCurrTab).SelectedItem = lvwItem(iCurrTab).ListItems(1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    '����Ԫ��ѡ�
    On Error Resume Next
    '   ��ʾѡ�
    For i = 0 To lvwItem.Count - 1
        lvwItem(i).Visible = IIf(i = iCurrTab, True, False)
        With cmdTab(i)
            If Len(pElementType) > 0 Then
                .Top = 0
            Else
                If i <= iCurrTab Then
                    .Top = i * (cmdTab(0).Height - 15)
                Else
                    .Top = Me.ScaleHeight - (lvwItem.Count - i) * (cmdTab(0).Height - 15)
                End If
            End If
            
            .Width = Me.ScaleWidth
        End With
    Next
    
    With lvwItem(iCurrTab)
        .Left = 0
        .Top = cmdTab(iCurrTab).Top + cmdTab(iCurrTab).Height
        .Width = Me.ScaleWidth
        If Len(pElementType) > 0 Then
            .Height = Me.ScaleHeight - .Top
        Else
            .Height = Me.ScaleHeight - (lvwItem.Count - iCurrTab - 1) * (cmdTab(0).Height - 15) - .Top
        End If
    End With
End Sub

Private Sub lvwItem_DblClick(Index As Integer)
    If lvwItem(Index).SelectedItem Is Nothing Then Exit Sub
    With lvwItem(Index).SelectedItem
        pElementID = Mid(.Key, 4)
        pElementName = .SubItems(2)
        pElementType = .SubItems(8)
    End With
    Unload Me
End Sub

Private Sub ListItem()
    Dim rsItem As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim i As Integer

    On Error Resume Next
    lvwItem(iCurrTab).ListItems.Clear
    
    If Len(pDepartID) = 0 Then
        If pFileType = -1 Then
            gstrSql = "Select a.*,decode(b.ID,'','',b.ID||'|'||b.����) As ��������,Decode(a.����,0,'�ı���',1,'���ӱ�',2,'������',3,'���ͼ',4,'ר��ֽ') As ����1 From ����Ԫ��Ŀ¼ a,���ű� b Where a.����=[1] And a.����ID=b.ID(+) Order by a.����"
            Set rsItem = zldatabase.OpenSQLRecord(gstrSql, "��ѯԪ��Ŀ¼", iCurrTab)
        Else
            gstrSql = "Select a.*,decode(b.ID,'','',b.ID||'|'||b.����) As ��������,Decode(a.����,0,'�ı���',1,'���ӱ�',2,'������',3,'���ͼ',4,'ר��ֽ') As ����1 From ����Ԫ��Ŀ¼ a,���ű� b Where a.����=[1] And ���� Like [2] And a.����ID=b.ID(+) Order by a.����"
            Set rsItem = zldatabase.OpenSQLRecord(gstrSql, "��ѯԪ��Ŀ¼", iCurrTab, String(pFileType, "_") + "1" + String(4 - pFileType, "_"))
        End If
    Else
        If pFileType = -1 Then
            gstrSql = "Select a.*,decode(b.ID,'','',b.ID||'|'||b.����) As ��������,Decode(a.����,0,'�ı���',1,'���ӱ�',2,'������',3,'���ͼ',4,'ר��ֽ') As ����1 From ����Ԫ��Ŀ¼ a,���ű� b Where a.����=[1] And (a.����ID Is Null Or a.����ID=[2]) And a.����ID=b.ID(+) Order by a.����"
            Set rsItem = zldatabase.OpenSQLRecord(gstrSql, "��ѯԪ��Ŀ¼", iCurrTab, pDepartID)
        Else
            gstrSql = "Select a.*,decode(b.ID,'','',b.ID||'|'||b.����) As ��������,Decode(a.����,0,'�ı���',1,'���ӱ�',2,'������',3,'���ͼ',4,'ר��ֽ') As ����1 From ����Ԫ��Ŀ¼ a,���ű� b Where a.����=[1] And (a.����ID Is Null Or a.����ID=[2]) And ���� Like [3] And a.����ID=b.ID(+) Order by a.����"
            Set rsItem = zldatabase.OpenSQLRecord(gstrSql, "��ѯԪ��Ŀ¼", iCurrTab, pDepartID, String(pFileType, "_") + "1" + String(4 - pFileType, "_"))
        End If
    End If

    Do While Not rsItem.EOF
        Set tmpItem = lvwItem(iCurrTab).ListItems.Add(, "Key" & rsItem("ID"), rsItem("����"))
        tmpItem.SubItems(1) = rsItem("����")
        tmpItem.SubItems(2) = rsItem("����")
        tmpItem.SubItems(3) = rsItem("˵��")
        tmpItem.SubItems(4) = rsItem("����")
        tmpItem.SubItems(5) = rsItem("�ֺ�")
        tmpItem.SubItems(6) = rsItem("ת�ı�")
        tmpItem.SubItems(7) = rsItem("��������")
        tmpItem.SubItems(8) = rsItem("����1")

        rsItem.MoveNext
    Loop
End Sub
